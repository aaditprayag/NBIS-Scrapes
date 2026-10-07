import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import unittest

from nbis_scrapes.third_party_prices import (
    _zip_member,
    canonical_sku,
    corroborate,
    corroborated_days,
    nebius_list,
    parse_gpuhunt,
    parse_skypilot,
)

FIXTURES = Path(__file__).resolve().parents[1] / "data" / "fixtures"

DOCS = [
    {"captured_on": "2026-07-15", "platform_id": "nvidia-h100-nvlink", "item": "NVIDIA® H100 NVLink", "preemptible": "False",
     "vintage": "before", "vintage_date": "2026-06-01", "usd_per_gpu_hour": "2.95"},
    {"captured_on": "2026-07-15", "platform_id": "nvidia-h100-nvlink", "item": "NVIDIA® H100 NVLink", "preemptible": "False",
     "vintage": "from", "vintage_date": "2026-06-01", "usd_per_gpu_hour": "3.85"},
    {"captured_on": "2026-10-07", "platform_id": "", "item": "NVIDIA® H100 NVLink", "preemptible": "False",
     "vintage": "from", "vintage_date": "2026-10-01", "usd_per_gpu_hour": "4.5"},
    {"captured_on": "2026-05-27", "platform_id": "nvidia-h100-nvlink", "item": "NVIDIA® H100 NVLink. GPU", "preemptible": "False",
     "vintage": "before", "vintage_date": "2025-10-01", "usd_per_gpu_hour": "2.118"},
]

PAGE = [
    {"captured_on": "2024-11-15", "sku": "H100", "gpus": "1", "kind": "on_demand", "usd_per_gpu_hour": "2.95"},
    {"captured_on": "2026-10-06", "sku": "H100", "gpus": "1", "kind": "on_demand", "usd_per_gpu_hour": "3.85"},
]


def _observation(date: str, source: str, item: str, price: str, price_type: str = "on_demand") -> dict:
    return {
        "date": date, "source": source, "source_url": "", "observed_ts": date, "item": item, "region": "eu-north1",
        "price_type": price_type, "gpu_count": 1, "price_usd": price, "usd_per_gpu_hour": price, "raw": "",
    }


class ThirdPartyPriceTests(unittest.TestCase):
    def test_gpuhunt_keeps_the_cheapest_preset_per_gpu(self):
        text = (FIXTURES / "gpuhunt_nebius.csv").read_text(encoding="utf-8")
        rows = parse_gpuhunt(text, "2026-10-06", "catalog.zip", "20261006000000")
        found = {(row["item"], row["price_type"]): row["usd_per_gpu_hour"] for row in rows}
        self.assertEqual(found[("H100", "on_demand")], "4.5")
        self.assertEqual(found[("H100", "preemptible")], "2.15")
        self.assertEqual(found[("L40S", "on_demand")], "1.5484")
        self.assertNotIn(("", "on_demand"), found)

    def test_skypilot_splits_on_demand_and_spot_columns(self):
        text = (FIXTURES / "skypilot_nebius_vms.csv").read_text(encoding="utf-8")
        rows = parse_skypilot(text, "2025-05-13", "vms.csv", "20250513000000")
        eight = [row for row in rows if row["gpu_count"] == 8 and row["price_type"] == "on_demand"]
        self.assertEqual(eight[0]["usd_per_gpu_hour"], "2.95")
        spot = [row for row in rows if row["item"] == "H100" and row["price_type"] == "preemptible"]
        self.assertEqual(spot[0]["usd_per_gpu_hour"], "1.25")
        self.assertFalse([row for row in rows if row["item"] == "L40S" and row["price_type"] == "preemptible"])

    def test_names_follow_the_rate_card(self):
        self.assertEqual(canonical_sku("H100_SXM"), "H100")
        self.assertEqual(canonical_sku("h200-nvl"), "H200")
        self.assertEqual(canonical_sku("RTXPRO6000"), "RTX PRO 6000")
        self.assertEqual(canonical_sku("L40S", "gpu-l40s-a_1gpu-8vcpu-32gb"), "L40S-Intel")
        self.assertEqual(canonical_sku("L40S", "gpu-l40s-d 1gpu-16vcpu-96gb"), "L40S-AMD")
        self.assertEqual(canonical_sku("A100"), "A100")

    def test_rate_card_wins_over_a_stale_page_column(self):
        steps = nebius_list(DOCS, PAGE)
        rows = corroborate([
            _observation("2025-03-01", "gpuhunt", "H100", "2.95"),
            _observation("2026-06-10", "gpuhunt", "H100", "3.85"),
            _observation("2026-10-06", "gpuhunt", "H100", "4.5"),
            _observation("2026-10-06", "tracker", "H100", "3.85"),
            _observation("2024-01-10", "gpuhunt", "H100", "4.12"),
        ], steps)
        verdicts = [row["matches_nebius_list"] for row in rows]
        self.assertEqual(verdicts, ["yes", "yes", "yes", "no", "no Nebius list on file for that day"])
        self.assertEqual(rows[3]["nebius_list_usd"], "4.5")

    def test_only_a_match_makes_a_day_and_monthly_rollups_never_do(self):
        steps = nebius_list(DOCS, PAGE)
        rows = corroborate([
            _observation("2026-06-10", "SkyPilot catalog (git)", "H100", "3.85"),
            _observation("2026-06-10", "tracker", "H100", "2.15"),
            _observation("2026-06-11", "tracker", "H100", "2.15"),
            _observation("2026-07-01", "gpufinder.dev API (monthly floor)", "H100", "3.85"),
        ], steps)
        days = corroborated_days(rows, {"2026-06-10"})
        self.assertEqual([day["date"] for day in days], ["2026-06-10"])
        self.assertEqual(days[0]["compared_prices"], 2)
        self.assertEqual(days[0]["nebius_capture_that_day"], "yes")

    def test_zip_member_reads_a_stored_file(self):
        import io
        import zipfile
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
            archive.writestr("nebius.csv", "instance_name,price\n")
        self.assertEqual(_zip_member(buffer.getvalue(), "nebius.csv"), b"instance_name,price\n")


if __name__ == "__main__":
    unittest.main()
