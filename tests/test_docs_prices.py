import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import unittest

from nbis_scrapes.docs_prices import parse_changelog, parse_docs_html, parse_markdown

FIXTURES = Path(__file__).resolve().parents[1] / "data" / "fixtures"


class DocsPriceTests(unittest.TestCase):
    def test_live_markdown_keeps_october_pair_and_skips_shekels(self):
        rows = parse_markdown(
            (FIXTURES / "docs_compute_oct.md").read_text(encoding="utf-8"),
            "https://docs.nebius.com/compute/resources/pricing.md",
            "2026-10-06",
        )
        b300 = [row for row in rows if row["platform_id"] == "gpu-b300-sxm" and not row["preemptible"]]
        self.assertEqual(
            [(row["vintage"], row["vintage_date"], row["usd_per_gpu_hour"]) for row in b300],
            [("from", "2026-10-01", 9.50), ("before", "2026-10-01", 7.85)],
        )
        self.assertEqual(b300[0]["regions"], "uk-south1,eu-west2,us-north1")
        self.assertTrue(all(row["usd_per_gpu_hour"] != 29.45 for row in rows))
        h200 = next(row for row in rows if "H200" in row["item"] and row["vintage"] == "from")
        self.assertEqual(h200["usd_per_gpu_hour"], 5.40)
        self.assertIn("eu-north1", h200["regions"])

    def test_archived_docs_html_keeps_june_pair(self):
        rows = parse_docs_html(
            (FIXTURES / "docs_compute_june.html").read_text(encoding="utf-8"),
            "https://docs.nebius.com/compute/resources/pricing",
            "2026-07-15",
        )
        b300 = [row for row in rows if row["platform_id"].endswith("b300-sxm") and not row["preemptible"]]
        self.assertEqual(
            [(row["vintage"], row["usd_per_gpu_hour"]) for row in b300],
            [("from", 7.85), ("before", 6.10)],
        )
        h100_before = next(row for row in rows if row["platform_id"] == "nvidia-h100-nvlink" and row["vintage"] == "before")
        self.assertEqual(h100_before["usd_per_gpu_hour"], 2.95)

    def test_changelog_dates_the_two_compute_changes(self):
        events = parse_changelog((FIXTURES / "docs_changelog_snippet.md").read_text(encoding="utf-8"))
        labels = [event["changelog_label"] for event in events]
        self.assertEqual(labels, ["September 28 — October 4, 2026", "May 25–31, 2026"])
        self.assertIn("June 1, 2026", events[1]["bullet"])


if __name__ == "__main__":
    unittest.main()
