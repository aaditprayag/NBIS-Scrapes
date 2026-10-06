import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import unittest

from nbis_scrapes.demand import (
    parse_nebius_platforms,
    parse_nebius_prices,
    parse_openrouter_provider_page,
    summarize_vast_metrics,
)

FIXTURES = Path(__file__).resolve().parents[1] / "data" / "fixtures"


class DemandTests(unittest.TestCase):
    def test_price_hikes(self):
        parsed = parse_nebius_prices((FIXTURES / "nebius_prices.md").read_text(encoding="utf-8"))
        by_sku = {row["sku"]: row for row in parsed["on_demand"]}
        self.assertEqual(by_sku["NVIDIA HGX H200"]["usd_per_gpu_hour"], 4.50)
        self.assertEqual(by_sku["NVIDIA HGX H200"]["next_usd_per_gpu_hour"], 5.40)
        self.assertTrue(by_sku["NVIDIA GB300 NVL72"]["contact_sales"])
        hikes = {row["sku"]: row["change_pct"] for row in parsed["scheduled_hikes"]}
        self.assertAlmostEqual(hikes["NVIDIA HGX H200"], 0.20, places=2)
        self.assertNotIn("NVIDIA RTX PRO 6000", hikes)
        spot = {row["sku"]: row["spot_floor_usd_per_gpu_hour"] for row in parsed["spot"]}
        self.assertEqual(spot["NVIDIA HGX H100"], 0.79)

    def test_platforms_by_region(self):
        rows = parse_nebius_platforms((FIXTURES / "nebius_platforms.md").read_text(encoding="utf-8"))
        by_id = {row["platform"]: row["regions"] for row in rows}
        self.assertEqual(by_id["gpu-h100-sxm"], ["eu-north1"])
        self.assertIn("eu-north2", by_id["gpu-h200-sxm"])
        self.assertIn("us-central1", by_id["gpu-h200-sxm"])
        self.assertEqual(by_id["gpu-b200-sxm-a"], ["me-west1"])
        self.assertIn("us-north1", by_id["gpu-b300-sxm"])

    def test_html_pages_not_only_markdown(self):
        price_html = (
            "<div><p>NVIDIA GPU Instances</p><p>NVIDIA HGX H200</p><p>16</p>"
            "<p>200</p><p>$4.50</p><p>$5.40</p><p>Preemptible</p></div>"
        )
        parsed = parse_nebius_prices(price_html)
        row = parsed["on_demand"][0]
        self.assertEqual(row["usd_per_gpu_hour"], 4.50)
        self.assertEqual(row["next_usd_per_gpu_hour"], 5.40)
        platform_html = (
            "<tr><td>NVIDIA H200</td><td>gpu-h200-sxm</td><td>CPU</td>"
            "<td><code>eu-north1</code>, <code>us-central1</code></td></tr>"
        )
        platforms = parse_nebius_platforms(platform_html)
        self.assertEqual(platforms[0]["regions"], ["eu-north1", "us-central1"])

    def test_openrouter_count(self):
        parsed = parse_openrouter_provider_page((FIXTURES / "openrouter_nebius.html").read_text(encoding="utf-8"))
        self.assertEqual(parsed["models_listed"], 12)
        self.assertTrue(parsed["names_nebius"])

    def test_vast_summary(self):
        payload = json.loads((FIXTURES / "vast_metrics.json").read_text(encoding="utf-8"))
        summary = summarize_vast_metrics(payload)
        self.assertEqual(summary["count"], 2)
        h100 = next(row for row in summary["gpus"] if row["gpu"] == "H100_SXM")
        self.assertEqual(h100["supply"], 120)
        self.assertEqual(h100["demand"], 340)
        self.assertEqual(h100["price_median"], 2.1)


if __name__ == "__main__":
    unittest.main()
