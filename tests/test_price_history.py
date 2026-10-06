import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import unittest

from nbis_scrapes.price_history import extract_rates, on_demand_changes

LAYOUT_A = """
GPU prices and configurations
On demand
Reserve
1 × H200 SXM
141
200
16
$3.50
$2.30
1 × H100 SXM
80
200
16
$2.95
$2.00
8 × H200 SXM
$28.00
$18.40
"""

LAYOUT_B = """
Price per GPU-hour
NVIDIA HGX B200
16
200
$5.50
NVIDIA HGX H200
16
200
$3.50
NVIDIA HGX H100
16
200
$2.95
NVIDIA H200 GPU
$2.30
/ hour
Access improved cost savings on NVIDIA H200 GPUs with a commitment of hundreds of units.
"""

LAYOUT_C = """
NVIDIA GPU Instances
Preemptible, GPU-hour
On-demand, GPU-hour
NVIDIA HGX H200
16
200
$2.45
$4.50
NVIDIA HGX H100
16
200
$2.15
$3.85
NVIDIA GB200 NVL72
112
800
Contact us
NVIDIA HGX H100
16
200
$2.15
$3.85
"""


class HistoryParserTests(unittest.TestCase):
    def _one(self, text, sku, kind):
        hits = [row for row in extract_rates(text) if row["sku"] == sku and row["kind"] == kind and row["gpus"] == 1]
        self.assertTrue(hits, f"missing {sku} {kind}")
        return hits[0]

    def test_2024_table_splits_on_demand_and_reserve(self):
        self.assertEqual(self._one(LAYOUT_A, "H200", "on_demand")["usd_per_gpu_hour"], 3.50)
        self.assertEqual(self._one(LAYOUT_A, "H200", "reserve")["usd_per_gpu_hour"], 2.30)
        self.assertEqual(self._one(LAYOUT_A, "H100", "on_demand")["usd_per_gpu_hour"], 2.95)
        eight = [row for row in extract_rates(LAYOUT_A) if row["gpus"] == 8 and row["kind"] == "on_demand"]
        self.assertEqual(eight[0]["usd_per_gpu_hour"], 3.50)

    def test_2025_single_column_and_commitment_card(self):
        self.assertEqual(self._one(LAYOUT_B, "B200", "on_demand")["usd_per_gpu_hour"], 5.50)
        self.assertEqual(self._one(LAYOUT_B, "H100", "on_demand")["usd_per_gpu_hour"], 2.95)
        self.assertEqual(self._one(LAYOUT_B, "H200", "commitment")["usd_per_gpu_hour"], 2.30)

    def test_2026_preemptible_column(self):
        self.assertEqual(self._one(LAYOUT_C, "H200", "preemptible")["usd_per_gpu_hour"], 2.45)
        self.assertEqual(self._one(LAYOUT_C, "H200", "on_demand")["usd_per_gpu_hour"], 4.50)
        self.assertEqual(self._one(LAYOUT_C, "H100", "on_demand")["usd_per_gpu_hour"], 3.85)
        self.assertEqual(self._one(LAYOUT_C, "H100", "preemptible")["usd_per_gpu_hour"], 2.15)
        contact = [row for row in extract_rates(LAYOUT_C) if row["sku"] == "GB200"]
        self.assertEqual(contact[0]["kind"], "contact_sales")

    def test_single_on_demand_column(self):
        text = """
NVIDIA GPU Instances
On-demand, GPU-hour
NVIDIA HGX B300
24
346
$6.10
NVIDIA HGX H200
16
200
$3.50
"""
        self.assertEqual(self._one(text, "B300", "on_demand")["usd_per_gpu_hour"], 6.10)
        self.assertEqual(self._one(text, "H200", "on_demand")["usd_per_gpu_hour"], 3.50)

    def test_change_log_keeps_first_print_and_moves(self):
        rows = [
            {"sku": "H200", "kind": "on_demand", "usd_per_gpu_hour": 3.5, "captured_on": "2025-06-29", "wayback_url": "a"},
            {"sku": "H200", "kind": "on_demand", "usd_per_gpu_hour": 3.5, "captured_on": "2026-03-04", "wayback_url": "b"},
            {"sku": "H200", "kind": "on_demand", "usd_per_gpu_hour": 4.5, "captured_on": "2026-07-12", "wayback_url": "c"},
        ]
        changes = on_demand_changes(rows)
        self.assertEqual([row["captured_on"] for row in changes], ["2025-06-29", "2026-07-12"])
        self.assertEqual(changes[1]["previous_usd"], 3.5)


if __name__ == "__main__":
    unittest.main()
