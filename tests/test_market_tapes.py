import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import unittest

from nbis_scrapes.market_tapes import parse_coreweave, parse_crusoe, parse_lambda

CRUSOE = """
NVIDIA H200
141GB
HGX
$4.29/GPU-hr
NVIDIA H100
80GB
HGX
$3.90/GPU-hr
NVIDIA L40S
48GB
$1.50/GPU-hr
"""

COREWEAVE = """
NVIDIA HGX B300
8
270
$35.84
NVIDIA HGX B300
On-Demand Price:
Contact sales
NVIDIA HGX H100
8
80
On-Demand Price:
$49.24
NVIDIA HGX H200
8
141
On-Demand Price:
$50.44
NVIDIA HGX B200
8
180
On-Demand Price:
$68.80
"""

LAMBDA = """
Instances pricing
NVIDIA B200 SXM6
180 GB
$6.69
NVIDIA H100 SXM
80 GB
$3.99
"""


class MarketTapeTests(unittest.TestCase):
    def test_crusoe_keeps_the_per_gpu_card(self):
        rows = {row["sku"]: row["usd_per_gpu_hour"] for row in parse_crusoe(CRUSOE)}
        self.assertEqual(rows["H200"], 4.29)
        self.assertEqual(rows["H100"], 3.90)
        self.assertEqual(rows["L40S"], 1.50)

    def test_coreweave_divides_the_node_price(self):
        rows = {row["sku"]: row["usd_per_gpu_hour"] for row in parse_coreweave(COREWEAVE)}
        self.assertAlmostEqual(rows["H100"], 49.24 / 8, places=2)
        self.assertAlmostEqual(rows["H200"], 50.44 / 8, places=2)
        self.assertAlmostEqual(rows["B200"], 68.80 / 8, places=2)
        self.assertNotIn("B300", rows)

    def test_lambda_takes_the_instances_section(self):
        rows = {row["sku"]: row["usd_per_gpu_hour"] for row in parse_lambda(LAMBDA)}
        self.assertEqual(rows["B200"], 6.69)
        self.assertEqual(rows["H100"], 3.99)


if __name__ == "__main__":
    unittest.main()
