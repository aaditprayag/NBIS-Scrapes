import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import unittest

from nbis_scrapes.peer_history import extract_nvidia_ondemand

CRUSOE_2024 = """
On-Demand
6-month reserved
NVIDIA A100 Tensor Core GPU
40GB
PCIe
$1.45
$1.16
NVIDIA A100 Tensor Core GPU
80GB
SXM
$1.95
$1.56
NVIDIA L40S Tensor Core GPU
48GB
$1.45
"""

COREWEAVE_CLASSIC = """
NVIDIA HGX H100
$4.76
NVIDIA H100 PCIe
$4.25
A100 80GB PCIe
$2.21
A100 40GB PCIe
$2.06
"""

LAMBDA_OLD = """
1x NVIDIA A100
$1.10 / hr
8x NVIDIA A100
$8.80 / hr
1x NVIDIA RTX A6000
$0.80 / hr
On-demand 8x NVIDIA H100 SXM
$2.99
"""

COREWEAVE_NODE = """
NVIDIA HGX H100
8
80
On-Demand Price:
$49.24
NVIDIA HGX B300
On-Demand Price:
Contact sales
"""


class PeerHistoryTests(unittest.TestCase):
    def test_crusoe_keeps_the_on_demand_column_and_the_memory_variant(self):
        rows = {row["sku"]: row["usd_per_gpu_hour"] for row in extract_nvidia_ondemand(CRUSOE_2024)}
        self.assertEqual(rows["A100 40GB PCIe"], 1.45)
        self.assertEqual(rows["A100 80GB SXM"], 1.95)
        self.assertEqual(rows["L40S"], 1.45)

    def test_classic_coreweave_is_already_per_gpu(self):
        rows = {row["sku"]: row["usd_per_gpu_hour"] for row in extract_nvidia_ondemand(COREWEAVE_CLASSIC)}
        self.assertEqual(rows["H100 SXM"], 4.76)
        self.assertEqual(rows["H100 PCIe"], 4.25)
        self.assertEqual(rows["A100 80GB PCIe"], 2.21)

    def test_lambda_divides_a_node_and_keeps_a_small_8x_rate(self):
        rows = {row["sku"]: row["usd_per_gpu_hour"] for row in extract_nvidia_ondemand(LAMBDA_OLD)}
        self.assertEqual(rows["A100"], 1.10)
        self.assertEqual(rows["RTX A6000"], 0.80)
        self.assertEqual(rows["H100 SXM"], 2.99)

    def test_current_coreweave_divides_the_node_and_skips_contact_sales(self):
        rows = {row["sku"]: row["usd_per_gpu_hour"] for row in extract_nvidia_ondemand(COREWEAVE_NODE)}
        self.assertAlmostEqual(rows["H100 SXM"], 6.155, places=2)
        self.assertNotIn("B300", rows)


if __name__ == "__main__":
    unittest.main()
