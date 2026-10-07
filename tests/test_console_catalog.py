import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import unittest

from nbis_scrapes.console_catalog import gpu_platform, parse_full_view, region_of, region_table

FIXTURES = Path(__file__).resolve().parents[1] / "data" / "fixtures"
DATA = Path(__file__).resolve().parents[1] / "data"


class ConsoleCatalogTests(unittest.TestCase):
    def setUp(self):
        pages = json.loads((FIXTURES / "console_list_full_view.json").read_text(encoding="utf-8"))
        self.rows = {row["sku_name"]: row for row in parse_full_view(pages)}

    def test_price_and_dates_come_from_the_price_record(self):
        h100 = self.rows["compute/vm.gpu-h100-sxm/eu-north1"]
        self.assertEqual(h100["price_usd"], "4.5")
        self.assertEqual(h100["effective_from"], "2026-10-01")
        self.assertEqual(h100["price_created"], "2026-09-29")
        self.assertEqual(h100["sku_created"], "2025-09-11")
        self.assertEqual(h100["region"], "eu-north1")
        self.assertEqual(h100["pricing_unit"], "gpu-unit")

    def test_unsold_sku_is_priced_at_zero(self):
        rubin = self.rows["compute/vm.gpu-vera-rubin/us-central2"]
        self.assertEqual(rubin["price_usd"], "0")
        self.assertEqual(rubin["tiers"], "0:0")
        self.assertEqual(rubin["sku_created"], "2026-09-24")

    def test_names(self):
        self.assertEqual(region_of("bare-metal/server.gpu-vera-rubin/eu-north3"), "eu-north3")
        self.assertEqual(region_of("billing/e2e/fact.simple"), "")
        self.assertEqual(gpu_platform("compute/vm.gpu-h100-sxm.preemptible.gpu/eu-north1"), "gpu-h100-sxm")
        self.assertEqual(gpu_platform("compute/vm.standard-v3-l40s.gpu/eu-north1"), "standard-v3-l40s")
        self.assertEqual(gpu_platform("compute/vm.cpu-d3.ram/eu-north1"), "")

    def test_docs_regions_keep_the_crosswalk_and_new_ones_use_the_inference_file(self):
        crosswalk = json.loads((DATA / "region_crosswalk.json").read_text(encoding="utf-8"))
        inferences = json.loads((DATA / "console_region_inferences.json").read_text(encoding="utf-8"))
        table = {row["region"]: row for row in region_table(list(self.rows.values()), crosswalk, inferences)}
        self.assertEqual(table["eu-north1"]["place"], "Finland")
        self.assertEqual(table["eu-north1"]["on_docs_regions_page"], "yes")
        self.assertEqual(table["us-central2"]["on_docs_regions_page"], "no")
        self.assertTrue(table["us-central2"]["place_basis"].startswith("inferred"))
        self.assertEqual(table["us-central2"]["first_vera_rubin_sku"], "2026-09-24")
        self.assertEqual(table["us-central2"]["gpu_skus_at_zero"], 1)

    def test_inferences_never_repeat_a_docs_region(self):
        crosswalk = {entry["region"] for entry in json.loads((DATA / "region_crosswalk.json").read_text(encoding="utf-8"))}
        inferred = {entry["region"] for entry in json.loads((DATA / "console_region_inferences.json").read_text(encoding="utf-8"))}
        self.assertFalse(crosswalk & inferred)


if __name__ == "__main__":
    unittest.main()
