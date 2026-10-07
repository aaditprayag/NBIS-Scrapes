import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import unittest

from nbis_scrapes.token_factory import parse_catalog, parse_models_info, price_changes

FIXTURES = Path(__file__).resolve().parents[1] / "data" / "fixtures"


class TokenFactoryTests(unittest.TestCase):
    def test_catalog_unescapes_price_fields(self):
        text = (FIXTURES / "token_factory_catalog.md").read_text(encoding="utf-8")
        first = parse_catalog(text, "2026-07-19", "https://tokenfactory.nebius.com/model-catalog.md")[0]
        self.assertEqual(first["model_id"], "Qwen/Qwen3-32B")
        self.assertEqual(first["input_usd_per_million"], 0.1)
        self.assertEqual(first["output_usd_per_million"], 0.3)
        self.assertEqual(first["regions"], "eu-north1")

    def test_models_info_is_the_live_row(self):
        payload = json.loads((FIXTURES / "token_factory_models.json").read_text(encoding="utf-8"))
        row = parse_models_info(payload, "2026-10-07", "https://tokenfactory.nebius.com/api/public/models_info")[0]
        self.assertEqual(row["input_usd_per_million"], 0.15)
        self.assertEqual(row["quantization"], "fp8")

    def test_change_log_records_only_moves(self):
        text = (FIXTURES / "token_factory_catalog.md").read_text(encoding="utf-8")
        early = parse_catalog(text, "2026-07-19", "catalog")[0]
        later = parse_catalog(text, "2026-09-09", "catalog")[1]
        same = dict(early)
        same["captured_on"] = "2026-08-06"
        changes = price_changes([early, same, later])
        self.assertEqual(len(changes), 1)
        self.assertEqual(changes[0]["previous_input_usd_per_million"], 0.1)
        self.assertEqual(changes[0]["input_usd_per_million"], 0.12)
        self.assertEqual(changes[0]["captured_on"], "2026-09-09")


if __name__ == "__main__":
    unittest.main()
