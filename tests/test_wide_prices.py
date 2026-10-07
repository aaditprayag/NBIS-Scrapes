import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import unittest

from nbis_scrapes.wide_prices import docs_column, fiscal_quarter, pivot_tokens


class WidePriceTests(unittest.TestCase):
    def test_fiscal_quarters_match_nebius_calendar_year(self):
        self.assertEqual(fiscal_quarter("2025-04-02"), "Q2 '25")
        self.assertEqual(fiscal_quarter("2025-06-30"), "Q2 '25")
        self.assertEqual(fiscal_quarter("2025-07-01"), "Q3 '25")
        self.assertEqual(fiscal_quarter("2025-10-01"), "Q4 '25")
        self.assertEqual(fiscal_quarter("2026-01-15"), "Q1 '26")
        self.assertEqual(fiscal_quarter("2026-07-19"), "Q3 '26")

    def test_before_a_rate_change_lands_on_the_prior_day(self):
        self.assertEqual(docs_column("before", "2026-06-01", "2026-07-15"), "2026-05-31")
        self.assertEqual(docs_column("from", "2026-06-01", "2026-07-15"), "2026-06-01")
        self.assertEqual(fiscal_quarter("2026-05-31"), "Q2 '26")
        self.assertEqual(fiscal_quarter("2026-06-01"), "Q2 '26")

    def test_token_rows_are_models_and_columns_are_dates(self):
        rows = [
            {"model_id": "Qwen/Qwen3-32B", "captured_on": "2026-07-19", "input_usd_per_million": "0.1"},
            {"model_id": "Qwen/Qwen3-32B", "captured_on": "2026-10-07", "input_usd_per_million": "0.1"},
            {"model_id": "zai-org/GLM-5.1", "captured_on": "2026-07-19", "input_usd_per_million": "1.4"},
        ]
        columns, grid = pivot_tokens(rows, "input_usd_per_million")
        self.assertEqual(columns, ["2026-07-19", "2026-10-07"])
        self.assertEqual(list(grid), ["Qwen/Qwen3-32B", "zai-org/GLM-5.1"])
        self.assertEqual(grid["zai-org/GLM-5.1"].get("2026-10-07", ""), "")
        self.assertEqual([fiscal_quarter(day) for day in columns], ["Q3 '26", "Q4 '26"])


if __name__ == "__main__":
    unittest.main()
