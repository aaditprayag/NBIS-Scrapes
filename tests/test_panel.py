import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import unittest

from nbis_scrapes.panel import (
    company_rows,
    load_ledger,
    quarter_index,
    summary_markdown,
    sum_energized_base,
    utilization_rows,
)


class PanelTests(unittest.TestCase):
    def setUp(self):
        self.ledger = load_ledger()
        self.rows = {row["quarter"]: row for row in company_rows(self.ledger)}

    def test_only_q4_2025_has_reported_active(self):
        reported = [row["quarter"] for row in self.rows.values() if row["reported_active_mw"] is not None]
        self.assertEqual(reported, ["Q4 2025"])
        self.assertEqual(self.rows["Q4 2025"]["reported_active_mw"], 170)
        self.assertEqual(self.rows["Q4 2025"]["reported_active_tag"], "reported")

    def test_connected_at_year_end_2025_is_a_target(self):
        row = self.rows["Q4 2025"]
        self.assertEqual(row["connected_base_mw"], 220)
        self.assertEqual(row["connected_tag"], "target")
        self.assertAlmostEqual(row["reported_active_over_connected"], 170 / 220, places=4)

    def test_q1_and_q2_2026_have_no_company_active_or_connected_actual(self):
        for quarter in ("Q1 2026", "Q2 2026"):
            row = self.rows[quarter]
            self.assertIsNone(row["reported_active_mw"])
            self.assertIsNone(row["connected_base_mw"])
            self.assertTrue(row["in_history"])

    def test_q3_2026_is_forward(self):
        self.assertFalse(self.rows["Q3 2026"]["in_history"])

    def test_site_sum_matches_identified_base(self):
        idx = quarter_index(self.ledger, "Q2 2026")
        identified = self.ledger["company"]["identified_energized_base"][idx]
        self.assertAlmostEqual(sum_energized_base(self.ledger, "Q2 2026"), identified, places=2)

    def test_ye26_guide_is_not_stored_as_reported_active(self):
        self.assertIsNone(self.rows["Q4 2026"]["reported_active_mw"])
        labels = [row["label"] for row in utilization_rows(self.ledger)]
        self.assertIn("YE2026 billing vs connected guide low", labels)
        low = next(row for row in utilization_rows(self.ledger) if row["denominator_mw"] == 800)
        self.assertGreater(low["ratio"], 0.2)
        self.assertLess(low["ratio"], 0.5)
        text = summary_markdown(self.ledger, list(self.rows.values()), utilization_rows(self.ledger))
        self.assertIn("77.3%", text)
        self.assertIn("32.2%", text)

    def test_unlocated_sites_have_no_coordinates(self):
        missing = [site["name"] for site in self.ledger["sites"] if site["lat"] is None]
        self.assertTrue(any(name.startswith("Oklahoma") for name in missing))
        self.assertTrue(any(name.startswith("France 2nd") for name in missing))


if __name__ == "__main__":
    unittest.main()
