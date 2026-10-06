import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import unittest

from nbis_scrapes.mapbuild import build_features, render_html
from nbis_scrapes.panel import load_ledger


class MapTests(unittest.TestCase):
    def setUp(self):
        self.ledger = load_ledger()
        self.geo = build_features(self.ledger)

    def test_finland_is_pinned_and_oklahoma_is_not(self):
        names = [feature["properties"]["name"] for feature in self.geo["features"]]
        self.assertTrue(any(name.startswith("Mantsala FI #1") for name in names))
        self.assertFalse(any(name.startswith("Oklahoma") for name in names))
        unmapped = [site["name"] for site in self.geo["unmapped"]]
        self.assertTrue(any(name.startswith("Oklahoma") for name in unmapped))

    def test_unconfirmed_lead_stays_flagged(self):
        lead = next(
            feature["properties"]
            for feature in self.geo["features"]
            if feature["properties"]["name"].startswith("UK 3rd")
        )
        self.assertFalse(lead["confirmed"])

    def test_html_contains_the_two_quarters(self):
        html = render_html(self.ledger)
        self.assertIn("Q2 2026", html)
        self.assertIn("Q4 2026", html)
        self.assertIn("Mantsala FI #1", html)
        self.assertNotIn("/*__DATA__*/", html)


if __name__ == "__main__":
    unittest.main()
