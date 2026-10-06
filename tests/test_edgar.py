import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import unittest

from nbis_scrapes.edgar import power_sentences

FIXTURE = Path(__file__).resolve().parents[1] / "data" / "fixtures" / "q1_2026_power.txt"


class EdgarTests(unittest.TestCase):
    def test_definitions_and_guide(self):
        hits = power_sentences(FIXTURE.read_text(encoding="utf-8"))
        joined = " ".join(hits).lower()
        self.assertIn("contracted power", joined)
        self.assertIn("connected power", joined)
        self.assertIn("active power", joined)
        self.assertIn("fully built and equipped", joined)
        self.assertTrue(any("800mw to 1gw" in hit.lower() for hit in hits))


if __name__ == "__main__":
    unittest.main()
