import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import unittest

from nbis_scrapes.competitors import (
    campuses,
    merge_projects,
    quarter_index,
    quarterly_mw,
    ring,
    rtt_ms,
    site_rows,
)

SITE = {"key": "Kansas City", "name": "Kansas City MO (Patmos)", "country": "USA", "lat": 39.0997, "lon": -94.5786, "pin": "ledger"}


def _project(pid, lat, lon, phases, operator="Google", name="Google Kansas City"):
    return {"project_id": pid, "project": name, "operator": operator, "lat": lat, "lon": lon, "country": "US",
            "type": "hyperscaler", "phases": phases}


class CompetitorTests(unittest.TestCase):
    def test_rings_and_round_trip(self):
        self.assertEqual(ring(12.0), 1)
        self.assertEqual(ring(50.0), 1)
        self.assertEqual(ring(50.1), 2)
        self.assertEqual(ring(200.0), 2)
        self.assertIsNone(ring(200.1))
        self.assertEqual(rtt_ms(37.0), 0.37)

    def test_quarters(self):
        self.assertEqual(quarter_index("pre-2024"), 0)
        self.assertEqual(quarter_index("Q4-2019"), 0)
        self.assertEqual(quarter_index("Q3-2026"), 10)
        self.assertEqual(quarter_index("Q4-2030"), 27)
        self.assertIsNone(quarter_index(None))
        self.assertIsNone(quarter_index("2027"))

    def test_megawatts_build_up_by_quarter_and_stay_on(self):
        projects = [
            _project("near", 39.18, -94.47, [
                {"phase": "Building 1", "mw": 105, "online_quarter": "Q3-2026"},
                {"phase": "Building 2", "mw": 100, "online_quarter": "pre-2024"},
                {"phase": "Building 3", "mw": None, "online_quarter": "Q1-2027"},
                {"phase": "Building 4", "mw": 50, "online_quarter": None},
            ]),
            _project("far", 38.0, -94.6, [{"phase": "Campus", "mw": 300, "online_quarter": "Q1-2028"}], name="Ring two"),
            _project("out", 30.0, -94.6, [{"phase": "Campus", "mw": 900, "online_quarter": "Q1-2025"}], name="Too far"),
        ]
        rows = site_rows(projects, [SITE])
        self.assertEqual({row["project_id"] for row in rows}, {"near", "far"})
        counted = {row["phase"]: row["counted"] for row in rows if row["project_id"] == "near"}
        self.assertEqual(counted["Building 3"], "MW not public")
        self.assertEqual(counted["Building 4"], "no online date public")
        columns, ring_one = quarterly_mw(rows, [SITE], 1)
        cells = ring_one[SITE["name"]]
        self.assertEqual(columns[0], "2024-03-31")
        self.assertEqual(cells["2024-03-31"], "100")
        self.assertEqual(cells["2026-06-30"], "100")
        self.assertEqual(cells["2026-09-30"], "205")
        self.assertEqual(cells["2030-12-31"], "205")
        _columns, ring_two = quarterly_mw(rows, [SITE], 2)
        self.assertEqual(ring_two[SITE["name"]]["2027-12-31"], "0")
        self.assertEqual(ring_two[SITE["name"]]["2028-03-31"], "300")

    def test_merge_drops_repeats_and_the_same_project_from_another_slice(self):
        first = [_project("google-kc", 39.18, -94.47, [{"mw": 105, "online_quarter": "Q3-2026"}])]
        second = [
            _project("google-kc", 39.18, -94.47, [{"mw": 105, "online_quarter": "Q3-2026"}]),
            _project("google-kc-east", 39.181, -94.471, [{"mw": 105, "online_quarter": "Q3-2026"}]),
            _project("google-kc-mica", 39.32, -94.57, [{"mw": 700, "online_quarter": None}]),
        ]
        kept, dropped = merge_projects([first, second])
        self.assertEqual([project["project_id"] for project in kept], ["google-kc", "google-kc-mica"])
        self.assertEqual(len(dropped), 2)

    def test_campus_pins_come_from_the_ledger(self):
        sites = {site["key"]: site for site in campuses()}
        self.assertEqual((sites["Kansas City"]["lat"], sites["Kansas City"]["lon"]), (39.0997, -94.5786))
        self.assertTrue(sites["Kansas City"]["pin"].startswith("data/ledger.json"))
        self.assertFalse(sites["Hyderabad"]["pin"].startswith("data/ledger.json"))
        self.assertEqual(len(sites), 22)


if __name__ == "__main__":
    unittest.main()
