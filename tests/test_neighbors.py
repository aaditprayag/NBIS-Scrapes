import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import unittest

from nbis_scrapes.neighbors import haversine_km, nearby


class NeighborTests(unittest.TestCase):
    def test_kansas_city_and_independence_are_close(self):
        km = haversine_km(39.0997, -94.5786, 39.0956, -94.3553)
        self.assertGreater(km, 10)
        self.assertLess(km, 30)

    def test_same_pin_is_dropped_and_the_next_building_is_kept(self):
        nebius = [{"name": "Mantsala", "country": "Finland", "lat": 60.63, "lon": 25.32, "confirmed": True}]
        facilities = [
            {"name": "Same hall", "owner": "Nebius", "city": "", "country": "FI", "lat": 60.6301, "lon": 25.3201,
             "website": "", "power_mw": "75", "power_basis": "Epoch", "source": "Epoch AI"},
            {"name": "Other hall", "owner": "Equinix", "city": "Helsinki", "country": "FI", "lat": 60.2, "lon": 24.9,
             "website": "", "power_mw": "", "power_basis": "", "source": "PeeringDB"},
        ]
        detail, summary = nearby(nebius, facilities)
        self.assertEqual(summary[0]["facilities_within_60km"], 1)
        self.assertEqual(detail[0]["neighbor"], "Other hall")
        self.assertIn("t=k", detail[0]["satellite_url"])


if __name__ == "__main__":
    unittest.main()
