import json
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import unittest

from nbis_scrapes.availability import (
    build_wide,
    farm_queries,
    farm_rows_from_csv,
    ionet_rows,
    parse_500farm_counts,
    parse_500farm_prices,
    parse_akash_breakdown,
    parse_ionet_market,
    import_observations,
    merge_rows,
    normalize_gpu,
    parse_digitized_chart,
    parse_gpuhunt_runpod,
    parse_runpod_stock,
    parse_shadeform_types,
    read_long,
    read_versions,
    runpod_rows,
    shadeform_rows,
    summarize_vast_offers,
    vast_metric_rows,
    vast_rows,
    write_long,
)
from nbis_scrapes.demand import summarize_vast_metrics

FIXTURES = Path(__file__).resolve().parents[1] / "data" / "fixtures"


def _read(name: str) -> str:
    return (FIXTURES / name).read_text(encoding="utf-8")


def _value(rows, **match):
    hits = [row["value"] for row in rows if all(row[key] == value for key, value in match.items())]
    return hits[0] if len(hits) == 1 else hits


class AvailabilityTests(unittest.TestCase):
    def test_gpu_names(self):
        self.assertEqual(normalize_gpu("NVIDIA H100 80GB HBM3"), "H100")
        self.assertEqual(normalize_gpu("H100 SXM"), "H100")
        self.assertEqual(normalize_gpu("H100_nvl"), "H100 NVL")
        self.assertEqual(normalize_gpu("NVIDIA H100 PCIe"), "H100 PCIe")
        self.assertEqual(normalize_gpu("NVIDIA B200"), "B200")
        self.assertEqual(normalize_gpu("GB200 NVL72"), "GB200")
        self.assertEqual(normalize_gpu("A6000"), "A6000")

    def test_chart_keeps_only_read_points(self):
        rows = parse_digitized_chart(_read("availability_3fourteen_b200.csv"))
        days = [row["date"] for row in rows]
        self.assertEqual(days, ["2026-01-01", "2026-01-07", "2026-01-08", "2026-01-09"])
        self.assertEqual(rows[0]["value"], "62.28")
        self.assertEqual(rows[0]["gpu"], "B200")
        self.assertEqual(rows[0]["unit"], "pct")

    def test_gpuhunt_runpod_counts_stock(self):
        seen = set()
        rows = parse_gpuhunt_runpod(_read("availability_gpuhunt_runpod.csv"), "2026-08-01", seen=seen)
        self.assertEqual(_value(rows, gpu="B200", metric="secure_data_centers_in_stock"), "3")
        self.assertEqual(_value(rows, gpu="B200", metric="offers_in_stock"), "4")
        self.assertEqual(_value(rows, gpu="B200", metric="largest_gpu_count_in_stock"), "2")
        self.assertEqual(_value(rows, gpu="B200", metric="community_countries_in_stock"), "0")
        self.assertEqual(_value(rows, gpu="H200", metric="community_countries_in_stock"), "1")
        self.assertEqual(_value(rows, gpu="H200", metric="secure_data_centers_in_stock"), "0")
        # A100 is not tracked, and H100 has never appeared, so neither gets a zero.
        self.assertFalse([row for row in rows if row["gpu"] in ("A100", "H100")])
        self.assertEqual(seen, {"B200", "H200"})

    def test_gpuhunt_zero_after_first_sighting(self):
        header = "instance_name,location,price,cpu,memory,gpu_count,gpu_name,gpu_memory,spot,disk_size\n"
        text = header + "NVIDIA H200,EU-SE-1,3.99,16,200,1,H200,141,False,\n"
        rows = parse_gpuhunt_runpod(text, "2025-01-01", seen={"B200"})
        self.assertEqual(_value(rows, gpu="B200", metric="secure_data_centers_in_stock"), "0")
        self.assertEqual(_value(rows, gpu="H200", metric="secure_data_centers_in_stock"), "1")
        # No community offers that day, so no community count at all.
        self.assertFalse([row for row in rows if row["metric"] == "community_countries_in_stock"])
        self.assertEqual(parse_gpuhunt_runpod(header, "2025-01-02", seen={"B200"}), [])

    def test_import_keeps_tracked_chips(self):
        text = (
            "date,source,source_url,observed_ts,provider,gpu,region,metric,value,unit,raw\n"
            "2026-09-08,github:x/y,u,t,lambda,NVIDIA B200,,regions_with_capacity_8x,2,count,gpu_8x_b200\n"
            "2026-09-08,github:x/y,u,t,lambda,A10,,regions_with_capacity_1x,5,count,gpu_1x_a10\n"
            "2026-09-09,github:x/y,u,t,lambda,H100 SXM,,regions_with_capacity_8x,,count,empty\n"
            "2026-09-10,github:x/z,u,t,runpod,B200,secure,stock_status,Low,text,r\n"
            "2026-09-10,github:x/w,u,t,vast,B200,global,offers,12,count,r\n"
        )
        rows = import_observations(text)
        self.assertEqual(len(rows), 3)
        self.assertEqual((rows[1]["metric"], rows[1]["region"]), ("stock_status_secure", ""))
        self.assertEqual(rows[2]["region"], "")
        self.assertEqual(rows[0]["gpu"], "B200")
        self.assertEqual(rows[0]["note"], "gpu_8x_b200")

    def test_500farm_counts_and_prices(self):
        answers = json.loads(_read("availability_500farm_query_range.json"))
        rows = parse_500farm_counts(answers["count_rented"], answers["count_available"])
        self.assertEqual(_value(rows, date="2026-10-07", gpu="B200", metric="gpus_rented"), "313")
        self.assertEqual(_value(rows, date="2026-10-07", gpu="B200", metric="gpus_available"), "44")
        self.assertEqual(_value(rows, date="2026-10-07", gpu="B200", metric="gpus_listed"), "357")
        self.assertEqual(_value(rows, date="2026-10-07", gpu="B200", metric="pct_available"), "12.32")
        self.assertEqual(_value(rows, date="2026-10-06", gpu="H100", metric="gpus_listed"), "213")
        prices = parse_500farm_prices(answers["price_rented"], answers["price_available"])
        self.assertEqual(_value(prices, date="2026-10-07", gpu="B200", metric="median_usd_per_gpu_hour_rented"), "7.556")
        self.assertEqual(_value(prices, date="2026-10-06", gpu="H100", metric="median_usd_per_gpu_hour_available"), "3.4093")
        queries = farm_queries(["B200"])
        self.assertIn('rented="yes",gpu_name=~"B200"', queries["count_rented"])
        self.assertIn('gpu_count_range="8+"', queries["price_available"])

    def test_500farm_csv_rows(self):
        counts = (
            "date,gpu,gpus_rented,gpus_available,gpus_listed,pct_available,metric\n"
            "2026-10-07,B200,313,44,357,12.32,vastai_gpu_count\n"
            "2026-10-07,RTX 4090,1,1,2,50,vastai_gpu_count\n"
        )
        prices = (
            "date,gpu,datacenter,gpu_count_range,offer_state,stat,usd_per_gpu_hour,metric\n"
            "2026-10-07,B200,yes,8+,rented,median,7.5559722,vastai_v2\n"
            "2026-10-07,B200,yes,8+,rented,p10,6.1,vastai_v2\n"
            "2026-10-07,B200,no,8+,available,median,5.0,vastai_v2\n"
        )
        rows = farm_rows_from_csv(counts, prices)
        self.assertEqual(len([row for row in rows if row["metric"] == "pct_available"]), 1)
        self.assertEqual(_value(rows, metric="median_usd_per_gpu_hour_rented"), "7.556")
        self.assertFalse([row for row in rows if row["gpu"] == "RTX 4090"])
        self.assertFalse([row for row in rows if row["metric"] == "median_usd_per_gpu_hour_available"])

    def test_akash_breakdown(self):
        rows = parse_akash_breakdown(json.loads(_read("availability_akash_gpu_breakdown.json")))
        self.assertEqual(_value(rows, date="2026-10-07", gpu="H100", metric="gpus_total"), "72")
        self.assertEqual(_value(rows, date="2026-10-07", gpu="H100", metric="gpus_leased"), "46")
        self.assertEqual(_value(rows, date="2026-10-07", gpu="H100", metric="utilization_pct"), "63.89")
        self.assertEqual(_value(rows, date="2026-10-06", gpu="B300", metric="utilization_pct"), "100")
        self.assertEqual(_value(rows, date="2026-10-07", gpu="A100", metric="gpus_total"), "201")

    def test_ionet_market(self):
        summary = parse_ionet_market(json.loads(_read("availability_ionet_market_snapshot.json")))
        self.assertEqual(summary["H100"]["gpus_total"], 512)
        self.assertEqual(summary["H100"]["hired_pct"], 100.0)
        self.assertEqual(summary["H100 PCIe"]["gpus_hired"], 8)
        self.assertNotIn("RTX 4090", summary)
        rows = ionet_rows("2026-10-07", summary)
        self.assertEqual(_value(rows, gpu="B300", metric="price_usd_per_gpu_hour"), "4.48")

    def test_versions_list(self):
        pairs = read_versions("# comment\n2026-08-01 https://x/v2/20260801-1/catalog.zip\n\n")
        self.assertEqual(pairs, [("2026-08-01", "https://x/v2/20260801-1/catalog.zip")])

    def test_runpod_stock(self):
        levels = parse_runpod_stock(json.loads(_read("availability_runpod_stock.json")))
        self.assertIsNone(levels["B200"])
        self.assertEqual(levels["H100"], "Low")
        self.assertEqual(levels["H200"], "Low")
        self.assertNotIn("A100", levels)
        rows = runpod_rows("2026-10-07", levels, {"US-CA-2": levels, "EU-RO-1": {"B200": "High"}}, {"US": {}})
        self.assertEqual(_value(rows, gpu="B200", metric="secure_stock_level"), "None")
        self.assertEqual(_value(rows, gpu="B200", metric="secure_data_centers_in_stock"), "1")
        self.assertEqual(_value(rows, gpu="H100", metric="secure_data_centers_in_stock"), "1")
        self.assertEqual(_value(rows, gpu="B200", metric="stock_level", region="EU-RO-1"), "High")
        self.assertEqual(_value(rows, gpu="H100", metric="community_countries_in_stock"), "0")

    def test_vast_offers(self):
        summary = summarize_vast_offers(
            json.loads(_read("availability_vast_rentable.json")),
            json.loads(_read("availability_vast_unrentable.json")),
        )
        b200 = summary["B200"]
        self.assertEqual(b200["machines_listed"], 6)
        self.assertEqual(b200["gpus_listed"], 48)
        self.assertEqual(b200["gpus_rentable"], 15)
        self.assertEqual(b200["rentable_pct"], 31.25)
        self.assertEqual(b200["gpus_not_rentable"], 33)
        self.assertEqual(b200["offers_rentable"], 6)
        # Six rentable offers, three near $6.25 and three near $7.50 a GPU-hour.
        self.assertAlmostEqual(b200["median_usd_per_gpu_hour"], 6.8767, places=4)
        rows = vast_rows("2026-10-07", summary)
        self.assertEqual(_value(rows, gpu="B200", metric="rentable_pct"), "31.25")

    def test_vast_metrics_rows(self):
        parsed = summarize_vast_metrics(json.loads(_read("vast_metrics.json")))
        rows = vast_metric_rows("2026-10-07", parsed)
        self.assertEqual(_value(rows, gpu="H100", metric="supply"), "120")
        self.assertEqual(_value(rows, gpu="H100", metric="demand"), "340")

    def test_shadeform(self):
        records = parse_shadeform_types(json.loads(_read("availability_shadeform_types.json")))
        self.assertFalse([record for record in records if record["gpu"] == "A6000"])
        rows = shadeform_rows("2026-10-07", records)
        self.assertEqual(_value(rows, gpu="H100", provider="all", metric="listings"), "28")
        self.assertEqual(_value(rows, gpu="H100", provider="all", metric="listings_available"), "3")
        self.assertEqual(_value(rows, gpu="H100", provider="all", metric="listings_available_pct"), "10.71")
        self.assertEqual(_value(rows, gpu="H100", provider="nebius", metric="listings_available"), "0")
        self.assertEqual(_value(rows, gpu="B200", provider="lambdalabs", metric="listings_available"), "0")
        self.assertEqual(_value(rows, gpu="H100", provider="lambdalabs", metric="available", region="saltlakecity-usa-1"), "1")
        self.assertEqual(_value(rows, gpu="H100", provider="lambdalabs", metric="available", region="austin-usa-1"), "0")

    def test_merge_has_no_duplicates(self):
        first = parse_digitized_chart(_read("availability_3fourteen_b200.csv"))
        again = parse_digitized_chart(_read("availability_3fourteen_b200.csv"))
        merged = merge_rows(first, again)
        self.assertEqual(len(merged), len(first))
        changed = dict(again[0], value="61")
        merged = merge_rows(merged, [changed])
        self.assertEqual(len(merged), len(first))
        self.assertEqual(merged[0]["value"], "61")

    def test_wide_files(self):
        rows = parse_digitized_chart(_read("availability_3fourteen_b200.csv"))
        rows += shadeform_rows("2026-01-08", parse_shadeform_types(json.loads(_read("availability_shadeform_types.json"))))
        with tempfile.TemporaryDirectory() as folder:
            long_path = Path(folder) / "long.csv"
            write_long(rows, long_path)
            self.assertEqual(len(read_long(long_path)), len(rows))
            written = build_wide(read_long(long_path), Path(folder))
            names = sorted(path.name for path in written)
            self.assertEqual(names, ["gpu_availability_b200.csv", "gpu_availability_h100.csv"])
            lines = (Path(folder) / "gpu_availability_b200.csv").read_text(encoding="utf-8").splitlines()
            self.assertTrue(lines[0].startswith("Fiscal quarter,Q1 '26"))
            self.assertEqual(lines[1], "Availability measure,2026-01-01,2026-01-07,2026-01-08,2026-01-09")
            chart = next(line for line in lines if line.startswith("3Fourteen"))
            self.assertTrue(chart.endswith(",62.28,56.86,55.86,53.64"))
            shade = next(line for line in lines if "(all clouds): listings available (count)" in line)
            self.assertEqual(shade.split(",")[-4:], ["", "", "0", ""])
            # Region-level flags stay in the long file.
            self.assertFalse([line for line in lines if "saltlakecity" in line])


if __name__ == "__main__":
    unittest.main()
