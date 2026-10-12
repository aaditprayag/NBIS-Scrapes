"""Build the quarterly megawatt tables from the site ledger."""

from __future__ import annotations

import csv
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
LEDGER_PATH = ROOT / "data" / "ledger.json"
OUT_DIR = ROOT / "data" / "out"

# Company guide from the 12 Feb 2026 and 13 May 2026 letters. Not a quarterly path.
YE26_CONNECTED_GUIDE_MW = (800, 1000)
# "More than 1 GW per year" from 2027, Q2 2026 letter. A pace, not a quarter.
FROM_2027_DEPLOYMENT_PACE_MW = 1000


def load_ledger(path: Path | None = None) -> dict:
    path = path or LEDGER_PATH
    return json.loads(path.read_text(encoding="utf-8"))


def quarter_index(ledger: dict, quarter: str) -> int:
    return ledger["quarters"].index(quarter)


def in_history(ledger: dict, quarter: str) -> bool:
    return quarter_index(ledger, quarter) <= quarter_index(ledger, ledger["history_end"])


def _lookup(rows: list[dict], quarter: str) -> dict | None:
    for row in rows:
        if row["quarter"] == quarter:
            return row
    return None


def _blank(value):
    return "" if value is None else value


def company_rows(ledger: dict) -> list[dict]:
    company = ledger["company"]
    reported = {row["quarter"]: row for row in company["reported_active"]}
    connected = {row["quarter"]: row for row in company["connected_base"]}
    contracted = {row["quarter"]: row for row in company["contracted_company"]}
    identified = company["identified_energized_base"]
    billing = company["billing_base"]
    rows = []
    for i, quarter in enumerate(ledger["quarters"]):
        rep = reported.get(quarter)
        con = connected.get(quarter)
        ctr = contracted.get(quarter)
        ratio = None
        ratio_basis = ""
        if rep and con and con["value"]:
            ratio = round(rep["value"] / con["value"], 4)
            ratio_basis = f"reported active / connected ({con['tag']})"
        rows.append({
            "quarter": quarter,
            "in_history": in_history(ledger, quarter),
            "identified_energized_base_mw": identified[i],
            "reported_active_mw": None if rep is None else rep["value"],
            "reported_active_tag": "" if rep is None else rep["tag"],
            "connected_base_mw": None if con is None else con["value"],
            "connected_tag": "" if con is None else con["tag"],
            "contracted_company_mw": None if ctr is None else ctr["value"],
            "contracted_tag": "" if ctr is None else ctr["tag"],
            "billing_base_mw": billing[i],
            "reported_active_over_connected": ratio,
            "ratio_basis": ratio_basis,
        })
    return rows


def site_rows(ledger: dict) -> list[dict]:
    rows = []
    for site in ledger["sites"]:
        for i, quarter in enumerate(ledger["quarters"]):
            rows.append({
                "site_id": site["id"],
                "site": site["name"],
                "country": site["country"],
                "ownership": site["ownership"],
                "confirmed": site["confirmed"],
                "quarter": quarter,
                "in_history": in_history(ledger, quarter),
                "basis_energized": site["basis_energized"],
                "contracted_mw": site["contracted_mw"][i],
                "energized_low_mw": site["energized_low"][i],
                "energized_base_mw": site["energized_base"][i],
                "energized_high_mw": site["energized_high"][i],
            })
    return rows


def sum_energized_base(ledger: dict, quarter: str) -> float:
    idx = quarter_index(ledger, quarter)
    total = 0.0
    for site in ledger["sites"]:
        value = site["energized_base"][idx]
        if isinstance(value, (int, float)):
            total += value
    return round(total, 4)


def utilization_rows(ledger: dict) -> list[dict]:
    """Ratios worth reading. Each row names the numerator and the denominator."""
    rows = []
    by_q = {row["quarter"]: row for row in company_rows(ledger)}
    ye25 = by_q["Q4 2025"]
    rows.append({
        "label": "YE2025 reported active vs connected target",
        "quarter": "Q4 2025",
        "numerator_mw": ye25["reported_active_mw"],
        "numerator": "reported active power",
        "denominator_mw": ye25["connected_base_mw"],
        "denominator": "year-end 2025 connected target (220 MW). The actual was not disclosed.",
        "ratio": ye25["reported_active_over_connected"],
        "tag": "reported_over_target",
    })

    idx = quarter_index(ledger, "Q4 2026")
    billing = ledger["company"]["billing_base"][idx]
    connected_model = by_q["Q4 2026"]["connected_base_mw"]
    pairs = [
        ("YE2026 billing vs connected guide low", 800, "company connected guide, low end"),
        ("YE2026 billing vs connected guide high", 1000, "company connected guide, high end"),
        ("YE2026 billing vs model connected base", connected_model, "capacity-model connected base, estimated"),
    ]
    for label, denom, denom_name in pairs:
        rows.append({
            "label": label,
            "quarter": "Q4 2026",
            "numerator_mw": round(billing, 4) if billing is not None else None,
            "numerator": "model billing MW (IT), estimated, about one quarter after energization",
            "denominator_mw": denom,
            "denominator": denom_name,
            "ratio": round(billing / denom, 4) if billing and denom else None,
            "tag": "estimated",
        })

    ident_26 = ledger["company"]["identified_energized_base"][idx]
    ident_27 = ledger["company"]["identified_energized_base"][quarter_index(ledger, "Q4 2027")]
    addition = None
    if ident_26 is not None and ident_27 is not None:
        addition = round(ident_27 - ident_26, 4)
    rows.append({
        "label": "2027 identified-energized addition vs >1 GW/year pace",
        "quarter": "Q4 2027",
        "numerator_mw": addition,
        "numerator": "change in identified energized base from Q4 2026 to Q4 2027",
        "denominator_mw": FROM_2027_DEPLOYMENT_PACE_MW,
        "denominator": "company pace: more than 1 GW per year brought to market from 2027",
        "ratio": round(addition / FROM_2027_DEPLOYMENT_PACE_MW, 4) if addition else None,
        "tag": "estimated_vs_guide",
    })
    return rows


def _write_csv(path: Path, rows: list[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if not rows:
        path.write_text("", encoding="utf-8")
        return
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        for row in rows:
            writer.writerow({key: _blank(value) for key, value in row.items()})


def _fmt(value) -> str:
    if value is None or value == "":
        return "—"
    if isinstance(value, float):
        return f"{value:.1f}".rstrip("0").rstrip(".")
    return str(value)


def _pct(value) -> str:
    if not isinstance(value, (int, float)):
        return "—"
    return f"{value * 100:.1f}%"


def summary_markdown(ledger: dict, company: list[dict], util: list[dict]) -> str:
    lines = [
        "# Nebius quarterly power panel",
        "",
        "History runs through Q2 2026. Later quarters are the capacity model's forward path.",
        "Reported active power is blank in every quarter except Q4 2025.",
        "Connected figures are a target or an estimate. Nebius has not printed a connected-power actual.",
        "",
        "| Quarter | History | Energized base | Reported active | Connected | Tag | Contracted | Tag | Billing |",
        "|---|---|---:|---:|---:|---|---:|---|---:|",
    ]
    for row in company:
        if row["quarter"].endswith("2028") or row["quarter"].endswith("2029") or row["quarter"].endswith("2030"):
            continue
        if row["quarter"] not in {
            "Q1 2024", "Q2 2024", "Q3 2024", "Q4 2024",
            "Q1 2025", "Q2 2025", "Q3 2025", "Q4 2025",
            "Q1 2026", "Q2 2026", "Q3 2026", "Q4 2026",
            "Q2 2027", "Q4 2027",
        } and not row["quarter"].startswith("Q"):
            continue
        # Keep every quarter through 2027 Q4, skip intra-2028+ already handled.
        if quarter_index(ledger, row["quarter"]) > quarter_index(ledger, "Q4 2027"):
            continue
        lines.append(
            "| {quarter} | {hist} | {en} | {act} | {con} | {ctag} | {ctr} | {kgtag} | {bill} |".format(
                quarter=row["quarter"],
                hist="yes" if row["in_history"] else "forward",
                en=_fmt(row["identified_energized_base_mw"]),
                act=_fmt(row["reported_active_mw"]),
                con=_fmt(row["connected_base_mw"]),
                ctag=row["connected_tag"] or "—",
                ctr=_fmt(row["contracted_company_mw"]),
                kgtag=row["contracted_tag"] or "—",
                bill=_fmt(row["billing_base_mw"]),
            )
        )
    lines.extend(["", "## Utilization", ""])
    for row in util:
        lines.append(
            f"- **{row['label']}:** {_fmt(row['numerator_mw'])} / {_fmt(row['denominator_mw'])} = {_pct(row['ratio'])} ({row['tag']}). {row['denominator']}"
        )
    lines.extend([
        "",
        "## How to read energized base",
        "",
        "Identified energized base is the sum of the site ledger. It is an estimate of power the model treats as live.",
        "It is the closest series to active power. It is not active power, except at Q4 2025 where the company printed about 170 MW and the identified sum is allowed to differ by the plug for undisclosed sites.",
        "",
    ])
    return "\n".join(lines) + "\n"


def write_outputs(ledger: dict | None = None, out_dir: Path | None = None) -> Path:
    ledger = ledger or load_ledger()
    out_dir = out_dir or OUT_DIR
    company = company_rows(ledger)
    sites = site_rows(ledger)
    util = utilization_rows(ledger)
    _write_csv(out_dir / "quarterly_company.csv", company)
    _write_csv(out_dir / "quarterly_sites.csv", sites)
    _write_csv(out_dir / "utilization.csv", util)
    summary = out_dir / "SUMMARY.md"
    summary.write_text(summary_markdown(ledger, company, util), encoding="utf-8")
    return out_dir
