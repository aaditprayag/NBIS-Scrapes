"""World map of Nebius campus capacity and estimated power being pulled.

Pins are city or campus-town coordinates, not surveyed parcels. "Pulling" is
the ledger's energized-base megawatts, which is an estimate of live IT load.
It is not a utility-meter reading.
"""

from __future__ import annotations

import json
from pathlib import Path

from nbis_scrapes.panel import ROOT, in_history, load_ledger, quarter_index

OUT_HTML = ROOT / "data" / "out" / "map.html"
OUT_GEOJSON = ROOT / "data" / "out" / "sites.geojson"
CROSSWALK_PATH = ROOT / "data" / "region_crosswalk.json"
FORWARD_QUARTER = "Q4 2026"


def _num(value):
    return value if isinstance(value, (int, float)) else None


def _capacity(site: dict, index: int) -> float | None:
    contracted = _num(site["contracted_mw"][index])
    if contracted:
        return contracted
    announced = site.get("announced_mw")
    if isinstance(announced, (int, float)) and announced:
        return announced
    return contracted


def build_features(ledger: dict) -> dict:
    history = ledger["history_end"]
    hist_i = quarter_index(ledger, history)
    fwd_i = quarter_index(ledger, FORWARD_QUARTER)
    features = []
    unmapped = []
    for site in ledger["sites"]:
        props = {
            "id": site["id"],
            "name": site["name"],
            "country": site["country"],
            "ownership": site["ownership"],
            "customer": site["customer"],
            "confirmed": site["confirmed"],
            "geo_precision": site["geo_precision"],
            "basis": site["basis_energized"],
            "note": site["note"],
            "history_quarter": history,
            "forward_quarter": FORWARD_QUARTER,
            "energized_history": _num(site["energized_base"][hist_i]),
            "energized_forward": _num(site["energized_base"][fwd_i]),
            "contracted_history": _num(site["contracted_mw"][hist_i]),
            "contracted_forward": _num(site["contracted_mw"][fwd_i]),
            "capacity_history": _capacity(site, hist_i),
            "capacity_forward": _capacity(site, fwd_i),
            "announced_mw": site.get("announced_mw"),
        }
        if site["lat"] is None or site["lon"] is None:
            unmapped.append(props)
            continue
        features.append({
            "type": "Feature",
            "geometry": {"type": "Point", "coordinates": [site["lon"], site["lat"]]},
            "properties": props,
        })
    return {
        "type": "FeatureCollection",
        "features": features,
        "unmapped": unmapped,
        "history_quarter": history,
        "forward_quarter": FORWARD_QUARTER,
        "history_is_closed": in_history(ledger, history),
    }


def _crosswalk_rows() -> list[dict]:
    if not CROSSWALK_PATH.exists():
        return []
    return json.loads(CROSSWALK_PATH.read_text(encoding="utf-8"))


def render_html(ledger: dict) -> str:
    payload = build_features(ledger)
    crosswalk = _crosswalk_rows()
    data = json.dumps({"map": payload, "regions": crosswalk})
    return HTML.replace("/*__DATA__*/", data)


def write_map(ledger: dict | None = None) -> Path:
    ledger = ledger or load_ledger()
    payload = build_features(ledger)
    OUT_GEOJSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_GEOJSON.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    OUT_HTML.write_text(render_html(ledger), encoding="utf-8")
    return OUT_HTML


HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Nebius campuses — capacity and power pulling</title>
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
<style>
  :root {
    --ink: #1c1915;
    --muted: #5e584f;
    --paper: #f4f0e6;
    --line: #d9d1c3;
    --pull: #0f6e56;
    --idle: #9a6b12;
    --lead: #8d867c;
  }
  * { box-sizing: border-box; }
  html, body { margin: 0; height: 100%; background: var(--paper); color: var(--ink); }
  body { font-family: Georgia, "Iowan Old Style", "Palatino Linotype", serif; }
  .layout { display: grid; grid-template-columns: 380px 1fr; grid-template-rows: 1fr auto; height: 100%; }
  aside { grid-row: 1 / span 2; overflow: auto; border-right: 1px solid var(--line); padding: 22px 20px 32px; }
  #map { width: 100%; height: 100%; background: #e7e1d4; }
  h1 { font-size: 22px; font-weight: 600; line-height: 1.25; margin: 0 0 8px; }
  p.lead, p.note { font-family: "Segoe UI", sans-serif; font-size: 13px; line-height: 1.45; color: var(--muted); margin: 0 0 12px; }
  .toggle { display: flex; gap: 6px; margin: 12px 0; }
  button { font-family: "Segoe UI", sans-serif; font-size: 12px; border: 1px solid var(--line); background: white; color: var(--ink); padding: 6px 10px; cursor: pointer; }
  button.on { background: var(--ink); color: var(--paper); border-color: var(--ink); }
  .legend { display: flex; gap: 14px; font-family: "Segoe UI", sans-serif; font-size: 12px; margin-bottom: 12px; }
  .swatch { display: inline-block; width: 9px; height: 9px; border-radius: 50%; margin-right: 5px; }
  ul.sites { list-style: none; padding: 0; margin: 0; }
  ul.sites li { border-top: 1px solid var(--line); padding: 8px 0; cursor: pointer; }
  ul.sites li strong { display: block; font-size: 14px; }
  ul.sites li span { font-family: "Segoe UI", sans-serif; font-size: 12px; color: var(--muted); }
  h2 { font-size: 14px; margin: 22px 0 8px; }
  table { width: 100%; border-collapse: collapse; font-family: "Segoe UI", sans-serif; font-size: 12px; }
  td { padding: 4px 4px 4px 0; vertical-align: top; border-top: 1px solid var(--line); }
  .footer { grid-column: 2; padding: 8px 14px; font-family: "Segoe UI", sans-serif; font-size: 11px; color: var(--muted); border-top: 1px solid var(--line); }
  @media (max-width: 860px) {
    .layout { grid-template-columns: 1fr; grid-template-rows: 58vh auto auto; }
    aside { grid-row: auto; border-right: 0; border-top: 1px solid var(--line); }
    .footer { grid-column: 1; }
  }
</style>
</head>
<body>
<div class="layout">
  <aside>
    <h1>Nebius campuses</h1>
    <p class="lead">Circle area is contracted megawatts, or the announced figure when the quarter has no contract yet. Green means the model has energized megawatts in the selected quarter. Amber means capacity is on the books and nothing is modeled as pulling power.</p>
    <div class="toggle">
      <button id="btn-hist" class="on" type="button">Q2 2026 pulling</button>
      <button id="btn-fwd" type="button">Q4 2026 path</button>
    </div>
    <div class="legend">
      <span><i class="swatch" style="background:#0f6e56"></i>Pulling</span>
      <span><i class="swatch" style="background:#e2b657"></i>Contracted, not pulling</span>
      <span><i class="swatch" style="background:#c8c2b8"></i>Unconfirmed lead</span>
    </div>
    <p class="note" id="totals"></p>
    <ul class="sites" id="list"></ul>
    <h2>Not pinned</h2>
    <p class="note" id="unmapped"></p>
    <h2>GPUs listed for sale, by region</h2>
    <p class="note">This is the public docs matrix, not a map of buildings. A region id is not a campus.</p>
    <table id="regions"></table>
  </aside>
  <div id="map"></div>
  <div class="footer">Pins are city-level. Energized MW is an estimate of live IT load from the site ledger, not a meter. Connected power and contracted power are different stocks. Open this file in a browser. Tiles load from CARTO.</div>
</div>
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<script>
const PACK = /*__DATA__*/;
const mapData = PACK.map;
let mode = "history";

const map = L.map("map", { worldCopyJump: true }).setView([30, 10], 2);
L.tileLayer("https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png", {
  attribution: "&copy; OpenStreetMap &copy; CARTO",
  subdomains: "abcd",
  maxZoom: 18
}).addTo(map);

function esc(value) {
  return String(value == null ? "" : value).replace(/[&<>]/g, function (ch) {
    return { "&": "&amp;", "<": "&lt;", ">": "&gt;" }[ch];
  });
}
function num(value) {
  return typeof value === "number" ? value : 0;
}
function fmt(value) {
  if (value == null || value === "") return "—";
  const n = Number(value);
  if (Number.isNaN(n)) return "—";
  return (Math.round(n * 10) / 10).toString();
}
function radius(mw) {
  const v = Math.max(mw, 0);
  if (!v) return 5;
  return 6 + 24 * Math.sqrt(v) / Math.sqrt(1200);
}
function pullOf(props) {
  return mode === "history" ? props.energized_history : props.energized_forward;
}
function capOf(props) {
  return mode === "history" ? props.capacity_history : props.capacity_forward;
}
function styleOf(feature) {
  const props = feature.properties;
  const pulling = num(pullOf(props)) > 0;
  const lead = !props.confirmed;
  return {
    radius: radius(Math.max(num(capOf(props)), num(pullOf(props)))),
    color: lead ? "#8d867c" : (pulling ? "#0f6e56" : "#9a6b12"),
    weight: 1,
    fillColor: lead ? "#c8c2b8" : (pulling ? "#1f8a70" : "#e2b657"),
    fillOpacity: 0.8
  };
}
function popup(props) {
  const q = mode === "history" ? props.history_quarter : props.forward_quarter;
  return "<strong>" + esc(props.name) + "</strong><br>"
    + esc(props.country) + " · " + esc(props.ownership) + "<br>"
    + esc(q) + "<br>"
    + "Pulling (energized base): " + fmt(pullOf(props)) + " MW<br>"
    + "Capacity: " + fmt(capOf(props)) + " MW<br>"
    + "Customer: " + esc(props.customer) + "<br>"
    + "Pin: " + esc(props.geo_precision) + (props.confirmed ? "" : " · unconfirmed lead") + "<br>"
    + "<span style='color:#5e584f'>" + esc(props.note) + "</span>";
}

const layer = L.geoJSON(mapData.features, {
  pointToLayer: function (feature, latlng) {
    return L.circleMarker(latlng, styleOf(feature));
  },
  onEachFeature: function (feature, marker) {
    marker.bindPopup(popup(feature.properties));
    feature._marker = marker;
  }
}).addTo(map);

function redraw() {
  layer.eachLayer(function (marker) {
    const feature = marker.feature;
    marker.setStyle(styleOf(feature));
    marker.setPopupContent(popup(feature.properties));
  });
  renderList();
  document.getElementById("btn-hist").className = mode === "history" ? "on" : "";
  document.getElementById("btn-fwd").className = mode === "forward" ? "on" : "";
}
function renderList() {
  const items = mapData.features.map(function (feature) { return feature.properties; });
  items.sort(function (a, b) { return num(pullOf(b)) - num(pullOf(a)) || num(capOf(b)) - num(capOf(a)); });
  let pullSum = 0;
  let capSum = 0;
  const list = document.getElementById("list");
  list.innerHTML = "";
  items.forEach(function (props) {
    pullSum += num(pullOf(props));
    capSum += num(capOf(props));
    const li = document.createElement("li");
    li.innerHTML = "<strong>" + esc(props.name) + "</strong><span>"
      + esc(props.country) + " · pulling " + fmt(pullOf(props))
      + " MW · capacity " + fmt(capOf(props)) + " MW</span>";
    li.addEventListener("click", function () {
      const feature = mapData.features.find(function (item) { return item.properties.id === props.id; });
      if (feature && feature._marker) {
        map.setView(feature._marker.getLatLng(), 5);
        feature._marker.openPopup();
      }
    });
    list.appendChild(li);
  });
  const q = mode === "history" ? mapData.history_quarter : mapData.forward_quarter;
  document.getElementById("totals").textContent =
    q + " pinned sites: " + fmt(pullSum) + " MW modeled as pulling, " + fmt(capSum) + " MW of capacity on the pin. Unlocated sites are excluded from both sums.";
}
document.getElementById("btn-hist").addEventListener("click", function () { mode = "history"; redraw(); });
document.getElementById("btn-fwd").addEventListener("click", function () { mode = "forward"; redraw(); });

const unmapped = mapData.unmapped || [];
document.getElementById("unmapped").textContent = unmapped.length
  ? unmapped.map(function (site) { return site.name; }).join("; ")
  : "Every site in the ledger has a pin.";

const regionTable = document.getElementById("regions");
(PACK.regions || []).forEach(function (row) {
  const tr = document.createElement("tr");
  tr.innerHTML = "<td>" + esc(row.region) + "</td><td>" + esc(row.candidate) + "</td><td>" + esc(row.confidence) + "</td>";
  regionTable.appendChild(tr);
});
redraw();
</script>
</body>
</html>
"""
