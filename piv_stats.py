#!/usr/bin/env python3
# encoding: utf-8
"""
PIV Software Academic Usage Statistics
=======================================
Fetches data from OpenAlex to compare how often PIV software packages
appear in the scientific literature, and in which research fields.

Unified methodology (fair, objective comparison across all software):
  - Full-text boolean search using quoted exact phrases:
      default.search:"SoftwareName" AND "particle image velocimetry"
  - "particle image velocimetry" must appear literally in the text —
    this ensures PIV context without relying on ML concept tags
  - Spelling variants covered with OR:
      ("OpenPIV" OR "open piv") AND "particle image velocimetry"
  - Commercial software: company name must also appear:
      "Dantec" AND "Dynamic Studio" AND "particle image velocimetry"

This ensures:
  (a) Names are found in full text (title + abstract + body where indexed)
  (b) PIV context is guaranteed by the literal phrase in text
  (c) Same metric for every package — enables fair, objective comparison
  (d) Quoted phrases prevent stemming false-positives

Run:  python piv_stats.py
Output: output/piv_report.html   (interactive Plotly report)
        output/piv_data.xlsx      (all data as Excel workbook)
"""

import json
import re
import sys
import time
import urllib.parse
from datetime import date
from pathlib import Path

import requests
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
EMAIL = "thielickeoptolution@gmail.com"   # OpenAlex polite-pool identifier
YEAR_MIN = 2010
YEAR_MAX = date.today().year - 1  # last fully completed year
CACHE_DIR = Path("cache")
OUTPUT_DIR = Path("output")
BASE_URL = "https://api.openalex.org"
REQUEST_DELAY = 0.15          # seconds between API calls (polite crawling)
TOP_N_FIELDS = 12             # fields shown in distributions
TOP_N_JOURNALS = 10
TOP_N_COUNTRIES = 15

# OpenAlex concept ID for "Particle image velocimetry" — used to filter out
# false positives when software names are ambiguous (e.g. GPIV=medical term,
# Flownizer≈flown, Dantec≈company with many product lines).
PIV_CONCEPT_FILTER = "concepts.id:C207857233"

# ---------------------------------------------------------------------------
# Software registry
# ---------------------------------------------------------------------------
# Unified metric for ALL software:
#   default.search:<fulltext_query>,concepts.id:C207857233
#
# fulltext_query : term(s) passed to OpenAlex default.search (searches title,
#                 abstract, and full-text body where indexed).  Quote phrases
#                 to avoid stemming false-positives (e.g. "PaIRS-UniNa").
#
# The PIV concept filter (C207857233) is always applied.  It restricts
# results to papers OpenAlex has ML-tagged as "Particle image velocimetry",
# ensuring multi-purpose software names are counted only in a PIV context.
#
# cite_ids : kept for reference / future supplementary analysis only.
# ---------------------------------------------------------------------------
SOFTWARE = [
    # ---- Open-source ----
    {
        "id": "pivlab",
        "name": "PIVlab",
        "category": "Open Source",
        "developer": "W. Thielicke",
        "url": "https://github.com/Shrediquette/PIVlab",
        "color": "#1565C0",
        "fulltext_query": '("PIVlab" OR "PIV lab") AND "particle image velocimetry"',
        "cite_ids": ["W2138221697", "W3172557238"],
        "active_since": 2010,
        "note": 'Exact phrase "PIVlab" + PIV phrase',
    },
    {
        "id": "matpiv",
        "name": "MatPIV",
        "category": "Open Source",
        "developer": "J.K. Sveen",
        "url": "https://github.com/alexlib/matpiv",
        "color": "#2E7D32",
        "fulltext_query": '"MatPIV" AND "particle image velocimetry"',
        "cite_ids": ["W593841176"],
        "active_since": 2000,
        "note": 'Exact phrase "MatPIV" + PIV phrase',
    },
    {
        "id": "geopiv",
        "name": "GeoPIV / GeoPIV-RG",
        "category": "Free / Academic",
        "developer": "White, Stanier et al.",
        "url": "https://www.geopivrg.com/",
        "color": "#6A1B9A",
        "fulltext_query": '("GeoPIV" OR "GeoPIV-RG" OR "Geo PIV") AND "particle image velocimetry"',
        "cite_ids": ["W2199344813"],
        "active_since": 2003,
        "note": '"GeoPIV" OR "GeoPIV-RG" OR "Geo PIV" (spacing variant; all geotechnical) + PIV phrase',
    },
    {
        "id": "openpiv",
        "name": "OpenPIV",
        "category": "Open Source",
        "developer": "OpenPIV Team",
        "url": "https://www.openpiv.net/",
        "color": "#00838F",
        "fulltext_query": '("OpenPIV" OR "open piv" OR "openpiv-matlab" OR "openpiv matlab") AND "particle image velocimetry"',
        "cite_ids": [],
        "active_since": 2010,
        "note": '"OpenPIV" OR "open piv" OR "openpiv-matlab" OR "openpiv matlab" + PIV phrase',
    },
    {
        "id": "pairs_unina",
        "name": "PaIRS-UniNa",
        "category": "Free / Academic",
        "developer": "Univ. of Naples",
        "url": "https://www.pairs.unina.it/",
        "color": "#E65100",
        "fulltext_query": '"PaIRS-UniNa" AND "particle image velocimetry"',
        "cite_ids": ["W4400499824"],
        "active_since": 2016,
        "note": 'Exact phrase "PaIRS-UniNa" + PIV phrase (unquoted "PaIRS" matches "pairs")',
    },
    {
        "id": "fluidimage",
        "name": "FluidImage",
        "category": "Open Source",
        "developer": "P. Augier et al.",
        "url": "https://github.com/fluiddyn/fluidimage",
        "color": "#558B2F",
        "fulltext_query": '("fluidimage" OR "fluiddyn") AND "particle image velocimetry"',
        "cite_ids": [],
        "active_since": 2016,
        "note": '"fluidimage" OR "fluiddyn" (parent project) + PIV phrase',
    },
    {
        "id": "uvmat",
        "name": "UVMAT / CIVx",
        "category": "Open Source",
        "developer": "J.-M. Foucaut et al.",
        "url": "https://legi.gricad-pages.univ-grenoble-alpes.fr/soft/uvmat-doc/",
        "color": "#4527A0",
        "fulltext_query": '"UVMAT" AND "particle image velocimetry"',
        "cite_ids": [],
        "active_since": 2003,
        "note": 'Exact phrase "UVMAT" + PIV phrase',
    },
    {
        "id": "jpiv",
        "name": "JPIV",
        "category": "Open Source",
        "developer": "P. Vennemann",
        "url": "https://github.com/eguvep/jpiv",
        "color": "#AD1457",
        "fulltext_query": '"JPIV" AND "particle image velocimetry"',
        "cite_ids": [],
        "active_since": 2007,
        "note": 'Exact phrase "JPIV" + PIV phrase',
    },
    {
        "id": "gpiv",
        "name": "GPIV",
        "category": "Open Source",
        "developer": "G.H. de Graaf",
        "url": "https://gpiv.sourceforge.net/",
        "color": "#37474F",
        # Requiring "particle image velocimetry" in text eliminates medical false positives
        "fulltext_query": '"GPIV" AND "particle image velocimetry"',
        "cite_ids": [],
        "active_since": 2005,
        "note": 'Exact phrase "GPIV" + PIV phrase (removes medical-literature false positives)',
    },
    {
        "id": "prana",
        "name": "PRANA",
        "category": "Open Source",
        "developer": "Eckstein & Vlachos",
        "url": "https://github.com/aether-lab/prana",
        "color": "#FF8F00",
        # "particle image velocimetry" phrase eliminates yoga/Sanskrit uses of "prana"
        "fulltext_query": '"PRANA" AND "particle image velocimetry"',
        "cite_ids": [],
        "active_since": 2009,
        "note": 'Exact phrase "PRANA" + PIV phrase (removes non-PIV uses)',
    },
    {
        "id": "pivmat",
        "name": "PIVmat",
        "category": "Open Source",
        "developer": "F. Moisy",
        "url": "http://www.fast.u-psud.fr/pivmat/",
        "color": "#BF360C",
        "fulltext_query": '"PIVmat" AND "particle image velocimetry"',
        "cite_ids": [],
        "active_since": 2005,
        "note": 'Exact phrase "PIVmat" + PIV phrase',
    },
    {
        "id": "digiflow",
        "name": "DigiFlow",
        "category": "Commercial",
        "developer": "Dalziel Research Partners",
        "url": "https://www.damtp.cam.ac.uk/user/fdl/digiflow/digiflow.htm",
        "color": "#1B5E20",
        "fulltext_query": '"DigiFlow" AND "particle image velocimetry"',
        "cite_ids": [],
        "active_since": 2000,
        "note": 'Exact phrase "DigiFlow" + PIV phrase',
    },
    {
        "id": "civx",
        "name": "CIVx",
        "category": "Open Source",
        "developer": "Gostiaux, Salort (ENS Lyon)",
        "url": "https://sourceforge.net/projects/civx/",
        "color": "#880E4F",
        "fulltext_query": '"CIVx" AND "particle image velocimetry"',
        "cite_ids": [],
        "active_since": 2005,
        "note": 'Exact phrase "CIVx" + PIV phrase',
    },
    # ---- Commercial (company name required) ----
    {
        "id": "lavision_davis",
        "name": "LaVision DaVis",
        "category": "Commercial",
        "developer": "LaVision GmbH",
        "url": "https://www.lavision.de/en/products/davis-software/",
        "color": "#C62828",
        # Company "LaVision" AND software "DaVis" as separate terms (not adjacent phrase).
        # Reason: papers routinely write "DaVis 8.3 (LaVision, Göttingen)" where the
        # words are not adjacent.  Adjacent-phrase search misses ~3× papers.
        # This aligns with the methodology of the PIVlab popularity blog post (2022).
        "fulltext_query": '"LaVision" AND "DaVis" AND "particle image velocimetry"',
        "cite_ids": [],
        "active_since": 1995,
        "note": '"LaVision" (company) AND "DaVis" (software, non-adjacent) AND PIV phrase',
    },
    {
        "id": "dantec_ds",
        "name": "Dantec Dynamic Studio",
        "category": "Commercial",
        "developer": "Dantec Dynamics",
        "url": "https://www.dantecdynamics.com/dynamicstudio/",
        "color": "#FF6F00",
        # Company "Dantec" AND software "Dynamic Studio" AND PIV phrase
        "fulltext_query": '"Dantec" AND ("Dynamic Studio" OR "dynamicstudio") AND "particle image velocimetry"',
        "cite_ids": [],
        "active_since": 1998,
        "note": '"Dantec" AND ("Dynamic Studio" OR "dynamicstudio") AND PIV phrase',
    },
    {
        "id": "tsi_insight",
        "name": "TSI Insight",
        "category": "Commercial",
        "developer": "TSI Inc.",
        "url": "https://tsi.com/",
        "color": "#00695C",
        # "TSI Insight" phrase includes company + software; "Insight 4G" is a newer version name.
        # "TSI" AND "Insight" separately is risky: "insight" is a very common English word.
        "fulltext_query": '("TSI Insight" OR "Insight 4G" OR "Insight4G") AND "particle image velocimetry"',
        "cite_ids": [],
        "active_since": 1999,
        "note": 'Phrase "TSI Insight" OR "Insight 4G"/"Insight4G" (version name) + PIV phrase',
    },
    {
        "id": "pivview",
        "name": "PIVview",
        "category": "Commercial",
        "developer": "PIVTEC GmbH",
        "url": "https://www.pivtec.com/pivview.html",
        "color": "#4A148C",
        # Software name "PIVview" OR company "PIVTEC"
        "fulltext_query": '("PIVview" OR "PIVTEC") AND "particle image velocimetry"',
        "cite_ids": [],
        "active_since": 2000,
        "note": '"PIVview" (software) OR "PIVTEC" (company) + PIV phrase',
    },
    {
        "id": "idt_provision",
        "name": "IDT ProVision",
        "category": "Commercial",
        "developer": "IDT",
        "url": "https://idtvision.com/products/software/provision/",
        "color": "#01579B",
        # Company "IDT" AND software "ProVision" AND PIV phrase
        "fulltext_query": '"IDT" AND "ProVision" AND "particle image velocimetry"',
        "cite_ids": [],
        "active_since": 2005,
        "note": '"IDT" (company) AND "ProVision" (software) AND PIV phrase',
    },
    {
        "id": "microvec",
        "name": "MicroVec",
        "category": "Commercial",
        "developer": "MicroVec Inc.",
        "url": "https://piv.com.sg/piv-products/microvec-piv-software/",
        "color": "#1A237E",
        "fulltext_query": '"MicroVec" AND "particle image velocimetry"',
        "cite_ids": [],
        "active_since": 2005,
        "note": 'Exact phrase "MicroVec" + PIV phrase',
    },
    {
        "id": "flownizer",
        "name": "Flownizer",
        "category": "Commercial",
        "developer": "Ditect Co. Ltd.",
        "url": "https://www.ditect.co.jp/en/software/flownizer.html",
        "color": "#E91E63",
        # Quoted exact phrase prevents stemming ("Flownizer" → "flown" without quotes).
        # "Ditect" is excluded: Ditect makes cameras used in many PIV labs regardless
        # of which PIV software they use — including it inflates the count ~10x.
        "fulltext_query": '"Flownizer" AND "particle image velocimetry"',
        "cite_ids": [],
        "active_since": 2005,
        "note": 'Exact phrase "Flownizer" + PIV phrase; "Ditect" excluded (camera brand, not software)',
    },
    {
        "id": "ftrpiv",
        "name": "FTR-PIV",
        "category": "Commercial",
        "developer": "FlowTech Research",
        "url": "https://www.ft-r.jp/en/piv-soft/",
        "color": "#795548",
        "fulltext_query": '("FTR-PIV" OR "FTRPIV") AND "particle image velocimetry"',
        "cite_ids": [],
        "active_since": 2005,
        "note": '"FTR-PIV" OR "FTRPIV" (one-word form; 12/13 co-occur with FlowTech) + PIV phrase',
    },
    {
        "id": "vidpiv",
        "name": "VidPIV",
        "category": "Commercial",
        "developer": "Oxford Lasers",
        "url": "https://www.oxfordlasers.com/imaging/particle-image-velocimetry-piv/particle-image-velocimetry-software-vidpiv/",
        "color": "#9E9D24",
        "fulltext_query": '"VidPIV" AND "particle image velocimetry"',
        "cite_ids": [],
        "active_since": 1996,
        "note": 'Exact phrase "VidPIV" + PIV phrase',
    },
    {
        "id": "ila",
        "name": "ILA PIVtec",
        "category": "Commercial",
        "developer": "ILA GmbH / PIVtec",
        "url": "https://www.ila5150.de/en/components/software",
        "color": "#006064",
        # "PIVtec" is the specific product brand; "ILA" alone is too ambiguous
        "fulltext_query": '"PIVtec" AND "particle image velocimetry"',
        "cite_ids": [],
        "active_since": 2000,
        "note": 'Exact phrase "PIVtec" (product brand) + PIV phrase',
    },
    {
        "id": "ysc",
        "name": "YSC Flow Analyzer",
        "category": "Commercial",
        "developer": "YSC Corp.",
        "url": "http://ysctech.com/Image-Analysis-Software-Flow-Analyzer.html",
        "color": "#4E342E",
        "fulltext_query": '"YSC" AND "particle image velocimetry"',
        "cite_ids": [],
        "active_since": 2005,
        "note": 'Exact phrase "YSC" + PIV phrase',
    },
]

# ---------------------------------------------------------------------------
# OpenAlex API helper
# ---------------------------------------------------------------------------

_session = requests.Session()
_session.headers.update({"User-Agent": f"PIVStats/1.0 (mailto:{EMAIL})"})


def _openalex_get(endpoint: str, params: dict) -> dict:
    """Raw GET with polite delay."""
    p = {"mailto": EMAIL, **params}
    url = f"{BASE_URL}{endpoint}"
    resp = _session.get(url, params=p, timeout=60)
    resp.raise_for_status()
    time.sleep(REQUEST_DELAY)
    return resp.json()


def _cache_path(key: str) -> Path:
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    return CACHE_DIR / f"{key}.json"


def cached_get(cache_key: str, endpoint: str, params: dict, force: bool = False) -> dict:
    """GET with filesystem cache (one JSON file per query)."""
    cp = _cache_path(cache_key)
    if cp.exists() and not force:
        return json.loads(cp.read_text(encoding="utf-8"))
    data = _openalex_get(endpoint, params)
    cp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return data


def group_by_query(filter_str: str, group_by: str, cache_key: str,
                   extra_filter: str = "", per_page: int = 200) -> list[dict]:
    """
    Run an OpenAlex works query with group_by and return the group list.
    Also returns the total count as group_list[0]["_total"] = meta.count.
    """
    year_filter = f"publication_year:{YEAR_MIN}-{YEAR_MAX}"
    full_filter = ",".join(f for f in [filter_str, year_filter, extra_filter] if f)
    params = {
        "filter": full_filter,
        "group_by": group_by,
        "per_page": per_page,
    }
    data = cached_get(cache_key, "/works", params)
    groups = data.get("group_by", [])
    total = data.get("meta", {}).get("count", 0)
    # tag total onto first element for convenience
    if groups:
        groups[0]["_total"] = total
    else:
        groups = [{"_total": total}]
    return groups, total


def build_filter(sw: dict) -> str:
    """Build the OpenAlex filter string for a software entry.

    Unified strategy: full-text boolean search combining the software name
    (exact quoted phrase) with "particle image velocimetry" (also quoted).
    The PIV phrase is baked into every fulltext_query — no separate concept
    filter needed.
    """
    q = sw["fulltext_query"]
    return f"default.search:{q}"


# ---------------------------------------------------------------------------
# Data collection
# ---------------------------------------------------------------------------

def collect_yearly(sw: dict) -> dict[str, int]:
    """Return {year_str: paper_count} for the given software."""
    filt = build_filter(sw)
    cache_key = f"yearly_{sw['id']}"
    groups, _ = group_by_query(filt, "publication_year", cache_key)
    result = {}
    for g in groups:
        k = g.get("key", "")
        if k and k.isdigit():
            result[k] = g.get("count", 0)
    return result


def collect_field_distribution(sw: dict) -> list[dict]:
    """Return top research fields for citing papers."""
    filt = build_filter(sw)
    cache_key = f"fields_{sw['id']}"
    # Use primary_topic.field.id (valid key); display name is in key_display_name
    groups, _ = group_by_query(filt, "primary_topic.field.id", cache_key, per_page=50)
    return [{"field": g.get("key_display_name", g["key"]), "count": g["count"]}
            for g in groups if g.get("key") and g.get("key_display_name")]


def collect_country_distribution(sw: dict) -> list[dict]:
    """Return top countries of citing paper authors."""
    filt = build_filter(sw)
    cache_key = f"countries_{sw['id']}"
    groups, _ = group_by_query(filt, "authorships.institutions.country_code",
                               cache_key, per_page=100)
    result = []
    for g in groups:
        key = g.get("key", "")
        if not key or key == "unknown":
            continue
        # key is a URI like https://openalex.org/countries/US — extract code
        iso2 = key.split("/")[-1] if "/" in key else key
        result.append({
            "country_code": iso2,
            "country": g.get("key_display_name", iso2),
            "count": g["count"],
        })
    return result


def collect_journal_distribution(sw: dict) -> list[dict]:
    """Return top journals/venues for citing papers."""
    filt = build_filter(sw)
    cache_key = f"journals_{sw['id']}"
    # Use source.id (valid); display name is in key_display_name
    groups, _ = group_by_query(filt, "primary_location.source.id",
                               cache_key, per_page=50)
    return [{"journal": g.get("key_display_name", g["key"]),
             "source_id": g.get("key", ""),
             "count": g["count"]}
            for g in groups if g.get("key") and g.get("key_display_name")]


def fetch_source_impacts(all_data: dict) -> dict[str, float]:
    """Fetch 2yr_mean_citedness for every unique journal seen across all software.
    Returns {openalex_source_uri: float}. Results are persisted in a single cache file."""
    cache_path = _cache_path("source_impacts")
    cache: dict[str, float] = (
        json.loads(cache_path.read_text(encoding="utf-8")) if cache_path.exists() else {}
    )

    all_ids = {
        j["source_id"]
        for dat in all_data.values()
        for j in dat.get("journals", [])
        if j.get("source_id")
    }
    missing = [sid for sid in all_ids if sid not in cache]

    if missing:
        print(f"\n    fetching {len(missing)} journal impact scores", end="", flush=True)
        for uri in missing:
            short_id = uri.split("/")[-1]
            try:
                data = _openalex_get(f"/sources/{short_id}", {})
                val = (data.get("summary_stats") or {}).get("2yr_mean_citedness") or 0.0
                cache[uri] = float(val)
            except Exception:
                cache[uri] = 0.0
        cache_path.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")
        print(" done")

    return cache


def _safe_collect(fn, sw: dict, default):
    """Run a collection function; return default on error."""
    try:
        return fn(sw)
    except Exception as exc:
        print(f"\n    (warning: {fn.__name__} failed: {exc})", end="")
        return default


def collect_piv_baseline() -> dict[str, int]:
    """Total PIV papers per year — papers containing the phrase
    'particle image velocimetry' in full text.  Used as denominator for
    market-share normalization (consistent with the per-software queries)."""
    cache_key = "piv_baseline_total"
    piv_filter = 'default.search:"particle image velocimetry"'
    groups, _ = group_by_query(piv_filter, "publication_year", cache_key)
    return {g["key"]: g["count"] for g in groups if g.get("key", "").isdigit()}


def collect_all(force: bool = False) -> dict:
    """Collect all data for every software entry. Returns nested dict."""
    all_data = {}
    total_sw = len(SOFTWARE)
    for i, sw in enumerate(SOFTWARE, 1):
        name = sw["name"]
        print(f"  [{i:2d}/{total_sw}] {name} ...", end="", flush=True)
        yearly = _safe_collect(collect_yearly, sw, {})
        total = sum(yearly.values())
        fields = _safe_collect(collect_field_distribution, sw, [])
        countries = _safe_collect(collect_country_distribution, sw, [])
        journals = _safe_collect(collect_journal_distribution, sw, [])
        all_data[sw["id"]] = {
            "meta": sw,
            "yearly": yearly,
            "total": total,
            "fields": fields,
            "countries": countries,
            "journals": journals,
        }
        print(f" {total:5d} papers")
    return all_data


# ---------------------------------------------------------------------------
# Data frames
# ---------------------------------------------------------------------------

def build_yearly_df(all_data: dict) -> pd.DataFrame:
    """Long-form DataFrame: software, year, count."""
    rows = []
    years = list(range(YEAR_MIN, YEAR_MAX + 1))
    for sw_id, dat in all_data.items():
        yearly = dat["yearly"]
        sw = dat["meta"]
        for y in years:
            rows.append({
                "software": sw["name"],
                "id": sw_id,
                "year": int(y),
                "count": yearly.get(str(y), 0),
                "category": sw["category"],
            })
    return pd.DataFrame(rows)


def build_totals_df(all_data: dict) -> pd.DataFrame:
    rows = []
    for sw_id, dat in all_data.items():
        sw = dat["meta"]
        rows.append({
            "software": sw["name"],
            "id": sw_id,
            "total": dat["total"],
            "category": sw["category"],
            "developer": sw["developer"],
            "note": sw.get("note", ""),
        })
    df = pd.DataFrame(rows).sort_values("total", ascending=False)
    return df


def build_fields_df(all_data: dict) -> pd.DataFrame:
    rows = []
    for sw_id, dat in all_data.items():
        sw = dat["meta"]
        for f in dat["fields"][:TOP_N_FIELDS]:
            rows.append({
                "software": sw["name"],
                "field": f["field"],
                "count": f["count"],
                "total": dat["total"],
            })
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df["pct"] = df["count"] / df["total"].replace(0, 1) * 100
    return df


def build_countries_df(all_data: dict) -> pd.DataFrame:
    rows = []
    for sw_id, dat in all_data.items():
        sw = dat["meta"]
        for c in dat["countries"]:
            rows.append({
                "software": sw["name"],
                "country": c["country"],
                "country_code": c["country_code"],
                "count": c["count"],
            })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Visualizations
# ---------------------------------------------------------------------------

PLOTLY_TEMPLATE = "plotly_white"
FONT_FAMILY = "Arial, sans-serif"


def fig_yearly_lines(df_yearly: pd.DataFrame, df_totals: pd.DataFrame) -> go.Figure:
    """Interactive line chart: papers per year for each software."""
    # Sort legend by total papers
    order = df_totals.sort_values("total", ascending=False)["software"].tolist()

    # Color map
    color_map = {sw["name"]: sw["color"] for sw in SOFTWARE}

    fig = px.line(
        df_yearly,
        x="year", y="count",
        color="software",
        color_discrete_map=color_map,
        category_orders={"software": order},
        labels={"count": "Papers per year", "year": "Year", "software": "Software"},
        template=PLOTLY_TEMPLATE,
        title="PIV Software — Papers per Year (OpenAlex)",
    )
    fig.update_traces(line_width=2, mode="lines+markers", marker_size=4)
    fig.update_layout(
        font_family=FONT_FAMILY,
        legend_title_text="Software",
        hovermode="x unified",
        height=550,
        legend=dict(orientation="v", x=1.01, y=1),
        margin=dict(r=220),
    )
    fig.add_annotation(
        text="Unified metric: full-text search (title + abstract + body) + PIV concept filter (OpenAlex C207857233).",
        xref="paper", yref="paper", x=0, y=-0.14,
        showarrow=False, font=dict(size=10), align="left",
    )
    return fig


def fig_yearly_log(df_yearly: pd.DataFrame, df_totals: pd.DataFrame) -> go.Figure:
    """Same as above but log-scale Y to show small software."""
    color_map = {sw["name"]: sw["color"] for sw in SOFTWARE}
    order = df_totals.sort_values("total", ascending=False)["software"].tolist()
    df_plot = df_yearly[df_yearly["count"] > 0].copy()
    fig = px.line(
        df_plot,
        x="year", y="count",
        color="software",
        color_discrete_map=color_map,
        category_orders={"software": order},
        log_y=True,
        labels={"count": "Papers per year (log)", "year": "Year", "software": "Software"},
        template=PLOTLY_TEMPLATE,
        title="PIV Software — Papers per Year (log scale)",
    )
    fig.update_traces(line_width=2, mode="lines+markers", marker_size=4)
    fig.update_layout(
        font_family=FONT_FAMILY,
        legend_title_text="Software",
        hovermode="x unified",
        height=550,
        legend=dict(orientation="v", x=1.01, y=1),
        margin=dict(r=220),
    )
    return fig


def fig_total_bars(df_totals: pd.DataFrame) -> go.Figure:
    """Horizontal bar chart: total papers per software."""
    df = df_totals[df_totals["total"] > 0].sort_values("total")
    color_map = {sw["name"]: sw["color"] for sw in SOFTWARE}
    colors = [color_map.get(n, "#888") for n in df["software"]]
    fig = go.Figure(go.Bar(
        x=df["total"],
        y=df["software"],
        orientation="h",
        marker_color=colors,
        text=df["total"],
        textposition="outside",
        hovertemplate="<b>%{y}</b><br>Total papers: %{x}<extra></extra>",
    ))
    fig.update_layout(
        title=f"PIV Software — Total Papers {YEAR_MIN}–{YEAR_MAX} (OpenAlex)",
        xaxis_title="Number of papers",
        yaxis_title="",
        template=PLOTLY_TEMPLATE,
        font_family=FONT_FAMILY,
        height=max(400, 28 * len(df)),
        margin=dict(r=80, l=200),
        showlegend=False,
    )
    return fig


def fig_field_heatmap(df_fields: pd.DataFrame) -> go.Figure:
    """Heatmap: software × research field (percentage of papers)."""
    if df_fields.empty:
        return go.Figure().update_layout(title="No field data available")

    # Keep top N most common fields overall
    top_fields = (
        df_fields.groupby("field")["count"].sum()
        .nlargest(TOP_N_FIELDS).index.tolist()
    )
    df_h = df_fields[df_fields["field"].isin(top_fields)].copy()
    pivot = df_h.pivot_table(index="software", columns="field", values="pct", fill_value=0)
    # Sort software by total papers
    order = [sw["name"] for sw in SOFTWARE if sw["name"] in pivot.index]
    pivot = pivot.reindex([n for n in order if n in pivot.index])

    fig = go.Figure(go.Heatmap(
        z=pivot.values,
        x=pivot.columns.tolist(),
        y=pivot.index.tolist(),
        colorscale="Blues",
        text=[[f"{v:.1f}%" for v in row] for row in pivot.values],
        texttemplate="%{text}",
        textfont_size=9,
        hovertemplate="<b>%{y}</b><br>Field: %{x}<br>Share: %{z:.1f}%<extra></extra>",
        colorbar_title="% of papers",
    ))
    fig.update_layout(
        title="Research Field Distribution per Software (% of papers)",
        template=PLOTLY_TEMPLATE,
        font_family=FONT_FAMILY,
        height=max(400, 30 * len(pivot)),
        xaxis_tickangle=-35,
        margin=dict(b=160, l=220),
    )
    return fig


def fig_top_fields_treemap(df_fields: pd.DataFrame) -> go.Figure:
    """Treemap of field distribution across all software (total paper counts)."""
    if df_fields.empty:
        return go.Figure().update_layout(title="No field data available")
    df_agg = df_fields.groupby("field")["count"].sum().reset_index()
    df_agg = df_agg[df_agg["count"] > 0].sort_values("count", ascending=False)
    fig = px.treemap(
        df_agg, path=["field"], values="count",
        title="Overall Research Field Distribution (all PIV software combined)",
        template=PLOTLY_TEMPLATE,
        color="count",
        color_continuous_scale="Blues",
    )
    fig.update_layout(font_family=FONT_FAMILY, height=450)
    return fig


def fig_country_choropleth(df_countries: pd.DataFrame) -> go.Figure:
    """World map: total papers per country across all software."""
    if df_countries.empty:
        return go.Figure().update_layout(title="No country data available")
    df_agg = df_countries.groupby(["country", "country_code"])["count"].sum().reset_index()
    df_agg = df_agg.sort_values("count", ascending=False)
    fig = px.choropleth(
        df_agg,
        locations="country",
        locationmode="country names",
        color="count",
        hover_name="country",
        color_continuous_scale="Blues",
        title=f"Geographic Distribution of PIV Research ({YEAR_MIN}–{YEAR_MAX})",
        template=PLOTLY_TEMPLATE,
    )
    fig.update_layout(font_family=FONT_FAMILY, height=450,
                      geo=dict(showframe=False, showcoastlines=True))
    return fig


def fig_country_per_software(df_countries: pd.DataFrame, df_totals: pd.DataFrame,
                              top_n: int = 8) -> go.Figure:
    """Grouped bar: top countries for each software (normalized %)."""
    if df_countries.empty:
        return go.Figure().update_layout(title="No country data")

    # Total per software
    sw_totals = df_totals.set_index("software")["total"]

    # Keep top N countries overall
    top_c = df_countries.groupby("country")["count"].sum().nlargest(top_n).index.tolist()
    df_p = df_countries[df_countries["country"].isin(top_c)].copy()
    df_p["pct"] = df_p.apply(
        lambda r: r["count"] / sw_totals.get(r["software"], 1) * 100, axis=1
    )

    # Sort software by total
    sw_order = df_totals.sort_values("total", ascending=False)["software"].tolist()
    fig = px.bar(
        df_p, x="software", y="pct", color="country",
        category_orders={"software": sw_order},
        labels={"pct": "% of papers", "software": "Software", "country": "Country"},
        title=f"Top {top_n} Countries per Software (% of papers)",
        template=PLOTLY_TEMPLATE,
        barmode="group",
        height=500,
    )
    fig.update_layout(font_family=FONT_FAMILY, xaxis_tickangle=-30, margin=dict(b=120))
    return fig


def fig_category_comparison(df_yearly: pd.DataFrame) -> go.Figure:
    """Area chart: Open Source vs Commercial over time (combined totals)."""
    df_agg = (df_yearly.groupby(["year", "category"])["count"]
              .sum().reset_index())
    fig = px.area(
        df_agg, x="year", y="count", color="category",
        labels={"count": "Papers per year", "year": "Year"},
        title="Open Source vs. Commercial PIV Software — Combined Paper Counts",
        template=PLOTLY_TEMPLATE,
        color_discrete_map={
            "Open Source": "#1565C0",
            "Commercial": "#C62828",
            "Free / Academic": "#558B2F",
        },
    )
    fig.update_layout(font_family=FONT_FAMILY, height=400, hovermode="x unified")
    return fig


def fig_growth_heatmap(df_yearly: pd.DataFrame) -> go.Figure:
    """Year-over-year growth heatmap."""
    pivot = df_yearly.pivot_table(index="software", columns="year",
                                  values="count", fill_value=0)
    # YoY growth rate (%) - skip first year
    growth = pivot.pct_change(axis=1) * 100
    growth = growth.iloc[:, 1:]  # drop first year (NaN)
    # Clip extreme values for readability
    growth = growth.clip(-100, 200)
    growth = growth.fillna(0)

    # Sort by total
    sw_totals = pivot.sum(axis=1).sort_values(ascending=False)
    growth = growth.reindex(sw_totals.index)

    fig = go.Figure(go.Heatmap(
        z=growth.values,
        x=[str(c) for c in growth.columns],
        y=growth.index.tolist(),
        colorscale=[
            [0.0, "#C62828"], [0.4, "#FFEB3B"], [0.5, "#F5F5F5"],
            [0.6, "#90CAF9"], [1.0, "#1565C0"],
        ],
        zmid=0,
        zmin=-100, zmax=200,
        hovertemplate="<b>%{y}</b><br>Year: %{x}<br>YoY growth: %{z:.0f}%<extra></extra>",
        colorbar_title="YoY growth %",
    ))
    fig.update_layout(
        title="Year-over-Year Growth Rate per Software (%)",
        template=PLOTLY_TEMPLATE,
        font_family=FONT_FAMILY,
        height=max(400, 28 * len(growth)),
        xaxis_tickangle=-45,
        margin=dict(l=220),
    )
    return fig


def _add_partial_year_shape(fig: go.Figure, current_year: int) -> go.Figure:
    """Add a vertical dashed line + annotation marking the current year as partial."""
    fig.add_vline(
        x=current_year, line_dash="dot", line_color="#999", line_width=1.5,
    )
    fig.add_annotation(
        x=current_year, y=1, xref="x", yref="paper",
        text=f"{current_year} (partial)", showarrow=False,
        font=dict(size=9, color="#999"), xanchor="left", yanchor="top",
    )
    return fig


def fig_market_share(df_yearly: pd.DataFrame, piv_baseline: dict) -> go.Figure:
    """Line chart: each software's papers as % of total PIV papers that year."""
    if not piv_baseline:
        return go.Figure().update_layout(title="No baseline data")

    color_map = {sw["name"]: sw["color"] for sw in SOFTWARE}
    rows = []
    for _, row in df_yearly.iterrows():
        y = str(row["year"])
        base = piv_baseline.get(y, 0)
        if base > 0 and row["count"] > 0:
            rows.append({
                "software": row["software"],
                "year": row["year"],
                "share_pct": row["count"] / base * 100,
                "category": row["category"],
            })
    if not rows:
        return go.Figure().update_layout(title="No data")

    df_share = pd.DataFrame(rows)
    # Only include top 10 software by share (to keep chart readable)
    top_sw = (df_share.groupby("software")["share_pct"].max()
              .nlargest(10).index.tolist())
    df_share = df_share[df_share["software"].isin(top_sw)]

    fig = px.line(
        df_share, x="year", y="share_pct", color="software",
        color_discrete_map=color_map,
        labels={"share_pct": "% of total PIV papers", "year": "Year"},
        title="Software Market Share (papers mentioning software name as % of total PIV papers/year)",
        template=PLOTLY_TEMPLATE,
    )
    fig.update_traces(line_width=2, mode="lines+markers", marker_size=5)
    fig.update_layout(
        font_family=FONT_FAMILY,
        hovermode="x unified",
        height=500,
        legend=dict(orientation="v", x=1.01, y=1),
        margin=dict(r=220),
        yaxis_ticksuffix="%",
    )
    fig.add_annotation(
        text="Denominator: total papers tagged with 'Particle image velocimetry' concept in OpenAlex.",
        xref="paper", yref="paper", x=0, y=-0.12,
        showarrow=False, font=dict(size=10), align="left",
    )
    _add_partial_year_shape(fig, YEAR_MAX)
    return fig


def build_impact_df(all_data: dict, source_impacts: dict) -> pd.DataFrame:
    """Weighted mean 2yr_mean_citedness per software (weighted by paper count per journal)."""
    rows = []
    for sw_id, dat in all_data.items():
        sw = dat["meta"]
        journals = dat.get("journals", [])
        if not journals:
            continue
        total_papers = 0
        weighted_sum = 0.0
        for j in journals:
            impact = source_impacts.get(j.get("source_id", ""), 0.0) or 0.0
            n = j["count"]
            weighted_sum += impact * n
            total_papers += n
        if total_papers > 0:
            rows.append({
                "software": sw["name"],
                "category": sw["category"],
                "weighted_mean_impact": round(weighted_sum / total_papers, 2),
                "total": dat["total"],
            })
    return pd.DataFrame(rows).sort_values("weighted_mean_impact", ascending=False)


def fig_journal_impact(df_impact: pd.DataFrame) -> go.Figure:
    """Horizontal bar chart: weighted mean journal 2yr_mean_citedness per software."""
    if df_impact.empty:
        return go.Figure().update_layout(title="No journal impact data available")

    df = df_impact.sort_values("weighted_mean_impact")
    cat_colors = {"Open Source": "#1565C0", "Commercial": "#C62828", "Free / Academic": "#558B2F"}
    colors = [cat_colors.get(c, "#888") for c in df["category"]]

    fig = go.Figure(go.Bar(
        x=df["weighted_mean_impact"],
        y=df["software"],
        orientation="h",
        marker_color=colors,
        text=[f"{v:.1f}" for v in df["weighted_mean_impact"]],
        textposition="outside",
        hovertemplate="<b>%{y}</b><br>Weighted mean impact: %{x:.2f}<extra></extra>",
    ))
    fig.update_layout(
        title="Journal Quality — Weighted Mean 2-Year Citedness of Publishing Journals",
        xaxis_title="Weighted mean 2yr mean citedness (OpenAlex proxy for impact factor)",
        yaxis_title="",
        template=PLOTLY_TEMPLATE,
        font_family=FONT_FAMILY,
        height=max(400, 28 * len(df)),
        margin=dict(r=80, l=200),
        showlegend=False,
    )
    fig.add_annotation(
        text=(
            "OpenAlex '2yr_mean_citedness' per journal, weighted by paper count per software. "
            "Higher = published in journals that receive more citations on average. "
            "Blue = Open Source, Red = Commercial."
        ),
        xref="paper", yref="paper", x=0, y=-0.12,
        showarrow=False, font=dict(size=10), align="left",
    )
    return fig


# ---------------------------------------------------------------------------
# Summary table
# ---------------------------------------------------------------------------

def build_summary_html(df_totals: pd.DataFrame, all_data: dict) -> str:
    """Build an HTML summary table."""
    rows = []
    for _, row in df_totals.iterrows():
        sw_id = row["id"]
        dat = all_data.get(sw_id, {})
        yearly = dat.get("yearly", {})
        top_field = ""
        if dat.get("fields"):
            top_field = dat["fields"][0]["field"] if dat["fields"] else ""
        top_country = ""
        if dat.get("countries"):
            top_country = dat["countries"][0]["country"] if dat["countries"] else ""

        # Peak year
        peak_year = max(yearly, key=yearly.get) if yearly else "N/A"
        peak_count = yearly.get(peak_year, 0) if yearly else 0

        # Recent (last 3 years)
        recent_years = [str(y) for y in range(YEAR_MAX - 2, YEAR_MAX + 1)]
        recent = sum(yearly.get(y, 0) for y in recent_years)

        rows.append({
            "Software": row["software"],
            "Category": row["category"],
            "Developer": row["developer"],
            f"Total ({YEAR_MIN}-{YEAR_MAX})": row["total"],
            f"Recent ({YEAR_MAX-2}-{YEAR_MAX})": recent,
            "Peak year": peak_year,
            "Peak count": peak_count,
            "Top field": top_field,
            "Top country": top_country,
        })

    url_map = {sw["name"]: sw.get("url", "") for sw in SOFTWARE}
    df = pd.DataFrame(rows)
    df["Software"] = df["Software"].apply(
        lambda n: f'<a href="{url_map[n]}" target="_blank">{n}</a>'
        if url_map.get(n) else n
    )
    html = df.to_html(
        index=False,
        border=0,
        classes="summary-table",
        escape=False,
    )
    # Inject onclick + sort icon into every <th> to enable client-side sorting
    html = re.sub(
        r"<th>(.*?)</th>",
        lambda m: (
            f'<th onclick="sortTable(this)" '
            f'style="cursor:pointer;user-select:none">'
            f'{m.group(1)}'
            f'<span class="sort-icon" style="font-size:10px;opacity:0.5;margin-left:4px">⇅</span>'
            f"</th>"
        ),
        html,
    )
    return html


# ---------------------------------------------------------------------------
# HTML Report
# ---------------------------------------------------------------------------

def fig_to_html(fig: go.Figure) -> str:
    """Convert figure to embedded HTML div (no full-page wrapper)."""
    return fig.to_html(full_html=False, include_plotlyjs=False, div_id=None,
                       config={"displayModeBar": True, "responsive": True})


def generate_html_report(all_data: dict, df_yearly: pd.DataFrame,
                         df_totals: pd.DataFrame, df_fields: pd.DataFrame,
                         df_countries: pd.DataFrame,
                         piv_baseline: dict | None = None,
                         df_impact: pd.DataFrame | None = None) -> str:
    """Build complete HTML report string."""
    today = date.today().strftime("%Y-%m-%d")
    piv_baseline = piv_baseline or {}

    # ---- Generate figures ----
    print("  Generating figures...")
    f_linear = fig_yearly_lines(df_yearly, df_totals)
    _add_partial_year_shape(f_linear, YEAR_MAX)
    f_log = fig_yearly_log(df_yearly, df_totals)
    _add_partial_year_shape(f_log, YEAR_MAX)
    figs = {
        "yearly_linear": f_linear,
        "yearly_log": f_log,
        "total_bars": fig_total_bars(df_totals),
        "market_share": fig_market_share(df_yearly, piv_baseline),
        "field_heatmap": fig_field_heatmap(df_fields),
        "field_treemap": fig_top_fields_treemap(df_fields),
        "country_map": fig_country_choropleth(df_countries),
        "country_bars": fig_country_per_software(df_countries, df_totals),
        "category": fig_category_comparison(df_yearly),
        "journal_impact": fig_journal_impact(df_impact if df_impact is not None else pd.DataFrame()),
    }

    summary_html = build_summary_html(df_totals, all_data)

    # ---- Search term table ----
    search_rows = ""
    for sw in SOFTWARE:
        q = sw["fulltext_query"]
        search_rows += (
            f"<tr><td style='padding:2px 8px'>{sw['name']}</td>"
            f"<td style='font-family:monospace;font-size:11px'>{q}</td>"
            f"<td style='font-size:11px;color:#555'>{sw.get('note','')}</td></tr>"
        )

    # ---- HTML template ----
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>PIV Software Usage Statistics</title>
<script src="https://cdn.plot.ly/plotly-2.35.0.min.js"></script>
<style>
  body {{ font-family: Arial, sans-serif; margin: 0; background: #f8f9fa; color: #212529; }}
  .header {{ background: #1565C0; color: white; padding: 24px 36px; }}
  .header h1 {{ margin: 0 0 6px; font-size: 28px; }}
  .header p {{ margin: 0; opacity: 0.85; font-size: 14px; }}
  .content {{ max-width: 1400px; margin: 0 auto; padding: 24px 24px 60px; }}
  .section {{ background: white; border-radius: 8px; box-shadow: 0 1px 4px rgba(0,0,0,0.1);
              margin-bottom: 24px; padding: 20px; }}
  h2 {{ color: #1565C0; margin-top: 0; border-bottom: 2px solid #e9ecef; padding-bottom: 8px; }}
  h3 {{ color: #333; }}
  .grid-2 {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }}
  .summary-table {{ border-collapse: collapse; width: 100%; font-size: 13px; }}
  .summary-table th {{ background: #1565C0; color: white; padding: 8px 12px; text-align: left; }}
  .summary-table td {{ padding: 6px 12px; border-bottom: 1px solid #e9ecef; }}
  .summary-table tr:hover td {{ background: #f0f4ff; }}
  .badge-oss {{ background:#1565C0; color:white; padding:2px 6px; border-radius:3px; font-size:11px }}
  .badge-com {{ background:#C62828; color:white; padding:2px 6px; border-radius:3px; font-size:11px }}
  .badge-free {{ background:#558B2F; color:white; padding:2px 6px; border-radius:3px; font-size:11px }}
  .note-box {{ background:#fff3cd; border:1px solid #ffc107; border-radius:6px;
               padding:12px 16px; margin-bottom:16px; font-size:13px; }}
  table.metric-legend {{ font-size: 12px; border-collapse: collapse; }}
  table.metric-legend td {{ padding: 3px 8px; }}
  .footer {{ text-align: center; color: #888; font-size: 12px; margin-top: 32px; }}
  .section h2 {{ cursor: pointer; display: flex; justify-content: space-between;
                 align-items: center; margin-bottom: 0; }}
  .section h2 .toggle-icon {{ font-size: 14px; opacity: 0.5; transition: transform 0.2s; }}
  .section h2.collapsed .toggle-icon {{ transform: rotate(-90deg); }}
  .section-body {{ margin-top: 16px; }}
  .section-body.hidden {{ display: none; }}
</style>
<script>
function sortTable(th) {{
  const table = th.closest('table');
  const tbody = table.querySelector('tbody');
  const rows = Array.from(tbody.querySelectorAll('tr'));
  const headers = Array.from(th.closest('tr').querySelectorAll('th'));
  const col = headers.indexOf(th);
  const asc = th.dataset.sort !== 'asc';
  headers.forEach(h => {{
    h.dataset.sort = '';
    h.querySelector('.sort-icon').textContent = '⇅';
  }});
  th.dataset.sort = asc ? 'asc' : 'desc';
  th.querySelector('.sort-icon').textContent = asc ? '▲' : '▼';
  rows.sort((a, b) => {{
    const av = a.cells[col].textContent.trim();
    const bv = b.cells[col].textContent.trim();
    const an = parseFloat(av.replace(/,/g, ''));
    const bn = parseFloat(bv.replace(/,/g, ''));
    if (!isNaN(an) && !isNaN(bn)) return asc ? an - bn : bn - an;
    return asc ? av.localeCompare(bv) : bv.localeCompare(av);
  }});
  rows.forEach(r => tbody.appendChild(r));
}}
function toggleSection(h2) {{
  const body = h2.nextElementSibling;
  const collapsed = body.classList.toggle('hidden');
  h2.classList.toggle('collapsed', collapsed);
}}
</script>
</head>
<body>
<div class="header">
  <h1>PIV Software — Academic Usage Statistics</h1>
  <p>Based on OpenAlex data &bull; Generated {today} &bull; Years {YEAR_MIN}–{YEAR_MAX}</p>
</div>
<div class="content">

<div class="section">
  <h2 onclick="toggleSection(this)" class="collapsed">Methodology <span class="toggle-icon">▼</span></h2>
  <div class="section-body hidden">
  <div class="note-box">
    <strong>Unified methodology (fair, objective comparison):</strong>
    Every software package is counted with the <em>same metric</em>:
    <ul style="margin:4px 0">
      <li><strong>Exact-phrase full-text search</strong> via OpenAlex <code>default.search</code>:
          software names are searched as <em>quoted exact phrases</em> in the full text
          (title + abstract + body where indexed). Quotes prevent stemming false-positives.</li>
      <li><strong>"particle image velocimetry" required</strong>: this phrase must appear
          literally in the text. Guarantees PIV context without relying on ML tags.
          Eliminates false positives from multi-purpose names (e.g. DaVis, TSI, GPIV).</li>
      <li><strong>Spelling variants</strong>: alternate forms are OR-combined
          (e.g. "OpenPIV" OR "open piv").</li>
      <li><strong>Commercial software</strong>: the company name must also appear in the text
          (e.g. "Dantec" AND "Dynamic Studio").</li>
    </ul>
    Example filter: <code>default.search:"LaVision" AND "DaVis" AND "particle image velocimetry",publication_year:2010-{YEAR_MAX}</code><br>
    <strong>Known limitation:</strong> open-source tools (PIVlab, OpenPIV, …) are almost always
    named explicitly by authors; commercial tools are sometimes described only as "commercial software"
    without naming them. This means commercial tools may still be <em>undercounted</em> relative
    to their true usage, even with full-text search.<br>
    Note: OpenAlex covers ~96% of the literature. Only fully completed years ({YEAR_MIN}&ndash;{YEAR_MAX}) are shown.
  </div>
  <h3>Search terms per software</h3>
  <table class="metric-legend"><thead><tr>
    <th style="padding:2px 8px;text-align:left">Software</th>
    <th style="padding:2px 8px;text-align:left">Search term</th>
    <th style="padding:2px 8px;text-align:left">Note</th>
  </tr></thead><tbody>{search_rows}</tbody></table>
  </div>
</div>

<div class="section">
  <h2 onclick="toggleSection(this)">Summary Table <span class="toggle-icon">▼</span></h2>
  <div class="section-body">
  {summary_html}
  </div>
</div>

<div class="section">
  <h2 onclick="toggleSection(this)">Papers per Year (linear scale) <span class="toggle-icon">▼</span></h2>
  <div class="section-body">
  {fig_to_html(figs["yearly_linear"])}
  </div>
</div>

<div class="section">
  <h2 onclick="toggleSection(this)">Papers per Year (log scale — shows smaller software) <span class="toggle-icon">▼</span></h2>
  <div class="section-body">
  {fig_to_html(figs["yearly_log"])}
  </div>
</div>

<div class="section">
  <h2 onclick="toggleSection(this)">Total Papers {YEAR_MIN}–{YEAR_MAX} <span class="toggle-icon">▼</span></h2>
  <div class="section-body">
  {fig_to_html(figs["total_bars"])}
  </div>
</div>

<div class="section">
  <h2 onclick="toggleSection(this)">Journal Quality (Impact Indicator) <span class="toggle-icon">▼</span></h2>
  <div class="section-body">
  <div class="note-box">
    Shows the weighted mean <strong>2-year mean citedness</strong> of the journals in which
    papers mentioning each software are published. This is OpenAlex&rsquo;s open proxy for
    journal impact factor. A higher score means the software tends to appear in
    higher-cited journals. This can reveal whether a software is used only in
    niche or low-visibility outlets.
  </div>
  {fig_to_html(figs["journal_impact"])}
  </div>
</div>

<div class="section">
  <h2 onclick="toggleSection(this)">Market Share (% of Total PIV Literature) <span class="toggle-icon">▼</span></h2>
  <div class="section-body">
  {fig_to_html(figs["market_share"])}
  </div>
</div>

<div class="section">
  <h2 onclick="toggleSection(this)">Open Source vs. Commercial (combined) <span class="toggle-icon">▼</span></h2>
  <div class="section-body">
  {fig_to_html(figs["category"])}
  </div>
</div>

<div class="section">
  <h2 onclick="toggleSection(this)">Research Field Distribution <span class="toggle-icon">▼</span></h2>
  <div class="section-body">
  {fig_to_html(figs["field_treemap"])}
  </div>
</div>

<div class="section">
  <h2 onclick="toggleSection(this)">Field Distribution per Software (% of papers) <span class="toggle-icon">▼</span></h2>
  <div class="section-body">
  {fig_to_html(figs["field_heatmap"])}
  </div>
</div>

<div class="section">
  <h2 onclick="toggleSection(this)">Top Countries per Software <span class="toggle-icon">▼</span></h2>
  <div class="section-body">
  {fig_to_html(figs["country_bars"])}
  </div>
</div>

<div class="section">
  <h2 onclick="toggleSection(this)">Geographic Distribution (all software) <span class="toggle-icon">▼</span></h2>
  <div class="section-body">
  {fig_to_html(figs["country_map"])}
  </div>
</div>


<div class="footer">
  Data source: <a href="https://openalex.org">OpenAlex</a> &mdash;
  Generated by piv_stats.py &mdash; {today}
</div>

</div>
</body>
</html>"""
    return html


# ---------------------------------------------------------------------------
# Excel export
# ---------------------------------------------------------------------------

def export_excel(all_data: dict, df_yearly: pd.DataFrame,
                 df_totals: pd.DataFrame, df_fields: pd.DataFrame,
                 df_countries: pd.DataFrame,
                 df_impact: pd.DataFrame | None = None) -> Path:
    """Write an Excel workbook with all data."""
    path = OUTPUT_DIR / "piv_data.xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df_totals.to_excel(writer, sheet_name="Summary", index=False)

        # Wide yearly table
        pivot = df_yearly.pivot_table(
            index="software", columns="year", values="count", fill_value=0
        )
        pivot.to_excel(writer, sheet_name="Yearly counts")

        df_fields.to_excel(writer, sheet_name="Field distribution", index=False)

        df_countries_agg = (
            df_countries.groupby(["country", "country_code"])["count"]
            .sum().reset_index().sort_values("count", ascending=False)
        )
        df_countries_agg.to_excel(writer, sheet_name="Countries (all)", index=False)

        # Per-software country tables
        for sw_id, dat in all_data.items():
            sw_name = dat["meta"]["name"][:25]
            countries = dat.get("countries", [])
            if countries:
                df_c = pd.DataFrame(countries).head(TOP_N_COUNTRIES)
                sheet = f"Country_{sw_id[:15]}"
                df_c.to_excel(writer, sheet_name=sheet, index=False)

        # Journals
        journal_rows = []
        for sw_id, dat in all_data.items():
            sw_name = dat["meta"]["name"]
            for j in dat.get("journals", []):
                journal_rows.append({"software": sw_name,
                                     "journal": j["journal"],
                                     "count": j["count"]})
        if journal_rows:
            pd.DataFrame(journal_rows).to_excel(
                writer, sheet_name="Journals", index=False
            )

        if df_impact is not None and not df_impact.empty:
            df_impact.to_excel(writer, sheet_name="Journal Impact", index=False)

    return path


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    CACHE_DIR.mkdir(parents=True, exist_ok=True)

    print(f"\n=== PIV Software Usage Statistics ===")
    print(f"  Range: {YEAR_MIN}-{YEAR_MAX}  |  {len(SOFTWARE)} software packages\n")
    print("Fetching data from OpenAlex (cached results reused)...")
    all_data = collect_all()

    print("  [baseline] Total PIV papers per year ...", end="", flush=True)
    try:
        piv_baseline = collect_piv_baseline()
    except Exception as exc:
        print(f" (warning: {exc})", end="")
        piv_baseline = {}
    total_piv = sum(piv_baseline.values())
    print(f" {total_piv} total PIV papers in OpenAlex")

    print("\nBuilding data frames...")
    df_yearly = build_yearly_df(all_data)
    df_totals = build_totals_df(all_data)
    df_fields = build_fields_df(all_data)
    df_countries = build_countries_df(all_data)

    print("  Fetching journal impact scores ...", end="", flush=True)
    try:
        source_impacts = fetch_source_impacts(all_data)
        df_impact = build_impact_df(all_data, source_impacts)
        print(f" {len(source_impacts)} journals looked up")
    except Exception as exc:
        print(f" (warning: {exc})")
        df_impact = pd.DataFrame()

    print("\nGenerating HTML report...")
    html = generate_html_report(all_data, df_yearly, df_totals, df_fields,
                                df_countries, piv_baseline=piv_baseline,
                                df_impact=df_impact)
    report_path = OUTPUT_DIR / "piv_report.html"
    report_path.write_text(html, encoding="utf-8")
    print(f"  Report: {report_path}")

    print("Exporting Excel workbook...")
    xl_path = export_excel(all_data, df_yearly, df_totals, df_fields, df_countries, df_impact)
    print(f"  Excel:  {xl_path}")

    print("\nDone! Top 10 by total papers:")
    top10 = df_totals.head(10)[["software", "category", "total"]]
    for _, r in top10.iterrows():
        print(f"  {r['software']:<30} {r['total']:>6}  ({r['category']})")

    return report_path, xl_path


if __name__ == "__main__":
    main()
