import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import json
import urllib.request
import urllib.parse
import numpy as np
import base64
from pathlib import Path

st.set_page_config(page_title="USDA County Production Dashboard", layout="wide")
_CACHE_VERSION = "v9"   # bump to invalidate all @st.cache_data on deploy

# ── Paths ─────────────────────────────────────────────────────────────────────
HERE       = Path(__file__).parent
DATA_PATH  = HERE / "data" / "2025 RMA Production Data.xlsx"
LOGO_50YR  = HERE / "assets" / "50 Year logo JSA.png"
LOGO_TRANS = HERE / "assets" / "Transparent Smal logo.png"
LOGO_FULL  = HERE / "assets" / "logo-full.png"

# ── NASS API ───────────────────────────────────────────────────────────────────
NASS_API_KEY  = "9A6D1EB8-4D94-3221-BA0C-ADD4533EA0C1"
NASS_BASE_URL = "https://quickstats.nass.usda.gov/api/api_GET/"
NASS_YEARS             = [2025, 2024, 2023, 2022]
_NASS_BENCHMARK_YEAR   = 2023   # most-complete county year — used for % reporting KPI

# Metrics available in the NASS tab
NASS_METRICS     = ["Production (bu)", "Planted Acres", "Harvested Acres",
                    "Yield (bu/ac)", "Prevent Plant Acres"]
NASS_CHANGE_OPTS = ["Absolute", "vs Prior Year", "vs Selected Year", "vs 3-Yr Avg"]

_METRIC_TO_STAT = {
    "Production (bu)":    "production",
    "Planted Acres":      "planted",
    "Harvested Acres":    "harvested",
    "Yield (bu/ac)":      "yield",
    "Prevent Plant Acres": "prevent_plant",
}

# Per-crop API params for each stat type
NASS_CROP_STAT_PARAMS = {
    "Corn": {
        "production":    {"commodity_desc": "CORN", "util_practice_desc": "GRAIN"},
        "planted":       {"commodity_desc": "CORN"},
        "harvested":     {"commodity_desc": "CORN", "util_practice_desc": "GRAIN"},
        "yield":         {"commodity_desc": "CORN", "util_practice_desc": "GRAIN"},
        "prevent_plant": {"commodity_desc": "CORN"},
    },
    "Soybeans": {
        "production":    {"commodity_desc": "SOYBEANS"},
        "planted":       {"commodity_desc": "SOYBEANS"},
        "harvested":     {"commodity_desc": "SOYBEANS"},
        "yield":         {"commodity_desc": "SOYBEANS"},
        "prevent_plant": {"commodity_desc": "SOYBEANS"},
    },
    "Wheat": {
        "production":    {"commodity_desc": "WHEAT", "class_desc": "ALL CLASSES"},
        "planted":       {"commodity_desc": "WHEAT", "class_desc": "ALL CLASSES"},
        "harvested":     {"commodity_desc": "WHEAT", "class_desc": "ALL CLASSES"},
        "yield":         {"commodity_desc": "WHEAT", "class_desc": "ALL CLASSES"},
        "prevent_plant": {"commodity_desc": "WHEAT", "class_desc": "ALL CLASSES"},
    },
    "Sorghum": {
        "production":    {"commodity_desc": "SORGHUM", "util_practice_desc": "GRAIN"},
        "planted":       {"commodity_desc": "SORGHUM"},
        "harvested":     {"commodity_desc": "SORGHUM", "util_practice_desc": "GRAIN"},
        "yield":         {"commodity_desc": "SORGHUM", "util_practice_desc": "GRAIN"},
        "prevent_plant": {"commodity_desc": "SORGHUM"},
    },
}

# Base API params per stat type
# prevent_plant uses AREA PLANTED — rows are filtered to "PREVENTED" in load_nass_stat
NASS_STAT_BASE = {
    "production":    {"statisticcat_desc": "PRODUCTION",     "unit_desc": "BU"},
    "planted":       {"statisticcat_desc": "AREA PLANTED",    "unit_desc": "ACRES"},
    "harvested":     {"statisticcat_desc": "AREA HARVESTED",  "unit_desc": "ACRES"},
    "yield":         {"statisticcat_desc": "YIELD",           "unit_desc": "BU / ACRE"},
    "prevent_plant": {"statisticcat_desc": "AREA PLANTED",    "unit_desc": "ACRES"},
}

# Legacy — kept for backward compat with any cached references
NASS_CROP_PARAMS = {
    "Corn":     {"commodity_desc": "CORN",    "util_practice_desc": "GRAIN"},
    "Soybeans": {"commodity_desc": "SOYBEANS"},
    "Wheat":    {"commodity_desc": "WHEAT",   "class_desc": "ALL CLASSES"},
    "Sorghum":  {"commodity_desc": "SORGHUM", "util_practice_desc": "GRAIN"},
}

# ── State lookups ─────────────────────────────────────────────────────────────
# RMA subset (used for county FIPS name-lookup only)
STATE_FIPS = {
    "AL": "01", "AR": "05", "CO": "08", "GA": "13", "IA": "19",
    "ID": "16", "IL": "17", "IN": "18", "KS": "20", "KY": "21",
    "MD": "24", "MI": "26", "MN": "27", "MO": "29", "MS": "28",
    "MT": "30", "NC": "37", "ND": "38", "NE": "31", "OH": "39",
    "OK": "40", "OR": "41", "PA": "42", "SC": "45", "SD": "46",
    "TN": "47", "TX": "48", "VA": "51", "WA": "53", "WI": "55",
    "WY": "56",
}

# Full 50-state FIPS — used for NASS county map lookups
STATE_FIPS_ALL = {
    "AL": "01", "AK": "02", "AZ": "04", "AR": "05", "CA": "06",
    "CO": "08", "CT": "09", "DE": "10", "FL": "12", "GA": "13",
    "HI": "15", "ID": "16", "IL": "17", "IN": "18", "IA": "19",
    "KS": "20", "KY": "21", "LA": "22", "ME": "23", "MD": "24",
    "MA": "25", "MI": "26", "MN": "27", "MS": "28", "MO": "29",
    "MT": "30", "NE": "31", "NV": "32", "NH": "33", "NJ": "34",
    "NM": "35", "NY": "36", "NC": "37", "ND": "38", "OH": "39",
    "OK": "40", "OR": "41", "PA": "42", "RI": "44", "SC": "45",
    "SD": "46", "TN": "47", "TX": "48", "UT": "49", "VT": "50",
    "VA": "51", "WA": "53", "WV": "54", "WI": "55", "WY": "56",
}

ABBR_TO_NAME = {
    "AL": "Alabama",        "AK": "Alaska",         "AZ": "Arizona",
    "AR": "Arkansas",       "CA": "California",     "CO": "Colorado",
    "CT": "Connecticut",    "DE": "Delaware",       "FL": "Florida",
    "GA": "Georgia",        "HI": "Hawaii",         "ID": "Idaho",
    "IL": "Illinois",       "IN": "Indiana",        "IA": "Iowa",
    "KS": "Kansas",         "KY": "Kentucky",       "LA": "Louisiana",
    "ME": "Maine",          "MD": "Maryland",       "MA": "Massachusetts",
    "MI": "Michigan",       "MN": "Minnesota",      "MS": "Mississippi",
    "MO": "Missouri",       "MT": "Montana",        "NE": "Nebraska",
    "NV": "Nevada",         "NH": "New Hampshire",  "NJ": "New Jersey",
    "NM": "New Mexico",     "NY": "New York",       "NC": "North Carolina",
    "ND": "North Dakota",   "OH": "Ohio",           "OK": "Oklahoma",
    "OR": "Oregon",         "PA": "Pennsylvania",   "RI": "Rhode Island",
    "SC": "South Carolina", "SD": "South Dakota",   "TN": "Tennessee",
    "TX": "Texas",          "UT": "Utah",           "VT": "Vermont",
    "VA": "Virginia",       "WA": "Washington",     "WV": "West Virginia",
    "WI": "Wisconsin",      "WY": "Wyoming",
}

STATE_CENTROIDS = {
    "AL": (-86.8,  32.8), "AR": (-92.4,  34.9), "AZ": (-111.6, 34.3),
    "CA": (-119.4, 37.2), "CO": (-105.5, 39.0), "CT": (-72.7,  41.6),
    "DE": (-75.5,  38.9), "FL": (-81.5,  27.8), "GA": (-83.4,  32.7),
    "IA": (-93.1,  42.0), "ID": (-114.5, 44.4), "IL": (-89.2,  40.0),
    "IN": (-86.3,  40.3), "KS": (-98.4,  38.5), "KY": (-84.9,  37.5),
    "LA": (-92.1,  30.5), "MD": (-76.8,  39.0), "ME": (-69.2,  44.7),
    "MI": (-84.5,  44.3), "MN": (-94.3,  46.4), "MO": (-92.5,  38.4),
    "MS": (-89.7,  32.7), "MT": (-110.5, 46.9), "NC": (-79.4,  35.6),
    "ND": (-100.5, 47.5), "NE": (-99.9,  41.5), "NJ": (-74.4,  40.1),
    "NM": (-106.0, 34.5), "NY": (-75.5,  42.9), "OH": (-82.8,  40.4),
    "OK": (-97.5,  35.5), "OR": (-120.6, 44.1), "PA": (-77.2,  40.9),
    "SC": (-80.9,  33.8), "SD": (-100.2, 44.4), "TN": (-86.7,  35.8),
    "TX": (-99.3,  31.5), "UT": (-111.1, 39.3), "VA": (-78.7,  37.5),
    "WA": (-120.5, 47.4), "WI": (-89.8,  44.5), "WV": (-80.4,  38.7),
    "WY": (-107.6, 43.0),
}

# ── RMA metric mappings ───────────────────────────────────────────────────────
METRIC_COL = {
    "Production":            "Reported Production",
    "Production Acres":      "Reported Production Acres",
    "Yield":                 "Reported Yield Mean",
    "Prevent Planted Acres": "Prev Plant Acres",
}
METRIC_UNIT = {
    "Production": "bu", "Production Acres": "ac",
    "Yield": "bu/ac",   "Prevent Planted Acres": "ac",
}
METRIC_FMT = {
    "Production": ",.0f", "Production Acres": ",.0f",
    "Yield": ".1f",        "Prevent Planted Acres": ",.0f",
}
COLOR_SCALE = {
    "Production": "YlOrBr", "Production Acres": "YlGn",
    "Yield": "RdYlGn",      "Prevent Planted Acres": "OrRd",
}

DISPLAY_DIVISOR = {
    "Production": 1_000_000, "Production Acres": 100_000,
    "Yield": 1,               "Prevent Planted Acres": 100_000,
}
DISPLAY_UNIT = {
    "Production": "M bu",     "Production Acres": "×100K ac",
    "Yield": "bu/ac",          "Prevent Planted Acres": "×100K ac",
}

# ── JPSI brand palette ────────────────────────────────────────────────────────
DARK    = "#0e1614"
PANEL   = "#162019"
SURFACE = "#1e2e2a"
BORDER  = "#243328"
TEXT    = "#e4e8f0"
MUTED   = "#7a9990"
ACCENT  = "#4ade80"
LAND    = "#1a2720"


# ── RMA Data ──────────────────────────────────────────────────────────────────
@st.cache_data
def load_data():
    out = {}
    for crop in ["Corn", "Soybeans", "Wheat"]:
        try:
            df = pd.read_excel(DATA_PATH, sheet_name=crop)
        except Exception:
            continue
        df.columns  = df.columns.str.strip()
        df["State"]    = df["State"].str.strip()
        df["County"]   = df["County"].str.strip()
        df["Practice"] = df["Practice"].str.strip()
        df["PG"] = np.where(
            df["Practice"].str.startswith("Irrigated"),          "Irrigated",
            np.where(df["Practice"].str.startswith("Non-Irrigated"), "Non-Irrigated", "Invalid"),
        )
        df = df[df["PG"] != "Invalid"].copy()
        if "Type" in df.columns:
            df["Type"] = df["Type"].str.strip()
        for col in METRIC_COL.values():
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        key_cols = [c for c in ["State", "County", "PG", "Type", "Yield Year"] if c in df.columns]
        idx_keep = df.groupby(key_cols)["Reported Production"].idxmax()
        df = df.loc[idx_keep].reset_index(drop=True)
        out[crop] = df
    return out


# ── NASS Data ─────────────────────────────────────────────────────────────────
@st.cache_data
def load_nass_stat(crop: str, year: int, stat_type: str,
                   cache_ver: str = _CACHE_VERSION) -> pd.DataFrame:
    """Generic NASS county-level loader for any stat type.
    stat_type: 'production' | 'planted' | 'harvested' | 'yield' | 'prevent_plant'
    Returns DataFrame with [State, County, fips, Value].
    """
    params = {
        "key":            NASS_API_KEY,
        "source_desc":    "SURVEY",
        "sector_desc":    "CROPS",
        "agg_level_desc": "COUNTY",
        "year":           str(year),
        "format":         "JSON",
    }
    params.update(NASS_STAT_BASE[stat_type])
    params.update(NASS_CROP_STAT_PARAMS[crop][stat_type])
    url = NASS_BASE_URL + "?" + urllib.parse.urlencode(params)
    try:
        with urllib.request.urlopen(url, timeout=45) as r:
            raw = json.load(r)
    except Exception as e:
        st.warning(f"NASS API error for {crop} {year} {stat_type}: {e}")
        return pd.DataFrame(columns=["State", "County", "fips", "Value"])

    records = raw.get("data", [])
    if not records:
        return pd.DataFrame(columns=["State", "County", "fips", "Value"])

    df = pd.DataFrame(records)
    needed = ["state_alpha", "county_name", "state_fips_code",
              "county_ansi", "prodn_practice_desc", "short_desc", "Value"]
    df = df[[c for c in needed if c in df.columns]].copy()

    # For prevent_plant: keep only rows whose short_desc contains "PREVENT"
    # (AREA PLANTED queries return both regular planted and prevented-planted rows)
    if stat_type == "prevent_plant":
        if "short_desc" in df.columns:
            df = df[df["short_desc"].str.upper().str.contains("PREVENT", na=False)]
        if df.empty:
            return pd.DataFrame(columns=["State", "County", "fips", "Value"])

    # Drop state-level and aggregate rows
    df = df[~df["county_ansi"].isin(["998", "000", "999"])]
    df = df[~df["county_name"].str.strip().str.lower().str.startswith("other")]

    df["Value"] = pd.to_numeric(
        df["Value"].str.replace(",", "", regex=False).str.strip(),
        errors="coerce",
    ).fillna(0)
    df["fips"]   = df["state_fips_code"].str.zfill(2) + df["county_ansi"].str.zfill(3)
    df["State"]  = df["state_alpha"].str.strip()
    df["County"] = df["county_name"].str.strip().str.title()

    # Dedup: prefer ALL PRODUCTION PRACTICES row; fallback to max-value row
    key      = ["State", "County", "fips"]
    all_prac = "ALL PRODUCTION PRACTICES"
    if "prodn_practice_desc" in df.columns:
        has_all = df[df["prodn_practice_desc"] == all_prac].copy()
        no_all  = df[~df["fips"].isin(has_all["fips"].unique())].copy()
        if not no_all.empty:
            no_all = no_all.loc[no_all.groupby(key)["Value"].idxmax()]
        df = pd.concat([has_all, no_all], ignore_index=True)

    return df[key + ["Value"]].reset_index(drop=True)


@st.cache_data
def load_nass_state(crop: str, year: int, stat_type: str,
                    cache_ver: str = _CACHE_VERSION) -> pd.DataFrame:
    """Query NASS at STATE level for official state-reported totals.
    Uses domain_desc=TOTAL so we get the main survey figures, not organic
    or other sub-domain breakdowns.
    Returns DataFrame with [State, Value] — one row per state.
    """
    params = {
        "key":                   NASS_API_KEY,
        "source_desc":           "SURVEY",
        "sector_desc":           "CROPS",
        "agg_level_desc":        "STATE",
        "domain_desc":           "TOTAL",
        # NASS stores multiple estimates per crop year (Aug/Sep/Nov forecasts
        # plus the January Annual Summary).  reference_period_desc=YEAR
        # isolates the final Annual Summary and ignores the in-season forecasts.
        "reference_period_desc": "YEAR",
        "year":                  str(year),
        "format":                "JSON",
    }
    params.update(NASS_STAT_BASE[stat_type])
    params.update(NASS_CROP_STAT_PARAMS[crop][stat_type])
    url = NASS_BASE_URL + "?" + urllib.parse.urlencode(params)
    try:
        with urllib.request.urlopen(url, timeout=45) as r:
            raw = json.load(r)
    except Exception as e:
        st.warning(f"NASS state API error for {crop} {year} {stat_type}: {e}")
        return pd.DataFrame(columns=["State", "Value"])

    records = raw.get("data", [])
    if not records:
        return pd.DataFrame(columns=["State", "Value"])

    df = pd.DataFrame(records)
    needed = ["state_alpha", "prodn_practice_desc", "short_desc", "Value"]
    df = df[[c for c in needed if c in df.columns]].copy()

    # For prevent_plant: keep only rows whose short_desc contains "PREVENT"
    if stat_type == "prevent_plant":
        if "short_desc" in df.columns:
            df = df[df["short_desc"].str.upper().str.contains("PREVENT", na=False)]
        if df.empty:
            return pd.DataFrame(columns=["State", "Value"])

    df["Value"] = pd.to_numeric(
        df["Value"].str.replace(",", "", regex=False).str.strip(),
        errors="coerce",
    ).fillna(0)
    df["State"] = df["state_alpha"].str.strip()

    # Keep only known 50-state abbreviations — explicitly excludes "US" (national
    # total), "PR", "GU", "VI", "OTHER STATES", etc.  If we kept "US" it would be
    # summed alongside the 50 states and inflate the national total ~2×.
    df = df[df["State"].isin(set(STATE_FIPS_ALL.keys()))]

    # Dedup step 1: prefer ALL PRODUCTION PRACTICES row; fallback to max-value row
    all_prac = "ALL PRODUCTION PRACTICES"
    if "prodn_practice_desc" in df.columns:
        has_all = df[df["prodn_practice_desc"] == all_prac].copy()
        no_all  = df[~df["State"].isin(has_all["State"].unique())].copy()
        if not no_all.empty:
            no_all = no_all.loc[no_all.groupby("State")["Value"].idxmax()]
        df = pd.concat([has_all, no_all], ignore_index=True)

    # Dedup step 2: guarantee exactly ONE row per state.  NASS can return multiple
    # "ALL PRODUCTION PRACTICES" rows per state (e.g. different short_desc values);
    # keeping them all causes duplicate map labels and an inflated national total.
    df = df.loc[df.groupby("State")["Value"].idxmax()].reset_index(drop=True)

    return df[["State", "Value"]].reset_index(drop=True)


@st.cache_data
def load_nass_county(crop: str, year: int = 2025,
                     cache_ver: str = _CACHE_VERSION) -> pd.DataFrame:
    """Load county-level production data.  Kept as a standalone function
    (not delegating to load_nass_stat) to avoid Streamlit cache-within-cache
    issues that can cause stale or incorrect return values.
    Returns [State, County, fips, Production].
    """
    params = {
        "key":               NASS_API_KEY,
        "source_desc":       "SURVEY",
        "sector_desc":       "CROPS",
        "statisticcat_desc": "PRODUCTION",
        "unit_desc":         "BU",
        "agg_level_desc":    "COUNTY",
        "year":              str(year),
        "format":            "JSON",
    }
    params.update(NASS_CROP_PARAMS[crop])
    url = NASS_BASE_URL + "?" + urllib.parse.urlencode(params)
    try:
        with urllib.request.urlopen(url, timeout=45) as r:
            raw = json.load(r)
    except Exception as e:
        st.warning(f"NASS API error for {crop} {year}: {e}")
        return pd.DataFrame(columns=["State", "County", "fips", "Production"])

    records = raw.get("data", [])
    if not records:
        return pd.DataFrame(columns=["State", "County", "fips", "Production"])

    df = pd.DataFrame(records)
    needed = ["state_alpha", "county_name", "state_fips_code",
              "county_ansi", "prodn_practice_desc", "Value"]
    df = df[[c for c in needed if c in df.columns]].copy()

    df = df[~df["county_ansi"].isin(["998", "000", "999"])]
    df = df[~df["county_name"].str.strip().str.lower().str.startswith("other")]

    df["Production"] = pd.to_numeric(
        df["Value"].str.replace(",", "", regex=False).str.strip(),
        errors="coerce",
    ).fillna(0)
    df["fips"]   = df["state_fips_code"].str.zfill(2) + df["county_ansi"].str.zfill(3)
    df["State"]  = df["state_alpha"].str.strip()
    df["County"] = df["county_name"].str.strip().str.title()

    key      = ["State", "County", "fips"]
    all_prac = "ALL PRODUCTION PRACTICES"
    if "prodn_practice_desc" in df.columns:
        has_all = df[df["prodn_practice_desc"] == all_prac].copy()
        no_all  = df[~df["fips"].isin(has_all["fips"].unique())].copy()
        if not no_all.empty:
            no_all = no_all.loc[no_all.groupby(key)["Production"].idxmax()]
        df = pd.concat([has_all, no_all], ignore_index=True)

    return df[key + ["Production"]].reset_index(drop=True)


# ── GeoJSON & lookups ─────────────────────────────────────────────────────────
@st.cache_data
def load_geojson():
    url = "https://raw.githubusercontent.com/plotly/datasets/master/geojson-counties-fips.json"
    with urllib.request.urlopen(url) as r:
        return json.load(r)


@st.cache_data
def get_state_geojson(_geo, sfips: str) -> dict:
    """Return a cached FeatureCollection filtered to one state.
    Returns the original feature objects unmodified so that every field
    (including the top-level 'id' that Plotly uses to match choropleth
    locations) is preserved exactly as it is in the source GeoJSON.
    """
    return {
        "type": "FeatureCollection",
        "features": [f for f in _geo["features"]
                     if f["properties"]["STATE"] == sfips],
    }


@st.cache_data
def build_fips_lookup(_geo):
    inv = {v: k for k, v in STATE_FIPS_ALL.items()}  # all 50 states — needed for NASS
    lk  = {}
    for feat in _geo["features"]:
        p = feat["properties"]
        if p["STATE"] not in inv:
            continue
        abbr  = inv[p["STATE"]]
        name  = p["NAME"]
        fips5 = p["STATE"] + p["COUNTY"]
        lk[(abbr, name.lower())] = fips5
        for suf in [" county", " parish", " borough", " city", " census area"]:
            if name.lower().endswith(suf):
                lk[(abbr, name.lower()[: -len(suf)])] = fips5
    return lk


def resolve_fips(state, county, lk):
    c = county.lower()
    return (
        lk.get((state, c))
        or lk.get((state, c + " county"))
        or lk.get((state, c.replace("st.", "saint").replace(".", "").strip()))
        or lk.get((state, c.replace(" ", "")))
    )


def _poly_centroid_area(coords):
    x, y = coords[:, 0], coords[:, 1]
    a    = x[:-1] * y[1:] - x[1:] * y[:-1]
    A    = 0.5 * a.sum()
    area = abs(A)
    if area < 1e-10:
        return float(x.mean()), float(y.mean()), area
    cx = float(((x[:-1] + x[1:]) * a).sum() / (6 * A))
    cy = float(((y[:-1] + y[1:]) * a).sum() / (6 * A))
    return cx, cy, area


@st.cache_data
def build_centroid_lookup(_geo):
    centroids = {}
    for feat in _geo["features"]:
        fips5 = feat["properties"]["STATE"] + feat["properties"]["COUNTY"]
        geom  = feat["geometry"]
        try:
            if geom["type"] == "Polygon":
                coords = np.array(geom["coordinates"][0])
                centroids[fips5] = _poly_centroid_area(coords)
            elif geom["type"] == "MultiPolygon":
                best, best_area = None, 0
                for part in geom["coordinates"]:
                    c = np.array(part[0])
                    a = abs(0.5 * (c[:-1, 0] * c[1:, 1] - c[1:, 0] * c[:-1, 1]).sum())
                    if a > best_area:
                        best, best_area = c, a
                if best is not None:
                    centroids[fips5] = _poly_centroid_area(best)
        except Exception:
            pass
    return centroids


# ── Formatting helpers ────────────────────────────────────────────────────────
def format_label(val, metric):
    if pd.isna(val) or val == 0:
        return ""
    if metric == "Yield":
        return f"{val:.0f}"
    if metric == "Production":
        m = val / 1_000_000
        return f"{m:.2f}" if m >= 0.005 else f"{val / 100_000:.2f}"
    return f"{val / 100_000:.2f}"


def format_state_label(val, metric):
    if pd.isna(val) or val == 0:
        return ""
    if metric == "Yield":
        return f"{val:.0f}"
    if metric == "Production":
        return f"{val / 1_000_000:.1f}"
    return f"{val / 100_000:.1f}"


def format_nass_label(val):
    if pd.isna(val) or val == 0:
        return ""
    return f"{val / 1_000_000:.1f}"


def format_nass_chg_label(val):
    """Format a % change value for map labels, e.g. '+12.3%'."""
    if pd.isna(val) or abs(val) < 0.05:
        return ""
    return f"{val:+.1f}%"


def format_nass_acres_label(val):
    if pd.isna(val) or val < 500:
        return ""
    return f"{val / 1_000:,.0f}K"


def format_nass_yield_label(val):
    if pd.isna(val) or val == 0:
        return ""
    return f"{val:.0f}"


# ── NASS view helpers ─────────────────────────────────────────────────────────
def _nass_view_cfg(metric: str, change_view: str) -> dict:
    """Return render-config dict given a metric and change_view."""
    if change_view != "Absolute":
        return {
            "cscale": "RdYlGn", "diverging": True, "clabel": "Change (%)",
            "hover_fmt": ":+.1f", "hover_sfx": "%", "label_unit": "% chg",
            "label_fn": format_nass_chg_label,
            "rank_unit": "%", "rank_div": 1, "rank_fmt": "+.1f",
        }
    _abs_cfgs = {
        "Production (bu)": {
            "cscale": "YlOrBr", "diverging": False, "clabel": "Production<br>(bu)",
            "hover_fmt": ":,.0f", "hover_sfx": " bu", "label_unit": "M bu",
            "label_fn": format_nass_label,
            "rank_unit": "M bu", "rank_div": 1_000_000, "rank_fmt": ",.2f",
        },
        "Planted Acres": {
            "cscale": "YlGn", "diverging": False, "clabel": "Planted<br>Acres",
            "hover_fmt": ":,.0f", "hover_sfx": " ac", "label_unit": "K ac",
            "label_fn": format_nass_acres_label,
            "rank_unit": "K ac", "rank_div": 1_000, "rank_fmt": ",.1f",
        },
        "Harvested Acres": {
            "cscale": "BuGn", "diverging": False, "clabel": "Harvested<br>Acres",
            "hover_fmt": ":,.0f", "hover_sfx": " ac", "label_unit": "K ac",
            "label_fn": format_nass_acres_label,
            "rank_unit": "K ac", "rank_div": 1_000, "rank_fmt": ",.1f",
        },
        "Yield (bu/ac)": {
            "cscale": "RdYlGn", "diverging": False, "clabel": "Yield<br>(bu/ac)",
            "hover_fmt": ":.1f", "hover_sfx": " bu/ac", "label_unit": "bu/ac",
            "label_fn": format_nass_yield_label,
            "rank_unit": "bu/ac", "rank_div": 1, "rank_fmt": ".1f",
        },
        "Prevent Plant Acres": {
            "cscale": "OrRd", "diverging": False, "clabel": "Prevent<br>Plant Acres",
            "hover_fmt": ":,.0f", "hover_sfx": " ac", "label_unit": "K ac",
            "label_fn": format_nass_acres_label,
            "rank_unit": "K ac", "rank_div": 1_000, "rank_fmt": ",.1f",
        },
    }
    return _abs_cfgs[metric]


def _state_pct_change(cur_df: pd.DataFrame, cmp_df: pd.DataFrame) -> pd.DataFrame:
    """Compute state-level % change: (cur - cmp) / cmp * 100.
    Both inputs are [State, Value]. Returns [State, Value] with % change."""
    mc = cur_df.merge(cmp_df.rename(columns={"Value": "Base"}), on="State", how="inner")
    mc["Value"] = (mc["Value"] - mc["Base"]) / mc["Base"].replace(0, np.nan) * 100
    return mc[["State", "Value"]].dropna(subset=["Value"])


def _state_pct_change_avg(cur_df: pd.DataFrame, frames: list) -> pd.DataFrame:
    """Compute state-level % change vs average of a list of prior-year DataFrames."""
    if not frames:
        return pd.DataFrame(columns=["State", "Value"])
    avg = (pd.concat(frames).groupby("State")["Value"].mean()
           .reset_index().rename(columns={"Value": "Base"}))
    mc = cur_df.merge(avg, on="State", how="inner")
    mc["Value"] = (mc["Value"] - mc["Base"]) / mc["Base"].replace(0, np.nan) * 100
    return mc[["State", "Value"]].dropna(subset=["Value"])


def _load_for_metric(crop: str, year: int, stat_type: str) -> pd.DataFrame:
    """Return [State, County, fips, Value] for any stat type.
    Production is routed through load_nass_county (its own validated API call)
    to guarantee correct figures; all other stats go through load_nass_stat.
    """
    if stat_type == "production":
        return load_nass_county(crop, year, _CACHE_VERSION).rename(columns={"Production": "Value"})
    return load_nass_stat(crop, year, stat_type, _CACHE_VERSION)


def get_nass_view_data(crop: str, year: int, metric: str, change_view: str, comp_year=None):
    """
    Load and compute the view metric for any metric + change_view combination.
    Returns (county_df [State, County, Value], state_df [State, Value]).
    Absolute view  → Value = raw stat (bu / acres / bu per ac).
    Change views   → Value = % change vs comparison period.
    """
    stat_type = _METRIC_TO_STAT[metric]
    df_cur    = _load_for_metric(crop, year, stat_type)

    def _agg_c(df):
        if metric == "Yield (bu/ac)":
            return df.groupby(["State", "County"])["Value"].mean().reset_index()
        return df.groupby(["State", "County"])["Value"].sum().reset_index()

    def _agg_s(df):
        if metric == "Yield (bu/ac)":
            return df.groupby("State")["Value"].mean().reset_index()
        return df.groupby("State")["Value"].sum().reset_index()

    if change_view == "Absolute" or df_cur.empty:
        return _agg_c(df_cur), _agg_s(df_cur)

    def _pct(cur_s, cmp_s):
        return (cur_s - cmp_s) / cmp_s.replace(0, np.nan) * 100

    if change_view == "vs Prior Year":
        df_cmp = _load_for_metric(crop, year - 1, stat_type)
    elif change_view == "vs Selected Year":
        df_cmp = _load_for_metric(crop, comp_year, stat_type) if comp_year else df_cur
    else:  # vs 3-Yr Avg
        prior_years = [y for y in [year - 1, year - 2, year - 3] if y >= 2022]
        frames = [_load_for_metric(crop, y, stat_type) for y in prior_years]
        frames = [f for f in frames if not f.empty]
        if not frames:
            return pd.DataFrame(columns=["State", "County", "Value"]), pd.DataFrame(columns=["State", "Value"])
        avg_c = (pd.concat([_agg_c(d) for d in frames])
                 .groupby(["State", "County"])["Value"].mean()
                 .reset_index().rename(columns={"Value": "Base"}))
        cur_c = _agg_c(df_cur)
        mc = cur_c.merge(avg_c, on=["State", "County"], how="inner")
        mc["Value"] = _pct(mc["Value"], mc["Base"])
        avg_s = (pd.concat([_agg_s(d) for d in frames])
                 .groupby("State")["Value"].mean()
                 .reset_index().rename(columns={"Value": "Base"}))
        cur_s = _agg_s(df_cur)
        ms = cur_s.merge(avg_s, on="State", how="inner")
        ms["Value"] = _pct(ms["Value"], ms["Base"])
        return mc[["State", "County", "Value"]].dropna(subset=["Value"]), ms[["State", "Value"]].dropna(subset=["Value"])

    # Prior Year / Selected Year shared path
    cur_c = _agg_c(df_cur)
    cmp_c = _agg_c(df_cmp).rename(columns={"Value": "Base"})
    mc    = cur_c.merge(cmp_c, on=["State", "County"], how="inner")
    mc["Value"] = _pct(mc["Value"], mc["Base"])

    cur_s = _agg_s(df_cur)
    cmp_s = _agg_s(df_cmp).rename(columns={"Value": "Base"})
    ms    = cur_s.merge(cmp_s, on="State", how="inner")
    ms["Value"] = _pct(ms["Value"], ms["Base"])
    return mc[["State", "County", "Value"]].dropna(subset=["Value"]), ms[["State", "Value"]].dropna(subset=["Value"])


# ── Aggregation ───────────────────────────────────────────────────────────────
def filter_practice(df, practice):
    return df if practice == "All" else df[df["PG"] == practice]


def agg_data(df, practice, metric, group_cols):
    col = METRIC_COL[metric]
    df  = filter_practice(df, practice)
    if metric == "Yield":
        prod   = df.groupby(group_cols)["Reported Production"].sum()
        acres  = df.groupby(group_cols)["Reported Production Acres"].sum()
        result = (prod / acres.replace(0, np.nan)).reset_index()
        result.columns = group_cols + [col]
    else:
        result = df.groupby(group_cols)[col].sum().reset_index()
    return result


# ── Logo helpers ──────────────────────────────────────────────────────────────
@st.cache_data
def load_logo(path):
    with open(path, "rb") as f:
        return "data:image/png;base64," + base64.b64encode(f.read()).decode()


def _add_logo(fig, logo_src, size=0.13, opacity=0.92, x=0.99, y=0.01,
              yanchor="bottom", layer="above"):
    fig.add_layout_image(
        source=logo_src, xref="paper", yref="paper",
        x=x, y=y, xanchor="right", yanchor=yanchor,
        sizex=size, sizey=size, sizing="contain",
        opacity=opacity, layer=layer,
    )


def _base_layout(title):
    return dict(
        paper_bgcolor=DARK, plot_bgcolor=SURFACE,
        font=dict(color=TEXT, family="Arial"),
        title=dict(text=title, font=dict(size=15, color=ACCENT)),
        margin=dict(l=0, r=0, t=50, b=0),
    )


# ── Shared county-label placement ────────────────────────────────────────────
def _place_labels(fig, fips_list, value_series, centroids, metric_or_fn):
    """Add adaptive Scattergeo text labels to a county fig.
    metric_or_fn: RMA metric string OR callable(val)->str for NASS.
    """
    county_areas = [centroids[f][2] for f in fips_list if f in centroids]
    if county_areas:
        avg_area   = float(np.mean(county_areas))
        label_size = int(np.clip(9 + np.log(max(avg_area, 0.01) / 0.05) * 2.0, 9, 15))
    else:
        avg_area   = 0.1
        label_size = 10

    fmt_fn = (lambda v: format_label(v, metric_or_fn)) if isinstance(metric_or_fn, str) \
             else metric_or_fn

    candidates = []
    for fips, val in zip(fips_list, value_series):
        label = fmt_fn(val)
        if label and fips in centroids:
            cx, cy, area = centroids[fips]
            candidates.append((area, cx, cy, label))
    candidates.sort(reverse=True)

    min_sep = float(np.clip(0.15 + avg_area * 0.8, 0.20, 0.45))
    placed, lons, lats, texts = [], [], [], []
    for area, cx, cy, label in candidates:
        if not any((cx - px) ** 2 + (cy - py) ** 2 < min_sep ** 2 for px, py in placed):
            placed.append((cx, cy))
            lons.append(cx); lats.append(cy); texts.append(label)

    if lons:
        fig.add_trace(go.Scattergeo(
            lon=lons, lat=lats, text=texts, mode="text",
            textfont=dict(color="#aaaaaa", size=label_size, family="Arial Black"),
            showlegend=False, hoverinfo="skip",
        ))


# ── RMA figure builders ───────────────────────────────────────────────────────
def build_state_fig(agg, metric, crop_label, practice, logo_50yr):
    col       = METRIC_COL[metric]
    unit      = METRIC_UNIT[metric]
    fmt       = METRIC_FMT[metric]
    disp_unit = DISPLAY_UNIT[metric]
    df = agg.copy()
    df["StateName"] = df["State"].map(ABBR_TO_NAME)

    title_text = (
        f"{crop_label} — {metric} | Practice: {practice}"
        f"<br><sup>Map labels in {disp_unit}</sup>"
    )
    fig = px.choropleth(
        df, locations="State", locationmode="USA-states", color=col,
        scope="usa", color_continuous_scale=COLOR_SCALE[metric],
        hover_name="StateName",
        hover_data={col: f":{fmt}", "State": False},
        labels={col: f"{metric} ({unit})"},
    )
    fig.update_layout(
        **_base_layout(title_text), height=520,
        geo=dict(showlakes=False, bgcolor=DARK, landcolor=LAND, showland=True, showframe=False),
        coloraxis_colorbar=dict(
            title=dict(text=f"{metric}<br>({unit})", font=dict(color=TEXT)),
            tickfont=dict(color=TEXT),
        ),
    )
    lons, lats, texts = [], [], []
    for _, row in df.iterrows():
        label = format_state_label(row[col], metric)
        if label and row["State"] in STATE_CENTROIDS:
            lon, lat = STATE_CENTROIDS[row["State"]]
            lons.append(lon); lats.append(lat); texts.append(label)
    if lons:
        fig.add_trace(go.Scattergeo(
            lon=lons, lat=lats, text=texts, mode="text", geo="geo",
            textfont=dict(color="#cccccc", size=11, family="Arial Black"),
            showlegend=False, hoverinfo="skip",
        ))
    _add_logo(fig, logo_50yr, size=0.30, opacity=1.0)
    return fig


def build_county_fig(agg, geo, fips_lk, centroids, state, metric, crop_label, practice, logo_50yr):
    col   = METRIC_COL[metric]
    unit  = METRIC_UNIT[metric]
    fmt   = METRIC_FMT[metric]
    sfips = STATE_FIPS.get(state)
    if sfips is None:
        return None

    df = agg.copy()
    df["fips"] = df["County"].apply(lambda c: resolve_fips(state, c, fips_lk))
    df = df.dropna(subset=["fips"])

    state_geo = get_state_geojson(geo, sfips)  # cached
    state_name = ABBR_TO_NAME.get(state, state)
    all_fips   = [f["properties"]["STATE"] + f["properties"]["COUNTY"]
                  for f in state_geo["features"]]

    z_vals = df[col].tolist()
    z_min  = df[col].min() if z_vals else 0
    z_max  = df[col].max() if z_vals else 1
    if z_min == z_max:
        z_min = 0

    county_line = dict(color="#3d5248", width=0.8)
    fig = go.Figure()
    fig.add_trace(go.Choropleth(
        geojson=state_geo, featureidkey="id",
        locations=all_fips, z=[0] * len(all_fips),
        colorscale=[[0, PANEL], [1, PANEL]], showscale=False,
        marker=dict(line=county_line), hoverinfo="skip",
    ))
    fig.add_trace(go.Choropleth(
        geojson=state_geo, featureidkey="id",
        locations=df["fips"].tolist(), z=z_vals,
        colorscale=COLOR_SCALE[metric], zmin=z_min, zmax=z_max,
        colorbar=dict(
            title=dict(text=f"{metric}<br>({unit})", font=dict(color=TEXT)),
            tickfont=dict(color=TEXT),
        ),
        marker=dict(line=county_line),
        text=df["County"].tolist(),
        hovertemplate=f"%{{text}}: %{{z:{fmt}}}<extra></extra>",
    ))

    disp_unit  = DISPLAY_UNIT[metric]
    title_text = (
        f"{crop_label} — {metric} | {state_name} Counties | Practice: {practice}"
        f"<br><sup>Map labels in {disp_unit}</sup>"
    )
    fig.update_geos(fitbounds="locations", visible=False,
                    bgcolor=DARK, landcolor=LAND)
    fig.update_layout(**_base_layout(title_text), height=620)
    _add_logo(fig, logo_50yr, size=0.15, opacity=1.0, x=0.99, y=0.03, yanchor="bottom")
    _place_labels(fig, df["fips"].tolist(), df[col].tolist(), centroids, metric)
    return fig


def build_ranking_chart(agg, metric, state):
    col        = METRIC_COL[metric]
    state_name = ABBR_TO_NAME.get(state, state)
    divisor    = DISPLAY_DIVISOR[metric]
    disp_unit  = DISPLAY_UNIT[metric]

    ranked   = agg.dropna(subset=[col]).sort_values(col, ascending=True)
    raw_avg  = ranked[col].mean()
    x_vals   = ranked[col] / divisor
    avg_disp = raw_avg / divisor

    colors = [ACCENT if v >= raw_avg else "#e05252" for v in ranked[col]]
    fmt    = ".1f" if metric == "Yield" else ",.2f"
    labels = [f"{v:{fmt}}" for v in x_vals]

    fig = go.Figure(go.Bar(
        x=x_vals, y=ranked["County"], orientation="h",
        marker_color=colors, marker_line_width=0,
        text=labels, textposition="outside",
        textfont=dict(color=TEXT, size=8), cliponaxis=False,
        hovertemplate=f"%{{y}}: %{{x:{fmt}}} {disp_unit}<extra></extra>",
    ))
    fig.add_vline(
        x=avg_disp, line_color="#f5a623", line_width=1.5, line_dash="dash",
        annotation_text=f"  Avg: {avg_disp:{fmt}} {disp_unit}",
        annotation_position="top left",
        annotation_font=dict(color="#f5a623", size=10),
    )
    fig.update_layout(
        paper_bgcolor=DARK, plot_bgcolor=SURFACE,
        font=dict(color=TEXT, family="Arial"),
        title=dict(text=f"{state_name} County Rankings — {metric}", font=dict(size=14, color=ACCENT)),
        height=max(380, len(ranked) * 22 + 80),
        margin=dict(l=10, r=90, t=50, b=20), bargap=0.18,
        xaxis=dict(title=f"{metric} ({disp_unit})", gridcolor=BORDER,
                   tickfont=dict(color=MUTED), title_font=dict(color=MUTED), zeroline=False),
        yaxis=dict(gridcolor=BORDER, tickfont=dict(color=TEXT, size=9), automargin=True),
    )
    return fig


# ── NASS figure builders ──────────────────────────────────────────────────────
def build_nass_state_fig(state_vdf, crop, year, metric, change_view, logo_50yr):
    """state_vdf has columns [State, Value] — pre-computed by get_nass_view_data."""
    cfg        = _nass_view_cfg(metric, change_view)
    view_label = metric if change_view == "Absolute" else f"{change_view} — {metric}"
    agg        = state_vdf.copy()
    agg["StateName"] = agg["State"].map(ABBR_TO_NAME)

    title_text = (
        f"NASS {year} {crop} — {view_label}"
        f"<br><sup>Map labels in {cfg['label_unit']}</sup>"
    )

    px_kwargs = {}
    if cfg["diverging"] and not agg["Value"].empty:
        abs_max = max(float(agg["Value"].abs().max()), 1.0)
        px_kwargs["range_color"]               = [-abs_max, abs_max]
        px_kwargs["color_continuous_midpoint"] = 0.0

    fig = px.choropleth(
        agg, locations="State", locationmode="USA-states", color="Value",
        scope="usa", color_continuous_scale=cfg["cscale"],
        hover_name="StateName",
        hover_data={"Value": cfg["hover_fmt"], "State": False},
        labels={"Value": view_label},
        **px_kwargs,
    )
    fig.update_layout(
        **_base_layout(title_text), height=520,
        geo=dict(showlakes=False, bgcolor=DARK, landcolor=LAND, showland=True, showframe=False),
        coloraxis_colorbar=dict(
            title=dict(text=cfg["clabel"], font=dict(color=TEXT)),
            tickfont=dict(color=TEXT),
        ),
    )
    lons, lats, texts = [], [], []
    for _, row in agg.iterrows():
        label = cfg["label_fn"](row["Value"])
        if label and row["State"] in STATE_CENTROIDS:
            lon, lat = STATE_CENTROIDS[row["State"]]
            lons.append(lon); lats.append(lat); texts.append(label)
    if lons:
        fig.add_trace(go.Scattergeo(
            lon=lons, lat=lats, text=texts, mode="text", geo="geo",
            textfont=dict(color="#cccccc", size=11, family="Arial Black"),
            showlegend=False, hoverinfo="skip",
        ))
    _add_logo(fig, logo_50yr, size=0.30, opacity=1.0)
    return fig


def build_nass_county_fig(county_vdf, geo, state, crop, year, metric, change_view, logo_50yr, centroids, fips_lk):
    """county_vdf has columns [State, County, Value] — pre-computed by get_nass_view_data."""
    cfg        = _nass_view_cfg(metric, change_view)
    view_label = metric if change_view == "Absolute" else f"{change_view} — {metric}"
    state_df   = county_vdf[county_vdf["State"] == state].copy()
    if state_df.empty:
        return None

    sfips = STATE_FIPS_ALL.get(state)
    if sfips is None:
        return None

    # Resolve FIPS from county names via GeoJSON lookup — guarantees the codes
    # match the feature "id" that Plotly uses to locate choropleth polygons.
    state_df["fips"] = state_df["County"].apply(lambda c: resolve_fips(state, c, fips_lk))
    state_df = state_df.dropna(subset=["fips"])
    if state_df.empty:
        return None

    state_geo = get_state_geojson(geo, sfips)  # cached
    all_fips  = [f["properties"]["STATE"] + f["properties"]["COUNTY"]
                 for f in state_geo["features"]]

    z_vals = state_df["Value"].tolist()
    if cfg["diverging"]:
        abs_max = max((abs(v) for v in z_vals if not pd.isna(v)), default=1.0)
        abs_max = max(abs_max, 1.0)
        z_min, z_max = -abs_max, abs_max
    else:
        z_min = min(z_vals) if z_vals else 0
        z_max = max(z_vals) if z_vals else 1
        if z_min == z_max:
            z_min = 0

    state_name = ABBR_TO_NAME.get(state, state)
    title_text = (
        f"NASS {year} {crop} — {view_label} | {state_name} Counties"
        f"<br><sup>Map labels in {cfg['label_unit']}</sup>"
    )

    county_line = dict(color="#3d5248", width=0.8)
    fig = go.Figure()
    fig.add_trace(go.Choropleth(
        geojson=state_geo, featureidkey="id",
        locations=all_fips, z=[0] * len(all_fips),
        colorscale=[[0, PANEL], [1, PANEL]], showscale=False,
        marker=dict(line=county_line), hoverinfo="skip",
    ))
    fig.add_trace(go.Choropleth(
        geojson=state_geo, featureidkey="id",
        locations=state_df["fips"].tolist(), z=z_vals,
        colorscale=cfg["cscale"], zmin=z_min, zmax=z_max,
        colorbar=dict(
            title=dict(text=cfg["clabel"], font=dict(color=TEXT)),
            tickfont=dict(color=TEXT),
        ),
        marker=dict(line=county_line),
        text=state_df["County"].tolist(),
        hovertemplate=f"%{{text}}: %{{z{cfg['hover_fmt']}}}{cfg['hover_sfx']}<extra></extra>",
    ))
    fig.update_geos(fitbounds="locations", visible=False, bgcolor=DARK, landcolor=LAND)
    fig.update_layout(**_base_layout(title_text), height=620)
    _add_logo(fig, logo_50yr, size=0.15, opacity=1.0, x=0.99, y=0.03, yanchor="bottom")
    _place_labels(fig, state_df["fips"].tolist(), state_df["Value"].tolist(),
                  centroids, cfg["label_fn"])
    return fig


def build_nass_ranking_chart(ranked_df, state, crop, year, metric, change_view):
    """ranked_df: DataFrame with [County, Value] pre-filtered to a single state."""
    cfg        = _nass_view_cfg(metric, change_view)
    view_label = metric if change_view == "Absolute" else f"{change_view} — {metric}"
    state_name = ABBR_TO_NAME.get(state, state)
    ranked     = ranked_df.dropna(subset=["Value"]).sort_values("Value", ascending=True)
    if ranked.empty:
        return go.Figure()

    raw_avg  = ranked["Value"].mean()
    x_vals   = ranked["Value"] / cfg["rank_div"]
    avg_disp = raw_avg / cfg["rank_div"]
    fmt      = cfg["rank_fmt"]

    colors = [ACCENT if v >= raw_avg else "#e05252" for v in ranked["Value"]]
    labels = [f"{v:{fmt}}" for v in x_vals]

    fig = go.Figure(go.Bar(
        x=x_vals, y=ranked["County"], orientation="h",
        marker_color=colors, marker_line_width=0,
        text=labels, textposition="outside",
        textfont=dict(color=TEXT, size=8), cliponaxis=False,
        hovertemplate=f"%{{y}}: %{{x:{fmt}}} {cfg['rank_unit']}<extra></extra>",
    ))
    fig.add_vline(
        x=avg_disp, line_color="#f5a623", line_width=1.5, line_dash="dash",
        annotation_text=f"  Avg: {avg_disp:{fmt}} {cfg['rank_unit']}",
        annotation_position="top left",
        annotation_font=dict(color="#f5a623", size=10),
    )
    fig.update_layout(
        paper_bgcolor=DARK, plot_bgcolor=SURFACE,
        font=dict(color=TEXT, family="Arial"),
        title=dict(
            text=f"{state_name} County Rankings — {crop} {view_label} (NASS {year})",
            font=dict(size=14, color=ACCENT),
        ),
        height=max(380, len(ranked) * 22 + 80),
        margin=dict(l=10, r=90, t=50, b=20), bargap=0.18,
        xaxis=dict(
            title=f"{view_label} ({cfg['rank_unit']})", gridcolor=BORDER,
            tickfont=dict(color=MUTED), title_font=dict(color=MUTED),
            zeroline=True, zerolinecolor=MUTED,
        ),
        yaxis=dict(gridcolor=BORDER, tickfont=dict(color=TEXT, size=9), automargin=True),
    )
    return fig


# ── Cached county figure wrappers ─────────────────────────────────────────────
# Cache key = hashable args only.  _geo / _centroids / _logo / _fips_lk are
# excluded (underscore prefix) so large objects aren't hashed into the key.

@st.cache_data(show_spinner=False)
def cached_nass_county_fig(state: str, crop: str, year: int,
                            metric: str, change_view: str,
                            comp_year: int, cache_ver: str,
                            _geo, _centroids, _logo_50yr, _fips_lk):
    county_vdf, _ = get_nass_view_data(
        crop, year, metric, change_view,
        comp_year if comp_year > 0 else None,
    )
    return build_nass_county_fig(
        county_vdf, _geo, state, crop, year, metric, change_view, _logo_50yr, _centroids, _fips_lk
    )


@st.cache_data(show_spinner=False)
def cached_rma_county_fig(state: str, crop: str, metric: str,
                           practice: str, wheat_type, cache_ver: str,
                           _geo, _fips_lk, _centroids, _logo_50yr):
    rma_data = load_data()
    if crop not in rma_data:
        return None
    df = rma_data[crop].copy()
    if crop == "Wheat" and wheat_type:
        df = df[df["Type"] == wheat_type]
    agg        = agg_data(df[df["State"] == state], practice, metric, ["County"])
    crop_label = f"Wheat — {wheat_type}" if crop == "Wheat" else crop
    return build_county_fig(agg, _geo, _fips_lk, _centroids, state,
                            metric, crop_label, practice, _logo_50yr)


# ── App ────────────────────────────────────────────────────────────────────────
def main():
    st.markdown(
        f"""
        <style>
        .stApp {{ background-color: {DARK}; color: {TEXT}; }}
        [data-testid="stSidebar"] {{ background-color: {PANEL}; border-right: 1px solid {BORDER}; }}
        .block-container {{ padding-top: 1rem; max-width: 1400px; }}
        h1, h2, h3 {{ color: {ACCENT} !important; letter-spacing: 0.02em; }}
        p, label, .stCaption, [data-testid="stCaptionContainer"] {{ color: {MUTED} !important; }}
        [data-testid="stSelectbox"] label {{ color: {MUTED} !important; font-size: 0.8rem; }}
        div[data-baseweb="select"] > div {{
            background-color: {PANEL} !important; border-color: {BORDER} !important; color: {TEXT} !important;
        }}
        div[data-baseweb="popover"] * {{ background-color: {PANEL} !important; color: {TEXT} !important; }}
        [data-testid="metric-container"] {{
            background-color: {PANEL}; border: 1px solid {BORDER};
            border-radius: 8px; padding: 12px 16px;
        }}
        [data-testid="stMetricValue"] {{ color: {ACCENT} !important; font-size: 1.35rem; font-weight: 700; }}
        [data-testid="stMetricLabel"] {{
            color: {MUTED} !important; font-size: 0.78rem;
            text-transform: uppercase; letter-spacing: 0.06em;
        }}
        [data-testid="stExpander"] {{ background-color: {PANEL}; border: 1px solid {BORDER}; border-radius: 6px; }}
        [data-testid="stDataFrame"] {{ background-color: {PANEL}; }}
        hr {{ border-color: {BORDER}; }}
        [data-testid="stSpinner"] p {{ color: {MUTED} !important; }}
        .stTabs [data-baseweb="tab-list"] {{ background-color: {PANEL}; border-radius: 6px 6px 0 0; gap: 4px; }}
        .stTabs [data-baseweb="tab"] {{ color: {MUTED}; font-size: 0.92rem; padding: 8px 20px; }}
        .stTabs [aria-selected="true"] {{ color: {ACCENT} !important; border-bottom: 2px solid {ACCENT} !important; }}
        /* NASS view toggle — style horizontal radio as a button group */
        div[data-testid="stRadio"] > label {{ display: none; }}
        div[data-testid="stRadio"] > div[role="radiogroup"] {{
            display: flex; flex-direction: row; gap: 6px; flex-wrap: wrap;
        }}
        div[data-testid="stRadio"] > div[role="radiogroup"] > label {{
            background-color: {PANEL} !important;
            border: 1px solid {BORDER} !important;
            border-radius: 6px !important;
            padding: 6px 16px !important;
            cursor: pointer !important;
            color: {MUTED} !important;
            font-size: 0.84rem !important;
            font-weight: 500 !important;
            margin: 0 !important;
        }}
        div[data-testid="stRadio"] > div[role="radiogroup"] > label:has(input:checked) {{
            border-color: {ACCENT} !important;
            color: {ACCENT} !important;
            background-color: {SURFACE} !important;
        }}
        div[data-testid="stRadio"] > div[role="radiogroup"] > label > div:first-child {{
            display: none !important;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.title("USDA County Production Dashboard")
    st.markdown(
        f"""<p style='color:{MUTED};font-size:0.80rem;margin-top:-10px;margin-bottom:6px;line-height:1.6;'>
        ℹ️ <b style='color:{TEXT};'>NASS tab</b>: USDA survey-based final production figures (post-harvest, all acres). &nbsp;|&nbsp;
        <b style='color:{TEXT};'>RMA tab</b>: Estimated production for federally insured acres (insured acres × projected yield) — figures will differ from NASS.
        </p>""",
        unsafe_allow_html=True,
    )

    # Load base resources shared across both tabs
    with st.spinner("Loading..."):
        rma_data   = load_data()
        geo        = load_geojson()
        fips_lk    = build_fips_lookup(geo)
        centroids  = build_centroid_lookup(geo)
        logo_50yr  = load_logo(LOGO_50YR)
        logo_trans = load_logo(LOGO_TRANS)
        logo_full  = load_logo(LOGO_FULL)

    st.markdown(
        f"""<style>
        [data-testid="stHeader"] {{
            background-image: url('{logo_full}');
            background-repeat: no-repeat;
            background-position: right 90px center;
            background-size: auto 68%;
        }}
        </style>""",
        unsafe_allow_html=True,
    )

    tab_nass, tab_rma, tab_about = st.tabs(["🌾  NASS Production", "📋  RMA", "📖  About the Data"])

    # ══════════════════════════════════════════════════════════════════════════
    # NASS TAB
    # ══════════════════════════════════════════════════════════════════════════
    with tab_nass:
        if "nass_sel_state" not in st.session_state:
            st.session_state.nass_sel_state = None

        # Row 1 — Crop, Year, State drill-down, Refresh
        nc1, nc2, nc3, nc4 = st.columns([1, 0.75, 1.8, 0.55])
        with nc1:
            nass_crop = st.selectbox("Crop", list(NASS_CROP_STAT_PARAMS.keys()), key="nass_crop")
        with nc2:
            nass_year = st.selectbox("Year", NASS_YEARS, index=0, key="nass_year")
        with nc4:
            st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
            if st.button("🔄 Refresh", use_container_width=True, key="nass_refresh"):
                st.cache_data.clear()
                st.rerun()

        # Row 2 — Metric selector
        st.markdown(
            f"<p style='color:{MUTED};font-size:0.78rem;margin:4px 0 2px 0;'>Metric</p>",
            unsafe_allow_html=True,
        )
        nass_metric = st.radio(
            "Metric", NASS_METRICS, horizontal=True,
            key="nass_metric", label_visibility="collapsed",
        )

        # Row 3 — Change view selector
        st.markdown(
            f"<p style='color:{MUTED};font-size:0.78rem;margin:4px 0 2px 0;'>View</p>",
            unsafe_allow_html=True,
        )
        nass_change = st.radio(
            "Change View", NASS_CHANGE_OPTS, horizontal=True,
            key="nass_change", label_visibility="collapsed",
        )

        # Row 4 (conditional) — Compare-year picker
        if nass_change == "vs Selected Year":
            cy_col, _ = st.columns([0.55, 2.6])
            with cy_col:
                avail_comp   = [y for y in NASS_YEARS if y != nass_year]
                nass_comp_yr = st.selectbox("Compare to Year", avail_comp, key="nass_comp_year")
        else:
            nass_comp_yr = None

        stat_type = _METRIC_TO_STAT[nass_metric]

        with st.spinner(f"Loading NASS {nass_year} {nass_crop} {nass_metric}..."):
            nass_df = load_nass_county(nass_crop, nass_year, _CACHE_VERSION)   # county data — for drill-down & coverage KPI
            # Pre-warm 2023 benchmark county count
            load_nass_county(nass_crop, _NASS_BENCHMARK_YEAR, _CACHE_VERSION)
            # Pre-warm comparison years for county and state data
            if nass_change == "vs Prior Year":
                _load_for_metric(nass_crop, nass_year - 1, stat_type)
                load_nass_state(nass_crop, nass_year - 1, stat_type, _CACHE_VERSION)
            elif nass_change == "vs Selected Year" and nass_comp_yr:
                _load_for_metric(nass_crop, nass_comp_yr, stat_type)
                load_nass_state(nass_crop, nass_comp_yr, stat_type, _CACHE_VERSION)
            elif nass_change == "vs 3-Yr Avg":
                for _y in [nass_year - 1, nass_year - 2, nass_year - 3]:
                    if _y >= 2022:
                        _load_for_metric(nass_crop, _y, stat_type)
                        load_nass_state(nass_crop, _y, stat_type, _CACHE_VERSION)
            # County-level data for the county map
            county_vdf, _ = get_nass_view_data(
                nass_crop, nass_year, nass_metric, nass_change, nass_comp_yr
            )
            # Official state-level data for the state choropleth map
            _st_cur = load_nass_state(nass_crop, nass_year, stat_type, _CACHE_VERSION)
            if nass_change == "Absolute" or _st_cur.empty:
                state_vdf = _st_cur
            elif nass_change == "vs Prior Year":
                state_vdf = _state_pct_change(
                    _st_cur,
                    load_nass_state(nass_crop, nass_year - 1, stat_type, _CACHE_VERSION),
                )
            elif nass_change == "vs Selected Year" and nass_comp_yr:
                state_vdf = _state_pct_change(
                    _st_cur,
                    load_nass_state(nass_crop, nass_comp_yr, stat_type, _CACHE_VERSION),
                )
            else:  # vs 3-Yr Avg
                _sy = [y for y in [nass_year - 1, nass_year - 2, nass_year - 3] if y >= 2022]
                state_vdf = _state_pct_change_avg(
                    _st_cur,
                    [load_nass_state(nass_crop, y, stat_type, _CACHE_VERSION) for y in _sy],
                )

        if nass_df.empty:
            st.warning(
                f"No NASS {nass_year} county-level production data returned for {nass_crop}. "
                "The data may not yet be published or the API parameters may need adjustment."
            )
        else:
            states_avail_nass = sorted(nass_df["State"].unique())
            with nc3:
                state_opts_nass = ["— US Overview —"] + [
                    f"{a}  —  {ABBR_TO_NAME.get(a, a)}" for a in states_avail_nass
                ]
                default_nass = 0
                if st.session_state.nass_sel_state:
                    try:
                        default_nass = states_avail_nass.index(st.session_state.nass_sel_state) + 1
                    except ValueError:
                        default_nass = 0
                nass_sel = st.selectbox(
                    "State Drill-Down", state_opts_nass,
                    index=default_nass, key="nass_state_dd"
                )
                st.session_state.nass_sel_state = (
                    None if nass_sel.startswith("—") else nass_sel[:2]
                )

            # ── Summary metrics ───────────────────────────────────────────────
            sel_st = st.session_state.nass_sel_state

            # Official state-level totals — used for KPI (state map already uses state_vdf)
            _kpi_state = _st_cur if sel_st is None else _st_cur[_st_cur["State"] == sel_st]

            # County coverage % — current year vs 2023 benchmark
            _bench_df  = load_nass_county(nass_crop, _NASS_BENCHMARK_YEAR, _CACHE_VERSION)
            if sel_st is None:
                _bench_n = len(_bench_df)
                _curr_n  = len(nass_df)
            else:
                _bench_n = len(_bench_df[_bench_df["State"] == sel_st])
                _curr_n  = len(nass_df[nass_df["State"] == sel_st])
            _pct_rep = _curr_n / _bench_n * 100 if _bench_n > 0 else 0.0

            scope_v = county_vdf if sel_st is None else county_vdf[county_vdf["State"] == sel_st]

            def _official_kpi_str(df_state):
                """Format the official KPI value from state-level NASS data."""
                if nass_metric == "Yield (bu/ac)":
                    v = df_state["Value"].mean()
                    return f"{v:.1f} bu/ac" if not pd.isna(v) else "—"
                if nass_metric in ("Planted Acres", "Harvested Acres", "Prevent Plant Acres"):
                    v = df_state["Value"].sum()
                    return f"{v/1e6:.1f}M ac"
                # Production
                v = df_state["Value"].sum()
                return f"{v/1e9:.2f}B bu" if v >= 1e9 else f"{v/1e6:.1f}M bu"

            if nass_change == "Absolute":
                nm1, nm2, nm3, nm4 = st.columns(4)
                nm1.metric(f"{nass_year} {nass_metric}", _official_kpi_str(_kpi_state))
                nm2.metric("County Coverage",
                           f"{_pct_rep:.0f}%",
                           help=f"{_curr_n:,} counties reporting vs {_bench_n:,} in {_NASS_BENCHMARK_YEAR} benchmark")
                nm3.metric("Counties Reporting", f"{_curr_n:,}")
                nm4.metric("States in Data", f"{_st_cur['State'].nunique():,}")
            else:
                nm1, nm2, nm3, nm4 = st.columns(4)
                nm1.metric(f"{nass_year} {nass_metric}", _official_kpi_str(_kpi_state))
                nm2.metric("County Coverage", f"{_pct_rep:.0f}%",
                           help=f"{_curr_n:,} of {_bench_n:,} counties reporting")
                valid_v  = scope_v["Value"].dropna()
                avg_chg  = valid_v.mean() if not valid_v.empty else float("nan")
                improved = int((valid_v > 0).sum())
                declined = int((valid_v < 0).sum())
                nm3.metric("Counties Above Prior", f"{improved:,} ▲")
                nm4.metric("Counties Below Prior", f"{declined:,} ▼")

            # ── Map ───────────────────────────────────────────────────────────
            if sel_st is None:
                if state_vdf.empty:
                    st.info(
                        "No comparison data available for the selected view and year range. "
                        "Try selecting a different year or view."
                    )
                else:
                    nass_fig = build_nass_state_fig(
                        state_vdf, nass_crop, nass_year, nass_metric, nass_change, logo_50yr
                    )
                    nass_fig.update_layout(dragmode=False)
                    st.plotly_chart(nass_fig, use_container_width=True,
                                    key="nass_state_map",
                                    config={"scrollZoom": False, "displayModeBar": False})
                    st.caption("Use the State Drill-Down dropdown above to view county detail.")

            else:
                nass_state = sel_st
                state_df   = nass_df[nass_df["State"] == nass_state].copy()

                if st.button("← Back to US Map", key="nass_back_btn"):
                    st.session_state.nass_sel_state = None
                    st.rerun()

                if state_df.empty or state_df["Production"].sum() == 0:
                    st.warning(
                        f"No NASS {nass_year} county data available for "
                        f"{ABBR_TO_NAME.get(nass_state, nass_state)}. "
                        "This crop may not be produced in this state or data has not been published."
                    )
                else:
                    with st.spinner(
                        f"Building {ABBR_TO_NAME.get(nass_state, nass_state)} county map…"
                    ):
                        nass_county_fig = cached_nass_county_fig(
                            nass_state, nass_crop, nass_year,
                            nass_metric, nass_change,
                            nass_comp_yr if nass_comp_yr else 0,
                            _CACHE_VERSION, geo, centroids, logo_50yr, fips_lk
                        )
                    if nass_county_fig is None:
                        st.info(
                            f"County map not available for "
                            f"{ABBR_TO_NAME.get(nass_state, nass_state)}."
                        )
                    else:
                        st.plotly_chart(nass_county_fig, use_container_width=True,
                                        key="nass_county_map")
                        st.caption("Use ← Back to US Map to return to the national overview.")

                st.markdown(f"<hr style='border-color:{BORDER};margin:8px 0'>",
                            unsafe_allow_html=True)

                state_county_v = county_vdf[county_vdf["State"] == nass_state].copy()
                if not state_county_v.empty:
                    ranking_nass = build_nass_ranking_chart(
                        state_county_v, nass_state, nass_crop, nass_year,
                        nass_metric, nass_change,
                    )
                    st.plotly_chart(ranking_nass, use_container_width=True, key="nass_ranking")

                with st.expander(
                    f"County Data Table — {ABBR_TO_NAME.get(nass_state, nass_state)}",
                    expanded=False,
                ):
                    # Build display table from state_county_v (selected metric/view values)
                    tbl = state_county_v[["County", "Value"]].dropna(subset=["Value"]).sort_values(
                        "Value", ascending=False
                    ).copy()

                    if nass_change != "Absolute":
                        col_label = f"% Change ({nass_metric})"
                        tbl[col_label] = tbl["Value"].apply(
                            lambda v: f"{v:+.1f}%" if pd.notna(v) else "—"
                        )
                    elif nass_metric == "Yield (bu/ac)":
                        col_label = "Yield (bu/ac)"
                        tbl[col_label] = tbl["Value"].apply(
                            lambda v: f"{v:.1f}" if pd.notna(v) else "—"
                        )
                    elif nass_metric in ("Planted Acres", "Harvested Acres", "Prevent Plant Acres"):
                        col_label = nass_metric
                        tbl[col_label] = tbl["Value"].apply(
                            lambda v: f"{v:,.0f}" if pd.notna(v) else "—"
                        )
                    else:
                        col_label = "Production (bu)"
                        tbl[col_label] = tbl["Value"].apply(
                            lambda v: f"{v:,.0f}" if pd.notna(v) else "—"
                        )

                    show_tbl = tbl[["County", col_label]].copy()
                    # Add production as context when showing a non-production metric
                    if nass_metric != "Production (bu)":
                        prod_ctx = state_df[["County", "Production"]].copy()
                        prod_ctx["Production (bu)"] = prod_ctx["Production"].apply(
                            lambda v: f"{v:,.0f}" if pd.notna(v) else "—"
                        )
                        show_tbl = show_tbl.merge(
                            prod_ctx[["County", "Production (bu)"]], on="County", how="left"
                        )
                    st.dataframe(show_tbl, use_container_width=True, hide_index=True)

    # ══════════════════════════════════════════════════════════════════════════
    # RMA TAB
    # ══════════════════════════════════════════════════════════════════════════
    with tab_rma:
        if "rma_sel_state" not in st.session_state:
            st.session_state.rma_sel_state = None

        crops_available = [c for c in ["Corn", "Wheat"] if c in rma_data]
        c1, c2, c3, c4, c5, c6 = st.columns([1, 1.2, 1.2, 1.2, 1.5, 0.6])

        with c1:
            crop = st.selectbox("Crop", crops_available, key="rma_crop")
        with c2:
            metric = st.selectbox("Metric", list(METRIC_COL.keys()), key="rma_metric")
        with c3:
            practice = st.selectbox("Practice", ["All", "Irrigated", "Non-Irrigated"],
                                    key="rma_practice")
        with c4:
            if crop == "Wheat":
                wheat_types = sorted(
                    t for t in rma_data["Wheat"]["Type"].dropna().unique()
                    if "khor" not in t.lower()
                )
                default_wt = next(
                    (i for i, t in enumerate(wheat_types) if "winter" in t.lower()), 0
                )
                wheat_type = st.selectbox("Wheat Type ✱", wheat_types,
                                          index=default_wt, key="rma_wheat_type")
            else:
                wheat_type = None
        with c6:
            st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
            if st.button("🔄 Refresh Data", use_container_width=True, key="rma_refresh"):
                st.cache_data.clear()
                st.rerun()

        df = rma_data[crop].copy()
        if crop == "Wheat" and wheat_type:
            df = df[df["Type"] == wheat_type]
        crop_label = f"Wheat — {wheat_type}" if crop == "Wheat" else crop

        with c5:
            states_avail = sorted(df["State"].unique())
            state_opts   = ["— US Overview —"] + [
                f"{a}  —  {ABBR_TO_NAME.get(a, a)}" for a in states_avail
            ]
            default_idx = 0
            if st.session_state.rma_sel_state:
                try:
                    default_idx = states_avail.index(st.session_state.rma_sel_state) + 1
                except ValueError:
                    default_idx = 0
            sel = st.selectbox("State Drill-Down", state_opts,
                               index=default_idx, key="rma_state_dd")
            st.session_state.rma_sel_state = (
                None if sel.startswith("—") else sel[:2]
            )

        col  = METRIC_COL[metric]
        unit = METRIC_UNIT[metric]
        fmt  = METRIC_FMT[metric]

        scope_df = filter_practice(df, practice)
        if st.session_state.rma_sel_state:
            scope_df = scope_df[scope_df["State"] == st.session_state.rma_sel_state]

        if metric == "Yield":
            p = scope_df["Reported Production"].sum()
            a = scope_df["Reported Production Acres"].sum()
            summary_val = p / a if a > 0 else 0.0
        else:
            summary_val = scope_df[col].sum()

        m1, m2, m3 = st.columns(3)
        lbl = "Avg Yield" if metric == "Yield" else f"Total {metric}"
        m1.metric(lbl, f"{summary_val:{fmt}} {unit}")
        m2.metric("Counties", f"{scope_df[['State','County']].drop_duplicates().shape[0]:,}")
        m3.metric("States",   f"{scope_df['State'].nunique():,}")

        if st.session_state.rma_sel_state is None:
            agg = agg_data(df, practice, metric, ["State"])
            fig = build_state_fig(agg, metric, crop_label, practice, logo_50yr)
            fig.update_layout(dragmode=False)
            st.plotly_chart(fig, use_container_width=True, key="rma_state_map",
                            config={"scrollZoom": False, "displayModeBar": False})
            st.caption("Use the State Drill-Down dropdown above to view county detail.")

        else:
            state = st.session_state.rma_sel_state
            agg   = agg_data(df[df["State"] == state], practice, metric, ["County"])

            if st.button("← Back to US Map", key="rma_back_btn"):
                st.session_state.rma_sel_state = None
                st.rerun()

            if agg.empty or agg[col].sum() == 0:
                st.warning(f"No data for {ABBR_TO_NAME.get(state, state)} with selected filters.")
            else:
                with st.spinner(f"Building {ABBR_TO_NAME.get(state, state)} county map…"):
                    fig = cached_rma_county_fig(
                        state, crop, metric, practice, wheat_type, _CACHE_VERSION,
                        geo, fips_lk, centroids, logo_50yr
                    )
                if fig is None:
                    st.info(f"County map not available for {ABBR_TO_NAME.get(state, state)}.")
                else:
                    st.plotly_chart(fig, use_container_width=True, key="rma_county_map")
                    st.caption("Use ← Back to US Map to return to the national overview.")

            state_name = ABBR_TO_NAME.get(state, state)
            st.markdown(f"<hr style='border-color:{BORDER};margin:8px 0'>",
                        unsafe_allow_html=True)
            ranking_fig = build_ranking_chart(agg, metric, state)
            st.plotly_chart(ranking_fig, use_container_width=True, key="rma_ranking_chart")

            with st.expander(f"County Data Table — {state_name}", expanded=False):
                disp = agg.sort_values(col, ascending=False).copy()
                disp.columns = ["County", f"{metric} ({unit})"]
                disp[f"{metric} ({unit})"] = disp[f"{metric} ({unit})"].apply(
                    lambda v: f"{v:,.1f}" if pd.notna(v) else "—"
                )
                st.dataframe(disp, use_container_width=True, hide_index=True)


    # ══════════════════════════════════════════════════════════════════════════
    # ABOUT THE DATA TAB
    # ══════════════════════════════════════════════════════════════════════════
    with tab_about:
        st.markdown(
            f"<h3 style='color:{ACCENT};margin-top:0;margin-bottom:4px;'>"
            "Understanding NASS vs RMA Production Data</h3>",
            unsafe_allow_html=True,
        )
        st.markdown(
            f"<p style='color:{MUTED};font-size:0.86rem;margin-top:0;margin-bottom:20px;'>"
            "Two federal agencies measure crop production using fundamentally different methodologies — "
            "knowing which lens you're looking through matters when interpreting the numbers.</p>",
            unsafe_allow_html=True,
        )

        # ── Side-by-side methodology cards ────────────────────────────────────
        ab_l, ab_r = st.columns(2, gap="large")
        with ab_l:
            st.markdown(
                f"""<div style='background:{PANEL};border:1px solid {BORDER};
                border-radius:8px;padding:20px 24px;height:100%;'>
                <h4 style='color:{ACCENT};margin-top:0;margin-bottom:10px;font-size:1.05rem;'>
                🌾&nbsp; NASS — Survey-Based Total Production</h4>
                <p style='color:{TEXT};font-size:0.91rem;line-height:1.75;margin:0 0 10px 0;'>
                The <b>USDA National Agricultural Statistics Service (NASS)</b> produces
                official crop estimates by surveying thousands of farm operators,
                grain elevators, and agribusinesses nationwide.
                </p>
                <ul style='color:{MUTED};font-size:0.88rem;line-height:1.85;
                margin:0;padding-left:18px;'>
                <li>Covers <b style='color:{TEXT};'>all planted acres</b> — insured and uninsured alike</li>
                <li>Final estimates released post-harvest (typically November)</li>
                <li>Represents <b style='color:{TEXT};'>actual harvested production</b> for the crop year</li>
                <li>The definitive benchmark for U.S. crop supply &amp; demand</li>
                <li>Feeds directly into the USDA WASDE monthly balance sheets</li>
                </ul>
                </div>""",
                unsafe_allow_html=True,
            )
        with ab_r:
            st.markdown(
                f"""<div style='background:{PANEL};border:1px solid {BORDER};
                border-radius:8px;padding:20px 24px;height:100%;'>
                <h4 style='color:{ACCENT};margin-top:0;margin-bottom:10px;font-size:1.05rem;'>
                📋&nbsp; RMA — Insurance Policy-Based Production</h4>
                <p style='color:{TEXT};font-size:0.91rem;line-height:1.75;margin:0 0 10px 0;'>
                The <b>USDA Risk Management Agency (RMA)</b> collects data through the
                federal crop insurance program, drawing from individual policy records
                filed by insured farmers and processed by approved insurance providers.
                </p>
                <ul style='color:{MUTED};font-size:0.88rem;line-height:1.85;
                margin:0;padding-left:18px;'>
                <li>Covers only <b style='color:{TEXT};'>federally insured acres</b></li>
                <li>Based on Actual Production History (APH) and policy-reported yields</li>
                <li>Breakdowns available by practice (Irrigated / Non-Irrigated)</li>
                <li>Published annually via the RMA Summary of Business</li>
                <li>Reflects <b style='color:{TEXT};'>insured-sector production</b>, not the full market</li>
                </ul>
                </div>""",
                unsafe_allow_html=True,
            )

        st.markdown("<div style='height:18px'></div>", unsafe_allow_html=True)

        # ── Why numbers differ callout ─────────────────────────────────────────
        st.markdown(
            f"""<div style='background:{SURFACE};border-left:3px solid {ACCENT};
            border-radius:0 6px 6px 0;padding:14px 20px;margin-bottom:24px;'>
            <h4 style='color:{ACCENT};margin:0 0 8px 0;font-size:1rem;'>
            ⚡&nbsp; Why RMA Figures Trend Higher Than NASS</h4>
            <p style='color:{TEXT};font-size:0.90rem;line-height:1.80;margin:0;'>
            Even for the same county and crop year, RMA reported production often runs
            <b>above the NASS estimate</b>. Several structural factors drive this gap:
            </p>
            <ul style='color:{MUTED};font-size:0.88rem;line-height:1.85;
            margin:8px 0 0 0;padding-left:18px;'>
            <li><b style='color:{TEXT};'>Larger, higher-yielding operations</b> — farms that purchase
            crop insurance tend to be bigger and more productive than the average uninsured acre.</li>
            <li><b style='color:{TEXT};'>APH yield smoothing</b> — RMA uses Actual Production History
            (a multi-year rolling average) rather than any single harvest year, which dampens
            downside and can overstate expected production in weak years.</li>
            <li><b style='color:{TEXT};'>Irrigated acre weighting</b> — insured acres skew toward
            irrigated, higher-yield ground, lifting the portfolio average above the county mean.</li>
            <li><b style='color:{TEXT};'>Convergence over time</b> — as participation has climbed to
            ~92–93 % of planted acres, the RMA dataset increasingly mirrors the full universe,
            and the gap with NASS has narrowed materially.</li>
            </ul>
            </div>""",
            unsafe_allow_html=True,
        )

        # ── Participation rate chart ───────────────────────────────────────────
        st.markdown(
            f"<h4 style='color:{ACCENT};margin-bottom:2px;'>"
            "From Optional to Universal: U.S. Crop Insurance Participation</h4>",
            unsafe_allow_html=True,
        )
        st.markdown(
            f"<p style='color:{MUTED};font-size:0.82rem;margin-top:0;margin-bottom:10px;'>"
            "Share of planted acres covered under federally reinsured crop insurance, 2000 – 2025.</p>",
            unsafe_allow_html=True,
        )

        # Approximate annual participation rates derived from
        # USDA RMA Summary of Business ÷ USDA NASS planted acres
        _PART_YEARS = list(range(2000, 2026))
        _CORN_PCT = [
            66, 67, 70, 72, 73, 74, 75, 77,   # 2000-2007
            80, 82, 83, 84, 85, 87, 89, 90,   # 2008-2015
            90, 91, 91, 91, 91, 91, 92, 92, 92, 92,  # 2016-2025
        ]
        _SOY_PCT = [
            64, 65, 67, 69, 71, 72, 73, 75,   # 2000-2007
            77, 79, 80, 82, 83, 85, 87, 89,   # 2008-2015
            90, 91, 91, 92, 92, 92, 93, 93, 93, 93,  # 2016-2025
        ]

        part_fig = go.Figure()

        # Shaded area beneath lines for visual weight
        part_fig.add_trace(go.Scatter(
            x=_PART_YEARS + _PART_YEARS[::-1],
            y=_CORN_PCT + [55] * len(_CORN_PCT),
            fill="toself",
            fillcolor=f"rgba(74,222,128,0.06)",
            line=dict(width=0),
            showlegend=False, hoverinfo="skip",
        ))
        part_fig.add_trace(go.Scatter(
            x=_PART_YEARS + _PART_YEARS[::-1],
            y=_SOY_PCT + [55] * len(_SOY_PCT),
            fill="toself",
            fillcolor=f"rgba(96,165,250,0.06)",
            line=dict(width=0),
            showlegend=False, hoverinfo="skip",
        ))

        # Corn line
        part_fig.add_trace(go.Scatter(
            x=_PART_YEARS, y=_CORN_PCT,
            mode="lines+markers",
            name="Corn",
            line=dict(color=ACCENT, width=2.5),
            marker=dict(size=5, color=ACCENT),
            hovertemplate="%{x}: <b>%{y}%</b><extra>Corn</extra>",
        ))

        # Soybeans line
        part_fig.add_trace(go.Scatter(
            x=_PART_YEARS, y=_SOY_PCT,
            mode="lines+markers",
            name="Soybeans",
            line=dict(color="#60a5fa", width=2.5),
            marker=dict(size=5, color="#60a5fa"),
            hovertemplate="%{x}: <b>%{y}%</b><extra>Soybeans</extra>",
        ))

        # 2008 Farm Bill reference
        part_fig.add_vline(
            x=2008, line_color=BORDER, line_width=1.5, line_dash="dot",
            annotation_text="2008 Farm Bill",
            annotation_position="top right",
            annotation_font=dict(color=MUTED, size=10),
        )
        # 2014 Farm Bill reference
        part_fig.add_vline(
            x=2014, line_color=BORDER, line_width=1.5, line_dash="dot",
            annotation_text="2014 Farm Bill",
            annotation_position="top right",
            annotation_font=dict(color=MUTED, size=10),
        )

        # Endpoint annotations — 2000
        part_fig.add_annotation(
            x=2000, y=66, text="66%", showarrow=False,
            xanchor="right", xshift=-6, yshift=8,
            font=dict(color=ACCENT, size=11, family="Arial Bold"),
        )
        part_fig.add_annotation(
            x=2000, y=64, text="64%", showarrow=False,
            xanchor="right", xshift=-6, yshift=-10,
            font=dict(color="#60a5fa", size=11, family="Arial Bold"),
        )
        # Endpoint annotations — 2025
        part_fig.add_annotation(
            x=2025, y=92, text="92%", showarrow=False,
            xanchor="left", xshift=6, yshift=-10,
            font=dict(color=ACCENT, size=11, family="Arial Bold"),
        )
        part_fig.add_annotation(
            x=2025, y=93, text="93%", showarrow=False,
            xanchor="left", xshift=6, yshift=8,
            font=dict(color="#60a5fa", size=11, family="Arial Bold"),
        )

        part_fig.update_layout(
            paper_bgcolor=DARK, plot_bgcolor=SURFACE,
            font=dict(color=TEXT, family="Arial"),
            margin=dict(l=60, r=70, t=30, b=50),
            height=380,
            xaxis=dict(
                title="Year",
                gridcolor=BORDER,
                tickfont=dict(color=MUTED),
                title_font=dict(color=MUTED),
                dtick=2,
                range=[1998.5, 2026.5],
                zeroline=False,
            ),
            yaxis=dict(
                title="% of Planted Acres Insured",
                gridcolor=BORDER,
                tickfont=dict(color=MUTED),
                title_font=dict(color=MUTED),
                range=[55, 100],
                ticksuffix="%",
                zeroline=False,
            ),
            legend=dict(
                orientation="h", yanchor="bottom", y=1.02,
                xanchor="left", x=0,
                font=dict(color=TEXT, size=11),
                bgcolor="rgba(0,0,0,0)",
            ),
            hovermode="x unified",
        )

        st.plotly_chart(
            part_fig, use_container_width=True,
            config={"displayModeBar": False},
        )

        st.markdown(
            f"<p style='color:{MUTED};font-size:0.78rem;margin-top:2px;'>"
            "Source: USDA RMA Summary of Business &nbsp;+&nbsp; USDA NASS Quick Stats · Annual, 2000–2025. "
            "Participation rates represent federally insured planted acres as a share of total NASS planted acres. "
            "Historical series are approximate; endpoint values (2000 &amp; 2025) are per RMA/NASS published data.</p>",
            unsafe_allow_html=True,
        )

        # ── Quick-reference comparison table ─────────────────────────────────
        st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
        st.markdown(
            f"<h4 style='color:{ACCENT};margin-bottom:8px;'>At a Glance: Key Differences</h4>",
            unsafe_allow_html=True,
        )
        cmp_df = pd.DataFrame({
            "": ["Data collection method", "Acre coverage", "Yield basis",
                 "Practice detail", "Primary use", "Publication timing"],
            "NASS": [
                "Farm & elevator surveys",
                "All planted acres (insured + uninsured)",
                "Actual harvested yield",
                "No (all-practice aggregate)",
                "Supply/demand balance sheets (WASDE)",
                "Monthly estimates; final in November",
            ],
            "RMA": [
                "Individual insurance policy records",
                "Federally insured acres only (~92–93%)",
                "APH / policy-reported yield (multi-year avg)",
                "Yes — Irrigated vs Non-Irrigated",
                "Crop insurance pricing & indemnity analysis",
                "Annual Summary of Business (spring/summer)",
            ],
        })
        st.dataframe(cmp_df, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
