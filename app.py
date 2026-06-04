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
from PIL import Image
import geopandas as gpd
from shapely.geometry import shape

_HERE = Path(__file__).parent
st.set_page_config(
    page_title="USDA County & ASD Production Dashboard",
    page_icon=Image.open(_HERE / "assets" / "Transparent Smal logo.png"),
    layout="wide",
)
_CACHE_VERSION = "v11"  # bump to invalidate all @st.cache_data on deploy

# ── Paths ─────────────────────────────────────────────────────────────────────
HERE       = Path(__file__).parent
DATA_PATH  = HERE / "data" / "2025 RMA Production Data.xlsx"
LOGO_50YR  = HERE / "assets" / "50 Year logo JSA.png"
LOGO_TRANS = HERE / "assets" / "Transparent Smal logo.png"
LOGO_FULL  = HERE / "assets" / "logo-full.png"

# ── NASS API ───────────────────────────────────────────────────────────────────
NASS_API_KEY  = "9A6D1EB8-4D94-3221-BA0C-ADD4533EA0C1"
NASS_BASE_URL = "https://quickstats.nass.usda.gov/api/api_GET/"
NASS_YEARS             = list(range(2025, 2014, -1))   # 2025 → 2015
_NASS_BENCHMARK_YEAR   = 2023   # most-complete county year — used for % reporting KPI

# Metrics available in the NASS tab
NASS_METRICS     = ["Production (bu)", "Planted Acres", "Harvested Acres",
                    "Yield (bu/ac)", "Prevent Plant Acres"]
NASS_CHANGE_OPTS = ["Current Year", "vs Prior Year", "vs Selected Year", "vs 3-Yr Avg"]

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
              "county_ansi", "prodn_practice_desc", "asd_desc", "asd_code", "Value"]
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

    # Include district fields when present in the API response
    extra = [c for c in ["asd_desc", "asd_code"] if c in df.columns]
    return df[key + ["Production"] + extra].reset_index(drop=True)


# ── Tier-1 county estimation (production, planted acres, harvested acres, yield)
# Missing counties are estimated via district-multiplier applied to olympic-avg
# historical shares, then scaled to reconcile with NASS OTHER COUNTIES totals.

@st.cache_data(show_spinner=False)
def _load_county_raw_for_est(crop: str, state: str, year: int,
                              stat_type: str, cache_ver: str) -> pd.DataFrame:
    """County rows INCLUDING the OTHER COUNTIES catch-all for any stat type.
    Used only for the estimation pipeline."""
    params = {
        "key": NASS_API_KEY, "source_desc": "SURVEY", "sector_desc": "CROPS",
        "agg_level_desc": "COUNTY", "domain_desc": "TOTAL",
        "state_alpha": state, "year": str(year), "format": "JSON",
    }
    params.update(NASS_STAT_BASE[stat_type])
    params.update(NASS_CROP_STAT_PARAMS[crop][stat_type])
    try:
        url = NASS_BASE_URL + "?" + urllib.parse.urlencode(params)
        with urllib.request.urlopen(url, timeout=45) as r:
            raw = json.load(r)
    except Exception:
        return pd.DataFrame()
    records = raw.get("data", [])
    if not records:
        return pd.DataFrame()
    df = pd.DataFrame(records)
    needed = ["state_alpha", "county_name", "county_ansi", "state_fips_code",
              "asd_desc", "asd_code", "prodn_practice_desc", "Value"]
    df = df[[c for c in needed if c in df.columns]].copy()
    df["Value_num"] = pd.to_numeric(
        df["Value"].str.replace(",", "", regex=False).str.strip(), errors="coerce"
    ).fillna(0)
    df["fips"]    = df["state_fips_code"].str.zfill(2) + df["county_ansi"].str.zfill(3)
    df["State"]   = df["state_alpha"].str.strip()
    df["County"]  = df["county_name"].str.strip().str.title()
    df["ANSI"]    = df["county_ansi"].str.strip().str.zfill(3)
    df["District"]     = df["asd_desc"].str.strip().str.title() if "asd_desc"  in df.columns else ""
    df["DistrictCode"] = df["asd_code"].str.strip()              if "asd_code"  in df.columns else ""
    return df


@st.cache_data(show_spinner=False)
def _build_hist_shares(crop: str, state: str, stat_type: str,
                       history_years: tuple, cache_ver: str) -> pd.DataFrame:
    """Olympic-average historical share (county_val / state_total) per county.
    Works for any stat_type (production / planted / harvested / yield).
    history_years is a tuple so it's hashable as a cache key."""
    rows = []
    for yr in history_years:
        st_total = load_nass_state(crop, yr, stat_type, cache_ver)
        if st_total.empty or "State" not in st_total.columns:
            continue
        st_row = st_total[st_total["State"] == state]
        if st_row.empty:
            continue
        st_val = float(st_row["Value"].iloc[0])
        if st_val <= 0:
            continue
        df = _load_county_raw_for_est(crop, state, yr, stat_type, cache_ver)
        if df.empty:
            continue
        named = df[
            (~df["ANSI"].isin(["998", "000", "999"])) &
            (~df["County"].str.lower().str.startswith("other", na=False))
        ].copy()
        if "prodn_practice_desc" in named.columns:
            named = named[named["prodn_practice_desc"] == "ALL PRODUCTION PRACTICES"]
        named = named.loc[named.groupby("fips")["Value_num"].idxmax()]
        for _, row in named.iterrows():
            rows.append({
                "fips":         row["fips"],
                "County":       row["County"],
                "District":     row.get("District", ""),
                "DistrictCode": row.get("DistrictCode", ""),
                "year":         yr,
                "share":        row["Value_num"] / st_val,
            })
    if not rows:
        return pd.DataFrame()
    hist = pd.DataFrame(rows)

    def _olympic(vals):
        v = sorted(vals)
        return float(np.mean(v[1:-1] if len(v) >= 4 else v))

    result = (
        hist.groupby(["fips", "County"])
        .apply(lambda g: pd.Series({
            "hist_share":    _olympic(g["share"].tolist()),
            "n_obs":         len(g),
            "District":      g["District"].mode().iloc[0]     if len(g) > 0 else "",
            "DistrictCode":  g["DistrictCode"].mode().iloc[0] if len(g) > 0 else "",
        }), include_groups=False)
        .reset_index()
    )
    return result


@st.cache_data(show_spinner=False)
def _build_adj(state_fips: str, cache_ver: str, _geo: dict) -> dict:
    """County adjacency map {fips: [neighbor_fips,...]} via geopandas."""
    feats = [f for f in _geo["features"] if f["properties"]["STATE"] == state_fips]
    if not feats:
        return {}
    rows, geoms = [], []
    for f in feats:
        try:
            geoms.append(shape(f["geometry"]).buffer(0.001))
            rows.append({"fips": f["properties"]["STATE"] + f["properties"]["COUNTY"]})
        except Exception:
            continue
    gdf = gpd.GeoDataFrame(rows, geometry=geoms, crs="EPSG:4326")
    joined = gpd.sjoin(gdf, gdf, how="left", predicate="intersects")
    joined = joined[joined["fips_left"] != joined["fips_right"]]
    adj: dict = {}
    for _, row in joined.iterrows():
        adj.setdefault(row["fips_left"], []).append(row["fips_right"])
    return adj


@st.cache_data(show_spinner=False)
def get_completed_county_data(crop: str, state: str, year: int,
                               stat_type: str, cache_ver: str) -> pd.DataFrame:
    """
    Generic Tier-1 estimation for any additive stat type:
      production / planted / harvested

    For yield: share = county_yield / state_yield; district multiplier applied
    but NO OTHER-COUNTIES reconciliation (yield is a ratio, not additive).

    Returns [fips, County, District, DistrictCode, Value, is_estimated]
    where  named + estimated  sums (or means for yield) match the state total.
    """
    raw = _load_county_raw_for_est(crop, state, year, stat_type, cache_ver)
    if raw.empty:
        return pd.DataFrame()

    state_total_df = load_nass_state(crop, year, stat_type, cache_ver)
    if state_total_df.empty or "State" not in state_total_df.columns:
        return pd.DataFrame()
    st_row = state_total_df[state_total_df["State"] == state]
    if st_row.empty:
        return pd.DataFrame()
    state_total = float(st_row["Value"].iloc[0])

    is_other  = raw["County"].str.lower().str.startswith("other", na=False)
    other_tot = raw[is_other]["Value_num"].sum()
    named     = raw[~is_other & ~raw["ANSI"].isin(["998","000","999"])].copy()
    if "prodn_practice_desc" in named.columns:
        named = named[named["prodn_practice_desc"] == "ALL PRODUCTION PRACTICES"]
    named = named.loc[named.groupby("fips")["Value_num"].idxmax()]
    named["is_estimated"] = False

    # For yield: OTHER COUNTIES doesn't exist as an additive bucket.
    # Fall back to raw if nothing is missing.
    is_yield = (stat_type == "yield")
    if not is_yield and other_tot <= 0:
        named["Value"] = named["Value_num"]
        return named[["fips","County","District","DistrictCode","Value","is_estimated"]]
    if is_yield and named.empty:
        named["Value"] = named["Value_num"]
        return named[["fips","County","District","DistrictCode","Value","is_estimated"]]

    hist_yrs    = tuple(range(max(year - 8, 2015), year))
    hist_shares = _build_hist_shares(crop, state, stat_type, hist_yrs, cache_ver)
    if hist_shares.empty:
        named["Value"] = named["Value_num"]
        return named[["fips","County","District","DistrictCode","Value","is_estimated"]]

    named_fips = set(named["fips"])
    missing = hist_shares[~hist_shares["fips"].isin(named_fips)].copy()
    if missing.empty:
        named["Value"] = named["Value_num"]
        return named[["fips","County","District","DistrictCode","Value","is_estimated"]]

    # Deviation ratios from reported counties
    hs_map = dict(zip(hist_shares["fips"], hist_shares["hist_share"]))
    dev = {
        row["fips"]: row["Value_num"] / (hs_map[row["fips"]] * state_total)
        for _, row in named.iterrows()
        if row["fips"] in hs_map and hs_map[row["fips"]] > 0
    }
    named_r    = named.copy(); named_r["ratio"] = named_r["fips"].map(dev)
    dist_ratio = named_r.dropna(subset=["ratio"]).groupby("District")["ratio"].mean().to_dict()
    state_ratio= float(np.mean(list(dev.values()))) if dev else 1.0

    missing = missing.copy()
    missing["mult"]    = missing["District"].map(lambda d: dist_ratio.get(d, state_ratio))
    missing["raw_est"] = missing["hist_share"] * state_total * missing["mult"]

    if not is_yield:
        # Scale estimates to sum exactly to OTHER COUNTIES
        raw_sum = missing["raw_est"].sum()
        scale   = (other_tot / raw_sum) if raw_sum > 0 else 1.0
        missing["Value"] = missing["raw_est"] * scale
    else:
        # Yield: no reconciliation — use raw estimate directly
        missing["Value"] = missing["raw_est"]

    missing["is_estimated"] = True
    named["Value"] = named["Value_num"]

    return pd.concat([
        named[["fips","County","District","DistrictCode","Value","is_estimated"]],
        missing[["fips","County","District","DistrictCode","Value","is_estimated"]],
    ], ignore_index=True)


def get_completed_county_production(crop: str, state: str, year: int,
                                    cache_ver: str) -> pd.DataFrame:
    """Backward-compat wrapper — returns Production column instead of Value."""
    df = get_completed_county_data(crop, state, year, "production", cache_ver)
    if df.empty:
        return df
    return df.rename(columns={"Value": "Production"})


# ── ASD district boundary helpers ────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_boundary_fips_map(crop: str, state_fips: str,
                           cache_ver: str, _geo: dict) -> dict:
    """Walk back from 2023 to build a complete fips→(District, DistrictCode) map.
    Uses historical years so the district boundaries are full regardless of
    how many counties have reported in the currently-selected year.
    Returns { fips_5: (district_name, district_code) }
    """
    all_geo_fips = {
        f["properties"]["STATE"] + f["properties"]["COUNTY"]
        for f in _geo["features"]
        if f["properties"]["STATE"] == state_fips
    }
    inv_fips = {v: k for k, v in STATE_FIPS_ALL.items()}
    state_alpha = inv_fips.get(state_fips, "")

    fips_map: dict = {}
    for yr in range(2023, 2016, -1):
        if len(fips_map) >= len(all_geo_fips):
            break
        df = load_nass_county(crop, yr, cache_ver)
        if df.empty or "State" not in df.columns:
            continue
        df_s = df[df["State"] == state_alpha]
        if df_s.empty or "asd_desc" not in df_s.columns:
            continue
        for _, row in df_s.iterrows():
            fips = row.get("fips", "")
            dist = str(row.get("asd_desc", "")).strip().title()
            code = str(row.get("asd_code", "")).strip()
            if fips and dist and dist.lower() not in ("", "nan") and fips not in fips_map:
                fips_map[fips] = (dist, code)
    return fips_map


@st.cache_data(show_spinner=False)
def build_nass_district_gdf(state_fips: str, cache_ver: str,
                            _fips_map: dict, _geo: dict) -> gpd.GeoDataFrame:
    """Dissolve county polygons → district polygons using the static fips_map."""
    rows, geoms = [], []
    for feat in _geo["features"]:
        if feat["properties"]["STATE"] != state_fips:
            continue
        fips  = feat["properties"]["STATE"] + feat["properties"]["COUNTY"]
        entry = _fips_map.get(fips)
        if entry is None:
            continue
        dist_name, dist_code = entry
        try:
            geoms.append(shape(feat["geometry"]))
            rows.append({"District": dist_name, "DistrictCode": dist_code})
        except Exception:
            continue

    if not rows:
        return gpd.GeoDataFrame()

    gdf = gpd.GeoDataFrame(rows, geometry=geoms, crs="EPSG:4326")
    dissolved = (
        gdf.dissolve(by=["District", "DistrictCode"])
        .reset_index()[["District", "DistrictCode", "geometry"]]
    )
    dissolved["centroid_lon"] = dissolved.geometry.centroid.x
    dissolved["centroid_lat"] = dissolved.geometry.centroid.y
    return dissolved.sort_values("DistrictCode").reset_index(drop=True)


def get_nass_district_view_data(crop: str, year: int, metric: str,
                                 change_view: str, fips_map: dict,
                                 state: str, comp_year=None,
                                 _geo=None) -> pd.DataFrame:
    """
    Returns [District, Value] aggregated from county-level NASS data,
    computing proper district-level % change for non-absolute views.
    For Current Year, uses Tier-1 estimated county data for all metrics so
    district totals reconcile to the official state total.
    """
    stat_type = _METRIC_TO_STAT[metric]
    _estimable = {"production", "planted", "harvested", "yield"}
    use_est    = (stat_type in _estimable and change_view == "Current Year")

    def _load_state(yr, use_estimation=False):
        if use_estimation:
            df = get_completed_county_data(crop, state, yr, stat_type, _CACHE_VERSION)
            if df.empty:
                return pd.DataFrame(columns=["fips", "Value"])
        else:
            df = _load_for_metric(crop, yr, stat_type)
            if df.empty or "State" not in df.columns:
                return pd.DataFrame(columns=["fips", "Value"])
            df = df[df["State"] == state].copy()
        df["District"] = df["fips"].map(lambda f: fips_map.get(f, (None, None))[0])
        return df.dropna(subset=["District"])

    def _agg(df):
        if metric == "Yield (bu/ac)":
            return df.groupby("District")["Value"].mean()
        return df.groupby("District")["Value"].sum()

    cur = _agg(_load_state(year, use_estimation=use_est))

    if change_view == "Current Year" or cur.empty:
        return cur.reset_index()

    def _pct(cur_s, base_s):
        return ((cur_s - base_s) / base_s.replace(0, np.nan) * 100).dropna()

    if change_view == "vs Prior Year":
        base = _agg(_load_state(year - 1))
    elif change_view == "vs Selected Year" and comp_year:
        base = _agg(_load_state(comp_year))
    else:  # vs 3-Yr Avg — average each prior year's DISTRICT totals
        prior = [y for y in [year-1, year-2, year-3] if y >= 2015]
        frames = [_agg(_load_state(y)) for y in prior]
        frames = [f for f in frames if not f.empty]
        if not frames:
            return cur.reset_index()
        base = pd.concat(frames, axis=1).mean(axis=1)

    result = _pct(cur, base)
    return result.reset_index().rename(columns={0: "Value"})


def build_nass_district_fig(dist_view_df: pd.DataFrame,
                             dist_raw_df,          # absolute metric values for labels
                             dist_gdf: gpd.GeoDataFrame,
                             state: str, crop: str, year: int,
                             metric: str, change_view: str,
                             logo_50yr, geo=None,
                             estimated_districts: set = None) -> go.Figure:
    """Build a state choropleth coloured by ASD district with county outlines,
    bold district boundaries, and labels showing metric value + % change."""
    if dist_gdf.empty or dist_view_df.empty:
        return None

    cfg        = _nass_view_cfg(metric, change_view)
    view_label = metric if change_view == "Current Year" else f"{change_view} — {metric}"
    state_name = ABBR_TO_NAME.get(state, state)

    dist_val_map = dict(zip(dist_view_df["District"], dist_view_df["Value"]))
    # Raw (absolute) values for the metric label — always show regardless of view
    raw_map = (
        dict(zip(dist_raw_df["District"], dist_raw_df["Value"]))
        if dist_raw_df is not None and not dist_raw_df.empty else {}
    )
    # Always use the absolute-metric config for formatting raw values so we
    # never accidentally apply the % change formatter to a bushel/acre figure
    abs_cfg    = _nass_view_cfg(metric, "Current Year")
    state_fips = STATE_FIPS_ALL.get(state)

    # Convert dissolved GeoDataFrame to GeoJSON for Plotly
    dist_geojson = json.loads(dist_gdf.to_json())

    districts = dist_gdf["District"].tolist()
    z_vals    = [dist_val_map.get(d, 0) for d in districts]

    # Hover: always show raw metric value + % change in comparison modes
    def _hover(d, z):
        rv   = raw_map.get(d)
        rv_s = abs_cfg["label_fn"](rv) if rv is not None else ""
        if change_view != "Current Year" and z:
            return f"<b>{d}</b><br>{rv_s}<br>{'+'if z>=0 else ''}{z:.1f}%"
        return f"<b>{d}</b><br>{rv_s}"

    hover_texts = [_hover(d, z) for d, z in zip(districts, z_vals)]

    _z_pos = [v for v in z_vals if v > 0]
    if cfg["diverging"]:
        _abs = max((abs(v) for v in z_vals), default=1.0)
        z_min, z_max = -max(_abs, 1.0), max(_abs, 1.0)
    else:
        z_min = 0
        z_max = max(_z_pos) if _z_pos else 1

    fig = go.Figure()

    # Layer 1 — district fill (coloured polygons)
    fig.add_trace(go.Choropleth(
        geojson=dist_geojson,
        featureidkey="properties.District",
        locations=districts, z=z_vals,
        colorscale=cfg["cscale"], zmin=z_min, zmax=z_max,
        colorbar=dict(
            title=dict(text=cfg["clabel"], font=dict(color=TEXT)),
            tickfont=dict(color=TEXT),
        ),
        marker=dict(line=dict(color=BORDER, width=0.3)),
        text=hover_texts,
        hovertemplate="%{text}<extra></extra>",
    ))

    # Layer 2 — county outlines (transparent fill, faint grid lines within districts)
    if geo is not None and state_fips:
        _county_feats = [f for f in geo["features"]
                         if f["properties"]["STATE"] == state_fips]
        _county_fips  = [f["properties"]["STATE"] + f["properties"]["COUNTY"]
                         for f in _county_feats]
        _county_geo   = {"type": "FeatureCollection", "features": _county_feats}
        fig.add_trace(go.Choropleth(
            geojson=_county_geo, featureidkey="id",
            locations=_county_fips, z=[0] * len(_county_fips),
            colorscale=[[0, "rgba(0,0,0,0)"], [1, "rgba(0,0,0,0)"]],
            showscale=False,
            marker=dict(line=dict(color="rgba(90,90,90,0.55)", width=0.5)),
            hoverinfo="skip",
        ))

    # Layer 3 — district boundary lines (bold white, drawn over county grid)
    all_lons, all_lats   = [], []
    lbl_lons, lbl_lats, lbl_texts = [], [], []
    for _, row in dist_gdf.iterrows():
        geom  = row.geometry
        polys = [geom] if geom.geom_type == "Polygon" else list(geom.geoms)
        for poly in polys:
            xs, ys = poly.exterior.coords.xy
            all_lons.extend(list(xs) + [None])
            all_lats.extend(list(ys) + [None])

        _dn  = row["District"]
        _dv  = dist_val_map.get(_dn)
        _rv  = raw_map.get(_dn)
        _rv_s = abs_cfg["label_fn"](_rv) if _rv is not None else ""

        # Label line 1: district name
        # Label line 2: metric value (always)
        # Add "Est" marker to districts with estimated counties
        _est_flag = (
            "<br>Est" if (estimated_districts and _dn in estimated_districts) else ""
        )
        if change_view != "Current Year" and _dv is not None:
            _sign = "+" if _dv >= 0 else ""
            _lbl  = f"{_dn.upper()}{_est_flag}<br>{_sign}{_dv:.1f}%"
        else:
            _lbl  = f"{_dn.upper()}{_est_flag}<br>{_rv_s}"

        lbl_lons.append(row["centroid_lon"])
        lbl_lats.append(row["centroid_lat"])
        lbl_texts.append(_lbl)

    fig.add_trace(go.Scattergeo(
        lon=all_lons, lat=all_lats, mode="lines",
        line=dict(color="white", width=1.8),
        showlegend=False, hoverinfo="skip",
    ))
    # Layer 4 — district name + value labels
    fig.add_trace(go.Scattergeo(
        lon=lbl_lons, lat=lbl_lats, mode="text",
        text=lbl_texts,
        textfont=dict(color="black", size=9, family="Arial Black"),
        showlegend=False, hoverinfo="skip",
    ))

    fig.update_geos(fitbounds="locations", visible=False,
                    bgcolor=DARK, landcolor=LAND, showframe=False)
    _layout = _base_layout(
        f"NASS {year} {crop} — {view_label} | {state_name} AG Districts"
    )
    _layout.update(height=620, geo=dict(showlakes=False))
    fig.update_layout(**_layout)
    _add_logo(fig, logo_50yr, size=0.15, opacity=1.0, x=0.99, y=0.03, yanchor="bottom")
    return fig


@st.cache_data(show_spinner=False)
def cached_nass_district_fig(state: str, crop: str, year: int,
                              metric: str, change_view: str,
                              comp_year: int, cache_ver: str,
                              _geo, _logo_50yr, _fips_map):
    # View data (absolute or % change depending on change_view)
    # Pass _geo so production "Current Year" can use Tier-2 estimated counties
    dist_view_df = get_nass_district_view_data(
        crop, year, metric, change_view, _fips_map, state,
        comp_year if comp_year > 0 else None,
        _geo=_geo,
    )
    # Raw absolute values for labels — always "Current Year" absolute
    dist_raw_df = get_nass_district_view_data(
        crop, year, metric, "Current Year", _fips_map, state, None,
        _geo=_geo,
    ) if change_view != "Current Year" else dist_view_df

    dist_gdf = build_nass_district_gdf(
        STATE_FIPS_ALL.get(state, ""), cache_ver, _fips_map, _geo
    )
    # Build set of districts that contain at least one estimated county
    if metric == "Production (bu)":
        _comp = get_completed_county_production(crop, state, year, cache_ver)
        estimated_districts = (
            set(_comp[_comp["is_estimated"]]["District"].dropna().unique())
            if not _comp.empty else set()
        )
    else:
        estimated_districts = set()

    return build_nass_district_fig(
        dist_view_df, dist_raw_df, dist_gdf,
        state, crop, year, metric, change_view, _logo_50yr, _geo,
        estimated_districts=estimated_districts,
    )


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
    if change_view != "Current Year":
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

    if change_view == "Current Year" or df_cur.empty:
        return _agg_c(df_cur), _agg_s(df_cur)

    def _pct(cur_s, cmp_s):
        return (cur_s - cmp_s) / cmp_s.replace(0, np.nan) * 100

    if change_view == "vs Prior Year":
        df_cmp = _load_for_metric(crop, year - 1, stat_type)
    elif change_view == "vs Selected Year":
        df_cmp = _load_for_metric(crop, comp_year, stat_type) if comp_year else df_cur
    else:  # vs 3-Yr Avg
        prior_years = [y for y in [year - 1, year - 2, year - 3] if y >= 2015]
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
    view_label = metric if change_view == "Current Year" else f"{change_view} — {metric}"
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
    view_label = metric if change_view == "Current Year" else f"{change_view} — {metric}"
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


def build_nass_county_fig_with_est(completed_df: pd.DataFrame, geo, state: str,
                                    crop: str, year: int, metric: str,
                                    logo_50yr, centroids) -> go.Figure:
    """
    County choropleth using completed (reported + Tier-1 estimated) data for
    any metric. Estimated counties get a small 'Est' label and hover suffix.
    """
    if completed_df.empty:
        return None

    sfips = STATE_FIPS_ALL.get(state)
    if not sfips:
        return None

    cfg        = _nass_view_cfg(metric, "Current Year")
    state_name = ABBR_TO_NAME.get(state, state)
    state_geo  = get_state_geojson(geo, sfips)
    all_fips   = [f["properties"]["STATE"] + f["properties"]["COUNTY"]
                  for f in state_geo["features"]]

    fips_val = dict(zip(completed_df["fips"], completed_df["Value"]))
    fips_cty = dict(zip(completed_df["fips"], completed_df["County"]))
    fips_est = dict(zip(completed_df["fips"], completed_df["is_estimated"]))

    z_vals, hover_texts = [], []
    for fips in all_fips:
        val = fips_val.get(fips, 0)
        z_vals.append(val)
        cty = fips_cty.get(fips, fips)
        sfx = " (Est)" if fips_est.get(fips, False) else ""
        hover_texts.append(
            f"{cty}{sfx}: {cfg['label_fn'](val)}{cfg['hover_sfx']}" if val
            else f"{cty}: No data"
        )

    _pos = [v for v in z_vals if v > 0]
    z_min = min(_pos) if _pos else 0
    z_max = max(_pos) if _pos else 1

    county_line = dict(color="#3d5248", width=0.8)
    fig = go.Figure()

    # Background (uncoloured county outlines)
    fig.add_trace(go.Choropleth(
        geojson=state_geo, featureidkey="id",
        locations=all_fips, z=[0] * len(all_fips),
        colorscale=[[0, PANEL], [1, PANEL]], showscale=False,
        marker=dict(line=county_line), hoverinfo="skip",
    ))

    # Production values (reported + estimated, same colour scale)
    fig.add_trace(go.Choropleth(
        geojson=state_geo, featureidkey="id",
        locations=all_fips, z=z_vals,
        colorscale=cfg["cscale"], zmin=z_min, zmax=z_max,
        colorbar=dict(
            title=dict(text=cfg["clabel"], font=dict(color=TEXT)),
            tickfont=dict(color=TEXT),
        ),
        marker=dict(line=county_line),
        text=hover_texts,
        hovertemplate="%{text}<extra></extra>",
    ))

    # "Est" labels on estimated county centroids
    est_fips = [f for f in all_fips if fips_est.get(f, False)]
    if est_fips and centroids:
        _elons = [centroids[f][0] for f in est_fips if f in centroids]
        _elats = [centroids[f][1] for f in est_fips if f in centroids]
        if _elons:
            fig.add_trace(go.Scattergeo(
                lon=_elons, lat=_elats, mode="text",
                text=["Est"] * len(_elons),
                textfont=dict(color="white", size=7, family="Arial Bold"),
                showlegend=False, hoverinfo="skip",
            ))

    title_text = (
        f"NASS {year} {crop} — {metric} | {state_name} Counties"
        f"<br><sup>Map labels in {cfg['label_unit']}  ·  Est = county value estimated</sup>"
    )
    fig.update_geos(fitbounds="locations", visible=False, bgcolor=DARK, landcolor=LAND)
    fig.update_layout(**_base_layout(title_text), height=620)
    _add_logo(fig, logo_50yr, size=0.15, opacity=1.0, x=0.99, y=0.03, yanchor="bottom")

    # Production labels for reported counties only (same as existing county fig)
    rep_fips = [f for f in all_fips if f in fips_val and not fips_est.get(f, False)]
    rep_vals = [fips_val[f] for f in rep_fips]
    _place_labels(fig, rep_fips, rep_vals, centroids, cfg["label_fn"])

    return fig


def build_nass_ranking_chart(ranked_df, state, crop, year, metric, change_view):
    """ranked_df: DataFrame with [County, Value] pre-filtered to a single state."""
    cfg        = _nass_view_cfg(metric, change_view)
    view_label = metric if change_view == "Current Year" else f"{change_view} — {metric}"
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


# One color per ASD district code (10–90)
_ASD_PALETTE = {
    "10": "#4ade80", "20": "#60a5fa", "30": "#f59e0b",
    "40": "#f87171", "50": "#a78bfa", "60": "#34d399",
    "70": "#fb923c", "80": "#38bdf8", "90": "#e879f9",
}


def build_nass_asd_ranking_chart(state_county_v, fips_map, fips_lk,
                                  state, crop, year, metric, change_view):
    """Bar chart of all counties in a state, sorted and coloured by ASD district."""
    cfg        = _nass_view_cfg(metric, change_view)
    view_label = metric if change_view == "Current Year" else f"{change_view} — {metric}"
    state_name = ABBR_TO_NAME.get(state, state)

    df = state_county_v.dropna(subset=["Value"]).copy()
    if df.empty:
        return go.Figure()

    # Attach district info via fips_map
    df["fips"]         = df["County"].apply(lambda c: resolve_fips(state, c, fips_lk))
    df["District"]     = df["fips"].map(lambda f: (fips_map.get(f) or (None,None))[0])
    df["DistrictCode"] = df["fips"].map(lambda f: (fips_map.get(f) or (None,""))[1])

    # Sort: district code ascending, then value ascending within district
    df = df.sort_values(
        ["DistrictCode", "Value"], ascending=[True, True]
    ).reset_index(drop=True)

    df["Color"] = df["DistrictCode"].map(_ASD_PALETTE).fillna(MUTED)

    x_vals   = df["Value"] / cfg["rank_div"]
    raw_avg  = df["Value"].mean()
    avg_disp = raw_avg / cfg["rank_div"]
    fmt      = cfg["rank_fmt"]

    fig = go.Figure(go.Bar(
        x=x_vals, y=df["County"], orientation="h",
        marker_color=df["Color"].tolist(), marker_line_width=0,
        text=[f"{v:{fmt}}" for v in x_vals],
        textposition="outside",
        textfont=dict(color=TEXT, size=8), cliponaxis=False,
        customdata=df["District"].fillna("—").tolist(),
        hovertemplate=(
            f"%{{y}} — %{{customdata}}: %{{x:{fmt}}} {cfg['rank_unit']}<extra></extra>"
        ),
    ))

    # State-average reference line
    fig.add_vline(
        x=avg_disp, line_color="#f5a623", line_width=1.5, line_dash="dash",
        annotation_text=f"  Avg: {avg_disp:{fmt}} {cfg['rank_unit']}",
        annotation_position="top left",
        annotation_font=dict(color="#f5a623", size=10),
    )

    # Horizontal separator lines + district label between groups
    cumulative = 0
    for code in sorted(df["DistrictCode"].dropna().unique()):
        grp = df[df["DistrictCode"] == code]
        if grp.empty:
            continue
        dist_name = grp.iloc[0]["District"] or code
        color     = _ASD_PALETTE.get(code, MUTED)

        if cumulative > 0:
            fig.add_hline(y=cumulative - 0.5,
                          line_color=BORDER, line_width=1.5, line_dash="dot")

        fig.add_annotation(
            x=0, y=cumulative + len(grp) / 2 - 0.5,
            xref="paper", yref="y",
            text=f"<b>{dist_name}</b>",
            font=dict(color=color, size=8),
            showarrow=False, xanchor="right", xshift=-4,
        )
        cumulative += len(grp)

    fig.update_layout(
        paper_bgcolor=DARK, plot_bgcolor=SURFACE,
        font=dict(color=TEXT, family="Arial"),
        title=dict(
            text=f"{state_name} County Rankings by AG District — {crop} {view_label} (NASS {year})",
            font=dict(size=14, color=ACCENT),
        ),
        height=max(400, len(df) * 22 + 80),
        margin=dict(l=110, r=90, t=50, b=20), bargap=0.18,
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
        if "nass_map_view" not in st.session_state:
            st.session_state.nass_map_view = "ASD District"

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
                    if _y >= 2015:
                        _load_for_metric(nass_crop, _y, stat_type)
                        load_nass_state(nass_crop, _y, stat_type, _CACHE_VERSION)
            # County-level data for the county map
            county_vdf, _ = get_nass_view_data(
                nass_crop, nass_year, nass_metric, nass_change, nass_comp_yr
            )
            # Official state-level data for the state choropleth map
            _st_cur = load_nass_state(nass_crop, nass_year, stat_type, _CACHE_VERSION)
            if nass_change == "Current Year" or _st_cur.empty:
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
                _sy = [y for y in [nass_year - 1, nass_year - 2, nass_year - 3] if y >= 2015]
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

            if nass_change == "Current Year":
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
                    nass_fig.update_layout(
                        dragmode=False,
                        geo=dict(
                            projection_type="albers usa",
                            showframe=False,
                            bgcolor=DARK,
                            landcolor=LAND,
                        ),
                    )
                    st.plotly_chart(nass_fig, use_container_width=True,
                                    key="nass_state_map",
                                    config={
                                        "scrollZoom":    False,
                                        "displayModeBar": False,
                                        "doubleClick":   False,
                                    })
                    st.caption("Use the State Drill-Down dropdown above to view county detail.")

            else:
                nass_state = sel_st
                state_df   = nass_df[nass_df["State"] == nass_state].copy()

                # Back button + map-view toggle on the same row
                _back_col, _view_col, _ = st.columns([0.7, 2.2, 2])
                with _back_col:
                    if st.button("← Back", key="nass_back_btn"):
                        st.session_state.nass_sel_state = None
                        st.rerun()
                with _view_col:
                    nass_map_view = st.radio(
                        "Map View",
                        ["ASD District", "County"],
                        horizontal=True,
                        key="nass_map_view",
                        index=0,
                    )

                if state_df.empty or state_df["Production"].sum() == 0:
                    st.warning(
                        f"No NASS {nass_year} county data available for "
                        f"{ABBR_TO_NAME.get(nass_state, nass_state)}. "
                        "This crop may not be produced in this state or data has not been published."
                    )
                else:
                    # Load fips_map always — used by both the ASD map and the
                    # ASD-grouped ranking chart regardless of current map view.
                    _sfips    = STATE_FIPS_ALL.get(nass_state, "")
                    _fips_map = load_boundary_fips_map(
                        nass_crop, _sfips, _CACHE_VERSION, geo
                    )

                    # ── Map ───────────────────────────────────────────────────
                    if nass_map_view == "ASD District":
                        # Current-year absolute district values for the side table
                        # Pass geo so production uses Tier-2 estimated counties
                        _dist_abs  = get_nass_district_view_data(
                            nass_crop, nass_year, nass_metric, "Current Year",
                            _fips_map, nass_state, None, _geo=geo,
                        )
                        # Estimated county count badge — all estimable metrics
                        _est_stat = _METRIC_TO_STAT.get(nass_metric, "")
                        if _est_stat in ("production", "planted", "harvested", "yield"):
                            _comp_data = get_completed_county_data(
                                nass_crop, nass_state, nass_year, _est_stat, _CACHE_VERSION
                            )
                            _n_est = int(_comp_data["is_estimated"].sum()) if not _comp_data.empty else 0
                        else:
                            _comp_data = pd.DataFrame()
                            _n_est = 0
                        _abs_cfg_d = _nass_view_cfg(nass_metric, "Current Year")
                        _is_yield  = nass_metric == "Yield (bu/ac)"

                        # Build name → ASD code lookup from fips_map
                        _name_to_code = {
                            v[0]: v[1] for v in _fips_map.values() if v[0] and v[1]
                        }

                        if not _dist_abs.empty:
                            _dt = _dist_abs.copy()
                            _dt["ASD"] = _dt["District"].map(_name_to_code).fillna("")
                            _dt = _dt.sort_values("ASD").reset_index(drop=True)
                            _dt[nass_metric] = _dt["Value"].apply(_abs_cfg_d["label_fn"])
                            _st_total = (
                                _dt["Value"].mean() if _is_yield
                                else _dt["Value"].sum()
                            )
                            _dt["% of State"] = (
                                (_dt["Value"] / _st_total * 100
                                 ).round(1).astype(str)
                                if not _is_yield and _st_total > 0
                                else "—"
                            )

                            # Per-district estimated county count
                            if not _comp_data.empty and _est_stat in ("production","planted","harvested","yield"):
                                _est_by_dist = (
                                    _comp_data[_comp_data["is_estimated"]]
                                    .groupby("District")["is_estimated"]
                                    .count()
                                    .reset_index()
                                    .rename(columns={"is_estimated": "Est. Counties"})
                                )
                                _dt = _dt.merge(_est_by_dist, on="District", how="left")
                                _dt["Est. Counties"] = _dt["Est. Counties"].fillna(0).astype(int)
                                _dt["Est. Counties"] = _dt["Est. Counties"].apply(
                                    lambda n: str(n) if n > 0 else "—"
                                )
                                _tbl_cols = ["District", "ASD", nass_metric,
                                             "% of State", "Est. Counties"]
                            else:
                                _tbl_cols = ["District", "ASD", nass_metric, "% of State"]

                            _tbl_col, _map_col = st.columns([1, 2.5])
                            with _tbl_col:
                                st.markdown(
                                    f"<p style='color:{MUTED};font-size:0.78rem;"
                                    f"margin:0 0 4px 0;'>"
                                    f"{ABBR_TO_NAME.get(nass_state,nass_state)} "
                                    f"{nass_year} {nass_crop}</p>",
                                    unsafe_allow_html=True,
                                )
                                st.dataframe(
                                    _dt[_tbl_cols],
                                    use_container_width=True, hide_index=True,
                                )
                                if _n_est > 0:
                                    st.caption(
                                        "Est = Some counties not yet final. Production is "
                                        "estimated using each county's historical share of "
                                        "state output, adjusted for current district "
                                        "performance, and scaled to reconcile with the "
                                        "NASS state total."
                                    )
                                _km1, _km2 = st.columns(2)
                                _km1.metric(
                                    "State Avg" if _is_yield else "State Total",
                                    _abs_cfg_d["label_fn"](_st_total),
                                )
                                _km2.metric("Districts", f"{len(_dt)}")
                        else:
                            _map_col = st.container()

                        with _map_col:
                            with st.spinner(
                                f"Building {ABBR_TO_NAME.get(nass_state, nass_state)}"
                                " ASD district map…"
                            ):
                                nass_dist_fig = cached_nass_district_fig(
                                    nass_state, nass_crop, nass_year,
                                    nass_metric, nass_change,
                                    nass_comp_yr if nass_comp_yr else 0,
                                    _CACHE_VERSION, geo, logo_50yr, _fips_map,
                                )
                            if nass_dist_fig is None:
                                st.info(f"ASD district map not available for "
                                        f"{ABBR_TO_NAME.get(nass_state, nass_state)}.")
                            else:
                                st.plotly_chart(nass_dist_fig, use_container_width=True,
                                                key="nass_district_map")
                    else:
                        _cty_est_stat = _METRIC_TO_STAT.get(nass_metric, "")
                        _use_est_county = (
                            _cty_est_stat in ("production","planted","harvested","yield")
                            and nass_change == "Current Year"
                        )
                        with st.spinner(
                            f"Building {ABBR_TO_NAME.get(nass_state, nass_state)} county map…"
                        ):
                            if _use_est_county:
                                _comp_cty = get_completed_county_data(
                                    nass_crop, nass_state, nass_year,
                                    _cty_est_stat, _CACHE_VERSION
                                )
                                nass_county_fig = build_nass_county_fig_with_est(
                                    _comp_cty, geo, nass_state, nass_crop,
                                    nass_year, nass_metric, logo_50yr, centroids,
                                )
                                _cty_n_est = int(_comp_cty["is_estimated"].sum()) \
                                    if not _comp_cty.empty else 0
                            else:
                                nass_county_fig = cached_nass_county_fig(
                                    nass_state, nass_crop, nass_year,
                                    nass_metric, nass_change,
                                    nass_comp_yr if nass_comp_yr else 0,
                                    _CACHE_VERSION, geo, centroids, logo_50yr, fips_lk
                                )
                                _cty_n_est = 0
                        if nass_county_fig is None:
                            st.info(
                                f"County map not available for "
                                f"{ABBR_TO_NAME.get(nass_state, nass_state)}."
                            )
                        else:
                            st.plotly_chart(nass_county_fig, use_container_width=True,
                                            key="nass_county_map")
                            if _cty_n_est > 0:
                                st.caption(
                                    f"Est = {_cty_n_est} counties not yet final. "
                                    "Production is estimated using each county's "
                                    "historical share of state output, adjusted for "
                                    "current district performance, and scaled to "
                                    "reconcile with the NASS state total."
                                )

                    st.markdown(f"<hr style='border-color:{BORDER};margin:8px 0'>",
                                unsafe_allow_html=True)

                    # ── Historical Summary Table ───────────────────────────────
                    _state_full = ABBR_TO_NAME.get(nass_state, nass_state)

                    # Scope selector: State (default), ASD District, or County
                    _hsc1, _hsc2 = st.columns([1.4, 2.6])
                    with _hsc1:
                        _htbl_scope = st.radio(
                            "Table Scope",
                            ["State", "ASD District", "County"],
                            horizontal=True, key="htbl_scope",
                        )
                    with _hsc2:
                        if _htbl_scope == "ASD District":
                            _dist_opts = sorted(
                                set(v[0] for v in _fips_map.values() if v[0])
                            )
                            _htbl_dist = st.selectbox(
                                "District", _dist_opts, key="htbl_district"
                            )
                            _htbl_county = None
                            _scope_label = f"{_state_full} — {_htbl_dist}"
                        elif _htbl_scope == "County":
                            _county_opts = sorted(
                                state_df["County"].dropna().unique()
                            )
                            _htbl_county = st.selectbox(
                                "County", _county_opts, key="htbl_county"
                            )
                            _htbl_dist = None
                            _scope_label = f"{_state_full} — {_htbl_county} County"
                        else:
                            _htbl_dist = _htbl_county = None
                            _scope_label = _state_full

                    st.markdown(
                        f"<p style='color:{MUTED};font-size:0.82rem;font-weight:600;"
                        f"margin:0 0 6px 0;letter-spacing:0.04em;'>"
                        f"📅 HISTORICAL SUMMARY — {_scope_label} | "
                        f"{nass_year - 5}–{nass_year}</p>",
                        unsafe_allow_html=True,
                    )

                    # Rolling 6-year window ending at the selected year
                    _HIST_YEARS = list(range(nass_year - 5, nass_year + 1))
                    _HIST_STATTYPES = ["planted", "harvested", "yield", "production"]
                    _HIST_ROW_LBL   = {
                        "planted":    "Planted Acres (000 ac)",
                        "harvested":  "Harvested Acres (000 ac)",
                        "yield":      "Yield (bu/ac)",
                        "production": "Production (M bu)",
                    }

                    with st.spinner("Loading historical data…"):
                        _hist: dict = {}
                        _tbl_est_years: set = set()   # years where production is estimated
                        for _hyr in _HIST_YEARS:
                            _hist[_hyr] = {}
                            for _hst in _HIST_STATTYPES:
                                try:
                                    if _htbl_scope == "State":
                                        # Official state totals from state-level API
                                        _hdf = load_nass_state(
                                            nass_crop, _hyr, _hst, _CACHE_VERSION
                                        )
                                        _hdf_s = (
                                            _hdf[_hdf["State"] == nass_state]
                                            if not _hdf.empty and "State" in _hdf.columns
                                            else pd.DataFrame()
                                        )
                                        _hist[_hyr][_hst] = (
                                            float(_hdf_s["Value"].iloc[0])
                                            if not _hdf_s.empty else None
                                        )
                                    elif (_htbl_scope == "ASD District"
                                          and _htbl_dist
                                          and _hst in ("production","planted","harvested","yield")):
                                        # Use Tier-1 estimated county data for all metrics
                                        _comp_yr = get_completed_county_data(
                                            nass_crop, nass_state, _hyr, _hst, _CACHE_VERSION
                                        )
                                        if _comp_yr.empty:
                                            _hist[_hyr][_hst] = None
                                            continue
                                        _comp_dist = _comp_yr[
                                            _comp_yr["District"] == _htbl_dist
                                        ]
                                        if _comp_dist.empty:
                                            _hist[_hyr][_hst] = None
                                        else:
                                            _hist[_hyr][_hst] = float(
                                                _comp_dist["Value"].mean()
                                                if _hst == "yield"
                                                else _comp_dist["Value"].sum()
                                            )
                                            if _comp_dist["is_estimated"].any():
                                                _tbl_est_years.add(_hyr)
                                    else:
                                        # County-level data filtered to district or county
                                        _hdf = _load_for_metric(
                                            nass_crop, _hyr, _hst
                                        )
                                        if _hdf.empty or "State" not in _hdf.columns:
                                            _hist[_hyr][_hst] = None
                                            continue
                                        _hdf_s = _hdf[_hdf["State"] == nass_state].copy()
                                        if _htbl_scope == "ASD District" and _htbl_dist:
                                            _hdf_s["_Dist"] = _hdf_s["fips"].map(
                                                lambda f: (_fips_map.get(f) or (None,))[0]
                                            )
                                            _hdf_s = _hdf_s[_hdf_s["_Dist"] == _htbl_dist]
                                        elif _htbl_scope == "County" and _htbl_county:
                                            _hdf_s = _hdf_s[_hdf_s["County"] == _htbl_county]
                                        if _hdf_s.empty:
                                            _hist[_hyr][_hst] = None
                                        elif _hst == "yield":
                                            _hist[_hyr][_hst] = float(_hdf_s["Value"].mean())
                                        else:
                                            _hist[_hyr][_hst] = float(_hdf_s["Value"].sum())
                                except Exception:
                                    _hist[_hyr][_hst] = None

                    def _hfmt(stype, v):
                        if v is None or (isinstance(v, float) and np.isnan(v)):
                            return "—"
                        if stype == "production":   return f"{v/1e6:,.0f}"
                        if stype in ("planted","harvested"): return f"{v/1e3:,.0f}"
                        if stype == "yield":        return f"{v:.1f}"
                        return f"{v:,.0f}"

                    def _hdelta_str(v):
                        if v is None or (isinstance(v, float) and np.isnan(v)):
                            return ""
                        return f" ({'+'if v>=0 else ''}{v:.1f}%)"

                    # Compute deltas using same nass_change view as map
                    _hdelta: dict = {yr: {} for yr in _HIST_YEARS}
                    if nass_change == "vs Prior Year":
                        for _hi, _hyr in enumerate(_HIST_YEARS):
                            if _hi == 0:
                                continue
                            _hprev = _HIST_YEARS[_hi - 1]
                            for _hst in _HIST_STATTYPES:
                                _hc, _hp = _hist[_hyr].get(_hst), _hist[_hprev].get(_hst)
                                _hdelta[_hyr][_hst] = (
                                    (_hc - _hp) / abs(_hp) * 100
                                    if _hc and _hp else None
                                )
                    elif nass_change == "vs 3-Yr Avg":
                        for _hst in _HIST_STATTYPES:
                            # 3-yr avg base = 3 years prior to the selected year
                            _avg_base = [nass_year - 3, nass_year - 2, nass_year - 1]
                            _hvals = [_hist[y].get(_hst) for y in _avg_base
                                      if _hist.get(y, {}).get(_hst)]
                            _havg  = sum(_hvals) / len(_hvals) if _hvals else None
                            for _hyr in _HIST_YEARS:
                                _hc = _hist[_hyr].get(_hst)
                                _hdelta[_hyr][_hst] = (
                                    (_hc - _havg) / abs(_havg) * 100
                                    if _hc and _havg else None
                                )
                    elif nass_change == "vs Selected Year" and nass_comp_yr:
                        for _hyr in _HIST_YEARS:
                            for _hst in _HIST_STATTYPES:
                                _hc   = _hist[_hyr].get(_hst)
                                _hbase= _hist[nass_comp_yr].get(_hst)
                                _hdelta[_hyr][_hst] = (
                                    (_hc - _hbase) / abs(_hbase) * 100
                                    if _hc and _hbase and _hyr != nass_comp_yr else None
                                )

                    _htbl_rows = []
                    for _hst in _HIST_STATTYPES:
                        _hrow = {"": _HIST_ROW_LBL[_hst]}
                        for _hyr in _HIST_YEARS:
                            _raw_s = _hfmt(_hst, _hist[_hyr].get(_hst))
                            _dlt_s = (
                                _hdelta_str(_hdelta[_hyr].get(_hst))
                                if nass_change != "Current Year" else ""
                            )
                            _hrow[str(_hyr)] = f"{_raw_s}{_dlt_s}"
                        _htbl_rows.append(_hrow)

                    _htbl_df = pd.DataFrame(_htbl_rows).set_index("")

                    # Mark estimated year columns with " (Est)"
                    if _tbl_est_years:
                        _htbl_df = _htbl_df.rename(
                            columns={str(y): f"{y} (Est)" for y in _tbl_est_years}
                        )

                    def _htbl_style(val):
                        if "(+" in str(val): return f"color:{ACCENT};font-weight:600"
                        if "(-" in str(val): return "color:#ef4444;font-weight:600"
                        return ""

                    if nass_change != "Current Year":
                        st.dataframe(
                            _htbl_df.style.map(_htbl_style),
                            use_container_width=True,
                        )
                    else:
                        st.dataframe(_htbl_df, use_container_width=True)

                    if _tbl_est_years:
                        st.caption(
                            "Est = Some counties not yet final. Production is estimated "
                            "using each county's historical share of state output, "
                            "adjusted for current district performance, and scaled to "
                            "reconcile with the NASS state total."
                        )

                    st.markdown(f"<hr style='border-color:{BORDER};margin:8px 0'>",
                                unsafe_allow_html=True)

                    # ── Ranking Chart ─────────────────────────────────────────
                    state_county_v = county_vdf[county_vdf["State"] == nass_state].copy()
                    if not state_county_v.empty:
                        if nass_map_view == "ASD District":
                            ranking_nass = build_nass_asd_ranking_chart(
                                state_county_v, _fips_map, fips_lk,
                                nass_state, nass_crop, nass_year,
                                nass_metric, nass_change,
                            )
                        else:
                            ranking_nass = build_nass_ranking_chart(
                                state_county_v, nass_state, nass_crop, nass_year,
                                nass_metric, nass_change,
                            )
                        st.plotly_chart(ranking_nass, use_container_width=True,
                                        key="nass_ranking")

                    # ── County Data Table ─────────────────────────────────────
                    with st.expander(
                        f"County Data Table — {ABBR_TO_NAME.get(nass_state, nass_state)}",
                        expanded=False,
                    ):
                        tbl = state_county_v[["County", "Value"]].dropna(
                            subset=["Value"]
                        ).sort_values("Value", ascending=False).copy()

                        if nass_change != "Current Year":
                            col_label = f"% Change ({nass_metric})"
                            tbl[col_label] = tbl["Value"].apply(
                                lambda v: f"{v:+.1f}%" if pd.notna(v) else "—"
                            )
                        elif nass_metric == "Yield (bu/ac)":
                            col_label = "Yield (bu/ac)"
                            tbl[col_label] = tbl["Value"].apply(
                                lambda v: f"{v:.1f}" if pd.notna(v) else "—"
                            )
                        elif nass_metric in ("Planted Acres", "Harvested Acres",
                                             "Prevent Plant Acres"):
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
                        if nass_metric != "Production (bu)":
                            prod_ctx = state_df[["County", "Production"]].copy()
                            prod_ctx["Production (bu)"] = prod_ctx["Production"].apply(
                                lambda v: f"{v:,.0f}" if pd.notna(v) else "—"
                            )
                            show_tbl = show_tbl.merge(
                                prod_ctx[["County", "Production (bu)"]],
                                on="County", how="left",
                            )
                        st.dataframe(show_tbl, use_container_width=True, hide_index=True)

    # ══════════════════════════════════════════════════════════════════════════
    # RMA TAB
    # ══════════════════════════════════════════════════════════════════════════
    with tab_rma:
        if "rma_sel_state" not in st.session_state:
            st.session_state.rma_sel_state = None

        crops_available = [c for c in ["Corn", "Soybeans", "Wheat"] if c in rma_data]
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
