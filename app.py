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

# ── Paths ─────────────────────────────────────────────────────────────────────
HERE       = Path(__file__).parent
DATA_PATH  = HERE / "data" / "2025 RMA Production Data.xlsx"
LOGO_50YR  = HERE / "assets" / "50 Year logo JSA.png"
LOGO_TRANS = HERE / "assets" / "Transparent Smal logo.png"
LOGO_FULL  = HERE / "assets" / "logo-full.png"

# ── NASS API ───────────────────────────────────────────────────────────────────
NASS_API_KEY  = "9A6D1EB8-4D94-3221-BA0C-ADD4533EA0C1"
NASS_BASE_URL = "https://quickstats.nass.usda.gov/api/api_GET/"
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
def load_nass_county(crop: str, year: int = 2025) -> pd.DataFrame:
    # Do NOT filter prodn_practice_desc here — some counties only report
    # IRRIGATED / NON-IRRIGATED without an aggregate row; we dedup below.
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
        st.warning(f"NASS API error for {crop}: {e}")
        return pd.DataFrame(columns=["State", "County", "fips", "Production"])

    records = raw.get("data", [])
    if not records:
        return pd.DataFrame(columns=["State", "County", "fips", "Production"])

    df = pd.DataFrame(records)
    needed = ["state_alpha", "county_name", "state_fips_code",
              "county_ansi", "prodn_practice_desc", "Value"]
    df = df[[c for c in needed if c in df.columns]].copy()

    # Drop state-level aggregate rows (county_ansi flags)
    df = df[~df["county_ansi"].isin(["998", "000", "999"])]

    df["Production"] = pd.to_numeric(
        df["Value"].str.replace(",", "", regex=False).str.strip(),
        errors="coerce",
    ).fillna(0)

    df["fips"]   = df["state_fips_code"].str.zfill(2) + df["county_ansi"].str.zfill(3)
    df["State"]  = df["state_alpha"].str.strip()
    df["County"] = df["county_name"].str.strip().str.title()

    # Dedup: prefer the "ALL PRODUCTION PRACTICES" row per county; if absent,
    # keep the single highest-value row (avoids double-counting irrigated +
    # non-irrigated where only a breakdown exists).
    key = ["State", "County", "fips"]
    all_prac = "ALL PRODUCTION PRACTICES"
    if "prodn_practice_desc" in df.columns:
        has_all = df[df["prodn_practice_desc"] == all_prac].copy()
        no_all  = df[~df["fips"].isin(has_all["fips"].unique())].copy()
        # For counties without an aggregate row, keep the max-production row
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
    """Return a FeatureCollection for one state with coordinates rounded to
    3 decimal places (~100 m precision).  Rounding cuts Plotly figure JSON
    by ~50 %, reducing serialisation time and browser render time.
    """
    features = []
    for f in _geo["features"]:
        if f["properties"]["STATE"] != sfips:
            continue
        gtype = f["geometry"]["type"]
        raw   = f["geometry"]["coordinates"]
        if gtype == "Polygon":
            coords = [[[round(x, 3), round(y, 3)] for x, y in ring]
                      for ring in raw]
        elif gtype == "MultiPolygon":
            coords = [[[[round(x, 3), round(y, 3)] for x, y in ring]
                       for ring in poly]
                      for poly in raw]
        else:
            coords = raw
        features.append({
            "type": "Feature",
            "id": f.get("id"),          # Plotly matches choropleth locations by this field
            "properties": f["properties"],
            "geometry": {"type": gtype, "coordinates": coords},
        })
    return {"type": "FeatureCollection", "features": features}


@st.cache_data
def build_fips_lookup(_geo):
    inv = {v: k for k, v in STATE_FIPS.items()}
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
        geojson=state_geo, locations=all_fips, z=[0] * len(all_fips),
        colorscale=[[0, PANEL], [1, PANEL]], showscale=False,
        marker=dict(line=county_line), hoverinfo="skip",
    ))
    fig.add_trace(go.Choropleth(
        geojson=state_geo, locations=df["fips"].tolist(), z=z_vals,
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
    fig.update_geos(fitbounds="locations", visible=False)
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
def build_nass_state_fig(df, crop, logo_50yr):
    agg = df.groupby("State")["Production"].sum().reset_index()
    agg["StateName"] = agg["State"].map(ABBR_TO_NAME)

    title_text = f"NASS 2025 {crop} — Production<br><sup>Map labels in M bu</sup>"
    fig = px.choropleth(
        agg, locations="State", locationmode="USA-states", color="Production",
        scope="usa", color_continuous_scale="YlOrBr",
        hover_name="StateName",
        hover_data={"Production": ":,.0f", "State": False},
        labels={"Production": "Production (bu)"},
    )
    fig.update_layout(
        **_base_layout(title_text), height=520,
        geo=dict(showlakes=False, bgcolor=DARK, landcolor=LAND, showland=True, showframe=False),
        coloraxis_colorbar=dict(
            title=dict(text="Production<br>(bu)", font=dict(color=TEXT)),
            tickfont=dict(color=TEXT),
        ),
    )
    lons, lats, texts = [], [], []
    for _, row in agg.iterrows():
        label = format_nass_label(row["Production"])
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


def build_nass_county_fig(state_df, geo, state, crop, logo_50yr, centroids):
    if state_df.empty:
        return None

    sfips = STATE_FIPS_ALL.get(state)
    if sfips is None and not state_df.empty:
        sfips = state_df["fips"].iloc[0][:2]
    if sfips is None:
        return None

    state_geo = get_state_geojson(geo, sfips)  # cached
    all_fips = [f["properties"]["STATE"] + f["properties"]["COUNTY"]
                for f in state_geo["features"]]

    z_vals = state_df["Production"].tolist()
    z_min  = state_df["Production"].min()
    z_max  = state_df["Production"].max()
    if z_min == z_max:
        z_min = 0

    county_line = dict(color="#3d5248", width=0.8)
    fig = go.Figure()
    fig.add_trace(go.Choropleth(
        geojson=state_geo, locations=all_fips, z=[0] * len(all_fips),
        colorscale=[[0, PANEL], [1, PANEL]], showscale=False,
        marker=dict(line=county_line), hoverinfo="skip",
    ))
    fig.add_trace(go.Choropleth(
        geojson=state_geo, locations=state_df["fips"].tolist(), z=z_vals,
        colorscale="YlOrBr", zmin=z_min, zmax=z_max,
        colorbar=dict(
            title=dict(text="Production<br>(bu)", font=dict(color=TEXT)),
            tickfont=dict(color=TEXT),
        ),
        marker=dict(line=county_line),
        text=state_df["County"].tolist(),
        hovertemplate="%{text}: %{z:,.0f} bu<extra></extra>",
    ))

    state_name = ABBR_TO_NAME.get(state, state)
    title_text = (
        f"NASS 2025 {crop} — Production | {state_name} Counties"
        f"<br><sup>Map labels in M bu</sup>"
    )
    fig.update_geos(fitbounds="locations", visible=False)
    fig.update_layout(**_base_layout(title_text), height=620)
    _add_logo(fig, logo_50yr, size=0.15, opacity=1.0, x=0.99, y=0.03, yanchor="bottom")
    _place_labels(fig, state_df["fips"].tolist(), state_df["Production"].tolist(),
                  centroids, format_nass_label)
    return fig


def build_nass_ranking_chart(state_df, state, crop):
    ranked     = state_df.dropna(subset=["Production"]).sort_values("Production", ascending=True)
    state_name = ABBR_TO_NAME.get(state, state)
    raw_avg    = ranked["Production"].mean()
    x_vals     = ranked["Production"] / 1_000_000
    avg_disp   = raw_avg / 1_000_000

    colors = [ACCENT if v >= raw_avg else "#e05252" for v in ranked["Production"]]
    labels = [f"{v:,.2f}" for v in x_vals]

    fig = go.Figure(go.Bar(
        x=x_vals, y=ranked["County"], orientation="h",
        marker_color=colors, marker_line_width=0,
        text=labels, textposition="outside",
        textfont=dict(color=TEXT, size=8), cliponaxis=False,
        hovertemplate="%{y}: %{x:,.2f} M bu<extra></extra>",
    ))
    fig.add_vline(
        x=avg_disp, line_color="#f5a623", line_width=1.5, line_dash="dash",
        annotation_text=f"  Avg: {avg_disp:,.2f} M bu",
        annotation_position="top left",
        annotation_font=dict(color="#f5a623", size=10),
    )
    fig.update_layout(
        paper_bgcolor=DARK, plot_bgcolor=SURFACE,
        font=dict(color=TEXT, family="Arial"),
        title=dict(
            text=f"{state_name} County Rankings — {crop} Production (NASS 2025)",
            font=dict(size=14, color=ACCENT),
        ),
        height=max(380, len(ranked) * 22 + 80),
        margin=dict(l=10, r=90, t=50, b=20), bargap=0.18,
        xaxis=dict(title="Production (M bu)", gridcolor=BORDER,
                   tickfont=dict(color=MUTED), title_font=dict(color=MUTED), zeroline=False),
        yaxis=dict(gridcolor=BORDER, tickfont=dict(color=TEXT, size=9), automargin=True),
    )
    return fig


# ── Cached county figure wrappers ─────────────────────────────────────────────
# Cache key = hashable args only.  _geo / _centroids / _logo are excluded
# (underscore prefix) so large objects aren't hashed or pickled into the key.
# The Figure object is stored in Streamlit's in-memory cache; repeat clicks on
# the same state+crop are instant after the first render.

@st.cache_data(show_spinner=False)
def cached_nass_county_fig(state: str, crop: str, year: int,
                            _geo, _centroids, _logo_50yr):
    df_all   = load_nass_county(crop, year)
    state_df = df_all[df_all["State"] == state].copy()
    return build_nass_county_fig(state_df, _geo, state, crop, _logo_50yr, _centroids)


@st.cache_data(show_spinner=False)
def cached_rma_county_fig(state: str, crop: str, metric: str,
                           practice: str, wheat_type,
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

    tab_nass, tab_rma = st.tabs(["🌾  NASS Production", "📋  RMA"])

    # ══════════════════════════════════════════════════════════════════════════
    # NASS TAB
    # ══════════════════════════════════════════════════════════════════════════
    with tab_nass:
        if "nass_sel_state" not in st.session_state:
            st.session_state.nass_sel_state = None

        nc1, nc2, nc3 = st.columns([1, 1.8, 0.6])
        with nc1:
            nass_crop = st.selectbox("Crop", list(NASS_CROP_PARAMS.keys()), key="nass_crop")
        with nc3:
            st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
            if st.button("🔄 Refresh", use_container_width=True, key="nass_refresh"):
                st.cache_data.clear()
                st.rerun()

        with st.spinner(f"Loading NASS 2025 {nass_crop} data..."):
            nass_df = load_nass_county(nass_crop)

        if nass_df.empty:
            st.warning(
                f"No NASS 2025 county-level production data returned for {nass_crop}. "
                "The data may not yet be published or the API parameters may need adjustment."
            )
        else:
            states_avail_nass = sorted(nass_df["State"].unique())
            with nc2:
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

            # Summary metrics
            scope_nass = (
                nass_df if st.session_state.nass_sel_state is None
                else nass_df[nass_df["State"] == st.session_state.nass_sel_state]
            )
            nm1, nm2, nm3 = st.columns(3)
            nm1.metric("Total Production", f"{scope_nass['Production'].sum():,.0f} bu")
            nm2.metric("Counties Reporting",
                       f"{scope_nass[['State','County']].drop_duplicates().shape[0]:,}")
            nm3.metric("States", f"{scope_nass['State'].nunique():,}")

            # Map / county drill-down
            if st.session_state.nass_sel_state is None:
                nass_fig = build_nass_state_fig(nass_df, nass_crop, logo_50yr)
                st.plotly_chart(nass_fig, use_container_width=True, key="nass_state_map")
                st.caption("Use the State Drill-Down dropdown above to view county detail.")

            else:
                nass_state = st.session_state.nass_sel_state
                state_df   = nass_df[nass_df["State"] == nass_state].copy()

                if st.button("← Back to US Map", key="nass_back_btn"):
                    st.session_state.nass_sel_state = None
                    st.rerun()

                if state_df.empty or state_df["Production"].sum() == 0:
                    st.warning(
                        f"No NASS 2025 county data available for "
                        f"{ABBR_TO_NAME.get(nass_state, nass_state)}. "
                        "This crop may not be produced in this state or data has not been published."
                    )
                else:
                    with st.spinner(f"Building {ABBR_TO_NAME.get(nass_state, nass_state)} county map…"):
                        nass_county_fig = cached_nass_county_fig(
                            nass_state, nass_crop, 2025, geo, centroids, logo_50yr
                        )
                    if nass_county_fig is None:
                        st.info(f"County map not available for "
                                f"{ABBR_TO_NAME.get(nass_state, nass_state)}.")
                    else:
                        st.plotly_chart(nass_county_fig, use_container_width=True,
                                        key="nass_county_map")
                        st.caption("Use ← Back to US Map to return to the national overview.")

                st.markdown(f"<hr style='border-color:{BORDER};margin:8px 0'>",
                            unsafe_allow_html=True)
                ranking_nass = build_nass_ranking_chart(state_df, nass_state, nass_crop)
                st.plotly_chart(ranking_nass, use_container_width=True, key="nass_ranking")

                with st.expander(
                    f"County Data Table — {ABBR_TO_NAME.get(nass_state, nass_state)}",
                    expanded=False
                ):
                    disp = (state_df[["County", "Production"]]
                            .sort_values("Production", ascending=False).copy())
                    disp["Production (bu)"] = disp["Production"].apply(
                        lambda v: f"{v:,.0f}" if pd.notna(v) else "—"
                    )
                    st.dataframe(disp[["County", "Production (bu)"]],
                                 use_container_width=True, hide_index=True)

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
            st.plotly_chart(fig, use_container_width=True, key="rma_state_map")
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
                        state, crop, metric, practice, wheat_type,
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


if __name__ == "__main__":
    main()
