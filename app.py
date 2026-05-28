import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import json
import urllib.request
import numpy as np
import base64
from pathlib import Path

st.set_page_config(page_title="USDA RMA Yield and Production", layout="wide")

# ── Paths (relative to this file — works locally and on Streamlit Cloud) ──────
HERE      = Path(__file__).parent
DATA_PATH = HERE / "data" / "2025 RMA Production Data.xlsx"
LOGO_50YR = HERE / "assets" / "50 Year logo JSA.png"
LOGO_TRANS = HERE / "assets" / "Transparent Smal logo.png"
LOGO_FULL  = HERE / "assets" / "logo-full.png"

STATE_FIPS = {
    "AL": "01", "AR": "05", "CO": "08", "GA": "13", "IA": "19",
    "IL": "17", "IN": "18", "KS": "20", "KY": "21", "MI": "26",
    "MN": "27", "MO": "29", "MS": "28", "ND": "38", "NE": "31",
    "OH": "39", "OK": "40", "PA": "42", "SC": "45", "SD": "46",
    "TN": "47", "TX": "48", "WI": "55",
}

# Approximate geographic centroids (lon, lat) for label placement on US map
STATE_CENTROIDS = {
    "AL": (-86.8,  32.8), "AR": (-92.4,  34.9), "CO": (-105.5, 39.0),
    "GA": (-83.4,  32.7), "IA": (-93.1,  42.0), "IL": (-89.2,  40.0),
    "IN": (-86.3,  40.3), "KS": (-98.4,  38.5), "KY": (-84.9,  37.5),
    "MI": (-84.5,  44.3), "MN": (-94.3,  46.4), "MO": (-92.5,  38.4),
    "MS": (-89.7,  32.7), "ND": (-100.5, 47.5), "NE": (-99.9,  41.5),
    "OH": (-82.8,  40.4), "OK": (-97.5,  35.5), "PA": (-77.2,  40.9),
    "SC": (-80.9,  33.8), "SD": (-100.2, 44.4), "TN": (-86.7,  35.8),
    "TX": (-99.3,  31.5), "WI": (-89.8,  44.5),
}

ABBR_TO_NAME = {
    "AL": "Alabama", "AR": "Arkansas", "CO": "Colorado", "GA": "Georgia",
    "IA": "Iowa", "IL": "Illinois", "IN": "Indiana", "KS": "Kansas",
    "KY": "Kentucky", "MI": "Michigan", "MN": "Minnesota", "MO": "Missouri",
    "MS": "Mississippi", "ND": "North Dakota", "NE": "Nebraska",
    "OH": "Ohio", "OK": "Oklahoma", "PA": "Pennsylvania",
    "SC": "South Carolina", "SD": "South Dakota", "TN": "Tennessee",
    "TX": "Texas", "WI": "Wisconsin",
}

METRIC_COL = {
    "Production":             "Reported Production",
    "Production Acres":       "Reported Production Acres",
    "Yield":                  "Reported Yield Mean",
    "Prevent Planted Acres":  "Prev Plant Acres",
}
METRIC_UNIT = {
    "Production": "bu", "Production Acres": "ac",
    "Yield": "bu/ac", "Prevent Planted Acres": "ac",
}
METRIC_FMT = {
    "Production": ",.0f", "Production Acres": ",.0f",
    "Yield": ".1f", "Prevent Planted Acres": ",.0f",
}
COLOR_SCALE = {
    "Production": "YlOrBr", "Production Acres": "YlGn",
    "Yield": "RdYlGn", "Prevent Planted Acres": "OrRd",
}

# ── Data ───────────────────────────────────────────────────────────────────────

@st.cache_data
def load_data():
    out = {}
    for crop in ["Corn", "Soybeans"]:
        df = pd.read_excel(DATA_PATH, sheet_name=crop)
        df.columns = df.columns.str.strip()
        df["State"] = df["State"].str.strip()
        df["County"] = df["County"].str.strip()
        df["Practice"] = df["Practice"].str.strip()
        df["PG"] = np.where(
            df["Practice"].str.startswith("Irrigated"), "Irrigated",
            np.where(df["Practice"].str.startswith("Non-Irrigated"), "Non-Irrigated", "Invalid"),
        )
        df = df[df["PG"] != "Invalid"].copy()
        for col in METRIC_COL.values():
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        out[crop] = df
    return out


@st.cache_data
def load_geojson():
    url = "https://raw.githubusercontent.com/plotly/datasets/master/geojson-counties-fips.json"
    with urllib.request.urlopen(url) as r:
        return json.load(r)


@st.cache_data
def build_fips_lookup(_geo):
    inv = {v: k for k, v in STATE_FIPS.items()}
    lk = {}
    for feat in _geo["features"]:
        p = feat["properties"]
        if p["STATE"] not in inv:
            continue
        abbr = inv[p["STATE"]]
        name = p["NAME"]
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
    a = x[:-1] * y[1:] - x[1:] * y[:-1]
    A = 0.5 * a.sum()
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
        geom = feat["geometry"]
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


DISPLAY_DIVISOR = {
    "Production": 1_000_000,
    "Production Acres": 100_000,
    "Yield": 1,
    "Prevent Planted Acres": 100_000,
}
DISPLAY_UNIT = {
    "Production": "M bu",
    "Production Acres": "×100K ac",
    "Yield": "bu/ac",
    "Prevent Planted Acres": "×100K ac",
}


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
    """Compact label for US overview map."""
    if pd.isna(val) or val == 0:
        return ""
    if metric == "Yield":
        return f"{val:.0f}"
    if metric == "Production":
        return f"{val / 1_000_000:.1f}"
    return f"{val / 100_000:.1f}"


# ── Aggregation ────────────────────────────────────────────────────────────────

def filter_practice(df, practice):
    return df if practice == "All" else df[df["PG"] == practice]


def agg_data(df, practice, metric, group_cols):
    col = METRIC_COL[metric]
    df = filter_practice(df, practice)
    if metric == "Yield":
        prod = df.groupby(group_cols)["Reported Production"].sum()
        acres = df.groupby(group_cols)["Reported Production Acres"].sum()
        result = (prod / acres.replace(0, np.nan)).reset_index()
        result.columns = group_cols + [col]
    else:
        result = df.groupby(group_cols)[col].sum().reset_index()
    return result


# ── JPSI brand palette ────────────────────────────────────────────────────────
DARK    = "#0e1614"
PANEL   = "#162019"
SURFACE = "#1e2e2a"
BORDER  = "#243328"
TEXT    = "#e4e8f0"
MUTED   = "#7a9990"
ACCENT  = "#4ade80"
LAND    = "#1a2720"


@st.cache_data
def load_logo(path):
    with open(path, "rb") as f:
        return "data:image/png;base64," + base64.b64encode(f.read()).decode()


def _add_logo(fig, logo_src, size=0.13, opacity=0.92, x=0.99, y=0.01, yanchor="bottom", layer="above"):
    fig.add_layout_image(
        source=logo_src,
        xref="paper", yref="paper",
        x=x, y=y,
        xanchor="right", yanchor=yanchor,
        sizex=size, sizey=size,
        sizing="contain",
        opacity=opacity,
        layer=layer,
    )


# ── Figures ────────────────────────────────────────────────────────────────────

def build_ranking_chart(agg, metric, state):
    col = METRIC_COL[metric]
    state_name = ABBR_TO_NAME.get(state, state)
    divisor = DISPLAY_DIVISOR[metric]
    disp_unit = DISPLAY_UNIT[metric]

    ranked = agg.dropna(subset=[col]).sort_values(col, ascending=True)
    raw_avg = ranked[col].mean()
    x_vals = ranked[col] / divisor
    avg_disp = raw_avg / divisor

    colors = [ACCENT if v >= raw_avg else "#e05252" for v in ranked[col]]
    fmt = ".1f" if metric == "Yield" else ",.2f"
    labels = [f"{v:{fmt}}" for v in x_vals]

    fig = go.Figure(go.Bar(
        x=x_vals,
        y=ranked["County"],
        orientation="h",
        marker_color=colors,
        marker_line_width=0,
        text=labels,
        textposition="outside",
        textfont=dict(color=TEXT, size=8),
        cliponaxis=False,
        hovertemplate=f"%{{y}}: %{{x:{fmt}}} {disp_unit}<extra></extra>",
    ))

    fig.add_vline(
        x=avg_disp,
        line_color="#f5a623",
        line_width=1.5,
        line_dash="dash",
        annotation_text=f"  Avg: {avg_disp:{fmt}} {disp_unit}",
        annotation_position="top left",
        annotation_font=dict(color="#f5a623", size=10),
    )

    height = max(380, len(ranked) * 22 + 80)

    fig.update_layout(
        paper_bgcolor=DARK,
        plot_bgcolor=SURFACE,
        font=dict(color=TEXT, family="Arial"),
        title=dict(
            text=f"{state_name} County Rankings — {metric}",
            font=dict(size=14, color=ACCENT),
        ),
        height=height,
        margin=dict(l=10, r=90, t=50, b=20),
        bargap=0.18,
        xaxis=dict(
            title=f"{metric} ({disp_unit})",
            gridcolor=BORDER,
            tickfont=dict(color=MUTED),
            title_font=dict(color=MUTED),
            zeroline=False,
        ),
        yaxis=dict(
            gridcolor=BORDER,
            tickfont=dict(color=TEXT, size=9),
            automargin=True,
        ),
    )
    return fig


def _base_layout(title):
    return dict(
        paper_bgcolor=DARK,
        plot_bgcolor=SURFACE,
        font=dict(color=TEXT, family="Arial"),
        title=dict(text=title, font=dict(size=15, color=ACCENT)),
        margin=dict(l=0, r=0, t=50, b=0),
    )


def build_state_fig(agg, metric, crop, practice, logo_50yr):
    col = METRIC_COL[metric]
    unit = METRIC_UNIT[metric]
    fmt = METRIC_FMT[metric]
    disp_unit = DISPLAY_UNIT[metric]
    df = agg.copy()
    df["StateName"] = df["State"].map(ABBR_TO_NAME)

    title_text = (
        f"{crop} — {metric} | Practice: {practice}"
        f"<br><sup>Map labels in {disp_unit}</sup>"
    )

    fig = px.choropleth(
        df,
        locations="State",
        locationmode="USA-states",
        color=col,
        scope="usa",
        color_continuous_scale=COLOR_SCALE[metric],
        hover_name="StateName",
        hover_data={col: f":{fmt}", "State": False},
        labels={col: f"{metric} ({unit})"},
    )
    fig.update_layout(
        **_base_layout(title_text),
        height=520,
        geo=dict(
            showlakes=False,
            bgcolor=DARK,
            landcolor=LAND,
            showland=True,
            showframe=False,
        ),
        coloraxis_colorbar=dict(
            title=dict(text=f"{metric}<br>({unit})", font=dict(color=TEXT)),
            tickfont=dict(color=TEXT),
        ),
    )

    # State data labels
    lons, lats, texts = [], [], []
    for _, row in df.iterrows():
        label = format_state_label(row[col], metric)
        if label and row["State"] in STATE_CENTROIDS:
            lon, lat = STATE_CENTROIDS[row["State"]]
            lons.append(lon)
            lats.append(lat)
            texts.append(label)

    if lons:
        fig.add_trace(go.Scattergeo(
            lon=lons,
            lat=lats,
            text=texts,
            mode="text",
            geo="geo",
            textfont=dict(color="#cccccc", size=11, family="Arial Black"),
            showlegend=False,
            hoverinfo="skip",
        ))

    _add_logo(fig, logo_50yr, size=0.30, opacity=1.0)
    return fig


def build_county_fig(agg, geo, fips_lk, centroids, state, metric, crop, practice, logo_50yr):
    col = METRIC_COL[metric]
    unit = METRIC_UNIT[metric]
    fmt = METRIC_FMT[metric]
    sfips = STATE_FIPS[state]

    df = agg.copy()
    df["fips"] = df["County"].apply(lambda c: resolve_fips(state, c, fips_lk))
    df = df.dropna(subset=["fips"])

    state_geo = {
        "type": "FeatureCollection",
        "features": [f for f in geo["features"] if f["properties"]["STATE"] == sfips],
    }
    state_name = ABBR_TO_NAME.get(state, state)

    all_fips = [
        f["properties"]["STATE"] + f["properties"]["COUNTY"]
        for f in state_geo["features"]
    ]

    z_vals = df[col].tolist()
    z_min = df[col].min() if z_vals else 0
    z_max = df[col].max() if z_vals else 1
    if z_min == z_max:
        z_min = 0

    county_line = dict(color="#3d5248", width=0.8)

    fig = go.Figure()

    # Layer 1: every county outlined with uniform dark fill
    fig.add_trace(go.Choropleth(
        geojson=state_geo,
        locations=all_fips,
        z=[0] * len(all_fips),
        colorscale=[[0, PANEL], [1, PANEL]],
        showscale=False,
        marker=dict(line=county_line),
        hoverinfo="skip",
    ))

    # Layer 2: counties with data, colored by metric
    fig.add_trace(go.Choropleth(
        geojson=state_geo,
        locations=df["fips"].tolist(),
        z=z_vals,
        colorscale=COLOR_SCALE[metric],
        zmin=z_min,
        zmax=z_max,
        colorbar=dict(
            title=dict(text=f"{metric}<br>({unit})", font=dict(color=TEXT)),
            tickfont=dict(color=TEXT),
        ),
        marker=dict(line=county_line),
        text=df["County"].tolist(),
        hovertemplate=f"%{{text}}: %{{z:{fmt}}}<extra></extra>",
    ))

    disp_unit = DISPLAY_UNIT[metric]
    title_text = (
        f"{crop} — {metric} | {state_name} Counties | Practice: {practice}"
        f"<br><sup>Map labels in {disp_unit}</sup>"
    )
    fig.update_geos(fitbounds="locations", visible=False)
    fig.update_layout(**_base_layout(title_text), height=620)
    _add_logo(fig, logo_50yr, size=0.15, opacity=1.0, x=0.99, y=0.03, yanchor="bottom")

    # Adaptive font size — scale with average county area so labels feel anchored
    county_areas = [centroids[f][2] for f in all_fips if f in centroids]
    if county_areas:
        avg_area = float(np.mean(county_areas))
        # Small counties (~0.05 sq°) → size 9, large counties (~0.5 sq°) → size 14
        label_size = int(np.clip(9 + np.log(max(avg_area, 0.01) / 0.05) * 2.0, 9, 15))
    else:
        label_size = 10

    # County data labels — larger counties get priority
    candidates = []
    for _, row in df.iterrows():
        fips = row["fips"]
        label = format_label(row[col], metric)
        if label and fips in centroids:
            cx, cy, area = centroids[fips]
            candidates.append((area, cx, cy, label))
    candidates.sort(reverse=True)

    # Deconfliction gap also scales with county size so dense states skip more
    min_sep = np.clip(0.15 + avg_area * 0.8, 0.20, 0.45) if county_areas else 0.30
    placed, lons, lats, texts = [], [], [], []
    for area, cx, cy, label in candidates:
        if not any((cx - px) ** 2 + (cy - py) ** 2 < min_sep ** 2 for px, py in placed):
            placed.append((cx, cy))
            lons.append(cx)
            lats.append(cy)
            texts.append(label)

    if lons:
        fig.add_trace(go.Scattergeo(
            lon=lons,
            lat=lats,
            text=texts,
            mode="text",
            textfont=dict(color="#aaaaaa", size=label_size, family="Arial Black"),
            showlegend=False,
            hoverinfo="skip",
        ))

    return fig


# ── App ────────────────────────────────────────────────────────────────────────

def main():
    st.markdown(
        f"""
        <style>
        .stApp {{ background-color: {DARK}; color: {TEXT}; }}
        [data-testid="stSidebar"] {{
            background-color: {PANEL};
            border-right: 1px solid {BORDER};
        }}
        .block-container {{ padding-top: 1rem; max-width: 1400px; }}
        h1, h2, h3 {{ color: {ACCENT} !important; letter-spacing: 0.02em; }}
        p, label, .stCaption, [data-testid="stCaptionContainer"] {{
            color: {MUTED} !important;
        }}
        [data-testid="stSelectbox"] label {{ color: {MUTED} !important; font-size: 0.8rem; }}
        div[data-baseweb="select"] > div {{
            background-color: {PANEL} !important;
            border-color: {BORDER} !important;
            color: {TEXT} !important;
        }}
        div[data-baseweb="popover"] * {{
            background-color: {PANEL} !important;
            color: {TEXT} !important;
        }}
        [data-testid="metric-container"] {{
            background-color: {PANEL};
            border: 1px solid {BORDER};
            border-radius: 8px;
            padding: 12px 16px;
        }}
        [data-testid="stMetricValue"] {{ color: {ACCENT} !important; font-size: 1.35rem; font-weight: 700; }}
        [data-testid="stMetricLabel"] {{ color: {MUTED} !important; font-size: 0.78rem; text-transform: uppercase; letter-spacing: 0.06em; }}
        [data-testid="stExpander"] {{
            background-color: {PANEL};
            border: 1px solid {BORDER};
            border-radius: 6px;
        }}
        [data-testid="stDataFrame"] {{ background-color: {PANEL}; }}
        hr {{ border-color: {BORDER}; }}
        [data-testid="stSpinner"] p {{ color: {MUTED} !important; }}
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.title("USDA RMA Yield and Production")

    if "sel_state" not in st.session_state:
        st.session_state.sel_state = None

    with st.spinner("Loading..."):
        data = load_data()
        geo = load_geojson()
        fips_lk = build_fips_lookup(geo)
        centroids = build_centroid_lookup(geo)
        logo_50yr = load_logo(LOGO_50YR)
        logo_trans = load_logo(LOGO_TRANS)
        logo_full  = load_logo(LOGO_FULL)

    st.markdown(
        f"""
        <style>
        [data-testid="stHeader"] {{
            background-image: url('{logo_full}');
            background-repeat: no-repeat;
            background-position: right 90px center;
            background-size: auto 68%;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )

    c1, c2, c3, c4, c5 = st.columns([1, 1.2, 1.2, 1.5, 0.6])

    with c1:
        crop = st.selectbox("Crop", ["Corn", "Soybeans"])
    with c2:
        metric = st.selectbox("Metric", list(METRIC_COL.keys()))
    with c3:
        practice = st.selectbox("Practice", ["All", "Irrigated", "Non-Irrigated"])
    with c5:
        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
        if st.button("🔄 Refresh Data", use_container_width=True):
            st.cache_data.clear()
            st.rerun()
    with c4:
        df = data[crop]
        states_avail = sorted(df["State"].unique())
        state_opts = ["— US Overview —"] + [
            f"{a}  —  {ABBR_TO_NAME.get(a, a)}" for a in states_avail
        ]
        default_idx = 0
        if st.session_state.sel_state:
            try:
                default_idx = states_avail.index(st.session_state.sel_state) + 1
            except ValueError:
                default_idx = 0
        sel = st.selectbox("State Drill-Down", state_opts, index=default_idx)
        st.session_state.sel_state = None if sel.startswith("—") else sel[:2]

    df = data[crop]
    col = METRIC_COL[metric]
    unit = METRIC_UNIT[metric]
    fmt = METRIC_FMT[metric]

    scope_df = filter_practice(df, practice)
    if st.session_state.sel_state:
        scope_df = scope_df[scope_df["State"] == st.session_state.sel_state]

    if metric == "Yield":
        p = scope_df["Reported Production"].sum()
        a = scope_df["Reported Production Acres"].sum()
        summary_val = p / a if a > 0 else 0.0
    else:
        summary_val = scope_df[col].sum()

    m1, m2, m3 = st.columns(3)
    label = "Avg Yield" if metric == "Yield" else f"Total {metric}"
    m1.metric(f"{label}", f"{summary_val:{fmt}} {unit}")
    m2.metric("Counties", f"{scope_df[['State','County']].drop_duplicates().shape[0]:,}")
    m3.metric("States", f"{scope_df['State'].nunique():,}")

    if st.session_state.sel_state is None:
        agg = agg_data(df, practice, metric, ["State"])
        fig = build_state_fig(agg, metric, crop, practice, logo_50yr)
        event = st.plotly_chart(
            fig, use_container_width=True, on_select="rerun", key="state_map"
        )
        if event and hasattr(event, "selection") and event.selection.points:
            loc = event.selection.points[0].get("location")
            if loc and loc in states_avail:
                st.session_state.sel_state = loc
                st.rerun()
        st.caption("Click any state on the map or use the State Drill-Down dropdown to view county detail.")

    else:
        state = st.session_state.sel_state
        agg = agg_data(df[df["State"] == state], practice, metric, ["County"])

        if agg.empty or agg[col].sum() == 0:
            st.warning(
                f"No data for {ABBR_TO_NAME.get(state, state)} with the selected filters."
            )
        else:
            fig = build_county_fig(agg, geo, fips_lk, centroids, state, metric, crop, practice, logo_50yr)
            st.plotly_chart(fig, use_container_width=True, key="county_map")

        state_name = ABBR_TO_NAME.get(state, state)

        st.markdown(f"<hr style='border-color:{BORDER};margin:8px 0'>", unsafe_allow_html=True)
        ranking_fig = build_ranking_chart(agg, metric, state)
        st.plotly_chart(ranking_fig, use_container_width=True, key="ranking_chart")

        with st.expander(f"County Data Table — {state_name}", expanded=False):
            disp = agg.sort_values(col, ascending=False).copy()
            disp.columns = ["County", f"{metric} ({unit})"]
            disp[f"{metric} ({unit})"] = disp[f"{metric} ({unit})"].apply(
                lambda v: f"{v:,.1f}" if pd.notna(v) else "—"
            )
            st.dataframe(disp, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
