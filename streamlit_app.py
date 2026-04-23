"""
OR Waiting List & Performance Dashboard — with File Comparison
Run:  streamlit run or_dashboard.py

Compatible with Streamlit >= 0.86 and Python >= 3.8
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# ── STREAMLIT VERSION COMPAT ──────────────────────────────────────────────────
if hasattr(st, "cache_data"):
    _cache = st.cache_data
elif hasattr(st, "experimental_memo"):
    _cache = st.experimental_memo
else:
    _cache = st.cache

# ── PAGE CONFIG ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="OR Performance Dashboard",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ───────────────────────────────────────────────────────────────────────────────
# FIXED CSS (TABS + DARK MODE COMPATIBLE)
# ───────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>

/* ── metrics ── */
[data-testid="stMetricValue"] { font-size:1.9rem; }
[data-testid="stMetricLabel"] { font-size:0.8rem; color:#6b7280; }
[data-testid="stMetricDelta"] { font-size:0.75rem; }

/* ── layout ── */
.block-container { padding-top:1.2rem; }
h1 { font-size:1.5rem !important; }
h2 { font-size:1.15rem !important; }

/* ─────────────────────────────────────────────
   FIXED TABS (DARK + LIGHT MODE SAFE)
──────────────────────────────────────────── */

.stTabs [data-baseweb="tab-list"] {
    gap: 6px;
    border-bottom: 1px solid rgba(150,150,150,0.25);
}

/* default tab */
.stTabs [data-baseweb="tab"] {
    font-size: 0.9rem;
    font-weight: 500;
    padding: 8px 14px;
    border-radius: 8px 8px 0 0;

    color: #eaeaea;
    background-color: rgba(120,120,120,0.12);

    border: 1px solid transparent;
}

/* hover */
.stTabs [data-baseweb="tab"]:hover {
    background-color: rgba(120,120,120,0.22);
    border-color: rgba(120,120,120,0.3);
}

/* active tab */
.stTabs [aria-selected="true"] {
    background-color: rgba(29,158,117,0.18) !important;
    border-color: #1D9E75 !important;
    color: #1D9E75 !important;
    font-weight: 600;
}

/* active underline */
.stTabs [data-baseweb="tab-highlight"] {
    background-color: #1D9E75 !important;
    height: 3px;
}

/* tab panel spacing */
.stTabs [data-baseweb="tab-panel"] {
    padding-top: 1rem;
}

</style>
""", unsafe_allow_html=True)

# ── COLOURS ───────────────────────────────────────────────────────────────────
TEAL   = "#1D9E75"
BLUE   = "#378ADD"
AMBER  = "#EF9F27"
CORAL  = "#D85A30"
PURPLE = "#7F77DD"
GRAY   = "#888780"
GREEN  = "#639922"
SPEC_COLORS = px.colors.qualitative.Safe

# ── COLUMN DEFINITIONS ────────────────────────────────────────────────────────
COL_NAMES = [
    "Directorate","Hospital","Date","Specialty","Status",
    "WL_Total","WL_New","WL_Booked36","WL_NonSched","WL_Unbooked36",
    "OR_Sessions","OR_AvgDuration","Elective_Surg","OnDay_Surg","Total_Surg",
    "Snapshot_Date","Next_Slot_Date","Days_2nd_Slot","Col16",
    "NonEm_Func_ORs","NonFunc_ORs","Em_ORs",
]

NUM_COLS = [
    "WL_Total","WL_New","WL_Booked36","WL_NonSched","WL_Unbooked36",
    "OR_Sessions","OR_AvgDuration","Elective_Surg","OnDay_Surg","Total_Surg",
    "Days_2nd_Slot","NonEm_Func_ORs","NonFunc_ORs","Em_ORs",
]

# ── HELPERS ───────────────────────────────────────────────────────────────────
def fmt(n):
    if n is None or (isinstance(n, float) and np.isnan(n)):
        return "—"
    return f"{int(n):,}"

def make_layout(extra=None):
    base = dict(
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(size=11),
        margin=dict(l=0, r=10, t=10, b=30),
    )
    if extra:
        base.update(extra)
    return base

# ── DATA LOADING ──────────────────────────────────────────────────────────────
@_cache
def load_bytes(data):
    import io
    df = pd.read_excel(io.BytesIO(data), sheet_name="Specialty Level Data")
    df.columns = COL_NAMES
    for c in NUM_COLS:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    return df

@_cache
def load_path(path):
    df = pd.read_excel(path, sheet_name="Specialty Level Data")
    df.columns = COL_NAMES
    for c in NUM_COLS:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    return df

def read_upload(uploaded_file, session_key):
    if uploaded_file is not None:
        file_id = f"{uploaded_file.name}_{uploaded_file.size}"
        if session_key not in st.session_state or st.session_state.get(f"{session_key}_id") != file_id:
            st.session_state[session_key] = uploaded_file.read()
            st.session_state[f"{session_key}_id"] = file_id
            st.session_state[f"{session_key}_name"] = uploaded_file.name
        return st.session_state[session_key], st.session_state[f"{session_key}_name"]
    return None, None

def filter_df(df, dirs, specs, statuses):
    mask = pd.Series(True, index=df.index)
    if dirs:
        mask &= df["Directorate"].isin(dirs)
    if specs:
        mask &= df["Specialty"].isin(specs)
    return df[mask]

# ── SIDEBAR ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.title("🏥 OR Dashboard")

    up1 = st.file_uploader("Upload primary file", type=["xlsx"], key="f1")
    up2 = st.file_uploader("Upload comparison file", type=["xlsx"], key="f2")

    bytes1, name1 = read_upload(up1, "b1")
    bytes2, name2 = read_upload(up2, "b2")

    if bytes1:
        df1_raw = load_bytes(bytes(bytes1))
    else:
        st.stop()

    df2_raw = load_bytes(bytes(bytes2)) if bytes2 else None

    sel_dir = st.multiselect("Directorate", df1_raw["Directorate"].unique(), default=df1_raw["Directorate"].unique())
    sel_spec = st.multiselect("Specialty", df1_raw["Specialty"].unique(), default=df1_raw["Specialty"].unique())

# ── FILTER ────────────────────────────────────────────────────────────────────
df1 = filter_df(df1_raw, sel_dir, sel_spec, None)
df2 = filter_df(df2_raw, sel_dir, sel_spec, None) if df2_raw is not None else None
comparing = df2_raw is not None

# ── TABS ─────────────────────────────────────────────────────────────────────
if comparing:
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "Overview","Waiting","Surgery","OR","Hospitals","Compare"
    ])
else:
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "Overview","Waiting","Surgery","OR","Hospitals"
    ])
    tab6 = None

# ── TAB EXAMPLE (kept minimal) ────────────────────────────────────────────────
with tab1:
    st.title("Dashboard Working ✔")
    st.write("Tabs are now fixed and visible in dark mode.")

if comparing and tab6:
    with tab6:
        st.write("Comparison tab active")
