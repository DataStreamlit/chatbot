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

  /* ── tabs — visible in both light and dark mode ── */
  .stTabs [data-baseweb="tab-list"] {
    gap: 4px;
    background-color: transparent;
    border-bottom: 2px solid rgba(128,128,128,0.25);
    padding-bottom: 0;
  }

  .stTabs [aria-selected="true"] {
    background-color: rgba(29,158,117,0.15) !important;
    border-color: #1D9E75 !important;
    border-bottom-color: transparent !important;
    color: #1D9E75 !important;
  }
  /* active tab bottom marker line */
  .stTabs [data-baseweb="tab-highlight"] {
    background-color: #1D9E75 !important;
    height: 3px;
  }
  /* tab panel top padding */
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
    """Build a fresh layout dict each time — avoids the duplicate-key margin error."""
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
    """
    Read an uploaded file's bytes immediately and cache them in session_state.
    This prevents the 403 error that occurs when Streamlit tries to re-read
    a file object on subsequent script reruns after the upload has expired.
    """
    if uploaded_file is not None:
        file_id = f"{uploaded_file.name}_{uploaded_file.size}"
        if session_key not in st.session_state or st.session_state.get(f"{session_key}_id") != file_id:
            st.session_state[session_key]          = uploaded_file.read()
            st.session_state[f"{session_key}_id"]  = file_id
            st.session_state[f"{session_key}_name"] = uploaded_file.name
        return st.session_state[session_key], st.session_state[f"{session_key}_name"]
    # clear stale cache when file is removed
    for k in [session_key, f"{session_key}_id", f"{session_key}_name"]:
        st.session_state.pop(k, None)
    return None, None

def filter_df(df, dirs, specs, statuses):
    mask = pd.Series(True, index=df.index)
    if dirs:
        mask = mask & df["Directorate"].isin(dirs)
    if specs:
        mask = mask & df["Specialty"].isin(specs)
    include_na  = "N/A" in statuses
    status_vals = [s for s in statuses if s != "N/A"]
    if status_vals and include_na:
        mask = mask & (df["Status"].isin(status_vals) | df["Status"].isna())
    elif status_vals:
        mask = mask & df["Status"].isin(status_vals)
    elif include_na:
        mask = mask & df["Status"].isna()
    return df[mask]

# ── SIDEBAR ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.title("🏥 OR Dashboard")
    st.markdown("---")

    st.subheader("File 1 — Primary")
    up1 = st.file_uploader("Upload primary file (.xlsx)", type=["xlsx"], key="f1")

    st.subheader("File 2 — Comparison (optional)")
    up2 = st.file_uploader(
        "Upload comparison file (.xlsx)", type=["xlsx"], key="f2",
        help="Upload a second Aggregated file to compare week-over-week or version differences",
    )

    # ── Read bytes immediately to avoid 403 on rerun ─────────────────────────
    bytes1, name1 = read_upload(up1, "bytes_f1")
    bytes2, name2 = read_upload(up2, "bytes_f2")

    # ── Load primary ──────────────────────────────────────────────────────────
    if bytes1 is not None:
        df1_raw = load_bytes(bytes(bytes1))
        label1  = name1
        st.success(f"✓ File 1 loaded — {len(df1_raw):,} rows")
    else:
        try:
            df1_raw = load_path("Aggregated_-_Specialty_level____1_.xlsx")
            label1  = "Default aggregated file"
            st.info("File 1: using default file in same folder.")
        except Exception:
            st.warning("⚠️ Upload your Aggregated Excel file to begin.")
            st.stop()

    # ── Load comparison ───────────────────────────────────────────────────────
    df2_raw = None
    label2  = None
    if bytes2 is not None:
        df2_raw = load_bytes(bytes(bytes2))
        label2  = name2
        st.success(f"✓ File 2 loaded — {len(df2_raw):,} rows")

    st.markdown("---")
    st.subheader("Filters")

    all_dir  = sorted(df1_raw["Directorate"].dropna().unique().tolist())
    sel_dir  = st.multiselect("Directorate", all_dir, default=all_dir)

    all_spec = sorted(df1_raw["Specialty"].dropna().unique().tolist())
    sel_spec = st.multiselect("Specialty", all_spec, default=all_spec)

    sel_status = st.multiselect(
        "Status", ["Available","On Hold","N/A"],
        default=["Available","On Hold"],
    )

    st.markdown("---")
    st.caption("OR Analytics · Filters apply to all tabs")

# ── APPLY FILTERS ─────────────────────────────────────────────────────────────
df1    = filter_df(df1_raw, sel_dir, sel_spec, sel_status)
avail1 = df1[df1["Status"] == "Available"]

if df2_raw is not None:
    df2    = filter_df(df2_raw, sel_dir, sel_spec, sel_status)
    avail2 = df2[df2["Status"] == "Available"]
    comparing = True
else:
    df2    = None
    avail2 = None
    comparing = False

# ══════════════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════════════
if comparing:
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "📊 Overview", "⏳ Waiting List", "🔪 Surgical Activity",
        "🏨 OR Rooms", "🔍 Hospital Explorer", "⚖️ File Comparison",
    ])
else:
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📊 Overview", "⏳ Waiting List", "🔪 Surgical Activity",
        "🏨 OR Rooms", "🔍 Hospital Explorer",
    ])
    tab6 = None

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1  OVERVIEW
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.header("Executive Overview")
    if comparing:
        st.caption(f"File 1: **{label1}**   |   File 2: **{label2}**")

    def kpi_row(av, av2=None):
        wl   = av["WL_Total"].sum()
        surg = av["Total_Surg"].sum()
        nh   = av["Hospital"].nunique()
        wait = av["Days_2nd_Slot"].mean()
        oh   = (av.index.map(lambda i: df1.loc[i, "Status"] if i in df1.index else None) == "On Hold").sum() \
               if av2 is None else None

        def delta(v1, v2):
            if v2 is None or pd.isna(v2):
                return None
            return f"{v1 - v2:+,.0f}"

        wl2   = av2["WL_Total"].sum()   if av2 is not None else None
        surg2 = av2["Total_Surg"].sum() if av2 is not None else None
        nh2   = av2["Hospital"].nunique() if av2 is not None else None
        wait2 = av2["Days_2nd_Slot"].mean() if av2 is not None else None

        c1,c2,c3,c4,c5 = st.columns(5)
        c1.metric("Total WL Patients",    fmt(wl),   delta=delta(wl, wl2))
        c2.metric("Total Surgeries",       fmt(surg), delta=delta(surg, surg2))
        c3.metric("Hospitals Reporting",   fmt(nh),   delta=delta(nh, nh2))
        wait_str  = f"{wait:.1f}d"  if not pd.isna(wait)  else "—"
        wait2_str = f"{wait2-wait:+.1f}d" if av2 is not None and not pd.isna(wait2) else None
        c4.metric("Avg Days to Next Slot", wait_str,  delta=wait2_str)
        on_hold_n = (df1["Status"]=="On Hold").sum()
        c5.metric("Specialties On Hold",   fmt(on_hold_n))

    kpi_row(avail1, avail2)

    st.markdown("---")
    left, right = st.columns(2)

    with left:
        st.subheader("Waiting List by Directorate")
        dir_wl1 = avail1.groupby("Directorate")["WL_Total"].sum().reset_index(name="WL_Total")
        if comparing:
            dir_wl2 = avail2.groupby("Directorate")["WL_Total"].sum().reset_index(name="WL_Total2")
            dir_wl  = dir_wl1.merge(dir_wl2, on="Directorate", how="outer").fillna(0)
            dir_wl  = dir_wl.sort_values("WL_Total")
            fig = go.Figure()
            fig.add_trace(go.Bar(y=dir_wl["Directorate"], x=dir_wl["WL_Total"],
                                 name=label1, orientation="h", marker_color=TEAL))
            fig.add_trace(go.Bar(y=dir_wl["Directorate"], x=dir_wl["WL_Total2"],
                                 name=label2, orientation="h", marker_color=BLUE, opacity=0.7))
            fig.update_layout(make_layout({"barmode":"group","height":440,
                                           "legend":dict(orientation="h",y=1.05)}))
        else:
            dir_wl = dir_wl1.sort_values("WL_Total")
            fig = px.bar(dir_wl, x="WL_Total", y="Directorate", orientation="h",
                         color="WL_Total",
                         color_continuous_scale=[[0,"#E1F5EE"],[1,TEAL]],
                         labels={"WL_Total":"Patients","Directorate":""}, height=440)
            fig.update_coloraxes(showscale=False)
            fig.update_layout(make_layout())
        fig.update_xaxes(gridcolor="#f0f0f0")
        st.plotly_chart(fig, use_container_width=True)

    with right:
        st.subheader("Waiting List by Specialty")
        spec_wl = avail1.groupby("Specialty")["WL_Total"].sum().sort_values(ascending=False).reset_index()
        fig2 = px.pie(spec_wl, values="WL_Total", names="Specialty", hole=0.42,
                      color_discrete_sequence=SPEC_COLORS, height=440)
        fig2.update_traces(textposition="outside", textinfo="percent+label")
        fig2.update_layout(showlegend=False, margin=dict(l=0,r=0,t=10,b=10))
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown("---")
    st.subheader("Specialty Status Distribution by Directorate")
    sd = df1.copy()
    sd["Status"] = sd["Status"].fillna("N/A")
    sc = sd.groupby(["Directorate","Status"]).size().reset_index(name="Count")
    fig3 = px.bar(sc, x="Directorate", y="Count", color="Status",
                  color_discrete_map={"Available":TEAL,"On Hold":AMBER,"N/A":"#D3D1C7"},
                  barmode="stack", height=360,
                  labels={"Count":"Specialty Slots","Directorate":""})
    fig3.update_layout(make_layout({"margin":dict(l=0,r=0,t=10,b=90),
                                    "legend":dict(orientation="h",y=1.05),
                                    "xaxis_tickangle":-40}))
    fig3.update_yaxes(gridcolor="#f0f0f0")
    st.plotly_chart(fig3, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2  WAITING LIST
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.header("Waiting List Deep Dive")

    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Total Patients",           fmt(avail1["WL_Total"].sum()))
    c2.metric("Booked within 36 days",    fmt(avail1["WL_Booked36"].sum()))
    c3.metric("Unbooked",                 fmt(avail1["WL_Unbooked36"].sum()))
    c4.metric("Non-Scheduled (shortage)", fmt(avail1["WL_NonSched"].sum()))

    st.markdown("---")
    left, right = st.columns(2)

    with left:
        st.subheader("Avg Days to 2nd Slot — by Specialty")
        ws1 = avail1.groupby("Specialty")["Days_2nd_Slot"].mean().sort_values(ascending=False).reset_index()
        ws1.columns = ["Specialty","Days"]
        if comparing:
            ws2 = avail2.groupby("Specialty")["Days_2nd_Slot"].mean().reset_index()
            ws2.columns = ["Specialty","Days2"]
            ws  = ws1.merge(ws2, on="Specialty", how="left")
            ws  = ws.sort_values("Days", ascending=True)
            fig = go.Figure()
            fig.add_trace(go.Bar(y=ws["Specialty"], x=ws["Days"],
                                 name=label1, orientation="h", marker_color=TEAL, text=ws["Days"].round(1),
                                 texttemplate="%{text:.1f}d", textposition="outside"))
            fig.add_trace(go.Bar(y=ws["Specialty"], x=ws["Days2"],
                                 name=label2, orientation="h", marker_color=BLUE, opacity=0.7))
            fig.update_layout(make_layout({"barmode":"group","height":420,
                                           "margin":dict(l=0,r=60,t=10,b=30),
                                           "legend":dict(orientation="h",y=1.05)}))
        else:
            ws1 = ws1.sort_values("Days", ascending=True)
            fig = px.bar(ws1, x="Days", y="Specialty", orientation="h", color="Days",
                         color_continuous_scale=[[0,"#E1F5EE"],[0.5,AMBER],[1,CORAL]],
                         text="Days", height=420,
                         labels={"Days":"Avg Days","Specialty":""})
            fig.update_coloraxes(showscale=False)
            fig.update_traces(texttemplate="%{text:.1f}d", textposition="outside")
            fig.update_layout(make_layout({"margin":dict(l=0,r=60,t=10,b=30)}))
        fig.update_xaxes(gridcolor="#f0f0f0")
        st.plotly_chart(fig, use_container_width=True)

    with right:
        st.subheader("WL Volume vs Booking Rate")
        sc2 = avail1.groupby("Specialty").agg(
            WL_Total=("WL_Total","sum"),
            WL_Booked36=("WL_Booked36","sum"),
            Total_Surg=("Total_Surg","sum"),
        ).reset_index()
        sc2["Booked_Rate"] = (sc2["WL_Booked36"] / sc2["WL_Total"] * 100).round(1)
        sc2 = sc2.dropna(subset=["Booked_Rate"])
        fig2 = px.scatter(sc2, x="WL_Total", y="Booked_Rate", size="Total_Surg",
                          color="Specialty", hover_name="Specialty",
                          color_discrete_sequence=SPEC_COLORS, height=420,
                          labels={"WL_Total":"Total WL","Booked_Rate":"% Booked 36d","Total_Surg":"Surgeries"})
        fig2.update_layout(make_layout({"showlegend":False}))
        fig2.update_xaxes(gridcolor="#f0f0f0"); fig2.update_yaxes(gridcolor="#f0f0f0")
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown("---")
    st.subheader("Waiting List Heatmap — Directorate × Specialty")
    heat = avail1.groupby(["Directorate","Specialty"])["WL_Total"].sum().reset_index()
    hpiv = heat.pivot(index="Directorate", columns="Specialty", values="WL_Total").fillna(0)
    fig3 = px.imshow(hpiv, color_continuous_scale=[[0,"#FFFFFF"],[0.3,"#9FE1CB"],[1,TEAL]],
                     aspect="auto", height=520, labels=dict(color="Patients"), text_auto=True)
    fig3.update_layout(margin=dict(l=0,r=0,t=10,b=0))
    fig3.update_xaxes(tickangle=-35)
    st.plotly_chart(fig3, use_container_width=True)

    st.markdown("---")
    st.subheader("On Hold Specialties")
    oh = df1[df1["Status"]=="On Hold"][["Directorate","Hospital","Specialty","WL_Total","Days_2nd_Slot"]].copy()
    oh = oh.sort_values("Directorate")
    oh.columns = ["Directorate","Hospital","Specialty","WL Total","Days to Slot"]
    st.dataframe(oh.reset_index(drop=True), use_container_width=True, height=260)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3  SURGICAL ACTIVITY
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.header("Surgical Activity")

    tot_s = avail1["Total_Surg"].sum()
    el_s  = avail1["Elective_Surg"].sum()
    od_s  = avail1["OnDay_Surg"].sum()
    od_rt = od_s / tot_s * 100 if tot_s else 0

    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Total Surgeries",    fmt(tot_s),
              delta=fmt(avail2["Total_Surg"].sum()) + " (file 2)" if comparing else None)
    c2.metric("Elective",           fmt(el_s))
    c3.metric("One-Day (Day Case)", fmt(od_s))
    c4.metric("Day Case Rate",      f"{od_rt:.1f}%")

    st.markdown("---")
    left, right = st.columns(2)

    with left:
        st.subheader("Surgery Mix by Directorate")
        ds = avail1.groupby("Directorate").agg(
            Elective=("Elective_Surg","sum"),
            OneDay=("OnDay_Surg","sum"),
            Total=("Total_Surg","sum"),
        ).reset_index()
        ds["Other"] = (ds["Total"] - ds["Elective"] - ds["OneDay"]).clip(lower=0)
        ds = ds.sort_values("Elective")
        fig = go.Figure()
        fig.add_trace(go.Bar(y=ds["Directorate"], x=ds["Elective"],  name="Elective",
                             orientation="h", marker_color=TEAL))
        fig.add_trace(go.Bar(y=ds["Directorate"], x=ds["OneDay"],    name="One-Day",
                             orientation="h", marker_color=BLUE))
        fig.add_trace(go.Bar(y=ds["Directorate"], x=ds["Other"],     name="Emergency/Other",
                             orientation="h", marker_color=CORAL))
        fig.update_layout(make_layout({"barmode":"stack","height":460,
                                       "legend":dict(orientation="h",y=1.05)}))
        fig.update_xaxes(gridcolor="#f0f0f0", title="Surgeries")
        st.plotly_chart(fig, use_container_width=True)

    with right:
        st.subheader("Surgery Volume by Specialty")
        ss2 = avail1.groupby("Specialty").agg(
            Elective=("Elective_Surg","sum"), OneDay=("OnDay_Surg","sum"),
        ).reset_index().sort_values("Elective", ascending=False)
        fig2 = go.Figure()
        fig2.add_trace(go.Bar(x=ss2["Specialty"], y=ss2["Elective"], name="Elective", marker_color=TEAL))
        fig2.add_trace(go.Bar(x=ss2["Specialty"], y=ss2["OneDay"],   name="One-Day",  marker_color=BLUE))
        fig2.update_layout(make_layout({"barmode":"group","height":460,
                                        "margin":dict(l=0,r=0,t=10,b=90),
                                        "legend":dict(orientation="h",y=1.05),
                                        "xaxis_tickangle":-40}))
        fig2.update_yaxes(gridcolor="#f0f0f0", title="Surgeries")
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown("---")
    st.subheader("OR Sessions vs Surgeries — by Directorate")
    bub = avail1.groupby("Directorate").agg(
        OR_Sessions=("OR_Sessions","sum"),
        Total_Surg=("Total_Surg","sum"),
        WL_Total=("WL_Total","sum"),
    ).reset_index()
    fig3 = px.scatter(bub, x="OR_Sessions", y="Total_Surg", size="WL_Total",
                      color="Directorate", hover_name="Directorate", size_max=50,
                      color_discrete_sequence=SPEC_COLORS, height=400, text="Directorate",
                      labels={"OR_Sessions":"OR Sessions/week","Total_Surg":"Total Surgeries","WL_Total":"WL Size"})
    fig3.update_traces(textposition="top center", textfont_size=9)
    fig3.update_layout(make_layout({"showlegend":False}))
    fig3.update_xaxes(gridcolor="#f0f0f0"); fig3.update_yaxes(gridcolor="#f0f0f0")
    st.plotly_chart(fig3, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4  OR ROOMS
# ══════════════════════════════════════════════════════════════════════════════
with tab4:
    st.header("OR Rooms & Capacity")

    or_h = df1.groupby("Hospital").agg(
        NonEm_Func_ORs=("NonEm_Func_ORs","first"),
        NonFunc_ORs=("NonFunc_ORs","first"),
        Em_ORs=("Em_ORs","first"),
        Directorate=("Directorate","first"),
    ).reset_index()
    for c in ["NonEm_Func_ORs","NonFunc_ORs","Em_ORs"]:
        or_h[c] = pd.to_numeric(or_h[c], errors="coerce").fillna(0)

    tot_func  = or_h["NonEm_Func_ORs"].sum()
    tot_nfunc = or_h["NonFunc_ORs"].sum()
    tot_em    = or_h["Em_ORs"].sum()
    nf_rate   = tot_nfunc / (tot_func + tot_nfunc) * 100 if (tot_func + tot_nfunc) else 0

    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Functioning ORs (Non-Em)", fmt(tot_func))
    c2.metric("Non-Functioning ORs",       fmt(tot_nfunc))
    c3.metric("Non-Function Rate",         f"{nf_rate:.1f}%")
    c4.metric("Emergency ORs Reported",    fmt(tot_em))

    st.markdown("---")
    left, right = st.columns(2)

    with left:
        st.subheader("OR Rooms by Directorate")
        dir_or = or_h.groupby("Directorate")[["NonEm_Func_ORs","NonFunc_ORs","Em_ORs"]].sum()
        dir_or = dir_or.sort_values("NonEm_Func_ORs").reset_index()
        fig = go.Figure()
        fig.add_trace(go.Bar(y=dir_or["Directorate"], x=dir_or["NonEm_Func_ORs"],
                             name="Functioning",     orientation="h", marker_color=TEAL))
        fig.add_trace(go.Bar(y=dir_or["Directorate"], x=dir_or["NonFunc_ORs"],
                             name="Non-Functioning", orientation="h", marker_color=CORAL))
        fig.add_trace(go.Bar(y=dir_or["Directorate"], x=dir_or["Em_ORs"],
                             name="Emergency",       orientation="h", marker_color=AMBER))
        fig.update_layout(make_layout({"barmode":"stack","height":480,
                                       "legend":dict(orientation="h",y=1.05)}))
        fig.update_xaxes(gridcolor="#f0f0f0", title="OR Rooms")
        st.plotly_chart(fig, use_container_width=True)

    with right:
        st.subheader("Non-Functioning OR Rate by Directorate")
        dir_or["Total"] = dir_or["NonEm_Func_ORs"] + dir_or["NonFunc_ORs"]
        dir_or["NF_Rate"] = (dir_or["NonFunc_ORs"] / dir_or["Total"] * 100).round(1)
        dir_or2 = dir_or[dir_or["Total"] > 0].sort_values("NF_Rate", ascending=False)
        fig2 = px.bar(dir_or2, x="NF_Rate", y="Directorate", orientation="h",
                      color="NF_Rate",
                      color_continuous_scale=[[0,"#E1F5EE"],[0.4,AMBER],[1,CORAL]],
                      text="NF_Rate", height=480,
                      labels={"NF_Rate":"Non-Function Rate (%)","Directorate":""})
        fig2.update_coloraxes(showscale=False)
        fig2.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
        fig2.update_layout(make_layout({"margin":dict(l=0,r=60,t=10,b=30)}))
        fig2.update_xaxes(gridcolor="#f0f0f0")
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown("---")
    st.subheader("Hospital-level OR Room Detail")
    sort_opt = st.selectbox("Sort by",
        ["Functioning ORs","Non-Functioning","Emergency ORs","Non-Func Rate%","Hospital"])
    or_det = or_h[or_h["NonEm_Func_ORs"] > 0].copy()
    or_det["NF_Rate"] = (or_det["NonFunc_ORs"] / (or_det["NonEm_Func_ORs"] + or_det["NonFunc_ORs"]) * 100).round(1)
    or_show = or_det[["Directorate","Hospital","NonEm_Func_ORs","NonFunc_ORs","Em_ORs","NF_Rate"]].copy()
    or_show.columns = ["Directorate","Hospital","Functioning ORs","Non-Functioning","Emergency ORs","Non-Func Rate%"]
    or_show = or_show.sort_values(sort_opt, ascending=False)
    st.dataframe(or_show.reset_index(drop=True), use_container_width=True, height=380)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 5  HOSPITAL EXPLORER
# ══════════════════════════════════════════════════════════════════════════════
with tab5:
    st.header("Hospital Explorer")

    ca, cb = st.columns([1,3])
    with ca:
        dir_choice = st.selectbox("Directorate",
            sorted(df1_raw["Directorate"].dropna().unique().tolist()), key="hex_dir")
    with cb:
        hosp_opts   = sorted(df1_raw[df1_raw["Directorate"]==dir_choice]["Hospital"].unique().tolist())
        hosp_choice = st.selectbox("Hospital", hosp_opts, key="hex_hosp")

    hdf = df1_raw[df1_raw["Hospital"]==hosp_choice].copy()
    for c in NUM_COLS:
        hdf[c] = pd.to_numeric(hdf[c], errors="coerce")
    hav = hdf[hdf["Status"]=="Available"]

    st.markdown("---")
    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric("Total WL",        fmt(hav["WL_Total"].sum()))
    c2.metric("Total Surgeries", fmt(hav["Total_Surg"].sum()))
    c3.metric("Available Specs", str(len(hav)))
    c4.metric("On Hold Specs",   str((hdf["Status"]=="On Hold").sum()))
    c5.metric("Functioning ORs", fmt(hdf["NonEm_Func_ORs"].iloc[0] if len(hdf) else None))

    st.markdown("---")
    left, right = st.columns(2)

    with left:
        st.subheader("Waiting List by Specialty")
        if len(hav) > 0:
            fig = px.bar(hav.sort_values("WL_Total"), x="WL_Total", y="Specialty",
                         orientation="h", color="WL_Total",
                         color_continuous_scale=[[0,"#E1F5EE"],[1,TEAL]],
                         text="WL_Total", height=380,
                         labels={"WL_Total":"Patients","Specialty":""})
            fig.update_coloraxes(showscale=False)
            fig.update_traces(texttemplate="%{text:.0f}", textposition="outside")
            fig.update_layout(make_layout({"margin":dict(l=0,r=50,t=10,b=30)}))
            fig.update_xaxes(gridcolor="#f0f0f0")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No available specialties for this hospital.")

    with right:
        st.subheader("Days to Next Slot by Specialty")
        wh = hav.dropna(subset=["Days_2nd_Slot"]).sort_values("Days_2nd_Slot", ascending=False)
        if len(wh) > 0:
            fig2 = px.bar(wh, x="Days_2nd_Slot", y="Specialty", orientation="h",
                          color="Days_2nd_Slot",
                          color_continuous_scale=[[0,"#E1F5EE"],[0.5,AMBER],[1,CORAL]],
                          text="Days_2nd_Slot", height=380,
                          labels={"Days_2nd_Slot":"Days","Specialty":""})
            fig2.update_coloraxes(showscale=False)
            fig2.update_traces(texttemplate="%{text:.0f}d", textposition="outside")
            fig2.update_layout(make_layout({"margin":dict(l=0,r=50,t=10,b=30)}))
            fig2.update_xaxes(gridcolor="#f0f0f0")
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.info("No appointment slot data for this hospital.")

    st.markdown("---")
    st.subheader("Full Specialty Detail")
    disp_cols = ["Specialty","Status","WL_Total","WL_New","WL_Booked36","WL_Unbooked36",
                 "Elective_Surg","OnDay_Surg","Total_Surg","OR_Sessions","Days_2nd_Slot"]
    disp = hdf[disp_cols].copy()
    disp.columns = ["Specialty","Status","WL Total","New Patients","Booked 36d","Unbooked",
                    "Elective","One-Day","Total Surg","OR Sessions/wk","Days to Slot"]
    disp["Status"] = disp["Status"].fillna("N/A")

    def hl(row):
        if row["Status"] == "Available": return ["background-color:#E1F5EE"]*len(row)
        if row["Status"] == "On Hold":   return ["background-color:#FAEEDA"]*len(row)
        return [""]*len(row)

    st.dataframe(disp.reset_index(drop=True).style.apply(hl, axis=1),
                 use_container_width=True, height=440)

    st.markdown("---")
    st.subheader("Compare with Directorate Average")
    if len(hav) > 0:
        dir_data = df1_raw[
            (df1_raw["Directorate"]==dir_choice) & (df1_raw["Status"]=="Available")
        ].copy()
        for c in NUM_COLS:
            dir_data[c] = pd.to_numeric(dir_data[c], errors="coerce")
        metric_map = {
            "Avg WL Patients":    "WL_Total",
            "Avg Total Surgeries":"Total_Surg",
            "Avg Days to Slot":   "Days_2nd_Slot",
            "Avg OR Sessions/wk": "OR_Sessions",
        }
        sel_m = st.selectbox("Metric", list(metric_map.keys()))
        col_n = metric_map[sel_m]
        davg  = dir_data.groupby("Specialty")[col_n].mean().reset_index()
        davg.columns = ["Specialty","Dir_Avg"]
        hvals = hav[["Specialty",col_n]].copy()
        hvals.columns = ["Specialty","Hospital"]
        comp  = davg.merge(hvals, on="Specialty").dropna()
        if len(comp) > 0:
            fig3 = go.Figure()
            fig3.add_trace(go.Bar(x=comp["Specialty"], y=comp["Hospital"],
                                  name=hosp_choice, marker_color=TEAL))
            fig3.add_trace(go.Bar(x=comp["Specialty"], y=comp["Dir_Avg"],
                                  name=f"{dir_choice} avg", marker_color=GRAY, opacity=0.65))
            fig3.update_layout(make_layout({"barmode":"group","height":360,
                                            "margin":dict(l=0,r=0,t=10,b=70),
                                            "legend":dict(orientation="h",y=1.05),
                                            "xaxis_tickangle":-35,
                                            "yaxis_title":sel_m}))
            fig3.update_yaxes(gridcolor="#f0f0f0")
            st.plotly_chart(fig3, use_container_width=True)
        else:
            st.info("Not enough shared specialties for comparison.")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 6  FILE COMPARISON  (only shown when two files are loaded)
# ══════════════════════════════════════════════════════════════════════════════
if comparing and tab6 is not None:
    with tab6:
        st.header("File Comparison")
        st.caption(f"**File 1:** {label1}   |   **File 2:** {label2}")
        st.markdown("*Positive delta = File 1 is higher. Negative = File 2 is higher.*")

        # ── Summary KPI comparison ────────────────────────────────────────────
        st.subheader("Key Metric Comparison")
        metrics = {
            "Total WL Patients":   ("WL_Total",     avail1, avail2),
            "Total Surgeries":     ("Total_Surg",   avail1, avail2),
            "Elective Surgeries":  ("Elective_Surg",avail1, avail2),
            "One-Day Surgeries":   ("OnDay_Surg",   avail1, avail2),
            "Avg Days to Slot":    ("Days_2nd_Slot",avail1, avail2),
            "Hospitals Reporting": (None,           avail1, avail2),
        }
        cols = st.columns(3)
        for idx, (label, (col, a1, a2)) in enumerate(metrics.items()):
            if col is None:
                v1 = a1["Hospital"].nunique()
                v2 = a2["Hospital"].nunique()
            else:
                v1 = a1[col].mean() if "Avg" in label else a1[col].sum()
                v2 = a2[col].mean() if "Avg" in label else a2[col].sum()
            d = v1 - v2
            fmt_v = f"{v1:.1f}" if "Avg" in label else fmt(v1)
            fmt_d = f"{d:+.1f}" if "Avg" in label else f"{d:+,.0f}"
            cols[idx % 3].metric(label, fmt_v, delta=fmt_d)

        st.markdown("---")

        # ── WL by specialty side-by-side ──────────────────────────────────────
        st.subheader("Waiting List by Specialty")
        sp1 = avail1.groupby("Specialty")["WL_Total"].sum().reset_index(name="File1")
        sp2 = avail2.groupby("Specialty")["WL_Total"].sum().reset_index(name="File2")
        sp  = sp1.merge(sp2, on="Specialty", how="outer").fillna(0)
        sp["Delta"] = sp["File1"] - sp["File2"]
        sp = sp.sort_values("Delta", ascending=True)

        fig_sp = go.Figure()
        fig_sp.add_trace(go.Bar(y=sp["Specialty"], x=sp["File1"],
                                name=label1, orientation="h", marker_color=TEAL))
        fig_sp.add_trace(go.Bar(y=sp["Specialty"], x=sp["File2"],
                                name=label2, orientation="h", marker_color=BLUE, opacity=0.7))
        fig_sp.update_layout(make_layout({"barmode":"group","height":440,
                                          "legend":dict(orientation="h",y=1.05)}))
        fig_sp.update_xaxes(gridcolor="#f0f0f0", title="Patients")
        st.plotly_chart(fig_sp, use_container_width=True)

        # ── Delta waterfall by Directorate ────────────────────────────────────
        st.subheader("WL Change by Directorate (File 1 minus File 2)")
        d1 = avail1.groupby("Directorate")["WL_Total"].sum().reset_index(name="F1")
        d2 = avail2.groupby("Directorate")["WL_Total"].sum().reset_index(name="F2")
        dd = d1.merge(d2, on="Directorate", how="outer").fillna(0)
        dd["Delta"] = dd["F1"] - dd["F2"]
        dd = dd.sort_values("Delta")
        colors_d = [TEAL if v >= 0 else CORAL for v in dd["Delta"]]
        fig_dd = go.Figure(go.Bar(
            x=dd["Delta"], y=dd["Directorate"], orientation="h",
            marker_color=colors_d,
            text=dd["Delta"].apply(lambda v: f"{v:+,.0f}"),
            textposition="outside",
        ))
        fig_dd.update_layout(make_layout({"height":440,"margin":dict(l=0,r=60,t=10,b=30)}))
        fig_dd.update_xaxes(gridcolor="#f0f0f0", title="Change in Patients (File 1 − File 2)")
        fig_dd.add_vline(x=0, line_width=1.5, line_color=GRAY)
        st.plotly_chart(fig_dd, use_container_width=True)

        # ── Avg wait days comparison ──────────────────────────────────────────
        st.subheader("Avg Days to 2nd Slot — Specialty Comparison")
        w1 = avail1.groupby("Specialty")["Days_2nd_Slot"].mean().reset_index(name="F1")
        w2 = avail2.groupby("Specialty")["Days_2nd_Slot"].mean().reset_index(name="F2")
        ww = w1.merge(w2, on="Specialty", how="outer")
        ww["Delta"] = (ww["F1"] - ww["F2"]).round(1)
        ww = ww.sort_values("Delta")
        colors_w = [TEAL if v >= 0 else GREEN for v in ww["Delta"].fillna(0)]
        fig_ww = go.Figure(go.Bar(
            x=ww["Delta"], y=ww["Specialty"], orientation="h",
            marker_color=colors_w,
            text=ww["Delta"].apply(lambda v: f"{v:+.1f}d" if not pd.isna(v) else ""),
            textposition="outside",
        ))
        fig_ww.update_layout(make_layout({"height":420,"margin":dict(l=0,r=70,t=10,b=30)}))
        fig_ww.update_xaxes(gridcolor="#f0f0f0", title="Change in Days (File 1 − File 2)")
        fig_ww.add_vline(x=0, line_width=1.5, line_color=GRAY)
        st.plotly_chart(fig_ww, use_container_width=True)

        # ── Hospitals in one file but not the other ────────────────────────────
        st.subheader("Hospital Coverage Differences")
        h1_set = set(df1_raw["Hospital"].unique())
        h2_set = set(df2_raw["Hospital"].unique())
        only1  = sorted(h1_set - h2_set)
        only2  = sorted(h2_set - h1_set)
        both   = len(h1_set & h2_set)
        ca2, cb2, cc2 = st.columns(3)
        ca2.metric(f"Only in File 1",  str(len(only1)))
        cb2.metric("In Both Files",    str(both))
        cc2.metric(f"Only in File 2",  str(len(only2)))
        if only1 or only2:
            exp = st.expander("See which hospitals differ")
            with exp:
                if only1:
                    st.write(f"**Only in {label1}:** " + ", ".join(only1))
                if only2:
                    st.write(f"**Only in {label2}:** " + ", ".join(only2))

        # ── Full side-by-side table ───────────────────────────────────────────
        st.subheader("Full Directorate × Specialty Comparison Table")
        grp_cols = ["Directorate","Specialty"]
        agg_cols = ["WL_Total","Total_Surg","Elective_Surg","Days_2nd_Slot"]
        t1 = avail1.groupby(grp_cols)[agg_cols].sum().reset_index()
        t2 = avail2.groupby(grp_cols)[agg_cols].sum().reset_index()
        tm = t1.merge(t2, on=grp_cols, suffixes=("_f1","_f2"), how="outer").fillna(0)
        for c in agg_cols:
            tm[f"{c}_delta"] = (tm[f"{c}_f1"] - tm[f"{c}_f2"]).round(1)
        show_cols2 = (
            grp_cols +
            [f"{c}_f1" for c in agg_cols] +
            [f"{c}_f2" for c in agg_cols] +
            [f"{c}_delta" for c in agg_cols]
        )
        tm_show = tm[show_cols2].copy()
        tm_show.columns = (
            grp_cols +
            [f"{c} (F1)" for c in ["WL","Surgeries","Elective","Days"]] +
            [f"{c} (F2)" for c in ["WL","Surgeries","Elective","Days"]] +
            [f"Δ {c}" for c in ["WL","Surgeries","Elective","Days"]]
        )
        st.dataframe(tm_show.reset_index(drop=True), use_container_width=True, height=400)

