# tat_gui_app.py — BI-grade TAT dashboard (Streamlit + Plotly)
# Polished visuals, reliable heatmap, no double-counting in Hold stacks,
# graceful handling of single-category cases, fixed-path Excel load.

import io
import os
import sys
from datetime import timedelta, date
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

# ---------- Page ----------
st.set_page_config(page_title="TAT Analysis GUI", layout="wide")
st.markdown(
    """
    <style>
      .block-container {max-width: 1420px; padding-top: 0.6rem;}
      h1, h2, h3 {font-weight: 700;}
      .stMetric {background: #f7f9fc; border-radius: 12px; padding: 8px 12px;}
      .small-note {color:#667085;font-size:0.9rem;}
    </style>
    """,
    unsafe_allow_html=True,
)
st.title("📊 TAT Analysis & Reporting")

# ---------- Load (fixed path; no uploader) ----------
st.sidebar.header("1) Data source")

def resolve_xlsx_path():
    # CLI: streamlit run tat_gui_app.py -- --xlsx /full/path.xlsx
    for a in sys.argv:
        if a.startswith("--xlsx="):
            p = a.split("=", 1)[1].strip().strip('"').strip("'")
            p = Path(p).expanduser()
            if p.exists():
                return str(p)
    # ENV: export TAT_XLSX=/full/path.xlsx
    envp = os.getenv("TAT_XLSX")
    if envp:
        p = Path(envp).expanduser()
        if p.exists():
            return str(p)
    # Defaults
    here = Path(__file__).parent
    for p in [
        here / "data" / "TAT Analysis by CompleteDate.xlsx",
        here / "TAT Analysis by CompletedDate.xlsx",
        Path.home() / "tat_gui" / "data" / "TAT Analysis by CompleteDate.xlsx",
    ]:
        if p.exists():
            return str(p)
    return None

@st.cache_data(show_spinner=False)
def load_sheet(file_path, sheet="AllTasks"):
    df = pd.read_excel(file_path, sheet_name=sheet)
    df.columns = [str(c).strip() for c in df.columns]
    # Coerce likely date columns
    for c in df.columns:
        lc = c.lower()
        if "date" in lc or "hit" in lc or "target" in lc or "complete" in lc:
            df[c] = pd.to_datetime(df[c], errors="coerce")
    return df

xlsx_path = resolve_xlsx_path()
if not xlsx_path:
    st.error(
        "Excel not found.\n"
        "Put it at `~/tat_gui/data/TAT Analysis by CompleteDate.xlsx`, or\n"
        "`streamlit run tat_gui_app.py -- --xlsx /full/path.xlsx`, or\n"
        "`export TAT_XLSX=/full/path.xlsx` then run the app."
    )
    st.stop()

df = load_sheet(xlsx_path, "AllTasks")
st.caption(f"Loaded: **{xlsx_path}**  (sheet: AllTasks)")

# ---------- Column detection ----------
def first_existing(names):
    for n in names:
        if n in df.columns:
            return n
    return None

col_sector   = first_existing(["Sector"])
col_owner    = first_existing(["Task Owner", "Owner", "Engineer", "Assignee"])
col_status   = first_existing(["Redash Status", "Task Source Status", "Status"])
col_priority = first_existing(["Priority"])
col_hit      = first_existing(["Hit Date"])
col_target   = first_existing(["Target Date", "Initial Target Date"])
col_complete = first_existing(["Task Complete Date", "Validation Complete Date"])

col_tat_pref = [
    "Total Task TAT",
    "NoSTR Total Task TAT",
    "MinusHold Total Task TAT",
    "NoHold NoSTR Total Adjusted TAT",
]
col_tat_total = first_existing(col_tat_pref)

# On-Time (prefer explicit OTD flag)
col_otd_flag = first_existing(["OTD PassFail", "OTD_PassFail", "OnTime"])
if col_otd_flag and df[col_otd_flag].notna().any():
    norm = {1: True, "1": True, "Y": True, "YES": True, "True": True,
            0: False, "0": False, "N": False, "NO": False, "False": False}
    s = df[col_otd_flag].map(norm)
    df["_OnTime"] = s.where(~s.isna(), df[col_otd_flag]).astype("boolean", errors="ignore")
elif col_complete and col_target:
    df["_OnTime"] = (df[col_complete] <= df[col_target]).astype("boolean")
else:
    df["_OnTime"] = pd.Series([pd.NA]*len(df), dtype="boolean")

# Hold columns (exclude "Total Hold TAT" from stacked components)
all_hold_cols = [c for c in df.columns if c.endswith("Hold TAT") or " Hold TAT" in c]
hold_stack_cols = [c for c in all_hold_cols if c.lower() != "total hold tat"]
for c in all_hold_cols + ([col_tat_total] if col_tat_total else []):
    if c in df.columns:
        df[c] = pd.to_numeric(df[c], errors="coerce")

# ---------- Filters ----------
st.sidebar.header("2) Filters")
if st.sidebar.button("Reset filters"):
    st.rerun()

def ms(col, label):
    if col and col in df.columns and df[col].notna().any():
        opts = sorted(map(str, df[col].dropna().unique()))
        return st.sidebar.multiselect(label, opts, default=[])
    return []

sel_sector   = ms(col_sector,   "Sector")
sel_owner    = ms(col_owner,    "Owner / Engineer")
sel_status   = ms(col_status,   "Status")
sel_priority = ms(col_priority, "Priority")

def date_range_for(col, label):
    if not col: return None
    s = pd.to_datetime(df[col], errors="coerce").dropna()
    if s.empty: return None
    return st.sidebar.date_input(label, (s.min().date(), s.max().date()))

hit_range  = date_range_for(col_hit, "Hit Date")
tgt_range  = date_range_for(col_target, "Target Date")
comp_range = date_range_for(col_complete, "Complete Date")

# Presets for Complete Date
if col_complete and st.sidebar.button("Last 90 days"):
    mx = df[col_complete].dropna().max().date()
    st.session_state["_comp_preset"] = (mx - timedelta(days=90), mx)
if "_comp_preset" in st.session_state:
    comp_range = st.session_state["_comp_preset"]

# Display options
st.sidebar.header("2a) Display")
top_n = st.sidebar.slider("Top N", 5, 50, 10, 1)
value_mode = st.sidebar.radio("Bar value", ["Count", "Percent of total"], horizontal=True, index=0)
rotate_labels = st.sidebar.checkbox("Rotate X labels 45°", True)

# ---------- Apply filters ----------
fdf = df.copy()

def filt_in(col, selected):
    global fdf
    if col and selected:
        fdf = fdf[fdf[col].astype(str).isin(selected)]

for c, sel in [(col_sector, sel_sector), (col_owner, sel_owner), (col_status, sel_status), (col_priority, sel_priority)]:
    filt_in(c, sel)

def filt_date(col, dr):
    global fdf
    if not col or not dr: return
    start = pd.to_datetime(dr[0])
    end   = pd.to_datetime(dr[1]) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
    s = pd.to_datetime(fdf[col], errors="coerce")
    fdf = fdf[(s >= start) & (s <= end)]

filt_date(col_hit, hit_range)
filt_date(col_target, tgt_range)
filt_date(col_complete, comp_range)

st.caption(f"Showing **{len(fdf):,}** of **{len(df):,}** rows after filters.")

# ---------- Helpers ----------
def style_fig(fig, height=460):
    fig.update_layout(template="plotly_white", height=height,
                      margin=dict(l=16, r=16, t=56, b=80))
    if rotate_labels:
        fig.update_xaxes(tickangle=45)
    return fig

def add_labels(fig, pct=False):
    txt = "%{y:.1f}%" if pct else "%{y}"
    fig.update_traces(texttemplate=txt, textposition="outside", cliponaxis=False)
    return fig

def apply_hover_format(fig, is_pct=False):
    if is_pct:
        fig.update_traces(hovertemplate="%{x}<br>%{y:.1f}%<extra></extra>")
    else:
        fig.update_traces(hovertemplate="%{x}<br>%{y:,}<extra></extra>")
    return fig

def bar_rank(frame, group_col, title):
    """Robust ranked bar: KPI fallback when only one category."""
    if not group_col or frame.empty:
        st.info("No data for current filters."); return
    g = frame.groupby(group_col).size().reset_index(name="Count").sort_values("Count", ascending=False)
    if g.empty:
        st.info("No data for current filters."); return

    if g[group_col].nunique() == 1:
        val = int(g["Count"].iloc[0])
        st.markdown(f"**{title}**")
        st.metric(label=str(g[group_col].iloc[0]), value=f"{val:,}")
        return

    if value_mode == "Percent of total":
        total = g["Count"].sum() or 1
        g["Percent"] = g["Count"] / total * 100
        ycol, is_pct = "Percent", True
    else:
        ycol, is_pct = "Count", False

    top = g.head(top_n)
    # Horizontal bars for long labels
    horiz = top[group_col].astype(str).str.len().mean() > 12
    fig = px.bar(
        top,
        x=(ycol if horiz else group_col),
        y=(group_col if horiz else ycol),
        orientation="h" if horiz else "v",
        title=title,
    )
    fig = add_labels(fig, is_pct)
    fig = apply_hover_format(fig, is_pct)
    st.plotly_chart(style_fig(fig), use_container_width=True)

# ---------- KPI cards ----------
k1, k2, k3, k4 = st.columns(4)
with k1:
    st.metric("Total Tasks", f"{len(fdf):,}")
with k2:
    st.metric("Avg TAT (days)", f"{fdf[col_tat_total].mean():.1f}" if col_tat_total and fdf[col_tat_total].notna().any() else "—")
with k3:
    st.metric("On-Time Rate", f"{(fdf['_OnTime']==True).mean()*100:.1f}%" if fdf['_OnTime'].notna().any() else "—")
with k4:
    if col_priority:
        pct_high = (fdf[col_priority].astype(str).str.contains("Urgent|High", case=False, na=False)).mean()*100
        st.metric("High Priority %", f"{pct_high:.1f}%")
    else:
        st.metric("High Priority %", "—")

# ---------- Tabs ----------
tabs = st.tabs([
    "Overview","By Category","Timeline","Spread & Outliers",
    "SLA Mix","Hold Breakdown","OTD Analysis","Heatmap",
    "Treemap","Pareto","Pivot Builder","Data & Export"
])
(tab_overview, tab_bycat, tab_timeline, tab_spread,
 tab_sla, tab_hold, tab_otd, tab_heat,
 tab_tree, tab_pareto, tab_pivot, tab_data) = tabs

# ---- Overview ----
with tab_overview:
    st.subheader("Highlights")
    pref_group = col_owner or col_sector or col_status or col_priority
    bar_rank(fdf, pref_group, f"Top {('Owners' if pref_group==col_owner else 'Groups')} by Volume")
    if col_sector and pref_group != col_sector:
        bar_rank(fdf, col_sector, "Sectors by Task Volume")

# ---- By Category ----
with tab_bycat:
    c1, c2 = st.columns(2)
    if col_sector:
        with c1: bar_rank(fdf, col_sector, "Sectors by Task Volume")
    if col_owner:
        with c2: bar_rank(fdf, col_owner, "Owners by Task Volume")

    # Average TAT by group (graceful with single category)
    if col_tat_total:
        g = col_owner or col_sector
        if g:
            agg = (
                fdf.groupby(g, as_index=False)[col_tat_total]
                   .mean()
                   .sort_values(col_tat_total, ascending=False)
            )
            if agg[g].nunique() == 1:
                st.markdown(f"**Average {col_tat_total} by {g}**")
                st.metric(label=str(agg[g].iloc[0]), value=f"{agg[col_tat_total].iloc[0]:.1f} days")
            else:
                fig = px.bar(
                    agg.head(top_n), x=g, y=col_tat_total,
                    title=f"Average {col_tat_total} by {g}"
                )
                fig = add_labels(fig)
                fig = apply_hover_format(fig, is_pct=False)
                st.plotly_chart(style_fig(fig), use_container_width=True)

# ---- Timeline ----
with tab_timeline:
    base = col_complete or col_hit or col_target
    if base and fdf[base].notna().any():
        tmp = fdf.dropna(subset=[base]).copy()
        tmp["Month"] = pd.to_datetime(tmp[base]).dt.to_period("M").dt.to_timestamp()
        ts = tmp.groupby("Month").size().reset_index(name="Completed Tasks")
        fig1 = px.line(ts, x="Month", y="Completed Tasks", markers=True, title="Completed Tasks over Time")
        st.plotly_chart(style_fig(fig1, height=420), use_container_width=True)
        if col_tat_total:
            tat_ts = tmp.groupby("Month", as_index=False)[col_tat_total].mean()
            fig2 = px.line(tat_ts, x="Month", y=col_tat_total, markers=True, title=f"Average {col_tat_total} over Time")
            st.plotly_chart(style_fig(fig2, height=420), use_container_width=True)
    else:
        st.info("No date column available for a timeline.")

# ---- Spread & Outliers ----
with tab_spread:
    if col_tat_total and fdf[col_tat_total].notna().any():
        g = col_owner or col_sector
        if g and fdf[g].dropna().nunique() > 1:
            sub = fdf[[g, col_tat_total]].dropna()
            fig = px.box(sub, x=g, y=col_tat_total, points="suspectedoutliers",
                         title=f"{col_tat_total} Distribution by {g}")
        else:
            fig = px.histogram(fdf, x=col_tat_total, nbins=40,
                               title=f"Distribution of {col_tat_total}")
        st.plotly_chart(style_fig(fig, height=520), use_container_width=True)
    else:
        st.info("No TAT column found for spread analysis.")

# ---- SLA Mix ----
with tab_sla:
    if col_tat_total and fdf[col_tat_total].notna().any():
        def sla_bucket(x):
            if pd.isna(x): return "Unknown"
            if x <= 5: return "≤5"
            if x <= 10: return "6–10"
            if x <= 20: return "11–20"
            return ">20"
        tmp = fdf.copy()
        tmp["_SLA"] = tmp[col_tat_total].apply(sla_bucket)
        g = col_owner or col_sector
        if g:
            sla = tmp.groupby([g, "_SLA"]).size().reset_index(name="Count")
            totals = sla.groupby(g)["Count"].transform("sum")
            sla["Pct"] = (sla["Count"] / totals) * 100
            sla["_SLA"] = pd.Categorical(sla["_SLA"], categories=["≤5", "6–10", "11–20", ">20", "Unknown"], ordered=True)
            fig = px.bar(sla, x=g, y="Pct", color="_SLA", barmode="stack",
                         title=f"SLA Bucket Mix by {g}")
            st.plotly_chart(style_fig(fig), use_container_width=True)
    else:
        st.info("No TAT data to compute SLA buckets.")

# ---- Hold Breakdown ----
with tab_hold:
    if hold_stack_cols:
        g = col_owner or col_sector
        if g:
            # drop empty hold columns for a clean legend
            nonempty = [c for c in hold_stack_cols if fdf[c].notna().any() and fdf[c].abs().sum() > 0]
            if not nonempty:
                st.info("No Hold data available for current filters.")
            else:
                agg = fdf.groupby(g)[nonempty].mean(numeric_only=True).reset_index()
                melted = agg.melt(id_vars=[g], value_vars=nonempty,
                                  var_name="Hold Type", value_name="Avg Days")
                fig = px.bar(melted, x=g, y="Avg Days", color="Hold Type",
                             barmode="stack", title=f"Average Hold Time by {g}")
                st.plotly_chart(style_fig(fig), use_container_width=True)
        if "Total Hold TAT" in all_hold_cols and g:
            tot = fdf.groupby(g)["Total Hold TAT"].mean().reset_index()
            fig2 = px.bar(tot, x=g, y="Total Hold TAT", title=f"Total Hold TAT (Avg) by {g}")
            fig2 = add_labels(fig2)
            fig2 = apply_hover_format(fig2, is_pct=False)
            st.plotly_chart(style_fig(fig2), use_container_width=True)
    else:
        st.info("No Hold columns detected.")

# ---- OTD Analysis ----
with tab_otd:
    if fdf["_OnTime"].notna().any():
        g = col_owner or col_sector
        if g:
            t = (
                fdf.dropna(subset=["_OnTime"])
                   .groupby(g)["_OnTime"]
                   .mean()
                   .mul(100)
                   .reset_index(name="On-Time %")
                   .sort_values("On-Time %", ascending=False)
                   .head(top_n)
            )
            fig = px.bar(t, x=g, y="On-Time %", title=f"On-Time % by {g}")
            fig = add_labels(fig, True)
            fig = apply_hover_format(fig, is_pct=True)
            st.plotly_chart(style_fig(fig), use_container_width=True)
    else:
        st.info("On-Time status not available.")

# ---- Heatmap (reliable) ----
with tab_heat:
    base = col_complete or col_hit or col_target
    g = col_owner or col_sector
    if base and g and fdf[base].notna().any():
        tdf = fdf.dropna(subset=[base]).copy()
        tdf["Month"] = pd.to_datetime(tdf[base]).dt.to_period("M").dt.to_timestamp()
        mat = tdf.groupby([g, "Month"]).size().unstack("Month", fill_value=0).sort_index()
        if mat.empty or (mat.values.sum() == 0):
            st.info("No data to display for the selected filters.")
        else:
            fig = px.imshow(
                mat,
                aspect="auto",
                color_continuous_scale="Blues",
                labels=dict(color="Count"),
                title=f"Volume Heatmap: {g} × Month",
            )
            fig.update_yaxes(autorange="reversed")
            fig.update_traces(hovertemplate="Month=%{x}<br>"+g+"=%{y}<br>Count=%{z}<extra></extra>")
            st.plotly_chart(style_fig(fig, height=520), use_container_width=True)
    else:
        st.info("Need a date column and a grouping for a heatmap.")

# ---- Treemap (hardened) ----
with tab_tree:
    g = col_owner or col_sector
    if not g:
        st.info("No grouping column available.")
    elif fdf.empty:
        st.info("No data for current filters.")
    else:
        tree = (
            fdf.groupby(g, dropna=False)
               .size()
               .reset_index(name="Count")
        )
        tree = tree[tree["Count"] > 0]
        if tree.empty:
            st.info("No non-zero groups to display.")
        else:
            try:
                fig = px.treemap(
                    tree,
                    path=[px.Constant("All"), g],  # root node prevents ShapeError
                    values="Count",
                    title=f"Treemap of Volume by {g}",
                )
                st.plotly_chart(style_fig(fig, height=520), use_container_width=True)
            except Exception as e:
                st.warning(f"Treemap rendering failed ({type(e).__name__}). Showing ranked bar instead.")
                bar_rank(tree.rename(columns={g: g}), g, f"Top {g} by Volume")

# ---- Pareto ----
with tab_pareto:
    g = col_owner or col_sector
    if g:
        t = fdf.groupby(g).size().reset_index(name="Count").sort_values("Count", ascending=False)
        t["Cum%"] = t["Count"].cumsum() / max(t["Count"].sum(), 1) * 100
        top = t.head(top_n)
        fig = px.bar(top, x=g, y="Count", title=f"Pareto — Top {top_n} by Volume")
        fig = add_labels(fig)
        line = px.line(top, x=g, y="Cum%", markers=True)
        line.update_traces(yaxis="y2")
        fig.update_layout(yaxis2=dict(overlaying="y", side="right", range=[0, 100], title="Cum %"))
        for tr in line.data:
            fig.add_trace(tr)
        st.plotly_chart(style_fig(fig), use_container_width=True)
    else:
        st.info("No grouping column available.")

# ---- Pivot Builder ----
with tab_pivot:
    st.subheader("Pivot Builder")
    groups = [c for c in [col_owner, col_sector, col_status, col_priority] if c]
    if not groups:
        st.info("No grouping fields available.")
    else:
        row_col = st.selectbox("Rows", groups, index=0, key="pivot_row")
        metric = st.selectbox(
            "Metric",
            ["Task Count"] + ([f"Average {col_tat_total}"] if col_tat_total else []),
            index=0, key="pivot_metric"
        )
        if metric == "Task Count":
            pivot = fdf.groupby(row_col).size().to_frame("Count").sort_values("Count", ascending=False)
        else:
            pivot = fdf.groupby(row_col)[col_tat_total].mean().to_frame(f"Avg {col_tat_total}").sort_values(f"Avg {col_tat_total}", ascending=False)
        st.dataframe(pivot, use_container_width=True, height=420)

# ---- Data & Export ----
with tab_data:
    st.subheader("Filtered Data")
    safe = fdf.copy()
    for c in safe.columns:
        if safe[c].dtype == "object":
            safe[c] = safe[c].astype(str)
    st.dataframe(safe, use_container_width=True, height=420)

    def build_exports(data: pd.DataFrame):
        sheets = {"FilteredData": data}
        if col_sector:
            sheets["BySector_Count"] = data.groupby(col_sector).size().to_frame("Count")
            if col_tat_total and data[col_tat_total].notna().any():
                sheets["BySector_AvgTAT"] = data.groupby(col_sector)[col_tat_total].mean().to_frame("Avg TAT (days)")
        if col_owner:
            sheets["ByOwner_Count"] = data.groupby(col_owner).size().to_frame("Count")
            if col_tat_total and data[col_tat_total].notna().any():
                sheets["ByOwner_AvgTAT"] = data.groupby(col_owner)[col_tat_total].mean().to_frame("Avg TAT (days)")
        if col_status:
            sheets["ByStatus_Count"] = data.groupby(col_status).size().to_frame("Count")
        # SLA table (compute if available)
        if col_tat_total:
            tmp = data.copy()
            def sla_bucket(x):
                if pd.isna(x): return "Unknown"
                if x <= 5: return "≤5"
                if x <= 10: return "6–10"
                if x <= 20: return "11–20"
                return ">20"
            tmp["_SLA"] = tmp[col_tat_total].apply(sla_bucket)
            gcol = col_owner or col_sector
            if gcol:
                sheets["SLA_ByGroup"] = tmp.groupby([gcol,"_SLA"]).size().unstack("_SLA", fill_value=0).sort_index()
        # Hold breakdown averages
        if hold_stack_cols:
            gcol = col_owner or col_sector
            if gcol:
                nonempty = [c for c in hold_stack_cols if data[c].notna().any() and data[c].abs().sum() > 0]
                if nonempty:
                    sheets["Hold_Breakdown_Avg"] = data.groupby(gcol)[nonempty].mean(numeric_only=True)
        if "Total Hold TAT" in all_hold_cols:
            gcol = col_owner or col_sector
            if gcol:
                sheets["Total_Hold_TAT_Avg"] = data.groupby(gcol)["Total Hold TAT"].mean().to_frame("Total Hold TAT (Avg)")
        return sheets

    def to_excel_bytes(sheets: dict):
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
            for name, frame in sheets.items():
                if isinstance(frame, pd.DataFrame):
                    frame.to_excel(writer, sheet_name=str(name)[:31])
        out.seek(0)
        return out

    dl = to_excel_bytes(build_exports(fdf.copy()))
    st.download_button(
        "Download Excel (filtered + pivots/SLA/holds)",
        data=dl,
        file_name="tat_analysis_export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.markdown(
    '<div class="small-note">Tip: use the sidebar “Top N” and “Bar value” to change how rankings look. Horizontal bars appear automatically when labels are long.</div>',
    unsafe_allow_html=True,
)
