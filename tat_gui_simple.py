
import io
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
from datetime import timedelta

st.set_page_config(page_title="TAT Analysis – Simple Mode", layout="wide")
st.title("📊 TAT Analysis – Simple Mode")

# ----------------- Helpers -----------------
def coerce_dates(df: pd.DataFrame) -> pd.DataFrame:
    for col in df.columns:
        lc = str(col).lower()
        if "date" in lc or "hit" in lc or "target" in lc or "complete" in lc:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    return df

def default_last_n_days(series: pd.Series, n=90):
    s = pd.to_datetime(series, errors="coerce").dropna()
    if s.empty: return None
    max_d = s.max().date()
    min_d = s.min().date()
    if (max_d - min_d).days > n:
        return (max_d - timedelta(days=n), max_d)
    return (min_d, max_d)

def group_small(df, value_col, label_col, topn=10, other_label="Other"):
    if df.empty: return df
    df = df.sort_values(value_col, ascending=False)
    if len(df) <= topn: return df
    top = df.iloc[:topn].copy()
    other_sum = df.iloc[topn:][value_col].sum()
    other_row = {label_col: other_label, value_col: other_sum}
    return pd.concat([top, pd.DataFrame([other_row])], ignore_index=True)

def sla_bucket(x):
    if pd.isna(x): return "Unknown"
    if x <= 5: return "≤5"
    if x <= 10: return "6–10"
    if x <= 20: return "11–20"
    return ">20"
SLA_ORDER = ["≤5","6–10","11–20",">20","Unknown"]

# ----------------- Load -----------------
st.sidebar.header("1) Load")
uploaded = st.sidebar.file_uploader("Upload .xlsx", type=["xlsx"])

@st.cache_data(show_spinner=False)
def load_alltasks(file_or_path):
    df = pd.read_excel(file_or_path, sheet_name="AllTasks")
    df.columns = [str(c).strip() for c in df.columns]
    df = coerce_dates(df)
    return df

if uploaded:
    df = load_alltasks(uploaded)
else:
    try:
        df = load_alltasks("TAT Analysis by CompleteDate.xlsx")
        st.info("Using local 'TAT Analysis by CompleteDate.xlsx' → AllTasks (upload to override).")
    except Exception:
        st.warning("Upload your Excel to continue.")
        st.stop()

# ----------------- Mapping -----------------
st.sidebar.header("2) Map columns")
cols = df.columns.tolist()
def pick(label, *cands):
    opts = ["— none —"] + cols
    default = next((c for c in cands if c in cols), "— none —")
    val = st.sidebar.selectbox(label, opts, index=opts.index(default) if default in opts else 0)
    return None if val == "— none —" else val

sector   = pick("Sector", "Sector")
owner    = pick("Owner/Engineer", "Task Owner","Owner","Engineer","Assignee")
status   = pick("Status", "Redash Status","Task Source Status","Status")
priority = pick("Priority", "Priority")
program  = pick("Program (optional)", "Program","Programs")
team     = pick("Team (optional)", "Team","Teams")

hit_date      = pick("Hit Date","Hit Date")
target_date   = pick("Target Date","Target Date","Initial Target Date")
complete_date = pick("Complete Date","Task Complete Date","Validation Complete Date")

tat = pick("Total TAT (days)","Total Task TAT","NoSTR Total Task TAT","MinusHold Total Task TAT","NoHold NoSTR Total Adjusted TAT")
otd_flag = pick("On-Time flag (optional)","OTD PassFail","OTD_PassFail","OnTime")

# Hold columns
hold_cols = [c for c in cols if c.endswith("Hold TAT") or " Hold TAT" in c]

# Coerce types
if tat: df[tat] = pd.to_numeric(df[tat], errors="coerce")
for c in hold_cols: df[c] = pd.to_numeric(df[c], errors="coerce")

if otd_flag and otd_flag in df.columns:
    df["_OnTime"] = df[otd_flag].map({1:True,"1":True,"Y":True,"Yes":True,"TRUE":True,True:True,0:False,"0":False,"N":False,"No":False,"FALSE":False,False:False}).astype("boolean")
elif complete_date and target_date:
    df["_OnTime"] = (pd.to_datetime(df[complete_date]) <= pd.to_datetime(df[target_date])).astype("boolean")
else:
    df["_OnTime"] = pd.NA

# ----------------- Filters -----------------
st.sidebar.header("3) Filters")
def multi(col, label):
    if not col: return []
    vals = df[col].dropna().astype(str).unique()
    vals.sort()
    return st.sidebar.multiselect(label, vals, default=[])

sel_sector   = multi(sector, "Sector")
sel_owner    = multi(owner, "Owner/Engineer")
sel_status   = multi(status, "Status")
sel_priority = multi(priority, "Priority")
sel_program  = multi(program, "Program")
sel_team     = multi(team, "Team")

def daterange(col, label):
    if not col: return None
    s = pd.to_datetime(df[col], errors="coerce").dropna()
    if s.empty: return None
    return st.sidebar.date_input(label, value=default_last_n_days(s, n=90))

hit_rng  = daterange(hit_date, "Hit range")
tgt_rng  = daterange(target_date, "Target range")
cmp_rng  = daterange(complete_date, "Complete range")

f = df.copy()

def apply_in(c, sel):
    global f
    if c and sel: f = f[f[c].astype(str).isin(sel)]

for c, sel in [(sector,sel_sector),(owner,sel_owner),(status,sel_status),(priority,sel_priority),(program,sel_program),(team,sel_team)]:
    apply_in(c, sel)

def apply_range(c, dr):
    global f
    if c and dr and isinstance(dr, tuple) and len(dr)==2 and all(dr):
        start = pd.to_datetime(dr[0]); end = pd.to_datetime(dr[1]) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
        f = f[(pd.to_datetime(f[c], errors="coerce") >= start) & (pd.to_datetime(f[c], errors="coerce") <= end)]

apply_range(hit_date, hit_rng)
apply_range(target_date, tgt_rng)
apply_range(complete_date, cmp_rng)

st.caption(f"Showing **{len(f):,}** of **{len(df):,}** rows after filters.")

# ----------------- Simple Controls -----------------
st.sidebar.header("4) View")
view = st.sidebar.selectbox("Choose view", [
    "Overview KPIs",
    "Top Sectors (Count)",
    "Top Owners (Count)",
    "Avg TAT by Sector (days)",
    "Completed per Month",
    "Avg TAT per Month (days)",
    "SLA Mix by Sector",
    "Hold Breakdown by Sector (days)",
    "On-Time % by Sector",
    "Data Table"
])
topN = st.sidebar.slider("Top N", 5, 30, 10, 1)
rotate = st.sidebar.checkbox("Rotate labels 45°", value=True)
show_labels = st.sidebar.checkbox("Show data labels", value=False)
counts_vs_pct = st.sidebar.selectbox("Bar value", ["Counts","Percentages"], index=0)

def style_fig(fig, ytitle=None):
    fig.update_layout(template="plotly_white", height=460, margin=dict(l=20,r=20,t=60,b=60))
    if ytitle:
        fig.update_yaxes(title=ytitle, ticks="outside")
    if rotate:
        fig.update_xaxes(tickangle=45)
    if show_labels:
        fig.update_traces(texttemplate="%{y:.1f}" if ytitle and "%%" in ytitle else "%{y}", textposition="outside", cliponaxis=False)
    return fig

# ----------------- KPIs -----------------
def draw_kpis(frame):
    k1,k2,k3,k4 = st.columns(4)
    k1.metric("Total Tasks", f"{len(frame):,}")
    avg_t = frame[tat].mean() if tat and frame[tat].notna().any() else None
    k2.metric("Avg TAT", f"{avg_t:.1f} days" if avg_t is not None else "—")
    ontime = (frame["_OnTime"]==True).mean()*100 if frame["_OnTime"].notna().any() else None
    k3.metric("On-Time Rate", f"{ontime:.1f}%" if ontime is not None else "—")
    hpct = (frame[priority].astype(str).str.contains("Urgent|High", case=False, na=False)).mean()*100 if priority else None
    k4.metric("High Priority %", f"{hpct:.1f}%" if hpct is not None else "—")

# ----------------- Views -----------------
if view == "Overview KPIs":
    draw_kpis(f)

elif view == "Top Sectors (Count)":
    if sector:
        data = f.groupby(sector).size().reset_index(name="Count").sort_values("Count", ascending=False)
        data = group_small(data, "Count", sector, topn=topN)
        fig = px.bar(data, x=sector, y="Count", title="Top Sectors by Count", text="Count")
        st.plotly_chart(style_fig(fig, "Count"), use_container_width=True)
    else:
        st.info("Please map the Sector column.")

elif view == "Top Owners (Count)":
    if owner:
        data = f.groupby(owner).size().reset_index(name="Count").sort_values("Count", ascending=False)
        data = group_small(data, "Count", owner, topn=topN)
        fig = px.bar(data, x=owner, y="Count", title="Top Owners by Count", text="Count")
        st.plotly_chart(style_fig(fig, "Count"), use_container_width=True)
    else:
        st.info("Please map the Owner/Engineer column.")

elif view == "Avg TAT by Sector (days)":
    if sector and tat and f[tat].notna().any():
        data = f.groupby(sector, as_index=False)[tat].mean().sort_values(tat, ascending=False)
        data = group_small(data, tat, sector, topn=topN)
        fig = px.bar(data, x=sector, y=tat, title="Average TAT by Sector (days)", text=tat)
        st.plotly_chart(style_fig(fig, "Days"), use_container_width=True)
    else:
        st.info("Need Sector and Total TAT mapped.")

elif view == "Completed per Month":
    if complete_date and f[complete_date].notna().any():
        t = f.dropna(subset=[complete_date]).copy()
        t["Month"] = pd.to_datetime(t[complete_date]).dt.to_period("M").dt.to_timestamp()
        data = t.groupby("Month").size().reset_index(name="Completed Tasks")
        fig = px.bar(data, x="Month", y="Completed Tasks", title="Completed Tasks per Month", text="Completed Tasks")
        st.plotly_chart(style_fig(fig, "Count"), use_container_width=True)
    else:
        st.info("Please map the Complete Date column.")

elif view == "Avg TAT per Month (days)":
    if complete_date and tat and f[complete_date].notna().any():
        t = f.dropna(subset=[complete_date, tat]).copy()
        t["Month"] = pd.to_datetime(t[complete_date]).dt.to_period("M").dt.to_timestamp()
        data = t.groupby("Month", as_index=False)[tat].mean()
        fig = px.line(data, x="Month", y=tat, markers=True, title="Average TAT per Month (days)")
        st.plotly_chart(style_fig(fig, "Days"), use_container_width=True)
    else:
        st.info("Need Complete Date and Total TAT mapped.")

elif view == "SLA Mix by Sector":
    if sector and tat and f[tat].notna().any():
        tmp = f.copy()
        tmp["_SLA"] = pd.Categorical(tmp[tat].apply(sla_bucket), categories=SLA_ORDER, ordered=True)
        sla = tmp.groupby([sector, "_SLA"]).size().reset_index(name="Count")
        if counts_vs_pct == "Percentages":
            sla["Pct"] = sla.groupby(sector)["Count"].transform(lambda x: x/x.sum()*100)
            fig = px.bar(sla, x=sector, y="Pct", color="_SLA", barmode="stack", title="SLA Mix by Sector (%)", category_orders={"_SLA":SLA_ORDER})
            st.plotly_chart(style_fig(fig, "Percent"), use_container_width=True)
        else:
            fig = px.bar(sla, x=sector, y="Count", color="_SLA", barmode="stack", title="SLA Count by Sector", category_orders={"_SLA":SLA_ORDER})
            st.plotly_chart(style_fig(fig, "Count"), use_container_width=True)
    else:
        st.info("Need Sector and Total TAT mapped.")

elif view == "Hold Breakdown by Sector (days)":
    if sector and hold_cols:
        agg = f.groupby(sector)[hold_cols].mean(numeric_only=True).reset_index()
        melted = agg.melt(id_vars=[sector], value_vars=hold_cols, var_name="Hold Type", value_name="Avg Days")
        fig = px.bar(melted, x=sector, y="Avg Days", color="Hold Type", barmode="stack", title="Average Hold Time by Sector (days)")
        st.plotly_chart(style_fig(fig, "Days"), use_container_width=True)
    else:
        st.info("Need Sector mapped and Hold columns present.")

elif view == "On-Time % by Sector":
    if sector and f["_OnTime"].notna().any():
        d = f.dropna(subset=["_OnTime"]).groupby(sector)["_OnTime"].mean().reset_index(name="On-Time %")
        d["On-Time %"] = d["On-Time %"]*100
        d = d.sort_values("On-Time %", ascending=False)
        d = group_small(d, "On-Time %", sector, topn=topN)
        fig = px.bar(d, x=sector, y="On-Time %", title="On-Time % by Sector", text="On-Time %")
        st.plotly_chart(style_fig(fig, "Percent"), use_container_width=True)
    else:
        st.info("Need Sector mapped and On-Time data available.")

elif view == "Data Table":
    st.subheader("Filtered Data")
    search = st.text_input("Search (contains, case-insensitive)")
    table = f.copy()
    if search:
        mask = pd.Series(False, index=table.index)
        for c in table.columns:
            mask |= table[c].astype(str).str.contains(search, case=False, na=False)
        table = table[mask]
    st.dataframe(table, use_container_width=True, height=480)

# Footer tip
st.caption("Tip: Map columns in the sidebar. Use Top N and the 90-day default date to keep charts readable.")
