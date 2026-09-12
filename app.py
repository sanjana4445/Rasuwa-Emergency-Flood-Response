from pathlib import Path
from datetime import date

import pandas as pd
import plotly.express as px
import streamlit as st


st.set_page_config(
    page_title="UNICEF Flood Response Dashboard",
    page_icon="💧",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    :root { --ink:#1f2937; --muted:#64748b; --unicef-blue:#00aeef; --unicef-deep:#0072bc; --teal:#0072bc; --green:#20965a; --line:#cbd5e1; --paper:#fff; --canvas:#f2f8fc; }
    .stApp { background:var(--canvas); color:var(--ink); }
    [data-testid="stHeader"] { background:transparent; }
    .block-container { max-width:1400px; padding:22px 32px 56px; }
    .hero { background:linear-gradient(115deg, var(--unicef-deep), var(--unicef-blue)); border-radius:18px; color:white; padding:25px 28px; margin-bottom:22px; }
    .hero h1 { margin:0; color:white; font-size:29px; }
    .hero p { margin:7px 0 0; color:#d8e8e9; font-size:15px; }
    .metric-card { background:var(--paper); border:1px solid var(--line); border-radius:12px; padding:17px; min-height:104px; }
    .metric-title { color:#64748b; font-size:12px; font-weight:700; text-transform:uppercase; letter-spacing:.4px; }
    .metric-value { color:#0f172a; font-size:27px; font-weight:750; margin-top:6px; }
    .metric-sub { color:#1594a2; font-size:12px; font-weight:600; margin-top:3px; }
    .section { background:#f8fafc; border:1px solid #e2e8f0; border-radius:10px; padding:13px 18px 4px; margin:20px 0 10px; }
    .section h4 { color:#0f172a; font-size:16px; margin:0 0 2px; }
    .section p { color:#64748b; font-size:13px; margin:0 0 9px; }
    [data-testid="stSidebar"] { background:#eef7fc; border-right:1px solid #c7e7f5; }
    [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 { color:var(--unicef-deep); }
    [data-testid="stPills"] button[aria-pressed="true"] { background:var(--green) !important; border-color:var(--green) !important; color:white !important; }
    [data-testid="stPills"] button[aria-pressed="false"] { background:white !important; border-color:var(--line) !important; color:var(--ink) !important; }
    [data-testid="stPills"] button:hover { border-color:var(--unicef-blue) !important; }
    .filter-help { color:var(--muted); font-size:12px; margin:0 0 8px; }
    </style>
    """,
    unsafe_allow_html=True,
)

PARTNER_FILES = {
    "CDC": "partner_files/CDC_Monitoring Matrix.xlsx",
    "Chaya": "partner_files/Chaya_Monitoring Matrix.xlsx",
    "COSOC": "partner_files/COSOC_Monitoring Matrix.xlsx",
    "SHANTI": "partner_files/SHANTI_Monitoring Matrix.xlsx",
}

PROJECT_START = date(2026, 9, 15)
PROJECT_END = date(2026, 12, 31)
TODAY = date.today()
PROJECT_DAYS = (PROJECT_END - PROJECT_START).days + 1
if TODAY < PROJECT_START:
    PROJECT_STATUS = f"Starts in {(PROJECT_START - TODAY).days} days"
    DAYS_REMAINING = PROJECT_DAYS
elif TODAY <= PROJECT_END:
    PROJECT_STATUS = "Project active"
    DAYS_REMAINING = (PROJECT_END - TODAY).days + 1
else:
    PROJECT_STATUS = "Project ended"
    DAYS_REMAINING = 0


def numeric_series(series):
    values = series.astype("string").str.replace(",", "", regex=False)
    extracted = values.str.extract(r"(-?\d+(?:\.\d+)?)", expand=False)
    return pd.to_numeric(extracted, errors="coerce").fillna(0)


def find_header_row(raw):
    for row_number, row in raw.iterrows():
        if row.astype("string").str.strip().str.upper().eq("SN").any():
            return row_number
    return None


def read_partner_sheet(file_path, partner, sheet_name):
    raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    header_row = find_header_row(raw)
    if header_row is None:
        return pd.DataFrame()

    header_values = raw.iloc[header_row].astype("string").str.strip().str.upper()
    sn_column = header_values.eq("SN").idxmax()
    data = raw.iloc[header_row + 1:].copy()
    is_daily = "daily" in sheet_name.lower()
    relative_columns = list(range(8)) + (list(range(8, 12)) if is_daily else [9, 10, 11, 12])
    source_columns = [sn_column + offset for offset in relative_columns]
    if max(source_columns) >= data.shape[1]:
        return pd.DataFrame()

    data = data.iloc[:, source_columns].copy()
    data.columns = [
        "SN", "Agency", "Province", "District", "Municipality", "Ward",
        "Result Statement", "Indicator", "Unit", "Target", "Progress", "Activities",
    ]
    data = data[data["Indicator"].notna()].copy()
    data = data[
        ~data["Indicator"].astype("string").str.strip().str.lower().eq("performance indicator/s")
    ]
    data["Result Statement"] = data["Result Statement"].ffill()
    data["Result Area"] = data["Result Statement"].astype("string").str.split("\n").str[0]
    data["Target"] = numeric_series(data["Target"])
    data["Progress"] = numeric_series(data["Progress"])
    data["Partner"] = partner
    data["Frequency"] = "Daily" if is_daily else "Weekly"
    data["Source Sheet"] = sheet_name
    data["Row ID"] = [f"{partner}|{sheet_name}|{index}" for index in data.index]
    return data.reset_index(drop=True)


@st.cache_data
def load_all_partners():
    frames = []
    errors = []
    for partner, relative_path in PARTNER_FILES.items():
        file_path = Path(relative_path)
        if not file_path.exists():
            errors.append(f"{partner}: file not found ({relative_path})")
            continue
        try:
            workbook = pd.ExcelFile(file_path)
            for sheet_name in workbook.sheet_names:
                sheet_data = read_partner_sheet(file_path, partner, sheet_name)
                if not sheet_data.empty:
                    frames.append(sheet_data)
        except Exception as error:
            errors.append(f"{partner}: {error}")

    if not frames:
        return pd.DataFrame(), errors
    return pd.concat(frames, ignore_index=True), errors


def chart_theme(figure):
    figure.update_layout(
        template="plotly_white",
        paper_bgcolor="#ffffff",
        plot_bgcolor="#ffffff",
        font=dict(color="#243331"),
        margin=dict(l=10, r=20, t=45, b=20),
    )
    return figure


df_master, load_errors = load_all_partners()

st.markdown(
    """
    <div class="hero">
        <h1>💧 UNICEF Flood Response Dashboard</h1>
        <p>Multi-partner WASH monitoring across target districts, municipalities, and reporting periods.</p>
        <p>📅 Project period: 15 September 2026 - 31 December 2026</p>
    </div>
    """,
    unsafe_allow_html=True,
)

if load_errors:
    with st.expander("Data loading warnings"):
        for error in load_errors:
            st.warning(error)

if df_master.empty:
    st.error("No valid monitoring rows were loaded. Check the Excel files and refresh the page.")
    st.stop()

with st.sidebar:
    st.header("🔎 Filters")
    st.metric("⏳ Days remaining", f"{DAYS_REMAINING:,}", PROJECT_STATUS)
    st.caption(f"Project dates: {PROJECT_START.strftime('%d %b %Y')} - {PROJECT_END.strftime('%d %b %Y')}")
    if st.button("Reload Excel files", width="stretch"):
        st.cache_data.clear()
        st.rerun()
    partner_values = sorted(df_master["Partner"].dropna().astype(str).unique())
    selected_partner = st.selectbox("🤝 Partner", ["All partners"] + partner_values)
    selected_frequency = st.pills(
        "🗓️ Reporting period",
        ["Daily", "Weekly"],
        default=["Daily", "Weekly"],
        selection_mode="multi",
    )

    district_values = sorted(value for value in df_master["District"].dropna().astype(str).unique() if value.strip())
    selected_district = st.selectbox("📍 District", ["All districts"] + district_values)

    partner_df = df_master if selected_partner == "All partners" else df_master[df_master["Partner"] == selected_partner]
    district_df = partner_df if selected_district == "All districts" else partner_df[partner_df["District"].astype(str) == selected_district]
    palika_values = sorted(value for value in district_df["Municipality"].dropna().astype(str).unique() if value.strip() and value.lower() != "nan")
    selected_palika = st.selectbox("🏘️ Palika / Municipality", ["All Palikas"] + palika_values)

    output_values = sorted(df_master["Result Area"].dropna().astype(str).unique())
    selected_outputs = st.multiselect("🎯 Output areas", output_values, default=output_values)

filtered_df = df_master[
    df_master["Frequency"].isin(selected_frequency)
    & df_master["Result Area"].isin(selected_outputs)
].copy()
if selected_partner != "All partners":
    filtered_df = filtered_df[filtered_df["Partner"] == selected_partner]
if selected_district != "All districts":
    filtered_df = filtered_df[filtered_df["District"].astype(str) == selected_district]
if selected_palika != "All Palikas":
    filtered_df = filtered_df[filtered_df["Municipality"].astype(str) == selected_palika]

if filtered_df.empty:
    st.info("No rows match the selected filters.")
    st.stop()

filtered_df["Completion %"] = (
    filtered_df["Progress"] / filtered_df["Target"].replace(0, 1) * 100
).clip(0, 100).round(1)
total_target = filtered_df["Target"].sum()
total_progress = filtered_df["Progress"].sum()
completion = total_progress / total_target * 100 if total_target else 0
remaining = max(total_target - total_progress, 0)

st.markdown('<div class="section"><h4>Response at a glance</h4><p>All values reflect the selected partners, reporting periods, districts, and outputs.</p></div>', unsafe_allow_html=True)
m1, m2, m3, m4 = st.columns(4)
metrics = [
    ("Total target", f"{total_target:,.0f}", f"{len(filtered_df):,} indicators"),
    ("Cumulative progress", f"{total_progress:,.0f}", "From partner matrices"),
    ("Completion rate", f"{completion:.1f}%", "Weighted by target"),
    ("Remaining gap", f"{remaining:,.0f}", "Target minus progress"),
]
for column, (title, value, subtitle) in zip((m1, m2, m3, m4), metrics):
    with column:
        st.markdown(f'<div class="metric-card"><div class="metric-title">{title}</div><div class="metric-value">{value}</div><div class="metric-sub">{subtitle}</div></div>', unsafe_allow_html=True)

tab_overview, tab_partner, tab_period, tab_trends, tab_monitoring, tab_data = st.tabs(
    ["Overview", "Partners", "Daily / Weekly", "Partner trends", "Monitoring", "Data & export"]
)

with tab_overview:
    st.markdown('<div class="section"><h4>Target versus progress by output</h4><p>Use this view to see which response areas have the largest gaps.</p></div>', unsafe_allow_html=True)
    output_summary = filtered_df.groupby("Result Area", as_index=False)[["Target", "Progress"]].sum()
    chart_data = output_summary.melt("Result Area", var_name="Measure", value_name="Value")
    chart_data["Measure"] = chart_data["Measure"].replace({"Progress": "Achieved"})
    figure = px.bar(
        chart_data, x="Value", y="Result Area", color="Measure", orientation="h", barmode="group",
        color_discrete_map={"Target": "#e4a11b", "Achieved": "#0072bc"}, text="Value",
    )
    figure.update_traces(texttemplate="%{text:,.0f}", textposition="outside", cliponaxis=False)
    st.plotly_chart(chart_theme(figure), width="stretch")

with tab_partner:
    partner_summary = filtered_df.groupby("Partner", as_index=False)[["Target", "Progress"]].sum()
    partner_summary["Completion %"] = (partner_summary["Progress"] / partner_summary["Target"].replace(0, 1) * 100).clip(0, 100).round(1)
    left, right = st.columns([6, 4])
    with left:
        partner_chart = px.bar(partner_summary, x="Partner", y=["Target", "Progress"], barmode="group", color_discrete_map={"Target": "#e4a11b", "Progress": "#0072bc"})
        st.plotly_chart(chart_theme(partner_chart), width="stretch")
    with right:
        st.dataframe(partner_summary, width="stretch", hide_index=True, column_config={"Completion %": st.column_config.NumberColumn(format="%.1f%%")})

with tab_period:
    period_summary = filtered_df.groupby("Frequency", as_index=False)[["Target", "Progress"]].sum()
    period_summary["Completion %"] = (period_summary["Progress"] / period_summary["Target"].replace(0, 1) * 100).clip(0, 100).round(1)
    period_chart = px.bar(period_summary, x="Frequency", y=["Target", "Progress"], barmode="group", color_discrete_map={"Target": "#e4a11b", "Progress": "#0072bc"})
    st.plotly_chart(chart_theme(period_chart), width="stretch")
    st.dataframe(filtered_df[["Partner", "Frequency", "District", "Municipality", "Indicator", "Target", "Progress", "Completion %", "Activities"]], width="stretch", hide_index=True)

with tab_trends:
    st.subheader("Partner trends and target gaps")
    st.caption("This compares Daily and Weekly reporting periods. Add a date column to the Excel sheets when a chronological time-lapse is required.")
    partner_period = filtered_df.groupby(["Partner", "Frequency"], as_index=False)[["Target", "Progress"]].sum()
    partner_period["Completion %"] = (
        partner_period["Progress"] / partner_period["Target"].replace(0, 1) * 100
    ).clip(0, 100).round(1)
    partner_period["Remaining gap"] = (partner_period["Target"] - partner_period["Progress"]).clip(lower=0)
    partner_period["Frequency"] = pd.Categorical(partner_period["Frequency"], ["Daily", "Weekly"], ordered=True)

    trend_left, trend_right = st.columns([6, 4])
    with trend_left:
        completion_trend = px.line(
            partner_period.sort_values("Frequency"),
            x="Frequency",
            y="Completion %",
            color="Partner",
            markers=True,
            range_y=[0, 100],
            title="Partner completion by reporting period",
            color_discrete_sequence=["#0072bc", "#00aeef", "#20965a", "#e4a11b"],
        )
        completion_trend.update_layout(yaxis_title="Completion (%)", xaxis_title="Reporting period")
        st.plotly_chart(chart_theme(completion_trend), width="stretch")
    with trend_right:
        gap_trend = px.bar(
            partner_period.sort_values("Remaining gap", ascending=True),
            x="Remaining gap",
            y="Partner",
            color="Frequency",
            orientation="h",
            barmode="group",
            text="Remaining gap",
            color_discrete_map={"Daily": "#e4a11b", "Weekly": "#0072bc"},
            title="Partner gap against target",
        )
        gap_trend.update_traces(texttemplate="%{text:,.0f}", textposition="outside", cliponaxis=False)
        st.plotly_chart(chart_theme(gap_trend), width="stretch")
    st.dataframe(
        partner_period,
        width="stretch",
        hide_index=True,
        column_config={
            "Target": st.column_config.NumberColumn(format="%,.0f"),
            "Progress": st.column_config.NumberColumn(format="%,.0f"),
            "Completion %": st.column_config.NumberColumn(format="%.1f%%"),
            "Remaining gap": st.column_config.NumberColumn(format="%,.0f"),
        },
    )

with tab_monitoring:
    st.subheader("Progress monitoring and follow-up risk")
    st.caption("Use the gap and risk views to prioritize support. Rows without a target are shown separately and are not treated as high risk.")

    monitoring = filtered_df.groupby(["Partner", "Frequency", "Result Area"], as_index=False)[["Target", "Progress"]].sum()
    monitoring["Remaining gap"] = (monitoring["Target"] - monitoring["Progress"]).clip(lower=0)
    monitoring["Completion %"] = (
        monitoring["Progress"] / monitoring["Target"].replace(0, 1) * 100
    ).clip(0, 100).round(1)
    monitoring["Priority"] = monitoring.apply(
        lambda row: "No target" if row["Target"] <= 0 else (
            "High" if row["Completion %"] < 25 else (
                "Watch" if row["Completion %"] < 75 else "On track"
            )
        ),
        axis=1,
    )
    risk_rows = monitoring[monitoring["Priority"] != "No target"].copy()
    high_risk_count = int((risk_rows["Priority"] == "High").sum())
    watch_count = int((risk_rows["Priority"] == "Watch").sum())
    no_target_count = int((monitoring["Priority"] == "No target").sum())

    if total_progress == 0:
        st.warning("No progress has been recorded in the selected data yet. Targets are loaded, but progress values are still zero or blank.")

    risk_metric_1, risk_metric_2, risk_metric_3, risk_metric_4 = st.columns(4)
    with risk_metric_1:
        st.metric("Indicators reviewed", f"{len(monitoring):,}")
    with risk_metric_2:
        st.metric("High risk", f"{high_risk_count:,}", help="Targeted indicators below 25% completion.")
    with risk_metric_3:
        st.metric("Watch", f"{watch_count:,}", help="Targeted indicators between 25% and 75% completion.")
    with risk_metric_4:
        st.metric("Without target", f"{no_target_count:,}", help="Rows needing a target before risk can be assessed.")

    if risk_rows.empty:
        st.info("Add target values to the Excel sheets to activate risk and gap analysis.")
    else:
        risk_rows["Area"] = risk_rows["Partner"] + " · " + risk_rows["Result Area"] + " · " + risk_rows["Frequency"]
        gap_left, gap_right = st.columns([6, 4])
        with gap_left:
            gap_chart_data = risk_rows.sort_values("Remaining gap", ascending=True).tail(12)
            gap_chart = px.bar(
                gap_chart_data,
                x="Remaining gap",
                y="Area",
                color="Priority",
                orientation="h",
                text="Remaining gap",
                color_discrete_map={"High": "#dc2626", "Watch": "#e4a11b", "On track": "#20965a"},
                title="Largest remaining gaps",
            )
            gap_chart.update_traces(texttemplate="%{text:,.0f}", textposition="outside", cliponaxis=False)
            st.plotly_chart(chart_theme(gap_chart), width="stretch")
        with gap_right:
            risk_chart = px.scatter(
                risk_rows,
                x="Completion %",
                y="Remaining gap",
                size="Target",
                color="Priority",
                hover_name="Area",
                hover_data={"Target": ":,.0f", "Progress": ":,.0f", "Completion %": ":.1f"},
                range_x=[0, 100],
                color_discrete_map={"High": "#dc2626", "Watch": "#e4a11b", "On track": "#20965a"},
                title="Completion versus remaining gap",
            )
            risk_chart.update_layout(xaxis_title="Completion (%)", yaxis_title="Remaining gap")
            st.plotly_chart(chart_theme(risk_chart), width="stretch")

    display_columns = ["Partner", "Frequency", "Result Area", "Target", "Progress", "Remaining gap", "Completion %", "Priority"]
    st.markdown("#### Follow-up priority list")
    st.dataframe(
        monitoring.sort_values(["Priority", "Remaining gap"], ascending=[True, False])[display_columns],
        width="stretch",
        hide_index=True,
        column_config={
            "Target": st.column_config.NumberColumn(format="%,.0f"),
            "Progress": st.column_config.NumberColumn(format="%,.0f"),
            "Remaining gap": st.column_config.NumberColumn(format="%,.0f"),
            "Completion %": st.column_config.NumberColumn(format="%.1f%%"),
        },
    )

with tab_data:
    st.subheader("Data review and export")
    st.caption("Update the four Excel files for permanent changes, then click Reload Excel files in the sidebar. This table is a filtered review of the current source data.")
    export_columns = ["Row ID", "Partner", "Frequency", "District", "Municipality", "Indicator", "Target", "Progress", "Activities"]
    st.dataframe(filtered_df[export_columns], width="stretch", hide_index=True)
    st.download_button("Download filtered CSV", filtered_df[export_columns].to_csv(index=False).encode("utf-8"), "flood_response_filtered.csv", "text/csv")
