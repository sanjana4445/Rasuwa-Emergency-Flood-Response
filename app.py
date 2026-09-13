from pathlib import Path
from datetime import date, datetime

import pandas as pd
import plotly.express as px
import streamlit as st


st.set_page_config(
    page_title="Rasuwa Emergency Flood Response",
    page_icon="💧",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    :root { --ink:#1f2937; --muted:#64748b; --rasuwa-blue:#00aeef; --rasuwa-deep:#0072bc; --teal:#0072bc; --green:#20965a; --line:#cbd5e1; --paper:#fff; --canvas:#f2f8fc; }
    .stApp { background:var(--canvas); color:var(--ink); }
    [data-testid="stHeader"] { background:transparent; }
    .block-container { max-width:1400px; padding:16px 32px 44px; }
    .hero { background:linear-gradient(115deg, var(--rasuwa-deep), var(--rasuwa-blue)); border-radius:16px; color:white; padding:18px 24px; margin-bottom:14px; }
    .hero h1 { margin:0; color:white; font-size:26px; }
    .hero p { margin:4px 0 0; color:#d8e8e9; font-size:14px; }
    .metric-card { background:var(--paper); border:1px solid var(--line); border-radius:10px; padding:12px 15px; min-height:88px; }
    .metric-title { color:#64748b; font-size:12px; font-weight:700; text-transform:uppercase; letter-spacing:.4px; }
    .metric-value { color:#0f172a; font-size:25px; font-weight:750; margin-top:4px; }
    .metric-sub { color:#1594a2; font-size:12px; font-weight:600; margin-top:3px; }
    .output-breakdown { background:var(--paper); border:1px solid var(--line); border-radius:10px; padding:11px 13px; margin-bottom:8px; }
    .output-breakdown-title { color:var(--rasuwa-deep); font-size:13px; font-weight:750; margin-bottom:7px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
    .output-breakdown-label { color:#64748b; font-size:11px; font-weight:700; text-transform:uppercase; letter-spacing:.35px; }
    .output-breakdown-value { color:#0f172a; font-size:19px; font-weight:750; margin-top:2px; }
    .output-breakdown-divider { border-left:1px solid var(--line); }
    .progress-focus { background:#ffffff; border:1px solid #b9ddec; border-left:6px solid var(--rasuwa-blue); border-radius:14px; padding:18px 22px; margin:18px 0; }
    .progress-focus-label { color:#64748b; font-size:12px; font-weight:700; letter-spacing:.5px; text-transform:uppercase; }
    .progress-focus-value { color:var(--rasuwa-deep); font-size:42px; font-weight:800; line-height:1.1; margin-top:3px; }
    .progress-track { background:#e2e8f0; border-radius:999px; height:10px; margin-top:12px; overflow:hidden; }
    .progress-fill { background:linear-gradient(90deg, var(--rasuwa-blue), var(--green)); border-radius:999px; height:100%; }
    .section { background:#f8fafc; border:1px solid #e2e8f0; border-radius:10px; padding:10px 16px 2px; margin:12px 0 7px; }
    .section h4 { color:#0f172a; font-size:16px; margin:0 0 2px; }
    .section p { color:#64748b; font-size:13px; margin:0 0 9px; }
    [data-testid="stSidebar"] { background:#eef7fc; border-right:1px solid #c7e7f5; }
    [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 { color:var(--rasuwa-deep); }
    [data-testid="stPills"] button[aria-pressed="true"] { background:var(--green) !important; border-color:var(--green) !important; color:white !important; }
    [data-testid="stPills"] button[aria-pressed="false"] { background:white !important; border-color:var(--line) !important; color:var(--ink) !important; }
    [data-testid="stPills"] button:hover { border-color:var(--rasuwa-blue) !important; }
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
STAFF_FILE = Path("partner_files/Staff_Roster.xlsx")
STAFF_SAVE_FILE = Path("partner_files/Staff_Roster.csv")
DATA_ENTRY_SAVE_FILE = Path("partner_files/Monitoring_Data_Entry.csv")
STAFF_COLUMNS = [
    "Name", "Position", "Duty Station", "Partner", "District", "Palika",
    "Phone", "Email", "Status",
]

PROJECT_START = date(2026, 9, 15)
PROJECT_END = date(2026, 12, 31)
TODAY = date.today()
PROJECT_DAYS = (PROJECT_END - PROJECT_START).days + 1
if TODAY < PROJECT_START:
    PROJECT_STATUS = f"Starts in {(PROJECT_START - TODAY).days} days"
    DAYS_REMAINING = PROJECT_DAYS
    TIME_LAPSE_PERCENTAGE = 0
elif TODAY <= PROJECT_END:
    PROJECT_STATUS = "Project active"
    DAYS_REMAINING = (PROJECT_END - TODAY).days + 1
    TIME_LAPSE_PERCENTAGE = ((TODAY - PROJECT_START).days / PROJECT_DAYS * 100)
else:
    PROJECT_STATUS = "Project ended"
    DAYS_REMAINING = 0
    TIME_LAPSE_PERCENTAGE = 100


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


def latest_source_update():
    source_paths = [Path(path) for path in PARTNER_FILES.values() if Path(path).exists()]
    if STAFF_FILE.exists():
        source_paths.append(STAFF_FILE)
    if not source_paths:
        return "No source files found"
    latest_time = max(path.stat().st_mtime for path in source_paths)
    return datetime.fromtimestamp(latest_time).strftime("%d %b %Y, %H:%M")


def read_staff_roster(uploaded_file=None):
    try:
        if uploaded_file is not None:
            if uploaded_file.name.lower().endswith(".csv"):
                roster = pd.read_csv(uploaded_file)
            else:
                roster = pd.read_excel(uploaded_file)
        elif STAFF_FILE.exists():
            roster = pd.read_excel(STAFF_FILE)
        elif STAFF_SAVE_FILE.exists():
            roster = pd.read_csv(STAFF_SAVE_FILE)
        else:
            return pd.DataFrame(columns=STAFF_COLUMNS), None
    except Exception as error:
        return pd.DataFrame(columns=STAFF_COLUMNS), str(error)

    normalized = {str(column).strip().lower(): column for column in roster.columns}
    renamed = {}
    for column in STAFF_COLUMNS:
        source_column = normalized.get(column.lower())
        if source_column is not None:
            renamed[source_column] = column
    roster = roster.rename(columns=renamed)
    for column in STAFF_COLUMNS:
        if column not in roster.columns:
            roster[column] = ""
    return roster[STAFF_COLUMNS].fillna(""), None


df_master, load_errors = load_all_partners()
staff_master, staff_load_error = read_staff_roster()
last_updated = latest_source_update()

st.markdown(
    """
    <div class="hero">
        <h1>💧 Rasuwa Emergency Flood Response</h1>
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

    monitoring_districts = df_master["District"].dropna().astype(str).tolist()
    staff_districts = staff_master["District"].dropna().astype(str).tolist()
    district_values = sorted({value.strip() for value in monitoring_districts + staff_districts if value.strip() and value.lower() != "nan"})
    selected_district = st.selectbox("📍 District", ["All districts"] + district_values)

    partner_df = df_master if selected_partner == "All partners" else df_master[df_master["Partner"] == selected_partner]
    district_df = partner_df if selected_district == "All districts" else partner_df[partner_df["District"].astype(str) == selected_district]
    staff_scope = staff_master
    if selected_partner != "All partners":
        staff_scope = staff_scope[staff_scope["Partner"].astype(str) == selected_partner]
    if selected_district != "All districts":
        staff_scope = staff_scope[staff_scope["District"].astype(str) == selected_district]
    monitoring_palikas = district_df["Municipality"].dropna().astype(str).tolist()
    staff_palikas = staff_scope["Palika"].dropna().astype(str).tolist()
    palika_values = sorted({value.strip() for value in monitoring_palikas + staff_palikas if value.strip() and value.lower() != "nan"})
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

m1 = st.columns(1)
with m1[0]:
    st.markdown(f'<div class="metric-card"><div class="metric-title">Last updated</div><div class="metric-value">{last_updated}</div><div class="metric-sub">Partner source files</div></div>', unsafe_allow_html=True)

output_totals = filtered_df.groupby("Result Area", as_index=False)[["Target", "Progress"]].sum()
columns_per_row = min(4, len(output_totals))
for output_start in range(0, len(output_totals), columns_per_row):
    output_columns = st.columns(columns_per_row)
    for output_column, output_row in zip(output_columns, output_totals.iloc[output_start:output_start + columns_per_row].itertuples(index=False)):
        with output_column:
            st.markdown(
                f'''<div class="output-breakdown"><div class="output-breakdown-title">{output_row[0]}</div><div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;"><div><div class="output-breakdown-label">Total target</div><div class="output-breakdown-value">{output_row[1]:,.0f}</div></div><div class="output-breakdown-divider" style="padding-left:16px;"><div class="output-breakdown-label">Met target</div><div class="output-breakdown-value">{output_row[2]:,.0f}</div></div></div></div>''',
                unsafe_allow_html=True,
            )

tab_overview, tab_partner, tab_period, tab_trends, tab_monitoring, tab_staff, tab_data = st.tabs(
    ["Overview", "Partners", "Daily / Weekly", "Partner trends", "Monitoring", "Staff roster", "Data & export"]
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
    partner_summary = filtered_df.groupby(["Partner", "Result Area"], as_index=False)[["Target", "Progress"]].sum()
    partner_summary["Completion %"] = (partner_summary["Progress"] / partner_summary["Target"].replace(0, 1) * 100).clip(0, 100).round(1)
    left, right = st.columns([6, 4])
    with left:
        partner_chart = px.bar(
            partner_summary,
            x="Partner",
            y=["Target", "Progress"],
            color="Result Area",
            barmode="group",
            color_discrete_sequence=["#0072bc", "#00aeef", "#20965a", "#e4a11b", "#64748b", "#0f766e"],
        )
        st.plotly_chart(chart_theme(partner_chart), width="stretch")
    with right:
        st.dataframe(partner_summary, width="stretch", hide_index=True, column_config={"Completion %": st.column_config.NumberColumn(format="%.1f%%")})

with tab_period:
    period_summary = filtered_df.groupby(["Frequency", "Result Area"], as_index=False)[["Target", "Progress"]].sum()
    period_summary["Completion %"] = (period_summary["Progress"] / period_summary["Target"].replace(0, 1) * 100).clip(0, 100).round(1)
    period_chart = px.bar(
        period_summary,
        x="Frequency",
        y=["Target", "Progress"],
        color="Result Area",
        barmode="group",
        color_discrete_sequence=["#0072bc", "#00aeef", "#20965a", "#e4a11b", "#64748b", "#0f766e"],
    )
    st.plotly_chart(chart_theme(period_chart), width="stretch")
    st.dataframe(filtered_df[["Partner", "Frequency", "District", "Municipality", "Result Area", "Indicator", "Target", "Progress", "Completion %", "Activities"]], width="stretch", hide_index=True)

with tab_trends:
    st.subheader("Partner trends and target gaps")
    st.caption("This compares Daily and Weekly reporting periods at output level so each target is tracked separately by result area.")
    partner_period = filtered_df.groupby(["Partner", "Frequency", "Result Area"], as_index=False)[["Target", "Progress"]].sum()
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
            color="Result Area",
            line_group="Partner",
            markers=True,
            range_y=[0, 100],
            title="Output completion by reporting period",
            color_discrete_sequence=["#0072bc", "#00aeef", "#20965a", "#e4a11b", "#64748b", "#0f766e"],
        )
        completion_trend.update_layout(yaxis_title="Completion (%)", xaxis_title="Reporting period")
        st.plotly_chart(chart_theme(completion_trend), width="stretch")
    with trend_right:
        gap_trend = px.bar(
            partner_period.sort_values("Remaining gap", ascending=True),
            x="Remaining gap",
            y="Partner",
            color="Result Area",
            orientation="h",
            barmode="group",
            text="Remaining gap",
            color_discrete_sequence=["#0072bc", "#00aeef", "#20965a", "#e4a11b", "#64748b", "#0f766e"],
            title="Output gaps by partner",
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

with tab_staff:
    st.subheader("Staff currently assigned to flood response")
    st.caption("Roster fields: name, position, duty station, partner, district, Palika, phone, email, and status.")
    uploaded_staff = st.file_uploader("Upload staff roster", type=["xlsx", "csv"], help="Use this for a temporary review, or save the file as partner_files/Staff_Roster.xlsx for automatic loading.")
    staff_df, staff_error = read_staff_roster(uploaded_staff)
    if staff_error:
        st.error(f"Could not read the staff roster: {staff_error}")
    if staff_df.empty:
        st.info("No staff roster has been added yet. Add partner_files/Staff_Roster.xlsx or upload an Excel/CSV roster above.")
        template = pd.DataFrame(columns=STAFF_COLUMNS)
        manual_staff = st.data_editor(template, num_rows="dynamic", width="stretch", hide_index=True, key="manual_staff_roster")
        if st.button("Save staff roster", type="primary"):
            STAFF_SAVE_FILE.parent.mkdir(parents=True, exist_ok=True)
            manual_staff.to_csv(STAFF_SAVE_FILE, index=False)
            st.success("Staff roster saved. It will load automatically next time.")
        st.download_button(
            "Download staff roster template",
            template.to_csv(index=False).encode("utf-8"),
            "Staff_Roster_Template.csv",
            "text/csv",
        )
    else:
        staff_filtered = staff_df.copy()
        if selected_partner != "All partners":
            staff_filtered = staff_filtered[staff_filtered["Partner"].astype(str) == selected_partner]
        if selected_district != "All districts":
            staff_filtered = staff_filtered[staff_filtered["District"].astype(str) == selected_district]
        if selected_palika != "All Palikas":
            staff_filtered = staff_filtered[staff_filtered["Palika"].astype(str) == selected_palika]
        active_status = staff_filtered["Status"].astype(str).str.strip().str.lower().isin(["active", "onboarded", "assigned", "current"])
        if active_status.any():
            staff_filtered = staff_filtered[active_status]
        st.metric("Staff shown", f"{len(staff_filtered):,}")
        edited_staff = st.data_editor(staff_filtered, num_rows="dynamic", width="stretch", hide_index=True, key="staff_roster_editor")
        if st.button("Save staff roster", type="primary"):
            STAFF_SAVE_FILE.parent.mkdir(parents=True, exist_ok=True)
            edited_staff.to_csv(STAFF_SAVE_FILE, index=False)
            st.success("Staff roster saved. Click Reload Excel files to refresh shared filters.")
        st.download_button(
            "Download filtered staff roster",
            staff_filtered.to_csv(index=False).encode("utf-8"),
            "Staff_Roster_Filtered.csv",
            "text/csv",
        )
with tab_data:
    st.subheader("Data entry and export")
    st.caption("Edit existing values directly, add rows with the plus button, or remove rows from this filtered view.")
    export_columns = ["Row ID", "Partner", "Frequency", "District", "Municipality", "Indicator", "Target", "Progress", "Activities"]
    editable_data = st.data_editor(
        filtered_df[export_columns],
        num_rows="dynamic",
        width="stretch",
        hide_index=True,
        key="monitoring_data_editor",
        column_config={
            "Target": st.column_config.NumberColumn(format="%,.0f"),
            "Progress": st.column_config.NumberColumn(format="%,.0f"),
        },
    )
    save_data, download_data = st.columns(2)
    with save_data:
        if st.button("Save entered data", type="primary", width="stretch"):
            DATA_ENTRY_SAVE_FILE.parent.mkdir(parents=True, exist_ok=True)
            editable_data.to_csv(DATA_ENTRY_SAVE_FILE, index=False)
            st.success(f"Saved {len(editable_data):,} rows to {DATA_ENTRY_SAVE_FILE}.")
    with download_data:
        st.download_button(
            "Download filtered CSV",
            editable_data.to_csv(index=False).encode("utf-8"),
            "flood_response_filtered.csv",
            "text/csv",
            width="stretch",
        )
