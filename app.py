from pathlib import Path

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
    :root { --ink:#243331; --muted:#687673; --teal:#0f626b; --line:#d9e2e0; --paper:#fff; --canvas:#f3f6f5; }
    .stApp { background:var(--canvas); color:var(--ink); }
    [data-testid="stHeader"] { background:transparent; }
    .block-container { max-width:1400px; padding:22px 32px 56px; }
    .hero { background:var(--teal); border-radius:18px; color:white; padding:25px 28px; margin-bottom:22px; }
    .hero h1 { margin:0; color:white; font-size:29px; }
    .hero p { margin:7px 0 0; color:#d8e8e9; font-size:15px; }
    .metric-card { background:var(--paper); border:1px solid var(--line); border-radius:12px; padding:17px; min-height:104px; }
    .metric-title { color:#64748b; font-size:12px; font-weight:700; text-transform:uppercase; letter-spacing:.4px; }
    .metric-value { color:#0f172a; font-size:27px; font-weight:750; margin-top:6px; }
    .metric-sub { color:#1594a2; font-size:12px; font-weight:600; margin-top:3px; }
    .section { background:#f8fafc; border:1px solid #e2e8f0; border-radius:10px; padding:13px 18px 4px; margin:20px 0 10px; }
    .section h4 { color:#0f172a; font-size:16px; margin:0 0 2px; }
    .section p { color:#64748b; font-size:13px; margin:0 0 9px; }
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
        <h1>UNICEF Flood Response Dashboard</h1>
        <p>Multi-partner WASH monitoring across target districts, municipalities, and reporting periods.</p>
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
    st.header("Filters")
    if st.button("Reload Excel files", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
    selected_partners = st.multiselect(
        "Partners", sorted(df_master["Partner"].unique()), default=sorted(df_master["Partner"].unique())
    )
    selected_frequency = st.multiselect(
        "Reporting period", ["Daily", "Weekly"], default=["Daily", "Weekly"]
    )
    district_values = sorted(value for value in df_master["District"].dropna().astype(str).unique() if value.strip())
    selected_districts = st.multiselect("Districts", district_values, default=district_values)
    output_values = sorted(df_master["Result Area"].dropna().astype(str).unique())
    selected_outputs = st.multiselect("Output areas", output_values, default=output_values)

filtered_df = df_master[
    df_master["Partner"].isin(selected_partners)
    & df_master["Frequency"].isin(selected_frequency)
    & df_master["Result Area"].isin(selected_outputs)
].copy()
if district_values:
    filtered_df = filtered_df[filtered_df["District"].astype(str).isin(selected_districts)]

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

tab_overview, tab_partner, tab_period, tab_monitoring, tab_data = st.tabs(
    ["Overview", "Partners", "Daily / Weekly", "Monitoring", "Data & export"]
)

with tab_overview:
    st.markdown('<div class="section"><h4>Target versus progress by output</h4><p>Use this view to see which response areas have the largest gaps.</p></div>', unsafe_allow_html=True)
    output_summary = filtered_df.groupby("Result Area", as_index=False)[["Target", "Progress"]].sum()
    chart_data = output_summary.melt("Result Area", var_name="Measure", value_name="Value")
    chart_data["Measure"] = chart_data["Measure"].replace({"Progress": "Achieved"})
    figure = px.bar(
        chart_data, x="Value", y="Result Area", color="Measure", orientation="h", barmode="group",
        color_discrete_map={"Target": "#e4a11b", "Achieved": "#0f626b"}, text="Value",
    )
    figure.update_traces(texttemplate="%{text:,.0f}", textposition="outside", cliponaxis=False)
    st.plotly_chart(chart_theme(figure), use_container_width=True)

with tab_partner:
    partner_summary = filtered_df.groupby("Partner", as_index=False)[["Target", "Progress"]].sum()
    partner_summary["Completion %"] = (partner_summary["Progress"] / partner_summary["Target"].replace(0, 1) * 100).clip(0, 100).round(1)
    left, right = st.columns([6, 4])
    with left:
        partner_chart = px.bar(partner_summary, x="Partner", y=["Target", "Progress"], barmode="group", color_discrete_map={"Target": "#e4a11b", "Progress": "#0f626b"})
        st.plotly_chart(chart_theme(partner_chart), use_container_width=True)
    with right:
        st.dataframe(partner_summary, use_container_width=True, hide_index=True, column_config={"Completion %": st.column_config.NumberColumn(format="%.1f%%")})

with tab_period:
    period_summary = filtered_df.groupby("Frequency", as_index=False)[["Target", "Progress"]].sum()
    period_summary["Completion %"] = (period_summary["Progress"] / period_summary["Target"].replace(0, 1) * 100).clip(0, 100).round(1)
    period_chart = px.bar(period_summary, x="Frequency", y=["Target", "Progress"], barmode="group", color_discrete_map={"Target": "#e4a11b", "Progress": "#0f626b"})
    st.plotly_chart(chart_theme(period_chart), use_container_width=True)
    st.dataframe(filtered_df[["Partner", "Frequency", "District", "Municipality", "Indicator", "Target", "Progress", "Completion %", "Activities"]], use_container_width=True, hide_index=True)

with tab_monitoring:
    monitoring = filtered_df.groupby(["Partner", "Result Area"], as_index=False)[["Target", "Progress"]].sum()
    monitoring["Completion %"] = (monitoring["Progress"] / monitoring["Target"].replace(0, 1) * 100).clip(0, 100).round(1)
    monitoring["Priority"] = monitoring["Completion %"].apply(lambda value: "High" if value < 25 else ("Watch" if value < 75 else "On track"))
    st.dataframe(monitoring.sort_values(["Priority", "Completion %"]), use_container_width=True, hide_index=True)

with tab_data:
    st.subheader("Data review and export")
    st.caption("Update the four Excel files for permanent changes, then click Reload Excel files in the sidebar. This table is a filtered review of the current source data.")
    export_columns = ["Row ID", "Partner", "Frequency", "District", "Municipality", "Indicator", "Target", "Progress", "Activities"]
    st.dataframe(filtered_df[export_columns], use_container_width=True, hide_index=True)
    st.download_button("Download filtered CSV", filtered_df[export_columns].to_csv(index=False).encode("utf-8"), "flood_response_filtered.csv", "text/csv")
