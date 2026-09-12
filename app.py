import streamlit as st
import pandas as pd
import glob
import os

# Page Configuration
st.set_page_config(
    page_title="UNICEF Emergency Flood Response Dashboard",
    page_icon="💧",
    layout="wide"
)

# UNICEF Brand CSS Styling
st.markdown("""
    <style>
    .overall-card {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 24px;
        box-shadow: 0px 4px 6px rgba(0, 0, 0, 0.02);
        margin-bottom: 20px;
    }
    .big-percent {
        font-size: 48px;
        font-weight: 800;
        color: #00ADEF;
        line-height: 1;
    }
    .card-subtitle {
        font-size: 14px;
        color: #64748B;
        margin-top: 8px;
    }
    .status-text {
        font-size: 14px;
        font-weight: 500;
        color: #334155;
    }
    .kpi-card {
        background-color: #FFFFFF;
        padding: 16px;
        border-radius: 8px;
        border-left: 5px solid #00ADEF;
        box-shadow: 0px 2px 4px rgba(0,0,0,0.05);
    }
    .kpi-title { font-size: 13px; color: #555555; margin-bottom: 4px; }
    .kpi-value { font-size: 22px; font-weight: bold; color: #00ADEF; }
    </style>
""", unsafe_allow_html=True)

# Helper function to safely find column names regardless of casing/naming
def find_column(df, possible_names):
    for col in df.columns:
        clean_col = str(col).strip().lower()
        for name in possible_names:
            if name.lower() in clean_col:
                return col
    return None

# Load data from all 4 partner matrices and clean header/blank rows
@st.cache_data
def load_all_partners():
    partner_mapping = {
        "CDC": "partner_files/CDC_Monitoring Matrix.xlsx",
        "Chaya": "partner_files/Chaya_Monitoring Matrix.xlsx",
        "COSOC": "partner_files/COSOC_Monitoring Matrix.xlsx",
        "SHANTI": "partner_files/SHANTI_Monitoring Matrix.xlsx"
    }
    
    combined_dfs = []
    
    for partner_name, file_path in partner_mapping.items():
        if os.path.exists(file_path):
            try:
                xls = pd.ExcelFile(file_path)
                for sheet in xls.sheet_names:
                    raw_df = pd.read_excel(file_path, sheet_name=sheet, header=None)
                    
                    # Locate header row
                    header_idx = None
                    for idx, row in raw_df.iterrows():
                        row_str = [str(x).upper().strip() for x in row.to_list()]
                        if any(term in row_str for term in ["DISTRICT", "AGENCY NAME", "AGENCY", "SN", "PROVINCE"]):
                            header_idx = idx
                            break
                            
                    if header_idx is not None:
                        df = pd.read_excel(file_path, sheet_name=sheet, skiprows=header_idx)
                    else:
                        df = raw_df
                        
                    if not df.empty and len(df.columns) > 1:
                        # Clean unnamed index columns
                        df = df.loc[:, ~df.columns.astype(str).str.contains('^Unnamed')]
                        
                        # Remove rows that are entirely empty or summary 'TOTAL' rows
                        df = df.dropna(how='all')
                        
                        # Identify main text columns to filter out header repetition and summary totals
                        agency_col = find_column(df, ["agency", "ageny"])
                        if agency_col:
                            df = df[~df[agency_col].astype(str).str.upper().str.contains("TOTAL|AGENCY|AGENV", na=False)]
                        
                        df['Partner'] = partner_name
                        combined_dfs.append(df)
            except Exception:
                pass
            
    if combined_dfs:
        res = pd.concat(combined_dfs, ignore_index=True)
        # Drop completely empty rows
        res = res.dropna(how='all')
        return res
    else:
        return pd.DataFrame({
            "District": ["Rasuwa", "Dhading", "Nuwakot", "Gorkha", "Tanahu", "Rasuwa", "Dhading"],
            "Municipality": ["Gosaikunda", "Nilkanth", "Bidur", "Gorkha", "Vyas", "Uttargaya", "Gajuri"],
            "Partner": ["Chaya", "CDC", "COSOC", "SHANTI", "CDC", "Chaya", "CDC"],
            "Target_Beneficiaries": [1000, 1500, 1200, 2000, 800, 1100, 950],
            "Reached_Beneficiaries": [850, 1200, 600, 1900, 800, 950, 400],
            "Kits_Distributed": [200, 300, 150, 400, 160, 210, 80],
            "Status": ["Achieved", "In Progress", "In Progress", "Achieved", "Achieved", "Achieved", "Not Started"]
        })

df_master = load_all_partners()

# Standardize Key Column Name Detection
district_col = find_column(df_master, ["district", "dist"]) or "District"
status_col = find_column(df_master, ["status", "progress", "state", "achievement", "result"])
target_col = find_column(df_master, ["target", "planned"])
reached_col = find_column(df_master, ["reached", "achieved", "beneficiaries"])
kits_col = find_column(df_master, ["kit", "material", "item"])

# Header
st.title("💧 UNICEF Emergency Flood Response Dashboard")
st.markdown("Multi-Partner & Multi-District Monitoring across **Rasuwa, Dhading, Nuwakot, Gorkha, and Tanahu**.")

# Filter by District
st.write("**Filter by District:**")
district_list = ["All Districts", "Rasuwa", "Dhading", "Nuwakot", "Gorkha", "Tanahu"]
selected_district = st.pills("District Selector", district_list, default="All Districts", label_visibility="collapsed")

# Sidebar Filter for Partners
st.sidebar.title("Partner Filter")
partners_list = ["All Partners", "CDC", "Chaya", "COSOC", "SHANTI"]
selected_partner = st.sidebar.selectbox("Select Partner:", partners_list)

# Data Filtering Logic
filtered_df = df_master.copy()

if selected_district != "All Districts" and district_col in filtered_df.columns:
    filtered_df = filtered_df[filtered_df[district_col].astype(str).str.lower().str.contains(selected_district.lower(), na=False)]

if selected_partner != "All Partners" and "Partner" in filtered_df.columns:
    filtered_df = filtered_df[filtered_df["Partner"] == selected_partner]

# Safe Progress Calculations
total_records = len(filtered_df)

if status_col and status_col in filtered_df.columns:
    status_series = filtered_df[status_col].astype(str).str.lower()
    achieved_count = len(filtered_df[status_series.str.contains("achieved|completed|100%|done|yes", na=False)])
    in_progress_count = len(filtered_df[status_series.str.contains("progress|ongoing|started", na=False)])
    not_started_count = len(filtered_df[status_series.str.contains("not started|pending|0%|no", na=False)])
else:
    achieved_count, in_progress_count, not_started_count = 0, 0, total_records

overall_pct = int((achieved_count / total_records * 100)) if total_records > 0 else 0

# Visual Overall Progress Card
st.markdown(f"""
    <div class="overall-card">
        <div style="display: flex; justify-content: space-between; align-items: center;">
            <div>
                <div class="big-percent">{overall_pct}%</div>
                <div class="card-subtitle">Overall progress across {total_records} response indicators</div>
            </div>
            <div style="text-align: right;">
                <p class="status-text"><span style="color: #10B981;">✓</span> &nbsp; <b>{achieved_count}</b> achieved</p>
                <p class="status-text"><span style="color: #00ADEF;">⚡</span> &nbsp; <b>{in_progress_count}</b> in progress</p>
                <p class="status-text"><span style="color: #6B7280;">ⓘ</span> &nbsp; <b>{not_started_count}</b> not started</p>
            </div>
        </div>
    </div>
""", unsafe_allow_html=True)

# Navigation Tabs
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "Overview", 
    "Activities", 
    "Daily Log", 
    "Trends", 
    "Monitoring", 
    "Data editor"
])

# ---------------- Tab 1: Overview ----------------
with tab1:
    st.subheader("Key Response Metrics")
    
    col1, col2, col3, col4 = st.columns(4)
    total_target = pd.to_numeric(filtered_df[target_col], errors='coerce').sum() if target_col else 0
    total_reached = pd.to_numeric(filtered_df[reached_col], errors='coerce').sum() if reached_col else 0
    total_kits = pd.to_numeric(filtered_df[kits_col], errors='coerce').sum() if kits_col else 0
    reach_pct = (total_reached / total_target * 100) if total_target > 0 else 0
    
    with col1:
        st.markdown(f'<div class="kpi-card"><div class="kpi-title">Target Beneficiaries</div><div class="kpi-value">{int(total_target):,}</div></div>', unsafe_allow_html=True)
    with col2:
        st.markdown(f'<div class="kpi-card"><div class="kpi-title">Reached Beneficiaries</div><div class="kpi-value">{int(total_reached):,}</div></div>', unsafe_allow_html=True)
    with col3:
        st.markdown(f'<div class="kpi-card"><div class="kpi-title">Kits Distributed</div><div class="kpi-value">{int(total_kits):,}</div></div>', unsafe_allow_html=True)
    with col4:
        st.markdown(f'<div class="kpi-card"><div class="kpi-title">Beneficiary Reach (%)</div><div class="kpi-value">{reach_pct:.1f}%</div></div>', unsafe_allow_html=True)
        
    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("Partner & District Data Matrix")
    st.dataframe(filtered_df, use_container_width=True)

# ---------------- Tab 2: Activities ----------------
with tab2:
    st.subheader("Activities Breakdown")
    st.dataframe(filtered_df, use_container_width=True)

# ---------------- Tab 3: Daily Log ----------------
with tab3:
    st.subheader("Daily Field Submissions")
    st.info("Field updates logged per district and partner.")

# ---------------- Tab 4: Trends ----------------
with tab4:
    st.subheader("Progress Trends")
    st.info("Timeline tracking of distribution and beneficiary reach.")

# ---------------- Tab 5: Monitoring ----------------
with tab5:
    st.subheader("Field Monitoring & Quality Assurance")
    st.info("Quality checks and field verification reports.")

# ---------------- Tab 6: Data Editor ----------------
with tab6:
    st.subheader("Live Data Editor")
    st.data_editor(filtered_df, use_container_width=True)
