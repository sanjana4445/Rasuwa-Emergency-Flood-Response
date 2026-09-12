import streamlit as st
import pandas as pd
import os

# Page Setup
st.set_page_config(
    page_title="UNICEF Flood Response Dashboard",
    page_icon="💧",
    layout="wide"
)

# -------------------------------------------------------------
# 1. INDIVIDUAL PARTNER LOADERS (Clean & Specific for Each)
# -------------------------------------------------------------

def load_chaya():
    path = "partner_files/Chaya_Monitoring Matrix.xlsx"
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        # Read Excel skipping initial title rows if needed (adjust skiprows=2 or 3 based on Chaya file)
        df = pd.read_excel(path, skiprows=2)
        df = df.dropna(how='all')
        df['Partner'] = 'Chaya'
        return df
    except Exception as e:
        return pd.DataFrame()

def load_cdc():
    path = "partner_files/CDC_Monitoring Matrix.xlsx"
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        df = pd.read_excel(path, skiprows=2)
        df = df.dropna(how='all')
        df['Partner'] = 'CDC'
        return df
    except Exception as e:
        return pd.DataFrame()

def load_cosoc():
    path = "partner_files/COSOC_Monitoring Matrix.xlsx"
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        df = pd.read_excel(path, skiprows=2)
        df = df.dropna(how='all')
        df['Partner'] = 'COSOC'
        return df
    except Exception as e:
        return pd.DataFrame()

def load_shanti():
    path = "partner_files/SHANTI_Monitoring Matrix.xlsx"
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        df = pd.read_excel(path, skiprows=2)
        df = df.dropna(how='all')
        df['Partner'] = 'SHANTI'
        return df
    except Exception as e:
        return pd.DataFrame()


# -------------------------------------------------------------
# 2. COMBINE ALL PARTNERS INTO ONE MASTER DATAFRAME
# -------------------------------------------------------------
@st.cache_data
def load_all_combined_data():
    df_chaya = load_chaya()
    df_cdc = load_cdc()
    df_cosoc = load_cosoc()
    df_shanti = load_shanti()
    
    # Merge all valid dataframes together
    dfs = [df for df in [df_chaya, df_cdc, df_cosoc, df_shanti] if not df.empty]
    
    if dfs:
        master_df = pd.concat(dfs, ignore_index=True)
        # Drop columns named 'Unnamed'
        master_df = master_df.loc[:, ~master_df.columns.astype(str).str.contains('^Unnamed')]
        return master_df
    else:
        return pd.DataFrame()

# -------------------------------------------------------------
# 3. DASHBOARD INTERFACE
# -------------------------------------------------------------
st.title("💧 UNICEF Emergency Flood Response Dashboard")
st.markdown("Multi-Partner Monitoring Dashboard")

df = load_all_combined_data()

if df.empty:
    st.error("No data could be loaded. Please check that Excel files are placed inside the 'partner_files' folder.")
else:
    # Top Navigation / Filter by Partner or Combined
    partner_option = st.selectbox(
        "Select View:",
        ["All Combined Matrix", "Chaya (Rasuwa)", "CDC", "COSOC", "SHANTI"]
    )
    
    # Filter dataset based on selection
    if partner_option == "Chaya (Rasuwa)":
        display_df = df[df["Partner"] == "Chaya"]
    elif partner_option == "CDC":
        display_df = df[df["Partner"] == "CDC"]
    elif partner_option == "COSOC":
        display_df = df[df["Partner"] == "COSOC"]
    elif partner_option == "SHANTI":
        display_df = df[df["Partner"] == "SHANTI"]
    else:
        display_df = df

    st.subheader(f"Data Matrix: {partner_option}")
    st.dataframe(display_df, use_container_width=True)
