import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import os

# 1. Page Configuration & Styling
st.set_page_config(page_title="Advanced Leave Tracker", layout="wide")
st.markdown("""
    <style>
    .main-title { font-size:32px; font-weight:bold; color:#1F497D; margin-bottom:10px; }
    .metric-box { padding: 15px; background-color: #F8FAFC; border-radius: 8px; border-left: 5px solid #1F497D; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
    .metric-num { font-size: 28px; font-weight: bold; color: #1F497D; }
    .metric-lbl { font-size: 14px; color: #595959; font-weight: 500; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">🎯 Enterprise Leave Tracking & Analytics Tool</div>', unsafe_allow_html=True)

# 📂 OFFICE EXCEL SHEETS PATH CONFIGURATION
# Aap in file names ko apni original office files ke naam se badal sakte hain
MASTER_FILE = "Employee_Master.xlsx"  # Isme saare employees ki details hain
LEAVE_FILE = "Leave_Records.xlsx"    # Isme saari leaves track hoti hain

# 2. Dynamic Data Loading from Office Excel Sheets
def load_employee_data():
    if os.path.exists(MASTER_FILE):
        return pd.read_excel(MASTER_FILE)
    else:
        st.error(f"Error: '{MASTER_FILE}' nahi mili! Kripya file ko sahi folder mein rakhein.")
        return pd.DataFrame(columns=["Emp ID", "Name", "Team", "PL_Quota", "SL_Quota"])

def load_leave_records():
    if os.path.exists(LEAVE_FILE):
        df = pd.read_excel(LEAVE_FILE)
        df['Date'] = pd.to_datetime(df['Date']).dt.date
        return df
    else:
        # Agar file nahi hai toh blank sheet bana dega unhi columns ke sath
        return pd.DataFrame(columns=["Emp ID", "Date", "Leave Type"])

# Load Data Dynamically
df_emp = load_employee_data()

if 'df_leaves' not in st.session_state:
    st.session_state.df_leaves = load_leave_records()

# 3. Sidebar Navigation
st.sidebar.header("🛠️ Admin Control Center")
menu = st.sidebar.radio("Navigate Menu", ["Live Dashboard", "Log New Leave Request", "Employee Directory"])
selected_date = st.sidebar.date_input("📅 Operations Date Check", date(2026, 1, 2))

# --- MENU 1: LIVE DASHBOARD ---
if menu == "Live Dashboard" and not df_emp.empty:
    active_leaves = st.session_state.df_leaves[st.session_state.df_leaves["Date"] == selected_date]
    df_dashboard = df_emp.merge(active_leaves, on="Emp ID", how="left")
    
    is_weekend = selected_date.weekday() in [5, 6]
    df_dashboard["Status Today"] = "WO" if is_weekend else df_dashboard["Leave Type"].fillna("Present")
    
    # Live Quota Calculation from Excel Data
    pl_counts = st.session_state.df_leaves[st.session_state.df_leaves["Leave Type"] == "PL"].groupby("Emp ID").size().to_dict()
    sl_counts_full = st.session_state.df_leaves[st.session_state.df_leaves["Leave Type"] == "SL"].groupby("Emp ID").size().to_dict()
    sl_counts_half = (st.session_state.df_leaves[st.session_state.df_leaves["Leave Type"] == "SL 1/2"].groupby("Emp ID").size() * 0.5).to_dict()
    
    df_dashboard["PL Taken"] = df_dashboard["Emp ID"].map(pl_counts).fillna(0)
    df_dashboard["SL Taken"] = df_dashboard["Emp ID"].map(sl_counts_full).fillna(0) + df_dashboard["Emp ID"].map(sl_counts_half).fillna(0)
    df_dashboard["Total Balance Remaining"] = (df_dashboard["PL_Quota"] + df_dashboard["SL_Quota"]) - (df_dashboard["PL Taken"] + df_dashboard["SL Taken"])

    on_leave = df_dashboard[df_dashboard["Status Today"].isin(["PL", "SL", "SL 1/2", "PH"])].shape[0]
    
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Staff", len(df_emp))
    col2.metric("Present Today", len(df_emp) - on_leave if not is_weekend else 0)
    col3.metric("On Leave", on_leave)

    st.markdown("### 📋 Live Status Table (Fetched from Office Sheets)")
    st.dataframe(df_dashboard[["Emp ID", "Name", "Team", "Status Today", "PL Taken", "SL Taken", "Total Balance Remaining"]], use_container_width=True, hide_index=True)

# --- MENU 2: LOG NEW LEAVE (Updates Office Excel Sheet) ---
elif menu == "Log New Leave Request" and not df_emp.empty:
    st.markdown("### 📝 Apply & Update New Leave in Office Records")
    
    with st.form("leave_form", clear_on_submit=True):
        emp_name = st.selectbox("Select Employee", df_emp["Name"].tolist())
        l_type = st.selectbox("Type", ["PL", "SL", "SL 1/2", "PH"])
        d_start = st.date_input("Start Date", date.today())
        d_end = st.date_input("End Date", date.today())
        
        if st.form_submit_button("Sanction & Save to Excel"):
            emp_id = df_emp[df_emp["Name"] == emp_name]["Emp ID"].values[0]
            
            new_entries = []
            current_day = d_start
            while current_day <= d_end:
                new_entries.append({
                    "Emp ID": int(emp_id),
                    "Date": current_day,
                    "Leave Type": l_type
                })
                current_day += timedelta(days=1)
            
            new_df = pd.DataFrame(new_entries)
            updated = pd.concat([st.session_state.df_leaves, new_df]).drop_duplicates(subset=["Emp ID", "Date"], keep="last")
            
            # 💾 DIRECT UPDATE TO OFFICE EXCEL FILE
            updated.to_excel(LEAVE_FILE, index=False)
            st.session_state.df_leaves = updated
            st.success(f"Success! Data directly updated in '{LEAVE_FILE}' for {emp_name}!")
            st.balloons()

elif menu == "Employee Directory":
    st.dataframe(df_emp, use_container_width=True, hide_index=True)