import streamlit as st
import pandas as pd
import io
import plotly.express as px
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# --- 1. UI CONFIGURATION ---
st.set_page_config(page_title="VMS Universal Reporting", layout="wide")
MASTER_PASSWORD = "VMS@123"

st.markdown("""
    <style>
    .stApp {
        background: linear-gradient(rgba(15, 23, 42, 0.92), rgba(15, 23, 42, 0.92)), 
                    url("https://images.unsplash.com/photo-1451187580459-43490279c0fa?ixlib=rb-1.2.1&auto=format&fit=crop&w=1950&q=80");
        background-size: cover; background-attachment: fixed;
    }
    .welcome-note { 
        background: linear-gradient(to right, #00d2ff, #92fe9d); 
        -webkit-background-clip: text; -webkit-text-fill-color: transparent; 
        font-size: 48px !important; font-weight: 700; text-align: center; margin: 40px 0 10px 0;
    }
    .glass-metric {
        background: rgba(255, 255, 255, 0.03); backdrop-filter: blur(15px);
        border: 1px solid rgba(255, 255, 255, 0.1); border-radius: 15px;
        padding: 25px; margin: 10px 0; text-align: center;
        box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.37);
    }
    .metric-value { font-size: 42px; font-weight: 800; color: #92fe9d; }
    .metric-title { color: #ffffff; font-size: 14px; font-weight: 600; text-transform: uppercase; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. AUTHENTICATION ---
if 'authenticated' not in st.session_state: st.session_state.authenticated = False
if not st.session_state.authenticated:
    st.markdown('<p class="welcome-note">VMS Reporting System</p>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.4, 1])
    with col2:
        p = st.text_input("Password", type="password")
        if st.button("Access Dashboard", use_container_width=True):
            if p == MASTER_PASSWORD: st.session_state.authenticated = True; st.rerun()
            else: st.error("Access Denied")
    st.stop()

# --- 3. CORE LOGIC ---
KEYWORDS_TO_IGNORE = ["BADMINTON", "BASKETBALL", "CROSS FITNESS", "SWIMMING", "ZUMBA", "TABLE TENNIS", 
                      "FREESLOT", "FREE SLOT", "SOFT SKILL", "ATOM", "DSA"]
ATT_COL_NAME = "Attended Hours with Approved Leave Percentage"

def is_valid_subject(subject_name):
    s_upper = str(subject_name).upper()
    return not any(bad in s_upper for bad in KEYWORDS_TO_IGNORE)

def apply_styles(ws):
    thin = Side(style='thin', color="4D4D4D")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    h_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col)
        cell.font, cell.fill, cell.border = Font(bold=True, color="FFFFFF"), h_fill, border
        ws.column_dimensions[cell.column_letter].width = 20

def process_grid(data_df, cols, batch_subjects, low_thresh, high_thresh):
    if data_df.empty: return None
    df_proc = data_df.copy()
    df_proc[cols['attendance']] = pd.to_numeric(df_proc[cols['attendance']], errors='coerce').round(2)
    
    grid = df_proc.pivot_table(index=[cols['roll'], cols['name'], cols['batch']],
                                columns=cols['subject'], values=cols['attendance'], sort=False).reset_index()
    
    # Filter for students who have AT LEAST ONE subject in range
    mask = (grid[batch_subjects] >= low_thresh) & (grid[batch_subjects] <= high_thresh)
    grid = grid[mask.any(axis=1)].copy()
    
    if not grid.empty:
        grid.insert(0, 'Sl No.', range(1, len(grid) + 1))
        return grid
    return None

# --- 4. DASHBOARD INTERFACE ---
uploaded_file = st.file_uploader("📂 Upload Universal Attendance File", type=["xlsx"])

if uploaded_file:
    df_raw = pd.read_excel(uploaded_file, header=None).head(15)
    h_row = 0
    for i, row in df_raw.iterrows():
        if any("ROLL NO" in str(x).upper() for x in row.values):
            h_row = i; break
    
    df = pd.read_excel(uploaded_file, header=h_row)
    c_map = {}
    for c in df.columns:
        cs = str(c).strip()
        if "Roll No" in cs: c_map['roll'] = c
        elif "Student Name" in cs: c_map['name'] = c
        elif "Batch" in cs: c_map['batch'] = c
        elif any(x in cs for x in ["Course", "Subject"]): c_map['subject'] = c
        elif ATT_COL_NAME in cs: c_map['attendance'] = c

    df = df[df[c_map['subject']].apply(is_valid_subject)]
    df['Dept'] = df[c_map['batch']].astype(str).apply(lambda x: x.split()[0].upper())
    
    with st.sidebar:
        st.markdown("### 🛠️ Parameters")
        low_v = st.number_input("From (%)", 0.0, 100.0, 0.0)
        high_v = st.number_input("To (%)", 0.0, 100.0, 75.0)
        dept_choice = st.selectbox("Select Department", ["All"] + sorted(df['Dept'].unique()))
        if st.button("Logout"): st.session_state.authenticated = False; st.rerun()

    active_depts = [dept_choice] if dept_choice != "All" else sorted(df['Dept'].unique())
    output = io.BytesIO()
    
    # Pre-calculate data for the Pivot Summary
    range_df = df[(pd.to_numeric(df[c_map['attendance']], errors='coerce') >= low_v) & 
                  (pd.to_numeric(df[c_map['attendance']], errors='coerce') <= high_v)]

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # 1. ALWAYS WRITE A SUMMARY SHEET FIRST (Prevents your IndexError)
        summary_list = []
        for dept in active_depts:
            count = len(range_df[range_df['Dept'] == dept][c_map['roll']].unique())
            summary_list.append({"Department": dept, "Total Students in Range": count})
        pd.DataFrame(summary_list).to_excel(writer, sheet_name="OVERALL_SUMMARY", index=False)

        tabs = st.tabs(["📊 COMMAND CENTER"] + [f"💎 {d}" for d in active_depts])

        with tabs[0]:
            st.subheader("Global Statistics")
            st.dataframe(pd.DataFrame(summary_list), hide_index=True, use_container_width=True)

        for d_idx, dept in enumerate(active_depts):
            dept_df = df[df['Dept'] == dept]
            dept_range_df = range_df[range_df['Dept'] == dept]
            dept_subs = sorted([s for s in dept_df[c_map['subject']].unique() if is_valid_subject(s)])
            
            with tabs[d_idx+1]:
                st.markdown(f"### 📋 {dept} - Subject wise Summary")
                
                pivot_data = []
                for i, sub in enumerate(dept_subs, 1):
                    sub_students = dept_range_df[dept_range_df[c_map['subject']] == sub]
                    count = len(sub_students)
                    pivot_data.append({"Sl No.": i, "Subject": sub, "Total Students": count})
                    
                    # Expandable list for each subject
                    label = f"{sub} — ({count} Students)"
                    with st.expander(label):
                        if count > 0:
                            st.dataframe(sub_students[[c_map['roll'], c_map['name'], c_map['batch'], c_map['attendance']]], hide_index=True)
                        else:
                            st.info("No students found.")

                # Write Dept Pivot to Excel
                dept_pivot_df = pd.DataFrame(pivot_data)
                dept_pivot_df.to_excel(writer, sheet_name=f"{dept}_Pivot", index=False)
                apply_styles(writer.sheets[f"{dept}_Pivot"])

    st.download_button(f"📥 Download Report", output.getvalue(), "VMS_Report.xlsx", use_container_width=True)
