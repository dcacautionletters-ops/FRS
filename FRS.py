import streamlit as st
import pandas as pd
import io
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# --- 1. UI CONFIGURATION ---
st.set_page_config(page_title="FRS Report Center", layout="wide")
MASTER_PASSWORD = "FRS@123"

st.markdown("""
    <style>
    .stApp {
        background: linear-gradient(rgba(15, 23, 42, 0.95), rgba(15, 23, 42, 0.95)), 
                    url("https://images.unsplash.com/photo-1451187580459-43490279c0fa?ixlib=rb-1.2.1&auto=format&fit=crop&w=1950&q=80");
        background-size: cover; background-attachment: fixed;
    }
    .welcome-note { 
        background: linear-gradient(to right, #00d2ff, #92fe9d); 
        -webkit-background-clip: text; -webkit-text-fill-color: transparent; 
        font-size: 42px !important; font-weight: 700; text-align: center; margin-bottom: 30px;
    }
    .glass-metric {
        background: rgba(255, 255, 255, 0.05); backdrop-filter: blur(15px);
        border: 1px solid rgba(255, 255, 255, 0.1); border-radius: 15px;
        padding: 25px; text-align: center;
        box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.37);
    }
    .metric-title { color: #8899A6; font-size: 14px; text-transform: uppercase; letter-spacing: 1px; }
    .metric-value { font-size: 48px; font-weight: 800; color: #92fe9d; margin-top: 5px; }
    
    /* Center the Download Button & Style it */
    div.stDownloadButton { text-align: center; margin-top: 40px; }
    .stDownloadButton button {
        background: linear-gradient(45deg, #00d2ff, #3a7bd5) !important;
        color: white !important;
        font-size: 20px !important;
        font-weight: bold !important;
        padding: 15px 50px !important;
        border-radius: 12px !important;
        border: none !important;
        box-shadow: 0 4px 15px rgba(0, 210, 255, 0.4);
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. AUTHENTICATION ---
if 'authenticated' not in st.session_state: st.session_state.authenticated = False
if not st.session_state.authenticated:
    st.markdown('<p class="welcome-note">FRS Secure Login</p>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        p = st.text_input("Administrator Password", type="password")
        if st.button("Access Reports", use_container_width=True):
            if p == MASTER_PASSWORD: st.session_state.authenticated = True; st.rerun()
            else: st.error("Incorrect Password")
    st.stop()

# --- 3. CORE LOGIC (Blacklist & Processing) ---
KEYWORDS_TO_IGNORE = ["BADMINTON", "BASKETBALL", "CROSS FITNESS", "SWIMMING", "ZUMBA", "TABLE TENNIS", "FREESLOT", "FREE SLOT", "SOFT SKILL", "ATOM", "DSA"]
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

def process_grid(data_df, cols, batch_subjects, threshold):
    if data_df.empty: return None
    data_df[cols['attendance']] = pd.to_numeric(data_df[cols['attendance']], errors='coerce')
    full_grid = data_df.pivot_table(index=[cols['roll'], cols['name'], cols['batch'], cols['sem']],
                                    columns=cols['subject'], values=cols['attendance'], sort=False).reset_index()
    final_subjects = [s for s in batch_subjects if is_valid_subject(s)]
    
    mask = (full_grid[final_subjects] < threshold).any(axis=1)
    shortage_grid = full_grid[mask].copy()
    if shortage_grid.empty: return None
    
    shortage_grid.insert(0, 'Sl No.', range(1, len(shortage_grid) + 1))
    return shortage_grid

# --- 4. REPORT HUB ---
st.markdown('<p class="welcome-note">FRS Universal Report Center</p>', unsafe_allow_html=True)
uploaded_file = st.file_uploader("📂 Upload Attendance Data (Excel)", type=["xlsx"])

if uploaded_file:
    # Auto-detect header
    df_preview = pd.read_excel(uploaded_file, header=None).head(10)
    h_row = 0
    for i, row in df_preview.iterrows():
        if any("ROLL NO" in str(x).upper() for x in row.values):
            h_row = i; break
    
    df = pd.read_excel(uploaded_file, header=h_row)
    c_map = {'roll': None, 'name': None, 'batch': None, 'subject': None, 'attendance': None, 'sem': df.columns[5]}
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
        st.markdown("### 🛠️ Report Config")
        threshold = st.slider("Shortage Threshold (%)", 50, 95, 75)
        dept_choice = st.selectbox("Select Department", ["All Departments"] + sorted(df['Dept'].unique().tolist()))
        if st.button("Logout"): st.session_state.authenticated = False; st.rerun()

    active_depts = sorted(df['Dept'].unique()) if dept_choice == "All Departments" else [dept_choice]
    
    # Process Data
    output = io.BytesIO()
    summaries = []

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for dept in active_depts:
            d_df = df[df['Dept'] == dept]
            d_subs = sorted([s for s in d_df[c_map['subject']].unique() if is_valid_subject(s)])
            grid = process_grid(d_df, c_map, d_subs, threshold)
            
            if grid is not None:
                sn = f"{dept}_Shortage"[:31]
                grid.to_excel(writer, sheet_name=sn, index=False)
                apply_styles(writer.sheets[sn])
                summaries.append({'Section': dept, 'Count': len(grid)})

    # --- FINAL DISPLAY ---
    if summaries:
        # Glass Metric Summary Row
        cols = st.columns(len(summaries))
        for idx, row in enumerate(summaries):
            with cols[idx]:
                st.markdown(f'''
                    <div class="glass-metric">
                        <div class="metric-title">{row["Section"]}</div>
                        <div class="metric-value">{row["Count"]}</div>
                        <div style="font-size:12px; color:#92fe9d; opacity:0.8;">Total Shortages</div>
                    </div>''', unsafe_allow_html=True)
        
        # Big Download Button
        st.write("###")
        st.download_button(
            label=f"📥 DOWNLOAD {dept_choice.upper()} MASTER REPORT",
            data=output.getvalue(),
            file_name=f"FRS_{dept_choice}_Report.xlsx",
            use_container_width=True
        )
    else:
        st.balloons()
        st.success("Analysis Complete: No shortages found!")
