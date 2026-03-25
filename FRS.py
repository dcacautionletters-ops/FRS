import streamlit as st
import pandas as pd
import io
import plotly.express as px
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# --- 1. UI CONFIGURATION ---
st.set_page_config(page_title="FRS Universal Reporting", layout="wide")
MASTER_PASSWORD = "FRS@123"

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
        font-size: 42px !important; font-weight: 700; text-align: center; margin-bottom: 20px;
    }
    .glass-metric {
        background: rgba(255, 255, 255, 0.05); backdrop-filter: blur(15px);
        border: 1px solid rgba(255, 255, 255, 0.1); border-radius: 15px;
        padding: 20px; text-align: center;
        box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.37);
    }
    .metric-title { color: #8899A6; font-size: 14px; text-transform: uppercase; }
    .metric-value { font-size: 36px; font-weight: 800; color: #92fe9d; }
    
    /* Clean up the Download Button */
    .stDownloadButton button {
        background-color: #00d2ff !important;
        color: black !important;
        font-weight: bold !important;
        border-radius: 10px !important;
        border: none !important;
        padding: 15px !important;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. AUTHENTICATION ---
if 'authenticated' not in st.session_state: st.session_state.authenticated = False
if not st.session_state.authenticated:
    st.markdown('<p class="welcome-note">FRS Reporting System</p>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.4, 1])
    with col2:
        p = st.text_input("Security Key", type="password")
        if st.button("Unlock Dashboard", use_container_width=True):
            if p == MASTER_PASSWORD: st.session_state.authenticated = True; st.rerun()
            else: st.error("Access Denied")
    st.stop()

# --- 3. CORE LOGIC (Preserved) ---
KEYWORDS_TO_IGNORE = ["BADMINTON", "BASKETBALL", "CROSS FITNESS", "SWIMMING", "ZUMBA", "TABLE TENNIS", "FREESLOT", "FREE SLOT", "SOFT SKILL", "ATOM", "DSA"]
ATT_COL_NAME = "Attended Hours with Approved Leave Percentage"

def is_valid_subject(subject_name):
    s_upper = str(subject_name).upper()
    return not any(bad in s_upper for bad in KEYWORDS_TO_IGNORE)

def get_bracket_summary(data_df, cols, subjects):
    summary_data = []
    for sub in subjects:
        sub_vals = pd.to_numeric(data_df[data_df[cols['subject']] == sub][cols['attendance']], errors='coerce').dropna()
        summary_data.append({
            "Subject": sub,
            "0.00-49.99": len(sub_vals[sub_vals < 50]),
            "50.00-59.99": len(sub_vals[(sub_vals >= 50) & (sub_vals < 60)]),
            "60.00-69.99": len(sub_vals[(sub_vals >= 60) & (sub_vals < 70)]),
            "70.00-74.99": len(sub_vals[(sub_vals >= 70) & (sub_vals < 75)]),
            "Total": len(sub_vals[sub_vals < 75])
        })
    return pd.DataFrame(summary_data)

def apply_styles(ws, threshold, is_summary=False):
    thin = Side(style='thin', color="4D4D4D")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    h_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col)
        cell.font, cell.fill, cell.border = Font(bold=True, color="FFFFFF"), h_fill, border
        ws.column_dimensions[cell.column_letter].width = 18

def process_grid(data_df, cols, batch_subjects, threshold):
    if data_df.empty: return None, None
    data_df[cols['attendance']] = pd.to_numeric(data_df[cols['attendance']], errors='coerce')
    full_grid = data_df.pivot_table(index=[cols['roll'], cols['name'], cols['batch'], cols['sem']],
                                    columns=cols['subject'], values=cols['attendance'], sort=False).reset_index()
    final_subjects = [s for s in batch_subjects if is_valid_subject(s)]
    for sub in final_subjects:
        if sub not in full_grid.columns: full_grid[sub] = None
    
    mask = (full_grid[final_subjects] < threshold).any(axis=1)
    shortage_grid = full_grid[mask].copy()
    if shortage_grid.empty: return None, None
    
    sub_counts = (shortage_grid[final_subjects] < threshold).sum()
    shortage_grid.insert(0, 'Sl No.', range(1, len(shortage_grid) + 1))
    return shortage_grid, sub_counts

# --- 4. DASHBOARD INTERFACE ---
st.markdown('<p class="welcome-note">Universal Report Command Center</p>', unsafe_allow_html=True)
uploaded_file = st.file_uploader("📂 Drop the attendance Excel here", type=["xlsx"])

if uploaded_file:
    df_preview = pd.read_excel(uploaded_file, header=None).head(15)
    h_row = 0
    for i, row in df_preview.iterrows():
        if any("ROLL NO" in str(x).upper() for x in row.values):
            h_row = i
            break
    
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
        st.markdown("### ⚙️ Adjust Filters")
        threshold = st.slider("Shortage Limit (%)", 50, 95, 75)
        dept_list = ["All Departments"] + sorted(df['Dept'].unique().tolist())
        dept_choice = st.selectbox("Department Focus", dept_list)
        if st.button("Logout"): st.session_state.authenticated = False; st.rerun()

    active_depts = sorted(df['Dept'].unique()) if dept_choice == "All Departments" else [dept_choice]
    
    # Process Data & Visuals
    output = io.BytesIO()
    summaries = []
    subject_impact = pd.Series(dtype=float)

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for dept in active_depts:
            d_df = df[df['Dept'] == dept]
            d_subs = sorted([s for s in d_df[c_map['subject']].unique() if is_valid_subject(s)])
            grid, counts = process_grid(d_df, c_map, d_subs, threshold)
            
            if grid is not None:
                sn = f"{dept}_Report"[:31]
                grid.to_excel(writer, sheet_name=sn, index=False)
                apply_styles(writer.sheets[sn], threshold)
                summaries.append({'Section': dept, 'Count': len(grid)})
                subject_impact = subject_impact.add(counts, fill_value=0)

    # --- UI RENDERING ---
    if summaries:
        sum_df = pd.DataFrame(summaries)
        
        # 1. Glass Highlighters
        m_cols = st.columns(len(sum_df))
        for idx, row in sum_df.iterrows():
            with m_cols[idx]:
                st.markdown(f'''
                    <div class="glass-metric">
                        <div class="metric-title">{row["Section"]}</div>
                        <div class="metric-value">{row["Count"]}</div>
                        <div style="font-size:12px; color:#aaa;">Shortages</div>
                    </div>''', unsafe_allow_html=True)
        
        # 2. Charts Row
        st.write("###")
        c1, c2 = st.columns(2)
        with c1:
            st.plotly_chart(px.bar(sum_df, x='Section', y='Count', title="Shortages by Dept", template="plotly_dark", color_discrete_sequence=['#00d2ff']), use_container_width=True)
        with c2:
            if not subject_impact.empty:
                imp_df = subject_impact.reset_index().rename(columns={'index':'Subject', 0:'Students'})
                st.plotly_chart(px.pie(imp_df.head(10), names='Subject', values='Students', hole=0.4, title="Top Subject Impact", template="plotly_dark"), use_container_width=True)
        
        # 3. Download Button (Center stage)
        st.divider()
        st.download_button(
            label=f"🚀 DOWNLOAD FINAL {dept_choice.upper()} REPORT (EXCEL)",
            data=output.getvalue(),
            file_name=f"FRS_{dept_choice}_Analysis.xlsx",
            use_container_width=True
        )

        # 4. Hidden Data Preview (Optional)
        with st.expander("🔍 View Raw Shortage Lists"):
            st.info("The tables below are included in your Excel download.")
            st.dataframe(df[df[c_map['attendance']] < threshold].head(100), use_container_width=True)
    else:
        st.balloons()
        st.success("Everything looks good! No shortages found for the selected criteria.")
