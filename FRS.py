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
    /* Custom style for interactive glass buttons */
    div.stButton > button {
        background: rgba(255, 255, 255, 0.05) !important;
        backdrop-filter: blur(15px);
        border: 1px solid rgba(255, 255, 255, 0.1) !important;
        border-radius: 15px !important;
        padding: 20px !important;
        color: white !important;
        width: 100%;
        transition: 0.3s;
    }
    div.stButton > button:hover {
        background: rgba(255, 255, 255, 0.15) !important;
        border: 1px solid #92fe9d !important;
    }
    .metric-val { font-size: 32px; font-weight: 800; color: #92fe9d; display: block; }
    .metric-lbl { font-size: 12px; text-transform: uppercase; color: #ffffff; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. AUTHENTICATION & STATE ---
if 'authenticated' not in st.session_state: st.session_state.authenticated = False
if 'selected_section' not in st.session_state: st.session_state.selected_section = None

if not st.session_state.authenticated:
    st.markdown('<p class="welcome-note">VMS Reporting System</p>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.4, 1])
    with col2:
        p = st.text_input("Password", type="password")
        if st.button("Access Dashboard"):
            if p == MASTER_PASSWORD: 
                st.session_state.authenticated = True
                st.rerun()
            else: st.error("Access Denied")
    st.stop()

# --- 3. CORE LOGIC ---
KEYWORDS_TO_IGNORE = ["BADMINTON", "BASKETBALL", "CROSS FITNESS", "SWIMMING", "ZUMBA", "TABLE TENNIS", "FREESLOT", "FREE SLOT", "SOFT SKILL", "ATOM", "DSA"]
ATT_COL_NAME = "Attended Hours with Approved Leave Percentage"

def is_valid_subject(subject_name):
    s_upper = str(subject_name).upper()
    return not any(bad in s_upper for bad in KEYWORDS_TO_IGNORE)

def get_bracket_summary(data_df, cols, subjects, threshold):
    summary_data = []
    for sub in subjects:
        sub_vals = pd.to_numeric(data_df[data_df[cols['subject']] == sub][cols['attendance']], errors='coerce').dropna().round(2)
        b1 = len(sub_vals[(sub_vals >= 0) & (sub_vals < 50)])
        b2 = len(sub_vals[(sub_vals >= 50) & (sub_vals < 60)])
        b3a = len(sub_vals[(sub_vals >= 60) & (sub_vals < 64.5)])
        b3b = len(sub_vals[(sub_vals >= 64.5) & (sub_vals < 70)])
        b4 = len(sub_vals[(sub_vals >= 70) & (sub_vals < 75)])
        
        row = {"Subject": sub}
        total = 0
        if threshold > 0: row["0.00-49.99"] = b1; total += b1
        if threshold > 50: row["50.00-59.99"] = b2; total += b2
        if threshold > 60:
            row["60.00-64.49"] = b3a
            row["64.50-69.99"] = b3b
            total += (b3a + b3b)
        if threshold > 70: row["70.00-74.99"] = b4; total += b4
        row["Total"] = total
        summary_data.append(row)
    return pd.DataFrame(summary_data)

def apply_styles(ws, threshold, is_summary=False):
    thin = Side(style='thin', color="4D4D4D")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    h_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
    crit_fill = PatternFill(start_color="C0392B", end_color="C0392B", fill_type="solid") 
    warn_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid") 
    
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col)
        cell.font, cell.fill, cell.border = Font(bold=True, color="FFFFFF"), h_fill, border
        ws.column_dimensions[cell.column_letter].width = 20

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.border, cell.alignment = border, Alignment(horizontal="center")

def process_grid(data_df, cols, batch_subjects, low_thresh, high_thresh, show_all=False):
    if data_df.empty: return None, None
    data_df = data_df.copy()
    data_df[cols['attendance']] = pd.to_numeric(data_df[cols['attendance']], errors='coerce').round(2)
    
    full_grid = data_df.pivot_table(index=[cols['roll'], cols['name'], cols['batch'], cols['sem']],
                                    columns=cols['subject'], values=cols['attendance'], sort=False).reset_index()
    
    final_subjects = [s for s in batch_subjects if is_valid_subject(s)]
    for sub in final_subjects:
        if sub not in full_grid.columns: full_grid[sub] = None

    theory_cols = [c for c in final_subjects if c in full_grid.columns and not any(x in str(c).upper() for x in ["LAB", "PRACTICAL", "WORKSHOP"])]
    full_grid['Theory Avg'] = full_grid[theory_cols].mean(axis=1).round(2) if theory_cols else 0
    full_grid['Final Avg'] = full_grid[final_subjects].mean(axis=1).round(2)
    
    grid_mask = (full_grid[final_subjects] >= low_thresh) & (full_grid[final_subjects] <= high_thresh)
    shortage_grid = full_grid if show_all else full_grid[grid_mask.any(axis=1)].copy()
    
    if shortage_grid.empty: return None, None
    
    shortage_grid.insert(0, 'Sl No.', range(1, len(shortage_grid) + 1))
    return shortage_grid, grid_mask.sum()

# --- 4. DASHBOARD INTERFACE ---
uploaded_file = st.file_uploader("📂 Upload Universal Attendance File", type=["xlsx"])

if uploaded_file:
    df_raw = pd.read_excel(uploaded_file, header=None).head(15)
    h_row = 0
    for i, row in df_raw.iterrows():
        if any("ROLL NO" in str(x).upper() for x in row.values):
            h_row = i; break
    
    df = pd.read_excel(uploaded_file, header=h_row)
    c_map = {'sem': df.columns[5]} 
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
        st.markdown("### 🛠️ Report Controls")
        low_v = st.number_input("Min %", 0.0, 100.0, 0.0)
        high_v = st.number_input("Max % (Limit)", 0.0, 100.0, 75.0)
        dept_choice = st.selectbox("Department", ["All"] + sorted(df['Dept'].unique()))
        
        all_subs = sorted(df[c_map['subject']].unique())
        subject_filter = st.selectbox("Subject-Wise Report (Optional)", ["None"] + all_subs)
        
        if st.button("Reset View"): st.session_state.selected_section = None; st.rerun()
        if st.button("Logout"): st.session_state.authenticated = False; st.rerun()

    active_depts = [dept_choice] if dept_choice != "All" else sorted(df['Dept'].unique())
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        summaries = []
        subject_impact = pd.Series(dtype=float)
        
        tabs = st.tabs(["📊 COMMAND CENTER"] + [f"💎 {d}" for d in active_depts])

        # Logic per Department
        for d_idx, dept in enumerate(active_depts):
            d_df = df[df['Dept'] == dept]
            batches = sorted(d_df[c_map['batch']].unique())
            
            with tabs[d_idx+1]:
                for batch in batches:
                    b_df = d_df[d_df[c_map['batch']] == batch]
                    b_subs = sorted(b_df[c_map['subject']].unique())
                    
                    grid, counts = process_grid(b_df, c_map, b_subs, low_v, high_v)
                    if grid is not None:
                        # Apply Subject Filter if selected
                        display_grid = grid
                        if subject_filter != "None" and subject_filter in grid.columns:
                            display_grid = grid[pd.to_numeric(grid[subject_filter], errors='coerce').between(low_v, high_v)]
                        
                        with st.expander(f"Section: {batch}"):
                            st.dataframe(display_grid, hide_index=True)
                        
                        summaries.append({'Section': batch, 'Count': len(display_grid), 'Data': display_grid})
                        subject_impact = subject_impact.add(counts, fill_value=0)
                        
                        # Excel Export
                        sn = str(batch).replace("/", "-")[:31]
                        display_grid.to_excel(writer, sheet_name=sn, index=False)

        # Command Center Dashboard
        with tabs[0]:
            if summaries:
                sum_df = pd.DataFrame(summaries)
                st.markdown("### Interactive Section Overview")
                m_cols = st.columns(4)
                for idx, row in sum_df.iterrows():
                    with m_cols[idx % 4]:
                        # INTERACTIVE GLASS HEADER
                        if st.button(f"{row['Section']}", key=f"btn_{idx}"):
                            st.session_state.selected_section = row['Section']
                        st.markdown(f'<div style="text-align:center; margin-top:-45px; pointer-events:none;"><span class="metric-val">{row["Count"]}</span><span class="metric-lbl">Shortages</span></div>', unsafe_allow_html=True)

                # Drill-down display
                if st.session_state.selected_section:
                    st.divider()
                    st.subheader(f"Detailed View: {st.session_state.selected_section}")
                    selected_data = next(item['Data'] for item in summaries if item['Section'] == st.session_state.selected_section)
                    st.dataframe(selected_data, use_container_width=True)

                # Charts
                c1, c2 = st.columns(2)
                with c1:
                    st.plotly_chart(px.bar(sum_df, x='Section', y='Count', title="Shortage by Section", template="plotly_dark"), use_container_width=True)
                with c2:
                    impact_df = subject_impact.reset_index()
                    impact_df.columns = ['Subject', 'Count']
                    st.plotly_chart(px.pie(impact_df[impact_df['Count']>0], names='Subject', values='Count', hole=0.4, title="Subject Impact", template="plotly_dark"), use_container_width=True)
            else:
                st.info("No shortages found in the selected range.")

    st.download_button("📥 Download Comprehensive Excel Report", output.getvalue(), "VMS_Shortage_Report.xlsx", use_container_width=True)
