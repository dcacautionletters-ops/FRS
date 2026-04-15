import streamlit as st
import pandas as pd
import io
import plotly.express as px
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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
        background: rgba(255, 255, 255, 0.05); backdrop-filter: blur(15px);
        border: 1px solid rgba(255, 255, 255, 0.1); border-radius: 15px;
        padding: 20px; margin: 10px 0; text-align: center;
        box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.3);
    }
    .metric-value { font-size: 38px; font-weight: 800; color: #92fe9d; }
    .metric-title { color: #ffffff; font-size: 13px; font-weight: 600; text-transform: uppercase; letter-spacing: 1px;}
    </style>
    """, unsafe_allow_html=True)

# --- 2. AUTHENTICATION ---
if 'authenticated' not in st.session_state: st.session_state.authenticated = False
if not st.session_state.authenticated:
    st.markdown('<p class="welcome-note">VMS Reporting System</p>', unsafe_allow_html=True)
    _, col2, _ = st.columns([1, 1.4, 1])
    with col2:
        p = st.text_input("Security Key", type="password")
        if st.button("Unlock Dashboard", use_container_width=True):
            if p == MASTER_PASSWORD: 
                st.session_state.authenticated = True
                st.rerun()
            else: st.error("Access Denied")
    st.stop()

# --- 3. CORE LOGIC ---
KEYWORDS_TO_IGNORE = ["BADMINTON", "BASKETBALL", "CROSS FITNESS", "SWIMMING", "ZUMBA", "TABLE TENNIS", 
                      "FREESLOT", "FREE SLOT", "SOFT SKILL", "ATOM", "DSA"]
ATT_COL_NAME = "Attended Hours with Approved Leave Percentage"

def is_valid_subject(subject_name):
    s_upper = str(subject_name).upper()
    return not any(bad in s_upper for bad in KEYWORDS_TO_IGNORE)

def get_bracket_summary(data_df, cols, subjects, threshold):
    summary_data = []
    for sub in subjects:
        sub_vals = pd.to_numeric(data_df[data_df[cols['subject']] == sub][cols['attendance']], errors='coerce').dropna().round(2)
        
        row = {
            "Subject": sub,
            "0-49.9%": len(sub_vals[sub_vals < 50]),
            "50-59.9%": len(sub_vals[(sub_vals >= 50) & (sub_vals < 60)]),
            "60-64.4%": len(sub_vals[(sub_vals >= 60) & (sub_vals < 64.5)]),
            "64.5-69.9%": len(sub_vals[(sub_vals >= 64.5) & (sub_vals < 70)]),
            "70-74.9%": len(sub_vals[(sub_vals >= 70) & (sub_vals < 75)]),
            "Total Students": len(sub_vals)
        }
        summary_data.append(row)
    return pd.DataFrame(summary_data)

def apply_styles(ws, threshold, is_summary=False):
    thin = Side(style='thin', color="4D4D4D")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    h_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
    crit_fill = PatternFill(start_color="E74C3C", end_color="E74C3C", fill_type="solid") 
    warn_fill = PatternFill(start_color="F1C40F", end_color="F1C40F", fill_type="solid") 
    
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col)
        cell.font, cell.fill, cell.border = Font(bold=True, color="FFFFFF"), h_fill, border
        ws.column_dimensions[get_column_letter(col)].width = 18

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.border, cell.alignment = border, Alignment(horizontal="center")
            if not is_summary and isinstance(cell.value, (int, float)):
                if cell.value < 70:
                    cell.fill, cell.font = crit_fill, Font(bold=True, color="FFFFFF")
                elif cell.value < threshold:
                    cell.fill, cell.font = warn_fill, Font(bold=True, color="000000")

def process_grid(data_df, cols, batch_subjects, low_thresh, high_thresh, show_all=False):
    if data_df.empty: return None, None
    df_proc = data_df.copy()
    df_proc[cols['attendance']] = pd.to_numeric(df_proc[cols['attendance']], errors='coerce').round(2)
    
    full_grid = df_proc.pivot_table(index=[cols['roll'], cols['name'], cols['batch'], cols['sem']],
                                    columns=cols['subject'], values=cols['attendance'], sort=False).reset_index()
    
    final_subjects = [s for s in batch_subjects if is_valid_subject(s)]
    for sub in final_subjects:
        if sub not in full_grid.columns: full_grid[sub] = None

    # Performance calculation
    theory_cols = [c for c in final_subjects if not any(x in str(c).upper() for x in ["LAB", "PRACTICAL", "WORKSHOP"])]
    full_grid['Theory Avg'] = full_grid[theory_cols].mean(axis=1).round(2)
    full_grid['Overall Avg'] = full_grid[final_subjects].mean(axis=1).round(2)
    
    grid_mask = (full_grid[final_subjects] >= low_thresh) & (full_grid[final_subjects] <= high_thresh)
    shortage_grid = full_grid.copy() if show_all else full_grid[grid_mask.any(axis=1)].copy()
    
    if shortage_grid.empty: return None, None
    
    active_mask = (shortage_grid[final_subjects] >= low_thresh) & (shortage_grid[final_subjects] <= high_thresh)
    shortage_grid['Count (In Range)'] = active_mask.sum(axis=1)
    sub_counts = active_mask.sum()
    
    if not show_all:
        for sub in final_subjects:
            shortage_grid[sub] = shortage_grid[sub].apply(lambda x: x if (pd.notnull(x) and low_thresh <= x <= high_thresh) else "-")
    
    shortage_grid.insert(0, 'Sl No.', range(1, len(shortage_grid) + 1))
    return shortage_grid, sub_counts

# --- 4. DASHBOARD INTERFACE ---
uploaded_file = st.file_uploader("📂 Upload Universal Attendance Excel File", type=["xlsx"])

if uploaded_file:
    # Header logic to find the 'Roll No' row
    df_peek = pd.read_excel(uploaded_file, header=None, nrows=15)
    h_row = 0
    for i, row in df_peek.iterrows():
        if any("ROLL NO" in str(x).upper() for x in row.values):
            h_row = i; break
    
    df = pd.read_excel(uploaded_file, header=h_row)
    df.columns = [str(c).strip() for c in df.columns]
    
    # Intelligent Mapping
    c_map = {'sem': df.columns[5]} 
    for c in df.columns:
        if "Roll No" in c: c_map['roll'] = c
        elif "Student Name" in c: c_map['name'] = c
        elif "Batch" in c: c_map['batch'] = c
        elif any(x in c for x in ["Course", "Subject"]): c_map['subject'] = c
        elif ATT_COL_NAME in c: c_map['attendance'] = c

    df = df[df[c_map['subject']].apply(is_valid_subject)]
    df['Dept'] = df[c_map['batch']].astype(str).apply(lambda x: x.split()[0].upper())
    
    with st.sidebar:
        st.markdown("### 📊 Filter Criteria")
        low_v = st.number_input("Min %", 0.0, 100.0, 0.0)
        high_v = st.number_input("Max %", 0.0, 100.0, 75.0)
        dept_choice = st.selectbox("Department", ["All"] + sorted(df['Dept'].unique()))
        exclude_subs = st.multiselect("Exclude Specific Subjects", sorted(df[c_map['subject']].unique()))
        if st.button("Logout"): st.session_state.authenticated = False; st.rerun()

    if exclude_subs: df = df[~df[c_map['subject']].isin(exclude_subs)]
    active_depts = [dept_choice] if dept_choice != "All" else sorted(df['Dept'].unique())

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        summaries, total_subject_impact = [], pd.Series(dtype=float)
        tabs = st.tabs(["🚀 Command Center"] + [f"🏢 {d}" for d in active_depts])

        for d_idx, dept in enumerate(active_depts):
            d_df = df[df['Dept'] == dept]
            # Identify unique series (e.g., CSE 2022)
            batches = d_df[c_map['batch']].astype(str).unique()
            
            with tabs[d_idx+1]:
                for b_name in sorted(batches):
                    s_df = d_df[d_df[c_map['batch']] == b_name]
                    s_subs = sorted([s for s in s_df[c_map['subject']].unique() if is_valid_subject(s)])
                    
                    # Process Filtered Data
                    grid, counts = process_grid(s_df, c_map, s_subs, low_v, high_v)
                    if grid is not None:
                        st.subheader(f"📍 {b_name}")
                        st.dataframe(grid, hide_index=True)
                        
                        sheet_name = str(b_name).replace("/", "-")[:31]
                        grid.to_excel(writer, sheet_name=sheet_name, index=False)
                        
                        # Add Bracket Summary below the main table
                        bracket_df = get_bracket_summary(s_df, c_map, s_subs, high_v)
                        bracket_df.to_excel(writer, sheet_name=sheet_name, startrow=len(grid)+3, index=False)
                        
                        apply_styles(writer.sheets[sheet_name], high_v)
                        summaries.append({'Section': b_name, 'Count': len(grid)})
                        total_subject_impact = total_subject_impact.add(counts, fill_value=0)

        with tabs[0]:
            if summaries:
                sum_df = pd.DataFrame(summaries)
                cols = st.columns(4)
                for idx, row in sum_df.iterrows():
                    with cols[idx % 4]:
                        st.markdown(f'<div class="glass-metric"><div class="metric-title">{row["Section"]}</div><div class="metric-value">{row["Count"]}</div></div>', unsafe_allow_html=True)
                
                c1, c2 = st.columns(2)
                with c1: st.plotly_chart(px.bar(sum_df, x='Section', y='Count', title="Range Distribution", template="plotly_dark"), use_container_width=True)
                with c2:
                    if not total_subject_impact.empty:
                        impact_df = total_subject_impact.reset_index()
                        impact_df.columns = ['Subject', 'Count']
                        st.plotly_chart(px.pie(impact_df[impact_df['Count']>0], names='Subject', values='Count', hole=0.4, title="Subject Impact", template="plotly_dark"), use_container_width=True)
                
                sum_df.to_excel(writer, sheet_name='MASTER_SUMMARY', index=False)
            else:
                st.info("No students found in the specified percentage range.")

    st.divider()
    st.download_button(f"📥 Download Comprehensive Report", output.getvalue(), "VMS_Analytics_Report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
