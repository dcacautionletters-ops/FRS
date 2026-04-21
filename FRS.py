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
        position: relative;
    }
    .metric-value { font-size: 42px; font-weight: 800; color: #92fe9d; }
    .metric-title { color: #ffffff; font-size: 14px; font-weight: 600; text-transform: uppercase; }
    
    /* Interactive Overlay Logic */
    .stButton>button {
        background-color: transparent !important;
        border: none !important;
        color: inherit !important;
        width: 100% !important;
        height: 120px !important;
        margin-bottom: -120px !important;
        z-index: 10;
        position: relative;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. AUTHENTICATION & STATE ---
if 'authenticated' not in st.session_state: st.session_state.authenticated = False
if 'active_sec' not in st.session_state: st.session_state.active_sec = None
if 'active_sub' not in st.session_state: st.session_state.active_sub = None

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

def process_grid(data_df, cols, batch_subjects, low_thresh, high_thresh, show_all=False):
    if data_df.empty: return None, None
    data_df = data_df.copy()
    data_df[cols['attendance']] = pd.to_numeric(data_df[cols['attendance']], errors='coerce').round(2)
    full_grid = data_df.pivot_table(index=[cols['roll'], cols['name'], cols['batch'], cols['sem']],
                                    columns=cols['subject'], values=cols['attendance'], sort=False).reset_index()
    final_subjects = [s for s in batch_subjects if is_valid_subject(s)]
    for sub in final_subjects:
        if sub not in full_grid.columns: full_grid[sub] = None
    theory_cols = [c for c in final_subjects if not any(x in str(c).upper() for x in ["LAB", "PRACTICAL", "WORKSHOP"])]
    full_grid['Theory Avg'] = full_grid[theory_cols].mean(axis=1).round(2)
    full_grid['Final Avg'] = full_grid[final_subjects].mean(axis=1).round(2)
    grid_mask = (full_grid[final_subjects] >= low_thresh) & (full_grid[final_subjects] <= high_thresh)
    shortage_grid = full_grid if show_all else full_grid[grid_mask.any(axis=1)].copy()
    if shortage_grid.empty: return None, None
    active_mask = (shortage_grid[final_subjects] >= low_thresh) & (shortage_grid[final_subjects] <= high_thresh)
    shortage_grid['Subjects in Range'] = active_mask.sum(axis=1)
    sub_counts = active_mask.sum()
    if not show_all:
        for sub in final_subjects:
            shortage_grid[sub] = shortage_grid[sub].apply(lambda x: x if (pd.notnull(x) and low_thresh <= x <= high_thresh) else "")
    shortage_grid.insert(0, 'Sl No.', range(1, len(shortage_grid) + 1))
    return shortage_grid, sub_counts

def apply_styles(ws):
    thin = Side(style='thin', color="4D4D4D")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    h_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col)
        cell.font, cell.fill, cell.border = Font(bold=True, color="FFFFFF"), h_fill, border
        ws.column_dimensions[cell.column_letter].width = 20
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.border, cell.alignment = border, Alignment(horizontal="center")

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
        st.markdown("### 🛠️ Global Parameters")
        low_v = st.number_input("From (%)", 0.00, 100.00, 0.00, 0.01, format="%.2f")
        high_v = st.number_input("To (%)", 0.00, 100.00, 75.00, 0.01, format="%.2f")
        dept_choice = st.selectbox("Select Department", ["All Departments"] + sorted(df['Dept'].unique()))
        if st.button("Reset All Filters"):
            st.session_state.active_sec = None
            st.session_state.active_sub = None
            st.rerun()

    active_depts = [dept_choice] if dept_choice != "All Departments" else sorted(df['Dept'].unique())
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Prevent IndexError: Ensure at least one sheet exists
        pd.DataFrame([["VMS Reporting Generated"]]).to_excel(writer, sheet_name='SUMMARY', index=False, header=False)
        
        summaries, subject_impact = [], pd.Series(dtype=float)
        full_data_store = {}

        tabs = st.tabs(["📊 COMMAND CENTER"] + [f"💎 {d}" for d in active_depts])

        for d_idx, dept in enumerate(active_depts):
            d_df = df[df['Dept'] == dept]
            with tabs[d_idx+1]:
                sections = sorted(d_df[c_map['batch']].unique())
                for sec in sections:
                    sec_df = d_df[d_df[c_map['batch']] == sec]
                    s_subs = sorted([s for s in sec_df[c_map['subject']].unique() if is_valid_subject(s)])
                    grid, counts = process_grid(sec_df, c_map, s_subs, low_v, high_v)
                    if grid is not None:
                        full_data_store[sec] = {'grid': grid, 'counts': counts}
                        summaries.append({'Section': sec, 'Count': len(grid)-1})
                        subject_impact = subject_impact.add(counts, fill_value=0)
                        sn = str(sec).replace("/", "-")[:31]
                        grid.to_excel(writer, sheet_name=sn, index=False)
                        apply_styles(writer.sheets[sn])

        with tabs[0]:
            if summaries:
                sum_df = pd.DataFrame(summaries)
                sum_df.to_excel(writer, sheet_name='SUMMARY', index=False)
                
                # LAYER 2: SUBJECT FOCUS ROWS
                st.markdown("### 📚 Subject Focus")
                for year in ["2024", "2025"]:
                    year_subs = sorted(df[df[c_map['batch']].astype(str).str.contains(year)][c_map['subject']].unique())
                    if year_subs:
                        st.write(f"**MCA {year} Subjects:**")
                        s_cols = st.columns(min(len(year_subs), 6))
                        for i, s_name in enumerate(year_subs):
                            with s_cols[i % 6]:
                                if st.button(s_name, key=f"subbtn_{year}_{s_name}"):
                                    st.session_state.active_sub = s_name
                                    st.session_state.active_sec = None
                
                # LAYER 1: GLASS HEADERS
                st.markdown("### 🏢 Section Overview")
                m_cols = st.columns(min(len(sum_df), 4))
                for idx, row in sum_df.iterrows():
                    with m_cols[idx % 4]:
                        if st.button("", key=f"gls_{row['Section']}"):
                            st.session_state.active_sec = row['Section']
                            st.session_state.active_sub = None
                        st.markdown(f'<div class="glass-metric"><div class="metric-title">{row["Section"]}</div><div class="metric-value">{row["Count"]}</div></div>', unsafe_allow_html=True)
                
                # DYNAMIC RESULTS VIEW
                if st.session_state.active_sec:
                    st.divider()
                    st.subheader(f"📍 Section Report: {st.session_state.active_sec}")
                    st.dataframe(full_data_store[st.session_state.active_sec]['grid'], hide_index=True)
                elif st.session_state.active_sub:
                    st.divider()
                    st.subheader(f"📖 Subject Shortage: {st.session_state.active_sub}")
                    sub_list = [d['grid'][d['grid'][st.session_state.active_sub] != ""] for d in full_data_store.values() if st.session_state.active_sub in d['grid'].columns]
                    if sub_list: st.dataframe(pd.concat(sub_list), hide_index=True)

                # DYNAMIC CHARTS
                c1, c2 = st.columns(2)
                with c1:
                    st.plotly_chart(px.bar(sum_df, x='Section', y='Count', title="Section Distribution", color='Section', template="plotly_dark"), use_container_width=True)
                with c2:
                    impact_plot_df = subject_impact[subject_impact > 0].reset_index()
                    if not impact_plot_df.empty:
                        impact_plot_df.columns = ['Subject', 'Students']
                        st.plotly_chart(px.pie(impact_plot_df, names='Subject', values='Students', hole=0.4, title="Subject Impact", template="plotly_dark"), use_container_width=True)

    st.download_button(f"📥 Download Full Report", output.getvalue(), "VMS_Report.xlsx", use_container_width=True)
