import streamlit as st
import pandas as pd
import io
import time
import plotly.express as px
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# --- 1. UI CONFIGURATION ---
st.set_page_config(page_title="VMS Universal Reporting", layout="wide", page_icon="📊")
MASTER_PASSWORD = "VMS@123"

# Enhanced CSS with better contrast and animations
st.markdown("""
    <style>
    .stApp {
        background: linear-gradient(rgba(15, 23, 42, 0.95), rgba(15, 23, 42, 0.95)), 
                    url("https://images.unsplash.com/photo-1451187580459-43490279c0fa?q=80&w=2000");
        background-size: cover; background-attachment: fixed;
    }
    .welcome-note { 
        background: linear-gradient(to right, #00d2ff, #3a7bd5); 
        -webkit-background-clip: text; -webkit-text-fill-color: transparent; 
        font-size: 52px !important; font-weight: 800; text-align: center; margin: 30px 0;
    }
    .glass-metric {
        background: rgba(255, 255, 255, 0.05);
        backdrop-filter: blur(10px);
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 20px;
        padding: 20px;
        transition: transform 0.3s ease;
    }
    .glass-metric:hover { transform: translateY(-5px); background: rgba(255, 255, 255, 0.08); }
    .metric-title { color: #8892b0; font-size: 14px; text-transform: uppercase; letter-spacing: 1px; }
    .metric-value { font-size: 36px; font-weight: 700; color: #64ffda; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. AUTHENTICATION ---
if 'authenticated' not in st.session_state: st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown('<p class="welcome-note">VMS Reporting System</p>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        with st.form("login_form"):
            p = st.text_input("Administrator Password", type="password")
            if st.form_submit_button("Access Dashboard", use_container_width=True):
                if p == MASTER_PASSWORD: 
                    st.session_state.authenticated = True
                    st.rerun()
                else: st.error("Invalid Credentials")
    st.stop()

# --- 3. CORE LOGIC ---
KEYWORDS_TO_IGNORE = ["BADMINTON", "BASKETBALL", "CROSS FITNESS", "SWIMMING", "ZUMBA", "TABLE TENNIS", 
                      "FREESLOT", "FREE SLOT", "SOFT SKILL", "ATOM", "DSA"]
ATT_COL_NAME = "Attended Hours with Approved Leave Percentage"

def is_valid_subject(subject_name):
    s_upper = str(subject_name).upper().strip()
    return not any(bad in s_upper for bad in KEYWORDS_TO_IGNORE)

@st.cache_data
def load_and_clean_data(file):
    # Detect header row
    df_preview = pd.read_excel(file, header=None, nrows=20)
    h_row = 0
    for i, row in df_preview.iterrows():
        if any("ROLL NO" in str(x).upper() for x in row.values):
            h_row = i
            break
    
    df = pd.read_excel(file, header=h_row)
    df.columns = [str(c).strip() for c in df.columns]
    return df, h_row

def process_grid(data_df, cols, batch_subjects, threshold):
    if data_df.empty: return None, None
    
    data_df = data_df.copy()
    data_df[cols['attendance']] = pd.to_numeric(data_df[cols['attendance']], errors='coerce')
    
    # Pivot logic with error handling
    try:
        full_grid = data_df.pivot_table(
            index=[cols['roll'], cols['name'], cols['batch'], cols['sem']],
            columns=cols['subject'], 
            values=cols['attendance'], 
            sort=False
        ).reset_index()
    except Exception as e:
        st.error(f"Data Pivot Error: {e}")
        return None, None
    
    final_subjects = [s for s in batch_subjects if is_valid_subject(s)]
    
    for sub in final_subjects:
        if sub not in full_grid.columns: full_grid[sub] = None
        full_grid[sub] = pd.to_numeric(full_grid[sub], errors='coerce')

    theory_cols = [c for c in final_subjects if not any(x in str(c).upper() for x in ["LAB", "PRACTICAL", "WORKSHOP"])]
    
    full_grid['Theory Avg'] = full_grid[theory_cols].mean(axis=1).round(2)
    full_grid['Final Avg'] = full_grid[final_subjects].mean(axis=1).round(2)
    
    mask = (full_grid[final_subjects] < threshold).any(axis=1)
    shortage_grid = full_grid[mask].copy()
    
    if shortage_grid.empty: return None, pd.Series(0, index=final_subjects)
    
    shortage_grid['Subjects in Shortage'] = (shortage_grid[final_subjects] < threshold).sum(axis=1)
    sub_counts = (shortage_grid[final_subjects] < threshold).sum()
    
    # Cleaning display: Only show values if they are below threshold
    for sub in final_subjects:
        shortage_grid[sub] = shortage_grid[sub].apply(lambda x: x if (pd.notnull(x) and x < threshold) else "")
    
    shortage_grid.insert(0, 'Sl No.', range(1, len(shortage_grid) + 1))
    return shortage_grid, sub_counts

# --- 4. DASHBOARD INTERFACE ---
uploaded_file = st.file_uploader("📂 Upload Universal Attendance Excel File", type=["xlsx"])

if uploaded_file:
    with st.spinner("Analyzing Data..."):
        df, h_row = load_and_clean_data(uploaded_file)
        
        # Mapping Columns
        c_map = {'sem': df.columns[5]} 
        for c in df.columns:
            cs = c.upper()
            if "ROLL NO" in cs: c_map['roll'] = c
            elif "STUDENT NAME" in cs: c_map['name'] = c
            elif "BATCH" in cs: c_map['batch'] = c
            elif any(x in cs for x in ["COURSE", "SUBJECT"]): c_map['subject'] = c
            elif ATT_COL_NAME.upper() in cs: c_map['attendance'] = c

        df = df[df[c_map['subject']].apply(is_valid_subject)]
        df['Dept'] = df[c_map['batch']].astype(str).apply(lambda x: x.split()[0].upper())
        all_subjects = sorted(df[c_map['subject']].unique())
    
    with st.sidebar:
        st.image("https://cdn-icons-png.flaticon.com/512/1055/1055644.png", width=80)
        st.markdown("### ⚙️ Filters")
        threshold = st.slider("Shortage Threshold (%)", 50, 95, 75, 5)
        dept_choice = st.selectbox("Department Focus", ["All"] + sorted(df['Dept'].unique()))
        
        exclude_subjects = st.multiselect("Exclude Subjects", all_subjects)
        
        if st.button("🔴 Secure Logout", use_container_width=True):
            st.session_state.authenticated = False
            st.rerun()

    # Apply Filters
    if exclude_subjects:
        df = df[~df[c_map['subject']].isin(exclude_subjects)]
    if dept_choice != "All":
        df = df[df['Dept'] == dept_choice]

    active_depts = sorted(df['Dept'].unique())
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        summaries = []
        subject_impact = pd.Series(dtype=float)
        
        tabs = st.tabs(["🚀 Command Center"] + [f"🏢 {d}" for d in active_depts])

        for d_idx, dept in enumerate(active_depts):
            d_df = df[df['Dept'] == dept]
            series_list = sorted(list(set(b.split()[0] + " " + b.split()[1] for b in d_df[c_map['batch']].astype(str))))
            
            with tabs[d_idx+1]:
                for series in series_list:
                    s_df = d_df[d_df[c_map['batch']].astype(str).str.contains(series)]
                    s_subs = sorted([s for s in s_df[c_map['subject']].unique() if is_valid_subject(s)])
                    
                    # Section Processing
                    sections = sorted(s_df[c_map['batch']].unique())
                    for sec in sections:
                        sec_df = s_df[s_df[c_map['batch']] == sec]
                        grid, counts = process_grid(sec_df, c_map, s_subs, threshold)
                        
                        if grid is not None and not grid.empty:
                            with st.expander(f"📍 {sec} | {len(grid)-1} Students"):
                                st.dataframe(grid.style.highlight_null(color='#1e293b'), hide_index=True)
                            
                            # Excel Export Logic
                            sn_sec = str(sec).replace("/", "-")[:31]
                            grid.to_excel(writer, sheet_name=sn_sec, index=False)
                            summaries.append({'Section': sec, 'Count': len(grid)-1})
                            subject_impact = subject_impact.add(counts, fill_value=0)

        # Dashboard View
        with tabs[0]:
            if summaries:
                sum_df = pd.DataFrame(summaries)
                
                # Metrics Row
                m_cols = st.columns(4)
                for idx, row in sum_df.iterrows():
                    with m_cols[idx % 4]:
                        st.markdown(f"""
                            <div class="glass-metric">
                                <div class="metric-title">{row['Section']}</div>
                                <div class="metric-value">{row['Count']}</div>
                            </div>
                        """, unsafe_allow_html=True)
                
                st.divider()
                
                # Charts
                c1, c2 = st.columns([6, 4])
                with c1:
                    fig_bar = px.bar(sum_df, x='Section', y='Count', title="Shortage Count by Section",
                                     color='Count', color_continuous_scale='Viridis', template="plotly_dark")
                    st.plotly_chart(fig_bar, use_container_width=True)
                with c2:
                    if not subject_impact.empty:
                        impact_df = subject_impact.nlargest(10).reset_index()
                        impact_df.columns = ['Subject', 'Count']
                        fig_pie = px.pie(impact_df, names='Subject', values='Count', hole=0.5,
                                         title="Top 10 Problematic Subjects", template="plotly_dark")
                        st.plotly_chart(fig_pie, use_container_width=True)
            else:
                st.balloons()
                st.success("Perfect Attendance! No shortages detected.")

    # Download Button
    st.sidebar.divider()
    st.sidebar.download_button(
        label="📥 Download Full Report",
        data=output.getvalue(),
        file_name=f"VMS_Report_{int(time.time())}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
