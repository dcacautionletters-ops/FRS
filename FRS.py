import streamlit as st
import pandas as pd
import io
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- 1. UI CONFIGURATION ---
st.set_page_config(page_title="VMS Universal Reporting", layout="wide")
MASTER_PASSWORD = "VMS@123"

st.markdown("""
    <style>
    .stApp { background-color: #0f172a; color: white; }
    .welcome-note { font-size: 42px; font-weight: 700; text-align: center; color: #92fe9d; margin: 20px; }
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
    st.stop()

# --- 3. STYLING HELPER ---
def apply_nice_formatting(ws):
    """Applies professional borders, alignment, and auto-column width."""
    thin = Side(style='thin', color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    
    # Format Headers
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    # Format Data Rows & Auto-Fit Columns
    for column_cells in ws.columns:
        max_length = 0
        column_letter = get_column_letter(column_cells[0].column)
        
        for cell in column_cells:
            cell.border = border
            cell.alignment = Alignment(horizontal="center")
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except: pass
        
        # Drag the column width (with a little extra padding)
        adjusted_width = (max_length + 4)
        ws.column_dimensions[column_letter].width = adjusted_width

# --- 4. CORE LOGIC ---
KEYWORDS_TO_IGNORE = ["BADMINTON", "BASKETBALL", "CROSS FITNESS", "SWIMMING", "ZUMBA", "TABLE TENNIS", "FREESLOT", "FREE SLOT", "SOFT SKILL", "ATOM", "DSA"]
ATT_COL_NAME = "Attended Hours with Approved Leave Percentage"

def is_valid_subject(subject_name):
    s_upper = str(subject_name).upper()
    return not any(bad in s_upper for bad in KEYWORDS_TO_IGNORE)

# --- 5. DASHBOARD INTERFACE ---
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
    df[c_map['attendance']] = pd.to_numeric(df[c_map['attendance']], errors='coerce').fillna(0)
    
    with st.sidebar:
        st.markdown("### 🛠️ Configuration")
        low_v = st.number_input("Min %", 0.0, 100.0, 0.0)
        high_v = st.number_input("Max %", 0.0, 100.0, 75.0)

    # Split by Year/Batch (Universal Logic)
    df['YearBatch'] = df[c_map['batch']].astype(str).apply(lambda x: " ".join(x.split()[:2]))
    batches = sorted(df['YearBatch'].unique())

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        
        for batch_name in batches:
            batch_df = df[df['YearBatch'] == batch_name]
            subjects = sorted(batch_df[c_map['subject']].unique())
            pivot_rows = []

            # 1. Generate Individual Student List Sheets First
            for i, sub in enumerate(subjects, 1):
                sub_data = batch_df[(batch_df[c_map['subject']] == sub) & 
                                    (batch_df[c_map['attendance']] >= low_v) & 
                                    (batch_df[c_map['attendance']] <= high_v)]
                
                count = len(sub_data)
                safe_sub_name = "".join(x for x in str(sub) if x.isalnum())[:25]
                sheet_name = f"{batch_name[:5]}_{safe_sub_name}" 
                
                pivot_rows.append({"Sl No.": i, "Subject": sub, "Total Students": count, "Target": sheet_name})
                
                if not sub_data.empty:
                    sub_data[[c_map['roll'], c_map['name'], c_map['batch'], c_map['attendance']]].to_excel(writer, sheet_name=sheet_name, index=False)
                    apply_nice_formatting(writer.sheets[sheet_name])

            # 2. Create the Nice Pivot Summary Sheet for the Batch
            summary_name = f"{batch_name} Summary"
            pdf = pd.DataFrame(pivot_rows)
            pdf[["Sl No.", "Subject", "Total Students"]].to_excel(writer, sheet_name=summary_name, index=False)
            
            ws_sum = writer.sheets[summary_name]
            apply_nice_formatting(ws_sum)

            # 3. Add Hyperlinks to the Pivot Table
            for idx, row in enumerate(pivot_rows, start=2):
                if row['Total Students'] > 0:
                    cell = ws_sum.cell(row=idx, column=3)
                    cell.hyperlink = f"#'{row['Target']}'!A1"
                    cell.font = Font(color="0000FF", underline="single", bold=True)

    st.success(f"Report for {', '.join(batches)} generated successfully!")
    st.download_button(f"📥 Download Formatted Report", output.getvalue(), "VMS_Formatted_Report.xlsx", use_container_width=True)

    # --- UI DASHBOARD ---
    st.divider()
    curr_batch = st.selectbox("View Preview For:", batches)
    preview_df = df[(df['YearBatch'] == curr_batch) & (df[c_map['attendance']] <= high_v)]
    st.dataframe(preview_df[[c_map['roll'], c_map['name'], c_map['subject'], c_map['attendance']]], use_container_width=True, hide_index=True)
