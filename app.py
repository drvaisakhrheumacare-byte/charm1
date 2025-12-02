import streamlit as st
import pandas as pd
import docx
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
import secrets
import string
import io
import re
import datetime

# --- CONFIGURATION ---
DEFAULT_SHEET_URL = "https://docs.google.com/spreadsheets/d/1mYapaNzFhSdWTLWaK1cedvQVcjCgvr17EQ-SkEpwV24/edit#gid=0"
COUPON_VALUE = "₹10"

# --- HELPER FUNCTIONS ---

def get_sheet_csv_url(original_url):
    """Extracts ID and GID (Tab ID) to create a CSV export link."""
    try:
        sheet_id_match = re.search(r'/d/([a-zA-Z0-9-_]+)', original_url)
        if not sheet_id_match:
            return None, "Could not find Sheet ID in URL."
        sheet_id = sheet_id_match.group(1)

        gid = "0" 
        gid_match = re.search(r'[#&]gid=([0-9]+)', original_url)
        if gid_match:
            gid = gid_match.group(1)
            
        return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}", None
    except Exception:
        return None, "Invalid URL format."

def generate_secure_code(prefix):
    """Generates unique code: PREFIX-XXX-XXXX"""
    alphabet = string.ascii_uppercase + string.digits
    part1 = ''.join(secrets.choice(alphabet) for _ in range(3))
    part2 = ''.join(secrets.choice(alphabet) for _ in range(4))
    
    clean_prefix = str(prefix).upper().strip()
    if clean_prefix:
        return f"{clean_prefix}-{part1}-{part2}"
    return f"{part1}-{part2}"

def create_coupon_content(cell, name, emp_id, code, date_label):
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    
    # 1. Value
    p1 = cell.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run1 = p1.add_run(f"{COUPON_VALUE} Coupon")
    run1.bold = True
    run1.font.size = Pt(9)
    run1.font.color.rgb = RGBColor(0, 100, 0)
    
    # 2. Unique Code
    p2 = cell.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = p2.add_run(code)
    run2.font.name = 'Courier New'
    run2.font.size = Pt(10)
    run2.bold = True
    
    # 3. Employee Info
    p3 = cell.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run3 = p3.add_run(f"{name}\n({emp_id})")
    run3.font.size = Pt(7)
    
    # 4. Month/Year
    p4 = cell.add_paragraph()
    p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run4 = p4.add_run(date_label)
    run4.font.size = Pt(6)
    run4.font.italic = True

def generate_docx(df, selected_date, prefix):
    doc = docx.Document()
    
    # A4 Layout
    section = doc.sections[0]
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)
    section.top_margin = Cm(1.0)
    section.bottom_margin = Cm(1.0)
    section.left_margin = Cm(1.0)
    section.right_margin = Cm(1.0)

    # Column Detection
    col_name = next((c for c in df.columns if "name" in c.lower()), df.columns[0])
    col_id = next((c for c in df.columns if "code" in c.lower() or "id" in c.lower()), df.columns[1])
    
    st.info(f"Processing... Using column '{col_name}' for Names and '{col_id}' for IDs.")

    progress_bar = st.progress(0)
    total_rows = len(df)

    for index, row in df.iterrows():
        emp_name = str(row[col_name])
        emp_id = str(row[col_id])
        
        table = doc.add_table(rows=13, cols=5)
        table.style = 'Table Grid'
        
        # FIX: Reduced height to 2.0cm to prevent spilling onto next page
        for row_obj in table.rows:
            row_obj.height = Cm(2.0) 
            
        for r in range(13):
            for c in range(5):
                cell = table.cell(r, c)
                unique_code = generate_secure_code(prefix)
                create_coupon_content(cell, emp_name, emp_id, unique_code, selected_date)
        
        # Only add page break if it is NOT the last employee
        if index < len(df) - 1:
            doc.add_page_break()
        
        progress_bar.progress((index + 1) / total_rows)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- MAIN APP UI ---
st.set_page_config(page_title="Coupon Generator", page_icon="🎫")

# 1. PASSWORD CHECK
if "password_correct" not in st.session_state:
    st.session_state["password_correct"] = False

def check_password():
    if st.session_state["password_correct"]: return True
    if "password" not in st.secrets: return True 

    pwd = st.text_input("Enter App Password", type="password")
    if pwd == st.secrets["password"]:
        st.session_state["password_correct"] = True
        return True
    elif pwd:
        st.error("Incorrect Password")
    return False

if check_password():
    st.title("🎫 Monthly Coupon Generator")

    # 2. USER INPUTS
    with st.container():
        sheet_url = st.text_input("Google Sheet URL (Specific Tab)", value=DEFAULT_SHEET_URL)
        
        # NEW: Dropdown Logic
        now = datetime.datetime.now()
        current_year = now.year
        months = ["January", "February", "March", "April", "May", "June", 
                  "July", "August", "September", "October", "November", "December"]
        years = [current_year + i for i in range(11)] # Current to +10 years
        
        col1, col2, col3 = st.columns([1, 1, 2])
        
        with col1:
            sel_month = st.selectbox("Month", options=months, index=now.month - 1)
        with col2:
            sel_year = st.selectbox("Year", options=years, index=0)
        with col3:
            prefix = st.text_input("Code Prefix", value="EMP")

        selected_date = f"{sel_month} {sel_year}"

    # 3. GENERATE BUTTON
    if st.button("Generate Coupons", type="primary"):
        export_url, error = get_sheet_csv_url(sheet_url)
        
        if error:
            st.error(error)
        else:
            with st.spinner('Fetching Sheet & Generating Document...'):
                try:
                    df = pd.read_csv(export_url)
                    df.columns = [c.strip() for c in df.columns]
                    
                    docx_file = generate_docx(df, selected_date, prefix)
                    
                    safe_date = re.sub(r'[^\w\s-]', '', selected_date).strip().replace(' ', '_')
                    file_name = f"Coupons_{safe_date}.docx"

                    st.success(f"✅ Done! Generated for {len(df)} employees.")
                    
                    st.download_button(
                        label="📥 Download Coupon File (DOCX)",
                        data=docx_file,
                        file_name=file_name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                    
                except Exception as e:
                    st.error(f"Error. Please check the URL and permissions.\nDetails: {e}")
