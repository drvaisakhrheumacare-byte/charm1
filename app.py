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
# Default to the link you shared
DEFAULT_SHEET_URL = "https://docs.google.com/spreadsheets/d/1mYapaNzFhSdWTLWaK1cedvQVcjCgvr17EQ-SkEpwV24/edit?gid=0#gid=0"
COUPON_VALUE = "₹10"

# --- HELPER FUNCTIONS ---

def get_sheet_csv_url(original_url):
    """Extracts ID and GID (Tab ID) to create a CSV export link."""
    if not original_url: return None, "Empty URL"
    try:
        # Extract Sheet ID
        sheet_id_match = re.search(r'/d/([a-zA-Z0-9-_]+)', original_url)
        if not sheet_id_match:
            return None, "Could not find Sheet ID."
        sheet_id = sheet_id_match.group(1)

        # Extract GID (The specific tab ID)
        gid = "0" 
        gid_match = re.search(r'[#&?]gid=([0-9]+)', original_url)
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
    
    # Clean prefix
    clean_prefix = str(prefix).upper().strip()
    # Handle empty/invalid prefixes gracefully
    if clean_prefix in ['NAN', 'NONE', '', 'nan']:
        return f"{part1}-{part2}"
        
    return f"{clean_prefix}-{part1}-{part2}"

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

def generate_docx(df, selected_date, default_prefix_input):
    doc = docx.Document()
    
    section = doc.sections[0]
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)
    section.top_margin = Cm(1.0)
    section.bottom_margin = Cm(1.0)
    section.left_margin = Cm(1.0)
    section.right_margin = Cm(1.0)

    # --- INTELLIGENT COLUMN DETECTION ---
    col_name = next((c for c in df.columns if "name" in c.lower()), df.columns[0])
    col_id = next((c for c in df.columns if "code" in c.lower() or "id" in c.lower()), df.columns[1])
    # Look for a column containing "prefix"
    col_prefix = next((c for c in df.columns if "prefix" in c.lower()), None)
    
    st.info(f"Using columns: Name='{col_name}', ID='{col_id}', Prefix='{col_prefix if col_prefix else 'Using Default'}'")

    progress_bar = st.progress(0)
    total_rows = len(df)

    for index, row in df.iterrows():
        emp_name = str(row[col_name])
        emp_id = str(row[col_id])
        
        # --- DETERMINE PREFIX ---
        # 1. Start with the default manual input
        current_prefix = default_prefix_input
        
        # 2. If the sheet has a prefix column, check the value for this row
        if col_prefix:
            row_prefix = str(row[col_prefix])
            # If it's valid (not empty/nan), use it instead of default
            if row_prefix.lower() != 'nan' and row_prefix.strip():
                current_prefix = row_prefix

        # --- HEADER (Top of Page) ---
        header_p = doc.add_paragraph()
        header_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        display_prefix = current_prefix.upper().strip() if current_prefix else "NONE"
        
        run_h = header_p.add_run(f"Name: {emp_name} ({emp_id})  |  Prefix: {display_prefix}  |  Month: {selected_date}")
        run_h.bold = True
        run_h.font.size = Pt(12)
        run_h.font.name = 'Arial'
        header_p.paragraph_format.space_after = Pt(6)

        # --- TABLE (Coupons) ---
        table = doc.add_table(rows=13, cols=5)
        table.style = 'Table Grid'
        
        for row_obj in table.rows:
            row_obj.height = Cm(1.9) 
            
        for r in range(13):
            for c in range(5):
                cell = table.cell(r, c)
                unique_code = generate_secure_code(current_prefix)
                create_coupon_content(cell, emp_name, emp_id, unique_code, selected_date)
        
        # New Page for next employee
        if index < len(df) - 1:
            doc.add_page_break()
        
        progress_bar.progress((index + 1) / total_rows)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- MAIN APP UI ---
st.set_page_config(page_title="Coupon Generator", page_icon="🎫")

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
        st.
