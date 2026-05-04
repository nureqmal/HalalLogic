import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
import io

# --- SET PAGE CONFIG ---
st.set_page_config(page_title="HalalLogic 3.0", layout="wide")

# --- STYLE HELPERS (UNTUK TIRU DESIGN HJ JAIS) ---
def set_cell_background(cell, fill_value):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_value}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def add_header_box(doc, section, ref):
    table = doc.add_table(rows=2, cols=4)
    table.style = 'Table Grid'
    # Row 1
    table.cell(0, 0).text = "Section/Process"
    table.cell(0, 1).text = section
    table.cell(0, 2).text = "Doc Reference"
    table.cell(0, 3).text = ref
    # Row 2
    table.cell(1, 0).text = "Implementation"
    table.cell(1, 1).text = "04 May 2026"
    table.cell(1, 2).text = "Version"
    table.cell(1, 3).text = "1.0"
    
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.runs[0].font.size = Pt(8)
    doc.add_paragraph()

# --- THE ARCHITECT (GENERATOR) ---
def generate_hj_jais_manual(data):
    doc = Document()
    
    # --- 1. COVER PAGE ---
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("INTERNAL HALAL CONTROL SYSTEMS (IHCS) MANUAL")
    run.bold = True
    run.font.size = Pt(18)
    
    doc.add_paragraph("\n" * 5)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"PREPARED BY:\n{data['c_name']}\n{data['c_address']}")
    run.font.size = Pt(12)
    
    doc.add_page_break()

    # --- 2. COMPANY PROFILE (Section 1) ---
    add_header_box(doc, "Company Profile", "IHCS/MAN/PROF/0-2021")
    doc.add_heading("1. Company Information", level=1)
    
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    rows = [
        ["Name", data['c_name']],
        ["Address", data['c_address']],
        ["Annual Sales", f"RM {data['revenue']}"],
        ["Employees", str(data['staff'])],
        ["Market", data['market']]
    ]
    for r in rows:
        row_cells = table.add_row().cells
        row_cells[0].text, row_cells[1].text = r
    
    doc.add_page_break()

    # --- 3. PRODUCT LISTING (Section 1.2) ---
    add_header_box(doc, "Operation", "IHCS/MAN/PROF-PROD/00-2021")
    doc.add_heading("1.2 Halal Products Listings", level=1)
    
    prod_table = doc.add_table(rows=1, cols=4)
    prod_table.style = 'Table Grid'
    hdr_cells = prod_table.rows[0].cells
    for i, h in enumerate(["No", "Product Description", "Brand", "Cert Expiry"]):
        hdr_cells[i].text = h
        set_cell_background(hdr_cells[i], "D9D9D9")
    
    for i, p in enumerate(data['products']):
        row_cells = prod_table.add_row().cells
        row_cells[0].text = str(i+1)
        row_cells[1].text = p['Description']
        row_cells[2].text = p['Brand']
        row_cells[3].text = p['Expiry']

    doc.add_page_break()

    # --- 4. SELF-ASSESSMENT (Section 7 - The Long One) ---
    add_header_box(doc, "Evaluation & Review", "IHCS/MAN/EV-SELF/00-2021")
    doc.add_heading("7. IHCS Self-Assessment Tools", level=1)
    
    audit_table = doc.add_table(rows=1, cols=3)
    audit_table.style = 'Table Grid'
    hdr = audit_table.rows[0].cells
    hdr[0].text, hdr[1].text, hdr[2].text = "Question", "Assessment (Y/N)", "Remarks"
    set_cell_background(hdr[0], "D9D9D9"); set_cell_background(hdr[1], "D9D9D9"); set_cell_background(hdr[2], "D9D9D9")
    
    questions = [
        "Company Profile complete and updated?",
        "Halal Policy exist and exhibited?",
        "Raw Material Masterlist updated?",
        "SOP for Purchasing followed?",
        "No usage of brushes from animal source?",
        "Traceability label include batch record?",
        "Muslim worker available at every shift?"
    ]
    for q in questions:
        row = audit_table.add_row().cells
        row[0].text = q
    
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- STREAMLIT UI ---
st.title("🛡️ HalalLogic 3.0")
st.markdown("Automated IHCS Manual Architect [Hj Jais Format]")

col1, col2 = st.columns([1, 2])

with col1:
    st.header("📋 Info Premis")
    c_name = st.text_input("Nama Syarikat", "Eqmal Biotech Sdn Bhd")
    c_addr = st.text_area("Alamat", "Batu Caves, Selangor")
    rev = st.text_input("Annual Sales", "RM 12 Million")
    staff = st.number_input("Bilangan Staf", 30)
    market = st.text_input("Market", "Malaysia, Singapore")

with col2:
    st.header("📦 Product Masterlist")
    st.write("Masukkan produk untuk Section 1.2 manual.")
    items = [{"Description": "Dodol Asli", "Brand": "Warisan", "Expiry": "2027-01-01"}]
    edited_items = st.data_editor(items, num_rows="dynamic")

st.divider()
if st.button("🚀 Jana & Download Manual Lengkap"):
    payload = {
        "c_name": c_name.upper(), "c_address": c_addr, "revenue": rev,
        "staff": staff, "market": market, "products": edited_items
    }
    docx_file = generate_hj_jais_manual(payload)
    st.success("✅ Manual Hj Jais Edition siap dijahit!")
    st.download_button("📥 Download IHCS Manual", data=docx_file, file_name="Manual_IHCS_Pro.docx")
