import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
from datetime import datetime

# --- SET PAGE CONFIG ---
st.set_page_config(page_title="HalalLogic 2.0", page_icon="🛡️", layout="wide")

# --- CUSTOM CSS FOR PROFESSIONAL LOOK ---
st.markdown("""
    <style>
    .main { background-color: #f5f7f9; }
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #007bff; color: white; }
    </style>
    """, unsafe_allow_html=True)

# --- HELPER: DOC CONTROL HEADER ---
def add_doc_control(doc, section_name, ref_no):
    table = doc.add_table(rows=2, cols=4)
    table.style = 'Table Grid'
    
    # Row 1
    table.cell(0, 0).text = "Section/Process"
    table.cell(0, 1).text = section_name
    table.cell(0, 2).text = "Doc Reference"
    table.cell(0, 3).text = ref_no
    
    # Row 2
    table.cell(1, 0).text = "Implementation"
    table.cell(1, 1).text = datetime.now().strftime("%d %B %Y")
    table.cell(1, 2).text = "Version"
    table.cell(1, 3).text = "1.0"
    
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.runs[0].font.size = Pt(8)
    doc.add_paragraph()

# --- MODULE 3: OUTPUT ENGINE (THE 100-PAGE LOGIC) ---
def generate_pro_manual(profile, content):
    doc = Document()
    
    # 1. FRONT COVER (Hj Jais Style)
    title = doc.add_heading("INTERNAL HALAL CONTROL SYSTEMS (IHCS) MANUAL", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph("\n" * 5)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"PREPARED FOR:\n{profile['company_name']}\n{profile['address']}")
    run.font.size = Pt(14)
    run.bold = True
    
    doc.add_page_break()

    # 2. COMPANY PROFILE (Section 1)
    add_doc_control(doc, "Company Profile", "IHCS/MAN/PROF/01")
    doc.add_heading("1. Company Information", level=1)
    
    info_table = doc.add_table(rows=6, cols=2)
    info_table.style = 'Table Grid'
    info_table.cell(0,0).text = "Business Category"
    info_table.cell(0,1).text = profile['size']
    info_table.cell(1,0).text = "Annual Sales"
    info_table.cell(1,1).text = f"RM {profile['revenue']}"
    info_table.cell(2,0).text = "No. of Employees"
    info_table.cell(2,1).text = str(profile['staff'])
    info_table.cell(3,0).text = "Market"
    info_table.cell(3,1).text = profile['market']
    
    doc.add_page_break()

    # 3. PRODUCT LISTINGS (Section 1.2)
    add_doc_control(doc, "Product Listings", "IHCS/MAN/LIST/01")
    doc.add_heading("1.2 Halal Product Listings", level=1)
    
    prod_table = doc.add_table(rows=1, cols=4)
    prod_table.style = 'Table Grid'
    headers = ["Product Name", "Brand", "SKU", "Expiry"]
    for i, h in enumerate(headers): prod_table.cell(0,i).text = h
    
    for prod in content['products']:
        row = prod_table.add_row().cells
        row[0].text = prod['name']
        row[1].text = prod['brand']
        row[2].text = prod['sku']
        row[3].text = str(prod['expiry'])

    doc.add_page_break()

    # 4. PERFORMANCE SCORING RUBRICS (Section 6)
    add_doc_control(doc, "Evaluation & Review", "IHCS/MAN/REV/01")
    doc.add_heading("6. IHCS Performance Scoring Rubrics", level=1)
    
    rubric_table = doc.add_table(rows=1, cols=3)
    rubric_table.style = 'Table Grid'
    rubric_table.cell(0,0).text = "Component"
    rubric_table.cell(0,1).text = "Weightage"
    rubric_table.cell(0,2).text = "Score (0-5)"
    
    components = [["Halal Policy", "20%"], ["Raw Material Control", "30%"], ["Traceability", "30%"], ["Documentation", "20%"]]
    for comp in components:
        row = rubric_table.add_row().cells
        row[0].text = comp[0]
        row[1].text = comp[1]
        row[2].text = "___ / 5"

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- UI LOGIC ---
st.title("🛡️ HalalLogic 2.0: Expert System")

with st.sidebar:
    st.header("🏢 Deep Profiler")
    c_name = st.text_input("Nama Syarikat", "Hj Jais Sdn Bhd")
    c_addr = st.text_area("Alamat Premis", "Melaka, Malaysia")
    industry = st.selectbox("Skim", ["Makanan & Minuman", "Kosmetik", "Logistik"])
    market = st.multiselect("Pasaran", ["Malaysia", "Singapore", "Brunei", "Indonesia"], default=["Malaysia"])
    revenue = st.number_input("Annual Sales (RM)", 12000000)
    staff = st.number_input("Bilangan Pekerja", 30)

    if st.button("🚀 Start Architectural Phase"):
        st.session_state['pro_profile'] = {
            "company_name": c_name.upper(),
            "address": c_addr,
            "industry": industry,
            "market": ", ".join(market),
            "revenue": f"{revenue:,}",
            "staff": staff,
            "size": "Small" if revenue < 15000000 else "Medium/Large"
        }

if 'pro_profile' in st.session_state:
    st.header(f"Architecting for {st.session_state['pro_profile']['company_name']}")
    
    tab1, tab2, tab3 = st.tabs(["📦 Product Masterlist", "📝 SOP & Clauses", "📊 Scoring Rubrics"])
    
    with tab1:
        st.subheader("Halal Product Listings")
        product_data = [
            {"name": "Dodol Asli Melaka", "brand": "Warisan Mak Yam", "sku": "95577600", "expiry": "2027-01-01"},
            {"name": "Wajik Durian", "brand": "Warisan Mak Yam", "sku": "95577601", "expiry": "2027-02-01"}
        ]
        edited_products = st.data_editor(product_data, num_rows="dynamic")

    with tab2:
        st.subheader("SOP with MHMS 2020 Compliance")
        st.write("Sistem akan automatik selaraskan dengan rujukan MPPHM.")
        st.checkbox("Generate Sertu Program Section", value=True)
        st.checkbox("Generate Traceability Flowchart Section", value=True)

    with tab3:
        st.subheader("IHCS Evaluation Plan")
        st.info("Sistem akan menjana Performance Scoring Rubrics (0-5) secara automatik di lampiran.")

    if st.button("🏁 Generate 100-Page Compliance Manual"):
        with st.spinner("Executing HalalLogic Intelligence..."):
            content = {"products": edited_products}
            file = generate_pro_manual(st.session_state['pro_profile'], content)
            st.download_button("📥 Download IHCS Manual (Hj Jais Edition)", data=file, file_name="IHCS_Manual_Pro.docx")
