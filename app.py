import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt
import io

# --- SET PAGE CONFIG ---
st.set_page_config(page_title="HalalLogic Pro", page_icon="🛡️", layout="wide")

# --- MODULE 3: THE OUTPUT ENGINE (PRO) ---
def create_docx(profile, content):
    doc = Document()
    
    # 1. Cover Page
    doc.add_heading('MANUAL SISTEM JAMINAN HALAL (HAS)', 0)
    doc.add_heading(profile['company_name'], level=1)
    doc.add_paragraph(f"Industri: {profile['industry']}")
    doc.add_paragraph(f"Klasifikasi: {profile['size']}")
    doc.add_paragraph(f"Rujukan Standard: MHMS 2020, MPPHM 2020, MS 1500:2019")
    doc.add_page_break()

    # 2. SOP Section with Cross-References
    doc.add_heading('SEKSYEN A: STANDARD OPERATING PROCEDURES (SOP)', level=1)
    for sop in content['sops']:
        doc.add_heading(sop['tajuk'], level=2)
        doc.add_paragraph(sop['isi'])
        # Add Reference Note
        ref_para = doc.add_paragraph(f"Rujukan: {sop['ref']}")
        ref_para.runs[0].italic = True
        ref_para.runs[0].font.size = Pt(9)
    
    doc.add_page_break()

    # 3. HCP Table
    doc.add_heading('SEKSYEN B: ANALISIS HALAL CONTROL POINT (HCP)', level=1)
    table = doc.add_table(rows=1, cols=3)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Titik Proses'
    hdr_cells[1].text = 'Ancaman Halal (Threats)'
    hdr_cells[2].text = 'Tindakan Kawalan (Control)'

    hcp_df = pd.DataFrame(content['hcp'])
    for _, row in hcp_df.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = str(row['proses'])
        row_cells[1].text = str(row['ancaman'])
        row_cells[2].text = str(row['kawalan'])

    doc.add_page_break()

    # 4. Automated Records/Forms (The Pro Touch)
    doc.add_heading('SEKSYEN C: LAMPIRAN REKOD & BORANG', level=1)
    doc.add_paragraph("Borang-borang berikut dijana secara automatik untuk kegunaan operasi harian:")
    for rec in content['records']:
        doc.add_heading(f"Lampiran: {rec['nama']}", level=2)
        doc.add_paragraph(f"Tujuan: {rec['desc']}")
        # Create a simple blank log table for each record
        log_table = doc.add_table(rows=2, cols=3)
        log_table.style = 'Table Grid'
        log_table.rows[0].cells[0].text = "Tarikh"
        log_table.rows[0].cells[1].text = "Aktiviti/Catatan"
        log_table.rows[0].cells[2].text = "Pengesahan"

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- MODULE 2: THE LOGIC BUILDER (PRO) ---
def show_architect(profile):
    st.divider()
    st.subheader(f"🏗️ HalalLogic Architect: {profile['industry']}")
    
    # 1. SOP Library with Cross-Refs
    sop_db = {
        "Makanan & Minuman": [
            {"tajuk": "SOP Kawalan Bahan Mentah", "ref": "MPPHM 2020 Klausa 4.2", "isi": "Memastikan semua bahan mentah mempunyai sijil halal yang sah dari JAKIM/JAIN atau badan luar negara yang diiktiraf."},
            {"tajuk": "SOP Kebersihan Premis", "ref": "MS 1500:2019 Fasal 3.5", "isi": "Premis hendaklah bebas daripada binatang pengerat, serangga dan pencemaran silang material haram."},
        ],
        "Logistik": [
            {"tajuk": "SOP Kawalan Kenderaan", "ref": "MS 2400:2010 (Halalan-Toyyiban)", "isi": "Kenderaan tidak boleh digunakan untuk membawa bahan yang tidak halal secara bersama (mixed cargo)."},
        ]
    }

    # 2. Record/Forms Database
    records_db = [
        {"nama": "Borang Log Penerimaan Bahan Mentah", "desc": "Merekodkan setiap batch bahan yang sampai dari supplier."},
        {"nama": "Borang Kawalan Kebersihan Harian", "desc": "Checklist pembersihan harian premis dan peralatan."},
        {"nama": "Rekod Latihan Kesedaran Halal", "desc": "Merekodkan kehadiran staf ke taklimat Halal bulanan."}
    ]

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("### 📋 1. SOP & Rujukan Spesifik")
        relevant_sops = sop_db.get(profile['industry'], sop_db["Makanan & Minuman"])
        final_sops = []
        for i, sop in enumerate(relevant_sops):
            with st.expander(f"Edit: {sop['tajuk']}"):
                isi = st.text_area(f"Isi SOP ({sop['ref']})", value=sop['isi'], key=f"pro_sop_{i}")
                final_sops.append({"tajuk": sop['tajuk'], "isi": isi, "ref": sop['ref']})

    with col2:
        st.markdown("### 🔍 2. HCP & Flow Analysis")
        hcp_defaults = [
            {"proses": "Penerimaan", "ancaman": "Sijil Tamat Tempoh", "kawalan": "Semak Portal MyeHalal"},
            {"proses": "Penyimpanan", "ancaman": "Pencemaran Silang", "kawalan": "Pengasingan Rak (Separation)"}
        ]
        edited_hcp = st.data_editor(hcp_defaults, num_rows="dynamic", key="pro_hcp")

    st.markdown("### 📂 3. Penjanaan Rekod Automatik")
    st.write("Sistem akan menjana borang log berikut berdasarkan profil anda:")
    selected_records = []
    for rec in records_db:
        if st.checkbox(rec['nama'], value=True):
            selected_records.append(rec)

    return {"sops": final_sops, "hcp": edited_hcp, "records": selected_records}

# --- MAIN CONTROLLER (SIDEBAR PROFILER) ---
st.title("🛡️ HalalLogic Pro")
st.caption("Advanced MHMS 2020 Intelligence Engine")

with st.sidebar:
    st.header("🏢 Company Profiler")
    c_name = st.text_input("Nama Syarikat", "Eqmal Biotech")
    industry = st.selectbox("Skim Pensijilan", ["Makanan & Minuman", "Logistik", "Kosmetik"])
    revenue = st.number_input("Jualan Tahunan (RM)", 300000)
    staff = st.number_input("Bilangan Pekerja", 10)

    if revenue < 300000 or staff < 5: size, req = "Mikro", "Asas"
    else: size, req = "SME", "Penuh"

    if st.button("Jana Struktur Manual"):
        st.session_state['profile_pro'] = {"company_name": c_name.upper(), "industry": industry, "size": size, "req_level": req}

if 'profile_pro' in st.session_state:
    data = show_architect(st.session_state['profile_pro'])
    
    st.divider()
    if st.button("🚀 Download Full Compliance Manual (PRO)"):
        with st.spinner("Compiling cross-references and generating records..."):
            docx_file = create_docx(st.session_state['profile_pro'], data)
            st.download_button(
                label="📥 Download Manual Compliant JAKIM",
                data=docx_file,
                file_name=f"Manual_HAS_Pro_{st.session_state['profile_pro']['company_name']}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
else:
    st.warning("Isi profil di sidebar untuk bermula.")
