import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches
import io

# --- SET PAGE CONFIG ---
st.set_page_config(page_title="HalalLogic", page_icon="🛠️")

# --- FUNCTIONS: MODULE 3 (DOCX GENERATOR) ---
def create_docx(profile, content):
    doc = Document()
    
    # 1. Muka Depan (Cover Page)
    doc.add_heading(f"MANUAL SISTEM JAMINAN HALAL (HAS)", 0)
    doc.add_paragraph(f"SYARIKAT: {profile['company_name']}")
    doc.add_paragraph(f"INDUSTRI: {profile['industry']}")
    doc.add_paragraph(f"SAIZ: {profile['size']}")
    doc.add_page_break()

    # 2. Kandungan SOP
    doc.add_heading('SEKSYEN A: STANDARD OPERATING PROCEDURES (SOP)', level=1)
    for sop in content['sops']:
        doc.add_heading(sop['tajuk'], level=2)
        doc.add_paragraph(sop['isi'])
    
    doc.add_page_break()

    # 3. Jadual HCP (Analisis Risiko)
    doc.add_heading('SEKSYEN B: ANALISIS HALAL CONTROL POINT (HCP)', level=1)
    table = doc.add_table(rows=1, cols=3)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Proses'
    hdr_cells[1].text = 'Ancaman Halal'
    hdr_cells[2].text = 'Tindakan Kawalan'

    # Convert edited_hcp (list of dicts) to DataFrame for easier iteration
    hcp_df = pd.DataFrame(content['hcp'])

    for index, row in hcp_df.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = str(row['proses'])
        row_cells[1].text = str(row['ancaman'])
        row_cells[2].text = str(row['kawalan'])

    # Simpan ke memori untuk download
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- FUNCTIONS: MODULE 2 (ARCHITECT) ---
def show_architect(profile):
    st.divider()
    st.subheader(f"🏗️ HalalLogic Architect: {profile['industry']}")
    st.info(f"Requirement: {profile['req_level']}")

    # SOP Database
    sop_db = {
        "Makanan & Minuman": [
            {"tajuk": "SOP Kawalan Bahan Mentah", "isi": "Memastikan semua bahan mentah mempunyai sijil halal yang sah dan diiktiraf oleh JAKIM..."},
            {"tajuk": "SOP Kebersihan Premis", "isi": "Premis hendaklah sentiasa bersih dan bebas daripada pencemaran silang..."},
            {"tajuk": "SOP Penyimpanan", "isi": "Bahan halal hendaklah diasingkan dengan jelas daripada bahan yang diragui..."}
        ],
        "Logistik": [
            {"tajuk": "SOP Kawalan Kenderaan", "isi": "Kenderaan pengangkutan mestilah bersih dan tidak digunakan untuk membawa muatan haram..."},
            {"taj_uk": "SOP Kawalan Suhu", "isi": "Suhu mestilah dipantau bagi memastikan integriti produk terjaga..."}
        ]
    }

    relevant_sops = sop_db.get(profile['industry'], sop_db["Makanan & Minuman"])

    st.markdown("### 📋 Langkah 1: Pengesahan SOP")
    final_sops = []
    for index, sop in enumerate(relevant_sops):
        with st.expander(f"{index+1}. {sop['tajuk']}"):
            edited_isi = st.text_area("Edit SOP:", value=sop['isi'], key=f"sop_{index}", height=150)
            final_sops.append({"tajuk": sop['tajuk'], "isi": edited_isi})

    st.markdown("### 🔍 Langkah 2: Analisis Risiko (HCP)")
    hcp_defaults = [
        {"proses": "Penerimaan Bahan", "ancaman": "Sijil Halal Tamat Tempoh / Palsu", "kawalan": "Semak sijil di portal MyeHalal sebelum terima."},
        {"proses": "Penyimpanan", "ancaman": "Pencemaran Silang", "kawalan": "Sediakan rak berasingan dan label jelas."}
    ]
    
    edited_hcp = st.data_editor(hcp_defaults, num_rows="dynamic", key="hcp_editor")

    return {"sops": final_sops, "hcp": edited_hcp}

# --- MAIN APP LOGIC (MODULE 1: PROFILER) ---
st.title("🛡️ HalalLogic")
st.caption("Automated MHMS 2020 Documentation Engine")

# Sidebar Profiler
with st.sidebar:
    st.header("🏢 Company Profiler")
    c_name = st.text_input("Nama Syarikat", placeholder="cth: Eqmal Food Industries")
    industry = st.selectbox("Skim Pensijilan", 
                            ["Makanan & Minuman", "Kosmetik", "Farmaseutikal", "Logistik", "Barang Gunaan"])
    
    revenue = st.number_input("Jualan Tahunan (RM)", min_value=0, step=10000)
    staff = st.number_input("Bilangan Pekerja", min_value=0, step=1)

    if revenue < 300000 or staff < 5:
        size, req = "Mikro", "Asas (Basic)"
    elif revenue <= 15000000 or staff <= 75:
        size, req = "Kecil (Small)", "Sederhana (Intermediate)"
    else:
        size, req = "Sederhana/Besar", "Penuh (Full MHMS 2020)"

    if st.button("Simpan & Bina Manual"):
        st.session_state['profile'] = {
            "company_name": c_name if c_name else "Syarikat Tanpa Nama",
            "industry": industry,
            "size": size,
            "req_level": req
        }

# Logic Paparan
if 'profile' in st.session_state:
    content_data = show_architect(st.session_state['profile'])
    
    st.divider()
    if st.button("🚀 Generate & Download Full Manual"):
        with st.spinner("Jahit Manual..."):
            docx_file = create_docx(st.session_state['profile'], content_data)
            
            st.success("✅ Manual Berjaya Dijana!")
            st.download_button(
                label="📥 Download Manual HAS (.docx)",
                data=docx_file,
                file_name=f"Manual_HAS_{st.session_state['profile']['company_name']}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
else:
    st.warning("⬅️ Sila isi profil syarikat di sidebar dan klik 'Simpan & Bina Manual' untuk mula.")
