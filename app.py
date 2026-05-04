import streamlit as st

def show_profiler():
    st.subheader("🛠️ HalalLogic: Company Profiler")
    st.markdown("Sila lengkapkan profil untuk menentukan *requirement* manual anda.")

    with st.expander("1. Maklumat Asas Syarikat", expanded=True):
        c_name = st.text_input("Nama Syarikat")
        industry = st.selectbox("Skim Pensijilan", 
                                ["Makanan & Minuman", "Kosmetik", "Farmaseutikal", "Logistik", "Barang Gunaan"])

    with st.expander("2. Penentuan Saiz Syarikat (Kriteria SME)", expanded=True):
        revenue = st.number_input("Jualan Tahunan (RM)", min_value=0, step=50000)
        staff = st.number_input("Bilangan Pekerja Tetap", min_value=0, step=1)
        
        # Logic penentuan saiz mengikut standard Malaysia
        if revenue < 300000 or staff < 5:
            company_size = "Mikro"
            req_level = "Asas (Basic)"
        elif revenue <= 15000000 or staff <= 75:
            company_size = "Kecil (Small)"
            req_level = "Sederhana (Intermediate)"
        else:
            company_size = "Sederhana/Besar"
            req_level = "Penuh (Full MHMS 2020)"
            
        st.info(f"Klasifikasi: **{company_size}** | Tahap Keperluan Manual: **{req_level}**")

    return {
        "company_name": c_name,
        "industry": industry,
        "size": company_size,
        "req_level": req_level
    }

# Main App Logic
profile_data = show_profiler()

if st.button("Seterusnya: Bina Struktur Manual"):
    st.session_state['profile'] = profile_data
    st.success(f"Profil {profile_data['company_name']} disimpan. Menjana modul {profile_data['req_level']}...")

# --- MODULE 2: DOCUMENT ARCHITECT ---

def show_architect(profile):
    st.divider()
    st.subheader(f"🏗️ HalalLogic: Architecting for {profile['industry']}")
    st.info(f"Requirement Level: {profile['req_level']}")

    # 1. SOP Library (Database Ringkas)
    # Dalam realiti, kau boleh simpan ni dalam JSON atau Database.
    sop_db = {
        "Makanan & Minuman": [
            {"tajuk": "SOP Kawalan Bahan Mentah", "isi": "Memastikan semua bahan mentah mempunyai sijil halal yang sah dan diiktiraf oleh JAKIM..."},
            {"tajuk": "SOP Kebersihan Premis", "isi": "Premis hendaklah sentiasa bersih dan bebas daripada pencemaran silang..."},
            {"tajuk": "SOP Penyimpanan", "isi": "Bahan halal hendaklah diasingkan dengan jelas daripada bahan yang diragui..."}
        ],
        "Logistik": [
            {"tajuk": "SOP Kawalan Kenderaan", "isi": "Kenderaan pengangkutan mestilah bersih dan tidak digunakan untuk membawa muatan haram..."},
            {"tajuk": "SOP Kawalan Suhu", "isi": "Suhu mestilah dipantau bagi memastikan integriti produk terjaga..."}
        ]
    }

    # Ambil SOP yang relevan
    relevant_sops = sop_db.get(profile['industry'], sop_db["Makanan & Minuman"])

    st.markdown("### 📋 Langkah 1: Pengesahan SOP")
    st.write("Sila semak dan edit SOP yang dicadangkan di bawah:")
    
    final_sops = []
    for index, sop in enumerate(relevant_sops):
        with st.expander(f"{index+1}. {sop['tajuk']}"):
            edited_isi = st.text_area("Edit SOP jika perlu:", value=sop['isi'], key=f"sop_{index}")
            final_sops.append({"tajuk": sop['tajuk'], "isi": edited_isi})

    # 2. HCP Wizard (The Expert Logic)
    st.markdown("### 🔍 Langkah 2: Analisis Risiko (HCP)")
    st.write("Kami telah mencadangkan ancaman Halal berdasarkan proses standard:")

    hcp_data = [
        {"proses": "Penerimaan Bahan", "ancaman": "Sijil Halal Tamat Tempoh / Palsu", "kawalan": "Semak sijil di portal MyeHalal sebelum terima."},
        {"proses": "Penyimpanan", "ancaman": "Pencemaran Silang dengan bahan non-halal", "kawalan": "Sediakan rak berasingan dan label yang jelas."}
    ]

    # Paparkan dalam jadual yang boleh diedit
    edited_hcp = st.data_editor(hcp_data, num_rows="dynamic")

    return {"sops": final_sops, "hcp": edited_hcp}

# --- CONTROLLER LOGIC ---
if 'profile' in st.session_state:
    module2_data = show_architect(st.session_state['profile'])
    
    if st.button("Generate Full Draft Manual"):
        # Nanti kita akan sambungkan ke Docx Generator di Module 3
        st.success("Draft Manual sedang dijana mengikut logik HalalLogic...")
