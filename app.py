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
