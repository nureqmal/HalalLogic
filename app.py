import streamlit as st
import pandas as pd
from database import init_db, get_connection
from auth import register, login
from pdf_generator import generate_pdf

init_db()

# =========================
# SESSION
# =========================
if "user_id" not in st.session_state:
    st.session_state.user_id = None

# =========================
# AUTH UI
# =========================
if not st.session_state.user_id:

    st.title("IHCS System Login")

    menu = st.radio("Choose", ["Login", "Register"])

    username = st.text_input("Username")
    password = st.text_input("Password", type="password")

    if menu == "Register":
        if st.button("Register"):
            if register(username, password):
                st.success("Registered!")
            else:
                st.error("User exists")

    if menu == "Login":
        if st.button("Login"):
            user_id = login(username, password)
            if user_id:
                st.session_state.user_id = user_id
                st.rerun()
            else:
                st.error("Invalid login")

# =========================
# MAIN SYSTEM
# =========================
else:
    st.title("IHCS Manual Generator")

    conn = get_connection()
    c = conn.cursor()
    uid = st.session_state.user_id

    menu = st.sidebar.selectbox("Menu", [
        "Company",
        "Policy",
        "Materials",
        "Generate"
    ])

    if st.sidebar.button("Logout"):
        st.session_state.user_id = None
        st.rerun()

    # COMPANY
    if menu == "Company":
        name = st.text_input("Name")
        address = st.text_area("Address")
        contact = st.text_input("Contact")

        if st.button("Save"):
            c.execute("DELETE FROM company WHERE user_id=?", (uid,))
            c.execute("INSERT INTO company VALUES (?,?,?,?)", (uid, name, address, contact))
            conn.commit()
            st.success("Saved")

    # POLICY
    if menu == "Policy":
        policy = st.text_area("Halal Policy")

        if st.button("Save"):
            c.execute("DELETE FROM policy WHERE user_id=?", (uid,))
            c.execute("INSERT INTO policy VALUES (?,?)", (uid, policy))
            conn.commit()
            st.success("Saved")

    # MATERIAL
    if menu == "Materials":
        name = st.text_input("Material")
        supplier = st.text_input("Supplier")
        cert = st.text_input("Halal Cert")
        expiry = st.date_input("Expiry")

        if st.button("Add"):
            c.execute("INSERT INTO materials VALUES (?,?,?,?,?)",
                      (uid, name, supplier, cert, str(expiry)))
            conn.commit()

        df = pd.read_sql_query("SELECT name, supplier, halal_cert, expiry FROM materials WHERE user_id=?", conn, params=(uid,))
        st.dataframe(df)

    # GENERATE
    if menu == "Generate":

        if st.button("Generate PDF"):

            company = c.execute("SELECT name,address,contact FROM company WHERE user_id=?", (uid,)).fetchone()
            policy = c.execute("SELECT content FROM policy WHERE user_id=?", (uid,)).fetchone()
            materials = c.execute("SELECT name,supplier,halal_cert,expiry FROM materials WHERE user_id=?", (uid,)).fetchall()

            if not company:
                st.error("Complete data first")
            else:
                file = generate_pdf(company, policy[0] if policy else "", materials)

                with open(file, "rb") as f:
                    st.download_button("Download", f, file_name="IHCS_Manual.pdf")
