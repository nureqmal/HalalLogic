import streamlit as st
import sqlite3
import pandas as pd
import bcrypt

from reportlab.lib.pagesizes import A4
from reportlab.platypus import *
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

# =========================
# DB SETUP
# =========================
conn = sqlite3.connect("ihcs.db", check_same_thread=False)
c = conn.cursor()

c.execute("""CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT UNIQUE,
    password TEXT
)""")

c.execute("""CREATE TABLE IF NOT EXISTS company (
    user_id INTEGER,
    name TEXT,
    address TEXT,
    contact TEXT
)""")

c.execute("""CREATE TABLE IF NOT EXISTS policy (
    user_id INTEGER,
    content TEXT
)""")

c.execute("""CREATE TABLE IF NOT EXISTS materials (
    user_id INTEGER,
    name TEXT,
    supplier TEXT,
    halal_cert TEXT,
    expiry TEXT
)""")

c.execute("""CREATE TABLE IF NOT EXISTS sop (
    user_id INTEGER,
    purchasing TEXT,
    receiving TEXT,
    storage TEXT
)""")

c.execute("""CREATE TABLE IF NOT EXISTS traceability (
    user_id INTEGER,
    system_desc TEXT
)""")

c.execute("""CREATE TABLE IF NOT EXISTS recall (
    user_id INTEGER,
    procedure TEXT
)""")

c.execute("""CREATE TABLE IF NOT EXISTS evaluation (
    user_id INTEGER,
    method TEXT
)""")

conn.commit()

# =========================
# PDF GENERATOR
# =========================
def generate_pdf(company, policy, materials, sop, trace, recall, evaluation):

    filename = "IHCS_Manual.pdf"

    doc = SimpleDocTemplate(filename, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = []

    elements.append(Paragraph("INTERNAL HALAL CONTROL SYSTEM (IHCS) MANUAL", styles['Title']))
    elements.append(Spacer(1, 20))

    # Company
    elements.append(Paragraph("1. Company Profile", styles['Heading2']))
    elements.append(Paragraph(f"Name: {company[0]}", styles['Normal']))
    elements.append(Paragraph(f"Address: {company[1]}", styles['Normal']))
    elements.append(Paragraph(f"Contact: {company[2]}", styles['Normal']))
    elements.append(Spacer(1, 15))

    # Policy
    elements.append(Paragraph("2. Halal Policy", styles['Heading2']))
    elements.append(Paragraph(policy, styles['Normal']))
    elements.append(Spacer(1, 15))

    # Materials
    elements.append(Paragraph("3. Raw Material Masterlist", styles['Heading2']))
    table_data = [["Material", "Supplier", "Halal Cert", "Expiry"]]
    for m in materials:
        table_data.append([m[0], m[1], m[2], m[3]])

    table = Table(table_data)
    table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 1, colors.black),
        ('BACKGROUND', (0,0), (-1,0), colors.grey),
    ]))

    elements.append(table)
    elements.append(Spacer(1, 15))

    # SOP
    elements.append(Paragraph("4. SOP", styles['Heading2']))
    elements.append(Paragraph(f"Purchasing: {sop[0]}", styles['Normal']))
    elements.append(Paragraph(f"Receiving: {sop[1]}", styles['Normal']))
    elements.append(Paragraph(f"Storage: {sop[2]}", styles['Normal']))
    elements.append(Spacer(1, 15))

    # Traceability
    elements.append(Paragraph("5. Traceability", styles['Heading2']))
    elements.append(Paragraph(trace, styles['Normal']))
    elements.append(Spacer(1, 15))

    # Recall
    elements.append(Paragraph("6. Product Recall", styles['Heading2']))
    elements.append(Paragraph(recall, styles['Normal']))
    elements.append(Spacer(1, 15))

    # Evaluation
    elements.append(Paragraph("7. Evaluation", styles['Heading2']))
    elements.append(Paragraph(evaluation, styles['Normal']))

    doc.build(elements)
    return filename

# =========================
# SESSION
# =========================
if "user_id" not in st.session_state:
    st.session_state.user_id = None

# =========================
# AUTH
# =========================
if not st.session_state.user_id:

    st.title("IHCS System Login")

    option = st.radio("Select", ["Login", "Register"])

    username = st.text_input("Username")
    password = st.text_input("Password", type="password")

    if option == "Register":
        if st.button("Register"):
            hashed = bcrypt.hashpw(password.encode(), bcrypt.gensalt())
            try:
                c.execute("INSERT INTO users (username, password) VALUES (?, ?)", (username, hashed))
                conn.commit()
                st.success("Registered!")
            except:
                st.error("Username exists")

    if option == "Login":
        if st.button("Login"):
            c.execute("SELECT id, password FROM users WHERE username=?", (username,))
            user = c.fetchone()

            if user and bcrypt.checkpw(password.encode(), user[1]):
                st.session_state.user_id = user[0]
                st.rerun()
            else:
                st.error("Invalid login")

# =========================
# MAIN APP
# =========================
else:
    st.title("IHCS Manual Generator")
    uid = st.session_state.user_id

    menu = st.sidebar.selectbox("Menu", [
        "Company",
        "Policy",
        "Materials",
        "SOP",
        "Traceability",
        "Recall",
        "Evaluation",
        "Generate"
    ])

    if st.sidebar.button("Logout"):
        st.session_state.user_id = None
        st.rerun()

    # COMPANY
    if menu == "Company":
        name = st.text_input("Company Name")
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

    # MATERIALS
    if menu == "Materials":
        name = st.text_input("Material")
        supplier = st.text_input("Supplier")
        cert = st.text_input("Halal Cert")
        expiry = st.date_input("Expiry")

        if st.button("Add"):
            c.execute("INSERT INTO materials VALUES (?,?,?,?,?)",
                      (uid, name, supplier, cert, str(expiry)))
            conn.commit()

        df = pd.read_sql_query(
            "SELECT name, supplier, halal_cert, expiry FROM materials WHERE user_id=?",
            conn, params=(uid,)
        )
        st.dataframe(df)

    # SOP
    if menu == "SOP":
        purchasing = st.text_area("Purchasing SOP")
        receiving = st.text_area("Receiving SOP")
        storage = st.text_area("Storage SOP")

        if st.button("Save"):
            c.execute("DELETE FROM sop WHERE user_id=?", (uid,))
            c.execute("INSERT INTO sop VALUES (?,?,?,?)",
                      (uid, purchasing, receiving, storage))
            conn.commit()
            st.success("Saved")

    # TRACEABILITY
    if menu == "Traceability":
        desc = st.text_area("Traceability Description")

        if st.button("Save"):
            c.execute("DELETE FROM traceability WHERE user_id=?", (uid,))
            c.execute("INSERT INTO traceability VALUES (?,?)", (uid, desc))
            conn.commit()
            st.success("Saved")

    # RECALL
    if menu == "Recall":
        rec = st.text_area("Recall Procedure")

        if st.button("Save"):
            c.execute("DELETE FROM recall WHERE user_id=?", (uid,))
            c.execute("INSERT INTO recall VALUES (?,?)", (uid, rec))
            conn.commit()
            st.success("Saved")

    # EVALUATION
    if menu == "Evaluation":
        eval_text = st.text_area("Evaluation Method")

        if st.button("Save"):
            c.execute("DELETE FROM evaluation WHERE user_id=?", (uid,))
            c.execute("INSERT INTO evaluation VALUES (?,?)", (uid, eval_text))
            conn.commit()
            st.success("Saved")

    # GENERATE PDF
    if menu == "Generate":
        if st.button("Generate PDF"):

            company = c.execute("SELECT name,address,contact FROM company WHERE user_id=?", (uid,)).fetchone()
            policy = c.execute("SELECT content FROM policy WHERE user_id=?", (uid,)).fetchone()
            materials = c.execute("SELECT name,supplier,halal_cert,expiry FROM materials WHERE user_id=?", (uid,)).fetchall()
            sop = c.execute("SELECT purchasing,receiving,storage FROM sop WHERE user_id=?", (uid,)).fetchone()
            trace = c.execute("SELECT system_desc FROM traceability WHERE user_id=?", (uid,)).fetchone()
            recall = c.execute("SELECT procedure FROM recall WHERE user_id=?", (uid,)).fetchone()
            evaluation = c.execute("SELECT method FROM evaluation WHERE user_id=?", (uid,)).fetchone()

            if not company:
                st.error("Fill Company first")
            else:
                file = generate_pdf(
                    company,
                    policy[0] if policy else "",
                    materials,
                    sop if sop else ("","",""),
                    trace[0] if trace else "",
                    recall[0] if recall else "",
                    evaluation[0] if evaluation else ""
                )

                with open(file, "rb") as f:
                    st.download_button("Download PDF", f, file_name="IHCS_Manual.pdf")
