import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import os
import zipfile
import json
from io import BytesIO
from datetime import datetime

# === Styling ===
st.markdown("""
    <style>
    .stApp {
        background: linear-gradient(135deg, #0f2027, #203a43, #1D4ED8);
        background-size: 400% 400%;
        animation: gradientShift 20s ease infinite;
        color: #f0f0f0;
    }

    @keyframes gradientShift {
        0% {background-position: 0% 50%;}
        50% {background-position: 100% 50%;}
        100% {background-position: 0% 50%;}
    }

    h1, h2, h3, p, label, .stTextInput label {
        color: #ffffff !important;
        font-family: 'Segoe UI', sans-serif;
    }

    .stTextInput input, .stFileUploader, .stDownloadButton button, .stButton button {
        background-color: #1e1e1e !important;
        color: #ffffff !important;
        border: 1px solid #00ffff !important;
        border-radius: 5px;
    }

    .stMetric label {
        color: #ffffff !important;
    }

    hr {
        border: 1px solid #00ffff;
    }
    </style>
""", unsafe_allow_html=True)

# === Title ===
st.markdown("""
    <h1 style='text-align: center;'>🎓 Certificate Generator</h1>
    <h3 style='text-align: center;'>PHN Technology Robotics Scholarship Portal</h3>
    <p style='text-align: center;'>Empowering Innovation • AI • IoT • Robotics • Future Talent</p>
    <hr>
""", unsafe_allow_html=True)

# === Common Password Authentication ===
COMMON_PASSWORD = "phnsecure2025"

email = st.text_input("📧 Enter your PHN Technology email:")
password = st.text_input("🔑 Enter shared password:", type="password")

if not email.endswith("@phntechnology.com"):
    st.error("Access restricted to PHN Technology emails only.")
    st.stop()

if password != COMMON_PASSWORD:
    st.error("Incorrect password.")
    st.stop()

st.success("✅ Access granted.")

# === Load or initialize stats ===
stats_file = "stats.json"
if os.path.exists(stats_file):
    with open(stats_file, "r") as f:
        stats = json.load(f)
else:
    stats = {
        "total_students": 0,
        "qualified": 0,
        "appeared": 0,
        "issued_by": {}
    }

# === File Upload ===
uploaded_file = st.file_uploader("📄 Upload Excel File (.xlsx)", type=["xlsx"])

# === Configuration ===
font_path = os.path.join("fonts", "DancingScript-VariableFont_wght.ttf")
font_name = "DancingScript"
font_size = 23
pdf_file = "phnscholar certificate 6.pdf"

if uploaded_file:
    df = pd.read_excel(uploaded_file, header=1, engine='openpyxl')
    df.columns = df.columns.str.strip()

    if not os.path.exists(font_path):
        st.error(f"Font file not found at: {font_path}")
        st.stop()

    if not os.path.exists(pdf_file):
        st.error(f"Certificate template not found: {pdf_file}")
        st.stop()

    doc = fitz.open(pdf_file)
    page_1 = doc.load_page(0)
    page_2 = doc.load_page(1)

    qualified_pdfs = {}
    not_qualified_pdfs = {}

    if st.button("🎯 Generate Certificates"):
        for _, row in df.iterrows():
            name = str(row['Name']).strip()
            status = str(row['Qualified for Level 2']).strip() if pd.notna(row['Qualified for Level 2']) else ""

            if status == "Qualified":
                template_page = page_2
                y_coord = 401
                target_dict = qualified_pdfs
            else:
                template_page = page_1
                y_coord = 383
                target_dict = not_qualified_pdfs

            new_pdf = fitz.open()
            new_page = new_pdf.new_page(width=template_page.rect.width, height=template_page.rect.height)
            new_page.show_pdf_page(new_page.rect, doc, template_page.number)

            new_page.insert_font(font_name, font_path)

            font = fitz.Font(fontfile=font_path)
            text_width = font.text_length(name, fontsize=font_size)
            page_width = template_page.rect.width
            x_center = (page_width - text_width) / 2

            new_page.insert_text(
                (x_center, y_coord),
                name,
                fontsize=font_size,
                fontname=font_name,
                fill=(0, 0, 0)
            )

            output = BytesIO()
            new_pdf.save(output)
            new_pdf.close()
            target_dict[name] = output.getvalue()

        st.success(f"✅ Certificates generated for {len(df)} students.")

        # === Update persistent stats ===
        qualified_count = df[df['Qualified for Level 2'] == 'Qualified']['Name'].nunique()
        appeared_count = df['Name'].nunique()

        stats["total_students"] += len(df)
        stats["qualified"] += qualified_count
        stats["appeared"] += appeared_count
        stats["issued_by"][email] = stats["issued_by"].get(email, 0) + len(df)

        with open(stats_file, "w") as f:
            json.dump(stats, f)

        # === Log per email
        log_entry = {
            "email": email,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "qualified": qualified_count,
            "appeared": appeared_count,
            "total": len(df)
        }

        log_file = "usage_log.json"
        if os.path.exists(log_file):
            with open(log_file, "r") as f:
                usage_log = json.load(f)
        else:
            usage_log = []

        usage_log.append(log_entry)

        with open(log_file, "w") as f:
            json.dump(usage_log, f, indent=2)

        # === Show stats
        st.subheader("📊 Scholarship Stats (All-Time)")
        st.metric("Total Students Processed", stats["total_students"])
        st.metric("Qualified Students", stats["qualified"])
        st.metric("Appeared Students", stats["appeared"])

        # === Individual download buttons
        for name, pdf_bytes in {**qualified_pdfs, **not_qualified_pdfs}.items():
            st.download_button(
                label=f"📥 Download {name}'s Certificate",
                data=pdf_bytes,
                file_name=f"{name}_certificate.pdf",
                mime="application/pdf"
            )

        # === ZIP download button
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for name, pdf_bytes in {**qualified_pdfs, **not_qualified_pdfs}.items():
                safe_name = "".join(c for c in name if c.isalnum() or c in " _-")
                zip_file.writestr(f"{safe_name}_certificate.pdf", pdf_bytes)
        zip_buffer.seek(0)

        st.download_button(
            label="📦 Download All Certificates as ZIP",
            data=zip_buffer,
            file_name="certificates.zip",
            mime="application/zip"
        )

# 🗑️ Clear button to reset the app safely
if st.button("🗑️ Clear"):
    st.write("🔄 Resetting app...")
    st.stop()
