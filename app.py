import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import os
import zipfile
from io import BytesIO

st.set_page_config(page_title="🎓 Certificate Generator")

st.title("🎓 Certificate Generator")
st.write("Upload an Excel file to generate certificates based on qualification status.")

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

        # Individual download buttons
        for name, pdf_bytes in {**qualified_pdfs, **not_qualified_pdfs}.items():
            st.download_button(
                label=f"📥 Download {name}'s Certificate",
                data=pdf_bytes,
                file_name=f"{name}_certificate.pdf",
                mime="application/pdf"
            )

        # ZIP download button
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

# 🗑️ Clear button to reset the app
clear = st.button("🗑️ Clear")
if clear:
    st.session_state.clear()
    st.experimental_rerun()

