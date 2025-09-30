import os
import pandas as pd
import fitz  # PyMuPDF
import streamlit as st
from io import BytesIO

# === Streamlit UI ===
st.set_page_config(page_title="Certificate Generator", layout="centered")
st.title("🎓 Certificate Generator")
st.markdown("Upload an Excel file and generate certificates for your students.")

# === File Upload ===
uploaded_excel = st.file_uploader("📄 Upload Excel File (.xlsx)", type=["xlsx"])
pdf_template = "phnscholar certificate 6.pdf"
font_path = "fonts/DancingScript-VariableFont_wght.ttf"
font_name = "DancingScript"
font_size = 23

if uploaded_excel:
    df = pd.read_excel(uploaded_excel, header=1, engine='openpyxl')
    df.columns = df.columns.str.strip()

    # === Load certificate template ===
    doc = fitz.open(pdf_template)
    page_1 = doc.load_page(0)
    page_2 = doc.load_page(1)

    # === Create output folders ===
    base_folder = os.path.splitext(pdf_template)[0]
    qualified_folder = os.path.join(base_folder, "Qualified")
    not_qualified_folder = os.path.join(base_folder, "Not_Qualified")
    os.makedirs(qualified_folder, exist_ok=True)
    os.makedirs(not_qualified_folder, exist_ok=True)

    # === Generate certificates ===
    with st.spinner("Generating certificates..."):
        for _, row in df.iterrows():
            name = str(row['Name']).strip()
            status = str(row['Qualified for Level 2']).strip() if pd.notna(row['Qualified for Level 2']) else ""

            if status == "Qualified":
                template_page = page_2
                output_folder = qualified_folder
                y_coord = 401
            else:
                template_page = page_1
                output_folder = not_qualified_folder
                y_coord = 383

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

            safe_name = "".join(c for c in name if c.isalnum() or c in " _-")
            output_path = os.path.join(output_folder, f"{safe_name}.pdf")
            new_pdf.save(output_path)
            new_pdf.close()

    st.success(f"✅ Certificates generated for {len(df)} students.")
    st.markdown(f"📁 Qualified certificates saved in: `{qualified_folder}`")
    st.markdown(f"📁 Not Qualified certificates saved in: `{not_qualified_folder}`")
