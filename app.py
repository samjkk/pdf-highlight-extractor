import streamlit as st
import fitz
from docx import Document
import re
import os

st.title("PDF Highlight Extractor")

# Upload PDF
uploaded_file = st.file_uploader(
    "Upload a highlighted PDF",
    type=["pdf"]
)

# Function to clean text
def clean_text(text):

    text = text.replace("\n", " ")
    text = re.sub(r'\s+', ' ', text)

    return text.strip()

# When file uploaded
if uploaded_file is not None:

    # Save uploaded PDF temporarily
    pdf_path = uploaded_file.name

    with open(pdf_path, "wb") as f:
        f.write(uploaded_file.read())

    # Open PDF
    pdf = fitz.open(pdf_path)

    # Output filename
    pdf_name = os.path.splitext(pdf_path)[0]
    output_file = f"{pdf_name}_Highlights.docx"

    # Create DOCX
    doc = Document()

    doc.add_heading(
        f"{pdf_name} - Extracted Highlights",
        level=1
    )

    seen_text = set()

    # Extract highlights
    for page in pdf:

        annotations = page.annots()

        if annotations:

            for annot in annotations:

                if annot.type[1] == "Highlight":

                    highlighted_text = page.get_textbox(
                        annot.rect
                    )

                    highlighted_text = clean_text(
                        highlighted_text
                    )

                    if (
                        highlighted_text
                        and highlighted_text not in seen_text
                    ):

                        seen_text.add(highlighted_text)

                        # Heading detection
                        if highlighted_text.isupper():

                            doc.add_heading(
                                highlighted_text,
                                level=2
                            )

                        else:

                            doc.add_paragraph(
                                highlighted_text
                            )

    # Save document
    doc.save(output_file)

    st.success("Highlights extracted!")

    # Download button
    with open(output_file, "rb") as file:

        st.download_button(
            label="Download Highlights DOCX",
            data=file,
            file_name=output_file,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )