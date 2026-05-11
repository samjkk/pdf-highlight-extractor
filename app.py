import streamlit as st
import fitz
from docx import Document
import re
import os
from io import BytesIO

# ---------------- PAGE CONFIG ---------------- #

st.set_page_config(
    page_title="PDF Highlight Extractor",
    page_icon="📚",
    layout="centered"
)

# ---------------- CUSTOM CSS ---------------- #

st.markdown("""
<style>

.main {
    padding-top: 2rem;
}

h1 {
    text-align: center;
    color: #FF4B4B;
}

.stDownloadButton button {
    width: 100%;
    border-radius: 10px;
    height: 3em;
    font-size: 18px;
}

.upload-text {
    text-align: center;
    color: gray;
    margin-bottom: 20px;
}

</style>
""", unsafe_allow_html=True)

# ---------------- TITLE ---------------- #

st.title("📚 PDF Highlight Extractor")

st.markdown(
    "<p class='upload-text'>Upload a highlighted PDF and extract clean notes instantly.</p>",
    unsafe_allow_html=True
)

# ---------------- SIDEBAR ---------------- #

st.sidebar.header("About")

st.sidebar.info(
    """
    This app extracts highlighted text
    from PDFs and converts it into
    clean notes.
    """
)

# ---------------- FILE UPLOAD ---------------- #

uploaded_file = st.file_uploader(
    "Upload PDF",
    type=["pdf"]
)

# ---------------- CLEAN FUNCTION ---------------- #

def clean_text(text):

    text = text.replace("\n", " ")

    text = re.sub(r'\s+', ' ', text)

    return text.strip()

# ---------------- PROCESS PDF ---------------- #

if uploaded_file is not None:

    with st.spinner("Extracting highlights..."):

        pdf_path = uploaded_file.name

        with open(pdf_path, "wb") as f:
            f.write(uploaded_file.read())

        pdf = fitz.open(pdf_path)

        pdf_name = os.path.splitext(pdf_path)[0]

        output_file = f"{pdf_name}_Highlights.docx"

        doc = Document()

        doc.add_heading(
            f"{pdf_name} - Extracted Highlights",
            level=1
        )

        seen_text = set()

        total_highlights = 0

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

                            total_highlights += 1

                            if highlighted_text.isupper():

                                doc.add_heading(
                                    highlighted_text,
                                    level=2
                                )

                            else:

                                doc.add_paragraph(
                                    highlighted_text
                                )

        # Save to memory
        doc_io = BytesIO()
        doc.save(doc_io)

        st.success(
            f"✅ Extracted {total_highlights} highlights successfully!"
        )

        st.download_button(
            label="📥 Download Highlights DOCX",
            data=doc_io.getvalue(),
            file_name=output_file,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )