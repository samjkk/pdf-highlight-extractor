import streamlit as st
import fitz
from docx import Document
import re
import os
from io import BytesIO

# ---------------- PAGE CONFIG ---------------- #

st.set_page_config(
    page_title="PDF Highlight Extractor",
    page_icon="📑",
    layout="centered",
    initial_sidebar_state="expanded"
)

# ---------------- THEME DETECTION & CSS ---------------- #

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=DM+Sans:ital,opsz,wght@0,9..40,300;0,9..40,400;0,9..40,500;0,9..40,600;1,9..40,300&display=swap');

/* ===== ROOT VARIABLES ===== */
:root {
    --accent: #E8572A;
    --accent-light: #FF7A52;
    --accent-muted: rgba(232, 87, 42, 0.12);
    --accent-border: rgba(232, 87, 42, 0.35);
    --radius-sm: 8px;
    --radius-md: 14px;
    --radius-lg: 20px;
    --radius-xl: 28px;
}

/* ===== GLOBAL FONT ===== */
html, body, [class*="css"], .stApp, .stMarkdown, p, span, div {
    font-family: 'DM Sans', sans-serif !important;
}

/* ===== HIDE STREAMLIT CHROME ===== */
#MainMenu, footer, header { visibility: hidden; }
.stDeployButton { display: none; }

/* ===== APP BACKGROUND ===== */
.stApp {
    background: var(--background-color, #F7F5F2);
}

/* ===== MAIN CONTAINER ===== */
.block-container {
    padding: 2.5rem 2rem 4rem !important;
    max-width: 720px !important;
}

/* ===== HERO HEADER ===== */
.hero-wrap {
    text-align: center;
    padding: 2.5rem 0 1rem;
    margin-bottom: 0.5rem;
}

.hero-badge {
    display: inline-block;
    background: var(--accent-muted);
    border: 1px solid var(--accent-border);
    color: var(--accent);
    font-size: 11px;
    font-weight: 600;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    padding: 5px 14px;
    border-radius: 999px;
    margin-bottom: 1.2rem;
}

.hero-title {
    font-family: 'DM Serif Display', serif !important;
    font-size: clamp(2.2rem, 5vw, 3.2rem);
    font-weight: 400;
    line-height: 1.15;
    margin: 0 0 0.75rem;
    letter-spacing: -0.02em;
}

.hero-title em {
    font-style: italic;
    color: var(--accent);
}

.hero-sub {
    font-size: 15px;
    font-weight: 300;
    opacity: 0.65;
    max-width: 400px;
    margin: 0 auto;
    line-height: 1.65;
}

/* ===== DIVIDER ===== */
.divider {
    height: 1px;
    background: linear-gradient(90deg, transparent, rgba(128,128,128,0.2) 30%, rgba(128,128,128,0.2) 70%, transparent);
    margin: 1.75rem 0;
}

/* ===== UPLOAD ZONE ===== */
.stFileUploader {
    margin-top: 1.5rem;
}

.stFileUploader > div {
    border: 1.5px dashed var(--accent-border) !important;
    border-radius: var(--radius-lg) !important;
    background: var(--accent-muted) !important;
    transition: all 0.2s ease;
}

.stFileUploader > div:hover {
    border-color: var(--accent) !important;
    background: rgba(232, 87, 42, 0.08) !important;
}

.stFileUploader label {
    font-size: 15px !important;
    font-weight: 500 !important;
}

/* Upload icon area */
[data-testid="stFileUploaderDropzone"] {
    padding: 2.5rem !important;
}

[data-testid="stFileUploaderDropzoneInstructions"] > div {
    gap: 0.5rem !important;
}

/* ===== CARDS ===== */
.info-card {
    background: var(--secondary-background-color, #EFEDE8);
    border-radius: var(--radius-md);
    padding: 1.1rem 1.25rem;
    border: 1px solid rgba(128,128,128,0.12);
    margin-bottom: 0.75rem;
}

.stat-row {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 10px;
    margin: 1.25rem 0;
}

.stat-card {
    background: var(--secondary-background-color, #EFEDE8);
    border-radius: var(--radius-md);
    padding: 1rem 1.25rem;
    border: 1px solid rgba(128,128,128,0.1);
    text-align: center;
}

.stat-num {
    font-family: 'DM Serif Display', serif !important;
    font-size: 2.25rem;
    color: var(--accent);
    line-height: 1.1;
    display: block;
    margin-bottom: 4px;
}

.stat-label {
    font-size: 12px;
    font-weight: 500;
    letter-spacing: 0.06em;
    text-transform: uppercase;
    opacity: 0.55;
}

/* ===== SUCCESS BANNER ===== */
.success-banner {
    background: rgba(34, 197, 94, 0.1);
    border: 1px solid rgba(34, 197, 94, 0.3);
    border-radius: var(--radius-md);
    padding: 1rem 1.25rem;
    display: flex;
    align-items: center;
    gap: 10px;
    margin: 1.25rem 0;
}

.success-icon {
    width: 36px;
    height: 36px;
    border-radius: 50%;
    background: rgba(34, 197, 94, 0.2);
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 18px;
    flex-shrink: 0;
}

.success-text-title {
    font-weight: 600;
    font-size: 14px;
    color: #16a34a;
    margin: 0 0 2px;
}

.success-text-sub {
    font-size: 13px;
    opacity: 0.7;
    margin: 0;
}

/* ===== DOWNLOAD BUTTON ===== */
.stDownloadButton {
    margin-top: 1rem;
}

.stDownloadButton button {
    width: 100% !important;
    background: var(--accent) !important;
    color: white !important;
    border: none !important;
    border-radius: var(--radius-md) !important;
    padding: 0.85rem 1.5rem !important;
    font-family: 'DM Sans', sans-serif !important;
    font-size: 15px !important;
    font-weight: 600 !important;
    letter-spacing: 0.02em !important;
    cursor: pointer !important;
    transition: all 0.2s ease !important;
    box-shadow: 0 4px 20px rgba(232, 87, 42, 0.3) !important;
}

.stDownloadButton button:hover {
    background: var(--accent-light) !important;
    box-shadow: 0 6px 28px rgba(232, 87, 42, 0.42) !important;
    transform: translateY(-1px) !important;
}

.stDownloadButton button:active {
    transform: translateY(0) !important;
}

/* ===== SPINNER ===== */
.stSpinner > div {
    border-top-color: var(--accent) !important;
}

/* ===== SIDEBAR ===== */
[data-testid="stSidebar"] {
    border-right: 1px solid rgba(128,128,128,0.12) !important;
}

[data-testid="stSidebar"] .block-container {
    padding: 2rem 1.25rem !important;
}

.sidebar-logo {
    display: flex;
    align-items: center;
    gap: 10px;
    margin-bottom: 1.75rem;
}

.sidebar-logo-icon {
    width: 36px;
    height: 36px;
    background: var(--accent);
    border-radius: 10px;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 18px;
}

.sidebar-logo-text {
    font-family: 'DM Serif Display', serif !important;
    font-size: 16px;
    font-weight: 400;
    line-height: 1.2;
}

.sidebar-logo-text small {
    display: block;
    font-family: 'DM Sans', sans-serif !important;
    font-size: 11px;
    opacity: 0.45;
    font-weight: 400;
}

.sidebar-section-title {
    font-size: 10px;
    font-weight: 600;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    opacity: 0.4;
    margin: 1.5rem 0 0.6rem;
}

.sidebar-step {
    display: flex;
    gap: 10px;
    align-items: flex-start;
    margin-bottom: 0.9rem;
    font-size: 13px;
    line-height: 1.5;
}

.step-num {
    flex-shrink: 0;
    width: 22px;
    height: 22px;
    background: var(--accent-muted);
    border: 1px solid var(--accent-border);
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 11px;
    font-weight: 600;
    color: var(--accent);
}

.sidebar-note {
    font-size: 12px;
    opacity: 0.5;
    line-height: 1.6;
    margin-top: 1.5rem;
    padding-top: 1.25rem;
    border-top: 1px solid rgba(128,128,128,0.12);
}

/* ===== HIDE DEFAULT STREAMLIT INFO BOX ===== */
.stAlert {
    border-radius: var(--radius-md) !important;
}

/* ===== PROGRESS / SPINNER TEXT ===== */
.stSpinner p {
    font-size: 14px !important;
    opacity: 0.6 !important;
}

/* ===== RESPONSIVE ===== */
@media (max-width: 600px) {
    .block-container {
        padding: 1.5rem 1rem 3rem !important;
    }
    .stat-row {
        grid-template-columns: 1fr;
    }
}
</style>
""", unsafe_allow_html=True)

# ---------------- SIDEBAR ---------------- #

with st.sidebar:
    st.markdown("""
    <div class="sidebar-logo">
        <div class="sidebar-logo-icon">📑</div>
        <div class="sidebar-logo-text">
            Highlight Extractor
            <small>PDF → DOCX</small>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="sidebar-section-title">How it works</div>', unsafe_allow_html=True)

    steps = [
        ("Upload", "Add any PDF with highlighted annotations"),
        ("Extract", "We detect all highlight annotations automatically"),
        ("Clean", "Duplicates removed, text normalized"),
        ("Download", "Get a structured DOCX with all your notes"),
    ]

    for i, (title, desc) in enumerate(steps, 1):
        st.markdown(f"""
        <div class="sidebar-step">
            <div class="step-num">{i}</div>
            <div><strong>{title}</strong><br><span style="opacity:0.6">{desc}</span></div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown('<div class="sidebar-section-title">Tips</div>', unsafe_allow_html=True)

    tips = [
        "🟡 All highlight colors are captured",
        "📖 ALL-CAPS highlights become headings",
        "🔄 Duplicate text is removed automatically",
        "📁 Output file matches your PDF name",
    ]
    for tip in tips:
        st.markdown(f'<div class="sidebar-step"><div style="font-size:13px; opacity:0.75">{tip}</div></div>', unsafe_allow_html=True)

    st.markdown("""
    <div class="sidebar-note">
        Works with annotations created in Adobe Acrobat, Preview, Foxit, PDF Expert, and most standard PDF readers.
    </div>
    """, unsafe_allow_html=True)

# ---------------- HERO ---------------- #

st.markdown("""
<div class="hero-wrap">
    <div class="hero-badge">✦ PDF Annotation Tool</div>
    <h1 class="hero-title">Extract your <em>highlights</em>,<br>instantly.</h1>
    <p class="hero-sub">Upload a highlighted PDF and get a clean, structured Word document in seconds.</p>
</div>
<div class="divider"></div>
""", unsafe_allow_html=True)

# ---------------- FILE UPLOAD ---------------- #

uploaded_file = st.file_uploader(
    "Drop your PDF here or click to browse",
    type=["pdf"],
    help="Supports PDF files with highlight annotations. Max 200MB."
)

# ---------------- HELPER FUNCTIONS ---------------- #

def clean_text(text):
    text = text.replace("\n", " ")
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def count_pages(pdf):
    return len(pdf)

# ---------------- PROCESS PDF ---------------- #

if uploaded_file is not None:

    file_size_kb = round(uploaded_file.size / 1024, 1)
    file_size_display = f"{file_size_kb} KB" if file_size_kb < 1024 else f"{round(file_size_kb/1024, 1)} MB"

    st.markdown(f"""
    <div class="info-card">
        <div style="display:flex; justify-content:space-between; align-items:center;">
            <div style="display:flex; align-items:center; gap:10px;">
                <span style="font-size:22px">📄</span>
                <div>
                    <div style="font-weight:600; font-size:14px">{uploaded_file.name}</div>
                    <div style="font-size:12px; opacity:0.5; margin-top:2px">{file_size_display} · PDF Document</div>
                </div>
            </div>
            <div style="font-size:11px; font-weight:600; letter-spacing:0.08em; text-transform:uppercase; opacity:0.4">Ready</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    with st.spinner("Scanning annotations and extracting highlights…"):

        pdf_bytes = uploaded_file.read()
        pdf_path = uploaded_file.name

        with open(pdf_path, "wb") as f:
            f.write(pdf_bytes)

        pdf = fitz.open(pdf_path)
        pdf_name = os.path.splitext(pdf_path)[0]
        output_file = f"{pdf_name}_Highlights.docx"
        total_pages = count_pages(pdf)

        doc = Document()

        # Style the heading
        heading = doc.add_heading(f"{pdf_name} — Extracted Highlights", level=1)

        seen_text = set()
        total_highlights = 0
        pages_with_highlights = 0

        for page in pdf:
            annotations = page.annots()
            page_had_highlight = False

            if annotations:
                for annot in annotations:
                    if annot.type[1] == "Highlight":
                        highlighted_text = page.get_textbox(annot.rect)
                        highlighted_text = clean_text(highlighted_text)

                        if highlighted_text and highlighted_text not in seen_text:
                            seen_text.add(highlighted_text)
                            total_highlights += 1
                            page_had_highlight = True

                            if highlighted_text.isupper():
                                doc.add_heading(highlighted_text, level=2)
                            else:
                                doc.add_paragraph(highlighted_text)

                if page_had_highlight:
                    pages_with_highlights += 1

        doc_io = BytesIO()
        doc.save(doc_io)

    # ---- Stats ---- #
    st.markdown(f"""
    <div class="stat-row">
        <div class="stat-card">
            <span class="stat-num">{total_highlights}</span>
            <span class="stat-label">Highlights found</span>
        </div>
        <div class="stat-card">
            <span class="stat-num">{pages_with_highlights}</span>
            <span class="stat-label">Pages annotated</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    if total_highlights > 0:
        # ---- Success Banner ---- #
        st.markdown(f"""
        <div class="success-banner">
            <div class="success-icon">✓</div>
            <div>
                <p class="success-text-title">Extraction complete</p>
                <p class="success-text-sub">{total_highlights} unique highlights extracted from {pages_with_highlights} page{"s" if pages_with_highlights != 1 else ""}</p>
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.download_button(
            label=f"  Download  {output_file}",
            data=doc_io.getvalue(),
            file_name=output_file,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    else:
        st.warning("⚠️ No highlight annotations found in this PDF. Make sure your highlights are saved as PDF annotations (not just colored text).")

else:
    # ---- Empty state hint ---- #
    st.markdown("""
    <div style="text-align:center; padding: 1.5rem 0; opacity:0.4; font-size:13px;">
        Supports PDFs with highlight annotations from Acrobat, Preview, Foxit & more
    </div>
    """, unsafe_allow_html=True)