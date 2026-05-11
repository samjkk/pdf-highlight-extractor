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

# ---------------- FONT (link tag avoids @import glitch) ---------------- #

st.markdown(
    '<link href="https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=DM+Sans:wght@300;400;500;600&display=swap" rel="stylesheet">',
    unsafe_allow_html=True
)

# ---------------- CSS ---------------- #

st.markdown("""
<style>

:root {
    --ac: #E8572A;
    --ac-l: #FF7A52;
    --ac-bg: rgba(232,87,42,0.10);
    --ac-br: rgba(232,87,42,0.30);
    --r-md: 12px;
    --r-lg: 18px;
}

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif !important;
}
#MainMenu, footer, header { visibility: hidden; }
.stDeployButton { display: none; }

.block-container {
    padding: 2rem 1.75rem 4rem !important;
    max-width: 700px !important;
}

/* HERO */
.hero { text-align: center; padding: 2rem 0 0.5rem; }
.badge {
    display: inline-block;
    background: var(--ac-bg);
    border: 1px solid var(--ac-br);
    color: var(--ac);
    font-size: 10px; font-weight: 700;
    letter-spacing: 0.13em; text-transform: uppercase;
    padding: 5px 16px; border-radius: 999px;
    margin-bottom: 1.1rem;
}
.h-title {
    font-family: 'DM Serif Display', Georgia, serif !important;
    font-size: 2.8rem; font-weight: 400;
    line-height: 1.15; margin: 0 0 0.6rem;
    letter-spacing: -0.02em;
}
.h-title em { font-style: italic; color: var(--ac); }
.h-sub {
    font-size: 14px; font-weight: 300; opacity: 0.55;
    margin: 0 auto; max-width: 380px; line-height: 1.7;
}
.div-line {
    height: 1px; background: rgba(128,128,128,0.15);
    margin: 1.5rem 0 1.75rem;
}

/* UPLOAD ZONE — fix duplicated label bug */
[data-testid="stFileUploaderDropzone"] {
    border: 1.5px dashed var(--ac-br) !important;
    border-radius: var(--r-lg) !important;
    background: var(--ac-bg) !important;
    padding: 2.2rem 1.5rem !important;
    transition: border-color 0.2s;
}
[data-testid="stFileUploaderDropzone"]:hover {
    border-color: var(--ac) !important;
}
[data-testid="stFileUploaderDropzone"] button {
    background: var(--ac) !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    font-size: 13px !important;
    font-weight: 600 !important;
    padding: 0.45rem 1.2rem !important;
}
[data-testid="stFileUploaderDropzone"] small {
    font-size: 12px !important; opacity: 0.45 !important;
}

/* CARDS */
.icard {
    border-radius: var(--r-md);
    border: 1px solid rgba(128,128,128,0.13);
    padding: 1rem 1.2rem; margin: 1.1rem 0;
    display: flex; align-items: center; gap: 12px;
}
.icard-icon { font-size: 22px; flex-shrink: 0; }
.icard-name { font-weight: 600; font-size: 14px; margin: 0; }
.icard-meta { font-size: 12px; opacity: 0.45; margin: 3px 0 0; }
.icard-status {
    margin-left: auto; font-size: 10px; font-weight: 700;
    letter-spacing: 0.1em; text-transform: uppercase; opacity: 0.35;
}

/* STAT GRID */
.stat-grid {
    display: grid; grid-template-columns: 1fr 1fr;
    gap: 10px; margin: 1.2rem 0;
}
.stat-card {
    border-radius: var(--r-md);
    border: 1px solid rgba(128,128,128,0.11);
    padding: 1rem 1.2rem; text-align: center;
}
.stat-n {
    font-family: 'DM Serif Display', Georgia, serif !important;
    font-size: 2.4rem; color: var(--ac);
    line-height: 1; display: block; margin-bottom: 6px;
}
.stat-l {
    font-size: 11px; font-weight: 600;
    letter-spacing: 0.08em; text-transform: uppercase; opacity: 0.4;
}

/* SUCCESS BANNER */
.ok-banner {
    border-radius: var(--r-md);
    border: 1px solid rgba(34,197,94,0.28);
    background: rgba(34,197,94,0.08);
    padding: 0.9rem 1.2rem;
    display: flex; align-items: center; gap: 12px;
    margin: 0.75rem 0 1rem;
}
.ok-dot {
    width: 32px; height: 32px; border-radius: 50%;
    background: rgba(34,197,94,0.18);
    display: flex; align-items: center; justify-content: center;
    font-size: 16px; flex-shrink: 0;
}
.ok-title { font-weight: 600; font-size: 14px; color: #16a34a; margin: 0 0 2px; }
.ok-sub { font-size: 12px; opacity: 0.6; margin: 0; }

/* DOWNLOAD BUTTON */
.stDownloadButton button {
    width: 100% !important;
    background: var(--ac) !important;
    color: white !important;
    border: none !important;
    border-radius: var(--r-md) !important;
    padding: 0.8rem 1.5rem !important;
    font-size: 15px !important;
    font-weight: 600 !important;
    box-shadow: 0 3px 16px rgba(232,87,42,0.28) !important;
    transition: background 0.18s, transform 0.12s !important;
}
.stDownloadButton button:hover {
    background: var(--ac-l) !important;
    transform: translateY(-1px) !important;
}

/* SIDEBAR */
[data-testid="stSidebar"] { border-right: 1px solid rgba(128,128,128,0.1) !important; }
[data-testid="stSidebar"] .block-container { padding: 1.75rem 1.1rem !important; }
.s-logo { display: flex; align-items: center; gap: 10px; margin-bottom: 1.75rem; }
.s-icon {
    width: 34px; height: 34px; background: var(--ac);
    border-radius: 9px; display: flex; align-items: center;
    justify-content: center; font-size: 17px;
}
.s-name { font-family: 'DM Serif Display', Georgia, serif !important; font-size: 15px; line-height: 1.25; }
.s-name small { font-family: 'DM Sans', sans-serif !important; font-size: 11px; opacity: 0.4; font-weight: 400; display: block; }
.s-sec { font-size: 10px; font-weight: 700; letter-spacing: 0.1em; text-transform: uppercase; opacity: 0.35; margin: 1.4rem 0 0.6rem; }
.s-step { display: flex; gap: 10px; align-items: flex-start; margin-bottom: 0.85rem; font-size: 13px; line-height: 1.5; }
.s-num {
    flex-shrink: 0; width: 21px; height: 21px;
    background: var(--ac-bg); border: 1px solid var(--ac-br);
    border-radius: 50%; display: flex; align-items: center;
    justify-content: center; font-size: 10px; font-weight: 700;
    color: var(--ac); margin-top: 1px;
}
.s-foot { font-size: 11.5px; opacity: 0.4; line-height: 1.65; margin-top: 1.5rem; padding-top: 1.2rem; border-top: 1px solid rgba(128,128,128,0.12); }
.empty { text-align: center; padding: 1rem 0 0.5rem; font-size: 13px; opacity: 0.35; }

</style>
""", unsafe_allow_html=True)

# ---------------- SIDEBAR ---------------- #

with st.sidebar:
    st.markdown("""
    <div class="s-logo">
        <div class="s-icon">📑</div>
        <div class="s-name">Highlight Extractor<small>PDF → DOCX</small></div>
    </div>
    <div class="s-sec">How it works</div>
    <div class="s-step"><div class="s-num">1</div><div><strong>Upload</strong><br><span style="opacity:.6">Add any annotated PDF</span></div></div>
    <div class="s-step"><div class="s-num">2</div><div><strong>Extract</strong><br><span style="opacity:.6">Highlights detected automatically</span></div></div>
    <div class="s-step"><div class="s-num">3</div><div><strong>Clean</strong><br><span style="opacity:.6">Duplicates removed, text normalized</span></div></div>
    <div class="s-step"><div class="s-num">4</div><div><strong>Download</strong><br><span style="opacity:.6">Structured DOCX ready to use</span></div></div>
    <div class="s-sec">Tips</div>
    <div class="s-step"><div style="font-size:13px;opacity:.7">🟡 All highlight colors captured</div></div>
    <div class="s-step"><div style="font-size:13px;opacity:.7">📖 ALL-CAPS highlights become headings</div></div>
    <div class="s-step"><div style="font-size:13px;opacity:.7">🔄 Duplicates removed automatically</div></div>
    <div class="s-step"><div style="font-size:13px;opacity:.7">📁 Output filename matches your PDF</div></div>
    <div class="s-foot">Works with Acrobat, Preview, Foxit, PDF Expert, and most standard PDF readers.</div>
    """, unsafe_allow_html=True)

# ---------------- HERO ---------------- #

st.markdown("""
<div class="hero">
    <div class="badge">✦ PDF Annotation Tool</div>
    <div class="h-title">Extract your <em>highlights</em>,<br>instantly.</div>
    <div class="h-sub">Upload a highlighted PDF and get a clean,<br>structured Word document in seconds.</div>
</div>
<div class="div-line"></div>
""", unsafe_allow_html=True)

# ---------------- UPLOAD LABEL (manual, above the widget) ---------------- #

st.markdown(
    '<p style="font-size:11px;font-weight:700;letter-spacing:0.1em;'
    'text-transform:uppercase;opacity:0.4;margin-bottom:4px">Upload PDF</p>',
    unsafe_allow_html=True
)

# label_visibility="collapsed" removes the built-in label so it doesn't double-render
uploaded_file = st.file_uploader(
    label="upload",
    type=["pdf"],
    label_visibility="collapsed",
    help="PDF files with highlight annotations. Max 200MB."
)

# ---------------- HELPERS ---------------- #

def clean_text(text):
    text = text.replace("\n", " ")
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

# ---------------- PROCESS ---------------- #

if uploaded_file is not None:

    size_kb = round(uploaded_file.size / 1024, 1)
    size_str = f"{size_kb} KB" if size_kb < 1024 else f"{round(size_kb/1024,1)} MB"

    st.markdown(f"""
    <div class="icard">
        <div class="icard-icon">📄</div>
        <div>
            <div class="icard-name">{uploaded_file.name}</div>
            <div class="icard-meta">{size_str} · PDF Document</div>
        </div>
        <div class="icard-status">Ready</div>
    </div>
    """, unsafe_allow_html=True)

    with st.spinner("Scanning annotations…"):
        pdf_bytes = uploaded_file.read()
        pdf_path = uploaded_file.name
        with open(pdf_path, "wb") as f:
            f.write(pdf_bytes)

        pdf = fitz.open(pdf_path)
        pdf_name = os.path.splitext(pdf_path)[0]
        output_file = f"{pdf_name}_Highlights.docx"

        doc = Document()
        doc.add_heading(f"{pdf_name} — Extracted Highlights", level=1)

        seen_text = set()
        total_highlights = 0
        pages_with_highlights = 0

        for page in pdf:
            page_had = False
            annotations = page.annots()
            if annotations:
                for annot in annotations:
                    if annot.type[1] == "Highlight":
                        txt = clean_text(page.get_textbox(annot.rect))
                        if txt and txt not in seen_text:
                            seen_text.add(txt)
                            total_highlights += 1
                            page_had = True
                            if txt.isupper():
                                doc.add_heading(txt, level=2)
                            else:
                                doc.add_paragraph(txt)
                if page_had:
                    pages_with_highlights += 1

        doc_io = BytesIO()
        doc.save(doc_io)

    st.markdown(f"""
    <div class="stat-grid">
        <div class="stat-card">
            <span class="stat-n">{total_highlights}</span>
            <span class="stat-l">Highlights found</span>
        </div>
        <div class="stat-card">
            <span class="stat-n">{pages_with_highlights}</span>
            <span class="stat-l">Pages annotated</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    if total_highlights > 0:
        pg_word = "page" if pages_with_highlights == 1 else "pages"
        st.markdown(f"""
        <div class="ok-banner">
            <div class="ok-dot">✓</div>
            <div>
                <div class="ok-title">Extraction complete</div>
                <div class="ok-sub">{total_highlights} unique highlights from {pages_with_highlights} {pg_word}</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.download_button(
            label=f"⬇  Download  {output_file}",
            data=doc_io.getvalue(),
            file_name=output_file,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    else:
        st.warning("⚠️ No highlight annotations found. Make sure highlights are saved as PDF annotations, not just colored text.")

else:
    st.markdown(
        '<div class="empty">Drag a PDF above, or click to browse · Supports Acrobat, Preview, Foxit & more</div>',
        unsafe_allow_html=True
    )