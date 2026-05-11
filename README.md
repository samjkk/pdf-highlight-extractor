# 📚 Highlight Extractor

A **multi-format annotation extraction tool** built with Streamlit that pulls highlights, comments, and marked content from documents and exports them into clean, organized notes.

Upload a PDF, DOCX, PPTX, EPUB, or TXT file — get your highlights back in **DOCX, PDF, TXT, or Markdown**.

---

## ✨ Features

| Capability | Details |
|---|---|
| **Multi-format input** | PDF · DOCX · PPTX · EPUB · TXT |
| **Multi-format output** | DOCX · PDF · TXT · Markdown |
| **Batch processing** | Upload and extract from multiple files at once |
| **Duplicate removal** | Automatically deduplicates identical highlights |
| **Smart headings** | ALL-CAPS highlights are promoted to section headings |
| **Styled PDF export** | ReportLab-powered output with custom typography and accent colors |
| **Responsive UI** | Mobile-friendly layout with sidebar documentation |

---

## 📖 Supported Formats

### Inputs

| Format | What gets extracted |
|---|---|
| **PDF** | Highlight annotations (quad-point intersection matching) |
| **DOCX** | Highlighted text runs + Word comments |
| **PPTX** | Highlighted runs and bold long-text heuristics |
| **EPUB** | `<mark>` tags and highlight-styled `<span>` elements |
| **TXT** | All non-empty lines (full text extraction) |

### Outputs

| Format | Description |
|---|---|
| **DOCX** | Structured Word document with headings per source file |
| **PDF** | Professionally styled document (accent colors, horizontal rules) |
| **TXT** | Plain text — one highlight per line |
| **MD** | Markdown with `## Headings` and `- bullet` items |

---

## 🚀 Getting Started

### Prerequisites

- **Python 3.9+**

### Installation

```bash
# Clone the repository
git clone https://github.com/samjkk/pdf-highlight-extractor.git
cd pdf-highlight-extractor

# Install dependencies
pip install -r requirements.txt
```

### Run the App

```bash
streamlit run app.py
```

The app will open at **http://localhost:8501**.

> **Note:** The repo also includes `main.py`, a minimal CLI script for quick single-PDF extraction to DOCX — useful for scripting or batch workflows outside the browser.

---

## 🖥️ Usage

1. **Upload** — Drag & drop one or more supported files into the upload zone.
2. **Pick format** — Select your preferred output format (DOCX, PDF, TXT, or MD).
3. **Extract** — Click **⚡ Extract highlights** to process all files.
4. **Download** — Grab your combined highlights file.

---

## 🏗️ Project Structure

```
pdf-highlight-extractor/
├── app.py                # Streamlit web app (main application)
├── main.py               # CLI script for single-PDF → DOCX extraction
├── requirements.txt      # Python dependencies
└── .devcontainer/        # GitHub Codespaces / VS Code dev container config
    └── devcontainer.json
```

---

## 🔧 Tech Stack

| Library | Purpose |
|---|---|
| [Streamlit](https://streamlit.io/) | Web UI framework |
| [PyMuPDF (fitz)](https://pymupdf.readthedocs.io/) | PDF annotation extraction |
| [python-docx](https://python-docx.readthedocs.io/) | DOCX read/write |
| [python-pptx](https://python-pptx.readthedocs.io/) | PowerPoint highlight extraction |
| [ebooklib](https://github.com/aerkalov/ebooklib) | EPUB parsing |
| [BeautifulSoup4](https://www.crummy.com/software/BeautifulSoup/) | HTML parsing for EPUB content |
| [ReportLab](https://www.reportlab.com/) | PDF output generation |
| [lxml](https://lxml.de/) | XML parsing for DOCX comments |

---

## ☁️ Deploy on Streamlit Cloud

1. Fork or push this repo to your GitHub account.
2. Go to [share.streamlit.io](https://share.streamlit.io/) → **New app**.
3. Point it at `app.py` — done.

---

## 📝 License

This project is open source and available under the [MIT License](LICENSE).
