# 📖 Tamil PDF to Word (.docx) Converter — Zero Data Loss

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://share.streamlit.io/)
[![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT)

A high-fidelity Tamil PDF to editable Microsoft Word (`.docx`) converter specifically designed to digitize and preserve Tamil books, poetry, literary works, and legacy publications without any loss of Tamil words, vowel signs, conjuncts, stanzas, or alignments.

---

## 🌟 Key Highlights (முக்கிய சிறப்பம்சங்கள்)

- **100% Pure Unicode Tamil (Zero Data Loss):** Converts legacy typesetting encodings (TAM, TAB, Elango, Bamini) into modern Unicode (`UTF-8`). Corrects known flaws in standard conversion tables (like `ங்`, `ளொ`, pulli ligatures).
- **Poetry & Book Layout Preservation:** Accurately preserves stanzas, line breaks, center alignments, and indentations matching the original print layout.
- **Optimized for MS Word:** Employs Windows `Nirmala UI` font with OpenXML Complex Script (`w:cs`) tags, ensuring flawless rendering on Microsoft Word, LibreOffice, and Google Docs across Windows, macOS, Android, and iOS.
- **Dual Spreads & Single Pages:** Seamlessly handles both 2-up landscape printer spreads (left & right book pages) and standard single-page portrait PDFs.
- **Instant Preview & Download:** View the converted Tamil text directly in the browser and download the `.docx` file in one click.

---


## 💻 Local Installation & Usage (உள்ளூர் இயக்கம்)

### 1. Clone the repository
```bash
git clone https://github.com/afeefzeed/tamil-pdf-to-word.git
cd tamil-pdf-to-word
```

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

### 3. Run the Streamlit Web App
```bash
streamlit run streamlit_app.py
```
*(Or double-click `start_streamlit.bat` on Windows)*

The app will open automatically in your browser at `http://localhost:8501`.

---

## 🛠️ CLI Usage (முனைய வழி பயன்பாடு)

You can also convert any Tamil PDF directly from the terminal:

```bash
python converter.py "path/to/tamil_book.pdf" "output_document.docx"
```

---

## 📄 License

MIT License — Feel free to use, share, and adapt for digitizing Tamil literature.
