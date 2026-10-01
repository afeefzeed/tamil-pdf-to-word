import os
import sys
import tempfile
import streamlit as st
import docx

# Import our high-fidelity converter
import converter

st.set_page_config(
    page_title="தமிழ் PDF to Word Converter",
    page_icon="📄",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Custom Styling
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Mukta+Malar:wght@400;600;700;800&family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', 'Mukta Malar', sans-serif;
    }
    
    .main-header {
        background: linear-gradient(135deg, #1e3a8a 0%, #2563eb 100%);
        color: white;
        padding: 30px;
        border-radius: 12px;
        text-align: center;
        margin-bottom: 24px;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.15);
    }
    .main-header h1 {
        font-size: 2.2rem;
        font-weight: 800;
        margin-bottom: 6px;
        color: #ffffff;
    }
    .main-header p {
        font-size: 1.05rem;
        opacity: 0.92;
        color: #dbeafe;
    }
    .fleuron {
        font-size: 1.5rem;
        margin-bottom: 8px;
        color: #bfdbfe;
    }
    .preview-container {
        background-color: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 18px;
        font-size: 0.95rem;
        line-height: 1.8;
        white-space: pre-wrap;
        max-height: 350px;
        overflow-y: auto;
    }
</style>
""", unsafe_allow_html=True)

# Main Header
st.markdown("""
<div class="main-header">
    <div class="fleuron">❖ ❖ ❖</div>
    <h1>தமிழ் PDF to Word Converter</h1>
    <p>Convert legacy Tamil PDFs to editable Word documents (.docx) with Zero Data Loss</p>
</div>
""", unsafe_allow_html=True)

st.markdown("""
பழைய தமிழ் எழுத்துருக்களில் (TAM, TAB, Elango, Bamini) அச்சுக்கோர்க்கப்பட்ட PDF ஆவணங்களையும் புத்தகங்களையும் எந்தவித எழுத்துப் பிழையோ, வரி அமைப்போ மாறாமல் துல்லியமான யூனிகோட் மைக்ரோசாப்ட் வேர்ட் (.docx) ஆவணமாக மாற்றலாம்.
""")

st.markdown("### 📄 தமிழ் PDF கோப்பைத் தேர்ந்தெடுக்கவும்")
uploaded_file = st.file_uploader(
    "PDF கோப்பை இங்கே பதிவேற்றவும் (Upload your Tamil PDF):",
    type=["pdf"],
    help="எந்தவொரு தமிழ் PDF கோப்பையும் பதிவேற்றலாம்."
)

col1, col2 = st.columns(2)
with col1:
    default_title = ""
    if uploaded_file:
        default_title = os.path.splitext(uploaded_file.name)[0]
    book_title = st.text_input("ஆவணத் தலைப்பு (Document Title):", value=default_title, placeholder="எ.கா: கவிதைத் தொகுதி")
with col2:
    author_name = st.text_input("ஆசிரியர் பெயர் (Author Name):", value="", placeholder="எ.கா: ஆசிரியர் பெயர் (விருப்பத்தேர்வு)")

if uploaded_file is not None:
    if st.button("⚡ Word (.docx) ஆக மாற்றுக (Convert to Word)", type="primary", use_container_width=True):
        with st.spinner("⏳ தமிழ் PDF கோப்பு மாற்றப்பட்டு வருகிறது... தயவுசெய்து சில விநாடிகள் காத்திருக்கவும்..."):
            try:
                # Save uploaded PDF to temporary file
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                    tmp_pdf.write(uploaded_file.getvalue())
                    tmp_pdf_path = tmp_pdf.name
                    
                # Create temporary output docx
                out_docx_path = tmp_pdf_path.replace(".pdf", ".docx")
                
                # Run conversion
                converter.convert_pdf_to_docx(
                    tmp_pdf_path,
                    out_docx_path,
                    book_title=book_title or uploaded_file.name,
                    author=author_name or ""
                )
                
                # Read docx bytes for download
                with open(out_docx_path, "rb") as f:
                    docx_bytes = f.read()
                    
                # Read text preview from generated docx
                d = docx.Document(out_docx_path)
                preview_paras = [p.text for p in d.paragraphs if p.text.strip() and p.text.strip() != "❖ ❖ ❖"]
                preview_text = "\n\n".join(preview_paras[1:25]) if len(preview_paras) > 1 else "மாற்றம் வெற்றிகரமாக முடிந்தது."
                
                st.success("🎉 வெற்றி! தமிழ் PDF வெற்றிகரமாக மைக்ரோசாப்ட் வேர்ட் (.docx) ஆவணமாக மாற்றப்பட்டது!")
                
                # Download button
                download_name = f"{book_title or 'Converted_Document'}.docx"
                st.download_button(
                    label=f"📥 Word கோப்பைப் பதிவிறக்குக: {download_name}",
                    data=docx_bytes,
                    file_name=download_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary",
                    use_container_width=True
                )
                
                # Preview section
                st.markdown("#### 📖 மாற்றப்பட்ட உரையின் முன்னோட்டம் (Converted Text Preview):")
                st.markdown(f'<div class="preview-container">{preview_text}</div>', unsafe_allow_html=True)
                
                # Clean up temp files
                try:
                    os.unlink(tmp_pdf_path)
                    os.unlink(out_docx_path)
                except Exception:
                    pass
                    
            except Exception as e:
                st.error(f"❌ பிழை ஏற்பட்டது (Error during conversion): {str(e)}")

st.divider()
st.caption("தமிழ் ஆவணக் காப்பகம் • Tamil Digital Preservation • 100% Free & Open Source")
