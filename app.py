import os
import sys
import subprocess
from flask import Flask, request, jsonify, send_file, render_template_string
from werkzeug.utils import secure_filename

# Import our converter
import converter

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100 MB max

DOCS_DIR = os.path.join(os.path.dirname(__file__), "documents")
WORD_DIR = os.path.join(os.path.dirname(__file__), "word_documents")
os.makedirs(DOCS_DIR, exist_ok=True)
os.makedirs(WORD_DIR, exist_ok=True)

HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="ta">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>தமிழ் PDF to Word Converter</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Mukta+Malar:wght@400;600;700;800&family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --primary: #1e3a8a;
            --primary-dark: #172554;
            --primary-light: #eff6ff;
            --accent: #2563eb;
            --bg: #f8fafc;
            --card-bg: #ffffff;
            --text-main: #1e293b;
            --text-muted: #64748b;
            --border: #e2e8f0;
            --success: #16a34a;
            --radius: 12px;
            --shadow: 0 4px 6px -1px rgba(0,0,0,0.06), 0 2px 4px -1px rgba(0,0,0,0.04);
            --shadow-lg: 0 10px 15px -3px rgba(0,0,0,0.08), 0 4px 6px -2px rgba(0,0,0,0.03);
        }

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }

        body {
            font-family: 'Plus Jakarta Sans', 'Mukta Malar', system-ui, -apple-system, sans-serif;
            background: var(--bg);
            color: var(--text-main);
            line-height: 1.6;
            padding: 40px 20px;
        }

        .container {
            max-width: 980px;
            margin: 0 auto;
        }

        header {
            text-align: center;
            margin-bottom: 40px;
            padding: 30px;
            background: linear-gradient(135deg, #742a2a 0%, #9b2c2c 60%, #c53030 100%);
            color: white;
            border-radius: var(--radius);
            box-shadow: var(--shadow-lg);
        }

        header h1 {
            font-size: 2.2rem;
            font-weight: 800;
            margin-bottom: 8px;
            letter-spacing: -0.5px;
        }

        header p {
            font-size: 1.05rem;
            opacity: 0.9;
            font-weight: 500;
        }

        .fleuron {
            font-size: 1.4rem;
            color: #feebc8;
            margin-bottom: 12px;
            display: inline-block;
        }

        .section-card {
            background: var(--card-bg);
            border-radius: var(--radius);
            padding: 32px;
            margin-bottom: 30px;
            border: 1px solid var(--border);
            box-shadow: var(--shadow);
        }

        .section-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 24px;
            padding-bottom: 16px;
            border-bottom: 2px solid var(--border);
        }

        .section-header h2 {
            font-size: 1.4rem;
            font-weight: 700;
            color: var(--primary);
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .btn {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            padding: 10px 18px;
            border-radius: 8px;
            font-weight: 600;
            font-size: 0.92rem;
            cursor: pointer;
            border: none;
            transition: all 0.2s ease;
            text-decoration: none;
        }

        .btn-primary {
            background: var(--primary);
            color: white;
        }

        .btn-primary:hover {
            background: var(--primary-dark);
            transform: translateY(-1px);
        }

        .btn-outline {
            background: transparent;
            color: var(--text-main);
            border: 1px solid var(--border);
        }

        .btn-outline:hover {
            background: #edf2f7;
            border-color: #cbd5e0;
        }

        .btn-success {
            background: var(--success);
            color: white;
        }

        .btn-success:hover {
            background: #2f855a;
        }

        /* Books Grid */
        .books-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
            gap: 20px;
        }

        .book-card {
            background: #fffaf0;
            border: 1px solid #feebc8;
            border-left: 5px solid var(--accent);
            border-radius: 10px;
            padding: 20px;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            transition: all 0.2s ease;
        }

        .book-card:hover {
            transform: translateY(-2px);
            box-shadow: var(--shadow);
            border-left-color: var(--primary);
        }

        .book-title {
            font-size: 1.25rem;
            font-weight: 700;
            color: #744210;
            margin-bottom: 6px;
        }

        .book-meta {
            font-size: 0.88rem;
            color: var(--text-muted);
            margin-bottom: 16px;
        }

        .book-badge {
            display: inline-block;
            background: #feebc8;
            color: #9c4221;
            padding: 3px 8px;
            border-radius: 6px;
            font-size: 0.75rem;
            font-weight: 600;
            margin-bottom: 12px;
        }

        /* Uploader Area */
        .upload-dropzone {
            border: 2px dashed #cbd5e0;
            border-radius: 12px;
            padding: 40px 20px;
            text-align: center;
            background: #f8fafc;
            cursor: pointer;
            transition: all 0.2s ease;
            position: relative;
        }

        .upload-dropzone:hover, .upload-dropzone.dragover {
            border-color: var(--primary);
            background: var(--primary-light);
        }

        .upload-icon {
            font-size: 3rem;
            margin-bottom: 12px;
            display: block;
        }

        .file-input {
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            opacity: 0;
            cursor: pointer;
        }

        .form-row {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 16px;
            margin-top: 20px;
        }

        .form-group label {
            display: block;
            font-size: 0.88rem;
            font-weight: 600;
            margin-bottom: 6px;
            color: var(--text-main);
        }

        .form-group input {
            width: 100%;
            padding: 10px 14px;
            border: 1px solid var(--border);
            border-radius: 8px;
            font-size: 0.95rem;
            font-family: inherit;
        }

        .form-group input:focus {
            outline: none;
            border-color: var(--primary);
            box-shadow: 0 0 0 3px rgba(155, 44, 44, 0.15);
        }

        .status-box {
            margin-top: 20px;
            padding: 16px;
            border-radius: 8px;
            display: none;
        }

        .status-box.active {
            display: block;
        }

        .status-box.loading {
            background: #ebf8ff;
            color: #2b6cb0;
            border: 1px solid #bee3f8;
        }

        .status-box.success {
            background: #f0fff4;
            color: #276749;
            border: 1px solid #c6f6d5;
        }

        .status-box.error {
            background: #fff5f5;
            color: #c53030;
            border: 1px solid #fed7d7;
        }

        /* Preview Panel */
        .preview-box {
            margin-top: 20px;
            background: #fffaf0;
            border: 1px solid #feebc8;
            border-radius: 8px;
            padding: 20px;
            max-height: 250px;
            overflow-y: auto;
            white-space: pre-wrap;
            font-size: 0.95rem;
            line-height: 1.7;
            display: none;
        }

        .preview-box.active {
            display: block;
        }

        .preview-title {
            font-weight: 700;
            color: #744210;
            margin-bottom: 10px;
            font-size: 0.9rem;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }

        footer {
            text-align: center;
            margin-top: 40px;
            color: var(--text-muted);
            font-size: 0.88rem;
        }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <div class="fleuron">❖ ❖ ❖</div>
            <h1>தமிழ் PDF to Word Converter</h1>
            <p>Convert legacy Tamil PDFs to editable Word documents (.docx) with Zero Data Loss</p>
        </header>

        <!-- Section: Convert Any New PDF -->
        <div class="section-card">
            <div class="section-header">
                <h2>✨ புதிய PDF ஆவணம் மாற்றுக (Convert Tamil PDF)</h2>
            </div>
            
            <form id="uploadForm">
                <div class="upload-dropzone" id="dropzone">
                    <input type="file" id="pdfFile" class="file-input" accept=".pdf" required onchange="handleFileSelect(event)">
                    <span class="upload-icon">📄</span>
                    <h3 id="dropzoneText">PDF கோப்பை இங்கே இழுத்துப் போடவும் அல்லது தேர்ந்தெடுக்கவும்</h3>
                    <p style="color: var(--text-muted); font-size: 0.88rem; margin-top: 6px;">
                        (Drag & drop your Tamil PDF book here or click to browse)
                    </p>
                </div>

                <div class="form-row">
                    <div class="form-group">
                        <label for="bookTitle">புத்தகத் தலைப்பு (Document Title):</label>
                        <input type="text" id="bookTitle" placeholder="எ.கா: கவிதைத் தொகுதி">
                    </div>
                    <div class="form-group">
                        <label for="authorName">ஆசிரியர் பெயர் (Author Name - விருப்பத்தேர்வு):</label>
                        <input type="text" id="authorName" value="" placeholder="எ.கா: ஆசிரியர் பெயர்">
                    </div>
                </div>

                <div style="margin-top: 24px; text-align: right;">
                    <button type="submit" class="btn btn-primary" id="convertBtn" style="padding: 12px 28px; font-size: 1rem;">
                        ⚡ Word (.docx) ஆக மாற்றுக (Convert)
                    </button>
                </div>
            </form>

            <div class="status-box" id="statusBox"></div>

            <div class="preview-box" id="previewBox">
                <div class="preview-title">முன்னோட்டம் (Tamil Text Preview):</div>
                <div id="previewText"></div>
            </div>
        </div>

        <footer>
            தமிழ் ஆவணக் காப்பகம் • Tamil Digital Preservation
        </footer>
    </div>

    <script>
        // Load existing books on page load
        async function loadBooks() {
            try {
                const res = await fetch('/api/books');
                const data = await res.json();
                const grid = document.getElementById('booksGrid');
                if (data.books.length === 0) {
                    grid.innerHTML = '<p>No Word documents generated yet.</p>';
                    return;
                }
                grid.innerHTML = data.books.map(b => `
                    <div class="book-card">
                        <div>
                            <span class="book-badge">✓ 100% தூய தமிழ் யூனிகோட்</span>
                            <div class="book-title">${b.title}</div>
                            <div class="book-meta">
                                <strong>ஆசிரியர்:</strong> ${b.author} <br>
                                <strong>அளவு:</strong> ${b.size_kb} KB • <strong>நிலை:</strong> தயார் (.docx)
                            </div>
                        </div>
                        <div>
                            <a href="/download/${encodeURIComponent(b.filename)}" class="btn btn-primary" style="width: 100%; justify-content: center;">
                                📥 Word கோப்பு பதிவிறக்குக (.docx)
                            </a>
                        </div>
                    </div>
                `).join('');
            } catch (err) {
                console.error(err);
            }
        }

        function handleFileSelect(e) {
            const file = e.target.files[0];
            if (file) {
                document.getElementById('dropzoneText').innerText = 'தேர்ந்தெடுக்கப்பட்டது: ' + file.name;
                const titleInput = document.getElementById('bookTitle');
                if (!titleInput.value) {
                    titleInput.value = file.name.replace(/\.[^/.]+$/, "");
                }
            }
        }

        async function openFolder() {
            await fetch('/api/open_folder', { method: 'POST' });
        }

        // Handle upload & conversion
        document.getElementById('uploadForm').addEventListener('submit', async (e) => {
            e.preventDefault();
            const fileInput = document.getElementById('pdfFile');
            if (!fileInput.files.length) return;

            const file = fileInput.files[0];
            const title = document.getElementById('bookTitle').value;
            const author = document.getElementById('authorName').value;

            const formData = new FormData();
            formData.append('pdf', file);
            formData.append('title', title);
            formData.append('author', author);

            const statusBox = document.getElementById('statusBox');
            const previewBox = document.getElementById('previewBox');
            const previewText = document.getElementById('previewText');
            const convertBtn = document.getElementById('convertBtn');

            statusBox.className = 'status-box active loading';
            statusBox.innerHTML = '⏳ தமிழ் PDF கோப்பு மாற்றப்பட்டு வருகிறது... தயவுசெய்து காத்திருக்கவும். (Converting PDF to Word, please wait...)';
            previewBox.className = 'preview-box';
            convertBtn.disabled = true;

            try {
                const res = await fetch('/api/convert', {
                    method: 'POST',
                    body: formData
                });
                const result = await res.json();

                if (result.success) {
                    statusBox.className = 'status-box active success';
                    statusBox.innerHTML = `
                        <strong>🎉 வெற்றி! (Conversion Completed Successfully!)</strong><br>
                        கோப்பு உருவாக்கப்பட்டது: <strong>${result.filename}</strong><br><br>
                        <a href="/download/${encodeURIComponent(result.filename)}" class="btn btn-success">
                            📥 புதிய Word கோப்பை பதிவிறக்குக (Download .docx)
                        </a>
                    `;
                    if (result.preview) {
                        previewText.innerText = result.preview;
                        previewBox.className = 'preview-box active';
                    }
                    loadBooks();
                } else {
                    statusBox.className = 'status-box active error';
                    statusBox.innerHTML = '❌ பிழை (Error): ' + (result.error || 'Conversion failed');
                }
            } catch (err) {
                statusBox.className = 'status-box active error';
                statusBox.innerHTML = '❌ பிழை (Network Error): ' + err.message;
            } finally {
                convertBtn.disabled = false;
            }
        });

        // Initialize
        loadBooks();
    </script>
</body>
</html>
"""

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/api/books')
def list_books():
    books = []
    files = [f for f in os.listdir(WORD_DIR) if f.endswith('.docx')]
    for f in sorted(files):
        fpath = os.path.join(WORD_DIR, f)
        size_kb = round(os.path.getsize(fpath) / 1024, 1)
        # Parse title and author from filename
        title = f.replace(".docx", "")
        author = "அனார்"
        if " - " in title:
            parts = title.split(" - ")
            title = parts[0]
            author = parts[1]
        books.append({
            "filename": f,
            "title": title,
            "author": author,
            "size_kb": size_kb
        })
    return jsonify({"books": books})

@app.route('/download/<path:filename>')
def download_file(filename):
    fpath = os.path.join(WORD_DIR, filename)
    if os.path.exists(fpath):
        return send_file(fpath, as_attachment=True)
    return "File not found", 404

@app.route('/api/open_folder', methods=['POST'])
def open_folder():
    try:
        if sys.platform == 'win32':
            os.startfile(WORD_DIR)
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})

@app.route('/api/convert', methods=['POST'])
def handle_convert():
    if 'pdf' not in request.files:
        return jsonify({"success": False, "error": "No PDF file uploaded"}), 400
    
    file = request.files['pdf']
    if file.filename == '':
        return jsonify({"success": False, "error": "No selected file"}), 400

    title = request.form.get('title', '').strip()
    author = request.form.get('author', '').strip()

    filename = secure_filename(file.filename) or "document.pdf"
    if not filename.lower().endswith(".pdf"):
        filename += ".pdf"
        
    save_path = os.path.join(DOCS_DIR, filename)
    file.save(save_path)

    # Output filename
    if not title:
        title = os.path.splitext(filename)[0]
    out_docx_name = f"{title} - {author}.docx" if author else f"{title}.docx"
    out_docx_path = os.path.join(WORD_DIR, out_docx_name)

    try:
        converter.convert_pdf_to_docx(save_path, out_docx_path, book_title=title, author=author)
        
        # Read small sample text from docx for preview
        import docx as dx
        d = dx.Document(out_docx_path)
        sample_paras = [p.text for p in d.paragraphs if p.text.strip() and p.text.strip() != "❖ ❖ ❖"]
        preview = "\n\n".join(sample_paras[1:8]) if len(sample_paras) > 1 else "மாற்றம் வெற்றிகரமாக முடிந்தது."

        return jsonify({
            "success": True,
            "filename": out_docx_name,
            "preview": preview
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

if __name__ == '__main__':
    port = 5000
    print(f"\n" + "=" * 60)
    print(f"  அனார் கவிதைகள் — TAMIL PDF TO WORD CONVERTER WEB APP")
    print(f"  App is running at: http://localhost:{port}")
    print(f"=" * 60 + "\n")
    app.run(host='127.0.0.1', port=port, debug=False)
