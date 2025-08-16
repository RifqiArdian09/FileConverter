from flask import Flask, request, send_file, render_template_string, jsonify
import tempfile
import os
import io
import subprocess
import mimetypes
import time
from pathlib import Path
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB max file size

# ==============================
# HTML (UI) — Modern Dark Theme
# ==============================
HTML_PAGE = """
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Advanced File Converter</title>
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" rel="stylesheet">
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #111 0%, #333 100%);
            min-height: 100vh; display: flex; align-items: center; justify-content: center;
            padding: 20px; color: #eee;
        }
        .container { background: #1e1e1e; border-radius: 20px; box-shadow: 0 20px 40px rgba(0,0,0,0.6);
            padding: 40px; width: 100%; max-width: 800px; animation: slideUp 0.5s ease-out; }
        @keyframes slideUp { from { opacity: 0; transform: translateY(30px);} to { opacity:1; transform: translateY(0);} }
        .header { text-align: center; margin-bottom: 40px; }
        .header h1 { font-size: 2.5rem; margin-bottom: 10px; color: #fff; }
        .header p { color: #bbb; font-size: 1.1rem; }
        .section-title { font-size: 1.5rem; margin-bottom: 20px; color: #fff; }
        .section-title i { margin-right: 10px; color: #aaa; }
        .file-upload-section { margin-bottom: 40px; }
        .file-upload { border: 3px dashed #666; border-radius: 15px; padding: 40px 20px; text-align: center;
            transition: all 0.3s ease; cursor: pointer; position: relative; overflow: hidden; color: #ccc; }
        .file-upload:hover { border-color: #999; background: rgba(255,255,255,0.05); }
        .file-upload.dragover { border-color: #fff; background: rgba(255,255,255,0.1); }
        .file-upload input[type="file"] { position:absolute; left:0; top:0; width:100%; height:100%; opacity:0; cursor:pointer; }
        .file-upload-icon { font-size: 3rem; color: #888; margin-bottom: 15px; }
        .file-info { margin-top: 15px; padding: 15px; background: #2a2a2a; border-radius: 10px; display: none; color: #ddd; }
        .file-info.show { display: block; }
        .format-selection { margin-bottom: 40px; }
        .format-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 15px; margin-bottom: 30px; }
        .format-card { background:#2a2a2a; border:2px solid #444; border-radius:15px; padding:20px; text-align:center; cursor:pointer; transition:all .3s ease; }
        .format-card:hover { border-color:#888; background:#333; }
        .format-card.selected { border-color:#fff; background:#444; color:#fff; }
        .format-icon { font-size:2.5rem; margin-bottom:10px; color:#aaa; }
        .format-card.selected .format-icon { color:#fff; }
        .format-name { font-weight:600; margin-bottom:5px; }
        .format-desc { font-size:.9rem; color:#bbb; }
        .convert-section { text-align:center; margin-bottom:30px; }
        .convert-btn { background:#444; color:#fff; border:none; padding:15px 40px; border-radius:50px; font-size:1.1rem; font-weight:600; cursor:pointer; transition:all .3s ease; }
        .convert-btn:hover:not(:disabled) { background:#666; transform: translateY(-2px); }
        .convert-btn:disabled { opacity:.4; cursor:not-allowed; }
        .status { display:none; padding:15px; border-radius:10px; margin-top:20px; text-align:center; }
        .status.show { display:block; }
        .status.success { background:#2f4f2f; color:#d4fcd4; border:1px solid #3a6; }
        .status.error { background:#4f2f2f; color:#fdd; border:1px solid #a66; }
        .status.loading { background:#2f2f4f; color:#ccd; border:1px solid #669; }
        .progress-bar { width:100%; height:6px; background:#444; border-radius:3px; margin-top:10px; overflow:hidden; }
        .progress-fill { height:100%; background: linear-gradient(90deg, #aaa, #fff); border-radius:3px; width:0%; transition: width .3s ease; }
        @media (max-width:768px){ .container{padding:20px;} .header h1{font-size:2rem;} .format-grid{grid-template-columns: repeat(auto-fit,minmax(120px,1fr));}}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1><i class="fas fa-magic"></i> Advanced File Converter</h1>
            <p>Convert your files to any format with just a few clicks</p>
        </div>

        <form id="convertForm">
            <div class="file-upload-section">
                <h3 class="section-title"><i class="fas fa-upload"></i> Upload File</h3>
                <div class="file-upload" id="fileUpload">
                    <div class="file-upload-icon"><i class="fas fa-cloud-upload-alt"></i></div>
                    <h4>Drag & Drop your file here</h4>
                    <p>atau klik untuk memilih file</p>
                    <input type="file" id="file" name="file" required>
                </div>
                <div class="file-info" id="fileInfo">
                    <i class="fas fa-file"></i>
                    <span id="fileName"></span> - <span id="fileSize"></span>
                </div>
            </div>

            <div class="format-selection">
                <h3 class="section-title"><i class="fas fa-exchange-alt"></i> Choose Output Format</h3>
                <div class="format-grid">
                    <div class="format-card" data-format="pdf"><div class="format-icon"><i class="fas fa-file-pdf"></i></div><div class="format-name">PDF</div><div class="format-desc">Portable Document</div></div>
                    <div class="format-card" data-format="docx"><div class="format-icon"><i class="fas fa-file-word"></i></div><div class="format-name">DOCX</div><div class="format-desc">Word Document</div></div>
                    <div class="format-card" data-format="pptx"><div class="format-icon"><i class="fas fa-file-powerpoint"></i></div><div class="format-name">PPTX</div><div class="format-desc">PowerPoint</div></div>
                    <div class="format-card" data-format="xlsx"><div class="format-icon"><i class="fas fa-file-excel"></i></div><div class="format-name">XLSX</div><div class="format-desc">Excel Spreadsheet</div></div>
                    <div class="format-card" data-format="txt"><div class="format-icon"><i class="fas fa-file-alt"></i></div><div class="format-name">TXT</div><div class="format-desc">Plain Text</div></div>
                    <div class="format-card" data-format="rtf"><div class="format-icon"><i class="fas fa-file-text"></i></div><div class="format-name">RTF</div><div class="format-desc">Rich Text Format</div></div>
                    <div class="format-card" data-format="odt"><div class="format-icon"><i class="fas fa-file-text"></i></div><div class="format-name">ODT</div><div class="format-desc">OpenDocument Text</div></div>
                    <div class="format-card" data-format="html"><div class="format-icon"><i class="fas fa-code"></i></div><div class="format-name">HTML</div><div class="format-desc">Web Page</div></div>
                    <div class="format-card" data-format="epub"><div class="format-icon"><i class="fas fa-book"></i></div><div class="format-name">EPUB</div><div class="format-desc">E-Book</div></div>
                    <div class="format-card" data-format="csv"><div class="format-icon"><i class="fas fa-table"></i></div><div class="format-name">CSV</div><div class="format-desc">Comma Separated</div></div>
                </div>
            </div>

            <div class="convert-section">
                <button type="submit" class="convert-btn" id="convertBtn" disabled>
                    <i class="fas fa-magic"></i> Convert File
                </button>
            </div>

            <div class="status" id="status">
                <div id="statusText"></div>
                <div class="progress-bar" id="progressBar" style="display:none;">
                    <div class="progress-fill" id="progressFill"></div>
                </div>
            </div>
        </form>
    </div>

    <script>
        let selectedFormat = '';
        let selectedFile = null;

        const fileUpload = document.getElementById('fileUpload');
        const fileInput = document.getElementById('file');
        const fileInfo = document.getElementById('fileInfo');
        const fileName = document.getElementById('fileName');
        const fileSize = document.getElementById('fileSize');

        fileUpload.addEventListener('click', (e) => { if (e.target !== fileInput) fileInput.click(); });
        fileUpload.addEventListener('dragover', (e) => { e.preventDefault(); fileUpload.classList.add('dragover'); });
        fileUpload.addEventListener('dragleave', (e) => { e.preventDefault(); fileUpload.classList.remove('dragover'); });
        fileUpload.addEventListener('drop', (e) => {
            e.preventDefault(); fileUpload.classList.remove('dragover');
            const files = e.dataTransfer.files;
            if (files.length > 0) { fileInput.files = files; handleFileSelect(files[0]); }
        });
        fileInput.addEventListener('change', (e) => { if (e.target.files.length > 0) handleFileSelect(e.target.files[0]); });

        function handleFileSelect(file) {
            selectedFile = file;
            fileName.textContent = file.name;
            fileSize.textContent = formatFileSize(file.size);
            fileInfo.classList.add('show');
            updateConvertButton();
        }
        function formatFileSize(bytes) {
            if (bytes === 0) return '0 Bytes';
            const k = 1024, sizes = ['Bytes','KB','MB','GB'];
            const i = Math.floor(Math.log(bytes) / Math.log(k));
            return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
        }
        const formatCards = document.querySelectorAll('.format-card');
        formatCards.forEach(card => {
            card.addEventListener('click', () => {
                formatCards.forEach(c => c.classList.remove('selected'));
                card.classList.add('selected');
                selectedFormat = card.dataset.format;
                updateConvertButton();
            });
        });
        function updateConvertButton() {
            const convertBtn = document.getElementById('convertBtn');
            convertBtn.disabled = !selectedFile || !selectedFormat;
        }
        document.getElementById('convertForm').addEventListener('submit', async (e) => {
            e.preventDefault(); await convertFile();
        });
        async function convertFile() {
            const status = document.getElementById('status');
            const statusText = document.getElementById('statusText');
            const progressBar = document.getElementById('progressBar');
            const progressFill = document.getElementById('progressFill');
            const convertBtn = document.getElementById('convertBtn');

            status.className = 'status show loading';
            statusText.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Converting file, please wait...';
            progressBar.style.display = 'block'; convertBtn.disabled = true;

            let progress = 0;
            const progressInterval = setInterval(() => {
                progress += Math.random() * 10; if (progress > 90) progress = 90;
                progressFill.style.width = progress + '%';
            }, 200);

            try {
                const formData = new FormData();
                formData.append('file', selectedFile);
                formData.append('to', selectedFormat);

                const response = await fetch('/convert', { method: 'POST', body: formData });

                clearInterval(progressInterval);
                progressFill.style.width = '100%';

                if (!response.ok) {
                    const error = await response.json().catch(() => ({}));
                    throw new Error(error.error || 'Conversion failed');
                }

                const cd = response.headers.get('content-disposition');
                let filename = `converted.${selectedFormat}`;
                if (cd) { const m = /filename="([^"]*)"/.exec(cd); if (m && m[1]) filename = m[1]; }

                const blob = await response.blob();
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a'); a.href = url; a.download = filename; document.body.appendChild(a); a.click();
                a.remove(); URL.revokeObjectURL(url);

                status.className = 'status show success';
                statusText.innerHTML = '<i class="fas fa-check-circle"></i> File converted and downloaded successfully!';
                setTimeout(() => { progressBar.style.display = 'none'; progressFill.style.width = '0%'; }, 2000);
            } catch (error) {
                clearInterval(progressInterval);
                status.className = 'status show error';
                statusText.innerHTML = `<i class="fas fa-exclamation-triangle"></i> Error: ${error.message}`;
                progressBar.style.display = 'none'; progressFill.style.width = '0%';
            }
            convertBtn.disabled = false;
        }
    </script>
</body>
</html>
"""

# =========================
# Helper & Convert Routines
# =========================
def detect_soffice_path() -> str:
    """Deteksi path LibreOffice di Windows, fallback ke 'soffice' (ada di PATH)."""
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "soffice"  # jika sudah ada di PATH
    ]
    for c in candidates:
        if c == "soffice":
            return c  # biarkan subprocess yang resolve via PATH
        if os.path.exists(c):
            return c
    return "soffice"

def convert_with_libreoffice(input_path: str, output_dir: str, to_format: str) -> bool:
    try:
        # Mapping format → filter LibreOffice
        format_mapping = {
            'pdf':  'pdf:writer_pdf_Export',
            'docx': 'docx',
            'odt':  'odt',
            'txt':  'txt:Text',
            'rtf':  'rtf:Rich Text Format',
            'html': 'html:XHTML Writer File',
            'epub': 'epub',
            'csv':  'csv:Text - txt - csv (StarCalc)',
            'xlsx': 'xlsx',
            'pptx': 'pptx'
        }
        libreoffice_format = format_mapping.get(to_format.lower(), to_format.lower())
        soffice = detect_soffice_path()

        cmd = [
            soffice, "--headless", "--invisible", "--nodefault",
            "--nolockcheck", "--nologo", "--norestore",
            "--convert-to", libreoffice_format,
            "--outdir", output_dir, input_path
        ]
        result = subprocess.run(
            cmd, check=True, capture_output=True, text=True, timeout=120
        )
        print(f"[LO] stdout: {result.stdout}")
        print(f"[LO] stderr: {result.stderr}")
        # beri sedikit jeda agar file lepas lock (Windows)
        time.sleep(0.5)
        return True
    except Exception as e:
        print(f"[LO] ERROR: {e}")
        return False

def get_file_extension(filename: str) -> str:
    return filename.rsplit('.', 1)[1].lower() if '.' in filename else ''

def is_supported_format(filename: str, target_format: str) -> bool:
    input_ext = get_file_extension(filename)
    document_formats = ['doc', 'docx', 'odt', 'rtf', 'txt', 'html', 'pdf']
    presentation_formats = ['ppt', 'pptx', 'odp']
    spreadsheet_formats = ['xls', 'xlsx', 'ods', 'csv']
    all_supported = document_formats + presentation_formats + spreadsheet_formats

    if input_ext not in all_supported:
        return False

    # Batasi konversi PDF → editable (tanpa OCR kualitas kurang)
    if input_ext == 'pdf' and target_format.lower() in ['docx', 'odt', 'txt']:
        return False

    return True

def get_mimetype(format_ext: str) -> str:
    mime_mapping = {
        'pdf': 'application/pdf',
        'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'txt': 'text/plain',
        'rtf': 'application/rtf',
        'odt': 'application/vnd.oasis.opendocument.text',
        'html': 'text/html',
        'epub': 'application/epub+zip',
        'csv': 'text/csv'
    }
    mime_type = mime_mapping.get(format_ext.lower())
    if not mime_type:
        mime_type, _ = mimetypes.guess_type(f"file.{format_ext}")
        if not mime_type:
            mime_type = 'application/octet-stream'
    return mime_type

def safe_open(path, mode="rb", retries=6, delay=0.5):
    """Open dengan retry untuk menghindari PermissionError (WinError 32)."""
    for i in range(retries):
        try:
            return open(path, mode)
        except PermissionError as e:
            if i < retries - 1:
                time.sleep(delay)
            else:
                raise e

def find_converted_file(tmpdir: str, target_ext: str, original_filename: str) -> str | None:
    """
    Cari file output:
    - Prioritaskan file dengan ekstensi target terbaru (mtime terbesar).
    - Hindari nama sama persis dengan input.
    """
    target_ext = target_ext.lower().lstrip('.')
    original_name = Path(original_filename).name

    candidates = []
    for f in os.listdir(tmpdir):
        p = Path(tmpdir) / f
        if p.is_file() and f.lower().endswith(f".{target_ext}") and f != original_name:
            candidates.append(p)

    if not candidates:
        return None

    candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return str(candidates[0])

# =========
# Endpoints
# =========
@app.route("/")
def index():
    return render_template_string(HTML_PAGE)

@app.route("/convert", methods=["POST"])
def convert_endpoint():
    try:
        if "file" not in request.files:
            return jsonify({"error": "No file part"}), 400

        file = request.files["file"]
        if file.filename == "":
            return jsonify({"error": "No selected file"}), 400

        to_format = (request.form.get("to") or "").lower().strip()
        if not to_format:
            return jsonify({"error": "No target format specified"}), 400

        if not is_supported_format(file.filename, to_format):
            return jsonify({
                "error": "Unsupported conversion for this input/output format."
            }), 400

        with tempfile.TemporaryDirectory() as tmpdir:
            input_path = os.path.join(tmpdir, secure_filename(file.filename))
            file.save(input_path)

            success = convert_with_libreoffice(input_path, tmpdir, to_format)
            if not success:
                return jsonify({
                    "error": "Conversion failed. Please ensure LibreOffice is installed and the file format is supported."
                }), 500

            # Cari file hasil konversi
            converted_file = find_converted_file(tmpdir, to_format, os.path.basename(input_path))
            if not converted_file or not os.path.exists(converted_file):
                # Debug isi folder temp
                print("Temp dir contents:", os.listdir(tmpdir))
                return jsonify({"error": "Converted file not found"}), 500

            # Baca file dengan retry
            try:
                with safe_open(converted_file, "rb") as f:
                    data = io.BytesIO(f.read())
                data.seek(0)
                print(f"Successfully read converted file: {converted_file}")
            except Exception as e:
                return jsonify({"error": f"Failed to read converted file: {str(e)}"}), 500

            output_filename = f"{os.path.splitext(file.filename)[0]}.{to_format}"
            return send_file(
                data,
                as_attachment=True,
                download_name=output_filename,
                mimetype=get_mimetype(to_format)
            )

    except Exception as e:
        # log stacktrace ke console
        import traceback
        traceback.print_exc()
        return jsonify({"error": f"Internal server error: {str(e)}"}), 500

@app.errorhandler(413)
def too_large(e):
    return jsonify({"error": "File too large. Maximum size is 100MB"}), 413

@app.errorhandler(500)
def internal_error(e):
    # Handler global (fallback), tapi detail error sudah ditangani di endpoint
    return jsonify({"error": "Internal server error"}), 500

# =====
# Main
# =====
if __name__ == "__main__":
    print("🚀 Starting Advanced File Converter...")
    print("📋 Supported conversions:")
    print("   - Documents: DOC/DOCX ↔ PDF/ODT/RTF/TXT/HTML")
    print("   - Spreadsheets: XLS/XLSX ↔ CSV/ODS")
    print("   - Presentations: PPT/PPTX ↔ PDF/ODP")
    print("⚠️  Note: LibreOffice must be installed for conversions to work")
    # Akses dari LAN: http://<IP>:5000/
    app.run(debug=True, host='0.0.0.0', port=5000)
