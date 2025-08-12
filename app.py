from flask import Flask, request, send_file, render_template_string
import tempfile
import os
import io
import subprocess

app = Flask(__name__)

# HTML (sama seperti sebelumnya, tidak diubah)
HTML_PAGE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>File Converter</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            background: #f7f7f7;
            display: flex;
            justify-content: center;
            align-items: center;
            height: 100vh;
        }
        .container {
            background: white;
            padding: 20px 30px;
            border-radius: 8px;
            box-shadow: 0 4px 8px rgba(0,0,0,0.1);
            width: 350px;
        }
        h2 {
            text-align: center;
        }
        label {
            font-weight: bold;
            display: block;
            margin-top: 10px;
        }
        select, input[type="file"], button {
            width: 100%;
            margin-top: 5px;
            padding: 8px;
            border-radius: 5px;
            border: 1px solid #ddd;
        }
        button {
            background: #007bff;
            color: white;
            border: none;
            margin-top: 15px;
            cursor: pointer;
        }
        button:hover {
            background: #0056b3;
        }
        #status {
            margin-top: 10px;
            text-align: center;
            font-size: 14px;
            color: green;
        }
    </style>
</head>
<body>
    <div class="container">
        <h2>File Converter</h2>
        <form id="convertForm">
            <label for="file">Pilih File:</label>
            <input type="file" id="file" name="file" required>

            <label for="to">Format Tujuan:</label>
            <select id="to" name="to" required>
                <option value="">-- Pilih --</option>
                <option value="pdf">PDF</option>
                <option value="docx">DOCX</option>
                <option value="pptx">PPTX</option>
                <option value="txt">TXT</option>
            </select>

            <button type="submit">Convert</button>
        </form>
        <div id="status"></div>
    </div>

    <script>
        document.getElementById("convertForm").addEventListener("submit", async function(e) {
            e.preventDefault();

            const fileInput = document.getElementById("file");
            const toFormat = document.getElementById("to").value;
            const status = document.getElementById("status");

            if (!fileInput.files.length) {
                alert("Pilih file terlebih dahulu");
                return;
            }
            if (!toFormat) {
                alert("Pilih format tujuan");
                return;
            }

            status.textContent = "Mengupload dan mengkonversi...";

            const formData = new FormData();
            formData.append("file", fileInput.files[0]);
            formData.append("to", toFormat);

            try {
                const res = await fetch("/convert", {
                    method: "POST",
                    body: formData
                });

                if (!res.ok) {
                    const err = await res.json();
                    throw new Error(err.error || "Konversi gagal");
                }

                const blob = await res.blob();
                const downloadUrl = URL.createObjectURL(blob);

                const a = document.createElement("a");
                a.href = downloadUrl;
                a.download = `converted.${toFormat}`;
                document.body.appendChild(a);
                a.click();
                a.remove();
                URL.revokeObjectURL(downloadUrl);

                status.textContent = "Konversi berhasil! File diunduh.";
            } catch (err) {
                console.error(err);
                status.textContent = "Gagal: " + err.message;
                status.style.color = "red";
            }
        });
    </script>
</body>
</html>
"""

def convert_with_libreoffice(input_path, output_dir, to_format):
    """Konversi file menggunakan LibreOffice CLI"""
    # Jalankan perintah LibreOffice headless
    subprocess.run([
        "soffice", "--headless", "--convert-to", to_format,
        "--outdir", output_dir, input_path
    ], check=True)

@app.route("/")
def index():
    return render_template_string(HTML_PAGE)

@app.route("/convert", methods=["POST"])
def convert_endpoint():
    try:
        file = request.files["file"]
        to_format = request.form["to"]

        with tempfile.TemporaryDirectory() as tmpdir:
            input_path = os.path.join(tmpdir, file.filename)
            file.save(input_path)

            # Format LibreOffice pakai ekstensi (contoh: pdf, docx, pptx, txt)
            output_ext = to_format.lower()
            convert_with_libreoffice(input_path, tmpdir, output_ext)

            # Cari file hasil konversi
            converted_file = None
            for f in os.listdir(tmpdir):
                if f.lower().endswith(f".{output_ext}"):
                    converted_file = os.path.join(tmpdir, f)
                    break

            if not converted_file:
                raise Exception("Gagal menemukan file hasil konversi.")

            with open(converted_file, "rb") as f:
                data = io.BytesIO(f.read())
            data.seek(0)

        return send_file(
            data,
            as_attachment=True,
            download_name=f"converted.{output_ext}"
        )
    except subprocess.CalledProcessError:
        return {"error": "Konversi gagal. Pastikan LibreOffice terinstall."}, 500
    except Exception as e:
        return {"error": str(e)}, 500

if __name__ == "__main__":
    app.run(debug=True)
