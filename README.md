# 🔮 Advanced File Converter

Aplikasi berbasis **Flask + LibreOffice** untuk mengonversi file (DOCX, PDF, PPTX, XLSX, TXT, CSV, dll.) dengan antarmuka web modern.  
Mendukung drag & drop, preview file, progress bar, dan hasil download otomatis.  

---

## 🚀 Fitur Utama
- 🌙 **UI Modern Dark Theme** dengan drag & drop upload  
- 📂 Dukungan berbagai format:
  - **Dokumen**: DOC, DOCX, ODT, RTF, TXT, HTML, PDF  
  - **Spreadsheet**: XLS, XLSX, ODS, CSV  
  - **Presentasi**: PPT, PPTX, ODP  
- ⚡ Progress bar animasi saat konversi  
- 📥 Hasil download otomatis  
- 🛡️ Validasi ukuran file max **100MB**  

---

## 📦 Instalasi

### 1. Clone Repository
```bash
git clone https://github.com/RifqiArdian09/FileConverter.git
cd FileConverter
```
### 2.Buat Virtual Environment
```bash
python -m venv venv
source venv/bin/activate   # Linux / Mac
venv\Scripts\activate      # Windows
```
### 3. Install Dependensi
```bash
pip install -r requirements.txt
```

### 4. Install LibreOffice
```bash
Windows: Download LibreOffice
Pastikan path soffice.exe ada di: [https://www.libreoffice.org/download/]
(https://www.libreoffice.org/download/)

C:\Program Files\LibreOffice\program\soffice.exe

atau tambahkan ke PATH

Linux / Mac:

sudo apt install libreoffice     # Ubuntu / Debian
brew install --cask libreoffice  # Mac (Homebrew)
```

### 5. Menjalankan Aplikasi
```bash
python app.py

```
