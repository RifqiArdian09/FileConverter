# 🔮 File Converter

Aplikasi berbasis **Flask + LibreOffice** untuk mengonversi file (DOCX, PDF, PPTX, XLSX, TXT, CSV, dll.) dengan antarmuka web modern.  
Mendukung drag & drop, preview file, progress bar, dan hasil download otomatis. 

---

## 🚀 Fitur Utama

* 🌙 **Antarmuka Pengguna Modern**: Tampilan tema gelap (*dark theme*) dengan fitur *drag and drop* untuk mengunggah file.

* 📂 **Dukungan Format Luas**: Mampu mengonversi berbagai format file, termasuk:

  * **Dokumen**: DOC, DOCX, ODT, RTF, TXT, HTML, PDF

  * **Spreadsheet**: XLS, XLSX, ODS, CSV

  * **Presentasi**: PPT, PPTX, ODP

* ⚡ **Pengalaman Pengguna Interaktif**: Menampilkan *progress bar* animasi yang informatif saat proses konversi berlangsung.

* 📥 **Kemudahan Penggunaan**: File hasil akan secara otomatis diunduh ke perangkat Anda setelah konversi selesai.

* 🛡️ **Validasi File Aman**: Memiliki validasi untuk membatasi ukuran file maksimum hingga **100MB** guna menjaga stabilitas dan performa aplikasi.

---

## 📦 Instalasi dan Penggunaan

Ikuti langkah-langkah di bawah ini untuk menginstal dan menjalankan aplikasi.

### 1. Kloning Repositori

```bash
git clone https://github.com/RifqiArdian09/FileConverter.git
cd FileConverter
```

### 2. Buat dan Aktifkan Virtual Environment

Virtual environment direkomendasikan untuk mengisolasi dependensi proyek.

```bash
# Membuat virtual environment
python -m venv venv

# Mengaktifkan virtual environment
source venv/bin/activate    # Linux / macOS
venv\Scripts\activate       # Windows
```

### 3. Instal Dependensi Python

Pasang semua pustaka Python yang diperlukan dari file `requirements.txt`.

```bash
pip install -r requirements.txt
```

### 4. Instalasi LibreOffice

Pastikan **LibreOffice** terinstal di sistem Anda, karena aplikasi ini menggunakannya sebagai *engine* konversi.

* **Windows**:

  * Unduh dan instal LibreOffice dari situs resminya: <https://www.libreoffice.org/download/>

* Secara default, path `soffice.exe` biasanya berada di `C:\Program Files\LibreOffice\program\soffice.exe`. Jika tidak, pastikan untuk menambahkannya ke variabel ***PATH*** sistem Anda.

* **Linux / macOS**:

  * **Ubuntu / Debian**:

    ```bash
    sudo apt install libreoffice
    ```

  * **macOS (menggunakan Homebrew)**:

    ```bash
    brew install --cask libreoffice
    ```

### 5. Menjalankan Aplikasi

Jalankan server aplikasi Flask dengan perintah berikut:

```bash
python app.py
```

Setelah server berjalan, buka *browser* Anda dan akses `http://127.0.0.1:5000` untuk mulai menggunakan aplikasi.

---
