# Aplikasi Catatan Praktik & Kalkulator PSA Nanomaterial

Aplikasi web berbasis Streamlit untuk mencatat hasil praktik dan mengkalkulasi hasil PSA (Particle Size Analysis) nanomaterial.

## 🚀 Fitur Utama

### 📝 Sistem Pencatatan Praktik
- Input data eksperimen nanomaterial
- Upload gambar hasil sintesis
- Penyimpanan data dalam database session
- Ekspor ke format Word (.docx)

### 🧮 Kalkulator PSA
- Input data distribusi ukuran (Diameter, % Volume, PDI)
- Perhitungan statistik otomatis
- Visualisasi distribusi ukuran
- Klasifikasi kualitas nanomaterial berdasarkan PDI

### 📊 Sistem Ekspor
- Ekspor catatan praktik ke Word
- Ekspor hasil PSA ke PDF
- Preview sebelum download

## 🛠️ Teknologi yang Digunakan

- **Streamlit** - Framework web application
- **Python-docx** - Generasi dokumen Word
- **ReportLab** - Generasi PDF
- **Plotly** - Visualisasi data interaktif
- **Pandas & NumPy** - Pengolahan data

## 📦 Instalasi

1. Clone repository:
```bash
git clone https://github.com/username/psa-nanomaterial-app.git
cd psa-nanomaterial-app

pip install -r requirements.txt

streamlit run app.py
