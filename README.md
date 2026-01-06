# 🔬 NaNote - Aplikasi Catatan & Kalkulator PSA Nanomaterial

![NaNote](https://img.shields.io/badge/NaNote-v1.0-blue)
![Streamlit](https://img.shields.io/badge/Streamlit-1.28.0-red)
![Python](https://img.shields.io/badge/Python-3.8%2B-green)

**NaNote** adalah aplikasi web berbasis Streamlit untuk mencatat hasil praktik dan mengkalkulasi hasil PSA (Particle Size Analysis) nanomaterial dalam bahasa Indonesia.

## ✨ Fitur Utama

### 📝 **Sistem Pencatatan Praktik**
- Form input lengkap untuk eksperimen nanomaterial
- Upload gambar hasil sintesis
- Penyimpanan data terstruktur
- **Ekspor ke format Word (.docx)**

### 🧮 **Kalkulator PSA**
- Input data distribusi ukuran (Diameter, % Volume, PDI)
- Analisis statistik lengkap (mean, median, std dev, variance)
- Visualisasi dengan Plotly
- Klasifikasi kualitas otomatis berdasarkan PDI
- **Ekspor ke format PDF**

### 📊 **Manajemen Data**
- Penyimpanan data dalam session
- Filter dan pencarian data
- Organisasi catatan dan hasil
- Backup data lokal

## 🚀 Instalasi & Deployment

### 1. Clone Repository
```bash
git clone https://github.com/username/nanote-app.git
cd nanote-app

streamlit run app.py
streamlit run app.py
