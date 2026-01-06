import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime
import tempfile
import os
import base64
from io import BytesIO
from utils.word_export import create_word_document
from utils.pdf_export import create_psa_pdf
from utils.data_handler import save_data, load_data, clear_data
from streamlit_option_menu import option_menu

# Konfigurasi halaman
st.set_page_config(
    page_title="Lab PSA Nano - Catatan & Kalkulator PSA",
    page_icon="🔬",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS kustom
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #2E86AB;
        text-align: center;
        padding: 1rem;
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        border-radius: 10px;
        margin-bottom: 2rem;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    
    .sub-header {
        font-size: 1.8rem;
        color: #2E86AB;
        border-left: 5px solid #2E86AB;
        padding-left: 1rem;
        margin: 1.5rem 0;
    }
    
    .info-box {
        background-color: #E8F4F8;
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 5px solid #2E86AB;
        margin: 1rem 0;
    }
    
    .success-box {
        background-color: #E8F8EF;
        padding: 1rem;
        border-radius: 8px;
        border-left: 5px solid #28A745;
        margin: 1rem 0;
    }
    
    .warning-box {
        background-color: #FFF3CD;
        padding: 1rem;
        border-radius: 8px;
        border-left: 5px solid #FFC107;
        margin: 1rem 0;
    }
    
    .stButton>button {
        background: linear-gradient(135deg, #2E86AB 0%, #A23B72 100%);
        color: white;
        border: none;
        padding: 0.5rem 2rem;
        border-radius: 5px;
        font-weight: bold;
        transition: all 0.3s ease;
    }
    
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(46, 134, 171, 0.4);
    }
    
    .card {
        background: white;
        padding: 1.5rem;
        border-radius: 10px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        margin: 1rem 0;
        border: 1px solid #e0e0e0;
    }
    
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 1.5rem;
        border-radius: 10px;
        text-align: center;
        margin: 0.5rem;
    }
    
    .sidebar .sidebar-content {
        background: linear-gradient(180deg, #2E86AB 0%, #A23B72 100%);
    }
</style>
""", unsafe_allow_html=True)

# Inisialisasi session state
def init_session_state():
    if 'catatan_list' not in st.session_state:
        st.session_state.catatan_list = load_data('catatan')
    
    if 'psa_results' not in st.session_state:
        st.session_state.psa_results = load_data('psa_results')
    
    if 'current_page' not in st.session_state:
        st.session_state.current_page = "beranda"
    
    if 'edit_mode' not in st.session_state:
        st.session_state.edit_mode = False
    
    if 'edit_index' not in st.session_state:
        st.session_state.edit_index = None

init_session_state()

# Fungsi navigasi
def set_page(page):
    st.session_state.current_page = page
    st.session_state.edit_mode = False
    st.session_state.edit_index = None

# Sidebar
with st.sidebar:
    st.markdown("""
    <div style="text-align: center; padding: 1rem;">
        <h1 style="color: white; font-size: 1.8rem;">🔬 Lab PSA Nano</h1>
        <p style="color: white; font-size: 0.9rem;">Aplikasi Catatan & Kalkulator PSA Nanomaterial</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Menu navigasi
    selected = option_menu(
        menu_title=None,
        options=["🏠 Beranda", "📝 Catatan Praktik", "🧮 Kalkulator PSA", "📊 Data Tersimpan", "📁 Ekspor", "⚙️ Panduan"],
        icons=['house', 'journal-text', 'calculator', 'database', 'download', 'info-circle'],
        menu_icon="cast",
        default_index=["🏠 Beranda", "📝 Catatan Praktik", "🧮 Kalkulator PSA", "📊 Data Tersimpan", "📁 Ekspor", "⚙️ Panduan"].index(st.session_state.current_page),
        styles={
            "container": {"padding": "0!important", "background-color": "rgba(255,255,255,0.1)"},
            "icon": {"color": "white", "font-size": "18px"}, 
            "nav-link": {"font-size": "16px", "text-align": "left", "margin": "0px", "color": "white"},
            "nav-link-selected": {"background-color": "rgba(255,255,255,0.2)"},
        }
    )
    
    if selected != st.session_state.current_page:
        set_page(selected)
    
    st.markdown("---")
    
    # Statistik
    st.markdown("### 📈 Statistik")
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Catatan", len(st.session_state.catatan_list))
    with col2:
        st.metric("Hasil PSA", len(st.session_state.psa_results))
    
    st.markdown("---")
    
    # Tools cepat
    st.markdown("### ⚡ Tools Cepat")
    if st.button("🗑️ Hapus Semua Data", use_container_width=True):
        if st.session_state.catatan_list or st.session_state.psa_results:
            if st.checkbox("Konfirmasi: Saya yakin ingin menghapus semua data"):
                clear_data()
                st.session_state.catatan_list = []
                st.session_state.psa_results = []
                st.success("Semua data berhasil dihapus!")
                st.rerun()
    
    st.markdown("---")
    
    # Info versi
    st.markdown("""
    <div style="text-align: center; color: white; font-size: 0.8rem;">
        <p>Lab PSA Nano v2.0</p>
        <p>© 2024 Laboratorium Nanomaterial</p>
    </div>
    """, unsafe_allow_html=True)

# Halaman Beranda
if st.session_state.current_page == "beranda":
    st.markdown('<div class="main-header">🔬 Lab PSA Nano</div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([2, 1, 2])
    
    with col2:
        st.image("https://img.icons8.com/color/300/000000/test-tube.png", width=200)
    
    with col1:
        st.markdown("### 📋 Tentang Aplikasi")
        st.markdown("""
        <div class="info-box">
        Aplikasi ini dirancang khusus untuk membantu praktikan nanomaterial dalam:
        
        • **Mencatat hasil praktik** dengan detail lengkap
        • **Mengkalkulasi hasil PSA** dari data PDI, %vol, dan diameter
        • **Menganalisis distribusi ukuran partikel**
        • **Menyimpan data** dalam format Word dan PDF
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown("### 🚀 Mulai Cepat")
        col_start1, col_start2 = st.columns(2)
        with col_start1:
            if st.button("📝 Buat Catatan", use_container_width=True):
                set_page("catatan_praktik")
        with col_start2:
            if st.button("🧮 Kalkulator PSA", use_container_width=True):
                set_page("kalkulator_psa")
    
    st.markdown("### ✨ Fitur Utama")
    
    feature_cols = st.columns(3)
    
    with feature_cols[0]:
        st.markdown("""
        <div class="card">
        <h4>📝 Sistem Pencatatan</h4>
        <ul>
        <li>Form input lengkap</li>
        <li>Upload gambar hasil</li>
        <li>Kategorisasi sampel</li>
        <li>Penyimpanan lokal</li>
        </ul>
        </div>
        """, unsafe_allow_html=True)
    
    with feature_cols[1]:
        st.markdown("""
        <div class="card">
        <h4>🧮 Kalkulator PSA</h4>
        <ul>
        <li>Input data distribusi</li>
        <li>Analisis statistik otomatis</li>
        <li>Visualisasi grafik</li>
        <li>Klasifikasi kualitas</li>
        </ul>
        </div>
        """, unsafe_allow_html=True)
    
    with feature_cols[2]:
        st.markdown("""
        <div class="card">
        <h4>📁 Sistem Ekspor</h4>
        <ul>
        <li>Export ke Word (.docx)</li>
        <li>Export ke PDF</li>
        <li>Preview sebelum download</li>
        <li>Multiple format</li>
        </ul>
        </div>
        """, unsafe_allow_html=True)
    
    # Recent Activity
    st.markdown("### 📈 Aktivitas Terbaru")
    if st.session_state.catatan_list or st.session_state.psa_results:
        recent_cols = st.columns(2)
        
        with recent_cols[0]:
            if st.session_state.catatan_list:
                st.markdown("**Catatan Terbaru:**")
                for catatan in list(reversed(st.session_state.catatan_list))[:3]:
                    with st.expander(f"📋 {catatan['judul'][:30]}..."):
                        st.write(f"**Praktikan:** {catatan['nama_praktikan']}")
                        st.write(f"**Tanggal:** {catatan['tanggal']}")
        
        with recent_cols[1]:
            if st.session_state.psa_results:
                st.markdown("**Hasil PSA Terbaru:**")
                for hasil in list(reversed(st.session_state.psa_results))[:3]:
                    with st.expander(f"📊 D={hasil['diameter_rerata']:.1f}nm, PDI={hasil['pdi_terhitung']:.3f}"):
                        st.write(f"**Klasifikasi:** {hasil['klasifikasi']}")
    else:
        st.info("Belum ada aktivitas. Mulai dengan membuat catatan atau menggunakan kalkulator PSA.")

# Halaman Catatan Praktik
elif st.session_state.current_page == "📝 Catatan Praktik":
    st.markdown('<div class="sub-header">📝 Catatan Praktik Nanomaterial</div>', unsafe_allow_html=True)
    
    # Mode Edit atau Tambah Baru
    if st.session_state.edit_mode and st.session_state.edit_index is not None:
        mode = "edit"
        catatan = st.session_state.catatan_list[st.session_state.edit_index]
        form_title = "✏️ Edit Catatan"
        button_label = "💾 Update Catatan"
    else:
        mode = "add"
        catatan = {}
        form_title = "📄 Buat Catatan Baru"
        button_label = "💾 Simpan Catatan"
    
    with st.form("form_catatan"):
        st.markdown(f"### {form_title}")
        
        # Bagian 1: Informasi Dasar
        st.markdown("#### 📋 Informasi Dasar")
        col1, col2 = st.columns(2)
        
        with col1:
            judul = st.text_input(
                "Judul Praktik*",
                value=catatan.get('judul', ''),
                placeholder="Contoh: Sintesis TiO₂ dengan Metode Sol-Gel"
            )
            nama_praktikan = st.text_input(
                "Nama Praktikan*",
                value=catatan.get('nama_praktikan', ''),
                placeholder="Nama lengkap praktikan"
            )
            tanggal = st.date_input(
                "Tanggal Praktik*",
                value=datetime.strptime(catatan.get('tanggal', str(datetime.now().date())), '%Y-%m-%d').date()
                if catatan.get('tanggal') else datetime.now().date()
            )
        
        with col2:
            institusi = st.text_input(
                "Institusi/Laboratorium",
                value=catatan.get('institusi', ''),
                placeholder="Universitas/Lab"
            )
            kelompok = st.text_input(
                "Kelompok Praktikum",
                value=catatan.get('kelompok', ''),
                placeholder="Kelompok/Shift"
            )
            supervisor = st.text_input(
                "Supervisor/Pembimbing",
                value=catatan.get('supervisor', ''),
                placeholder="Nama supervisor"
            )
        
        st.divider()
        
        # Bagian 2: Spesifikasi Nanomaterial
        st.markdown("#### 🔬 Spesifikasi Nanomaterial")
        col_nano1, col_nano2 = st.columns(2)
        
        with col_nano1:
            jenis_nanomaterial = st.selectbox(
                "Jenis Nanomaterial*",
                ["TiO₂ (Titanium Dioxide)", "SiO₂ (Silicon Dioxide)", "ZnO (Zinc Oxide)", 
                 "Ag (Silver Nanoparticles)", "Au (Gold Nanoparticles)", "Fe₃O₄ (Magnetite)",
                 "Al₂O₃ (Alumina)", "CuO (Copper Oxide)", "Lainnya"],
                index=0 if mode == "add" else ["TiO₂ (Titanium Dioxide)", "SiO₂ (Silicon Dioxide)", "ZnO (Zinc Oxide)", 
                 "Ag (Silver Nanoparticles)", "Au (Gold Nanoparticles)", "Fe₃O₄ (Magnetite)",
                 "Al₂O₃ (Alumina)", "CuO (Copper Oxide)", "Lainnya"].index(catatan.get('jenis_nanomaterial', "TiO₂ (Titanium Dioxide)"))
            )
            if jenis_nanomaterial == "Lainnya":
                jenis_nanomaterial = st.text_input("Sebutkan jenis nanomaterial", value=catatan.get('jenis_custom', ''))
        
        with col_nano2:
            metode_sintesis = st.selectbox(
                "Metode Sintesis*",
                ["Sol-Gel", "Hidrotermal", "Sonokimia", "Mekanokimia", 
                 "Chemical Vapor Deposition", "Electrospinning", "Co-precipitation", "Lainnya"],
                index=0 if mode == "add" else ["Sol-Gel", "Hidrotermal", "Sonokimia", "Mekanokimia", 
                 "Chemical Vapor Deposition", "Electrospinning", "Co-precipitation", "Lainnya"].index(catatan.get('metode_sintesis', "Sol-Gel"))
            )
            if metode_sintesis == "Lainnya":
                metode_sintesis = st.text_input("Sebutkan metode sintesis", value=catatan.get('metode_custom', ''))
        
        st.divider()
        
        # Bagian 3: Parameter Sintesis
        st.markdown("#### ⚙️ Parameter Sintesis")
        col_param1, col_param2, col_param3 = st.columns(3)
        
        with col_param1:
            suhu = st.number_input(
                "Suhu Sintesis (°C)*",
                min_value=-273.0,
                max_value=2000.0,
                value=float(catatan.get('suhu', 25.0)),
                step=0.1
            )
            waktu = st.number_input(
                "Waktu Sintesis (jam)*",
                min_value=0.0,
                max_value=500.0,
                value=float(catatan.get('waktu', 1.0)),
                step=0.1
            )
        
        with col_param2:
            tekanan = st.number_input(
                "Tekanan (atm)",
                min_value=0.0,
                max_value=1000.0,
                value=float(catatan.get('tekanan', 1.0)),
                step=0.1
            )
            ph = st.slider(
                "pH Larutan",
                0.0, 14.0, 
                value=float(catatan.get('ph', 7.0)),
                step=0.1
            )
        
        with col_param3:
            konsentrasi = st.number_input(
                "Konsentrasi (mg/mL)*",
                min_value=0.0,
                max_value=1000.0,
                value=float(catatan.get('konsentrasi', 1.0)),
                step=0.01
            )
            pelarut = st.text_input(
                "Jenis Pelarut*",
                value=catatan.get('pelarut', 'Aquades'),
                placeholder="Contoh: Aquades, Etanol, dll"
            )
        
        st.divider()
        
        # Bagian 4: Prosedur dan Hasil
        st.markdown("#### 📋 Prosedur Praktik")
        prosedur = st.text_area(
            "Tuliskan prosedur yang dilakukan*",
            value=catatan.get('prosedur', ''),
            height=150,
            placeholder="""Contoh:
1. Siapkan bahan awal: ...
2. Larutkan dalam pelarut: ...
3. Panaskan pada suhu ...°C selama ... jam
4. ..."""
        )
        
        st.markdown("#### 👁️ Hasil Pengamatan")
        hasil_pengamatan = st.text_area(
            "Tuliskan hasil pengamatan*",
            value=catatan.get('hasil_pengamatan', ''),
            height=150,
            placeholder="""Contoh:
- Warna akhir: putih susu
- Tekstur: koloidal
- Bau: tidak berbau
- ..."""
        )
        
        # Bagian 5: Upload Gambar
        st.markdown("#### 📸 Dokumentasi Visual")
        uploaded_images = st.file_uploader(
            "Upload gambar hasil sintesis (maks 5 gambar, format: JPG, PNG)",
            type=['jpg', 'jpeg', 'png'],
            accept_multiple_files=True,
            key=f"upload_{mode}_{st.session_state.edit_index if mode == 'edit' else 'new'}"
        )
        
        # Simpan gambar yang diupload sebelumnya jika edit mode
        existing_images = catatan.get('image_paths', []) if mode == "edit" else []
        
        # Bagian 6: Catatan Tambahan
        st.markdown("#### 📝 Catatan Tambahan")
        catatan_tambahan = st.text_area(
            "Catatan khusus, kendala, atau rencana tindak lanjut",
            value=catatan.get('catatan_tambahan', ''),
            height=100,
            placeholder="Masukkan catatan tambahan jika diperlukan..."
        )
        
        col_submit1, col_submit2 = st.columns([3, 1])
        
        with col_submit1:
            submitted = st.form_submit_button(
                button_label,
                type="primary",
                use_container_width=True
            )
        
        with col_submit2:
            if mode == "edit":
                if st.form_submit_button("❌ Batal Edit", use_container_width=True):
                    st.session_state.edit_mode = False
                    st.session_state.edit_index = None
                    st.rerun()
        
        if submitted:
            if judul and nama_praktikan and prosedur and hasil_pengamatan:
                # Proses upload gambar
                image_paths = existing_images.copy() if mode == "edit" else []
                if uploaded_images:
                    for img in uploaded_images:
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
                            tmp.write(img.getvalue())
                            image_paths.append(tmp.name)
                
                catatan_data = {
                    'id': catatan.get('id', len(st.session_state.catatan_list) + 1),
                    'judul': judul,
                    'nama_praktikan': nama_praktikan,
                    'tanggal': str(tanggal),
                    'institusi': institusi,
                    'kelompok': kelompok,
                    'supervisor': supervisor,
                    'jenis_nanomaterial': jenis_nanomaterial,
                    'metode_sintesis': metode_sintesis,
                    'suhu': suhu,
                    'waktu': waktu,
                    'tekanan': tekanan,
                    'ph': ph,
                    'konsentrasi': konsentrasi,
                    'pelarut': pelarut,
                    'prosedur': prosedur,
                    'hasil_pengamatan': hasil_pengamatan,
                    'catatan_tambahan': catatan_tambahan,
                    'image_paths': image_paths,
                    'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }
                
                if mode == "edit":
                    st.session_state.catatan_list[st.session_state.edit_index] = catatan_data
                    save_data('catatan', st.session_state.catatan_list)
                    st.session_state.edit_mode = False
                    st.session_state.edit_index = None
                    st.success(f"Catatan '{judul}' berhasil diperbarui!")
                else:
                    st.session_state.catatan_list.append(catatan_data)
                    save_data('catatan', st.session_state.catatan_list)
                    st.success(f"Catatan '{judul}' berhasil disimpan!")
                
                st.balloons()
                st.rerun()
            else:
                st.error("Harap isi semua field yang wajib (*)!")

# Halaman Kalkulator PSA
elif st.session_state.current_page == "🧮 Kalkulator PSA":
    st.markdown('<div class="sub-header">🧮 Kalkulator PSA Nanomaterial</div>', unsafe_allow_html=True)
    
    tab_kalkulator, tab_panduan = st.tabs(["🧮 Kalkulator", "📖 Panduan Perhitungan"])
    
    with tab_panduan:
        st.markdown("""
        ### 📖 Panduan Perhitungan PSA
        
        **Apa itu PSA?**
        Particle Size Analysis (PSA) adalah analisis untuk menentukan distribusi ukuran partikel dalam sampel nanomaterial.
        
        **Parameter yang Diperlukan:**
        1. **Diameter (nm)**: Ukuran partikel dalam nanometer
        2. **% Volume**: Persentase volume partikel pada ukuran tertentu
        3. **PDI**: Polydispersity Index (Indeks Polidispersitas)
        
        **Rumus Perhitungan:**
        - **Diameter Rata-rata**: Σ(Diameter × %Volume) / Σ(%Volume)
        - **PDI Rata-rata**: Σ(PDI × %Volume) / Σ(%Volume)
        - **Standard Deviation**: √[Σ(%Volume × (Diameter - Rata-rata)²) / Σ(%Volume)]
        - **PDI Terhitung**: Variance / (Diameter Rata-rata)²
        
        **Klasifikasi Berdasarkan PDI:**
        - **🟢 PDI < 0.1**: Monodispersi (Sangat Baik)
        - **🟡 PDI 0.1-0.2**: Hampir Monodispersi (Baik)
        - **🟠 PDI 0.2-0.3**: Polydispersi Sedang
        - **🔴 PDI > 0.3**: Polydispersi Tinggi
        """)
    
    with tab_kalkulator:
        # Metode Input
        input_method = st.radio(
            "Pilih Metode Input:",
            ["📝 Input Manual", "📁 Upload File Excel/CSV"],
            horizontal=True
        )
        
        if input_method == "📁 Upload File Excel/CSV":
            uploaded_file = st.file_uploader(
                "Upload file data (Excel atau CSV)",
                type=['xlsx', 'xls', 'csv'],
                help="File harus memiliki kolom: Diameter (nm), % Volume, dan PDI"
            )
            
            if uploaded_file:
                try:
                    if uploaded_file.name.endswith('.csv'):
                        df = pd.read_csv(uploaded_file)
                    else:
                        df = pd.read_excel(uploaded_file)
                    
                    # Validasi kolom
                    required_cols = ['Diameter (nm)', '% Volume', 'PDI']
                    if all(col in df.columns for col in required_cols):
                        st.session_state.psa_data = df[required_cols].copy()
                        st.success(f"File berhasil diupload! {len(df)} data ditemukan.")
                    else:
                        st.error("File harus mengandung kolom: 'Diameter (nm)', '% Volume', 'PDI'")
                except Exception as e:
                    st.error(f"Error membaca file: {str(e)}")
        
        # Input Manual
        else:
            col_input1, col_input2 = st.columns([2, 1])
            
            with col_input1:
                num_points = st.number_input(
                    "Jumlah Titik Data Distribusi",
                    min_value=3,
                    max_value=100,
                    value=10,
                    step=1,
                    help="Jumlah data titik untuk distribusi ukuran"
                )
            
            with col_input2:
                st.write("")
                st.write("")
                if st.button("🔄 Generate Data Contoh", use_container_width=True):
                    st.session_state.generate_table = True
            
            if 'psa_data' not in st.session_state or st.session_state.get('generate_table', False):
                # Generate data contoh yang realistis
                np.random.seed(42)
                diameters = np.sort(np.random.normal(50, 20, num_points))
                diameters = np.clip(diameters, 1, 200)
                
                # Buat distribusi normal untuk % volume
                volumes = np.exp(-(diameters - diameters.mean())**2 / (2 * (diameters.std()**2)))
                volumes = volumes / volumes.sum() * 100
                
                # Generate PDI yang berhubungan dengan diameter
                pdis = 0.1 + (np.abs(diameters - diameters.mean()) / diameters.max()) * 0.3
                
                data = {
                    'Diameter (nm)': np.round(diameters, 2),
                    '% Volume': np.round(volumes, 2),
                    'PDI': np.round(pdis, 3)
                }
                st.session_state.psa_data = pd.DataFrame(data)
                st.session_state.generate_table = False
        
        # Tampilkan dan edit data
        if 'psa_data' in st.session_state and not st.session_state.psa_data.empty:
            st.markdown("### 📊 Data Distribusi")
            
            # Editor tabel
            edited_df = st.data_editor(
                st.session_state.psa_data,
                use_container_width=True,
                num_rows="dynamic",
                column_config={
                    "Diameter (nm)": st.column_config.NumberColumn(
                        format="%.2f",
                        min_value=0.1,
                        max_value=10000.0,
                        required=True
                    ),
                    "% Volume": st.column_config.NumberColumn(
                        format="%.2f",
                        min_value=0.0,
                        max_value=100.0,
                        required=True
                    ),
                    "PDI": st.column_config.NumberColumn(
                        format="%.3f",
                        min_value=0.001,
                        max_value=1.0,
                        required=True
                    )
                }
            )
            
            # Validasi data
            total_volume = edited_df['% Volume'].sum()
            if abs(total_volume - 100) > 0.1:
                st.warning(f"Total % Volume = {total_volume:.2f}%. Disarankan total mendekati 100%.")
            
            # Tombol kalkulasi
            col_calc1, col_calc2, col_calc3 = st.columns([2, 1, 1])
            
            with col_calc2:
                calculate_btn = st.button(
                    "🧮 Hitung PSA",
                    type="primary",
                    use_container_width=True
                )
            
            with col_calc3:
                if st.button("💾 Simpan Data", use_container_width=True):
                    # Validasi
                    if edited_df['% Volume'].sum() > 0:
                        st.session_state.psa_data = edited_df.copy()
                        st.success("Data berhasil disimpan!")
                    else:
                        st.error("Total % Volume tidak boleh nol!")
            
            if calculate_btn:
                with st.spinner("Menghitung hasil PSA..."):
                    try:
                        # Validasi data
                        if edited_df.empty:
                            st.error("Data tidak boleh kosong!")
                        elif edited_df['% Volume'].sum() == 0:
                            st.error("Total % Volume tidak boleh nol!")
                        else:
                            # Normalisasi % Volume jika perlu
                            df_calc = edited_df.copy()
                            if abs(total_volume - 100) > 0.1:
                                df_calc['% Volume Normalized'] = (df_calc['% Volume'] / total_volume * 100)
                            else:
                                df_calc['% Volume Normalized'] = df_calc['% Volume']
                            
                            # Hitung statistik
                            # Diameter rata-rata berbobot
                            diameter_weighted = np.average(
                                df_calc['Diameter (nm)'],
                                weights=df_calc['% Volume Normalized']
                            )
                            
                            # PDI rata-rata berbobot
                            pdi_weighted = np.average(
                                df_calc['PDI'],
                                weights=df_calc['% Volume Normalized']
                            )
                            
                            # Variance dan standard deviation
                            variance = np.average(
                                (df_calc['Diameter (nm)'] - diameter_weighted) ** 2,
                                weights=df_calc['% Volume Normalized']
                            )
                            std_dev = np.sqrt(variance)
                            
                            # PDI terhitung
                            pdi_calculated = variance / (diameter_weighted ** 2)
                            
                            # Koefisien variasi
                            cv = (std_dev / diameter_weighted) * 100
                            
                            # Mode (diameter dengan % volume tertinggi)
                            mode_idx = df_calc['% Volume Normalized'].idxmax()
                            mode_diameter = df_calc.loc[mode_idx, 'Diameter (nm)']
                            mode_percentage = df_calc.loc[mode_idx, '% Volume Normalized']
                            
                            # Klasifikasi berdasarkan PDI
                            if pdi_calculated < 0.05:
                                klasifikasi = "Sangat Monodispersi (Excellent)"
                                warna = "🟢"
                                grade = "A+"
                            elif pdi_calculated < 0.1:
                                klasifikasi = "Monodispersi (Sangat Baik)"
                                warna = "🟢"
                                grade = "A"
                            elif pdi_calculated < 0.2:
                                klasifikasi = "Hampir Monodispersi (Baik)"
                                warna = "🟡"
                                grade = "B"
                            elif pdi_calculated < 0.3:
                                klasifikasi = "Polydispersi Sedang (Cukup)"
                                warna = "🟠"
                                grade = "C"
                            else:
                                klasifikasi = "Polydispersi Tinggi (Perlu Optimasi)"
                                warna = "🔴"
                                grade = "D"
                            
                            # Simpan hasil
                            hasil_psa = {
                                'dataframe': df_calc.copy(),
                                'diameter_rerata': diameter_weighted,
                                'pdi_rerata': pdi_weighted,
                                'pdi_terhitung': pdi_calculated,
                                'std_dev': std_dev,
                                'variance': variance,
                                'cv': cv,
                                'mode_diameter': mode_diameter,
                                'mode_percentage': mode_percentage,
                                'klasifikasi': klasifikasi,
                                'warna': warna,
                                'grade': grade,
                                'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                'total_points': len(df_calc)
                            }
                            
                            st.session_state.psa_results.append(hasil_psa)
                            save_data('psa_results', st.session_state.psa_results)
                            
                            st.success("✅ Perhitungan PSA berhasil!")
                            
                            # Tampilkan hasil
                            st.markdown("### 📈 Hasil Kalkulasi PSA")
                            
                            # Metrics cards
                            col_metric1, col_metric2, col_metric3, col_metric4 = st.columns(4)
                            
                            with col_metric1:
                                st.markdown(f"""
                                <div class="metric-card">
                                <h4>Diameter Rata-rata</h4>
                                <h2>{diameter_weighted:.1f} nm</h2>
                                <p>± {std_dev:.1f} nm</p>
                                </div>
                                """, unsafe_allow_html=True)
                            
                            with col_metric2:
                                st.markdown(f"""
                                <div class="metric-card">
                                <h4>PDI Terhitung</h4>
                                <h2>{pdi_calculated:.3f}</h2>
                                <p>{grade} {warna}</p>
                                </div>
                                """, unsafe_allow_html=True)
                            
                            with col_metric3:
                                st.markdown(f"""
                                <div class="metric-card">
                                <h4>Mode Diameter</h4>
                                <h2>{mode_diameter:.1f} nm</h2>
                                <p>{mode_percentage:.1f}% volume</p>
                                </div>
                                """, unsafe_allow_html=True)
                            
                            with col_metric4:
                                st.markdown(f"""
                                <div class="metric-card">
                                <h4>Koef. Variasi</h4>
                                <h2>{cv:.1f}%</h2>
                                <p>Std Dev / Mean</p>
                                </div>
                                """, unsafe_allow_html=True)
                            
                            # Detail hasil
                            with st.expander("📊 Detail Hasil Perhitungan", expanded=True):
                                col_detail1, col_detail2 = st.columns(2)
                                
                                with col_detail1:
                                    st.markdown("**📋 Statistik Deskriptif**")
                                    stats_data = {
                                        'Parameter': ['Minimum', 'Maksimum', 'Median', 'Mean', 'Std Dev', 'Variance'],
                                        'Nilai (nm)': [
                                            f"{df_calc['Diameter (nm)'].min():.2f}",
                                            f"{df_calc['Diameter (nm)'].max():.2f}",
                                            f"{df_calc['Diameter (nm)'].median():.2f}",
                                            f"{df_calc['Diameter (nm)'].mean():.2f}",
                                            f"{std_dev:.2f}",
                                            f"{variance:.2f}"
                                        ]
                                    }
                                    st.table(pd.DataFrame(stats_data))
                                
                                with col_detail2:
                                    st.markdown("**🎯 Kualitas Nanomaterial**")
                                    quality_data = {
                                        'Parameter': ['PDI Terhitung', 'Klasifikasi', 'Grade', 'Koef. Variasi', 'Distribusi Utama'],
                                        'Nilai': [
                                            f"{pdi_calculated:.3f}",
                                            f"{warna} {klasifikasi}",
                                            grade,
                                            f"{cv:.1f}%",
                                            f"{mode_diameter:.1f} nm ({mode_percentage:.1f}%)"
                                        ]
                                    }
                                    st.table(pd.DataFrame(quality_data))
                                
                                # Rekomendasi
                                st.markdown("**💡 Rekomendasi & Interpretasi**")
                                if pdi_calculated < 0.1:
                                    st.success("""
                                    **Kualitas Sangat Baik!** Nanomaterial Anda memiliki distribusi ukuran yang sangat seragam.
                                    **Rekomendasi:**
                                    - Lanjutkan metode sintesis dengan parameter yang sama
                                    - Cocok untuk aplikasi biomedis dan elektronik presisi
                                    - Pertimbangkan untuk publikasi hasil
                                    """)
                                elif pdi_calculated < 0.2:
                                    st.info("""
                                    **Kualitas Baik.** Distribusi ukuran cukup seragam untuk kebanyakan aplikasi.
                                    **Rekomendasi:**
                                    - Dapat digunakan untuk aplikasi katalisis dan coating
                                    - Optimasi kecil dapat meningkatkan monodispersitas
                                    - Evaluasi efek pH dan konsentrasi
                                    """)
                                elif pdi_calculated < 0.3:
                                    st.warning("""
                                    **Perlu Optimasi.** Distribusi ukuran cukup lebar.
                                    **Rekomendasi:**
                                    - Evaluasi parameter sintesis (suhu, waktu, stirring rate)
                                    - Pertimbangkan penggunaan surfaktan atau stabilizer
                                    - Cocok untuk aplikasi bulk material
                                    """)
                                else:
                                    st.error("""
                                    **Perlu Optimasi Signifikan.** Distribusi ukuran sangat lebar.
                                    **Rekomendasi:**
                                    - Evaluasi ulang metode sintesis
                                    - Optimasi parameter utama
                                    - Pertimbangkan metode purifikasi
                                    - Cocok untuk aplikasi konstruksi
                                    """)
                            
                            # Visualisasi
                            st.markdown("### 📊 Visualisasi Distribusi")
                            
                            tab_viz1, tab_viz2, tab_viz3 = st.tabs(["📈 Distribusi Ukuran", "📊 Histogram", "🎯 Scatter Plot"])
                            
                            with tab_viz1:
                                fig1 = go.Figure()
                                
                                # Bar chart untuk distribusi
                                fig1.add_trace(go.Bar(
                                    x=df_calc['Diameter (nm)'],
                                    y=df_calc['% Volume Normalized'],
                                    name='% Volume',
                                    marker_color='royalblue',
                                    opacity=0.8,
                                    hovertemplate='Diameter: %{x:.1f} nm<br>% Volume: %{y:.1f}%<extra></extra>'
                                ))
                                
                                # Garis untuk rata-rata
                                fig1.add_vline(
                                    x=diameter_weighted,
                                    line_dash="dash",
                                    line_color="red",
                                    annotation_text=f"Rata-rata: {diameter_weighted:.1f} nm",
                                    annotation_position="top right"
                                )
                                
                                # Area untuk standard deviation
                                fig1.add_vrect(
                                    x0=diameter_weighted - std_dev,
                                    x1=diameter_weighted + std_dev,
                                    fillcolor="rgba(255, 0, 0, 0.1)",
                                    line_width=0,
                                    annotation_text=f"± {std_dev:.1f} nm",
                                    annotation_position="bottom right"
                                )
                                
                                fig1.update_layout(
                                    title='Distribusi Ukuran Partikel Nanomaterial',
                                    xaxis_title='Diameter (nm)',
                                    yaxis_title='% Volume',
                                    template='plotly_white',
                                    hovermode='x unified',
                                    height=500
                                )
                                st.plotly_chart(fig1, use_container_width=True)
                            
                            with tab_viz2:
                                fig2 = px.histogram(
                                    df_calc,
                                    x='Diameter (nm)',
                                    y='% Volume Normalized',
                                    nbins=20,
                                    title='Histogram Distribusi Ukuran',
                                    labels={'Diameter (nm)': 'Diameter (nm)', '% Volume Normalized': '% Volume'}
                                )
                                
                                fig2.update_layout(
                                    template='plotly_white',
                                    height=500
                                )
                                st.plotly_chart(fig2, use_container_width=True)
                            
                            with tab_viz3:
                                fig3 = px.scatter(
                                    df_calc,
                                    x='Diameter (nm)',
                                    y='PDI',
                                    size='% Volume Normalized',
                                    color='% Volume Normalized',
                                    hover_data=['% Volume Normalized'],
                                    title='Hubungan Diameter vs PDI',
                                    size_max=30,
                                    color_continuous_scale=px.colors.sequential.Viridis
                                )
                                
                                fig3.update_layout(
                                    xaxis_title='Diameter (nm)',
                                    yaxis_title='PDI',
                                    template='plotly_white',
                                    height=500
                                )
                                st.plotly_chart(fig3, use_container_width=True)
                            
                            # Tombol ekspor
                            col_export1, col_export2 = st.columns(2)
                            
                            with col_export1:
                                if st.button("📥 Ekspor ke PDF", use_container_width=True):
                                    try:
                                        pdf_path = create_psa_pdf(hasil_psa, len(st.session_state.psa_results))
                                        with open(pdf_path, 'rb') as f:
                                            pdf_data = f.read()
                                        
                                        st.download_button(
                                            label="⬇️ Download Laporan PDF",
                                            data=pdf_data,
                                            file_name=f"laporan_psa_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                                            mime="application/pdf",
                                            use_container_width=True
                                        )
                                        os.unlink(pdf_path)
                                    except Exception as e:
                                        st.error(f"Error: {str(e)}")
                            
                            with col_export2:
                                # Ekspor data ke Excel
                                excel_buffer = BytesIO()
                                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                                    df_calc.to_excel(writer, sheet_name='Data Distribusi', index=False)
                                    
                                    # Buat sheet hasil
                                    hasil_df = pd.DataFrame([{
                                        'Parameter': 'Diameter Rata-rata (nm)',
                                        'Nilai': diameter_weighted
                                    }, {
                                        'Parameter': 'PDI Terhitung',
                                        'Nilai': pdi_calculated
                                    }, {
                                        'Parameter': 'Standard Deviation (nm)',
                                        'Nilai': std_dev
                                    }, {
                                        'Parameter': 'Klasifikasi',
                                        'Nilai': klasifikasi
                                    }])
                                    hasil_df.to_excel(writer, sheet_name='Hasil PSA', index=False)
                                
                                excel_data = excel_buffer.getvalue()
                                
                                st.download_button(
                                    label="📊 Ekspor Data ke Excel",
                                    data=excel_data,
                                    file_name=f"data_psa_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True
                                )
                    
                    except Exception as e:
                        st.error(f"Terjadi kesalahan dalam perhitungan: {str(e)}")
                        st.exception(e)

# Halaman Data Tersimpan
elif st.session_state.current_page == "📊 Data Tersimpan":
    st.markdown('<div class="sub-header">📊 Data Tersimpan</div>', unsafe_allow_html=True)
    
    tab_catatan, tab_psa = st.tabs(["📝 Catatan Praktik", "📊 Hasil PSA"])
    
    with tab_catatan:
        if st.session_state.catatan_list:
            st.markdown(f"### 📚 Total {len(st.session_state.catatan_list)} Catatan Praktik")
            
            # Filter dan pencarian
            col_filter1, col_filter2, col_filter3 = st.columns(3)
            
            with col_filter1:
                search_term = st.text_input("🔍 Cari catatan...", placeholder="Judul atau nama praktikan")
            
            with col_filter2:
                filter_nanomaterial = st.multiselect(
                    "Filter Nanomaterial",
                    options=list(set([c['jenis_nanomaterial'] for c in st.session_state.catatan_list])),
                    default=[]
                )
            
            with col_filter3:
                filter_method = st.multiselect(
                    "Filter Metode",
                    options=list(set([c['metode_sintesis'] for c in st.session_state.catatan_list])),
                    default=[]
                )
            
            # Filter data
            filtered_catatan = st.session_state.catatan_list
            
            if search_term:
                filtered_catatan = [
                    c for c in filtered_catatan 
                    if search_term.lower() in c['judul'].lower() 
                    or search_term.lower() in c['nama_praktikan'].lower()
                ]
            
            if filter_nanomaterial:
                filtered_catatan = [
                    c for c in filtered_catatan 
                    if c['jenis_nanomaterial'] in filter_nanomaterial
                ]
            
            if filter_method:
                filtered_catatan = [
                    c for c in filtered_catatan 
                    if c['metode_sintesis'] in filter_method
                ]
            
            # Tampilkan catatan
            if filtered_catatan:
                for idx, catatan in enumerate(filtered_catatan):
                    original_idx = st.session_state.catatan_list.index(catatan)
                    
                    with st.expander(f"📋 {catatan['judul']} - {catatan['tanggal']}", expanded=False):
                        col_info, col_action = st.columns([3, 1])
                        
                        with col_info:
                            st.write(f"**Praktikan:** {catatan['nama_praktikan']}")
                            st.write(f"**Nanomaterial:** {catatan['jenis_nanomaterial']}")
                            st.write(f"**Metode:** {catatan['metode_sintesis']}")
                            st.write(f"**Suhu:** {catatan['suhu']}°C | **Waktu:** {catatan['waktu']} jam")
                            st.write(f"**pH:** {catatan['ph']} | **Konsentrasi:** {catatan['konsentrasi']} mg/mL")
                            
                            if catatan['catatan_tambahan']:
                                with st.expander("Catatan Tambahan"):
                                    st.write(catatan['catatan_tambahan'])
                            
                            if catatan['image_paths']:
                                st.write("**Gambar Hasil:**")
                                cols_img = st.columns(min(3, len(catatan['image_paths'])))
                                for img_idx, img_path in enumerate(catatan['image_paths'][:3]):
                                    with cols_img[img_idx]:
                                        try:
                                            st.image(img_path, use_column_width=True)
                                        except:
                                            st.warning("Gambar tidak dapat ditampilkan")
                        
                        with col_action:
                            if st.button("✏️ Edit", key=f"edit_{original_idx}", use_container_width=True):
                                st.session_state.edit_mode = True
                                st.session_state.edit_index = original_idx
                                set_page("catatan_praktik")
                                st.rerun()
                            
                            if st.button("📥 Word", key=f"word_{original_idx}", use_container_width=True):
                                try:
                                    doc_path = create_word_document(catatan)
                                    with open(doc_path, 'rb') as f:
                                        doc_data = f.read()
                                    
                                    st.download_button(
                                        label="⬇️ Download",
                                        data=doc_data,
                                        file_name=f"catatan_{catatan['judul'][:20]}_{catatan['tanggal']}.docx",
                                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                        key=f"download_{original_idx}"
                                    )
                                    os.unlink(doc_path)
                                except Exception as e:
                                    st.error(f"Error: {str(e)}")
                            
                            if st.button("🗑️", key=f"delete_{original_idx}", use_container_width=True):
                                st.session_state.catatan_list.pop(original_idx)
                                save_data('catatan', st.session_state.catatan_list)
                                st.success("Catatan berhasil dihapus!")
                                st.rerun()
            
            else:
                st.info("Tidak ada catatan yang sesuai dengan filter.")
        
        else:
            st.info("Belum ada catatan praktik. Silahkan buat catatan baru di halaman Catatan Praktik.")
    
    with tab_psa:
        if st.session_state.psa_results:
            st.markdown(f"### 📊 Total {len(st.session_state.psa_results)} Hasil PSA")
            
            # Filter hasil
            col_filter_psa1, col_filter_psa2 = st.columns(2)
            
            with col_filter_psa1:
                pdi_min = st.slider("PDI Minimum", 0.0, 1.0, 0.0, 0.01)
                pdi_max = st.slider("PDI Maksimum", 0.0, 1.0, 1.0, 0.01)
            
            with col_filter_psa2:
                diameter_min = st.number_input("Diameter Min (nm)", 0.0, 1000.0, 0.0, 1.0)
                diameter_max = st.number_input("Diameter Max (nm)", 0.0, 1000.0, 500.0, 1.0)
            
            # Filter data
            filtered_results = [
                r for r in st.session_state.psa_results
                if pdi_min <= r['pdi_terhitung'] <= pdi_max
                and diameter_min <= r['diameter_rerata'] <= diameter_max
            ]
            
            if filtered_results:
                for idx, hasil in enumerate(filtered_results):
                    original_idx = st.session_state.psa_results.index(hasil)
                    
                    with st.expander(f"📊 Hasil PSA #{original_idx + 1} - {hasil['timestamp']}", expanded=False):
                        col_res, col_act = st.columns([3, 1])
                        
                        with col_res:
                            st.write(f"**Diameter Rata-rata:** {hasil['diameter_rerata']:.2f} nm")
                            st.write(f"**PDI Terhitung:** {hasil['pdi_terhitung']:.3f}")
                            st.write(f"**Standard Deviation:** {hasil['std_dev']:.2f} nm")
                            st.write(f"**Klasifikasi:** {hasil['warna']} {hasil['klasifikasi']}")
                            st.write(f"**Jumlah Data:** {hasil['total_points']} titik")
                            
                            # Tampilkan preview data
                            with st.expander("Preview Data"):
                                st.dataframe(hasil['dataframe'].head(), use_container_width=True)
                        
                        with col_act:
                            if st.button("📥 PDF", key=f"pdf_{original_idx}", use_container_width=True):
                                try:
                                    pdf_path = create_psa_pdf(hasil, original_idx + 1)
                                    with open(pdf_path, 'rb') as f:
                                        pdf_data = f.read()
                                    
                                    st.download_button(
                                        label="⬇️ Download",
                                        data=pdf_data,
                                        file_name=f"hasil_psa_{original_idx + 1}_{hasil['timestamp'].replace(':', '-')}.pdf",
                                        mime="application/pdf",
                                        key=f"download_pdf_{original_idx}"
                                    )
                                    os.unlink(pdf_path)
                                except Exception as e:
                                    st.error(f"Error: {str(e)}")
                            
                            if st.button("🗑️", key=f"del_psa_{original_idx}", use_container_width=True):
                                st.session_state.psa_results.pop(original_idx)
                                save_data('psa_results', st.session_state.psa_results)
                                st.success("Hasil PSA berhasil dihapus!")
                                st.rerun()
            
            else:
                st.info("Tidak ada hasil PSA yang sesuai dengan filter.")
        
        else:
            st.info("Belum ada hasil PSA. Silahkan gunakan kalkulator PSA terlebih dahulu.")

# Halaman Ekspor
elif st.session_state.current_page == "📁 Ekspor":
    st.markdown('<div class="sub-header">📁 Ekspor Data</div>', unsafe_allow_html=True)
    
    tab_batch, tab_single = st.tabs(["📦 Ekspor Batch", "📄 Ekspor Per Item"])
    
    with tab_batch:
        st.markdown("### 📦 Ekspor Data dalam Batch")
        
        col_batch1, col_batch2 = st.columns(2)
        
        with col_batch1:
            st.markdown("#### Catatan Praktik")
            if st.session_state.catatan_list:
                catatan_options = {f"{c['id']}: {c['judul'][:50]}": c for c in st.session_state.catatan_list}
                selected_catatan = st.multiselect(
                    "Pilih catatan untuk diekspor",
                    options=list(catatan_options.keys()),
                    default=[]
                )
                
                if selected_catatan:
                    if st.button("📥 Ekspor Catatan Terpilih ke Word", use_container_width=True):
                        try:
                            # Gabungkan catatan terpilih
                            selected_data = [catatan_options[key] for key in selected_catatan]
                            
                            # Buat dokumen gabungan
                            from utils.word_export import create_batch_word_document
                            doc_path = create_batch_word_document(selected_data)
                            
                            with open(doc_path, 'rb') as f:
                                doc_data = f.read()
                            
                            st.download_button(
                                label=f"⬇️ Download {len(selected_data)} Catatan",
                                data=doc_data,
                                file_name=f"batch_catatan_{datetime.now().strftime('%Y%m%d')}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True
                            )
                            os.unlink(doc_path)
                        except Exception as e:
                            st.error(f"Error: {str(e)}")
            
            else:
                st.info("Belum ada catatan untuk diekspor")
        
        with col_batch2:
            st.markdown("#### Hasil PSA")
            if st.session_state.psa_results:
                psa_options = {f"Hasil #{i+1} - D={r['diameter_rerata']:.1f}nm": r 
                              for i, r in enumerate(st.session_state.psa_results)}
                selected_psa = st.multiselect(
                    "Pilih hasil PSA untuk diekspor",
                    options=list(psa_options.keys()),
                    default=[]
                )
                
                if selected_psa:
                    if st.button("📥 Ekspor Hasil PSA ke PDF", use_container_width=True):
                        try:
                            # Gabungkan hasil terpilih
                            selected_data = [psa_options[key] for key in selected_psa]
                            
                            # Buat PDF gabungan
                            from utils.pdf_export import create_batch_pdf
                            pdf_path = create_batch_pdf(selected_data)
                            
                            with open(pdf_path, 'rb') as f:
                                pdf_data = f.read()
                            
                            st.download_button(
                                label=f"⬇️ Download {len(selected_data)} Hasil PSA",
                                data=pdf_data,
                                file_name=f"batch_psa_{datetime.now().strftime('%Y%m%d')}.pdf",
                                mime="application/pdf",
                                use_container_width=True
                            )
                            os.unlink(pdf_path)
                        except Exception as e:
                            st.error(f"Error: {str(e)}")
            
            else:
                st.info("Belum ada hasil PSA untuk diekspor")
    
    with tab_single:
        st.markdown("### 📄 Ekspor Individual")
        
        col_single1, col_single2 = st.columns(2)
        
        with col_single1:
            st.markdown("#### Ekspor Catatan ke Word")
            if st.session_state.catatan_list:
                catatan_list = [f"{c['id']}: {c['judul'][:50]} - {c['tanggal']}" 
                               for c in st.session_state.catatan_list]
                selected_note = st.selectbox("Pilih catatan", catatan_list)
                
                if selected_note:
                    note_idx = int(selected_note.split(":")[0]) - 1
                    catatan = st.session_state.catatan_list[note_idx]
                    
                    # Preview
                    with st.expander("👁️ Preview Catatan"):
                        st.write(f"**Judul:** {catatan['judul']}")
                        st.write(f"**Praktikan:** {catatan['nama_praktikan']}")
                        st.write(f"**Tanggal:** {catatan['tanggal']}")
                    
                    # Tombol ekspor
                    if st.button("📥 Ekspor ke Word", use_container_width=True):
                        try:
                            doc_path = create_word_document(catatan)
                            with open(doc_path, 'rb') as f:
                                doc_data = f.read()
                            
                            st.download_button(
                                label="⬇️ Download File Word",
                                data=doc_data,
                                file_name=f"catatan_{catatan['judul'][:20]}_{catatan['tanggal']}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True
                            )
                            os.unlink(doc_path)
                        except Exception as e:
                            st.error(f"Error: {str(e)}")
            
            else:
                st.info("Belum ada catatan untuk diekspor")
        
        with col_single2:
            st.markdown("#### Ekspor Hasil PSA ke PDF")
            if st.session_state.psa_results:
                psa_list = [f"Hasil #{i+1} - D={r['diameter_rerata']:.1f}nm, PDI={r['pdi_terhitung']:.3f}" 
                           for i, r in enumerate(st.session_state.psa_results)]
                selected_psa = st.selectbox("Pilih hasil PSA", psa_list)
                
                if selected_psa:
                    psa_idx = int(selected_psa.split("#")[1].split(" ")[0]) - 1
                    hasil = st.session_state.psa_results[psa_idx]
                    
                    # Preview
                    with st.expander("👁️ Preview Hasil"):
                        st.write(f"**Diameter Rata-rata:** {hasil['diameter_rerata']:.2f} nm")
                        st.write(f"**PDI Terhitung:** {hasil['pdi_terhitung']:.3f}")
                        st.write(f"**Klasifikasi:** {hasil['klasifikasi']}")
                    
                    # Tombol ekspor
                    if st.button("📥 Ekspor ke PDF", use_container_width=True, key="export_single_pdf"):
                        try:
                            pdf_path = create_psa_pdf(hasil, psa_idx + 1)
                            with open(pdf_path, 'rb') as f:
                                pdf_data = f.read()
                            
                            st.download_button(
                                label="⬇️ Download File PDF",
                                data=pdf_data,
                                file_name=f"hasil_psa_{psa_idx + 1}_{hasil['timestamp'].replace(':', '-')}.pdf",
                                mime="application/pdf",
                                use_container_width=True
                            )
                            os.unlink(pdf_path)
                        except Exception as e:
                            st.error(f"Error: {str(e)}")
            
            else:
                st.info("Belum ada hasil PSA untuk diekspor")

# Halaman Panduan
elif st.session_state.current_page == "⚙️ Panduan":
    st.markdown('<div class="sub-header">⚙️ Panduan Penggunaan</div>', unsafe_allow_html=True)
    
    tab_guide, tab_about = st.tabs(["📖 Panduan", "ℹ️ Tentang"])
    
    with tab_guide:
        st.markdown("""
        ### 🎯 Panduan Lengkap Penggunaan Aplikasi
        
        #### 1. 📝 Modul Catatan Praktik
        **Fungsi:** Mencatat seluruh proses sintesis nanomaterial
        
        **Langkah-langkah:**
        1. Buka halaman **"📝 Catatan Praktik"**
        2. Isi semua field yang wajib (*)
        3. Upload gambar hasil sintesis (opsional)
        4. Klik **"💾 Simpan Catatan"**
        5. Catatan akan tersimpan dan dapat diekspor ke Word
        
        **Tips:**
        - Gunakan deskripsi yang detail pada bagian prosedur
        - Catat semua parameter sintesis dengan teliti
        - Upload gambar untuk dokumentasi visual
        
        #### 2. 🧮 Modul Kalkulator PSA
        **Fungsi:** Menganalisis distribusi ukuran partikel nanomaterial
        
        **Langkah-langkah:**
        1. Buka halaman **"🧮 Kalkulator PSA"**
        2. Pilih metode input (manual atau upload file)
        3. Input data Diameter (nm), % Volume, dan PDI
        4. Klik **"🧮 Hitung PSA"**
        5. Hasil akan tampil dengan statistik lengkap
        6. Ekspor hasil ke PDF jika diperlukan
        
        **Parameter Input:**
        - **Diameter (nm):** Ukuran partikel dalam nanometer
        - **% Volume:** Persentase volume pada ukuran tertentu
        - **PDI:** Polydispersity Index (0-1)
        
        #### 3. 📊 Interpretasi Hasil PSA
        
        **Klasifikasi PDI:**
        - **🟢 PDI < 0.1:** Monodispersi - Sangat baik untuk aplikasi presisi
        - **🟡 PDI 0.1-0.2:** Hampir monodispersi - Baik untuk kebanyakan aplikasi
        - **🟠 PDI 0.2-0.3:** Polydispersi sedang - Perlu optimasi
        - **🔴 PDI > 0.3:** Polydispersi tinggi - Perlu optimasi signifikan
        
        #### 4. 📁 Modul Ekspor Data
        
        **Ekspor ke Word (.docx):**
        - Format profesional untuk laporan praktikum
        - Termasuk semua data dan gambar
        - Siap untuk dicetak atau disubmit
        
        **Ekspor ke PDF:**
        - Format fixed untuk hasil PSA
        - Termasuk grafik dan statistik
        - Cocok untuk publikasi atau presentasi
        
        #### 5. 💾 Manajemen Data
        
        **Penyimpanan:**
        - Data disimpan secara lokal dalam session
        - Bertahan selama aplikasi berjalan
        - Ekspor untuk penyimpanan permanen
        
        **Keamanan:**
        - Backup data dengan ekspor reguler
        - Simpan file Word/PDF di lokasi aman
        """)
    
    with tab_about:
        st.markdown("""
        ### ℹ️ Tentang Aplikasi
        
        **Lab PSA Nano v2.0**
        
        **Deskripsi:**
        Aplikasi web berbasis Streamlit untuk mencatat hasil praktik dan mengkalkulasi 
        hasil PSA (Particle Size Analysis) nanomaterial.
        
        **Fitur Utama:**
        - Sistem pencatatan praktik nanomaterial
        - Kalkulator PSA dengan analisis statistik
        - Visualisasi distribusi ukuran partikel
        - Ekspor data ke Word dan PDF
        - Manajemen data terintegrasi
        
        **Teknologi:**
        - **Frontend:** Streamlit
        - **Backend:** Python
        - **Visualisasi:** Plotly, Matplotlib
        - **Export:** python-docx, ReportLab
        
        **Pengembang:**
        Aplikasi ini dikembangkan untuk mendukung penelitian dan praktikum nanomaterial
        di laboratorium akademik dan industri.
        
        **Lisensi:**
        Open Source - MIT License
        
        **Kontak:**
        - Email: support@labpsanano.com
        - GitHub: github.com/labpsanano
        - Website: labpsanano.streamlit.app
        
        **Versi:** 2.0.0
        **Update Terakhir:** Oktober 2024
        """)

# Footer
st.markdown("---")
footer_cols = st.columns([2, 1, 2])
with footer_cols[0]:
    st.caption("🔬 **Lab PSA Nano** - Aplikasi Catatan & Kalkulator PSA Nanomaterial")
with footer_cols[1]:
    st.caption("📧 support@labpsanano.com")
with footer_cols[2]:
    st.caption("© 2024 Laboratorium Nanomaterial Indonesia")

# Auto-save reminder
if st.session_state.catatan_list or st.session_state.psa_results:
    st.toast("💡 Ingat untuk mengekspor data penting Anda!", icon="⚠️")
