import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime
import tempfile
import os
from utils.word_export import create_word_document
from utils.pdf_export import create_psa_pdf

# Konfigurasi halaman
st.set_page_config(
    page_title="Catatan Praktik & Kalkulator PSA Nanomaterial",
    page_icon="🔬",
    layout="wide"
)

# Inisialisasi session state
if 'catatan_list' not in st.session_state:
    st.session_state.catatan_list = []

if 'psa_results' not in st.session_state:
    st.session_state.psa_results = []

# Fungsi untuk navigasi
def set_page(page):
    st.session_state.current_page = page

if 'current_page' not in st.session_state:
    st.session_state.current_page = "beranda"

# Sidebar navigasi
with st.sidebar:
    st.image("https://img.icons8.com/color/96/000000/test-tube.png", width=80)
    st.title("🔬 NanoLab PSA")
    st.write("Aplikasi Catatan Praktik & Kalkulator PSA Nanomaterial")
    
    st.divider()
    
    # Menu navigasi
    menu_options = {
        "🏠 Beranda": "beranda",
        "📝 Catatan Praktik": "catatan",
        "🧮 Kalkulator PSA": "kalkulator",
        "📊 Hasil PSA": "hasil",
        "📁 Ekspor Data": "ekspor"
    }
    
    for label, page in menu_options.items():
        if st.button(label, use_container_width=True, key=f"btn_{page}"):
            set_page(page)
    
    st.divider()
    
    # Informasi
    st.caption("**Informasi Aplikasi**")
    st.caption("""
    Aplikasi ini membantu praktikan dalam:
    1. Mencatat hasil praktik nanomaterial
    2. Mengkalkulasi hasil PSA (Particle Size Analysis)
    3. Menyimpan catatan dalam format Word
    4. Menyimpan hasil kalkulasi dalam format PDF
    """)

# Halaman Beranda
if st.session_state.current_page == "beranda":
    st.title("🏠 Selamat Datang di NanoLab PSA")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        ### 🔬 Tentang Aplikasi
        Aplikasi ini dirancang untuk membantu praktikan nanomaterial dalam:
        
        - **Mencatat hasil praktik** dengan detail lengkap
        - **Mengkalkulasi hasil PSA** dari data PDI, %vol, dan diameter
        - **Menganalisis distribusi ukuran partikel**
        - **Menyimpan data** dalam format Word dan PDF
        
        ### 📋 Fitur Utama
        1. **Sistem Pencatatan Praktik**
           - Input data eksperimen
           - Upload gambar hasil sintesis
           - Kategorisasi sampel
        
        2. **Kalkulator PSA**
           - Input data distribusi ukuran
           - Analisis statistik
           - Visualisasi distribusi
        
        3. **Sistem Ekspor**
           - Export catatan ke Word (.docx)
           - Export hasil PSA ke PDF
        """)
    
    with col2:
        st.image("https://img.icons8.com/color/300/000000/laboratory.png", width=300)
        
        st.info("""
        **Panduan Cepat:**
        1. Gunakan menu sidebar untuk navigasi
        2. Mulai dengan membuat catatan praktik
        3. Hitung hasil PSA menggunakan kalkulator
        4. Simpan hasil dalam format yang diinginkan
        """)

# Halaman Catatan Praktik
elif st.session_state.current_page == "catatan":
    st.title("📝 Catatan Praktik Nanomaterial")
    
    tab1, tab2 = st.tabs(["📄 Buat Catatan Baru", "📚 Lihat Catatan"])
    
    with tab1:
        with st.form("form_catatan"):
            col1, col2 = st.columns(2)
            
            with col1:
                judul = st.text_input("Judul Praktik*")
                nama_praktikan = st.text_input("Nama Praktikan*")
                tanggal = st.date_input("Tanggal Praktik*")
                jenis_nanomaterial = st.selectbox(
                    "Jenis Nanomaterial*",
                    ["TiO₂", "SiO₂", "ZnO", "Ag", "Au", "Fe₃O₄", "Lainnya"]
                )
                if jenis_nanomaterial == "Lainnya":
                    jenis_nanomaterial = st.text_input("Sebutkan jenis nanomaterial")
            
            with col2:
                metode_sintesis = st.selectbox(
                    "Metode Sintesis*",
                    ["Sol-Gel", "Hidrotermal", "Sonokimia", "Mekanokimia", "Lainnya"]
                )
                if metode_sintesis == "Lainnya":
                    metode_sintesis = st.text_input("Sebutkan metode sintesis")
                
                suhu = st.number_input("Suhu Sintesis (°C)", min_value=0, max_value=1000, value=25)
                waktu = st.number_input("Waktu Sintesis (jam)", min_value=0.0, max_value=100.0, value=1.0)
            
            st.subheader("Prosedur Praktik")
            prosedur = st.text_area(
                "Tuliskan prosedur yang dilakukan*",
                height=150,
                placeholder="1. Siapkan bahan...\n2. Campurkan...\n3. Panaskan pada suhu..."
            )
            
            st.subheader("Hasil Pengamatan")
            hasil_pengamatan = st.text_area(
                "Tuliskan hasil pengamatan*",
                height=150,
                placeholder="Warna, bentuk, karakteristik nanomaterial..."
            )
            
            st.subheader("Upload Gambar Hasil")
            uploaded_images = st.file_uploader(
                "Upload gambar hasil sintesis (format: JPG, PNG)",
                type=['jpg', 'jpeg', 'png'],
                accept_multiple_files=True
            )
            
            st.subheader("Parameter Tambahan")
            col_params1, col_params2 = st.columns(2)
            
            with col_params1:
                ph = st.slider("pH Larutan", 0.0, 14.0, 7.0)
                konsentrasi = st.number_input("Konsentrasi (mg/mL)", min_value=0.0, value=1.0)
            
            with col_params2:
                pelarut = st.text_input("Jenis Pelarut", "Air/Aqua")
                catatan_tambahan = st.text_area("Catatan Tambahan", height=100)
            
            submitted = st.form_submit_button("💾 Simpan Catatan")
            
            if submitted:
                if judul and nama_praktikan and prosedur and hasil_pengamatan:
                    # Simpan gambar ke temporary file
                    image_paths = []
                    if uploaded_images:
                        for img in uploaded_images:
                            with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
                                tmp.write(img.getvalue())
                                image_paths.append(tmp.name)
                    
                    catatan = {
                        'id': len(st.session_state.catatan_list) + 1,
                        'judul': judul,
                        'nama_praktikan': nama_praktikan,
                        'tanggal': str(tanggal),
                        'jenis_nanomaterial': jenis_nanomaterial,
                        'metode_sintesis': metode_sintesis,
                        'suhu': suhu,
                        'waktu': waktu,
                        'prosedur': prosedur,
                        'hasil_pengamatan': hasil_pengamatan,
                        'ph': ph,
                        'konsentrasi': konsentrasi,
                        'pelarut': pelarut,
                        'catatan_tambahan': catatan_tambahan,
                        'image_paths': image_paths,
                        'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    }
                    
                    st.session_state.catatan_list.append(catatan)
                    st.success(f"Catatan '{judul}' berhasil disimpan!")
                    st.balloons()
                else:
                    st.error("Harap isi semua field yang wajib (*)!")

    with tab2:
        if st.session_state.catatan_list:
            st.subheader("Daftar Catatan Praktik")
            
            for idx, catatan in enumerate(st.session_state.catatan_list):
                with st.expander(f"📋 {catatan['judul']} - {catatan['tanggal']}"):
                    col_info, col_action = st.columns([3, 1])
                    
                    with col_info:
                        st.write(f"**Praktikan:** {catatan['nama_praktikan']}")
                        st.write(f"**Nanomaterial:** {catatan['jenis_nanomaterial']}")
                        st.write(f"**Metode:** {catatan['metode_sintesis']}")
                        st.write(f"**Suhu:** {catatan['suhu']}°C")
                        st.write(f"**Waktu:** {catatan['waktu']} jam")
                        
                        if catatan['image_paths']:
                            st.write("**Gambar Hasil:**")
                            cols = st.columns(min(3, len(catatan['image_paths'])))
                            for img_idx, img_path in enumerate(catatan['image_paths'][:3]):
                                with cols[img_idx]:
                                    try:
                                        st.image(img_path, use_column_width=True)
                                    except:
                                        st.warning("Gambar tidak dapat ditampilkan")
                    
                    with col_action:
                        if st.button("📥 Ekspor ke Word", key=f"export_{idx}"):
                            try:
                                doc_path = create_word_document(catatan)
                                with open(doc_path, 'rb') as f:
                                    st.download_button(
                                        label="⬇️ Download Word",
                                        data=f,
                                        file_name=f"catatan_{catatan['judul']}_{catatan['tanggal']}.docx",
                                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                                    )
                                os.unlink(doc_path)  # Hapus temporary file
                            except Exception as e:
                                st.error(f"Error: {str(e)}")
                        
                        if st.button("🗑️ Hapus", key=f"delete_{idx}"):
                            st.session_state.catatan_list.pop(idx)
                            st.rerun()
        else:
            st.info("Belum ada catatan praktik. Silahkan buat catatan baru.")

# Halaman Kalkulator PSA
elif st.session_state.current_page == "kalkulator":
    st.title("🧮 Kalkulator PSA Nanomaterial")
    
    st.markdown("""
    ### Panduan Penggunaan:
    1. Masukkan jumlah titik data (minimum 3 titik)
    2. Input data Diameter (nm), % Volume, dan PDI untuk setiap titik
    3. Tekan tombol **Hitung Hasil PSA** untuk analisis
    4. Hasil akan ditampilkan dan dapat diekspor ke PDF
    """)
    
    # Input jumlah data
    col1, col2 = st.columns([2, 1])
    
    with col1:
        num_points = st.number_input(
            "Jumlah Titik Data Distribusi",
            min_value=3,
            max_value=50,
            value=5,
            step=1,
            help="Masukkan jumlah titik data untuk distribusi ukuran partikel"
        )
    
    with col2:
        st.write("###")
        if st.button("🔄 Generate Tabel Input", use_container_width=True):
            st.session_state.generate_table = True
    
    # Input data dalam bentuk tabel
    st.subheader("📊 Input Data Distribusi")
    
    # Inisialisasi dataframe
    if 'psa_data' not in st.session_state or st.session_state.get('generate_table', False):
        data = {
            'Diameter (nm)': np.round(np.linspace(10, 100, num_points), 2),
            '% Volume': np.round(np.random.dirichlet(np.ones(num_points)) * 100, 2),
            'PDI': np.round(np.random.uniform(0.05, 0.3, num_points), 3)
        }
        st.session_state.psa_data = pd.DataFrame(data)
        st.session_state.generate_table = False
    
    # Tabel input data yang bisa diedit
    edited_df = st.data_editor(
        st.session_state.psa_data,
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "Diameter (nm)": st.column_config.NumberColumn(
                format="%.2f",
                min_value=0.1,
                max_value=10000.0
            ),
            "% Volume": st.column_config.NumberColumn(
                format="%.2f",
                min_value=0.0,
                max_value=100.0
            ),
            "PDI": st.column_config.NumberColumn(
                format="%.3f",
                min_value=0.001,
                max_value=1.0
            )
        }
    )
    
    # Tombol kalkulasi
    if st.button("🧮 Hitung Hasil PSA", type="primary", use_container_width=True):
        with st.spinner("Menghitung hasil PSA..."):
            try:
                # Validasi data
                if edited_df.empty:
                    st.error("Data tidak boleh kosong!")
                elif edited_df['% Volume'].sum() == 0:
                    st.error("Total % Volume tidak boleh nol!")
                else:
                    # Normalisasi % Volume
                    edited_df['% Volume Normalized'] = (edited_df['% Volume'] / edited_df['% Volume'].sum() * 100)
                    
                    # Hitung statistik
                    diameter_weighted = np.average(
                        edited_df['Diameter (nm)'],
                        weights=edited_df['% Volume']
                    )
                    
                    # Hitung rata-rata PDI
                    pdi_weighted = np.average(
                        edited_df['PDI'],
                        weights=edited_df['% Volume']
                    )
                    
                    # Hitung standard deviation
                    variance = np.average(
                        (edited_df['Diameter (nm)'] - diameter_weighted) ** 2,
                        weights=edited_df['% Volume']
                    )
                    std_dev = np.sqrt(variance)
                    
                    # Hitung polydispersity index
                    pdi_calculated = variance / (diameter_weighted ** 2)
                    
                    # Klasifikasi berdasarkan PDI
                    if pdi_calculated < 0.1:
                        klasifikasi = "Monodispersi (Sangat Baik)"
                        warna = "🟢"
                    elif pdi_calculated < 0.2:
                        klasifikasi = "Hampir Monodispersi (Baik)"
                        warna = "🟡"
                    elif pdi_calculated < 0.3:
                        klasifikasi = "Polydispersi Sedang"
                        warna = "🟠"
                    else:
                        klasifikasi = "Polydispersi Tinggi (Perlu Optimasi)"
                        warna = "🔴"
                    
                    # Simpan hasil
                    hasil_psa = {
                        'dataframe': edited_df.copy(),
                        'diameter_rerata': diameter_weighted,
                        'pdi_rerata': pdi_weighted,
                        'pdi_terhitung': pdi_calculated,
                        'std_dev': std_dev,
                        'variance': variance,
                        'klasifikasi': klasifikasi,
                        'warna': warna,
                        'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    }
                    
                    st.session_state.psa_results.append(hasil_psa)
                    st.success("Perhitungan PSA berhasil!")
                    
                    # Tampilkan hasil
                    st.subheader("📈 Hasil Kalkulasi PSA")
                    
                    col_result1, col_result2, col_result3 = st.columns(3)
                    
                    with col_result1:
                        st.metric(
                            label="Diameter Rata-rata (nm)",
                            value=f"{diameter_weighted:.2f}",
                            delta=f"± {std_dev:.2f} nm"
                        )
                    
                    with col_result2:
                        st.metric(
                            label="PDI Terhitung",
                            value=f"{pdi_calculated:.3f}",
                            delta=klasifikasi.split('(')[-1].replace(')', '')
                        )
                    
                    with col_result3:
                        st.metric(
                            label="PDI Rata-rata",
                            value=f"{pdi_weighted:.3f}"
                        )
                    
                    # Visualisasi
                    tab_viz1, tab_viz2 = st.tabs(["📊 Distribusi Ukuran", "📈 Grafik PDI"])
                    
                    with tab_viz1:
                        fig1 = go.Figure()
                        fig1.add_trace(go.Bar(
                            x=edited_df['Diameter (nm)'],
                            y=edited_df['% Volume'],
                            name='% Volume',
                            marker_color='royalblue'
                        ))
                        
                        fig1.add_vline(
                            x=diameter_weighted,
                            line_dash="dash",
                            line_color="red",
                            annotation_text=f"Rata-rata: {diameter_weighted:.2f} nm"
                        )
                        
                        fig1.update_layout(
                            title='Distribusi Ukuran Partikel',
                            xaxis_title='Diameter (nm)',
                            yaxis_title='% Volume',
                            template='plotly_white'
                        )
                        st.plotly_chart(fig1, use_container_width=True)
                    
                    with tab_viz2:
                        fig2 = px.scatter(
                            edited_df,
                            x='Diameter (nm)',
                            y='PDI',
                            size='% Volume',
                            color='% Volume',
                            hover_data=['% Volume'],
                            title='Hubungan Diameter vs PDI'
                        )
                        
                        fig2.update_traces(
                            marker=dict(line=dict(width=2, color='DarkSlateGrey'))
                        )
                        
                        fig2.update_layout(
                            xaxis_title='Diameter (nm)',
                            yaxis_title='PDI',
                            template='plotly_white'
                        )
                        st.plotly_chart(fig2, use_container_width=True)
                    
                    # Detail hasil
                    with st.expander("🔍 Detail Hasil Perhitungan"):
                        col_detail1, col_detail2 = st.columns(2)
                        
                        with col_detail1:
                            st.write("**Statistik Deskriptif:**")
                            stats_df = edited_df['Diameter (nm)'].describe()
                            st.dataframe(stats_df)
                        
                        with col_detail2:
                            st.write("**Parameter Kualitas:**")
                            st.write(f"- Standard Deviation: {std_dev:.2f} nm")
                            st.write(f"- Variance: {variance:.2f}")
                            st.write(f"- Range: {edited_df['Diameter (nm)'].min():.2f} - {edited_df['Diameter (nm)'].max():.2f} nm")
                            st.write(f"- Koefisien Variasi: {(std_dev/diameter_weighted*100):.2f}%")
                            st.write(f"- **Klasifikasi: {warna} {klasifikasi}**")
                            
                            # Rekomendasi berdasarkan hasil
                            st.write("**Rekomendasi:**")
                            if pdi_calculated < 0.1:
                                st.success("Kualitas nanomaterial sangat baik. Lanjutkan metode sintesis ini.")
                            elif pdi_calculated < 0.2:
                                st.info("Kualitas baik. Pertimbangkan optimasi kecil untuk meningkatkan monodispersitas.")
                            elif pdi_calculated < 0.3:
                                st.warning("Perlu optimasi proses sintesis untuk mengurangi polydispersitas.")
                            else:
                                st.error("Perlu evaluasi ulang metode sintesis. Pertimbangkan untuk mengubah parameter atau metode.")
            
            except Exception as e:
                st.error(f"Terjadi kesalahan dalam perhitungan: {str(e)}")

# Halaman Hasil PSA
elif st.session_state.current_page == "hasil":
    st.title("📊 Hasil PSA Tersimpan")
    
    if st.session_state.psa_results:
        st.subheader(f"Total {len(st.session_state.psa_results)} Hasil PSA")
        
        for idx, hasil in enumerate(st.session_state.psa_results):
            with st.expander(f"📋 Hasil PSA #{idx+1} - {hasil['timestamp']}"):
                col_res1, col_res2 = st.columns([2, 1])
                
                with col_res1:
                    st.write(f"**Diameter Rata-rata:** {hasil['diameter_rerata']:.2f} nm")
                    st.write(f"**PDI Terhitung:** {hasil['pdi_terhitung']:.3f}")
                    st.write(f"**Standard Deviation:** {hasil['std_dev']:.2f} nm")
                    st.write(f"**Klasifikasi:** {hasil['warna']} {hasil['klasifikasi']}")
                    
                    # Tampilkan data
                    st.write("**Data Distribusi:**")
                    st.dataframe(hasil['dataframe'][['Diameter (nm)', '% Volume', 'PDI']])
                
                with col_res2:
                    # Tombol ekspor
                    if st.button("📥 Ekspor ke PDF", key=f"export_pdf_{idx}", use_container_width=True):
                        try:
                            pdf_path = create_psa_pdf(hasil, idx+1)
                            with open(pdf_path, 'rb') as f:
                                st.download_button(
                                    label="⬇️ Download PDF",
                                    data=f,
                                    file_name=f"hasil_psa_{idx+1}_{hasil['timestamp'].replace(':', '-')}.pdf",
                                    mime="application/pdf",
                                    key=f"download_{idx}"
                                )
                            os.unlink(pdf_path)  # Hapus temporary file
                        except Exception as e:
                            st.error(f"Error: {str(e)}")
                    
                    # Tombol hapus
                    if st.button("🗑️ Hapus", key=f"delete_res_{idx}", use_container_width=True):
                        st.session_state.psa_results.pop(idx)
                        st.rerun()
    else:
        st.info("Belum ada hasil PSA. Silahkan gunakan kalkulator PSA terlebih dahulu.")

# Halaman Ekspor Data
elif st.session_state.current_page == "ekspor":
    st.title("📁 Ekspor Data")
    
    tab_export1, tab_export2 = st.tabs(["📝 Ekspor Catatan", "📊 Ekspor Hasil PSA"])
    
    with tab_export1:
        st.subheader("Ekspor Catatan Praktik ke Word")
        
        if st.session_state.catatan_list:
            # Pilih catatan untuk diekspor
            catatan_options = [f"{c['id']}: {c['judul']} - {c['tanggal']}" for c in st.session_state.catatan_list]
            selected_note = st.selectbox("Pilih Catatan untuk Diekspor", catatan_options)
            
            if selected_note:
                note_id = int(selected_note.split(":")[0]) - 1
                catatan = st.session_state.catatan_list[note_id]
                
                # Preview catatan
                with st.expander("👁️ Preview Catatan"):
                    col_preview1, col_preview2 = st.columns(2)
                    
                    with col_preview1:
                        st.write("**Informasi Praktik:**")
                        st.write(f"- Judul: {catatan['judul']}")
                        st.write(f"- Praktikan: {catatan['nama_praktikan']}")
                        st.write(f"- Tanggal: {catatan['tanggal']}")
                        st.write(f"- Nanomaterial: {catatan['jenis_nanomaterial']}")
                        st.write(f"- Metode: {catatan['metode_sintesis']}")
                    
                    with col_preview2:
                        st.write("**Parameter:**")
                        st.write(f"- Suhu: {catatan['suhu']}°C")
                        st.write(f"- Waktu: {catatan['waktu']} jam")
                        st.write(f"- pH: {catatan['ph']}")
                        st.write(f"- Konsentrasi: {catatan['konsentrasi']} mg/mL")
                        st.write(f"- Pelarut: {catatan['pelarut']}")
                
                # Tombol ekspor
                if st.button("📥 Ekspor ke Word Document", type="primary"):
                    try:
                        doc_path = create_word_document(catatan)
                        with open(doc_path, 'rb') as f:
                            st.download_button(
                                label="⬇️ Download File Word",
                                data=f,
                                file_name=f"catatan_{catatan['judul']}_{catatan['tanggal']}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                        os.unlink(doc_path)
                    except Exception as e:
                        st.error(f"Error dalam ekspor: {str(e)}")
        
        else:
            st.info("Belum ada catatan untuk diekspor")
    
    with tab_export2:
        st.subheader("Ekspor Hasil PSA ke PDF")
        
        if st.session_state.psa_results:
            # Pilih hasil PSA untuk diekspor
            psa_options = [f"Hasil #{i+1} - D={r['diameter_rerata']:.2f}nm, PDI={r['pdi_terhitung']:.3f}" 
                          for i, r in enumerate(st.session_state.psa_results)]
            selected_psa = st.selectbox("Pilih Hasil PSA untuk Diekspor", psa_options)
            
            if selected_psa:
                psa_idx = int(selected_psa.split("#")[1].split(" ")[0]) - 1
                hasil = st.session_state.psa_results[psa_idx]
                
                # Preview hasil
                with st.expander("👁️ Preview Hasil PSA"):
                    col_prev1, col_prev2 = st.columns(2)
                    
                    with col_prev1:
                        st.write("**Statistik Utama:**")
                        st.write(f"- Diameter Rata-rata: {hasil['diameter_rerata']:.2f} nm")
                        st.write(f"- PDI Terhitung: {hasil['pdi_terhitung']:.3f}")
                        st.write(f"- Standard Deviation: {hasil['std_dev']:.2f} nm")
                        st.write(f"- Klasifikasi: {hasil['klasifikasi']}")
                    
                    with col_prev2:
                        st.write("**Data Sample:**")
                        st.dataframe(hasil['dataframe'].head(3))
                
                # Tombol ekspor
                if st.button("📥 Ekspor ke PDF Document", type="primary", key="export_pdf_btn"):
                    try:
                        pdf_path = create_psa_pdf(hasil, psa_idx+1)
                        with open(pdf_path, 'rb') as f:
                            st.download_button(
                                label="⬇️ Download File PDF",
                                data=f,
                                file_name=f"hasil_psa_{psa_idx+1}_{hasil['timestamp'].replace(':', '-')}.pdf",
                                mime="application/pdf"
                            )
                        os.unlink(pdf_path)
                    except Exception as e:
                        st.error(f"Error dalam ekspor: {str(e)}")
        
        else:
            st.info("Belum ada hasil PSA untuk diekspor")

# Footer
st.divider()
footer_col1, footer_col2, footer_col3 = st.columns(3)
with footer_col1:
    st.caption("🔬 **NanoLab PSA** v1.0")
with footer_col2:
    st.caption("📧 Contact: nanolab@example.com")
with footer_col3:
    st.caption("© 2024 Laboratorium Nanomaterial")
