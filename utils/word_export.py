from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
import os
from datetime import datetime

def create_word_document(catatan):
    """
    Membuat dokumen Word dari catatan praktik
    """
    doc = Document()
    
    # Tambahkan judul
    title = doc.add_heading('CATATAN PRAKTIK NANOMATERIAL', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Informasi metadata
    doc.add_paragraph(f"Judul Praktik: {catatan['judul']}")
    doc.add_paragraph(f"Nama Praktikan: {catatan['nama_praktikan']}")
    doc.add_paragraph(f"Tanggal: {catatan['tanggal']}")
    doc.add_paragraph(f"Jenis Nanomaterial: {catatan['jenis_nanomaterial']}")
    doc.add_paragraph(f"Metode Sintesis: {catatan['metode_sintesis']}")
    
    doc.add_paragraph()  # Spasi
    
    # Prosedur
    doc.add_heading('PROSEDUR PRAKTIK', level=1)
    for line in catatan['prosedur'].split('\n'):
        if line.strip():
            doc.add_paragraph(line.strip())
    
    doc.add_paragraph()  # Spasi
    
    # Hasil Pengamatan
    doc.add_heading('HASIL PENGAMATAN', level=1)
    for line in catatan['hasil_pengamatan'].split('\n'):
        if line.strip():
            doc.add_paragraph(line.strip())
    
    doc.add_paragraph()  # Spasi
    
    # Parameter
    doc.add_heading('PARAMETER SINTESIS', level=1)
    doc.add_paragraph(f"Suhu: {catatan['suhu']} °C")
    doc.add_paragraph(f"Waktu: {catatan['waktu']} jam")
    doc.add_paragraph(f"pH: {catatan['ph']}")
    doc.add_paragraph(f"Konsentrasi: {catatan['konsentrasi']} mg/mL")
    doc.add_paragraph(f"Pelarut: {catatan['pelarut']}")
    
    # Catatan Tambahan
    if catatan['catatan_tambahan']:
        doc.add_heading('CATATAN TAMBAHAN', level=1)
        doc.add_paragraph(catatan['catatan_tambahan'])
    
    # Footer
    doc.add_paragraph()  # Spasi
    doc.add_paragraph(f"Dokumen dibuat pada: {catatan['timestamp']}")
    
    # Simpan ke temporary file
    temp_dir = tempfile.gettempdir()
    filename = f"catatan_{catatan['judul']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    filepath = os.path.join(temp_dir, filename)
    doc.save(filepath)
    
    return filepath
