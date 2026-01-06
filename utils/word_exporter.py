from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_PARAGRAPH_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
import os
from datetime import datetime
import tempfile

def create_word_note(catatan):
    """
    Membuat dokumen Word dari catatan praktik
    """
    # Buat dokumen baru
    doc = Document()
    
    # Set margin
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
    
    # Header dengan judul
    title = doc.add_heading('LAPORAN PRAKTIKUM NANOMATERIAL', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.runs[0].font.color.rgb = RGBColor(46, 134, 171)
    title.runs[0].font.size = Pt(16)
    title.runs[0].bold = True
    
    # Sub judul
    subtitle = doc.add_heading(catatan['judul'], 1)
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.runs[0].font.size = Pt(14)
    
    doc.add_paragraph()  # Spasi
    
    # Informasi metadata
    doc.add_heading('INFORMASI UMUM', level=2)
    
    info_data = [
        ("Nama Praktikan", catatan['nama_praktikan']),
        ("Tanggal Praktikum", catatan['tanggal']),
        ("Institusi/Laboratorium", catatan.get('institusi', '-')),
        ("Kelompok/Shift", catatan.get('kelompok', '-')),
        ("Supervisor", catatan.get('supervisor', '-')),
        ("Waktu Penyimpanan", catatan['timestamp'])
    ]
    
    for label, value in info_data:
        p = doc.add_paragraph()
        p.add_run(f"{label}: ").bold = True
        p.add_run(value)
    
    doc.add_paragraph()  # Spasi
    
    # Spesifikasi nanomaterial
    doc.add_heading('SPESIFIKASI NANOMATERIAL', level=2)
    
    spec_data = [
        ("Jenis Nanomaterial", catatan['jenis_nanomaterial']),
        ("Metode Sintesis", catatan['metode_sintesis']),
        ("Pelarut", catatan['pelarut'])
    ]
    
    for label, value in spec_data:
        p = doc.add_paragraph()
        p.add_run(f"{label}: ").bold = True
        p.add_run(value)
    
    doc.add_paragraph()  # Spasi
    
    # Parameter sintesis
    doc.add_heading('PARAMETER SINTESIS', level=2)
    
    param_data = [
        ("Suhu Sintesis", f"{catatan['suhu']} °C"),
        ("Waktu Sintesis", f"{catatan['waktu']} jam"),
        ("Tekanan", f"{catatan.get('tekanan', '-')} atm"),
        ("pH Larutan", f"{catatan['ph']}"),
        ("Konsentrasi", f"{catatan['konsentrasi']} mg/mL")
    ]
    
    for label, value in param_data:
        p = doc.add_paragraph()
        p.add_run(f"{label}: ").bold = True
        p.add_run(value)
    
    doc.add_page_break()
    
    # Prosedur
    doc.add_heading('PROSEDUR PRAKTIKUM', level=2)
    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    p.add_run(catatan['prosedur'])
    
    doc.add_paragraph()  # Spasi
    
    # Hasil pengamatan
    doc.add_heading('HASIL PENGAMATAN', level=2)
    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    p.add_run(catatan['hasil_pengamatan'])
    
    # Gambar jika ada
    if catatan.get('image_path') and os.path.exists(catatan['image_path']):
        doc.add_paragraph()  # Spasi
        doc.add_heading('DOKUMENTASI HASIL', level=2)
        try:
            doc.add_picture(catatan['image_path'], width=Inches(4))
            doc.add_paragraph("Gambar hasil sintesis nanomaterial")
        except:
            pass
    
    # Footer
    doc.add_paragraph()  # Spasi
    footer = doc.add_paragraph()
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer_run = footer.add_run(f"Dokumen dibuat dengan NaNote • {datetime.now().strftime('%d %B %Y %H:%M:%S')}")
    footer_run.font.size = Pt(9)
    footer_run.font.color.rgb = RGBColor(128, 128, 128)
    footer_run.italic = True
    
    # Simpan file
    temp_dir = tempfile.gettempdir()
    safe_title = "".join(c for c in catatan['judul'] if c.isalnum() or c in (' ', '-', '_')).rstrip()
    filename = f"NaNote_{safe_title}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    filepath = os.path.join(temp_dir, filename)
    doc.save(filepath)
    
    return filepath
