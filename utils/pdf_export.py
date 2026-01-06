from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.graphics.charts.lineplots import LinePlot
from reportlab.graphics.shapes import Drawing
import tempfile
import os

def create_psa_pdf(hasil_psa, result_id):
    """
    Membuat PDF dari hasil kalkulasi PSA
    """
    # Setup document
    temp_dir = tempfile.gettempdir()
    filename = f"hasil_psa_{result_id}_{hasil_psa['timestamp'].replace(':', '-')}.pdf"
    filepath = os.path.join(temp_dir, filename)
    
    doc = SimpleDocTemplate(filepath, pagesize=letter)
    story = []
    styles = getSampleStyleSheet()
    
    # Judul
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=16,
        alignment=1,  # Center
        spaceAfter=30
    )
    
    story.append(Paragraph("LAPORAN HASIL PSA NANOMATERIAL", title_style))
    story.append(Paragraph(f"ID Hasil: #{result_id}", styles["Normal"]))
    story.append(Paragraph(f"Tanggal Analisis: {hasil_psa['timestamp']}", styles["Normal"]))
    story.append(Spacer(1, 20))
    
    # Ringkasan Hasil
    story.append(Paragraph("RINGKASAN HASIL", styles["Heading2"]))
    
    summary_data = [
        ["Parameter", "Nilai", "Keterangan"],
        ["Diameter Rata-rata", f"{hasil_psa['diameter_rerata']:.2f} nm", "Weighted average"],
        ["PDI Terhitung", f"{hasil_psa['pdi_terhitung']:.3f}", hasil_psa['klasifikasi']],
        ["Standard Deviation", f"{hasil_psa['std_dev']:.2f} nm", "Distribusi ukuran"],
        ["Variance", f"{hasil_psa['variance']:.2f}", "Variansi distribusi"],
        ["PDI Rata-rata Input", f"{hasil_psa['pdi_rerata']:.3f}", "Rata-rata PDI input"]
    ]
    
    summary_table = Table(summary_data, colWidths=[2*inch, 1.5*inch, 2*inch])
    summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    
    story.append(summary_table)
    story.append(Spacer(1, 30))
    
    # Data Distribusi
    story.append(Paragraph("DATA DISTRIBUSI UKURAN", styles["Heading2"]))
    
    # Siapkan data tabel
    df = hasil_psa['dataframe']
    table_data = [["Diameter (nm)", "% Volume", "PDI"]]
    
    for _, row in df.iterrows():
        table_data.append([
            f"{row['Diameter (nm)']:.2f}",
            f"{row['% Volume']:.2f}",
            f"{row['PDI']:.3f}"
        ])
    
    # Tambahkan total
    table_data.append([
        "TOTAL/RATA-RATA",
        f"{df['% Volume'].sum():.2f}",
        f"{df['PDI'].mean():.3f}"
    ])
    
    dist_table = Table(table_data, colWidths=[1.5*inch, 1.5*inch, 1.5*inch])
    dist_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 11),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -2), colors.whitesmoke),
        ('BACKGROUND', (0, -1), (-1, -1), colors.lightgrey),
        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey)
    ]))
    
    story.append(dist_table)
    story.append(Spacer(1, 30))
    
    # Interpretasi Hasil
    story.append(Paragraph("INTERPRETASI HASIL", styles["Heading2"]))
    
    interpretation = f"""
    Hasil analisis PSA menunjukkan nanomaterial dengan diameter rata-rata {hasil_psa['diameter_rerata']:.2f} nm 
    dan PDI sebesar {hasil_psa['pdi_terhitung']:.3f}. Berdasarkan nilai PDI, sampel diklasifikasikan sebagai:
    
    {hasil_psa['klasifikasi']}
    
    Standard deviation sebesar {hasil_psa['std_dev']:.2f} nm menunjukkan tingkat dispersi ukuran partikel. 
    Koefisien variasi sebesar {(hasil_psa['std_dev']/hasil_psa['diameter_rerata']*100):.2f}% 
    mengindikasikan {'distribusi yang homogen' if hasil_psa['pdi_terhitung'] < 0.2 else 'distribusi yang heterogen'}.
    """
    
    story.append(Paragraph(interpretation, styles["Normal"]))
    story.append(Spacer(1, 30))
    
    # Rekomendasi
    story.append(Paragraph("REKOMENDASI", styles["Heading2"]))
    
    if hasil_psa['pdi_terhitung'] < 0.1:
        recommendation = """
        • Kualitas nanomaterial sangat baik dengan distribusi ukuran yang seragam
        • Lanjutkan metode sintesis dengan parameter yang sama
        • Pertimbangkan untuk publikasi hasil
        """
    elif hasil_psa['pdi_terhitung'] < 0.2:
        recommendation = """
        • Kualitas baik, dapat diterima untuk kebanyakan aplikasi
        • Optimasi kecil dapat meningkatkan monodispersitas
        • Pertimbangkan variasi waktu atau suhu sintesis
        """
    elif hasil_psa['pdi_terhitung'] < 0.3:
        recommendation = """
        • Perlu optimasi proses sintesis
        • Evaluasi parameter seperti pH, konsentrasi, atau metode pencampuran
        • Pertimbangkan penggunaan surfaktan atau stabilizer
        """
    else:
        recommendation = """
        • Perlu evaluasi ulang metode sintesis
        • Pertimbangkan untuk mengubah metode sintesis
        • Optimasi parameter utama (suhu, waktu, konsentrasi)
        • Lakukan analisis penyebab polydispersitas tinggi
        """
    
    story.append(Paragraph(recommendation, styles["Normal"]))
    
    # Catatan
    story.append(Spacer(1, 30))
    story.append(Paragraph("CATATAN:", styles["Heading3"]))
    story.append(Paragraph("""
    1. PDI (Polydispersity Index) mengindikasikan keseragaman ukuran partikel
    2. PDI < 0.1: Monodispersi (ideal)
    3. PDI 0.1-0.2: Hampir monodispersi (baik)
    4. PDI > 0.3: Polydispersi (perlu optimasi)
    """, styles["Normal"]))
    
    # Build PDF
    doc.build(story)
    
    return filepath
