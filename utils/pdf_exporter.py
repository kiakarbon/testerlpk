from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, cm
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.pdfgen import canvas
from reportlab.graphics.charts.lineplots import LinePlot
from reportlab.graphics.shapes import Drawing, String
from reportlab.graphics.charts.barcharts import VerticalBarChart
import tempfile
import os
from datetime import datetime
import matplotlib.pyplot as plt
import numpy as np
from io import BytesIO

def create_psa_pdf(hasil_psa, result_id):
    """
    Membuat PDF profesional untuk hasil PSA
    """
    # Setup document
    temp_dir = tempfile.gettempdir()
    filename = f"Laporan_PSA_{result_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    filepath = os.path.join(temp_dir, filename)
    
    doc = SimpleDocTemplate(
        filepath,
        pagesize=A4,
        topMargin=1*cm,
        bottomMargin=1*cm,
        leftMargin=1.5*cm,
        rightMargin=1.5*cm
    )
    
    story = []
    styles = getSampleStyleSheet()
    
    # Custom styles
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=18,
        alignment=TA_CENTER,
        textColor=colors.HexColor('#2E86AB'),
        spaceAfter=20
    )
    
    heading_style = ParagraphStyle(
        'CustomHeading',
        parent=styles['Heading2'],
        fontSize=14,
        textColor=colors.HexColor('#2E86AB'),
        spaceAfter=10,
        spaceBefore=20
    )
    
    normal_style = ParagraphStyle(
        'CustomNormal',
        parent=styles['Normal'],
        fontSize=10,
        leading=14
    )
    
    # Header dengan logo dan judul
    header_table = Table([
        [Paragraph("LABORATORIUM NANOMATERIAL", title_style)],
        [Paragraph("LAPORAN ANALISIS DISTRIBUSI UKURAN PARTIKEL", heading_style)]
    ], colWidths=[16*cm])
    
    header_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    
    story.append(header_table)
    story.append(Spacer(1, 0.5*cm))
    
    # Informasi laporan
    info_data = [
        ["ID Laporan", f"PSA-{result_id:03d}"],
        ["Tanggal Analisis", hasil_psa['timestamp']],
        ["Jumlah Data", str(hasil_psa['total_points']) + " titik"],
        ["Dienerate oleh", "Lab PSA Nano v2.0"]
    ]
    
    info_table = Table(info_data, colWidths=[4*cm, 12*cm])
    info_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#F0F8FF')),
        ('TEXTCOLOR', (0, 0), (0, -1), colors.HexColor('#2E86AB')),
        ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('PADDING', (0, 0), (-1, -1), 6),
    ]))
    
    story.append(info_table)
    story.append(Spacer(1, 1*cm))
    
    # Ringkasan Hasil
    story.append(Paragraph("RINGKASAN HASIL", heading_style))
    
    summary_data = [
        ["PARAMETER", "NILAI", "SATUAN", "KETERANGAN"],
        ["Diameter Rata-rata", f"{hasil_psa['diameter_rerata']:.2f}", "nm", "Weighted average"],
        ["PDI Terhitung", f"{hasil_psa['pdi_terhitung']:.3f}", "", hasil_psa['klasifikasi']],
        ["Standard Deviation", f"{hasil_psa['std_dev']:.2f}", "nm", f"± {hasil_psa['std_dev']:.1f} nm"],
        ["Koefisien Variasi", f"{hasil_psa['cv']:.1f}", "%", "CV = (σ/μ)×100%"],
        ["Mode Diameter", f"{hasil_psa['mode_diameter']:.1f}", "nm", f"{hasil_psa['mode_percentage']:.1f}% volume"],
        ["Variance", f"{hasil_psa['variance']:.2f}", "nm²", "σ²"],
        ["PDI Input Rata-rata", f"{hasil_psa['pdi_rerata']:.3f}", "", "Weighted average"],
        ["Grade Kualitas", hasil_psa['grade'], "", hasil_psa['warna'] + " " + hasil_psa['klasifikasi']]
    ]
    
    summary_table = Table(summary_data, colWidths=[4*cm, 2.5*cm, 1.5*cm, 7*cm])
    summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2E86AB')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('PADDING', (0, 0), (-1, -1), 6),
        ('ALIGN', (1, 0), (2, -1), 'CENTER'),
        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F9F9F9')]),
    ]))
    
    story.append(summary_table)
    story.append(Spacer(1, 1*cm))
    
    # Interpretasi Kualitas
    story.append(Paragraph("INTERPRETASI KUALITAS", heading_style))
    
    # Warna berdasarkan grade
    if hasil_psa['grade'] in ['A+', 'A']:
        color_box = colors.green
        quality_text = "Kualitas Sangat Baik"
    elif hasil_psa['grade'] == 'B':
        color_box = colors.yellow
        quality_text = "Kualitas Baik"
    elif hasil_psa['grade'] == 'C':
        color_box = colors.orange
        quality_text = "Kualitas Cukup"
    else:
        color_box = colors.red
        quality_text = "Perlu Optimasi"
    
    quality_data = [
        ["INDIKATOR", "NILAI", "STATUS"],
        ["PDI Score", f"{hasil_psa['pdi_terhitung']:.3f}", quality_text],
        ["Uniformitas", f"{100 - hasil_psa['cv']:.1f}%", "Tingkat keseragaman"],
        ["Distribusi", "Normal", "Bentuk distribusi"],
        ["Rekomendasi", "Lihat tabel", "Saran penggunaan"]
    ]
    
    quality_table = Table(quality_data, colWidths=[5*cm, 4*cm, 7*cm])
    quality_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), color_box),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('PADDING', (0, 0), (-1, -1), 8),
        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
    ]))
    
    story.append(quality_table)
    story.append(Spacer(1, 0.5*cm))
    
    # Rekomendasi penggunaan
    if hasil_psa['pdi_terhitung'] < 0.1:
        recommendation = """
        • Cocok untuk aplikasi biomedis (drug delivery, imaging)
        • Ideal untuk aplikasi elektronik presisi
        • Dapat digunakan untuk katalisis selektif
        • Rekomendasi: Lanjutkan metode sintesis
        """
    elif hasil_psa['pdi_terhitung'] < 0.2:
        recommendation = """
        • Cocok untuk coating dan film tipis
        • Dapat digunakan untuk katalisis umum
        • Baik untuk aplikasi sensor
        • Rekomendasi: Optimasi kecil untuk meningkatkan uniformitas
        """
    elif hasil_psa['pdi_terhitung'] < 0.3:
        recommendation = """
        • Cocok untuk aplikasi konstruksi
        • Dapat digunakan untuk bulk material
        • Perlu purifikasi untuk aplikasi presisi
        • Rekomendasi: Evaluasi parameter sintesis
        """
    else:
        recommendation = """
        • Cocok untuk aplikasi yang tidak memerlukan uniformitas tinggi
        • Perlu optimasi signifikan
        • Pertimbangkan metode purifikasi
        • Rekomendasi: Evaluasi ulang metode sintesis
        """
    
    story.append(Paragraph(recommendation, normal_style))
    story.append(PageBreak())
    
    # Data Distribusi
    story.append(Paragraph("DATA DISTRIBUSI UKURAN", heading_style))
    
    # Siapkan data untuk tabel
    df = hasil_psa['dataframe']
    table_data = [["No", "Diameter (nm)", "% Volume", "PDI", "Kumulatif %"]]
    
    cumulative = 0
    for idx, row in df.iterrows():
        cumulative += row['% Volume Normalized']
        table_data.append([
            str(idx + 1),
            f"{row['Diameter (nm)']:.2f}",
            f"{row['% Volume Normalized']:.2f}",
            f"{row['PDI']:.3f}",
            f"{cumulative:.2f}"
        ])
    
    # Tambahkan summary row
    table_data.append([
        "SUMMARY",
        f"Avg: {hasil_psa['diameter_rerata']:.2f}",
        f"Total: {df['% Volume Normalized'].sum():.2f}",
        f"Avg: {df['PDI'].mean():.3f}",
        "100.00"
    ])
    
    dist_table = Table(table_data, colWidths=[1*cm, 3*cm, 3*cm, 3*cm, 3*cm])
    dist_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2E86AB')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -2), 0.5, colors.grey),
        ('GRID', (0, -1), (-1, -1), 1, colors.black),
        ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#F0F0F0')),
        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
        ('ROWBACKGROUNDS', (0, 1), (-1, -2), [colors.white, colors.HexColor('#F9F9F9')]),
        ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
        ('PADDING', (0, 0), (-1, -1), 4),
    ]))
    
    story.append(dist_table)
    story.append(Spacer(1, 1*cm))
    
    # Statistik Deskriptif
    story.append(Paragraph("STATISTIK DESKRIPTIF", heading_style))
    
    stats_data = [
        ["Statistik", "Diameter (nm)", "% Volume", "PDI"],
        ["Minimum", f"{df['Diameter (nm)'].min():.2f}", f"{df['% Volume Normalized'].min():.2f}", f"{df['PDI'].min():.3f}"],
        ["Maksimum", f"{df['Diameter (nm)'].max():.2f}", f"{df['% Volume Normalized'].max():.2f}", f"{df['PDI'].max():.3f}"],
        ["Rata-rata", f"{df['Diameter (nm)'].mean():.2f}", f"{df['% Volume Normalized'].mean():.2f}", f"{df['PDI'].mean():.3f}"],
        ["Median", f"{df['Diameter (nm)'].median():.2f}", f"{df['% Volume Normalized'].median():.2f}", f"{df['PDI'].median():.3f}"],
        ["Std Dev", f"{hasil_psa['std_dev']:.2f}", f"{df['% Volume Normalized'].std():.2f}", f"{df['PDI'].std():.3f}"],
        ["Variance", f"{hasil_psa['variance']:.2f}", f"{df['% Volume Normalized'].var():.2f}", f"{df['PDI'].var():.3f}"]
    ]
    
    stats_table = Table(stats_data, colWidths=[3*cm, 3*cm, 3*cm, 3*cm])
    stats_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#A23B72')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F9F9F9')]),
        ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
        ('PADDING', (0, 0), (-1, -1), 6),
    ]))
    
    story.append(stats_table)
    
    # Footer dengan catatan
    story.append(Spacer(1, 2*cm))
    footer_text = """
    <b>CATATAN:</b><br/>
    1. PDI (Polydispersity Index) mengindikasikan keseragaman ukuran partikel<br/>
    2. PDI < 0.1: Monodispersi (ideal untuk aplikasi presisi)<br/>
    3. PDI 0.1-0.2: Hampir monodispersi (baik untuk kebanyakan aplikasi)<br/>
    4. PDI 0.2-0.3: Polydispersi sedang (perlu optimasi)<br/>
    5. PDI > 0.3: Polydispersi tinggi (perlu optimasi signifikan)<br/>
    6. Laporan ini dibuat otomatis oleh sistem Lab PSA Nano
    """
    
    story.append(Paragraph(footer_text, normal_style))
    
    # Build PDF
    doc.build(story)
    
    return filepath

def create_batch_pdf(hasil_list):
    """
    Membuat PDF gabungan untuk multiple hasil PSA
    """
    temp_dir = tempfile.gettempdir()
    filename = f"Batch_PSA_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    filepath = os.path.join(temp_dir, filename)
    
    doc = SimpleDocTemplate(
        filepath,
        pagesize=A4,
        topMargin=1*cm,
        bottomMargin=1*cm
    )
    
    story = []
    styles = getSampleStyleSheet()
    
    # Title
    title_style = ParagraphStyle(
        'BatchTitle',
        parent=styles['Heading1'],
        fontSize=16,
        alignment=TA_CENTER,
        spaceAfter=30
    )
    
    story.append(Paragraph("BATCH REPORT - ANALISIS PSA NANOMATERIAL", title_style))
    story.append(Paragraph(f"Total Laporan: {len(hasil_list)} | Tanggal: {datetime.now().strftime('%d %B %Y')}", styles['Normal']))
    story.append(Spacer(1, 20))
    
    # Summary table
    summary_data = [["No", "ID", "Diameter (nm)", "PDI", "Klasifikasi", "Grade"]]
    
    for idx, hasil in enumerate(hasil_list, 1):
        summary_data.append([
            str(idx),
            f"PSA-{idx:03d}",
            f"{hasil['diameter_rerata']:.1f}",
            f"{hasil['pdi_terhitung']:.3f}",
            hasil['klasifikasi'].split('(')[0].strip(),
            hasil['grade']
        ])
    
    summary_table = Table(summary_data, colWidths=[1*cm, 2*cm, 3*cm, 2*cm, 5*cm, 2*cm])
    summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
    ]))
    
    story.append(summary_table)
    
    # Build PDF
    doc.build(story)
    
    return filepath
