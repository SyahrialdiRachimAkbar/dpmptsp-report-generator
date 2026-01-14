"""
Enhanced PDF Exporter Module for DPMPTSP Reporting System

This module generates professional PDF reports following the template structure:
- Cover Page with logo and branding
- Table of Contents
- Section separators
- Data sections with charts and narratives
- Closing page with contact info
"""

import io
from pathlib import Path
from typing import Dict, Optional, List
from datetime import datetime

# Try to import PDF generation libraries
try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm, mm, inch
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, 
        Image, PageBreak, KeepTogether, Frame, PageTemplate
    )
    from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
    from reportlab.graphics.shapes import Drawing, Rect, String
    from reportlab.graphics.charts.barcharts import VerticalBarChart
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False


class EnhancedPDFExporter:
    """
    Enhanced PDF exporter with template-matching features.
    
    Features:
    - Cover page with logo
    - Table of contents
    - Section separators
    - Charts with percentage indicators
    - Professional styling
    - Closing page
    """
    
    # Color scheme
    COLORS = {
        'primary': colors.HexColor('#1e3a5f') if REPORTLAB_AVAILABLE else None,
        'secondary': colors.HexColor('#3d7ea6') if REPORTLAB_AVAILABLE else None,
        'accent': colors.HexColor('#5cb85c') if REPORTLAB_AVAILABLE else None,
        'warning': colors.HexColor('#f0ad4e') if REPORTLAB_AVAILABLE else None,
        'danger': colors.HexColor('#d9534f') if REPORTLAB_AVAILABLE else None,
        'light': colors.HexColor('#f8f9fa') if REPORTLAB_AVAILABLE else None,
        'white': colors.white if REPORTLAB_AVAILABLE else None,
    }
    
    def __init__(self, logo_path: Optional[Path] = None, contact_info: Optional[Dict] = None):
        self.logo_path = logo_path
        self.contact_info = contact_info or {
            'nama': 'DPMPTSP Provinsi Lampung',
            'alamat': 'Jl. Wolter Monginsidi No. 69, Bandar Lampung',
            'telepon': '(0721) 123456',
            'website': 'https://dpmptsp.lampungprov.go.id',
            'email': 'dpmptsp@lampungprov.go.id'
        }
        self.styles = None
        if REPORTLAB_AVAILABLE:
            self._setup_styles()
    
    def _setup_styles(self):
        """Setup paragraph styles for the PDF."""
        self.styles = getSampleStyleSheet()
        
        # Cover title style
        self.styles.add(ParagraphStyle(
            name='CoverTitle',
            parent=self.styles['Heading1'],
            fontSize=24,
            alignment=TA_CENTER,
            spaceAfter=20,
            textColor=self.COLORS['primary'],
            leading=30
        ))
        
        # Cover subtitle style
        self.styles.add(ParagraphStyle(
            name='CoverSubtitle',
            parent=self.styles['Normal'],
            fontSize=14,
            alignment=TA_CENTER,
            spaceAfter=10,
            textColor=self.COLORS['secondary']
        ))
        
        # Title style
        self.styles.add(ParagraphStyle(
            name='ReportTitle',
            parent=self.styles['Heading1'],
            fontSize=16,
            alignment=TA_CENTER,
            spaceAfter=20,
            textColor=self.COLORS['primary']
        ))
        
        # Section title style
        self.styles.add(ParagraphStyle(
            name='SectionTitle',
            parent=self.styles['Heading2'],
            fontSize=14,
            spaceBefore=20,
            spaceAfter=10,
            textColor=self.COLORS['primary'],
        ))
        
        # Subsection style
        self.styles.add(ParagraphStyle(
            name='SubSection',
            parent=self.styles['Heading3'],
            fontSize=12,
            spaceBefore=15,
            spaceAfter=8,
            textColor=self.COLORS['secondary']
        ))
        
        # Body text style
        self.styles.add(ParagraphStyle(
            name='ReportBody',
            parent=self.styles['Normal'],
            fontSize=10,
            alignment=TA_JUSTIFY,
            spaceAfter=12,
            leading=14
        ))
        
        # TOC style
        self.styles.add(ParagraphStyle(
            name='TOCEntry',
            parent=self.styles['Normal'],
            fontSize=12,
            spaceAfter=8,
            leftIndent=20
        ))
        
        # Section separator title
        self.styles.add(ParagraphStyle(
            name='SeparatorTitle',
            parent=self.styles['Heading1'],
            fontSize=28,
            alignment=TA_CENTER,
            textColor=self.COLORS['primary'],
            spaceBefore=100
        ))
    
    def is_available(self) -> bool:
        """Check if PDF export is available."""
        return REPORTLAB_AVAILABLE
    
    def export_report(
        self,
        report,
        stats: Dict,
        narratives,
        charts: Dict[str, bytes],
        output_path: Optional[Path] = None
    ) -> bytes:
        """Export a complete report to PDF with enhanced template."""
        if not REPORTLAB_AVAILABLE:
            raise ImportError("ReportLab is not installed.")
        
        buffer = io.BytesIO()
        
        doc = SimpleDocTemplate(
            buffer,
            pagesize=A4,
            rightMargin=1.5*cm,
            leftMargin=1.5*cm,
            topMargin=2*cm,
            bottomMargin=2*cm
        )
        
        story = []
        
        # 1. Cover Page
        story.extend(self._create_cover_page(report))
        story.append(PageBreak())
        
        # 2. Table of Contents
        story.extend(self._create_table_of_contents())
        story.append(PageBreak())
        
        # 3. Section 1: Nomor Induk Berusaha (NIB)
        story.extend(self._create_section_separator("1", "NOMOR INDUK BERUSAHA (NIB)"))
        story.append(PageBreak())
        
        # 3.1 Pendahuluan
        story.append(Paragraph("Pendahuluan", self.styles['SectionTitle']))
        story.append(Paragraph(narratives.pendahuluan.replace('\n', '<br/>'), self.styles['ReportBody']))
        story.append(Spacer(1, 10))
        
        # Add metrics summary
        story.extend(self._create_metrics_section(stats))
        
        # 3.2 Rekapitulasi NIB
        story.append(Paragraph("1.1 Rekapitulasi Data NIB", self.styles['SectionTitle']))
        
        if 'monthly' in charts:
            story.append(self._create_chart_image(charts['monthly'], width=16*cm))
            story.append(Spacer(1, 10))
        
        story.append(Paragraph(narratives.rekapitulasi_nib.replace('\n', '<br/>'), self.styles['ReportBody']))
        
        # 3.3 Per Kabupaten/Kota
        story.append(Paragraph("1.2 Rekapitulasi per Kabupaten/Kota", self.styles['SectionTitle']))
        
        if 'kab_kota' in charts:
            story.append(self._create_chart_image(charts['kab_kota'], width=16*cm))
            story.append(Spacer(1, 10))
        
        story.append(Paragraph(narratives.rekapitulasi_kab_kota.replace('\n', '<br/>'), self.styles['ReportBody']))
        
        # Add data table
        story.append(Spacer(1, 10))
        story.append(Paragraph("Tabel Data per Kabupaten/Kota", self.styles['SubSection']))
        story.extend(self._create_data_table(report, stats))
        
        # 3.4 Status PM
        story.append(PageBreak())
        story.append(Paragraph("1.3 Status Penanaman Modal", self.styles['SectionTitle']))
        
        if 'pm' in charts:
            story.append(self._create_chart_image(charts['pm'], width=12*cm))
            story.append(Spacer(1, 10))
        
        story.append(Paragraph(narratives.status_pm.replace('\n', '<br/>'), self.styles['ReportBody']))
        
        # 3.5 Pelaku Usaha
        story.append(Paragraph("1.4 Kategori Pelaku Usaha", self.styles['SectionTitle']))
        
        if 'pelaku' in charts:
            story.append(self._create_chart_image(charts['pelaku'], width=12*cm))
            story.append(Spacer(1, 10))
        
        story.append(Paragraph(narratives.pelaku_usaha.replace('\n', '<br/>'), self.styles['ReportBody']))
        
        # 3.6 Sektor & Risiko (if data available)
        sektor_risiko = stats.get('sektor_risiko', {})
        if sektor_risiko:
            story.append(Paragraph("1.5 Perizinan Berdasarkan Risiko dan Sektor", self.styles['SectionTitle']))
            
            if 'risk' in charts:
                story.append(self._create_chart_image(charts['risk'], width=12*cm))
                story.append(Spacer(1, 10))
            
            if 'sector' in charts:
                story.append(self._create_chart_image(charts['sector'], width=14*cm))
                story.append(Spacer(1, 10))
            
            # Generate sektor risiko narrative
            sektor_narrative = self._generate_sektor_risiko_narrative(sektor_risiko)
            story.append(Paragraph(sektor_narrative.replace('\n', '<br/>'), self.styles['ReportBody']))
        
        # 4. Kesimpulan
        story.append(PageBreak())
        story.append(Paragraph("Kesimpulan", self.styles['SectionTitle']))
        story.append(Paragraph(narratives.kesimpulan.replace('\n', '<br/>'), self.styles['ReportBody']))
        
        # 5. Closing Page
        story.append(PageBreak())
        story.extend(self._create_closing_page())
        
        # Build PDF
        doc.build(story)
        
        pdf_bytes = buffer.getvalue()
        buffer.close()
        
        if output_path:
            with open(output_path, 'wb') as f:
                f.write(pdf_bytes)
        
        return pdf_bytes
    
    def _create_cover_page(self, report) -> list:
        """Create cover page with logo, title, and metadata."""
        elements = []
        
        elements.append(Spacer(1, 2*cm))
        
        # Add logo if available
        if self.logo_path and self.logo_path.exists():
            try:
                logo = Image(str(self.logo_path), width=12*cm, height=3*cm)
                logo.hAlign = 'CENTER'
                elements.append(logo)
                elements.append(Spacer(1, 2*cm))
            except Exception:
                pass
        
        # Main title
        elements.append(Paragraph(
            "LAPORAN REKAPITULASI<br/>DATA NOMOR INDUK BERUSAHA (NIB)",
            self.styles['CoverTitle']
        ))
        
        elements.append(Spacer(1, 1*cm))
        
        # Period
        elements.append(Paragraph(
            f"PERIODE {report.period_name} TAHUN {report.year}",
            self.styles['CoverSubtitle']
        ))
        
        elements.append(Spacer(1, 3*cm))
        
        # Metadata table
        metadata = [
            ["Tim Pengolah Data", ":", "DPMPTSP Provinsi Lampung"],
            ["Sumber Data", ":", "OSS-RBA (Online Single Submission)"],
            ["Tanggal Penarikan Data", ":", datetime.now().strftime("%d %B %Y")],
        ]
        
        meta_table = Table(metadata, colWidths=[5*cm, 0.5*cm, 8*cm])
        meta_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (0, -1), 'RIGHT'),
            ('ALIGN', (1, 0), (1, -1), 'CENTER'),
            ('ALIGN', (2, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 11),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
        ]))
        elements.append(meta_table)
        
        return elements
    
    def _create_table_of_contents(self) -> list:
        """Create table of contents page."""
        elements = []
        
        elements.append(Paragraph("DAFTAR ISI", self.styles['ReportTitle']))
        elements.append(Spacer(1, 1*cm))
        
        toc_items = [
            ("1.", "Nomor Induk Berusaha (NIB)"),
            ("   1.1", "Rekapitulasi Data NIB"),
            ("   1.2", "Rekapitulasi per Kabupaten/Kota"),
            ("   1.3", "Status Penanaman Modal"),
            ("   1.4", "Kategori Pelaku Usaha"),
            ("2.", "Kesimpulan"),
        ]
        
        for num, title in toc_items:
            elements.append(Paragraph(f"{num} {title}", self.styles['TOCEntry']))
        
        return elements
    
    def _create_section_separator(self, number: str, title: str) -> list:
        """Create section separator page."""
        elements = []
        
        elements.append(Spacer(1, 5*cm))
        
        # Section number in circle
        elements.append(Paragraph(
            f"<font size='48' color='#1e3a5f'><b>{number}</b></font>",
            ParagraphStyle('BigNumber', alignment=TA_CENTER)
        ))
        
        elements.append(Spacer(1, 1*cm))
        
        # Section title
        elements.append(Paragraph(title, self.styles['SeparatorTitle']))
        
        return elements
    
    def _create_metrics_section(self, stats: Dict) -> list:
        """Create metrics summary section with percentage indicators."""
        elements = []
        
        total_nib = stats.get('total_nib', 0)
        pm_dist = stats.get('pm_distribution', {})
        pelaku = stats.get('pelaku_usaha_distribution', {})
        change_pct = stats.get('change_percentage')
        
        # Format change indicator
        if change_pct is not None:
            if change_pct > 0:
                indicator = f"▲ +{change_pct:.1f}%"
                indicator_color = '#5cb85c'
            elif change_pct < 0:
                indicator = f"▼ {change_pct:.1f}%"
                indicator_color = '#d9534f'
            else:
                indicator = "— 0%"
                indicator_color = '#f0ad4e'
        else:
            indicator = ""
            indicator_color = '#666666'
        
        # Create metrics table with indicators
        header = ['Total NIB', 'PMDN', 'PMA', 'UMK', 'Perubahan']
        values = [
            f"{total_nib:,}".replace(',', '.'),
            f"{pm_dist.get('PMDN', 0):,}".replace(',', '.'),
            f"{pm_dist.get('PMA', 0):,}".replace(',', '.'),
            f"{pelaku.get('UMK', 0):,}".replace(',', '.'),
            indicator
        ]
        
        data = [header, values]
        
        table = Table(data, colWidths=[3.2*cm, 3.2*cm, 3.2*cm, 3.2*cm, 3.2*cm])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), self.COLORS['primary']),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('FONTSIZE', (0, 1), (-1, -1), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('TOPPADDING', (0, 1), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 10),
            ('BACKGROUND', (0, 1), (-1, -1), self.COLORS['light']),
            ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#dee2e6')),
            # Color the change indicator
            ('TEXTCOLOR', (4, 1), (4, 1), colors.HexColor(indicator_color)),
            ('FONTNAME', (4, 1), (4, 1), 'Helvetica-Bold'),
        ]))
        
        elements.append(table)
        elements.append(Spacer(1, 20))
        
        return elements
    
    def _create_chart_image(self, chart_bytes: bytes, width=15*cm) -> Image:
        """Create an image element from chart bytes."""
        img_buffer = io.BytesIO(chart_bytes)
        img = Image(img_buffer, width=width, height=width*0.6)
        img.hAlign = 'CENTER'
        return img
    
    def _create_data_table(self, report, stats: Dict) -> list:
        """Create data table for kabupaten/kota."""
        elements = []
        
        top_5 = stats.get('top_5_locations', [])
        if not top_5:
            return elements
        
        header = ['No', 'Kabupaten/Kota', 'Total NIB', 'Persentase']
        data = [header]
        
        total_nib = stats.get('total_nib', 1)
        
        for i, loc in enumerate(top_5, 1):
            pct = (loc['Total'] / total_nib * 100) if total_nib > 0 else 0
            data.append([
                str(i),
                loc['Kabupaten/Kota'],
                f"{loc['Total']:,}".replace(',', '.'),
                f"{pct:.1f}%"
            ])
        
        table = Table(data, colWidths=[1*cm, 6*cm, 4*cm, 3*cm])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), self.COLORS['primary']),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (1, 1), (1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
            ('TOPPADDING', (0, 1), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#dee2e6')),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, self.COLORS['light']])
        ]))
        
        elements.append(table)
        
        return elements
    
    def _create_closing_page(self) -> list:
        """Create closing page with thank you message and contact info."""
        elements = []
        
        elements.append(Spacer(1, 4*cm))
        
        # Thank you message
        elements.append(Paragraph(
            "<font size='36' color='#1e3a5f'><b>TERIMA KASIH</b></font>",
            ParagraphStyle('ThankYou', alignment=TA_CENTER)
        ))
        
        elements.append(Spacer(1, 2*cm))
        
        # Add logo if available
        if self.logo_path and self.logo_path.exists():
            try:
                logo = Image(str(self.logo_path), width=8*cm, height=2*cm)
                logo.hAlign = 'CENTER'
                elements.append(logo)
                elements.append(Spacer(1, 1*cm))
            except Exception:
                pass
        
        # Contact info
        contact_text = f"""
        <b>{self.contact_info.get('nama', '')}</b><br/>
        {self.contact_info.get('alamat', '')}<br/>
        Telepon: {self.contact_info.get('telepon', '')}<br/>
        Website: {self.contact_info.get('website', '')}<br/>
        Email: {self.contact_info.get('email', '')}
        """
        
        elements.append(Paragraph(
            contact_text,
            ParagraphStyle('ContactInfo', alignment=TA_CENTER, fontSize=10, leading=14)
        ))
        
        return elements
    
    def _generate_sektor_risiko_narrative(self, sektor_risiko_data: dict) -> str:
        """Generate narrative for Sektor & Risiko section."""
        total_risiko = (
            sektor_risiko_data.get('risiko_rendah', 0) +
            sektor_risiko_data.get('risiko_menengah_rendah', 0) +
            sektor_risiko_data.get('risiko_menengah_tinggi', 0) +
            sektor_risiko_data.get('risiko_tinggi', 0)
        )
        
        if total_risiko == 0:
            return "Data perizinan berdasarkan risiko belum tersedia."
        
        # Find dominant risk level
        risk_levels = {
            'Rendah': sektor_risiko_data.get('risiko_rendah', 0),
            'Menengah Rendah': sektor_risiko_data.get('risiko_menengah_rendah', 0),
            'Menengah Tinggi': sektor_risiko_data.get('risiko_menengah_tinggi', 0),
            'Tinggi': sektor_risiko_data.get('risiko_tinggi', 0),
        }
        dominant_risk = max(risk_levels, key=risk_levels.get)
        dominant_pct = (risk_levels[dominant_risk] / total_risiko * 100) if total_risiko > 0 else 0
        
        # Find dominant sector
        sectors = {
            'Perindustrian': sektor_risiko_data.get('sektor_perindustrian', 0),
            'Kelautan & Perikanan': sektor_risiko_data.get('sektor_kelautan', 0),
            'Pertanian': sektor_risiko_data.get('sektor_pertanian', 0),
            'Energi': sektor_risiko_data.get('sektor_energi', 0),
            'Kesehatan': sektor_risiko_data.get('sektor_kesehatan', 0),
            'Perhubungan': sektor_risiko_data.get('sektor_perhubungan', 0),
            'Pariwisata': sektor_risiko_data.get('sektor_pariwisata', 0),
            'Komunikasi': sektor_risiko_data.get('sektor_komunikasi', 0),
        }
        sectors_filtered = {k: v for k, v in sectors.items() if v > 0}
        
        if sectors_filtered:
            total_sektor = sum(sectors_filtered.values())
            dominant_sector = max(sectors_filtered, key=sectors_filtered.get)
            dominant_sector_pct = (sectors_filtered[dominant_sector] / total_sektor * 100) if total_sektor > 0 else 0
            sector_text = f"Berdasarkan sektor usaha, sektor {dominant_sector} menempati posisi tertinggi dengan {sectors_filtered[dominant_sector]:,} perizinan ({dominant_sector_pct:.1f}%)."
        else:
            sector_text = ""
        
        narrative = f"""Berdasarkan tingkat risiko, perizinan dengan kategori "{dominant_risk}" mendominasi dengan {risk_levels[dominant_risk]:,} perizinan ({dominant_pct:.1f}%) dari total {total_risiko:,} perizinan.

{sector_text}

Hal ini menunjukkan bahwa sebagian besar kegiatan usaha yang dilaksanakan di Provinsi Lampung berada pada kategori risiko yang relatif dapat dikelola."""
        
        return narrative.replace(',', '.')
