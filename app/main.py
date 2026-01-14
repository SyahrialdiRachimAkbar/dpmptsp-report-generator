"""
DPMPTSP Automated Reporting System - Streamlit Application

Main entry point for the web application that allows users to:
1. Upload monthly Excel data files
2. Select reporting period (Triwulan/Semester/Tahunan)
3. Preview generated reports with charts and narratives
4. Export reports to PDF or Word
"""

import streamlit as st
import pandas as pd
from pathlib import Path
import sys
import io
from datetime import datetime

# Add app directory to path
sys.path.insert(0, str(Path(__file__).parent.parent))

from app.data.loader import DataLoader
from app.data.aggregator import DataAggregator
from app.visualization.charts import ChartGenerator
from app.narrative.generator import NarrativeGenerator
from app.config import LOGO_PATH, TRIWULAN_KE_BULAN, NAMA_BULAN


# Page configuration
st.set_page_config(
    page_title="Laporan Otomatis DPMPTSP Lampung",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS - Modern UI/UX Design
st.markdown("""
<style>
    /* ===== ROOT VARIABLES ===== */
    :root {
        --primary-color: #1e3a5f;
        --primary-light: #2d5a87;
        --secondary-color: #3d7ea6;
        --accent-color: #5cb85c;
        --warning-color: #f0ad4e;
        --danger-color: #d9534f;
        --background-light: #f0f4f8;
        --card-bg: rgba(255, 255, 255, 0.95);
        --text-primary: #2c3e50;
        --text-secondary: #6c757d;
        --shadow-soft: 0 4px 20px rgba(0, 0, 0, 0.08);
        --shadow-hover: 0 8px 30px rgba(0, 0, 0, 0.12);
        --gradient-primary: linear-gradient(135deg, #1e3a5f 0%, #3d7ea6 100%);
        --gradient-accent: linear-gradient(135deg, #5cb85c 0%, #3d9e52 100%);
        --border-radius: 12px;
        --transition: all 0.3s ease;
    }
    
    /* ===== GLOBAL STYLES ===== */
    .stApp {
        background: linear-gradient(180deg, #f0f4f8 0%, #e8eef3 100%);
    }
    
    /* ===== SIDEBAR IMPROVEMENTS ===== */
    [data-testid="stSidebar"] {
        background: var(--gradient-primary);
        padding-top: 0;
    }
    
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] .stRadio label span {
        color: white !important;
    }
    
    [data-testid="stSidebar"] .stSelectbox label,
    [data-testid="stSidebar"] .stMultiSelect label,
    [data-testid="stSidebar"] .stFileUploader label {
        color: white !important;
        font-weight: 500;
    }
    
    [data-testid="stSidebar"] hr {
        border-color: rgba(255, 255, 255, 0.2);
    }
    
    [data-testid="stSidebar"] .stButton button {
        width: 100%;
        background: rgba(255, 255, 255, 0.15);
        border: 1px solid rgba(255, 255, 255, 0.3);
        color: white;
        transition: var(--transition);
    }
    
    [data-testid="stSidebar"] .stButton button:hover {
        background: rgba(255, 255, 255, 0.25);
        border-color: rgba(255, 255, 255, 0.5);
        transform: translateY(-2px);
    }
    
    /* ===== HEADER STYLES ===== */
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        background: var(--gradient-primary);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        text-align: center;
        padding: 1rem;
        margin-bottom: 0;
    }
    
    .sub-header {
        font-size: 1.1rem;
        color: var(--text-secondary);
        text-align: center;
        margin-bottom: 2rem;
        font-weight: 400;
    }
    
    /* ===== METRIC CARDS - GLASSMORPHISM ===== */
    .metric-card {
        background: var(--card-bg);
        backdrop-filter: blur(10px);
        border-radius: var(--border-radius);
        padding: 1.5rem;
        box-shadow: var(--shadow-soft);
        border: 1px solid rgba(255, 255, 255, 0.8);
        transition: var(--transition);
        position: relative;
        overflow: hidden;
    }
    
    .metric-card:hover {
        transform: translateY(-5px);
        box-shadow: var(--shadow-hover);
    }
    
    .metric-card::before {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        height: 4px;
        background: var(--gradient-primary);
    }
    
    .metric-card.accent::before {
        background: var(--gradient-accent);
    }
    
    /* ===== CUSTOM METRIC DISPLAY ===== */
    .custom-metric {
        background: var(--card-bg);
        backdrop-filter: blur(10px);
        border-radius: var(--border-radius);
        padding: 1.5rem;
        box-shadow: var(--shadow-soft);
        border: 1px solid rgba(255, 255, 255, 0.8);
        text-align: center;
        transition: var(--transition);
        height: 100%;
    }
    
    .custom-metric:hover {
        transform: translateY(-3px);
        box-shadow: var(--shadow-hover);
    }
    
    .metric-icon {
        font-size: 2rem;
        margin-bottom: 0.5rem;
    }
    
    .metric-value {
        font-size: 2rem;
        font-weight: 700;
        color: var(--primary-color);
        margin-bottom: 0.25rem;
    }
    
    .metric-label {
        font-size: 0.9rem;
        color: var(--text-secondary);
        font-weight: 500;
    }
    
    .metric-delta {
        font-size: 0.85rem;
        font-weight: 600;
        margin-top: 0.5rem;
        padding: 0.25rem 0.5rem;
        border-radius: 4px;
    }
    
    .metric-delta.positive {
        background: rgba(92, 184, 92, 0.15);
        color: #27a745;
    }
    
    .metric-delta.negative {
        background: rgba(217, 83, 79, 0.15);
        color: #dc3545;
    }
    
    /* ===== NARRATIVE BOX ===== */
    .narrative-box {
        background: var(--card-bg);
        backdrop-filter: blur(10px);
        border-left: 4px solid var(--primary-color);
        padding: 1.25rem 1.5rem;
        margin: 1rem 0;
        border-radius: 0 var(--border-radius) var(--border-radius) 0;
        box-shadow: var(--shadow-soft);
        line-height: 1.7;
        color: var(--text-primary);
        font-size: 0.95rem;
    }
    
    /* ===== SECTION TITLE ===== */
    .section-title {
        color: var(--primary-color);
        font-size: 1.4rem;
        font-weight: 700;
        margin-top: 2.5rem;
        margin-bottom: 1.25rem;
        padding-bottom: 0.75rem;
        border-bottom: 3px solid;
        border-image: var(--gradient-primary) 1;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }
    
    .section-title::before {
        content: '📋';
        font-size: 1.2rem;
    }
    
    /* ===== DATA TABLES ===== */
    [data-testid="stDataFrame"] {
        border-radius: var(--border-radius);
        overflow: hidden;
        box-shadow: var(--shadow-soft);
    }
    
    /* ===== CHARTS CONTAINER ===== */
    .chart-container {
        background: var(--card-bg);
        border-radius: var(--border-radius);
        padding: 1rem;
        box-shadow: var(--shadow-soft);
        margin: 1rem 0;
    }
    
    /* ===== EXPORT BUTTONS ===== */
    .export-section {
        background: var(--card-bg);
        border-radius: var(--border-radius);
        padding: 1.5rem;
        box-shadow: var(--shadow-soft);
        margin-top: 2rem;
    }
    
    .stDownloadButton button {
        background: var(--gradient-primary) !important;
        color: white !important;
        border: none !important;
        padding: 0.75rem 1.5rem !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        transition: var(--transition) !important;
    }
    
    .stDownloadButton button:hover {
        opacity: 0.9;
        transform: translateY(-2px);
        box-shadow: 0 4px 15px rgba(30, 58, 95, 0.3);
    }
    
    /* ===== LOADING ANIMATION ===== */
    @keyframes pulse {
        0%, 100% { opacity: 1; }
        50% { opacity: 0.5; }
    }
    
    @keyframes slideIn {
        from {
            opacity: 0;
            transform: translateY(20px);
        }
        to {
            opacity: 1;
            transform: translateY(0);
        }
    }
    
    .animate-slide-in {
        animation: slideIn 0.5s ease forwards;
    }
    
    /* ===== SUCCESS/INFO MESSAGES ===== */
    .stSuccess, .stInfo, .stWarning {
        border-radius: 8px !important;
    }
    
    /* ===== FILE UPLOADER ===== */
    [data-testid="stFileUploader"] {
        padding: 0.5rem;
    }
    
    /* ===== LOGO HEADER ===== */
    .logo-container {
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 1rem;
        padding: 1rem;
        background: var(--card-bg);
        border-radius: var(--border-radius);
        box-shadow: var(--shadow-soft);
        margin-bottom: 1.5rem;
    }
    
    /* ===== RESPONSIVE DESIGN ===== */
    @media (max-width: 768px) {
        .main-header {
            font-size: 1.6rem;
        }
        
        .metric-value {
            font-size: 1.5rem;
        }
        
        .section-title {
            font-size: 1.2rem;
        }
        
        .custom-metric {
            padding: 1rem;
        }
    }
    
    /* ===== TABS STYLING ===== */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background: var(--card-bg);
        border-radius: var(--border-radius);
        padding: 0.5rem;
    }
    
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        padding: 0.5rem 1rem;
        font-weight: 500;
    }
    
    .stTabs [data-baseweb="tab"][aria-selected="true"] {
        background: var(--gradient-primary);
        color: white;
    }
    
    /* ===== DIVIDER ===== */
    hr {
        border: none;
        height: 1px;
        background: linear-gradient(90deg, transparent, var(--secondary-color), transparent);
        margin: 2rem 0;
    }
    
    /* ===== HIDE STREAMLIT BRANDING ===== */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
    /* ===== SPINNER ===== */
    .stSpinner > div {
        border-top-color: var(--primary-color) !important;
    }
</style>
""", unsafe_allow_html=True)


def init_session_state():
    """Initialize session state variables."""
    if 'aggregator' not in st.session_state:
        st.session_state.aggregator = DataAggregator()
    if 'loaded_files' not in st.session_state:
        st.session_state.loaded_files = []
    if 'report' not in st.session_state:
        st.session_state.report = None
    if 'stats' not in st.session_state:
        st.session_state.stats = None


def render_header():
    """Render the application header."""
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        # Try to load logo
        if LOGO_PATH.exists():
            st.image(str(LOGO_PATH), width=600)
        else:
            st.markdown('<div class="main-header">DPMPTSP PROVINSI LAMPUNG</div>', 
                       unsafe_allow_html=True)
        
        st.markdown(
            '<div class="sub-header">Sistem Laporan Otomatis Perizinan Berusaha</div>',
            unsafe_allow_html=True
        )


def render_sidebar():
    """Render sidebar with file upload and period selection."""
    with st.sidebar:
        st.header("📁 Upload Data")
        
        # File uploader
        uploaded_files = st.file_uploader(
            "Upload file Excel bulanan",
            type=['xlsx', 'xls'],
            accept_multiple_files=True,
            help="Upload file OLAH DATA OSS BULAN *.xlsx"
        )
        
        if uploaded_files:
            for file in uploaded_files:
                if file.name not in [f.name for f in st.session_state.loaded_files]:
                    st.session_state.loaded_files.append(file)
            
            st.success(f"✅ {len(st.session_state.loaded_files)} file diupload")
            
            # Show uploaded files
            with st.expander("File yang diupload"):
                for f in st.session_state.loaded_files:
                    st.text(f"• {f.name}")
        
        st.divider()
        
        # Period selection
        st.header("📅 Pilih Periode")
        
        tahun = st.selectbox(
            "Tahun",
            options=[2025, 2024, 2023, 2026],
            index=0
        )
        
        jenis_periode = st.radio(
            "Jenis Periode",
            options=["Triwulan", "Semester", "Tahunan"],
            index=0
        )
        
        if jenis_periode == "Triwulan":
            periode = st.selectbox(
                "Pilih Triwulan",
                options=["TW I", "TW II", "TW III", "TW IV"]
            )
        elif jenis_periode == "Semester":
            periode = st.selectbox(
                "Pilih Semester",
                options=["Semester I", "Semester II"]
            )
        else:
            periode = str(tahun)
        
        st.divider()
        
        # Generate button
        if st.button("🚀 Generate Laporan", type="primary", use_container_width=True):
            if not st.session_state.loaded_files:
                st.error("⚠️ Upload file data terlebih dahulu!")
            else:
                with st.spinner("Memproses data..."):
                    process_data(st.session_state.loaded_files, jenis_periode, periode, tahun)
                st.success("✅ Laporan berhasil dibuat!")
                st.rerun()
        
        # Clear button
        if st.button("🗑️ Clear Data", use_container_width=True):
            st.session_state.loaded_files = []
            st.session_state.report = None
            st.session_state.stats = None
            st.session_state.aggregator = DataAggregator()
            st.rerun()
        
        return jenis_periode, periode, tahun


def process_data(uploaded_files, jenis_periode: str, periode: str, tahun: int):
    """Process uploaded files and generate report."""
    loader = DataLoader()
    aggregator = DataAggregator()
    
    # Storage for sektor risiko data
    all_sektor_risiko = []
    
    # Load each file directly from memory using BytesIO
    for file in uploaded_files:
        try:
            # Read file content into BytesIO
            file_content = io.BytesIO(file.getvalue())
            
            # Load NIB data using pandas directly from BytesIO
            data = loader.load_from_bytes(file_content, file.name)
            
            # Check if this is a quarterly file
            if data.get('is_quarterly'):
                # Process quarterly file - has multiple months
                file_content.seek(0)  # Reset file pointer
                monthly_data = loader.load_quarterly_file(file_content, file.name)
                
                # Add each month's data to aggregator
                for month, month_data in monthly_data.items():
                    if month_data.get('nib'):
                        aggregator.loaded_data[month] = month_data
                        
                # Load sektor risiko data from quarterly file sheets
                file_content.seek(0)
                xl = pd.ExcelFile(file_content)
                for sheet_name in xl.sheet_names:
                    if 'RESIKO' in sheet_name.upper() or 'RISIKO' in sheet_name.upper():
                        df = pd.read_excel(xl, sheet_name=sheet_name, header=None)
                        sektor_data = loader.parse_sektor_resiko_sheet(df)
                        all_sektor_risiko.extend(sektor_data)
            else:
                # Regular monthly file
                month = data.get('month')
                if month:
                    aggregator.loaded_data[month] = data
                
                # Also try to load sektor risiko data
                file_content.seek(0)  # Reset file pointer
                xl = pd.ExcelFile(file_content)
                for sheet_name in xl.sheet_names:
                    if 'RESIKO' in sheet_name.upper() or 'RISIKO' in sheet_name.upper():
                        df = pd.read_excel(xl, sheet_name=sheet_name, header=None)
                        sektor_data = loader.parse_sektor_resiko_sheet(df)
                        all_sektor_risiko.extend(sektor_data)
                        break
                    
        except Exception as e:
            st.warning(f"⚠️ Error loading {file.name}: {str(e)}")
    
    # Generate report based on period type
    if jenis_periode == "Triwulan":
        report = aggregator.aggregate_triwulan(periode, tahun)
    elif jenis_periode == "Semester":
        report = aggregator.aggregate_semester(periode, tahun)
    else:
        report = aggregator.aggregate_tahunan(tahun)
    
    stats = aggregator.get_summary_stats(report)
    
    # Aggregate sektor risiko data and add to stats
    if all_sektor_risiko:
        sektor_totals = {
            'risiko_rendah': sum(d.risiko_rendah for d in all_sektor_risiko),
            'risiko_menengah_rendah': sum(d.risiko_menengah_rendah for d in all_sektor_risiko),
            'risiko_menengah_tinggi': sum(d.risiko_menengah_tinggi for d in all_sektor_risiko),
            'risiko_tinggi': sum(d.risiko_tinggi for d in all_sektor_risiko),
            'sektor_energi': sum(d.sektor_energi for d in all_sektor_risiko),
            'sektor_kelautan': sum(d.sektor_kelautan for d in all_sektor_risiko),
            'sektor_kesehatan': sum(d.sektor_kesehatan for d in all_sektor_risiko),
            'sektor_komunikasi': sum(d.sektor_komunikasi for d in all_sektor_risiko),
            'sektor_pariwisata': sum(d.sektor_pariwisata for d in all_sektor_risiko),
            'sektor_perhubungan': sum(d.sektor_perhubungan for d in all_sektor_risiko),
            'sektor_perindustrian': sum(d.sektor_perindustrian for d in all_sektor_risiko),
            'sektor_pertanian': sum(d.sektor_pertanian for d in all_sektor_risiko),
            'total': sum(d.total for d in all_sektor_risiko),
        }
        stats['sektor_risiko'] = sektor_totals
    
    st.session_state.report = report
    st.session_state.stats = stats
    st.session_state.aggregator = aggregator


def render_metrics(stats: dict):
    """Render key metrics with custom styled cards."""
    total_nib = stats.get('total_nib', 0)
    pm_dist = stats.get('pm_distribution', {})
    pelaku = stats.get('pelaku_usaha_distribution', {})
    change_pct = stats.get('change_percentage')
    
    # Helper to formatting percentage
    def format_pct(val):
        if val > 99 and val < 100:
            return f"{val:.2f}%"
        elif val < 1 and val > 0:
            return f"{val:.2f}%"
        else:
            return f"{val:.1f}%"

    # Format delta
    if change_pct is not None:
        delta_class = "positive" if change_pct >= 0 else "negative"
        delta_symbol = "▲" if change_pct >= 0 else "▼"
        delta_html = f'<div class="metric-delta {delta_class}">{delta_symbol} {abs(change_pct):.1f}%</div>'
    else:
        delta_html = ""
    
    # Create 4 columns
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(f"""
        <div class="custom-metric">
            <div class="metric-icon">📊</div>
            <div class="metric-value">{total_nib:,}</div>
            <div class="metric-label">Total NIB</div>
            {delta_html}
        </div>
        """.replace(",", "."), unsafe_allow_html=True)
    
    with col2:
        pmdn_val = pm_dist.get('PMDN', 0)
        pmdn_pct = pm_dist.get('PMDN_pct', 0)
        st.markdown(f"""
        <div class="custom-metric">
            <div class="metric-icon">🏢</div>
            <div class="metric-value">{pmdn_val:,}</div>
            <div class="metric-label">PMDN</div>
            <div class="metric-delta positive">{format_pct(pmdn_pct)}</div>
        </div>
        """.replace(",", "."), unsafe_allow_html=True)
    
    with col3:
        pma_val = pm_dist.get('PMA', 0)
        pma_pct = pm_dist.get('PMA_pct', 0)
        st.markdown(f"""
        <div class="custom-metric">
            <div class="metric-icon">🌍</div>
            <div class="metric-value">{pma_val:,}</div>
            <div class="metric-label">PMA</div>
            <div class="metric-delta positive">{format_pct(pma_pct)}</div>
        </div>
        """.replace(",", "."), unsafe_allow_html=True)
    
    with col4:
        umk_val = pelaku.get('UMK', 0)
        umk_pct = pelaku.get('UMK_pct', 0)
        st.markdown(f"""
        <div class="custom-metric">
            <div class="metric-icon">🏪</div>
            <div class="metric-value">{umk_val:,}</div>
            <div class="metric-label">UMK</div>
            <div class="metric-delta positive">{format_pct(umk_pct)}</div>
        </div>
        """.replace(",", "."), unsafe_allow_html=True)


def generate_sektor_risiko_narrative(sektor_risiko_data: dict) -> str:
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
    
    narrative = f"""Berdasarkan tingkat risiko, perizinan dengan kategori "{dominant_risk}" mendominasi dengan {risk_levels[dominant_risk]:,} perizinan ({dominant_pct:.1f}%) dari total {total_risiko:,} perizinan yang diterbitkan.

{sector_text}

Hal ini menunjukkan bahwa sebagian besar kegiatan usaha yang dilaksanakan di Provinsi Lampung berada pada kategori risiko yang relatif dapat dikelola dengan prosedur pengawasan standar."""
    
    return narrative.replace(',', '.')


def render_report(report, stats: dict):
    """Render the full report with charts and narratives."""
    chart_gen = ChartGenerator()
    narrative_gen = NarrativeGenerator()
    
    # Generate narratives
    narratives = narrative_gen.generate_full_narrative(report, stats)
    
    # Section 1: Pendahuluan
    st.markdown('<div class="section-title">Pendahuluan</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="narrative-box">{narratives.pendahuluan}</div>', 
                unsafe_allow_html=True)
    
    # Section 2: Rekapitulasi NIB Total
    st.markdown('<div class="section-title">1.1 Rekapitulasi Data NIB</div>', 
                unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Monthly bar chart with trendline
        monthly_data = stats.get('monthly_totals', {})
        if monthly_data:
            fig_monthly = chart_gen.create_monthly_bar_with_trendline(
                monthly_data,
                title="NIB per Bulan",
                show_trendline=True
            )
            st.plotly_chart(fig_monthly, use_container_width=True)
    
    with col2:
        # Q-o-Q comparison
        if stats.get('prev_period_total'):
            prev_data = {'prev': stats['prev_period_total']}
            current_data = {'current': stats['total_nib']}
            fig_qoq = chart_gen.create_qoq_comparison_bar(
                current_data=current_data,
                previous_data=prev_data,
                current_label=report.period_name,
                previous_label="Periode Sebelumnya"
            )
            st.plotly_chart(fig_qoq, use_container_width=True)
        else:
            st.info("Data periode sebelumnya tidak tersedia untuk perbandingan Q-o-Q")
    
    st.markdown(f'<div class="narrative-box">{narratives.rekapitulasi_nib}</div>', 
                unsafe_allow_html=True)
    
    # Section 3: Per Kabupaten/Kota
    st.markdown('<div class="section-title">1.2 Rekapitulasi per Kabupaten/Kota</div>', 
                unsafe_allow_html=True)
    
    col1, col2 = st.columns([1.5, 1])
    
    with col1:
        df = st.session_state.aggregator.to_dataframe(report)
        if not df.empty:
            fig_kab = chart_gen.create_horizontal_bar_gradient(
                df,
                title="NIB per Kabupaten/Kota"
            )
            st.plotly_chart(fig_kab, use_container_width=True)
    
    with col2:
        if not df.empty:
            # Display table
            table_cols = ['Kabupaten/Kota', 'Total']
            for month in report.months_included:
                if month in df.columns:
                    table_cols.append(month)
            
            display_df = df[table_cols].head(15)
            st.dataframe(
                display_df,
                hide_index=True,
                use_container_width=True
            )
    
    st.markdown(f'<div class="narrative-box">{narratives.rekapitulasi_kab_kota}</div>', 
                unsafe_allow_html=True)
    
    # Section 4: Status PM
    st.markdown('<div class="section-title">1.3 Status Penanaman Modal</div>', 
                unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        pm_dist = stats.get('pm_distribution', {})
        fig_pm = chart_gen.create_pm_comparison_chart(
            pma_total=pm_dist.get('PMA', 0),
            pmdn_total=pm_dist.get('PMDN', 0)
        )
        st.plotly_chart(fig_pm, use_container_width=True)
    
    with col2:
        # PM table per kab/kota
        if not df.empty:
            pm_df = df[['Kabupaten/Kota', 'PMA', 'PMDN']].head(15)
            st.dataframe(pm_df, hide_index=True, use_container_width=True)
    
    st.markdown(f'<div class="narrative-box">{narratives.status_pm}</div>', 
                unsafe_allow_html=True)
    
    # Section 5: Pelaku Usaha
    st.markdown('<div class="section-title">1.4 Kategori Pelaku Usaha</div>', 
                unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        pelaku = stats.get('pelaku_usaha_distribution', {})
        fig_pelaku = chart_gen.create_pelaku_usaha_chart(
            umk_total=pelaku.get('UMK', 0),
            non_umk_total=pelaku.get('NON_UMK', 0)
        )
        st.plotly_chart(fig_pelaku, use_container_width=True)
    
    with col2:
        # Pelaku usaha table
        if not df.empty:
            pelaku_df = df[['Kabupaten/Kota', 'UMK', 'NON-UMK']].head(15)
            st.dataframe(pelaku_df, hide_index=True, use_container_width=True)
    
    st.markdown(f'<div class="narrative-box">{narratives.pelaku_usaha}</div>', 
                unsafe_allow_html=True)
    
    # Section 6: Sektor & Risiko (if data available)
    sektor_risiko_data = stats.get('sektor_risiko', {})
    if sektor_risiko_data:
        st.markdown('<div class="section-title">1.5 Perizinan Berdasarkan Risiko dan Sektor</div>', 
                    unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Risk distribution chart
            fig_risk = chart_gen.create_risk_donut_chart(
                rendah=sektor_risiko_data.get('risiko_rendah', 0),
                menengah_rendah=sektor_risiko_data.get('risiko_menengah_rendah', 0),
                menengah_tinggi=sektor_risiko_data.get('risiko_menengah_tinggi', 0),
                tinggi=sektor_risiko_data.get('risiko_tinggi', 0)
            )
            st.plotly_chart(fig_risk, use_container_width=True)
        
        with col2:
            # Sector distribution chart
            sector_data = {
                'Energi': sektor_risiko_data.get('sektor_energi', 0),
                'Kelautan': sektor_risiko_data.get('sektor_kelautan', 0),
                'Kesehatan': sektor_risiko_data.get('sektor_kesehatan', 0),
                'Komunikasi': sektor_risiko_data.get('sektor_komunikasi', 0),
                'Pariwisata': sektor_risiko_data.get('sektor_pariwisata', 0),
                'Perhubungan': sektor_risiko_data.get('sektor_perhubungan', 0),
                'Perindustrian': sektor_risiko_data.get('sektor_perindustrian', 0),
                'Pertanian': sektor_risiko_data.get('sektor_pertanian', 0),
            }
            # Filter out zeros
            sector_data = {k: v for k, v in sector_data.items() if v > 0}
            if sector_data:
                fig_sector = chart_gen.create_sector_distribution_chart(sector_data)
                st.plotly_chart(fig_sector, use_container_width=True)
        
        # Generate sektor risiko narrative
        sektor_narrative = generate_sektor_risiko_narrative(sektor_risiko_data)
        st.markdown(f'<div class="narrative-box">{sektor_narrative}</div>', 
                    unsafe_allow_html=True)
    
    # Section 7: Kesimpulan
    st.markdown('<div class="section-title">Kesimpulan</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="narrative-box">{narratives.kesimpulan}</div>', 
                unsafe_allow_html=True)


def render_export_section(report, stats):
    """Render export options."""
    st.divider()
    st.subheader("📥 Export Laporan")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if st.button("📄 Export PDF", use_container_width=True):
            try:
                with st.spinner("Generating PDF..."):
                    pdf_bytes = generate_pdf(report, stats)
                    
                st.download_button(
                    label="⬇️ Download PDF",
                    data=pdf_bytes,
                    file_name=f"Laporan_NIB_{report.period_name}_{report.year}.pdf",
                    mime="application/pdf",
                    key="pdf_download"
                )
            except ImportError as e:
                st.error(f"⚠️ {str(e)}")
            except Exception as e:
                st.error(f"⚠️ Error generating PDF: {str(e)}")
    
    with col2:
        if st.button("📝 Export Word", use_container_width=True):
            try:
                with st.spinner("Generating Word document..."):
                    docx_bytes = generate_word(report, stats)
                    
                st.download_button(
                    label="⬇️ Download Word",
                    data=docx_bytes,
                    file_name=f"Laporan_NIB_{report.period_name}_{report.year}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="word_download"
                )
            except ImportError as e:
                st.error(f"⚠️ {str(e)}")
            except Exception as e:
                st.error(f"⚠️ Error generating Word: {str(e)}")
    
    with col3:
        if st.button("📊 Export Excel Summary", use_container_width=True):
            df = st.session_state.aggregator.to_dataframe(report)
            if not df.empty:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Rekapitulasi')
                
                st.download_button(
                    label="⬇️ Download Excel",
                    data=output.getvalue(),
                    file_name=f"Rekapitulasi_NIB_{report.period_name}_{report.year}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )


def generate_pdf(report, stats) -> bytes:
    """Generate PDF report with charts and narratives."""
    from app.export.pdf_exporter import EnhancedPDFExporter as PDFExporter
    
    chart_gen = ChartGenerator()
    narrative_gen = NarrativeGenerator()
    
    # Generate narratives
    narratives = narrative_gen.generate_full_narrative(report, stats)
    
    # Generate chart images
    charts = {}
    
    # Monthly chart
    monthly_data = stats.get('monthly_totals', {})
    if monthly_data:
        fig = chart_gen.create_monthly_bar_with_trendline(monthly_data, show_trendline=True)
        charts['monthly'] = fig.to_image(format='png', scale=2)
    
    # Kab/Kota chart
    df = st.session_state.aggregator.to_dataframe(report)
    if not df.empty:
        fig = chart_gen.create_horizontal_bar_gradient(df, title="NIB per Kabupaten/Kota")
        charts['kab_kota'] = fig.to_image(format='png', scale=2)
    
    # PM chart
    pm_dist = stats.get('pm_distribution', {})
    fig = chart_gen.create_pm_comparison_chart(
        pma_total=pm_dist.get('PMA', 0),
        pmdn_total=pm_dist.get('PMDN', 0)
    )
    charts['pm'] = fig.to_image(format='png', scale=2)
    
    # Pelaku usaha chart
    pelaku = stats.get('pelaku_usaha_distribution', {})
    fig = chart_gen.create_pelaku_usaha_chart(
        umk_total=pelaku.get('UMK', 0),
        non_umk_total=pelaku.get('NON_UMK', 0)
    )
    charts['pelaku'] = fig.to_image(format='png', scale=2)
    
    # Risk and Sector charts (if data available)
    sektor_risiko = stats.get('sektor_risiko', {})
    if sektor_risiko:
        # Risk donut chart
        fig = chart_gen.create_risk_donut_chart(
            rendah=sektor_risiko.get('risiko_rendah', 0),
            menengah_rendah=sektor_risiko.get('risiko_menengah_rendah', 0),
            menengah_tinggi=sektor_risiko.get('risiko_menengah_tinggi', 0),
            tinggi=sektor_risiko.get('risiko_tinggi', 0)
        )
        charts['risk'] = fig.to_image(format='png', scale=2)
        
        # Sector distribution chart
        sector_data = {
            'Energi': sektor_risiko.get('sektor_energi', 0),
            'Kelautan': sektor_risiko.get('sektor_kelautan', 0),
            'Kesehatan': sektor_risiko.get('sektor_kesehatan', 0),
            'Komunikasi': sektor_risiko.get('sektor_komunikasi', 0),
            'Pariwisata': sektor_risiko.get('sektor_pariwisata', 0),
            'Perhubungan': sektor_risiko.get('sektor_perhubungan', 0),
            'Perindustrian': sektor_risiko.get('sektor_perindustrian', 0),
            'Pertanian': sektor_risiko.get('sektor_pertanian', 0),
        }
        # Filter zeros
        sector_data = {k: v for k, v in sector_data.items() if v > 0}
        if sector_data:
            fig = chart_gen.create_sector_distribution_chart(sector_data)
            charts['sector'] = fig.to_image(format='png', scale=2)
    
    # Create PDF exporter
    exporter = PDFExporter(logo_path=LOGO_PATH)
    
    if not exporter.is_available():
        raise ImportError("ReportLab tidak terinstall. Jalankan: pip install reportlab")
    
    return exporter.export_report(report, stats, narratives, charts)


def generate_word(report, stats) -> bytes:
    """Generate Word document with charts and narratives."""
    from app.export.docx_exporter import WordExporter
    
    chart_gen = ChartGenerator()
    narrative_gen = NarrativeGenerator()
    
    # Generate narratives
    narratives = narrative_gen.generate_full_narrative(report, stats)
    
    # Generate chart images
    charts = {}
    
    # Monthly chart
    monthly_data = stats.get('monthly_totals', {})
    if monthly_data:
        fig = chart_gen.create_monthly_bar_with_trendline(monthly_data, show_trendline=True)
        charts['monthly'] = fig.to_image(format='png', scale=2)
    
    # Kab/Kota chart
    df = st.session_state.aggregator.to_dataframe(report)
    if not df.empty:
        fig = chart_gen.create_horizontal_bar_gradient(df, title="NIB per Kabupaten/Kota")
        charts['kab_kota'] = fig.to_image(format='png', scale=2)
    
    # PM chart
    pm_dist = stats.get('pm_distribution', {})
    fig = chart_gen.create_pm_comparison_chart(
        pma_total=pm_dist.get('PMA', 0),
        pmdn_total=pm_dist.get('PMDN', 0)
    )
    charts['pm'] = fig.to_image(format='png', scale=2)
    
    # Pelaku usaha chart
    pelaku = stats.get('pelaku_usaha_distribution', {})
    fig = chart_gen.create_pelaku_usaha_chart(
        umk_total=pelaku.get('UMK', 0),
        non_umk_total=pelaku.get('NON_UMK', 0)
    )
    charts['pelaku'] = fig.to_image(format='png', scale=2)
    
    # Risk and Sector charts (if data available)
    sektor_risiko = stats.get('sektor_risiko', {})
    if sektor_risiko:
        # Risk donut chart
        fig = chart_gen.create_risk_donut_chart(
            rendah=sektor_risiko.get('risiko_rendah', 0),
            menengah_rendah=sektor_risiko.get('risiko_menengah_rendah', 0),
            menengah_tinggi=sektor_risiko.get('risiko_menengah_tinggi', 0),
            tinggi=sektor_risiko.get('risiko_tinggi', 0)
        )
        charts['risk'] = fig.to_image(format='png', scale=2)
        
        # Sector distribution chart
        sector_data = {
            'Energi': sektor_risiko.get('sektor_energi', 0),
            'Kelautan': sektor_risiko.get('sektor_kelautan', 0),
            'Kesehatan': sektor_risiko.get('sektor_kesehatan', 0),
            'Komunikasi': sektor_risiko.get('sektor_komunikasi', 0),
            'Pariwisata': sektor_risiko.get('sektor_pariwisata', 0),
            'Perhubungan': sektor_risiko.get('sektor_perhubungan', 0),
            'Perindustrian': sektor_risiko.get('sektor_perindustrian', 0),
            'Pertanian': sektor_risiko.get('sektor_pertanian', 0),
        }
        # Filter zeros
        sector_data = {k: v for k, v in sector_data.items() if v > 0}
        if sector_data:
            fig = chart_gen.create_sector_distribution_chart(sector_data)
            charts['sector'] = fig.to_image(format='png', scale=2)
    
    # Create Word exporter
    exporter = WordExporter(logo_path=LOGO_PATH)
    
    if not exporter.is_available():
        raise ImportError("python-docx tidak terinstall. Jalankan: pip install python-docx")
    
    return exporter.export_report(report, stats, narratives, charts)


def main():
    """Main application entry point."""
    init_session_state()
    render_header()
    jenis_periode, periode, tahun = render_sidebar()
    
    # Main content
    if st.session_state.report and st.session_state.stats:
        render_metrics(st.session_state.stats)
        st.divider()
        render_report(st.session_state.report, st.session_state.stats)
        render_export_section(st.session_state.report, st.session_state.stats)
    else:
        # Welcome message
        st.info("""
        👋 **Selamat datang di Sistem Laporan Otomatis DPMPTSP Provinsi Lampung!**
        
        Untuk memulai:
        1. Upload file Excel data bulanan di sidebar kiri
        2. Pilih periode laporan (Triwulan/Semester/Tahunan)
        3. Klik tombol "Generate Laporan"
        
        Sistem akan menghasilkan laporan lengkap dengan grafik dan narasi secara otomatis.
        """)
        
        # Show demo with existing data
        if st.button("🎮 Demo dengan Data Contoh"):
            # Load sample data from existing files
            data_dir = Path(__file__).parent.parent / "DATA OSS 2025" / "TW III"
            if data_dir.exists():
                sample_files = list(data_dir.glob("OLAH DATA OSS BULAN *.xlsx"))
                if sample_files:
                    st.session_state.aggregator = DataAggregator()
                    loader = DataLoader()
                    
                    for file_path in sample_files:
                        data = loader.load_monthly_data(file_path)
                        month = data.get('month')
                        if month:
                            st.session_state.aggregator.loaded_data[month] = data
                    
                    report = st.session_state.aggregator.aggregate_triwulan("TW III", 2025)
                    stats = st.session_state.aggregator.get_summary_stats(report)
                    
                    st.session_state.report = report
                    st.session_state.stats = stats
                    st.rerun()


if __name__ == "__main__":
    main()
