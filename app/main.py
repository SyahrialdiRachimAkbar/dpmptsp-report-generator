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

from app.data.loader import DataLoader, InvestmentReport, InvestmentData, TWSummary
from app.data.aggregator import DataAggregator, PeriodReport, AggregatedNIBData
from app.data.reference_loader import ReferenceDataLoader
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

# Custom CSS - Modern UI/UX Design with Dark Theme Support
st.markdown("""
<style>
    /* ===== ROOT VARIABLES - LIGHT MODE ===== */
    :root {
        --primary-color: #1e3a5f;
        --primary-light: #2d5a87;
        --secondary-color: #3d7ea6;
        --accent-color: #5cb85c;
        --warning-color: #f0ad4e;
        --danger-color: #d9534f;
        --background-light: #f0f4f8;
        --background-gradient-start: #f0f4f8;
        --background-gradient-end: #e8eef3;
        --card-bg: rgba(255, 255, 255, 0.95);
        --card-border: rgba(255, 255, 255, 0.8);
        --text-primary: #2c3e50;
        --text-secondary: #6c757d;
        --shadow-soft: 0 4px 20px rgba(0, 0, 0, 0.08);
        --shadow-hover: 0 8px 30px rgba(0, 0, 0, 0.12);
        --gradient-primary: linear-gradient(135deg, #1e3a5f 0%, #3d7ea6 100%);
        --gradient-accent: linear-gradient(135deg, #5cb85c 0%, #3d9e52 100%);
        --border-radius: 12px;
        --transition: all 0.3s ease;
        --divider-gradient: linear-gradient(90deg, transparent, var(--secondary-color), transparent);
        --narrative-border: var(--primary-color);
        --table-header-bg: #f8f9fa;
        --table-row-hover: #f1f3f4;
        --input-bg: #ffffff;
        --input-border: #dee2e6;
    }
    
    /* ===== DARK MODE VARIABLES ===== */
    @media (prefers-color-scheme: dark) {
        :root {
            --primary-color: #4da6d9;
            --primary-light: #6bb8e6;
            --secondary-color: #5cbddb;
            --accent-color: #6fcf6f;
            --warning-color: #ffcc66;
            --danger-color: #ff6b6b;
            --background-light: #1a1d23;
            --background-gradient-start: #1a1d23;
            --background-gradient-end: #12151a;
            --card-bg: rgba(30, 35, 45, 0.95);
            --card-border: rgba(60, 70, 90, 0.6);
            --text-primary: #e8eaed;
            --text-secondary: #9aa0a8;
            --shadow-soft: 0 4px 20px rgba(0, 0, 0, 0.35);
            --shadow-hover: 0 8px 30px rgba(0, 0, 0, 0.5);
            --gradient-primary: linear-gradient(135deg, #2d5a87 0%, #4da6d9 100%);
            --gradient-accent: linear-gradient(135deg, #4caf50 0%, #6fcf6f 100%);
            --divider-gradient: linear-gradient(90deg, transparent, var(--secondary-color), transparent);
            --narrative-border: var(--secondary-color);
            --table-header-bg: #252a33;
            --table-row-hover: #2d323d;
            --input-bg: #252a33;
            --input-border: #3d4450;
        }
    }
    
    /* ===== GLOBAL STYLES ===== */
    .stApp {
        background: linear-gradient(180deg, var(--background-gradient-start) 0%, var(--background-gradient-end) 100%);
    }
    
    /* Dark mode main content area */
    @media (prefers-color-scheme: dark) {
        .stApp {
            color: var(--text-primary);
        }
        
        .stApp [data-testid="stAppViewContainer"] {
            background: linear-gradient(180deg, var(--background-gradient-start) 0%, var(--background-gradient-end) 100%);
        }
        
        /* Streamlit native elements text color */
        .stApp p, .stApp span, .stApp label, .stApp div {
            color: var(--text-primary);
        }
        
        /* Info/Warning/Success boxes */
        .stAlert {
            background: var(--card-bg) !important;
            border: 1px solid var(--card-border) !important;
        }
        
        .stAlert p {
            color: var(--text-primary) !important;
        }
        
        /* DataFrames / Tables */
        [data-testid="stDataFrame"] {
            background: var(--card-bg);
        }
        
        [data-testid="stDataFrame"] th {
            background: var(--table-header-bg) !important;
            color: var(--text-primary) !important;
        }
        
        [data-testid="stDataFrame"] td {
            background: var(--card-bg) !important;
            color: var(--text-primary) !important;
        }
        
        [data-testid="stDataFrame"] tr:hover td {
            background: var(--table-row-hover) !important;
        }
        
        /* Streamlit DataFrame with glide-data-grid */
        [data-testid="stDataFrameResizable"],
        [data-testid="stDataFrame"] > div,
        .dvn-scroller,
        .dvn-underlay,
        .dvn-scroll-inner {
            background: var(--card-bg) !important;
        }
        
        /* Glide Data Editor (Streamlit's table component) */
        .glideDataEditor,
        .gdg-style {
            background: var(--card-bg) !important;
        }
        
        canvas + div {
            background: var(--card-bg) !important; 
        }
        
        /* Expander */
        .streamlit-expanderHeader {
            background: var(--card-bg) !important;
            color: var(--text-primary) !important;
        }
        
        .streamlit-expanderContent {
            background: var(--card-bg) !important;
            color: var(--text-primary) !important;
        }
        
        /* Select boxes and inputs */
        .stSelectbox > div > div,
        .stMultiSelect > div > div,
        .stTextInput > div > div > input {
            background: var(--input-bg) !important;
            color: var(--text-primary) !important;
            border-color: var(--input-border) !important;
        }
        
        /* Radio buttons and checkboxes */
        .stRadio label, .stCheckbox label {
            color: var(--text-primary) !important;
        }
        
        /* Plotly charts - transparent background */
        .js-plotly-plot .plotly .main-svg {
            background: transparent !important;
        }
        
        /* Legend text in Plotly */
        .js-plotly-plot .plotly .legend text {
            fill: var(--text-primary) !important;
        }
        
        /* Chart annotations */
        .js-plotly-plot .annotation-text {
            fill: var(--text-primary) !important;
        }
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
    
    /* Dark mode header override */
    @media (prefers-color-scheme: dark) {
        .main-header {
            background: linear-gradient(135deg, #6bb8e6 0%, #4da6d9 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
        }
    }
    
    /* ===== METRIC CARDS - GLASSMORPHISM ===== */
    .metric-card {
        background: var(--card-bg);
        backdrop-filter: blur(10px);
        border-radius: var(--border-radius);
        padding: 1.5rem;
        box-shadow: var(--shadow-soft);
        border: 1px solid var(--card-border);
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
        border: 1px solid var(--card-border);
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
        color: var(--accent-color);
    }
    
    .metric-delta.negative {
        background: rgba(217, 83, 79, 0.15);
        color: var(--danger-color);
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
    
    /* ===== STYLED HTML TABLES ===== */
    .styled-table-container {
        overflow-x: auto;
        border-radius: var(--border-radius);
        box-shadow: var(--shadow-soft);
        margin: 0.5rem 0;
    }
    
    .styled-table {
        width: 100%;
        border-collapse: collapse;
        background: var(--card-bg);
        font-size: 0.85rem;
    }
    
    .styled-table thead th {
        background: var(--table-header-bg);
        color: var(--text-primary);
        padding: 0.75rem 0.5rem;
        text-align: left;
        font-weight: 600;
        border-bottom: 2px solid var(--secondary-color);
        white-space: nowrap;
    }
    
    .styled-table tbody td {
        padding: 0.6rem 0.5rem;
        border-bottom: 1px solid var(--card-border);
        color: var(--text-primary);
    }
    
    .styled-table tbody tr:hover {
        background: var(--table-row-hover);
    }
    
    .styled-table tbody tr:last-child td {
        border-bottom: none;
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

<script>
// Detect dark mode preference and apply styles to DataFrame components
(function() {
    function applyDarkModeToDataFrames() {
        const isDarkMode = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
        
        if (isDarkMode) {
            // Apply dark background to DataFrame containers
            const dataFrames = document.querySelectorAll('[data-testid="stDataFrame"], [data-testid="stDataFrameResizable"]');
            dataFrames.forEach(df => {
                df.style.backgroundColor = 'rgba(30, 35, 45, 0.95)';
                df.style.borderRadius = '12px';
                df.style.overflow = 'hidden';
                
                // Apply to all child divs
                const childDivs = df.querySelectorAll('div');
                childDivs.forEach(div => {
                    if (!div.querySelector('canvas')) {
                        div.style.backgroundColor = 'rgba(30, 35, 45, 0.95)';
                    }
                });
            });
            
            // Apply to canvas containers with a slight delay for Streamlit rendering
            setTimeout(() => {
                const canvasContainers = document.querySelectorAll('[data-testid="stDataFrame"] > div > div');
                canvasContainers.forEach(container => {
                    container.style.backgroundColor = 'rgba(30, 35, 45, 0.95)';
                });
            }, 100);
        }
    }
    
    // Run on load
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', applyDarkModeToDataFrames);
    } else {
        applyDarkModeToDataFrames();
    }
    
    // Re-run when new content is added (Streamlit re-renders)
    const observer = new MutationObserver(function(mutations) {
        applyDarkModeToDataFrames();
    });
    
    observer.observe(document.body, { childList: true, subtree: true });
    
    // Listen for theme changes
    if (window.matchMedia) {
        window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', applyDarkModeToDataFrames);
    }
})();
</script>
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
    if 'investment_reports' not in st.session_state:
        st.session_state.investment_reports = None  # Dict[str, InvestmentReport]
    if 'investment_file' not in st.session_state:
        st.session_state.investment_file = None
    if 'prev_year_investment_file' not in st.session_state:
        st.session_state.prev_year_investment_file = None  # For Y-o-Y comparison
    if 'tw_summary' not in st.session_state:
        st.session_state.tw_summary = None  # Dict[str, TWSummary] for current year
    if 'prev_year_tw_summary' not in st.session_state:
        st.session_state.prev_year_tw_summary = None  # Dict[str, TWSummary] for previous year


def df_to_html_table(df: pd.DataFrame, max_rows: int = 15) -> str:
    """Convert DataFrame to styled HTML table for dark mode compatibility.
    
    Args:
        df: DataFrame to convert
        max_rows: Maximum number of rows to display
        
    Returns:
        HTML string of styled table
    """
    display_df = df.head(max_rows)
    
    # Build HTML table
    html = '<div class="styled-table-container"><table class="styled-table">'
    
    # Header
    html += '<thead><tr>'
    for col in display_df.columns:
        html += f'<th>{col}</th>'
    html += '</tr></thead>'
    
    # Body
    html += '<tbody>'
    for _, row in display_df.iterrows():
        html += '<tr>'
        for val in row:
            if isinstance(val, (int, float)):
                formatted_val = f"{val:,.0f}".replace(",", ".")
            else:
                formatted_val = str(val)
            html += f'<td>{formatted_val}</td>'
        html += '</tr>'
    html += '</tbody>'
    
    html += '</table></div>'
    return html


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
        st.header("📁 Upload Data Reference")
        
        # 1. Upload NIB
        nib_file = st.file_uploader(
            "Upload File NIB (.xlsx)",
            type=['xlsx', 'xls'],
            key="nib_uploader",
            help="File yang berisi data NIB (Sheet 1)"
        )
        if nib_file:
            st.session_state.nib_ref_file = nib_file
            st.success(f"✅ NIB: {nib_file.name}")
            
        # 2. Upload PB OSS
        pb_oss_file = st.file_uploader(
            "Upload File PB OSS (.xlsx)",
            type=['xlsx', 'xls'],
            key="pb_oss_uploader",
            help="File yang berisi data Perizinan Berusaha (RISIKO/SEKTOR/Sheet 1)"
        )
        if pb_oss_file:
            st.session_state.pb_oss_ref_file = pb_oss_file
            st.success(f"✅ PB OSS: {pb_oss_file.name}")
            
        # 3. Upload PROYEK
        proyek_file = st.file_uploader(
            "Upload File PROYEK (.xlsx)",
            type=['xlsx', 'xls'],
            key="proyek_uploader", 
            help="File yang berisi data Realisasi Investasi"
        )
        if proyek_file:
            st.session_state.proyek_ref_file = proyek_file
            st.success(f"✅ PROYEK: {proyek_file.name}")
            
        st.divider()
        
        # Period selection
        st.header("📅 Pilih Periode")
        
        # Year selection (Auto-detect from files if possible, else manual)
        detected_year = 2025
        # Try to detect from files
        loader = ReferenceDataLoader()
        if st.session_state.get('nib_ref_file'):
            y = loader.extract_year_from_filename(st.session_state.nib_ref_file.name)
            if y: detected_year = y
        elif st.session_state.get('proyek_ref_file'):
            y = loader.extract_year_from_filename(st.session_state.proyek_ref_file.name)
            if y: detected_year = y
            
        tahun_options = [detected_year] + [y for y in [2026, 2025, 2024, 2023] if y != detected_year]
        
        tahun = st.selectbox("Tahun", options=tahun_options)
        
        jenis_periode = st.radio(
            "Jenis Periode",
            options=["Triwulan", "Semester", "Tahunan"],
            index=0
        )
        
        periode = str(tahun)
        if jenis_periode == "Triwulan":
            periode = st.selectbox("Pilih Triwulan", options=["TW I", "TW II", "TW III", "TW IV"])
        elif jenis_periode == "Semester":
            periode = st.selectbox("Pilih Semester", options=["Semester I", "Semester II"])
            
        st.divider()
        
        if st.button("🚀 Generate Laporan", type="primary", use_container_width=True):
             # Check if at least one file is uploaded
             if not (st.session_state.get('nib_ref_file') or 
                     st.session_state.get('pb_oss_ref_file') or 
                     st.session_state.get('proyek_ref_file')):
                 st.error("⚠️ Upload minimal satu file referensi!")
             else:
                 with st.spinner("Memproses data..."):
                     # Pass empty list for uploaded_files since we use session_state logic now
                     success = process_data([], jenis_periode, periode, tahun)
                 
                 if success:
                     st.success("✅ Laporan berhasil dibuat!")
                     st.rerun()
                 else:
                     st.error("❌ Gagal membuat laporan. Periksa pesan error di atas.")
                 
        # Clear button
        if st.button("🗑️ Clear Data", use_container_width=True):
            cols_to_clear = ['nib_ref_file', 'pb_oss_ref_file', 'proyek_ref_file', 
                             'report', 'stats', 'aggregator', 'investment_reports', 
                             'tw_summary', 'prev_year_tw_summary']
            for col in cols_to_clear:
                if col in st.session_state:
                    del st.session_state[col]
            # Re-init basic state
            st.session_state.loaded_files = [] # Keep for legacy compatibility if needed
            if 'aggregator' not in st.session_state:
                st.session_state.aggregator = DataAggregator()
            st.rerun()
            
        return jenis_periode, periode, tahun


@st.cache_data(show_spinner=False)
def _cached_load_nib(file_content: bytes, filename: str, year: int):
    """Cached NIB loader - only reloads when file content changes."""
    from io import BytesIO
    loader = ReferenceDataLoader()
    return loader.load_nib(BytesIO(file_content), filename, year)

@st.cache_data(show_spinner=False)
def _cached_load_pb_oss(file_content: bytes, filename: str, year: int):
    """Cached PB OSS loader - only reloads when file content changes."""
    from io import BytesIO
    loader = ReferenceDataLoader()
    return loader.load_pb_oss(BytesIO(file_content), filename, year)

@st.cache_data(show_spinner=False)
def _cached_load_proyek(file_content: bytes, filename: str, year: int):
    """Cached PROYEK loader - only reloads when file content changes."""
    from io import BytesIO
    loader = ReferenceDataLoader()
    return loader.load_proyek(BytesIO(file_content), filename, year)


def process_data(uploaded_files, jenis_periode: str, periode: str, tahun: int):
    """Process uploaded reference files and generate report."""
    loader = ReferenceDataLoader()
    aggregator = DataAggregator()
    
    # Initialize containers
    report = None
    stats = {}
    
    # Determine months included in the period
    months = loader.get_months_for_period(jenis_periode, periode)
    
    # 1. Process NIB Data (if uploaded)
    nib_file = st.session_state.get('nib_ref_file')
    if nib_file:
        try:
            # Use cached loader for performance
            nib_data = _cached_load_nib(nib_file.getvalue(), nib_file.name, tahun)
            
            if nib_data:
                # Create PeriodReport structure manually
                report = PeriodReport(
                    period_type=jenis_periode,
                    period_name=periode,
                    year=tahun,
                    months_included=months
                )
                
                # Populate monthly totals
                for m in months:
                    report.monthly_totals[m] = nib_data.monthly_totals.get(m, 0)
                
                # Aggregate totals
                report.total_nib = nib_data.get_period_total(months)
                
                pm_totals = nib_data.get_period_by_pm_status(months)
                report.total_pma = pm_totals.get('PMA', 0)
                report.total_pmdn = pm_totals.get('PMDN', 0)
                
                skala_totals = nib_data.get_period_by_skala_usaha(months)
                # Map various spellings if needed
                for k, v in skala_totals.items():
                    k_lower = k.lower()
                    if 'mikro' in k_lower: report.total_umk += v
                    elif 'kecil' in k_lower: report.total_umk += v
                    elif 'menengah' in k_lower: report.total_non_umk += v
                    elif 'besar' in k_lower: report.total_non_umk += v
                
                # Populate data_by_location (AggregatedNIBData)
                # Iterate over all known locations from data
                all_locations = set(nib_data.by_kab_kota.keys())
                
                for kab in all_locations:
                    agg_data = AggregatedNIBData(kabupaten_kota=kab)
                    
                    # Period total
                    agg_data.grand_total = sum(nib_data.by_kab_kota[kab].get(m, 0) for m in months)
                    
                    # Monthly breakdown
                    for m in months:
                        agg_data.period_data[m] = nib_data.by_kab_kota[kab].get(m, 0)
                    
                    # PM breakdown for this kab/kota (using new granular data)
                    if hasattr(nib_data, 'kab_pm_monthly') and kab in nib_data.kab_pm_monthly:
                        for m in months:
                            if m in nib_data.kab_pm_monthly[kab]:
                                for pm_status, count in nib_data.kab_pm_monthly[kab][m].items():
                                    if 'PMA' in str(pm_status).upper(): agg_data.pma_total += count
                                    elif 'PMDN' in str(pm_status).upper(): agg_data.pmdn_total += count
                    
                    # Skala breakdown for this kab/kota
                    if hasattr(nib_data, 'kab_skala_monthly') and kab in nib_data.kab_skala_monthly:
                        for m in months:
                            if m in nib_data.kab_skala_monthly[kab]:
                                for skala, count in nib_data.kab_skala_monthly[kab][m].items():
                                    s_lower = str(skala).lower()
                                    if 'mikro' in s_lower: agg_data.usaha_mikro_total += count
                                    elif 'kecil' in s_lower: agg_data.usaha_kecil_total += count
                                    elif 'menengah' in s_lower: agg_data.usaha_menengah_total += count
                                    elif 'besar' in s_lower: agg_data.usaha_besar_total += count
                    
                    report.data_by_location[kab] = agg_data
                    
                # Generate base stats
                stats = aggregator.get_summary_stats(report)
                
        except Exception as e:
            st.error(f"Error loading NIB file: {str(e)}")
            print(f"Detailed error NIB: {e}")
            
    # 2. Process PB OSS Data (if uploaded)
    pb_file = st.session_state.get('pb_oss_ref_file')
    if pb_file:
        try:
            # Use cached loader for performance
            pb_data = _cached_load_pb_oss(pb_file.getvalue(), pb_file.name, tahun)
            
            if pb_data:
                # Get risk and sector distribution for selected period
                risk_dist = pb_data.get_period_risk(months)
                sector_dist = pb_data.get_period_sector(months)
                
                # Map to stats structure expected by charts
                sektor_totals = {
                    'risiko_rendah': risk_dist.get('Rendah', 0),
                    'risiko_menengah_rendah': risk_dist.get('Menengah Rendah', 0),
                    'risiko_menengah_tinggi': risk_dist.get('Menengah Tinggi', 0),
                    'risiko_tinggi': risk_dist.get('Tinggi', 0),
                    'total': sum(risk_dist.values())
                }
                
                # Add sector specific keys if available (simple mapping)
                for sector, count in sector_dist.items():
                    # Sanitize key for convenience
                    key = 'sektor_' + sector.lower().split()[0] # e.g. sektor_pertanian
                    sektor_totals[key] = count
                
                stats['sektor_risiko'] = sektor_totals
                
        except Exception as e:
            st.warning(f"Error loading PB OSS file: {str(e)}")
            
    # 3. Process PROYEK Data (if uploaded)
    proyek_file = st.session_state.get('proyek_ref_file')
    if proyek_file:
        try:
            # Use cached loader for performance
            proyek_data = _cached_load_proyek(proyek_file.getvalue(), proyek_file.name, tahun)
            
            if proyek_data:
                investment_reports = {} # Dict[periode_name, InvestmentReport]
                tw_summary = {} # Dict[triwulan, TWSummary] -> needed for projections
                
                # Create InvestmentReport for the CURRENT period
                current_inv_report = InvestmentReport(
                    triwulan=periode,
                    year=tahun
                )
                
                # Populate data
                current_inv_report.pma_total = proyek_data.get_period_pma(months)
                current_inv_report.pmdn_total = proyek_data.get_period_pmdn(months)
                current_inv_report.pma_tki = proyek_data.get_period_tki(months) # Simplified labor assignment
                current_inv_report.pma_tka = proyek_data.get_period_tka(months)
                current_inv_report.pma_proyek = proyek_data.get_period_pma_projects(months)
                current_inv_report.pmdn_proyek = proyek_data.get_period_pmdn_projects(months)
                
                # Populate Wilayah breakdown (InvestmentData objects)
                wilayah_data = proyek_data.get_period_by_wilayah(months)
                pma_wil_list = []
                
                for wil, inv in wilayah_data.items():
                     # Create generic InvestmentData
                     inv_obj = InvestmentData(name=wil, jumlah_rp=inv)
                     # Add to PMA list (temporary hack until granular split available)
                     pma_wil_list.append(inv_obj)
                     
                current_inv_report.pma_by_wilayah = pma_wil_list
                # current_inv_report.pmdn_by_wilayah remains empty or we split logic
                
                investment_reports[periode] = current_inv_report
                st.session_state.investment_reports = investment_reports
                
                # Project projection (from TW summary)
                # Create dummy summary for sections using it
                summary_obj = TWSummary(triwulan=periode, year=tahun)
                summary_obj.proyek = proyek_data.get_period_projects(months)
                summary_obj.total_rp = current_inv_report.total_investasi
                tw_summary[periode] = summary_obj
                st.session_state.tw_summary = tw_summary
                
        except Exception as e:
            st.warning(f"Error loading PROYEK file: {str(e)}")
            print(f"Detailed Proyek error: {e}")

    # Set session state
    st.session_state.report = report
    st.session_state.stats = stats
    st.session_state.aggregator = aggregator
    
    return report is not None


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
    st.markdown('<div class="section-title">1. Nomor Induk Berusaha</div>', 
                unsafe_allow_html=True)
    
    # Summary metrics row
    total_nib = stats.get('total_nib', 0)
    pm_dist = stats.get('pm_distribution', {})
    pma_count = pm_dist.get('PMA', 0)
    pmdn_count = pm_dist.get('PMDN', 0)
    pelaku = stats.get('pelaku_usaha_distribution', {})
    umk_count = pelaku.get('UMK', 0)
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown(f'''
        <div class="metric-card">
            <div class="metric-value">{total_nib:,}</div>
            <div class="metric-label">Total NIB</div>
        </div>
        ''', unsafe_allow_html=True)
    with col2:
        st.markdown(f'''
        <div class="metric-card">
            <div class="metric-value">{pma_count:,}</div>
            <div class="metric-label">PMA (Asing)</div>
        </div>
        ''', unsafe_allow_html=True)
    with col3:
        st.markdown(f'''
        <div class="metric-card">
            <div class="metric-value">{pmdn_count:,}</div>
            <div class="metric-label">PMDN (Domestik)</div>
        </div>
        ''', unsafe_allow_html=True)
    with col4:
        st.markdown(f'''
        <div class="metric-card">
            <div class="metric-value">{umk_count:,}</div>
            <div class="metric-label">UMK</div>
        </div>
        ''', unsafe_allow_html=True)
    
    # 1.1 Rekapitulasi Data NIB
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
            
            display_df = df[table_cols]
            st.markdown(df_to_html_table(display_df), unsafe_allow_html=True)
    
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
            pm_df = df[['Kabupaten/Kota', 'PMA', 'PMDN']]
            st.markdown(df_to_html_table(pm_df), unsafe_allow_html=True)
    
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
            pelaku_df = df[['Kabupaten/Kota', 'UMK', 'NON-UMK']]
            st.markdown(df_to_html_table(pelaku_df), unsafe_allow_html=True)
    
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
    
    # Section: Realisasi Investasi (if data available)
    investment_reports = st.session_state.get('investment_reports', None)
    if investment_reports:
        st.markdown('<div class="section-title">2. Rencana Proyek</div>', 
                    unsafe_allow_html=True)
        
        # Get current period's investment data
        periode_name = report.period_name  # e.g., "TW I", "TW II"
        current_investment = investment_reports.get(periode_name)
        
        if current_investment:
            # Metrics for investment
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.markdown(f'''
                <div class="metric-card">
                    <div class="metric-value">Rp {current_investment.total_investasi/1e9:,.1f} M</div>
                    <div class="metric-label">Total Investasi</div>
                </div>
                ''', unsafe_allow_html=True)
            
            with col2:
                st.markdown(f'''
                <div class="metric-card">
                    <div class="metric-value">Rp {current_investment.pma_total/1e9:,.1f} M</div>
                    <div class="metric-label">PMA (Asing)</div>
                </div>
                ''', unsafe_allow_html=True)
            
            with col3:
                st.markdown(f'''
                <div class="metric-card">
                    <div class="metric-value">Rp {current_investment.pmdn_total/1e9:,.1f} M</div>
                    <div class="metric-label">PMDN (Domestik)</div>
                </div>
                ''', unsafe_allow_html=True)
            
            with col4:
                st.markdown(f'''
                <div class="metric-card">
                    <div class="metric-value">{current_investment.total_proyek:,}</div>
                    <div class="metric-label">Total Proyek</div>
                </div>
                ''', unsafe_allow_html=True)
            
            st.markdown('<div class="section-title">2.1 Rekapitulasi Data Proyek Berdasarkan Periode dan Kabupaten/Kota</div>', 
                        unsafe_allow_html=True)
            
            # Monthly project chart and Kab/Kota side by side
            proyek_file = st.session_state.get('proyek_ref_file')
            col1, col2 = st.columns(2)
            
            if proyek_file:
                from app.data.reference_loader import ReferenceDataLoader
                loader = ReferenceDataLoader()
                months = loader.get_months_for_period(report.period_type, report.period_name)
                proyek_data = _cached_load_proyek(proyek_file.getvalue(), proyek_file.name, report.year)
                
                if proyek_data:
                    with col1:
                        # Monthly project chart
                        monthly_project_data = {}
                        for month in months:
                            if month in proyek_data.monthly_projects:
                                monthly_project_data[month] = proyek_data.monthly_projects[month]
                        
                        if monthly_project_data:
                            fig_monthly_proj = chart_gen.create_monthly_bar_with_trendline(
                                monthly_project_data,
                                title="Jumlah Proyek per Bulan",
                                show_trendline=True
                            )
                            fig_monthly_proj.update_layout(height=400)
                            st.plotly_chart(fig_monthly_proj, use_container_width=True)
                    
                    with col2:
                        # Project count by Kab/Kota (horizontal bar)
                        import plotly.graph_objects as go
                        projects_by_kab = proyek_data.get_period_projects_by_wilayah(months)
                        if projects_by_kab:
                            sorted_kab = dict(sorted(projects_by_kab.items(), key=lambda x: x[1], reverse=True)[:15])
                            fig_kab = go.Figure(data=[go.Bar(
                                x=list(sorted_kab.values()), 
                                y=list(sorted_kab.keys()), 
                                orientation='h', 
                                marker_color='#3B82F6'
                            )])
                            fig_kab.update_layout(
                                title='Jumlah Proyek per Kabupaten/Kota',
                                template='plotly_dark',
                                height=400,
                                yaxis={'categoryorder': 'total ascending'}
                            )
                            st.plotly_chart(fig_kab, use_container_width=True)
                    
                    # Interpretation in Indonesian
                    total_proyek = sum(monthly_project_data.values()) if monthly_project_data else 0
                    top_month = max(monthly_project_data.items(), key=lambda x: x[1]) if monthly_project_data else ("", 0)
                    top_kab = list(sorted_kab.items())[0] if projects_by_kab and sorted_kab else ("", 0)
                    
                    interpretation = f"""
                    <b>Analisis dan Interpretasi:</b><br>
                    Rekapitulasi data proyek di Provinsi Lampung periode {report.period_name} Tahun {report.year} 
                    menunjukkan total sebanyak <b>{total_proyek:,}</b> proyek. 
                    Jumlah proyek tertinggi tercatat pada bulan <b>{top_month[0]}</b> dengan <b>{top_month[1]:,}</b> proyek. 
                    Berdasarkan lokasi, <b>{top_kab[0]}</b> mencatatkan jumlah proyek tertinggi sebanyak <b>{top_kab[1]:,}</b> proyek.
                    """
                    st.markdown(f'<div class="narrative-box">{interpretation}</div>', unsafe_allow_html=True)
            
            st.markdown('<div class="section-title">2.2 Rekapitulasi Proyek Berdasarkan Status Penanaman Modal</div>', 
                        unsafe_allow_html=True)
            
            col1, col2 = st.columns(2)
            
            with col1:
                # PMA vs PMDN donut
                fig_pma_pmdn = chart_gen.create_pma_pmdn_comparison_chart(
                    current_investment.pma_total,
                    current_investment.pmdn_total
                )
                st.plotly_chart(fig_pma_pmdn, use_container_width=True)
            
            with col2:
                # Labor absorption
                fig_labor = chart_gen.create_labor_absorption_chart(
                    current_investment.total_tki,
                    current_investment.total_tka
                )
                st.plotly_chart(fig_labor, use_container_width=True)
            
            # Narratives for PMA/PMDN and Labor charts
            col1, col2 = st.columns(2)
            with col1:
                pma_pmdn_narr = narrative_gen.generate_pma_pmdn_comparison_narrative(
                    current_investment.pma_total,
                    current_investment.pmdn_total
                )
                if pma_pmdn_narr:
                    st.markdown(f'<div class="narrative-box">{pma_pmdn_narr}</div>', unsafe_allow_html=True)
            with col2:
                labor_narr = narrative_gen.generate_labor_narrative(
                    current_investment.total_tki,
                    current_investment.total_tka
                )
                if labor_narr:
                    st.markdown(f'<div class="narrative-box">{labor_narr}</div>', unsafe_allow_html=True)
            
            # TW Comparison chart (if multiple TW data available)
            if len(investment_reports) > 1:
                st.markdown('<div class="section-title">2.3 Perbandingan Antar Triwulan (Investasi)</div>', 
                            unsafe_allow_html=True)
                fig_tw_comp = chart_gen.create_investment_tw_comparison_chart(investment_reports)
                st.plotly_chart(fig_tw_comp, use_container_width=True)
                
                # Narrative for TW comparison
                tw_comp_narr = narrative_gen.generate_tw_comparison_narrative(investment_reports)
                if tw_comp_narr:
                    st.markdown(f'<div class="narrative-box">{tw_comp_narr}</div>', unsafe_allow_html=True)
            
            # Investment Narrative Interpretation
            st.markdown('<div class="section-title">2.4 Interpretasi Data Investasi</div>', 
                        unsafe_allow_html=True)
            tw_summary_data = st.session_state.get('tw_summary', None)
            prev_year_data = st.session_state.get('prev_year_tw_summary', None)
            investment_narrative = narrative_gen.generate_investment_narrative(
                report=report,
                current_investment=current_investment,
                tw_summary=tw_summary_data,
                prev_year_summary=prev_year_data
            )
            st.markdown(f'<div class="narrative-box">{investment_narrative}</div>', 
                        unsafe_allow_html=True)
        else:
            st.info(f"Data investasi untuk {periode_name} tidak tersedia dalam file yang diupload.")
    
    # Section: Rencana Proyek (if TW summary data available)
    tw_summary = st.session_state.get('tw_summary', None)
    if tw_summary:
        # This section content is now part of Section 2
        # Removing old Section 3 header since it's merged into Section 2
        
        # Get current period's summary
        periode_name = report.period_name
        current_summary = tw_summary.get(periode_name)
        
        if current_summary:
            # 2.3 Skala Usaha visualization
            st.markdown('<div class="section-title">2.3 Rekapitulasi Data Proyek Berdasarkan Skala Usaha</div>', 
                        unsafe_allow_html=True)
            
            # Get proyek data
            proyek_data = None
            proyek_file = st.session_state.get('proyek_ref_file')
            if proyek_file:
                from app.data.reference_loader import ReferenceDataLoader
                loader = ReferenceDataLoader()
                months = loader.get_months_for_period(report.period_type, report.period_name)
                proyek_data = _cached_load_proyek(proyek_file.getvalue(), proyek_file.name, report.year)
                
                if proyek_data:
                    # Skala Usaha chart (from uraian_skala_usaha column)
                    skala_data = proyek_data.get_period_by_skala_usaha(months)
                    
                    if skala_data:
                        import plotly.graph_objects as go
                        
                        # Define order for skala categories
                        skala_order = ['Usaha Mikro', 'Usaha Kecil', 'Usaha Menengah', 'Usaha Besar']
                        ordered_data = {k: skala_data.get(k, 0) for k in skala_order if k in skala_data}
                        # Add any remaining categories
                        for k, v in skala_data.items():
                            if k not in ordered_data:
                                ordered_data[k] = v
                        
                        fig_skala = go.Figure(data=[
                            go.Bar(
                                x=list(ordered_data.keys()),
                                y=list(ordered_data.values()),
                                marker_color=['#10B981', '#3B82F6', '#8B5CF6', '#F59E0B'][:len(ordered_data)],
                                text=[f'{v:,}' for v in ordered_data.values()],
                                textposition='outside'
                            )
                        ])
                        fig_skala.update_layout(
                            title='Jumlah Proyek Berdasarkan Skala Usaha',
                            xaxis_title='Skala Usaha',
                            yaxis_title='Jumlah Proyek',
                            template='plotly_dark',
                            showlegend=False,
                            height=400
                        )
                        st.plotly_chart(fig_skala, use_container_width=True)
                        
                        # Interpretation for Skala Usaha
                        total_skala = sum(ordered_data.values())
                        top_skala = max(ordered_data.items(), key=lambda x: x[1]) if ordered_data else ("", 0)
                        interpretation_skala = f"""
                        <b>Analisis dan Interpretasi:</b><br>
                        Berdasarkan skala usaha, mayoritas proyek di Provinsi Lampung berada pada kategori 
                        <b>{top_skala[0]}</b> dengan jumlah <b>{top_skala[1]:,}</b> proyek ({top_skala[1]/total_skala*100:.1f}% dari total).
                        """
                        st.markdown(f'<div class="narrative-box">{interpretation_skala}</div>', unsafe_allow_html=True)
                    else:
                        st.info("Data skala usaha tidak tersedia dalam file PROYEK.")
            
            # 2.4 Jumlah Investasi visualization
            st.markdown('<div class="section-title">2.4 Rekapitulasi Data Proyek Berdasarkan Jumlah Investasi</div>', 
                        unsafe_allow_html=True)
            
            if proyek_file and proyek_data:
                # Investment by Kab/Kota
                import plotly.graph_objects as go
                inv_by_wilayah = proyek_data.get_period_by_wilayah(months)
                if inv_by_wilayah:
                    sorted_inv = dict(sorted(inv_by_wilayah.items(), key=lambda x: x[1], reverse=True)[:15])
                    fig_inv = go.Figure(data=[go.Bar(
                        x=list(sorted_inv.values()), 
                        y=list(sorted_inv.keys()), 
                        orientation='h', 
                        marker_color='#10B981'
                    )])
                    fig_inv.update_layout(
                        title='Jumlah Investasi per Kabupaten/Kota (Rupiah)',
                        template='plotly_dark',
                        height=400,
                        yaxis={'categoryorder': 'total ascending'}
                    )
                    st.plotly_chart(fig_inv, use_container_width=True)
                    
                    # Interpretation
                    top_inv = list(sorted_inv.items())[0] if sorted_inv else ("", 0)
                    total_inv = sum(sorted_inv.values())
                    interpretation_inv = f"""
                    <b>Analisis dan Interpretasi:</b><br>
                    <b>{top_inv[0]}</b> mencatatkan investasi tertinggi dengan nilai 
                    <b>Rp {top_inv[1]/1e9:,.2f} Miliar</b> ({top_inv[1]/total_inv*100:.1f}% dari total investasi).
                    """
                    st.markdown(f'<div class="narrative-box">{interpretation_inv}</div>', unsafe_allow_html=True)
            
            # 2.5 Tenaga Kerja visualization
            st.markdown('<div class="section-title">2.5 Rekapitulasi Data Proyek Berdasarkan Tenaga Kerja</div>', 
                        unsafe_allow_html=True)
            
            if proyek_file and proyek_data:
                # Labor (TKI+TKA) by Kab/Kota
                import plotly.graph_objects as go
                labor_by_wilayah = proyek_data.get_period_labor_by_wilayah(months)
                if labor_by_wilayah:
                    sorted_labor = dict(sorted(labor_by_wilayah.items(), key=lambda x: x[1], reverse=True)[:15])
                    fig_labor = go.Figure(data=[go.Bar(
                        x=list(sorted_labor.values()), 
                        y=list(sorted_labor.keys()), 
                        orientation='h', 
                        marker_color='#F59E0B'
                    )])
                    fig_labor.update_layout(
                        title='Jumlah Tenaga Kerja per Kabupaten/Kota',
                        template='plotly_dark',
                        height=400,
                        yaxis={'categoryorder': 'total ascending'},
                        xaxis_title='Jumlah Tenaga Kerja'
                    )
                    st.plotly_chart(fig_labor, use_container_width=True)
                    
                    # Interpretation
                    top_labor = list(sorted_labor.items())[0] if sorted_labor else ("", 0)
                    total_labor = sum(sorted_labor.values())
                    interpretation_labor = f"""
                    <b>Analisis dan Interpretasi:</b><br>
                    <b>{top_labor[0]}</b> mencatatkan penyerapan tenaga kerja tertinggi sebanyak 
                    <b>{top_labor[1]:,}</b> orang ({top_labor[1]/total_labor*100:.1f}% dari total {total_labor:,} tenaga kerja).
                    """
                    st.markdown(f'<div class="narrative-box">{interpretation_labor}</div>', unsafe_allow_html=True)
                else:
                    st.info("Data tenaga kerja tidak tersedia dalam file PROYEK.")
            
            # Q-o-Q Comparison (if previous TW exists)
            tw_order = ["TW I", "TW II", "TW III", "TW IV"]
            current_idx = tw_order.index(periode_name) if periode_name in tw_order else -1
            
            # Debug: show available TW data
            available_tws = list(tw_summary.keys()) if tw_summary else []
            
            if current_idx > 0:
                previous_tw = tw_order[current_idx - 1]
                previous_summary = tw_summary.get(previous_tw) if tw_summary else None
                
                if previous_summary:
                    st.markdown('<div class="section-title">3.4 Perbandingan Q-o-Q (Quarter-over-Quarter)</div>', 
                                unsafe_allow_html=True)
                    
                    # Calculate project counts for Q-o-Q
                    prev_inv = st.session_state.investment_reports.get(previous_tw) if st.session_state.investment_reports else None
                    prev_pma = prev_inv.pma_proyek if prev_inv else 0
                    prev_pmdn = prev_inv.pmdn_proyek if prev_inv else 0
                    
                    # Fallback estimation
                    if prev_pma == 0 and prev_pmdn == 0 and previous_summary.proyek > 0:
                        total_inv = previous_summary.pma_rp + previous_summary.pmdn_rp
                        if total_inv > 0:
                            prev_pma = int(previous_summary.proyek * (previous_summary.pma_rp / total_inv))
                            prev_pmdn = previous_summary.proyek - prev_pma
                    
                    fig_qoq = chart_gen.create_qoq_comparison_chart(
                        current_tw=periode_name,
                        current_data={"pma": pma_proyek, "pmdn": pmdn_proyek},
                        previous_tw=previous_tw,
                        previous_data={"pma": prev_pma, "pmdn": prev_pmdn}
                    )
                    st.plotly_chart(fig_qoq, use_container_width=True)
                    
                    # Q-o-Q Narrative
                    qoq_narr = narrative_gen.generate_qoq_narrative(
                        current_tw=periode_name,
                        current_proyek=current_summary.proyek,
                        prev_tw=previous_tw,
                        prev_proyek=previous_summary.proyek
                    )
                    if qoq_narr:
                        st.markdown(f'<div class="narrative-box">{qoq_narr}</div>', unsafe_allow_html=True)
                else:
                    st.info(f"Data {previous_tw} tidak tersedia untuk perbandingan Q-o-Q. Data yang tersedia: {available_tws}")
            else:
                if periode_name == "TW I":
                    st.info("Perbandingan Q-o-Q tidak tersedia untuk TW I (tidak ada triwulan sebelumnya).")
            
            # Y-o-Y Comparison (if previous year data available)
            prev_year_summary = st.session_state.get('prev_year_tw_summary', None)
            if prev_year_summary:
                prev_year_tw = prev_year_summary.get(periode_name)
                
                if prev_year_tw:
                    st.markdown('<div class="section-title">3.5 Perbandingan Y-o-Y (Year-over-Year)</div>', 
                                unsafe_allow_html=True)
                    
                    # Estimate project counts for previous year
                    prev_year_pma = 0
                    prev_year_pmdn = 0
                    if prev_year_tw.proyek > 0:
                        total_inv = prev_year_tw.pma_rp + prev_year_tw.pmdn_rp
                        if total_inv > 0:
                            prev_year_pma = int(prev_year_tw.proyek * (prev_year_tw.pma_rp / total_inv))
                            prev_year_pmdn = prev_year_tw.proyek - prev_year_pma
                    
                    fig_yoy = chart_gen.create_yoy_comparison_chart(
                        tw_name=periode_name,
                        current_year=current_summary.year,
                        current_data={"pma": pma_proyek, "pmdn": pmdn_proyek},
                        previous_year=prev_year_tw.year,
                        previous_data={"pma": prev_year_pma, "pmdn": prev_year_pmdn}
                    )
                    st.plotly_chart(fig_yoy, use_container_width=True)
                    
                    # Y-o-Y Narrative
                    yoy_narr = narrative_gen.generate_yoy_narrative(
                        tw_name=periode_name,
                        current_year=current_summary.year,
                        current_proyek=current_summary.proyek,
                        prev_year=prev_year_tw.year,
                        prev_proyek=prev_year_tw.proyek
                    )
                    if yoy_narr:
                        st.markdown(f'<div class="narrative-box">{yoy_narr}</div>', unsafe_allow_html=True)
            
            # Project Narrative Interpretation
            st.markdown('<div class="section-title">3.6 Interpretasi Data Proyek</div>', 
                        unsafe_allow_html=True)
            prev_year_data = st.session_state.get('prev_year_tw_summary', None)
            project_narrative = narrative_gen.generate_project_narrative(
                report=report,
                current_summary=current_summary,
                tw_summary=tw_summary,
                prev_year_summary=prev_year_data
            )
            st.markdown(f'<div class="narrative-box">{project_narrative}</div>', 
                        unsafe_allow_html=True)
        else:
            st.info(f"Data ringkasan proyek untuk {periode_name} tidak tersedia.")
    
    # ===========================================
    # Section 3: Perizinan Berusaha Berbasis Risiko (PB OSS data)
    # ===========================================
    pb_oss_file = st.session_state.get('pb_oss_ref_file')
    if pb_oss_file:
        st.markdown('<div class="section-title">3. Perizinan Berusaha Berbasis Risiko Provinsi Lampung</div>', 
                    unsafe_allow_html=True)
        
        # Load PB OSS data using cached loader
        pb_data = _cached_load_pb_oss(pb_oss_file.getvalue(), pb_oss_file.name, report.year)
        
        if pb_data:
            from app.data.reference_loader import ReferenceDataLoader
            loader = ReferenceDataLoader()
            months = loader.get_months_for_period(report.period_type, report.period_name)
            
            # 3.1 Period and Location
            st.markdown('<div class="section-title">3.1 Rekapitulasi Berdasarkan Periode dan Lokasi Usaha di Kabupaten/Kota</div>', 
                        unsafe_allow_html=True)
            kab_data = pb_data.get_period_by_kab_kota(months)
            if kab_data:
                import plotly.graph_objects as go
                sorted_kab = dict(sorted(kab_data.items(), key=lambda x: x[1], reverse=True)[:15])
                fig = go.Figure(data=[go.Bar(x=list(sorted_kab.values()), y=list(sorted_kab.keys()), orientation='h', marker_color='#3B82F6')])
                fig.update_layout(title='Perizinan per Kabupaten/Kota', template='plotly_dark', height=400, yaxis={'categoryorder': 'total ascending'})
                st.plotly_chart(fig, use_container_width=True)
            
            # 3.2 Status PM
            st.markdown('<div class="section-title">3.2 Rekapitulasi Berdasarkan Status Penanaman Modal</div>', 
                        unsafe_allow_html=True)
            pm_data = pb_data.get_period_status_pm(months)
            if pm_data:
                fig = go.Figure(data=[go.Bar(x=list(pm_data.keys()), y=list(pm_data.values()), marker_color=['#10B981', '#F59E0B'][:len(pm_data)])])
                fig.update_layout(title='Perizinan per Status PM', template='plotly_dark', height=350)
                st.plotly_chart(fig, use_container_width=True)
            
            # 3.3 Risk Level
            st.markdown('<div class="section-title">3.3 Rekapitulasi Berdasarkan Tingkat Risiko</div>', 
                        unsafe_allow_html=True)
            risk_data = pb_data.get_period_risk(months)
            if risk_data:
                fig = go.Figure(data=[go.Bar(x=list(risk_data.keys()), y=list(risk_data.values()), marker_color=['#22C55E', '#EAB308', '#EF4444', '#8B5CF6'][:len(risk_data)])])
                fig.update_layout(title='Perizinan per Tingkat Risiko', template='plotly_dark', height=350)
                st.plotly_chart(fig, use_container_width=True)
            
            # 3.4 Sector
            st.markdown('<div class="section-title">3.4 Rekapitulasi Berdasarkan Sektor Kementerian/Lembaga</div>', 
                        unsafe_allow_html=True)
            sector_data = pb_data.get_period_sector(months)
            if sector_data:
                sorted_sector = dict(sorted(sector_data.items(), key=lambda x: x[1], reverse=True)[:10])
                fig = go.Figure(data=[go.Bar(x=list(sorted_sector.values()), y=list(sorted_sector.keys()), orientation='h', marker_color='#8B5CF6')])
                fig.update_layout(title='Perizinan per Sektor (Top 10)', template='plotly_dark', height=400, yaxis={'categoryorder': 'total ascending'})
                st.plotly_chart(fig, use_container_width=True)
            
            # 3.5 Jenis Perizinan
            st.markdown('<div class="section-title">3.5 Rekapitulasi Berdasarkan Jenis Perizinan</div>', 
                        unsafe_allow_html=True)
            jenis_data = pb_data.get_period_jenis_perizinan(months)
            if jenis_data:
                sorted_jenis = dict(sorted(jenis_data.items(), key=lambda x: x[1], reverse=True)[:10])
                fig = go.Figure(data=[go.Bar(x=list(sorted_jenis.values()), y=list(sorted_jenis.keys()), orientation='h', marker_color='#06B6D4')])
                fig.update_layout(title='Perizinan per Jenis (Top 10)', template='plotly_dark', height=400, yaxis={'categoryorder': 'total ascending'})
                st.plotly_chart(fig, use_container_width=True)
            
            # 3.6 Status Perizinan (NO Gubernur filter - all data)
            st.markdown('<div class="section-title">3.6 Rekapitulasi Berdasarkan Status Respon</div>', 
                        unsafe_allow_html=True)
            status_data = pb_data.get_period_status_perizinan(months)
            if status_data:
                fig = go.Figure(data=[go.Pie(labels=list(status_data.keys()), values=list(status_data.values()), hole=0.4)])
                fig.update_layout(title='Perizinan per Status Respon', template='plotly_dark', height=400)
                st.plotly_chart(fig, use_container_width=True)
            
            # 3.7 Kewenangan (NO Gubernur filter - all data)
            st.markdown('<div class="section-title">3.7 Rekapitulasi Berdasarkan Kewenangan</div>', 
                        unsafe_allow_html=True)
            kew_data = pb_data.get_period_kewenangan(months)
            if kew_data:
                fig = go.Figure(data=[go.Bar(x=list(kew_data.keys()), y=list(kew_data.values()), marker_color=['#3B82F6', '#10B981', '#F59E0B'][:len(kew_data)])])
                fig.update_layout(title='Perizinan per Kewenangan', template='plotly_dark', height=350)
                st.plotly_chart(fig, use_container_width=True)
    
    # Section: Kesimpulan
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
