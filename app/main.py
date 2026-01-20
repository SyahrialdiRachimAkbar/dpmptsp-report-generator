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
        
        # === Previous Year Files (Optional for Y-o-Y) ===
        with st.expander("📁 Previous Year (Y-o-Y)", expanded=False):
            st.caption("Upload previous year files for Year-over-Year comparison")
            
            # Previous Year NIB
            nib_prev_file = st.file_uploader(
                "NIB Previous Year (.xlsx)",
                type=['xlsx', 'xls'],
                key="nib_prev_uploader"
            )
            if nib_prev_file:
                st.session_state.nib_prev_ref_file = nib_prev_file
                st.success(f"✅ {nib_prev_file.name}")
            
            # Previous Year PB OSS
            pb_prev_file = st.file_uploader(
                "PB OSS Previous Year (.xlsx)",
                type=['xlsx', 'xls'],
                key="pb_oss_prev_uploader"
            )
            if pb_prev_file:
                st.session_state.pb_oss_prev_ref_file = pb_prev_file
                st.success(f"✅ {pb_prev_file.name}")
            
            # Previous Year PROYEK
            proyek_prev_file = st.file_uploader(
                "PROYEK Previous Year (.xlsx)",
                type=['xlsx', 'xls'],
                key="proyek_prev_uploader"
            )
            if proyek_prev_file:
                st.session_state.proyek_prev_ref_file = proyek_prev_file
                st.success(f"✅ {proyek_prev_file.name}")
            
        st.divider()
        
        # Period selection
        st.header("📅 Pilih Periode")
        
        # Year selection (Auto-detect from files if possible, else current year)
        from datetime import datetime
        current_year = datetime.now().year
        detected_year = current_year
        # Try to detect from files
        loader = ReferenceDataLoader()
        if st.session_state.get('nib_ref_file'):
            y = loader.extract_year_from_filename(st.session_state.nib_ref_file.name)
            if y: detected_year = y
        elif st.session_state.get('proyek_ref_file'):
            y = loader.extract_year_from_filename(st.session_state.proyek_ref_file.name)
            if y: detected_year = y
            
        # Generate year options dynamically
        year_range = [current_year + 1, current_year, current_year - 1, current_year - 2]
        tahun_options = [detected_year] + [y for y in year_range if y != detected_year]
        
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
                # Iterate through all quarters to populate available history
                from app.config import TRIWULAN_KE_BULAN
                
                for period_name, period_months in TRIWULAN_KE_BULAN.items():
                    # Calculate stats for this period
                    period_proyek_count = proyek_data.get_period_projects(period_months)
                    
                    if period_proyek_count > 0:
                        # Create summary if data exists
                        sum_obj = TWSummary(triwulan=period_name, year=tahun)
                        sum_obj.proyek = period_proyek_count
                        
                        # Populate investment values
                        curr_pma_val = proyek_data.get_period_pma(period_months)
                        curr_pmdn_val = proyek_data.get_period_pmdn(period_months)
                        
                        sum_obj.pma_rp = curr_pma_val
                        sum_obj.pmdn_rp = curr_pmdn_val
                        sum_obj.total_rp = curr_pma_val + curr_pmdn_val
                        
                        # Populate other fields if needed (Labor etc.)
                        # sum_obj.tki = ... 
                        
                        tw_summary[period_name] = sum_obj
                
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
    
    # === Load Previous Year Data for Comparison ===
    # === Load Previous Year Data for Comparison ===
    # Use ReferenceDataLoader because NIB files are Reference/Master files, not PB OSS files
    from app.data.reference_loader import ReferenceDataLoader
    from app.config import TRIWULAN_KE_BULAN
    
    ref_loader = ReferenceDataLoader()
    
    current_nib_file = st.session_state.get('nib_ref_file')
    prev_nib_file = st.session_state.get('nib_prev_ref_file')
    
    # 1. Load Full Data for Current and Previous Year
    current_full_data = None
    if current_nib_file:
         try:
             current_full_data = ref_loader.load_nib(current_nib_file.getvalue(), current_nib_file.name)
         except Exception as e:
             st.error(f"Error loading current NIB file: {e}")

    prev_full_data = None
    if prev_nib_file:
         try:
             prev_full_data = ref_loader.load_nib(prev_nib_file.getvalue(), prev_nib_file.name)
         except Exception as e:
             st.error(f"Error loading previous NIB file: {e}")

    # 2. Determine Target Months
    # Define Semester mappings (not in config, so define here)
    SEMESTER_KE_BULAN = {
        "Semester I": ["Januari", "Februari", "Maret", "April", "Mei", "Juni"],
        "Semester II": ["Juli", "Agustus", "September", "Oktober", "November", "Desember"],
    }
    
    target_months = []
    if report.period_type == "Triwulan" and report.period_name in TRIWULAN_KE_BULAN:
        target_months = TRIWULAN_KE_BULAN[report.period_name]
    elif report.period_type == "Semester" and report.period_name in SEMESTER_KE_BULAN:
        target_months = SEMESTER_KE_BULAN[report.period_name]
    elif report.period_type == "Tahunan":
        target_months = [m for sublist in TRIWULAN_KE_BULAN.values() for m in sublist]

    # 3. Calculate Totals
    # Current Period
    current_total = 0
    if current_full_data:
        current_total = sum(current_full_data.monthly_totals.get(m, 0) for m in target_months)
        
    # Previous Year (Y-o-Y)
    prev_year_total = 0
    if prev_full_data:
        prev_year_total = sum(prev_full_data.monthly_totals.get(m, 0) for m in target_months)

    # Previous Quarter (Q-o-Q)
    prev_q_total = 0
    prev_q_label = ""
    has_prev_q_data = False
    
    if report.period_type == "Triwulan":
        tw_list = ["TW I", "TW II", "TW III", "TW IV"]
        try:
            curr_idx = tw_list.index(report.period_name)
            if curr_idx > 0:
                # Same year, prev quarter
                prev_q_name = tw_list[curr_idx - 1]
                prev_q_months = TRIWULAN_KE_BULAN[prev_q_name]
                if current_full_data:
                    prev_q_total = sum(current_full_data.monthly_totals.get(m, 0) for m in prev_q_months)
                    has_prev_q_data = True
                prev_q_label = f"{prev_q_name} {report.year}"
            else:
                # Prev year, TW IV
                prev_q_name = "TW IV"
                prev_q_months = TRIWULAN_KE_BULAN[prev_q_name]
                if prev_full_data:
                     prev_q_total = sum(prev_full_data.monthly_totals.get(m, 0) for m in prev_q_months)
                     has_prev_q_data = True
                prev_q_label = f"{prev_q_name} {report.year - 1}"
        except ValueError:
            pass
    
    elif report.period_type == "Semester":
        # For Semester periods:
        # Y-o-Y: Compare last TW of semester (TW II or TW IV) with same TW previous year
        # Q-o-Q: Compare first TW vs last TW within semester (TW I vs TW II, or TW III vs TW IV)
        
        if report.period_name == "Semester I":
            tw1_months = TRIWULAN_KE_BULAN["TW I"]
            tw2_months = TRIWULAN_KE_BULAN["TW II"]
            
            # Q-o-Q: TW I vs TW II (same year)
            if current_full_data:
                tw1_total = sum(current_full_data.monthly_totals.get(m, 0) for m in tw1_months)
                tw2_total = sum(current_full_data.monthly_totals.get(m, 0) for m in tw2_months)
                prev_q_total = tw1_total  # TW I is "previous quarter"
                has_prev_q_data = True
                prev_q_label = f"TW I {report.year}"
                # Store current quarter total for Q-o-Q chart (TW II)
                current_q_total = tw2_total
            
            # Y-o-Y: TW II this year vs TW II previous year
            if prev_full_data:
                prev_year_tw2 = sum(prev_full_data.monthly_totals.get(m, 0) for m in tw2_months)
                prev_year_total = prev_year_tw2
                
        elif report.period_name == "Semester II":
            tw3_months = TRIWULAN_KE_BULAN["TW III"]
            tw4_months = TRIWULAN_KE_BULAN["TW IV"]
            
            # Q-o-Q: TW III vs TW IV (same year)
            if current_full_data:
                tw3_total = sum(current_full_data.monthly_totals.get(m, 0) for m in tw3_months)
                tw4_total = sum(current_full_data.monthly_totals.get(m, 0) for m in tw4_months)
                prev_q_total = tw3_total  # TW III is "previous quarter"
                has_prev_q_data = True
                prev_q_label = f"TW III {report.year}"
                # Store current quarter total for Q-o-Q chart (TW IV)
                current_q_total = tw4_total
            
            # Y-o-Y: TW IV this year vs TW IV previous year
            if prev_full_data:
                prev_year_tw4 = sum(prev_full_data.monthly_totals.get(m, 0) for m in tw4_months)
                prev_year_total = prev_year_tw4
    
    # === Top Row: Monthly Chart + Narrative ===
    col_top_left, col_top_right = st.columns([1, 1])
    
    with col_top_left:
        # Monthly bar chart with trendline
        monthly_data = stats.get('monthly_totals', {})
        if monthly_data:
            fig_monthly = chart_gen.create_monthly_bar_with_trendline(
                monthly_data,
                f"JUMLAH NIB PER-BULAN TAHUN {report.year}\nDI PROVINSI LAMPUNG"
            )
            st.plotly_chart(fig_monthly, use_container_width=True)
            
    with col_top_right:
        st.markdown(f'<div class="narrative-box">{narratives.rekapitulasi_nib}</div>', 
                    unsafe_allow_html=True)

    # === Bottom Row: Y-o-Y + Q-o-Q ===
    col_btm_left, col_btm_right = st.columns(2)
    
    # Determine labels based on period type
    if report.period_type == "Semester":
        if report.period_name == "Semester I":
            yoy_current_label = f"TW II {report.year}"
            yoy_prev_label = f"TW II {report.year - 1}"
            qoq_current_label = f"TW II {report.year}"
            qoq_prev_label = f"TW I {report.year}"
            yoy_title = f"JUMLAH NIB DI PROVINSI LAMPUNG\nPERIODE TRIWULAN II {report.year - 1} & TRIWULAN II {report.year} (y-o-y)"
            qoq_title = f"JUMLAH NIB DI PROVINSI LAMPUNG\nPERIODE TRIWULAN I {report.year} & TRIWULAN II {report.year} (q-o-q)"
            # Use TW II totals for Y-o-Y and Q-o-Q
            current_yoy_value = current_q_total if 'current_q_total' in dir() else current_total
            current_qoq_value = current_q_total if 'current_q_total' in dir() else current_total
        else:  # Semester II
            yoy_current_label = f"TW IV {report.year}"
            yoy_prev_label = f"TW IV {report.year - 1}"
            qoq_current_label = f"TW IV {report.year}"
            qoq_prev_label = f"TW III {report.year}"
            yoy_title = f"JUMLAH NIB DI PROVINSI LAMPUNG\nPERIODE TRIWULAN IV {report.year - 1} & TRIWULAN IV {report.year} (y-o-y)"
            qoq_title = f"JUMLAH NIB DI PROVINSI LAMPUNG\nPERIODE TRIWULAN III {report.year} & TRIWULAN IV {report.year} (q-o-q)"
            current_yoy_value = current_q_total if 'current_q_total' in dir() else current_total
            current_qoq_value = current_q_total if 'current_q_total' in dir() else current_total
    else:
        yoy_current_label = f"{report.period_name} {report.year}"
        yoy_prev_label = f"{report.period_name} {report.year - 1}"
        qoq_current_label = f"{report.period_name} {report.year}"
        qoq_prev_label = prev_q_label
        yoy_title = f"JUMLAH NIB DI PROVINSI LAMPUNG\nPERIODE {report.period_name} {report.year-1} & {report.period_name} {report.year} (y-o-y)"
        qoq_title = f"JUMLAH NIB DI PROVINSI LAMPUNG\nPERIODE {prev_q_label} & {report.period_name} {report.year} (q-o-q)"
        current_yoy_value = current_total
        current_qoq_value = current_total
    
    with col_btm_left:
        # Y-o-Y Chart
        if prev_full_data:
             fig_yoy = chart_gen.create_qoq_comparison_bar(
                current_data={yoy_current_label: current_yoy_value},
                previous_data={yoy_prev_label: prev_year_total},
                current_label=yoy_current_label,
                previous_label=yoy_prev_label,
                title=yoy_title
             )
             st.plotly_chart(fig_yoy, use_container_width=True)
        else:
             st.info("Upload file tahun sebelumnya untuk melihat perbandingan Y-o-Y")

    with col_btm_right:
        # Q-o-Q Chart
        if has_prev_q_data:
             fig_qoq = chart_gen.create_qoq_comparison_bar(
                current_data={qoq_current_label: current_qoq_value},
                previous_data={qoq_prev_label: prev_q_total},
                current_label=qoq_current_label,
                previous_label=qoq_prev_label,
                title=qoq_title
             )
             st.plotly_chart(fig_qoq, use_container_width=True)
        else:
             st.info("Data triwulan sebelumnya tidak tersedia untuk perbandingan Q-o-Q")
    
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
    
    # Section 4: Status PM - Redesigned Layout
    st.markdown('<div class="section-title">1.3 Rekapitulasi Data NIB Berdasarkan Status Penanaman Modal</div>', 
                unsafe_allow_html=True)
    
    # Get PM distribution from stats
    pm_dist = stats.get('pm_distribution', {})
    current_pma = pm_dist.get('PMA', 0)
    current_pmdn = pm_dist.get('PMDN', 0)
    
    # Calculate TW-level PM data for comparisons (for Semester periods)
    tw1_pma, tw1_pmdn, tw2_pma, tw2_pmdn = 0, 0, 0, 0
    prev_year_tw_pma, prev_year_tw_pmdn = 0, 0
    
    if report.period_type == "Semester" and report.period_name == "Semester I":
        tw1_months = TRIWULAN_KE_BULAN["TW I"]
        tw2_months = TRIWULAN_KE_BULAN["TW II"]
        
        if current_full_data:
            # Get TW I PM totals
            tw1_pm = current_full_data.get_period_by_pm_status(tw1_months)
            tw1_pma = tw1_pm.get('PMA', 0)
            tw1_pmdn = tw1_pm.get('PMDN', 0)
            
            # Get TW II PM totals
            tw2_pm = current_full_data.get_period_by_pm_status(tw2_months)
            tw2_pma = tw2_pm.get('PMA', 0)
            tw2_pmdn = tw2_pm.get('PMDN', 0)
        
        if prev_full_data:
            # Get TW II prev year PM totals
            prev_tw2_pm = prev_full_data.get_period_by_pm_status(tw2_months)
            prev_year_tw_pma = prev_tw2_pm.get('PMA', 0)
            prev_year_tw_pmdn = prev_tw2_pm.get('PMDN', 0)
            
    elif report.period_type == "Semester" and report.period_name == "Semester II":
        tw3_months = TRIWULAN_KE_BULAN["TW III"]
        tw4_months = TRIWULAN_KE_BULAN["TW IV"]
        
        if current_full_data:
            tw3_pm = current_full_data.get_period_by_pm_status(tw3_months)
            tw1_pma = tw3_pm.get('PMA', 0)  # Using tw1 var for "first TW of semester"
            tw1_pmdn = tw3_pm.get('PMDN', 0)
            
            tw4_pm = current_full_data.get_period_by_pm_status(tw4_months)
            tw2_pma = tw4_pm.get('PMA', 0)  # Using tw2 var for "second TW of semester"
            tw2_pmdn = tw4_pm.get('PMDN', 0)
        
        if prev_full_data:
            prev_tw4_pm = prev_full_data.get_period_by_pm_status(tw4_months)
            prev_year_tw_pma = prev_tw4_pm.get('PMA', 0)
            prev_year_tw_pmdn = prev_tw4_pm.get('PMDN', 0)
    
    elif report.period_type == "Triwulan":
        # For Triwulan, use current period as current and prev Q for comparison
        if current_full_data:
            curr_pm = current_full_data.get_period_by_pm_status(target_months)
            tw2_pma = curr_pm.get('PMA', 0)
            tw2_pmdn = curr_pm.get('PMDN', 0)
        
        if has_prev_q_data and current_full_data:
            prev_q_pm_months = TRIWULAN_KE_BULAN.get(prev_q_label.split()[0] + " " + prev_q_label.split()[1], []) if len(prev_q_label.split()) >= 2 else []
            if prev_q_label.startswith("TW"):
                tw_name = " ".join(prev_q_label.split()[:2])
                if tw_name in TRIWULAN_KE_BULAN:
                    prev_q_pm_data = current_full_data.get_period_by_pm_status(TRIWULAN_KE_BULAN[tw_name])
                    tw1_pma = prev_q_pm_data.get('PMA', 0)
                    tw1_pmdn = prev_q_pm_data.get('PMDN', 0)
        
        if prev_full_data:
            prev_year_pm = prev_full_data.get_period_by_pm_status(target_months)
            prev_year_tw_pma = prev_year_pm.get('PMA', 0)
            prev_year_tw_pmdn = prev_year_pm.get('PMDN', 0)
    
    # === Row 1: PM Bar Chart + Table ===
    col_pm1, col_pm2 = st.columns([1, 1.5])
    
    with col_pm1:
        # Horizontal bar chart for current period PM distribution
        fig_pm_bar = chart_gen.create_pm_horizontal_bar(
            pma_total=current_pma,
            pmdn_total=current_pmdn,
            title=f"Status PM - {report.period_name} {report.year}"
        )
        st.plotly_chart(fig_pm_bar, use_container_width=True)
    
    with col_pm2:
        # Detailed table with PM breakdown
        if not df.empty and 'Kabupaten/Kota' in df.columns:
            pm_table_cols = ['Kabupaten/Kota', 'PMA', 'PMDN', 'Total']
            if all(c in df.columns for c in ['PMA', 'PMDN']):
                pm_df = df[pm_table_cols].copy() if 'Total' in df.columns else df[['Kabupaten/Kota', 'PMA', 'PMDN']].copy()
                if 'Total' not in pm_df.columns:
                    pm_df['Total'] = pm_df['PMA'] + pm_df['PMDN']
                st.markdown(df_to_html_table(pm_df, max_rows=15), unsafe_allow_html=True)
    
    # === Row 2: Y-o-Y and Q-o-Q PM Comparison Charts ===
    col_pm_yoy, col_pm_qoq = st.columns(2)
    
    with col_pm_yoy:
        # Y-o-Y PM Comparison
        if prev_full_data and (prev_year_tw_pma > 0 or prev_year_tw_pmdn > 0):
            if report.period_type == "Semester":
                if report.period_name == "Semester I":
                    yoy_title = f"Status PM: TW II {report.year - 1} vs TW II {report.year} (Y-o-Y)"
                    curr_label = f"TW II {report.year}"
                    prev_label = f"TW II {report.year - 1}"
                    yoy_curr_pma, yoy_curr_pmdn = tw2_pma, tw2_pmdn
                else:
                    yoy_title = f"Status PM: TW IV {report.year - 1} vs TW IV {report.year} (Y-o-Y)"
                    curr_label = f"TW IV {report.year}"
                    prev_label = f"TW IV {report.year - 1}"
                    yoy_curr_pma, yoy_curr_pmdn = tw2_pma, tw2_pmdn
            else:
                yoy_title = f"Status PM: {report.period_name} {report.year - 1} vs {report.year} (Y-o-Y)"
                curr_label = f"{report.period_name} {report.year}"
                prev_label = f"{report.period_name} {report.year - 1}"
                yoy_curr_pma, yoy_curr_pmdn = tw2_pma, tw2_pmdn
            
            fig_pm_yoy = chart_gen.create_pm_grouped_comparison(
                current_pma=yoy_curr_pma,
                current_pmdn=yoy_curr_pmdn,
                prev_pma=prev_year_tw_pma,
                prev_pmdn=prev_year_tw_pmdn,
                current_label=curr_label,
                prev_label=prev_label,
                title=yoy_title
            )
            st.plotly_chart(fig_pm_yoy, use_container_width=True)
        else:
            st.info("Upload file tahun sebelumnya untuk Y-o-Y")
    
    with col_pm_qoq:
        # Q-o-Q PM Comparison
        if tw1_pma > 0 or tw1_pmdn > 0 or tw2_pma > 0 or tw2_pmdn > 0:
            if report.period_type == "Semester":
                if report.period_name == "Semester I":
                    qoq_title = f"Status PM: TW I vs TW II {report.year} (Q-o-Q)"
                    curr_label = f"TW II {report.year}"
                    prev_label = f"TW I {report.year}"
                else:
                    qoq_title = f"Status PM: TW III vs TW IV {report.year} (Q-o-Q)"
                    curr_label = f"TW IV {report.year}"
                    prev_label = f"TW III {report.year}"
            else:
                qoq_title = f"Status PM: {prev_q_label} vs {report.period_name} {report.year} (Q-o-Q)"
                curr_label = f"{report.period_name} {report.year}"
                prev_label = prev_q_label
            
            fig_pm_qoq = chart_gen.create_pm_grouped_comparison(
                current_pma=tw2_pma,
                current_pmdn=tw2_pmdn,
                prev_pma=tw1_pma,
                prev_pmdn=tw1_pmdn,
                current_label=curr_label,
                prev_label=prev_label,
                title=qoq_title
            )
            st.plotly_chart(fig_pm_qoq, use_container_width=True)
        else:
            st.info("Data triwulan sebelumnya tidak tersedia untuk perbandingan Q-o-Q per Status PM")
    
    st.markdown(f'<div class="narrative-box">{narratives.status_pm}</div>', 
                unsafe_allow_html=True)
    
    # Section 1.4: Pelaku Usaha - Redesigned Layout
    st.markdown('<div class="section-title">1.4 Rekapitulasi Data NIB Berdasarkan Pelaku Usaha</div>', 
                unsafe_allow_html=True)
    
    # Helper to aggregate Skala Usaha (Mikro+Kecil -> UMK, Menengah+Besar -> NON-UMK)
    def aggregate_pelaku_usaha(full_data_obj, months_list):
        if not full_data_obj:
            return 0, 0
            
        skala_data = full_data_obj.get_period_by_skala_usaha(months_list)
        umk_val = 0
        non_umk_val = 0
        
        # Iterate through all keys to perform robust matching
        for key, val in skala_data.items():
            k_upper = str(key).upper()
            if 'MIKRO' in k_upper or 'KECIL' in k_upper or 'UMK' in k_upper:
                # Avoid double counting if key is "NON UMK"
                if 'NON' in k_upper:
                    non_umk_val += val
                else:
                    umk_val += val
            elif 'MENENGAH' in k_upper or 'BESAR' in k_upper:
                non_umk_val += val
            elif 'NON' in k_upper: # Fallback for "NON UMK" or similar
                non_umk_val += val
            
        return umk_val, non_umk_val

    # Calculate current period totals using robust helper
    current_umk, current_non_umk = 0, 0
    if current_full_data:
        current_umk, current_non_umk = aggregate_pelaku_usaha(current_full_data, target_months)

    # Calculate TW-level Pelaku Usaha data for comparisons
    tw1_umk, tw1_non_umk, tw2_umk, tw2_non_umk = 0, 0, 0, 0
    prev_year_tw_umk, prev_year_tw_non_umk = 0, 0
    
    if report.period_type == "Semester" and report.period_name == "Semester I":
        tw1_months = TRIWULAN_KE_BULAN["TW I"]
        tw2_months = TRIWULAN_KE_BULAN["TW II"]
        
        if current_full_data:
            tw1_umk, tw1_non_umk = aggregate_pelaku_usaha(current_full_data, tw1_months)
            tw2_umk, tw2_non_umk = aggregate_pelaku_usaha(current_full_data, tw2_months)
        
        if prev_full_data:
            prev_year_tw_umk, prev_year_tw_non_umk = aggregate_pelaku_usaha(prev_full_data, tw2_months)
            
    elif report.period_type == "Semester" and report.period_name == "Semester II":
        tw3_months = TRIWULAN_KE_BULAN["TW III"]
        tw4_months = TRIWULAN_KE_BULAN["TW IV"]
        
        if current_full_data:
            tw1_umk, tw1_non_umk = aggregate_pelaku_usaha(current_full_data, tw3_months) # Using tw1 var for first TW 
            tw2_umk, tw2_non_umk = aggregate_pelaku_usaha(current_full_data, tw4_months) # Using tw2 var for second TW
        
        if prev_full_data:
            prev_year_tw_umk, prev_year_tw_non_umk = aggregate_pelaku_usaha(prev_full_data, tw4_months)
    
    elif report.period_type == "Triwulan":
        if current_full_data:
            tw2_umk, tw2_non_umk = aggregate_pelaku_usaha(current_full_data, target_months)
        
        # Q-o-Q Logic: Handle cross-year (TW I) vs same-year (TW II-IV)
        if has_prev_q_data:
             # prev_q_label is like "TW IV 2024" or "TW I 2025"
             parts = prev_q_label.split()
             if len(parts) >= 3:
                 prev_q_name = f"{parts[0]} {parts[1]}" # "TW IV"
                 prev_q_year_str = parts[2]
                 
                 # Decide source: current year file or prev year file?
                 source_data = current_full_data if str(report.year) == prev_q_year_str else prev_full_data
                 
                 if source_data and prev_q_name in TRIWULAN_KE_BULAN:
                     tw1_umk, tw1_non_umk = aggregate_pelaku_usaha(source_data, TRIWULAN_KE_BULAN[prev_q_name])
        
        if prev_full_data:
            prev_year_tw_umk, prev_year_tw_non_umk = aggregate_pelaku_usaha(prev_full_data, target_months)

    # === Row 1: Pelaku Usaha Bar Chart + Table ===
    col_pelaku1, col_pelaku2 = st.columns([1, 1.5])
    
    with col_pelaku1:
        # Horizontal bar chart
        fig_pelaku_bar = chart_gen.create_pelaku_usaha_horizontal_bar(
            umk_total=current_umk,
            non_umk_total=current_non_umk,
            title=f"Kategori Pelaku Usaha - {report.period_name} {report.year}"
        )
        st.plotly_chart(fig_pelaku_bar, use_container_width=True)
    
    with col_pelaku2:
        # Detailed table with Per-District breakdown
        if not df.empty and 'Kabupaten/Kota' in df.columns:
            # Check if we have UMK/NON-UMK columns
            pelaku_cols = ['Kabupaten/Kota', 'UMK', 'NON-UMK', 'Total']
            # If columns might be named 'NON_UMK' (underscore), handle that
            available_cols = df.columns.tolist()
            non_umk_col = 'NON_UMK' if 'NON_UMK' in available_cols else 'NON-UMK'
            
            if 'UMK' in available_cols and non_umk_col in available_cols:
                pelaku_df = df[['Kabupaten/Kota', 'UMK', non_umk_col]].copy()
                if non_umk_col != 'NON-UMK':
                    pelaku_df = pelaku_df.rename(columns={non_umk_col: 'NON-UMK'})
                pelaku_df['Total'] = pelaku_df['UMK'] + pelaku_df['NON-UMK']
                st.markdown(df_to_html_table(pelaku_df, max_rows=15), unsafe_allow_html=True)
    
    # === Row 2: Y-o-Y and Q-o-Q Pelaku Usaha Comparisons ===
    col_pelaku_yoy, col_pelaku_qoq = st.columns(2)
    
    with col_pelaku_yoy:
        # Y-o-Y Comparison
        if prev_full_data and (prev_year_tw_umk > 0 or prev_year_tw_non_umk > 0):
            if report.period_type == "Semester":
                if report.period_name == "Semester I":
                    yoy_title = f"Kategori Pelaku Usaha: TW II {report.year - 1} vs TW II {report.year} (Y-o-Y)"
                    curr_label = f"TW II {report.year}"
                    prev_label = f"TW II {report.year - 1}"
                    yoy_curr_umk, yoy_curr_non_umk = tw2_umk, tw2_non_umk
                else:
                    yoy_title = f"Kategori Pelaku Usaha: TW IV {report.year - 1} vs TW IV {report.year} (Y-o-Y)"
                    curr_label = f"TW IV {report.year}"
                    prev_label = f"TW IV {report.year - 1}"
                    yoy_curr_umk, yoy_curr_non_umk = tw2_umk, tw2_non_umk
            else:
                yoy_title = f"Kategori Pelaku Usaha: {report.period_name} {report.year - 1} vs {report.year} (Y-o-Y)"
                curr_label = f"{report.period_name} {report.year}"
                prev_label = f"{report.period_name} {report.year - 1}"
                yoy_curr_umk, yoy_curr_non_umk = tw2_umk, tw2_non_umk
            
            fig_pelaku_yoy = chart_gen.create_pelaku_grouped_comparison(
                current_umk=yoy_curr_umk,
                current_non_umk=yoy_curr_non_umk,
                prev_umk=prev_year_tw_umk,
                prev_non_umk=prev_year_tw_non_umk,
                current_label=curr_label,
                prev_label=prev_label,
                title=yoy_title
            )
            st.plotly_chart(fig_pelaku_yoy, use_container_width=True)
            
        else:
            st.info("Upload file triwulan tahun sebelumnya untuk Y-o-Y")
    
    with col_pelaku_qoq:
        # Q-o-Q Comparison
        if tw1_umk > 0 or tw1_non_umk > 0 or tw2_umk > 0 or tw2_non_umk > 0:
            if report.period_type == "Semester":
                if report.period_name == "Semester I":
                    qoq_title = f"Kategori Pelaku Usaha: TW I vs TW II {report.year} (Q-o-Q)"
                    curr_label = f"TW II {report.year}"
                    prev_label = f"TW I {report.year}"
                else:
                    qoq_title = f"Kategori Pelaku Usaha: TW III vs TW IV {report.year} (Q-o-Q)"
                    curr_label = f"TW IV {report.year}"
                    prev_label = f"TW III {report.year}"
            else:
                qoq_title = f"Kategori Pelaku Usaha: {prev_q_label} vs {report.period_name} {report.year} (Q-o-Q)"
                curr_label = f"{report.period_name} {report.year}"
                prev_label = prev_q_label
            
            fig_pelaku_qoq = chart_gen.create_pelaku_grouped_comparison(
                current_umk=tw2_umk,
                current_non_umk=tw2_non_umk,
                prev_umk=tw1_umk,
                prev_non_umk=tw1_non_umk,
                current_label=curr_label,
                prev_label=prev_label,
                title=qoq_title
            )
            st.plotly_chart(fig_pelaku_qoq, use_container_width=True)
        else:
            st.info("Data triwulan sebelumnya tidak tersedia untuk Q-o-Q")
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
            
            # Load Previous Year Project File if available
            proyek_file = st.session_state.get('proyek_ref_file')
            proyek_prev_file = st.session_state.get('proyek_prev_ref_file')
            
            if proyek_file:
                from app.data.reference_loader import ReferenceDataLoader
                loader = ReferenceDataLoader()
                
                # Load Current Data
                current_proyek_data = _cached_load_proyek(proyek_file.getvalue(), proyek_file.name, report.year)
                
                # Load Previous Data (if available)
                prev_proyek_data = None
                if proyek_prev_file:
                    prev_proyek_data = _cached_load_proyek(proyek_prev_file.getvalue(), proyek_prev_file.name, report.year - 1)
                
                # Calculate Stats
                target_months = loader.get_months_for_period(report.period_type, report.period_name)
                
                def get_proyek_total(data_obj, months):
                   if not data_obj: return 0
                   return sum(data_obj.monthly_projects.get(m, 0) for m in months)

                # Current Period Total
                curr_total_proyek = get_proyek_total(current_proyek_data, target_months)
                
                # Y-o-Y Stats
                prev_year_total_proyek = get_proyek_total(prev_proyek_data, target_months)
                
                # Q-o-Q Stats
                prev_q_total_proyek = 0
                prev_q_source_data = None
                prev_q_label_text = prev_q_label # e.g. "TW IV 2024" or "TW I 2025" from earlier global var
                
                if has_prev_q_data: # Use global flag as hint that we should look for Q-o-Q
                     # Determine source for Prev Q (Same Year or Prev Year)
                     parts = prev_q_label.split()
                     if len(parts) >= 3:
                         prev_q_name = f"{parts[0]} {parts[1]}"
                         prev_q_year_str = parts[2]
                         prev_q_source_data = current_proyek_data if str(report.year) == prev_q_year_str else prev_proyek_data
                         
                         if prev_q_source_data and prev_q_name in TRIWULAN_KE_BULAN:
                             prev_q_total_proyek = get_proyek_total(prev_q_source_data, TRIWULAN_KE_BULAN[prev_q_name])
                
                if current_proyek_data:
                    # Layout: Left Col (3 Charts), Right Col (1 Chart)
                    col_left, col_right = st.columns([1, 1.2]) # Right column slightly wider for long names
                    
                    with col_left:
                        # 1. Monthly Chart
                        monthly_project_data = {}
                        for month in target_months:
                            if month in current_proyek_data.monthly_projects:
                                monthly_project_data[month] = current_proyek_data.monthly_projects[month]
                        
                        if monthly_project_data:
                            fig_monthly_proj = chart_gen.create_monthly_bar_with_trendline(
                                monthly_project_data,
                                title=f"Jumlah Proyek Per-Bulan Tahun {report.year}",
                                show_trendline=False # Reference image implies simple bars
                            )
                            # Customize to match blue pillars in reference
                            fig_monthly_proj.update_traces(marker_color='#3498db', opacity=0.9)
                            fig_monthly_proj.update_layout(height=350, margin=dict(l=20, r=20, t=40, b=20))
                            st.plotly_chart(fig_monthly_proj, use_container_width=True)
                        
                        # 2. Y-o-Y Chart
                        if prev_proyek_data:
                             # Logic for label
                             if report.period_type == "Semester":
                                 pass # Logic handled by calc, just label
                             
                             yoy_title = f"Jumlah Proyek: {report.period_name} {report.year - 1} & {report.period_name} {report.year} (y-o-y)"
                             
                             fig_yoy = chart_gen.create_comparison_bar_chart(
                                 current_val=curr_total_proyek,
                                 prev_val=prev_year_total_proyek,
                                 current_label=f"{report.period_name} {report.year}",
                                 prev_label=f"{report.period_name} {report.year - 1}",
                                 title=yoy_title
                             )
                             st.plotly_chart(fig_yoy, use_container_width=True)
                        else:
                             st.info("Upload file proyek tahun sebelumnya untuk Y-o-Y")

                        # 3. Q-o-Q Chart
                        if prev_q_total_proyek > 0 or curr_total_proyek > 0 and has_prev_q_data:
                             qoq_title = f"Jumlah Proyek: {prev_q_label_text} & {report.period_name} {report.year} (q-o-q)"
                             
                             fig_qoq = chart_gen.create_comparison_bar_chart(
                                 current_val=curr_total_proyek,
                                 prev_val=prev_q_total_proyek,
                                 current_label=f"{report.period_name} {report.year}",
                                 prev_label=prev_q_label_text,
                                 title=qoq_title
                             )
                             st.plotly_chart(fig_qoq, use_container_width=True)
                        else:
                             st.info(f"Data {prev_q_label_text} tidak tersedia untuk Q-o-Q")

                    with col_right:
                        # District (Kab/Kota) Chart - Tall
                        import plotly.graph_objects as go
                        projects_by_kab = current_proyek_data.get_period_projects_by_wilayah(target_months)
                        
                        if projects_by_kab:
                            # Show ALL districts (or top 15 if too many) - reference showing many
                            sorted_kab = dict(sorted(projects_by_kab.items(), key=lambda x: x[1], reverse=True)) # All sorted
                            
                            fig_kab = go.Figure(data=[go.Bar(
                                x=list(sorted_kab.values()), 
                                y=list(sorted_kab.keys()), 
                                orientation='h', 
                                marker_color='#4a90e2',
                                text=[f"{val:,}".replace(",", ".") for val in sorted_kab.values()],
                                textposition='outside'
                            )])
                            
                            fig_kab.update_layout(
                                title='Jumlah Proyek Berdasarkan Kabupaten/Kota',
                                template='plotly_dark',
                                height=750, # Taller chart to match reference
                                yaxis={'categoryorder': 'total ascending'}, # High at top
                                margin=dict(l=0, r=0, t=40, b=0),
                                xaxis_title="Jumlah Proyek"
                            )
                            st.plotly_chart(fig_kab, use_container_width=True)
                    
                    # Logic for narrative
                    total_proyek = curr_total_proyek
                    
                    # Calculate growth stats for narrative
                    if prev_year_total_proyek > 0:
                        yoy_growth = ((curr_total_proyek - prev_year_total_proyek) / prev_year_total_proyek) * 100
                        yoy_text = f"{'meningkat' if yoy_growth >= 0 else 'menurun'} sebesar <b>{abs(yoy_growth):.2f}%</b>"
                    else:
                        yoy_text = "tidak dapat dibandingkan (data tahun lalu tidak tersedia)"
                        
                    if prev_q_total_proyek > 0:
                        qoq_growth = ((curr_total_proyek - prev_q_total_proyek) / prev_q_total_proyek) * 100
                        qoq_text = f"{'meningkat' if qoq_growth >= 0 else 'menurun'} sebesar <b>{abs(qoq_growth):.2f}%</b>"
                    else:
                        qoq_text = "tidak dapat dibandingkan (data triwulan lalu tidak tersedia)"

                    top_kab = list(sorted_kab.items())[0] if projects_by_kab else ("-", 0)
                    
                    interpretation = f"""
                    <b>Analisis dan Interpretasi:</b><br>
                    Rekapitulasi jumlah proyek di Provinsi Lampung periode {report.period_name} Tahun {report.year} 
                    adalah sebanyak <b>{total_proyek:,}</b> proyek. <br>
                    Proyek tertinggi berada di lokasi <b>{top_kab[0]}</b> sebanyak <b>{top_kab[1]:,}</b> proyek.
                    Jika dibandingkan dengan tahun sebelumnya ({report.period_name} {report.year-1}), mengalami {yoy_text}.
                    Dan jika dibandingkan dengan triwulan sebelumnya ({prev_q_label_text}), mengalami {qoq_text}.
                    """
                    st.markdown(f'<div class="narrative-box">{interpretation}</div>', unsafe_allow_html=True)
            
            st.markdown('<div class="section-title">2.2 Rekapitulasi Proyek Berdasarkan Status Penanaman Modal</div>', 
                        unsafe_allow_html=True)
            
            # --- CALCULATE PMA/PMDN STATS (PROJECT COUNTS) ---
            current_pma = current_investment.pma_proyek
            current_pmdn = current_investment.pmdn_proyek
            
            # Y-o-Y Stats
            prev_year_pma = 0
            prev_year_pmdn = 0
            if 'prev_proyek_data' in locals() and prev_proyek_data:
                # Use helper in reference_loader via object method
                 # Note: prev_proyek_data is ProyekReferenceData object
                 prev_year_pma = prev_proyek_data.get_period_pma_projects(target_months)
                 prev_year_pmdn = prev_proyek_data.get_period_pmdn_projects(target_months)
            
            # Q-o-Q Stats
            prev_q_pma = 0
            prev_q_pmdn = 0
            if 'prev_q_source_data' in locals() and prev_q_source_data and 'prev_q_name' in locals():
                 prev_q_months = TRIWULAN_KE_BULAN[prev_q_name]
                 prev_q_pma = prev_q_source_data.get_period_pma_projects(prev_q_months)
                 prev_q_pmdn = prev_q_source_data.get_period_pmdn_projects(prev_q_months)

            # --- RENDER 2.2 CHARTS ---
            col1, col2, col3 = st.columns(3)
            
            with col1:
                # Current Status Bar Chart (Replaces Donut)
                fig_status = chart_gen.create_simple_bar_chart(
                    labels=['PMA', 'PMDN'],
                    values=[current_pma, current_pmdn],
                    title=f"Jumlah Proyek Berdasarkan Status PM {report.period_name} {report.year}",
                    color='#9b59b6' # Purple
                )
                st.plotly_chart(fig_status, use_container_width=True)
            
            with col2:
                # Y-o-Y Comparison
                if 'prev_proyek_data' in locals() and prev_proyek_data:
                     # Logic for label
                     if report.period_type == "Semester":
                         pass
                     
                     yoy_title = f"PMA & PMDN (y-o-y)"
                     
                     fig_yoy = chart_gen.create_grouped_comparison_two_categories(
                         curr_val1=current_pma,
                         curr_val2=current_pmdn,
                         prev_val1=prev_year_pma,
                         prev_val2=prev_year_pmdn,
                         cat1_label="PMA",
                         cat2_label="PMDN",
                         current_period_label=f"{report.period_name} {report.year}",
                         prev_period_label=f"{report.period_name} {report.year - 1}",
                         title=yoy_title,
                         y_axis_title="Jumlah Proyek"
                     )
                     st.plotly_chart(fig_yoy, use_container_width=True)
                else:
                     st.info("Upload file proyek tahun sebelumnya untuk Y-o-Y")

            with col3:
                # Q-o-Q Comparison
                if (prev_q_pma > 0 or current_pma > 0) and has_prev_q_data:
                     qoq_title = f"PMA & PMDN (q-o-q)"
                     
                     fig_qoq = chart_gen.create_grouped_comparison_two_categories(
                         curr_val1=current_pma,
                         curr_val2=current_pmdn,
                         prev_val1=prev_q_pma,
                         prev_val2=prev_q_pmdn,
                         cat1_label="PMA",
                         cat2_label="PMDN",
                         current_period_label=f"{report.period_name} {report.year}",
                         prev_period_label=prev_q_label_text,
                         title=qoq_title,
                         y_axis_title="Jumlah Proyek"
                     )
                     st.plotly_chart(fig_qoq, use_container_width=True)
                else:
                     st.info(f"Data {prev_q_label_text} tidak tersedia untuk Q-o-Q")

            # Narrative for 2.2
            pma_pmdn_narr = narrative_gen.generate_status_pm_narrative(
                current_pma,
                current_pmdn,
                unit_type="proyek"
            )
            if pma_pmdn_narr:
                st.markdown(f'<div class="narrative-box">{pma_pmdn_narr}</div>', unsafe_allow_html=True)




            
            # TW Comparison chart (if multiple TW data available)
            if len(investment_reports) > 1:
                st.markdown('<div class="section-title">Perbandingan Antar Triwulan (Investasi)</div>', 
                            unsafe_allow_html=True)
                fig_tw_comp = chart_gen.create_investment_tw_comparison_chart(investment_reports)
                st.plotly_chart(fig_tw_comp, use_container_width=True)
                
                # Narrative for TW comparison
                tw_comp_narr = narrative_gen.generate_tw_comparison_narrative(investment_reports)
                if tw_comp_narr:
                    st.markdown(f'<div class="narrative-box">{tw_comp_narr}</div>', unsafe_allow_html=True)
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
            # 2.3 Skala Usaha visualization (Redesigned with Y-o-Y & Q-o-Q)
            st.markdown('<div class="section-title">2.3 Rekapitulasi Data Proyek Berdasarkan Skala Usaha</div>', 
                        unsafe_allow_html=True)
            
            # Get proyek data
            proyek_data = None
            proyek_file = st.session_state.get('proyek_ref_file')
            proyek_prev_file = st.session_state.get('proyek_prev_ref_file')
            
            if proyek_file:
                from app.data.reference_loader import ReferenceDataLoader
                loader = ReferenceDataLoader()
                months = loader.get_months_for_period(report.period_type, report.period_name)
                
                # Load Current Data
                proyek_data = _cached_load_proyek(proyek_file.getvalue(), proyek_file.name, report.year)
                
                # Load Previous Year Data (for Y-o-Y)
                prev_proyek_data = None
                if proyek_prev_file:
                    prev_proyek_data = _cached_load_proyek(proyek_prev_file.getvalue(), proyek_prev_file.name, report.year - 1)
                
                # Determine Previous Quarter Data (for Q-o-Q)
                prev_q_source_data = None
                prev_q_name_str = None
                
                if has_prev_q_data:
                     try:
                         parts = prev_q_label.split() # e.g. "TW I 2025" or "TW IV 2024"
                         if len(parts) >= 3:
                             prev_q_name_str = f"{parts[0]} {parts[1]}"
                             prev_q_year_str = parts[2]
                             # Logic: If prev q year == current year, use current data. Else use prev data.
                             prev_q_source_data = proyek_data if str(report.year) == prev_q_year_str else prev_proyek_data
                     except Exception:
                         pass

                if proyek_data:
                    # Current Skala Usaha Data
                    skala_data = proyek_data.get_period_by_skala_usaha(months)
                    
                    if skala_data:
                        # Define standard keys and sort order
                        std_keys = ['Usaha Mikro', 'Usaha Kecil', 'Usaha Menengah', 'Usaha Besar']
                        
                        # --- 1. Current Period Chart ---
                        ordered_vals = [skala_data.get(k, 0) for k in std_keys]
                        
                        # Use generic simple bar chart logic or custom
                        fig_skala = go.Figure(data=[
                            go.Bar(
                                x=std_keys,
                                y=ordered_vals,
                                marker_color=['#3498db', '#e67e22', '#2ecc71', '#9b59b6'],
                                text=[f'{v:,.0f}'.replace(",", ".") for v in ordered_vals],
                                textposition='outside'
                            )
                        ])
                        fig_skala.update_layout(
                            title=f"Jumlah Proyek {report.period_name} {report.year} Berdasarkan Skala Usaha",
                            yaxis_title='Jumlah Proyek',
                            template='plotly_dark',
                            height=400,
                            **chart_gen.layout_defaults
                        )
                        st.plotly_chart(fig_skala, use_container_width=True)
                        
                        # --- 2. Comparison Charts (Bottom Row) ---
                        col_yoy, col_qoq = st.columns(2)
                        
                        with col_yoy:
                            if prev_proyek_data:
                                prev_skala_data = prev_proyek_data.get_period_by_skala_usaha(months)
                                prev_vals = [prev_skala_data.get(k, 0) for k in std_keys]
                                
                                fig_yoy_skala = chart_gen.create_grouped_comparison_multi_category(
                                    categories=[k.replace("Usaha ", "").upper() for k in std_keys], # Shorten labels
                                    current_values=ordered_vals,
                                    prev_values=prev_vals,
                                    current_label=f"{report.year}",
                                    prev_label=f"{report.year - 1}",
                                    title="Jumlah Proyek (y-o-y)",
                                    y_axis_title="Jumlah"
                                )
                                st.plotly_chart(fig_yoy_skala, use_container_width=True)
                            else:
                                st.info("Upload file proyek tahun sebelumnya untuk Y-o-Y")
                        
                        with col_qoq:
                            if prev_q_source_data and prev_q_name_str and prev_q_name_str in TRIWULAN_KE_BULAN:
                                pq_months = TRIWULAN_KE_BULAN[prev_q_name_str]
                                pq_skala_data = prev_q_source_data.get_period_by_skala_usaha(pq_months)
                                pq_vals = [pq_skala_data.get(k, 0) for k in std_keys]
                                
                                fig_qoq_skala = chart_gen.create_grouped_comparison_multi_category(
                                    categories=[k.replace("Usaha ", "").upper() for k in std_keys],
                                    current_values=ordered_vals,
                                    prev_values=pq_vals,
                                    current_label=f"{report.period_name}",
                                    prev_label=prev_q_name_str,
                                    title="Jumlah Proyek (q-o-q)",
                                    y_axis_title="Jumlah"
                                )
                                st.plotly_chart(fig_qoq_skala, use_container_width=True)
                            else:
                                st.info(f"Data {prev_q_label_text if 'prev_q_label_text' in locals() else 'Q-o-Q'} tidak tersedia")

                        
                        # Interpretation for Skala Usaha (Comparison Narrative)
                        interpretation_skala = narrative_gen.generate_skala_usaha_comparison_narrative(
                            current_data=skala_data,
                            prev_year_data=prev_proyek_data.get_period_by_skala_usaha(months) if prev_proyek_data else {},
                            prev_q_data={}, # Optional
                            period_name=report.period_name,
                            year=report.year
                        )
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
            
            # Q-o-Q and Y-o-Y Comparisons removed from Section 2.5 as per request
            
            # Project Narrative Interpretation
            st.markdown('<div class="section-title">Interpretasi Data Proyek</div>', 
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
        
        pb_data = _cached_load_pb_oss(pb_oss_file.getvalue(), pb_oss_file.name, report.year)
        
        if pb_data:
            from app.data.reference_loader import ReferenceDataLoader
            loader = ReferenceDataLoader()
            months = loader.get_months_for_period(report.period_type, report.period_name)
            
            # Summary metrics for Section 3
            total_permits = pb_data.get_period_permits(months)
            gubernur_permits = sum(pb_data.get_period_by_kab_kota(months).values()) if pb_data.get_period_by_kab_kota(months) else 0
            status_pm = pb_data.get_period_status_pm(months)
            pma_permits = status_pm.get('PMA', 0)
            pmdn_permits = status_pm.get('PMDN', 0)
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.markdown(f'''
                <div class="metric-card">
                    <div class="metric-value">{total_permits:,}</div>
                    <div class="metric-label">Total Perizinan</div>
                </div>
                ''', unsafe_allow_html=True)
            with col2:
                st.markdown(f'''
                <div class="metric-card">
                    <div class="metric-value">{gubernur_permits:,}</div>
                    <div class="metric-label">Kewenangan Gubernur</div>
                </div>
                ''', unsafe_allow_html=True)
            with col3:
                st.markdown(f'''
                <div class="metric-card">
                    <div class="metric-value">{pma_permits:,}</div>
                    <div class="metric-label">PMA (Asing)</div>
                </div>
                ''', unsafe_allow_html=True)
            with col4:
                st.markdown(f'''
                <div class="metric-card">
                    <div class="metric-value">{pmdn_permits:,}</div>
                    <div class="metric-label">PMDN (Domestik)</div>
                </div>
                ''', unsafe_allow_html=True)
            
            # 3.1 Period and Location
            st.markdown('<div class="section-title">3.1 Rekapitulasi Berdasarkan Periode dan Lokasi Usaha di Kabupaten/Kota</div>', 
                        unsafe_allow_html=True)
            
            # --- Load Previous Data for Comparisons ---
            pb_prev_file = st.session_state.get('pb_oss_prev_ref_file')
            prev_pb_data = None
            if pb_prev_file:
                prev_pb_data = _cached_load_pb_oss(pb_prev_file.getvalue(), pb_prev_file.name, report.year - 1)
            
            prev_q_pb_data = None
            prev_q_name_pb = None
            if has_prev_q_data:
                 try:
                     parts = prev_q_label.split()
                     if len(parts) >= 3:
                         prev_q_name_pb = f"{parts[0]} {parts[1]}"
                         prev_q_year_str = parts[2]
                         prev_q_pb_data = pb_data if str(report.year) == prev_q_year_str else prev_pb_data
                 except Exception:
                     pass

            # Calculate Totals
            curr_permits = pb_data.get_period_permits(months)
            
            prev_year_permits = 0
            if prev_pb_data:
                prev_year_permits = prev_pb_data.get_period_permits(months)
            
            prev_q_permits = 0
            if prev_q_pb_data and prev_q_name_pb in TRIWULAN_KE_BULAN:
                prev_q_months_pb = TRIWULAN_KE_BULAN[prev_q_name_pb]
                prev_q_permits = prev_q_pb_data.get_period_permits(prev_q_months_pb)

            # --- Render 3.1 Charts ---
            # --- Render 3.1 Charts ---
            
            # Row 1: Monthly & Location
            col_row1_1, col_row1_2 = st.columns(2)
            
            with col_row1_1:
                # Monthly Chart
                monthly_permits = pb_data.get_period_permits_by_month(months) if hasattr(pb_data, 'get_period_permits_by_month') else {}
                if not monthly_permits: monthly_permits = {}
                
                if monthly_permits:
                    fig_monthly = chart_gen.create_simple_bar_chart(
                        labels=list(monthly_permits.keys()),
                        values=list(monthly_permits.values()),
                        title="Jumlah Perizinan per Bulan",
                        color='#3498db'
                    )
                    st.plotly_chart(fig_monthly, use_container_width=True)
                else:
                    st.info("Data bulanan tidak tersedia")

            with col_row1_2:
                # Kab/Kota chart
                kab_data = pb_data.get_period_by_kab_kota(months)
                if kab_data:
                    import plotly.graph_objects as go
                    # Top 10 Kab/Kota
                    sorted_kab = dict(sorted(kab_data.items(), key=lambda x: x[1], reverse=True)[:10])
                    fig_kab = go.Figure(data=[go.Bar(
                        x=list(sorted_kab.values()), 
                        y=list(sorted_kab.keys()), 
                        orientation='h', 
                        marker_color='#3B82F6'
                    )])
                    fig_kab.update_layout(
                        title='Top 10 Lokasi Usaha (Kab/Kota)', 
                        template='plotly_dark', 
                        height=400, 
                        yaxis={'categoryorder': 'total ascending'},
                        margin=dict(l=10, r=10, t=40, b=10)
                    )
                    st.plotly_chart(fig_kab, use_container_width=True)
                else:
                    st.info("Data Kab/Kota tidak tersedia")

            # Row 2: Comparisons (Y-o-Y & Q-o-Q)
            col_row2_1, col_row2_2 = st.columns(2)
            with col_row2_1:
                # Y-o-Y
                if prev_pb_data:
                    fig_yoy_pb = chart_gen.create_comparison_bar_chart(
                        current_val=curr_permits,
                        prev_val=prev_year_permits,
                        current_label=f"{report.year}",
                        prev_label=f"{report.year - 1}",
                        title="Total Perizinan (y-o-y)"
                    )
                    st.plotly_chart(fig_yoy_pb, use_container_width=True)
                else:
                    st.info("Upload PB OSS tahun lalu untuk Y-o-Y")
            
            with col_row2_2:
                # Q-o-Q
                if prev_q_pb_data:
                    fig_qoq_pb = chart_gen.create_comparison_bar_chart(
                        current_val=curr_permits,
                        prev_val=prev_q_permits,
                        current_label=f"{report.period_name}",
                        prev_label=f"{prev_q_name_pb}",
                        title="Total Perizinan (q-o-q)"
                    )
                    st.plotly_chart(fig_qoq_pb, use_container_width=True)
                else:
                    st.info("Data Q-o-Q tidak tersedia")
            
            # 3.2 Status PM
            st.markdown('<div class="section-title">3.2 Rekapitulasi Berdasarkan Status Penanaman Modal</div>', 
                        unsafe_allow_html=True)
            pm_data = pb_data.get_period_status_pm(months)
            if pm_data:
                import plotly.graph_objects as go
                fig = go.Figure(data=[go.Bar(x=list(pm_data.keys()), y=list(pm_data.values()), marker_color=['#10B981', '#F59E0B'][:len(pm_data)])])
                fig.update_layout(title='Perizinan per Status PM', template='plotly_dark', height=350)
                st.plotly_chart(fig, use_container_width=True)
            
            # 3.3 Risk Level
            st.markdown('<div class="section-title">3.3 Rekapitulasi Berdasarkan Tingkat Risiko</div>', 
                        unsafe_allow_html=True)
            risk_data = pb_data.get_period_risk(months)
            if risk_data:
                import plotly.graph_objects as go
                fig = go.Figure(data=[go.Bar(x=list(risk_data.keys()), y=list(risk_data.values()), marker_color=['#22C55E', '#EAB308', '#EF4444', '#8B5CF6'][:len(risk_data)])])
                fig.update_layout(title='Perizinan per Tingkat Risiko', template='plotly_dark', height=350)
                st.plotly_chart(fig, use_container_width=True)
            
            # 3.4 Sector
            st.markdown('<div class="section-title">3.4 Rekapitulasi Berdasarkan Sektor Kementerian/Lembaga</div>', 
                        unsafe_allow_html=True)
            sector_data = pb_data.get_period_sector(months)
            if sector_data and sum(sector_data.values()) > 0:
                import plotly.graph_objects as go
                # Show ALL sectors, sorted
                sorted_sector = dict(sorted(sector_data.items(), key=lambda x: x[1], reverse=True))
                
                # Dynamic height based on number of sectors
                num_sectors = len(sorted_sector)
                chart_height = max(400, num_sectors * 30 + 100)
                
                fig = go.Figure(data=[go.Bar(x=list(sorted_sector.values()), y=list(sorted_sector.keys()), orientation='h', marker_color='#8B5CF6')])
                fig.update_layout(title='Perizinan per Sektor', template='plotly_dark', height=chart_height, yaxis={'categoryorder': 'total ascending'})
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Data sektor kementerian/lembaga tidak tersedia atau kosong.")
            
            # 3.5 Jenis Perizinan
            st.markdown('<div class="section-title">3.5 Rekapitulasi Berdasarkan Jenis Perizinan</div>', 
                        unsafe_allow_html=True)
            jenis_data = pb_data.get_period_jenis_perizinan(months)
            if jenis_data:
                import plotly.graph_objects as go
                sorted_jenis = dict(sorted(jenis_data.items(), key=lambda x: x[1], reverse=True)[:10])
                fig = go.Figure(data=[go.Bar(x=list(sorted_jenis.values()), y=list(sorted_jenis.keys()), orientation='h', marker_color='#06B6D4')])
                fig.update_layout(title='Perizinan per Jenis (Top 10)', template='plotly_dark', height=400, yaxis={'categoryorder': 'total ascending'})
                st.plotly_chart(fig, use_container_width=True)
            
            # 3.6 Status Perizinan (NO Gubernur filter - all data)
            st.markdown('<div class="section-title">3.6 Rekapitulasi Berdasarkan Status Respon</div>', 
                        unsafe_allow_html=True)
            status_data = pb_data.get_period_status_perizinan(months)
            if status_data:
                import plotly.graph_objects as go
                
                col1, col2 = st.columns([1.2, 1])
                with col1:
                    # Bar chart for Status Respon
                    status_colors = {
                        'Izin Terbit/SS Terverifikasi': '#22C55E',
                        'Menunggu Verifikasi Persyaratan': '#EAB308', 
                        'Terbit Otomatis': '#3B82F6'
                    }
                    colors = [status_colors.get(k, '#8B5CF6') for k in status_data.keys()]
                    
                    fig = go.Figure(data=[go.Bar(
                        x=list(status_data.keys()), 
                        y=list(status_data.values()), 
                        marker_color=colors,
                        text=[f'{v:,}' for v in status_data.values()],
                        textposition='outside'
                    )])
                    fig.update_layout(
                        title=f'Jumlah Perizinan Berdasarkan Status Respon<br>Periode {report.period_name} Tahun {report.year}',
                        template='plotly_dark', 
                        height=400,
                        showlegend=False
                    )
                    st.plotly_chart(fig, use_container_width=True)
                
                with col2:
                    # Narrative interpretation
                    total_status = sum(status_data.values())
                    status_items = list(status_data.items())
                    
                    narrative = f"""
                    <b>Rekapitulasi Perizinan Berusaha Berbasis Risiko</b> Kewenangan Gubernur Provinsi Lampung 
                    periode {report.period_name} Tahun {report.year} berdasarkan Status Respon:<br><br>
                    """
                    
                    for status_name, count in status_items:
                        pct = count / total_status * 100 if total_status > 0 else 0
                        narrative += f"• Status <b>{status_name}</b> sebanyak <b>{count:,}</b> pemohon ({pct:.1f}%)<br>"
                    
                    narrative += f"<br>Total keseluruhan sebanyak <b>{total_status:,}</b> perizinan."
                    
                    st.markdown(f'<div class="narrative-box" style="margin-top: 1rem;">{narrative}</div>', unsafe_allow_html=True)
            
            # 3.7 Kewenangan (NO Gubernur filter - all data, grouped by DPMPTSP)
            st.markdown('<div class="section-title">3.7 Rekapitulasi Berdasarkan Kewenangan</div>', 
                        unsafe_allow_html=True)
            kew_data = pb_data.get_period_kewenangan(months)
            if kew_data:
                import plotly.graph_objects as go
                import pandas as pd
                
                # Sort all entries by total count
                sorted_kew = dict(sorted(kew_data.items(), key=lambda x: x[1], reverse=True))
                top_kew = dict(list(sorted_kew.items())[:20])  # Top 20 for chart
                total = sum(kew_data.values())
                sorted_items = sorted(kew_data.items(), key=lambda x: x[1], reverse=True)
                
                # Build monthly breakdown for each kewenangan
                kew_monthly = {}
                for month in months:
                    month_data = pb_data.monthly_kewenangan.get(month, {})
                    for kew, count in month_data.items():
                        if kew not in kew_monthly:
                            kew_monthly[kew] = {m: 0 for m in months}
                        kew_monthly[kew][month] = count
                
                # ========== HORIZONTAL BAR CHART (Full Width) ==========
                chart_height = max(500, len(top_kew) * 28)
                fig = go.Figure(data=[go.Bar(
                    x=list(top_kew.values()), 
                    y=list(top_kew.keys()), 
                    orientation='h',
                    marker=dict(
                        color=list(top_kew.values()),
                        colorscale=[[0, '#60A5FA'], [0.5, '#3B82F6'], [1, '#1E40AF']],
                        showscale=False
                    ),
                    text=[f'{v:,}' for v in top_kew.values()],
                    textposition='outside',
                    textfont=dict(size=11)
                )])
                fig.update_layout(
                    title=dict(
                        text=f'<b>JUMLAH PERIZINAN BERUSAHA BERBASIS RISIKO</b><br>PERIODE {report.period_name.upper()} TAHUN {report.year} BERDASARKAN KEWENANGAN',
                        font=dict(size=14)
                    ),
                    template='plotly_dark', 
                    height=chart_height,
                    yaxis=dict(categoryorder='total ascending', tickfont=dict(size=10)),
                    xaxis=dict(title='Jumlah Perizinan', tickformat=','),
                    margin=dict(l=10, r=60, t=80, b=40),
                    showlegend=False
                )
                st.plotly_chart(fig, use_container_width=True)
                
                # ========== DATA TABLE WITH MONTHLY BREAKDOWN ==========
                st.markdown(f'<div style="background: linear-gradient(90deg, #1E3A5F, #3B82F6); padding: 10px; border-radius: 8px 8px 0 0; margin-top: 1rem;"><b style="color: white;">📊 Tabel Rekapitulasi: {len(sorted_items)} Kewenangan | Total: {total:,} Perizinan</b></div>', unsafe_allow_html=True)
                
                # Build DataFrame with monthly columns
                table_data = []
                for idx, (kew, count) in enumerate(sorted_items, 1):
                    row = {'No': idx, 'Kewenangan': kew}
                    # Add monthly columns
                    for month in months:
                        row[month] = kew_monthly.get(kew, {}).get(month, 0)
                    row['JUMLAH'] = count
                    table_data.append(row)
                
                kew_df = pd.DataFrame(table_data)

                # Display with Plotly Table for better visibility and control
                header_values = ['<b>NO</b>', '<b>KEWENANGAN</b>'] + [f'<b>{m.upper()}</b>' for m in months] + ['<b>JUMLAH</b>']
                
                # Prepare column data
                cell_values = [
                    kew_df['No'].tolist(),
                    kew_df['Kewenangan'].tolist()
                ]
                for m in months:
                    cell_values.append(kew_df[m].tolist())
                cell_values.append(kew_df['JUMLAH'].tolist())
                
                # Create table figure
                table_fig = go.Figure(data=[go.Table(
                    header=dict(
                        values=header_values,
                        fill_color='#1E40AF',  # Blue header
                        align=['center', 'left'] + ['center'] * (len(months) + 1),
                        font=dict(color='white', size=12),
                        height=40
                    ),
                    cells=dict(
                        values=cell_values,
                        fill_color=['#0F172A', '#1E293B'], # Alternating dark rows
                        align=['center', 'left'] + ['center'] * (len(months) + 1),
                        font=dict(color='white', size=11),
                        height=30,
                        line_color='#334155'
                    )
                )])
                
                table_fig.update_layout(
                    margin=dict(l=0, r=0, t=10, b=0),
                    height=min(600, len(sorted_items) * 35 + 50),
                    paper_bgcolor='rgba(0,0,0,0)',
                    plot_bgcolor='rgba(0,0,0,0)'
                )
                
                st.plotly_chart(table_fig, use_container_width=True)
                
                # ========== NARRATIVE INTERPRETATION ==========
                top_3 = sorted_items[:3] if len(sorted_items) >= 3 else sorted_items
                narrative = f"""
                <b>Rekapitulasi Perizinan Berusaha Berbasis Risiko Berdasarkan Kewenangan</b> di Provinsi Lampung 
                periode {report.period_name} Tahun {report.year}. Dari rekapitulasi data tersebut, 
                kewenangan tertinggi adalah dari <b>{top_3[0][0]}</b> berjumlah <b>{top_3[0][1]:,}</b> perizinan
                """
                if len(top_3) > 1:
                    narrative += f", serta <b>{top_3[1][0]}</b> berjumlah <b>{top_3[1][1]:,}</b>"
                if len(top_3) > 2:
                    narrative += f" dan <b>{top_3[2][0]}</b> berjumlah <b>{top_3[2][1]:,}</b>"
                narrative += "."
                
                st.markdown(f'<div class="narrative-box">{narrative}</div>', unsafe_allow_html=True)
    
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
