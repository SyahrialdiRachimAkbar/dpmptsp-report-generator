# DPMPTSP Automated Reporting System - Configuration

from pathlib import Path
from typing import Dict, List

# Base paths
BASE_DIR = Path(__file__).parent.parent
DATA_DIR = BASE_DIR / "DATA OSS 2025"
STATIC_DIR = Path(__file__).parent / "static"
STORAGE_DIR = BASE_DIR / "storage" / "uploads"

# Logo path
LOGO_PATH = DATA_DIR / "logo.webp"

# Kabupaten/Kota di Provinsi Lampung
KABUPATEN_KOTA: List[str] = [
    "Kab. Lampung Barat",
    "Kab. Lampung Selatan", 
    "Kab. Lampung Tengah",
    "Kab. Lampung Timur",
    "Kab. Lampung Utara",
    "Kab. Mesuji",
    "Kab. Pesawaran",
    "Kab. Pesisir Barat",
    "Kab. Pringsewu",
    "Kab. Tanggamus",
    "Kab. Tulang Bawang",
    "Kab. Tulang Bawang Barat",
    "Kab. Way Kanan",
    "Kota Bandar Lampung",
    "Kota Metro",
]

# Mapping bulan ke Triwulan
BULAN_KE_TRIWULAN: Dict[str, str] = {
    "Januari": "TW I",
    "Februari": "TW I", 
    "Maret": "TW I",
    "April": "TW II",
    "Mei": "TW II",
    "Juni": "TW II",
    "Juli": "TW III",
    "Agustus": "TW III",
    "September": "TW III",
    "Oktober": "TW IV",
    "November": "TW IV",
    "Desember": "TW IV",
}

# Mapping bulan ke Semester
BULAN_KE_SEMESTER: Dict[str, str] = {
    "Januari": "Semester I",
    "Februari": "Semester I",
    "Maret": "Semester I",
    "April": "Semester I",
    "Mei": "Semester I",
    "Juni": "Semester I",
    "Juli": "Semester II",
    "Agustus": "Semester II",
    "September": "Semester II",
    "Oktober": "Semester II",
    "November": "Semester II",
    "Desember": "Semester II",
}

# Triwulan ke bulan
TRIWULAN_KE_BULAN: Dict[str, List[str]] = {
    "TW I": ["Januari", "Februari", "Maret"],
    "TW II": ["April", "Mei", "Juni"],
    "TW III": ["Juli", "Agustus", "September"],
    "TW IV": ["Oktober", "November", "Desember"],
}

# Nama bulan dalam Bahasa Indonesia
NAMA_BULAN: List[str] = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
]

# Skala Usaha
SKALA_USAHA: List[str] = [
    "Usaha Mikro",
    "Usaha Kecil", 
    "Usaha Menengah",
    "Usaha Besar"
]

# Status Penanaman Modal
STATUS_PM: List[str] = ["PMA", "PMDN"]

# Jenis Pelaku Usaha
JENIS_PELAKU_USAHA: List[str] = ["UMK", "NON-UMK"]

# Warna untuk visualisasi
COLORS = {
    "primary": "#1e3a5f",      # Biru tua
    "secondary": "#3d7ea6",    # Biru medium
    "accent": "#5cb85c",       # Hijau
    "warning": "#f0ad4e",      # Kuning/Orange
    "danger": "#d9534f",       # Merah
    "gradient_start": "#e8f4f8",  # Biru muda
    "gradient_end": "#1e3a5f",    # Biru tua
}

# Konfigurasi export
EXPORT_CONFIG = {
    "page_size": "A4",
    "margin_top": "1cm",
    "margin_bottom": "1cm",
    "margin_left": "1.5cm",
    "margin_right": "1.5cm",
}
