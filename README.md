# DPMPTSP Report Generator

<div align="center">

![Version](https://img.shields.io/badge/version-1.0.0-blue.svg)
![Python](https://img.shields.io/badge/python-3.8%2B-blue.svg)
![Streamlit](https://img.shields.io/badge/streamlit-1.28%2B-red.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)

**Aplikasi Streamlit untuk menghasilkan laporan triwulan dan tahunan otomatis**  
**Dinas Penanaman Modal dan Pelayanan Terpadu Satu Pintu (DPMPTSP) Provinsi Lampung**

[Fitur](#-fitur) • [Instalasi](#-instalasi) • [Penggunaan](#-penggunaan) • [Struktur Data](#-struktur-data) • [Dokumentasi](#-dokumentasi)

</div>

---

## 📋 Tentang Proyek

Sistem pelaporan otomatis untuk DPMPTSP Provinsi Lampung yang memproses data OSS (Online Single Submission) dan menghasilkan laporan periodik dengan visualisasi, narasi otomatis, dan ekspor ke PDF/Word.

### Apa yang Dilaporkan?

1. **NIB (Nomor Induk Berusaha)** - Data registrasi usaha
2. **Perizinan Berusaha** - Data izin usaha berdasarkan risiko dan sektor
3. **Realisasi Investasi** - Data proyek investasi dan penyerapan tenaga kerja

---

## ✨ Fitur

### 🎨 Antarmuka Modern
- **Dark Mode Support** - Deteksi otomatis tema sistem
- **Glassmorphism Design** - UI modern dengan gradien dan shadow halus
- **Responsive Layout** - Tampilan optimal di berbagai ukuran layar
- **Interactive Charts** - Visualisasi interaktif dengan Plotly

### 📊 Visualisasi Data
- Bar chart dengan trendline
- Donut chart untuk distribusi
- Comparison charts (Q-o-Q, Y-o-Y)
- Geographic distribution maps
- Horizontal bar rankings

### 📝 Narasi Otomatis
- Analisis otomatis dalam Bahasa Indonesia formal
- Insight per-chart yang kontekstual
- Ringkasan periode dan tren
- Perbandingan quarter-over-quarter dan year-over-year

### 📄 Ekspor Laporan
- **PDF Export** - Format profesional dengan WeasyPrint
- **Word/DOCX Export** - Format editable dengan python-docx
- Embedded charts dan tabel
- Layout terstruktur dengan nomor section

### ⚡ Performa
- Streamlit caching untuk loading cepat
- Smart sheet detection
- Efficient data aggregation

---

## 🏗️ Struktur Laporan

Laporan terbagi menjadi 3 section utama:

### Section 1: NIB (Nomor Induk Berusaha)
- 1.1 Rekapitulasi NIB per Bulan
- 1.2 Rekapitulasi NIB per Kabupaten/Kota
- 1.3 Status Penanaman Modal (PMA/PMDN)
- 1.4 Pelaku Usaha (UMK/NON-UMK)
- 1.5 Summary Metrics

### Section 2: Realisasi Investasi (Proyek)
- 2.1 Tren Proyek Bulanan + Distribusi per Kab/Kota
- 2.2 Investasi per Wilayah
- 2.3 Distribusi Skala Usaha
- 2.4 Perbandingan PMA vs PMDN
- 2.5 Penyerapan Tenaga Kerja per Kab/Kota
- 2.6 Perbandingan Quarter-over-Quarter
- 2.7 Perbandingan Year-over-Year

### Section 3: Perizinan Berusaha (PB OSS)
- 3.1 Tren Perizinan Bulanan
- 3.2 Distribusi per Kabupaten/Kota
- 3.3 Status Penanaman Modal
- 3.4 Tingkat Risiko
- 3.5 Jenis Perizinan
- 3.6 Status Perizinan
- 3.7 Kewenangan (Detail DPMPTSP)

---

## 🚀 Instalasi

### Prerequisites
- Python 3.8 atau lebih tinggi
- pip (Python package manager)

### Langkah Instalasi

1. **Clone repository:**
```bash
git clone https://github.com/SyahrialdiRachimAkbar/dpmptsp-report-generator.git
cd dpmptsp-report-generator
```

2. **Install dependencies:**
```bash
pip install -r requirements.txt
```

3. **Jalankan aplikasi:**
```bash
streamlit run app/main.py
```

4. **Akses di browser:**
```
http://localhost:8501
```

---

## 💻 Penggunaan

### 1. Upload File Referensi

Upload 3 file Excel (.xlsx):

- **NIB {Tahun}.xlsx** - Data NIB dari OSS
- **PB OSS {Tahun}.xlsx** - Data Perizinan Berusaha
- **PROYEK OSS {Tahun}.xlsx** - Data Realisasi Investasi

> **Catatan:** Sistem akan otomatis mendeteksi tahun dari nama file

### 2. Pilih Periode Laporan

**Triwulan:**
- TW I (Januari - Maret)
- TW II (April - Juni)
- TW III (Juli - September)
- TW IV (Oktober - Desember)

**Semester:**
- Semester I (Januari - Juni)
- Semester II (Juli - Desember)

**Tahunan:**
- Tahun penuh (Januari - Desember)

### 3. Generate Laporan

Klik tombol **"🚀 Generate Laporan"** untuk memproses data.

### 4. Review & Export

- Preview laporan di dashboard
- Export ke PDF atau Word
- Download untuk distribusi

---

## 📁 Struktur Proyek

```
dpmptsp-report-generator/
│
├── app/
│   ├── main.py                    # Aplikasi Streamlit utama (2,337 lines)
│   ├── config.py                  # Konfigurasi dan konstanta
│   │
│   ├── data/
│   │   ├── loader.py              # Legacy monthly data loader
│   │   ├── aggregator.py          # Period aggregation logic
│   │   └── reference_loader.py    # NIB/PB/PROYEK file parser ⭐
│   │
│   ├── visualization/
│   │   └── charts.py              # Plotly chart generation (1,127 lines)
│   │
│   ├── narrative/
│   │   └── generator.py           # Auto-narrative generation (639 lines)
│   │
│   └── export/
│       ├── pdf_exporter.py        # PDF report export
│       └── docx_exporter.py       # Word report export
│
├── DATA OSS 2025/                 # Sample data directory
├── NIB 2025.xlsx                  # Reference data files
├── PB OSS 2025.xlsx
├── PROYEK OSS 2025.xlsx
│
├── create_dummy_data.py           # Utility: Generate test data
├── verify_nib_2024.py             # Utility: Verify NIB formatting
├── requirements.txt               # Python dependencies
└── README.md                      # This file
```

---

## 📊 Struktur Data

### Format File NIB
**File:** `NIB {Tahun}.xlsx`  
**Sheet:** Sheet 1 (raw data)

| Kolom | Deskripsi |
|-------|-----------|
| `nib` | Nomor Induk Berusaha (string) |
| `Day of tanggal_terbit_oss` | Tanggal terbit |
| `kab_kota` | Kabupaten/Kota |
| `status_penanaman_modal` | PMA atau PMDN |
| `uraian_skala_usaha` | Mikro/Kecil/Menengah/Besar |

### Format File PB OSS
**File:** `PB OSS {Tahun}.xlsx`  
**Sheet:** Sheet 1 (raw data)

| Kolom | Deskripsi |
|-------|-----------|
| `Day of tgl_izin` | Tanggal izin |
| `risiko` | R/MR/MT/T |
| `sektor` | Sektor usaha |
| `kewenangan` | Gubernur/Pusat/Kab-Kota |
| `status_perizinan` | Status izin |

### Format File PROYEK
**File:** `PROYEK OSS {Tahun}.xlsx`  
**Sheet:** Sheet 1 (raw data)

| Kolom | Deskripsi |
|-------|-----------|
| `tanggal_pengajuan_proyek` | Tanggal pengajuan |
| `Jumlah Investasi` | Nilai investasi (Rp) |
| `Status PM` | PMA atau PMDN |
| `Kab Kota Usaha` | Lokasi proyek |
| `TKI` | Tenaga Kerja Indonesia |
| `TKA` | Tenaga Kerja Asing |

---

## 🔧 Business Rules

### NIB Processing
- **Deduplication:** Per-month (NIB yang sama di bulan berbeda dihitung terpisah)
- **Aggregation:** Sum of monthly unique NIB counts
- **Data Type:** NIB dibaca sebagai string (preserve leading zeros)

### PB OSS Processing
- **Filtering:** Sebagian besar section filter `Kewenangan = Gubernur`
- **Exception:** Section 3.6 (Status Perizinan) dan 3.7 (Kewenangan) menggunakan data unfiltered
- **Counting:** Total permits (tidak unique NIB)

### PROYEK Processing
- **Filtering:** `Kewenangan = Gubernur`
- **Aggregation:** SUM all investment amounts (no deduplication)
- **Labor:** TKI + TKA = Total tenaga kerja

---

## 🛠️ Development

### Tech Stack

```python
# Core
streamlit >= 1.28.0          # Web framework
pandas >= 2.0.0              # Data manipulation
openpyxl >= 3.1.0            # Excel support

# Visualization
plotly >= 5.18.0             # Interactive charts
folium >= 0.15.0             # Maps

# Export
weasyprint >= 60.0           # HTML to PDF
jinja2 >= 3.1.0              # Templating
python-docx >= 1.1.0         # Word export

# Utils
Pillow >= 10.0.0             # Image processing
```

### Key Components

**ReferenceDataLoader** (`app/data/reference_loader.py`)
- Smart sheet detection algorithm
- Date parsing with Indonesian month names
- Monthly aggregation with flexible queries
- Dual-pass filtering for kewenangan

**ChartGenerator** (`app/visualization/charts.py`)
- Plotly-based visualizations
- Dark mode support
- Gradient colors and trendlines
- Indonesian language labels

**NarrativeGenerator** (`app/narrative/generator.py`)
- Context-aware narratives
- Formal Indonesian language
- Per-chart interpretations
- Automated insights

---

## 🧪 Utilities

### Generate Test Data (2024)
```bash
python create_dummy_data.py
```
Membuat data 2024 dari data 2025 dengan:
- Shift tanggal 366 hari (leap year)
- Randomisasi nilai ±20%
- Preserve NIB formatting

### Verify NIB Format
```bash
python verify_nib_2024.py
```
Memverifikasi NIB tersimpan sebagai string (bukan float).

---

## 📝 Recent Updates

### Version 1.0.0 (January 2026)

**Major Features:**
- ✅ Reference data loader architecture
- ✅ 3-section report structure
- ✅ Dark mode support
- ✅ Per-chart narrative generation
- ✅ Dual-pass filtering for kewenangan
- ✅ Summary metric cards

**Bug Fixes:**
- 🐛 Fixed kewenangan data source (Section 3.7)
- 🐛 Fixed UMK and PM metric data sources
- 🐛 Improved dark mode table visibility (Plotly tables)
- 🐛 Enhanced column detection algorithm

**Performance:**
- ⚡ Streamlit caching for data loaders
- ⚡ Optimized chart rendering
- ⚡ Efficient period aggregation

---

## 🤝 Contributing

Contributions are welcome! Please:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'feat: add AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

### Commit Convention
- `feat:` - New feature
- `fix:` - Bug fix
- `perf:` - Performance improvement
- `docs:` - Documentation update
- `style:` - Code style changes
- `refactor:` - Code refactoring
- `test:` - Test updates

---

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

## 👨‍💻 Author

**Syahrialdi Rachim Akbar**

- GitHub: [@SyahrialdiRachimAkbar](https://github.com/SyahrialdiRachimAkbar)
- Repository: [dpmptsp-report-generator](https://github.com/SyahrialdiRachimAkbar/dpmptsp-report-generator)

---

## 🙏 Acknowledgments

- **DPMPTSP Provinsi Lampung** - For the use case and data requirements
- **Streamlit** - For the amazing web framework
- **Plotly** - For beautiful interactive visualizations

---

## 📞 Support

Jika Anda mengalami masalah atau memiliki pertanyaan:

1. Buka [GitHub Issues](https://github.com/SyahrialdiRachimAkbar/dpmptsp-report-generator/issues)
2. Cek [Documentation](#-dokumentasi)
3. Contact the maintainer

---

## 🗺️ Roadmap

### Future Enhancements
- [ ] Database integration (replace Excel with DB)
- [ ] Automated anomaly detection
- [ ] Predictive analytics for trend forecasting
- [ ] API layer for external integrations
- [ ] Batch report generation
- [ ] Email notification system
- [ ] Multi-user authentication
- [ ] Report scheduling

---

<div align="center">

**Made with ❤️ for DPMPTSP Provinsi Lampung**

⭐ Star this repo if you find it helpful!

</div>
