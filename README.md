# DPMPTSP Report Generator

Aplikasi Streamlit untuk menghasilkan laporan triwulan dan tahunan Dinas Penanaman Modal dan Pelayanan Terpadu Satu Pintu (DPMPTSP) Provinsi Lampung.

## Fitur

- 📊 **Dashboard Interaktif** - Visualisasi data NIB (Nomor Induk Berusaha) per bulan
- 📈 **Grafik & Chart** - Distribusi PMA/PMDN, Skala Usaha, Sektor Risiko
- 📝 **Generasi Narasi Otomatis** - Analisis tren dan insight
- 📄 **Ekspor Laporan** - Format PDF dan Word (DOCX)
- 🗓️ **Laporan Triwulan & Tahunan** - Agregasi data per periode

## Struktur Proyek

```
├── app/
│   ├── main.py              # Aplikasi Streamlit utama
│   ├── config.py            # Konfigurasi aplikasi
│   ├── data/
│   │   ├── loader.py        # Pembaca file Excel
│   │   └── aggregator.py    # Agregator data
│   ├── export/
│   │   ├── pdf_exporter.py  # Ekspor ke PDF
│   │   └── docx_exporter.py # Ekspor ke Word
│   ├── narrative/
│   │   └── generator.py     # Generator narasi otomatis
│   └── visualization/
│       └── charts.py        # Komponen chart/grafik
└── requirements.txt
```

## Instalasi

1. Clone repository:
```bash
git clone https://github.com/SyahrialdiRachimAkbar/dpmptsp-report-generator.git
cd dpmptsp-report-generator
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Jalankan aplikasi:
```bash
streamlit run app/main.py
```

## Penggunaan

1. Upload file Excel data OSS (format bulanan atau triwulan)
2. Pilih periode laporan (Triwulan I/II/III/IV atau Tahunan)
3. Lihat dashboard dan visualisasi data
4. Ekspor laporan ke PDF atau Word

## Requirements

- Python 3.8+
- Streamlit
- Pandas
- Plotly
- ReportLab (PDF)
- python-docx (Word)

## Lisensi

MIT License
