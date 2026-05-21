import io
import struct
import unittest
import zipfile
import zlib
import xml.etree.ElementTree as ET
from types import SimpleNamespace

from app.export.docx_exporter import DOCX_AVAILABLE, WordExporter
from app.narrative.generator import Narrative


def _png_bytes(width=12, height=12, color=(30, 58, 95)):
    """Return a tiny valid RGB PNG without requiring Pillow."""
    raw_rows = [b"\x00" + bytes(color) * width for _ in range(height)]
    raw = b"".join(raw_rows)

    def chunk(tag, data):
        checksum = zlib.crc32(tag + data) & 0xFFFFFFFF
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", checksum)

    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(raw))
        + chunk(b"IEND", b"")
    )


def _base_report():
    return SimpleNamespace(period_name="TW I", year=2025)


def _base_stats():
    return {
        "total_nib": 10,
        "pm_distribution": {"PMDN": 8, "PMA": 2},
        "pelaku_usaha_distribution": {"UMK": 7},
        "top_5_locations": [],
    }


def _base_narrative(**overrides):
    values = {
        "pendahuluan": "Pendahuluan.",
        "rekapitulasi_nib": "Rekapitulasi NIB.",
        "rekapitulasi_kab_kota": "Rekapitulasi kabupaten/kota.",
        "status_pm": "Status PM.",
        "pelaku_usaha": "Pelaku usaha.",
        "kesimpulan": "Kesimpulan.",
    }
    values.update(overrides)
    return Narrative(**values)


def _docx_text_and_media(docx_bytes):
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as docx:
        document_xml = docx.read("word/document.xml")
        root = ET.fromstring(document_xml)
        ns_text = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"
        text = "\n".join(node.text or "" for node in root.iter(ns_text))
        media = [name for name in docx.namelist() if name.startswith("word/media/")]
    return text, media


def _chart_images(keys):
    charts = {}
    for index, key in enumerate(keys):
        charts[key] = _png_bytes(
            color=(
                30 + (index * 29) % 180,
                58 + (index * 17) % 160,
                95 + (index * 23) % 130,
            )
        )
    return charts


@unittest.skipUnless(DOCX_AVAILABLE, "python-docx is required for Word export tests")
class WordExporterParityTests(unittest.TestCase):
    def setUp(self):
        self.exporter = WordExporter(logo_path=None)

    def export(self, charts, narratives=None):
        return self.exporter.export_report(
            _base_report(),
            _base_stats(),
            narratives or _base_narrative(),
            charts,
        )

    def test_section_3_1_includes_main_comparison_narrative_and_table(self):
        charts = _chart_images([
            "pb_monthly",
            "pb_kab_kota",
            "pb_total_yoy",
            "pb_total_qoq",
            "pb_kab_table",
        ])
        narratives = _base_narrative(
            pb_periode_lokasi="Narasi periode dan lokasi PB OSS."
        )

        text, media = _docx_text_and_media(self.export(charts, narratives))

        self.assertIn("3.1 Rekapitulasi Berdasarkan Periode dan Lokasi Usaha di Kabupaten/Kota", text)
        self.assertIn("Narasi periode dan lokasi PB OSS.", text)
        self.assertEqual(len(media), len(charts))

    def test_sections_3_2_and_3_3_include_three_visuals_interpretation_and_tables(self):
        charts = _chart_images([
            "pb_pm_monthly",
            "pb_pm_yoy",
            "pb_pm_qoq",
            "pb_pm_table",
            "pb_risk",
            "pb_risk_yoy",
            "pb_risk_qoq",
            "pb_risk_table",
        ])
        narratives = _base_narrative(
            pb_status_pm="Interpretasi status PM PB OSS.",
            pb_risiko="Interpretasi risiko PB OSS.",
        )

        text, media = _docx_text_and_media(self.export(charts, narratives))

        self.assertIn("3.2 Rekapitulasi Berdasarkan Status Penanaman Modal", text)
        self.assertIn("Interpretasi status PM PB OSS.", text)
        self.assertIn("3.3 Rekapitulasi Berdasarkan Tingkat Risiko", text)
        self.assertIn("Interpretasi risiko PB OSS.", text)
        self.assertEqual(len(media), len(charts))

    def test_section_3_4_renders_independently_without_risk_section(self):
        charts = _chart_images(["pb_sector", "pb_sector_table"])
        narratives = _base_narrative(pb_sektor="Interpretasi sektor PB OSS.")

        text, media = _docx_text_and_media(self.export(charts, narratives))
        body_section_3 = text.split("3. Perizinan Berusaha Berbasis Risiko", 2)[-1]

        self.assertNotIn("3.3 Rekapitulasi Berdasarkan Tingkat Risiko", body_section_3)
        self.assertIn("3.4 Top 10 Sektor Perizinan", body_section_3)
        self.assertIn("Interpretasi sektor PB OSS.", body_section_3)
        self.assertEqual(len(media), len(charts))

    def test_optional_export_charts_are_skipped_cleanly(self):
        text, media = _docx_text_and_media(self.export({}))

        self.assertIn("Kesimpulan", text)
        self.assertEqual(media, [])

    def test_export_narrative_html_is_cleaned_for_word_text(self):
        narratives = _base_narrative(
            proyek_skala_usaha="<b>Skala Usaha</b><br>Naik &amp; terkendali."
        )

        text, _ = _docx_text_and_media(self.export({}, narratives))

        self.assertIn("Skala Usaha", text)
        self.assertIn("Naik & terkendali.", text)
        self.assertNotIn("<b>", text)


if __name__ == "__main__":
    unittest.main()
