import tempfile
import unittest
from pathlib import Path

import pandas as pd

from app.cache import CACHE_VERSION, get_cache_key, load_or_build
from app.data.reference_loader import (
    NIBReferenceData,
    PBOSSReferenceData,
    ProyekReferenceData,
    ReferenceDataLoader,
)


class PersistentCacheTests(unittest.TestCase):
    def test_cache_key_is_stable_and_versioned(self):
        content = b"same file bytes"

        key_one = get_cache_key("nib", content, 2025)
        key_two = get_cache_key("nib", content, 2025)
        key_new_version = get_cache_key("nib", content, 2025, version=f"{CACHE_VERSION}-next")

        self.assertEqual(key_one, key_two)
        self.assertNotEqual(key_one, key_new_version)

    def test_cache_miss_then_hit_preserves_reference_object_methods(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            import app.cache as cache_module

            old_cache_dir = cache_module.CACHE_DIR
            cache_module.CACHE_DIR = Path(tmpdir)
            calls = []

            def builder(content, filename, year):
                calls.append((content, filename, year))
                data = NIBReferenceData(year=year)
                data.monthly_totals = {"Januari": 3, "Februari": 4}
                return data

            try:
                first = load_or_build("nib", b"nib-bytes", "NIB 2025.xlsx", 2025, builder)
                second = load_or_build("nib", b"nib-bytes", "NIB 2025.xlsx", 2025, builder)
            finally:
                cache_module.CACHE_DIR = old_cache_dir

        self.assertEqual(first.status, "parsed")
        self.assertEqual(second.status, "cache")
        self.assertEqual(len(calls), 1)
        self.assertEqual(second.data.get_period_total(["Januari", "Februari"]), 7)

    def test_cached_pb_and_proyek_objects_keep_breakdown_methods(self):
        pb = PBOSSReferenceData(year=2025)
        pb.monthly_permits = {"Januari": 5}
        pb.monthly_status_pm = {"Januari": {"PMA": 2, "PMDN": 3}}

        proyek = ProyekReferenceData(year=2025)
        proyek.monthly_projects = {"Januari": 9}
        proyek.monthly_by_wilayah = {"Januari": {"Kota Metro": 1000.0}}

        self.assertEqual(pb.get_period_permits(["Januari"]), 5)
        self.assertEqual(pb.get_period_status_pm(["Januari"]), {"PMA": 2, "PMDN": 3})
        self.assertEqual(proyek.get_period_projects(["Januari"]), 9)
        self.assertEqual(proyek.get_period_by_wilayah(["Januari"]), {"Kota Metro": 1000.0})


class ReferenceLoaderDateTests(unittest.TestCase):
    def test_vectorized_date_parsing_supports_existing_formats(self):
        loader = ReferenceDataLoader()
        excel_serial_april = (pd.Timestamp("2025-04-01") - pd.Timestamp("1899-12-30")).days
        parsed_months = loader._month_series(pd.Series([
            pd.Timestamp("2025-01-02"),
            "02/01/2025",
            "2025-03-10",
            excel_serial_april,
            "18 Agustus 2025",
        ])).tolist()

        self.assertEqual(parsed_months, ["Januari", "Februari", "Maret", "April", "Agustus"])


if __name__ == "__main__":
    unittest.main()
