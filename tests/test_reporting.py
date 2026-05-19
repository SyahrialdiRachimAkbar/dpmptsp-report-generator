import unittest

from app.config import NAMA_BULAN, TRIWULAN_KE_BULAN
from app.reporting import (
    build_comparison_context,
    resolve_reference_data,
    sum_month_values,
    validate_report_inputs,
)


class FakeUpload:
    def __init__(self, name, content=b"payload"):
        self.name = name
        self._content = content

    def getvalue(self):
        return self._content


class ReportingHelperTests(unittest.TestCase):
    def test_triwulan_context_uses_canonical_titlecase_months(self):
        context = build_comparison_context("Triwulan", "TW III", 2025)

        self.assertEqual(context["main_target_months"], ["Juli", "Agustus", "September"])
        self.assertEqual(context["yoy_curr_label"], "TW III 2025")
        self.assertEqual(context["qoq_prev_label"], "TW II 2025")
        self.assertNotIn("juli", context["main_target_months"])

    def test_semester_and_tahunan_contexts_use_config_months(self):
        semester = build_comparison_context("Semester", "Semester I", 2025)
        tahunan = build_comparison_context("Tahunan", "2025", 2025)

        self.assertEqual(semester["main_target_months"], TRIWULAN_KE_BULAN["TW I"] + TRIWULAN_KE_BULAN["TW II"])
        self.assertEqual(semester["qoq_curr_label"], "TW II 2025")
        self.assertEqual(semester["qoq_prev_label"], "TW I 2025")
        self.assertEqual(tahunan["main_target_months"], NAMA_BULAN)
        self.assertEqual(tahunan["yoy_curr_label"], "Semester II 2025")

    def test_nib_month_sum_does_not_regress_to_lowercase_keys(self):
        context = build_comparison_context("Triwulan", "TW III", 2025)
        monthly_totals = {"Juli": 7, "Agustus": 8, "September": 9}

        self.assertEqual(sum_month_values(monthly_totals, context["main_target_months"]), 24)
        self.assertEqual(sum_month_values(monthly_totals, ["juli", "agustus", "september"]), 0)

    def test_nib_is_required_for_report_generation(self):
        valid, message = validate_report_inputs({"nib_ref_file": FakeUpload("NIB 2025.xlsx")})
        self.assertTrue(valid)
        self.assertEqual(message, "")

        valid, message = validate_report_inputs({
            "nib_ref_file": FakeUpload("NIB 2025.xlsx"),
            "pb_oss_ref_file": FakeUpload("PB OSS 2025.xlsx"),
            "proyek_ref_file": FakeUpload("PROYEK OSS 2025.xlsx"),
        })
        self.assertTrue(valid)
        self.assertEqual(message, "")

        valid, message = validate_report_inputs({
            "pb_oss_ref_file": FakeUpload("PB OSS 2025.xlsx"),
            "proyek_ref_file": FakeUpload("PROYEK OSS 2025.xlsx"),
        })
        self.assertFalse(valid)
        self.assertIn("NIB", message)

    def test_previous_pb_reference_data_uses_cache_or_real_loader(self):
        cached = object()
        state = {"prev_pb_data": cached}
        calls = []

        result = resolve_reference_data(
            state,
            "prev_pb_data",
            "pb_oss_prev_ref_file",
            lambda content, name, year: calls.append((content, name, year)),
            2024,
        )

        self.assertIs(result, cached)
        self.assertEqual(calls, [])

        loaded = object()
        state = {"pb_oss_prev_ref_file": FakeUpload("PB OSS 2024.xlsx", b"pb")}

        def loader(content, name, year):
            calls.append((content, name, year))
            return loaded

        result = resolve_reference_data(state, "prev_pb_data", "pb_oss_prev_ref_file", loader, 2024)

        self.assertIs(result, loaded)
        self.assertIs(state["prev_pb_data"], loaded)
        self.assertEqual(calls[-1], (b"pb", "PB OSS 2024.xlsx", 2024))

    def test_previous_proyek_reference_data_uses_real_loader_contract(self):
        loaded = object()
        state = {"proyek_prev_ref_file": FakeUpload("PROYEK OSS 2024.xlsx", b"proyek")}
        calls = []

        def loader(content, name, year):
            calls.append((content, name, year))
            return loaded

        result = resolve_reference_data(state, "prev_proyek_data", "proyek_prev_ref_file", loader, 2024)

        self.assertIs(result, loaded)
        self.assertIs(state["prev_proyek_data"], loaded)
        self.assertEqual(calls, [(b"proyek", "PROYEK OSS 2024.xlsx", 2024)])


if __name__ == "__main__":
    unittest.main()
