"""Small reporting helpers shared by the Streamlit UI and exporters."""

from typing import Any, Callable, Dict, Iterable, Optional

from app.config import NAMA_BULAN, TRIWULAN_KE_BULAN


TRIWULAN_ORDER = ["TW I", "TW II", "TW III", "TW IV"]

SEMESTER_KE_BULAN = {
    "Semester I": TRIWULAN_KE_BULAN["TW I"] + TRIWULAN_KE_BULAN["TW II"],
    "Semester II": TRIWULAN_KE_BULAN["TW III"] + TRIWULAN_KE_BULAN["TW IV"],
}

PERIOD_KE_BULAN = {
    **TRIWULAN_KE_BULAN,
    **SEMESTER_KE_BULAN,
}


def build_comparison_context(period_type: str, period_name: str, year: int) -> Dict[str, Any]:
    """Build canonical month ranges and labels for YoY/QoQ comparisons."""
    context = {
        "main_target_months": [],
        "yoy_curr_months": [],
        "yoy_prev_months": [],
        "yoy_curr_label": "",
        "yoy_prev_label": "",
        "qoq_curr_months": [],
        "qoq_prev_months": [],
        "qoq_curr_label": "",
        "qoq_prev_label": "",
        "qoq_prev_year_required": False,
        "has_prev_q_data": False,
    }

    if period_type == "Triwulan":
        context["main_target_months"] = TRIWULAN_KE_BULAN.get(period_name, [])
        context["yoy_curr_months"] = context["main_target_months"]
        context["yoy_prev_months"] = context["main_target_months"]
        context["yoy_curr_label"] = f"{period_name} {year}"
        context["yoy_prev_label"] = f"{period_name} {year - 1}"

        if period_name in TRIWULAN_ORDER:
            current_idx = TRIWULAN_ORDER.index(period_name)
            context["qoq_curr_months"] = context["main_target_months"]
            context["qoq_curr_label"] = f"{period_name} {year}"
            prev_period = TRIWULAN_ORDER[current_idx - 1] if current_idx > 0 else "TW IV"
            prev_year = year if current_idx > 0 else year - 1
            context["qoq_prev_months"] = TRIWULAN_KE_BULAN[prev_period]
            context["qoq_prev_label"] = f"{prev_period} {prev_year}"
            context["qoq_prev_year_required"] = current_idx == 0

    elif period_type == "Semester":
        context["main_target_months"] = SEMESTER_KE_BULAN.get(period_name, [])
        if period_name == "Semester I":
            current_tw = "TW II"
            previous_tw = "TW I"
        else:
            current_tw = "TW IV"
            previous_tw = "TW III"

        context["yoy_curr_months"] = TRIWULAN_KE_BULAN[current_tw]
        context["yoy_prev_months"] = TRIWULAN_KE_BULAN[current_tw]
        context["yoy_curr_label"] = f"{current_tw} {year}"
        context["yoy_prev_label"] = f"{current_tw} {year - 1}"
        context["qoq_curr_months"] = TRIWULAN_KE_BULAN[current_tw]
        context["qoq_prev_months"] = TRIWULAN_KE_BULAN[previous_tw]
        context["qoq_curr_label"] = f"{current_tw} {year}"
        context["qoq_prev_label"] = f"{previous_tw} {year}"

    elif period_type == "Tahunan":
        context["main_target_months"] = list(NAMA_BULAN)
        context["yoy_curr_months"] = SEMESTER_KE_BULAN["Semester II"]
        context["yoy_prev_months"] = SEMESTER_KE_BULAN["Semester II"]
        context["yoy_curr_label"] = f"Semester II {year}"
        context["yoy_prev_label"] = f"Semester II {year - 1}"
        context["qoq_curr_months"] = SEMESTER_KE_BULAN["Semester II"]
        context["qoq_prev_months"] = SEMESTER_KE_BULAN["Semester I"]
        context["qoq_curr_label"] = f"Semester II {year}"
        context["qoq_prev_label"] = f"Semester I {year}"

    return context


def sum_month_values(monthly_values: Dict[str, float], months: Iterable[str]) -> float:
    """Sum a month-keyed mapping using canonical Indonesian month names."""
    return sum(monthly_values.get(month, 0) for month in months)


def has_required_nib(session_state: Any) -> bool:
    """Return whether the base NIB upload required for a report is present."""
    return bool(session_state.get("nib_ref_file"))


def validate_report_inputs(session_state: Any):
    """Validate uploaded files for report generation."""
    if not has_required_nib(session_state):
        return False, "Upload file NIB terlebih dahulu. PB OSS dan PROYEK bersifat opsional."
    return True, ""


def resolve_reference_data(
    session_state: Any,
    data_key: str,
    file_key: str,
    loader: Callable[[bytes, str, int], Any],
    year: int,
) -> Optional[Any]:
    """Return cached reference data, or load it from an uploaded file."""
    cached = session_state.get(data_key)
    if cached is not None:
        return cached

    uploaded_file = session_state.get(file_key)
    if not uploaded_file:
        return None

    loaded = loader(uploaded_file.getvalue(), uploaded_file.name, year)
    if loaded is not None:
        try:
            session_state[data_key] = loaded
        except Exception:
            pass
    return loaded


def report_to_dataframe(report: Any, aggregator: Optional[Any] = None):
    """Convert a PeriodReport to a DataFrame with a local fallback aggregator."""
    if aggregator is None:
        from app.data.aggregator import DataAggregator

        aggregator = DataAggregator()
    return aggregator.to_dataframe(report)
