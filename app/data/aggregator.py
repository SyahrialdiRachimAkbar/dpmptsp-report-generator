"""
Data Aggregator Module for DPMPTSP Reporting System

This module handles aggregation of NIB data for different periods:
- Triwulan (Quarterly): Q1, Q2, Q3, Q4
- Semester: S1, S2  
- Tahunan (Annual): Full year
"""

import pandas as pd
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass, field
from collections import defaultdict

from .loader import DataLoader, NIBData


@dataclass 
class AggregatedNIBData:
    """Aggregated NIB data for a specific period"""
    kabupaten_kota: str
    period_data: Dict[str, int] = field(default_factory=dict)  # month -> total
    pma_total: int = 0
    pmdn_total: int = 0
    usaha_mikro_total: int = 0
    usaha_kecil_total: int = 0
    usaha_menengah_total: int = 0
    usaha_besar_total: int = 0
    grand_total: int = 0
    
    @property
    def umk_total(self) -> int:
        return self.usaha_mikro_total + self.usaha_kecil_total
    
    @property
    def non_umk_total(self) -> int:
        return self.usaha_menengah_total + self.usaha_besar_total


@dataclass
class PeriodReport:
    """Complete report data for a specific period"""
    period_type: str  # "Triwulan", "Semester", "Tahunan"
    period_name: str  # "TW I", "TW II", etc.
    year: int
    months_included: List[str]
    
    # Aggregated data by Kabupaten/Kota
    data_by_location: Dict[str, AggregatedNIBData] = field(default_factory=dict)
    
    # Summary totals
    total_nib: int = 0
    total_pma: int = 0
    total_pmdn: int = 0
    total_umk: int = 0
    total_non_umk: int = 0
    
    # Monthly breakdown
    monthly_totals: Dict[str, int] = field(default_factory=dict)
    
    # Comparison with previous period (if available)
    prev_period_total: Optional[int] = None
    change_percentage: Optional[float] = None


class DataAggregator:
    """
    Aggregates NIB data from multiple monthly files into period reports.
    
    Supports:
    - Triwulan (Quarterly) aggregation
    - Semester aggregation
    - Tahunan (Annual) aggregation
    - Quarter-over-Quarter (Q-o-Q) comparison
    """
    
    TRIWULAN_MONTHS = {
        "TW I": ["Januari", "Februari", "Maret"],
        "TW II": ["April", "Mei", "Juni"],
        "TW III": ["Juli", "Agustus", "September"],
        "TW IV": ["Oktober", "November", "Desember"],
    }
    
    SEMESTER_MONTHS = {
        "Semester I": ["Januari", "Februari", "Maret", "April", "Mei", "Juni"],
        "Semester II": ["Juli", "Agustus", "September", "Oktober", "November", "Desember"],
    }
    
    def __init__(self):
        self.loader = DataLoader()
        self.loaded_data: Dict[str, Dict] = {}  # month -> data
    
    def load_files(self, file_inputs: List[Any]) -> None:
        """
        Load multiple Excel files and store their data.
        
        Args:
            file_inputs: List of file paths or Streamlit UploadedFile objects
        """
        for file_input in file_inputs:
            try:
                # Handle Streamlit UploadedFile
                if hasattr(file_input, 'getvalue') and hasattr(file_input, 'name'):
                    data = self.loader.load_monthly_data(file_input.getvalue(), filename=file_input.name)
                    filename_for_log = file_input.name
                # Handle Path or str
                else:
                    data = self.loader.load_monthly_data(file_input)
                    filename_for_log = str(file_input)
                
                month = data.get('month')
                year = data.get('year')
                if month and year:
                    key = f"{month}_{year}"
                    self.loaded_data[key] = data
                    print(f"Loaded: {filename_for_log} -> {key}")
                elif month: # Fallback if year missing
                     self.loaded_data[month] = data # Warning: unsafe
            except Exception as e:
                print(f"Error loading {file_input}: {e}")
    
    def aggregate_triwulan(self, triwulan: str, year: int) -> PeriodReport:
        """
        Aggregate data for a specific Triwulan (Quarter).
        
        Args:
            triwulan: Quarter name (e.g., "TW I", "TW II")
            year: Year
            
        Returns:
            PeriodReport for the quarter
        """
        months = self.TRIWULAN_MONTHS.get(triwulan, [])
        return self._aggregate_period(
            period_type="Triwulan",
            period_name=triwulan,
            year=year,
            months=months
        )
    
    def aggregate_semester(self, semester: str, year: int) -> PeriodReport:
        """
        Aggregate data for a specific Semester.
        
        Args:
            semester: Semester name (e.g., "Semester I", "Semester II")
            year: Year
            
        Returns:
            PeriodReport for the semester
        """
        months = self.SEMESTER_MONTHS.get(semester, [])
        return self._aggregate_period(
            period_type="Semester",
            period_name=semester,
            year=year,
            months=months
        )
    
    def aggregate_tahunan(self, year: int) -> PeriodReport:
        """
        Aggregate data for a full year.
        
        Args:
            year: Year
            
        Returns:
            PeriodReport for the year
        """
        all_months = [
            "Januari", "Februari", "Maret", "April", "Mei", "Juni",
            "Juli", "Agustus", "September", "Oktober", "November", "Desember"
        ]
        return self._aggregate_period(
            period_type="Tahunan",
            period_name=str(year),
            year=year,
            months=all_months
        )
    
    def _aggregate_period(
        self, 
        period_type: str, 
        period_name: str, 
        year: int, 
        months: List[str]
    ) -> PeriodReport:
        """
        Internal method to aggregate data for any period.
        
        Args:
            period_type: Type of period
            period_name: Name of period
            year: Year
            months: List of month names to include
            
        Returns:
            PeriodReport with aggregated data
        """
        report = PeriodReport(
            period_type=period_type,
            period_name=period_name,
            year=year,
            months_included=months
        )
        
        # Aggregate data by location
        location_data: Dict[str, AggregatedNIBData] = defaultdict(
            lambda: AggregatedNIBData(kabupaten_kota="")
        )
        
        for month in months:
            # Try specific year first
            key = f"{month}_{year}"
            month_data = self.loaded_data.get(key)
            
            # Fallback to old behavior (just month) if not found (backward compat)
            if not month_data:
                month_data = self.loaded_data.get(month)
            
            if not month_data:
                continue
            
            nib_list = month_data.get('nib', [])
            month_total = 0
            
            for nib in nib_list:
                kab_kota = nib.kabupaten_kota
                
                if kab_kota not in location_data:
                    location_data[kab_kota] = AggregatedNIBData(
                        kabupaten_kota=kab_kota
                    )
                
                agg = location_data[kab_kota]
                agg.period_data[month] = nib.total
                agg.pma_total += nib.pma
                agg.pmdn_total += nib.pmdn
                agg.usaha_mikro_total += nib.usaha_mikro
                agg.usaha_kecil_total += nib.usaha_kecil
                agg.usaha_menengah_total += nib.usaha_menengah
                agg.usaha_besar_total += nib.usaha_besar
                agg.grand_total += nib.total
                
                month_total += nib.total
            
            report.monthly_totals[month] = month_total
        
        # Store aggregated data
        report.data_by_location = dict(location_data)
        
        # Calculate summary totals
        for agg in location_data.values():
            report.total_nib += agg.grand_total
            report.total_pma += agg.pma_total
            report.total_pmdn += agg.pmdn_total
            report.total_umk += agg.umk_total
            report.total_non_umk += agg.non_umk_total
        
        return report
    
    def get_qoq_comparison(
        self, 
        current_triwulan: str, 
        year: int
    ) -> Tuple[PeriodReport, Optional[PeriodReport], Optional[float]]:
        """
        Get Quarter-over-Quarter comparison.
        
        Args:
            current_triwulan: Current quarter (e.g., "TW II")
            year: Year
            
        Returns:
            Tuple of (current_report, previous_report, change_percentage)
        """
        triwulan_order = ["TW I", "TW II", "TW III", "TW IV"]
        
        current_report = self.aggregate_triwulan(current_triwulan, year)
        
        # Determine previous quarter
        current_idx = triwulan_order.index(current_triwulan)
        if current_idx == 0:
            # TW I -> previous is TW IV of last year
            prev_triwulan = "TW IV"
            prev_year = year - 1
        else:
            prev_triwulan = triwulan_order[current_idx - 1]
            prev_year = year
        
        # Try to get previous quarter data
        try:
            prev_report = self.aggregate_triwulan(prev_triwulan, prev_year)
            
            if prev_report.total_nib > 0:
                change_pct = (
                    (current_report.total_nib - prev_report.total_nib) 
                    / prev_report.total_nib * 100
                )
                current_report.prev_period_total = prev_report.total_nib
                current_report.change_percentage = change_pct
                return current_report, prev_report, change_pct
        except Exception:
            pass
        
        return current_report, None, None
    
    def to_dataframe(self, report: PeriodReport) -> pd.DataFrame:
        """
        Convert PeriodReport to a pandas DataFrame.
        
        Args:
            report: PeriodReport to convert
            
        Returns:
            DataFrame with aggregated data
        """
        data = []
        for kab_kota, agg in report.data_by_location.items():
            row = {
                'Kabupaten/Kota': kab_kota,
                'PMA': agg.pma_total,
                'PMDN': agg.pmdn_total,
                'Usaha Mikro': agg.usaha_mikro_total,
                'Usaha Kecil': agg.usaha_kecil_total,
                'Usaha Menengah': agg.usaha_menengah_total,
                'Usaha Besar': agg.usaha_besar_total,
                'UMK': agg.umk_total,
                'NON-UMK': agg.non_umk_total,
                'Total': agg.grand_total,
            }
            
            # Add monthly breakdown
            for month in report.months_included:
                row[month] = agg.period_data.get(month, 0)
            
            data.append(row)
        
        df = pd.DataFrame(data)
        
        # Sort by total descending
        if not df.empty and 'Total' in df.columns:
            df = df.sort_values('Total', ascending=False).reset_index(drop=True)
        
        return df
    
    def get_summary_stats(self, report: PeriodReport) -> Dict:
        """
        Get summary statistics for a report.
        
        Args:
            report: PeriodReport to analyze
            
        Returns:
            Dictionary with summary statistics
        """
        df = self.to_dataframe(report)
        
        if df.empty:
            return {}
        
        # Top 5 locations by total
        top_5 = df.nlargest(5, 'Total')[['Kabupaten/Kota', 'Total']].to_dict('records')
        
        # PM distribution - use PMA+PMDN as base for percentage calculation
        pm_total = report.total_pma + report.total_pmdn
        pm_dist = {
            'PMA': report.total_pma,
            'PMDN': report.total_pmdn,
            'PMA_pct': (report.total_pma / pm_total * 100) if pm_total > 0 else 0,
            'PMDN_pct': (report.total_pmdn / pm_total * 100) if pm_total > 0 else 0,
        }
        
        # Pelaku Usaha distribution - use UMK+NON_UMK as base
        pelaku_total = report.total_umk + report.total_non_umk
        pelaku_dist = {
            'UMK': report.total_umk,
            'NON_UMK': report.total_non_umk,
            'UMK_pct': (report.total_umk / pelaku_total * 100) if pelaku_total > 0 else 0,
            'NON_UMK_pct': (report.total_non_umk / pelaku_total * 100) if pelaku_total > 0 else 0,
        }
        
        return {
            'total_nib': report.total_nib,
            'top_5_locations': top_5,
            'pm_distribution': pm_dist,
            'pelaku_usaha_distribution': pelaku_dist,
            'monthly_totals': report.monthly_totals,
            'change_percentage': report.change_percentage,
            'prev_period_total': report.prev_period_total,
        }
