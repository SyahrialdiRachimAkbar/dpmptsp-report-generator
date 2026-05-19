"""
Reference Data Loader Module for DPMPTSP Reporting System

This module handles loading and parsing reference Excel files:
1. NIB file - Business registration data
2. PB OSS file - Permit/Risk/Sector data  
3. PROYEK file - Project/Investment data

Key Business Rules:
- NIB deduplication is per-month (same NIB in different months = counted separately)
- PROYEK sums all investment amounts (no deduplication)
- PB OSS counts all permits (not unique NIB)
"""

import pandas as pd
import re
import math
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any
from dataclasses import dataclass, field
from datetime import datetime

from app.config import NAMA_BULAN, TRIWULAN_KE_BULAN


@dataclass
class NIBReferenceData:
    """Data structure for NIB reference file data."""
    year: int
    total_nib: int = 0  # Sum of monthly unique NIB counts
    monthly_totals: Dict[str, int] = field(default_factory=dict)  # Month → unique NIB count
    by_kab_kota: Dict[str, Dict[str, int]] = field(default_factory=dict)  # Kab → Month → count
    by_pm_status: Dict[str, Dict[str, int]] = field(default_factory=dict)  # PM → Month → count
    by_skala_usaha: Dict[str, Dict[str, int]] = field(default_factory=dict)  # Skala → Month → count
    # Detailed breakdowns for cross-tabulation (Kab/Kota → Month → Category → Count)
    kab_pm_monthly: Dict[str, Dict[str, Dict[str, int]]] = field(default_factory=dict)
    kab_skala_monthly: Dict[str, Dict[str, Dict[str, int]]] = field(default_factory=dict)
    
    def get_period_total(self, months: List[str]) -> int:
        """Get total NIB for specified months."""
        return sum(self.monthly_totals.get(m, 0) for m in months)
    
    def get_period_by_kab_kota(self, months: List[str]) -> Dict[str, int]:
        """Get Kab/Kota totals for specified months."""
        result = {}
        for kab, month_data in self.by_kab_kota.items():
            result[kab] = sum(month_data.get(m, 0) for m in months)
        return result
    
    def get_period_by_pm_status(self, months: List[str]) -> Dict[str, int]:
        """Get PM status totals for specified months."""
        result = {}
        for pm, month_data in self.by_pm_status.items():
            result[pm] = sum(month_data.get(m, 0) for m in months)
        return result
    
    def get_period_by_skala_usaha(self, months: List[str]) -> Dict[str, int]:
        """Get skala usaha totals for specified months."""
        result = {}
        for skala, month_data in self.by_skala_usaha.items():
            result[skala] = sum(month_data.get(m, 0) for m in months)
        return result


@dataclass
class PBOSSReferenceData:
    """Data structure for PB OSS reference file data."""
    year: int
    monthly_risk: Dict[str, Dict[str, int]] = field(default_factory=dict)  # Month → Risk → count
    monthly_sector: Dict[str, Dict[str, int]] = field(default_factory=dict)  # Month → Sector → count
    monthly_by_kab_kota: Dict[str, Dict[str, int]] = field(default_factory=dict)  # Month → Kab/Kota → count
    monthly_status_pm: Dict[str, Dict[str, int]] = field(default_factory=dict)  # Month → Status PM → count
    monthly_jenis_perizinan: Dict[str, Dict[str, int]] = field(default_factory=dict)  # Month → Jenis → count
    monthly_status_perizinan: Dict[str, Dict[str, int]] = field(default_factory=dict)  # Month → Status → count
    monthly_kewenangan: Dict[str, Dict[str, int]] = field(default_factory=dict)  # Month → Kewenangan → count
    monthly_permits: Dict[str, int] = field(default_factory=dict)  # Month → permit count
    total_permits: int = 0
    
    def get_period_risk(self, months: List[str]) -> Dict[str, int]:
        """Get risk distribution for specified months."""
        result = {}
        for month in months:
            if month in self.monthly_risk:
                for risk, count in self.monthly_risk[month].items():
                    result[risk] = result.get(risk, 0) + count
        return result
    
    def get_period_sector(self, months: List[str]) -> Dict[str, int]:
        """Get sector distribution for specified months."""
        result = {}
        for month in months:
            if month in self.monthly_sector:
                for sector, count in self.monthly_sector[month].items():
                    result[sector] = result.get(sector, 0) + count
        return result
    
    def get_period_by_kab_kota(self, months: List[str]) -> Dict[str, int]:
        """Get permits by Kab/Kota for specified months."""
        result = {}
        for month in months:
            if month in self.monthly_by_kab_kota:
                for kab, count in self.monthly_by_kab_kota[month].items():
                    result[kab] = result.get(kab, 0) + count
        return result
    
    def get_period_status_pm(self, months: List[str]) -> Dict[str, int]:
        """Get Status PM distribution for specified months."""
        result = {}
        for month in months:
            if month in self.monthly_status_pm:
                for status, count in self.monthly_status_pm[month].items():
                    result[status] = result.get(status, 0) + count
        return result
    
    def get_period_jenis_perizinan(self, months: List[str]) -> Dict[str, int]:
        """Get Jenis Perizinan distribution for specified months."""
        result = {}
        for month in months:
            if month in self.monthly_jenis_perizinan:
                for jenis, count in self.monthly_jenis_perizinan[month].items():
                    result[jenis] = result.get(jenis, 0) + count
        return result
    
    def get_period_status_perizinan(self, months: List[str]) -> Dict[str, int]:
        """Get Status Perizinan distribution for specified months."""
        result = {}
        for month in months:
            if month in self.monthly_status_perizinan:
                for status, count in self.monthly_status_perizinan[month].items():
                    result[status] = result.get(status, 0) + count
        return result
    
    def get_period_kewenangan(self, months: List[str]) -> Dict[str, int]:
        """Get Kewenangan distribution for specified months."""
        result = {}
        for month in months:
            if month in self.monthly_kewenangan:
                for kew, count in self.monthly_kewenangan[month].items():
                    result[kew] = result.get(kew, 0) + count
        return result
    
    def get_period_permits(self, months: List[str]) -> int:
        """Get total permits for specified months."""
        return sum(self.monthly_permits.get(m, 0) for m in months)

    def get_monthly_status_pm_breakdown(self, months: List[str]) -> Dict[str, Dict[str, int]]:
        """
        Get breakdown of Status PM (PMA/PMDN) for each month.
        Returns: { 'Januari': {'PMA': 10, 'PMDN': 20}, ... }
        """
        result = {}
        for month in months:
            if month in self.monthly_status_pm:
                result[month] = {
                    'PMA': self.monthly_status_pm[month].get('PMA', 0),
                    'PMDN': self.monthly_status_pm[month].get('PMDN', 0)
                }
            else:
                result[month] = {'PMA': 0, 'PMDN': 0}
        return result

    def get_period_permits_by_month(self, months: List[str]) -> Dict[str, int]:
        """Get permit counts by month for specified months."""
        result = {}
        for month in months:
            if month in self.monthly_permits:
                result[month] = self.monthly_permits[month]
        return result


@dataclass
class ProyekReferenceData:
    """Data structure for PROYEK reference file data."""
    year: int
    # Monthly investment data (sum all, no dedup)
    monthly_investment: Dict[str, float] = field(default_factory=dict)  # Month → total investment
    monthly_pma: Dict[str, float] = field(default_factory=dict)  # Month → PMA investment
    monthly_pmdn: Dict[str, float] = field(default_factory=dict)  # Month → PMDN investment
    monthly_tki: Dict[str, int] = field(default_factory=dict)  # Month → TKI count
    monthly_tka: Dict[str, int] = field(default_factory=dict)  # Month → TKA count
    monthly_projects: Dict[str, int] = field(default_factory=dict)  # Month → project count
    monthly_pma_projects: Dict[str, int] = field(default_factory=dict)  # Month → PMA project count
    monthly_pmdn_projects: Dict[str, int] = field(default_factory=dict)  # Month → PMDN project count
    monthly_by_wilayah: Dict[str, Dict[str, float]] = field(default_factory=dict)  # Month → Wilayah → investment
    monthly_by_skala_usaha: Dict[str, Dict[str, int]] = field(default_factory=dict)  # Month → Skala → project count
    monthly_labor_by_wilayah: Dict[str, Dict[str, int]] = field(default_factory=dict)  # Month → Wilayah → labor count (TKI+TKA)
    monthly_projects_by_wilayah: Dict[str, Dict[str, int]] = field(default_factory=dict)  # Month → Wilayah → project count
    
    def get_period_investment(self, months: List[str]) -> float:
        """Get total investment for specified months."""
        return sum(self.monthly_investment.get(m, 0) for m in months)
    
    def get_period_pma(self, months: List[str]) -> float:
        """Get PMA investment for specified months."""
        return sum(self.monthly_pma.get(m, 0) for m in months)
    
    def get_period_pmdn(self, months: List[str]) -> float:
        """Get PMDN investment for specified months."""
        return sum(self.monthly_pmdn.get(m, 0) for m in months)
    
    def get_period_tki(self, months: List[str]) -> int:
        """Get TKI count for specified months."""
        return sum(self.monthly_tki.get(m, 0) for m in months)
    
    def get_period_tka(self, months: List[str]) -> int:
        """Get TKA count for specified months."""
        return sum(self.monthly_tka.get(m, 0) for m in months)
    
    def get_period_projects(self, months: List[str]) -> int:
        """Get project count for specified months."""
        return sum(self.monthly_projects.get(m, 0) for m in months)
    
    def get_period_pma_projects(self, months: List[str]) -> int:
        """Get PMA project count for specified months."""
        return sum(self.monthly_pma_projects.get(m, 0) for m in months)
    
    def get_period_pmdn_projects(self, months: List[str]) -> int:
        """Get PMDN project count for specified months."""
        return sum(self.monthly_pmdn_projects.get(m, 0) for m in months)
    
    def get_period_by_wilayah(self, months: List[str]) -> Dict[str, float]:
        """Get investment by wilayah for specified months."""
        result = {}
        for month in months:
            if month in self.monthly_by_wilayah:
                for wilayah, investment in self.monthly_by_wilayah[month].items():
                    result[wilayah] = result.get(wilayah, 0) + investment
        return result
    
    def get_period_by_skala_usaha(self, months: List[str]) -> Dict[str, int]:
        """Get project count by skala usaha for specified months."""
        result = {}
        for month in months:
            if month in self.monthly_by_skala_usaha:
                for skala, count in self.monthly_by_skala_usaha[month].items():
                    result[skala] = result.get(skala, 0) + count
        return result
    
    def get_period_labor_by_wilayah(self, months: List[str]) -> Dict[str, int]:
        """Get total labor (TKI+TKA) by wilayah for specified months."""
        result = {}
        for month in months:
            if month in self.monthly_labor_by_wilayah:
                for wilayah, count in self.monthly_labor_by_wilayah[month].items():
                    result[wilayah] = result.get(wilayah, 0) + count
        return result
    
    def get_period_projects_by_wilayah(self, months: List[str]) -> Dict[str, int]:
        """Get project count by wilayah for specified months."""
        result = {}
        for month in months:
            if month in self.monthly_projects_by_wilayah:
                for wilayah, count in self.monthly_projects_by_wilayah[month].items():
                    result[wilayah] = result.get(wilayah, 0) + count
        return result


class ReferenceDataLoader:
    """
    Loader for reference Excel files (NIB, PB OSS, PROYEK).
    
    These files contain year-to-date data in a different format
    from the monthly operational files.
    """
    
    # Month name mapping for date parsing
    MONTH_MAP = {
        1: "Januari", 2: "Februari", 3: "Maret", 4: "April",
        5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus",
        9: "September", 10: "Oktober", 11: "November", 12: "Desember"
    }
    INDO_MONTHS = {
        'Januari': 1, 'Februari': 2, 'Maret': 3, 'April': 4,
        'Mei': 5, 'Juni': 6, 'Juli': 7, 'Agustus': 8,
        'September': 9, 'Oktober': 10, 'November': 11, 'Desember': 12
    }
    
    # Risk level mapping
    RISK_MAP = {
        'R': 'Rendah',
        'MR': 'Menengah Rendah', 
        'MT': 'Menengah Tinggi',
        'T': 'Tinggi'
    }
    
    def __init__(self):
        self.data_cache = {}
    
    def detect_file_type(self, file_bytes: BytesIO, filename: str) -> Optional[str]:
        """
        Detect file type from sheet names and structure.
        
        Returns: 'NIB', 'PB_OSS', 'PROYEK', or None
        """
        try:
            xl = pd.ExcelFile(file_bytes)
            sheet_names = [s.upper() for s in xl.sheet_names]
            
            # Check for PROYEK indicators
            if len(xl.sheet_names) == 1:
                df = pd.read_excel(xl, sheet_name=xl.sheet_names[0], nrows=5)
                cols = [str(c).upper() for c in df.columns]
                if any('PROYEK' in c or 'INVESTASI' in c for c in cols):
                    return 'PROYEK'
                if any('JUMLAH INVESTASI' in c for c in cols):
                    return 'PROYEK'
            
            # Check for PB OSS indicators (has RISIKO or SEKTOR sheets)
            if any('RISIKO' in s for s in sheet_names) or any('SEKTOR' in s for s in sheet_names):
                return 'PB_OSS'
            
            # Check for NIB indicators (has SKALA USAHA or PM sheets)
            if any('SKALA' in s for s in sheet_names) or any('JENIS PERUSAHAAN' in s for s in sheet_names):
                return 'NIB'
            
            # Fallback: check filename
            filename_upper = filename.upper()
            if 'NIB' in filename_upper:
                return 'NIB'
            if 'PB' in filename_upper or 'PERIZINAN' in filename_upper:
                return 'PB_OSS'
            if 'PROYEK' in filename_upper:
                return 'PROYEK'
            
            return None
            
        except Exception as e:
            print(f"Error detecting file type: {e}")
            return None
    
    def extract_year_from_filename(self, filename: str) -> Optional[int]:
        """Extract year from filename."""
        match = re.search(r'(20\d{2})', filename)
        if match:
            return int(match.group(1))
        return None
    
    def _parse_date_to_month(self, date_val) -> Optional[str]:
        """Parse a date value to month name."""
        if pd.isna(date_val):
            return None
        
        try:
            # If standard datetime object
            if isinstance(date_val, datetime):
                return self.MONTH_MAP.get(date_val.month)
            
            # If string
            if isinstance(date_val, str):
                date_str = date_val.strip()
                
                # Check for "Day Month Year" format (e.g., "18 November 2025")
                # This format often appears in the Excel files
                try:
                    # Try English parsing first
                    dt = datetime.strptime(date_str, '%d %B %Y')
                    return self.MONTH_MAP.get(dt.month)
                except ValueError:
                    pass
                
                # Try handling Indonesian month names manually
                # Map Indo month -> English month for parsing, or direct to index
                for indo, idx in self.INDO_MONTHS.items():
                    if indo in date_str:
                        return self.MONTH_MAP.get(idx)
                
                # Try standard numeric formats
                for fmt in ['%m/%d/%Y', '%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y']:
                    try:
                        dt = datetime.strptime(date_str, fmt)
                        return self.MONTH_MAP.get(dt.month)
                    except ValueError:
                        continue
                        
            return None
        except Exception:
            return None

    def _parse_date_obj(self, date_val) -> Optional[datetime]:
        """Parse a date value to datetime object."""
        if pd.isna(date_val):
            return None
        
        try:
            # If standard datetime object
            if isinstance(date_val, (datetime, pd.Timestamp)):
                return date_val

            if isinstance(date_val, (int, float)) and not math.isnan(float(date_val)):
                parsed = pd.to_datetime(date_val, unit='D', origin='1899-12-30', errors='coerce')
                if not pd.isna(parsed):
                    return parsed.to_pydatetime()
            
            # If string
            if isinstance(date_val, str):
                date_str = date_val.strip()
                
                # Check for "Day Month Year" format (e.g., "18 November 2025")
                try:
                    # Try English parsing first
                    return datetime.strptime(date_str, '%d %B %Y')
                except ValueError:
                    pass
                
                # Check for Indonesian month names manually
                # Check if string contains any Indo month
                for indo, idx in self.INDO_MONTHS.items():
                    if indo.lower() in date_str.lower():
                        # Try to replace indo month with english or parse manually
                        # Simple regex to extract Day and Year if format is "DD Month YYYY"
                        import re
                        match = re.search(r'(\d{1,2})\s+([a-zA-Z]+)\s+(\d{4})', date_str)
                        if match:
                            d, m, y = match.groups()
                            if m.lower() == indo.lower():
                                return datetime(int(y), idx, int(d))
                
                # Try standard numeric formats
                for fmt in ['%m/%d/%Y', '%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y']:
                    try:
                        return datetime.strptime(date_str, fmt)
                    except ValueError:
                        continue
                        
            return None
        except Exception:
            return None

    def _parse_date_series(self, series: pd.Series) -> pd.Series:
        """Parse a date Series with vectorized pandas parsing plus Indonesian fallback."""
        if pd.api.types.is_datetime64_any_dtype(series):
            return pd.to_datetime(series, errors='coerce')

        parsed = pd.Series(pd.NaT, index=series.index, dtype='datetime64[ns]')
        numeric_values = pd.to_numeric(series, errors='coerce')
        numeric_mask = series.map(lambda value: isinstance(value, (int, float)) and not isinstance(value, bool)) & numeric_values.notna()
        if numeric_mask.any():
            parsed.loc[numeric_mask] = pd.to_datetime(
                numeric_values.loc[numeric_mask],
                unit='D',
                origin='1899-12-30',
                errors='coerce'
            )

        remaining = parsed.isna() & series.notna()
        if remaining.any():
            parsed.loc[remaining] = pd.to_datetime(series.loc[remaining], errors='coerce', dayfirst=False)

        missing = parsed.isna() & series.notna()
        if missing.any():
            parsed.loc[missing] = series.loc[missing].apply(self._parse_date_obj)
        return parsed

    def _month_series(self, series: pd.Series) -> pd.Series:
        """Convert a date Series to canonical Indonesian month names."""
        parsed = self._parse_date_series(series)
        return parsed.dt.month.map(self.MONTH_MAP)

    def _assign_count_by_month(self, target: Dict[str, int], counts: pd.Series) -> None:
        for month, count in counts.items():
            if pd.notna(month):
                target[month] = int(count)

    def _assign_nested_month_counts(self, target: Dict[str, Dict[str, int]], counts: pd.Series, key_transform=None) -> None:
        for key, month, count in counts.reset_index().itertuples(index=False, name=None):
            if pd.isna(key) or pd.isna(month):
                continue
            key = key_transform(key) if key_transform else key
            target.setdefault(key, {})[month] = int(count)

    def _assign_month_nested_counts(self, target: Dict[str, Dict[str, int]], counts: pd.Series, value_transform=None, limit_per_month: Optional[int] = None) -> None:
        frame = counts.reset_index(name='count')
        if limit_per_month is not None and not frame.empty:
            frame = frame.sort_values(['_month', 'count'], ascending=[True, False]).groupby('_month').head(limit_per_month)
        for month, value, count in frame.itertuples(index=False, name=None):
            if pd.isna(month) or pd.isna(value):
                continue
            value = value_transform(value) if value_transform else value
            target.setdefault(month, {})[value] = int(count)

    def _assign_three_level_counts(self, target: Dict[str, Dict[str, Dict[str, int]]], counts: pd.Series, third_transform=None) -> None:
        for first, month, third, count in counts.reset_index().itertuples(index=False, name=None):
            if pd.isna(first) or pd.isna(month) or pd.isna(third):
                continue
            third = third_transform(third) if third_transform else third
            target.setdefault(first, {}).setdefault(month, {})[third] = int(count)

    def load_nib(self, file_bytes: BytesIO, filename: str, year: Optional[int] = None) -> Optional[NIBReferenceData]:
        """
        Load NIB reference file.
        
        Uses Sheet 1 (raw data) with columns:
        - nib: Business registration number
        - Day of tanggal_terbit_oss: Issue date
        - kab_kota: Location
        - status_penanaman_modal: PMA/PMDN
        - uraian_skala_usaha: Business scale
        """
        try:
            xl = pd.ExcelFile(file_bytes)
            
            # Find Sheet 1 or similar raw data sheet
            # Priority: 'Sheet 1' (with space) > 'sheet1' (no space) > last sheet
            raw_sheet = None
            for sheet in xl.sheet_names:
                if sheet.lower() == 'sheet 1':  # Exact match with space first
                    raw_sheet = sheet
                    break
            
            if not raw_sheet:
                for sheet in xl.sheet_names:
                    if 'sheet 1' in sheet.lower() or sheet.lower() == 'sheet1':
                        raw_sheet = sheet
                        break
            
            if not raw_sheet:
                raw_sheet = xl.sheet_names[-1]  # Fallback to last sheet
            
            df = pd.read_excel(xl, sheet_name=raw_sheet)
            
            # Normalize column names
            df.columns = [str(c).strip().lower() for c in df.columns]
            
            # Find relevant columns
            nib_col = self._find_column(df, ['nib'])
            date_col = self._find_column(df, ['tanggal_terbit', 'tanggal', 'day of tanggal'])
            kab_col = self._find_column(df, ['kab_kota', 'kab kota', 'kabupaten'])
            pm_col = self._find_column(df, ['status_penanaman_modal', 'status pm', 'penanaman_modal'])
            skala_col = self._find_column(df, ['uraian_skala_usaha', 'skala_usaha', 'skala usaha'])
            
            if not nib_col:
                print(f"NIB column not found in {filename}")
                return None
            
            # Extract year
            if year is None:
                year = self.extract_year_from_filename(filename) or datetime.now().year
            
            if date_col:
                df['_month'] = self._month_series(df[date_col])
                df = df[df['_month'].isin(NAMA_BULAN)].copy()
            else:
                df['_month'] = 'Januari'

            result = NIBReferenceData(year=year)

            monthly_counts = df.groupby('_month')[nib_col].nunique()
            self._assign_count_by_month(result.monthly_totals, monthly_counts)

            if kab_col:
                kab_counts = df.groupby([kab_col, '_month'])[nib_col].nunique()
                self._assign_nested_month_counts(result.by_kab_kota, kab_counts)

            if pm_col:
                df['_pm_status'] = df[pm_col].astype(str).str.upper().str.strip()
                pm_counts = df.groupby(['_pm_status', '_month'])[nib_col].nunique()
                self._assign_nested_month_counts(result.by_pm_status, pm_counts)

            if skala_col:
                skala_counts = df.groupby([skala_col, '_month'])[nib_col].nunique()
                self._assign_nested_month_counts(result.by_skala_usaha, skala_counts)

            if kab_col and pm_col:
                kab_pm = df.groupby([kab_col, '_month', '_pm_status'])[nib_col].nunique()
                self._assign_three_level_counts(result.kab_pm_monthly, kab_pm)

            if kab_col and skala_col:
                kab_skala = df.groupby([kab_col, '_month', skala_col])[nib_col].nunique()
                self._assign_three_level_counts(result.kab_skala_monthly, kab_skala)
            
            # Calculate total (sum of monthly counts)
            result.total_nib = sum(result.monthly_totals.values())
            
            return result
            
        except Exception as e:
            print(f"Error loading NIB file: {e}")
            return None
    
    def load_pb_oss(self, file_bytes: BytesIO, filename: str, year: Optional[int] = None) -> Optional[PBOSSReferenceData]:
        """
        Load PB OSS reference file.
        
        Uses Sheet 1 (raw data) to count all permits by risk/sector.
        """
        try:
            xl = pd.ExcelFile(file_bytes)
            
            # Smart sheet detection
            # Find all sheets that look like raw data
            candidate_sheets = []
            for sheet in xl.sheet_names:
                s_lower = sheet.lower()
                if 'sheet' in s_lower or 'data' in s_lower:
                    candidate_sheets.append(sheet)
            
            # If no apparent candidates, check all
            if not candidate_sheets:
                candidate_sheets = xl.sheet_names
            
            best_sheet = None
            max_rows = 0
            best_df = None
            
            # Evaluate candidates
            for sheet in candidate_sheets:
                try:
                    df = pd.read_excel(xl, sheet_name=sheet)
                    df.columns = [str(c).strip().lower() for c in df.columns]
                    
                    # Check for critical columns
                    has_risk = self._find_column(df, ['risiko', 'risk', 'kd_resiko', 'uraian risiko'])
                    has_nib = self._find_column(df, ['nib'])
                    
                    if has_risk or has_nib:
                        if len(df) > max_rows:
                            max_rows = len(df)
                            best_sheet = sheet
                            best_df = df
                except Exception:
                    continue
            
            if best_df is None:
                print(f"No valid data sheet found in {filename}")
                return None
            
            df = best_df
            
            # Find relevant columns
            date_col = self._find_column(df, ['day of tgl_izin', 'tanggal', 'tgl', 'date'])
            risk_col = self._find_column(df, ['risiko', 'risk', 'kd_resiko', 'uraian risiko'])
            sector_col = self._find_column(df, ['sektor', 'sector', 'judul_kbli', 'uraian_sektor'])
            
            if year is None:
                year = self.extract_year_from_filename(filename) or datetime.now().year
            
            if date_col:
                df['_month'] = self._month_series(df[date_col])
            else:
                # If no date, put all in first month
                df['_month'] = 'Januari'
            df = df[df['_month'].isin(NAMA_BULAN)].copy()
            
            result = PBOSSReferenceData(year=year)
            
            # Find additional columns for new breakdowns
            kab_kota_col = self._find_column(df, ['kab_kota', 'kab kota', 'kabupaten', 'kota'])
            status_pm_col = self._find_column(df, ['status pm', 'status_pm'])
            jenis_perizinan_col = self._find_column(df, ['uraian_jenis_perizinan', 'jenis_perizinan', 'jenis perizinan'])
            status_perizinan_col = self._find_column(df, ['status perizinan', 'status_perizinan'])
            
            # Find kewenangan columns - strict detection first
            uraian_kewenangan_col = self._find_exact_column(df, 'uraian_kewenangan')
            kewenangan_col = self._find_exact_column(df, 'kewenangan')
            
            # Fallback to fuzzy search if exact match not found
            if not kewenangan_col or not uraian_kewenangan_col:
                for col in df.columns:
                    col_lower = str(col).lower().strip()
                    if not uraian_kewenangan_col and col_lower == 'uraian_kewenangan':
                        uraian_kewenangan_col = col
                    elif not kewenangan_col and 'kewenangan' in col_lower and 'uraian' not in col_lower:
                        kewenangan_col = col
            

            
            if kewenangan_col:
                kew_counts = df.groupby(['_month', kewenangan_col]).size()
                self._assign_month_nested_counts(result.monthly_kewenangan, kew_counts)

            if status_perizinan_col:
                status_counts = df.groupby(['_month', status_perizinan_col]).size()
                self._assign_month_nested_counts(result.monthly_status_perizinan, status_counts)
            
            # Now apply Gubernur filter for remaining breakdowns
            if uraian_kewenangan_col:
                gubernur_mask = df[uraian_kewenangan_col].astype(str).str.upper().str.contains('GUBERNUR', na=False)
                df = df[gubernur_mask].copy()
            elif kewenangan_col:
                gubernur_mask = df[kewenangan_col].astype(str).str.upper().str.contains('GUBERNUR', na=False)
                df = df[gubernur_mask].copy()

            self._assign_count_by_month(result.monthly_permits, df.groupby('_month').size())

            if risk_col:
                risk_counts = df.groupby(['_month', risk_col]).size()
                def risk_name(value):
                    risk_str = str(value).strip().upper()
                    return self.RISK_MAP.get(risk_str, risk_str)
                self._assign_month_nested_counts(result.monthly_risk, risk_counts, risk_name)

            if sector_col:
                sector_counts = df.groupby(['_month', sector_col]).size()
                self._assign_month_nested_counts(result.monthly_sector, sector_counts)

            if kab_kota_col:
                kab_counts = df.groupby(['_month', kab_kota_col]).size()
                self._assign_month_nested_counts(result.monthly_by_kab_kota, kab_counts)

            if status_pm_col:
                pm_counts = df.groupby(['_month', status_pm_col]).size()
                self._assign_month_nested_counts(result.monthly_status_pm, pm_counts)

            if jenis_perizinan_col:
                jenis_counts = df.groupby(['_month', jenis_perizinan_col]).size()
                self._assign_month_nested_counts(result.monthly_jenis_perizinan, jenis_counts, limit_per_month=15)
            
            # Calculate total permits (from Gubernur-filtered data)
            result.total_permits = len(df)
            
            return result
            
        except Exception as e:
            print(f"Error loading PB OSS file: {e}")
            return None
    
    def load_proyek(self, file_bytes: BytesIO, filename: str, year: Optional[int] = None) -> Optional[ProyekReferenceData]:
        """
        Load PROYEK reference file.
        
        Uses Sheet 1 with columns:
        - tanggal_pengajuan_proyek: Project submission date
        - Jumlah Investasi: Investment amount
        - Status PM: PMA/PMDN
        - Kab Kota Usaha: Location
        - TKI, TKA: Labor counts
        """
        try:
            xl = pd.ExcelFile(file_bytes)
            
            # Use first/only sheet
            raw_sheet = xl.sheet_names[0]
            df = pd.read_excel(xl, sheet_name=raw_sheet)
            df.columns = [str(c).strip().lower() for c in df.columns]
            
            # Find relevant columns
            date_col = self._find_column(df, ['tanggal_pengajuan_proyek', 'tanggal_pengajuan', 'tanggal pengajuan'])
            investment_col = self._find_column(df, ['jumlah investasi', 'jumlah_investasi', 'investasi'])
            pm_col = self._find_column(df, ['status pm', 'status_pm', 'penanaman modal'])
            wilayah_col = self._find_column(df, ['kab kota usaha', 'kab_kota_usaha', 'kab kota', 'kabupaten'])
            tki_col = self._find_column(df, ['tki'])
            tka_col = self._find_column(df, ['tka'])
            skala_col = self._find_column(df, ['uraian_skala_usaha', 'skala_usaha', 'skala usaha'])
            
            if year is None:
                year = self.extract_year_from_filename(filename) or datetime.now().year
            
            if date_col:
                df['_date_obj'] = self._parse_date_series(df[date_col])
                
                # Filter by year if specified
                if year:
                    df = df[df['_date_obj'].dt.year == year].copy()
                
                # Then extract month names from the filtered data
                df['_month'] = df['_date_obj'].dt.month.map(self.MONTH_MAP)
            else:
                df['_month'] = 'Januari'
            df = df[df['_month'].isin(NAMA_BULAN)].copy()
            
            result = ProyekReferenceData(year=year)
            
            # Filter by Kewenangan = Gubernur
            kewenangan_col = self._find_column(df, ['kewenangan'])
            if kewenangan_col:
                gubernur_mask = df[kewenangan_col].astype(str).str.upper().str.contains('GUBERNUR', na=False)
                df = df[gubernur_mask].copy()
            
            if investment_col:
                df[investment_col] = pd.to_numeric(df[investment_col], errors='coerce').fillna(0)
            if tki_col:
                df[tki_col] = pd.to_numeric(df[tki_col], errors='coerce').fillna(0)
            if tka_col:
                df[tka_col] = pd.to_numeric(df[tka_col], errors='coerce').fillna(0)
            
            # Try to find a Project ID column for deduplication (to fix inflated labor counts)
            id_col = self._find_column(df, ['id_proyek', 'id proyek', 'nomor_proyek', 'kode_proyek', 'nib'])

            if investment_col:
                for month, value in df.groupby('_month')[investment_col].sum().items():
                    result.monthly_investment[month] = float(value)

                if wilayah_col:
                    wilayah_sums = df.groupby(['_month', wilayah_col])[investment_col].sum().reset_index(name='investment')
                    for month, wilayah, investment in wilayah_sums.itertuples(index=False, name=None):
                        result.monthly_by_wilayah.setdefault(month, {})[wilayah] = float(investment)

            if pm_col:
                pm_upper = df[pm_col].astype(str).str.upper()
                pma_invest_mask = pm_upper.str.contains('PMA', na=False)
                pmdn_mask = pm_upper.str.contains('PMDN', na=False)
                pma_project_mask = pma_invest_mask & ~pmdn_mask

                if investment_col:
                    for month, value in df.loc[pma_invest_mask].groupby('_month')[investment_col].sum().items():
                        result.monthly_pma[month] = float(value)
                    for month, value in df.loc[pmdn_mask].groupby('_month')[investment_col].sum().items():
                        result.monthly_pmdn[month] = float(value)

                self._assign_count_by_month(result.monthly_pma_projects, df.loc[pma_project_mask].groupby('_month').size())
                self._assign_count_by_month(result.monthly_pmdn_projects, df.loc[pmdn_mask].groupby('_month').size())

            if tki_col:
                self._assign_count_by_month(result.monthly_tki, df.groupby('_month')[tki_col].sum())
            if tka_col:
                self._assign_count_by_month(result.monthly_tka, df.groupby('_month')[tka_col].sum())

            self._assign_count_by_month(result.monthly_projects, df.groupby('_month').size())

            if skala_col:
                skala_counts = df.groupby(['_month', skala_col]).size()
                self._assign_month_nested_counts(result.monthly_by_skala_usaha, skala_counts)

            if wilayah_col and (tki_col or tka_col):
                if id_col:
                    dedup_cols = ['_month', wilayah_col, id_col]
                else:
                    dedup_cols = ['_month', wilayah_col]
                    name_col = self._find_column(df, ['nama_perusahaan', 'nama perusahaan', 'nama_perseroan', 'pelaku_usaha'])
                    if name_col:
                        dedup_cols.append(name_col)
                    if investment_col:
                        dedup_cols.append(investment_col)

                calc_df = df.drop_duplicates(subset=dedup_cols).copy()
                calc_df['_labor_total'] = 0
                if tki_col:
                    calc_df['_labor_total'] += calc_df[tki_col]
                if tka_col:
                    calc_df['_labor_total'] += calc_df[tka_col]

                labor_sums = calc_df.groupby(['_month', wilayah_col])['_labor_total'].sum().reset_index(name='labor')
                for month, wilayah, labor in labor_sums.itertuples(index=False, name=None):
                    result.monthly_labor_by_wilayah.setdefault(month, {})[wilayah] = int(labor)

            if wilayah_col:
                project_counts = df.groupby(['_month', wilayah_col]).size()
                self._assign_month_nested_counts(result.monthly_projects_by_wilayah, project_counts)
            
            return result
            
        except Exception as e:
            print(f"Error loading PROYEK file: {e}")
            return None
    
    def _find_column(self, df: pd.DataFrame, patterns: List[str]) -> Optional[str]:
        """Find column matching any of the patterns."""
        for col in df.columns:
            col_lower = str(col).lower()
            for pattern in patterns:
                if pattern.lower() in col_lower:
                    return col
        return None
    
    def _find_exact_column(self, df: pd.DataFrame, column_name: str) -> Optional[str]:
        """Find column with exact name match (case-insensitive)."""
        for col in df.columns:
            if str(col).lower().strip() == column_name.lower().strip():
                return col
        return None
    
    def get_months_for_period(self, period_type: str, period: str) -> List[str]:
        """Get list of months for a given period."""
        if period_type == "Tahunan":
            return NAMA_BULAN
        elif period_type == "Triwulan":
            return TRIWULAN_KE_BULAN.get(period, [])
        elif period_type == "Semester":
            if period == "Semester I":
                return NAMA_BULAN[:6]
            else:
                return NAMA_BULAN[6:]
        return NAMA_BULAN
