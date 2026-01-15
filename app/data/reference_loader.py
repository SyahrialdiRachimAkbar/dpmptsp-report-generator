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
    monthly_by_wilayah: Dict[str, Dict[str, float]] = field(default_factory=dict)  # Month → Wilayah → investment
    
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
    
    def get_period_by_wilayah(self, months: List[str]) -> Dict[str, float]:
        """Get investment by wilayah for specified months."""
        result = {}
        for month in months:
            if month in self.monthly_by_wilayah:
                for wilayah, investment in self.monthly_by_wilayah[month].items():
                    result[wilayah] = result.get(wilayah, 0) + investment
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
                indo_months = {
                    'Januari': 1, 'Februari': 2, 'Maret': 3, 'April': 4,
                    'Mei': 5, 'Juni': 6, 'Juli': 7, 'Agustus': 8,
                    'September': 9, 'Oktober': 10, 'November': 11, 'Desember': 12
                }
                
                for indo, idx in indo_months.items():
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
            raw_sheet = None
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
                year = self.extract_year_from_filename(filename) or 2025
            
            # Parse dates to months
            if date_col:
                df['_month'] = df[date_col].apply(self._parse_date_to_month)
            else:
                df['_month'] = None
            
            # Initialize result
            result = NIBReferenceData(year=year)
            
            # Process per month
            for month in NAMA_BULAN:
                month_df = df[df['_month'] == month] if date_col else df
                
                if month_df.empty:
                    continue
                
                # Count unique NIB per month
                unique_nib = month_df[nib_col].dropna().nunique()
                result.monthly_totals[month] = unique_nib
                
                # By Kab/Kota (unique NIB per kab per month)
                if kab_col:
                    kab_counts = month_df.groupby(kab_col)[nib_col].nunique()
                    for kab, count in kab_counts.items():
                        if kab not in result.by_kab_kota:
                            result.by_kab_kota[kab] = {}
                        result.by_kab_kota[kab][month] = count
                
                # By PM Status (unique NIB per PM per month)
                if pm_col:
                    pm_counts = month_df.groupby(pm_col)[nib_col].nunique()
                    for pm, count in pm_counts.items():
                        pm_key = str(pm).upper().strip()
                        if pm_key not in result.by_pm_status:
                            result.by_pm_status[pm_key] = {}
                        result.by_pm_status[pm_key][month] = count
                
                # By Skala Usaha (unique NIB per skala per month)
                if skala_col:
                    skala_counts = month_df.groupby(skala_col)[nib_col].nunique()
                    for skala, count in skala_counts.items():
                        if skala not in result.by_skala_usaha:
                            result.by_skala_usaha[skala] = {}
                        result.by_skala_usaha[skala][month] = count
                        
                # Detailed breakdown for Kab/Kota x PM
                if kab_col and pm_col:
                    kab_pm = month_df.groupby([kab_col, pm_col])[nib_col].nunique()
                    for (kab, pm), count in kab_pm.items():
                        if kab not in result.kab_pm_monthly:
                            result.kab_pm_monthly[kab] = {}
                        if month not in result.kab_pm_monthly[kab]:
                            result.kab_pm_monthly[kab][month] = {}
                        
                        pm_key = str(pm).upper().strip()
                        result.kab_pm_monthly[kab][month][pm_key] = count
                
                # Detailed breakdown for Kab/Kota x Skala
                if kab_col and skala_col:
                    kab_skala = month_df.groupby([kab_col, skala_col])[nib_col].nunique()
                    for (kab, skala), count in kab_skala.items():
                        if kab not in result.kab_skala_monthly:
                            result.kab_skala_monthly[kab] = {}
                        if month not in result.kab_skala_monthly[kab]:
                            result.kab_skala_monthly[kab][month] = {}
                            
                        result.kab_skala_monthly[kab][month][skala] = count
            
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
            sector_col = self._find_column(df, ['sektor', 'sector', 'kbli', 'judul_kbli'])
            
            if year is None:
                year = self.extract_year_from_filename(filename) or 2025
            
            # Parse dates
            if date_col:
                df['_month'] = df[date_col].apply(self._parse_date_to_month)
            else:
                # If no date, put all in first month
                df['_month'] = 'Januari'
            
            result = PBOSSReferenceData(year=year)
            
            # Process per month
            for month in NAMA_BULAN:
                month_df = df[df['_month'] == month]
                
                if month_df.empty:
                    continue
                
                # Count permits by risk level
                if risk_col:
                    risk_counts = month_df[risk_col].value_counts()
                    result.monthly_risk[month] = {}
                    for risk, count in risk_counts.items():
                        risk_str = str(risk).strip().upper()
                        # Map short codes to full names
                        risk_name = self.RISK_MAP.get(risk_str, risk_str)
                        result.monthly_risk[month][risk_name] = count
                
                # Count permits by sector
                if sector_col:
                    sector_counts = month_df[sector_col].value_counts().head(10)  # Top 10 sectors
                    result.monthly_sector[month] = dict(sector_counts)
            
            # Calculate total permits
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
            
            if year is None:
                year = self.extract_year_from_filename(filename) or 2025
            
            # Parse dates
            if date_col:
                df['_month'] = df[date_col].apply(self._parse_date_to_month)
            else:
                df['_month'] = 'Januari'
            
            result = ProyekReferenceData(year=year)
            
            # Process per month (sum all, no dedup)
            for month in NAMA_BULAN:
                month_df = df[df['_month'] == month]
                
                if month_df.empty:
                    continue
                
                # Total investment for month
                if investment_col:
                    result.monthly_investment[month] = month_df[investment_col].sum()
                    
                    # PMA investment
                    if pm_col:
                        pma_df = month_df[month_df[pm_col].str.upper().str.contains('PMA', na=False)]
                        pmdn_df = month_df[month_df[pm_col].str.upper().str.contains('PMDN', na=False)]
                        result.monthly_pma[month] = pma_df[investment_col].sum()
                        result.monthly_pmdn[month] = pmdn_df[investment_col].sum()
                    
                    # By wilayah
                    if wilayah_col:
                        wilayah_sums = month_df.groupby(wilayah_col)[investment_col].sum()
                        result.monthly_by_wilayah[month] = dict(wilayah_sums)
                
                # Labor counts
                if tki_col:
                    result.monthly_tki[month] = int(month_df[tki_col].sum())
                if tka_col:
                    result.monthly_tka[month] = int(month_df[tka_col].fillna(0).sum())
                
                # Project count
                result.monthly_projects[month] = len(month_df)
            
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
