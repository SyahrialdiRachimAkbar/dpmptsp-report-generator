"""
Data Loader Module for DPMPTSP Reporting System

This module handles reading and parsing Excel files containing
NIB (Nomor Induk Berusaha) data from DPMPTSP Provinsi Lampung.
"""

import pandas as pd
import re
import io
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass


@dataclass
class NIBData:
    """Data structure for NIB information per Kabupaten/Kota"""
    kabupaten_kota: str
    pma: int = 0
    pmdn: int = 0
    usaha_mikro: int = 0
    usaha_kecil: int = 0
    usaha_menengah: int = 0
    usaha_besar: int = 0
    total: int = 0
    
    @property
    def umk(self) -> int:
        """UMK = Usaha Mikro + Usaha Kecil"""
        return self.usaha_mikro + self.usaha_kecil
    
    @property
    def non_umk(self) -> int:
        """NON-UMK = Usaha Menengah + Usaha Besar"""
        return self.usaha_menengah + self.usaha_besar


@dataclass
class SektorResikoData:
    """Data structure for risk-based permit data per Kabupaten/Kota"""
    kabupaten_kota: str
    # Risk levels
    risiko_rendah: int = 0
    risiko_menengah_rendah: int = 0
    risiko_menengah_tinggi: int = 0
    risiko_tinggi: int = 0
    # Sectors
    sektor_energi: int = 0
    sektor_kelautan: int = 0
    sektor_kesehatan: int = 0
    sektor_komunikasi: int = 0
    sektor_pariwisata: int = 0
    sektor_perhubungan: int = 0
    sektor_perindustrian: int = 0
    sektor_pertanian: int = 0
    total: int = 0
    
    @property
    def total_risiko(self) -> int:
        """Total all risk levels"""
        return self.risiko_rendah + self.risiko_menengah_rendah + self.risiko_menengah_tinggi + self.risiko_tinggi


@dataclass
class InvestmentData:
    """Data structure for investment realization data per Wilayah/Sektor"""
    name: str  # Wilayah or Sektor name
    jumlah_rp: float = 0  # Investment value in Rupiah
    proyek: int = 0  # Number of projects
    tki: int = 0  # Tenaga Kerja Indonesia (local workers)
    tka: int = 0  # Tenaga Kerja Asing (foreign workers)
    
    @property
    def total_tenaga_kerja(self) -> int:
        """Total labor absorption"""
        return self.tki + self.tka


@dataclass
class TWSummary:
    """Summary data for a Triwulan from the main summary sheet"""
    triwulan: str  # e.g., "TW I", "TW II"
    year: int
    pma_rp: float = 0  # PMA investment value
    pmdn_rp: float = 0  # PMDN investment value
    total_rp: float = 0  # Total investment (PMA + PMDN)
    proyek: int = 0  # Total projects
    tki: int = 0  # Total TKI
    tka: int = 0  # Total TKA
    target_rp: float = 0  # Annual target (optional)
    percentage: float = 0  # % of target achieved


@dataclass
class InvestmentReport:
    """Investment realization report for a specific period (Triwulan)"""
    triwulan: str  # e.g., "TW I", "TW II"
    year: int
    # PMA data
    pma_total: float = 0  # Total PMA value in Rupiah
    pma_by_wilayah: list = None  # List[InvestmentData]
    pma_by_sektor: list = None  # List[InvestmentData]
    pma_proyek: int = 0
    pma_tki: int = 0
    pma_tka: int = 0
    # PMDN data
    pmdn_total: float = 0  # Total PMDN value in Rupiah
    pmdn_by_wilayah: list = None  # List[InvestmentData]
    pmdn_by_sektor: list = None  # List[InvestmentData]
    pmdn_proyek: int = 0
    pmdn_tki: int = 0
    pmdn_tka: int = 0
    # By country (for PMA)
    by_country: list = None  # List[InvestmentData]
    
    def __post_init__(self):
        if self.pma_by_wilayah is None:
            self.pma_by_wilayah = []
        if self.pma_by_sektor is None:
            self.pma_by_sektor = []
        if self.pmdn_by_wilayah is None:
            self.pmdn_by_wilayah = []
        if self.pmdn_by_sektor is None:
            self.pmdn_by_sektor = []
        if self.by_country is None:
            self.by_country = []
    
    @property
    def total_investasi(self) -> float:
        """Total investment (PMA + PMDN) in Rupiah"""
        return self.pma_total + self.pmdn_total
    
    @property
    def total_proyek(self) -> int:
        """Total project count"""
        return self.pma_proyek + self.pmdn_proyek
    
    @property
    def total_tki(self) -> int:
        """Total local workers"""
        return self.pma_tki + self.pmdn_tki
    
    @property
    def total_tka(self) -> int:
        """Total foreign workers"""
        return self.pma_tka + self.pmdn_tka


class DataLoader:
    """
    Loader for DPMPTSP Excel data files.
    
    Supports two file formats:
    1. Monthly files (e.g., "OLAH DATA OSS BULAN JULI 2025.xlsx")
    2. Quarterly aggregate files (e.g., "OLAH DATA OSS BULANAN TW I 2025.xlsx")
    """
    
    # Column mappings for NIB sheet
    NIB_COLUMNS = {
        'kabupaten_kota': 1,
        'pma': 2,
        'pmdn': 3,
        'usaha_besar': 4,
        'usaha_kecil': 5,
        'usaha_menengah': 6,
        'usaha_mikro': 7,
        'total': 8
    }
    
    # Known sheet names patterns
    MONTHLY_SHEET_PATTERNS = [
        r"PERIZINAN BERUSAHA (\w+)",
        r"PB (\w+)",
        r"SEKTOR & RESIKO (\w+)",
        r"SEKTOR RESIKO (\w+)"
    ]
    
    def __init__(self):
        self.data_cache: Dict[str, pd.DataFrame] = {}
    
    def load_file(self, file_path: Path) -> Dict[str, pd.DataFrame]:
        """
        Load an Excel file and return all relevant sheets as DataFrames.
        
        Args:
            file_path: Path to the Excel file
            
        Returns:
            Dictionary mapping sheet names to DataFrames
        """
        file_path = Path(file_path)
        if not file_path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")
        
        xl = pd.ExcelFile(file_path)
        sheets = {}
        
        for sheet_name in xl.sheet_names:
            try:
                df = pd.read_excel(xl, sheet_name=sheet_name, header=None, keep_default_na=False, na_values=[''])
                sheets[sheet_name] = df
            except Exception as e:
                print(f"Warning: Could not load sheet '{sheet_name}': {e}")
        
        return sheets
    
    def load_file_from_bytes(self, file_bytes, filename: str = "") -> Dict[str, pd.DataFrame]:
        """
        Load an Excel file from bytes/BytesIO and return all relevant sheets as DataFrames.
        
        Args:
            file_bytes: BytesIO object containing the Excel file
            filename: Original filename for reference
            
        Returns:
            Dictionary mapping sheet names to DataFrames
        """
        xl = pd.ExcelFile(file_bytes)
        sheets = {}
        
        for sheet_name in xl.sheet_names:
            try:
                df = pd.read_excel(xl, sheet_name=sheet_name, header=None, keep_default_na=False, na_values=[''])
                sheets[sheet_name] = df
            except Exception as e:
                print(f"Warning: Could not load sheet '{sheet_name}': {e}")
        
        return sheets
    
    def load_from_bytes(self, file_bytes, filename: str) -> Dict[str, List]:
        """
        Load a monthly data file from bytes and extract NIB data.
        
        Args:
            file_bytes: BytesIO object containing the Excel file
            filename: Original filename for month/year extraction
            
        Returns:
            Dictionary with 'nib' key containing list of NIBData
        """
        sheets = self.load_file_from_bytes(file_bytes, filename)
        result = {}
        
        # Check if this is a quarterly file (has multiple months in sheet names)
        is_quarterly = self._is_quarterly_file(filename, sheets)
        
        if is_quarterly:
            # Return special marker to indicate quarterly file processing needed
            result['is_quarterly'] = True
            result['sheets'] = sheets
            result['year'] = self.extract_year_from_filename(filename)
            return result
        
        # Regular monthly file processing
        # Possible NIB sheet names (case-insensitive search)
        nib_sheet_names = ["NIB", "SKALA NIB", "SKALA", "DATA NIB"]
        
        # Find and parse NIB sheet
        for sheet_name, df in sheets.items():
            sheet_upper = sheet_name.upper().strip()
            if any(nib_name in sheet_upper for nib_name in nib_sheet_names):
                result['nib'] = self.parse_nib_sheet(df)
                break
        
        # If no NIB sheet found, try first sheet that has appropriate structure
        if 'nib' not in result:
            for sheet_name, df in sheets.items():
                nib_data = self.parse_nib_sheet(df)
                if nib_data:  # If we got valid data
                    result['nib'] = nib_data
                    break
        
        # Extract metadata from filename
        result['month'] = self.extract_month_from_filename(filename)
        result['year'] = self.extract_year_from_filename(filename)
        
        return result
    
    def _is_quarterly_file(self, filename: str, sheets: Dict) -> bool:
        """Check if file is a quarterly aggregate file with multiple months."""
        # Check filename for quarterly indicators
        filename_upper = filename.upper()
        if "TW " in filename_upper or "TRIWULAN" in filename_upper:
            return True
        
        # Check if sheets have month names (indicating multiple months in one file)
        months_found = set()
        month_names = ["JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI",
                       "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"]
        
        for sheet_name in sheets.keys():
            sheet_upper = sheet_name.upper()
            for month in month_names:
                if month in sheet_upper:
                    months_found.add(month)
        
        # If more than one month found, it's a quarterly file
        return len(months_found) > 1
    
    def _detect_sheet_content_type(self, df: pd.DataFrame) -> str:
        """
        Detect content type based on headers.
        Returns: 'PB', 'SEKTOR_RESIKO', or 'UNKNOWN'
        """
        # Scan first few rows for keywords
        for idx in range(min(20, len(df))):
            row_str = str(df.iloc[idx].values).upper()
            if "RESIKO" in row_str or "RISIKO" in row_str or "SEKTOR" in row_str:
                return 'SEKTOR_RESIKO'
            if "PMA" in row_str and "PMDN" in row_str:
                return 'PB'
        return 'UNKNOWN'

    def load_quarterly_file(self, file_bytes, filename: str) -> Dict[str, Dict]:
        """
        Load a quarterly file and return data organized by month.
        """
        sheets = self.load_file_from_bytes(file_bytes, filename)
        year = self.extract_year_from_filename(filename)
        
        monthly_data = {}
        month_names = ["JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI",
                       "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"]
        
        # Temporary storage to gather all potential data for each month
        # Structure: {MonthName: {'pb': List[NIBData], 'sr': List[SektorResikoData]}}
        temp_month_data = {}

        for sheet_name, df in sheets.items():
            sheet_upper = sheet_name.upper()
            
            # Find which month this sheet belongs to
            found_month = None
            for month in month_names:
                if month in sheet_upper:
                    found_month = month.capitalize()
                    break
            
            if not found_month:
                continue
                
            if found_month not in temp_month_data:
                temp_month_data[found_month] = {'pb': [], 'sr': []}
            
            # Detect content type
            content_type = self._detect_sheet_content_type(df)
            
            # Parse based on content type, regardless of sheet name
            if content_type == 'SEKTOR_RESIKO':
                sr_data = self.parse_sektor_resiko_sheet(df)
                if sr_data:
                    temp_month_data[found_month]['sr'].extend(sr_data)
            
            elif content_type == 'PB':
                pb_data = self.parse_perizinan_berusaha_sheet(df)
                if pb_data:
                    temp_month_data[found_month]['pb'].extend(pb_data)
            
            else:
                # Fallback to name-based detection if content is ambiguous
                if "RESIKO" in sheet_upper or "RISIKO" in sheet_upper or "SEKTOR" in sheet_upper:
                    sr_data = self.parse_sektor_resiko_sheet(df)
                    if sr_data:
                        temp_month_data[found_month]['sr'].extend(sr_data)
                elif "PB" in sheet_upper or "PERIZINAN" in sheet_upper or "NIB" in sheet_upper:
                    pb_data = self.parse_perizinan_berusaha_sheet(df)
                    if pb_data:
                        temp_month_data[found_month]['pb'].extend(pb_data)

        # Finalize data for each month
        for month_name, data in temp_month_data.items():
            final_nib_data = []
            
            # Priority 1: Use PB data (contains PMA/PMDN breakdown)
            if data['pb']:
                final_nib_data = data['pb']
            
            # Priority 2: Use SR data (Convert to NIBData, contains valid Total)
            # Only if PB data is missing. 
            # This handles the case where "Perizinan Berusaha Mei" actually contains Risk data
            elif data['sr']:
                print(f"Using Sektor Resiko data as fallback for {month_name}")
                final_nib_data = [
                    NIBData(
                        kabupaten_kota=item.kabupaten_kota,
                        total=item.total,
                        # Unfortunately we lose PMA/PMDN distinction here, set to 0
                        pma=0, pmdn=0,
                        usaha_mikro=0, usaha_kecil=0, usaha_menengah=0, usaha_besar=0
                    )
                    for item in data['sr']
                ]
            
            if final_nib_data:
                monthly_data[month_name] = {
                    'month': month_name,
                    'year': year,
                    'nib': final_nib_data
                }
        
        return monthly_data
    
    def parse_perizinan_berusaha_sheet(self, df: pd.DataFrame) -> List[NIBData]:
        """
        Parse a PERIZINAN BERUSAHA sheet from quarterly files.
        Structure: Kab/Kota, PMA, PMDN, [other cols], JUMLAH (last col)
        """
        results = []
        
        # Find the data start row
        data_start_row, kab_col_idx = self._find_data_start_row(df)
        
        if data_start_row is None:
            return results
        
        # Offsets relative to Kab: PMA(+1), PMDN(+2)
        # Total is ALWAYS in the LAST column (varies by sheet)
        off_pma = 1
        off_pmdn = 2
        
        for idx in range(data_start_row, len(df)):
            row = df.iloc[idx]
            
            # Access Kab/Kota dynamically
            raw_kab_kota = row.iloc[kab_col_idx] if len(row) > kab_col_idx and pd.notna(row.iloc[kab_col_idx]) else None
            kab_kota_str = str(raw_kab_kota).strip() if raw_kab_kota is not None else ""
            
            is_valid_row = False
            
            # Check if it's a valid location or "Null" category with data
            # IMPORTANT: "Null" is a valid category label, not missing data!
            null_indicators = ["null", "none", "nan", "-", ""]
            is_null_loc = kab_kota_str.lower() in null_indicators
            
            # Safe access helper
            def get_val(offset):
                return self._safe_int(row.iloc[kab_col_idx + offset]) if len(row) > kab_col_idx + offset else 0
            
            # Get Total from LAST column (this is the reliable source)
            total_from_last_col = self._safe_int(row.iloc[-1]) if len(row) > 0 else 0
            
            if is_null_loc:
                # For "Null" labeled rows, include if they have data
                pma_check = get_val(off_pma)
                pmdn_check = get_val(off_pmdn)
                
                if pma_check + pmdn_check > 0 or total_from_last_col > 0:
                    kab_kota_str = "Tanpa Lokasi"  # Relabel for display
                    is_valid_row = True
            else:
                # Skip summary/total rows
                skip_keywords = ["JUMLAH", "TOTAL", "GRAND", "STATUS PM", "KABUPATEN", "NO", "URAIAN"]
                if any(x in kab_kota_str.upper() for x in skip_keywords):
                    is_valid_row = False
                else:
                    is_valid_row = True
            
            if not is_valid_row:
                continue
            
            # Extract values
            pma = get_val(off_pma)
            pmdn = get_val(off_pmdn)
            
            # Use last column for Total (most reliable)
            # Only fallback to PMA+PMDN if last column is empty
            final_total = total_from_last_col if total_from_last_col > 0 else (pma + pmdn)
            
            nib_data = NIBData(
                kabupaten_kota=kab_kota_str,
                pma=pma,
                pmdn=pmdn,
                usaha_mikro=0,
                usaha_kecil=0,
                usaha_menengah=0,
                usaha_besar=0,
                total=final_total
            )
            
            results.append(nib_data)
        
        return results
    
    def _merge_nib_data(self, existing: List[NIBData], new: List[NIBData]) -> List[NIBData]:
        """Merge two lists of NIBData by Kabupaten/Kota."""
        merged = {d.kabupaten_kota: d for d in existing}
        
        for item in new:
            if item.kabupaten_kota in merged:
                # Add values
                old = merged[item.kabupaten_kota]
                merged[item.kabupaten_kota] = NIBData(
                    kabupaten_kota=item.kabupaten_kota,
                    pma=old.pma + item.pma,
                    pmdn=old.pmdn + item.pmdn,
                    usaha_mikro=old.usaha_mikro + item.usaha_mikro,
                    usaha_kecil=old.usaha_kecil + item.usaha_kecil,
                    usaha_menengah=old.usaha_menengah + item.usaha_menengah,
                    usaha_besar=old.usaha_besar + item.usaha_besar,
                    total=old.total + item.total
                )
            else:
                merged[item.kabupaten_kota] = item
        
        return list(merged.values())
    
    def extract_month_from_filename(self, filename: str) -> Optional[str]:
        """
        Extract month name from filename.
        
        Examples:
            "OLAH DATA OSS BULAN JULI 2025.xlsx" -> "Juli"
            "OLAH DATA OSS BULAN SEPTEMBER 2025.xlsx" -> "September"
        """
        months = [
            "JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI",
            "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"
        ]
        
        filename_upper = filename.upper()
        for month in months:
            if month in filename_upper:
                return month.capitalize()
        
        return None
    
    def extract_year_from_filename(self, filename: str) -> Optional[int]:
        """Extract year from filename."""
        match = re.search(r'(\d{4})', filename)
        if match:
            return int(match.group(1))
        return None
    
    def parse_nib_sheet(self, df: pd.DataFrame) -> List[NIBData]:
        """
        Parse the NIB sheet to extract data per Kabupaten/Kota.
        """
        results = []
        
        # Find the data start row (after headers) and column index for Kab/Kota
        data_start_row, kab_col_idx = self._find_data_start_row(df)
        
        if data_start_row is None:
            return results
        
        # Determine column offsets assuming standard relative structure
        # Standard: Kab(0), PMA(1), PMDN(2), UB(3), UK(4), UM(5), UMi(6), Total(7)
        # Offsets relative to Kab: +1, +2, etc.
        off_pma = 1
        off_pmdn = 2
        off_ub = 3
        off_uk = 4
        off_um = 5
        off_umi = 6
        off_total = 7
        
        # Process each data row
        for idx in range(data_start_row, len(df)):
            row = df.iloc[idx]
            
            # Handle potential None/NaN values safely
            raw_kab_kota = row.iloc[kab_col_idx] if len(row) > kab_col_idx and pd.notna(row.iloc[kab_col_idx]) else None
            kab_kota_str = str(raw_kab_kota).strip() if raw_kab_kota is not None else ""
            
            is_valid_row = False
            kab_kota = kab_kota_str
            
            # Check if it's a valid location or "Null"/"Empty" with data
            null_indicators = ["null", "none", "nan", "-", ""]
            is_null_loc = kab_kota_str.lower() in null_indicators
            
            if is_null_loc:
                # Calculate sum of components to check if row has data
                comp_sum = 0
                comp_sum += self._safe_int(row.iloc[kab_col_idx + off_pma]) if len(row) > kab_col_idx + off_pma else 0
                comp_sum += self._safe_int(row.iloc[kab_col_idx + off_pmdn]) if len(row) > kab_col_idx + off_pmdn else 0
                comp_sum += self._safe_int(row.iloc[kab_col_idx + off_ub]) if len(row) > kab_col_idx + off_ub else 0
                comp_sum += self._safe_int(row.iloc[kab_col_idx + off_uk]) if len(row) > kab_col_idx + off_uk else 0
                comp_sum += self._safe_int(row.iloc[kab_col_idx + off_um]) if len(row) > kab_col_idx + off_um else 0
                comp_sum += self._safe_int(row.iloc[kab_col_idx + off_umi]) if len(row) > kab_col_idx + off_umi else 0
                
                # Also check explicit total column
                explicit_total = self._safe_int(row.iloc[kab_col_idx + off_total]) if len(row) > kab_col_idx + off_total else 0
                
                if comp_sum > 0 or explicit_total > 0:
                    kab_kota = "Tanpa Lokasi"
                    is_valid_row = True
            else:
                # Normal location
                skip_keywords = ["JUMLAH", "TOTAL", "GRAND", "STATUS PM", "SKALA USAHA", "KABUPATEN", "NO"]
                if any(x in kab_kota_str.upper() for x in skip_keywords):
                    is_valid_row = False
                else:
                    is_valid_row = True
            
            if not is_valid_row:
                continue
            
            # Extract values with safe conversion
            pma = self._safe_int(row.iloc[kab_col_idx + off_pma]) if len(row) > kab_col_idx + off_pma else 0
            pmdn = self._safe_int(row.iloc[kab_col_idx + off_pmdn]) if len(row) > kab_col_idx + off_pmdn else 0
            u_besar = self._safe_int(row.iloc[kab_col_idx + off_ub]) if len(row) > kab_col_idx + off_ub else 0
            u_kecil = self._safe_int(row.iloc[kab_col_idx + off_uk]) if len(row) > kab_col_idx + off_uk else 0
            u_menengah = self._safe_int(row.iloc[kab_col_idx + off_um]) if len(row) > kab_col_idx + off_um else 0
            u_mikro = self._safe_int(row.iloc[kab_col_idx + off_umi]) if len(row) > kab_col_idx + off_umi else 0
            explicit_total = self._safe_int(row.iloc[kab_col_idx + off_total]) if len(row) > kab_col_idx + off_total else 0
            
            final_total = explicit_total if explicit_total > 0 else (pma + pmdn)
            
            nib_data = NIBData(
                kabupaten_kota=kab_kota,
                pma=pma,
                pmdn=pmdn,
                usaha_besar=u_besar,
                usaha_kecil=u_kecil,
                usaha_menengah=u_menengah,
                usaha_mikro=u_mikro,
                total=final_total
            )
            
            results.append(nib_data)
        
        return results
    
    def _find_data_start_row(self, df: pd.DataFrame) -> Tuple[Optional[int], int]:
        """
        Find the row where actual data starts and the column index of Kabupaten/Kota.
        Returns (row_idx, col_idx). Returns (None, -1) if not found.
        """
        # Strategy 1: Header search (dynamic column)
        for idx in range(min(50, len(df))):
            row = df.iloc[idx]
            # Check col 0 and 1
            for col_idx in [0, 1]:
                if col_idx < len(row):
                    cell = str(row.iloc[col_idx]).upper().strip()
                    if "KABUPATEN" in cell or "KAB/KOTA" in cell or "KAB. / KOTA" in cell:
                        return idx + 1, col_idx
        
        # Strategy 2: Fallback - First data row
        for idx in range(len(df)):
            row = df.iloc[idx]
            for col_idx in [0, 1]:
                if col_idx < len(row) and pd.notna(row.iloc[col_idx]):
                    cell = str(row.iloc[col_idx]).strip()
                    if (cell.startswith("Kab.") or cell.startswith("Kota") or 
                        cell.startswith("KAB.") or cell.startswith("KOTA")):
                        return idx, col_idx
        return None, -1
    
    def _safe_int(self, value) -> int:
        """Safely convert a value to integer."""
        if pd.isna(value):
            return 0
        try:
            return int(float(value))
        except (ValueError, TypeError):
            return 0
    
    def load_monthly_data(self, file_path: Path) -> Dict[str, List[NIBData]]:
        """
        Load a monthly data file and extract NIB data.
        
        Args:
            file_path: Path to monthly Excel file
            
        Returns:
            Dictionary with 'nib' key containing list of NIBData
        """
        sheets = self.load_file(file_path)
        result = {}
        
        # Possible NIB sheet names (case-insensitive search)
        nib_sheet_names = ["NIB", "SKALA NIB", "SKALA", "DATA NIB"]
        
        # Find and parse NIB sheet
        for sheet_name, df in sheets.items():
            sheet_upper = sheet_name.upper().strip()
            if any(nib_name in sheet_upper for nib_name in nib_sheet_names):
                result['nib'] = self.parse_nib_sheet(df)
                break
        
        # If no NIB sheet found, try first sheet that has appropriate structure
        if 'nib' not in result:
            for sheet_name, df in sheets.items():
                nib_data = self.parse_nib_sheet(df)
                if nib_data:  # If we got valid data
                    result['nib'] = nib_data
                    break
        
        # Extract metadata
        filename = Path(file_path).name
        result['month'] = self.extract_month_from_filename(filename)
        result['year'] = self.extract_year_from_filename(filename)
        
        return result
    
    def load_quarterly_data(self, file_path: Path) -> Dict[str, Dict[str, List[NIBData]]]:
        """
        Load a quarterly aggregate file containing multiple months.
        
        These files have separate sheets for each month like:
        - "PERIZINAN BERUSAHA JANUARI"
        - "SEKTOR & RESIKO JANUARI"
        - etc.
        
        Args:
            file_path: Path to quarterly Excel file
            
        Returns:
            Dictionary mapping month names to their data
        """
        sheets = self.load_file(file_path)
        result = {}
        
        # Group sheets by month
        months_found = set()
        for sheet_name in sheets.keys():
            for pattern in self.MONTHLY_SHEET_PATTERNS:
                match = re.search(pattern, sheet_name, re.IGNORECASE)
                if match:
                    month = match.group(1).capitalize()
                    months_found.add(month)
        
        # For each month found, try to find and parse the NIB data
        # Note: Quarterly files might have different structure
        # This is a simplified implementation
        
        result['year'] = self.extract_year_from_filename(Path(file_path).name)
        result['months'] = list(months_found)
        result['sheets'] = sheets
        
        return result
    
    def get_nib_dataframe(self, nib_data_list: List[NIBData]) -> pd.DataFrame:
        """
        Convert list of NIBData to a pandas DataFrame.
        
        Args:
            nib_data_list: List of NIBData objects
            
        Returns:
            DataFrame with NIB data
        """
        if not nib_data_list:
            return pd.DataFrame()
        
        data = []
        for nib in nib_data_list:
            data.append({
                'Kabupaten/Kota': nib.kabupaten_kota,
                'PMA': nib.pma,
                'PMDN': nib.pmdn,
                'Usaha Mikro': nib.usaha_mikro,
                'Usaha Kecil': nib.usaha_kecil,
                'Usaha Menengah': nib.usaha_menengah,
                'Usaha Besar': nib.usaha_besar,
                'UMK': nib.umk,
                'NON-UMK': nib.non_umk,
                'Total': nib.total,
            })
        
        return pd.DataFrame(data)
    
    def parse_sektor_resiko_sheet(self, df: pd.DataFrame) -> List[SektorResikoData]:
        """
        Parse the SEKTOR RESIKO sheet.
        """
        results = []
        data_start_row, kab_col_idx = self._find_data_start_row(df)
        
        if data_start_row is None:
            return results
        
        # Standard Sektor Resiko Offsets relative to Kab column
        # Based on: Kab(0), MR(1), MT(2), R(3), T(4), Eng(5)... Total(13)
        # So offsets: +1, +2, ... +13
        
        for idx in range(data_start_row, len(df)):
            row = df.iloc[idx]
            
            raw_kab_kota = row.iloc[kab_col_idx] if len(row) > kab_col_idx and pd.notna(row.iloc[kab_col_idx]) else None
            kab_kota_str = str(raw_kab_kota).strip() if raw_kab_kota is not None else ""
            
            is_valid_row = False
            kab_kota = kab_kota_str
            
            null_indicators = ["null", "none", "nan", "-", ""]
            is_null_loc = kab_kota_str.lower() in null_indicators
            
            # Safe row access Helper
            def get_val(offset):
                return self._safe_int(row.iloc[kab_col_idx + offset]) if len(row) > kab_col_idx + offset else 0
            
            if is_null_loc:
                # Check for total data (last column ~offset 13) or components
                # Check first few risk columns (offsets 1-4)
                risk_sum = get_val(1) + get_val(2) + get_val(3) + get_val(4)
                total_val = get_val(13)
                
                if total_val > 0 or risk_sum > 0:
                    kab_kota = "Tanpa Lokasi"
                    is_valid_row = True
            else:
                skip_keywords = ["JUMLAH", "TOTAL", "GRAND", "RISIKO", "SEKTOR", "NO"]
                if any(x in kab_kota_str.upper() for x in skip_keywords):
                    is_valid_row = False
                else:
                    is_valid_row = True
            
            if not is_valid_row:
                continue
            
            # Using offsets
            sektor_data = SektorResikoData(
                kabupaten_kota=str(kab_kota).strip(),
                risiko_menengah_rendah=get_val(1),
                risiko_menengah_tinggi=get_val(2),
                risiko_rendah=get_val(3),
                risiko_tinggi=get_val(4),
                sektor_energi=get_val(5),
                sektor_kelautan=get_val(6),
                sektor_kesehatan=get_val(7),
                sektor_komunikasi=get_val(8),
                sektor_pariwisata=get_val(9),
                sektor_perhubungan=get_val(10),
                sektor_perindustrian=get_val(11),
                sektor_pertanian=get_val(12),
                total=get_val(13),
            )
            
            results.append(sektor_data)
        
        return results
    
    def get_sektor_resiko_dataframe(self, sektor_data_list: List[SektorResikoData]) -> pd.DataFrame:
        """
        Convert list of SektorResikoData to a pandas DataFrame.
        
        Args:
            sektor_data_list: List of SektorResikoData objects
            
        Returns:
            DataFrame with risk-based permit data
        """
        if not sektor_data_list:
            return pd.DataFrame()
        
        data = []
        for item in sektor_data_list:
            data.append({
                'Kabupaten/Kota': item.kabupaten_kota,
                'Risiko Rendah': item.risiko_rendah,
                'Risiko Menengah Rendah': item.risiko_menengah_rendah,
                'Risiko Menengah Tinggi': item.risiko_menengah_tinggi,
                'Risiko Tinggi': item.risiko_tinggi,
                'Energi': item.sektor_energi,
                'Kelautan': item.sektor_kelautan,
                'Kesehatan': item.sektor_kesehatan,
                'Komunikasi': item.sektor_komunikasi,
                'Pariwisata': item.sektor_pariwisata,
                'Perhubungan': item.sektor_perhubungan,
                'Perindustrian': item.sektor_perindustrian,
                'Pertanian': item.sektor_pertanian,
                'Total': item.total,
            })
        
        return pd.DataFrame(data)
    
    def load_realisasi_investasi(self, file_bytes, filename: str = "") -> Dict[str, InvestmentReport]:
        """
        Load REALISASI INVESTASI file and parse investment data by Triwulan.
        
        The file contains sheets like:
        - REALISASI INVESTASI 2025: Summary per TW
        - PMA SEKTOR TW I: PMA by sector for TW I
        - PMA WILAYAH TW I: PMA by region for TW I
        - PMDN SEKTOR TW I: PMDN by sector for TW I
        - PMDN WILAYAH TW I: PMDN by region for TW I
        
        Returns:
            Dictionary mapping Triwulan name to InvestmentReport
        """
        sheets = self.load_file_from_bytes(file_bytes, filename)
        year = self.extract_year_from_filename(filename) or 2025
        
        # Initialize reports for each Triwulan
        reports = {}
        triwulan_list = ["TW I", "TW II", "TW III", "TW IV"]
        
        for tw in triwulan_list:
            reports[tw] = InvestmentReport(triwulan=tw, year=year)
        
        # Parse each sheet
        for sheet_name, df in sheets.items():
            sheet_upper = sheet_name.upper()
            
            # Determine which Triwulan this sheet belongs to
            # Use regex to match exact triwulan (avoid TW I matching TW II, TW III, TW IV)
            tw = None
            # Check TW IV first (most specific), then TW III, TW II, TW I
            for t in ["TW IV", "TW III", "TW II", "TW I"]:
                # Pattern variations: "TW I", "TWI", "TW1"
                tw_clean = t.replace(" ", "")  # TWI, TWII, TWIII, TWIV
                tw_roman = t  # TW I, TW II, TW III, TW IV
                
                if tw_clean in sheet_upper.replace(" ", ""):
                    tw = t
                    break
                if tw_roman in sheet_upper:
                    tw = t
                    break
            
            if tw is None:
                continue
            
            # Determine sheet type and parse
            if "PMA" in sheet_upper and "SEKTOR" in sheet_upper:
                reports[tw].pma_by_sektor = self._parse_investment_sheet(df)
                reports[tw].pma_total = sum(d.jumlah_rp for d in reports[tw].pma_by_sektor)
                reports[tw].pma_proyek = sum(d.proyek for d in reports[tw].pma_by_sektor)
                reports[tw].pma_tki = sum(d.tki for d in reports[tw].pma_by_sektor)
                reports[tw].pma_tka = sum(d.tka for d in reports[tw].pma_by_sektor)
                
            elif "PMA" in sheet_upper and "WILAYAH" in sheet_upper:
                reports[tw].pma_by_wilayah = self._parse_investment_sheet(df)
                # If totals not set from sektor sheet, use wilayah data
                if reports[tw].pma_total == 0:
                    reports[tw].pma_total = sum(d.jumlah_rp for d in reports[tw].pma_by_wilayah)
                    reports[tw].pma_proyek = sum(d.proyek for d in reports[tw].pma_by_wilayah)
                    reports[tw].pma_tki = sum(d.tki for d in reports[tw].pma_by_wilayah)
                    reports[tw].pma_tka = sum(d.tka for d in reports[tw].pma_by_wilayah)
                    
            elif "PMDN" in sheet_upper and "SEKTOR" in sheet_upper:
                reports[tw].pmdn_by_sektor = self._parse_investment_sheet(df)
                reports[tw].pmdn_total = sum(d.jumlah_rp for d in reports[tw].pmdn_by_sektor)
                reports[tw].pmdn_proyek = sum(d.proyek for d in reports[tw].pmdn_by_sektor)
                reports[tw].pmdn_tki = sum(d.tki for d in reports[tw].pmdn_by_sektor)
                reports[tw].pmdn_tka = sum(d.tka for d in reports[tw].pmdn_by_sektor)
                
            elif "PMDN" in sheet_upper and "WILAYAH" in sheet_upper:
                reports[tw].pmdn_by_wilayah = self._parse_investment_sheet(df)
                if reports[tw].pmdn_total == 0:
                    reports[tw].pmdn_total = sum(d.jumlah_rp for d in reports[tw].pmdn_by_wilayah)
                    reports[tw].pmdn_proyek = sum(d.proyek for d in reports[tw].pmdn_by_wilayah)
                    reports[tw].pmdn_tki = sum(d.tki for d in reports[tw].pmdn_by_wilayah)
                    reports[tw].pmdn_tka = sum(d.tka for d in reports[tw].pmdn_by_wilayah)
                    
            elif "NEGARA" in sheet_upper:
                reports[tw].by_country = self._parse_investment_sheet(df)
        
        # Filter out empty reports
        reports = {tw: r for tw, r in reports.items() 
                   if r.pma_total > 0 or r.pmdn_total > 0 or r.pma_by_wilayah or r.pmdn_by_wilayah}
        
        return reports
    
    def _parse_investment_sheet(self, df: pd.DataFrame) -> List[InvestmentData]:
        """
        Parse an investment sheet (PMA/PMDN by wilayah or sektor).
        
        Standard structure:
        NO | WILAYAH/SEKTOR | JUMLAH | PROYEK | TKI | TKA
                            (Rp.)
        """
        results = []
        
        # Find header row with WILAYAH/SEKTOR as anchor
        # Must also have NO in the same row to distinguish from title
        header_row = None
        name_col = None
        
        for idx in range(min(20, len(df))):
            row = df.iloc[idx]
            row_str = ' '.join(str(v).upper() for v in row)
            
            # Look for row that has both "NO" and "WILAYAH" or "SEKTOR" - this is the true header
            has_no = any(str(v).upper().strip() == "NO" for v in row)
            has_wilayah_or_sektor = any("WILAYAH" in str(v).upper() or "SEKTOR" in str(v).upper() for v in row)
            
            if has_no and has_wilayah_or_sektor:
                header_row = idx
                # Find the column with WILAYAH or SEKTOR
                for col_idx, val in enumerate(row):
                    val_str = str(val).upper().strip()
                    if "WILAYAH" in val_str or "SEKTOR" in val_str:
                        name_col = col_idx
                        break
                break
        
        if header_row is None or name_col is None:
            return results
        
        # Determine column positions based on the same row as header
        # Structure: NO(0), WILAYAH(1), JUMLAH(2), PROYEK(3), TKI(4), TKA(5)
        # name_col is WILAYAH position, so:
        jumlah_col = name_col + 1  # JUMLAH is right after WILAYAH
        proyek_col = name_col + 2  # PROYEK is next
        tki_col = name_col + 3     # TKI
        tka_col = name_col + 4     # TKA
        
        # Skip to data rows (header row + 2 to skip the subheader with "(Rp.)")
        data_start = header_row + 2
        
        # Parse data rows
        for idx in range(data_start, len(df)):
            row = df.iloc[idx]
            
            # Get name
            raw_name = row.iloc[name_col] if name_col is not None and len(row) > name_col else None
            if pd.isna(raw_name):
                continue
            name = str(raw_name).strip()
            
            # Skip summary rows
            skip_keywords = ["JUMLAH", "TOTAL", "GRAND", "NO"]
            if any(x in name.upper() for x in skip_keywords) or not name:
                continue
            
            # Get values
            def safe_float(val):
                if pd.isna(val):
                    return 0.0
                try:
                    return float(val)
                except (ValueError, TypeError):
                    return 0.0
            
            def safe_int(val):
                if pd.isna(val):
                    return 0
                try:
                    return int(float(val))
                except (ValueError, TypeError):
                    return 0
            
            jumlah = safe_float(row.iloc[jumlah_col]) if jumlah_col is not None and len(row) > jumlah_col else 0.0
            proyek = safe_int(row.iloc[proyek_col]) if proyek_col is not None and len(row) > proyek_col else 0
            tki = safe_int(row.iloc[tki_col]) if tki_col is not None and len(row) > tki_col else 0
            tka = safe_int(row.iloc[tka_col]) if tka_col is not None and len(row) > tka_col else 0
            
            if jumlah > 0 or proyek > 0:
                results.append(InvestmentData(
                    name=name,
                    jumlah_rp=jumlah,
                    proyek=proyek,
                    tki=tki,
                    tka=tka
                ))
        
        return results
    
    def parse_investment_summary(self, file_bytes: io.BytesIO, filename: str = "") -> Dict[str, TWSummary]:
        """
        Parse the summary sheet from REALISASI INVESTASI file.
        Returns Dict mapping TW name to TWSummary.
        
        Args:
            file_bytes: BytesIO object of the Excel file
            filename: Original filename to detect year
            
        Returns:
            Dict[str, TWSummary]: e.g., {"TW I": TWSummary(...), "TW II": TWSummary(...)}
        """
        results = {}
        
        # Detect year from filename
        year = 2025
        import re
        year_match = re.search(r'(\d{4})', filename)
        if year_match:
            year = int(year_match.group(1))
        
        try:
            xl = pd.ExcelFile(file_bytes)
        except Exception as e:
            print(f"Error reading Excel file: {e}")
            return results
        
        # Find summary sheet (pattern: REALISASI INVESTASI YYYY or similar)
        summary_sheet = None
        for sheet in xl.sheet_names:
            if 'REALISASI INVESTASI' in sheet.upper() and any(c.isdigit() for c in sheet):
                summary_sheet = sheet
                break
            elif sheet.upper().startswith('REALISASI INVESTASI'):
                summary_sheet = sheet
                break
        
        if not summary_sheet:
            return results
        
        try:
            df = pd.read_excel(xl, sheet_name=summary_sheet, header=None)
        except Exception as e:
            print(f"Error reading summary sheet: {e}")
            return results
        
        # Find target value (usually in row with TARGET)
        target_rp = 0
        for idx in range(min(10, len(df))):
            row = df.iloc[idx]
            for col_idx, val in enumerate(row):
                if pd.notna(val) and 'TARGET' in str(val).upper():
                    # Target value is usually in column 1
                    try:
                        target_rp = float(df.iloc[idx, 1]) if pd.notna(df.iloc[idx, 1]) else 0
                    except:
                        pass
                    break
        
        # Parse TW rows
        # Structure: PERIODE column has TW I, TW II, etc.
        # Columns: NO, TARGET, PERIODE, PMA(Rp.), PMDN(Rp.), JUMLAH(Rp.), %, PROYEK, TKI, TKA
        tw_patterns = ["TW I", "TW II", "TW III", "TW IV"]
        
        for idx in range(len(df)):
            row = df.iloc[idx]
            row_str = ' '.join(str(v).upper().strip() for v in row if pd.notna(v))
            
            for tw in tw_patterns:
                # Check if this row contains exactly this TW (not as part of another)
                if tw.upper() in row_str:
                    # Find which cell has the TW value
                    periode_col = None
                    for col_idx, val in enumerate(row):
                        if pd.notna(val) and tw.upper() == str(val).upper().strip():
                            periode_col = col_idx
                            break
                    
                    if periode_col is None:
                        continue
                    
                    # Extract values based on column positions relative to PERIODE
                    # Usually: PERIODE(2), PMA(3), PMDN(4), JUMLAH(5), %(6), PROYEK(7), TKI(8), TKA(9)
                    def safe_float(val):
                        if pd.isna(val):
                            return 0.0
                        try:
                            return float(val)
                        except:
                            return 0.0
                    
                    def safe_int(val):
                        if pd.isna(val):
                            return 0
                        try:
                            return int(float(val))
                        except:
                            return 0
                    
                    # Get values - adjust indices based on actual structure
                    pma_col = periode_col + 1
                    pmdn_col = periode_col + 2
                    total_col = periode_col + 3
                    pct_col = periode_col + 4
                    proyek_col = periode_col + 5
                    tki_col = periode_col + 6
                    tka_col = periode_col + 7
                    
                    pma_rp = safe_float(row.iloc[pma_col]) if len(row) > pma_col else 0
                    pmdn_rp = safe_float(row.iloc[pmdn_col]) if len(row) > pmdn_col else 0
                    total_rp = safe_float(row.iloc[total_col]) if len(row) > total_col else 0
                    percentage = safe_float(row.iloc[pct_col]) if len(row) > pct_col else 0
                    proyek = safe_int(row.iloc[proyek_col]) if len(row) > proyek_col else 0
                    tki = safe_int(row.iloc[tki_col]) if len(row) > tki_col else 0
                    tka = safe_int(row.iloc[tka_col]) if len(row) > tka_col else 0
                    
                    # Only add if we have meaningful data
                    if pma_rp > 0 or pmdn_rp > 0 or proyek > 0:
                        results[tw] = TWSummary(
                            triwulan=tw,
                            year=year,
                            pma_rp=pma_rp,
                            pmdn_rp=pmdn_rp,
                            total_rp=total_rp if total_rp > 0 else (pma_rp + pmdn_rp),
                            proyek=proyek,
                            tki=tki,
                            tka=tka,
                            target_rp=target_rp,
                            percentage=percentage
                        )
                    break  # Found this TW, move to next row
        
        return results
def load_excel_file(file_path: str | Path) -> Dict:
    """
    Convenience function to load an Excel file.
    
    Args:
        file_path: Path to Excel file
        
    Returns:
        Parsed data dictionary
    """
    loader = DataLoader()
    return loader.load_monthly_data(Path(file_path))
