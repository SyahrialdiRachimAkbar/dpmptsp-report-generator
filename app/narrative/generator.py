"""
Narrative Generator Module for DPMPTSP Reporting System

This module auto-generates Indonesian language narratives/interpretations
for NIB data reports. Designed to be insight-focused and easy to understand
for general audiences.
"""

from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass
import sys
sys.path.append('..')
from app.data.aggregator import PeriodReport


@dataclass
class Narrative:
    """Container for generated narrative sections"""
    pendahuluan: str
    rekapitulasi_nib: str
    rekapitulasi_kab_kota: str
    status_pm: str
    pelaku_usaha: str
    kesimpulan: str


class NarrativeGenerator:
    """
    Generates automatic narratives in formal Indonesian language.
    
    Features:
    - Insight-focused descriptions
    - Easy to understand for general audiences
    - Formal language style
    - Automatic percentage calculations
    - Trend analysis
    """
    
    # Month names for Indonesian
    BULAN = [
        "", "Januari", "Februari", "Maret", "April", "Mei", "Juni",
        "Juli", "Agustus", "September", "Oktober", "November", "Desember"
    ]
    
    TRIWULAN_BULAN = {
        "TW I": "Januari - Maret",
        "TW II": "April - Juni", 
        "TW III": "Juli - September",
        "TW IV": "Oktober - Desember"
    }
    
    def __init__(self):
        pass
    
    def generate_full_narrative(self, report: PeriodReport, stats: Dict) -> Narrative:
        """
        Generate complete narrative for a period report.
        
        Args:
            report: PeriodReport object
            stats: Summary statistics dictionary
            
        Returns:
            Narrative object with all sections
        """
        return Narrative(
            pendahuluan=self._generate_pendahuluan(report),
            rekapitulasi_nib=self._generate_rekapitulasi_nib(report, stats),
            rekapitulasi_kab_kota=self._generate_rekapitulasi_kab_kota(report, stats),
            status_pm=self._generate_status_pm(report, stats),
            pelaku_usaha=self._generate_pelaku_usaha(report, stats),
            kesimpulan=self._generate_kesimpulan(report, stats)
        )
    
    def _generate_pendahuluan(self, report: PeriodReport) -> str:
        """Generate introduction paragraph."""
        periode_text = self._get_periode_text(report)
        bulan_range = self.TRIWULAN_BULAN.get(report.period_name, "")
        
        text = f"""Laporan ini menyajikan rekapitulasi data Nomor Induk Berusaha (NIB) yang diterbitkan melalui sistem Online Single Submission Risk Based Approach (OSS-RBA) di Provinsi Lampung. Periode laporan mencakup {periode_text} ({bulan_range} {report.year}).

Data yang disajikan meliputi distribusi NIB berdasarkan kabupaten/kota, status penanaman modal (PMA dan PMDN), serta kategori pelaku usaha (UMK dan Non-UMK). Laporan ini bertujuan untuk memberikan gambaran menyeluruh mengenai perkembangan perizinan berusaha di Provinsi Lampung."""
        
        return text
    
    def _generate_rekapitulasi_nib(self, report: PeriodReport, stats: Dict) -> str:
        """Generate NIB summary narrative."""
        total = stats.get('total_nib', 0)
        monthly = stats.get('monthly_totals', {})
        change_pct = stats.get('change_percentage')
        prev_total = stats.get('prev_period_total')
        
        # Format total with thousands separator
        total_formatted = f"{total:,}".replace(",", ".")
        
        # Monthly breakdown text
        monthly_text = ""
        if monthly:
            monthly_parts = []
            for bulan, nilai in monthly.items():
                nilai_formatted = f"{nilai:,}".replace(",", ".")
                monthly_parts.append(f"{bulan} ({nilai_formatted} NIB)")
            monthly_text = ", ".join(monthly_parts)
        
        # Trend analysis
        trend_text = ""
        if change_pct is not None and prev_total is not None:
            prev_formatted = f"{prev_total:,}".replace(",", ".")
            if change_pct > 0:
                trend_text = f"\n\nDibandingkan dengan periode sebelumnya ({prev_formatted} NIB), terjadi peningkatan sebesar {change_pct:.1f}%. Hal ini menunjukkan pertumbuhan positif dalam aktivitas perizinan berusaha di Provinsi Lampung."
            elif change_pct < 0:
                trend_text = f"\n\nDibandingkan dengan periode sebelumnya ({prev_formatted} NIB), terjadi penurunan sebesar {abs(change_pct):.1f}%. Penurunan ini perlu dikaji lebih lanjut untuk mengetahui faktor-faktor yang mempengaruhinya."
            else:
                trend_text = f"\n\nJumlah NIB stabil dibandingkan periode sebelumnya ({prev_formatted} NIB)."
        
        # Monthly trend within period
        monthly_trend = ""
        if len(monthly) > 1:
            values = list(monthly.values())
            if values[-1] > values[0]:
                monthly_trend = f" Terlihat tren peningkatan dari bulan ke bulan dalam periode ini."
            elif values[-1] < values[0]:
                monthly_trend = f" Terlihat tren penurunan dari bulan ke bulan dalam periode ini."
        
        text = f"""Pada {self._get_periode_text(report)}, total NIB yang diterbitkan di Provinsi Lampung mencapai {total_formatted} NIB. Rincian per bulan adalah sebagai berikut: {monthly_text}.{monthly_trend}{trend_text}"""
        
        return text
    
    def _generate_rekapitulasi_kab_kota(self, report: PeriodReport, stats: Dict) -> str:
        """Generate location-based summary narrative."""
        total = stats.get('total_nib', 0)
        top_5 = stats.get('top_5_locations', [])
        
        if not top_5:
            return "Data per kabupaten/kota belum tersedia."
        
        # Top performer
        top_1 = top_5[0]
        top_1_name = top_1['Kabupaten/Kota']
        top_1_total = top_1['Total']
        top_1_pct = (top_1_total / total * 100) if total > 0 else 0
        top_1_formatted = f"{top_1_total:,}".replace(",", ".")
        
        # Other top performers
        others_text = ""
        if len(top_5) > 1:
            others = []
            for loc in top_5[1:3]:  # Top 2-3
                name = loc['Kabupaten/Kota']
                val = loc['Total']
                pct = (val / total * 100) if total > 0 else 0
                val_formatted = f"{val:,}".replace(",", ".")
                others.append(f"{name} ({val_formatted} NIB, {pct:.1f}%)")
            others_text = f" Urutan selanjutnya ditempati oleh {' dan '.join(others)}."
        
        text = f"""Berdasarkan distribusi per kabupaten/kota, {top_1_name} menempati posisi tertinggi dengan {top_1_formatted} NIB ({top_1_pct:.1f}% dari total).{others_text}

Distribusi ini menunjukkan bahwa aktivitas perizinan berusaha terkonsentrasi di beberapa wilayah strategis, terutama di daerah dengan tingkat aktivitas ekonomi yang tinggi."""
        
        return text
    
    def _generate_status_pm(self, report: PeriodReport, stats: Dict) -> str:
        """Generate investment status (PMA/PMDN) narrative."""
        pm_dist = stats.get('pm_distribution', {})
        pma = pm_dist.get('PMA', 0)
        pmdn = pm_dist.get('PMDN', 0)
        pma_pct = pm_dist.get('PMA_pct', 0)
        pmdn_pct = pm_dist.get('PMDN_pct', 0)
        
        pma_formatted = f"{pma:,}".replace(",", ".")
        pmdn_formatted = f"{pmdn:,}".replace(",", ".")
        
        # Format percentage with appropriate precision
        def format_pct(pct):
            if pct < 0.1 and pct > 0:
                return f"{pct:.2f}"
            elif pct < 1 and pct > 0:
                return f"{pct:.2f}"
            elif pct > 99 and pct < 100:
                return f"{pct:.2f}"
            else:
                return f"{pct:.1f}"
        
        pma_pct_str = format_pct(pma_pct)
        pmdn_pct_str = format_pct(pmdn_pct)
        
        # Determine dominant type
        if pmdn > pma:
            dominant = "PMDN (Penanaman Modal Dalam Negeri)"
            dominant_val = pmdn_formatted
            dominant_pct = pmdn_pct_str
            other = "PMA (Penanaman Modal Asing)"
            other_val = pma_formatted
            other_pct = pma_pct_str
        else:
            dominant = "PMA (Penanaman Modal Asing)"
            dominant_val = pma_formatted
            dominant_pct = pma_pct_str
            other = "PMDN (Penanaman Modal Dalam Negeri)"
            other_val = pmdn_formatted
            other_pct = pmdn_pct_str
        
        text = f"""Berdasarkan status penanaman modal, {dominant} mendominasi dengan {dominant_val} NIB ({dominant_pct}%), sedangkan {other} tercatat sebanyak {other_val} NIB ({other_pct}%).

Dominasi investasi dalam negeri menunjukkan tingginya partisipasi pelaku usaha lokal dalam mengembangkan kegiatan ekonomi di Provinsi Lampung."""
        
        return text
    
    def _generate_pelaku_usaha(self, report: PeriodReport, stats: Dict) -> str:
        """Generate business actor category narrative."""
        pelaku = stats.get('pelaku_usaha_distribution', {})
        umk = pelaku.get('UMK', 0)
        non_umk = pelaku.get('NON_UMK', 0)
        umk_pct = pelaku.get('UMK_pct', 0)
        non_umk_pct = pelaku.get('NON_UMK_pct', 0)
        
        umk_formatted = f"{umk:,}".replace(",", ".")
        non_umk_formatted = f"{non_umk:,}".replace(",", ".")
        
        text = f"""Ditinjau dari kategori pelaku usaha, UMK (Usaha Mikro dan Kecil) menjadi kontributor utama dengan {umk_formatted} NIB ({umk_pct:.1f}%). Sementara itu, Non-UMK (Usaha Menengah dan Besar) tercatat sebanyak {non_umk_formatted} NIB ({non_umk_pct:.1f}%).

Tingginya proporsi UMK menunjukkan bahwa sektor usaha mikro dan kecil memegang peran penting dalam perekonomian Provinsi Lampung. Hal ini sejalan dengan karakteristik ekonomi daerah yang didominasi oleh usaha-usaha skala kecil dan menengah."""
        
        return text
    
    def _generate_kesimpulan(self, report: PeriodReport, stats: Dict) -> str:
        """Generate conclusion paragraph."""
        total = stats.get('total_nib', 0)
        total_formatted = f"{total:,}".replace(",", ".")
        top_5 = stats.get('top_5_locations', [])
        change_pct = stats.get('change_percentage')
        
        top_location = top_5[0]['Kabupaten/Kota'] if top_5 else "N/A"
        
        # Trend conclusion
        trend_conclusion = ""
        if change_pct is not None:
            if change_pct > 0:
                trend_conclusion = f" dengan pertumbuhan positif {change_pct:.1f}% dari periode sebelumnya"
            elif change_pct < 0:
                trend_conclusion = f" dengan penurunan {abs(change_pct):.1f}% dari periode sebelumnya"
        
        text = f"""Berdasarkan data yang telah disajikan, dapat disimpulkan bahwa {self._get_periode_text(report)} mencatat {total_formatted} penerbitan NIB di Provinsi Lampung{trend_conclusion}. 

{top_location} menjadi wilayah dengan aktivitas perizinan tertinggi, sementara investasi didominasi oleh PMDN dengan pelaku usaha mayoritas berasal dari kategori UMK.

DPMPTSP Provinsi Lampung terus berkomitmen untuk meningkatkan pelayanan perizinan guna mendukung iklim investasi yang kondusif dan pertumbuhan ekonomi daerah."""
        
        return text
    
    def _get_periode_text(self, report: PeriodReport) -> str:
        """Get formatted period text."""
        if report.period_type == "Triwulan":
            return f"{report.period_name} Tahun {report.year}"
        elif report.period_type == "Semester":
            return f"{report.period_name} Tahun {report.year}"
        else:
            return f"Tahun {report.year}"
    
    def generate_section(
        self, 
        section_type: str, 
        report: PeriodReport, 
        stats: Dict
    ) -> str:
        """
        Generate a specific narrative section.
        
        Args:
            section_type: Type of section ('pendahuluan', 'nib', 'kab_kota', 'pm', 'pelaku', 'kesimpulan')
            report: PeriodReport object
            stats: Summary statistics
            
        Returns:
            Narrative text for the section
        """
        generators = {
            'pendahuluan': self._generate_pendahuluan,
            'nib': lambda r, s: self._generate_rekapitulasi_nib(r, s),
            'kab_kota': lambda r, s: self._generate_rekapitulasi_kab_kota(r, s),
            'pm': lambda r, s: self._generate_status_pm(r, s),
            'pelaku': lambda r, s: self._generate_pelaku_usaha(r, s),
            'kesimpulan': lambda r, s: self._generate_kesimpulan(r, s),
        }
        
        generator = generators.get(section_type)
        if generator:
            if section_type == 'pendahuluan':
                return generator(report)
            return generator(report, stats)
        
        return ""
