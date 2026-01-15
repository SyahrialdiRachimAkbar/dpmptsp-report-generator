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
    
    def generate_investment_narrative(
        self,
        report,  # PeriodReport
        current_investment,  # InvestmentReport
        tw_summary=None,  # Dict[str, TWSummary]
        prev_year_summary=None  # Dict[str, TWSummary] for Y-o-Y
    ) -> str:
        """
        Generate narrative for investment realization section.
        
        Args:
            report: PeriodReport object
            current_investment: InvestmentReport for current period
            tw_summary: Optional dict of TWSummary for Q-o-Q comparison
            prev_year_summary: Optional dict of TWSummary for Y-o-Y comparison
        
        Returns:
            Narrative text for investment section
        """
        if not current_investment:
            return "Data realisasi investasi belum tersedia."
        
        periode_name = report.period_name
        year = report.year
        
        # Format investment values
        pma = current_investment.pma_total
        pmdn = current_investment.pmdn_total
        total = pma + pmdn
        
        # Convert to readable format (Miliar/Triliun)
        def format_rupiah(val):
            if val >= 1e12:
                return f"Rp {val/1e12:.2f} Triliun"
            elif val >= 1e9:
                return f"Rp {val/1e9:.1f} Miliar"
            else:
                return f"Rp {val/1e6:.1f} Juta"
        
        pma_str = format_rupiah(pma)
        pmdn_str = format_rupiah(pmdn)
        total_str = format_rupiah(total)
        
        # Calculate percentages
        pma_pct = (pma / total * 100) if total > 0 else 0
        pmdn_pct = (pmdn / total * 100) if total > 0 else 0
        
        # Dominant type analysis
        if pmdn > pma:
            dominant = "PMDN (Penanaman Modal Dalam Negeri)"
            dominant_pct = pmdn_pct
            insight = "Hal ini menunjukkan tingginya kepercayaan investor domestik terhadap potensi ekonomi Provinsi Lampung."
        else:
            dominant = "PMA (Penanaman Modal Asing)"
            dominant_pct = pma_pct
            insight = "Hal ini menunjukkan daya tarik Provinsi Lampung bagi investor asing."
        
        # Q-o-Q comparison
        qoq_text = ""
        if tw_summary and periode_name in ["TW II", "TW III", "TW IV"]:
            tw_order = ["TW I", "TW II", "TW III", "TW IV"]
            idx = tw_order.index(periode_name)
            prev_tw = tw_order[idx - 1]
            prev_data = tw_summary.get(prev_tw)
            curr_data = tw_summary.get(periode_name)
            
            if prev_data and curr_data:
                prev_total = prev_data.total_rp
                curr_total = curr_data.total_rp
                if prev_total > 0:
                    change = ((curr_total - prev_total) / prev_total) * 100
                    if change > 0:
                        qoq_text = f"\n\nDibandingkan dengan {prev_tw}, realisasi investasi mengalami peningkatan sebesar {change:.1f}%. "
                    else:
                        qoq_text = f"\n\nDibandingkan dengan {prev_tw}, realisasi investasi mengalami penurunan sebesar {abs(change):.1f}%. "
        
        # Y-o-Y comparison
        yoy_text = ""
        if prev_year_summary and periode_name in prev_year_summary:
            prev_year_data = prev_year_summary.get(periode_name)
            curr_data = tw_summary.get(periode_name) if tw_summary else None
            
            if prev_year_data and curr_data:
                prev_total = prev_year_data.total_rp
                curr_total = curr_data.total_rp
                prev_year = prev_year_data.year
                if prev_total > 0:
                    change = ((curr_total - prev_total) / prev_total) * 100
                    if change > 0:
                        yoy_text = f"Secara tahunan (Y-o-Y), realisasi investasi {periode_name} meningkat {change:.1f}% dibandingkan periode yang sama tahun {prev_year}."
                    else:
                        yoy_text = f"Secara tahunan (Y-o-Y), realisasi investasi {periode_name} menurun {abs(change):.1f}% dibandingkan periode yang sama tahun {prev_year}."
        
        # Labor absorption
        tki = getattr(current_investment, 'total_tki', 0)
        tka = getattr(current_investment, 'total_tka', 0)
        total_labor = tki + tka
        labor_text = ""
        if total_labor > 0:
            tki_formatted = f"{tki:,}".replace(",", ".")
            tka_formatted = f"{tka:,}".replace(",", ".")
            labor_text = f"\n\nDari segi penyerapan tenaga kerja, investasi pada {periode_name} menyerap {tki_formatted} Tenaga Kerja Indonesia (TKI) dan {tka_formatted} Tenaga Kerja Asing (TKA)."
        
        text = f"""Realisasi investasi di Provinsi Lampung pada {periode_name} {year} mencapai {total_str}, terdiri dari PMA sebesar {pma_str} ({pma_pct:.1f}%) dan PMDN sebesar {pmdn_str} ({pmdn_pct:.1f}%).

{dominant} mendominasi dengan kontribusi {dominant_pct:.1f}%. {insight}{qoq_text}{yoy_text}{labor_text}"""
        
        return text
    
    def generate_project_narrative(
        self,
        report,  # PeriodReport
        current_summary,  # TWSummary
        tw_summary=None,  # Dict[str, TWSummary]
        prev_year_summary=None  # Dict[str, TWSummary]
    ) -> str:
        """
        Generate narrative for project realization section (Rencana Proyek).
        
        Args:
            report: PeriodReport object
            current_summary: TWSummary for current period
            tw_summary: Optional dict of TWSummary for Q-o-Q
            prev_year_summary: Optional dict for Y-o-Y
        
        Returns:
            Narrative text for project section
        """
        if not current_summary:
            return "Data rekapitulasi proyek belum tersedia."
        
        periode_name = report.period_name
        year = report.year
        
        # Project count
        total_proyek = current_summary.proyek
        target_pct = current_summary.percentage
        
        proyek_formatted = f"{total_proyek:,}".replace(",", ".")
        
        # Target achievement analysis
        if target_pct >= 100:
            target_insight = f"Pencapaian ini telah melampaui target tahunan ({target_pct:.1f}%)."
        elif target_pct >= 75:
            target_insight = f"Pencapaian sudah mencapai {target_pct:.1f}% dari target tahunan."
        elif target_pct >= 50:
            target_insight = f"Pencapaian baru mencapai {target_pct:.1f}% dari target, perlu akselerasi di periode berikutnya."
        else:
            target_insight = f"Pencapaian masih {target_pct:.1f}% dari target, perlu upaya signifikan untuk mencapai target."
        
        # Q-o-Q comparison
        qoq_text = ""
        if tw_summary and periode_name in ["TW II", "TW III", "TW IV"]:
            tw_order = ["TW I", "TW II", "TW III", "TW IV"]
            idx = tw_order.index(periode_name)
            prev_tw = tw_order[idx - 1]
            prev_data = tw_summary.get(prev_tw)
            
            if prev_data:
                prev_proyek = prev_data.proyek
                if prev_proyek > 0:
                    change = ((total_proyek - prev_proyek) / prev_proyek) * 100
                    prev_formatted = f"{prev_proyek:,}".replace(",", ".")
                    if change > 0:
                        qoq_text = f"\n\nDibandingkan dengan {prev_tw} ({prev_formatted} proyek), jumlah proyek meningkat {change:.1f}%."
                    else:
                        qoq_text = f"\n\nDibandingkan dengan {prev_tw} ({prev_formatted} proyek), jumlah proyek menurun {abs(change):.1f}%."
        
        # Y-o-Y comparison
        yoy_text = ""
        if prev_year_summary and periode_name in prev_year_summary:
            prev_year_data = prev_year_summary.get(periode_name)
            
            if prev_year_data:
                prev_proyek = prev_year_data.proyek
                prev_year = prev_year_data.year
                if prev_proyek > 0:
                    change = ((total_proyek - prev_proyek) / prev_proyek) * 100
                    prev_formatted = f"{prev_proyek:,}".replace(",", ".")
                    if change > 0:
                        yoy_text = f"\n\nSecara tahunan, jumlah proyek {periode_name} {year} meningkat {change:.1f}% dari periode yang sama tahun {prev_year} ({prev_formatted} proyek)."
                    else:
                        yoy_text = f"\n\nSecara tahunan, jumlah proyek {periode_name} {year} menurun {abs(change):.1f}% dari periode yang sama tahun {prev_year} ({prev_formatted} proyek)."
        
        text = f"""Pada {periode_name} {year}, tercatat {proyek_formatted} proyek investasi di Provinsi Lampung. {target_insight}{qoq_text}{yoy_text}

Data ini mencerminkan dinamika investasi di wilayah Lampung dan menjadi indikator penting dalam perencanaan kebijakan investasi ke depan."""
        
        return text
    
    # === Per-Chart Narrative Methods ===
    
    def generate_wilayah_narrative(self, investment_data, investment_type: str = "PMA") -> str:
        """Generate narrative for investment by wilayah chart."""
        if not investment_data:
            return ""
        
        # Sort by value
        sorted_data = sorted(investment_data, key=lambda x: x.jumlah_rp, reverse=True)
        if not sorted_data:
            return ""
        
        top_wilayah = sorted_data[0]
        total = sum(d.jumlah_rp for d in sorted_data)
        top_pct = (top_wilayah.jumlah_rp / total * 100) if total > 0 else 0
        
        # Format value
        val = top_wilayah.jumlah_rp
        if val >= 1e12:
            val_str = f"Rp {val/1e12:.2f} Triliun"
        elif val >= 1e9:
            val_str = f"Rp {val/1e9:.1f} Miliar"
        else:
            val_str = f"Rp {val/1e6:.1f} Juta"
        
        text = f"Investasi {investment_type} tertinggi berada di wilayah {top_wilayah.name} dengan nilai {val_str} ({top_pct:.1f}% dari total)."
        
        # Add second if exists
        if len(sorted_data) > 1:
            second = sorted_data[1]
            second_pct = (second.jumlah_rp / total * 100) if total > 0 else 0
            text += f" Posisi kedua ditempati oleh {second.name} ({second_pct:.1f}%)."
        
        return text
    
    def generate_pma_pmdn_comparison_narrative(self, pma_total: float, pmdn_total: float) -> str:
        """Generate narrative for PMA vs PMDN comparison chart."""
        total = pma_total + pmdn_total
        if total <= 0:
            return ""
        
        pma_pct = (pma_total / total * 100)
        pmdn_pct = (pmdn_total / total * 100)
        
        if pmdn_total > pma_total:
            dominant = "PMDN"
            dominant_pct = pmdn_pct
            ratio = pmdn_total / pma_total if pma_total > 0 else 0
            insight = f"Investasi domestik {ratio:.1f}x lebih besar dari investasi asing, menunjukkan kuatnya partisipasi pengusaha dalam negeri."
        else:
            dominant = "PMA"
            dominant_pct = pma_pct
            ratio = pma_total / pmdn_total if pmdn_total > 0 else 0
            insight = f"Investasi asing {ratio:.1f}x lebih besar dari domestik, menunjukkan daya tarik daerah bagi investor luar negeri."
        
        return f"Berdasarkan proporsi, {dominant} mendominasi dengan {dominant_pct:.1f}% dari total investasi. {insight}"
    
    def generate_labor_narrative(self, tki: int, tka: int) -> str:
        """Generate narrative for labor absorption chart."""
        total = tki + tka
        if total <= 0:
            return ""
        
        tki_pct = (tki / total * 100)
        tka_pct = (tka / total * 100)
        
        tki_formatted = f"{tki:,}".replace(",", ".")
        tka_formatted = f"{tka:,}".replace(",", ".")
        
        text = f"Penyerapan tenaga kerja mencapai {tki_formatted} TKI ({tki_pct:.1f}%) dan {tka_formatted} TKA ({tka_pct:.1f}%)."
        
        if tki > tka * 10:
            text += " Dominasi TKI menunjukkan investasi berhasil membuka lapangan kerja bagi tenaga lokal."
        elif tka > tki:
            text += " Tingginya proporsi TKA mengindikasikan kebutuhan tenaga ahli dari luar negeri."
        
        return text
    
    def generate_tw_comparison_narrative(self, investment_reports: dict) -> str:
        """Generate narrative for TW comparison chart."""
        if not investment_reports or len(investment_reports) < 2:
            return ""
        
        # Get ordered TW data
        tw_order = ["TW I", "TW II", "TW III", "TW IV"]
        data = []
        for tw in tw_order:
            if tw in investment_reports:
                report = investment_reports[tw]
                total = report.pma_total + report.pmdn_total
                data.append((tw, total))
        
        if len(data) < 2:
            return ""
        
        # Find trend
        first_val = data[0][1]
        last_val = data[-1][1]
        
        if last_val > first_val:
            trend = "tren peningkatan"
            change = ((last_val - first_val) / first_val * 100) if first_val > 0 else 0
            insight = f"meningkat {change:.1f}% dari {data[0][0]} ke {data[-1][0]}"
        elif last_val < first_val:
            trend = "tren penurunan"
            change = ((first_val - last_val) / first_val * 100) if first_val > 0 else 0
            insight = f"menurun {change:.1f}% dari {data[0][0]} ke {data[-1][0]}"
        else:
            trend = "stabil"
            insight = "relatif stabil sepanjang periode"
        
        # Find peak
        peak_tw, peak_val = max(data, key=lambda x: x[1])
        
        return f"Perbandingan antar Triwulan menunjukkan {trend}, {insight}. Investasi tertinggi tercatat pada {peak_tw}."
    
    def generate_qoq_narrative(self, current_tw: str, current_proyek: int, prev_tw: str, prev_proyek: int) -> str:
        """Generate narrative for Q-o-Q comparison chart."""
        if prev_proyek <= 0:
            return ""
        
        change = ((current_proyek - prev_proyek) / prev_proyek * 100)
        
        curr_formatted = f"{current_proyek:,}".replace(",", ".")
        prev_formatted = f"{prev_proyek:,}".replace(",", ".")
        
        if change > 0:
            trend = "peningkatan"
            insight = "menunjukkan pertumbuhan aktivitas investasi"
        else:
            trend = "penurunan"
            insight = "perlu evaluasi faktor-faktor yang mempengaruhi"
        
        return f"Secara Q-o-Q, jumlah proyek mengalami {trend} {abs(change):.1f}% dari {prev_tw} ({prev_formatted}) ke {current_tw} ({curr_formatted}). Hal ini {insight}."
    
    def generate_yoy_narrative(self, tw_name: str, current_year: int, current_proyek: int, 
                                prev_year: int, prev_proyek: int) -> str:
        """Generate narrative for Y-o-Y comparison chart."""
        if prev_proyek <= 0:
            return ""
        
        change = ((current_proyek - prev_proyek) / prev_proyek * 100)
        
        curr_formatted = f"{current_proyek:,}".replace(",", ".")
        prev_formatted = f"{prev_proyek:,}".replace(",", ".")
        
        if change > 0:
            trend = "pertumbuhan"
            insight = "menunjukkan perbaikan iklim investasi dari tahun ke tahun"
        else:
            trend = "penurunan"
            insight = "perlu strategi untuk meningkatkan daya tarik investasi"
        
        return f"Perbandingan Y-o-Y menunjukkan {trend} {abs(change):.1f}% untuk {tw_name} ({prev_year}: {prev_formatted} vs {current_year}: {curr_formatted}). {insight.capitalize()}."

