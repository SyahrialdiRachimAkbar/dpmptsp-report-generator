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
    
    # New fields for comprehensive export
    investasi_wilayah: str = ""
    investasi_tenaga_kerja: str = ""
    investasi_sektor: str = ""
    pb_jenis: str = ""
    pb_status_respon: str = ""
    pb_kewenangan: str = ""
    pb_sektor: str = ""
    pb_risiko: str = ""
    pb_kab_kota_narrative: str = ""  # Explicit field if needed beyond rekapitulasi_kab_kota


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
        """Generate NIB summary narrative with enhanced insights."""
        total = stats.get('total_nib', 0)
        monthly = stats.get('monthly_totals', {})
        change_pct = stats.get('change_percentage')
        prev_total = stats.get('prev_period_total')
        
        # Format total with thousands separator
        total_formatted = f"{total:,}".replace(",", ".")
        
        # Monthly breakdown with percentages and insights
        monthly_text = ""
        peak_month = ""
        peak_value = 0
        monthly_growth_insight = ""
        
        if monthly:
            monthly_parts = []
            month_values = list(monthly.items())
            
            # Identify peak month
            for bulan, nilai in month_values:
                if nilai > peak_value:
                    peak_value = nilai
                    peak_month = bulan
            
            # Build monthly breakdown with percentages
            for bulan, nilai in month_values:
                nilai_formatted = f"{nilai:,}".replace(",", ".")
                pct_of_total = (nilai / total * 100) if total > 0 else 0
                monthly_parts.append(f"{bulan} ({nilai_formatted} NIB, {pct_of_total:.1f}%)")
            
            monthly_text = ", ".join(monthly_parts)
            
            # Calculate month-over-month growth for multi-month periods
            if len(month_values) >= 2:
                first_month_val = month_values[0][1]
                last_month_val = month_values[-1][1]
                
                if first_month_val > 0:
                    mom_growth = ((last_month_val - first_month_val) / first_month_val) * 100
                    
                    if mom_growth > 10:
                        monthly_growth_insight = f" Terdapat akselerasi positif dengan pertumbuhan {mom_growth:.1f}% dari awal ke akhir periode, dengan {peak_month} mencatat kinerja tertinggi."
                    elif mom_growth > 0:
                        monthly_growth_insight = f" Pertumbuhan moderat sebesar {mom_growth:.1f}% terlihat dari bulan pertama ke bulan terakhir periode ini."
                    elif mom_growth < -10:
                        monthly_growth_insight = f" Teridentifikasi penurunan {abs(mom_growth):.1f}% dari awal ke akhir periode, yang memerlukan perhatian khusus."
                    elif mom_growth < 0:
                        monthly_growth_insight = f" Penurunan ringan {abs(mom_growth):.1f}% tercatat dari bulan pertama ke akhir periode."
                    else:
                        monthly_growth_insight = f" Konsistensi performa terlihat sepanjang periode dengan fluktuasi minimal."
        
        # Enhanced trend analysis with actionable insights
        trend_text = ""
        if change_pct is not None and prev_total is not None:
            prev_formatted = f"{prev_total:,}".replace(",", ".")
            abs_change = total - prev_total
            abs_change_formatted = f"{abs(abs_change):,}".replace(",", ".")
            
            if change_pct > 15:
                trend_text = f"\n\nKinerja periode ini menunjukkan pertumbuhan signifikan sebesar {change_pct:.1f}% ({abs_change_formatted} NIB) dibanding periode sebelumnya ({prev_formatted} NIB). Momentum positif ini mengindikasikan meningkatnya minat investasi dan aktivitas ekonomi di Provinsi Lampung, yang perlu dipertahankan melalui kebijakan yang kondusif."
            elif change_pct > 5:
                trend_text = f"\n\nTercatat peningkatan moderat sebesar {change_pct:.1f}% ({abs_change_formatted} NIB) dibandingkan periode sebelumnya ({prev_formatted} NIB). Pertumbuhan stabil ini menunjukkan iklim investasi yang kondusif."
            elif change_pct > 0:
                trend_text = f"\n\nPertumbuhan ringan sebesar {change_pct:.1f}% ({abs_change_formatted} NIB) dari periode sebelumnya ({prev_formatted} NIB) mengindikasikan stabilitas dengan potensi peningkatan lebih lanjut."
            elif change_pct > -5:
                trend_text = f"\n\nPenurunan minor sebesar {abs(change_pct):.1f}% ({abs_change_formatted} NIB) dari periode sebelumnya ({prev_formatted} NIB). Fluktuasi ini masih dalam batas normal dan perlu dipantau."
            elif change_pct > -15:
                trend_text = f"\n\nTerjadi penurunan sebesar {abs(change_pct):.1f}% ({abs_change_formatted} NIB) dibandingkan periode sebelumnya ({prev_formatted} NIB). Evaluasi mendalam diperlukan untuk mengidentifikasi faktor penyebab dan strategi perbaikan."
            else:
                trend_text = f"\n\nPenurunan signifikan {abs(change_pct):.1f}% ({abs_change_formatted} NIB) dari periode sebelumnya ({prev_formatted} NIB) memerlukan perhatian serius. Rekomendasi: analisis komprehensif terhadap hambatan investasi dan revisi strategi promosi."
        
        # Build final narrative with enhanced structure
        text = f"""Pada {self._get_periode_text(report)}, total NIB yang diterbitkan di Provinsi Lampung mencapai {total_formatted} NIB. Rincian distribusi per bulan: {monthly_text}.{monthly_growth_insight}{trend_text}"""
        
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
    
    def generate_status_pm_narrative(self, pma_total: float, pmdn_total: float, unit_type: str = "investasi") -> str:
        """
        Generate narrative for PMA vs PMDN comparison chart.
        
        Args:
            pma_total: Total PMA value/count
            pmdn_total: Total PMDN value/count
            unit_type: "investasi" (default) or "proyek"
        """
        total = pma_total + pmdn_total
        if total <= 0:
            return ""
        
        pma_pct = (pma_total / total * 100)
        pmdn_pct = (pmdn_total / total * 100)
        
        # Context-aware text
        if unit_type == "proyek":
            context_noun = "proyek"
            partisipasi_msg = "partisipasi"
            attraction_msg = "minat investor"
        else:
            context_noun = "investasi"
            partisipasi_msg = "partisipasi pengusaha"
            attraction_msg = "daya tarik daerah bagi investor"

        if pmdn_total > pma_total:
            dominant = "PMDN"
            dominant_pct = pmdn_pct
            ratio = pmdn_total / pma_total if pma_total > 0 else 0
            insight = f"{context_noun.capitalize()} domestik {ratio:.1f}x lebih besar dari asing, menunjukkan kuatnya {partisipasi_msg} dalam negeri."
        else:
            dominant = "PMA"
            dominant_pct = pma_pct
            ratio = pma_total / pmdn_total if pmdn_total > 0 else 0
            insight = f"{context_noun.capitalize()} asing {ratio:.1f}x lebih besar dari domestik, menunjukkan tingginya {attraction_msg} luar negeri."
        
        return f"Berdasarkan proporsi, {dominant} mendominasi dengan {dominant_pct:.1f}% dari total {context_noun}. {insight}"
    
    def generate_labor_narrative(self, tki_total: int, tka_total: int) -> str:
        """Generate narrative for labor absorption."""
        total = tki_total + tka_total
        
        if total == 0:
            return "Belum ada data penyerapan tenaga kerja yang tercatat pada periode ini."
            
        tki_pct = (tki_total / total) * 100
        tka_pct = (tka_total / total) * 100
        
        text = f"""
        Total penyerapan tenaga kerja pada periode ini mencapai {total:,} orang.
        Dari jumlah tersebut, sebanyak {tki_total:,} orang ({tki_pct:.1f}%) merupakan Tenaga Kerja Indonesia (TKI),
        sedangkan {tka_total:,} orang ({tka_pct:.1f}%) merupakan Tenaga Kerja Asing (TKA).
        """.replace(",", ".")
        
        return text

    def generate_skala_usaha_comparison_narrative(
        self,
        current_data: Dict[str, int],
        prev_year_data: Dict[str, int],
        prev_q_data: Dict[str, int],
        period_name: str,
        year: int
    ) -> str:
        """
        Generate narrative for Skala Usaha distribution and comparison.
        Matches the style in the reference image.
        """
        total_proyek = sum(current_data.values())
        
        text = f"Rekapitulasi jumlah proyek di provinsi lampung periode {period_name} tahun {year} berdasarkan skala usaha berjumlah <b>{total_proyek:,.0f}</b>.".replace(",", ".")
        
        # Detail breakdown
        details = []
        # Standardize keys
        std_keys = ['MIKRO', 'KECIL', 'MENENGAH', 'BESAR']
        
        # Helper to find key
        def find_val(data, key_part):
            for k, v in data.items():
                if key_part in str(k).upper():
                    return v
            return 0
            
        for key in std_keys:
            count = find_val(current_data, key)
            if count > 0:
                details.append(f"yang berstatus tingkat risiko <b>USAHA {key}</b> berjumlah <b>{count:,.0f}</b> proyek".replace(",", "."))
        
        if details:
            text += ", " + ", ".join(details) + "."
            
        # Comparison Y-o-Y
        text += f" Jika dibandingkan dengan tahun sebelumnya ({period_name} tahun {year-1}), "
        yoy_details = []
        for key in std_keys:
            curr = find_val(current_data, key)
            prev = find_val(prev_year_data, key)
            
            if curr > 0 or prev > 0:
                if prev == 0:
                    growth = 100.0 if curr > 0 else 0
                    trend = "peningkatan"
                else:
                    growth = ((curr - prev) / prev) * 100
                    trend = "peningkatan" if growth >= 0 else "penurunan"
                
                yoy_details.append(f"yang berstatus tingkat risiko <b>USAHA {key}</b> mengalami {trend} sebesar <b>{abs(growth):.2f}%</b>")
        
        if yoy_details:
             text += ", ".join(yoy_details) + "."
             
        # Comparison Q-o-Q
        # (Optional to add logic for prev quarter name, simpler to just list stats)
        # For brevity matching the image style which just flows.
        
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

    def generate_pb_oss_narrative(
        self,
        report,  # PeriodReport
        total_permits: int,
        monthly_permits: Dict[str, int],
        location_data: Dict[str, int],
        prev_year_total: int,
        prev_q_total: int,
        prev_q_label: str
    ) -> str:
        """
        Generate summary narrative for Section 3.1 (PB OSS).
        Summarizes: Monthly Trend, Top Location, Y-o-Y, Q-o-Q.
        """
        if total_permits <= 0:
            return "Data perizinan belum tersedia."

        period_text = self._get_periode_text(report)
        total_formatted = f"{total_permits:,}".replace(",", ".")

        # 1. Monthly Peak
        peak_month = ""
        peak_val = 0
        if monthly_permits:
            peak_month = max(monthly_permits, key=monthly_permits.get)
            peak_val = monthly_permits[peak_month]
        
        peak_text = ""
        if peak_month:
            peak_val_fmt = f"{peak_val:,}".replace(",", ".")
            peak_text = f" Aktivitas tertinggi tercatat pada bulan {peak_month} dengan {peak_val_fmt} perizinan."

        # 2. Top Location (All locations considered)
        loc_text = ""
        if location_data:
            top_loc = max(location_data, key=location_data.get)
            top_loc_val = location_data[top_loc]
            top_loc_pct = (top_loc_val / total_permits * 100)
            top_loc_fmt = f"{top_loc_val:,}".replace(",", ".")
            loc_text = f" Lokasi usaha didominasi oleh {top_loc} dengan {top_loc_fmt} perizinan ({top_loc_pct:.1f}%)."

        # 3. Y-o-Y Comparison
        yoy_text = ""
        if prev_year_total > 0:
            change = ((total_permits - prev_year_total) / prev_year_total) * 100
            trend = "meningkat" if change >= 0 else "menurun"
            yoy_text = f" Secara tahunan (Y-o-Y), terjadi {trend} sebesar {abs(change):.1f}% dibandingkan tahun sebelumnya."

        # 4. Q-o-Q Comparison
        qoq_text = ""
        if prev_q_total > 0:
            change = ((total_permits - prev_q_total) / prev_q_total) * 100
            trend = "peningkatan" if change >= 0 else "penurunan"
            qoq_text = f" Dibandingkan dengan {prev_q_label}, terjadi {trend} sebesar {abs(change):.1f}%."

        # Combine
        text = f"""Sepanjang {period_text}, tercatat {total_formatted} perizinan berusaha lintas sektor di Provinsi Lampung (Kewenangan Gubernur).{peak_text}{loc_text}{yoy_text}{qoq_text}"""
        
        return text

    def generate_status_pm_comparison_narrative(
        self,
        report,  # PeriodReport
        curr_pma: int,
        curr_pmdn: int,
        prev_year_pma: int,
        prev_year_pmdn: int,
        prev_q_pma: int,
        prev_q_pmdn: int,
        prev_q_label: str,
        monthly_breakdown: Optional[Dict[str, Dict[str, int]]] = None
    ) -> str:
        """
        Generate summary narrative for Section 3.2 (Status PM).
        Summarizes: Monthly Peak, Dominance, YoY Growth, QoQ Growth.
        """
        total = curr_pma + curr_pmdn
        if total <= 0:
            return "Data penanaman modal belum tersedia."

        period_text = self._get_periode_text(report)

        # 1. Monthly Peak Analysis
        peak_text = ""
        if monthly_breakdown:
            # Find month with max Total (PMA + PMDN)
            month_totals = {
                m: (d.get('PMA', 0) + d.get('PMDN', 0)) 
                for m, d in monthly_breakdown.items()
            }
            if month_totals:
                peak_month = max(month_totals, key=month_totals.get)
                peak_val = month_totals[peak_month]
                peak_pma = monthly_breakdown[peak_month].get('PMA', 0)
                peak_pmdn = monthly_breakdown[peak_month].get('PMDN', 0)
                
                if peak_val > 0:
                    peak_text = (f"Aktivitas perizinan tertinggi tercatat pada bulan {peak_month} "
                                 f"dengan total {peak_val} perizinan ({peak_pma} PMA, {peak_pmdn} PMDN). ")
        
        # 2. Dominance
        if curr_pmdn > curr_pma:
            dom = "PMDN"
            val = curr_pmdn
            pct = (curr_pmdn / total * 100)
        else:
            dom = "PMA"
            val = curr_pma
            pct = (curr_pma / total * 100)
        
        val_fmt = f"{val:,}".replace(",", ".")
        dom_text = f"Secara keseluruhan pada {period_text}, didominasi oleh {dom} dengan {val_fmt} perizinan ({pct:.1f}%)."

        # 3. Y-o-Y Analysis
        yoy_text = ""
        if prev_year_pma > 0 or prev_year_pmdn > 0:
            # Change for PMA
            pma_chg = 0
            if prev_year_pma > 0:
                pma_chg = ((curr_pma - prev_year_pma) / prev_year_pma) * 100
                pma_trend = "meningkat" if pma_chg >= 0 else "menurun"
                pma_str = f"PMA {pma_trend} {abs(pma_chg):.1f}%"
            else:
                pma_str = "PMA baru tercatat" if curr_pma > 0 else "PMA tetap nihil"

            # Change for PMDN
            pmdn_chg = 0
            if prev_year_pmdn > 0:
                pmdn_chg = ((curr_pmdn - prev_year_pmdn) / prev_year_pmdn) * 100
                pmdn_trend = "meningkat" if pmdn_chg >= 0 else "menurun"
                pmdn_str = f"PMDN {pmdn_trend} {abs(pmdn_chg):.1f}%"
            else:
                pmdn_str = "PMDN baru tercatat" if curr_pmdn > 0 else "PMDN tetap nihil"
            
            yoy_text = f" Secara tahunan (Y-o-Y), {pma_str} dan {pmdn_str} dibandingkan periode yang sama tahun sebelumnya."

        # 4. Q-o-Q Analysis
        qoq_text = ""
        if (prev_q_pma > 0 or prev_q_pmdn > 0) and prev_q_label:
            # Change for PMA
            pma_chg = 0
            if prev_q_pma > 0:
                pma_chg = ((curr_pma - prev_q_pma) / prev_q_pma) * 100
                pma_trend = "naik" if pma_chg >= 0 else "turun"
                pma_str = f"PMA {pma_trend} {abs(pma_chg):.1f}%"
            else:
                pma_str = ""

            # Change for PMDN
            pmdn_chg = 0
            if prev_q_pmdn > 0:
                pmdn_chg = ((curr_pmdn - prev_q_pmdn) / prev_q_pmdn) * 100
                pmdn_trend = "naik" if pmdn_chg >= 0 else "turun"
                pmdn_str = f"PMDN {pmdn_trend} {abs(pmdn_chg):.1f}%"
            else:
                pmdn_str = ""
            
            detail_list = [s for s in [pma_str, pmdn_str] if s]
            if detail_list:
                qoq_text = f" Dibandingkan dengan {prev_q_label}, {' dan '.join(detail_list)}."

        return f"{peak_text}{dom_text}{yoy_text}{qoq_text}"

    def generate_risk_comparison_narrative(
        self,
        report,
        current_data: Dict[str, int],
        prev_year_data: Dict[str, int],
        prev_q_data: Dict[str, int],
        prev_q_label: str
    ) -> str:
        """
        Generate summary narrative for Section 3.3 (Risk Levels).
        Summarizes: Dominant Risk Level, YoY for Dominant, QoQ for Dominant.
        """
        if not current_data:
            return "Data tingkat risiko belum tersedia."

        period_text = self._get_periode_text(report)
        total = sum(current_data.values())
        
        # 1. Dominance
        dom_risk = max(current_data, key=current_data.get)
        dom_val = current_data[dom_risk]
        dom_pct = (dom_val / total * 100) if total > 0 else 0
        
        dom_formatted = f"{dom_val:,}".replace(",", ".")
        dom_text = f"Pada {period_text}, perizinan berusaha didominasi oleh tingkat risiko {dom_risk} dengan {dom_formatted} perizinan ({dom_pct:.1f}%)."

        # 2. Comparison for Dominant Risk
        yoy_text = ""
        qoq_text = ""
        
        # YoY
        prev_y_val = prev_year_data.get(dom_risk, 0)
        if prev_y_val > 0:
            chg = ((dom_val - prev_y_val) / prev_y_val) * 100
            trend = "naik" if chg >= 0 else "turun"
            yoy_text = f" Secara tahunan (Y-o-Y), kategori ini {trend} {abs(chg):.1f}% dibandingkan tahun sebelumnya."
        
        # QoQ
        prev_q_val = prev_q_data.get(dom_risk, 0)
        if prev_q_val > 0 and prev_q_label:
            chg = ((dom_val - prev_q_val) / prev_q_val) * 100
            trend = "meningkat" if chg >= 0 else "menurun"
            qoq_text = f" Dibandingkan dengan {prev_q_label} (Q-o-Q), tercatat {trend} sebesar {abs(chg):.1f}%."

        return f"{dom_text}{yoy_text}{qoq_text}"

