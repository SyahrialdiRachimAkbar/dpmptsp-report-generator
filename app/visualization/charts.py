"""
Chart Generator Module for DPMPTSP Reporting System

This module creates visualizations for NIB data including:
- Bar charts with trendlines
- Horizontal bar charts with gradient colors
- Stacked bar charts for comparisons
- Donut charts for proportions
"""

import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import pandas as pd
from typing import Dict, List, Optional, Tuple
import numpy as np


class ChartGenerator:
    """
    Generates various charts for DPMPTSP reports using Plotly.
    
    Features:
    - Trendlines on bar charts
    - Gradient color bars
    - Data labels
    - Conditional formatting
    - Indonesian language labels
    """
    
    # Color scheme
    COLORS = {
        'primary': '#1e3a5f',
        'secondary': '#3d7ea6',
        'tertiary': '#6db3d5',
        'accent': '#5cb85c',
        'warning': '#f0ad4e',
        'danger': '#d9534f',
        'light': '#e8f4f8',
        'pma': '#2ecc71',
        'pmdn': '#3498db',
        'umk': '#9b59b6',
        'non_umk': '#e74c3c',
    }
    
    # Gradient colors (light to dark)
    GRADIENT = [
        '#e8f4f8', '#cce5ef', '#a8d4e6', '#84c3dd',
        '#60b2d4', '#3d9fc9', '#2d8bb8', '#1e7aa6',
        '#1a6a94', '#165a82', '#124a70', '#1e3a5f'
    ]
    
    def __init__(self, width: int = 800, height: int = 500):
        self.width = width
        self.height = height
        self.layout_defaults = {
            'font': {'family': 'Arial, sans-serif', 'size': 12, 'color': '#e8eaed'},
            'paper_bgcolor': 'rgba(0,0,0,0)',
            'plot_bgcolor': 'rgba(0,0,0,0)',
            'margin': {'l': 50, 'r': 50, 't': 60, 'b': 50},
        }
    
    def _get_gradient_colors(self, n: int) -> List[str]:
        """Generate n colors from the gradient palette."""
        if n <= 0:
            return []
        if n >= len(self.GRADIENT):
            return self.GRADIENT[:n]
        # Sample evenly from gradient
        step = len(self.GRADIENT) / n
        return [self.GRADIENT[int(i * step)] for i in range(n)]
    
    def create_monthly_bar_with_trendline(
        self,
        data: Dict[str, int],
        title: str = "Rekapitulasi NIB per Bulan",
        show_trendline: bool = True
    ) -> go.Figure:
        """
        Create a bar chart with optional trendline showing monthly NIB data.
        
        Args:
            data: Dictionary mapping month names to values
            title: Chart title
            show_trendline: Whether to show a trendline
            
        Returns:
            Plotly Figure object
        """
        months = list(data.keys())
        values = list(data.values())
        
        # Create figure with secondary y-axis for trendline
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        
        # Add bar chart
        fig.add_trace(
            go.Bar(
                x=months,
                y=values,
                name='Total NIB',
                marker_color=self.COLORS['primary'],
                text=values,
                textposition='outside',
                textfont={'size': 11, 'color': self.COLORS['primary']},
            ),
            secondary_y=False
        )
        
        # Add trendline
        if show_trendline and len(values) > 1:
            x_numeric = list(range(len(values)))
            z = np.polyfit(x_numeric, values, 1)
            p = np.poly1d(z)
            trendline_values = [p(x) for x in x_numeric]
            
            fig.add_trace(
                go.Scatter(
                    x=months,
                    y=trendline_values,
                    name='Trendline',
                    mode='lines+markers',
                    line={'color': self.COLORS['danger'], 'width': 2, 'dash': 'dash'},
                    marker={'size': 6},
                ),
                secondary_y=False
            )
        
        # Update layout
        fig.update_layout(
            title={'text': title, 'x': 0.5, 'xanchor': 'center'},
            xaxis_title='Bulan',
            yaxis_title='Jumlah NIB',
            width=self.width,
            height=self.height,
            showlegend=True,
            legend={'x': 0.8, 'y': 1.1, 'orientation': 'h'},
            **self.layout_defaults
        )
        
        # Set y-axis to start from 0
        max_val = max(values) if values else 0
        fig.update_yaxes(range=[0, max_val * 1.2], gridcolor='rgba(150,150,150,0.3)', title_font={'color': '#e8eaed'}, tickfont={'color': '#e8eaed'})
        fig.update_xaxes(tickfont={'color': '#e8eaed'})
        
        return fig
    
    def create_qoq_comparison_bar(
        self,
        current_data: Dict[str, int],
        previous_data: Optional[Dict[str, int]] = None,
        current_label: str = "TW II 2025",
        previous_label: str = "TW I 2025",
        title: str = "Perbandingan Quarter-over-Quarter"
    ) -> go.Figure:
        """
        Create a grouped bar chart comparing two quarters.
        
        Args:
            current_data: Current quarter data
            previous_data: Previous quarter data (optional)
            current_label: Label for current quarter
            previous_label: Label for previous quarter
            title: Chart title
            
        Returns:
            Plotly Figure object
        """
        fig = go.Figure()
        
        # Current quarter
        current_total = sum(current_data.values())
        fig.add_trace(go.Bar(
            x=[current_label],
            y=[current_total],
            name=current_label,
            marker_color=self.COLORS['primary'],
            text=[f"{current_total:,}"],
            textposition='outside',
            textfont={'size': 14, 'color': self.COLORS['primary']},
        ))
        
        # Previous quarter (if available)
        if previous_data:
            previous_total = sum(previous_data.values())
            fig.add_trace(go.Bar(
                x=[previous_label],
                y=[previous_total],
                name=previous_label,
                marker_color=self.COLORS['tertiary'],
                text=[f"{previous_total:,}"],
                textposition='outside',
                textfont={'size': 14, 'color': self.COLORS['tertiary']},
            ))
            
            # Calculate change percentage
            if previous_total > 0:
                change_pct = ((current_total - previous_total) / previous_total) * 100
                change_text = f"{'▲' if change_pct > 0 else '▼'} {abs(change_pct):.1f}%"
                change_color = self.COLORS['accent'] if change_pct > 0 else self.COLORS['danger']
                
                fig.add_annotation(
                    x=0.5,
                    y=max(current_total, previous_total) * 1.15,
                    text=change_text,
                    showarrow=False,
                    font={'size': 18, 'color': change_color, 'family': 'Arial Black'},
                    xref='paper',
                )
        
        fig.update_layout(
            title={'text': title, 'x': 0.5, 'xanchor': 'center'},
            yaxis_title='Jumlah NIB',
            width=self.width,
            height=self.height,
            showlegend=False,
            barmode='group',
            **self.layout_defaults
        )
        
        return fig
    
    def create_horizontal_bar_gradient(
        self,
        df: pd.DataFrame,
        x_col: str = 'Total',
        y_col: str = 'Kabupaten/Kota',
        title: str = "Rekapitulasi NIB per Kabupaten/Kota",
        top_n: int = 15
    ) -> go.Figure:
        """
        Create a horizontal bar chart with gradient colors based on value.
        
        Args:
            df: DataFrame with data
            x_col: Column name for x values
            y_col: Column name for y labels
            title: Chart title
            top_n: Number of top items to show
            
        Returns:
            Plotly Figure object
        """
        # Sort and take top N
        df_sorted = df.nlargest(top_n, x_col)[[y_col, x_col]].copy()
        df_sorted = df_sorted.sort_values(x_col, ascending=True)  # For horizontal bar
        
        # Create gradient colors based on value
        values = df_sorted[x_col].values
        max_val = values.max() if len(values) > 0 else 1
        colors = []
        for val in values:
            # Map value to gradient index
            idx = int((val / max_val) * (len(self.GRADIENT) - 1))
            colors.append(self.GRADIENT[idx])
        
        fig = go.Figure()
        
        fig.add_trace(go.Bar(
            x=df_sorted[x_col],
            y=df_sorted[y_col],
            orientation='h',
            marker_color=colors,
            text=df_sorted[x_col].apply(lambda x: f"{x:,}"),
            textposition='outside',
            textfont={'size': 10},
        ))
        
        fig.update_layout(
            title={'text': title, 'x': 0.5, 'xanchor': 'center'},
            xaxis_title='Jumlah NIB',
            width=self.width,
            height=max(400, top_n * 30),  # Dynamic height
            **self.layout_defaults
        )
        
        fig.update_xaxes(gridcolor='rgba(150,150,150,0.3)', title_font={'color': '#e8eaed'}, tickfont={'color': '#e8eaed'})
        fig.update_yaxes(tickfont={'size': 10, 'color': '#e8eaed'})
        
        return fig
    
    def create_pm_comparison_chart(
        self,
        pma_total: int,
        pmdn_total: int,
        title: str = "Distribusi Status Penanaman Modal"
    ) -> go.Figure:
        """
        Create a donut chart showing PMA vs PMDN distribution.
        
        Args:
            pma_total: Total PMA count
            pmdn_total: Total PMDN count
            title: Chart title
            
        Returns:
            Plotly Figure object
        """
        labels = ['PMA', 'PMDN']
        values = [pma_total, pmdn_total]
        colors = [self.COLORS['pma'], self.COLORS['pmdn']]
        
        # Calculate percentages manually for precise formatting
        total = sum(values)
        percentages = [v/total*100 if total > 0 else 0 for v in values]
        
        # Custom formatting logic
        text_labels = []
        for p in percentages:
            if p < 0.01 and p > 0:
                p_str = "< 0.01%"
            elif p < 1 and p > 0:
                p_str = f"{p:.2f}%" # 2 decimals for < 1% to show detail
            elif p > 99 and p < 100:
                p_str = f"{p:.2f}%" # 2 decimals for > 99% to avoid rounding to 100%
            else:
                p_str = f"{p:.1f}%" # 1 decimal for others
            text_labels.append(p_str)
        
        fig = go.Figure(data=[go.Pie(
            labels=labels,
            values=values,
            hole=0.5,
            marker_colors=colors,
            text=text_labels,
            textinfo='label+text',
            textfont={'size': 14},
            hovertemplate="<b>%{label}</b><br>Jumlah: %{value:,}<br>Persentase: %{text}<extra></extra>"
        )])
        
        # Add center annotation
        total = pma_total + pmdn_total
        fig.add_annotation(
            text=f"<b>Total</b><br>{total:,}",
            showarrow=False,
            font={'size': 16}
        )
        
        fig.update_layout(
            title={'text': title, 'x': 0.5, 'xanchor': 'center'},
            width=self.width,
            height=self.height,
            showlegend=True,
            legend={'x': 0.8, 'y': 0.5},
            **self.layout_defaults
        )
        
        return fig
    
    def create_pelaku_usaha_chart(
        self,
        umk_total: int,
        non_umk_total: int,
        title: str = "Distribusi Pelaku Usaha"
    ) -> go.Figure:
        """
        Create a bar chart comparing UMK vs NON-UMK.
        
        Args:
            umk_total: Total UMK (Usaha Mikro Kecil)
            non_umk_total: Total NON-UMK
            title: Chart title
            
        Returns:
            Plotly Figure object
        """
        labels = ['UMK', 'NON-UMK']
        values = [umk_total, non_umk_total]
        colors = [self.COLORS['umk'], self.COLORS['non_umk']]
        
        # Calculate percentages
        total = umk_total + non_umk_total
        percentages = [v/total*100 if total > 0 else 0 for v in values]
        
        fig = go.Figure()
        
        fig.add_trace(go.Bar(
            x=labels,
            y=values,
            marker_color=colors,
            text=[f"{v:,}<br>({p:.1f}%)" for v, p in zip(values, percentages)],
            textposition='outside',
            textfont={'size': 12},
        ))
        
        fig.update_layout(
            title={'text': title, 'x': 0.5, 'xanchor': 'center'},
            xaxis_title='Kategori Pelaku Usaha',
            yaxis_title='Jumlah NIB',
            width=self.width,
            height=self.height,
            **self.layout_defaults
        )
        
        max_val = max(values) if values else 0
        fig.update_yaxes(range=[0, max_val * 1.3], gridcolor='rgba(150,150,150,0.3)', title_font={'color': '#e8eaed'}, tickfont={'color': '#e8eaed'})
        fig.update_xaxes(tickfont={'color': '#e8eaed'})
        
        return fig
    
    def create_stacked_bar_pm(
        self,
        df: pd.DataFrame,
        title: str = "Status PM per Kabupaten/Kota"
    ) -> go.Figure:
        """
        Create a stacked horizontal bar chart showing PMA vs PMDN by location.
        
        Args:
            df: DataFrame with PMA and PMDN columns
            title: Chart title
            
        Returns:
            Plotly Figure object
        """
        # Sort by total
        df = df.copy()
        df['_total'] = df['PMA'] + df['PMDN']
        df = df.nlargest(15, '_total').sort_values('_total', ascending=True)
        
        fig = go.Figure()
        
        fig.add_trace(go.Bar(
            y=df['Kabupaten/Kota'],
            x=df['PMA'],
            name='PMA',
            orientation='h',
            marker_color=self.COLORS['pma'],
            text=df['PMA'].apply(lambda x: f"{x:,}" if x > 0 else ""),
            textposition='inside',
        ))
        
        fig.add_trace(go.Bar(
            y=df['Kabupaten/Kota'],
            x=df['PMDN'],
            name='PMDN',
            orientation='h',
            marker_color=self.COLORS['pmdn'],
            text=df['PMDN'].apply(lambda x: f"{x:,}" if x > 0 else ""),
            textposition='inside',
        ))
        
        fig.update_layout(
            title={'text': title, 'x': 0.5, 'xanchor': 'center'},
            xaxis_title='Jumlah NIB',
            barmode='stack',
            width=self.width,
            height=max(400, len(df) * 30),
            legend={'x': 0.7, 'y': 1.05, 'orientation': 'h'},
            **self.layout_defaults
        )
        
        return fig
    
    def save_chart(self, fig: go.Figure, filepath: str, format: str = 'png') -> str:
        """
        Save a chart to file.
        
        Args:
            fig: Plotly Figure object
            filepath: Output file path
            format: Output format ('png', 'html', 'svg')
            
        Returns:
            Path to saved file
        """
        if format == 'html':
            fig.write_html(filepath)
        else:
            fig.write_image(filepath, format=format, scale=2)
        
        return filepath
    
    def fig_to_bytes(self, fig: go.Figure, format: str = 'png') -> bytes:
        """
        Convert figure to bytes for embedding in reports.
        
        Args:
            fig: Plotly Figure object
            format: Output format
            
        Returns:
            Image bytes
        """
        return fig.to_image(format=format, scale=2)
    
    def create_risk_distribution_chart(
        self,
        rendah: int,
        menengah_rendah: int,
        menengah_tinggi: int,
        tinggi: int,
        title: str = "Distribusi Perizinan Berdasarkan Tingkat Risiko"
    ) -> go.Figure:
        """
        Create a bar chart showing distribution by risk level.
        
        Args:
            rendah: Low risk count
            menengah_rendah: Medium-low risk count
            menengah_tinggi: Medium-high risk count
            tinggi: High risk count
            title: Chart title
            
        Returns:
            Plotly Figure object
        """
        labels = ['Rendah', 'Menengah Rendah', 'Menengah Tinggi', 'Tinggi']
        values = [rendah, menengah_rendah, menengah_tinggi, tinggi]
        colors = ['#27ae60', '#f1c40f', '#e67e22', '#e74c3c']
        
        total = sum(values)
        percentages = [(v/total*100) if total > 0 else 0 for v in values]
        
        fig = go.Figure()
        
        fig.add_trace(go.Bar(
            x=labels,
            y=values,
            marker_color=colors,
            text=[f"{v:,}<br>({p:.1f}%)" for v, p in zip(values, percentages)],
            textposition='outside',
            textfont={'size': 11},
        ))
        
        fig.update_layout(
            title={'text': title, 'x': 0.5, 'xanchor': 'center'},
            xaxis_title='Tingkat Risiko',
            yaxis_title='Jumlah Perizinan',
            width=self.width,
            height=self.height,
            **self.layout_defaults
        )
        
        max_val = max(values) if values else 0
        fig.update_yaxes(range=[0, max_val * 1.25], gridcolor='rgba(150,150,150,0.3)', title_font={'color': '#e8eaed'}, tickfont={'color': '#e8eaed'})
        fig.update_xaxes(tickfont={'color': '#e8eaed'})
        
        return fig
    
    def create_sector_distribution_chart(
        self,
        sector_data: Dict[str, int],
        title: str = "Distribusi Perizinan Berdasarkan Sektor Usaha"
    ) -> go.Figure:
        """
        Create a horizontal bar chart showing distribution by business sector.
        
        Args:
            sector_data: Dictionary mapping sector names to counts
            title: Chart title
            
        Returns:
            Plotly Figure object
        """
        # Filter out zeros and sort by value ascending (smallest at top)
        filtered_data = {k: v for k, v in sector_data.items() if v > 0}
        sorted_data = dict(sorted(filtered_data.items(), key=lambda x: x[1], reverse=False))
        
        labels = list(sorted_data.keys())
        values = list(sorted_data.values())
        
        # Specific colors for each sector matching the template style
        sector_color_map = {
            'Kelautan': '#4dbfa4',      # Teal/green
            'Perindustrian': '#f4a261', # Orange
            'Pertanian': '#6a9bd1',     # Blue
            'Perhubungan': '#d4a5c9',   # Light purple/pink
            'Kesehatan': '#a8d08d',     # Light green
            'Komunikasi': '#e6c566',    # Yellow/gold
            'Energi': '#d4c17a',        # Khaki/tan
            'Pariwisata': '#c9c9c9',    # Gray
        }
        
        colors = [sector_color_map.get(label, '#999999') for label in labels]
        
        fig = go.Figure()
        
        fig.add_trace(go.Bar(
            x=values,
            y=labels,
            orientation='h',
            marker_color=colors,
            text=[f"{v:,}".replace(',', '.') for v in values],
            textposition='inside',
            textfont={'size': 12, 'color': 'white'},
            insidetextanchor='end',
        ))
        
        fig.update_layout(
            title={'text': title, 'x': 0.5, 'xanchor': 'center', 'font': {'size': 14, 'color': '#e8eaed'}},
            xaxis_title='Jumlah Perizinan',
            width=self.width,
            height=max(350, len(labels) * 45),
            margin={'l': 100, 'r': 50, 't': 60, 'b': 50},
            font={'family': 'Arial, sans-serif', 'size': 12, 'color': '#e8eaed'},
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
        )
        
        fig.update_xaxes(gridcolor='rgba(150,150,150,0.3)', showgrid=True, title_font={'color': '#e8eaed'}, tickfont={'color': '#e8eaed'})
        fig.update_yaxes(tickfont={'size': 11, 'color': '#e8eaed'})
        
        return fig
    
    def create_risk_donut_chart(
        self,
        rendah: int,
        menengah_rendah: int,
        menengah_tinggi: int,
        tinggi: int,
        title: str = "Proporsi Tingkat Risiko"
    ) -> go.Figure:
        """
        Create a donut chart showing risk level proportions.
        """
        labels = ['Rendah', 'Menengah Rendah', 'Menengah Tinggi', 'Tinggi']
        values = [rendah, menengah_rendah, menengah_tinggi, tinggi]
        colors = ['#27ae60', '#f1c40f', '#e67e22', '#e74c3c']
        
        # Calculate percentages manually for precise formatting
        total = sum(values)
        percentages = [v/total*100 if total > 0 else 0 for v in values]
        
        # Custom formatting logic
        text_labels = []
        for p in percentages:
            if p < 0.01 and p > 0:
                p_str = "< 0.01%"
            elif p < 1 and p > 0:
                p_str = f"{p:.2f}%" 
            elif p > 99 and p < 100:
                p_str = f"{p:.2f}%"
            else:
                p_str = f"{p:.1f}%"
            text_labels.append(p_str)
        
        fig = go.Figure(data=[go.Pie(
            labels=labels,
            values=values,
            hole=0.5,
            marker_colors=colors,
            text=text_labels,
            textinfo='label+text',
            textfont={'size': 12},
        )])
        
        total = sum(values)
        fig.add_annotation(
            text=f"<b>Total</b><br>{total:,}",
            showarrow=False,
            font={'size': 14}
        )
        
        fig.update_layout(
            title={'text': title, 'x': 0.5, 'xanchor': 'center'},
            width=self.width,
            height=self.height,
            showlegend=True,
            legend={'x': 0.85, 'y': 0.5},
            **self.layout_defaults
        )
        
        return fig
    
    def create_investment_by_wilayah_chart(
        self,
        data: List,  # List[InvestmentData]
        title: str = "Realisasi Investasi per Wilayah",
        top_n: int = 10
    ) -> go.Figure:
        """
        Create horizontal bar chart showing investment by wilayah.
        
        Args:
            data: List of InvestmentData objects
            title: Chart title
            top_n: Number of top regions to show
            
        Returns:
            Plotly Figure object
        """
        # Sort by value and take top N
        sorted_data = sorted(data, key=lambda x: x.jumlah_rp, reverse=True)[:top_n]
        
        # Prepare data (reverse for correct display order)
        names = [d.name for d in sorted_data][::-1]
        values = [d.jumlah_rp / 1e9 for d in sorted_data][::-1]  # Convert to Billions
        
        # Create gradient colors
        n_bars = len(names)
        colors = self._get_gradient_colors(n_bars)
        
        fig = go.Figure()
        
        fig.add_trace(go.Bar(
            x=values,
            y=names,
            orientation='h',
            marker_color=colors,
            text=[f'Rp {v:,.1f}M' for v in values],
            textposition='outside',
            textfont={'size': 10, 'color': '#e8eaed'},
        ))
        
        fig.update_layout(
            title={'text': title, 'x': 0.5, 'xanchor': 'center', 'font': {'size': 14, 'color': '#e8eaed'}},
            xaxis_title='Nilai Investasi (Miliar Rupiah)',
            width=self.width,
            height=max(350, len(names) * 40),
            **self.layout_defaults
        )
        
        fig.update_xaxes(
            gridcolor='rgba(150,150,150,0.3)',
            title_font={'color': '#e8eaed'},
            tickfont={'color': '#e8eaed'}
        )
        fig.update_yaxes(tickfont={'color': '#e8eaed'})
        
        return fig
    
    def create_pma_pmdn_comparison_chart(
        self,
        pma_total: float,
        pmdn_total: float,
        title: str = "Perbandingan Investasi PMA vs PMDN"
    ) -> go.Figure:
        """
        Create a grouped bar or pie chart comparing PMA and PMDN investments.
        
        Args:
            pma_total: Total PMA investment in Rupiah
            pmdn_total: Total PMDN investment in Rupiah
            title: Chart title
            
        Returns:
            Plotly Figure object
        """
        total = pma_total + pmdn_total
        pma_pct = (pma_total / total * 100) if total > 0 else 0
        pmdn_pct = (pmdn_total / total * 100) if total > 0 else 0
        
        fig = go.Figure()
        
        fig.add_trace(go.Pie(
            labels=['PMA', 'PMDN'],
            values=[pma_total, pmdn_total],
            hole=0.6,
            marker_colors=[self.COLORS['pma'], self.COLORS['pmdn']],
            textinfo='label+percent',
            textposition='outside',
            textfont={'size': 12, 'color': '#e8eaed'},
            hovertemplate='%{label}: Rp %{value:,.0f}<extra></extra>'
        ))
        
        # Add center annotation
        fig.add_annotation(
            x=0.5, y=0.5,
            text=f"<b>Total</b><br>Rp {total/1e12:.2f}T",
            showarrow=False,
            font={'size': 14, 'color': '#e8eaed'}
        )
        
        fig.update_layout(
            title={'text': title, 'x': 0.5, 'xanchor': 'center', 'font': {'size': 14, 'color': '#e8eaed'}},
            width=self.width,
            height=self.height,
            showlegend=True,
            legend={'x': 0.85, 'y': 0.5},
            **self.layout_defaults
        )
        
        return fig
    
    def create_investment_tw_comparison_chart(
        self,
        tw_data: Dict,  # Dict[str, InvestmentReport]
        title: str = "Perbandingan Investasi per Triwulan"
    ) -> go.Figure:
        """
        Create a grouped bar chart comparing investment across Triwulans.
        
        Args:
            tw_data: Dictionary mapping TW name to InvestmentReport
            title: Chart title
            
        Returns:
            Plotly Figure object
        """
        tw_names = []
        pma_values = []
        pmdn_values = []
        
        # Sort by TW order
        for tw in ["TW I", "TW II", "TW III", "TW IV"]:
            if tw in tw_data:
                report = tw_data[tw]
                tw_names.append(tw)
                pma_values.append(report.pma_total / 1e9)  # Convert to Billions
                pmdn_values.append(report.pmdn_total / 1e9)
        
        fig = go.Figure()
        
        fig.add_trace(go.Bar(
            name='PMA',
            x=tw_names,
            y=pma_values,
            marker_color=self.COLORS['pma'],
            text=[f'{v:,.0f}M' for v in pma_values],
            textposition='outside',
            textfont={'size': 10, 'color': '#e8eaed'},
        ))
        
        fig.add_trace(go.Bar(
            name='PMDN',
            x=tw_names,
            y=pmdn_values,
            marker_color=self.COLORS['pmdn'],
            text=[f'{v:,.0f}M' for v in pmdn_values],
            textposition='outside',
            textfont={'size': 10, 'color': '#e8eaed'},
        ))
        
        fig.update_layout(
            title={'text': title, 'x': 0.5, 'xanchor': 'center', 'font': {'size': 14, 'color': '#e8eaed'}},
            xaxis_title='Triwulan',
            yaxis_title='Nilai Investasi (Miliar Rupiah)',
            barmode='group',
            width=self.width,
            height=self.height,
            **self.layout_defaults
        )
        
        fig.update_xaxes(
            tickfont={'color': '#e8eaed'},
            title_font={'color': '#e8eaed'}
        )
        fig.update_yaxes(
            gridcolor='rgba(150,150,150,0.3)',
            tickfont={'color': '#e8eaed'},
            title_font={'color': '#e8eaed'}
        )
        
        return fig
    
    def create_labor_absorption_chart(
        self,
        tki: int,
        tka: int,
        title: str = "Penyerapan Tenaga Kerja"
    ) -> go.Figure:
        """
        Create a chart showing labor absorption (TKI vs TKA).
        
        Args:
            tki: Total domestic workers
            tka: Total foreign workers
            title: Chart title
            
        Returns:
            Plotly Figure object
        """
        fig = go.Figure()
        
        fig.add_trace(go.Bar(
            x=['TKI (Indonesia)', 'TKA (Asing)'],
            y=[tki, tka],
            marker_color=[self.COLORS['accent'], self.COLORS['warning']],
            text=[f'{tki:,}', f'{tka:,}'],
            textposition='outside',
            textfont={'size': 12, 'color': '#e8eaed'},
        ))
        
        fig.update_layout(
            title={'text': title, 'x': 0.5, 'xanchor': 'center', 'font': {'size': 14, 'color': '#e8eaed'}},
            yaxis_title='Jumlah Tenaga Kerja',
            width=self.width,
            height=400,
            **self.layout_defaults
        )
        
        fig.update_yaxes(
            gridcolor='rgba(150,150,150,0.3)',
            tickfont={'color': '#e8eaed'},
            title_font={'color': '#e8eaed'}
        )
        fig.update_xaxes(tickfont={'color': '#e8eaed'})
        
        return fig

