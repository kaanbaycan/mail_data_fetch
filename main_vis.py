import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("TkAgg") # Force TkAgg backend for macOS compatibility
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import matplotlib.patheffects as pe
from datetime import datetime, timedelta
import os
import glob
from outlook_bot import check_and_download_specific_mails
from veri_isleyici import process_and_merge_files

# -----------------------------------------------------------------------------
# 1. CONSTANTS & CONFIGURATION
# -----------------------------------------------------------------------------
DATA_FILENAME_PATTERN = "jet_brent*.csv"
PERSISTENT_FILE = "jet_brent_current.csv"

AUTHORITY_FORECASTS = [
    {"inst": "Goldman Sachs", "val": 90.00, "disp": "$90.00", "date": "26 Nisan", "note": "Hürmüz normalleşmesi"},
    {"inst": "EIA", "val": 88.00, "disp": "$88.00", "date": "7 Nisan", "note": "2026 sonu toparlanma"},
    {"inst": "UBS", "val": 90.00, "disp": "$90.00", "date": "13 Nisan", "note": "Fiziksel-vadeli kopukluğu"},
    {"inst": "ANZ", "val": 92.00, "disp": "$92.00", "date": "9 Nisan", "note": "OPEC+ kısıtlı kapasite"},
    {"inst": "Barclays", "val": 85.00, "disp": "$85.00", "date": "13 Mart", "note": "Hürmüz 2-3 haftada düzelir"},
    {"inst": "StanChart", "val": 80.50, "disp": "$80.50", "date": "19 Mart", "note": "Envanter yeniden yapılanma"},
    {"inst": "Citi", "val": 80.00, "disp": "$80.00", "date": "26 Nisan", "note": "%50 olasılık baz senaryo"},
    {"inst": "J.P. Morgan", "val": 80.00, "disp": "$80.00", "date": "20 Mart", "note": "Kesintiler fiyatı korur"},
    {"inst": "Morgan Stanley", "val": 80.00, "disp": ">$80.00", "date": "24 Mart", "note": "2026 boyu yüksek kalır"},
    {"inst": "BMI (Fitch)", "val": 69.00, "disp": "$69.00", "date": "12 Mart", "note": "Talep düşüşü etkili olur"},
]

THY_RED = "#C70A0C"       # Official Brand Red
THY_RED_HOVER = "#E53935" # Lighter Red for Hover
THY_NAVY = "#00204E"      # Deep Professional Navy
THY_NAVY_HOVER = "#003366" # Lighter Navy for Hover
THY_WHITE = "#FFFFFF"
THY_BG_LIGHT = "#F8F9FA"  # Clean background
THY_SILVER = "#A7A9AC"    # Secondary Silver
THY_TEXT_DARK = "#1F2937" # High readability dark text

# -----------------------------------------------------------------------------
# 2. CORE LOGIC & HELPERS
# -----------------------------------------------------------------------------
def get_delta(current, previous):
    if previous == 0 or pd.isna(previous): return None
    pct = ((current - previous) / previous) * 100
    return f"{pct:+.1f}%"

def find_last_time_similar(df, current_val, col, threshold=0.03):
    """Finds the most recent date (before the last 7 days) with a similar value."""
    history = df.iloc[:-7]
    lower = current_val * (1 - threshold)
    upper = current_val * (1 + threshold)
    similar = history[(history[col] >= lower) & (history[col] <= upper)]
    if not similar.empty:
        return similar.iloc[-1]['tarih']
    return None

def analyze_and_plot(df, params, figure):
    figure.clear()
    plt.style.use('seaborn-v0_8-white')
    
    # Extract Params
    window = int(params['window'])
    point_size = float(params['point_size']) 
    line_width = float(params['line_width']) 
    period = params['period']
    
    shifts = {'brent': float(params['brent_mult']), 'cif': float(params['cif_shift'])}
    thresholds = {'brent': float(params['th_brent']), 'cif': float(params['th_cif']), 'katsayi': float(params['th_kat'])}
    colors = {'brent': params['c_brent'], 'cif': params['c_cif'], 'katsayi': params['c_kat']}
    
    temp_df = df.copy()
    temp_df['tarih'] = pd.to_datetime(temp_df['tarih'], errors='coerce')
    temp_df = temp_df.dropna(subset=['tarih']).sort_values('tarih')

    if params.get('show_ma_brent'):
        ma_win = int(params.get('ma_brent_win', 5))
        if ma_win > 0:
            temp_df['brent_ma'] = temp_df['brent'].rolling(window=ma_win, min_periods=1).mean()
    
    if params.get('show_ma_cif'):
        ma_win = int(params.get('ma_cif_win', 5))
        if ma_win > 0:
            temp_df['cif_ma'] = temp_df['cif med'].rolling(window=ma_win, min_periods=1).mean()

    start_date = pd.to_datetime(params['start_date'])
    end_date = pd.to_datetime(params['end_date'])
    plot_df = temp_df[(temp_df['tarih'] >= start_date) & (temp_df['tarih'] <= end_date)].copy()
    
    if plot_df.empty: return

    if period == 'Haftalık':
        plot_df = plot_df.resample('W', on='tarih').mean().reset_index()
    elif period == 'Aylık':
        plot_df = plot_df.resample('M', on='tarih').mean().reset_index()
    elif period == 'Yıllık':
        plot_df = plot_df.resample('A', on='tarih').mean().reset_index()
    
    plot_df = plot_df.sort_values('tarih').reset_index(drop=True)
    plot_df['x_idx'] = range(len(plot_df))
    
    plot_df['brent_katsayisi'] = plot_df['cif med'] / plot_df['brent']
    plot_df['brent_visual'] = plot_df['brent'] * shifts['brent']
    plot_df['cif_visual'] = plot_df['cif med'] + shifts['cif']
    
    if 'brent_ma' in plot_df.columns:
        plot_df['brent_ma_visual'] = plot_df['brent_ma'] * shifts['brent']
    if 'cif_ma' in plot_df.columns:
        plot_df['cif_ma_visual'] = plot_df['cif_ma'] + shifts['cif']

    for col in ['brent', 'cif med', 'brent_katsayisi']:
        plot_df[f'{col}_degisim'] = plot_df[col].pct_change().abs()

    ax1 = figure.add_subplot(111, facecolor='#ffffff')
    ax2 = ax1.twinx()
    
    ax1.set_yticks([]); ax2.set_yticks([])
    for ax in [ax1, ax2]:
        for s in ax.spines.values(): s.set_visible(False)
    ax1.grid(True, axis='x', linestyle=':', alpha=0.6, color='#D1D5DB')
            
    configs = []
    if params.get('show_cif', True):
        configs.append({'col': 'cif med', 'vis_col': 'cif_visual', 'ax': ax1, 'color': colors['cif'], 
                        'label': 'Jet Yakıt (Ton/$)', 'fmt': '{:,.0f}', 'thresh': thresholds['cif'], 'type': 'line', 'zorder': 3})
    if params.get('show_brent', True):
        configs.append({'col': 'brent', 'vis_col': 'brent_visual', 'ax': ax1, 'color': colors['brent'], 
                        'label': 'Brent', 'fmt': '{:.0f}', 'thresh': thresholds['brent'], 'type': 'line', 'zorder': 2})
    if params.get('show_kat', True):
        configs.append({'col': 'brent_katsayisi', 'vis_col': 'brent_katsayisi', 'ax': ax2, 'color': colors['katsayi'], 
                        'label': 'Katsayı', 'fmt': '{:.1f}', 'thresh': thresholds['katsayi'], 'type': 'scatter', 'zorder': 4})

    plotted_handles = []

    def get_peaks_and_troughs(series, win):
        if len(series) < win*2+1: return pd.Series(False, index=series.index), pd.Series(False, index=series.index)
        is_peak = series.rolling(window=win*2+1, center=True).max() == series
        is_trough = series.rolling(window=win*2+1, center=True).min() == series
        return is_peak, is_trough

    cutoff_date = pd.Timestamp('2026-01-01')
    last_year_mask = plot_df['tarih'] >= cutoff_date
    last_year_indices = plot_df.index[last_year_mask].tolist()
    split_idx = last_year_indices[0] if last_year_indices else len(plot_df)

    for cfg in configs:
        if cfg['type'] == 'line':
            x = plot_df['x_idx']
            y = plot_df[cfg['vis_col']]
            cfg['ax'].plot(x, y, color=cfg['color'], 
                           linewidth=line_width, alpha=0.2, zorder=cfg['zorder']-1, 
                           path_effects=[pe.SimpleLineShadow(offset=(0, -2), alpha=0.1), pe.Normal()])
            handle, = cfg['ax'].plot([], [], label=cfg['label'], color=cfg['color'], linewidth=line_width)
            use_dotted = params.get('dotted_2026', True)
            if split_idx > 0:
                cfg['ax'].plot(x.iloc[:split_idx], y.iloc[:split_idx],
                               color=cfg['color'], linewidth=line_width, alpha=1.0, zorder=cfg['zorder'])
                if split_idx < len(plot_df) + 1:
                    ls = ':' if use_dotted else '-'
                    cfg['ax'].plot(x.iloc[split_idx-1:], y.iloc[split_idx-1:], 
                                   color=cfg['color'], linewidth=line_width, alpha=1.0, 
                                   zorder=cfg['zorder'], linestyle=ls)
            else:
                ls = ':' if use_dotted else '-'
                cfg['ax'].plot(x, y, color=cfg['color'], linewidth=line_width, alpha=1.0, 
                               zorder=cfg['zorder'], linestyle=ls)
        else:
            handle = cfg['ax'].scatter(plot_df['x_idx'], plot_df[cfg['vis_col']], label=cfg['label'],
                                      color=cfg['color'], s=point_size, alpha=0.95, 
                                      edgecolor='white', linewidth=1.5, zorder=cfg['zorder'])
        plotted_handles.append(handle)

    if params.get('show_ma_brent') and 'brent_ma_visual' in plot_df.columns:
        ma_handle, = ax1.plot(plot_df['x_idx'], plot_df['brent_ma_visual'], 
                             label=f'Brent MA ({params["ma_brent_win"]})',
                             color=colors['brent'], linewidth=line_width*0.7, linestyle='--', alpha=0.5)
        plotted_handles.append(ma_handle)

    if params.get('show_ma_cif') and 'cif_ma_visual' in plot_df.columns:
        ma_handle, = ax1.plot(plot_df['x_idx'], plot_df['cif_ma_visual'], 
                             label=f'CIF MA ({params["ma_cif_win"]})',
                             color=colors['cif'], linewidth=line_width*0.7, linestyle='--', alpha=0.5)
        plotted_handles.append(ma_handle)

    label_interval = int(params.get('label_interval', 0))
    combined_mask = pd.Series(False, index=plot_df.index)
    if label_interval > 0:
        combined_mask.iloc[::-label_interval] = True
    else:
        for cfg in configs:
            is_p, is_t = get_peaks_and_troughs(plot_df[cfg['col']], window)
            threshold_mask = plot_df[f"{cfg['col']}_degisim"] > cfg['thresh']
            combined_mask = combined_mask | is_p | is_t | threshold_mask

    if not plot_df.empty:
        combined_mask.iloc[-1] = True

    labeled_indices = plot_df[combined_mask].index.tolist()
    fs_labels = int(params.get('font_size_labels', 10))
    fs_axis = int(params.get('font_size_axis', 10))
    fs_legend = int(params.get('font_size_legend', 10))
    
    for cfg in configs:
        for i, idx in enumerate(labeled_indices):
            row = plot_df.loc[idx]
            val_real = row[cfg['col']]
            val_vis = row[cfg['vis_col']]
            x_pos = row['x_idx']
            formatted_val = cfg["fmt"].format(val_real).replace(",",".")
            is_p, is_t = get_peaks_and_troughs(plot_df[cfg['col']], window)
            if period == 'Yıllık':
                offset = 15 if cfg['col'] != 'brent' else -22
            else:
                offset = 12 if is_p.get(idx, True) else -22
                if i > 0:
                    prev_idx = labeled_indices[i-1]
                    if (idx - prev_idx) <= 1:
                        offset = 15 if i == len(labeled_indices) - 1 else -28 
            
            t = cfg['ax'].annotate(formatted_val, xy=(x_pos, val_vis), 
                                   xytext=(0, offset), textcoords='offset points', ha='center',
                                   fontsize=fs_labels, fontweight='800', color=cfg['color'], zorder=5)
            t.set_path_effects([pe.withStroke(linewidth=3, foreground='white', alpha=0.9)])

    y1_min, y1_max = ax1.get_ylim()
    ax1_upper = float(params.get('ax1_upper_factor', 1.8))
    ax1.set_ylim(y1_min*0.80, y1_max * ax1_upper) 
    y2_min, y2_max = plot_df['brent_katsayisi'].min(), plot_df['brent_katsayisi'].max()
    ax2_low = float(params.get('ax2_lower_factor', 0.1))
    ax2_high = float(params.get('ax2_upper_factor', 1.2))
    ax2.set_ylim(y2_min * ax2_low, y2_max * ax2_high) 

    tr_months = {1:"Ocak",2:"Şubat",3:"Mart",4:"Nisan",5:"Mayıs",6:"Haziran",7:"Temmuz",8:"Ağustos",9:"Eylül",10:"Ekim",11:"Kasım",12:"Aralık"}

    if period == "Yıllık":
        tick_indices = plot_df['x_idx'].tolist()
        tick_labels = plot_df['tarih'].dt.strftime('%Y').tolist()
    elif period == "Aylık":
        tick_indices = plot_df['x_idx'].tolist()
        tick_labels = [tr_months[d.month] for d in plot_df['tarih']]
    else:
        step = max(1, len(plot_df) // 10)
        tick_indices = plot_df['x_idx'][::step].tolist()
        last_idx = plot_df['x_idx'].iloc[-1]
        if last_idx not in tick_indices:
            if len(tick_indices) > 0 and (last_idx - tick_indices[-1]) < (step * 0.7):
                tick_indices[-1] = last_idx
            else:
                tick_indices.append(last_idx)
        tick_labels = [f"{d.day} {tr_months[d.month]}" for d in plot_df.loc[tick_indices, 'tarih']]

    ax1.set_xticks(tick_indices)
    fw_axis = 'bold' if params.get('bold_axis', False) else '500'
    rot_axis = int(params.get('x_rotation', 0))
    ax1.set_xticklabels(tick_labels, rotation=rot_axis, fontsize=fs_axis, fontweight=fw_axis, color='black')

    if params.get('show_legend', True):
        ncols = len(plotted_handles) if params.get('legend_horizontal', False) else 1
        figure.legend(plotted_handles, [h.get_label() for h in plotted_handles], 
                   loc='upper left', frameon=True, framealpha=0.9, edgecolor='#E5E7EB', fancybox=True, fontsize=fs_legend, ncol=ncols)
    
    if params.get('show_forecasts', False):
        forecast_x = len(plot_df) - 1
        if forecast_x >= 0:
            for f in AUTHORITY_FORECASTS:
                y_vis = f['val'] * shifts['brent']
                ax1.plot(forecast_x, y_vis, marker='*', markersize=10, color='gold', markeredgecolor='black', markeredgewidth=0.5, zorder=10)
                txt = ax1.annotate(f"{f['inst']}: {f['disp']}", xy=(forecast_x, y_vis), xytext=(15, 0), textcoords='offset points', fontsize=fs_labels*0.8, fontweight='bold', color='#4B5563', va='center', zorder=11)
                txt.set_path_effects([pe.withStroke(linewidth=2, foreground='white')])
    figure.tight_layout(pad=2.0)

class MarketApp:
    def __init__(self, root):
        self.root = root
        self.root.title("THY Market Intelligence - Desktop v1.0")
        self.root.geometry("1200x800")
        self.df = None
        self.current_file = None
        self.colors = {'c_brent': "#000000", 'c_cif': THY_NAVY, 'c_kat': THY_RED}
        self.setup_ui()
        self.load_on_startup()

    def setup_ui(self):
        main_container = tk.Frame(self.root, bg=THY_WHITE)
        main_container.pack(fill=tk.BOTH, expand=True)
        self.sidebar_container = tk.Frame(main_container, width=320, bg=THY_BG_LIGHT)
        self.sidebar_container.pack(side=tk.LEFT, fill=tk.Y)
        self.sidebar_container.pack_propagate(False)
        self.sidebar = tk.Canvas(self.sidebar_container, bg=THY_BG_LIGHT, highlightthickness=0)
        self.sidebar_scroll = tk.Scrollbar(self.sidebar_container, orient="vertical", command=self.sidebar.yview)
        self.scrollable_frame = tk.Frame(self.sidebar, bg=THY_BG_LIGHT)
        self.scrollable_frame.bind("<Configure>", lambda e: self.sidebar.configure(scrollregion=self.sidebar.bbox("all")))
        self.sidebar_window = self.sidebar.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.sidebar.configure(yscrollcommand=self.sidebar_scroll.set)
        def _on_mousewheel(event):
            if event.num == 4 or event.delta > 0: self.sidebar.yview_scroll(-1, "units")
            elif event.num == 5 or event.delta < 0: self.sidebar.yview_scroll(1, "units")
        self.sidebar.bind_all("<MouseWheel>", _on_mousewheel)
        self.sidebar.bind_all("<Button-4>", _on_mousewheel)
        self.sidebar.bind_all("<Button-5>", _on_mousewheel)
        self.sidebar.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.sidebar_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        header_f = tk.Frame(self.scrollable_frame, bg=THY_NAVY, pady=20)
        header_f.pack(fill=tk.X)
        tk.Label(header_f, text="✈️ THY MARKET", font=("Segoe UI", 16, "bold"), bg=THY_NAVY, fg=THY_WHITE).pack()
        tk.Label(header_f, text="INTELLIGENCE DASHBOARD", font=("Segoe UI", 8, "bold"), bg=THY_NAVY, fg=THY_SILVER, pady=5).pack()
        self.setup_sidebar_sections()
        self.frame_right = tk.Frame(main_container, bg=THY_WHITE)
        self.frame_right.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        self.setup_insight_bar()
        self.notebook = ttk.Notebook(self.frame_right)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        self.tab_chart = tk.Frame(self.notebook, bg=THY_WHITE)
        self.notebook.add(self.tab_chart, text=" 📈 Analiz Grafiği ")
        self.tab_dash = tk.Frame(self.notebook, bg=THY_WHITE)
        self.notebook.add(self.tab_dash, text=" 📊 Özet Dashboard ")
        self.setup_dashboard_tab()
        self.fig = plt.Figure(figsize=(8, 6), dpi=100, facecolor=THY_WHITE) 
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.tab_chart)
        self.canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.toolbar = NavigationToolbar2Tk(self.canvas, self.tab_chart)
        self.toolbar.config(background=THY_WHITE)
        self.toolbar.update()

    def setup_dashboard_tab(self):
        self.dash_content = tk.Frame(self.tab_dash, bg=THY_WHITE, padx=30, pady=30)
        self.dash_content.pack(fill=tk.BOTH, expand=True)
        self.lbl_dash_title = tk.Label(self.dash_content, text="PİYASA PERFORMANS ÖZETİ", font=("Segoe UI", 16, "bold"), bg=THY_WHITE, fg=THY_NAVY)
        self.lbl_dash_title.pack(anchor="w", pady=(0, 20))
        self.dash_stats_frame = tk.Frame(self.dash_content, bg=THY_WHITE)
        self.dash_stats_frame.pack(fill=tk.X)

    def update_dashboard_tab(self):
        if self.df is None: return
        for widget in self.dash_stats_frame.winfo_children(): widget.destroy()
        df = self.df.sort_values('tarih')
        last_date = df['tarih'].max()
        this_month = df[(df['tarih'].dt.month == last_date.month) & (df['tarih'].dt.year == last_date.year)]
        prev_month_date = last_date.replace(day=1) - timedelta(days=1)
        prev_month = df[(df['tarih'].dt.month == prev_month_date.month) & (df['tarih'].dt.year == prev_month_date.year)]
        last_week = df[df['tarih'] > (last_date - timedelta(days=7))]
        prev_week = df[(df['tarih'] <= (last_date - timedelta(days=7))) & (df['tarih'] > (last_date - timedelta(days=14)))]
        def create_metric_card(parent, title, value, delta, row, col):
            card = tk.Frame(parent, bg="#F9FAFB", highlightbackground="#E5E7EB", highlightthickness=1, padx=20, pady=20)
            card.grid(row=row, column=col, padx=10, pady=10, sticky="nsew")
            tk.Label(card, text=title, font=("Segoe UI", 10, "bold"), bg="#F9FAFB", fg="#6B7280").pack(anchor="w")
            tk.Label(card, text=value, font=("Segoe UI", 18, "bold"), bg="#F9FAFB", fg=THY_NAVY).pack(anchor="w", pady=5)
            if delta:
                color = "#059669" if "+" not in delta else "#DC2626"
                tk.Label(card, text=f"Değişim: {delta}", font=("Segoe UI", 10, "bold"), bg="#F9FAFB", fg=color).pack(anchor="w")
        create_metric_card(self.dash_stats_frame, "AYLIK JET YAKITI ORT", f"${this_month['cif med'].mean():.1f}", get_delta(this_month['cif med'].mean(), prev_month['cif med'].mean()) if not prev_month.empty else None, 0, 0)
        create_metric_card(self.dash_stats_frame, "AYLIK BRENT ORT", f"${this_month['brent'].mean():.2f}", get_delta(this_month['brent'].mean(), prev_month['brent'].mean()) if not prev_month.empty else None, 0, 1)
        create_metric_card(self.dash_stats_frame, "AYLIK KATSAYI ORT", f"{this_month['brent_katsayisi'].mean():.2f}", None, 0, 2)
        create_metric_card(self.dash_stats_frame, "HAFTALIK JET YAKITI ORT", f"${last_week['cif med'].mean():.1f}", get_delta(last_week['cif med'].mean(), prev_week['cif med'].mean()) if not prev_week.empty else None, 1, 0)
        create_metric_card(self.dash_stats_frame, "HAFTALIK BRENT ORT", f"${last_week['brent'].mean():.2f}", get_delta(last_week['brent'].mean(), prev_week['brent'].mean()) if not prev_week.empty else None, 1, 1)
        create_metric_card(self.dash_stats_frame, "HAFTALIK KATSAYI ORT", f"{last_week['brent_katsayisi'].mean():.2f}", None, 1, 2)
        hist_f = tk.Frame(self.dash_content, bg="#EFF6FF", padx=20, pady=20)
        hist_f.pack(fill=tk.X, pady=20)
        last_seen = find_last_time_similar(df, df['cif med'].iloc[-1], 'cif med')
        ath_row = df.loc[df['cif med'].idxmax()]
        tk.Label(hist_f, text="🕰️ TARİHSEL ANALİZ VE TREND", font=("Segoe UI", 11, "bold"), bg="#EFF6FF", fg="#1E40AF").pack(anchor="w")
        if last_seen: tk.Label(hist_f, text=f"• Mevcut fiyat seviyeleri en son {last_seen.strftime('%d %B %Y')} tarihinde görüldü.", font=("Segoe UI", 10), bg="#EFF6FF", fg="#1E3A8A").pack(anchor="w", pady=2)
        tk.Label(hist_f, text=f"• Tarihsel zirve (ATH): ${ath_row['cif med']:.0f} ({ath_row['tarih'].strftime('%d.%m.%Y')})", font=("Segoe UI", 10), bg="#EFF6FF", fg="#1E3A8A").pack(anchor="w", pady=2)
        forecast_f = tk.Frame(self.dash_content, bg=THY_WHITE)
        forecast_f.pack(fill=tk.BOTH, expand=True, pady=(20, 0))
        tk.Label(forecast_f, text="🏛️ OTORİTE BRENT TAHMİNLERİ (2026 Q4)", font=("Segoe UI", 12, "bold"), bg=THY_WHITE, fg=THY_NAVY).pack(anchor="w", pady=(0, 10))
        table_inner = tk.Frame(forecast_f, bg="#E5E7EB")
        table_inner.pack(fill=tk.X)
        headers = ["Kurum / Otorite", "Brent Tahmini", "Rapor Tarihi", "Önemli Not / Varsayım"]
        for c, h in enumerate(headers): tk.Label(table_inner, text=h, font=("Segoe UI", 9, "bold"), bg="#F3F4F6", fg="#374151", padx=10, pady=5).grid(row=0, column=c, sticky="nsew")
        for r, f in enumerate(AUTHORITY_FORECASTS):
            bg = THY_WHITE if r % 2 == 0 else "#F9FAFB"
            tk.Label(table_inner, text=f['inst'], font=("Segoe UI", 9), bg=bg, fg=THY_TEXT_DARK, padx=10, pady=4).grid(row=r+1, column=0, sticky="nsew")
            tk.Label(table_inner, text=f['disp'], font=("Segoe UI", 9, "bold"), bg=bg, fg=THY_RED, padx=10, pady=4).grid(row=r+1, column=1, sticky="nsew")
            tk.Label(table_inner, text=f['date'], font=("Segoe UI", 9), bg=bg, fg="#6B7280", padx=10, pady=4).grid(row=r+1, column=2, sticky="nsew")
            tk.Label(table_inner, text=f['note'], font=("Segoe UI", 9), bg=bg, fg="#6B7280", padx=10, pady=4, anchor="w").grid(row=r+1, column=3, sticky="nsew")
        for i in range(4): table_inner.columnconfigure(i, weight=1)
        for i in range(3): self.dash_stats_frame.columnconfigure(i, weight=1)

    def create_section(self, text, is_collapsed=False):
        section_frame = tk.Frame(self.scrollable_frame, bg=THY_BG_LIGHT)
        section_frame.pack(fill=tk.X, pady=(10, 2))
        header_f = tk.Frame(section_frame, bg="#D1D5DB", pady=1)
        header_f.pack(fill=tk.X)
        content_frame = tk.Frame(section_frame, bg=THY_BG_LIGHT)
        if not is_collapsed: content_frame.pack(fill=tk.X, pady=5)
        def toggle():
            if content_frame.winfo_viewable(): content_frame.pack_forget(); btn.config(text=f"▶ {text}")
            else: content_frame.pack(fill=tk.X, pady=5); btn.config(text=f"▼ {text}")
            self.root.update_idletasks(); self.sidebar.configure(scrollregion=self.sidebar.bbox("all"))
        prefix = "▼ " if not is_collapsed else "▶ "
        btn = tk.Label(header_f, text=f"{prefix}{text}", font=("Segoe UI", 11, "bold"), bg=THY_BG_LIGHT, fg=THY_NAVY, padx=10, pady=8, cursor="hand2", anchor="w")
        btn.pack(fill=tk.X)
        btn.bind("<Button-1>", lambda e: toggle())
        return content_frame

    def setup_sidebar_sections(self):
        data_f = self.create_section("VERİ YÖNETİMİ", is_collapsed=False)
        auto_f = tk.Frame(data_f, bg=THY_BG_LIGHT); auto_f.pack(fill=tk.X, padx=20, pady=5)
        tk.Label(auto_f, text="🚀 OUTLOOK OTOMASYONU", font=("Segoe UI", 10, "bold"), bg=THY_BG_LIGHT, fg=THY_RED).pack(pady=(5,5))
        tk.Label(auto_f, text="Outlook Klasör Adı:", bg=THY_BG_LIGHT, fg="#111827", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.ent_folder = tk.Entry(auto_f, font=("Segoe UI", 9)); self.ent_folder.insert(0, "jet fuel"); self.ent_folder.pack(fill=tk.X, pady=(2,10))
        self.btn_auto_update = self.create_button(data_f, "📩 OUTLOOK'TAN GÜNCELLE", self.run_outlook_pipeline, THY_RED, THY_WHITE, hover_color=THY_RED_HOVER)
        tk.Frame(data_f, height=2, bg="#E5E7EB").pack(fill=tk.X, padx=20, pady=10)
        self.btn_load_csv = self.create_button(data_f, "1. Ana CSV Yükle", self.load_initial_csv, THY_WHITE, THY_NAVY, is_outline=True)
        self.btn_update_excel = self.create_button(data_f, "2. Excel Güncelle (Manuel)", self.load_update_excel, THY_WHITE, THY_NAVY, is_outline=True)
        self.btn_save_csv = self.create_button(data_f, "3. CSV Kaydet", self.save_updated_csv, THY_NAVY, THY_WHITE, hover_color=THY_NAVY_HOVER)
        self.btn_baz = self.create_button(data_f, "Baz Senaryo Göster", self.toggle_baz, THY_WHITE, THY_NAVY, is_outline=True)
        self.lbl_file_status = tk.Label(data_f, text="Veri Yüklenmedi", bg=THY_BG_LIGHT, fg="#374151", font=("Segoe UI", 10, "italic")); self.lbl_file_status.pack(pady=5)
        self.vars = {'period': tk.StringVar(value="Günlük"), 'start_date': tk.StringVar(value="2026-02-01"), 'end_date': tk.StringVar(value=datetime.now().strftime('%Y-%m-%d')), 'brent_mult': tk.StringVar(value="10"), 'cif_shift': tk.StringVar(value="200"), 'window': tk.StringVar(value="0"), 'label_interval': tk.StringVar(value="0"), 'th_brent': tk.StringVar(value="0.03"), 'th_cif': tk.StringVar(value="0.03"), 'th_kat': tk.StringVar(value="0.05"), 'point_size': tk.DoubleVar(value=15), 'line_width': tk.DoubleVar(value=2.5), 'show_ma_brent': tk.BooleanVar(value=False), 'ma_brent_win': tk.StringVar(value="5"), 'show_ma_cif': tk.BooleanVar(value=False), 'ma_cif_win': tk.StringVar(value="5"), 'width': tk.StringVar(value="35"), 'height': tk.StringVar(value="17"), 'font_size_labels': tk.StringVar(value="12"), 'font_size_axis': tk.StringVar(value="10"), 'font_size_legend': tk.StringVar(value="12"), 'show_legend': tk.BooleanVar(value=True), 'legend_horizontal': tk.BooleanVar(value=False), 'bold_axis': tk.BooleanVar(value=False), 'x_rotation': tk.IntVar(value=0), 'ax1_upper_factor': tk.StringVar(value="1.8"), 'ax2_lower_factor': tk.StringVar(value="0.1"), 'ax2_upper_factor': tk.StringVar(value="1.2"), 'show_brent': tk.BooleanVar(value=True), 'show_cif': tk.BooleanVar(value=True), 'show_kat': tk.BooleanVar(value=True), 'dotted_2026': tk.BooleanVar(value=True), 'show_baz': tk.BooleanVar(value=False), 'show_forecasts': tk.BooleanVar(value=False)}
        for var in self.vars.values(): var.trace_add("write", lambda *args: self.update_plot())
        time_section_f = self.create_section("ZAMAN & PERİYOT", is_collapsed=False)
        time_f = tk.Frame(time_section_f, bg=THY_BG_LIGHT, pady=5); time_f.pack(fill=tk.X, padx=20)
        self.create_sidebar_label(time_f, "Veri Seçimi", 0)
        sel_f = tk.Frame(time_f, bg=THY_BG_LIGHT); sel_f.grid(row=1, column=0, sticky="w", pady=(0, 10))
        tk.Checkbutton(sel_f, text="Jet Yakıt", variable=self.vars['show_cif'], bg=THY_BG_LIGHT).pack(side=tk.LEFT)
        tk.Checkbutton(sel_f, text="Brent", variable=self.vars['show_brent'], bg=THY_BG_LIGHT).pack(side=tk.LEFT)
        tk.Checkbutton(sel_f, text="Katsayı", variable=self.vars['show_kat'], bg=THY_BG_LIGHT).pack(side=tk.LEFT)
        tk.Checkbutton(sel_f, text="Otorite", variable=self.vars['show_forecasts'], bg=THY_BG_LIGHT).pack(side=tk.LEFT)
        self.create_sidebar_label(time_f, "Görünüm Periyodu", 2)
        ttk.Combobox(time_f, textvariable=self.vars['period'], values=["Günlük", "Haftalık", "Aylık", "Yıllık"], width=15).grid(row=3, column=0, pady=(0, 10), sticky="w")
        self.create_sidebar_label(time_f, "Analiz Başlangıç", 4)
        tk.Entry(time_f, textvariable=self.vars['start_date'], width=18).grid(row=5, column=0, pady=(0, 10), sticky="w")
        self.create_sidebar_label(time_f, "Analiz Bitiş", 6)
        tk.Entry(time_f, textvariable=self.vars['end_date'], width=18).grid(row=7, column=0, sticky="w")
        ana_section_f = self.create_section("ANALİZ AYARLARI", is_collapsed=True)
        ana_f = tk.Frame(ana_section_f, bg=THY_BG_LIGHT, pady=5); ana_f.pack(fill=tk.X, padx=20)
        tk.Checkbutton(ana_f, text="Brent MA (Gün)", variable=self.vars['show_ma_brent'], bg=THY_BG_LIGHT, font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w")
        tk.Entry(ana_f, textvariable=self.vars['ma_brent_win'], width=5).grid(row=0, column=1)
        tk.Checkbutton(ana_f, text="CIF MA (Gün)", variable=self.vars['show_ma_cif'], bg=THY_BG_LIGHT, font=("Segoe UI", 10, "bold")).grid(row=1, column=0, sticky="w")
        tk.Entry(ana_f, textvariable=self.vars['ma_cif_win'], width=5).grid(row=1, column=1)
        self.create_sidebar_label(ana_f, "Etiket Sıklığı (X Günde Bir)", 2); tk.Entry(ana_f, textvariable=self.vars['label_interval'], width=8).grid(row=3, column=0, sticky="w")
        self.create_sidebar_label(ana_f, "Pik/Dip Hassasiyeti", 4); tk.Entry(ana_f, textvariable=self.vars['window'], width=8).grid(row=5, column=0, sticky="w")
        vis_section_f = self.create_section("GÖRSEL AYARLAR", is_collapsed=True)
        vis_f = tk.Frame(vis_section_f, bg=THY_BG_LIGHT, pady=5); vis_f.pack(fill=tk.X, padx=20)
        self.create_vis_input(vis_f, "Brent Çarpanı", 'brent_mult', 0); self.create_vis_input(vis_f, "CIF Kaydırma", 'cif_shift', 1)
        self.create_sidebar_label(vis_f, "Grafik Yerleşimi", 2)
        self.create_vis_input(vis_f, "Üst Boşluk", 'ax1_upper_factor', 3)
        self.create_vis_input(vis_f, "Katsayı Alt", 'ax2_lower_factor', 4)
        self.create_vis_input(vis_f, "Katsayı Üst", 'ax2_upper_factor', 5)
        self.create_sidebar_label(vis_f, "Etiket Font", 6); tk.Entry(vis_f, textvariable=self.vars['font_size_labels'], width=8).grid(row=7, column=0, sticky="w")
        self.create_sidebar_label(vis_f, "X Eksen Font", 8); tk.Entry(vis_f, textvariable=self.vars['font_size_axis'], width=8).grid(row=9, column=0, sticky="w")
        self.create_sidebar_label(vis_f, "Dönüş", 10); tk.Scale(vis_f, variable=self.vars['x_rotation'], from_=0, to=90, orient=tk.HORIZONTAL, bg=THY_BG_LIGHT).grid(row=11, column=0, sticky="w")
        tk.Checkbutton(vis_f, text="Legend Göster", variable=self.vars['show_legend'], bg=THY_BG_LIGHT).grid(row=15, column=0, sticky="w")
        dim_section_f = self.create_section("BOYUT & STİL", is_collapsed=True)
        dim_f = tk.Frame(dim_section_f, bg=THY_BG_LIGHT, pady=5); dim_f.pack(fill=tk.X, padx=20)
        self.create_vis_input(dim_f, "Genişlik (cm)", 'width', 0); self.create_vis_input(dim_f, "Yükseklik (cm)", 'height', 1)
        act_section_f = self.create_section("İŞLEMLER", is_collapsed=False)
        self.create_button(act_section_f, "🔄 GRAFİĞİ GÜNCELLE", self.update_plot, THY_RED, THY_WHITE, hover_color=THY_RED_HOVER)
        self.create_button(act_section_f, "📸 RESMİ KAYDET", self.save_image, THY_NAVY, THY_WHITE, hover_color=THY_NAVY_HOVER)

    def create_sidebar_label(self, parent, text, row): tk.Label(parent, text=text, bg=THY_BG_LIGHT, fg="#111827", font=("Segoe UI", 10, "bold")).grid(row=row, column=0, sticky="w", pady=(8, 2))
    def create_button(self, parent, text, command, bg, fg, hover_color=None, is_outline=False):
        f = tk.Frame(parent, bg=THY_NAVY if is_outline else bg, padx=1, pady=1)
        btn = tk.Label(f, text=text, bg=bg, fg=fg, font=("Segoe UI", 11, "bold"), pady=12, cursor="hand2"); btn.pack(fill=tk.X)
        btn.bind("<Enter>", lambda e: btn.config(bg=hover_color) if hover_color else None); btn.bind("<Leave>", lambda e: btn.config(bg=bg))
        btn.bind("<Button-1>", lambda e: command()); f.pack(fill=tk.X, padx=20, pady=4)
        return f
    def create_vis_input(self, parent, label, var_key, row):
        tk.Label(parent, text=label+":", bg=THY_BG_LIGHT, fg="#4B5563").grid(row=row, column=0, sticky="w"); tk.Entry(parent, textvariable=self.vars[var_key], width=8).grid(row=row, column=1, pady=4)

    def setup_insight_bar(self):
        self.insight_frame = tk.Frame(self.frame_right, bg=THY_NAVY, pady=20); self.insight_frame.pack(side=tk.TOP, fill=tk.X)
        self.lbl_metrics = tk.Label(self.insight_frame, text="Yükleniyor...", font=("Segoe UI", 14, "bold"), bg=THY_NAVY, fg=THY_WHITE); self.lbl_metrics.pack()
        self.lbl_history = tk.Label(self.insight_frame, text="", font=("Segoe UI", 10), bg=THY_NAVY, fg="#94A3B8"); self.lbl_history.pack()

    def load_on_startup(self):
        if os.path.exists(PERSISTENT_FILE): self.load_csv(PERSISTENT_FILE)
        else:
            files = glob.glob(DATA_FILENAME_PATTERN)
            if files: self.load_csv(max(files, key=os.path.getmtime))

    def load_csv(self, path):
        try:
            df = pd.read_csv(path); df["tarih"] = pd.to_datetime(df["tarih"], errors='coerce'); df = df.dropna(subset=["tarih"])
            self.df = df[["tarih","brent","cif med"]].copy(); self.df['brent_katsayisi'] = self.df['cif med'] / self.df['brent']
            self.lbl_file_status.config(text=f"📂 {os.path.basename(path)}"); self.calculate_insights(); self.update_dashboard_tab(); self.update_plot()
        except: pass

    def load_initial_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")]); 
        if path: self.load_csv(path)

    def calculate_insights(self):
        if self.df is None: return
        df = self.df.sort_values('tarih'); last_date = df['tarih'].max()
        this_month = df[(df['tarih'].dt.month == last_date.month) & (df['tarih'].dt.year == last_date.year)]
        prev_month = df[(df['tarih'].dt.month == (last_date.replace(day=1)-timedelta(days=1)).month)]
        self.lbl_metrics.config(text=f"JET: ${df['cif med'].iloc[-1]:.0f} | BRENT: ${df['brent'].iloc[-1]:.2f} | AY ORT: ${this_month['cif med'].mean():.1f}")

    def load_update_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if path: self.load_any_excel(path)

    def load_any_excel(self, path):
        try:
            df_up = pd.read_excel(path, skiprows=4)
            # Find columns dynamically to avoid "Position 1" error
            cols = [str(c).strip().upper() for c in df_up.columns]
            # Try to find date column (first column usually)
            df_up = df_up.iloc[:, [0, 1, 2]]
            df_up.columns = ["tarih","brent","cif med"]
            
            # THE FIX: Ensure clean conversion with specific format handling if needed
            df_up["tarih"] = pd.to_datetime(df_up["tarih"], errors='coerce')
            df_up = df_up.dropna(subset=["tarih"])
            
            df_up["brent"] = pd.to_numeric(df_up["brent"], errors='coerce')
            df_up["cif med"] = pd.to_numeric(df_up["cif med"], errors='coerce')
            df_up = df_up.dropna(subset=["brent", "cif med"])

            if self.df is None: self.df = df_up.copy()
            else:
                for _, row in df_up.iterrows():
                    match = self.df.index[self.df['tarih'] == row['tarih']].tolist()
                    if match:
                        idx = match[0]
                        self.df.at[idx, 'brent'] = row['brent']
                        self.df.at[idx, 'cif med'] = row['cif med']
                    else:
                        self.df = pd.concat([self.df, pd.DataFrame([row])], ignore_index=True)

            self.df['brent_katsayisi'] = self.df['cif med'] / self.df['brent']
            self.df = self.df.sort_values('tarih').reset_index(drop=True)
            self.df.to_csv(PERSISTENT_FILE, index=False)
            self.calculate_insights(); self.update_dashboard_tab(); self.update_plot()
            messagebox.showinfo("Başarılı", "Veriler güncellendi.")
        except Exception as e:
            messagebox.showerror("Hata", f"Excel işleme hatası: {e}")

    def run_outlook_pipeline(self):
        folder = self.ent_folder.get().strip() or "jet fuel"
        self.lbl_file_status.config(text="⏳ Outlook taranıyor...")
        self.root.update()
        if check_and_download_specific_mails(folder):
            process_and_merge_files()
            self.load_any_excel("Guncel_Master_Veri.xlsx")
            self.lbl_file_status.config(text="✅ Güncellendi")
        else:
            self.lbl_file_status.config(text="❌ Bulunamadı")

    def toggle_baz(self):
        self.vars['show_baz'].set(not self.vars['show_baz'].get()); self.update_plot()

    def update_plot(self):
        if self.df is None: return
        params = {k: v.get() for k, v in self.vars.items()}; params.update(self.colors)
        analyze_and_plot(self.df, params, self.fig); self.canvas.draw()

    def save_image(self):
        path = filedialog.asksaveasfilename(defaultextension=".png")
        if path: self.fig.savefig(path, dpi=300); messagebox.showinfo("Başarılı", "Kaydedildi")

if __name__ == "__main__":
    root = tk.Tk(); app = MarketApp(root); root.mainloop()
