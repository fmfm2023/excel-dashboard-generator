#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel Dashboard Generator - Flask API
Génère des dashboards Excel professionnels à partir de fichiers CSV/XLSX.
Auteur: Excel Dashboard Generator
Version: 1.0.0
"""

import os
import io
import base64
import json
import logging
import traceback
from datetime import datetime
from flask import Flask, request, jsonify
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side
)
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
import chardet
import warnings
warnings.filterwarnings('ignore')

# ─────────────────────────────────────────────────────────────────────────────
# Configuration Flask
# ─────────────────────────────────────────────────────────────────────────────
app = Flask(__name__)
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
logger = logging.getLogger(__name__)

# Palette de couleurs professionnelle
COLORS = {
    'primary_dark':  '1F4E79',
    'primary':       '2E75B6',
    'primary_light': 'D6E4F0',
    'accent':        '70AD47',
    'accent_light':  'E2EFDA',
    'warning':       'ED7D31',
    'danger':        'C00000',
    'success':       '375623',
    'bg_header':     '1F4E79',
    'bg_subheader':  '2E75B6',
    'bg_alt':        'F2F7FC',
    'white':         'FFFFFF',
    'light_gray':    'F5F5F5',
    'border':        'BFBFBF',
    'text_dark':     '1F1F1F',
    'text_light':    '595959',
}

# ─────────────────────────────────────────────────────────────────────────────
# Helpers de style
# ─────────────────────────────────────────────────────────────────────────────

def hex_fill(hex_color):
    return PatternFill(fill_type='solid', fgColor=hex_color)

def thin_border(color=COLORS['border']):
    side = Side(style='thin', color=color)
    return Border(left=side, right=side, top=side, bottom=side)

def header_border():
    return Border(
        bottom=Side(style='medium', color=COLORS['primary_dark'])
    )

def apply_header_style(cell, bg=COLORS['bg_header'], fg=COLORS['white'],
                        size=11, bold=True, center=True):
    cell.fill = hex_fill(bg)
    cell.font = Font(name='Calibri', size=size, bold=bold, color=fg)
    if center:
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    else:
        cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

def apply_cell_style(cell, bg=None, fg=COLORS['text_dark'], size=10,
                     bold=False, center=False, wrap=True):
    if bg:
        cell.fill = hex_fill(bg)
    cell.font = Font(name='Calibri', size=size, bold=bold, color=fg)
    h = 'center' if center else 'left'
    cell.alignment = Alignment(horizontal=h, vertical='center', wrap_text=wrap)

def apply_number_format(cell, fmt):
    cell.number_format = fmt

def set_col_width(ws, col, width):
    ws.column_dimensions[get_column_letter(col)].width = width

def set_row_height(ws, row, height):
    ws.row_dimensions[row].height = height

def merge_and_style(ws, start_row, start_col, end_row, end_col,
                    value='', bg=COLORS['bg_header'], fg=COLORS['white'],
                    size=11, bold=True):
    ws.merge_cells(
        start_row=start_row, start_column=start_col,
        end_row=end_row, end_column=end_col
    )
    cell = ws.cell(row=start_row, column=start_col)
    cell.value = value
    apply_header_style(cell, bg=bg, fg=fg, size=size, bold=bold)
    return cell


# ─────────────────────────────────────────────────────────────────────────────
# Chargement et détection des données
# ─────────────────────────────────────────────────────────────────────────────

def detect_encoding(raw_bytes):
    result = chardet.detect(raw_bytes)
    enc = result.get('encoding', 'utf-8') or 'utf-8'
    # Normaliser
    enc = enc.upper().replace('-', '')
    mapping = {'UTF8': 'utf-8', 'UTF8BOM': 'utf-8-sig',
                'ISO88591': 'iso-8859-1', 'WINDOWS1252': 'cp1252'}
    return mapping.get(enc, enc.lower())

def detect_separator(text_sample):
    counts = {';': text_sample.count(';'), ',': text_sample.count(','),
              '\t': text_sample.count('\t'), '|': text_sample.count('|')}
    return max(counts, key=counts.get)

def load_dataframe(file_bytes, filename):
    """Charge un fichier CSV ou XLSX en DataFrame pandas."""
    ext = filename.lower().rsplit('.', 1)[-1] if '.' in filename else ''
    if ext in ('xlsx', 'xls', 'xlsm'):
        df = pd.read_excel(io.BytesIO(file_bytes), engine='openpyxl')
    else:
        encoding = detect_encoding(file_bytes)
        text = file_bytes.decode(encoding, errors='replace')
        sep = detect_separator(text[:2000])
        df = pd.read_csv(io.StringIO(text), sep=sep, engine='python')

    # Nettoyage de base
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how='all').reset_index(drop=True)
    if len(df) > 10000:
        df = df.head(10000)
        logger.warning('Fichier tronqué à 10 000 lignes')
    return df

def classify_columns(df):
    """Détecte automatiquement le type de chaque colonne."""
    numeric_cols, date_cols, cat_cols = [], [], []
    for col in df.columns:
        series = df[col].dropna()
        if series.empty:
            continue
        # Essayer conversion date
        if series.dtype == object:
            try:
                converted = pd.to_datetime(series, infer_datetime_format=True, errors='coerce')
                if converted.notna().sum() / len(series) > 0.7:
                    df[col] = pd.to_datetime(df[col], infer_datetime_format=True, errors='coerce')
                    date_cols.append(col)
                    continue
            except Exception:
                pass
        if pd.api.types.is_numeric_dtype(series):
            numeric_cols.append(col)
        elif pd.api.types.is_datetime64_any_dtype(series):
            date_cols.append(col)
        else:
            # Essayer conversion numérique sur colonnes object
            try:
                converted = pd.to_numeric(series.str.replace(',', '.', regex=False), errors='coerce')
                if converted.notna().sum() / len(series) > 0.7:
                    df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', '.', regex=False),
                                            errors='coerce')
                    numeric_cols.append(col)
                    continue
            except Exception:
                pass
            cat_cols.append(col)
    return df, numeric_cols, date_cols, cat_cols


# ─────────────────────────────────────────────────────────────────────────────
# Calcul des KPIs
# ─────────────────────────────────────────────────────────────────────────────

def compute_kpis(df, numeric_cols, date_cols, cat_cols):
    kpis = {
        'total_rows': len(df),
        'total_columns': len(df.columns),
        'numeric_columns': len(numeric_cols),
        'categorical_columns': len(cat_cols),
        'date_columns': len(date_cols),
        'numeric_stats': {},
        'cat_stats': {},
        'date_range': {},
        'growth': {},
    }

    # Stats numériques
    for col in numeric_cols:
        s = df[col].dropna()
        if s.empty:
            continue
        kpis['numeric_stats'][col] = {
            'total': float(s.sum()),
            'mean': float(s.mean()),
            'median': float(s.median()),
            'min': float(s.min()),
            'max': float(s.max()),
            'std': float(s.std()),
            'count': int(s.count()),
            'pct_null': round(df[col].isna().mean() * 100, 1),
        }

    # Stats catégorielles (top 10 par fréquence)
    for col in cat_cols:
        s = df[col].dropna()
        if s.empty:
            continue
        vc = s.value_counts()
        top10 = vc.head(10)
        kpis['cat_stats'][col] = {
            'unique_values': int(s.nunique()),
            'top_values': [
                {'value': str(k), 'count': int(v),
                 'pct': round(v / len(s) * 100, 1)}
                for k, v in top10.items()
            ],
            'mode': str(vc.index[0]) if len(vc) > 0 else None,
        }

    # Plage de dates
    for col in date_cols:
        s = df[col].dropna()
        if s.empty:
            continue
        kpis['date_range'][col] = {
            'min': str(s.min().date()) if hasattr(s.min(), 'date') else str(s.min()),
            'max': str(s.max().date()) if hasattr(s.max(), 'date') else str(s.max()),
        }

    # Calcul de croissance si dates + numériques
    if date_cols and numeric_cols:
        date_col = date_cols[0]
        num_col = numeric_cols[0]
        try:
            temp = df[[date_col, num_col]].dropna()
            temp[date_col] = pd.to_datetime(temp[date_col])
            temp = temp.sort_values(date_col)
            temp['period'] = temp[date_col].dt.to_period('M').astype(str)
            monthly = temp.groupby('period')[num_col].sum()
            if len(monthly) >= 2:
                growth = (monthly.iloc[-1] - monthly.iloc[-2]) / monthly.iloc[-2] * 100
                kpis['growth'] = {
                    'col': num_col,
                    'last_period': monthly.index[-1],
                    'prev_period': monthly.index[-2],
                    'growth_pct': round(float(growth), 1),
                    'monthly_data': monthly.to_dict(),
                }
        except Exception as e:
            logger.warning(f'Calcul croissance échoué: {e}')

    return kpis


# ─────────────────────────────────────────────────────────────────────────────
# Onglet Dashboard
# ─────────────────────────────────────────────────────────────────────────────

def build_dashboard_sheet(wb, df, kpis, numeric_cols, date_cols, cat_cols, filename):
    ws = wb.create_sheet('📊 Dashboard', 0)
    ws.sheet_view.showGridLines = False

    # Largeurs colonnes fixes
    col_widths = [2, 18, 14, 14, 14, 14, 14, 14, 14, 14, 2]
    for i, w in enumerate(col_widths, 1):
        set_col_width(ws, i, w)
    ws.row_dimensions[1].height = 8  # marge top

    # ── En-tête principal ─────────────────────────────────────────────────────
    merge_and_style(ws, 2, 2, 4, 10,
                    value=f'📊 DASHBOARD ANALYTIQUE — {filename.upper()}',
                    bg=COLORS['bg_header'], fg=COLORS['white'], size=16)
    merge_and_style(ws, 5, 2, 5, 10,
                    value=f'Généré le {datetime.now().strftime("%d/%m/%Y à %H:%M")}  |  '
                          f'{kpis["total_rows"]:,} enregistrements  |  '
                          f'{kpis["total_columns"]} colonnes',
                    bg=COLORS['bg_subheader'], fg=COLORS['white'], size=10, bold=False)
    ws.row_dimensions[2].height = 30
    ws.row_dimensions[5].height = 18

    # ── Section KPIs ──────────────────────────────────────────────────────────
    row = 7
    merge_and_style(ws, row, 2, row, 10,
                    value='🔑 INDICATEURS CLÉS DE PERFORMANCE',
                    bg=COLORS['primary'], fg=COLORS['white'], size=11)
    ws.row_dimensions[row].height = 20
    row += 1

    kpi_items = []
    # KPI 1 : total enregistrements
    kpi_items.append(('📋 Enregistrements', kpis['total_rows'], '#,##0', None, None))
    # KPI 2 : première colonne numérique total
    if numeric_cols:
        col = numeric_cols[0]
        stats = kpis['numeric_stats'].get(col, {})
        kpi_items.append((f'💰 Total {col[:15]}', stats.get('total', 0), '#,##0.00', None, None))
        kpi_items.append((f'📈 Moyenne {col[:12]}', stats.get('mean', 0), '#,##0.00', None, None))
        kpi_items.append((f'⬆️ Max {col[:15]}', stats.get('max', 0), '#,##0.00', None, None))
    # KPI 5 : croissance
    if kpis.get('growth'):
        g = kpis['growth']
        color_val = COLORS['accent'] if g['growth_pct'] >= 0 else COLORS['danger']
        sign = '+' if g['growth_pct'] >= 0 else ''
        kpi_items.append((f'📉 Croissance M/M', g['growth_pct'] / 100,
                          '+0.0%;-0.0%;0.0%', color_val, None))
    # KPI : catégories uniques
    if cat_cols:
        col = cat_cols[0]
        st = kpis['cat_stats'].get(col, {})
        kpi_items.append((f'🏷️ {col[:15]} uniques', st.get('unique_values', 0), '#,##0', None, None))

    # Afficher les KPIs en ligne (max 4 par ligne sur 2 colonnes larges)
    n_per_row = 4
    kpi_col_start = 2
    kpi_col_width = 2  # chaque KPI occupe 2 colonnes
    for idx, (label, value, fmt, color, _) in enumerate(kpi_items[:8]):
        row_offset = idx // n_per_row
        col_offset = idx % n_per_row
        r = row + row_offset * 3
        c = kpi_col_start + col_offset * 2

        # Titre KPI
        cell_title = ws.cell(row=r, column=c, value=label)
        apply_cell_style(cell_title, bg=COLORS['bg_alt'], size=9,
                         fg=COLORS['text_light'], bold=False, center=True)
        ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + 1)
        cell_title.border = thin_border()
        ws.row_dimensions[r].height = 14

        # Valeur KPI
        cell_val = ws.cell(row=r + 1, column=c, value=value)
        bg_color = color or COLORS['white']
        txt_color = COLORS['white'] if color else COLORS['primary_dark']
        apply_header_style(cell_val, bg=bg_color, fg=txt_color, size=18)
        apply_number_format(cell_val, fmt)
        ws.merge_cells(start_row=r + 1, start_column=c, end_row=r + 1, end_column=c + 1)
        cell_val.border = thin_border(COLORS['primary'])
        ws.row_dimensions[r + 1].height = 32

    row += (min(len(kpi_items), 8) // n_per_row + 1) * 3 + 1

    return ws, row


# ─────────────────────────────────────────────────────────────────────────────
# Graphiques natifs Excel
# ─────────────────────────────────────────────────────────────────────────────

def add_charts(wb, ws, df, kpis, numeric_cols, date_cols, cat_cols, start_row):
    """Ajoute 3 graphiques natifs Excel sur le Dashboard."""
    row = start_row
    merge_and_style(ws, row, 2, row, 10,
                    value='📈 GRAPHIQUES ANALYTIQUES',
                    bg=COLORS['primary'], fg=COLORS['white'], size=11)
    ws.row_dimensions[row].height = 20
    row += 1

    # ── Onglet données intermédiaires (caché) pour les graphiques ────────────
    ws_data = wb.create_sheet('_chart_data')
    ws_data.sheet_state = 'hidden'

    # ── 1) Graphique en barres : Top catégories ───────────────────────────────
    chart_row = 1
    if cat_cols and numeric_cols:
        cat_col = cat_cols[0]
        num_col = numeric_cols[0]
        top_data = df.groupby(cat_col)[num_col].sum().nlargest(10).reset_index()

        ws_data.cell(row=chart_row, column=1, value='Catégorie')
        ws_data.cell(row=chart_row, column=2, value=f'Total {num_col}')
        for i, (_, r_data) in enumerate(top_data.iterrows(), 1):
            ws_data.cell(row=chart_row + i, column=1, value=str(r_data[cat_col]))
            ws_data.cell(row=chart_row + i, column=2, value=float(r_data[num_col]))
        n_bars = len(top_data)

        bar_chart = BarChart()
        bar_chart.type = 'col'
        bar_chart.style = 10
        bar_chart.title = f'Top {n_bars} — {cat_col}'
        bar_chart.y_axis.title = f'Total {num_col}'
        bar_chart.x_axis.title = cat_col
        bar_chart.width = 16
        bar_chart.height = 11
        bar_chart.grouping = 'clustered'
        bar_chart.overlap = 0

        data_ref = Reference(ws_data, min_col=2, min_row=chart_row,
                             max_row=chart_row + n_bars)
        cats_ref = Reference(ws_data, min_col=1, min_row=chart_row + 1,
                             max_row=chart_row + n_bars)
        bar_chart.add_data(data_ref, titles_from_data=True)
        bar_chart.set_categories(cats_ref)
        bar_chart.series[0].graphicalProperties.solidFill = COLORS['primary']
        bar_chart.series[0].graphicalProperties.line.solidFill = COLORS['primary_dark']

        ws.add_chart(bar_chart, f'B{row}')
        chart_row += n_bars + 3

    # ── 2) Graphique courbe : évolution temporelle ────────────────────────────
    if date_cols and numeric_cols and kpis.get('growth', {}).get('monthly_data'):
        monthly = kpis['growth']['monthly_data']
        num_col = kpis['growth']['col']
        periods = list(monthly.keys())
        values = list(monthly.values())

        ws_data.cell(row=chart_row, column=4, value='Période')
        ws_data.cell(row=chart_row, column=5, value=f'Total {num_col}')
        for i, (p, v) in enumerate(zip(periods, values), 1):
            ws_data.cell(row=chart_row + i, column=4, value=p)
            ws_data.cell(row=chart_row + i, column=5, value=float(v))
        n_periods = len(periods)

        line_chart = LineChart()
        line_chart.style = 10
        line_chart.title = f'Évolution mensuelle — {num_col}'
        line_chart.y_axis.title = f'Total {num_col}'
        line_chart.x_axis.title = 'Période'
        line_chart.width = 16
        line_chart.height = 11

        data_ref = Reference(ws_data, min_col=5, min_row=chart_row,
                             max_row=chart_row + n_periods)
        cats_ref = Reference(ws_data, min_col=4, min_row=chart_row + 1,
                             max_row=chart_row + n_periods)
        line_chart.add_data(data_ref, titles_from_data=True)
        line_chart.set_categories(cats_ref)
        s = line_chart.series[0]
        s.graphicalProperties.line.solidFill = COLORS['accent']
        s.graphicalProperties.line.width = 25000
        s.smooth = True

        ws.add_chart(line_chart, f'F{row}')
        chart_row += n_periods + 3

    # ── 3) Camembert : répartition ────────────────────────────────────────────
    if cat_cols:
        cat_col = cat_cols[0]
        top5 = df[cat_col].value_counts().head(6)

        ws_data.cell(row=chart_row, column=7, value='Segment')
        ws_data.cell(row=chart_row, column=8, value='Nombre')
        for i, (val, cnt) in enumerate(top5.items(), 1):
            ws_data.cell(row=chart_row + i, column=7, value=str(val))
            ws_data.cell(row=chart_row + i, column=8, value=int(cnt))
        n_pie = len(top5)

        pie = PieChart()
        pie.style = 10
        pie.title = f'Répartition — {cat_col}'
        pie.width = 14
        pie.height = 11

        data_ref = Reference(ws_data, min_col=8, min_row=chart_row,
                             max_row=chart_row + n_pie)
        cats_ref = Reference(ws_data, min_col=7, min_row=chart_row + 1,
                             max_row=chart_row + n_pie)
        pie.add_data(data_ref, titles_from_data=True)
        pie.set_categories(cats_ref)
        pie.dataLabels = DataLabelList()
        pie.dataLabels.showPercent = True
        pie.dataLabels.showSerName = False

        # Position camembert : en dessous des barres
        pie_row = row + 22
        ws.add_chart(pie, f'B{pie_row}')

    return row + 22


# ─────────────────────────────────────────────────────────────────────────────
# Onglet Données brutes
# ─────────────────────────────────────────────────────────────────────────────

def build_raw_data_sheet(wb, df, numeric_cols, date_cols):
    ws = wb.create_sheet('📋 Données brutes')
    ws.sheet_view.showGridLines = False

    n_rows, n_cols = df.shape

    # En-tête du tableau
    for col_idx, col_name in enumerate(df.columns, 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        apply_header_style(cell, bg=COLORS['bg_header'], size=10)
        cell.border = thin_border(COLORS['primary_dark'])

    # Données avec alternance de couleurs
    for row_idx, (_, row_data) in enumerate(df.iterrows(), 2):
        bg = COLORS['bg_alt'] if row_idx % 2 == 0 else COLORS['white']
        for col_idx, (col_name, value) in enumerate(row_data.items(), 1):
            # Convertir valeurs numpy pour openpyxl
            if isinstance(value, (np.integer,)):
                value = int(value)
            elif isinstance(value, (np.floating,)):
                value = float(value) if not np.isnan(value) else None
            elif isinstance(value, (pd.Timestamp, np.datetime64)):
                try:
                    value = pd.Timestamp(value).to_pydatetime()
                except Exception:
                    value = str(value)

            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            apply_cell_style(cell, bg=bg, size=9)
            cell.border = thin_border()

            # Format selon type
            if col_name in numeric_cols:
                apply_number_format(cell, '#,##0.00')
                cell.alignment = Alignment(horizontal='right', vertical='center')
            elif col_name in date_cols:
                apply_number_format(cell, 'DD/MM/YYYY')

    # Ajuster largeurs automatiquement
    for col_idx in range(1, n_cols + 1):
        col_letter = get_column_letter(col_idx)
        max_len = max(
            len(str(df.columns[col_idx - 1])),
            df.iloc[:, col_idx - 1].astype(str).str.len().max()
        )
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 40)

    ws.row_dimensions[1].height = 20

    # Figer la ligne d'en-tête
    ws.freeze_panes = ws['A2']

    # Mise en forme conditionnelle sur colonnes numériques
    for col_name in numeric_cols[:3]:
        col_idx = list(df.columns).index(col_name) + 1
        col_letter = get_column_letter(col_idx)
        data_range = f'{col_letter}2:{col_letter}{n_rows + 1}'
        ws.conditional_formatting.add(data_range, ColorScaleRule(
            start_type='min', start_color='C00000',
            mid_type='percentile', mid_value=50, mid_color='FFFF00',
            end_type='max', end_color='375623'
        ))

    return ws


# ─────────────────────────────────────────────────────────────────────────────
# Onglet Analyse détaillée (TCD simulés)
# ─────────────────────────────────────────────────────────────────────────────

def build_analysis_sheet(wb, df, kpis, numeric_cols, date_cols, cat_cols):
    ws = wb.create_sheet('🔍 Analyse détaillée')
    ws.sheet_view.showGridLines = False

    row = 1

    def section_header(ws, row, title):
        merge_and_style(ws, row, 1, row, 8, value=title,
                        bg=COLORS['bg_header'], fg=COLORS['white'], size=11)
        ws.row_dimensions[row].height = 22
        return row + 1

    def table_header(ws, row, cols, bg=COLORS['bg_subheader']):
        for c_idx, col_name in enumerate(cols, 1):
            cell = ws.cell(row=row, column=c_idx, value=col_name)
            apply_header_style(cell, bg=bg, size=10)
            cell.border = thin_border()
        ws.row_dimensions[row].height = 18
        return row + 1

    # ── TCD 1 : Résumé par catégorie ──────────────────────────────────────────
    if cat_cols and numeric_cols:
        cat_col = cat_cols[0]
        row = section_header(ws, row, f'📊 Tableau croisé : {cat_col} × Numériques')

        num_subset = numeric_cols[:4]
        headers = [cat_col, 'Nb lignes'] + \
                  [f'Total {c[:10]}' for c in num_subset] + \
                  [f'Moy. {c[:10]}' for c in num_subset]
        row = table_header(ws, row, headers)

        grouped = df.groupby(cat_col)
        agg_dict = {c: ['sum', 'mean'] for c in num_subset}
        agg_dict['_count'] = lambda x: len(x)
        try:
            summary = df.groupby(cat_col).agg(
                {c: ['sum', 'mean'] for c in num_subset}
            )
            summary.columns = ['_'.join(c) for c in summary.columns]
            counts = df.groupby(cat_col).size().rename('count')
            summary = summary.join(counts)
            summary = summary.sort_values('count', ascending=False)

            for r_idx, (cat_val, r_data) in enumerate(summary.iterrows()):
                bg = COLORS['bg_alt'] if r_idx % 2 == 0 else COLORS['white']
                col_data = [str(cat_val), int(r_data['count'])]
                for nc in num_subset:
                    col_data.append(round(float(r_data.get(f'{nc}_sum', 0)), 2))
                for nc in num_subset:
                    col_data.append(round(float(r_data.get(f'{nc}_mean', 0)), 2))

                for c_idx, val in enumerate(col_data, 1):
                    cell = ws.cell(row=row + r_idx, column=c_idx, value=val)
                    apply_cell_style(cell, bg=bg, size=9)
                    cell.border = thin_border()
                    if c_idx > 2:
                        apply_number_format(cell, '#,##0.00')
                        cell.alignment = Alignment(horizontal='right', vertical='center')

            row += len(summary) + 2
        except Exception as e:
            logger.warning(f'TCD catégorie: {e}')
            row += 2

    # ── TCD 2 : Évolution temporelle ──────────────────────────────────────────
    if date_cols and numeric_cols:
        date_col = date_cols[0]
        num_col = numeric_cols[0]
        row = section_header(ws, row, f'📅 Évolution par période : {date_col}')
        row = table_header(ws, row, ['Période', 'Nb transactions',
                                      f'Total {num_col[:12]}',
                                      f'Moy. {num_col[:12]}',
                                      'Min', 'Max', '% du total'])
        try:
            temp = df[[date_col, num_col]].dropna()
            temp[date_col] = pd.to_datetime(temp[date_col])
            temp['period'] = temp[date_col].dt.to_period('M').astype(str)
            monthly = temp.groupby('period')[num_col].agg(
                ['count', 'sum', 'mean', 'min', 'max']
            ).reset_index()
            total_all = monthly['sum'].sum()

            for r_idx, r_data in monthly.iterrows():
                bg = COLORS['bg_alt'] if r_idx % 2 == 0 else COLORS['white']
                pct = round(r_data['sum'] / total_all * 100, 1) if total_all else 0
                row_vals = [r_data['period'], int(r_data['count']),
                            round(float(r_data['sum']), 2),
                            round(float(r_data['mean']), 2),
                            round(float(r_data['min']), 2),
                            round(float(r_data['max']), 2),
                            pct / 100]
                for c_idx, val in enumerate(row_vals, 1):
                    cell = ws.cell(row=row + r_idx, column=c_idx, value=val)
                    apply_cell_style(cell, bg=bg, size=9)
                    cell.border = thin_border()
                    if c_idx in (3, 4, 5, 6):
                        apply_number_format(cell, '#,##0.00')
                        cell.alignment = Alignment(horizontal='right', vertical='center')
                    elif c_idx == 7:
                        apply_number_format(cell, '0.0%')
                        cell.alignment = Alignment(horizontal='right', vertical='center')

            row += len(monthly) + 2
        except Exception as e:
            logger.warning(f'TCD temporel: {e}')
            row += 2

    # ── Statistiques descriptives ─────────────────────────────────────────────
    if numeric_cols:
        row = section_header(ws, row, '📐 Statistiques descriptives')
        row = table_header(ws, row, ['Colonne', 'Total', 'Moyenne', 'Médiane',
                                      'Min', 'Max', 'Écart-type', '% valeurs nulles'])
        for r_idx, nc in enumerate(numeric_cols):
            s = kpis['numeric_stats'].get(nc, {})
            if not s:
                continue
            bg = COLORS['bg_alt'] if r_idx % 2 == 0 else COLORS['white']
            row_vals = [nc, s.get('total', 0), s.get('mean', 0),
                        s.get('median', 0), s.get('min', 0), s.get('max', 0),
                        s.get('std', 0), s.get('pct_null', 0) / 100]
            for c_idx, val in enumerate(row_vals, 1):
                cell = ws.cell(row=row + r_idx, column=c_idx, value=val)
                apply_cell_style(cell, bg=bg, size=9)
                cell.border = thin_border()
                if c_idx > 1:
                    fmt = '0.0%' if c_idx == 8 else '#,##0.00'
                    apply_number_format(cell, fmt)
                    cell.alignment = Alignment(horizontal='right', vertical='center')
        row += len(numeric_cols) + 2

    # Ajuster largeurs
    col_widths_analysis = [25, 14, 14, 14, 12, 12, 12, 16]
    for i, w in enumerate(col_widths_analysis, 1):
        set_col_width(ws, i, w)

    return ws


# ─────────────────────────────────────────────────────────────────────────────
# Orchestration principale
# ─────────────────────────────────────────────────────────────────────────────

def generate_excel_dashboard(file_bytes, filename):
    """
    Charge les données, calcule les KPIs, génère le fichier Excel dashboard.
    Retourne (excel_bytes, kpis_dict) ou lève une exception.
    """
    logger.info(f'Traitement de {filename} ({len(file_bytes)} octets)')

    # 1) Chargement
    df = load_dataframe(file_bytes, filename)
    if df.empty:
        raise ValueError('Le fichier est vide ou ne contient aucune donnée valide.')
    if len(df.columns) < 1:
        raise ValueError('Aucune colonne détectée dans le fichier.')

    logger.info(f'DataFrame chargé : {df.shape[0]} lignes × {df.shape[1]} colonnes')

    # 2) Classification des colonnes
    df, numeric_cols, date_cols, cat_cols = classify_columns(df)
    logger.info(f'Colonnes — num: {numeric_cols}, dates: {date_cols}, cat: {cat_cols}')

    # 3) Calcul des KPIs
    kpis = compute_kpis(df, numeric_cols, date_cols, cat_cols)

    # 4) Construction du classeur Excel
    wb = Workbook()
    # Supprimer la feuille par défaut
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']

    # 4a) Onglet Dashboard
    ws_dash, next_row = build_dashboard_sheet(
        wb, df, kpis, numeric_cols, date_cols, cat_cols, filename
    )

    # 4b) Graphiques natifs
    add_charts(wb, ws_dash, df, kpis, numeric_cols, date_cols, cat_cols, next_row)

    # 4c) Onglet Données brutes
    build_raw_data_sheet(wb, df, numeric_cols, date_cols)

    # 4d) Onglet Analyse détaillée
    build_analysis_sheet(wb, df, kpis, numeric_cols, date_cols, cat_cols)

    # 5) Propriétés du classeur
    wb.properties.title = f'Dashboard — {filename}'
    wb.properties.subject = 'Dashboard analytique généré automatiquement'
    wb.properties.creator = 'Excel Dashboard Generator'
    wb.properties.description = (
        f'Généré le {datetime.now().strftime("%d/%m/%Y %H:%M")} '
        f'à partir de {filename}'
    )

    # 6) Sérialisation en mémoire
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    excel_bytes = buf.read()

    logger.info(f'Excel généré : {len(excel_bytes)} octets')
    return excel_bytes, kpis


# ─────────────────────────────────────────────────────────────────────────────
# Routes Flask
# ─────────────────────────────────────────────────────────────────────────────

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'version': '1.0.0', 'timestamp': str(datetime.now())}), 200


@app.route('/generate-dashboard', methods=['POST'])
def generate_dashboard():
    """
    POST JSON :
    {
      "filename": "data.csv",
      "file_data": "<base64>",
      "email": "user@example.com",
      "file_type": "csv"   # optionnel
    }
    Retourne JSON :
    {
      "status": "success" | "error",
      "excel_base64": "<base64>",
      "kpis": {...},
      "error_message": "..."
    }
    """
    try:
        payload = request.get_json(force=True)
        if not payload:
            return jsonify({'status': 'error',
                            'error_message': 'Corps JSON manquant ou invalide.'}), 400

        filename = payload.get('filename', 'upload.csv')
        file_data_b64 = payload.get('file_data', '')
        if not file_data_b64:
            return jsonify({'status': 'error',
                            'error_message': 'Champ file_data manquant ou vide.'}), 400

        # Décoder le base64
        try:
            # Nettoyer le préfixe data URI éventuel
            if ',' in file_data_b64:
                file_data_b64 = file_data_b64.split(',', 1)[1]
            # Padding
            file_data_b64 += '=' * (-len(file_data_b64) % 4)
            file_bytes = base64.b64decode(file_data_b64)
        except Exception as e:
            return jsonify({'status': 'error',
                            'error_message': f'Erreur de décodage base64 : {e}'}), 400

        if len(file_bytes) == 0:
            return jsonify({'status': 'error',
                            'error_message': 'Fichier vide reçu.'}), 400

        # Générer le dashboard
        excel_bytes, kpis = generate_excel_dashboard(file_bytes, filename)

        # Encoder la réponse
        excel_b64 = base64.b64encode(excel_bytes).decode('utf-8')

        return jsonify({
            'status': 'success',
            'excel_base64': excel_b64,
            'kpis': {
                'total_rows': kpis['total_rows'],
                'total_columns': kpis['total_columns'],
                'numeric_columns': kpis['numeric_columns'],
                'categorical_columns': kpis['categorical_columns'],
            },
            'filename': filename,
        }), 200

    except ValueError as ve:
        logger.warning(f'Erreur métier : {ve}')
        return jsonify({'status': 'error', 'error_message': str(ve)}), 422

    except Exception as exc:
        logger.error(f'Erreur inattendue : {exc}\n{traceback.format_exc()}')
        return jsonify({
            'status': 'error',
            'error_message': f'Erreur interne du serveur : {exc}'
        }), 500


@app.route('/generate-from-upload', methods=['POST'])
def generate_from_upload():
    """Route alternative : accepte un fichier multipart/form-data."""
    try:
        if 'file' not in request.files:
            return jsonify({'status': 'error', 'error_message': 'Champ file manquant.'}), 400

        f = request.files['file']
        email = request.form.get('email', '')
        filename = f.filename or 'upload'
        file_bytes = f.read()

        if len(file_bytes) == 0:
            return jsonify({'status': 'error', 'error_message': 'Fichier vide.'}), 400

        excel_bytes, kpis = generate_excel_dashboard(file_bytes, filename)
        excel_b64 = base64.b64encode(excel_bytes).decode('utf-8')

        return jsonify({
            'status': 'success',
            'excel_base64': excel_b64,
            'kpis': {
                'total_rows': kpis['total_rows'],
                'total_columns': kpis['total_columns'],
                'numeric_columns': kpis['numeric_columns'],
                'categorical_columns': kpis['categorical_columns'],
            },
            'filename': filename,
            'email': email,
        }), 200

    except Exception as exc:
        logger.error(f'Erreur upload: {exc}\n{traceback.format_exc()}')
        return jsonify({'status': 'error', 'error_message': str(exc)}), 500


# ─────────────────────────────────────────────────────────────────────────────
# Point d'entrée
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    debug = os.environ.get('DEBUG', 'false').lower() == 'true'
    logger.info(f'Démarrage serveur sur port {port} (debug={debug})')
    app.run(host='0.0.0.0', port=port, debug=debug)
