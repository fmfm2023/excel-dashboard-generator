"""
Dashboard Excel Pro v2.0 — Style Power BI
Génère un dashboard Excel professionnel avec KPIs, graphiques, analyses.
"""
import io, os, logging, base64, math
from datetime import datetime

import pandas as pd
import numpy as np
from flask import Flask, request, jsonify

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
from openpyxl.chart import BarChart, LineChart, PieChart, AreaChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
logger = logging.getLogger(__name__)
app = Flask(__name__)

# ══════════════════════════════════════════════════════════════════════════════
# PALETTE (Power BI / Microsoft Fluent)
# ══════════════════════════════════════════════════════════════════════════════
C = {
    'hdr':         '16213E', 'hdr2':        '0F3460', 'hdr_fg':     'FFFFFF',
    'blue':        '0078D4', 'blue2':       '2E86C1', 'blue_l':     'DDEEFF',
    'green':       '107C10', 'green_l':     'DFF6DD',
    'orange':      'D97706', 'orange_l':    'FEF3C7',
    'red':         'DC2626', 'red_l':       'FEE2E2',
    'purple':      '7C3AED', 'purple_l':    'EDE9FE',
    'teal':        '0D9488', 'teal_l':      'CCFBF1',
    'yellow':      'CA8A04', 'yellow_l':    'FEF9C3',
    'white':       'FFFFFF', 'g1':          'F8FAFC', 'g2':         'E2E8F0',
    'g3':          'CBD5E1', 'g4':          '94A3B8', 'g5':         '475569',
    'dark':        '1E293B', 'row_alt':     'F0F4F8',
    # KPI card backgrounds
    'k1': '0078D4', 'k2': '107C10', 'k3': '0D9488',
    'k4': 'D97706', 'k5': '7C3AED', 'k6': 'DC2626',
}
CHART_PAL = ['0078D4','107C10','D97706','7C3AED','DC2626',
             '0D9488','CA8A04','2E86C1','059669','9333EA']

# ══════════════════════════════════════════════════════════════════════════════
# UTILITAIRES STYLES
# ══════════════════════════════════════════════════════════════════════════════
def _hex(c):
    """Normalise une couleur hex : supprime '#', assure 6 chars."""
    c = str(c).lstrip('#').upper()
    return c if len(c) == 6 else c[:6] if len(c) > 6 else c.ljust(6, '0')

def _fill(hex_color):
    return PatternFill(fgColor=_hex(hex_color), fill_type='solid')

def _font(size=10, bold=False, color='1E293B', italic=False):
    return Font(name='Calibri', size=size, bold=bold, color=_hex(color), italic=italic)

def _align(h='left', v='center', wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap, indent=0)

def _border(color='CBD5E1', style='thin'):
    s = Side(style=style, color=_hex(color))
    return Border(left=s, right=s, top=s, bottom=s)

def _border_bottom(color='0078D4', top_color='CBD5E1'):
    return Border(
        left=Side(style='thin', color=top_color),
        right=Side(style='thin', color=top_color),
        top=Side(style='thin', color=top_color),
        bottom=Side(style='medium', color=color)
    )

def sc(cell, bg=None, fg='1E293B', size=10, bold=False,
       h='left', v='center', wrap=False, italic=False, border=True, nf=None):
    """Style a cell (shorthand)."""
    if bg:
        cell.fill = _fill(bg)
    cell.font = _font(size=size, bold=bold, color=fg, italic=italic)
    cell.alignment = _align(h=h, v=v, wrap=wrap)
    if border:
        cell.border = _border()
    if nf:
        cell.number_format = nf

def merge_sc(ws, r1, c1, r2, c2, val='', bg=None, fg='FFFFFF',
             size=12, bold=True, h='center', v='center', wrap=False):
    """Merge cells and style."""
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    cell = ws.cell(row=r1, column=c1, value=val)
    cell.font = _font(size=size, bold=bold, color=fg)
    cell.alignment = _align(h=h, v=v, wrap=wrap)
    if bg:
        for r in range(r1, r2+1):
            for c in range(c1, c2+1):
                ws.cell(row=r, column=c).fill = _fill(bg)
    return cell

def col_w(ws, idx, w):
    ws.column_dimensions[get_column_letter(idx)].width = w

def row_h(ws, idx, h):
    ws.row_dimensions[idx].height = h

def fmt_num(v, dec=0):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return '—'
    if abs(v) >= 1_000_000:
        return f'{v/1_000_000:.2f}M'
    if abs(v) >= 1_000:
        return f'{v/1_000:.1f}K'
    return f'{v:,.{dec}f}'

def fmt_eur(v):
    return fmt_num(v) + ' €' if v is not None else '—'

def fmt_pct(v):
    return f'{v:.1f}%' if v is not None else '—'

# ══════════════════════════════════════════════════════════════════════════════
# CHARGEMENT DES DONNÉES (avec détection automatique de la ligne d'en-tête)
# ══════════════════════════════════════════════════════════════════════════════
def find_header_row(file_bytes, filename):
    """Détecte la ligne qui contient les vrais en-têtes (ignore les titres)."""
    ext = filename.rsplit('.', 1)[-1].lower()
    try:
        if ext in ('xlsx', 'xls', 'xlsm'):
            raw = pd.read_excel(io.BytesIO(file_bytes), header=None, nrows=10)
        else:
            text = file_bytes.decode('utf-8-sig', errors='replace')
            raw = pd.read_csv(io.StringIO(text), header=None, nrows=10)
        for i, row in raw.iterrows():
            vals = [v for v in row if pd.notna(v)]
            texts = [v for v in vals if isinstance(v, str) and v.strip()]
            if len(texts) >= 3 and len(texts) >= len(vals) * 0.5:
                return i
        return 0
    except Exception:
        return 0


def load_dataframe(file_bytes, filename):
    """Charge le DataFrame avec détection intelligente des en-têtes."""
    ext = filename.rsplit('.', 1)[-1].lower()
    hdr = find_header_row(file_bytes, filename)
    try:
        if ext in ('xlsx', 'xls', 'xlsm'):
            df = pd.read_excel(io.BytesIO(file_bytes), header=hdr)
        elif ext == 'csv':
            text = file_bytes.decode('utf-8-sig', errors='replace')
            df = pd.read_csv(io.StringIO(text), header=hdr)
        else:
            raise ValueError(f'Format non supporté : {ext}')
    except Exception as e:
        raise ValueError(f'Impossible de lire le fichier : {e}')

    df = df.dropna(how='all').reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]
    # Supprimer colonnes "Unnamed:"
    named = [c for c in df.columns if not c.startswith('Unnamed:')]
    if named:
        df = df[named]
    df = df.dropna(how='all').reset_index(drop=True)
    return df


# ══════════════════════════════════════════════════════════════════════════════
# DÉTECTION DES RÔLES DE COLONNES
# ══════════════════════════════════════════════════════════════════════════════
def detect_columns(df):
    """Mappe les colonnes à des rôles sémantiques via mots-clés."""
    cols = {c.lower().strip(): c for c in df.columns}

    def find(*kws):
        for kw in kws:
            for cl, co in cols.items():
                if kw in cl:
                    return co
        return None

    return {
        'date':      find('date', 'jour', 'période', 'period'),
        'client':    find('client', 'customer', 'acheteur', 'société'),
        'produit':   find('produit', 'product', 'article', 'désignation', 'libellé'),
        'categorie': find('catégorie', 'categorie', 'category', 'famille', 'type'),
        'marque':    find('marque', 'brand', 'fabricant', 'fournisseur'),
        'quantite':  find('quantité', 'quantite', 'qty', 'qté', 'nb ', 'volume'),
        'prix_unit': find('prix unitaire', 'prix_unit', 'unit price', 'pu '),
        'remise':    find('remise', 'discount', 'réduction', 'rabais', 'taux rem'),
        'total_ht':  find('total ht', 'ht (', 'montant ht', 'ca ht', 'ht€'),
        'tva':       find('tva', 'tax', 'vat'),
        'total_ttc': find('total ttc', 'ttc', 'montant ttc', 'total tc'),
        'statut':    find('statut', 'status', 'état', 'etat'),
        'vendeur':   find('vendeur', 'commercial', 'sales rep', 'representant', 'conseiller'),
        'commande':  find('commande', 'order', 'n° cmd', 'numéro', 'n°'),
        'region':    find('région', 'region', 'zone', 'secteur', 'département'),
    }


def get_numeric_cols(df):
    return [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]

# ══════════════════════════════════════════════════════════════════════════════
# CALCUL DES KPIs (spécialisé ventes + générique)
# ══════════════════════════════════════════════════════════════════════════════
def compute_kpis(df, cm):
    """Calcule tous les KPIs à partir du DataFrame et de la carte de colonnes."""
    kpis = {'nb_rows': len(df), 'nb_cols': len(df.columns)}

    def to_num(col):
        if col and col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    # Convertir colonnes numériques
    for role in ('total_ttc', 'total_ht', 'tva', 'quantite', 'prix_unit', 'remise'):
        to_num(cm[role])

    # ── CA ──────────────────────────────────────────────────────────────────
    ca_col = cm['total_ttc'] or cm['total_ht']
    if ca_col:
        kpis['ca_col']   = ca_col
        kpis['ca_total'] = float(df[ca_col].sum())
        kpis['ca_mean']  = float(df[ca_col].mean())
        kpis['ca_max']   = float(df[ca_col].max())
        kpis['ca_min']   = float(df[ca_col].fillna(0).min())

    # ── Commandes ───────────────────────────────────────────────────────────
    kpis['nb_commandes'] = len(df)
    if cm['commande']:
        kpis['nb_commandes'] = df[cm['commande']].nunique()

    # ── Clients ─────────────────────────────────────────────────────────────
    if cm['client']:
        kpis['nb_clients'] = int(df[cm['client']].nunique())
        if ca_col:
            kpis['top_clients'] = (df.groupby(cm['client'])[ca_col]
                                   .sum().sort_values(ascending=False).head(8))
    # ── Produits ────────────────────────────────────────────────────────────
    if cm['produit']:
        kpis['nb_produits'] = int(df[cm['produit']].nunique())
        if ca_col:
            kpis['top_produits'] = (df.groupby(cm['produit'])[ca_col]
                                    .sum().sort_values(ascending=False).head(8))
    # ── Quantité ────────────────────────────────────────────────────────────
    if cm['quantite']:
        kpis['qty_total'] = float(df[cm['quantite']].sum())

    # ── Remise ──────────────────────────────────────────────────────────────
    if cm['remise']:
        kpis['remise_moy'] = float(df[cm['remise']].mean())
        kpis['remise_max'] = float(df[cm['remise']].max())

    # ── Statut ──────────────────────────────────────────────────────────────
    if cm['statut']:
        sc_col = cm['statut']
        counts = df[sc_col].value_counts()
        kpis['statut_counts'] = counts.to_dict()
        total = len(df)
        for label in counts.index:
            key = label.lower().replace(' ', '_').replace('é', 'e').replace('è', 'e')
            kpis[f'statut_{key}'] = int(counts[label])
            kpis[f'pct_{key}'] = round(counts[label] / total * 100, 1)
        # Taux livraison
        livres = counts.get('Livré', counts.get('Livre', counts.get('livré', 0)))
        annules = counts.get('Annulé', counts.get('Annule', counts.get('annulé', 0)))
        kpis['nb_livres']  = int(livres)
        kpis['nb_annules'] = int(annules)
        kpis['taux_livraison']  = round(livres / total * 100, 1) if total else 0
        kpis['taux_annulation'] = round(annules / total * 100, 1) if total else 0

    # ── Catégorie ──────────────────────────────────────────────────────────
    if cm['categorie'] and ca_col:
        try:
            gp = df.groupby(cm['categorie'])
            cat = gp[ca_col].agg(['sum', 'count', 'mean'])
            cat.columns = ['ca', 'nb', 'moy']
            if cm['quantite']:
                cat['qty'] = gp[cm['quantite']].sum()
            if cm['remise']:
                cat['remise_moy'] = gp[cm['remise']].mean()
            kpis['by_categorie'] = cat.sort_values('ca', ascending=False)
        except Exception as e:
            logger.warning(f'by_categorie: {e}')

    # ── Vendeur ─────────────────────────────────────────────────────────────
    if cm['vendeur'] and ca_col:
        try:
            gp = df.groupby(cm['vendeur'])
            vd = gp[ca_col].agg(['sum', 'count', 'mean'])
            vd.columns = ['ca', 'nb', 'moy']
            if cm['remise']:
                vd['remise_moy'] = gp[cm['remise']].mean()
            kpis['by_vendeur'] = vd.sort_values('ca', ascending=False)
            kpis['nb_vendeurs'] = len(vd)
        except Exception as e:
            logger.warning(f'by_vendeur: {e}')

    # ── Marque ──────────────────────────────────────────────────────────────
    if cm['marque'] and ca_col:
        try:
            kpis['by_marque'] = (df.groupby(cm['marque'])[ca_col]
                                 .sum().sort_values(ascending=False))
        except Exception as e:
            logger.warning(f'by_marque: {e}')

    # ── Évolution mensuelle ─────────────────────────────────────────────────
    if cm['date'] and ca_col:
        try:
            tmp = df[[cm['date'], ca_col]].copy()
            tmp[cm['date']] = pd.to_datetime(tmp[cm['date']], errors='coerce')
            tmp = tmp.dropna(subset=[cm['date']])
            tmp['mois'] = tmp[cm['date']].dt.to_period('M')
            monthly = tmp.groupby('mois')[ca_col].agg(['sum', 'count']).reset_index()
            monthly.columns = ['mois', 'ca', 'nb']
            monthly['mois_str'] = monthly['mois'].astype(str)
            kpis['monthly'] = monthly
        except Exception as e:
            logger.warning(f'monthly: {e}')

    # ── Statut × vendeur croisé ─────────────────────────────────────────────
    if cm['statut'] and cm['vendeur']:
        try:
            pivot = pd.crosstab(df[cm['vendeur']], df[cm['statut']])
            kpis['pivot_statut_vendeur'] = pivot
        except Exception as e:
            logger.warning(f'pivot_statut_vendeur: {e}')

    return kpis

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 1 : DASHBOARD PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════
def build_dashboard_sheet(wb, df, kpis, cm, filename):
    ws = wb.create_sheet('📊 Dashboard')
    ws.sheet_view.showGridLines = False
    ws.sheet_view.showRowColHeaders = False

    # ── Largeurs colonnes (26 colonnes pour layout 2×13) ────────────────────
    for i in range(1, 27):
        col_w(ws, i, 4.5)

    # ── HEADER BANNER (lignes 1-4) ──────────────────────────────────────────
    for r in range(1, 5):
        row_h(ws, r, 22 if r in (1, 4) else 18)
        for c in range(1, 27):
            ws.cell(row=r, column=c).fill = _fill(C['hdr'])

    # Titre principal
    merge_sc(ws, 1, 1, 1, 18,
             val=f'📊  DASHBOARD ANALYTIQUE — {filename.upper()}',
             bg=C['hdr'], fg=C['white'], size=16, bold=True, h='left')
    ws.cell(row=1, column=1).alignment = _align(h='left', v='center')
    ws.cell(row=1, column=1).alignment = Alignment(horizontal='left', vertical='center', indent=2)

    # Sous-titre (date + métriques)
    nb = kpis.get('nb_rows', 0)
    nb_col = kpis.get('nb_cols', 0)
    sub = f"Généré le {datetime.now().strftime('%d/%m/%Y %H:%M')}   |   {nb} enregistrements   |   {nb_col} colonnes"
    merge_sc(ws, 2, 1, 2, 18, val=sub, bg=C['hdr'], fg=C['g4'], size=10, bold=False, h='left')
    ws.cell(row=2, column=1).alignment = Alignment(horizontal='left', vertical='center', indent=2)

    # Barre accent bleue bas du header
    for c in range(1, 27):
        cell = ws.cell(row=4, column=c)
        cell.fill = _fill(C['blue'])

    row_h(ws, 5, 10)  # espace

    # ── KPI CARDS (lignes 6-13, 6 cartes de 4 colonnes chacune) ────────────
    kpi_defs = [
        ('k1', '💰', 'CA TOTAL TTC',
         fmt_eur(kpis.get('ca_total')),
         f"Panier moy : {fmt_eur(kpis.get('ca_mean'))}", 1),
        ('k2', '📦', 'COMMANDES',
         str(kpis.get('nb_commandes', kpis.get('nb_rows', 0))),
         f"Clients : {kpis.get('nb_clients', '—')}", 5),
        ('k3', '✅', 'TAUX LIVRAISON',
         fmt_pct(kpis.get('taux_livraison')),
         f"Livrés : {kpis.get('nb_livres', '—')}", 9),
        ('k4', '🏷️', 'REMISE MOYENNE',
         fmt_pct(kpis.get('remise_moy')),
         f"Max : {fmt_pct(kpis.get('remise_max'))}", 13),
        ('k5', '👥', 'CLIENTS UNIQUES',
         str(kpis.get('nb_clients', '—')),
         f"Produits : {kpis.get('nb_produits', '—')}", 17),
        ('k6', '❌', 'TAUX ANNULATION',
         fmt_pct(kpis.get('taux_annulation')),
         f"Annulés : {kpis.get('nb_annules', '—')}", 21),
    ]

    for color_key, icon, label, value, sub_val, start_col in kpi_defs:
        bg = C[color_key]
        end_col = start_col + 3
        # Fond de la carte
        for r in range(6, 14):
            for c in range(start_col, end_col + 1):
                ws.cell(row=r, column=c).fill = _fill(bg)
        # Icône + label
        merge_sc(ws, 6, start_col, 6, end_col, val=f'{icon}  {label}',
                 bg=bg, fg='FFFFFF', size=9, bold=True, h='left')
        ws.cell(row=6, column=start_col).alignment = Alignment(
            horizontal='left', vertical='center', indent=1)
        row_h(ws, 6, 18)
        # Valeur principale
        merge_sc(ws, 7, start_col, 10, end_col, val=value,
                 bg=bg, fg='FFFFFF', size=20, bold=True, h='center')
        for r in range(7, 11):
            row_h(ws, r, 20)
        # Sous-valeur
        merge_sc(ws, 11, start_col, 13, end_col, val=sub_val,
                 bg=bg, fg='FFFFFF', size=9, bold=False, h='center')
        for r in range(11, 14):
            row_h(ws, r, 14)

    row_h(ws, 14, 12)  # séparateur

    # ── SECTION : Analyse par catégorie (tableau, lignes 15-30) ────────────
    if 'by_categorie' in kpis:
        merge_sc(ws, 15, 1, 15, 13, val='🏷️  CHIFFRE D\'AFFAIRES PAR CATÉGORIE',
                 bg=C['hdr2'], fg=C['white'], size=11, bold=True, h='left')
        ws.cell(row=15, column=1).alignment = Alignment(horizontal='left', vertical='center', indent=2)
        row_h(ws, 15, 22)

        # En-têtes tableau
        hdrs = ['Catégorie', 'CA TTC (€)', 'Nb Cmd', 'Panier Moy (€)', '% du Total']
        hdr_cols = [1, 4, 7, 9, 12]
        hdr_spans = [3, 3, 2, 3, 2]
        for hdr, sc_start, span in zip(hdrs, hdr_cols, hdr_spans):
            merge_sc(ws, 16, sc_start, 16, sc_start + span - 1,
                     val=hdr, bg=C['blue'], fg='FFFFFF', size=9, bold=True, h='center')
        row_h(ws, 16, 18)

        cat_df = kpis['by_categorie']
        total_ca = cat_df['ca'].sum()
        r = 17
        for i, (cat_name, row_data) in enumerate(cat_df.iterrows()):
            bg = C['row_alt'] if i % 2 == 0 else C['white']
            ca = row_data.get('ca', 0)
            nb = row_data.get('nb', 0)
            moy = row_data.get('moy', 0)
            pct = ca / total_ca * 100 if total_ca else 0

            merge_sc(ws, r, 1, r, 3, val=str(cat_name), bg=bg, fg=C['dark'],
                     size=9, bold=False, h='left')
            ws.cell(row=r, column=1).alignment = Alignment(horizontal='left', vertical='center', indent=1)
            merge_sc(ws, r, 4, r, 6, val=round(float(ca), 2), bg=bg, fg=C['dark'],
                     size=9, bold=True, h='right')
            ws.cell(row=r, column=4).number_format = '#,##0.00'
            merge_sc(ws, r, 7, r, 8, val=int(nb), bg=bg, fg=C['dark'], size=9, h='center')
            merge_sc(ws, r, 9, r, 11, val=round(float(moy), 2), bg=bg, fg=C['dark'],
                     size=9, h='right')
            ws.cell(row=r, column=9).number_format = '#,##0.00'
            merge_sc(ws, r, 12, r, 13, val=round(pct, 1), bg=bg, fg=C['dark'],
                     size=9, h='center')
            ws.cell(row=r, column=12).number_format = '0.0"%"'
            row_h(ws, r, 16)
            r += 1
            if r > 30:
                break

    # ── SECTION : Statut résumé (lignes 15-30, colonnes 14-26) ─────────────
    if 'statut_counts' in kpis:
        merge_sc(ws, 15, 15, 15, 26, val='📋  RÉPARTITION PAR STATUT',
                 bg=C['hdr2'], fg=C['white'], size=11, bold=True, h='left')
        ws.cell(row=15, column=15).alignment = Alignment(horizontal='left', vertical='center', indent=2)

        statut_colors = {
            'livré': (C['green'], C['green_l']),
            'livre': (C['green'], C['green_l']),
            'en attente': (C['orange'], C['orange_l']),
            'annulé': (C['red'], C['red_l']),
            'annule': (C['red'], C['red_l']),
        }
        r = 16
        total = kpis['nb_rows']
        for statut, cnt in kpis['statut_counts'].items():
            key = str(statut).lower()
            main_c, light_c = statut_colors.get(key, (C['blue'], C['blue_l']))
            pct = cnt / total * 100 if total else 0

            # Couleur indicateur
            merge_sc(ws, r, 15, r + 1, 16, val='', bg=main_c, fg='FFFFFF', size=10)
            # Statut label
            merge_sc(ws, r, 17, r, 22, val=str(statut), bg=light_c, fg=C['dark'],
                     size=10, bold=True, h='left')
            ws.cell(row=r, column=17).alignment = Alignment(horizontal='left', vertical='center', indent=1)
            # Count
            merge_sc(ws, r, 23, r, 24, val=cnt, bg=light_c, fg=C['dark'],
                     size=14, bold=True, h='center')
            # Pct
            merge_sc(ws, r, 25, r, 26, val=f'{pct:.0f}%', bg=light_c, fg=main_c,
                     size=12, bold=True, h='center')
            row_h(ws, r, 20)
            row_h(ws, r + 1, 4)
            r += 2

    # ── Top vendeurs rapide (lignes 32+) ────────────────────────────────────
    row_h(ws, 31, 12)
    if 'by_vendeur' in kpis:
        merge_sc(ws, 32, 1, 32, 26, val='🏆  CLASSEMENT VENDEURS (CA TTC)',
                 bg=C['hdr2'], fg=C['white'], size=11, bold=True, h='left')
        ws.cell(row=32, column=1).alignment = Alignment(horizontal='left', vertical='center', indent=2)
        row_h(ws, 32, 22)

        vd = kpis['by_vendeur']
        max_ca = vd['ca'].max() if len(vd) > 0 else 1
        total_ca = vd['ca'].sum()

        hdrs2 = ['#', 'Vendeur', 'CA TTC (€)', 'Nb Cmd', 'Panier Moy (€)', 'Remise Moy', '% Total CA']
        hdr_c2 = [1, 3, 9, 14, 17, 21, 24]
        hdr_s2 = [2, 6, 5, 3, 4, 3, 3]
        for hdr, sc_s, sp in zip(hdrs2, hdr_c2, hdr_s2):
            merge_sc(ws, 33, sc_s, 33, sc_s + sp - 1,
                     val=hdr, bg=C['blue'], fg='FFFFFF', size=9, bold=True, h='center')
        row_h(ws, 33, 18)

        medal = {0: '🥇', 1: '🥈', 2: '🥉'}
        r = 34
        for i, (vend, vrow) in enumerate(vd.iterrows()):
            bg = [C['yellow_l'], C['blue_l'], C['g1']][min(i, 2)] if i < 3 else (C['row_alt'] if i % 2 == 0 else C['white'])
            rank = f"{medal.get(i, '')} {i+1}" if i < 3 else str(i + 1)
            ca = vrow.get('ca', 0)
            nb = vrow.get('nb', 0)
            moy = vrow.get('moy', 0)
            remise = vrow.get('remise_moy', None)
            pct = ca / total_ca * 100 if total_ca else 0

            merge_sc(ws, r, 1, r, 2, val=rank, bg=bg, fg=C['dark'], size=10, bold=(i<3), h='center')
            merge_sc(ws, r, 3, r, 8, val=str(vend), bg=bg, fg=C['dark'], size=10, bold=(i<3), h='left')
            ws.cell(row=r, column=3).alignment = Alignment(horizontal='left', vertical='center', indent=1)
            merge_sc(ws, r, 9, r, 13, val=round(float(ca), 2), bg=bg, fg=C['dark'],
                     size=10, bold=(i<3), h='right')
            ws.cell(row=r, column=9).number_format = '#,##0.00'
            merge_sc(ws, r, 14, r, 16, val=int(nb), bg=bg, fg=C['dark'], size=10, h='center')
            merge_sc(ws, r, 17, r, 20, val=round(float(moy), 2), bg=bg, fg=C['dark'],
                     size=9, h='right')
            ws.cell(row=r, column=17).number_format = '#,##0.00'
            rem_txt = fmt_pct(float(remise)) if remise is not None and not math.isnan(float(remise)) else '—'
            merge_sc(ws, r, 21, r, 23, val=rem_txt, bg=bg, fg=C['dark'], size=9, h='center')
            merge_sc(ws, r, 24, r, 26, val=f'{pct:.1f}%', bg=bg, fg=C['dark'], size=9, h='center')
            row_h(ws, r, 18)
            r += 1
            if r > 44:
                break

    return ws

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 2 : ÉVOLUTION MENSUELLE
# ══════════════════════════════════════════════════════════════════════════════
def build_evolution_sheet(wb, kpis, cm):
    ws = wb.create_sheet('📈 Évolution')
    ws.sheet_view.showGridLines = False
    ws.sheet_view.showRowColHeaders = False

    for i in range(1, 22):
        col_w(ws, i, 11)

    # Header
    for c in range(1, 22):
        ws.cell(row=1, column=c).fill = _fill(C['hdr'])
    merge_sc(ws, 1, 1, 1, 21, val='📈  ANALYSE DE L\'ÉVOLUTION MENSUELLE',
             bg=C['hdr'], fg='FFFFFF', size=14, bold=True, h='left')
    ws.cell(row=1, column=1).alignment = Alignment(horizontal='left', vertical='center', indent=2)
    for c in range(1, 22):
        ws.cell(row=2, column=c).fill = _fill(C['blue'])
    row_h(ws, 1, 26); row_h(ws, 2, 4); row_h(ws, 3, 10)

    if 'monthly' not in kpis:
        merge_sc(ws, 4, 1, 4, 10, val='Pas de colonne date détectée.',
                 bg=C['g1'], fg=C['g5'], size=10)
        return ws

    monthly = kpis['monthly']

    # ── Tableau mensuel détaillé ─────────────────────────────────────────────
    merge_sc(ws, 4, 1, 4, 21, val='📅  TABLEAU MENSUEL DÉTAILLÉ',
             bg=C['hdr2'], fg='FFFFFF', size=11, bold=True, h='left')
    ws.cell(row=4, column=1).alignment = Alignment(horizontal='left', vertical='center', indent=2)
    row_h(ws, 4, 22)

    hdrs = ['Période', 'CA TTC (€)', 'Nb Cmd', 'Panier Moy (€)',
            'CA Cumulé (€)', 'Croissance', '% du Total']
    for c_idx, hdr in enumerate(hdrs, 1):
        cell = ws.cell(row=5, column=c_idx, value=hdr)
        sc(cell, bg=C['blue'], fg='FFFFFF', size=9, bold=True, h='center', border=False)
    row_h(ws, 5, 18)

    total_ca = monthly['ca'].sum()
    ca_cum = 0
    prev_ca = None
    r = 6
    for i, row_data in monthly.iterrows():
        bg = C['row_alt'] if i % 2 == 0 else C['white']
        ca = float(row_data['ca'])
        nb = int(row_data['nb'])
        ca_cum += ca
        moy = ca / nb if nb else 0
        croiss = ((ca - prev_ca) / prev_ca * 100) if prev_ca and prev_ca > 0 else None
        pct = ca / total_ca * 100 if total_ca else 0

        vals = [str(row_data['mois_str']), ca, nb, round(moy, 2),
                round(ca_cum, 2), round(croiss, 1) if croiss is not None else '—',
                round(pct, 1)]
        fmts = [None, '#,##0.00', None, '#,##0.00', '#,##0.00', '0.0"%"', '0.0"%"']
        for c_idx, (v, fmt) in enumerate(zip(vals, fmts), 1):
            cell = ws.cell(row=r, column=c_idx, value=v)
            sc(cell, bg=bg, fg=C['dark'], size=9,
               h='right' if c_idx > 1 else 'left', border=False)
            if fmt and isinstance(v, (int, float)):
                cell.number_format = fmt
            # Couleur croissance
            if c_idx == 6 and isinstance(v, float):
                cell.font = _font(size=9, bold=True,
                                  color=C['green'] if v >= 0 else C['red'])
        row_h(ws, r, 16)
        r += 1
        prev_ca = ca

    # Ligne total
    for c_idx, v in enumerate([f'TOTAL ({len(monthly)} mois)',
                                round(total_ca, 2), int(monthly['nb'].sum()),
                                round(total_ca / max(monthly['nb'].sum(), 1), 2),
                                '', '', '100.0'], 1):
        cell = ws.cell(row=r, column=c_idx, value=v)
        sc(cell, bg=C['hdr2'], fg='FFFFFF', size=9, bold=True,
           h='right' if c_idx > 1 else 'left', border=False)
    row_h(ws, r, 18)

    # ── Data pour graphiques ─────────────────────────────────────────────────
    # Stocker les données mensuelles dans les colonnes 9-11 (pour référence chart)
    ws.cell(row=5, column=9).value  = 'Mois'
    ws.cell(row=5, column=10).value = 'CA'
    ws.cell(row=5, column=11).value = 'NbCmd'
    for i, row_data in enumerate(monthly.itertuples(), 6):
        ws.cell(row=i, column=9).value  = row_data.mois_str
        ws.cell(row=i, column=10).value = round(float(row_data.ca), 2)
        ws.cell(row=i, column=11).value = int(row_data.nb)

    n_months = len(monthly)

    # ── Graphique 1 : Évolution CA (LineChart) ───────────────────────────────
    chart1 = LineChart()
    chart1.title = 'Évolution du CA TTC Mensuel'
    chart1.style = 10
    chart1.grouping = 'standard'
    chart1.height = 14
    chart1.width  = 22
    chart1.y_axis.title = 'CA TTC (€)'
    chart1.x_axis.title = 'Mois'
    chart1.y_axis.numFmt = '#,##0'

    data_ref = Reference(ws, min_col=10, max_col=10, min_row=5, max_row=5 + n_months)
    chart1.add_data(data_ref, titles_from_data=True)
    cats_ref = Reference(ws, min_col=9, min_row=6, max_row=5 + n_months)
    chart1.set_categories(cats_ref)

    s = chart1.series[0]
    s.graphicalProperties.line.solidFill = C['blue']
    s.graphicalProperties.line.width = 25000
    s.marker.symbol = 'circle'
    s.marker.size = 6
    s.marker.graphicalProperties.fgColor = C['blue']
    s.dLbls = DataLabelList()
    s.dLbls.showVal = True
    s.dLbls.showLegendKey = False

    ws.add_chart(chart1, 'A' + str(r + 3))

    # ── Graphique 2 : Nb commandes (BarChart) ────────────────────────────────
    chart2 = BarChart()
    chart2.type = 'col'
    chart2.title = 'Nombre de Commandes par Mois'
    chart2.style = 10
    chart2.height = 14
    chart2.width  = 20
    chart2.y_axis.title = 'Nb Commandes'

    data_ref2 = Reference(ws, min_col=11, max_col=11, min_row=5, max_row=5 + n_months)
    chart2.add_data(data_ref2, titles_from_data=True)
    chart2.set_categories(cats_ref)
    s2 = chart2.series[0]
    s2.graphicalProperties.solidFill = C['teal']
    s2.dLbls = DataLabelList()
    s2.dLbls.showVal = True
    s2.dLbls.showLegendKey = False

    ws.add_chart(chart2, 'L' + str(r + 3))

    return ws


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 3 : PERFORMANCE VENDEURS & CATÉGORIES
# ══════════════════════════════════════════════════════════════════════════════
def build_performance_sheet(wb, kpis, cm):
    ws = wb.create_sheet('🏆 Performance')
    ws.sheet_view.showGridLines = False
    ws.sheet_view.showRowColHeaders = False

    for i in range(1, 26):
        col_w(ws, i, 9)

    # Header
    for c in range(1, 26):
        ws.cell(row=1, column=c).fill = _fill(C['hdr'])
    merge_sc(ws, 1, 1, 1, 25, val='🏆  PERFORMANCE VENDEURS & ANALYSE PAR CATÉGORIE',
             bg=C['hdr'], fg='FFFFFF', size=14, bold=True, h='left')
    ws.cell(row=1, column=1).alignment = Alignment(horizontal='left', vertical='center', indent=2)
    for c in range(1, 26):
        ws.cell(row=2, column=c).fill = _fill(C['blue'])
    row_h(ws, 1, 26); row_h(ws, 2, 4); row_h(ws, 3, 10)

    # ── Tableau vendeurs ─────────────────────────────────────────────────────
    r = 4
    if 'by_vendeur' in kpis:
        merge_sc(ws, r, 1, r, 25, val='👤  CLASSEMENT COMPLET DES VENDEURS',
                 bg=C['hdr2'], fg='FFFFFF', size=11, bold=True, h='left')
        ws.cell(row=r, column=1).alignment = Alignment(horizontal='left', vertical='center', indent=2)
        row_h(ws, r, 22); r += 1

        hdrs = ['Rang', 'Vendeur', 'CA TTC (€)', '% Total', 'Nb Cmd', 'Panier Moy (€)', 'Remise Moy (%)']
        col_starts = [1, 3, 9, 14, 17, 20, 24]
        col_spans =  [2, 6, 5, 3, 3, 4, 2]
        for hdr, cs, sp in zip(hdrs, col_starts, col_spans):
            merge_sc(ws, r, cs, r, cs + sp - 1,
                     val=hdr, bg=C['blue2'], fg='FFFFFF', size=9, bold=True, h='center')
        row_h(ws, r, 18); r += 1

        vd = kpis['by_vendeur']
        total_ca = vd['ca'].sum()
        medal = {0: '🥇', 1: '🥈', 2: '🥉'}

        # Data cols pour graphique (col 14+)
        ws.cell(row=r-1, column=14).value = None
        chart_row_start_vd = r

        for i, (vend, vrow) in enumerate(vd.iterrows()):
            bg = (C['yellow_l'] if i == 0 else C['g1'] if i == 1 else
                  C['row_alt'] if i % 2 == 0 else C['white'])
            ca = float(vrow['ca'])
            nb = int(vrow['nb'])
            moy = float(vrow['moy'])
            remise = vrow.get('remise_moy', None)
            pct = ca / total_ca * 100 if total_ca else 0

            rank_str = f"{medal.get(i, '')} {i+1}"
            vals_with_pos = [
                (rank_str, 1, 2), (str(vend), 3, 8),
                (round(ca, 2), 9, 13), (round(pct, 1), 14, 16),
                (nb, 17, 19), (round(moy, 2), 20, 23),
                (round(float(remise), 1) if remise is not None and not math.isnan(float(remise)) else '—', 24, 25),
            ]
            for val, cs, ce in vals_with_pos:
                merge_sc(ws, r, cs, r, ce, val=val, bg=bg, fg=C['dark'],
                         size=9, bold=(i < 3), h='right' if cs > 2 else ('center' if cs == 1 else 'left'))
                if cs == 3:
                    ws.cell(row=r, column=3).alignment = Alignment(horizontal='left', vertical='center', indent=1)
            ws.cell(row=r, column=9).number_format = '#,##0.00'
            ws.cell(row=r, column=20).number_format = '#,##0.00'
            row_h(ws, r, 17); r += 1

        row_h(ws, r, 12); r += 1

    # ── Tableau catégories ───────────────────────────────────────────────────
    cat_table_row = r
    if 'by_categorie' in kpis:
        merge_sc(ws, r, 1, r, 25, val='🏷️  ANALYSE PAR CATÉGORIE DE PRODUIT',
                 bg=C['hdr2'], fg='FFFFFF', size=11, bold=True, h='left')
        ws.cell(row=r, column=1).alignment = Alignment(horizontal='left', vertical='center', indent=2)
        row_h(ws, r, 22); r += 1

        hdrs2 = ['Catégorie', 'CA TTC (€)', '% Total', 'Nb Cmd', 'Panier Moy (€)', 'Quantité', 'Remise Moy (%)']
        col_s2 = [1, 6, 11, 14, 17, 21, 24]
        col_sp2 = [5, 5, 3, 3, 4, 3, 2]
        for hdr, cs, sp in zip(hdrs2, col_s2, col_sp2):
            merge_sc(ws, r, cs, r, cs + sp - 1,
                     val=hdr, bg=C['teal'], fg='FFFFFF', size=9, bold=True, h='center')
        row_h(ws, r, 18); r += 1

        cat_df = kpis['by_categorie']
        total_ca_c = cat_df['ca'].sum()
        chart_row_start_cat = r

        for i, (cat_name, crow) in enumerate(cat_df.iterrows()):
            bg = C['teal_l'] if i == 0 else (C['row_alt'] if i % 2 == 0 else C['white'])
            ca = float(crow['ca'])
            nb = int(crow['nb'])
            moy = float(crow['moy'])
            qty = int(crow['qty']) if 'qty' in crow.index else '—'
            remise = crow.get('remise_moy', None)
            pct = ca / total_ca_c * 100 if total_ca_c else 0

            vals2 = [
                (str(cat_name), 1, 5), (round(ca, 2), 6, 10),
                (round(pct, 1), 11, 13), (nb, 14, 16),
                (round(moy, 2), 17, 20), (qty, 21, 23),
                (round(float(remise), 1) if remise is not None and not math.isnan(float(remise)) else '—', 24, 25),
            ]
            for val, cs, ce in vals2:
                merge_sc(ws, r, cs, r, ce, val=val, bg=bg, fg=C['dark'],
                         size=9, bold=(i == 0), h='left' if cs == 1 else 'right')
                if cs == 1:
                    ws.cell(row=r, column=1).alignment = Alignment(horizontal='left', vertical='center', indent=1)
            ws.cell(row=r, column=6).number_format = '#,##0.00'
            ws.cell(row=r, column=17).number_format = '#,##0.00'
            row_h(ws, r, 17); r += 1

    # ── GRAPHIQUES (col 14+ zone, après les tableaux) ────────────────────────
    # Préparer data vendeurs dans cols cachées pour graphiques
    data_start = r + 2

    # Data vendeurs
    if 'by_vendeur' in kpis:
        vd = kpis['by_vendeur'].head(8)
        ws.cell(row=data_start, column=1).value = 'Vendeur'
        ws.cell(row=data_start, column=2).value = 'CA TTC'
        for i, (vend, vrow) in enumerate(vd.iterrows(), 1):
            ws.cell(row=data_start + i, column=1).value = str(vend)
            ws.cell(row=data_start + i, column=2).value = round(float(vrow['ca']), 2)

        chart_v = BarChart()
        chart_v.type = 'bar'
        chart_v.style = 10
        chart_v.title = 'CA par Vendeur (Top 8)'
        chart_v.height = 16
        chart_v.width  = 22
        chart_v.y_axis.title = 'CA TTC (€)'
        chart_v.y_axis.numFmt = '#,##0'

        data_ref_v = Reference(ws, min_col=2, max_col=2,
                               min_row=data_start, max_row=data_start + len(vd))
        chart_v.add_data(data_ref_v, titles_from_data=True)
        cats_ref_v = Reference(ws, min_col=1, min_row=data_start + 1,
                               max_row=data_start + len(vd))
        chart_v.set_categories(cats_ref_v)
        s_v = chart_v.series[0]
        s_v.graphicalProperties.solidFill = C['blue']
        s_v.dLbls = DataLabelList()
        s_v.dLbls.showVal = True
        s_v.dLbls.showLegendKey = False
        ws.add_chart(chart_v, f'E{data_start}')

    # Data catégories
    if 'by_categorie' in kpis:
        cat_d = kpis['by_categorie'].head(8)
        d2s = data_start + len(kpis.get('by_vendeur', pd.DataFrame())) + 4
        ws.cell(row=d2s, column=1).value = 'Catégorie'
        ws.cell(row=d2s, column=2).value = 'CA TTC'
        for i, (cat, crow) in enumerate(cat_d.iterrows(), 1):
            ws.cell(row=d2s + i, column=1).value = str(cat)
            ws.cell(row=d2s + i, column=2).value = round(float(crow['ca']), 2)

        chart_c = BarChart()
        chart_c.type = 'col'
        chart_c.style = 10
        chart_c.title = 'CA par Catégorie'
        chart_c.height = 16
        chart_c.width  = 22
        chart_c.y_axis.numFmt = '#,##0'

        data_ref_c = Reference(ws, min_col=2, max_col=2,
                               min_row=d2s, max_row=d2s + len(cat_d))
        chart_c.add_data(data_ref_c, titles_from_data=True)
        cats_ref_c = Reference(ws, min_col=1, min_row=d2s + 1, max_row=d2s + len(cat_d))
        chart_c.set_categories(cats_ref_c)
        s_c = chart_c.series[0]
        s_c.graphicalProperties.solidFill = C['teal']
        s_c.dLbls = DataLabelList()
        s_c.dLbls.showVal = True
        s_c.dLbls.showLegendKey = False
        ws.add_chart(chart_c, f'M{data_start}')

    return ws

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 4 : ANALYSE CLIENTS & PRODUITS
# ══════════════════════════════════════════════════════════════════════════════
def build_analyse_sheet(wb, df, kpis, cm):
    ws = wb.create_sheet('🔍 Analyse')
    ws.sheet_view.showGridLines = False
    ws.sheet_view.showRowColHeaders = False

    for i in range(1, 24):
        col_w(ws, i, 10)

    # Header
    for c in range(1, 24):
        ws.cell(row=1, column=c).fill = _fill(C['hdr'])
    merge_sc(ws, 1, 1, 1, 23, val='🔍  ANALYSE CLIENTS, PRODUITS & STATISTIQUES',
             bg=C['hdr'], fg='FFFFFF', size=14, bold=True, h='left')
    ws.cell(row=1, column=1).alignment = Alignment(horizontal='left', vertical='center', indent=2)
    for c in range(1, 24):
        ws.cell(row=2, column=c).fill = _fill(C['blue'])
    row_h(ws, 1, 26); row_h(ws, 2, 4); row_h(ws, 3, 10)

    r = 4

    # ── TOP CLIENTS ──────────────────────────────────────────────────────────
    if 'top_clients' in kpis:
        merge_sc(ws, r, 1, r, 23, val='👥  TOP CLIENTS (par CA TTC)',
                 bg=C['hdr2'], fg='FFFFFF', size=11, bold=True, h='left')
        ws.cell(row=r, column=1).alignment = Alignment(horizontal='left', vertical='center', indent=2)
        row_h(ws, r, 22); r += 1

        hdrs = ['Rang', 'Client', 'CA TTC (€)', '% du Total']
        col_s = [1, 3, 16, 21]
        col_sp = [2, 13, 5, 3]
        for hdr, cs, sp in zip(hdrs, col_s, col_sp):
            merge_sc(ws, r, cs, r, cs + sp - 1,
                     val=hdr, bg=C['purple'], fg='FFFFFF', size=9, bold=True, h='center')
        row_h(ws, r, 18); r += 1

        top_c = kpis['top_clients']
        total_ca = float(top_c.sum())
        chart_start_cli = r
        for i, (cli, ca) in enumerate(top_c.items()):
            bg = C['purple_l'] if i == 0 else (C['row_alt'] if i % 2 == 0 else C['white'])
            pct = float(ca) / total_ca * 100 if total_ca else 0
            merge_sc(ws, r, 1, r, 2, val=f'{i+1}', bg=bg, fg=C['dark'], size=10, bold=(i==0), h='center')
            merge_sc(ws, r, 3, r, 15, val=str(cli), bg=bg, fg=C['dark'], size=10, bold=(i==0), h='left')
            ws.cell(row=r, column=3).alignment = Alignment(horizontal='left', vertical='center', indent=1)
            merge_sc(ws, r, 16, r, 20, val=round(float(ca), 2), bg=bg, fg=C['dark'],
                     size=10, bold=(i==0), h='right')
            ws.cell(row=r, column=16).number_format = '#,##0.00'
            merge_sc(ws, r, 21, r, 23, val=f'{pct:.1f}%', bg=bg, fg=C['dark'], size=10, h='center')
            row_h(ws, r, 17); r += 1

        row_h(ws, r, 12); r += 1

    # ── TOP PRODUITS ─────────────────────────────────────────────────────────
    if 'top_produits' in kpis:
        merge_sc(ws, r, 1, r, 23, val='📦  TOP PRODUITS (par CA TTC)',
                 bg=C['hdr2'], fg='FFFFFF', size=11, bold=True, h='left')
        ws.cell(row=r, column=1).alignment = Alignment(horizontal='left', vertical='center', indent=2)
        row_h(ws, r, 22); r += 1

        hdrs2 = ['Rang', 'Produit', 'CA TTC (€)', 'Nb Cmd', '% du Total']
        col_s2 = [1, 3, 14, 19, 21]
        col_sp2 = [2, 11, 5, 2, 3]
        for hdr, cs, sp in zip(hdrs2, col_s2, col_sp2):
            merge_sc(ws, r, cs, r, cs + sp - 1,
                     val=hdr, bg=C['orange'], fg='FFFFFF', size=9, bold=True, h='center')
        row_h(ws, r, 18); r += 1

        top_p = kpis['top_produits']
        total_ca_p = float(top_p.sum())
        for i, (prod, ca) in enumerate(top_p.items()):
            bg = C['orange_l'] if i == 0 else (C['row_alt'] if i % 2 == 0 else C['white'])
            pct = float(ca) / total_ca_p * 100 if total_ca_p else 0
            # Nb commandes pour ce produit
            nb_cmd = 0
            if cm.get('produit') and kpis.get('ca_col'):
                try:
                    nb_cmd = int(df[df[cm['produit']] == prod].shape[0])
                except Exception:
                    nb_cmd = 0

            merge_sc(ws, r, 1, r, 2, val=f'{i+1}', bg=bg, fg=C['dark'], size=10, bold=(i==0), h='center')
            merge_sc(ws, r, 3, r, 13, val=str(prod), bg=bg, fg=C['dark'], size=9, bold=(i==0), h='left')
            ws.cell(row=r, column=3).alignment = Alignment(horizontal='left', vertical='center', indent=1)
            merge_sc(ws, r, 14, r, 18, val=round(float(ca), 2), bg=bg, fg=C['dark'],
                     size=10, bold=(i==0), h='right')
            ws.cell(row=r, column=14).number_format = '#,##0.00'
            merge_sc(ws, r, 19, r, 20, val=nb_cmd if nb_cmd else '—', bg=bg, fg=C['dark'], size=9, h='center')
            merge_sc(ws, r, 21, r, 23, val=f'{pct:.1f}%', bg=bg, fg=C['dark'], size=9, h='center')
            row_h(ws, r, 17); r += 1

        row_h(ws, r, 12); r += 1

    # ── STATISTIQUES DESCRIPTIVES ─────────────────────────────────────────────
    num_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    if num_cols:
        merge_sc(ws, r, 1, r, 23, val='📐  STATISTIQUES DESCRIPTIVES',
                 bg=C['hdr2'], fg='FFFFFF', size=11, bold=True, h='left')
        ws.cell(row=r, column=1).alignment = Alignment(horizontal='left', vertical='center', indent=2)
        row_h(ws, r, 22); r += 1

        hdrs3 = ['Colonne', 'Total', 'Moyenne', 'Médiane', 'Min', 'Max', 'Écart-type', '% Nulls']
        col_s3 = [1, 5, 8, 11, 13, 15, 17, 21]
        col_sp3 = [4, 3, 3, 2, 2, 2, 4, 3]
        for hdr, cs, sp in zip(hdrs3, col_s3, col_sp3):
            merge_sc(ws, r, cs, r, cs + sp - 1,
                     val=hdr, bg=C['green'], fg='FFFFFF', size=9, bold=True, h='center')
        row_h(ws, r, 18); r += 1

        for i, nc in enumerate(num_cols[:10]):
            bg = C['green_l'] if i == 0 else (C['row_alt'] if i % 2 == 0 else C['white'])
            s = df[nc].dropna()
            if len(s) == 0:
                continue
            pct_null = df[nc].isnull().sum() / len(df) * 100
            vals3 = [nc, round(float(s.sum()),2), round(float(s.mean()),2),
                     round(float(s.median()),2), round(float(s.min()),2),
                     round(float(s.max()),2), round(float(s.std()),2),
                     round(pct_null, 1)]
            poss3 = list(zip(col_s3, col_sp3))
            for j, (v, (cs, sp)) in enumerate(zip(vals3, poss3)):
                merge_sc(ws, r, cs, r, cs + sp - 1, val=v, bg=bg, fg=C['dark'],
                         size=9, h='left' if j == 0 else 'right')
                if j > 0 and isinstance(v, float):
                    ws.cell(row=r, column=cs).number_format = '#,##0.00'
            row_h(ws, r, 16); r += 1

    # ── Graphique Top clients & Top produits ─────────────────────────────────
    data_row = r + 2
    if 'top_clients' in kpis:
        top_c = kpis['top_clients'].head(8)
        ws.cell(row=data_row, column=1).value = 'Client'
        ws.cell(row=data_row, column=2).value = 'CA TTC'
        for i, (cli, ca) in enumerate(top_c.items(), 1):
            ws.cell(row=data_row + i, column=1).value = str(cli)[:20]
            ws.cell(row=data_row + i, column=2).value = round(float(ca), 2)

        chart_cli = BarChart()
        chart_cli.type = 'bar'
        chart_cli.style = 10
        chart_cli.title = 'Top Clients par CA TTC'
        chart_cli.height = 14; chart_cli.width = 20
        chart_cli.y_axis.numFmt = '#,##0'

        dr = Reference(ws, min_col=2, max_col=2, min_row=data_row, max_row=data_row + len(top_c))
        chart_cli.add_data(dr, titles_from_data=True)
        cr = Reference(ws, min_col=1, min_row=data_row + 1, max_row=data_row + len(top_c))
        chart_cli.set_categories(cr)
        chart_cli.series[0].graphicalProperties.solidFill = C['purple']
        chart_cli.series[0].dLbls = DataLabelList()
        chart_cli.series[0].dLbls.showVal = True
        chart_cli.series[0].dLbls.showLegendKey = False
        ws.add_chart(chart_cli, f'A{data_row + len(top_c) + 2}')

    if 'top_produits' in kpis:
        top_p = kpis['top_produits'].head(8)
        d2 = data_row
        ws.cell(row=d2, column=4).value = 'Produit'
        ws.cell(row=d2, column=5).value = 'CA TTC'
        for i, (prod, ca) in enumerate(top_p.items(), 1):
            ws.cell(row=d2 + i, column=4).value = str(prod)[:22]
            ws.cell(row=d2 + i, column=5).value = round(float(ca), 2)

        chart_p = BarChart()
        chart_p.type = 'bar'
        chart_p.style = 10
        chart_p.title = 'Top Produits par CA TTC'
        chart_p.height = 14; chart_p.width = 20
        chart_p.y_axis.numFmt = '#,##0'

        dr2 = Reference(ws, min_col=5, max_col=5, min_row=d2, max_row=d2 + len(top_p))
        chart_p.add_data(dr2, titles_from_data=True)
        cr2 = Reference(ws, min_col=4, min_row=d2 + 1, max_row=d2 + len(top_p))
        chart_p.set_categories(cr2)
        chart_p.series[0].graphicalProperties.solidFill = C['orange']
        chart_p.series[0].dLbls = DataLabelList()
        chart_p.series[0].dLbls.showVal = True
        chart_p.series[0].dLbls.showLegendKey = False
        ws.add_chart(chart_p, f'L{data_row + len(top_p) + 2}')

    return ws

# ══════════════════════════════════════════════════════════════════════════════
# GRAPHIQUES DASHBOARD (ajout après construction du sheet)
# ══════════════════════════════════════════════════════════════════════════════
def add_dashboard_charts(wb, ws_dash, kpis):
    """Ajoute les graphiques sur le Dashboard principal."""
    ws_dash.sheet_view.showGridLines = False

    # Trouver la dernière ligne utilisée
    max_r = ws_dash.max_row + 2

    # ── Préparer données pour les graphiques ─────────────────────────────────
    data_start = max_r + 1

    # Statut data (pour PieChart)
    pie_start = None
    if 'statut_counts' in kpis:
        ws_dash.cell(row=data_start, column=1).value = 'Statut'
        ws_dash.cell(row=data_start, column=2).value = 'Nb'
        sc_dict = kpis['statut_counts']
        for i, (stat, cnt) in enumerate(sc_dict.items(), 1):
            ws_dash.cell(row=data_start + i, column=1).value = str(stat)
            ws_dash.cell(row=data_start + i, column=2).value = int(cnt)
        pie_start = data_start
        pie_end = data_start + len(sc_dict)

        # PieChart statut
        chart_pie = PieChart()
        chart_pie.title = 'Répartition des Commandes par Statut'
        chart_pie.style = 10
        chart_pie.height = 14
        chart_pie.width  = 16

        data_ref_p = Reference(ws_dash, min_col=2, max_col=2,
                               min_row=pie_start, max_row=pie_end)
        chart_pie.add_data(data_ref_p, titles_from_data=True)
        cats_ref_p = Reference(ws_dash, min_col=1,
                               min_row=pie_start + 1, max_row=pie_end)
        chart_pie.set_categories(cats_ref_p)
        chart_pie.dataLabels = DataLabelList()
        chart_pie.dataLabels.showPercent = True
        chart_pie.dataLabels.showCatName = True
        chart_pie.dataLabels.showVal = False
        chart_pie.dataLabels.showLegendKey = False

        # Couleurs des tranches
        status_fill_colors = ['107C10', 'D97706', 'DC2626', '0078D4', '7C3AED']
        for i, color in enumerate(status_fill_colors[:len(sc_dict)]):
            pt = DataPoint(idx=i)
            pt.graphicalProperties.solidFill = color
            chart_pie.series[0].dPt.append(pt)

        ws_dash.add_chart(chart_pie, f'N46')

    # Évolution mensuelle pour le dashboard
    if 'monthly' in kpis:
        monthly = kpis['monthly']
        m_start = data_start + (len(kpis.get('statut_counts', {})) + 3)
        ws_dash.cell(row=m_start, column=1).value = 'Mois'
        ws_dash.cell(row=m_start, column=2).value = 'CA TTC'
        for i, row_data in enumerate(monthly.itertuples(), 1):
            ws_dash.cell(row=m_start + i, column=1).value = row_data.mois_str
            ws_dash.cell(row=m_start + i, column=2).value = round(float(row_data.ca), 2)
        m_end = m_start + len(monthly)

        chart_line = LineChart()
        chart_line.title = 'Évolution Mensuelle du CA TTC'
        chart_line.style = 10
        chart_line.grouping = 'standard'
        chart_line.height = 14
        chart_line.width  = 28
        chart_line.y_axis.title = 'CA TTC (€)'
        chart_line.y_axis.numFmt = '#,##0'

        data_ref_l = Reference(ws_dash, min_col=2, max_col=2, min_row=m_start, max_row=m_end)
        chart_line.add_data(data_ref_l, titles_from_data=True)
        cats_ref_l = Reference(ws_dash, min_col=1, min_row=m_start + 1, max_row=m_end)
        chart_line.set_categories(cats_ref_l)

        sl = chart_line.series[0]
        sl.graphicalProperties.line.solidFill = C['blue']
        sl.graphicalProperties.line.width = 25000
        sl.marker.symbol = 'circle'
        sl.marker.size = 7
        sl.marker.graphicalProperties.fgColor = C['blue']
        sl.dLbls = DataLabelList()
        sl.dLbls.showVal = True
        sl.dLbls.showLegendKey = False

        ws_dash.add_chart(chart_line, 'A46')

    # Catégorie bar chart sur dashboard
    if 'by_categorie' in kpis:
        cat_df = kpis['by_categorie'].head(6)
        c_start = data_start + (len(kpis.get('statut_counts', {})) + 3 +
                                len(kpis.get('monthly', pd.DataFrame())) + 4)
        ws_dash.cell(row=c_start, column=1).value = 'Catégorie'
        ws_dash.cell(row=c_start, column=2).value = 'CA TTC'
        for i, (cat, crow) in enumerate(cat_df.iterrows(), 1):
            ws_dash.cell(row=c_start + i, column=1).value = str(cat)
            ws_dash.cell(row=c_start + i, column=2).value = round(float(crow['ca']), 2)
        c_end = c_start + len(cat_df)

        chart_cat = BarChart()
        chart_cat.type = 'col'
        chart_cat.style = 10
        chart_cat.title = 'CA TTC par Catégorie'
        chart_cat.height = 14
        chart_cat.width  = 22
        chart_cat.y_axis.numFmt = '#,##0'

        data_ref_c = Reference(ws_dash, min_col=2, max_col=2,
                               min_row=c_start, max_row=c_end)
        chart_cat.add_data(data_ref_c, titles_from_data=True)
        cats_ref_c = Reference(ws_dash, min_col=1, min_row=c_start + 1, max_row=c_end)
        chart_cat.set_categories(cats_ref_c)

        # Couleurs différentes par barre
        for i, color in enumerate(CHART_PAL[:len(cat_df)]):
            pt = DataPoint(idx=i)
            pt.graphicalProperties.solidFill = color
            chart_cat.series[0].dPt.append(pt)
        chart_cat.series[0].dLbls = DataLabelList()
        chart_cat.series[0].dLbls.showVal = True
        chart_cat.series[0].dLbls.showLegendKey = False

        ws_dash.add_chart(chart_cat, 'A64')

    # Marque PieChart
    if 'by_marque' in kpis:
        marque_d = kpis['by_marque'].head(8)
        mk_start = data_start + 60
        ws_dash.cell(row=mk_start, column=1).value = 'Marque'
        ws_dash.cell(row=mk_start, column=2).value = 'CA TTC'
        for i, (m, ca) in enumerate(marque_d.items(), 1):
            ws_dash.cell(row=mk_start + i, column=1).value = str(m)
            ws_dash.cell(row=mk_start + i, column=2).value = round(float(ca), 2)
        mk_end = mk_start + len(marque_d)

        chart_mk = PieChart()
        chart_mk.title = 'Répartition CA par Marque'
        chart_mk.style = 10
        chart_mk.height = 14
        chart_mk.width  = 16

        dr_mk = Reference(ws_dash, min_col=2, max_col=2, min_row=mk_start, max_row=mk_end)
        chart_mk.add_data(dr_mk, titles_from_data=True)
        cr_mk = Reference(ws_dash, min_col=1, min_row=mk_start + 1, max_row=mk_end)
        chart_mk.set_categories(cr_mk)
        chart_mk.dataLabels = DataLabelList()
        chart_mk.dataLabels.showPercent = True
        chart_mk.dataLabels.showCatName = True
        chart_mk.dataLabels.showVal = False
        chart_mk.dataLabels.showLegendKey = False

        for i, color in enumerate(CHART_PAL[:len(marque_d)]):
            pt = DataPoint(idx=i)
            pt.graphicalProperties.solidFill = color
            chart_mk.series[0].dPt.append(pt)

        ws_dash.add_chart(chart_mk, 'N64')

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 5 : DONNÉES BRUTES (Table Excel)
# ══════════════════════════════════════════════════════════════════════════════
def build_raw_data_sheet(wb, df, cm):
    ws = wb.create_sheet('📋 Données brutes')
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = 'A2'

    n_rows, n_cols = df.shape
    numeric_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    date_cols    = [c for c in df.columns if pd.api.types.is_datetime64_any_dtype(df[c])]

    # Header banner
    for c in range(1, n_cols + 1):
        ws.cell(row=1, column=c).fill = _fill(C['hdr'])
    merge_sc(ws, 1, 1, 1, n_cols,
             val=f'📋  DONNÉES BRUTES — {n_rows} enregistrements × {n_cols} colonnes',
             bg=C['hdr'], fg='FFFFFF', size=11, bold=True, h='left')
    ws.cell(row=1, column=1).alignment = Alignment(horizontal='left', vertical='center', indent=2)
    row_h(ws, 1, 22)

    # En-têtes colonnes
    statut_col_idx = None
    for c_idx, col_name in enumerate(df.columns, 1):
        cell = ws.cell(row=2, column=c_idx, value=col_name)
        sc(cell, bg=C['blue'], fg='FFFFFF', size=9, bold=True, h='center', border=False)
        cell.border = _border(C['hdr2'])
        if cm.get('statut') and col_name == cm['statut']:
            statut_col_idx = c_idx
    row_h(ws, 2, 20)

    # Statut colors map
    statut_colors = {
        'livré': C['green_l'], 'livre': C['green_l'],
        'en attente': C['yellow_l'],
        'annulé': C['red_l'], 'annule': C['red_l'],
    }

    # Données
    for r_idx, (_, row_data) in enumerate(df.iterrows(), 3):
        bg = C['row_alt'] if r_idx % 2 == 0 else C['white']
        row_bg = bg

        # Colorier la ligne selon le statut
        if statut_col_idx and cm.get('statut'):
            sv = str(row_data.get(cm['statut'], '')).lower()
            row_bg = statut_colors.get(sv, bg)

        for c_idx, (col_name, value) in enumerate(row_data.items(), 1):
            # Conversion types numpy
            if isinstance(value, (int,)):
                pass
            elif hasattr(value, 'item'):
                try:
                    value = value.item()
                except Exception:
                    value = str(value)
            elif isinstance(value, float) and math.isnan(value):
                value = None

            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            is_num = col_name in numeric_cols
            is_date = col_name in date_cols

            cell.fill = _fill(row_bg)
            cell.font = _font(size=9)
            cell.border = _border()
            cell.alignment = _align(h='right' if is_num else 'left', v='center')

            if is_num:
                cell.number_format = '#,##0.00'
            elif is_date:
                cell.number_format = 'DD/MM/YYYY'

            # Colorer la cellule statut
            if c_idx == statut_col_idx:
                sv = str(value).lower() if value else ''
                col_map_stat = {
                    'livré': (C['green'], C['green_l']),
                    'livre': (C['green'], C['green_l']),
                    'en attente': (C['yellow'], C['yellow_l']),
                    'annulé': (C['red'], C['red_l']),
                    'annule': (C['red'], C['red_l']),
                }
                if sv in col_map_stat:
                    fc, bc = col_map_stat[sv]
                    cell.fill = _fill(bc)
                    cell.font = _font(size=9, bold=True, color=fc)
                    cell.alignment = _align(h='center', v='center')

        row_h(ws, r_idx, 15)

    # Largeurs auto
    for c_idx in range(1, n_cols + 1):
        col_letter = get_column_letter(c_idx)
        col_name = df.columns[c_idx - 1]
        try:
            max_len = max(len(str(col_name)),
                          df.iloc[:, c_idx - 1].astype(str).str.len().max())
        except Exception:
            max_len = len(str(col_name))
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 35)

    # Mise en forme conditionnelle sur colonnes numériques
    for col_name in numeric_cols[:3]:
        c_idx = list(df.columns).index(col_name) + 1
        col_letter = get_column_letter(c_idx)
        data_range = f'{col_letter}3:{col_letter}{n_rows + 2}'
        ws.conditional_formatting.add(data_range, ColorScaleRule(
            start_type='min', start_color='FEE2E2',
            mid_type='percentile', mid_value=50, mid_color='FEF9C3',
            end_type='max', end_color='DFF6DD'
        ))

    # Table Excel
    try:
        tab = Table(displayName='TableDonnees',
                    ref=f'A2:{get_column_letter(n_cols)}{n_rows + 2}')
        tab.tableStyleInfo = TableStyleInfo(
            name='TableStyleMedium2',
            showFirstColumn=False, showLastColumn=False,
            showRowStripes=True, showColumnStripes=False)
        ws.add_table(tab)
    except Exception as e:
        logger.warning(f'Table Données brutes: {e}')

    return ws


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 6 : SOURCE POUR TCD
# ══════════════════════════════════════════════════════════════════════════════
def build_tcd_source_sheet(wb, df, cm):
    ws = wb.create_sheet('📝 Source TCD')
    ws.sheet_view.showGridLines = False

    n_rows, n_cols = df.shape
    numeric_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    date_cols    = [c for c in df.columns if pd.api.types.is_datetime64_any_dtype(df[c])]

    # Bandeau instructions
    merge_sc(ws, 1, 1, 2, n_cols,
             val='💡 INSTRUCTIONS : Cliquer dans la table → Insertion → Tableau croisé dynamique → OK — '
                 'Glisser les champs dans les zones Lignes / Colonnes / Valeurs / Filtres',
             bg=C['yellow_l'], fg=C['dark'], size=10, bold=False, h='left')
    ws.cell(row=1, column=1).alignment = Alignment(horizontal='left', vertical='center', indent=2, wrap_text=True)
    for c in range(1, n_cols + 1):
        ws.cell(row=1, column=c).fill = _fill(C['yellow_l'])
        ws.cell(row=2, column=c).fill = _fill(C['yellow_l'])
    row_h(ws, 1, 22); row_h(ws, 2, 18)

    # En-têtes table (ligne 3)
    for c_idx, col_name in enumerate(df.columns, 1):
        cell = ws.cell(row=3, column=c_idx, value=col_name)
        cell.fill = _fill(C['hdr2'])
        cell.font = _font(size=10, bold=True, color='FFFFFF')
        cell.alignment = _align(h='center', v='center')
        cell.border = _border(C['hdr'])
    row_h(ws, 3, 20)

    # Données (à partir ligne 4)
    for r_idx, (_, row_data) in enumerate(df.iterrows(), 4):
        for c_idx, (col_name, value) in enumerate(row_data.items(), 1):
            if hasattr(value, 'item'):
                try:
                    value = value.item()
                except Exception:
                    value = str(value)
            elif isinstance(value, float) and math.isnan(value):
                value = None
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            cell.font = _font(size=9)
            cell.border = _border()
            if col_name in numeric_cols:
                cell.number_format = '#,##0.00'
                cell.alignment = _align(h='right')
            elif col_name in date_cols:
                cell.number_format = 'DD/MM/YYYY'
        row_h(ws, r_idx, 14)

    # Largeurs
    for c_idx in range(1, n_cols + 1):
        col_letter = get_column_letter(c_idx)
        col_name = df.columns[c_idx - 1]
        try:
            max_len = max(len(str(col_name)),
                          df.iloc[:, c_idx - 1].astype(str).str.len().max())
        except Exception:
            max_len = len(str(col_name))
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 12), 35)

    # Table Excel
    try:
        tab = Table(displayName='SourceTCD',
                    ref=f'A3:{get_column_letter(n_cols)}{n_rows + 3}')
        tab.tableStyleInfo = TableStyleInfo(
            name='TableStyleLight1',
            showFirstColumn=False, showLastColumn=False,
            showRowStripes=True, showColumnStripes=False)
        ws.add_table(tab)
    except Exception as e:
        logger.warning(f'Table Source TCD: {e}')

    ws.freeze_panes = 'A4'
    return ws

# ══════════════════════════════════════════════════════════════════════════════
# ORCHESTRATION PRINCIPALE
# ══════════════════════════════════════════════════════════════════════════════
def generate_excel_dashboard(file_bytes, filename):
    """
    Point d'entrée : charge les données, calcule les KPIs, génère le classeur.
    Retourne (excel_bytes, kpis_dict).
    """
    logger.info(f'Traitement de {filename} ({len(file_bytes)} octets)')

    # 1) Chargement avec détection d'en-tête intelligente
    df = load_dataframe(file_bytes, filename)
    if df.empty:
        raise ValueError('Le fichier est vide ou ne contient aucune donnée valide.')
    logger.info(f'DataFrame : {df.shape[0]} lignes × {df.shape[1]} colonnes')

    # 2) Détection des rôles de colonnes
    cm = detect_columns(df)
    logger.info(f'Colonnes détectées : {cm}')

    # 3) Conversion date si détectée
    if cm['date']:
        try:
            df[cm['date']] = pd.to_datetime(df[cm['date']], errors='coerce')
        except Exception:
            pass

    # 4) Conversion numériques
    for role in ('total_ttc', 'total_ht', 'tva', 'quantite', 'prix_unit', 'remise'):
        if cm[role]:
            df[cm[role]] = pd.to_numeric(df[cm[role]], errors='coerce')

    # 5) Calcul KPIs
    kpis = compute_kpis(df, cm)
    logger.info(f'KPIs calculés : {list(kpis.keys())}')

    # 6) Construction du classeur
    wb = Workbook()
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']

    # 6a) Dashboard principal
    ws_dash = build_dashboard_sheet(wb, df, kpis, cm, filename)

    # 6b) Graphiques dashboard
    try:
        add_dashboard_charts(wb, ws_dash, kpis)
    except Exception as e:
        logger.warning(f'Graphiques dashboard : {e}')

    # 6c) Évolution mensuelle
    try:
        build_evolution_sheet(wb, kpis, cm)
    except Exception as e:
        logger.warning(f'Feuille Évolution : {e}')

    # 6d) Performance vendeurs & catégories
    try:
        build_performance_sheet(wb, kpis, cm)
    except Exception as e:
        logger.warning(f'Feuille Performance : {e}')

    # 6e) Analyse clients & produits
    try:
        build_analyse_sheet(wb, df, kpis, cm)
    except Exception as e:
        logger.warning(f'Feuille Analyse : {e}')

    # 6f) Données brutes avec Table Excel
    try:
        build_raw_data_sheet(wb, df, cm)
    except Exception as e:
        logger.warning(f'Feuille Données brutes : {e}')

    # 6g) Source TCD
    try:
        build_tcd_source_sheet(wb, df, cm)
    except Exception as e:
        logger.warning(f'Feuille Source TCD : {e}')

    # 7) Activer le Dashboard au premier plan
    wb.active = wb['📊 Dashboard']

    # 8) Propriétés du classeur
    wb.properties.title    = f'Dashboard — {filename}'
    wb.properties.subject  = 'Dashboard analytique professionnel'
    wb.properties.creator  = 'Excel Dashboard Generator v2.0'
    wb.properties.description = (
        f'Généré le {datetime.now().strftime("%d/%m/%Y %H:%M")} '
        f'à partir de {filename}'
    )

    # 9) Sérialisation
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    excel_bytes = buf.read()
    logger.info(f'Excel généré : {len(excel_bytes)} octets — {len(wb.sheetnames)} onglets')
    return excel_bytes, kpis


# ══════════════════════════════════════════════════════════════════════════════
# ROUTES FLASK
# ══════════════════════════════════════════════════════════════════════════════
@app.route('/health', methods=['GET'])
def health():
    return jsonify({
        'status': 'ok',
        'version': '2.0.0',
        'timestamp': str(datetime.now())
    }), 200


@app.route('/generate-dashboard', methods=['POST'])
def generate_dashboard():
    """
    POST JSON :
    { "filename": "data.xlsx", "file_data": "<base64>", "email": "user@example.com" }
    Retourne JSON :
    { "status": "success", "excel_base64": "<base64>", "kpis": {...} }
    """
    try:
        payload = request.get_json(force=True)
        if not payload:
            return jsonify({'status': 'error',
                            'error_message': 'Corps JSON manquant.'}), 400

        filename     = payload.get('filename', 'upload.csv')
        file_data_b64 = payload.get('file_data', '')
        if not file_data_b64:
            return jsonify({'status': 'error',
                            'error_message': 'Champ file_data manquant.'}), 400

        file_bytes = base64.b64decode(file_data_b64)
        excel_bytes, kpis = generate_excel_dashboard(file_bytes, filename)
        excel_b64 = base64.b64encode(excel_bytes).decode('utf-8')

        # Sérialiser les KPIs (enlever objets pandas non JSON-sérialisables)
        kpis_out = {}
        for k, v in kpis.items():
            try:
                if isinstance(v, (int, float, str, bool, list, dict)):
                    kpis_out[k] = v
                elif isinstance(v, pd.DataFrame):
                    kpis_out[k] = v.reset_index().to_dict(orient='records')
                elif isinstance(v, pd.Series):
                    kpis_out[k] = v.to_dict()
                else:
                    kpis_out[k] = str(v)
            except Exception:
                kpis_out[k] = str(v)

        return jsonify({
            'status': 'success',
            'excel_base64': excel_b64,
            'kpis': kpis_out,
            'filename': f'Dashboard_{filename}'
        }), 200

    except Exception as e:
        logger.error(f'Erreur génération : {e}', exc_info=True)
        return jsonify({
            'status': 'error',
            'error_message': str(e)
        }), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
