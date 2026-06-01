import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import plotly.graph_objects as go
import plotly.express as px
import io
import re
from datetime import datetime

# ================================================================
# CONFIGURAÇÃO DA PÁGINA
# ================================================================
st.set_page_config(
    page_title="DRE Sienge — Processador de Ajustes | Campesatto",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ================================================================
# CSS CUSTOMIZADO
# ================================================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');

    html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }

    .main { background: #F8F9FB; }

    .block-container { padding: 1.5rem 2rem 3rem; max-width: 1300px; }

    .header-banner {
        background: linear-gradient(135deg, #1A2744 0%, #2E4A7A 100%);
        border-radius: 12px; padding: 24px 28px; margin-bottom: 24px;
        display: flex; align-items: center; justify-content: space-between;
    }
    .header-title { color: #fff; font-size: 22px; font-weight: 600; margin: 0; }
    .header-sub { color: rgba(255,255,255,0.6); font-size: 13px; margin: 4px 0 0; }
    .header-badge {
        background: rgba(255,255,255,0.12); color: rgba(255,255,255,0.9);
        border-radius: 20px; padding: 6px 14px; font-size: 12px;
    }

    .kpi-card {
        background: #fff; border-radius: 10px; padding: 18px 20px;
        border: 1px solid #E8ECF0; height: 100%;
    }
    .kpi-label { font-size: 12px; color: #6B7280; font-weight: 500; margin-bottom: 8px; }
    .kpi-value { font-size: 26px; font-weight: 600; line-height: 1; margin-bottom: 6px; }
    .kpi-sub { font-size: 12px; color: #9CA3AF; }
    .kpi-pos { color: #059669; }
    .kpi-neg { color: #DC2626; }
    .kpi-neu { color: #1A2744; }

    .ajuste-card {
        background: #fff; border-radius: 10px; padding: 16px 18px;
        border: 1px solid #E8ECF0; margin-bottom: 10px;
    }
    .ajuste-titulo { font-size: 14px; font-weight: 600; margin-bottom: 4px; }
    .ajuste-desc { font-size: 12px; color: #6B7280; line-height: 1.5; }

    .badge {
        display: inline-block; padding: 3px 10px; border-radius: 20px;
        font-size: 11px; font-weight: 600;
    }
    .badge-obra { background: #EFF6FF; color: #1D4ED8; }
    .badge-consol { background: #F0FDF4; color: #166534; }
    .badge-anos { background: #FDF4FF; color: #7E22CE; }

    .dre-table { width: 100%; border-collapse: collapse; font-size: 13px; }
    .dre-table th {
        background: #1A2744; color: #fff;
        padding: 8px 12px; text-align: left; font-weight: 500;
    }
    .dre-table th:not(:first-child) { text-align: right; }
    .dre-table td { padding: 7px 12px; border-bottom: 1px solid #F1F5F9; }
    .dre-table td:not(:first-child) { text-align: right; font-family: 'IBM Plex Mono', monospace; font-size: 12px; }
    .dre-table tr.total td { background: #F8FAFC; font-weight: 600; border-top: 2px solid #E2E8F0; }
    .dre-table tr.destaque td { background: #F0FDF4; font-weight: 600; color: #166534; }
    .dre-table tr.ajuste td { background: #FAF5FF; color: #7E22CE; font-size: 12px; }
    .dre-table tr.secao td { background: #1A2744; color: #fff; font-weight: 600; padding: 6px 12px; }
    .val-orig { color: #D97706; }
    .val-adj { color: #7E22CE; }
    .val-corr { color: #059669; font-weight: 600; }
    .val-neg { color: #DC2626; }
    .val-pos { color: #059669; }
    .indent { padding-left: 24px; color: #6B7280; }

    .alert-box {
        background: #FFFBEB; border: 1px solid #FDE68A;
        border-radius: 8px; padding: 12px 16px;
        font-size: 13px; color: #92400E; line-height: 1.6;
        margin: 12px 0;
    }
    .info-box {
        background: #EFF6FF; border: 1px solid #BFDBFE;
        border-radius: 8px; padding: 12px 16px;
        font-size: 13px; color: #1E40AF; line-height: 1.6;
        margin: 12px 0;
    }
    .success-box {
        background: #F0FDF4; border: 1px solid #BBF7D0;
        border-radius: 8px; padding: 12px 16px;
        font-size: 13px; color: #166534; line-height: 1.6;
        margin: 12px 0;
    }

    div[data-testid="stFileUploader"] { border: 2px dashed #CBD5E1; border-radius: 10px; padding: 8px; }
    div[data-testid="stFileUploader"]:hover { border-color: #059669; }

    .stButton > button {
        background: #1A2744; color: #fff; border: none;
        border-radius: 8px; font-weight: 500; font-size: 14px;
        padding: 10px 24px; width: 100%;
        transition: background 0.2s;
    }
    .stButton > button:hover { background: #2E4A7A; color: #fff; }

    .stDownloadButton > button {
        background: #059669; color: #fff; border: none;
        border-radius: 8px; font-weight: 500; font-size: 14px;
        padding: 10px 24px; width: 100%;
    }
    .stDownloadButton > button:hover { background: #047857; color: #fff; }

    hr { border: none; border-top: 1px solid #E2E8F0; margin: 20px 0; }

    .metric-diff {
        background: #F0FDF4; border-radius: 8px; padding: 10px 14px;
        display: flex; justify-content: space-between; align-items: center;
        margin: 6px 0;
    }
    .stMetric { background: #fff; border-radius: 10px; padding: 16px; border: 1px solid #E8ECF0; }
</style>
""", unsafe_allow_html=True)


# ================================================================
# FUNÇÕES DE PROCESSAMENTO
# ================================================================

def fmt_brl(v):
    """Formata valor em R$ brasileiro"""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "—"
    v = float(v)
    abs_v = abs(v)
    s = f"{abs_v:_.2f}".replace("_", ".").replace(".", ",", 1)
    s = s[::-1].replace(",", ".", 1)[::-1]
    # Simpler: use locale-style
    s = f"{abs_v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}" if v >= 0 else f"(R$ {s})"

def fmt_brl_k(v):
    """Formata em R$ mil"""
    if v is None: return "—"
    v = float(v)
    k = abs(v) / 1000
    s = f"{k:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}k" if v >= 0 else f"(R$ {s}k)"

def fmt_m(v):
    """Formata em R$ milhões"""
    if v is None: return "—"
    v = float(v)
    m = abs(v) / 1_000_000
    s = f"{m:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}M" if v >= 0 else f"(R$ {s}M)"

def pct(v, base):
    if not base or base == 0: return "—"
    p = v / base * 100
    return f"{p:+.1f}%".replace(".", ",")




def processar_dre_mensal(wb, nome_arquivo, ajustes, da_pct):
    """
    Processa DRE consolidada de UM único mês.
    Formato: aba 'Planilha1', col A=Código, col B=Conta, col C=valor do mês.
    Exportada diretamente do Sienge como relatório mensal.
    """
    # Tentar encontrar a aba correta
    aba = None
    for sname in wb.sheetnames:
        if sname.lower() in ('planilha1', 'sheet1', 'plan1', 'dre', 'resultado'):
            aba = sname
            break
    if not aba:
        aba = wb.sheetnames[0]

    ws = wb[aba]
    rows = list(ws.iter_rows(values_only=True))

    # Detectar período (header)
    periodo = 'Mês'
    if rows and rows[0][2]:
        periodo = str(rows[0][2]).strip()

    # Mapear código → valor (col C, índice 2)
    km = {}
    for i, row in enumerate(rows, 1):
        code = ' '.join(str(row[0]).strip().split()) if row[0] else ''
        val  = row[2]
        if code and val is not None:
            try:
                km[code] = (i, float(val))
            except (TypeError, ValueError):
                pass

    def gv(code):
        return km.get(code, (None, 0.0))[1]

    # Valores principais
    rec_bruta  = gv('01.01.01')
    deducoes   = gv('01.01.02')
    irpj_ded   = gv('01.01.02.01.06')   # IRPJ nas deduções (ajuste 2)
    csll_ded   = gv('01.01.02.01.05')   # CSLL nas deduções (ajuste 2, se existir)
    outras_rec = gv('01.01.03.01.04')   # Outras Receitas (aportes)
    custos_dir = gv('02.01.01')
    custos_ind = gv('02.01.02')
    capex      = gv('02.01.02.05.04')   # Aquisição de Imobilizado (ajuste 1)
    desp_adm   = gv('03.01')
    desp_trib  = gv('03.02')
    desp_fin   = gv('03.03')
    juros      = gv('03.03.01.03')      # Juros reais
    amort      = gv('03.03.01.07')      # Amortização de principal (ajuste 3)
    outras_nop = gv('04')
    resultado  = gv('05')

    lucro_bruto = (rec_bruta + deducoes) + custos_dir + custos_ind

    # Calcular ajustes
    da_val     = abs(rec_bruta) * da_pct

    aj1 = -capex                    if ajustes[0] else 0   # reverte CAPEX (negativo → positivo)
    aj2 = -(irpj_ded + csll_ded)    if ajustes[1] else 0   # reverte IRPJ+CSLL das deduções
    aj3 = -amort                    if ajustes[2] else 0   # reverte amortização de principal
    aj4 = da_val                    if ajustes[3] else 0   # D&A estimado

    res_corr    = resultado + aj1 + aj2 + aj3 + aj4
    ebitda_orig = resultado - desp_fin - outras_nop
    ebitda_corr = res_corr  - (desp_fin - amort if ajustes[2] else desp_fin) - outras_nop + aj4

    rec_liq_orig = rec_bruta + deducoes
    rec_liq_corr = rec_bruta + (deducoes + aj2)  # deduções sem IRPJ/CSLL

    return {
        'tipo': 'mensal',
        'arquivo': nome_arquivo,
        'periodo': periodo,
        'rec_bruta': rec_bruta,
        'deducoes': deducoes,
        'rec_liq_orig': rec_liq_orig,
        'rec_liq_corr': rec_liq_corr,
        'custos_dir': custos_dir,
        'custos_ind': custos_ind,
        'capex': capex,
        'lucro_bruto': lucro_bruto,
        'desp_adm': desp_adm,
        'desp_trib': desp_trib,
        'desp_fin': desp_fin,
        'juros': juros,
        'amort': amort,
        'outras_nop': outras_nop,
        'irpj_csll': irpj_ded + csll_ded,
        'resultado_original': resultado,
        'capex_aj': aj1,
        'impostos_aj': aj2,
        'amort_aj': aj3,
        'da_val': aj4,
        'resultado_corrigido': res_corr,
        'ebitda_original': ebitda_orig,
        'ebitda_corrigido': ebitda_corr,
        'margem_ebitda_orig': ebitda_orig / rec_bruta if rec_bruta else 0,
        'margem_ebitda_corr': ebitda_corr / rec_bruta if rec_bruta else 0,
        'margem_res_orig': resultado / rec_bruta if rec_bruta else 0,
        'margem_res_corr': res_corr / rec_bruta if rec_bruta else 0,
    }

def detectar_tipo_arquivo(nome):
    n = nome.lower()
    if any(k in n for k in ['resultado', 'obras', 'abril', 'janeiro', 'fevereiro',
                              'março', 'jan', 'fev', 'mar', 'abr', 'mai', 'jun',
                              'jul', 'ago', 'set', 'out', 'nov', 'dez']):
        return 'obras'
    if any(k in n for k in ['consolidada', 'anual', 'construtora', 'campesatto']):
        if any(str(y) in n for y in range(2020, 2030)):
            return 'consolidada'
        return 'consolidada'
    return 'obras'


def encontrar_linha(rows_data, termos, exatos=False):
    """Encontra número da linha (1-based) pelo conteúdo da coluna B"""
    for i, row in enumerate(rows_data, 1):
        b = str(row[1]).strip().lower() if len(row) > 1 and row[1] else ''
        for t in termos:
            if exatos:
                if b == t.lower(): return i
            else:
                if t.lower() in b: return i
    return None


def extrair_valores_linha(rows_data, row_num, col_inicio=2, col_fim=None):
    """Extrai valores numéricos de uma linha"""
    if row_num is None or row_num > len(rows_data):
        return []
    row = rows_data[row_num - 1]
    fim = col_fim if col_fim else len(row)
    vals = []
    for v in row[col_inicio:fim]:
        try:
            vals.append(float(v) if v is not None else 0.0)
        except (TypeError, ValueError):
            vals.append(0.0)
    return vals


def processar_dre_obras(wb, nome_arquivo, ajustes, da_pct):
    """Processa arquivo de DREs de obras (múltiplas abas)"""
    ABAS_EXCLUIR = {'obras em andamento', 'fluxo de caixa', 'rental',
                    'dashboard', 'fs energia-cnp', 'sheet1', 'plan1'}

    resultados = {}
    abas_obras = [s for s in wb.sheetnames
                  if s.lower() not in ABAS_EXCLUIR]

    for sname in abas_obras:
        ws = wb[sname]
        rows = list(ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True))

        # Header de colunas (row 8)
        header_row = rows[7] if len(rows) > 7 else []
        meses = []
        total_col_idx = None
        for ci, v in enumerate(header_row):
            sv = str(v).strip() if v else ''
            if re.match(r'(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)', sv.lower()):
                meses.append((ci, sv))
            elif sv.lower() == 'total':
                total_col_idx = ci

        if not meses and not total_col_idx:
            continue

        # Encontrar linhas-chave
        km = {}
        for i, row in enumerate(rows, 1):
            b = str(row[1]).strip().lower() if len(row) > 1 and row[1] else ''
            if b == 'faturamento total':           km['fat'] = i
            elif b == 'custos variaveis':          km['cv'] = i
            elif '200 - aquisição de imobilizado' in b: km['cap_imo'] = i
            elif '20702' in b and 'moveis' in b:   km['cap_mov'] = i
            elif '20704' in b and 'máquinas' in b: km['cap_maq'] = i
            elif '20706' in b and 'imóveis' in b:  km['cap_imov'] = i
            elif b == 'margem de contribuição':    km['mc'] = i
            elif b == 'despesas fixas':            km['df'] = i
            elif b == 'despesas com pessoal':      km['pessoal'] = i
            elif b == 'despesas administrativas':  km['adm'] = i
            elif b == 'despesas financeiras':      km['fin'] = i
            elif b == 'resultado antes dos impostos': km['antes_ir'] = i
            elif b == 'despesas com impostos':     km['impostos'] = i
            elif '20407' in b or 'ir pessoa jurídica' in b: km['irpj'] = i
            elif '20412' in b or b == 'csll':      km['csll'] = i
            elif b == 'resultado operacional':     km['res'] = i

        if 'fat' not in km or 'res' not in km:
            continue

        def gv(key, col_idx=None):
            r = km.get(key)
            if r is None: return 0.0
            row = rows[r - 1]
            if col_idx is not None:
                try: return float(row[col_idx]) if row[col_idx] else 0.0
                except: return 0.0
            # Soma todos os meses
            total = 0.0
            for ci, _ in meses:
                try: total += float(row[ci]) if row[ci] else 0.0
                except: pass
            return total

        fat    = gv('fat',  total_col_idx) if total_col_idx else gv('fat')
        cv     = gv('cv',   total_col_idx) if total_col_idx else gv('cv')
        mc     = gv('mc',   total_col_idx) if total_col_idx else gv('mc')
        df     = gv('df',   total_col_idx) if total_col_idx else gv('df')
        adm    = gv('adm',  total_col_idx) if total_col_idx else gv('adm')
        fin    = gv('fin',  total_col_idx) if total_col_idx else gv('fin')
        irpj   = gv('irpj', total_col_idx) if total_col_idx else gv('irpj')
        csll   = gv('csll', total_col_idx) if total_col_idx else gv('csll')
        res    = gv('res',  total_col_idx) if total_col_idx else gv('res')

        cap_imo  = gv('cap_imo',  total_col_idx) if total_col_idx else gv('cap_imo')
        cap_mov  = gv('cap_mov',  total_col_idx) if total_col_idx else gv('cap_mov')
        cap_maq  = gv('cap_maq',  total_col_idx) if total_col_idx else gv('cap_maq')
        cap_imov = gv('cap_imov', total_col_idx) if total_col_idx else gv('cap_imov')

        capex_total    = -(cap_imo + cap_mov + cap_maq + cap_imov)
        impostos_total = -(irpj + csll)
        da_val         = fat * da_pct

        aj1 = capex_total    if ajustes[0] else 0
        aj2 = impostos_total if ajustes[1] else 0
        aj3 = 0              # amortização não segregada por obra
        aj4 = da_val         if ajustes[3] else 0

        res_corr   = res + aj1 + aj2 + aj3 + aj4
        ebitda_orig = res + abs(fin) + abs(irpj + csll)
        ebitda_corr = res_corr + abs(fin) + aj4

        resultados[sname] = {
            'obra': sname,
            'faturamento': fat,
            'custos_variaveis': cv,
            'margem_contribuicao': mc,
            'despesas_fixas': df,
            'desp_adm': adm,
            'desp_fin': fin,
            'irpj_csll': irpj + csll,
            'resultado_original': res,
            'capex_total': capex_total,
            'impostos_total': impostos_total,
            'da_val': da_val,
            'aj1': aj1, 'aj2': aj2, 'aj3': aj3, 'aj4': aj4,
            'resultado_corrigido': res_corr,
            'ebitda_original': ebitda_orig,
            'ebitda_corrigido': ebitda_corr,
            'margem_orig': res / fat if fat else 0,
            'margem_corr': res_corr / fat if fat else 0,
            'meses': meses,
            'total_col': total_col_idx,
            'km': km,
            'rows': rows,
        }

    return resultados


def processar_dre_consolidada(wb, nome_arquivo, ajustes, da_pct):
    """Processa DRE consolidada anual (estrutura com colunas de meses)"""
    # Usar aba principal (primeira aba de dados)
    ABAS_EXCLUIR = {'obras em andamento', 'fluxo de caixa', 'rental', 'dashboard'}
    abas = [s for s in wb.sheetnames if s.lower() not in ABAS_EXCLUIR]
    if not abas:
        return None

    ws = wb[abas[0]]
    rows = list(ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True))

    # Tentar detectar estrutura
    km = {}
    for i, row in enumerate(rows, 1):
        b = str(row[1]).strip().lower() if len(row) > 1 and row[1] else ''
        a = str(row[0]).strip().lower() if row[0] else ''
        code = str(row[0]).strip().rstrip() if row[0] else ''
        code = ' '.join(code.split())  # normalize spaces

        # Detect by EXACT code match only (avoid substring false positives)
        # Use only if not already found (first match wins)
        if code in ('01.01.01', '01.01.01.01') and 'rec_bruta' not in km:
            km['rec_bruta'] = i
        elif code == '01.01.02' and 'deducoes' not in km:
            km['deducoes'] = i
        elif code == '02' and 'total_custos' not in km:
            km['total_custos'] = i
        elif code == '02.01.01' and 'custos_dir' not in km:
            km['custos_dir'] = i
        elif code == '02.01.02' and 'custos_ind' not in km:
            km['custos_ind'] = i
        elif code == '02.01.02.05.04' and 'cap_imo' not in km:
            km['cap_imo'] = i
        elif ('aquisição de imobilizado' in b and 'imóveis' not in b and 'cap_imo' not in km):
            km['cap_imo'] = i
        elif code == '03.01' and 'desp_adm' not in km:
            km['desp_adm'] = i
        elif code == '03.02' and 'desp_trib' not in km:
            km['desp_trib'] = i
        elif code == '03.03' and 'desp_fin' not in km:
            km['desp_fin'] = i
        elif code == '03.03.01.07' and 'amort_emprest' not in km:
            km['amort_emprest'] = i
        elif code == '04' and 'outras_nop' not in km:
            km['outras_nop'] = i
        elif code == '05' and 'resultado' not in km:
            km['resultado'] = i
        elif b in ['resultado operacional', '05 resultado operacional'] and 'resultado' not in km:
            km['resultado'] = i
        elif code in ('01.01.02.01.06', '20407') or 'ir pessoa jurídica' in b:
            if 'irpj' not in km: km['irpj'] = i
        elif code in ('01.01.02.01.05', '20412') or b == 'csll':
            if 'csll' not in km: km['csll'] = i

    # Se não encontrou pela estrutura 01.01.01, tentar pelo nome
    if 'rec_bruta' not in km:
        for i, row in enumerate(rows, 1):
            b = str(row[1]).strip().lower() if len(row) > 1 and row[1] else ''
            if 'receita bruta' in b and 'rec_bruta' not in km: km['rec_bruta'] = i
            elif 'receita líquida' in b and 'rec_liq' not in km: km['rec_liq'] = i
            elif 'resultado' in b and b not in ['resultado antes dos impostos']:
                if 'resultado' not in km: km['resultado'] = i

    # Encontrar coluna Total — busca em mais linhas
    total_col = None
    for row in rows[:15]:
        for ci, v in enumerate(row):
            sv = str(v).strip().lower() if v else ''
            if sv == 'total' or sv == 'total geral':
                total_col = ci
                break
        if total_col is not None: break
    # Se não encontrou "Total", usar última coluna não-zero com dados
    if total_col is None and rows:
        for row in rows[5:20]:
            for ci in range(min(len(row)-1, 30), 1, -1):
                try:
                    v = row[ci]
                    if v is not None and abs(float(v)) > 1000:
                        total_col = ci
                        break
                except: pass
            if total_col is not None: break
    # Last resort: col index 2 (col C) for single-column annual files
    if total_col is None:
        total_col = 2

    def gv(key):
        r = km.get(key)
        if r is None: return 0.0
        row = rows[r - 1]
        if total_col is not None:
            try: return float(row[total_col]) if row[total_col] else 0.0
            except: return 0.0
        # Tentar última coluna com dado
        for v in reversed(row[2:]):
            try:
                f = float(v)
                if f != 0: return f
            except: pass
        return 0.0

    rec_bruta  = gv('rec_bruta')
    deducoes   = gv('deducoes')
    rec_liq    = rec_bruta + deducoes
    custos_dir = gv('custos_dir')
    custos_ind = gv('custos_ind')
    cap_imo    = gv('cap_imo')
    desp_adm   = gv('desp_adm')
    desp_trib  = gv('desp_trib')
    desp_fin   = gv('desp_fin')
    amort      = gv('amort_emprest')
    outras_nop = gv('outras_nop')
    resultado  = gv('resultado')
    irpj       = gv('irpj')
    csll_val   = gv('csll')

    capex_aj    = -cap_imo         if ajustes[0] else 0
    impostos_aj = -(irpj + csll_val) if ajustes[1] else 0
    amort_aj    = -amort           if ajustes[2] else 0
    da_val      = rec_bruta * da_pct if ajustes[3] else 0

    res_corr    = resultado + capex_aj + impostos_aj + amort_aj + da_val
    juros_reais = desp_fin - amort if ajustes[2] else desp_fin
    ebitda_orig = resultado - desp_fin - outras_nop
    ebitda_corr = (resultado + capex_aj + impostos_aj + amort_aj
                   - juros_reais - outras_nop + da_val)

    return {
        'tipo': 'consolidada',
        'arquivo': nome_arquivo,
        'rec_bruta': rec_bruta,
        'deducoes': deducoes,
        'rec_liq': rec_liq,
        'custos_dir': custos_dir,
        'custos_ind': custos_ind,
        'cap_imo': cap_imo,
        'lucro_bruto': rec_liq + custos_dir + custos_ind,
        'desp_adm': desp_adm,
        'desp_trib': desp_trib,
        'desp_fin': desp_fin,
        'amort': amort,
        'outras_nop': outras_nop,
        'irpj_csll': irpj + csll_val,
        'resultado_original': resultado,
        'capex_aj': capex_aj,
        'impostos_aj': impostos_aj,
        'amort_aj': amort_aj,
        'da_val': da_val,
        'resultado_corrigido': res_corr,
        'ebitda_original': ebitda_orig,
        'ebitda_corrigido': ebitda_corr,
        'margem_ebitda_orig': ebitda_orig / rec_bruta if rec_bruta else 0,
        'margem_ebitda_corr': ebitda_corr / rec_bruta if rec_bruta else 0,
    }


def gerar_excel_corrigido(dados_processados, ajustes, da_pct):
    """Gera o Excel corrigido para download"""
    wb = Workbook()

    P = {
        'dark': '1A2744', 'med': '2E4A7A', 'light': 'BBDEFB',
        'white': 'FFFFFF', 'gray': 'F5F7FA', 'gray2': 'ECEFF1',
        'green': 'E8F5E9', '1B5E20': '1B5E20',
        'amber': 'FFF8E1', 'E65100': 'E65100',
        'purple': 'EDE7F6', '4A148C': '4A148C',
    }

    def fl(k): return PatternFill('solid', fgColor='FF' + P.get(k, k))
    def ft(bold=False, sz=9, color='333333', italic=False):
        c = color.lstrip('#')
        if len(c) == 6: c = 'FF' + c
        return Font(bold=bold, size=sz, color=c, name='Arial', italic=italic)
    def al(h='left', v='center', wrap=False):
        return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
    def bdr():
        return Border(bottom=Side(border_style='thin', color='E2E8F0'),
                      left=Side(border_style='thin', color='E2E8F0'))

    # ABA PAINEL
    ws = wb.active
    ws.title = 'Painel Geral'
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 35
    for col in ['C', 'D', 'E', 'F', 'G', 'H']:
        ws.column_dimensions[col].width = 18

    ws.merge_cells('A1:H1')
    c = ws['A1']
    c.value = f'PAINEL DRE CORRIGIDA — CAMPESATTO CONSTRUTORA | Gerado em {datetime.now().strftime("%d/%m/%Y %H:%M")}'
    c.font = ft(True, 12, 'FFFFFF')
    c.fill = fl('dark')
    c.alignment = al('center')
    ws.row_dimensions[1].height = 26

    headers = ['', 'OBRA / PERÍODO', 'FATURAMENTO', 'RES. ORIGINAL', '+CAPEX', '+IRPJ/CSLL', '+D&A', 'RES. CORRIGIDO']
    for ci, h in enumerate(headers, 1):
        c = ws.cell(4, ci, h)
        c.font = ft(True, 9, 'FFFFFF')
        c.fill = fl('med')
        c.alignment = al('center')
    ws.row_dimensions[4].height = 20
    ws.row_dimensions[2].height = 6
    ws.row_dimensions[3].height = 6

    row = 5
    for tipo, dados in dados_processados:
        if tipo == 'obras':
            for obra, d in dados.items():
                bg = 'white' if row % 2 == 0 else 'gray'
                for ci in range(1, 9):
                    ws.cell(row, ci).fill = fl(bg)
                ws.cell(row, 2, obra).font = ft(False, 9, '333333')
                ws.cell(row, 2).alignment = al('left')

                vals = [d['faturamento'], d['resultado_original'],
                        d['aj1'], d['aj2'], d['aj4'], d['resultado_corrigido']]
                for ci, v in zip(range(3, 9), vals):
                    c = ws.cell(row, ci, v)
                    c.font = ft(False, 9, '1B5E20' if (ci == 8 and v >= 0) else
                                ('E65100' if ci in [4, 5, 6] else
                                 ('4A148C' if ci == 7 else
                                  ('B71C1C' if v < 0 else '333333'))))
                    c.alignment = al('right')
                    c.number_format = '#,##0;(#,##0);"-"'
                row += 1

        elif tipo in ('consolidada', 'mensal'):
            d = dados
            periodo = d.get('periodo', '')
            label = f"{d['arquivo']}" + (f" — {periodo}" if periodo and tipo=='mensal' else '')
            ws.cell(row, 2, label).font = ft(True, 10, '1A2744')
            vals = [d['rec_bruta'], d['resultado_original'],
                    d.get('capex_aj', 0), d.get('impostos_aj', 0),
                    d.get('da_val', 0), d['resultado_corrigido']]
            for ci, v in zip(range(3, 9), vals):
                c = ws.cell(row, ci, v)
                c.font = ft(True, 10, '1B5E20' if (ci == 8 and v >= 0) else '333333')
                c.alignment = al('right')
                c.number_format = '#,##0;(#,##0);"-"'
            for ci in range(1, 9):
                ws.cell(row, ci).fill = fl('light')
            row += 1

    wb_bytes = io.BytesIO()
    wb.save(wb_bytes)
    wb_bytes.seek(0)
    return wb_bytes.getvalue()


# ================================================================
# INTERFACE STREAMLIT
# ================================================================

def main():
    # Header
    st.markdown("""
    <div class="header-banner">
        <div>
            <div class="header-title">📊 DRE Sienge — Processador de Ajustes</div>
            <div class="header-sub">Campesatto Construtora Ltda · Aplicação automática dos 4 ajustes de conciliação</div>
        </div>
        <div class="header-badge">v1.0 · 2024–2026</div>
    </div>
    """, unsafe_allow_html=True)

    # Sidebar — configuração dos ajustes
    with st.sidebar:
        st.markdown("### ⚙️ Configuração dos Ajustes")
        st.markdown("---")

        st.markdown("**Ative ou desative cada ajuste:**")

        aj1 = st.checkbox(
            "Ajuste 1 — Reverter CAPEX do custo",
            value=True,
            help="Remove Aquisição de Imobilizado, Máquinas, Móveis e Imóveis do custo. São ativos, não despesas."
        )
        aj2 = st.checkbox(
            "Ajuste 2 — Reposicionar IRPJ e CSLL",
            value=True,
            help="Move IRPJ e CSLL para provisão no final da DRE (lugar correto). Sai das deduções de receita."
        )
        aj3 = st.checkbox(
            "Ajuste 3 — Separar amortização de principal",
            value=True,
            help="Remove amortização de principal das despesas financeiras (vai para o Balanço, não é despesa)."
        )
        aj4 = st.checkbox(
            "Ajuste 4 — Adicionar D&A estimado",
            value=True,
            help="Adiciona estimativa de Depreciação e Amortização como despesa não-caixa."
        )
        ajustes = [aj1, aj2, aj3, aj4]

        st.markdown("---")
        st.markdown("**D&A — % da Receita Bruta:**")
        da_pct = st.slider(
            "Percentual D&A",
            min_value=0.5, max_value=3.0, value=1.0, step=0.1,
            format="%.1f%%",
            help="Para maior precisão, use o razão de depreciação do contador. 1% = estimativa conservadora para construtoras com frota pesada."
        ) / 100

        st.markdown("---")
        st.markdown("""
        <div style='font-size:11px;color:#6B7280;line-height:1.6'>
        <strong>Convenção de cores no Excel:</strong><br>
        🔵 Azul = input Sienge<br>
        🟠 Laranja = valor original<br>
        🟣 Roxo = ajuste<br>
        🟢 Verde = valor corrigido
        </div>
        """, unsafe_allow_html=True)

    # Área principal — Tabs
    tab_upload, tab_resultado, tab_obras, tab_comparativo, tab_ajuda = st.tabs([
        "📂 Upload", "📊 Resultado", "🏗️ Por Obra", "📈 Comparativo", "❓ Como funciona"
    ])

    # ── TAB UPLOAD ──
    with tab_upload:
        col1, col2 = st.columns([2, 1])

        with col1:
            st.markdown("#### Faça upload dos arquivos exportados do Sienge")
            st.markdown("""
            <div class="info-box">
            <strong>Aceita:</strong> arquivos <code>.xlsx</code>, <code>.xlsm</code>, <code>.xls</code>
            exportados do Sienge. Pode enviar vários de uma vez — obras individuais,
            consolidada anual, comparativo de anos.
            </div>
            """, unsafe_allow_html=True)

            uploaded_files = st.file_uploader(
                "Selecione os arquivos",
                type=["xlsx", "xlsm", "xls"],
                accept_multiple_files=True,
                label_visibility="collapsed"
            )

        with col2:
            st.markdown("#### Ajustes ativos")
            for i, (nome, ativo) in enumerate([
                ("CAPEX no custo", aj1),
                ("IRPJ/CSLL nas deduções", aj2),
                ("Amortização como despesa", aj3),
                ("D&A não segregado", aj4),
            ], 1):
                cor = "#059669" if ativo else "#9CA3AF"
                icone = "✓" if ativo else "—"
                st.markdown(f"""
                <div style='display:flex;align-items:center;gap:8px;padding:8px 12px;
                background:#fff;border-radius:8px;border:1px solid #E8ECF0;margin-bottom:6px'>
                <span style='color:{cor};font-weight:600;font-size:14px'>{icone}</span>
                <span style='font-size:13px;color:{"#1A2744" if ativo else "#9CA3AF"}'>{i}. {nome}</span>
                </div>
                """, unsafe_allow_html=True)
            st.markdown(f"""
            <div style='padding:8px 12px;background:#EFF6FF;border-radius:8px;
            border:1px solid #BFDBFE;font-size:12px;color:#1D4ED8;margin-top:4px'>
            D&A: <strong>{da_pct*100:.1f}%</strong> da receita bruta/mês
            </div>
            """, unsafe_allow_html=True)

        if uploaded_files:
            st.markdown("---")
            st.markdown(f"**{len(uploaded_files)} arquivo(s) carregado(s):**")

            tipos = []
            for f in uploaded_files:
                tipo = detectar_tipo_arquivo(f.name)
                icone = "🏗️" if tipo == 'obras' else "📋"
                badge = "badge-obra" if tipo == 'obras' else "badge-consol"
                label = "Obras" if tipo == 'obras' else "Consolidada"
                st.markdown(f"""
                <div style='display:flex;align-items:center;gap:10px;padding:10px 14px;
                background:#fff;border-radius:8px;border:1px solid #E8ECF0;margin-bottom:6px'>
                <span style='font-size:20px'>{icone}</span>
                <span style='font-size:13px;font-weight:500;flex:1'>{f.name}</span>
                <span class='badge {badge}'>{label}</span>
                <span style='font-size:12px;color:#9CA3AF'>{f.size/1024:.0f} KB</span>
                </div>
                """, unsafe_allow_html=True)
                tipos.append(tipo)

            st.markdown("")
            if st.button("⚡ Processar e aplicar ajustes", type="primary"):
                dados_todos = []
                progress = st.progress(0, text="Iniciando processamento...")

                for idx, (f, tipo) in enumerate(zip(uploaded_files, tipos)):
                    progress.progress(
                        int((idx / len(uploaded_files)) * 80),
                        text=f"Processando: {f.name}..."
                    )
                    try:
                        wb = load_workbook(
                            io.BytesIO(f.read()),
                            keep_vba=False, data_only=True
                        )
                        # Detecção robusta pelo CONTEÚDO do arquivo
                        n_abas = len(wb.sheetnames)
                        has_planilha1 = any(s.lower() in ('planilha1','sheet1','plan1')
                                           for s in wb.sheetnames)
                        ws_test = wb[wb.sheetnames[0]]
                        max_col_test = ws_test.max_column

                        if n_abas > 3 and not has_planilha1:
                            # Múltiplas abas com nomes de obras = formato obras
                            res = processar_dre_obras(wb, f.name, ajustes, da_pct)
                            dados_todos.append(('obras', res))
                        elif max_col_test <= 4 and has_planilha1:
                            # Aba única com poucas colunas = mensal
                            res = processar_dre_mensal(wb, f.name, ajustes, da_pct)
                            if res and res.get('rec_bruta', 0) != 0:
                                dados_todos.append(('mensal', res))
                        elif max_col_test > 10:
                            # Muitas colunas = anual com meses
                            res = processar_dre_consolidada(wb, f.name, ajustes, da_pct)
                            if res and res.get('rec_bruta', 0) != 0:
                                dados_todos.append(('consolidada', res))
                            else:
                                res = processar_dre_mensal(wb, f.name, ajustes, da_pct)
                                if res:
                                    dados_todos.append(('mensal', res))
                        else:
                            # Tentar mensal primeiro, depois consolidada
                            res = processar_dre_mensal(wb, f.name, ajustes, da_pct)
                            if res and res.get('rec_bruta', 0) != 0:
                                dados_todos.append(('mensal', res))
                            else:
                                res = processar_dre_consolidada(wb, f.name, ajustes, da_pct)
                                if res:
                                    dados_todos.append(('consolidada', res))
                    except Exception as e:
                        st.warning(f"⚠️ Erro ao processar {f.name}: {e}")

                progress.progress(95, text="Gerando relatórios...")
                st.session_state['dados_processados'] = dados_todos
                st.session_state['ajustes'] = ajustes
                st.session_state['da_pct'] = da_pct
                progress.progress(100, text="Concluído!")
                st.success(f"✅ {len(uploaded_files)} arquivo(s) processado(s) com sucesso!")
                st.info("👉 Acesse as abas **Resultado**, **Por Obra** ou **Comparativo** para ver os resultados.")

    # ── TAB RESULTADO ──
    with tab_resultado:
        if 'dados_processados' not in st.session_state:
            st.markdown("""
            <div style='text-align:center;padding:60px 0;color:#9CA3AF'>
            <div style='font-size:48px;margin-bottom:16px'>📂</div>
            <div style='font-size:16px;font-weight:500'>Faça o upload e processe os arquivos primeiro</div>
            <div style='font-size:13px;margin-top:8px'>Acesse a aba "Upload" para começar</div>
            </div>
            """, unsafe_allow_html=True)
        else:
            dados = st.session_state['dados_processados']
            da_pct_us = st.session_state['da_pct']
            ajustes_us = st.session_state['ajustes']

            # Consolidar totais
            total_fat = 0
            total_res_orig = 0
            total_res_corr = 0
            total_ebitda_corr = 0

            for tipo, d in dados:
                if tipo == 'obras':
                    for obra, od in d.items():
                        total_fat += od['faturamento']
                        total_res_orig += od['resultado_original']
                        total_res_corr += od['resultado_corrigido']
                        total_ebitda_corr += od['ebitda_corrigido']
                elif tipo in ('consolidada', 'mensal'):
                    total_fat += d['rec_bruta']
                    total_res_orig += d['resultado_original']
                    total_res_corr += d['resultado_corrigido']
                    total_ebitda_corr += d['ebitda_corrigido']

            # KPIs
            c1, c2, c3, c4 = st.columns(4)
            kpi_data = [
                (c1, "Faturamento Total", total_fat, None, ""),
                (c2, "Resultado Original (Sienge)", total_res_orig, None,
                 "💡 Antes dos ajustes"),
                (c3, "Resultado Corrigido", total_res_corr, None,
                 "Após 4 ajustes aplicados"),
                (c4, "EBITDA Corrigido", total_ebitda_corr,
                 total_ebitda_corr / total_fat if total_fat else 0,
                 "% da Receita Bruta"),
            ]
            for col, label, val, pct_v, sub in kpi_data:
                with col:
                    cor = "kpi-pos" if val >= 0 else "kpi-neg"
                    sub_txt = f"{pct_v*100:.1f}% da RB" if pct_v is not None else sub
                    st.markdown(f"""
                    <div class="kpi-card">
                        <div class="kpi-label">{label}</div>
                        <div class="kpi-value {cor}">{fmt_m(val)}</div>
                        <div class="kpi-sub">{sub_txt}</div>
                    </div>
                    """, unsafe_allow_html=True)

            st.markdown("---")

            # Gráfico cascata
            if total_res_orig != 0:
                st.markdown("#### Cascata de ajustes: do resultado Sienge ao resultado corrigido")

                aj_total = [0, 0, 0, 0]
                for tipo, d in dados:
                    if tipo == 'obras':
                        for obra, od in d.items():
                            aj_total[0] += od['aj1']
                            aj_total[1] += od['aj2']
                            aj_total[2] += od['aj3']
                            aj_total[3] += od['aj4']
                    elif tipo == 'consolidada':
                        aj_total[0] += d['capex_aj']
                        aj_total[1] += d['impostos_aj']
                        aj_total[2] += d['amort_aj']
                        aj_total[3] += d['da_val']

                measures = ["absolute", "relative", "relative", "relative", "relative", "total"]
                x_labels = [
                    "Resultado Sienge",
                    "Aj.1 CAPEX",
                    "Aj.2 IRPJ/CSLL",
                    "Aj.3 Amortização",
                    "Aj.4 D&A",
                    "Resultado Corrigido"
                ]
                y_vals = [total_res_orig, aj_total[0], aj_total[1], aj_total[2], aj_total[3], total_res_corr]

                fig = go.Figure(go.Waterfall(
                    name="Ajustes",
                    orientation="v",
                    measure=measures,
                    x=x_labels,
                    y=y_vals,
                    connector={"line": {"color": "#E2E8F0", "width": 1}},
                    decreasing={"marker": {"color": "#DC2626"}},
                    increasing={"marker": {"color": "#059669"}},
                    totals={"marker": {"color": "#1A2744"}},
                    text=[fmt_m(v) for v in y_vals],
                    textposition="outside",
                ))
                fig.update_layout(
                    height=380,
                    margin=dict(l=40, r=40, t=20, b=40),
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)",
                    font=dict(family="IBM Plex Sans", size=12, color="#374151"),
                    showlegend=False,
                    yaxis=dict(gridcolor="#F1F5F9", zeroline=True, zerolinecolor="#CBD5E1"),
                )
                st.plotly_chart(fig, use_container_width=True)

            # Exibir DRE mensal se houver
            for tipo, d in dados:
                if tipo == 'mensal':
                    st.markdown("---")
                    st.markdown(f"#### DRE Mensal Corrigida — {d['periodo']} · {d['arquivo']}")
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.markdown(f"""<div class="kpi-card">
                            <div class="kpi-label">Receita Bruta</div>
                            <div class="kpi-value kpi-neu">{fmt_m(d['rec_bruta'])}</div>
                            <div class="kpi-sub">{d['periodo']}</div>
                        </div>""", unsafe_allow_html=True)
                    with col2:
                        st.markdown(f"""<div class="kpi-card">
                            <div class="kpi-label">Resultado Original</div>
                            <div class="kpi-value {'kpi-pos' if d['resultado_original']>=0 else 'kpi-neg'}">{fmt_m(d['resultado_original'])}</div>
                            <div class="kpi-sub">DRE Sienge</div>
                        </div>""", unsafe_allow_html=True)
                    with col3:
                        st.markdown(f"""<div class="kpi-card">
                            <div class="kpi-label">Resultado Corrigido</div>
                            <div class="kpi-value {'kpi-pos' if d['resultado_corrigido']>=0 else 'kpi-neg'}">{fmt_m(d['resultado_corrigido'])}</div>
                            <div class="kpi-sub">Após 4 ajustes</div>
                        </div>""", unsafe_allow_html=True)

                    # Tabela resumida
                    linhas_mensal = [
                        ("RECEITA BRUTA", d['rec_bruta'], 0, d['rec_bruta'], "total"),
                        ("(-) Deduções s/ Receita", d['deducoes'], d['impostos_aj'], d['deducoes']+d['impostos_aj'], "normal"),
                        ("  ↳ Aj.2: IRPJ+CSLL reposicionados", d['irpj_csll'], -d['irpj_csll'], 0, "ajuste") if d['impostos_aj'] else None,
                        ("RECEITA LÍQUIDA", d['rec_liq_orig'], d['impostos_aj'], d['rec_liq_corr'], "sub"),
                        ("(-) Custos Diretos", d['custos_dir'], 0, d['custos_dir'], "normal"),
                        ("(-) Aquisição Imobilizado (CAPEX)", d['capex'], d['capex_aj'], 0, "ajuste") if d['capex_aj'] else None,
                        ("(-) Custos Indiretos", d['custos_ind'], d['capex_aj'], d['custos_ind']-d['capex_aj'], "normal"),
                        ("LUCRO BRUTO", d['lucro_bruto'], d['capex_aj']+d['impostos_aj'], d['lucro_bruto']+d['capex_aj']+d['impostos_aj'], "sub"),
                        ("(-) Desp Administrativas", d['desp_adm'], 0, d['desp_adm'], "normal"),
                        ("(-) Desp Tributárias", d['desp_trib'], 0, d['desp_trib'], "normal"),
                        ("(-) Desp Financeiras", d['desp_fin'], d['amort_aj'], d['desp_fin']+d['amort_aj'], "normal"),
                        ("  ↳ Aj.3: Amortiz. principal retirada", d['amort'], -d['amort'], 0, "ajuste") if d['amort_aj'] else None,
                        ("  ↳ Aj.4: D&A estimado adicionado", 0, d['da_val'], d['da_val'], "ajuste") if d['da_val'] else None,
                        ("(-) Outras NOP", d['outras_nop'], 0, d['outras_nop'], "normal"),
                        ("★ RESULTADO CORRIGIDO", d['resultado_original'], sum([d['capex_aj'],d['impostos_aj'],d['amort_aj'],d['da_val']]), d['resultado_corrigido'], "destaque"),
                        (f"  Margem ({d['periodo']})", pct(d['resultado_original'],d['rec_bruta']), "→", pct(d['resultado_corrigido'],d['rec_bruta']), "info"),
                    ]
                    html = """<table class='dre-table'>
                    <thead><tr><th>Linha</th><th>Original</th><th>Ajuste</th><th>Corrigido</th></tr></thead><tbody>"""
                    for linha in linhas_mensal:
                        if linha is None: continue
                        nome_, orig_, adj_, corr_, cls_ = linha
                        if cls_ == "info":
                            html += f"<tr class='ajuste'><td class='indent'>{nome_}</td><td class='val-orig'>{orig_}</td><td>→</td><td class='val-corr'>{corr_}</td></tr>"
                            continue
                        def fv2(v):
                            if isinstance(v,str): return v
                            cor = "val-pos" if float(v)>=0 else "val-neg"
                            return f"<span class='{cor}'>{fmt_brl_k(float(v))}</span>"
                        rc_ = {"total":"total","sub":"total","destaque":"destaque","ajuste":"ajuste","normal":""}.get(cls_,"")
                        html += f"<tr class='{rc_}'><td>{nome_}</td><td class='val-orig'>{fv2(orig_)}</td><td class='val-adj'>{fv2(adj_)}</td><td class='val-corr'>{fv2(corr_)}</td></tr>"
                    html += "</tbody></table>"
                    st.markdown(html, unsafe_allow_html=True)

            # Download
            st.markdown("---")
            excel_bytes = gerar_excel_corrigido(dados, ajustes_us, da_pct_us)
            st.download_button(
                label="⬇️ Baixar DRE Corrigida (.xlsx)",
                data=excel_bytes,
                file_name=f"DRE_Corrigida_Campesatto_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    # ── TAB POR OBRA ──
    with tab_obras:
        if 'dados_processados' not in st.session_state:
            st.info("Processe os arquivos primeiro na aba Upload.")
        else:
            dados = st.session_state['dados_processados']
            obras_dict = {}
            for tipo, d in dados:
                if tipo == 'obras':
                    obras_dict.update(d)

            if not obras_dict:
                st.info("Nenhuma DRE de obra encontrada nos arquivos processados.")
            else:
                obras_lista = list(obras_dict.keys())
                obra_sel = st.selectbox("Selecione a obra:", obras_lista)

                if obra_sel:
                    d = obras_dict[obra_sel]
                    fat = d['faturamento']
                    res_o = d['resultado_original']
                    res_c = d['resultado_corrigido']

                    c1, c2, c3 = st.columns(3)
                    with c1:
                        st.metric("Faturamento", fmt_m(fat))
                    with c2:
                        delta = res_c - res_o
                        st.metric("Resultado Original", fmt_m(res_o),
                                  delta=f"{fmt_m(delta)} pelos ajustes")
                    with c3:
                        cor_mg = "normal" if res_c >= 0 else "inverse"
                        st.metric("Resultado Corrigido", fmt_m(res_c),
                                  delta=f"{d['margem_corr']*100:.1f}% de margem")

                    st.markdown("---")
                    st.markdown("#### Memória de cálculo dos ajustes")

                    linhas = [
                        ("FATURAMENTO TOTAL", fat, 0, fat, "total"),
                        ("Custos Variáveis", d['custos_variaveis'], 0, d['custos_variaveis'], "normal"),
                        ("  ↳ Capex no custo (aj.1)", -d['aj1'], d['aj1'], 0, "ajuste") if d['aj1'] else None,
                        ("MARGEM DE CONTRIBUIÇÃO", d['margem_contribuicao'], d['aj1'], d['margem_contribuicao'] + d['aj1'], "total"),
                        ("Despesas Fixas", d['despesas_fixas'], 0, d['despesas_fixas'], "normal"),
                        ("Despesas Financeiras", d['desp_fin'], 0, d['desp_fin'], "normal"),
                        ("  ↳ IRPJ+CSLL (aj.2)", -(d['irpj_csll']), d['aj2'], 0, "ajuste") if d['aj2'] else None,
                        ("  ↳ D&A estimado (aj.4)", 0, d['aj4'], d['aj4'], "ajuste") if d['aj4'] else None,
                        ("RESULTADO OPERACIONAL", res_o, d['aj1'] + d['aj2'] + d['aj4'], res_c, "destaque"),
                        ("  Margem (%)", pct(res_o, fat), "→", pct(res_c, fat), "info"),
                    ]

                    html = """<table class='dre-table'>
                    <thead><tr>
                    <th>Linha</th><th>Original</th><th>Ajuste</th><th>Corrigido</th>
                    </tr></thead><tbody>"""

                    for linha in linhas:
                        if linha is None: continue
                        nome, orig, adj, corr, classe = linha
                        if classe == "info":
                            html += f"<tr class='ajuste'><td class='indent'>{nome}</td><td class='val-orig'>{orig}</td><td>→</td><td class='val-corr'>{corr}</td></tr>"
                            continue
                        def fv(v, cls=""):
                            if isinstance(v, str): return f"<span>{v}</span>"
                            cor = "val-pos" if float(v) >= 0 else "val-neg"
                            return f"<span class='{cor}'>{fmt_brl_k(v)}</span>"
                        row_cls = {"total": "total", "destaque": "destaque",
                                   "ajuste": "ajuste", "normal": ""}.get(classe, "")
                        html += f"<tr class='{row_cls}'><td>{nome}</td><td class='val-orig'>{fv(orig)}</td><td class='val-adj'>{fv(adj)}</td><td class='val-corr'>{fv(corr)}</td></tr>"

                    html += "</tbody></table>"
                    st.markdown(html, unsafe_allow_html=True)

    # ── TAB COMPARATIVO ──
    with tab_comparativo:
        if 'dados_processados' not in st.session_state:
            st.info("Processe os arquivos primeiro na aba Upload.")
        else:
            dados = st.session_state['dados_processados']

            consolidadas = [(tipo, d) for tipo, d in dados if tipo == 'consolidada']
            obras_dicts = [d for tipo, d in dados if tipo == 'obras']

            if consolidadas:
                st.markdown("#### Comparativo: DRE Sienge × DRE Corrigida × DRE do Balanço (2024)")

                linhas_comp = [
                    ("Receita Bruta", 'rec_bruta', 'rec_bruta', 91_355_260),
                    ("Receita Líquida", 'rec_liq', 'rec_liq', 84_575_090),
                    ("Lucro Bruto", 'lucro_bruto', 'lucro_bruto', 13_955_634),
                    ("EBITDA", 'ebitda_original', 'ebitda_corrigido', 12_363_432),
                    ("Resultado", 'resultado_original', 'resultado_corrigido', 3_993_382),
                ]

                html = """<table class='dre-table'>
                <thead><tr>
                <th>Indicador</th><th>🟠 Sienge</th><th>🟣 Corrigida</th><th>🟢 SPED/Balanço 2024</th>
                </tr></thead><tbody>"""

                for tipo, d in consolidadas:
                    html += f"<tr class='secao'><td colspan='4'>{d['arquivo']}</td></tr>"
                    for nome, k_orig, k_corr, sped_val in linhas_comp:
                        orig = d.get(k_orig, 0)
                        corr = d.get(k_corr, 0)
                        def cv(v):
                            c = "val-pos" if v >= 0 else "val-neg"
                            return f"<span class='{c}'>{fmt_brl_k(v)}</span>"
                        html += f"<tr><td>{nome}</td><td class='val-orig'>{cv(orig)}</td><td class='val-adj'>{cv(corr)}</td><td class='val-corr'>{cv(sped_val)}</td></tr>"

                html += "</tbody></table>"
                st.markdown(html, unsafe_allow_html=True)

                st.markdown("""
                <div class="success-box">
                ✅ <strong>A DRE Corrigida aproxima a visão gerencial da DRE do Balanço.</strong>
                Diferenças residuais se devem ao D&A estimado vs. real confirmado pelo contador.
                Para apresentar ao banco, use sempre a DRE do SPED (Balanço) — é o documento oficial.
                </div>
                """, unsafe_allow_html=True)

            if obras_dicts:
                st.markdown("#### Resumo de obras — resultado original vs. corrigido")
                rows_chart = []
                for od in obras_dicts:
                    for obra, d in od.items():
                        rows_chart.append({
                            'Obra': obra[:20] + ('...' if len(obra) > 20 else ''),
                            'Original': d['resultado_original'],
                            'Corrigido': d['resultado_corrigido'],
                        })
                if rows_chart:
                    df = pd.DataFrame(rows_chart)
                    fig = go.Figure()
                    fig.add_trace(go.Bar(
                        name='Original', x=df['Obra'], y=df['Original'],
                        marker_color='#F97316', opacity=0.7
                    ))
                    fig.add_trace(go.Bar(
                        name='Corrigido', x=df['Obra'], y=df['Corrigido'],
                        marker_color='#059669'
                    ))
                    fig.update_layout(
                        barmode='group', height=350,
                        margin=dict(l=20, r=20, t=20, b=100),
                        paper_bgcolor="rgba(0,0,0,0)",
                        plot_bgcolor="rgba(0,0,0,0)",
                        font=dict(family="IBM Plex Sans", size=11, color="#374151"),
                        legend=dict(orientation="h", y=1.1),
                        xaxis=dict(tickangle=-30),
                        yaxis=dict(gridcolor="#F1F5F9"),
                    )
                    st.plotly_chart(fig, use_container_width=True)

    # ── TAB AJUDA ──
    with tab_ajuda:
        st.markdown("#### Por que a DRE do Sienge difere da DRE do Balanço?")
        st.markdown("""
        O Sienge é um excelente sistema de gestão operacional para controlar obras, caixa e contratos.
        Porém ele mistura três tipos de eventos em uma mesma DRE:

        | Tipo de Evento | Onde deve ir | Onde o Sienge lança |
        |---|---|---|
        | Compra de máquina (CAPEX) | Ativo Imobilizado (Balanço) | Custo dos serviços (DRE) |
        | Amortização de dívida | Redução de Passivo (Balanço) | Despesa Financeira (DRE) |
        | Aporte de capital dos sócios | Patrimônio Líquido (Balanço) | Outras Receitas (DRE) |
        | CSLL e IRPJ | Provisão no final da DRE | Deduções da Receita |
        | Depreciação (D&A) | Despesa na DRE | Não segregado |
        """)

        st.markdown("---")
        st.markdown("#### Impacto confirmado em 2024 — Campesatto Construtora")

        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown("""
            <div class="kpi-card">
                <div class="kpi-label">DRE Sienge 2024</div>
                <div class="kpi-value kpi-neg">(R$ 34,9M)</div>
                <div class="kpi-sub">Prejuízo aparente</div>
            </div>
            """, unsafe_allow_html=True)
        with col2:
            st.markdown("""
            <div class="kpi-card">
                <div class="kpi-label">Diferença (4 ajustes)</div>
                <div class="kpi-value kpi-pos">+R$ 38,9M</div>
                <div class="kpi-sub">CAPEX + amort + aportes + IRPJ/CSLL</div>
            </div>
            """, unsafe_allow_html=True)
        with col3:
            st.markdown("""
            <div class="kpi-card">
                <div class="kpi-label">DRE Balanço (SPED) 2024</div>
                <div class="kpi-value kpi-pos">R$ 3,99M</div>
                <div class="kpi-sub">Lucro real — o que o banco vê</div>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("---")
        st.markdown("""
        <div class="alert-box">
        ⚠️ <strong>Limitações desta ferramenta:</strong><br>
        — O D&A é uma estimativa percentual. Para precisão total, solicite ao contador o razão da conta de depreciação acumulada.<br>
        — A separação de amortização vs. juros requer o cronograma detalhado de cada contrato bancário.<br>
        — Esta DRE corrigida é gerencial. A DRE oficial para o banco é sempre a do SPED/Balanço elaborada pelo contador.
        </div>
        """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
