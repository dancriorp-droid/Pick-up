"""
PICK-UP SNIPER - App Streamlit
Upload dos HITS → Atualiza planilha → Gera PPTXs → Download

Fluxo:
1. Upload da planilha Pick-up_sniper.xlsx (atual)
2. Upload dos HITS (Normandie e/ou Alto)
3. Clica "Atualizar"
4. Baixa planilha atualizada + PPTXs
"""
import streamlit as st
import os, sys, re, json, shutil, calendar, zipfile, tempfile, io
from datetime import datetime, date, timedelta
import pytz
from collections import OrderedDict
from copy import deepcopy

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from lxml import etree

# ══════════════════════════════════════════════════════════════════
# CONFIG
# ══════════════════════════════════════════════════════════════════
st.set_page_config(page_title="Pick-up Sniper", page_icon="🎯", layout="wide")

HOTEIS = OrderedDict([
    ('normandie', {
        'hits_file': 'hits_normandie.xlsx',
        'aba_planilha': 'Normandie',
        'template_pptx': 'pickup_normandie_template.pptx',
        'total_uhs': 164,
    }),
    ('alto', {
        'hits_file': 'hits_alto.xlsx',
        'aba_planilha': 'Alto da Boa Vista',
        'aba_metas': 'Alto da Boa Vista Metas',
        'template_pptx': 'pickup_alto_template.pptx',
        'total_uhs': 21,
    }),
])

MESES_PT = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho',
            'Julho','Agosto','Setembro','Outubro','Novembro','Dezembro']

FONT_DEFAULT = Font(name='Aptos Narrow', size=11)
FMT_RECEITA = '#,##0.00'
FMT_OCC = '#,##0.00'

METAS = {
    'normandie': {
        5: {'dm': 315.92, 'occ': 0.85, 'rec': 1390200},
        6: {'dm': 256.90, 'occ': 0.82, 'rec': 1055400},
        7: {'dm': 289.89, 'occ': 0.72, 'rec': 1080545},
        8: {'dm': 323.29, 'occ': 0.66, 'rec': 1104623.01},
        9: {'dm': 322.94, 'occ': 0.70, 'rec': 1132558.24},
        10: {'dm': 307.69, 'occ': 0.66, 'rec': 1051334.77},
        11: {'dm': 427.69, 'occ': 0.78, 'rec': 1671327.84},
        12: {'dm': 283.21, 'occ': 0.52, 'rec': 762401.25},
        1: {'dm': 329.13, 'occ': 0.41, 'rec': 698591.23},
    },
    'alto': {
        5: {'dm': 1239.98, 'occ': 0.39, 'rec': 314818.16},
        6: {'dm': 1638.57, 'occ': 0.38, 'rec': 392273.26},
        7: {'dm': 1559.03, 'occ': 0.67, 'rec': 680000},
        8: {'dm': 1250.82, 'occ': 0.42, 'rec': 342000},
        9: {'dm': 1328.23, 'occ': 0.30, 'rec': 251034.77},
        10: {'dm': 1168.04, 'occ': 0.24, 'rec': 182494.46},
        11: {'dm': 1191.08, 'occ': 0.32, 'rec': 240122.57},
        12: {'dm': 1365.42, 'occ': 0.36, 'rec': 320000},
        1: {'dm': 1114.39, 'occ': 0.31, 'rec': 224895},
    },
}

# ══════════════════════════════════════════════════════════════════
# FUNÇÕES - LEITURA HITS
# ══════════════════════════════════════════════════════════════════
def ler_hits(file_bytes):
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    ws = wb.worksheets[0]
    dados = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]: continue
        data_raw = row[2]; total = row[4]; ocupado = row[6]; receita = row[20]
        ocup_kpi = row[31] if len(row) > 31 else None
        if not data_raw or not total or ocupado is None: continue
        if isinstance(data_raw, datetime): dt = data_raw.date()
        elif isinstance(data_raw, date): dt = data_raw
        elif isinstance(data_raw, str):
            for fmt in ('%d/%m/%Y', '%Y-%m-%d'):
                try: dt = datetime.strptime(data_raw.strip(), fmt).date(); break
                except: continue
            else: continue
        else: continue
        occ = (float(ocupado) / float(total)) * 100 if float(total) > 0 else 0.
        kpi = int(ocup_kpi) if ocup_kpi and isinstance(ocup_kpi, (int, float)) else int(ocupado) if ocupado else 0
        dados[dt] = {'receita': float(receita) if receita else 0., 'ocupacao': occ, 'ocup_kpi': kpi}
    wb.close()
    return dados

# ══════════════════════════════════════════════════════════════════
# FUNÇÕES - MAPEAR MESES
# ══════════════════════════════════════════════════════════════════
def mapear_meses(ws):
    meses = []
    for c in range(1, ws.max_column + 1):
        if ws.cell(row=1, column=c).value == 'Dia':
            for r in range(3, ws.max_row + 1):
                val = ws.cell(row=r, column=c).value
                if val is None: continue
                if isinstance(val, str) and val.strip().upper() == 'TOTAL': break
                try:
                    dt = val if isinstance(val, (datetime, date)) else pd.to_datetime(val)
                    if isinstance(dt, datetime): dt = dt.date()
                    elif isinstance(dt, pd.Timestamp): dt = dt.date()
                    total_row = None
                    for tr in range(r, ws.max_row + 1):
                        tv = ws.cell(row=tr, column=c).value
                        if tv and isinstance(tv, str) and tv.strip().upper() == 'TOTAL':
                            total_row = tr; break
                    meses.append({'col_base': c, 'mes': dt.month, 'ano': dt.year, 'total_row': total_row})
                    break
                except: continue
    return meses

# ══════════════════════════════════════════════════════════════════
# FUNÇÕES - CRIAR MÊS NOVO
# ══════════════════════════════════════════════════════════════════
def criar_mes_novo(ws, meses_existentes, novo_mes, novo_ano, hotel_key):
    if meses_existentes:
        col_nova = meses_existentes[-1]['col_base'] + 8
    else:
        col_nova = 1
    n_dias = calendar.monthrange(novo_ano, novo_mes)[1]
    headers = ['Dia', 'Ontem', None, 'Hoje', None, 'Pick-up', None, None]
    for i, h in enumerate(headers):
        ws.cell(row=1, column=col_nova + i).value = h
        ws.cell(row=1, column=col_nova + i).font = FONT_DEFAULT
    sub = [None, 'Receita', 'Ocupação', 'Receita', 'Ocupação', 'Receita', 'Ocupação', None]
    for i, h in enumerate(sub):
        ws.cell(row=2, column=col_nova + i).value = h
        ws.cell(row=2, column=col_nova + i).font = FONT_DEFAULT
    for dia in range(1, n_dias + 1):
        r = 2 + dia
        ws.cell(row=r, column=col_nova).value = datetime(novo_ano, novo_mes, dia)
        ws.cell(row=r, column=col_nova).number_format = 'mm-dd-yy'
        for off in range(1, 7):
            cell = ws.cell(row=r, column=col_nova + off)
            cell.value = 0
            cell.number_format = FMT_RECEITA if off in (1,3,5) else FMT_OCC
    total_row = 3 + n_dias
    ws.cell(row=total_row, column=col_nova).value = 'TOTAL'
    ldr = 2 + n_dias
    for off, typ in [(1,'SUM'),(2,'AVERAGE'),(3,'SUM'),(4,'AVERAGE'),(5,'SUM'),(6,'AVERAGE')]:
        cl = get_column_letter(col_nova + off)
        ws.cell(row=total_row, column=col_nova+off).value = f'={typ}({cl}3:{cl}{ldr})'
        ws.cell(row=total_row, column=col_nova+off).number_format = FMT_RECEITA if off in (1,3,5) else FMT_OCC
    # Meta
    meta = METAS.get(hotel_key, {}).get(novo_mes, {})
    if meta:
        meta_str = f"Meta: R$ {meta['rec']:,.2f}".replace(',','X').replace('.',',').replace('X','.')
        ws.cell(row=total_row + 3, column=col_nova + 2).value = meta_str
        ws.cell(row=total_row + 3, column=col_nova + 2).font = Font(name='Aptos Narrow', size=11, color='FF0000')
    return {'col_base': col_nova, 'mes': novo_mes, 'ano': novo_ano, 'total_row': total_row}

# ══════════════════════════════════════════════════════════════════
# FUNÇÕES - ATUALIZAR ABA
# ══════════════════════════════════════════════════════════════════
def atualizar_aba(ws, dados_hits, hotel_key):
    meses = mapear_meses(ws)
    if not meses: return 0, {}
    datas_hits = sorted(dados_hits.keys())
    if not datas_hits: return 0, {}

    # Virada de mês (só futuros)
    meses_hits = set(d.month for d in datas_hits)
    meses_plan = set(m['mes'] for m in meses)
    max_mes = max(meses_plan) if meses_plan else 0
    for mn in sorted(meses_hits - meses_plan):
        if mn <= max_mes: continue
        an = next(d.year for d in datas_hits if d.month == mn)
        info = criar_mes_novo(ws, meses, mn, an, hotel_key)
        meses.append(info)

    atualizados = 0
    pickup_por_mes = {}

    for mi in meses:
        col = mi['col_base']
        pu_total = 0.
        for r in range(3, ws.max_row + 1):
            val = ws.cell(row=r, column=col).value
            if val is None: continue
            if isinstance(val, str) and val.strip().upper() == 'TOTAL': break
            try:
                if isinstance(val, datetime): dt = val.date()
                elif isinstance(val, date): dt = val
                else: dt = pd.to_datetime(val).date()
            except: continue
            if dt not in dados_hits: continue
            hit = dados_hits[dt]
            hoje_r = ws.cell(row=r, column=col+3).value
            hoje_o = ws.cell(row=r, column=col+4).value
            ontem_r = float(hoje_r) if hoje_r and isinstance(hoje_r, (int, float)) else 0.
            ontem_o = float(hoje_o) if hoje_o and isinstance(hoje_o, (int, float)) else 0.
            ws.cell(row=r, column=col+1).value = ontem_r
            ws.cell(row=r, column=col+1).number_format = FMT_RECEITA
            ws.cell(row=r, column=col+2).value = ontem_o
            ws.cell(row=r, column=col+2).number_format = FMT_OCC
            ws.cell(row=r, column=col+3).value = hit['receita']
            ws.cell(row=r, column=col+3).number_format = FMT_RECEITA
            ws.cell(row=r, column=col+4).value = hit['ocupacao']
            ws.cell(row=r, column=col+4).number_format = FMT_OCC
            pu_r = hit['receita'] - ontem_r
            pu_o = hit['ocupacao'] - ontem_o
            ws.cell(row=r, column=col+5).value = pu_r
            ws.cell(row=r, column=col+5).number_format = FMT_RECEITA
            ws.cell(row=r, column=col+6).value = pu_o
            ws.cell(row=r, column=col+6).number_format = FMT_OCC
            pu_total += pu_r
            atualizados += 1
        pickup_por_mes[mi['mes']] = pu_total

    # Recalcular TOTAIs
    for mi in meses:
        col = mi['col_base']; tr = mi['total_row']
        if not tr: continue
        sums = {1:0, 3:0, 5:0}
        avgs = {2:[], 4:[], 6:[]}
        for r in range(3, tr):
            val = ws.cell(row=r, column=col).value
            if val is None: continue
            if isinstance(val, str) and val.strip().upper() == 'TOTAL': break
            for off in [1,3,5]:
                v = ws.cell(row=r, column=col+off).value
                if v and isinstance(v, (int, float)): sums[off] += float(v)
            for off in [2,4,6]:
                v = ws.cell(row=r, column=col+off).value
                if v and isinstance(v, (int, float)): avgs[off].append(float(v))
        for off, total in sums.items():
            ws.cell(row=tr, column=col+off).value = total
            ws.cell(row=tr, column=col+off).number_format = FMT_RECEITA
        for off, vals in avgs.items():
            ws.cell(row=tr, column=col+off).value = sum(vals)/len(vals) if vals else 0.
            ws.cell(row=tr, column=col+off).number_format = FMT_OCC

    # DM (Receita / Ocup KPI)
    for mi in meses:
        col = mi['col_base']; tr = mi['total_row']
        total_rec = total_kpi = 0
        for r in range(3, ws.max_row + 1):
            val = ws.cell(row=r, column=col).value
            if val is None: continue
            if isinstance(val, str) and val.strip().upper() == 'TOTAL': break
            try:
                if isinstance(val, datetime): dt = val.date()
                elif isinstance(val, date): dt = val
                else: dt = pd.to_datetime(val).date()
            except: continue
            if dt in dados_hits:
                total_rec += dados_hits[dt]['receita']
                total_kpi += dados_hits[dt].get('ocup_kpi', 0)
        dm = total_rec / total_kpi if total_kpi > 0 else 0.
        if dm > 0 and tr:
            ws.cell(row=tr+2, column=col+5).value = round(dm, 2)
            ws.cell(row=tr+2, column=col+5).number_format = FMT_RECEITA

    return atualizados, pickup_por_mes

# ══════════════════════════════════════════════════════════════════
# FUNÇÕES - METAS ALTO
# ══════════════════════════════════════════════════════════════════
def atualizar_metas_alto(wb):
    if 'Alto da Boa Vista Metas' not in wb.sheetnames: return
    if 'Alto da Boa Vista' not in wb.sheetnames: return
    ws_dados = wb['Alto da Boa Vista']; ws_metas = wb['Alto da Boa Vista Metas']
    meses = mapear_meses(ws_dados)
    meta_pos = [
        {'rec_row':10,'occ_row':11,'dm_row':12,'dias_row':7},
        {'rec_row':17,'occ_row':18,'dm_row':19,'dias_row':14},
        {'rec_row':24,'occ_row':25,'dm_row':26,'dias_row':21},
        {'rec_row':31,'occ_row':32,'dm_row':33,'dias_row':28},
        {'rec_row':38,'occ_row':39,'dm_row':40,'dias_row':35},
        {'rec_row':45,'occ_row':46,'dm_row':47,'dias_row':42},
        {'rec_row':52,'occ_row':53,'dm_row':54,'dias_row':49},
        {'rec_row':59,'occ_row':60,'dm_row':61,'dias_row':56},
        {'rec_row':66,'occ_row':67,'dm_row':68,'dias_row':63},
    ]
    for idx, mi in enumerate(meses):
        if idx >= len(meta_pos): break
        mp = meta_pos[idx]; col = mi['col_base']
        total_r = total_o = 0.; n_dias = 0
        for r in range(3, ws_dados.max_row + 1):
            val = ws_dados.cell(row=r, column=col).value
            if val is None: continue
            if isinstance(val, str) and val.strip().upper() == 'TOTAL': break
            n_dias += 1
            rec = ws_dados.cell(row=r, column=col+3).value
            occ = ws_dados.cell(row=r, column=col+4).value
            if rec and isinstance(rec, (int, float)): total_r += float(rec)
            if occ and isinstance(occ, (int, float)): total_o += float(occ)
        n_dias_mes = calendar.monthrange(mi['ano'], mi['mes'])[1]
        occ_media = total_o / n_dias_mes if n_dias_mes > 0 else 0.
        ws_metas.cell(row=mp['rec_row'], column=12).value = total_r
        ws_metas.cell(row=mp['occ_row'], column=12).value = occ_media / 100 if occ_media > 1 else occ_media
        meta_r = ws_metas.cell(row=mp['rec_row'], column=13).value
        meta_o = ws_metas.cell(row=mp['occ_row'], column=13).value
        if meta_r and isinstance(meta_r, (int, float)) and float(meta_r) > 0:
            ws_metas.cell(row=mp['rec_row'], column=14).value = total_r / float(meta_r)
        if meta_o and isinstance(meta_o, (int, float)) and float(meta_o) > 0:
            ws_metas.cell(row=mp['occ_row'], column=14).value = (occ_media/100 if occ_media>1 else occ_media) / float(meta_o)
        # DM
        dm_val = ws_metas.cell(row=mp['dm_row'], column=12).value
        meta_dm = ws_metas.cell(row=mp['dm_row'], column=13).value
        if dm_val and meta_dm and isinstance(meta_dm, (int, float)) and float(meta_dm) > 0:
            ws_metas.cell(row=mp['dm_row'], column=14).value = float(dm_val) / float(meta_dm)
        # Dias faltando
        ultimo_dia = date(mi['ano'], mi['mes'], calendar.monthrange(mi['ano'], mi['mes'])[1])
        dias_faltam = max(1, (ultimo_dia - date.today()).days + 1)
        ws_metas.cell(row=mp['dias_row'], column=13).value = dias_faltam
        if meta_r and isinstance(meta_r, (int, float)) and float(meta_r) > 0:
            ws_metas.cell(row=mp['rec_row'], column=15).value = (float(meta_r) - total_r) / dias_faltam

# ══════════════════════════════════════════════════════════════════
# INTERFACE STREAMLIT
# ══════════════════════════════════════════════════════════════════
st.title("🎯 PICK-UP SNIPER")
st.markdown("**Upload dos HITS → Planilha atualizada + PPTXs gerados**")

# Upload da planilha
st.markdown("---")
planilha_file = st.file_uploader("📊 Planilha Pick-up_sniper.xlsx", type=['xlsx'], key="planilha")

# Upload dos HITS
col1, col2 = st.columns(2)
with col1:
    hits_norm = st.file_uploader("🏩 HITS Normandie", type=['xlsx'], key="hits_norm")
with col2:
    hits_alto = st.file_uploader("🏨 HITS Alto", type=['xlsx'], key="hits_alto")

# Upload do histórico (opcional)
hist_file = st.file_uploader("📋 pickup_historico.json (opcional)", type=['json'], key="hist")

# Upload dos templates
st.markdown("---")
st.markdown("**Templates PPTX** (se não enviar, usa os padrão do servidor)")
tcol1, tcol2 = st.columns(2)
with tcol1:
    tpl_norm = st.file_uploader("📄 Template Normandie", type=['pptx'], key="tpl_norm")
with tcol2:
    tpl_alto = st.file_uploader("📄 Template Alto", type=['pptx'], key="tpl_alto")

st.markdown("---")

if st.button("🚀 ATUALIZAR PICK-UP", type="primary", use_container_width=True):
    if not planilha_file:
        st.error("⚠️ Envie a planilha Pick-up_sniper.xlsx!")
    elif not hits_norm and not hits_alto:
        st.error("⚠️ Envie pelo menos um relatório HITS!")
    else:
        with st.spinner("⏳ Processando..."):
            # Criar pasta temporária
            tmp = tempfile.mkdtemp()

            # Salvar planilha
            planilha_path = os.path.join(tmp, 'Pick-up_sniper.xlsx')
            with open(planilha_path, 'wb') as f:
                f.write(planilha_file.getvalue())

            # Histórico
            hist = {}
            if hist_file:
                try:
                    hist = json.loads(hist_file.getvalue())
                except: pass

            # Data da emissão
            br_tz = pytz.timezone('America/Sao_Paulo')
            hoje = datetime.now(br_tz).date()
            dow = hoje.weekday()
            if dow == 0:
                sexta = hoje - timedelta(days=3)
                sabado = hoje - timedelta(days=2)
                domingo = hoje - timedelta(days=1)
                data_emissao = f"{sexta.day}, {sabado.day}, {domingo.strftime('%d/%m')}"
            else:
                data_emissao = (hoje - timedelta(days=1)).strftime('%d/%m')

            wb = load_workbook(planilha_path)
            resultados = {}

            for hk, hfile in [('normandie', hits_norm), ('alto', hits_alto)]:
                if not hfile: continue
                cfg = HOTEIS[hk]
                if cfg['aba_planilha'] not in wb.sheetnames: continue

                dados = ler_hits(hfile.getvalue())
                ws = wb[cfg['aba_planilha']]
                n, pu_mes = atualizar_aba(ws, dados, hk)

                # Registrar emissão no histórico
                if hk not in hist: hist[hk] = {}
                for mes, pu in pu_mes.items():
                    mes_key = str(mes)
                    if mes_key not in hist[hk]: hist[hk][mes_key] = []
                    # Evitar duplicata
                    if not hist[hk][mes_key] or hist[hk][mes_key][-1]['data'] != data_emissao:
                        hist[hk][mes_key].append({'data': data_emissao, 'pickup_total': round(pu, 2)})
                    hist[hk][mes_key] = hist[hk][mes_key][-30:]

                resultados[hk] = {'n': n, 'pu': pu_mes}

            if 'alto' in resultados:
                atualizar_metas_alto(wb)

            wb.save(planilha_path)

            # Salvar histórico
            hist_path = os.path.join(tmp, 'pickup_historico.json')
            with open(hist_path, 'w') as f:
                json.dump(hist, f, indent=2, ensure_ascii=False)

            # Salvar templates
            for hk, tpl_file, tpl_name in [('normandie', tpl_norm, 'pickup_normandie_template.pptx'),
                                             ('alto', tpl_alto, 'pickup_alto_template.pptx')]:
                tpl_path = os.path.join(tmp, tpl_name)
                if tpl_file:
                    with open(tpl_path, 'wb') as f:
                        f.write(tpl_file.getvalue())

            # Gerar PPTXs
            # Copiar gerar_pickup_pptx.py para tmp
            gerar_script = os.path.join(os.path.dirname(__file__), 'gerar_pickup_pptx.py')
            if os.path.exists(gerar_script):
                shutil.copy(gerar_script, tmp)

            pptxs_gerados = {}
            for hk in resultados:
                cfg = HOTEIS[hk]
                tpl_path = os.path.join(tmp, cfg['template_pptx'])
                out_path = os.path.join(tmp, f'pickup_{hk}_gerado.pptx')
                if os.path.exists(tpl_path):
                    import subprocess
                    r = subprocess.run(
                        [sys.executable, os.path.join(tmp, 'gerar_pickup_pptx.py'),
                         planilha_path, tpl_path, hk, out_path, hist_path],
                        capture_output=True, text=True, timeout=120
                    )
                    if os.path.exists(out_path):
                        pptxs_gerados[hk] = out_path

        # Resultados
        st.success("✅ Atualização concluída!")

        for hk, info in resultados.items():
            nome = hk.upper()
            st.markdown(f"**{nome}**: {info['n']} dias atualizados")
            for mes, pu in info['pu'].items():
                st.markdown(f"  - {MESES_PT[mes-1]}: pick-up R$ {pu:,.2f}")

        st.markdown("---")
        st.markdown("### 📥 Downloads")

        dcol1, dcol2, dcol3 = st.columns(3)

        # Planilha atualizada
        with open(planilha_path, 'rb') as f:
            dcol1.download_button("📊 Planilha", f.read(),
                                  "Pick-up_sniper.xlsx",
                                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # Histórico
        with open(hist_path, 'r') as f:
            dcol1.download_button("📋 Histórico", f.read(),
                                  "pickup_historico.json", "application/json")

        # PPTXs
        for hk, path in pptxs_gerados.items():
            with open(path, 'rb') as f:
                nome = "Normandie" if hk == 'normandie' else "Alto"
                dcol2.download_button(f"📥 PPTX {nome}", f.read(),
                                      f"pickup_{hk}_gerado.pptx",
                                      "application/vnd.openxmlformats-officedocument.presentationml.presentation")
