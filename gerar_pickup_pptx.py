"""
Gera PPTX de Pick-up Diario a partir da planilha Pick-up_sniper.xlsx
Uso: python gerar_pickup_pptx.py <planilha.xlsx> <template.pptx> <hotel> <saida.pptx>
hotel: 'alto', 'normandie' ou 'chales'

v2 - Correcoes:
- Pick-up ultimos 7 dias: preenche cada dia individualmente
- Meta: le corretamente para Normandie e Chales
- Diaria media e Ocupacao media: atualiza para todos os hoteis
- Pick-up do periodo no resumo: corrige substituicao de texto (sem duplicar)
- Campos extras (diaria media anterior, receita acum anterior): atualiza
- Receita total do TOTAL: usa valor correto da planilha
"""
import sys, re, zipfile, os
from collections import defaultdict

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

import pandas as pd
from lxml import etree

XLSX     = sys.argv[1] if len(sys.argv) > 1 else 'Pick-up_sniper.xlsx'
PPTX_IN  = sys.argv[2] if len(sys.argv) > 2 else 'template.pptx'
HOTEL    = (sys.argv[3] if len(sys.argv) > 3 else 'alto').lower()
PPTX_OUT = sys.argv[4] if len(sys.argv) > 4 else 'pickup_gerado.pptx'
HIST_FILE = sys.argv[5] if len(sys.argv) > 5 else 'pickup_historico.json'

MESES_PT = ['Janeiro','Fevereiro','Marco','Abril','Maio','Junho',
            'Julho','Agosto','Setembro','Outubro','Novembro','Dezembro']
NSA = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NSP = 'http://schemas.openxmlformats.org/presentationml/2006/main'

if HOTEL == 'alto':
    aba_dados = 'Alto da Boa Vista'
    aba_metas = 'Alto da Boa Vista Metas'
elif HOTEL == 'normandie':
    aba_dados = 'Normandie'
    aba_metas = None
elif HOTEL == 'chales':
    aba_dados = 'CHALES VILLA CARAVAGGIO'
    aba_metas = None
else:
    print(f"Hotel invalido: {HOTEL}. Use 'alto', 'normandie' ou 'chales'.")
    sys.exit(1)

import json

def ler_historico_emissoes(hotel):
    """Le historico de emissoes do arquivo JSON. Retorna {mes_num: [(data, pickup_total)]}"""
    if not os.path.exists(HIST_FILE): return {}
    with open(HIST_FILE, 'r', encoding='utf-8') as f:
        hist = json.load(f)
    if hotel not in hist: return {}
    result = {}
    for mes_key, emissoes in hist[hotel].items():
        result[int(mes_key)] = [(e['data'], e['pickup_total']) for e in emissoes]
    return result

historico_emissoes = ler_historico_emissoes(HOTEL)

xl = pd.read_excel(XLSX, sheet_name=None, header=None)

def parsear_meses(df):
    meses, col = [], 0
    while col < df.shape[1]:
        dias, nome = [], None
        t_r = t_o = pu_r = pu_o = r_ontem_t = o_ontem_t = 0.
        for i in range(2, len(df)):
            row = df.iloc[i]
            val = row.iloc[col]
            if pd.isna(val): continue
            if str(val).strip().upper() == 'TOTAL':
                def gv(c, row=row, col=col):
                    if col+c >= df.shape[1]: return 0.
                    v = row.iloc[col+c]; return float(v) if pd.notna(v) else 0.
                r_ontem_t=gv(1); o_ontem_t=gv(2)
                t_r=gv(3); t_o=gv(4); pu_r=gv(5); pu_o=gv(6); break
            try:
                dt = pd.to_datetime(val)
                if nome is None: nome = MESES_PT[dt.month-1]
                def fv(c, row=row, col=col):
                    if col+c >= df.shape[1]: return 0.
                    v = row.iloc[col+c]; return float(v) if pd.notna(v) else 0.
                dias.append({'data':dt,'dia':dt.day,
                    'r_ontem':fv(1),'o_ontem':fv(2),
                    'r_hoje':fv(3),'o_hoje':fv(4),'pu_r':fv(5),'pu_o':fv(6)})
            except: continue
        if dias:
            # Calcular ocupacao media correta: soma de todas as occ / dias do mes
            import calendar as _cal
            _mes_num = dias[0]['data'].month
            _ano = dias[0]['data'].year
            _n_dias_mes = _cal.monthrange(_ano, _mes_num)[1]
            _soma_occ = sum(d['o_hoje'] for d in dias)
            _occ_real = _soma_occ / _n_dias_mes if _n_dias_mes > 0 else 0.
            meses.append({'nome':nome,'dias':dias,
                'total_r':t_r,'total_o':t_o,'pickup_r':pu_r,'pickup_o':pu_o,
                'r_ontem_total':r_ontem_t,'o_ontem_total':o_ontem_t,
                'occ_real':_occ_real,'n_dias_mes':_n_dias_mes})
        col += 8
    return meses

def parsear_metas_alto(df):
    if df is None: return {}
    metas, mes = {}, None
    for i in range(len(df)):
        row = df.iloc[i]
        v9 = str(row.iloc[9]).strip() if pd.notna(row.iloc[9]) else ''
        v10 = str(row.iloc[10]).strip() if pd.notna(row.iloc[10]) else ''
        for m in MESES_PT:
            if v9 == m:
                mes = m; metas[m] = {}; break
        if mes and mes in metas:
            def fv(c, row=row):
                v=row.iloc[c]; return float(v) if pd.notna(v) and isinstance(v, (int,float)) else 0.
            if v10 == 'Receita':
                metas[mes].update({'acum_r':fv(11),'meta_r':fv(12),'pct_r':fv(13),'falta_dia':fv(14)})
            elif v10 in ('Ocupacao','Ocupação'):
                metas[mes].update({'acum_o':fv(11),'meta_o':fv(12),'pct_o':fv(13)})
            elif 'Di' in v10 and 'dia' in v10.lower():
                metas[mes].update({'acum_dm':fv(11),'meta_dm':fv(12),'pct_dm':fv(13)})
    return metas

def parsear_metas_inline(df):
    metas_por_idx = {}
    n_grupos = (df.shape[1] + 7) // 8
    for g in range(n_grupos):
        b = g * 8
        meta_r = 0.
        diaria_media = 0.
        for i in range(32, min(len(df), 40)):
            row = df.iloc[i]
            for c in range(b, min(b+8, df.shape[1])):
                val = row.iloc[c] if c < df.shape[1] else None
                if pd.notna(val):
                    s = str(val).strip()
                    m_meta = re.match(r'Meta:\s*(?:R\$\s*)?([\d.,]+)', s)
                    if m_meta:
                        meta_r = float(m_meta.group(1).replace('.','').replace(',','.'))
                    if s.rstrip() in ('Meta:', 'Meta'):
                        for cc in range(c+1, min(b+8, df.shape[1])):
                            vv = row.iloc[cc]
                            if pd.notna(vv) and isinstance(vv, (int, float)):
                                meta_r = float(vv); break
                            elif pd.notna(vv):
                                ss = str(vv).strip()
                                mm = re.match(r'(?:R\$\s*)?([\d.,]+)', ss)
                                if mm:
                                    meta_r = float(mm.group(1).replace('.','').replace(',','.')); break
                    if 'diaria media' in s.lower() or 'diária média' in s.lower():
                        for cc in range(c+1, min(b+8, df.shape[1])):
                            vv = row.iloc[cc]
                            if pd.notna(vv) and isinstance(vv, (int, float)):
                                diaria_media = float(vv); break
        metas_por_idx[g] = {'meta_r': meta_r, 'diaria_media': diaria_media}
    return metas_por_idx

def ler_diaria_media(xlsx_path, hotel):
    """Le diaria media de cada mes dinamicamente (TOTAL+2, col+5)."""
    from openpyxl import load_workbook as _lwb
    wb = _lwb(xlsx_path, read_only=True, data_only=True)
    aba = 'Normandie' if hotel == 'normandie' else 'CHALES VILLA CARAVAGGIO' if hotel == 'chales' else None
    if not aba or aba not in wb.sheetnames:
        wb.close(); return {}
    ws = wb[aba]
    rows = list(ws.iter_rows(values_only=True))
    result = {}
    mes_idx = 0
    c = 0
    while c < len(rows[0]) if rows else 0:
        if rows[0][c] == 'Dia':
            # Encontrar TOTAL
            for r in range(2, len(rows)):
                val = rows[r][c]
                if val and isinstance(val, str) and val.strip().upper() == 'TOTAL':
                    # DM está em TOTAL+2, col+5
                    dm_row = r + 2
                    dm_col = c + 5
                    if dm_row < len(rows) and dm_col < len(rows[dm_row]):
                        v = rows[dm_row][dm_col]
                        result[mes_idx] = float(v) if v and isinstance(v, (int, float)) else 0.
                    else:
                        result[mes_idx] = 0.
                    break
            mes_idx += 1
        c += 1
    wb.close()
    return result

meses = parsear_meses(xl[aba_dados])
if HOTEL == 'alto':
    metas_raw = parsear_metas_alto(xl.get(aba_metas))
    dm_por_idx = {}; metas_inline = {}
else:
    metas_raw = {}
    metas_inline = parsear_metas_inline(xl[aba_dados])
    dm_por_idx = ler_diaria_media(XLSX, HOTEL)

def get_meta(mes_nome, mes_idx):
    if HOTEL == 'alto':
        return dict(metas_raw.get(mes_nome, {}))
    mi = metas_inline.get(mes_idx, {})
    dm = dm_por_idx.get(mes_idx, 0.)
    if dm == 0: dm = mi.get('diaria_media', 0.)
    return {'meta_r': mi.get('meta_r', 0.), 'acum_dm': dm}

print(f"Meses encontrados: {[m['nome'] for m in meses]}")
for i, m in enumerate(meses):
    meta = get_meta(m['nome'], i)
    print(f"  {m['nome']}: total_r={m['total_r']:.2f} pu_r={m['pickup_r']:.2f} meta_r={meta.get('meta_r',0):.2f} dm={meta.get('acum_dm',0):.2f}")

def fmt_r(v):
    if v == 0: return "R$0"
    neg = v < 0; s = f"R${abs(v):,.2f}".replace(",","X").replace(".",",").replace("X",".")
    return f"-{s}" if neg else s

def fmt_r_esp(v):
    if v == 0: return "R$ 0"
    neg = v < 0; s = f"R$ {abs(v):,.2f}".replace(",","X").replace(".",",").replace("X",".")
    return f"-{s}" if neg else s

def fmt_o(v): return "0" if v == 0 else f"{abs(v):.2f}".replace(".",",")
def fmt_pct_raw(v): return f"{v:.2f}%".replace(".",",")
def fmt_pct_frac(v): return f"{v*100:.2f}%".replace(".",",")

def get_xy(elem):
    for xfrm in elem.iter(f'{{{NSA}}}xfrm'):
        off = xfrm.find(f'{{{NSA}}}off')
        if off is not None: return int(off.get('x',0)), int(off.get('y',0))
    return 0, 0

def para_text(p):
    return ''.join(t.text or '' for t in p.findall(f'.//{{{NSA}}}t'))

def set_para_value(p, new_value):
    all_t = p.findall(f'.//{{{NSA}}}t')
    if not all_t: return
    full = ''.join(t.text or '' for t in all_t)
    m = re.match(r'^((?:Receita|Occ do dia|[Oo]cc do dia|[Oo]cc|Ocupacao|Ocupação)[: ]\s*)', full, re.I)
    prefix = m.group(1) if m else ''
    all_t[0].text = prefix + new_value
    for t in all_t[1:]: t.text = ''

def set_all_text(p, new_text):
    all_t = p.findall(f'.//{{{NSA}}}t')
    if not all_t: return
    all_t[0].text = new_text
    for t in all_t[1:]: t.text = ''

def get_t_elems(elem):
    return [t for t in elem.iter(f'{{{NSA}}}t') if t.text and t.text.strip()]

def mapear_dia_card(root):
    spTree = root.find(f'.//{{{NSP}}}spTree')
    children = list(spTree)
    LINHAS_REF = [950000, 2560000, 4310000, 6060000, 7820000]
    def faixa_y(y): return min(LINHAS_REF, key=lambda l: abs(y-l))
    dias_por_linha = defaultdict(list)
    for child in children:
        tag = child.tag.split('}')[-1]
        texts = [t.text.strip() for t in get_t_elems(child)]
        if tag == 'sp' and texts and re.match(r'^\d{1,2}$', texts[0]):
            dia = int(texts[0])
            if 1 <= dia <= 31:
                x, y = get_xy(child)
                dias_por_linha[faixa_y(y)].append((dia, x))
    cards_por_linha = defaultdict(list)
    for ci, child in enumerate(children):
        tag = child.tag.split('}')[-1]
        texts = [t.text.strip() for t in get_t_elems(child)]
        if tag == 'grpSp' and any('Receita' in t for t in texts):
            if any('Pick-up total' in t for t in texts): continue
            if any('Receita total' in t for t in texts): continue
            x, y = get_xy(child)
            cards_por_linha[faixa_y(y)].append((x, child, ci))
    dia_card = {}; cards_usados = set()
    for fy in sorted(dias_por_linha.keys()):
        cards_linha = sorted([(x,e,ci) for x,e,ci in cards_por_linha[fy] if ci not in cards_usados], key=lambda c: c[0])
        if not cards_linha: continue
        dia_x = {}
        for dia, x in dias_por_linha[fy]:
            if dia not in dia_x or x < dia_x[dia]: dia_x[dia] = x
        dias_unicos = sorted(dia_x.items(), key=lambda d: d[0])
        dias_usados_linha = dias_unicos[:len(cards_linha)]
        dias_sorted = sorted(dias_usados_linha, key=lambda d: d[1])
        for (dia, _), (cx, card_elem, ci) in zip(dias_sorted, cards_linha):
            dia_card[dia] = card_elem; cards_usados.add(ci)
    return dia_card

def get_painel_total(root):
    spTree = root.find(f'.//{{{NSP}}}spTree')
    for child in spTree:
        texts = [t.text.strip() for t in get_t_elems(child)]
        if child.tag.split('}')[-1] == 'grpSp' and any('Pick-up total' in t for t in texts):
            return child
    return None

def calc_pickup_7dias(dias, mes_num):
    """Retorna (data_emissao, pickup_total) da emissao atual usando historico."""
    if mes_num in historico_emissoes and historico_emissoes[mes_num]:
        ultima = historico_emissoes[mes_num][-1]
        return ultima[0], float(ultima[1])  # (data, pickup_total)
    # Fallback
    total_pu = sum(d['pu_r'] for d in dias)
    ultimo = dias[-1] if dias else None
    data = ultimo['data'].strftime('%d/%m') if ultimo else '??/??'
    return data, total_pu

def atualizar_bloco_7dias(root, mes_data, mes_num):
    """Preenche bloco 'ultimos 7 dias' direto do historico.
    Linha 1: emissão mais recente
    Linha 2: emissão anterior
    Linha 3: emissões mais antigas agrupadas"""
    import re as _re
    NSA_local = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    NSP_local = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    
    # Pegar últimas emissões do histórico para este mês
    if mes_num not in historico_emissoes or not historico_emissoes[mes_num]:
        return root
    
    emissoes = historico_emissoes[mes_num]
    ultimas = emissoes[-7:]  # Últimas 7 emissões
    ultimas.reverse()  # Mais recente primeiro
    
    # Montar 3 linhas
    if len(ultimas) >= 1:
        linha1 = f"{ultimas[0][0]}: {fmt_r(float(ultimas[0][1]))}"
    else:
        linha1 = ""
    
    if len(ultimas) >= 2:
        linha2 = f"{ultimas[1][0]}: {fmt_r(float(ultimas[1][1]))}"
    else:
        linha2 = ""
    
    if len(ultimas) >= 3:
        restantes = ultimas[2:]
        soma = sum(float(e[1]) for e in restantes)
        datas = [e[0] for e in restantes]
        if len(datas) > 1:
            label = ', '.join(datas[:-1]) + ' e ' + datas[-1]
        else:
            label = datas[0]
        linha3 = f"{label}: {fmt_r(soma)}"
    else:
        linha3 = ""
    
    # Encontrar e preencher os parágrafos no slide
    spTree = root.find(f'.//{{{NSP_local}}}spTree')
    for child in spTree:
        texts_raw = [t.text.strip() for t in child.iter(f'{{{NSA_local}}}t') if t.text and t.text.strip()]
        if 'ltimos' not in ' '.join(texts_raw):
            continue
        
        paras = list(child.iter(f'{{{NSA_local}}}p'))
        data_paras = []
        for i, p in enumerate(paras):
            ptxt = ''.join(t.text or '' for t in p.findall(f'.//{{{NSA_local}}}t')).strip()
            if _re.match(r'^\d', ptxt) and 'R$' in ptxt:
                data_paras.append((i, p))
        
        linhas = [linha1, linha2, linha3]
        for idx, (pi, p) in enumerate(data_paras):
            if idx < len(linhas):
                all_t = p.findall(f'.//{{{NSA_local}}}t')
                if all_t:
                    all_t[0].text = linhas[idx]
                    for t in all_t[1:]:
                        t.text = ''
        
        # Se tem 4 parágrafos de dados no template, limpar o 4o
        if len(data_paras) > 3:
            for idx in range(3, len(data_paras)):
                pi, p = data_paras[idx]
                all_t = p.findall(f'.//{{{NSA_local}}}t')
                if all_t:
                    all_t[0].text = ''
                    for t in all_t[1:]:
                        t.text = ''
        
        break
    
    return root


def proc_calendario(root, mes_data):
    nome = mes_data['nome']
    dias_d = {d['dia']: d for d in mes_data['dias']}
    for t in root.iter(f'{{{NSA}}}t'):
        if t.text and t.text.strip().upper() in [m.upper() for m in MESES_PT]:
            t.text = nome.upper()
    painel = get_painel_total(root)
    if painel is not None:
        paras = list(painel.iter(f'{{{NSA}}}p'))
        for p in paras:
            txt = para_text(p).strip()
            if re.match(r'^Receita:', txt, re.I):
                set_para_value(p, fmt_r(mes_data['pickup_r']))
            elif re.match(r'^[\d,.-]+$', txt):
                set_all_text(p, fmt_o(mes_data['pickup_o']))
    dia_card = mapear_dia_card(root)
    for dia, card_elem in dia_card.items():
        if dia not in dias_d:
            for p in card_elem.iter(f'{{{NSA}}}p'):
                txt = para_text(p).strip()
                if re.match(r'^Receita:', txt, re.I): set_para_value(p, 'R$ 0')
                elif re.match(r'^[Oo]cc do dia:', txt): set_para_value(p, '0')
                elif (re.match(r'^[Oo]c+[: ]', txt) or txt.lower().startswith('occ')) and 'do dia' not in txt.lower():
                    set_para_value(p, '0')
            continue
        d = dias_d[dia]; cnt_r = cnt_o = cnt_od = 0
        for p in card_elem.iter(f'{{{NSA}}}p'):
            txt = para_text(p).strip()
            if not txt: continue
            if re.match(r'^Receita:', txt, re.I) and cnt_r == 0:
                set_para_value(p, fmt_r(d['pu_r']) if d['pu_r'] != 0 else 'R$ 0'); cnt_r += 1
            elif re.match(r'^[Oo]cc do dia:', txt) and cnt_od == 0:
                set_para_value(p, fmt_o(d['o_hoje'])); cnt_od += 1
            elif (re.match(r'^[Oo]c+[: ]', txt) or txt.lower().startswith('occ')) and 'do dia' not in txt.lower() and cnt_o == 0:
                set_para_value(p, fmt_o(d['pu_o'])); cnt_o += 1
    return root

def proc_resumo(root, mes_data, meta, mes_idx):
    nome = mes_data['nome']
    dias = mes_data['dias']
    mes_num = mes_data['dias'][0]['data'].month if mes_data['dias'] else 0

    # Pick-up: data e valor da ULTIMA EMISSAO do historico
    pu_data, pu_valor = calc_pickup_7dias(dias, mes_num)
    ultimo_dia_str = str(pu_data)
    pickup_total = float(pu_valor)

    dm = meta.get('acum_dm', 0.)
    meta_r_v = meta.get('meta_r', 0.)
    spTree = root.find(f'.//{{{NSP}}}spTree')

    for child in spTree:
        tag = child.tag.split('}')[-1]
        texts_raw = [t.text.strip() for t in get_t_elems(child)]
        texts_str = ' '.join(texts_raw)

        # Bloco TITULO (RESUMO DO PICK-UP + mes)
        if 'RESUMO' in texts_str.upper():
            for t in child.iter(f'{{{NSA}}}t'):
                if t.text and t.text.strip().upper() in [m.upper() for m in MESES_PT]:
                    t.text = nome.upper()

        # Bloco resumo principal
        elif 'Meta:' in texts_str and 'ltimos' not in texts_str:
            for p in child.iter(f'{{{NSA}}}p'):
                txt = para_text(p).strip()
                if not txt: continue
                if re.match(r'^Meta:', txt):
                    set_all_text(p, f"Meta: {fmt_r_esp(meta_r_v)}")
                elif 'Receita' in txt and 'total' in txt:
                    set_all_text(p, f"Receita total:  {fmt_r(mes_data['total_r'])}")
                elif re.match(r'^Pick-up\s+\d', txt):
                    set_all_text(p, f"Pick-up {ultimo_dia_str}:  {fmt_r(pickup_total)}")
                elif re.match(r'^R\$', txt) and 'Meta' not in txt:
                    # Valor solto de pick-up (abaixo do "Pick-up XX/XX:")
                    set_all_text(p, "")
                elif 'Di' in txt and 'dia' in txt.lower():
                    set_all_text(p, f"Diaria media: {fmt_r_esp(dm)}")
                elif 'Ocupa' in txt and 'dia' in txt.lower():
                    occ_val = mes_data.get('occ_real', mes_data['total_o'])
                    set_all_text(p, f"Ocupacao media: {fmt_pct_raw(occ_val)}")

        # Bloco "Pick-up ultimos 7 dias" - tratado por atualizar_bloco_7dias

        # sp com nome do mes
        elif tag == 'sp':
            for t in child.iter(f'{{{NSA}}}t'):
                if t.text and t.text.strip().upper() in [m.upper() for m in MESES_PT]:
                    t.text = nome.upper()

    # Campos extras soltos (fora dos grupos)
    for child in spTree:
        if child.tag.split('}')[-1] == 'sp':
            for t in child.iter(f'{{{NSA}}}t'):
                if not t.text: continue
                txt = t.text.strip()
                if re.match(r'^\d{2,3},\d+$', txt) and dm > 0:
                    t.text = fmt_o(dm)
                elif re.match(r'^R\$\s+[\d.,]+$', txt):
                    t.text = fmt_r_esp(mes_data['r_ontem_total'])

    # Atualizar bloco "Pick-up ultimos 7 dias" (le do template, rotaciona)
    root = atualizar_bloco_7dias(root, mes_data, mes_num)

    return root

def proc_metas(root, mes_data, meta):
    """Preenche slide de ACOMPANHAMENTO DE METAS do Alto.
    Remove a imagem (print) e cria uma tabela real com os dados."""
    import calendar as _cal
    from datetime import date as _date
    from copy import deepcopy
    
    nome = mes_data['nome']
    
    # Atualizar nome do mes nos textos
    for t in root.iter(f'{{{NSA}}}t'):
        if t.text and t.text.strip().upper() in [m.upper() for m in MESES_PT]:
            t.text = nome.upper()
        if t.text and t.text.strip() in MESES_PT:
            t.text = nome
    
    if not meta: return root
    
    acum_r = meta.get('acum_r', 0.)
    meta_r = meta.get('meta_r', 0.)
    pct_r = meta.get('pct_r', 0.)
    falta = meta.get('falta_dia', 0.)
    acum_o = meta.get('acum_o', 0.)
    meta_o = meta.get('meta_o', 0.)
    pct_o = meta.get('pct_o', 0.)
    acum_dm = meta.get('acum_dm', 0.)
    meta_dm = meta.get('meta_dm', 0.)
    pct_dm = meta.get('pct_dm', 0.)
    
    # Remover imagem (pic) existente
    spTree = root.find(f'{{{NSP}}}cSld/{{{NSP}}}spTree')
    if spTree is None:
        spTree = root.find(f'.//{{{NSP}}}spTree')
    
    # Remover imagens (print) E tabelas existentes (de execuções anteriores)
    to_remove = []
    for child in list(spTree):
        tag = child.tag.split('}')[-1]
        if tag == 'pic':
            to_remove.append(child)
        elif tag == 'graphicFrame':
            to_remove.append(child)
    for elem in to_remove:
        spTree.remove(elem)
    
    # Criar tabela com os dados
    NSA_URI = NSA
    
    # Definir dados da tabela
    occ_acum_str = fmt_pct_raw(acum_o) if acum_o > 1 else fmt_pct_frac(acum_o)
    occ_meta_str = fmt_pct_raw(meta_o * 100) if meta_o < 1 else fmt_pct_raw(meta_o)
    
    table_data = [
        ['', '', 'Acumulada', 'Meta', 'Atingimento', 'Falta por dia'],
        [nome, 'Receita', fmt_r_esp(acum_r), fmt_r_esp(meta_r),
         fmt_pct_frac(pct_r) if pct_r else '0%', fmt_r_esp(falta)],
        ['', 'Ocupação', occ_acum_str, occ_meta_str,
         fmt_pct_frac(pct_o) if pct_o else '0%', ''],
        ['', 'Diária Média', fmt_r_esp(acum_dm), fmt_r_esp(meta_dm),
         fmt_pct_frac(pct_dm) if pct_dm else '0%', ''],
    ]
    
    # Construir XML da tabela
    NSPC = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    gf = etree.SubElement(spTree, f'{{{NSPC}}}graphicFrame')
    
    # nvGraphicFramePr
    nvGfPr = etree.SubElement(gf, f'{{{NSPC}}}nvGraphicFramePr')
    cNvPr = etree.SubElement(nvGfPr, f'{{{NSPC}}}cNvPr')
    cNvPr.set('id', '100')
    cNvPr.set('name', 'Tabela Metas')
    etree.SubElement(nvGfPr, f'{{{NSPC}}}cNvGraphicFramePr')
    etree.SubElement(nvGfPr, f'{{{NSPC}}}nvPr')
    
    # xfrm
    xfrm = etree.SubElement(gf, f'{{{NSPC}}}xfrm')
    off = etree.SubElement(xfrm, f'{{{NSA_URI}}}off')
    off.set('x', '1500000')
    off.set('y', '4500000')
    ext = etree.SubElement(xfrm, f'{{{NSA_URI}}}ext')
    ext.set('cx', '9200000')
    ext.set('cy', '1800000')
    
    # graphic
    NSG = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    graphic = etree.SubElement(gf, f'{{{NSA_URI}}}graphic')
    graphicData = etree.SubElement(graphic, f'{{{NSA_URI}}}graphicData')
    graphicData.set('uri', 'http://schemas.openxmlformats.org/drawingml/2006/table')
    
    tbl = etree.SubElement(graphicData, f'{{{NSA_URI}}}tbl')
    tblPr = etree.SubElement(tbl, f'{{{NSA_URI}}}tblPr')
    tblPr.set('firstRow', '1')
    tblPr.set('bandRow', '1')
    
    tblGrid = etree.SubElement(tbl, f'{{{NSA_URI}}}tblGrid')
    col_widths = [1200000, 1200000, 1800000, 1800000, 1200000, 1800000]
    for w in col_widths:
        gc = etree.SubElement(tblGrid, f'{{{NSA_URI}}}gridCol')
        gc.set('w', str(w))
    
    for row_data in table_data:
        tr = etree.SubElement(tbl, f'{{{NSA_URI}}}tr')
        tr.set('h', '400000')
        for cell_text in row_data:
            tc = etree.SubElement(tr, f'{{{NSA_URI}}}tc')
            txBody = etree.SubElement(tc, f'{{{NSA_URI}}}txBody')
            bodyPr = etree.SubElement(txBody, f'{{{NSA_URI}}}bodyPr')
            p = etree.SubElement(txBody, f'{{{NSA_URI}}}p')
            pPr = etree.SubElement(p, f'{{{NSA_URI}}}pPr')
            pPr.set('algn', 'ctr')
            r = etree.SubElement(p, f'{{{NSA_URI}}}r')
            rPr = etree.SubElement(r, f'{{{NSA_URI}}}rPr')
            rPr.set('lang', 'pt-BR')
            rPr.set('sz', '1200')
            rPr.set('dirty', '0')
            t = etree.SubElement(r, f'{{{NSA_URI}}}t')
            t.text = str(cell_text)
            tcPr = etree.SubElement(tc, f'{{{NSA_URI}}}tcPr')
    
    return root


def detectar_estrutura(pptx_path):
    with zipfile.ZipFile(pptx_path) as zf:
        slides = [f for f in zf.namelist() if re.match(r'ppt/slides/slide\d+\.xml$', f)]
        n_slides = len(slides)
        tem_metas = any(
            any('ACOMPANHAMENTO' in (t.text or '').upper()
                for t in etree.parse(zf.open(sf)).iter(f'{{{NSA}}}t'))
            for sf in slides)
    return n_slides, 3 if tem_metas else 2

NSP_URI = 'http://schemas.openxmlformats.org/presentationml/2006/main'
def ajustar_nome_mes(root):
    """Padroniza o nome do mes em TODOS os slides: sem quebra de linha,
    centralizado e na mesma posicao, independente da fonte instalada no PC."""
    nomes = {m.upper() for m in MESES_PT} | {'MARCO', 'MAR\u00c7O'}
    for sp in root.iter(f'{{{NSP_URI}}}sp'):
        ts = [t for t in sp.iter(f'{{{NSA}}}t')]
        if len(ts) != 1 or not ts[0].text or ts[0].text.strip().upper() not in nomes:
            continue
        bp = next(sp.iter(f'{{{NSA}}}bodyPr'), None)
        off = next(sp.iter(f'{{{NSA}}}off'), None)
        ext = next(sp.iter(f'{{{NSA}}}ext'), None)
        if bp is None or off is None or ext is None:
            continue
        x = int(off.get('x'))
        if x < 4000000:                       # slide do calendario (cartao da esquerda)
            cx, y, w, h = 1840000, 4420000, 2700000, 420000
        else:                                 # resumo / metas: alinhar ao titulo do mesmo slide
            tit = None
            for sp2 in root.iter(f'{{{NSP_URI}}}sp'):
                tx = ''.join(t.text or '' for t in sp2.iter(f'{{{NSA}}}t')).upper()
                if tx.startswith(('RESUMO', 'ACOMPANHAMENTO')):
                    tit = sp2; break
            if tit is None:
                continue
            to = next(tit.iter(f'{{{NSA}}}off')); te = next(tit.iter(f'{{{NSA}}}ext'))
            tx0, ty0, tcx = int(to.get('x')), int(to.get('y')), int(te.get('cx'))
            k = tcx / 15310848.0              # escala (slides em grupo usam outra escala)
            cx = tx0 + tcx // 2
            y = int(ty0 + (2200000 - 1008555) * k)
            w, h = int(8000000 * k), int(660527 * k)
        bp.set('wrap', 'square')   # caixa larga: nao quebra e fica centralizada
        off.set('x', str(cx - w // 2)); off.set('y', str(y))
        ext.set('cx', str(w)); ext.set('cy', str(h))
        for p in sp.iter(f'{{{NSA}}}pPr'):
            p.set('algn', 'ctr')

n_slides, slides_por_mes = detectar_estrutura(PPTX_IN)
n_meses_template = (n_slides - 1) // slides_por_mes
print(f"Template: {n_slides} slides, {slides_por_mes}/mes, {n_meses_template} meses")

# Detectar quais meses estao no TEMPLATE (pelo nome nos slides)
meses_no_template = []
with zipfile.ZipFile(PPTX_IN, 'r') as _zi:
    for t_idx in range(n_meses_template):
        sn = 2 + t_idx * slides_por_mes
        _path = f'ppt/slides/slide{sn}.xml'
        with _zi.open(_path) as _f:
            _root = etree.fromstring(_f.read())
        for _t in _root.iter(f'{{{NSA}}}t'):
            if _t.text and _t.text.strip().upper() in [m.upper() for m in MESES_PT]:
                for _i, _m in enumerate(MESES_PT):
                    if _m.upper() == _t.text.strip().upper():
                        meses_no_template.append(_i + 1)
                        break
                break

# Mapear mes do template -> indice no dados da planilha
mes_para_idx = {m['dias'][0]['data'].month: i for i, m in enumerate(meses)}

slide_map = []
for t_idx in range(min(len(meses_no_template), n_meses_template)):
    base = 2 + t_idx * slides_por_mes
    mes_num = meses_no_template[t_idx]
    if mes_num in mes_para_idx:
        data_idx = mes_para_idx[mes_num]
    else:
        data_idx = t_idx  # fallback
    slide_map.append((base, 'cal', data_idx))
    slide_map.append((base + 1, 'res', data_idx))
    if slides_por_mes == 3:
        slide_map.append((base + 2, 'meta', data_idx))

slides_editar = {s for s, _, _ in slide_map}
tmp = PPTX_OUT + ".tmp"

with zipfile.ZipFile(PPTX_IN, 'r') as zi:
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zo:
        for item in zi.infolist():
            m = re.search(r'slide(\d+)\.xml$', item.filename)
            if not (m and int(m.group(1)) in slides_editar):
                with zi.open(item) as f: zo.writestr(item, f.read())
        for slide_num, tipo, idx in slide_map:
            mes_data = meses[idx]; meta = get_meta(mes_data['nome'], idx)
            path = f'ppt/slides/slide{slide_num}.xml'
            print(f"  Slide {slide_num}: {tipo.upper()} -- {mes_data['nome']}")
            with zi.open(path) as f: root = etree.fromstring(f.read())
            if tipo == 'cal':    root = proc_calendario(root, mes_data)
            elif tipo == 'res':  root = proc_resumo(root, mes_data, meta, idx)
            elif tipo == 'meta': root = proc_metas(root, mes_data, meta)
            ajustar_nome_mes(root)
            zo.writestr(path, etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True))

os.replace(tmp, PPTX_OUT)
print(f"\nGerado: {PPTX_OUT}")
