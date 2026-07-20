"""
SISTEMA DE PICK-UP AUTOMATICO v3

Fluxo:
1. Le relatorios HITS (todas as colunas, 1a aba)
2. Na planilha: Hoje -> Ontem, HITS -> Hoje, Pick-up = Hoje - Ontem
3. Pick-up total da emissao = soma de todos os pick-ups do mes
4. Salva historico de emissoes em pickup_historico.json
5. Calcula Diaria Media automaticamente (Receita / Ocup KPI)
6. Gera PPTXs

Arquivos HITS: hits_normandie.xlsx, hits_alto.xlsx, hits_chales.xlsx
Historico: pickup_historico.json (criado automaticamente)
"""
import sys, os, re, shutil, json, calendar
from datetime import datetime, date, timedelta
from collections import OrderedDict

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

HOTEIS = OrderedDict([
    ('normandie', {
        'hits_file': 'hits_normandie.xlsx',
        'aba_planilha': 'Normandie',
        'template_pptx': 'pickup_normandie_template.pptx',
    }),
    ('alto', {
        'hits_file': 'hits_alto.xlsx',
        'aba_planilha': 'Alto da Boa Vista',
        'template_pptx': 'pickup_alto_template.pptx',
    }),
    # ('chales', {  # REMOVIDO
    #     'hits_file': 'hits_chales.xlsx',
    #     'aba_planilha': 'CHALES VILLA CARAVAGGIO',
    #     'template_pptx': 'pickup_chales_vilaa_caravaggio_template.pptx',
    # }),
])

MESES_PT = ['Janeiro','Fevereiro','Marco','Abril','Maio','Junho',
            'Julho','Agosto','Setembro','Outubro','Novembro','Dezembro']

PASTA_PROGRAMA = sys.argv[2] if len(sys.argv) > 2 else '.'
PASTA_HITS = sys.argv[1] if len(sys.argv) > 1 else PASTA_PROGRAMA
PLANILHA = os.path.join(PASTA_PROGRAMA, 'Pick-up_sniper.xlsx')
HISTORICO_FILE = os.path.join(PASTA_PROGRAMA, 'pickup_historico.json')

FONT_DEFAULT = Font(name='Aptos Narrow', size=11)
FMT_RECEITA = '#,##0.00'
FMT_OCC = '#,##0.00'
FMT_DATA = 'mm-dd-yy'

# ══════════════════════════════════════════════════════════════════════════════
# HISTORICO DE EMISSOES
# ══════════════════════════════════════════════════════════════════════════════
def carregar_historico():
    if os.path.exists(HISTORICO_FILE):
        with open(HISTORICO_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

def salvar_historico(hist):
    with open(HISTORICO_FILE, 'w', encoding='utf-8') as f:
        json.dump(hist, f, indent=2, ensure_ascii=False)

def registrar_emissao(hist, hotel, data_emissao, pickup_por_mes):
    """Registra pick-up total de cada mes nesta emissao.
    pickup_por_mes = {mes_num: pickup_total_receita}"""
    if hotel not in hist:
        hist[hotel] = {}
    for mes, pu_total in pickup_por_mes.items():
        mes_key = str(mes)
        if mes_key not in hist[hotel]:
            hist[hotel][mes_key] = []
        hist[hotel][mes_key].append({
            'data': data_emissao,
            'pickup_total': round(pu_total, 2),
        })
        # Manter apenas ultimas 30 emissoes por mes
        hist[hotel][mes_key] = hist[hotel][mes_key][-30:]

# ══════════════════════════════════════════════════════════════════════════════
# LEITURA DO HITS
# ══════════════════════════════════════════════════════════════════════════════
def ler_hits(filepath):
    wb = load_workbook(filepath, data_only=True, read_only=True)
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

# ══════════════════════════════════════════════════════════════════════════════
# MAPEAMENTO DE MESES
# ══════════════════════════════════════════════════════════════════════════════
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

# ══════════════════════════════════════════════════════════════════════════════
# VIRADA DE MES
# ══════════════════════════════════════════════════════════════════════════════
def criar_mes_novo(ws, meses_existentes, novo_mes, novo_ano, meta_receita=0):
    if meses_existentes:
        col_nova = meses_existentes[-1]['col_base'] + 8
    else:
        col_nova = 1
    n_dias = calendar.monthrange(novo_ano, novo_mes)[1]
    nome_mes = MESES_PT[novo_mes - 1]
    print(f"    📅 Criando {nome_mes}/{novo_ano} ({n_dias} dias) na coluna {col_nova}")
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
        ws.cell(row=r, column=col_nova).number_format = FMT_DATA
        for offset in range(1, 7):
            cell = ws.cell(row=r, column=col_nova + offset)
            cell.value = 0
            cell.number_format = FMT_RECEITA if offset in (1,3,5) else FMT_OCC
    total_row = 3 + n_dias
    ws.cell(row=total_row, column=col_nova).value = 'TOTAL'
    ldr = 2 + n_dias
    for off, typ in [(1,'SUM'),(2,'AVERAGE'),(3,'SUM'),(4,'AVERAGE'),(5,'SUM'),(6,'AVERAGE')]:
        cl = get_column_letter(col_nova + off)
        ws.cell(row=total_row, column=col_nova+off).value = f'={typ}({cl}3:{cl}{ldr})'
        ws.cell(row=total_row, column=col_nova+off).number_format = FMT_RECEITA if off in (1,3,5) else FMT_OCC
    meta_row = total_row + 3
    meta_str = f'Meta: R$ {meta_receita:,.2f}'.replace(',','X').replace('.',',').replace('X','.')
    ws.cell(row=meta_row, column=col_nova + 2).value = meta_str
    return {'col_base': col_nova, 'mes': novo_mes, 'ano': novo_ano, 'total_row': total_row}

# ══════════════════════════════════════════════════════════════════════════════
# ATUALIZAR ABA
# ══════════════════════════════════════════════════════════════════════════════
def atualizar_aba(ws, dados_hits, meta_nova=None):
    meses = mapear_meses(ws)
    if not meses: return 0, False, {}
    datas_hits = sorted(dados_hits.keys())
    if not datas_hits: return 0, False, {}

    # Virada de mes
    meses_hits = set(d.month for d in datas_hits)
    meses_plan = set(m['mes'] for m in meses)
    mes_criado = False
    for mn in sorted(meses_hits - meses_plan):
        an = next(d.year for d in datas_hits if d.month == mn)
        info = criar_mes_novo(ws, meses, mn, an, meta_nova or 0)
        meses.append(info); mes_criado = True

    # Atualizar dias
    atualizados = 0
    pickup_por_mes = {}  # {mes_num: pickup_total_receita}

    for mi in meses:
        col = mi['col_base']
        pu_total_mes = 0.
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

            pu_total_mes += pu_r
            atualizados += 1

        pickup_por_mes[mi['mes']] = pu_total_mes

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

    # Diaria Media (Receita / Ocup KPI)
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
            print(f"    📊 DM {MESES_PT[mi['mes']-1]}: R$ {dm:.2f}")

    return atualizados, mes_criado, pickup_por_mes

# ══════════════════════════════════════════════════════════════════════════════
# ATUALIZAR METAS ALTO
# ══════════════════════════════════════════════════════════════════════════════
def atualizar_metas_alto(wb):
    if 'Alto da Boa Vista Metas' not in wb.sheetnames: return
    if 'Alto da Boa Vista' not in wb.sheetnames: return
    ws_dados = wb['Alto da Boa Vista']; ws_metas = wb['Alto da Boa Vista Metas']
    meses = mapear_meses(ws_dados)
    
    # Verificar se precisa criar seção de metas para mês novo
    # A aba de metas tem blocos de 7 linhas cada, começando na row 7 (col direita)
    # Cada bloco: row=dias_faltam, row+2=headers, row+3=receita, row+4=occ, row+5=dm
    n_blocos_existentes = 0
    for r in range(1, ws_metas.max_row + 1):
        v = ws_metas.cell(row=r, column=10).value  # Coluna J
        if v and isinstance(v, str) and 'Quantos dias' in v:
            n_blocos_existentes += 1
    
    if len(meses) > n_blocos_existentes:
        for idx in range(n_blocos_existentes, len(meses)):
            # Criar novo bloco na coluna direita
            base_row = 7 + idx * 7  # Row base do bloco
            mes_nome = MESES_PT[meses[idx]['mes'] - 1]
            
            ws_metas.cell(row=base_row, column=10).value = f'Quantos dias faltam - {mes_nome}'
            ws_metas.cell(row=base_row, column=13).value = 0  # Dias faltam (será calculado)
            
            ws_metas.cell(row=base_row + 2, column=10).value = mes_nome
            ws_metas.cell(row=base_row + 2, column=12).value = 'Acumulada'
            ws_metas.cell(row=base_row + 2, column=13).value = 'Meta'
            ws_metas.cell(row=base_row + 2, column=14).value = '% Atingimento'
            ws_metas.cell(row=base_row + 2, column=15).value = 'Falta por dia'
            
            ws_metas.cell(row=base_row + 3, column=11).value = 'Receita'
            ws_metas.cell(row=base_row + 3, column=12).value = 0  # Acumulada (será atualizada)
            ws_metas.cell(row=base_row + 3, column=13).value = 0  # ← PREENCHER META RECEITA
            ws_metas.cell(row=base_row + 3, column=13).font = Font(name='Aptos Narrow', size=11, color='FF0000')
            
            ws_metas.cell(row=base_row + 4, column=11).value = 'Ocupação'
            ws_metas.cell(row=base_row + 4, column=12).value = 0
            ws_metas.cell(row=base_row + 4, column=13).value = 0  # ← PREENCHER META OCUPAÇÃO
            ws_metas.cell(row=base_row + 4, column=13).font = Font(name='Aptos Narrow', size=11, color='FF0000')
            
            ws_metas.cell(row=base_row + 5, column=11).value = 'Diária Média'
            ws_metas.cell(row=base_row + 5, column=12).value = 0
            ws_metas.cell(row=base_row + 5, column=13).value = 0  # ← PREENCHER META DM
            ws_metas.cell(row=base_row + 5, column=13).font = Font(name='Aptos Narrow', size=11, color='FF0000')
            
            print(f"    🆕 Metas {mes_nome} criadas na planilha (preencha a coluna Meta em vermelho!)")
    
    meta_pos = [
        {'rec_row':10,'occ_row':11,'dm_row':12,'dias_row':7},
        {'rec_row':17,'occ_row':18,'dm_row':19,'dias_row':14},
        {'rec_row':24,'occ_row':25,'dm_row':26,'dias_row':21},
        {'rec_row':31,'occ_row':32,'dm_row':33,'dias_row':28},
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
        occ_media = total_o / n_dias if n_dias > 0 else 0.
        ws_metas.cell(row=mp['rec_row'], column=12).value = total_r
        ws_metas.cell(row=mp['occ_row'], column=12).value = occ_media/100 if occ_media > 1 else occ_media
        meta_r = ws_metas.cell(row=mp['rec_row'], column=13).value
        meta_o = ws_metas.cell(row=mp['occ_row'], column=13).value
        if meta_r and isinstance(meta_r, (int, float)) and float(meta_r) > 0:
            ws_metas.cell(row=mp['rec_row'], column=14).value = total_r / float(meta_r)
        if meta_o and isinstance(meta_o, (int, float)) and float(meta_o) > 0:
            ws_metas.cell(row=mp['occ_row'], column=14).value = (occ_media/100 if occ_media>1 else occ_media) / float(meta_o)
        # DM atingimento
        meta_dm_val = ws_metas.cell(row=mp['dm_row'], column=13).value
        acum_dm_val = ws_metas.cell(row=mp['dm_row'], column=12).value
        if meta_dm_val and acum_dm_val and isinstance(meta_dm_val, (int, float)) and float(meta_dm_val) > 0:
            ws_metas.cell(row=mp['dm_row'], column=14).value = float(acum_dm_val) / float(meta_dm_val)
        ultimo_dia = date(mi['ano'], mi['mes'], calendar.monthrange(mi['ano'], mi['mes'])[1])
        dias_faltam = max(0, (ultimo_dia - date.today()).days + 1)  # +1 inclui hoje
        ws_metas.cell(row=mp['dias_row'], column=13).value = dias_faltam
        if dias_faltam > 0 and meta_r and isinstance(meta_r, (int, float)):
            ws_metas.cell(row=mp['rec_row'], column=15).value = (float(meta_r) - total_r) / dias_faltam

# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════
def main():
    print("=" * 60)
    print("  SISTEMA DE PICK-UP AUTOMATICO v3")
    print(f"  Data: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    print("=" * 60)

    if not os.path.exists(PLANILHA):
        print(f"\n❌ Planilha nao encontrada: {PLANILHA}"); sys.exit(1)

    # Verificar HITS
    hits_disponiveis = {}
    for hk, cfg in HOTEIS.items():
        hp = os.path.join(PASTA_HITS, cfg['hits_file'])
        if os.path.exists(hp):
            dados = ler_hits(hp)
            if dados:
                datas = sorted(dados.keys())
                hits_disponiveis[hk] = {'dados': dados, 'datas': datas}
                print(f"\n  ✅ {cfg['hits_file']}: {len(dados)} dias ({datas[0].strftime('%d/%m')} a {datas[-1].strftime('%d/%m')})")
        else:
            print(f"\n  ⚠️  {cfg['hits_file']}: nao encontrado")

    if not hits_disponiveis:
        print("\n❌ Nenhum relatorio HITS encontrado!"); sys.exit(1)

    # Data da emissao = dia(s) anterior(es)
    # Segunda: emissao cobre sexta+sabado+domingo -> "24, 25, 26/04"
    # Terca a sexta: emissao cobre so o dia anterior -> "27/04"
    hoje = date.today()
    dow = hoje.weekday()  # 0=seg, 1=ter, ... 4=sex, 5=sab, 6=dom
    if dow == 0:  # Segunda - cobre sexta, sabado e domingo
        sexta = hoje - timedelta(days=3)
        sabado = hoje - timedelta(days=2)
        domingo = hoje - timedelta(days=1)
        data_emissao = f"{sexta.day}, {sabado.day}, {domingo.strftime('%d/%m')}"
    elif dow == 5:  # Sabado (não deveria emitir, mas por segurança)
        data_emissao = (hoje - timedelta(days=1)).strftime('%d/%m')
    elif dow == 6:  # Domingo
        data_emissao = (hoje - timedelta(days=2)).strftime('%d/%m')
    else:  # Terca a sexta - dia anterior
        data_emissao = (hoje - timedelta(days=1)).strftime('%d/%m')
    
    print(f"\n  📅 Data da emissao: {data_emissao}")

    # Backup
    backup = PLANILHA.replace('.xlsx', f'_backup_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
    shutil.copy2(PLANILHA, backup)
    print(f"  📋 Backup: {os.path.basename(backup)}")

    # Carregar historico
    hist = carregar_historico()

    # Verificar virada de mes
    wb = load_workbook(PLANILHA)
    metas_novas = {}
    for hk, hi in hits_disponiveis.items():
        cfg = HOTEIS[hk]
        if cfg['aba_planilha'] not in wb.sheetnames: continue
        ws = wb[cfg['aba_planilha']]
        mp = mapear_meses(ws)
        novos = set(d.month for d in hi['datas']) - set(m['mes'] for m in mp)
        if novos:
            for mn in sorted(novos):
                nome = MESES_PT[mn - 1]
                print(f"\n🆕 MES NOVO: {nome} ({hk.upper()})")
                try:
                    meta_str = input(f"   Meta de receita para {nome}: R$ ") if sys.stdin.isatty() else "0"
                    meta = float(meta_str.replace('.','').replace(',','.').replace('R$','').strip() or '0')
                except: meta = 0
                metas_novas[(hk, mn)] = meta

    # Processar
    print(f"\n{'='*50}\n📊 ATUALIZANDO PLANILHA...\n{'='*50}")

    for hk, hi in hits_disponiveis.items():
        cfg = HOTEIS[hk]
        if cfg['aba_planilha'] not in wb.sheetnames: continue
        print(f"\n  🏨 {hk.upper()}")
        ws = wb[cfg['aba_planilha']]
        meta = next((v for (k,m),v in metas_novas.items() if k==hk), None)
        n, mc, pu_mes = atualizar_aba(ws, hi['dados'], meta)
        print(f"  ✅ {n} dias atualizados")

        # Registrar emissao no historico
        registrar_emissao(hist, hk, data_emissao, pu_mes)
        for mes, pu in pu_mes.items():
            print(f"    📈 {MESES_PT[mes-1]}: pick-up total = R$ {pu:,.2f}".replace(',','X').replace('.',',').replace('X','.'))

    if 'alto' in hits_disponiveis:
        atualizar_metas_alto(wb)
        print(f"\n  📊 Metas do Alto atualizadas")

    wb.save(PLANILHA)
    salvar_historico(hist)
    print(f"\n💾 Planilha salva")
    print(f"💾 Historico salvo: {os.path.basename(HISTORICO_FILE)}")

    # Gerar PPTXs
    print(f"\n{'='*50}\n📊 GERANDO PPTXs...\n{'='*50}")
    script_pptx = os.path.join(PASTA_PROGRAMA, 'gerar_pickup_pptx.py')
    if not os.path.exists(script_pptx):
        print(f"  ⚠️  gerar_pickup_pptx.py nao encontrado")
    else:
        for hk in hits_disponiveis:
            cfg = HOTEIS[hk]
            tmpl = os.path.join(PASTA_PROGRAMA, cfg['template_pptx'])
            if not os.path.exists(tmpl): continue
            saida = os.path.join(PASTA_PROGRAMA, f'pickup_{hk}_gerado.pptx')
            print(f"\n  🔄 {hk.upper()}...")
            ret = os.system(f'python "{script_pptx}" "{PLANILHA}" "{tmpl}" "{hk}" "{saida}" "{HISTORICO_FILE}"')
            if ret == 0: print(f"  ✅ {os.path.basename(saida)}")
            else: print(f"  ❌ Erro ao gerar {hk}")

    print(f"\n{'='*60}\n  ✅ CONCLUIDO!\n{'='*60}")

if __name__ == '__main__':
    main()