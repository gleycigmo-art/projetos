"""Monta o painel HTML (protótipo) com os dados das regras novas, conferidos contra o arquivo.

Uso: python3 build_painel.py tabelas.pkl painel_template.html saida.html teste_excel.json
"""
import json, pickle, sys
from cap_engine import records
import painel_data as PD

TAB, TPL, OUT, TESTE = sys.argv[1:5]
t = pickle.load(open(TAB, 'rb'))
template = open(TPL, encoding='utf-8').read()
teste_excel = json.load(open(TESTE))

# ── base: a réplica (regras antigas) tem de bater com o arquivo ──
semana_arquivo = PD.semana(t, 202645.0, dias_fifo=22.97)
saved = records(t['CAP_PROPOSTA'])
calc = {(r['PROCESSO'], r['FAMILIA']): r for r in semana_arquivo['cap']}
vcols = ['CLUSTER', 'Categoria', 'CM (HRS)', 'RECURSO MÉDIO ATUAL', 'TOTAL RECURSO', 'HRS TURNO', 'OEE REAL',
         'OEE NEC', 'HC NECESSARIO', 'CARREGAMENTO', 'Soma Demanda Media', 'TAKT', 'TC MEDIO', 'FIFO']
dif = 0
for s in saved:
    c = calc[(s['PROCESSO'], s['FAMILIA'])]
    for col in vcols:
        a, b = s[col], c[col]
        if isinstance(a, (int, float)) and isinstance(b, (int, float)):
            dif += abs(a - b) > 1e-9 * max(1, abs(a))
        else:
            dif += not (a == b or (a in (None, '') and b in (None, '')))
checks = PD.conferencias(t)
assert len(saved) == len(calc) and dif == 0 and all(c['ok'] for c in checks), 'réplica não bate com o arquivo'
assert teste_excel['falhas'] == 0, 'fórmulas do PAINEL não bateram no teste'

# ── períodos com as regras novas ──
per = [PD.semana(t, 202645.0)] + [PD.mes(t, i) for i in range(12)]
COLS = ['_k', 'CM (HRS)', 'RECURSO MÉDIO ATUAL', 'TOTAL RECURSO', 'HRS TURNO', 'OEE REAL', 'OEE NEC',
        'HC NECESSARIO', 'CARREGAMENTO', 'Soma Demanda Media', 'TAKT', 'TC MEDIO', 'FIFO']
linhas, lidx, periodos = [], {}, {}
for p in per:
    rows = []
    for r in p['cap']:
        k = (r['PROCESSO'], r['FAMILIA'], r['CLUSTER'], r['Categoria'])
        if k not in lidx:
            lidx[k] = len(linhas); linhas.append(list(k))
        rows.append([lidx[k]] + [r[c] for c in COLS[1:]])
    rows.sort(key=lambda x: x[0])   # mesma ordem em todos os períodos (a média de intervalo usa a posição)
    periodos[p['id']] = dict(tipo=p['tipo'], rotulo=p['rotulo'], n_meses=1, rodape=p['rodape'],
                             complexidade=p['complexidade'], rows=rows)
assert len({tuple(x[0] for x in p['rows']) for p in periodos.values()}) == 1

s45, s45a = per[0], semana_arquivo
fifo_dis = [r['FIFO'] for r in s45['cap'] if r['FAMILIA'] == 'Fam Dissolutor'][0]
fifo_dis_a = [r['FIFO'] for r in s45a['cap'] if r['FAMILIA'] == 'Fam Dissolutor'][0]
setv = periodos['M9']['rodape']

exemplo = """8 lotes na semana · campanha de 4,22 lotes → ARRED.PARA.CIMA = 5 lotes/campanha → nCamp = ARRED.PARA.CIMA(8 ÷ 5) = 2

Etapa 1 · preparo        100 min, repete 0%    →  2 × 100 + (8 − 2) × 100 × 0%   =   200 min
Etapa 2 · prep. amostra   90 min, sem campanha →  90 × 8                         =   720 min
Etapa 3 · análise FQ       0 min               →                                       0 min
Etapa 6 · finalização     25 min, repete 50%   →  2 × 25 + (8 − 2) × 25 × 50%     =   125 min
MO BANCADA (Sólidos Dipirona)                                                    = 1.045 min   (sem campanha seriam 1.720)

Etapa 7 · conferência     25 min, repete 50%   →  2 × 25 + 6 × 25 × 50%           =   125 min  → MO CONFERÊNCIA BANCADA
Etapa 9 · montagem        45 min, repete 10%   →  2 × 45 + 6 × 45 × 10%           =   117 min  → MO EQUIPAMENTO
Etapa 14 · processamento  30 min               →  30 × 8                          =   240 min  → MO EQUIPAMENTO
HPLC CRLQ0043: padrão 6 inj × 10 min = 60 · amostra 10 inj × 10 min = 100
   por lote = (2 × 60 + 6 × 60 × 10%) ÷ 8 + 100 = 119,5 min  × 8 lotes             =   956 min  → linha CRLQ0043"""

def n2(v, d=2):
    return f'{v:,.{d}f}'.replace(',', 'X').replace('.', ',').replace('X', '.')

achados = [
    dict(nivel='ok', tag='Resolvido no PAINEL', titulo='O rodapé dependia do filtro manual da tabela',
         det='HC por turno real e HC necessário usavam SUBTOTAL (só linhas visíveis). No PAINEL, a soma é por Categoria = Mão de Obra, independente de filtro.'),
    dict(nivel='ok', tag='Resolvido no PAINEL', titulo='O bloco "Mês Setembro" misturava semana e mês',
         det=f'Usava os lotes de setembro com o HC da CAP da semana 45 (117,79%). Agora cada mês usa a própria CAP: setembro tem HC necessário por turno de {n2(setv["hc_turno_nec"])} e ocupação corrigida de {n2(setv["ocupacao_corr"] * 100, 1)}%. Intervalo de meses = média dos meses.'),
    dict(nivel='ok', tag='Resolvido no PAINEL', titulo='Números digitados a cada período',
         det=f'"Semana 45", 4,58, "Mês Setembro", 22,97 e as referências à coluna de setembro saíram. Dias úteis vêm do calendário: semana 45 = 4,58; setembro = 25. Com isso, lotes/dia de setembro passa de 73,6 para {n2(setv["lotes_dia"], 1)}.'),
    dict(nivel='ok', tag='Resolvido no PAINEL', titulo='FIFO usava a constante mensal também na semana',
         det=f'FIFO agora é demanda ÷ dias úteis do período. Fam Dissolutor na semana 45: {n2(fifo_dis_a)} → {n2(fifo_dis)}.'),
    dict(nivel='ok', tag='Resolvido no PAINEL', titulo='Categoria era digitada dentro da tabela da consulta',
         det='Na consulta nova, a Categoria vem da Tabela_FAM. Conferido: igual às 55 linhas atuais.'),
    dict(nivel='info', tag='Mantido (decisão)', titulo='Horas calculadas que não aparecem em nenhuma linha',
         det='Espectro (68,75 h na semana 45; 392 h em setembro) e CRGS0043 (25,75 h; 214 h) seguem fora da CAP porque não há recurso cadastrado para eles.'),
    dict(nivel='info', tag='Mantido (decisão)', titulo='HRS TURNO arredondado só no HC necessário',
         det='TOTAL RECURSO usa 104,22 h e HC NECESSARIO usa 104 h (diferença de ~0,2%).'),
    dict(nivel='warn', tag='Atenção', titulo='Critérios de lotes diferentes entre semana e mês',
         det='Na semana, "Lotes no período" conta só os lotes com TC (414 de 416). No mês, conta a demanda total CQ P30 (1.691). Mantido como no arquivo.'),
    dict(nivel='info', tag='Semanal', titulo='A visão semanal só enxerga as semanas que estão no SharePoint',
         det='A base de quarentena traz hoje só a semana 202645. Para semanas passadas é preciso guardar um retrato por semana.'),
    dict(nivel='info', tag='Arquivo', titulo='A versão de 25/09 ficou com a CAP desatualizada',
         det='A CAP salva não corresponde às bases do próprio arquivo. No PAINEL novo, trocar o período não exige atualizar consultas, o que reduz esse risco.'),
    dict(nivel='info', tag='Arquivo', titulo='Cálculo paralelo pesado na BASE TC',
         det='~194 mil fórmulas com DESLOC (colunas CB a DM) refazem campanha e TC × demanda por fórmula, com divisores fixos (÷12, ÷5, ÷6). Não alimentam a CAP e deixam o arquivo lento.'),
]
decisoes = [
    'Mês filtrado: CAP e indicadores daquele mês. Intervalo filtrado: média simples dos meses do intervalo.',
    'Dias úteis: calendário de cada mês e de cada semana (aba FAMÍLIAS). Vale para lotes/dia e FIFO.',
    'Espectro e CRGS0043: ficam como estão (sem recurso cadastrado).',
    'HRS TURNO: mantém o arredondamento para inteiro no HC necessário.',
    '× 3 turnos para todas as famílias de mão de obra.',
]
mudancas = [
    f'FIFO de todas as linhas: demanda ÷ dias úteis do período (antes ÷ 22,97). Ex.: Fam Dissolutor na semana 45, {n2(fifo_dis_a)} → {n2(fifo_dis)}.',
    'Categoria passa a vir da consulta; os valores são os mesmos de hoje.',
    'Visão mensal: cada mês usa a própria CAP e o próprio calendário; o bloco "Mês Setembro" do arquivo (117,79%) dá lugar ao setembro calculado de fato '
    f'({n2(setv["ocupacao_corr"] * 100, 1)}% de ocupação corrigida, {n2(setv["lotes_dia"], 1)} lotes/dia).',
    'Intervalo de meses: média simples dos meses (antes a consulta calculava um "mês médio").',
    'Todo o resto da semana 45 (CM, HC, carregamento, demanda, TAKT, TC médio e o rodapé) segue idêntico ao arquivo.',
]

data = dict(cols=COLS, linhas=linhas, periodos=periodos,
            semanas=[dict(id=per[0]['id'], rotulo=per[0]['rotulo'])],
            exemplo=exemplo, achados=achados, decisoes=decisoes, mudancas=mudancas,
            validacao=dict(linhas=len(saved), valores=len(saved) * len(vcols), diferencas=dif,
                           checks=checks, excel=teste_excel))
out = template.replace('/*__DATA__*/null', json.dumps(data, ensure_ascii=False, separators=(',', ':')))
open(OUT, 'w', encoding='utf-8').write(out)
print('ok', len(out) // 1024, 'KB')
