"""Monta o painel HTML com os dados recalculados (e conferidos) a partir do arquivo."""
import json, pickle, sys, io, contextlib
sys.path.insert(0, sys.argv[4] if len(sys.argv) > 4 else '.')
from cap_engine import Periodo, cap_proposta, records, MESES
import painel_data as PD

t = pickle.load(open(sys.argv[1], 'rb'))
template = open(sys.argv[2], encoding='utf-8').read()
OUT = sys.argv[3]

# ── períodos ──
sem = PD.periodo_semana(t, 202645.0)
per = [sem] + [PD.periodo_meses(t, i0, i1) for i0 in range(12) for i1 in range(i0, 12)]

COLS = ['_k', 'CM (HRS)', 'RECURSO MÉDIO ATUAL', 'TOTAL RECURSO', 'HRS TURNO', 'OEE REAL', 'OEE NEC',
        'HC NECESSARIO', 'CARREGAMENTO', 'Soma Demanda Media', 'TAKT', 'TC MEDIO', 'FIFO']
linhas, lidx = [], {}
periodos = {}
for p in per:
    rows = []
    for r in p['cap']:
        k = (r['PROCESSO'], r['FAMILIA'], r['CLUSTER'], r['Categoria'])
        if k not in lidx:
            lidx[k] = len(linhas); linhas.append(list(k))
        rows.append([lidx[k]] + [None if r[c] is None else (round(r[c], 12) if isinstance(r[c], float) else r[c]) for c in COLS[1:]])
    n = 1 if p['tipo'] == 'Semanal' else int(p['id'][1:].split('-')[1]) - int(p['id'][1:].split('-')[0]) + 1
    periodos[p['id']] = dict(tipo=p['tipo'], rotulo=p['rotulo'], n_meses=n, rodape=p['rodape'],
                             complexidade=p['complexidade'], rows=rows)

# ── validação (CAP semana 45 célula a célula + indicadores) ──
saved = records(t['CAP_PROPOSTA'])
calc = {(r['PROCESSO'], r['FAMILIA']): r for r in sem['cap']}
vcols = ['CLUSTER', 'Categoria', 'CM (HRS)', 'RECURSO MÉDIO ATUAL', 'TOTAL RECURSO', 'HRS TURNO', 'OEE REAL',
         'OEE NEC', 'HC NECESSARIO', 'CARREGAMENTO', 'Soma Demanda Media', 'TAKT', 'TC MEDIO', 'FIFO']
dif = 0
for s in saved:
    c = calc.get((s['PROCESSO'], s['FAMILIA']))
    for col in vcols:
        a, b = s[col], c[col] if c else None
        if isinstance(a, (int, float)) and isinstance(b, (int, float)):
            dif += abs(a - b) > 1e-9 * max(1, abs(a))
        else:
            dif += not (a == b or (a in (None, '') and b in (None, '')))
assert len(saved) == len(calc) and dif == 0, 'CAP semana 45 não bate com o arquivo'

buf = io.StringIO()
with contextlib.redirect_stdout(buf):
    sys.argv = ['painel_data.py', sys.argv[1], '/dev/null']
    exec(compile(open('painel_data.py').read().split("if __name__ == '__main__':")[1].replace('\n    ', '\n'), 'pd', 'exec'),
         dict(PD.__dict__, t=t))
checks = []
for line in buf.getvalue().splitlines():
    if line.startswith(('OK ', 'FALHA ')):
        ok = line.startswith('OK')
        nome = line[5:].split(' arquivo=')[0].strip()
        arq = float(line.split('arquivo=')[1].split()[0])
        calc_v = float(line.split('recalculado=')[1])
        checks.append(dict(nome=nome, arq=arq, calc=calc_v, ok=ok))
assert checks and all(c['ok'] for c in checks)

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

achados = [
    dict(nivel='bad', tag='Rodapé', titulo='O rodapé depende do filtro manual da tabela',
         det='HC por turno real e HC necessário usam SUBTOTAL, que soma só as linhas visíveis. Hoje a tabela está filtrada em Categoria = Mão de Obra (50,98). Sem o filtro, o mesmo campo vira 92,98 e a ocupação muda sem aviso.'),
    dict(nivel='bad', tag='Rodapé', titulo='O bloco "Mês Setembro" mistura semana e mês',
         det='Ele usa os lotes de setembro (1.691) com o HC necessário da CAP que estiver carregada (hoje, a da semana 45: 59,48). Por isso mostra 117,79% de ocupação corrigida. Com a CAP de setembro de fato, o HC necessário por turno é 48,68 e a ocupação corrigida fica em 96,4%.'),
    dict(nivel='warn', tag='Rodapé', titulo='Números digitados que precisam ser trocados a cada período',
         det='"Semana 45", 4,58 dias úteis (copiado do calendário), "Mês Setembro", 22,97 dias úteis no mês, as referências DEMANDA!P1459/P1460 (coluna de setembro) e o "× 3" turnos. O 22,97 também está fixo no FIFO da consulta e não bate com o calendário (setembro tem 25 dias úteis; a média semanal × 52 ÷ 12 dá 23,28).'),
    dict(nivel='warn', tag='Rodapé', titulo='Critérios de lotes diferentes entre semana e mês',
         det='Na semana, "Total lotes" conta só os lotes com TC (414 de 416). No mês, conta a demanda total (1.691, inclusive lotes sem TC). A correção pelo % de lotes com TC é aplicada nos dois casos.'),
    dict(nivel='warn', tag='CAP', titulo='Categoria é digitada dentro da tabela da consulta',
         det='A coluna não vem do Power Query. Se uma atualização mudar a quantidade ou a ordem das linhas, a categoria pode ficar desalinhada e o rodapé soma linhas erradas. Hoje está correta (conferi contra Tabela_FAM[CATEGORIA]).'),
    dict(nivel='warn', tag='CAP', titulo='Horas calculadas que não aparecem em nenhuma linha',
         det='Espectro (AA, ICP, IV, UV) gera "MO EQUIPAMENTO BANCADA", mas essa família tem recurso 0 na Tabela_FAM e a linha é descartada: 68,75 h na semana 45, 392 h em setembro. O cromatógrafo CRGS0043 não está na Tabela_FAM: 25,75 h na semana e 214 h em setembro somem da CAP.'),
    dict(nivel='info', tag='CAP', titulo='HRS TURNO é arredondado só no HC necessário',
         det='A consulta FAM_PROPOSTA converte HRS TURNO para inteiro (104,22 → 104). TOTAL RECURSO usa 104,22 e HC NECESSARIO usa 104, uma diferença de ~0,2% entre as duas colunas.'),
    dict(nivel='info', tag='CAP', titulo='FIFO usa a constante mensal também na semana',
         det='FIFO = TOTAL RECURSO ÷ TAKT ÷ 22,97, que é o mesmo que demanda ÷ 22,97. Na visão semanal a demanda é da semana, então o número não representa lotes por dia.'),
    dict(nivel='info', tag='Semanal', titulo='A visão semanal só enxerga a semana que está no SharePoint',
         det='A Tabela_DemandaSemanal tem hoje apenas a chave 202645. Escolher outra semana no filtro devolve uma CAP vazia, a menos que a base do SharePoint tenha aquela semana.'),
    dict(nivel='info', tag='Arquivo', titulo='A versão de 25/09 ficou com a CAP desatualizada',
         det='A CAP salva não corresponde às bases do próprio arquivo (recursos e horas diferentes da aba FAMÍLIAS). Isso acontece quando uma base é editada e o "Atualizar Tudo" não é rodado; o desenho atual depende disso a cada troca de filtro.'),
    dict(nivel='info', tag='Arquivo', titulo='Cálculo paralelo pesado na BASE TC',
         det='~194 mil fórmulas com DESLOC (colunas CB a DM) refazem campanha e TC × demanda por fórmula, com divisores fixos (÷12, ÷5, ÷6). Não alimentam a CAP, deixam o arquivo lento e dão resultados diferentes dela.'),
]
decisoes = [
    'O bloco mensal do rodapé deve usar o HC da CAP do mesmo mês (96,4% em setembro) ou continuar comparando com a CAP carregada (117,79%)?',
    'Dias úteis no mês: manter 22,97 fixo ou usar o calendário de cada mês (ex.: setembro = 25)? Afeta lotes/dia e o FIFO.',
    'Espectro e CRGS0043: cadastrar recurso na Tabela_FAM para que essas horas entrem na CAP?',
    'HRS TURNO: manter o arredondamento para inteiro no HC necessário ou usar o valor exato?',
    'Os 3 turnos do "× 3" valem para todas as famílias de mão de obra?',
]

data = dict(cols=COLS, linhas=linhas, periodos=periodos,
            semanas=[dict(id=sem['id'], rotulo=sem['rotulo'])],
            exemplo=exemplo, achados=achados, decisoes=decisoes,
            validacao=dict(linhas=len(saved), valores=len(saved) * len(vcols), diferencas=dif, checks=checks))
js = json.dumps(data, ensure_ascii=False, separators=(',', ':'))
out = template.replace('/*__DATA__*/null', js)
open(OUT, 'w', encoding='utf-8').write(out)
print('ok', len(out) // 1024, 'KB', len(periodos), 'períodos', len(linhas), 'linhas distintas', len(checks), 'checks')
