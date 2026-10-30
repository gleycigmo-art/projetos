"""Compara a aba PAINEL recalculada pelo LibreOffice com o motor Python (regras novas)."""
import json, sys
import openpyxl
from painel_data import media_meses
from gerar_painel_excel import INDICADORES, CAPCOLS, R_CAP0, N_CAP

per = {p['id']: p for p in json.load(open(sys.argv[1]))}
CASOS = {'sem': [per['S202645']], 'set': [per['M9']], 'jul_set': [per['M7'], per['M8'], per['M9']]}
CHAVES = ['hc_turno_real', 'hc_real_total', 'hc_turno_nec', 'hc_nec_total', 'ocupacao', 'pct_lotes_tc',
          'hc_turno_corr', 'hc_total_corr', 'ocupacao_corr', 'total_lotes', 'dias_uteis', 'lotes_dia', 'lotes_dia_100']
falhas = total = 0
for caso, meses in CASOS.items():
    ref = meses[0] if len(meses) == 1 else media_meses(meses)
    ws = openpyxl.load_workbook(f'{sys.argv[2]}/painel_{caso}.xlsx', data_only=True)['PAINEL']
    n = 0
    for i, k in enumerate(CHAVES):
        a, b = ws[f'C{9 + i}'].value, ref['rodape'][k]
        ok = isinstance(a, (int, float)) and abs(a - b) <= 1e-9 * max(1, abs(b))
        falhas += not ok; n += 1
        if not ok:
            print('FALHA', caso, INDICADORES[i][0], 'excel=', a, 'python=', b)
    for i, k in enumerate(['1', '2', '3', '4', '0', 'sem']):
        a, b = ws[f'G{10 + i}'].value, ref['complexidade'].get(k, 0)
        ok = isinstance(a, (int, float)) and abs(a - b) <= 1e-9 * max(1, abs(b))
        falhas += not ok; n += 1
        if not ok:
            print('FALHA', caso, 'complexidade', k, 'excel=', a, 'python=', b)
    py = {(r['PROCESSO'], r['FAMILIA']): r for r in ref['cap']}
    linhas = 0
    for r in range(R_CAP0, R_CAP0 + N_CAP):
        proc, fam = ws[f'B{r}'].value, ws[f'C{r}'].value
        if not fam:
            continue
        linhas += 1
        alvo = py[(proc, fam)]
        for col, nome, _ in CAPCOLS[2:]:
            a, b = ws[f'{col}{r}'].value, alvo[nome]
            if isinstance(b, (int, float)):
                ok = isinstance(a, (int, float)) and abs(a - b) <= 1e-9 * max(1, abs(b))
            else:
                ok = (a or None) == (b or None)
            falhas += not ok; n += 1
            if not ok:
                print('FALHA', caso, fam, nome, 'excel=', a, 'python=', b)
    conf = [ws[f'H{r}'].value for r in (19, 20, 21)]
    total += n
    print(f'{caso}: {n} valores conferidos, {linhas} linhas na CAP, conferência com CAP antiga (CM, HC, carreg.) = {conf}')
print('FALHAS:', falhas)
if len(sys.argv) > 3:
    json.dump(dict(valores=total, falhas=falhas, casos='semana 45, setembro e julho a setembro'), open(sys.argv[3], 'w'))
