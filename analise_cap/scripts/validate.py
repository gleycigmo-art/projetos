"""Compara a CAP recalculada (réplica) com a CAP_PROPOSTA salva no arquivo (valores do Excel)."""
import pickle, sys
sys.path.insert(0, sys.argv[2] if len(sys.argv) > 2 else '.')
from cap_engine import Periodo, cap_proposta, records

t = pickle.load(open(sys.argv[1], 'rb'))
filtro = {r['Parametro']: r['Valor'] for r in records(t['Tabela_FILTRO'])}
print('Filtro salvo:', filtro)
if filtro['Nível de Visão'] == 'Semanal':
    p = Periodo('Semanal', semana_chave=int(filtro['Semana (Chave)']))
else:
    p = Periodo('Mensal', filtro['Mês Início'], filtro['Mês Fim'])

calc = cap_proposta(t, p)
saved = records(t['CAP_PROPOSTA'])
cols = ['CLUSTER', 'Categoria', 'CM (HRS)', 'RECURSO MÉDIO ATUAL', 'TOTAL RECURSO', 'HRS TURNO', 'OEE REAL',
        'OEE NEC', 'HC NECESSARIO', 'CARREGAMENTO', 'Soma Demanda Media', 'TAKT', 'TC MEDIO', 'FIFO']
ck = {(r['PROCESSO'], r['FAMILIA']): r for r in calc}
sk = {(r['PROCESSO'], r['FAMILIA']): r for r in saved}
print('linhas salvas', len(saved), 'linhas calculadas', len(calc))
print('só no salvo:', set(sk) - set(ck))
print('só no calculado:', set(ck) - set(sk))
ndiff = 0
for k in sk:
    if k not in ck:
        continue
    for c in cols:
        a, b = sk[k][c], ck[k][c]
        if isinstance(a, (int, float)) and isinstance(b, (int, float)):
            ok = abs(a - b) <= 1e-9 * max(1, abs(a))
        else:
            ok = (a == b) or (a in (None, '') and b in (None, ''))
        if not ok:
            ndiff += 1
            print('DIF', k, c, 'salvo=', a, 'calc=', b)
print('TOTAL DIFERENÇAS:', ndiff)
