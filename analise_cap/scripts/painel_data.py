"""Dados do painel (CAP + indicadores + complexidade) e conferências contra o arquivo.

Regras novas (decididas em 30/10/2026):
- Mês único: CAP e indicadores daquele mês. Intervalo: média simples dos meses (feita no painel / no Excel).
- Dias úteis: calendário do período (semana: CALENDARIO_SEMANA; mês: CALENDARIO_MES). Vale para lotes/dia e FIFO.
- Demais regras iguais ao arquivo (HRS TURNO inteiro no HC, 3 turnos, espectro/CRGS0043 fora da CAP).
"""
import json, pickle, sys
from cap_engine import Periodo, cap_proposta, dias_uteis, records, MESES, to_num

MES_UTIL_ARQUIVO = 22.9716666666667   # constante digitada em CAP_PROPOSTA!N71 (regra antiga)
TURNOS = 3                            # "×3" das fórmulas C64 / L65 / L67


def rodape(cap, pct_tc, total_lotes, dias):
    mo = [r for r in cap if r['Categoria'] == 'Mão de Obra']
    hc_real = round(sum(r['RECURSO MÉDIO ATUAL'] for r in mo), 10)
    hc_nec = round(sum(r['HC NECESSARIO'] for r in mo), 10)
    out = dict(hc_turno_real=hc_real, hc_real_total=hc_real * TURNOS,
               hc_turno_nec=hc_nec, hc_nec_total=hc_nec * TURNOS,
               pct_lotes_tc=pct_tc, total_lotes=total_lotes, dias_uteis=dias)
    out['ocupacao'] = out['hc_nec_total'] / out['hc_real_total']
    out['lotes_dia'] = total_lotes / dias
    out['hc_turno_corr'] = hc_nec / pct_tc
    out['hc_total_corr'] = out['hc_turno_corr'] * TURNOS
    out['ocupacao_corr'] = out['hc_total_corr'] / out['hc_real_total']
    out['lotes_dia_100'] = out['lotes_dia'] / out['ocupacao_corr']
    return out


def _nivel(c):
    return str(int(c)) if isinstance(c, (int, float)) else 'sem'


def complexidade_semana(t, chave):
    acc = {}
    for r in records(t['Tabela_DemandaSemanal']):
        if r['Chave semana'] == chave:
            k = _nivel(r['Complexidade'])
            acc[k] = acc.get(k, 0) + (r['Lote semana'] or 0)
    return acc


def complexidade_mes(t, mes):
    acc = {}
    for r in records(t['Tabela_DEMANDA']):
        if r['Local de Análise'] == 'CQ P30':
            k = _nivel(r['complexidade'])
            acc[k] = acc.get(k, 0) + to_num(r[mes])
    return acc


def lotes_mes(t, mes):
    """(demanda total CQ P30, demanda CQ P30 com TC levantado) — mesmas regras das linhas 1459/1460 da aba DEMANDA."""
    tot = sim = 0
    for r in records(t['Tabela_DEMANDA']):
        if r['Local de Análise'] == 'CQ P30':
            v = to_num(r[mes])
            tot += v
            if str(r['TC levantados?']).upper() == 'SIM':
                sim += v
    return tot, sim


def semana(t, chave, dias_fifo=None):
    p = Periodo('Semanal', semana_chave=chave)
    dias = dias_uteis(t, p)
    cap = cap_proposta(t, p, dias_fifo=dias if dias_fifo is None else dias_fifo)
    ds = [r for r in records(t['Tabela_DemandaSemanal']) if r['Chave semana'] == chave]
    tot = sum(r['Lote semana'] or 0 for r in ds)
    ok = sum(r['Lote semana'] or 0 for r in ds if r['Status2'] == 'Tempos OK')
    rot = [r for r in records(t['MAPA_SEMANAS']) if r['Semana Chave'] == chave][0]['Rótulo (uso no filtro)']
    return dict(id=f'S{int(chave)}', tipo='Semanal', rotulo=rot, ordem=int(chave), cap=cap,
                rodape=rodape(cap, ok / tot, ok, dias), complexidade=complexidade_semana(t, chave))


def mes(t, i):
    m = MESES[i]
    p = Periodo('Mensal', m, m)
    dias = dias_uteis(t, p)
    cap = cap_proposta(t, p, dias_fifo=dias)
    tot, sim = lotes_mes(t, m)
    return dict(id=f'M{i + 1}', tipo='Mensal', rotulo=m, ordem=i + 1, cap=cap,
                rodape=rodape(cap, sim / tot, tot, dias), complexidade=complexidade_mes(t, m))


def close(a, b, tol=1e-9):
    return abs(a - b) <= tol * max(1, abs(a))


def conferencias(t):
    """Confere a réplica (com as regras ANTIGAS) contra os valores gravados no arquivo."""
    s = semana(t, 202645.0, dias_fifo=22.97)
    salvo_b = {r[0]: r[1] for r in t['CAP_RODAPE']['rows'] if isinstance(r[0], str)}
    rs = s['rodape']
    checks = [
        ('HC por turno Real', salvo_b['HC por turno Real'], rs['hc_turno_real']),
        ('HC real total', salvo_b['HC real total'], rs['hc_real_total']),
        ('% Lotes c/ TC (sem)', 0.9951923076923077, rs['pct_lotes_tc']),
        ('Total lotes (sem)', 414.0, rs['total_lotes']),
        ('Semana útil', 4.58045222465354, rs['dias_uteis']),
        ('Lotes/dia (sem)', 90.38408866524406, rs['lotes_dia']),
        ('HC p/ turno Nec', 59.48, rs['hc_turno_nec']),
        ('HC Nec total', 178.44, rs['hc_nec_total']),
        ('% Ocupação', 1.1667320517850137, rs['ocupacao']),
        ('HC p/ turno Corrigido', 59.76734299516908, rs['hc_turno_corr']),
        ('HC total Corrigido', 179.30202898550723, rs['hc_total_corr']),
        ('% Ocupação Corrigido', 1.1723684385086128, rs['ocupacao_corr']),
        ('Lotes/dia a 100%', 77.09529333647279, rs['lotes_dia_100']),
    ]
    tot, sim = lotes_mes(t, 'Set/2026')
    set_arq = rodape(s['cap'], sim / tot, tot, MES_UTIL_ARQUIVO)
    checks += [
        ('Set: % Lotes c/ TC', 0.9905381431105854, set_arq['pct_lotes_tc']),
        ('Set: lotes', 1691.0, set_arq['total_lotes']),
        ('Set: Lotes/dia', 73.61242109845462, set_arq['lotes_dia']),
        ('Set: HC p/ turno Corrigido', 60.048167164179105, set_arq['hc_turno_corr']),
        ('Set: HC total Corrigido', 180.1445014925373, set_arq['hc_total_corr']),
        ('Set: % Ocupação Corrigido', 1.1778769549662436, set_arq['ocupacao_corr']),
        ('Set: Lotes/dia a 100%', 62.495849662466874, set_arq['lotes_dia_100']),
    ]
    cs = s['complexidade']
    checks += [(f'Complex. semana nível {k}', v, cs[k]) for k, v in (('1', 18), ('2', 125), ('3', 166), ('4', 102))]
    cm = complexidade_mes(t, 'Set/2026')
    t14 = sum(cm[k] for k in '1234')
    checks += [(f'Complex. Set nível {k}', v, cm[k] / t14) for k, v in
               (('1', 0.06453715775749674), ('2', 0.2907431551499348), ('3', 0.39895697522816165), ('4', 0.2457627118644068))]
    return [dict(nome=n, arq=a, calc=b, ok=close(a, b)) for n, a, b in checks]


def media_meses(meses):
    """Regra de intervalo: média simples, mês a mês, de cada número (vazios ignorados, como MÉDIASES)."""
    def avg(vals):
        vals = [v for v in vals if v is not None]
        return sum(vals) / len(vals) if vals else None
    rod = {k: avg([m['rodape'][k] for m in meses]) for k in meses[0]['rodape']}
    niveis = set().union(*[m['complexidade'] for m in meses])
    cx = {k: avg([m['complexidade'].get(k, 0) for m in meses]) for k in niveis}
    cap = []
    for i, r in enumerate(meses[0]['cap']):
        linha = dict(r)
        for c, v in r.items():
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                linha[c] = avg([m['cap'][i][c] for m in meses])
        cap.append(linha)
    return dict(rodape=rod, complexidade=cx, cap=cap)


if __name__ == '__main__':
    t = pickle.load(open(sys.argv[1], 'rb'))
    falhas = 0
    for c in conferencias(t):
        falhas += not c['ok']
        print('OK   ' if c['ok'] else 'FALHA', f"{c['nome']:32s} arquivo={c['arq']!r:<22} recalculado={c['calc']!r}")
    print('FALHAS:', falhas)
    per = [semana(t, 202645.0)] + [mes(t, i) for i in range(12)]
    json.dump(per, open(sys.argv[2], 'w'), ensure_ascii=False, default=str)
    print('períodos:', len(per))
