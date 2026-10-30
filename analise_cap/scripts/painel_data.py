"""Gera os dados do painel (CAP + indicadores + complexidade) para todos os períodos
e confere, contra os valores salvos no arquivo, tudo o que o arquivo permite conferir."""
import json, pickle, sys
sys.path.insert(0, sys.argv[3] if len(sys.argv) > 3 else '.')
from cap_engine import Periodo, cap_proposta, records, MESES, to_num

MES_UTIL = 22.9716666666667   # constante digitada em CAP_PROPOSTA!N71 (mesmo 22,97 usado no FIFO do PQ)
TURNOS = 3                    # "×3" fixo das fórmulas C64 / L65 / L67


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
    out['lotes_dia_100'] = 1 * out['lotes_dia'] / out['ocupacao_corr']
    return out


def complexidade_semana(t, chave):
    acc = {}
    for r in records(t['Tabela_DemandaSemanal']):
        if r['Chave semana'] != chave:
            continue
        c = r['Complexidade']
        k = str(int(c)) if isinstance(c, (int, float)) else 'sem'
        acc[k] = acc.get(k, 0) + (r['Lote semana'] or 0)
    return acc


def complexidade_meses(t, meses):
    acc = {}
    for r in records(t['Tabela_DEMANDA']):
        if r['Local de Análise'] != 'CQ P30':
            continue
        c = r['complexidade']
        k = str(int(c)) if isinstance(c, (int, float)) else 'sem'
        for m in meses:
            acc[k] = acc.get(k, 0) + to_num(r[m])
    n = len(meses)
    return {k: v / n for k, v in acc.items()}


def linhas_rodape_demanda(t):
    rows = {r[0]: r[1:] for r in t['DEMANDA_RODAPE']['rows'] if isinstance(r[0], str)}
    return rows['Demanda total'], rows['Sim']


def periodo_semana(t, chave):
    p = Periodo('Semanal', semana_chave=chave)
    cap = cap_proposta(t, p)
    ds = [r for r in records(t['Tabela_DemandaSemanal']) if r['Chave semana'] == chave]
    tot = sum(r['Lote semana'] or 0 for r in ds)
    ok = sum(r['Lote semana'] or 0 for r in ds if r['Status2'] == 'Tempos OK')
    cal = [r for r in records(t['CALENDARIO_SEMANA']) if r['Semana Chave'] == chave][0]
    rot = [r for r in records(t['MAPA_SEMANAS']) if r['Semana Chave'] == chave][0]['Rótulo (uso no filtro)']
    return dict(id=f'S{int(chave)}', tipo='Semanal', rotulo=rot, cap=cap,
                rodape=rodape(cap, ok / tot, ok, cal['Dias Úteis']),
                complexidade=complexidade_semana(t, chave))


def periodo_meses(t, i0, i1):
    meses = MESES[i0:i1 + 1]
    p = Periodo('Mensal', meses[0], meses[-1])
    cap = cap_proposta(t, p)
    dem_tot, dem_sim = linhas_rodape_demanda(t)
    tot = sum(dem_tot[i] for i in range(i0, i1 + 1))
    sim = sum(dem_sim[i] for i in range(i0, i1 + 1))
    rot = meses[0] if i0 == i1 else f'{meses[0]} a {meses[-1]} (média mensal)'
    return dict(id=f'M{i0 + 1}-{i1 + 1}', tipo='Mensal', rotulo=rot, cap=cap,
                rodape=rodape(cap, sim / tot, tot / len(meses), MES_UTIL),
                complexidade=complexidade_meses(t, meses))


def close(a, b, tol=1e-9):
    return abs(a - b) <= tol * max(1, abs(a))


if __name__ == '__main__':
    t = pickle.load(open(sys.argv[1], 'rb'))
    sem = periodo_semana(t, 202645.0)

    # ── Conferências contra o arquivo ──
    salvo = {r[9]: (r[10], r[11], r[12]) for r in t['CAP_RODAPE']['rows'] if isinstance(r[9], str)}
    salvo_b = {r[0]: r[1] for r in t['CAP_RODAPE']['rows'] if isinstance(r[0], str)}
    rs = sem['rodape']
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
    # Bloco "Mês Setembro" exatamente como o arquivo faz (HC da CAP semanal + lotes de Setembro)
    dem_tot, dem_sim = linhas_rodape_demanda(t)
    set_file = rodape(sem['cap'], dem_sim[8] / dem_tot[8], dem_tot[8], MES_UTIL)
    checks += [
        ('Set: % Lotes c/ TC', 0.9905381431105854, set_file['pct_lotes_tc']),
        ('Set: lotes', 1691.0, set_file['total_lotes']),
        ('Set: Lotes/dia', 73.61242109845462, set_file['lotes_dia']),
        ('Set: HC p/ turno Corrigido', 60.048167164179105, set_file['hc_turno_corr']),
        ('Set: HC total Corrigido', 180.1445014925373, set_file['hc_total_corr']),
        ('Set: % Ocupação Corrigido', 1.1778769549662436, set_file['ocupacao_corr']),
        ('Set: Lotes/dia a 100%', 62.495849662466874, set_file['lotes_dia_100']),
    ]
    cs = sem['complexidade']
    checks += [(f'Complex. semana nível {k}', v, cs[k]) for k, v in (('1', 18), ('2', 125), ('3', 166), ('4', 102))]
    cm = complexidade_meses(t, ['Set/2026'])
    tot14 = sum(cm[k] for k in '1234')
    checks += [(f'Complex. Set nível {k}', v, cm[k] / tot14) for k, v in
               (('1', 0.06453715775749674), ('2', 0.2907431551499348), ('3', 0.39895697522816165), ('4', 0.2457627118644068))]
    falhas = 0
    for nome, a, b in checks:
        ok = close(a, b)
        falhas += not ok
        print(('OK   ' if ok else 'FALHA'), f'{nome:32s} arquivo={a!r:<22} recalculado={b!r}')
    print('FALHAS:', falhas)

    # ── Todos os períodos ──
    periodos = [sem] + [periodo_meses(t, i0, i1) for i0 in range(12) for i1 in range(i0, 12)]
    for p in periodos:
        for r in p['cap']:
            r.pop('_CM_MIN', None)
    json.dump({'periodos': periodos, 'set_como_arquivo': set_file}, open(sys.argv[2], 'w'), ensure_ascii=False)
    print('períodos gerados:', len(periodos))
