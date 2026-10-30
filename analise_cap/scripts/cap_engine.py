"""Réplica fiel (linha a linha) das consultas Power Query do MFV_CQPA:
FILTRO_PERIODO -> DEMANDA_MPR -> BASE_TC_CAMP / FONTES_TC -> TC_COM_CAMPANHA -> CAP_PROPOSTA.
Também reproduz as fórmulas de planilha que alimentam o PQ (HRS TURNO do Tabela_FAM) e o rodapé da CAP.
"""
import math
from decimal import Decimal, ROUND_HALF_EVEN, ROUND_HALF_UP

MESES = ["Jan/2026", "Fev/2026", "Mar/2026", "Abr/2026", "Mai/2026", "Jun/2026",
         "Jul/2026", "Ago/2026", "Set/2026", "Out/2026", "Nov/2026", "Dez/2026"]

E1 = "1 - PREPARO (PADRÃO E SOLUÇÕES) - BANCADA"
E2 = "2 - PREPARO AMOSTRA - BANCADA"
E3 = "3 - ANÁLISE FQ - BANCADA"
E6 = "6 - FINALIZAÇÃO DOS DADOS - BANCADA"
E7 = "7 - CONFERÊNCIA DO PREPARO - BANCADA"
E8 = "8 - TEMPO DO DISSOLUTOR - BANCADA"
E9 = "9 - MONTAGEM DO EQUIPAMENTO - LEITURA"
E10 = "10 - CONFERÊNCIA EQUIPAMENTO - LEITURA"
E14 = "14 - PROCESSAMENTO DA CORRIDA - LEITURA"
E15 = "15 - CONFERÊNCIA PROCESSAMENTO DA CORRIDA"
E16 = "16 - REVISÃO LAUDO - REVISÃO"
NI12 = "12 - Nº INJEÇÕES PADRÃO - LEITURA"
TC12 = "12 - TEMPO PADRÃO (CORRIDA) - LEITURA"
NI13 = "13 - Nº INJEÇÕES AMOSTRA/ DILUENTE/ PLACEBO - LEITURA"
TC13 = "13 - TEMPO AMOSTRA/ DILUENTE/ PLACEBO (CORRIDA) - LEITURA"
ESPECTRO = ["SPCT0071 - AA", "ICP", "SPCT - IV", "SPCT - UV"]
MPR = "Método e versão (DocNix)"
FAM = "FAM proposta"


def records(t):
    h = t['header']
    return [dict(zip(h, r)) for r in t['rows']]


def is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def to_num(v):
    """ToNum do M: null/'' -> 0; número -> ele; texto conversível -> número; senão 0."""
    if v is None or v == "":
        return 0
    if is_num(v):
        return v
    try:
        return float(str(v).strip().replace(',', '.'))
    except ValueError:
        return 0


def roundup0(x):
    return math.ceil(x - 1e-12) if x > 0 else -math.ceil(-x - 1e-12) if x < 0 else 0


def m_round(x, d=2):
    """Number.Round do M (padrão RoundingMode.ToEven)."""
    if x is None:
        return None
    return float(Decimal(repr(x)).quantize(Decimal(1).scaleb(-d), rounding=ROUND_HALF_EVEN))


def m_int64(x):
    if x is None:
        return None
    return int(Decimal(repr(x)).quantize(Decimal(1), rounding=ROUND_HALF_EVEN))


def ehnull(v):
    return v is None


# ─────────────────────────────── PERÍODO ────────────────────────────────
class Periodo:
    """nivel: 'Mensal' (mes_ini..mes_fim) ou 'Semanal' (semana_chave)."""

    def __init__(self, nivel, mes_ini="Jan/2026", mes_fim="Dez/2026", semana_chave=None):
        self.nivel = nivel
        self.mes_ini, self.mes_fim, self.semana_chave = mes_ini, mes_fim, semana_chave
        i0, i1 = MESES.index(mes_ini), MESES.index(mes_fim)
        self.meses_ativos = MESES[i0:i1 + 1]


# ───────────────────── HRS TURNO (fórmula da planilha) ───────────────────
def hrs_turno(t, periodo, categoria, turnos):
    """Fórmula J de Tabela_FAM: média, nos períodos 'Ativos', das horas acumuladas p/ N turnos."""
    if periodo.nivel == "Mensal":
        cal = records(t['CALENDARIO_MES'])
        ativos = [r for r in cal if MESES[int(r['Mês Nº']) - 1] in periodo.meses_ativos]
        cols = ['MO - 1', 'MO - 2', 'MO - 3', 'ADM'] if categoria == "Mão de Obra" else ['EQP - 1', 'EQP - 2', 'EQP - 3', 'ADM2']
    else:
        cal = records(t['CALENDARIO_SEMANA'])
        ativos = [r for r in cal if r['Semana Chave'] == periodo.semana_chave]
        cols = ['MO - 1', 'MO - 2', 'MO - 3'] if categoria == "Mão de Obra" else ['EQP - 1', 'EQP - 2', 'EQP - 3']
    col = cols[int(turnos) - 1]
    return sum(to_num(r[col]) for r in ativos) / len(ativos)


def tabela_fam(t, periodo, hrs_override=None):
    """Recalcula Tabela_FAM (colunas I..M) para o período, como o Excel faria."""
    out = []
    for r in records(t['Tabela_FAM']):
        r = dict(r)
        rec = (to_num(r['QTDE 1° TURNO']) + to_num(r['QTDE 2° TURNO']) + to_num(r['QTDE 3° TURNO'])) / 3
        hrs = hrs_override if hrs_override is not None else hrs_turno(t, periodo, r['CATEGORIA'], r['TURNOS'])
        r['RECURSO MÉDIO ATUAL'] = rec
        r['HRS TURNO'] = hrs
        r['TOTAL RECURSO'] = rec * hrs
        r['RECURSO DISP (MIN)'] = rec * hrs * to_num(r['OEE REAL']) * 60
        out.append(r)
    return out


def fam_proposta(fam_rows):
    """Query FAM_PROPOSTA: tipa HRS TURNO e TURNOS como Int64 (arredonda!)."""
    out = []
    for r in fam_rows:
        r = dict(r)
        r['HRS TURNO'] = m_int64(r['HRS TURNO'])
        r['TURNOS'] = m_int64(r['TURNOS'])
        out.append(r)
    return out


# ─────────────────────────────── DEMANDA ────────────────────────────────
def demanda_mpr(t, periodo):
    """Query DEMANDA_MPR -> {MPR: (Demanda_Total, Demanda_MPR, Num_Meses)}"""
    res = {}
    if periodo.nivel == "Semanal":
        for r in records(t['Tabela_DemandaSemanal']):
            if r['Chave semana'] is None or r['Cod'] in (None, ""):
                continue
            if int(r['Chave semana']) != periodo.semana_chave:
                continue
            k = r['MÉTODO DOCNIX']
            v = r['Lote semana']
            res[k] = res.get(k, 0) + (m_int64(v) if v is not None else 0)
        return {k: (v, v, 1) for k, v in res.items()}
    n = len(periodo.meses_ativos)
    for r in records(t['Tabela_DEMANDA']):
        if r['CONSIDERAR?'] != "SIM":
            continue
        k = r['MÉTODO DOCNIX']
        for m in periodo.meses_ativos:
            v = r[m]
            if v is None:  # UnpivotOtherColumns descarta nulos
                continue
            res[k] = res.get(k, 0) + v
    return {k: (v, v / n, n) for k, v in res.items()}


# ─────────────────────────── BASE_TC_CAMP / FONTES ──────────────────────
def fontes(t):
    camp_rows = records(t['Tabela_BD_CAMPANHA'])
    skip_rows = records(t['Tabela_SKIP_IMP'])
    camp_dedup, skip_dedup = {}, {}
    for r in camp_rows:
        camp_dedup.setdefault(r['CÓD'], r['QTD LOTES CAMPANHA'])
    for r in skip_rows:
        skip_dedup.setdefault(r['MPR'], r['Qtde Lotes Skip'])
    lotes_camp = {r['Etapa']: r['% tempo'] for r in records(t['Tabela_Lotes_Campanha'])}
    return camp_rows, skip_rows, camp_dedup, skip_dedup, lotes_camp


def base_tc_camp(t, camp_rows, skip_rows):
    """Merge SEM dedup (como no M): duplica linhas se CÓD/MPR repetir."""
    camp_idx, skip_idx = {}, {}
    for r in camp_rows:
        camp_idx.setdefault(r['CÓD'], []).append(r['QTD LOTES CAMPANHA'])
    for r in skip_rows:
        skip_idx.setdefault(r['MPR'], []).append(r['Qtde Lotes Skip'])
    out = []
    for r in records(t['Tabela_BASE_TC']):
        for lc in camp_idx.get(r[MPR], [None]) if r[MPR] is not None else [None]:
            for ls in skip_idx.get(r[MPR], [None]) if r[MPR] is not None else [None]:
                x = dict(r)
                x['Lotes_Campanha'] = lc
                x['Lotes_Skip'] = ls
                x['N_Camp'] = 1 if (lc is None or lc < 1) else lc
                x['N_Skip'] = 1 if (ls is None or ls < 1) else ls
                x['EhSkip'] = r[FAM] == "Impureza Skip Test"
                out.append(x)
    return out


# ───────────────────────────── TC_COM_CAMPANHA ───────────────────────────
def calc_camp(tc, lotes, lxc, pct, nm):
    ncamp = roundup0(lotes / lxc)
    comp = ncamp * tc
    resto = (lotes - ncamp) * tc * pct
    tcm = (comp + resto) / lotes if lotes > 0 else 0
    return tcm * lotes / nm if nm > 0 else 0


def calc_skip(tc, lotes, nskip, nm):
    nsc = roundup0(lotes / nskip) if nskip > 0 else 0
    return tc * nsc / nm if nm > 0 else 0


def tc_com_campanha(t, periodo, dem=None, detalhe=False):
    dem = dem if dem is not None else demanda_mpr(t, periodo)
    camp_rows, skip_rows, camp_dedup, skip_dedup, lotes_camp = fontes(t)
    base = base_tc_camp(t, camp_rows, skip_rows)
    base = [r for r in base if r[MPR] is not None]

    def pct(et):
        return lotes_camp.get(et + "_CAMPANHA")

    P1, P6, P7, P9, P10, P12 = pct(E1), pct(E6), pct(E7), pct(E9), pct(E10), pct(TC12)

    def dvals(r):
        d = dem.get(r[MPR])
        tot, dm, nm = d if d else (None, None, None)
        return to_num(tot), dm, to_num(nm)

    def lxc_of(r):
        return roundup0(to_num(camp_dedup.get(r[MPR])))

    def ncamp_of(lotes, lxc):
        return roundup0(lotes / lxc) if lxc > 0 else 0

    linhas = []  # (FAM proposta, PROCESSO, MPR, etapa, CM_Min, Demanda_MPR)

    def generic(r, tc, etapa_pct, lotes, lxc, nm, nskip, ehskip):
        ncamp = ncamp_of(lotes, lxc)
        if ehskip:
            return calc_skip(tc, lotes, nskip, nm)
        if ncamp == 0:
            return tc * lotes / nm if nm > 0 else 0
        if ncamp == 1:
            return (tc / lotes) * lotes / nm if (nm > 0 and lotes > 0) else 0
        if etapa_pct is not None and etapa_pct is not False:
            return calc_camp(tc, lotes, lxc, to_num(etapa_pct), nm)
        return tc * lotes / nm if nm > 0 else 0

    # Bloco A — MO BANCADA (1+2+3+6)
    for r in base:
        lotes, dm, nm = dvals(r)
        lxc = lxc_of(r)
        nskip, ehskip = to_num(r['N_Skip']), r['EhSkip'] is True
        ncamp = ncamp_of(lotes, lxc)
        tc1, tc2, tc3, tc6 = (to_num(r.get(e)) for e in (E1, E2, E3, E6))
        soma = tc1 + tc2 + tc3 + tc6
        if ehskip:
            cm = calc_skip(soma, lotes, nskip, nm)
        elif ncamp == 0:
            cm = soma * lotes / nm if nm > 0 else 0
        elif ncamp == 1:
            cm = (soma / lotes) * lotes / nm if (nm > 0 and lotes > 0) else 0
        else:
            cm1 = calc_camp(tc1, lotes, lxc, to_num(P1), nm) if P1 is not None else (tc1 * lotes / nm if nm > 0 else 0)
            cm2 = tc2 * lotes / nm if nm > 0 else 0
            cm3 = tc3 * lotes / nm if nm > 0 else 0
            cm6 = calc_camp(tc6, lotes, lxc, to_num(P6), nm) if P6 is not None else (tc6 * lotes / nm if nm > 0 else 0)
            cm = cm1 + cm2 + cm3 + cm6
        linhas.append((r[FAM], "MO BANCADA", r[MPR], "1+2+3+6", cm, dm))

    # Bloco B — Conferência (7) bancada x PI/TF
    for r in base:
        tc = to_num(r.get(E7))
        if tc == 0:
            continue
        lotes, dm, nm = dvals(r)
        lxc = lxc_of(r)
        cm = generic(r, tc, P7, lotes, lxc, nm, to_num(r['N_Skip']), r['EhSkip'] is True)
        f = "MO CONFERÊNCIA PI/TF" if r[FAM] == "Fam PI/TF" else "MO CONFERÊNCIA BANCADA"
        linhas.append((f, f, r[MPR], "7", cm, dm))

    # Bloco C — Espectro (MO EQUIPAMENTO BANCADA)
    for r in base:
        tc = sum(to_num(r.get(c)) for c in ESPECTRO)
        if tc == 0:
            continue
        lotes, dm, nm = dvals(r)
        if r['EhSkip'] is True:
            cm = calc_skip(tc, lotes, to_num(r['N_Skip']), nm)
        else:
            cm = tc * lotes / nm if nm > 0 else 0
        linhas.append(("MO EQUIPAMENTO BANCADA", "MO EQUIPAMENTO BANCADA", r[MPR], "espectro", cm, dm))

    # Blocos D/E — MO EQUIPAMENTO (9+14) e MO CONFERÊNCIA EQUIPAMENTO (10+15), via unpivot
    for fam_out, etapas, et_pct, pct_val in (("MO EQUIPAMENTO", (E9, E14), E9, P9),
                                             ("MO CONFERÊNCIA EQUIPAMENTO", (E10, E15), E10, P10)):
        for r in base:
            for et in etapas:
                v = r.get(et)
                if v is None:  # unpivot descarta nulos
                    continue
                if to_num(v) == 0 or str(v) == "S":
                    continue
                tc = to_num(v)
                lotes, dm, nm = dvals(r)
                lxc = lxc_of(r)
                ncamp = ncamp_of(lotes, lxc)
                p = to_num(pct_val) if et == et_pct else 0
                if r['EhSkip'] is True:
                    cm = calc_skip(tc, lotes, to_num(r['N_Skip']), nm)
                elif ncamp == 0:
                    cm = tc * lotes / nm if nm > 0 else 0
                elif ncamp == 1:
                    cm = (tc / lotes) * lotes / nm if (nm > 0 and lotes > 0) else 0
                elif p > 0:
                    cm = calc_camp(tc, lotes, lxc, p, nm)
                else:
                    cm = tc * lotes / nm if nm > 0 else 0
                linhas.append((fam_out, fam_out, r[MPR], et.split(' ')[0], cm, dm))

    # Blocos F/G — REVISÃO FINAL (16) e DISSOLUTOR (8): sem campanha
    for fam_out, et in (("REVISÃO FINAL", E16), ("DISSOLUTOR", E8)):
        for r in base:
            tc = to_num(r.get(et))
            if tc == 0:
                continue
            lotes, dm, nm = dvals(r)
            if r['EhSkip'] is True:
                cm = calc_skip(tc, lotes, to_num(r['N_Skip']), nm)
            else:
                cm = tc * lotes / nm if nm > 0 else 0
            linhas.append((fam_out, fam_out, r[MPR], et.split(' ')[0], cm, dm))

    # Bloco H — Equipamentos HPLC/CG (TC_Orig, sem BASE_TC_CAMP)
    for r in records(t['Tabela_BASE_TC']):
        ep = r.get('Equipamento Proposto')
        if ep is None or ep == "" or not isinstance(ep, str) or not (ep.startswith("CRLQ") or ep.startswith("CRGS")):
            continue
        qs = skip_dedup.get(r[MPR]) if r[MPR] is not None else None
        nskip = 1 if (qs is None or qs < 1) else qs
        ehskip = r[FAM] == "Impureza Skip Test" and qs is not None and qs >= 1
        lxc = roundup0(to_num(camp_dedup.get(r[MPR]) if r[MPR] is not None else None))
        d = dem.get(r[MPR]) if r[MPR] is not None else None
        tot, dm, nm = d if d else (None, None, None)
        lotes, nm = to_num(tot), to_num(nm)
        tc12 = to_num(r.get(NI12)) * to_num(r.get(TC12))
        tc13 = to_num(r.get(NI13)) * to_num(r.get(TC13))
        soma = tc12 + tc13
        ncamp = ncamp_of(lotes, lxc)
        if ehskip:
            nsc = roundup0(lotes / nskip) if nskip > 0 else 0
            s = soma * nsc / lotes if lotes > 0 else 0
        elif ncamp == 0:
            s = soma
        elif ncamp == 1:
            s = soma / lotes if lotes > 0 else 0
        else:
            cc = ncamp * tc12
            cr = (lotes - ncamp) * tc12 * to_num(P12)
            s = ((cc + cr) / lotes if lotes > 0 else 0) + tc13
        cm = s * lotes / nm if nm > 0 else 0
        proc = "EQUIPAMENTO HPLC" if ep.startswith("CRLQ") else "EQUIPAMENTO CG"
        linhas.append((ep, proc, r[MPR], "12+13", cm, dm))

    return linhas


# ─────────────────────────────── CAP_PROPOSTA ───────────────────────────
def dias_uteis(t, periodo):
    """Dias úteis do período no calendário: semana escolhida ou média dos meses ativos."""
    if periodo.nivel == "Semanal":
        cal = [r for r in records(t['CALENDARIO_SEMANA']) if r['Semana Chave'] == periodo.semana_chave]
    else:
        cal = [r for r in records(t['CALENDARIO_MES']) if MESES[int(r['Mês Nº']) - 1] in periodo.meses_ativos]
    return sum(to_num(r['Dias Úteis']) for r in cal) / len(cal)


def cap_proposta(t, periodo, fam_rows=None, dem=None, linhas=None, dias_fifo=22.97):
    """dias_fifo: 22.97 reproduz o arquivo original; a versão nova usa dias_uteis(t, periodo)."""
    dem = dem if dem is not None else demanda_mpr(t, periodo)
    linhas = linhas if linhas is not None else tc_com_campanha(t, periodo, dem)
    fam_rows = fam_rows if fam_rows is not None else tabela_fam(t, periodo)
    famp = fam_proposta(fam_rows)

    # Agrupa por (FAM proposta, PROCESSO)
    grp = {}
    for f, p, _m, _e, cm, _dm in linhas:
        grp[(f, p)] = grp.get((f, p), 0) + cm

    rows = []
    for (f, p), cm in grp.items():
        matches = [x for x in famp if x['FAMILIA'] == f and x['PROCESSO'] == p]
        if not matches:
            matches_cand = [None]
        else:
            matches_cand = matches
        for mt in matches_cand:
            fam_final = mt['FAMILIA'] if (mt and mt['FAMILIA'] is not None) else f
            if mt is not None and mt['TOTAL RECURSO'] is not None:
                rows.append(dict(FAMP=f, PROCESSO=p, CM=cm, FAMILIA=fam_final, CLUSTER=mt['CLUSTER'],
                                 REC=mt['RECURSO MÉDIO ATUAL'], HRS=mt['HRS TURNO'], TOT=mt['TOTAL RECURSO'],
                                 OEE=mt['OEE REAL'], CATEGORIA=mt['CATEGORIA']))
            else:
                m2 = [x for x in famp if x['FAMILIA'] == fam_final] or [None]
                for x in m2:
                    rows.append(dict(FAMP=f, PROCESSO=p, CM=cm, FAMILIA=fam_final,
                                     CLUSTER=x['CLUSTER'] if x else None, REC=x['RECURSO MÉDIO ATUAL'] if x else None,
                                     HRS=x['HRS TURNO'] if x else None, TOT=x['TOTAL RECURSO'] if x else None,
                                     OEE=x['OEE REAL'] if x else None, CATEGORIA=x['CATEGORIA'] if x else None))
    rows = [r for r in rows if r['TOT'] is not None and r['TOT'] != 0]

    for r in rows:
        r['CM_HRS'] = r['CM'] / 60
        r['OEE_NEC'] = r['CM_HRS'] / r['TOT'] if r['TOT'] else None
        r['CARREG'] = r['OEE_NEC'] / r['OEE'] if (r['OEE'] not in (None, 0) and r['OEE_NEC'] is not None) else None
        r['HC_NEC'] = r['CM_HRS'] / (r['HRS'] * r['OEE']) if (r['HRS'] and r['OEE']) else None

    # ── Demanda por família ──
    base = base_tc_camp(t, *fontes(t)[:2])
    base_dem = [(r[MPR], r[FAM], dem.get(r[MPR], (None, None, None))[1], r) for r in base]
    seen, dem_fam = set(), {}
    for mpr, fam, dm, _ in base_dem:
        if (mpr, fam) in seen:
            continue
        seen.add((mpr, fam))
        if dm is not None:
            dem_fam[fam] = (dem_fam.get(fam) or 0) + dm
        else:
            dem_fam.setdefault(fam, None)
    dem_fam = {k: v for k, v in dem_fam.items() if k is not None and k != "Terceirizado"}

    def nz(v):
        return v is not None and v != 0 and v != ""

    def soma_dem_mpr(cols, filtro=None):
        vistos, tot = set(), None
        for mpr, fam, dm, r in base_dem:
            if not any(nz(r.get(c)) for c in cols):
                continue
            if filtro == "<>PI" and fam == "Fam PI/TF":
                continue
            if filtro == "=PI" and fam != "Fam PI/TF":
                continue
            if mpr in vistos:
                continue
            vistos.add(mpr)
            if dm is not None:
                tot = (tot or 0) + dm
        return tot if tot is not None else None

    dem_conf = {
        "MO CONFERÊNCIA BANCADA": soma_dem_mpr([E7], "<>PI"),
        "MO CONFERÊNCIA PI/TF": soma_dem_mpr([E7], "=PI"),
        "MO CONFERÊNCIA EQUIPAMENTO": soma_dem_mpr([E10, E15]),
        "MO EQUIPAMENTO BANCADA": soma_dem_mpr(ESPECTRO),
        "REVISÃO FINAL": soma_dem_mpr([E16]),
        "DISSOLUTOR": soma_dem_mpr([E8]),
        "MO EQUIPAMENTO": soma_dem_mpr([E9, E14]),
    }
    tc_orig = records(t['Tabela_BASE_TC'])
    dem_equip = {}
    for ep in {r.get('Equipamento Proposto') for r in tc_orig}:
        if not isinstance(ep, str) or not (ep.startswith("CRLQ") or ep.startswith("CRGS")):
            continue
        mprs = {r[MPR] for r in tc_orig if r.get('Equipamento Proposto') == ep}
        s = sum(dem[m][1] for m in mprs if m in dem and dem[m][1] is not None)
        dem_equip[ep] = s

    for r in rows:
        v = dem_fam.get(r['FAMP'])
        if v is None:
            v = dem_conf.get(r['PROCESSO'])
        if v is None:
            v = dem_equip.get(r['FAMP'])
        r['DEM'] = v if v is not None else 0
        r['TAKT'] = r['TOT'] / r['DEM'] if r['DEM'] else None
        r['TC_MEDIO'] = r['CM_HRS'] / r['DEM'] if r['DEM'] else None
        r['FIFO'] = r['TOT'] / r['TAKT'] / dias_fifo if r['TAKT'] else None

    out = []
    for r in rows:
        out.append({
            'PROCESSO': r['PROCESSO'], 'FAMILIA': r['FAMILIA'], 'CLUSTER': r['CLUSTER'],
            'Categoria': r['CATEGORIA'],
            'CM (HRS)': m_round(r['CM_HRS']), 'RECURSO MÉDIO ATUAL': m_round(r['REC']),
            'TOTAL RECURSO': m_round(r['TOT']), 'HRS TURNO': r['HRS'], 'OEE REAL': r['OEE'],
            'OEE NEC': r['OEE_NEC'], 'HC NECESSARIO': m_round(r['HC_NEC']), 'CARREGAMENTO': r['CARREG'],
            'Soma Demanda Media': m_round(r['DEM']), 'TAKT': m_round(r['TAKT']),
            'TC MEDIO': m_round(r['TC_MEDIO']), 'FIFO': m_round(r['FIFO']),
            '_CM_MIN': r['CM'],
        })
    out.sort(key=lambda x: ((x['CLUSTER'] or ''), (x['PROCESSO'] or '')))
    return out
