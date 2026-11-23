"""Motor de cálculo de capacidade do laboratório de CQ (réplica das queries Power Query).

Fluxo (mesmos nomes das queries do Excel):
    DEMANDA_MPR  → lotes por MPR no período
    TC_COM_CAMPANHA → carga (min) por família/processo, com regras de campanha e skip
    CAP_PROPOSTA → carga × capacidade (pessoas, calendário, OEE)

O objeto `Regras` liga/desliga as correções combinadas com o negócio; com
`Regras.legado()` o resultado é o mesmo da planilha atual.
"""
from __future__ import annotations

import math
import os
from dataclasses import dataclass, field

import pandas as pd

MESES = ["Jan/2026", "Fev/2026", "Mar/2026", "Abr/2026", "Mai/2026", "Jun/2026",
         "Jul/2026", "Ago/2026", "Set/2026", "Out/2026", "Nov/2026", "Dez/2026"]

COL_MPR = "Método e versão (DocNix)"
COL_FAM = "FAM proposta"
COL_EQUIP = "Equipamento Proposto"

ET1 = "1 - PREPARO (PADRÃO E SOLUÇÕES) - BANCADA"
ET2 = "2 - PREPARO AMOSTRA - BANCADA"
ET3 = "3 - ANÁLISE FQ - BANCADA"
ET6 = "6 - FINALIZAÇÃO DOS DADOS - BANCADA"
ET7 = "7 - CONFERÊNCIA DO PREPARO - BANCADA"
ET8 = "8 - TEMPO DO DISSOLUTOR - BANCADA"
ET9 = "9 - MONTAGEM DO EQUIPAMENTO - LEITURA"
ET10 = "10 - CONFERÊNCIA EQUIPAMENTO - LEITURA"
ET12_N = "12 - Nº INJEÇÕES PADRÃO - LEITURA"
ET12_T = "12 - TEMPO PADRÃO (CORRIDA) - LEITURA"
ET13_N = "13 - Nº INJEÇÕES AMOSTRA/ DILUENTE/ PLACEBO - LEITURA"
ET13_T = "13 - TEMPO AMOSTRA/ DILUENTE/ PLACEBO (CORRIDA) - LEITURA"
ET14 = "14 - PROCESSAMENTO DA CORRIDA - LEITURA"
ET15 = "15 - CONFERÊNCIA PROCESSAMENTO DA CORRIDA"
ET16 = "16 - REVISÃO LAUDO - REVISÃO"
ESPECTRO = ["SPCT0071 - AA", "ICP", "SPCT - IV", "SPCT - UV"]

FAM_SKIP = "Impureza Skip Test"
FAM_PITF = "Fam PI/TF"
FIFO_DIAS_LEGADO = 22.97


# ─────────────────────────────────────────────────────────────────────────────
# Regras e período
# ─────────────────────────────────────────────────────────────────────────────
@dataclass
class Regras:
    # 1. Campanha: 1º lote de cada campanha com tempo cheio, demais com o % da etapa,
    #    inclusive quando todos os lotes cabem em uma campanha só.
    campanha_corrigida: bool = True
    # 2. FIFO usa os dias úteis do período (calendário) em vez da constante 22,97.
    fifo_dias_uteis_calendario: bool = True
    # 3. Skip = teste de impureza de MPR que consta na Tabela_SKIP_IMP (mesma regra para
    #    MO e equipamento; no legado a MO olhava só a família).
    skip_pela_tabela: bool = True

    @classmethod
    def legado(cls) -> "Regras":
        return cls(False, False, False)


@dataclass
class Periodo:
    nivel: str = "Mensal"              # "Mensal" ou "Semanal"
    meses: list[str] = field(default_factory=lambda: list(MESES))
    semana: int | None = None          # chave AAAASS, ex. 202638

    @property
    def rotulo(self) -> str:
        if self.nivel == "Semanal":
            return f"Sem {self.semana % 100}/{self.semana // 100}"
        return self.meses[0] if len(self.meses) == 1 else f"{self.meses[0]} a {self.meses[-1]}"


@dataclass
class Cenario:
    """Ajustes de simulação sobre os dados base."""
    fator_demanda: float = 1.0                                        # multiplica toda a demanda
    fator_demanda_familia: dict[str, float] = field(default_factory=dict)  # por FAM proposta
    pessoas: dict[str, float] = field(default_factory=dict)          # FAMILIA → recurso médio/turno
    oee: dict[str, float] = field(default_factory=dict)              # FAMILIA → OEE


# ─────────────────────────────────────────────────────────────────────────────
# Dados
# ─────────────────────────────────────────────────────────────────────────────
def _num_ou_texto(v):
    if isinstance(v, str):
        try:
            return float(v)
        except ValueError:
            return v
    if isinstance(v, float) and math.isnan(v):
        return None
    return v


def _ler(pasta, nome):
    df = pd.read_csv(os.path.join(pasta, f"{nome}.csv"), dtype=object, keep_default_na=False,
                     na_values=[""])
    return df.map(_num_ou_texto).astype(object).where(df.notna(), None)


@dataclass
class Dados:
    demanda_mensal: pd.DataFrame
    demanda_semanal: pd.DataFrame
    base_tc: pd.DataFrame
    campanha: pd.DataFrame
    skip: pd.DataFrame
    etapas_campanha: pd.DataFrame
    familias: pd.DataFrame
    calendario_mes: pd.DataFrame
    calendario_semana: pd.DataFrame
    filtro: pd.DataFrame

    @classmethod
    def carregar(cls, pasta="dados") -> "Dados":
        return cls(*(_ler(pasta, n) for n in (
            "demanda_mensal", "demanda_semanal", "base_tc", "campanha", "skip",
            "etapas_campanha", "familias", "calendario_mes", "calendario_semana", "filtro")))

    def periodo_do_filtro(self) -> Periodo:
        p = dict(zip(self.filtro["Parametro"], self.filtro["Valor"]))
        if p.get("Nível de Visão") == "Semanal":
            return Periodo("Semanal", semana=int(p["Semana (Chave)"]))
        i, f = MESES.index(p["Mês Início"]), MESES.index(p["Mês Fim"])
        return Periodo("Mensal", MESES[i:f + 1])


def to_num(v) -> float:
    """Equivalente ao ToNum das queries: nulo/vazio/texto não numérico → 0."""
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return 0.0 if isinstance(v, float) and math.isnan(v) else float(v)
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _vazio(v) -> bool:
    """Nas queries: v = null or v = 0 or v = "" (texto como "S" conta como preenchido)."""
    return v is None or v == "" or (isinstance(v, (int, float)) and v == 0)


def roundup(x: float) -> float:
    return float(math.ceil(x - 1e-12))


# ─────────────────────────────────────────────────────────────────────────────
# DEMANDA_MPR
# ─────────────────────────────────────────────────────────────────────────────
def demanda_mpr(dados: Dados, periodo: Periodo, cenario: Cenario | None = None) -> pd.DataFrame:
    """Lotes por MPR: Demanda_Total (no período), Num_Meses e Demanda_MPR (média por período)."""
    if periodo.nivel == "Semanal":
        d = dados.demanda_semanal
        d = d[d["Chave semana"].map(to_num) == periodo.semana]
        linhas = pd.DataFrame({"MPR": d["MÉTODO DOCNIX"], "COD": d["Cod"],
                               "Qtde": d["Lote semana"].map(to_num)})
        n = 1
    else:
        d = dados.demanda_mensal
        d = d[d["CONSIDERAR?"] == "SIM"]
        partes = []
        for m in periodo.meses:
            v = d[m]
            ok = v.notna()                      # Unpivot descarta nulos
            partes.append(pd.DataFrame({"MPR": d.loc[ok, "MÉTODO DOCNIX"], "COD": d.loc[ok, "CÓD"],
                                        "Qtde": v[ok].map(to_num)}))
        linhas = pd.concat(partes)
        n = len(periodo.meses)

    if cenario is not None and (cenario.fator_demanda != 1 or cenario.fator_demanda_familia):
        fam_mpr = (dados.base_tc.dropna(subset=[COL_MPR]).drop_duplicates(COL_MPR)
                   .set_index(COL_MPR)[COL_FAM])
        fator = linhas["MPR"].map(lambda m: cenario.fator_demanda
                                  * cenario.fator_demanda_familia.get(fam_mpr.get(m), 1.0))
        linhas = linhas.assign(Qtde=linhas["Qtde"] * fator)

    g = linhas.groupby("MPR", dropna=False)["Qtde"].sum().rename("Demanda_Total").to_frame()
    g["Num_Meses"] = n
    g["Demanda_MPR"] = g["Demanda_Total"] / n
    return g


# ─────────────────────────────────────────────────────────────────────────────
# TC_COM_CAMPANHA
# ─────────────────────────────────────────────────────────────────────────────
class _Ctx:
    def __init__(self, dados: Dados, regras: Regras, dem: pd.DataFrame):
        self.regras = regras
        self.dem = dem
        camp = dados.campanha.drop_duplicates("CÓD")
        self.lotes_camp = dict(zip(camp["CÓD"], camp["QTD LOTES CAMPANHA"]))
        sk = dados.skip.drop_duplicates("MPR")
        self.skip = {m: q for m, q in zip(sk["MPR"], sk["Qtde Lotes Skip"])
                     if q is not None and to_num(q) >= 1}
        pct = dict(zip(dados.etapas_campanha["Etapa"], dados.etapas_campanha["% tempo"]))
        self.pct = {e: pct.get(e + "_CAMPANHA") for e in (ET1, ET6, ET7, ET9, ET10, ET12_T)}

    def lotes(self, mpr):
        if mpr in self.dem.index:
            r = self.dem.loc[mpr]
            return to_num(r["Demanda_Total"]), to_num(r["Num_Meses"])
        return 0.0, 0.0

    def n_skip(self, mpr):
        q = self.skip.get(mpr)
        return to_num(q) if q is not None else 1.0

    def eh_skip(self, mpr, fam, bloco_equip=False):
        # O skip vale só para o teste de impureza (linha da família "Impureza Skip Test");
        # a Tabela_SKIP_IMP confirma se o MPR tem skip aprovado e define o N.
        if self.regras.skip_pela_tabela or bloco_equip:
            return fam == FAM_SKIP and mpr in self.skip
        return fam == FAM_SKIP

    def n_camp(self, mpr, lotes):
        lxc = roundup(to_num(self.lotes_camp.get(mpr)))
        return lxc, (roundup(lotes / lxc) if lxc > 0 else 0.0)

    def carga(self, mpr, fam, etapas: list[tuple[float, float | None]], com_campanha=True):
        """Carga (min por período) de um conjunto de etapas [(tc, pct_campanha|None)]."""
        lotes, meses = self.lotes(mpr)
        if meses <= 0:
            return 0.0
        tc_total = sum(tc for tc, _ in etapas)
        if self.eh_skip(mpr, fam):
            ns = self.n_skip(mpr)
            return tc_total * (roundup(lotes / ns) if ns > 0 else 0) / meses
        lxc, nc = self.n_camp(mpr, lotes) if com_campanha else (0, 0)
        if nc == 0:
            return tc_total * lotes / meses
        if not self.regras.campanha_corrigida:
            if nc == 1:
                return tc_total / meses              # legado: TC cobrado 1x para todos os lotes
            total = 0.0
            for tc, pct in etapas:
                if pct is not None:
                    total += nc * tc + (lotes - nc) * tc * pct
                else:
                    total += tc * lotes
            return total / meses
        # Corrigido: 1º lote de cada campanha cheio, demais com % da etapa
        primeiros = min(nc, lotes)
        total = 0.0
        for tc, pct in etapas:
            if pct is None:
                total += tc * lotes
            else:
                total += primeiros * tc + (lotes - primeiros) * tc * pct
        return total / meses


def _pct_bloco(pct, somente_positivo):
    """Blocos D/E/H usam o % só se > 0; blocos A/B usam se existir na tabela."""
    if pct is None:
        return None
    return pct if (pct > 0 or not somente_positivo) else None


def tc_com_campanha(dados: Dados, regras: Regras, dem: pd.DataFrame) -> pd.DataFrame:
    """Carga por linha MPR × teste, agrupada como na query: FAM proposta, PROCESSO, CM_Min."""
    cx = _Ctx(dados, regras, dem)
    tc = dados.base_tc[dados.base_tc[COL_MPR].notna()]
    cols = set(tc.columns)
    g = lambda r, c: to_num(r[c]) if c in cols else 0.0
    p = cx.pct
    saida = []  # (mpr, fam, processo, cm)

    for _, r in tc.iterrows():
        mpr, fam = r[COL_MPR], r[COL_FAM]

        # A — MO BANCADA (1+2+3+6), agrupado por família
        saida.append((mpr, fam, "MO BANCADA", cx.carga(mpr, fam, [
            (g(r, ET1), p[ET1]), (g(r, ET2), None), (g(r, ET3), None), (g(r, ET6), p[ET6])])))

        # B — Conferência do preparo (7)
        if not _vazio(r.get(ET7)) and g(r, ET7) != 0:
            proc = "MO CONFERÊNCIA PI/TF" if fam == FAM_PITF else "MO CONFERÊNCIA BANCADA"
            saida.append((mpr, proc, proc, cx.carga(mpr, fam, [(g(r, ET7), p[ET7])])))

        # C — MO EQUIPAMENTO BANCADA (espectro), sem campanha
        esp = sum(g(r, c) for c in ESPECTRO)
        if esp != 0:
            saida.append((mpr, "MO EQUIPAMENTO BANCADA", "MO EQUIPAMENTO BANCADA",
                          cx.carga(mpr, fam, [(esp, None)], com_campanha=False)))

        # D — MO EQUIPAMENTO (9 com campanha, 14) — cada etapa é uma linha
        for et, pct in ((ET9, _pct_bloco(p[ET9], True)), (ET14, None)):
            v = r.get(et)
            if v is not None and to_num(v) != 0 and str(v) != "S":
                saida.append((mpr, "MO EQUIPAMENTO", "MO EQUIPAMENTO",
                              cx.carga(mpr, fam, [(to_num(v), pct)])))

        # E — MO CONFERÊNCIA EQUIPAMENTO (10 com campanha, 15)
        for et, pct in ((ET10, _pct_bloco(p[ET10], True)), (ET15, None)):
            v = r.get(et)
            if v is not None and to_num(v) != 0 and str(v) != "S":
                saida.append((mpr, "MO CONFERÊNCIA EQUIPAMENTO", "MO CONFERÊNCIA EQUIPAMENTO",
                              cx.carga(mpr, fam, [(to_num(v), pct)])))

        # F — REVISÃO FINAL (16), G — DISSOLUTOR (8): sem campanha
        for et, proc in ((ET16, "REVISÃO FINAL"), (ET8, "DISSOLUTOR")):
            if g(r, et) != 0:
                saida.append((mpr, proc, proc, cx.carga(mpr, fam, [(g(r, et), None)],
                                                        com_campanha=False)))

    # H — EQUIPAMENTO HPLC / CG (tempo de máquina), agrupado por equipamento
    tco = dados.base_tc
    eq = tco[tco[COL_EQUIP].map(lambda e: isinstance(e, str) and e.startswith(("CRLQ", "CRGS")))]
    for _, r in eq.iterrows():
        mpr, fam, equip = r[COL_MPR], r[COL_FAM], r[COL_EQUIP]
        tc12 = g(r, ET12_N) * g(r, ET12_T)
        tc13 = g(r, ET13_N) * g(r, ET13_T)
        lotes, meses = cx.lotes(mpr)
        if meses <= 0:
            cm = 0.0
        elif cx.eh_skip(mpr, fam, bloco_equip=True):
            ns = cx.n_skip(mpr)
            cm = (tc12 + tc13) * (roundup(lotes / ns) if ns > 0 else 0) / meses
        else:
            pct12 = to_num(p[ET12_T])
            _, nc = cx.n_camp(mpr, lotes)
            if nc == 0:
                cm = (tc12 + tc13) * lotes / meses
            elif regras.campanha_corrigida:
                prim = min(nc, lotes)
                cm = (prim * tc12 + (lotes - prim) * tc12 * pct12 + tc13 * lotes) / meses
            elif nc == 1:
                cm = (tc12 + tc13) / meses
            else:
                cm = (nc * tc12 + (lotes - nc) * tc12 * pct12 + tc13 * lotes) / meses
        proc = "EQUIPAMENTO HPLC" if equip.startswith("CRLQ") else "EQUIPAMENTO CG"
        saida.append((mpr, equip, proc, cm))

    return pd.DataFrame(saida, columns=["MPR", "FAM proposta", "PROCESSO", "CM_Min"])


# ─────────────────────────────────────────────────────────────────────────────
# Capacidade (aba FAMÍLIAS + calendário)
# ─────────────────────────────────────────────────────────────────────────────
def _linhas_calendario(dados: Dados, periodo: Periodo) -> pd.DataFrame:
    if periodo.nivel == "Semanal":
        c = dados.calendario_semana
        return c[c["Semana Chave"].map(to_num) == periodo.semana]
    c = dados.calendario_mes
    nums = [MESES.index(m) + 1 for m in periodo.meses]
    return c[c["Mês Nº"].map(to_num).isin(nums)]


def dias_uteis_periodo(dados: Dados, periodo: Periodo) -> float:
    """Dias úteis médios por período (mês ou semana ISO) — denominador do FIFO."""
    return float(_linhas_calendario(dados, periodo)["Dias Úteis"].map(to_num).mean())


def capacidade(dados: Dados, periodo: Periodo, cenario: Cenario | None = None) -> pd.DataFrame:
    """Recalcula RECURSO MÉDIO, HRS TURNO e TOTAL RECURSO para o período."""
    cal = _linhas_calendario(dados, periodo)
    f = dados.familias.copy()
    turnos_cols = ["QTDE 1° TURNO", "QTDE 2° TURNO", "QTDE 3° TURNO"]

    def recurso(r):
        n = int(to_num(r["TURNOS"])) or 3
        return sum(to_num(r[c]) for c in turnos_cols[:n]) / n

    def horas(r):
        n = int(to_num(r["TURNOS"])) or 3
        pref = "EQP" if r["CATEGORIA"] == "Equipamento" else "MO"
        return float(cal[f"{pref} - {n}"].map(to_num).mean())

    f["RECURSO MÉDIO ATUAL"] = f.apply(recurso, axis=1)
    f["HRS TURNO"] = f.apply(horas, axis=1)
    f["OEE REAL"] = f["OEE REAL"].map(to_num)
    if cenario:
        for fam, v in cenario.pessoas.items():
            f.loc[f["FAMILIA"] == fam, "RECURSO MÉDIO ATUAL"] = v
        for fam, v in cenario.oee.items():
            f.loc[f["FAMILIA"] == fam, "OEE REAL"] = v
    f["TOTAL RECURSO"] = f["RECURSO MÉDIO ATUAL"] * f["HRS TURNO"]
    return f


# ─────────────────────────────────────────────────────────────────────────────
# CAP_PROPOSTA
# ─────────────────────────────────────────────────────────────────────────────
def _demanda_por_recurso(dados: Dados, dem: pd.DataFrame) -> tuple[dict, dict, dict]:
    tc = dados.base_tc[dados.base_tc[COL_MPR].notna()]
    dmpr = dem["Demanda_MPR"].to_dict()

    def soma(mprs):
        vals = [dmpr[m] for m in set(mprs) if m in dmpr and dmpr[m] is not None]
        return sum(vals) if vals else None

    por_fam = {}
    for fam, grp in tc.groupby(COL_FAM):
        if fam != "Terceirizado":
            por_fam[fam] = soma(grp[COL_MPR])

    def preenchido(cols, filtro=None):
        cols = [c for c in cols if c in tc.columns]
        m = tc[cols].apply(lambda s: s.map(lambda v: not _vazio(v))).any(axis=1)
        if filtro is not None:
            m &= filtro
        return soma(tc.loc[m, COL_MPR])

    pitf = tc[COL_FAM] == FAM_PITF
    por_proc = {
        "MO CONFERÊNCIA BANCADA": preenchido([ET7], ~pitf),
        "MO CONFERÊNCIA PI/TF": preenchido([ET7], pitf),
        "MO CONFERÊNCIA EQUIPAMENTO": preenchido([ET10, ET15]),
        "MO EQUIPAMENTO BANCADA": preenchido(ESPECTRO),
        "REVISÃO FINAL": preenchido([ET16]),
        "DISSOLUTOR": preenchido([ET8]),
        "MO EQUIPAMENTO": preenchido([ET9, ET14]),
    }
    eq = dados.base_tc[dados.base_tc[COL_EQUIP].map(
        lambda e: isinstance(e, str) and e.startswith(("CRLQ", "CRGS")))]
    por_equip = {e: (soma(g[COL_MPR]) or 0) for e, g in eq.groupby(COL_EQUIP)}
    return por_fam, por_proc, por_equip


def cap_proposta(dados: Dados, periodo: Periodo | None = None, regras: Regras | None = None,
                 cenario: Cenario | None = None, carga: pd.DataFrame | None = None) -> pd.DataFrame:
    periodo = periodo or dados.periodo_do_filtro()
    regras = regras or Regras()
    dem = demanda_mpr(dados, periodo, cenario)
    carga = carga if carga is not None else tc_com_campanha(dados, regras, dem)
    cap = capacidade(dados, periodo, cenario)

    ag = carga.groupby(["FAM proposta", "PROCESSO"], dropna=False)["CM_Min"].sum().reset_index()
    cap_idx = {(r["FAMILIA"], r["PROCESSO"]): r for _, r in cap.iterrows()}
    cap_fam = {}
    for _, r in cap.iterrows():
        cap_fam.setdefault(r["FAMILIA"], r)

    por_fam, por_proc, por_equip = _demanda_por_recurso(dados, dem)
    dias = (dias_uteis_periodo(dados, periodo) if regras.fifo_dias_uteis_calendario
            else FIFO_DIAS_LEGADO)

    linhas = []
    for _, a in ag.iterrows():
        fam, proc, cm = a["FAM proposta"], a["PROCESSO"], a["CM_Min"]
        c = cap_idx.get((fam, proc))
        if c is None or not to_num(c["TOTAL RECURSO"]):
            c = cap_fam.get(fam)
        if c is None or not to_num(c["TOTAL RECURSO"]):
            continue
        cm_h = cm / 60
        total, oee, hrs = c["TOTAL RECURSO"], c["OEE REAL"], c["HRS TURNO"]
        dem_rec = next((v for v in (por_fam.get(fam), por_proc.get(proc), por_equip.get(fam))
                        if v is not None), 0)
        oee_nec = cm_h / total
        linhas.append({
            "PROCESSO": proc, "FAMILIA": fam, "CLUSTER": c["CLUSTER"], "CATEGORIA": c["CATEGORIA"],
            "CM (HRS)": cm_h, "RECURSO MÉDIO ATUAL": c["RECURSO MÉDIO ATUAL"],
            "TOTAL RECURSO": total, "HRS TURNO": hrs, "OEE REAL": oee, "OEE NEC": oee_nec,
            "HC NECESSARIO": cm_h / (hrs * oee) if hrs and oee else None,
            "CARREGAMENTO": oee_nec / oee if oee else None,
            "Soma Demanda Media": dem_rec,
            "TAKT": total / dem_rec if dem_rec else None,
            "TC MEDIO": cm_h / dem_rec if dem_rec else None,
            "FIFO": dem_rec / dias if dem_rec else None,
            # Pergunta ao contrário: lotes/período que o recurso atende com carregamento 1,0
            "DEMANDA MAX": dem_rec / (oee_nec / oee) if dem_rec and oee_nec and oee else None,
        })
    out = pd.DataFrame(linhas)
    return out.sort_values(["CLUSTER", "PROCESSO", "FAMILIA"], na_position="last").reset_index(drop=True)


# ─────────────────────────────────────────────────────────────────────────────
# Visões derivadas
# ─────────────────────────────────────────────────────────────────────────────
def mes_a_mes(dados: Dados, meses: list[str] | None = None, regras: Regras | None = None,
              cenario: Cenario | None = None) -> pd.DataFrame:
    """CAP_PROPOSTA calculada para cada mês separadamente (capacidade e demanda do mês)."""
    meses = meses or MESES
    partes = []
    for m in meses:
        r = cap_proposta(dados, Periodo("Mensal", [m]), regras, cenario)
        r.insert(0, "PERIODO", m)
        partes.append(r)
    return pd.concat(partes, ignore_index=True)


def pior_periodo(por_periodo: pd.DataFrame) -> pd.DataFrame:
    """Para cada recurso: carregamento médio, pior período e saldo de horas nele."""
    df = por_periodo.assign(SALDO_HRS=por_periodo["TOTAL RECURSO"] * por_periodo["OEE REAL"]
                            - por_periodo["CM (HRS)"])
    chave = ["PROCESSO", "FAMILIA"]
    idx = df.groupby(chave)["CARREGAMENTO"].idxmax()
    pior = df.loc[idx, chave + ["CLUSTER", "PERIODO", "CARREGAMENTO", "SALDO_HRS", "HC NECESSARIO",
                                "RECURSO MÉDIO ATUAL"]]
    pior = pior.rename(columns={"PERIODO": "PIOR PERIODO", "CARREGAMENTO": "CARREG. PIOR",
                                "SALDO_HRS": "SALDO HRS PIOR", "HC NECESSARIO": "HC NEC. PIOR"})
    med = df.groupby(chave)["CARREGAMENTO"].mean().rename("CARREG. MÉDIO")
    return pior.join(med, on=chave).sort_values("CARREG. PIOR", ascending=False).reset_index(drop=True)


def capacidade_maxima(dados: Dados, periodo: Periodo | None = None, regras: Regras | None = None,
                      cenario: Cenario | None = None, limite: float = 1.0,
                      categorias: tuple[str, ...] = ("Mão de Obra",)) -> dict:
    """Pergunta ao contrário: até quanto a demanda pode crescer (fator uniforme) com os
    recursos atuais antes que o recurso gargalo das `categorias` passe de `limite`.

    A carga não é exatamente linear (arredondamento de campanhas e skip), então o fator
    é buscado por bisseção recalculando o motor.
    """
    periodo = periodo or dados.periodo_do_filtro()
    regras = regras or Regras()
    base = cenario or Cenario()

    def pior(f):
        c = Cenario(base.fator_demanda * f, base.fator_demanda_familia, base.pessoas, base.oee)
        r = cap_proposta(dados, periodo, regras, c)
        r = r[r["CATEGORIA"].isin(categorias)]
        i = r["CARREGAMENTO"].idxmax()
        return r.loc[i, "CARREGAMENTO"], r.loc[i, "FAMILIA"], r

    c1, gargalo1, r1 = pior(1.0)
    lo, hi = 0.0, 1.0
    while pior(hi)[0] <= limite and hi < 64:
        lo, hi = hi, hi * 2
    for _ in range(12):
        mid = (lo + hi) / 2
        if pior(mid)[0] <= limite:
            lo = mid
        else:
            hi = mid
    _, gargalo, r = pior(lo)
    return {"fator_maximo": lo, "gargalo": gargalo, "gargalo_hoje": gargalo1,
            "carregamento_gargalo_hoje": c1, "resultado_no_limite": r}
