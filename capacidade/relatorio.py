"""Gera o relatório de capacidade em Excel a partir dos CSVs de dados/.

Uso: python -m capacidade.relatorio [pasta_dados] [arquivo_saida.xlsx]
"""
import os
import sys

import pandas as pd

from .motor import (COL_EQUIP, COL_MPR, ET12_N, ET13_N, MESES, Dados, Regras, cap_proposta,
                    capacidade_maxima, demanda_mpr, mes_a_mes, pior_periodo, to_num)

CHAVE = ["PROCESSO", "FAMILIA"]


def alertas_dados(dados: Dados) -> pd.DataFrame:
    """Inconsistências de cadastro que fazem carga sumir do cálculo."""
    tc = dados.base_tc[dados.base_tc[COL_MPR].notna()]
    dem = demanda_mpr(dados, dados.periodo_do_filtro())
    com_dem = set(dem.index[dem["Demanda_Total"] > 0])
    alertas = []

    corre = tc[(tc[ET12_N].map(to_num) + tc[ET13_N].map(to_num) > 0)
               & tc[COL_EQUIP].map(lambda e: not (isinstance(e, str) and e.startswith(("CRLQ", "CRGS"))))]
    for mpr, g in corre.groupby(COL_MPR):
        if mpr in com_dem:
            alertas.append(("Tempo de HPLC/CG sem Equipamento Proposto (carga de máquina ignorada)",
                            mpr, f"{len(g)} teste(s)"))

    sem_tc = sorted(m for m in com_dem if m is not None and m not in set(tc[COL_MPR]))
    for mpr in sem_tc:
        texto = ("Códigos de produto com demanda e sem MPR (carga ignorada)"
                 if mpr == "Método não encontrado"
                 else "MPR com demanda e sem tempos na BASE TC (carga ignorada)")
        alertas.append((texto, mpr, f"{dem.loc[mpr, 'Demanda_Total']:.1f} lotes no período"))

    fam = dados.familias
    for _, r in fam[fam["CLUSTER"].isna()].iterrows():
        alertas.append(("Recurso sem CLUSTER na aba FAMÍLIAS", r["FAMILIA"], r["PROCESSO"]))
    return pd.DataFrame(alertas, columns=["Alerta", "Item", "Detalhe"])


def gerar(pasta="dados", saida="resultados/relatorio_capacidade.xlsx"):
    dados = Dados.carregar(pasta)
    periodo = dados.periodo_do_filtro()

    corr = cap_proposta(dados, periodo, Regras())
    leg = cap_proposta(dados, periodo, Regras.legado())
    comp = leg[CHAVE + ["CM (HRS)", "CARREGAMENTO", "HC NECESSARIO", "FIFO"]].merge(
        corr[CHAVE + ["CM (HRS)", "CARREGAMENTO", "HC NECESSARIO", "FIFO"]],
        on=CHAVE, suffixes=(" legado", " corrigido"))
    comp["Δ CM (HRS)"] = comp["CM (HRS) corrigido"] - comp["CM (HRS) legado"]

    mm = mes_a_mes(dados, periodo.meses if periodo.nivel == "Mensal" else MESES)
    pior = pior_periodo(mm)
    mapa = mm.pivot_table(index=CHAVE, columns="PERIODO", values="CARREGAMENTO")
    mapa = mapa[[m for m in MESES if m in mapa.columns]].reset_index()

    maxmo = capacidade_maxima(dados, periodo, categorias=("Mão de Obra",))
    alertas = alertas_dados(dados)

    resumo = pd.DataFrame([
        ("Período", periodo.rotulo),
        ("Regras", "corrigidas (campanha, FIFO por dias úteis, skip pela tabela)"),
        ("Carga total (h/mês)", round(corr["CM (HRS)"].sum(), 1)),
        ("Recursos sobrecarregados (> 1,0) na média", int((corr["CARREGAMENTO"] > 1).sum())),
        ("Recursos sobrecarregados (> 1,0) no pior mês", int((pior["CARREG. PIOR"] > 1).sum())),
        ("Mão de obra — gargalo hoje", maxmo["gargalo_hoje"]),
        ("Mão de obra — carregamento do gargalo", round(maxmo["carregamento_gargalo_hoje"], 2)),
        ("Mão de obra — demanda máxima atendível (× demanda atual)", round(maxmo["fator_maximo"], 2)),
        ("Alertas de cadastro", len(alertas)),
    ], columns=["Indicador", "Valor"])

    os.makedirs(os.path.dirname(saida) or ".", exist_ok=True)
    with pd.ExcelWriter(saida) as xw:
        resumo.to_excel(xw, sheet_name="Resumo", index=False)
        corr.round(3).to_excel(xw, sheet_name="CAP_PROPOSTA", index=False)
        pior.round(3).to_excel(xw, sheet_name="Pior mês", index=False)
        mapa.round(2).to_excel(xw, sheet_name="Carregamento mês a mês", index=False)
        mm.round(3).to_excel(xw, sheet_name="Mês a mês (detalhe)", index=False)
        comp.round(2).to_excel(xw, sheet_name="Legado x Corrigido", index=False)
        alertas.to_excel(xw, sheet_name="Alertas de dados", index=False)
        for ws in xw.book.worksheets:
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = min(
                    45, max(10, max(len(str(c.value or "")) for c in col[:60]) + 2))
            ws.freeze_panes = "A2"
    print(resumo.to_string(index=False))
    print(f"\nRelatório salvo em {saida}")


if __name__ == "__main__":
    gerar(*sys.argv[1:])
