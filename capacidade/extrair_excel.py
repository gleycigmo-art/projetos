"""Extrai as tabelas da planilha MFV (.xlsb) para CSVs em dados/.

Uso: python -m capacidade.extrair_excel "MFV_CQPA_2026_Oficial - 25-09-26.xlsb"

Os intervalos de cada tabela do Excel são lidos direto do arquivo, então linhas
novas nas tabelas entram automaticamente.
"""
import os
import re
import struct
import sys
import zipfile

import pandas as pd
from pyxlsb import open_workbook

# Tabelas usadas pelo motor de cálculo → nome do CSV gerado
TABELAS = {
    "Tabela_FILTRO": "filtro",
    "Tabela_DEMANDA": "demanda_mensal",
    "Tabela_DemandaSemanal": "demanda_semanal",
    "Tabela_FAM": "familias",
    "CALENDARIO_MES": "calendario_mes",
    "CALENDARIO_SEMANA": "calendario_semana",
    "Tabela_Lotes_Campanha": "etapas_campanha",
    "Tabela_BASE_TC": "base_tc",
    "Tabela_BD_CAMPANHA": "campanha",
    "Tabela_SKIP_IMP": "skip",
    "CAP_PROPOSTA": "cap_proposta_excel",
}


def _registros(b):
    """Itera os registros binários (tipo, dados) de uma parte .bin do XLSB."""
    i = 0
    while i < len(b):
        t = b[i] & 0x7F
        i += 1
        if b[i - 1] & 0x80:
            t |= (b[i] & 0x7F) << 7
            i += 1
        tam, desl = 0, 0
        for _ in range(4):
            x = b[i]
            i += 1
            tam |= (x & 0x7F) << desl
            desl += 7
            if not x & 0x80:
                break
        yield t, b[i:i + tam]
        i += tam


def _intervalos_tabelas(z):
    """{nome_tabela: (nº da aba, linha1, linha2, col1, col2)} lido das partes xl/tables."""
    aba_da_tabela = {}
    for nome in z.namelist():
        m = re.match(r"xl/worksheets/_rels/sheet(\d+)\.bin\.rels", nome)
        if m:
            for tb in re.findall(r"tables/(table\d+)\.bin", z.read(nome).decode()):
                aba_da_tabela[tb] = int(m.group(1))
    res = {}
    for nome in z.namelist():
        m = re.match(r"xl/tables/(table\d+)\.bin", nome)
        if not m:
            continue
        for tipo, d in _registros(z.read(nome)):
            if tipo == 343:  # BrtBeginList
                r1, r2, c1, c2 = struct.unpack("<4I", d[:16])
                n = struct.unpack("<I", d[64:68])[0]  # stName (após 12 campos de 4 bytes)
                nome_tab = d[68:68 + 2 * n].decode("utf-16").strip()
                res[nome_tab] = (aba_da_tabela[m.group(1)], r1, r2, c1, c2)
                break
    return res


def ler_tabelas(caminho):
    with zipfile.ZipFile(caminho) as z:
        intervalos = _intervalos_tabelas(z)
    saida = {}
    with open_workbook(caminho) as wb:
        for nome_tab, (aba, r1, r2, c1, c2) in intervalos.items():
            if nome_tab not in TABELAS:
                continue
            grade = {}
            with wb.get_sheet(aba) as sh:
                for row in sh.rows():
                    for c in row:
                        if r1 <= c.r <= r2 and c1 <= c.c <= c2:
                            grade[(c.r, c.c)] = c.v
            cab = [str(grade.get((r1, c), f"col{c}")).strip() for c in range(c1, c2 + 1)]
            linhas = [[grade.get((r, c)) for c in range(c1, c2 + 1)] for r in range(r1 + 1, r2 + 1)]
            saida[TABELAS[nome_tab]] = pd.DataFrame(linhas, columns=cab)
    return saida


def ler_blocos_familias(caminho):
    """Blocos soltos da aba FAMÍLIAS (fora das tabelas): ausências e peso dos dias úteis."""
    with open_workbook(caminho) as wb:
        with wb.get_sheet("FAMÍLIAS") as sh:
            celulas = {(c.r, c.c): c.v for row in sh.rows() for c in row if c.v not in (None, "")}
    pos = {v.strip(): k for k, v in celulas.items() if isinstance(v, str)}
    blocos = {}

    if "P/D" in pos:
        r0, c0 = pos["P/D"]
        cab_sem = str(celulas.get((r0 - 1, c0 + 1), "Semana"))
        sem = re.search(r"(\d+)", cab_sem)
        linhas, r = [], r0
        while isinstance(celulas.get((r, c0)), str):
            tipo = celulas[(r, c0)].strip().rstrip(":")
            linhas.append((tipo, celulas.get((r, c0 + 1)), celulas.get((r, c0 + 2))))
            r += 1
        df = pd.DataFrame(linhas, columns=["TIPO", "QTDE_SEMANA", "QTDE_MEDIA_ANO"])
        df["SEMANA_REF"] = int(sem.group(1)) if sem else None
        blocos["ausencias"] = df

    if "Dia útil" in pos:
        r0, c0 = pos["Dia útil"]
        linhas = []
        for r in range(r0 + 1, r0 + 8):
            dia = celulas.get((r, c0 - 1))
            if isinstance(dia, str):
                linhas.append((dia.strip(), celulas.get((r, c0))))
        blocos["dia_util"] = pd.DataFrame(linhas, columns=["DIA", "PESO"])
    return blocos


def main():
    caminho = sys.argv[1]
    destino = sys.argv[2] if len(sys.argv) > 2 else "dados"
    os.makedirs(destino, exist_ok=True)
    tabelas = ler_tabelas(caminho)
    tabelas.update(ler_blocos_familias(caminho))
    for nome, df in tabelas.items():
        df.to_csv(os.path.join(destino, f"{nome}.csv"), index=False, encoding="utf-8")
        print(f"{nome:22s} {df.shape[0]:5d} linhas x {df.shape[1]:3d} colunas")


if __name__ == "__main__":
    main()
