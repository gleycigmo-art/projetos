"""Lê as Tabelas do .xlsb com os valores CALCULADOS PELO EXCEL (cache), via pyxlsb."""
import sys, pickle
from pyxlsb import open_workbook
from openpyxl.utils import range_boundaries

REFS = {
    'Tabela_FILTRO': ('FILTRO_DATA', 'B92:C97'),
    'MAPA_SEMANAS': ('FILTRO_DATA', 'B36:I89'),
    'Tabela_DemandaSemanal': ('Tabela_DemandaSemanal', 'A1:H197'),
    'Tabela_DEMANDA': ('DEMANDA', 'A2:AC1455'),
    'DEMANDA_RODAPE': ('DEMANDA', 'G1457:S1465'),
    'FAMILIAS_EQUIPAMENTO': ('FAMILIAS_EQUIPAMENTO', 'A1:F38'),
    'CALENDARIO_MES': ('FAMÍLIAS', 'O2:AD14'),
    'CALENDARIO_SEMANA': ('FAMÍLIAS', 'AF2:AR55'),
    'Tabela_FAM': ('FAMÍLIAS', 'A2:M58'),
    'SOMA_TEMPOS_TOTAIS': ('SOMA TEMPOS TOTAIS', 'A1:I733'),
    'FAIXAS_COMPLEXIDADE': ('SOMA TEMPOS TOTAIS', 'L2:O5'),
    'CAP_PROPOSTA': ('CAP_PROPOSTA', 'A4:P59'),
    'CAP_RODAPE': ('CAP_PROPOSTA', 'B61:O76'),
    'Tabela_BASE_TC': ('BASE TC', 'A4:DM4723'),
    'Tabela_Lotes_Campanha': ('ETAPAS_CAMPANHAS', 'A1:B7'),
    'Tabela_BD_CAMPANHA': ('CAMPANHA', 'A1:B295'),
    'Tabela_SKIP_IMP': ('SKIP_IMPUREZA', 'A1:C23'),
}

def load(path):
    out = {}
    by_sheet = {}
    for name, (sh, ref) in REFS.items():
        by_sheet.setdefault(sh, []).append((name, ref))
    with open_workbook(path) as wb:
        for sh, items in by_sheet.items():
            grid = {}
            with wb.get_sheet(sh) as ws:
                for row in ws.rows(sparse=True):
                    for c in row:
                        if c.v is not None:
                            grid[(c.r, c.c)] = c.v
            for name, ref in items:
                c0, r0, c1, r1 = range_boundaries(ref)
                rows = [[grid.get((r - 1, c - 1)) for c in range(c0, c1 + 1)] for r in range(r0, r1 + 1)]
                out[name] = {'header': rows[0], 'rows': rows[1:], 'ref': ref, 'sheet': sh}
    return out

if __name__ == '__main__':
    t = load(sys.argv[1])
    pickle.dump(t, open(sys.argv[2], 'wb'))
    for k, v in t.items():
        print(k, len(v['rows']), v['header'][:8])
