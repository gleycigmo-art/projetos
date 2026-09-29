"""Converte a planilha "Status + Programação" em tabelas normalizadas.

Cada aba "Fam ..." é uma pilha de blocos, um por grupo de produto. Cada bloco tem:
  - linha de cabeçalho (Entrada CQ, Código, Produto, Lote PA, Lote PI, ..., testes)
  - linha "Padrão": preparo de padrão e fase móvel de cada teste do grupo
  - linhas de lote (SA/granel com Lote PI e PA/acabado com Lote PA)
Cada teste ocupa 3 colunas: Status | Prog. | Backup
  - Prog.  : 1 = programado
  - Backup : retorno de quem executou — 1 = feito, C = continuidade no próximo turno, N = não feito
  - "s" nas duas colunas = marcação antiga de teste já concluído

Saída: um CSV por tabela e `carga_powerapps.xlsx` (uma aba/tabela por entidade,
pronta para importar em Dataverse ou listas do SharePoint).

Uso:
    python -m programacao.extrair_programacao "Status + Programação V0.xlsx" dados_programacao
"""

import csv
import sys
from datetime import date, datetime
from pathlib import Path

import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo

IGNORAR_TESTES = {"OBSERVAÇÃO", ""}
STATUS_CONCLUIDO = {
    "APROVADO", "N/A", "FEITO NO PI - OK", "FEITO NA VALIDAÇÃO", "PEGAR DO COA",
}
RETORNO = {"1": "Feito", "C": "Continuidade", "N": "Não feito"}


def _txt(v):
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.date().isoformat() if v.time() == datetime.min.time() else v.isoformat(" ", "minutes")
    if isinstance(v, date):
        return v.isoformat()
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def _prog_backup(prog, backup):
    """Traduz o par Prog./Backup em (programado, retorno, códigos originais)."""
    p, b = _txt(prog).upper(), _txt(backup).upper()
    programado = "Sim" if p in ("1", "2-") else ""
    retorno = RETORNO.get(b, "")
    return programado, retorno, p, b


def _eh_cabecalho(ws, r):
    return _txt(ws.cell(r, 4).value) == "Entrada CQ" or _txt(ws.cell(r, 5).value) == "Código"


def extrair_familias(wb):
    grupos, catalogo, lotes, analises, preparos = [], [], [], [], []
    for ws in wb.worksheets:
        familia = ws.title.strip()
        if not familia.startswith("Fam"):
            continue
        testes, grupo = None, None
        for r in range(1, ws.max_row + 1):
            cel = lambda c: ws.cell(r, c).value  # noqa: E731
            if _eh_cabecalho(ws, r):
                testes = {}
                for c in range(13, ws.max_column + 1):
                    nome = _txt(ws.cell(r, c).value)
                    if nome.upper() not in IGNORAR_TESTES:
                        testes[c] = nome
                grupo = dict(grupo_id=f"{familia} #{len(grupos) + 1:03d}", familia=familia,
                             linha=r, produto_referencia="", codigos="")
                grupos.append(grupo)
                for ordem, nome in enumerate(testes.values(), start=1):
                    catalogo.append(dict(grupo_id=grupo["grupo_id"], familia=familia,
                                         ordem=ordem, teste=nome))
                continue
            if testes is None:
                continue
            produto = _txt(cel(6))
            if produto == "Padrão":
                for c, nome in testes.items():
                    status = _txt(cel(c)).upper()
                    programado, retorno, p, b = _prog_backup(cel(c + 1), cel(c + 2))
                    if status or programado:
                        preparos.append(dict(grupo_id=grupo["grupo_id"], familia=familia,
                                             teste=nome, status=status, programado=programado,
                                             retorno=retorno, prog_original=p, backup_original=b))
                continue
            if not produto or not (_txt(cel(7)) or _txt(cel(8))):
                continue
            lote_pa, lote_pi = _txt(cel(7)), _txt(cel(8))
            lote_id = lote_pa or lote_pi
            codigo = _txt(cel(5))
            if not grupo["produto_referencia"]:
                grupo["produto_referencia"] = produto
            if codigo and codigo not in grupo["codigos"].split(", "):
                grupo["codigos"] = ", ".join(filter(None, [grupo["codigos"], codigo]))
            lotes.append(dict(
                lote=lote_id, familia=familia, grupo_id=grupo["grupo_id"], linha=r,
                tipo="PA" if lote_pa else "SA", codigo=codigo, produto=produto,
                lote_pa=lote_pa, lote_pi=lote_pi, entrada_cq=_txt(cel(4)),
                dias_cq=_txt(cel(9)), data_lib_prevista=_txt(cel(10)),
                cluster_inicial=_txt(cel(11)), cluster_atual=_txt(cel(12)),
                analista_t1=_txt(cel(1)), analista_t2=_txt(cel(2)), analista_t3=_txt(cel(3)),
            ))
            for c, nome in testes.items():
                status = _txt(cel(c)).upper()
                programado, retorno, p, b = _prog_backup(cel(c + 1), cel(c + 2))
                if not status and not programado:
                    continue
                analises.append(dict(
                    analise_id=f"{lote_id}|{nome}", lote=lote_id, familia=familia,
                    grupo_id=grupo["grupo_id"], tipo="PA" if lote_pa else "SA",
                    lote_pi=lote_pi, teste=nome, status=status,
                    concluido="Sim" if status in STATUS_CONCLUIDO else "",
                    programado=programado, retorno=retorno,
                    prog_original=p, backup_original=b,
                    comentario=(ws.cell(r, c).comment.text.strip() if ws.cell(r, c).comment else ""),
                ))
    return grupos, catalogo, lotes, analises, preparos


def extrair_escala(wb):
    ws = wb["Lista Analista Fam"]
    escala = []
    for r in range(2, ws.max_row + 1):
        for turno, (ca, cf) in (("1", (1, 2)), ("2", (4, 5)), ("3", (7, 8))):
            nome = _txt(ws.cell(r, ca).value)
            if nome:
                escala.append(dict(turno=turno, analista=nome,
                                   alocacao=_txt(ws.cell(r, cf).value)))
    return escala


def extrair_equipamentos(wb):
    ws = wb["Equipamentos"]
    cab = [_txt(ws.cell(1, c).value) for c in range(1, ws.max_column + 1)]
    out = []
    for r in range(2, ws.max_row + 1):
        linha = {cab[c - 1]: _txt(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)}
        if linha.get("Equipamentos"):
            out.append(linha)
    return out


def extrair_dissolucao(wb):
    """Programação da dissolução: listas por turno a partir da linha 'Código'."""
    ws = wb["Dissolução"]
    blocos = {"1": 2, "2": 9, "3": 16}  # coluna de 'Código' de cada turno
    cab = next(r for r in range(1, ws.max_row + 1) if _txt(ws.cell(r, 2).value) == "Código")
    out = []
    for turno, c0 in blocos.items():
        for ordem, r in enumerate(range(cab + 1, ws.max_row + 1), start=1):
            lote = _txt(ws.cell(r, c0 + 2).value)
            if not lote:
                continue
            out.append(dict(turno=turno, ordem=ordem, codigo=_txt(ws.cell(r, c0).value),
                            produto=_txt(ws.cell(r, c0 + 1).value), lote=lote,
                            cluster=_txt(ws.cell(r, c0 + 3).value),
                            obs=_txt(ws.cell(r, c0 + 4).value)))
    return out


def _campos(linhas):
    return list(dict.fromkeys(k for l in linhas for k in l))


def _salvar_csv(pasta, nome, linhas):
    with open(pasta / f"{nome}.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=_campos(linhas), delimiter=";")
        w.writeheader()
        w.writerows(linhas)


def _salvar_xlsx(caminho, tabelas):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for nome, linhas in tabelas.items():
        ws = wb.create_sheet(nome)
        campos = _campos(linhas)
        ws.append(campos)
        for l in linhas:
            ws.append([l.get(k, "") for k in campos])
        ref = f"A1:{openpyxl.utils.get_column_letter(len(campos))}{len(linhas) + 1}"
        tab = Table(displayName=f"tb{nome.title().replace('_', '')}", ref=ref)
        tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(tab)
        for i, k in enumerate(campos, start=1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = max(12, min(40, len(k) + 4))
    wb.save(caminho)


def main(arquivo, pasta_saida):
    wb = openpyxl.load_workbook(arquivo, data_only=True)
    pasta = Path(pasta_saida)
    pasta.mkdir(parents=True, exist_ok=True)
    grupos, catalogo, lotes, analises, preparos = extrair_familias(wb)
    tabelas = {
        "grupos": grupos,
        "catalogo_testes": catalogo,
        "lotes": lotes,
        "analises": analises,
        "preparo_padrao": preparos,
        "escala": extrair_escala(wb),
        "equipamentos": extrair_equipamentos(wb),
        "programacao_dissolucao": extrair_dissolucao(wb),
    }
    for nome, linhas in tabelas.items():
        _salvar_csv(pasta, nome, linhas)
        print(f"{nome:24s} {len(linhas):5d} linhas")
    _salvar_xlsx(pasta / "carga_powerapps.xlsx", tabelas)

    pend = [a for a in analises if not a["concluido"]]
    prog = [a for a in analises if a["programado"]]
    ret = [a for a in prog if a["retorno"]]
    print(f"\n{len(lotes)} lotes, {len(analises)} análises, {len(pend)} pendentes")
    print(f"programadas: {len(prog)} | com retorno: {len(ret)} "
          f"(feito {sum(a['retorno'] == 'Feito' for a in ret)}, "
          f"continuidade {sum(a['retorno'] == 'Continuidade' for a in ret)}, "
          f"não feito {sum(a['retorno'] == 'Não feito' for a in ret)})")
    print(f"preparos de padrão/FM programados: {sum(1 for p in preparos if p['programado'])}")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    main(sys.argv[1], sys.argv[2])
