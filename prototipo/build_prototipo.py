"""Gera o protótipo navegável do app de programação com os dados reais da planilha.

    python -m prototipo.build_prototipo "Status + Programação V0.xlsx"

Saída: prototipo/programacao-cq.html (arquivo único, abre direto no navegador).
"""

import json
import re
import sys
from pathlib import Path

import openpyxl

from programacao.extrair_programacao import (
    extrair_equipamentos, extrair_escala, extrair_familias, STATUS_CONCLUIDO,
)

AQUI = Path(__file__).parent

# Nome da aba -> (id, nome curto, palavras para casar escala/equipamentos)
FAMILIAS = [
    ("Fam 1 - Torsilax", "tor", "Torsilax", ["fam 1", "torsilax"]),
    ("Fam  - Flavonid", "fla", "Flavonid", ["flavonid"]),
    ("Fam  - Dipirona", "dip", "Dipirona", ["dipirona"]),
    ("Fam 2 - Doralgina", "dor", "Doralgina", ["fam 2", "doralgina"]),
    ("Fam 3 - Neolefrin", "neo", "Neolefrin", ["fam 3", "neolefrin"]),
    ("Fam 4 - Losartana", "los", "Losartana", ["fam 4", "losartana"]),
    ("Fam 5 - Dermo - Semissólidos", "der", "Dermo / Semissólidos", ["fam 5", "dermo"]),
    ("Fam 6 - Liquidos", "liq", "Líquidos", ["fam 6", "líquidos", "liquidos"]),
    ("Fam 7 - Estéreis - Aerossóis", "est", "Estéreis / Aerossóis", ["fam 7", "estéreis", "estereis"]),
    ("Fam FQ", "fq", "FQ", ["fam fq", "fq"]),
    ("Fam 8 - Imp", "imp", "Impurezas", ["fam 8", "imp"]),
]
ID_POR_ABA = {aba: fid for aba, fid, _, _ in FAMILIAS}
AUSENCIAS = {"férias", "atestado", "treinamento"}


def _familia_por_texto(texto):
    t = re.sub(r"\s+", " ", texto.lower()).strip()
    for _, fid, _, chaves in FAMILIAS:
        if any(t == k or t.startswith(k + " ") or (len(k) > 4 and k in t) for k in chaves):
            return fid
    return ""


def montar_dados(arquivo):
    wb = openpyxl.load_workbook(arquivo, data_only=True)
    grupos, catalogo, lotes, analises, preparos = extrair_familias(wb)

    testes_por_grupo = {}
    for c in catalogo:
        testes_por_grupo.setdefault(c["grupo_id"], []).append(c["teste"])

    out_lotes, idx_lote = [], {}
    for l in lotes:
        chave = (l["grupo_id"], l["lote"])
        if chave in idx_lote:
            continue
        idx_lote[chave] = len(out_lotes)
        out_lotes.append(dict(
            id=len(out_lotes), g=l["grupo_id"], f=ID_POR_ABA[l["familia"]], lote=l["lote"],
            tipo=l["tipo"], cod=l["codigo"], prod=l["produto"], pi=l["lote_pi"],
            ent=l["entrada_cq"], dias=l["dias_cq"], lib=l["data_lib_prevista"],
            cl=l["cluster_atual"] or l["cluster_inicial"], a=[],
        ))

    itens = []
    for a in analises:
        lid = idx_lote.get((a["grupo_id"], a["lote"]))
        if lid is None:
            continue
        aid = f"a{lid}-{len(out_lotes[lid]['a'])}"
        out_lotes[lid]["a"].append([a["teste"], a["status"] or "AG. ANÁLISE"])
        if a["programado"]:
            itens.append(dict(tipo="A", ref=aid, turno=2 if a["prog_original"] == "2-" else 1,
                              ret=a["retorno"], origem="planilha"))

    out_grupos = {}
    for g in grupos:
        out_grupos[g["grupo_id"]] = dict(f=ID_POR_ABA[g["familia"]], nome=g["produto_referencia"],
                                         t=testes_por_grupo.get(g["grupo_id"], []), p={})
    for p in preparos:
        grp = out_grupos[p["grupo_id"]]
        grp["p"][p["teste"]] = p["status"] or "AG. ANÁLISE"
        if p["programado"]:
            itens.append(dict(tipo="P", ref=f"p|{p['grupo_id']}|{p['teste']}",
                              turno=2 if p["prog_original"] == "2-" else 1,
                              ret=p["retorno"], origem="planilha"))
    # só grupos que têm lote
    usados = {l["g"] for l in out_lotes}
    out_grupos = {k: v for k, v in out_grupos.items() if k in usados}

    analistas = []
    for e in extrair_escala(wb):
        aloc = e["alocacao"]
        analistas.append(dict(nome=re.sub(r"^\d\s*-\s*", "", e["analista"]).strip(),
                              turno=int(e["turno"]), aloc=aloc.strip(),
                              f=_familia_por_texto(aloc),
                              ausente=aloc.strip().lower() in AUSENCIAS))

    equipamentos = []
    for e in extrair_equipamentos(wb):
        equipamentos.append(dict(
            tag=e.get("Equipamentos", ""), desc=e.get("Caracteristicas", ""),
            fam_txt=e.get("Familia", ""), f=_familia_por_texto(e.get("Familia", "").replace("Sólidos ", "")),
            atual=e.get("Produto", ""), termino=e.get("Horário Término", ""),
            prox=e.get("Próximo Produto/Teste", ""), cal=e.get("Data de Calibração", ""),
            resp=[e.get("RESPONSÁVEL 1º TURNO", ""), e.get("RESPONSÁVEL 2º TURNO", ""),
                  e.get("RESPONSÁVEL 3º TURNO", "")],
        ))

    return dict(
        gerado="2026-09-29",
        concluidos=sorted(STATUS_CONCLUIDO),
        familias=[dict(id=fid, nome=nome) for _, fid, nome, _ in FAMILIAS],
        grupos=out_grupos, lotes=out_lotes, itens=itens,
        analistas=analistas, equipamentos=equipamentos,
    )


def main(arquivo):
    dados = montar_dados(arquivo)
    modelo = (AQUI / "app.html").read_text(encoding="utf-8")
    js = json.dumps(dados, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    html = modelo.replace("/*__DADOS__*/null", js)
    destino = AQUI / "programacao-cq.html"
    destino.write_text(html, encoding="utf-8")
    print(f"{destino}  ({len(html) / 1024:.0f} KB, {len(dados['lotes'])} lotes, "
          f"{len(dados['itens'])} itens programados, {len(dados['analistas'])} analistas)")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "Status + Programação V0.xlsx")
