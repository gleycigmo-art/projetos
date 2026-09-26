"""Build do FLOW LAB.

1. Converte dados/*.csv no conjunto compacto usado pelo app (app/dist/dados.json).
2. Gera o gabarito do motor Python para o teste de paridade (app/test/gabarito.json).
3. Monta o arquivo único app/dist/flow-lab.html (CSS, JS e dados embutidos).

Uso: python app/build.py [pasta_dados]
"""
import datetime as dt
import json
import math
import os
import sys

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, RAIZ)

from capacidade import motor  # noqa: E402

APP = os.path.join(RAIZ, "app")
TABELAS = ["base_tc", "demanda_mensal", "demanda_semanal", "familias", "calendario_mes",
           "calendario_semana", "campanha", "skip", "etapas_campanha", "ausencias", "dia_util"]

# Colunas mantidas por tabela (as demais não entram no app)
COLS_BASE_TC = [motor.COL_MPR, motor.COL_FAM, motor.COL_EQUIP, *motor.ESPECTRO, motor.ET1, motor.ET2,
                motor.ET3, motor.ET6, motor.ET7, motor.ET8, motor.ET9, motor.ET10, motor.ET12_N,
                motor.ET12_T, motor.ET13_N, motor.ET13_T, motor.ET14, motor.ET15, motor.ET16]
COLS = {
    "base_tc": COLS_BASE_TC,
    "demanda_mensal": ["CONSIDERAR?", "CÓD", "DESCRIÇÃO", "MÉTODO DOCNIX", *motor.MESES, "complexidade"],
    "demanda_semanal": ["Cod", "Chave semana", "STATUS", "Lote semana", "MÉTODO DOCNIX", "Status2", "Complexidade"],
}


def _limpo(v):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return None
    if isinstance(v, float) and v.is_integer() and abs(v) < 1e15:
        return int(v)
    if isinstance(v, float):
        return round(v, 10)
    return v


def dataset(pasta):
    out = {}
    for nome in TABELAS:
        df = motor._ler(pasta, nome)
        cols = [c for c in COLS.get(nome, list(df.columns))]
        faltando = [c for c in cols if c not in df.columns]
        if faltando:
            raise SystemExit(f"{nome}: colunas ausentes {faltando}")
        out[nome] = {"cols": cols, "rows": [[_limpo(v) for v in linha] for linha in df[cols].itertuples(index=False)]}
    return {"meta": {"fonte": "MFV_CQPA_2026_Oficial - 25-09-26.xlsb",
                     "gerado_em": dt.date.today().isoformat()}, "tabelas": out}


def gabarito(pasta):
    d = motor.Dados.carregar(pasta)
    casos = []
    periodos = [("Ano", motor.Periodo("Mensal", list(motor.MESES)))]
    periodos += [(m, motor.Periodo("Mensal", [m])) for m in motor.MESES]
    periodos += [("S38", motor.Periodo("Semanal", semana=202638))]
    for nome_reg, reg in (("corrigidas", motor.Regras()), ("legado", motor.Regras.legado())):
        for rot, per in periodos:
            r = motor.cap_proposta(d, per, reg)
            casos.append({"regras": nome_reg, "periodo": rot, "linhas": [
                {"PROCESSO": x["PROCESSO"], "FAMILIA": x["FAMILIA"], "CM_HRS": x["CM (HRS)"],
                 "HC_NECESSARIO": x["HC NECESSARIO"], "CARREGAMENTO": x["CARREGAMENTO"],
                 "DEMANDA": x["Soma Demanda Media"], "LOTES_DIA": x["FIFO"]}
                for _, x in r.iterrows()]})
    return casos


def html(dados_json):
    src = os.path.join(APP, "src")
    ler = lambda n: open(os.path.join(src, n), encoding="utf-8").read()
    pagina = ler("index.html")
    pagina = pagina.replace("/*__CSS__*/", ler("styles.css"))
    pagina = pagina.replace("/*__ENGINE__*/", ler("engine.js"))
    pagina = pagina.replace("/*__IMPORT__*/", ler("importar.js"))
    pagina = pagina.replace("/*__APP__*/", ler("app.js"))
    pagina = pagina.replace("/*__DADOS__*/", dados_json.replace("</", "<\\/"))
    return pagina


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    pasta = args[0] if args else os.path.join(RAIZ, "dados")
    os.makedirs(os.path.join(APP, "dist"), exist_ok=True)
    ds = json.dumps(dataset(pasta), ensure_ascii=False, separators=(",", ":"))
    with open(os.path.join(APP, "dist", "dados.json"), "w", encoding="utf-8") as f:
        f.write(ds)
    if "--sem-gabarito" not in sys.argv:
        with open(os.path.join(APP, "test", "gabarito.json"), "w", encoding="utf-8") as f:
            json.dump(gabarito(pasta), f, ensure_ascii=False)
    if os.path.exists(os.path.join(APP, "src", "index.html")):
        with open(os.path.join(APP, "dist", "flow-lab.html"), "w", encoding="utf-8") as f:
            f.write(html(ds))
    print(f"dados.json: {len(ds) / 1024:.0f} KB")


if __name__ == "__main__":
    main()
