// Teste de paridade: o motor JS precisa reproduzir o motor Python (capacidade/motor.py).
// Rode: python app/build.py && node --test app/test/
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const E = require("../src/engine.js");
const dados = JSON.parse(readFileSync(new URL("../dist/dados.json", import.meta.url)));
const gabarito = JSON.parse(readFileSync(new URL("./gabarito.json", import.meta.url)));
const ctx = E.preparar(dados);

function periodo(rot) {
  if (rot === "Ano") return { nivel: "Mensal", meses: E.MESES };
  if (rot === "S38") return { nivel: "Semanal", semana: 38, chave: 202638, fonte: "real" };
  return { nivel: "Mensal", mes: rot };
}
const perto = (a, b, tol = 1e-6) => (a === null || b === null ? a === b : Math.abs(a - b) <= tol * Math.max(1, Math.abs(b)));

for (const caso of gabarito) {
  test(`paridade ${caso.regras} · ${caso.periodo}`, () => {
    const regras = caso.regras === "legado" ? E.REGRAS_LEGADO : E.REGRAS_CORRIGIDAS;
    const r = E.capProposta(ctx, periodo(caso.periodo), { regras });
    const idx = new Map(r.linhas.map((l) => [`${l.PROCESSO}|${l.FAMILIA}`, l]));
    assert.equal(r.linhas.length, caso.linhas.length, "quantidade de linhas");
    for (const g of caso.linhas) {
      const l = idx.get(`${g.PROCESSO}|${g.FAMILIA}`);
      assert.ok(l, `linha ausente ${g.PROCESSO} ${g.FAMILIA}`);
      for (const k of ["CM_HRS", "HC_NECESSARIO", "CARREGAMENTO", "DEMANDA", "LOTES_DIA"])
        assert.ok(perto(l[k], g[k]), `${g.FAMILIA} ${k}: js=${l[k]} py=${g[k]}`);
    }
  });
}

test("projeção semanal distribui o plano do ano inteiro", () => {
  const { semanas } = E.listarPeriodos(ctx);
  let total = 0;
  for (const s of semanas) total += E.linhasDemanda(ctx, { ...s, fonte: "projecao" }).linhas.reduce((a, l) => a + l.qtde, 0);
  const ano = E.linhasDemanda(ctx, { nivel: "Mensal", meses: E.MESES }).linhas.reduce((a, l) => a + l.qtde, 0);
  assert.ok(Math.abs(total - ano) / ano < 1e-9, `semanas=${total} ano=${ano}`);
});

test("semana 38 confere com o exemplo do FLOW LAB (HC real e ausências)", () => {
  const r = E.calcular(ctx, { nivel: "Semanal", semana: 38, chave: 202638, fonte: "real" });
  assert.ok(Math.abs(r.kpi.hcRealTotal - 149.35) < 0.01);
  assert.ok(Math.abs(r.ausencias.total - 12.667) < 0.01);
});

test("cenário não altera os dados originais e OEE menor aumenta a ocupação", () => {
  const p = { nivel: "Mensal", mes: "Set/2026" };
  const antes = E.calcular(ctx, p);
  const sim = E.calcular(ctx, p, { cenario: { oeeMO: 0.6 } });
  const depois = E.calcular(ctx, p);
  assert.equal(antes.kpi.ocupGlobal, depois.kpi.ocupGlobal);
  assert.ok(Math.abs(sim.kpi.ocupGlobal / antes.kpi.ocupGlobal - 0.8 / 0.6) < 1e-9);
});

test("remanejamento move HC entre famílias sem mudar o total", () => {
  const p = { nivel: "Mensal", mes: "Set/2026" };
  const base = E.calcular(ctx, p);
  const sim = E.calcular(ctx, p, { cenario: { remanejamentos: [{ de: "Fam PI/TF", para: "Sólidos - Neolefrin", turno: "todos", hc: 1 }] } });
  assert.ok(Math.abs(sim.kpi.hcRealTotal - base.kpi.hcRealTotal) < 1e-9);
  const neo = (r) => r.mo.find((l) => l.FAMILIA === "Sólidos - Neolefrin").CARREGAMENTO;
  assert.ok(neo(sim) < neo(base));
});

test("capacidade máxima leva o gargalo a carregamento 1,0", () => {
  const p = { nivel: "Mensal", mes: "Set/2026" };
  const cm = E.capacidadeMaxima(ctx, p);
  const r = E.calcular(ctx, p, { cenario: { fatorDemanda: cm.fator } });
  const max = Math.max(...r.mo.map((l) => l.CARREGAMENTO));
  assert.ok(Math.abs(max - 1) < 1e-3, `max=${max}`);
});
