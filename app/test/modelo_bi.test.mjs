// Valida o desenho do modelo Power BI (powerbi/*.m): a mesma lógica escrita como
// tabelas longas (componentes × demanda por período) precisa dar o mesmo resultado
// que o motor do app. Cada passo abaixo corresponde a uma query M.
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
const require = createRequire(import.meta.url);
const E = require("../src/engine.js");
const ctx = E.preparar(JSON.parse(readFileSync(new URL("../dist/dados.json", import.meta.url))));
const { C, toNum } = E;
const t = ctx.t;
const vazio = (v) => v === null || v === undefined || v === "" || v === 0;
const up = (x) => Math.ceil(x - 1e-12);

// ── pCampanha / pSkip ─────────────────────────────────────────────────────
const camp = new Map(); for (const r of t.campanha) if (!camp.has(r["CÓD"])) camp.set(r["CÓD"], up(toNum(r["QTD LOTES CAMPANHA"])));
const skip = new Map(); for (const r of t.skip) if (!skip.has(r.MPR) && r["Qtde Lotes Skip"] !== null && toNum(r["Qtde Lotes Skip"]) >= 1) skip.set(r.MPR, toNum(r["Qtde Lotes Skip"]));
const pctT = new Map(t.etapas_campanha.map((r) => [r.Etapa.replace(/_CAMPANHA$/, ""), r["% tempo"] === null ? null : toNum(r["% tempo"])]));
const pct = (e) => (pctT.has(e) ? pctT.get(e) : null);
const pos = (p) => (p !== null && p > 0 ? p : null);

// ── fComponentes: MPR, FAM_LINHA, DESTINO, PROCESSO, TC, PCT, CAMPANHA ───
const comp = [];
const add = (r, dest, proc, tc, p, campanha) => comp.push({ mpr: r[C.MPR], fam: r[C.FAM], dest, proc, tc, pct: p, campanha });
const ok = (v) => v !== null && v !== undefined && toNum(v) !== 0 && String(v) !== "S";
for (const r of t.base_tc) {
  if (r[C.MPR] === null || r[C.MPR] === "") continue;
  const f = r[C.FAM];
  add(r, f, "MO BANCADA", toNum(r[C.ET1]), pct(C.ET1), true);
  add(r, f, "MO BANCADA", toNum(r[C.ET2]), null, true);
  add(r, f, "MO BANCADA", toNum(r[C.ET3]), null, true);
  add(r, f, "MO BANCADA", toNum(r[C.ET6]), pct(C.ET6), true);
  if (!vazio(r[C.ET7]) && toNum(r[C.ET7]) !== 0) { const p = f === "Fam PI/TF" ? "MO CONFERÊNCIA PI/TF" : "MO CONFERÊNCIA BANCADA"; add(r, p, p, toNum(r[C.ET7]), pct(C.ET7), true); }
  const esp = ["SPCT0071 - AA", "ICP", "SPCT - IV", "SPCT - UV"].reduce((a, c) => a + toNum(r[c]), 0);
  if (esp !== 0) add(r, "MO EQUIPAMENTO BANCADA", "MO EQUIPAMENTO BANCADA", esp, null, false);
  if (ok(r[C.ET9])) add(r, "MO EQUIPAMENTO", "MO EQUIPAMENTO", toNum(r[C.ET9]), pos(pct(C.ET9)), true);
  if (ok(r[C.ET14])) add(r, "MO EQUIPAMENTO", "MO EQUIPAMENTO", toNum(r[C.ET14]), null, true);
  if (ok(r[C.ET10])) add(r, "MO CONFERÊNCIA EQUIPAMENTO", "MO CONFERÊNCIA EQUIPAMENTO", toNum(r[C.ET10]), pos(pct(C.ET10)), true);
  if (ok(r[C.ET15])) add(r, "MO CONFERÊNCIA EQUIPAMENTO", "MO CONFERÊNCIA EQUIPAMENTO", toNum(r[C.ET15]), null, true);
  if (toNum(r[C.ET16]) !== 0) add(r, "REVISÃO FINAL", "REVISÃO FINAL", toNum(r[C.ET16]), null, false);
  if (toNum(r[C.ET8]) !== 0) add(r, "DISSOLUTOR", "DISSOLUTOR", toNum(r[C.ET8]), null, false);
}
for (const r of t.base_tc) {
  const e = r[C.EQUIP];
  if (!(typeof e === "string" && /^CR(LQ|GS)/.test(e))) continue;
  const proc = e.startsWith("CRLQ") ? "EQUIPAMENTO HPLC" : "EQUIPAMENTO CG";
  add(r, e, proc, toNum(r[C.ET12_N]) * toNum(r[C.ET12_T]), pct(C.ET12_T) ?? 0, true);
  add(r, e, proc, toNum(r[C.ET13_N]) * toNum(r[C.ET13_T]), null, true);
}

// ── fDemanda: Periodo × MPR → lotes (a partir das linhas de demanda do período) ─
function demandaPeriodo(p) {
  const m = new Map(); for (const l of E.linhasDemanda(ctx, p).linhas) m.set(l.mpr, (m.get(l.mpr) || 0) + l.qtde); return m;
}

// ── fCarga: componente × demanda → CM_Min, agrupado por DESTINO+PROCESSO ──
function carga(p) {
  const dem = demandaPeriodo(p), ag = new Map();
  for (const c of comp) {
    const k = c.dest + "|" + c.proc;
    if (!ag.has(k)) ag.set(k, 0);
    if (!dem.has(c.mpr)) continue;
    const L = dem.get(c.mpr);
    let cm;
    if (c.fam === "Impureza Skip Test" && skip.has(c.mpr)) cm = c.tc * up(L / skip.get(c.mpr));
    else {
      const lxc = camp.get(c.mpr) || 0, nc = c.campanha && lxc > 0 ? up(L / lxc) : 0;
      if (nc === 0 || c.pct === null) cm = c.tc * L;
      else { const pr = Math.min(nc, L); cm = pr * c.tc + (L - pr) * c.tc * c.pct; }
    }
    ag.set(k, ag.get(k) + cm);
  }
  return ag;
}

// ── fMprRecurso: (TIPO, CHAVE, MPR) distintos → demanda do recurso ────────
const mprRec = new Set();
const addR = (tipo, chave, mpr) => mprRec.add(`${tipo}|${chave}|${mpr}`);
for (const r of t.base_tc) {
  if (r[C.MPR] === null || r[C.MPR] === "") continue;
  const f = r[C.FAM], m = r[C.MPR], cheio = (cols) => cols.some((c) => !vazio(r[c]));
  if (f !== null && f !== "Terceirizado") addR("FAM", f, m);
  if (cheio([C.ET7])) addR("PROC", f === "Fam PI/TF" ? "MO CONFERÊNCIA PI/TF" : "MO CONFERÊNCIA BANCADA", m);
  if (cheio([C.ET10, C.ET15])) addR("PROC", "MO CONFERÊNCIA EQUIPAMENTO", m);
  if (cheio(["SPCT0071 - AA", "ICP", "SPCT - IV", "SPCT - UV"])) addR("PROC", "MO EQUIPAMENTO BANCADA", m);
  if (cheio([C.ET16])) addR("PROC", "REVISÃO FINAL", m);
  if (cheio([C.ET8])) addR("PROC", "DISSOLUTOR", m);
  if (cheio([C.ET9, C.ET14])) addR("PROC", "MO EQUIPAMENTO", m);
}
for (const r of t.base_tc) { const e = r[C.EQUIP]; if (typeof e === "string" && /^CR(LQ|GS)/.test(e)) addR("EQUIP", e, r[C.MPR]); }
function demRecurso(p) {
  const dem = demandaPeriodo(p), g = new Map();
  for (const s of mprRec) { const [tp, ch, m] = s.split("|"); if (!dem.has(m)) continue; const k = tp + "|" + ch; g.set(k, (g.get(k) || 0) + dem.get(m)); }
  return g;
}

// ── fCap: linhas de carga + recursos (junção FAMILIA+PROCESSO, senão FAMILIA) ─
function cap(p) {
  const cal = p.nivel === "Semanal" ? ctx.calSem.get(p.chave) : ctx.calMes.get(E.MESES.indexOf(p.mes) + 1);
  const fam = t.familias, porFP = new Map(fam.map((f) => [f.FAMILIA + "|" + f.PROCESSO, f])), porF = new Map();
  for (const f of fam) if (!porF.has(f.FAMILIA)) porF.set(f.FAMILIA, f);
  const hrs = (f) => toNum(cal[`${f.CATEGORIA === "Equipamento" ? "EQP" : "MO"} - ${toNum(f.TURNOS) || 3}`]);
  const rec = (f) => { const n = toNum(f.TURNOS) || 3; return [f["QTDE 1° TURNO"], f["QTDE 2° TURNO"], f["QTDE 3° TURNO"]].slice(0, n).reduce((a, b) => a + toNum(b), 0) / n; };
  const dr = demRecurso(p), out = [];
  for (const [k, cm] of carga(p)) {
    const [dest, proc] = k.split("|");
    let f = porFP.get(k);
    if (!f || !(rec(f) * hrs(f))) f = porF.get(dest) && rec(porF.get(dest)) * hrs(porF.get(dest)) ? porF.get(dest) : f || porF.get(dest);
    if (!f || !hrs(f) || !rec(f)) continue;
    const demanda = dr.get("FAM|" + dest) ?? dr.get("PROC|" + proc) ?? dr.get("EQUIP|" + dest) ?? 0;
    out.push({ FAMILIA: dest, PROCESSO: proc, CM_HRS: cm / 60, HRS: hrs(f), RECURSO: rec(f), OEE: toNum(f["OEE REAL"]), DEMANDA: demanda });
  }
  return out;
}

const { meses, semanas } = E.listarPeriodos(ctx);
const periodos = [...meses, ...semanas.map((s) => ({ ...s, fonte: "real" }))];
for (const p of periodos) {
  test(`modelo BI = motor · ${p.rotulo}`, () => {
    const a = E.capProposta(ctx, p).linhas, b = cap(p);
    const ib = new Map(b.map((l) => [l.PROCESSO + "|" + l.FAMILIA, l]));
    assert.equal(b.length, a.length, "linhas");
    for (const l of a) {
      const x = ib.get(l.PROCESSO + "|" + l.FAMILIA); assert.ok(x, "falta " + l.FAMILIA);
      for (const [ka, kb] of [["CM_HRS", "CM_HRS"], ["HRS_TURNO", "HRS"], ["RECURSO", "RECURSO"], ["OEE", "OEE"], ["DEMANDA", "DEMANDA"]])
        assert.ok(Math.abs(l[ka] - x[kb]) <= 1e-6 * Math.max(1, Math.abs(l[ka])), `${l.FAMILIA} ${ka}: ${l[ka]} vs ${x[kb]}`);
    }
  });
}
