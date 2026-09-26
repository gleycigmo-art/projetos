/*
 * FLOW LAB — motor de cálculo de capacidade (fonte única de verdade no navegador).
 *
 * Porta fiel de capacidade/motor.py (mesmas regras, validada por teste de paridade):
 *   demanda por MPR → carga (min) por família/processo com campanha e skip →
 *   capacidade (pessoas × horas do calendário × OEE) → indicadores.
 * Funções puras: nenhum acesso a DOM. Funciona no navegador (window.FlowEngine) e no Node.
 */
(function (root) {
  "use strict";

  const MESES = ["Jan/2026", "Fev/2026", "Mar/2026", "Abr/2026", "Mai/2026", "Jun/2026",
    "Jul/2026", "Ago/2026", "Set/2026", "Out/2026", "Nov/2026", "Dez/2026"];
  const ANO = 2026;

  const C = {
    MPR: "Método e versão (DocNix)", FAM: "FAM proposta", EQUIP: "Equipamento Proposto",
    ET1: "1 - PREPARO (PADRÃO E SOLUÇÕES) - BANCADA", ET2: "2 - PREPARO AMOSTRA - BANCADA",
    ET3: "3 - ANÁLISE FQ - BANCADA", ET6: "6 - FINALIZAÇÃO DOS DADOS - BANCADA",
    ET7: "7 - CONFERÊNCIA DO PREPARO - BANCADA", ET8: "8 - TEMPO DO DISSOLUTOR - BANCADA",
    ET9: "9 - MONTAGEM DO EQUIPAMENTO - LEITURA", ET10: "10 - CONFERÊNCIA EQUIPAMENTO - LEITURA",
    ET12_N: "12 - Nº INJEÇÕES PADRÃO - LEITURA", ET12_T: "12 - TEMPO PADRÃO (CORRIDA) - LEITURA",
    ET13_N: "13 - Nº INJEÇÕES AMOSTRA/ DILUENTE/ PLACEBO - LEITURA",
    ET13_T: "13 - TEMPO AMOSTRA/ DILUENTE/ PLACEBO (CORRIDA) - LEITURA",
    ET14: "14 - PROCESSAMENTO DA CORRIDA - LEITURA", ET15: "15 - CONFERÊNCIA PROCESSAMENTO DA CORRIDA",
    ET16: "16 - REVISÃO LAUDO - REVISÃO",
  };
  const ESPECTRO = ["SPCT0071 - AA", "ICP", "SPCT - IV", "SPCT - UV"];
  const FAM_SKIP = "Impureza Skip Test";
  const FAM_PITF = "Fam PI/TF";
  const FIFO_DIAS_LEGADO = 22.97;
  const TURNOS_COLS = ["QTDE 1° TURNO", "QTDE 2° TURNO", "QTDE 3° TURNO"];

  // Colunas mínimas de cada tabela (usadas também para validar importações)
  const COLUNAS = {
    base_tc: [C.MPR, C.FAM, C.EQUIP, ...ESPECTRO, C.ET1, C.ET2, C.ET3, C.ET6, C.ET7, C.ET8, C.ET9,
      C.ET10, C.ET12_N, C.ET12_T, C.ET13_N, C.ET13_T, C.ET14, C.ET15, C.ET16],
    demanda_mensal: ["CONSIDERAR?", "CÓD", "DESCRIÇÃO", "MÉTODO DOCNIX", ...MESES, "complexidade"],
    demanda_semanal: ["Cod", "Chave semana", "STATUS", "Lote semana", "MÉTODO DOCNIX", "Status2", "Complexidade"],
    familias: ["PROCESSO", "FAMILIA", "CLUSTER", "CATEGORIA", "TURNOS", ...TURNOS_COLS, "OEE REAL"],
    calendario_mes: ["Mês Nº", "Dias Úteis", "MO - 1", "MO - 2", "MO - 3", "EQP - 1", "EQP - 2", "EQP - 3"],
    calendario_semana: ["Semana Nº", "Semana Chave", "Dias Úteis", "MO - 1", "MO - 2", "MO - 3", "EQP - 1", "EQP - 2", "EQP - 3"],
    campanha: ["CÓD", "QTD LOTES CAMPANHA"],
    skip: ["MPR", "Qtde Lotes Skip"],
    etapas_campanha: ["Etapa", "% tempo"],
    ausencias: ["TIPO", "QTDE_SEMANA", "QTDE_MEDIA_ANO", "SEMANA_REF"],
    dia_util: ["DIA", "PESO"],
  };

  // ── utilitários ─────────────────────────────────────────────────────────
  function toNum(v) {
    if (v === null || v === undefined || v === "") return 0;
    if (typeof v === "number") return Number.isFinite(v) ? v : 0;
    const n = Number(String(v).replace(",", "."));
    return Number.isFinite(n) ? n : 0;
  }
  const vazio = (v) => v === null || v === undefined || v === "" || (typeof v === "number" && v === 0);
  const roundup = (x) => Math.ceil(x - 1e-12);
  const ehEquip = (e) => typeof e === "string" && (e.startsWith("CRLQ") || e.startsWith("CRGS"));
  const soma = (arr) => arr.reduce((a, b) => a + b, 0);

  /** Converte {cols, rows} (formato compacto) em lista de objetos. */
  function hidratar(t) {
    if (Array.isArray(t)) return t;
    return t.rows.map((r) => Object.fromEntries(t.cols.map((c, i) => [c, r[i]])));
  }

  // ── calendário ISO ──────────────────────────────────────────────────────
  function segundaDaSemanaISO(ano, semana) {
    const jan4 = new Date(Date.UTC(ano, 0, 4));
    const dia = (jan4.getUTCDay() + 6) % 7; // 0 = segunda
    const seg = new Date(jan4);
    seg.setUTCDate(jan4.getUTCDate() - dia + (semana - 1) * 7);
    return seg;
  }
  const chaveSemana = (ano, n) => Number(`${ano}${n}`);

  // ── preparação ──────────────────────────────────────────────────────────
  /** Indexa o conjunto de dados; chame uma vez por importação. */
  function preparar(dataset) {
    const t = {};
    for (const [k, v] of Object.entries(dataset.tabelas)) t[k] = hidratar(v);

    const tc = t.base_tc.filter((r) => r[C.MPR] !== null && r[C.MPR] !== undefined && r[C.MPR] !== "");
    const campanha = new Map();
    for (const r of t.campanha) if (!campanha.has(r["CÓD"])) campanha.set(r["CÓD"], r["QTD LOTES CAMPANHA"]);
    const skip = new Map();
    const vistosSkip = new Set();
    for (const r of t.skip) {
      if (vistosSkip.has(r.MPR)) continue;
      vistosSkip.add(r.MPR);
      const q = r["Qtde Lotes Skip"];
      if (q !== null && q !== undefined && toNum(q) >= 1) skip.set(r.MPR, toNum(q));
    }
    const pctTab = new Map(t.etapas_campanha.map((r) => [r.Etapa, r["% tempo"]]));
    const pct = {};
    for (const e of [C.ET1, C.ET6, C.ET7, C.ET9, C.ET10, C.ET12_T]) {
      const v = pctTab.get(e + "_CAMPANHA");
      pct[e] = v === undefined || v === null ? null : toNum(v);
    }

    const pesoDia = [1, 1, 1, 1, 1, 0, 0];
    (t.dia_util || []).forEach((r) => {
      const i = ["seg", "ter", "qua", "qui", "sex", "sáb", "dom"].indexOf(String(r.DIA).toLowerCase());
      if (i >= 0) pesoDia[i] = toNum(r.PESO);
    });

    const semanasReais = new Set(t.demanda_semanal.map((r) => toNum(r["Chave semana"])).filter(Boolean));
    const calSem = new Map(t.calendario_semana.map((r) => [toNum(r["Semana Chave"]), r]));
    const calMes = new Map(t.calendario_mes.map((r) => [toNum(r["Mês Nº"]), r]));

    return { t, tc, campanha, skip, pct, pesoDia, semanasReais, calSem, calMes,
      tcEquip: t.base_tc.filter((r) => ehEquip(r[C.EQUIP])), meta: dataset.meta || {} };
  }

  // ── períodos ────────────────────────────────────────────────────────────
  function listarPeriodos(ctx) {
    const meses = MESES.map((m, i) => ({ nivel: "Mensal", id: m, rotulo: m, mes: m, ordem: i }));
    const semanas = [];
    for (const [chave, r] of [...ctx.calSem.entries()].sort((a, b) => a[0] - b[0])) {
      const n = toNum(r["Semana Nº"]);
      const ini = segundaDaSemanaISO(ANO, n);
      const fim = new Date(ini); fim.setUTCDate(ini.getUTCDate() + 6);
      semanas.push({ nivel: "Semanal", id: chave, semana: n, chave, ordem: n,
        rotulo: `Semana ${String(n).padStart(2, "0")} · ${ANO}`, inicio: ini, fim,
        temReal: ctx.semanasReais.has(chave) });
    }
    return { meses, semanas };
  }

  /** Fração de cada mês de 2026 dentro de uma semana ISO, ponderada pelos dias úteis. */
  function fracoesSemana(ctx, n) {
    const ini = segundaDaSemanaISO(ANO, n);
    const pesoMesTotal = (m) => {
      let s = 0; const d = new Date(Date.UTC(ANO, m, 1));
      while (d.getUTCMonth() === m) { s += ctx.pesoDia[(d.getUTCDay() + 6) % 7]; d.setUTCDate(d.getUTCDate() + 1); }
      return s;
    };
    const fr = new Map();
    for (let i = 0; i < 7; i++) {
      const d = new Date(ini); d.setUTCDate(ini.getUTCDate() + i);
      if (d.getUTCFullYear() !== ANO) continue;
      const m = d.getUTCMonth();
      fr.set(m, (fr.get(m) || 0) + ctx.pesoDia[i]);
    }
    for (const [m, p] of fr) fr.set(m, p / pesoMesTotal(m));
    return fr;
  }

  // ── demanda ─────────────────────────────────────────────────────────────
  /**
   * Linhas de demanda do período: [{mpr, cod, qtde, complexidade, semTempos}].
   * Mensal: plano do mês. Semanal: quarentena real (se houver e fonte='real') ou
   * projeção do plano mensal rateada pelos dias úteis.
   */
  function linhasDemanda(ctx, periodo) {
    if (periodo.nivel === "Semanal") {
      const real = periodo.fonte !== "projecao" && ctx.semanasReais.has(periodo.chave);
      if (real) {
        return {
          fonte: "real", n: 1,
          linhas: ctx.t.demanda_semanal.filter((r) => toNum(r["Chave semana"]) === periodo.chave).map((r) => ({
            mpr: r["MÉTODO DOCNIX"], cod: r.Cod, qtde: toNum(r["Lote semana"]),
            complexidade: toNum(r.Complexidade), semTempos: r.Status2 === "Sem tempos",
          })),
        };
      }
      const fr = fracoesSemana(ctx, periodo.semana);
      const linhas = [];
      for (const r of ctx.t.demanda_mensal) {
        if (r["CONSIDERAR?"] !== "SIM") continue;
        let q = 0, tem = false;
        for (const [m, f] of fr) { const v = r[MESES[m]]; if (v !== null && v !== undefined) { tem = true; q += toNum(v) * f; } }
        if (tem) linhas.push({ mpr: r["MÉTODO DOCNIX"], cod: r["CÓD"], qtde: q,
          complexidade: toNum(r.complexidade), semTempos: r["MÉTODO DOCNIX"] === "Método não encontrado" });
      }
      return { fonte: "projecao", n: 1, linhas };
    }
    const meses = periodo.meses || [periodo.mes];
    const linhas = [];
    for (const r of ctx.t.demanda_mensal) {
      if (r["CONSIDERAR?"] !== "SIM") continue;
      for (const m of meses) {
        const v = r[m];
        if (v === null || v === undefined) continue; // o Unpivot descarta nulos
        linhas.push({ mpr: r["MÉTODO DOCNIX"], cod: r["CÓD"], qtde: toNum(v),
          complexidade: toNum(r.complexidade), semTempos: r["MÉTODO DOCNIX"] === "Método não encontrado" });
      }
    }
    return { fonte: "plano", n: meses.length, linhas };
  }

  function demandaMPR(dem) {
    const m = new Map();
    for (const l of dem.linhas) m.set(l.mpr, (m.get(l.mpr) || 0) + l.qtde);
    return m; // MPR → Demanda_Total
  }

  // ── regras e cenário ────────────────────────────────────────────────────
  const REGRAS_CORRIGIDAS = { campanhaCorrigida: true, fifoDiasUteis: true, skipPelaTabela: true };
  const REGRAS_LEGADO = { campanhaCorrigida: false, fifoDiasUteis: false, skipPelaTabela: false };

  function cenarioVazio() {
    return { fatorDemanda: 1, fatorFamilia: {}, turnos: {}, remanejamentos: [], oeeMO: null,
      oeeEquip: null, oee: {}, ausencias: null, lotesAdiados: 0 };
  }

  // ── carga (TC_COM_CAMPANHA) ─────────────────────────────────────────────
  function calcularCarga(ctx, regras, demTotal, nPer, fatorLinha) {
    const pct = ctx.pct;
    const lotesDe = (mpr, fam) => {
      if (!demTotal.has(mpr) || nPer <= 0) return [0, 0];
      return [demTotal.get(mpr) * (fatorLinha ? fatorLinha(fam) : 1), nPer];
    };
    const nSkip = (mpr) => (ctx.skip.has(mpr) ? ctx.skip.get(mpr) : 1);
    const ehSkip = (mpr, fam, blocoEquip) =>
      (regras.skipPelaTabela || blocoEquip) ? fam === FAM_SKIP && ctx.skip.has(mpr) : fam === FAM_SKIP;
    const nCamp = (mpr, lotes) => {
      const lxc = roundup(toNum(ctx.campanha.get(mpr)));
      return lxc > 0 ? roundup(lotes / lxc) : 0;
    };

    // etapas: [[tc, pct|null]]
    function carga(mpr, fam, etapas, comCampanha = true) {
      const [lotes, meses] = lotesDe(mpr, fam);
      if (meses <= 0) return 0;
      const tcTotal = soma(etapas.map((e) => e[0]));
      if (ehSkip(mpr, fam, false)) {
        const ns = nSkip(mpr);
        return tcTotal * (ns > 0 ? roundup(lotes / ns) : 0) / meses;
      }
      const nc = comCampanha ? nCamp(mpr, lotes) : 0;
      if (nc === 0) return tcTotal * lotes / meses;
      let total = 0;
      if (!regras.campanhaCorrigida) {
        if (nc === 1) return tcTotal / meses; // legado: TC cobrado 1x para todos os lotes
        for (const [tc, p] of etapas) total += p !== null ? nc * tc + (lotes - nc) * tc * p : tc * lotes;
        return total / meses;
      }
      const prim = Math.min(nc, lotes); // 1º lote de cada campanha com tempo cheio
      for (const [tc, p] of etapas) total += p === null ? tc * lotes : prim * tc + (lotes - prim) * tc * p;
      return total / meses;
    }
    const soPositivo = (p) => (p !== null && p > 0 ? p : null);
    const g = (r, c) => toNum(r[c]);
    const saida = []; // {mpr, fam, proc, cm}

    for (const r of ctx.tc) {
      const mpr = r[C.MPR], fam = r[C.FAM];
      saida.push({ mpr, fam, proc: "MO BANCADA", cm: carga(mpr, fam,
        [[g(r, C.ET1), pct[C.ET1]], [g(r, C.ET2), null], [g(r, C.ET3), null], [g(r, C.ET6), pct[C.ET6]]]) });

      if (!vazio(r[C.ET7]) && g(r, C.ET7) !== 0) {
        const proc = fam === FAM_PITF ? "MO CONFERÊNCIA PI/TF" : "MO CONFERÊNCIA BANCADA";
        saida.push({ mpr, fam: proc, proc, cm: carga(mpr, fam, [[g(r, C.ET7), pct[C.ET7]]]) });
      }
      const esp = soma(ESPECTRO.map((c) => g(r, c)));
      if (esp !== 0) saida.push({ mpr, fam: "MO EQUIPAMENTO BANCADA", proc: "MO EQUIPAMENTO BANCADA",
        cm: carga(mpr, fam, [[esp, null]], false) });

      for (const [et, p] of [[C.ET9, soPositivo(pct[C.ET9])], [C.ET14, null]]) {
        const v = r[et];
        if (v !== null && v !== undefined && toNum(v) !== 0 && String(v) !== "S")
          saida.push({ mpr, fam: "MO EQUIPAMENTO", proc: "MO EQUIPAMENTO", cm: carga(mpr, fam, [[toNum(v), p]]) });
      }
      for (const [et, p] of [[C.ET10, soPositivo(pct[C.ET10])], [C.ET15, null]]) {
        const v = r[et];
        if (v !== null && v !== undefined && toNum(v) !== 0 && String(v) !== "S")
          saida.push({ mpr, fam: "MO CONFERÊNCIA EQUIPAMENTO", proc: "MO CONFERÊNCIA EQUIPAMENTO",
            cm: carga(mpr, fam, [[toNum(v), p]]) });
      }
      for (const [et, proc] of [[C.ET16, "REVISÃO FINAL"], [C.ET8, "DISSOLUTOR"]]) {
        if (g(r, et) !== 0) saida.push({ mpr, fam: proc, proc, cm: carga(mpr, fam, [[g(r, et), null]], false) });
      }
    }

    // Bloco H — tempo de máquina HPLC/CG
    const pct12 = toNum(pct[C.ET12_T]);
    for (const r of ctx.tcEquip) {
      const mpr = r[C.MPR], fam = r[C.FAM], equip = r[C.EQUIP];
      const tc12 = g(r, C.ET12_N) * g(r, C.ET12_T);
      const tc13 = g(r, C.ET13_N) * g(r, C.ET13_T);
      const [lotes, meses] = lotesDe(mpr, fam);
      let cm = 0;
      if (meses > 0) {
        if (ehSkip(mpr, fam, true)) {
          const ns = nSkip(mpr);
          cm = (tc12 + tc13) * (ns > 0 ? roundup(lotes / ns) : 0) / meses;
        } else {
          const nc = nCamp(mpr, lotes);
          if (nc === 0) cm = (tc12 + tc13) * lotes / meses;
          else if (regras.campanhaCorrigida) {
            const prim = Math.min(nc, lotes);
            cm = (prim * tc12 + (lotes - prim) * tc12 * pct12 + tc13 * lotes) / meses;
          } else if (nc === 1) cm = (tc12 + tc13) / meses;
          else cm = (nc * tc12 + (lotes - nc) * tc12 * pct12 + tc13 * lotes) / meses;
        }
      }
      saida.push({ mpr, fam: equip, proc: equip.startsWith("CRLQ") ? "EQUIPAMENTO HPLC" : "EQUIPAMENTO CG", cm });
    }
    return saida;
  }

  // ── capacidade ──────────────────────────────────────────────────────────
  function linhaCalendario(ctx, periodo) {
    if (periodo.nivel === "Semanal") return [ctx.calSem.get(periodo.chave)].filter(Boolean);
    const meses = periodo.meses || [periodo.mes];
    return meses.map((m) => ctx.calMes.get(MESES.indexOf(m) + 1)).filter(Boolean);
  }

  function diasUteis(ctx, periodo) {
    const cal = linhaCalendario(ctx, periodo);
    return cal.length ? soma(cal.map((r) => toNum(r["Dias Úteis"]))) / cal.length : 0;
  }

  /** Turnos por família já com os ajustes e remanejamentos do cenário. */
  function turnosComCenario(ctx, cen) {
    const mapa = new Map();
    for (const r of ctx.t.familias) mapa.set(r.FAMILIA, TURNOS_COLS.map((c) => toNum(r[c])));
    for (const [fam, arr] of Object.entries(cen.turnos || {})) if (mapa.has(fam)) mapa.set(fam, arr.map(toNum));
    for (const mv of cen.remanejamentos || []) {
      const hc = toNum(mv.hc), t = mv.turno;
      if (!mapa.has(mv.de) || !mapa.has(mv.para) || !hc) continue;
      const idx = t === null || t === undefined || t === "todos" ? [0, 1, 2] : [Number(t)];
      const de = mapa.get(mv.de).slice(), para = mapa.get(mv.para).slice();
      for (const i of idx) { const q = Math.min(hc, de[i]); de[i] -= q; para[i] += q; }
      mapa.set(mv.de, de); mapa.set(mv.para, para);
    }
    return mapa;
  }

  function capacidade(ctx, periodo, cen) {
    const cal = linhaCalendario(ctx, periodo);
    const turnos = turnosComCenario(ctx, cen);
    return ctx.t.familias.map((r) => {
      const n = Math.round(toNum(r.TURNOS)) || 3;
      const pref = r.CATEGORIA === "Equipamento" ? "EQP" : "MO";
      const hrs = cal.length ? soma(cal.map((c) => toNum(c[`${pref} - ${n}`]))) / cal.length : 0;
      const q = turnos.get(r.FAMILIA);
      const qBase = TURNOS_COLS.map((c) => toNum(r[c]));
      const recurso = soma(q.slice(0, n)) / n;
      let oee = toNum(r["OEE REAL"]);
      if (r.CATEGORIA === "Mão de Obra" && cen.oeeMO !== null && cen.oeeMO !== undefined) oee = cen.oeeMO;
      if (r.CATEGORIA === "Equipamento" && cen.oeeEquip !== null && cen.oeeEquip !== undefined) oee *= cen.oeeEquip;
      if (cen.oee && cen.oee[r.FAMILIA] !== undefined) oee = cen.oee[r.FAMILIA];
      return { PROCESSO: r.PROCESSO, FAMILIA: r.FAMILIA, CLUSTER: r.CLUSTER, CATEGORIA: r.CATEGORIA,
        TURNOS: n, QTDE: q, QTDE_BASE: qBase, RECURSO: recurso, RECURSO_BASE: soma(qBase.slice(0, n)) / n,
        HRS_TURNO: hrs, TOTAL: recurso * hrs, OEE: oee, OEE_BASE: toNum(r["OEE REAL"]) };
    });
  }

  // ── demanda por recurso (Soma Demanda Media) ────────────────────────────
  function demandaPorRecurso(ctx, dmpr) {
    const somaMPR = (mprs) => {
      let s = 0, algum = false;
      for (const m of new Set(mprs)) if (dmpr.has(m)) { s += dmpr.get(m); algum = true; }
      return algum ? s : null;
    };
    const porFam = new Map();
    const grupos = new Map();
    for (const r of ctx.tc) {
      const f = r[C.FAM];
      if (f === null || f === undefined || f === "Terceirizado") continue;
      if (!grupos.has(f)) grupos.set(f, []);
      grupos.get(f).push(r[C.MPR]);
    }
    for (const [f, mprs] of grupos) porFam.set(f, somaMPR(mprs));

    const preenchido = (cols, filtro) =>
      somaMPR(ctx.tc.filter((r) => (!filtro || filtro(r)) && cols.some((c) => !vazio(r[c]))).map((r) => r[C.MPR]));
    const pitf = (r) => r[C.FAM] === FAM_PITF;
    const porProc = new Map([
      ["MO CONFERÊNCIA BANCADA", preenchido([C.ET7], (r) => !pitf(r))],
      ["MO CONFERÊNCIA PI/TF", preenchido([C.ET7], pitf)],
      ["MO CONFERÊNCIA EQUIPAMENTO", preenchido([C.ET10, C.ET15])],
      ["MO EQUIPAMENTO BANCADA", preenchido(ESPECTRO)],
      ["REVISÃO FINAL", preenchido([C.ET16])],
      ["DISSOLUTOR", preenchido([C.ET8])],
      ["MO EQUIPAMENTO", preenchido([C.ET9, C.ET14])],
    ]);
    const porEquip = new Map();
    const eqGr = new Map();
    for (const r of ctx.tcEquip) { if (!eqGr.has(r[C.EQUIP])) eqGr.set(r[C.EQUIP], []); eqGr.get(r[C.EQUIP]).push(r[C.MPR]); }
    for (const [e, mprs] of eqGr) porEquip.set(e, somaMPR(mprs) || 0);
    return { porFam, porProc, porEquip };
  }

  // ── CAP_PROPOSTA ────────────────────────────────────────────────────────
  function capProposta(ctx, periodo, opts = {}) {
    const regras = opts.regras || REGRAS_CORRIGIDAS;
    const cen = Object.assign(cenarioVazio(), opts.cenario || {});
    const dem = opts.demanda || linhasDemanda(ctx, periodo);
    const dmprBase = demandaMPR(dem);
    const lotesTotal = soma(dem.linhas.map((l) => l.qtde));
    let fator = cen.fatorDemanda;
    if (cen.lotesAdiados && lotesTotal > 0) fator *= Math.max(0, lotesTotal - cen.lotesAdiados) / lotesTotal;
    const fatorFam = cen.fatorFamilia || {};
    const temFatorFam = Object.keys(fatorFam).length > 0;
    const fatorLinha = (fator !== 1 || temFatorFam) ? (fam) => fator * (fatorFam[fam] ?? 1) : null;

    const carga = calcularCarga(ctx, regras, dmprBase, dem.n, fatorLinha);
    const cap = capacidade(ctx, periodo, cen);
    const dmprMedia = new Map([...dmprBase].map(([k, v]) => [k, v * fator / dem.n]));
    const { porFam, porProc, porEquip } = demandaPorRecurso(ctx, dmprMedia);
    const dias = regras.fifoDiasUteis ? diasUteis(ctx, periodo) : FIFO_DIAS_LEGADO;

    const ag = new Map();
    for (const c of carga) {
      const k = `${c.fam}\u0001${c.proc}`;
      ag.set(k, (ag.get(k) || 0) + c.cm);
    }
    const capIdx = new Map(cap.map((c) => [`${c.FAMILIA}\u0001${c.PROCESSO}`, c]));
    const capFam = new Map();
    for (const c of cap) if (!capFam.has(c.FAMILIA)) capFam.set(c.FAMILIA, c);

    const linhas = [], semRecurso = [];
    for (const [k, cm] of ag) {
      const [fam, proc] = k.split("\u0001");
      let c = capIdx.get(k);
      if (!c || !c.TOTAL) c = capFam.get(fam) && capFam.get(fam).TOTAL ? capFam.get(fam) : c || capFam.get(fam);
      if (!c || !c.HRS_TURNO || (!c.RECURSO_BASE && !c.RECURSO)) { if (cm > 0) semRecurso.push({ FAMILIA: fam, PROCESSO: proc, CM_HRS: cm / 60 }); continue; }
      const cmH = cm / 60;
      const total = c.TOTAL, oee = c.OEE, hrs = c.HRS_TURNO;
      const demRec = [porFam.get(fam), porProc.get(proc), porEquip.get(fam)].find((v) => v !== null && v !== undefined) ?? 0;
      const hcNec = hrs && oee ? cmH / (hrs * oee) : null;
      const carreg = total && oee ? cmH / total / oee : (cmH > 0 ? Infinity : 0);
      linhas.push({
        PROCESSO: proc, FAMILIA: fam, CLUSTER: c.CLUSTER, CATEGORIA: c.CATEGORIA,
        CM_HRS: cmH, RECURSO: c.RECURSO, RECURSO_BASE: c.RECURSO_BASE, QTDE: c.QTDE, TURNOS: c.TURNOS,
        TOTAL_RECURSO: total, HRS_TURNO: hrs, OEE: oee, OEE_NEC: total ? cmH / total : null,
        HC_NECESSARIO: hcNec, CARREGAMENTO: carreg, GAP_HC: hcNec !== null ? hcNec - c.RECURSO : null,
        DEMANDA: demRec, TAKT: demRec ? total / demRec : null, TC_MEDIO: demRec ? cmH / demRec : null,
        LOTES_DIA: demRec && dias ? demRec / dias : null,
        DEMANDA_MAX: demRec && carreg && Number.isFinite(carreg) ? demRec / carreg : null,
      });
    }
    linhas.sort((a, b) => String(a.CLUSTER).localeCompare(String(b.CLUSTER)) || a.PROCESSO.localeCompare(b.PROCESSO) || a.FAMILIA.localeCompare(b.FAMILIA));
    return { linhas, semRecurso, dias, dem, fator, lotesTotal, cap };
  }

  // ── indicadores do painel ───────────────────────────────────────────────
  function ausenciasPeriodo(ctx, periodo, cen) {
    const rows = ctx.t.ausencias || [];
    const ref = rows.length ? toNum(rows[0].SEMANA_REF) : null;
    const daSemana = periodo.nivel === "Semanal" && ref === periodo.semana;
    const tipos = rows.filter((r) => !/^P\/D$|^Total/i.test(String(r.TIPO))).map((r) => ({
      tipo: r.TIPO, semana: toNum(r.QTDE_SEMANA), ano: toNum(r.QTDE_MEDIA_ANO),
      valor: daSemana ? toNum(r.QTDE_SEMANA) : toNum(r.QTDE_MEDIA_ANO) }));
    const tot = rows.find((r) => /^Total/i.test(String(r.TIPO)));
    let total = tot ? (daSemana ? toNum(tot.QTDE_SEMANA) : toNum(tot.QTDE_MEDIA_ANO)) : soma(tipos.map((x) => x.valor));
    const mediaAno = tot ? toNum(tot.QTDE_MEDIA_ANO) : soma(tipos.map((x) => x.ano));
    const origem = daSemana ? "semana" : "média do ano";
    if (cen && cen.ausencias !== null && cen.ausencias !== undefined) total = toNum(cen.ausencias);
    return { tipos, total, mediaAno, origem, semanaRef: ref };
  }

  function complexidade(dem, fator = 1) {
    const cont = { 1: 0, 2: 0, 3: 0, 4: 0, sem: 0 };
    for (const l of dem.linhas) {
      const q = l.qtde * fator;
      if (l.complexidade >= 1 && l.complexidade <= 4) cont[Math.round(l.complexidade)] += q; else cont.sem += q;
    }
    const cls = cont[1] + cont[2] + cont[3] + cont[4];
    return { cont, classificados: cls, indice: cls ? (cont[1] + 2 * cont[2] + 3 * cont[3] + 4 * cont[4]) / cls : null,
      pct34: cls ? (cont[3] + cont[4]) / cls : null, lotes34: cont[3] + cont[4] };
  }

  const severidade = (x) => (x === null || x === undefined || !Number.isFinite(x) ? (x === Infinity ? "danger" : "neutral")
    : x > 1 ? "danger" : x >= 0.8 ? "warning" : "success");
  function veredito(ocup) {
    if (ocup === null || !Number.isFinite(ocup)) return "Sem dados suficientes";
    if (ocup <= 0.9) return "Atingível com folga";
    if (ocup <= 1.1) return "Atingível com baixa margem";
    return "Não atingível sem ação";
  }

  function calcular(ctx, periodo, opts = {}) {
    const cen = Object.assign(cenarioVazio(), opts.cenario || {});
    const cap = capProposta(ctx, periodo, { ...opts, cenario: cen });
    const mo = cap.linhas.filter((l) => l.CATEGORIA === "Mão de Obra");
    const eq = cap.linhas.filter((l) => l.CATEGORIA === "Equipamento");
    const famMO = cap.cap.filter((c) => c.CATEGORIA === "Mão de Obra");
    const nTurnos = Math.max(3, ...famMO.map((c) => c.TURNOS));
    const hcNecTurno = soma(mo.map((l) => l.HC_NECESSARIO || 0));
    const hcNecTotal = hcNecTurno * nTurnos;
    const hcRealTotal = soma(famMO.map((c) => soma(c.QTDE)));
    const hcTurnoReal = soma(famMO.map((c) => c.RECURSO));
    const aus = ausenciasPeriodo(ctx, periodo, cen);
    const hcEfetivo = hcRealTotal - aus.total;
    const ocupGlobal = hcRealTotal ? hcNecTotal / hcRealTotal : null;
    const ocupCorrigida = hcEfetivo > 0 ? hcNecTotal / hcEfetivo : null;
    const lotes = cap.lotesTotal * cap.fator;
    const cx = complexidade(cap.dem, cap.fator);
    const porTurno = [0, 1, 2].map((i) => soma(famMO.map((c) => c.QTDE[i] || 0)));
    const semTempos = soma(cap.dem.linhas.filter((l) => l.semTempos).map((l) => l.qtde)) * cap.fator;

    return {
      periodo, fonte: cap.dem.fonte, linhas: cap.linhas, mo, eq, semRecurso: cap.semRecurso, dias: cap.dias,
      kpi: {
        hcNecTurno, hcNecTotal, hcRealTotal, hcTurnoReal, hcEfetivo, nTurnos,
        gapHC: hcNecTotal - hcRealTotal, gapHCEfetivo: hcNecTotal - hcEfetivo,
        ocupGlobal, ocupCorrigida, severidade: severidade(ocupCorrigida), veredito: veredito(ocupCorrigida),
        lotes, lotesDia: cap.dias ? lotes / cap.dias : null, semTempos,
        cmMO: soma(mo.map((l) => l.CM_HRS)), cmEquip: soma(eq.map((l) => l.CM_HRS)),
        equipSobrecarga: eq.filter((l) => l.CARREGAMENTO > 1).length,
      },
      ausencias: aus, complexidade: cx, porTurno,
    };
  }

  // ── pergunta ao contrário ───────────────────────────────────────────────
  /** Maior fator de demanda com o qual toda a mão de obra fica com carregamento ≤ limite. */
  function capacidadeMaxima(ctx, periodo, opts = {}, limite = 1) {
    const base = Object.assign(cenarioVazio(), opts.cenario || {});
    const dem = linhasDemanda(ctx, periodo);
    const pior = (f) => {
      const r = capProposta(ctx, periodo, { ...opts, demanda: dem, cenario: { ...base, fatorDemanda: base.fatorDemanda * f } });
      let m = 0, fam = null;
      for (const l of r.linhas) if (l.CATEGORIA === "Mão de Obra" && l.CARREGAMENTO > m) { m = l.CARREGAMENTO; fam = l.FAMILIA; }
      return { m, fam, r };
    };
    const hoje = pior(1);
    if (!hoje.r.lotesTotal) return { fator: null, lotes: 0, gargalo: null };
    let lo = 0, hi = 1;
    while (pior(hi).m <= limite && hi < 64) { lo = hi; hi *= 2; }
    for (let i = 0; i < 22; i++) { const mid = (lo + hi) / 2; if (pior(mid).m <= limite) lo = mid; else hi = mid; }
    const lim = pior(lo === 0 ? 1e-6 : lo);
    const adiado = base.lotesAdiados && hoje.r.lotesTotal ? Math.max(0, hoje.r.lotesTotal - base.lotesAdiados) / hoje.r.lotesTotal : 1;
    const lotesAtuais = hoje.r.lotesTotal * base.fatorDemanda * adiado;
    return { fator: lo, lotes: lotesAtuais * lo, lotesAtuais, gargalo: lim.fam, gargaloHoje: hoje.fam, carregHoje: hoje.m };
  }

  // ── tendência ───────────────────────────────────────────────────────────
  function tendencia(ctx, nivel, opts = {}) {
    const { meses, semanas } = listarPeriodos(ctx);
    const lista = nivel === "Semanal" ? semanas.map((s) => ({ ...s, fonte: opts.fonteSemanal })) : meses;
    return lista.map((p) => {
      const r = calcular(ctx, p, opts);
      const piorMO = r.mo.reduce((a, l) => (!a || l.CARREGAMENTO > a.CARREGAMENTO ? l : a), null);
      return { periodo: p, fonte: r.fonte, ocupGlobal: r.kpi.ocupGlobal, ocupCorrigida: r.kpi.ocupCorrigida,
        gapHC: r.kpi.gapHC, lotes: r.kpi.lotes, indice: r.complexidade.indice, pct34: r.complexidade.pct34,
        ausencias: r.ausencias.total, piorFamilia: piorMO && piorMO.FAMILIA, piorCarreg: piorMO && piorMO.CARREGAMENTO,
        sobrecarregadas: r.mo.filter((l) => l.CARREGAMENTO > 1).map((l) => l.FAMILIA) };
    });
  }

  const api = { MESES, ANO, C, COLUNAS, REGRAS_CORRIGIDAS, REGRAS_LEGADO, toNum, preparar, listarPeriodos,
    linhasDemanda, capProposta, calcular, capacidadeMaxima, tendencia, cenarioVazio, severidade, veredito,
    diasUteis, fracoesSemana, segundaDaSemanaISO, chaveSemana, hidratar };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.FlowEngine = api;
})(typeof globalThis !== "undefined" ? globalThis : this);
