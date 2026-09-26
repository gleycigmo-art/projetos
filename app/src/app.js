/* FLOW LAB — interface. Todos os números vêm de FlowEngine; aqui só há apresentação. */
(function () {
  "use strict";
  const E = window.FlowEngine;
  const I = window.FlowImport;
  const $ = (s, el = document) => el.querySelector(s);
  const $$ = (s, el = document) => [...el.querySelectorAll(s)];

  // ── formatação pt-BR ────────────────────────────────────────────────────
  const nf = (d) => new Intl.NumberFormat("pt-BR", { minimumFractionDigits: d, maximumFractionDigits: d });
  const NF = [nf(0), nf(1), nf(2)];
  const num = (v, d = 1) => (v === null || v === undefined || Number.isNaN(v) ? "—" : !Number.isFinite(v) ? "∞" : NF[d].format(v));
  const pct = (v, d = 1) => (v === null || v === undefined || Number.isNaN(v) ? "—" : !Number.isFinite(v) ? "∞" : NF[d].format(v * 100) + "%");
  const sinal = (v, d = 1) => (v === null || v === undefined || !Number.isFinite(v) ? "—" : (v > 0 ? "+" : "") + NF[d].format(v));
  const dataBR = (d) => `${String(d.getUTCDate()).padStart(2, "0")}/${String(d.getUTCMonth() + 1).padStart(2, "0")}`;
  const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const sev = E.severidade;
  const dot = (s, extra = "") => `<span class="dot s-${s} ${extra}" aria-hidden="true"></span>`;
  const SEV_TXT = { success: "abaixo de 80%", warning: "entre 80% e 100%", danger: "acima de 100%", neutral: "sem dados" };

  // ── armazenamento local (conveniência; o app funciona sem ele) ───────────
  const store = {
    get(k, d) { try { const v = localStorage.getItem("flowlab:" + k); return v === null ? d : JSON.parse(v); } catch (e) { return d; } },
    set(k, v) { try { localStorage.setItem("flowlab:" + k, JSON.stringify(v)); return true; } catch (e) { return false; } },
    del(k) { try { localStorage.removeItem("flowlab:" + k); } catch (e) { /* sem armazenamento */ } },
  };

  // ── dados ───────────────────────────────────────────────────────────────
  const embutido = JSON.parse($("#dados-embutidos").textContent);
  let dataset = store.get("dataset", null) || embutido;
  let ctx = E.preparar(dataset);
  let periodos = E.listarPeriodos(ctx);

  const VIEWS = [
    ["painel", "Painel executivo", '<rect x="3" y="3" width="7" height="9" rx="1"/><rect x="14" y="3" width="7" height="5" rx="1"/><rect x="14" y="12" width="7" height="9" rx="1"/><rect x="3" y="16" width="7" height="5" rx="1"/>'],
    ["mapa", "Mapa de capacidade", '<path d="M3 5h18M3 12h18M3 19h18M8 5v14"/>'],
    ["gargalos", "Gargalos", '<path d="M4 4h16l-6 8v6l-4 2v-8z"/>'],
    ["mao", "Mão de obra", '<circle cx="9" cy="8" r="3"/><path d="M3 20c0-3.3 2.7-6 6-6s6 2.7 6 6"/><circle cx="17" cy="9" r="2.5"/><path d="M16 14.2c2.9.4 5 2.9 5 5.8"/>'],
    ["complexidade", "Complexidade", '<path d="M5 20V10M12 20V4M19 20v-7"/>'],
    ["tendencias", "Tendências", '<path d="M3 17l6-6 4 4 8-8"/><path d="M15 7h6v6"/>'],
    ["simulador", "Simulador", '<path d="M4 7h10M18 7h2M4 17h4M12 17h8"/><circle cx="16" cy="7" r="2"/><circle cx="10" cy="17" r="2"/>'],
    ["importacao", "Importação", '<path d="M12 16V4m0 0L8 8m4-4l4 4M5 20h14"/>'],
  ];

  function periodoPadrao() {
    const reais = periodos.semanas.filter((s) => s.temReal);
    const sem = (reais[reais.length - 1] || periodos.semanas[0]);
    const mes = E.MESES[sem.inicio ? Math.min(11, new Date(sem.inicio.getTime() + 3 * 864e5).getUTCMonth()) : 0];
    return { semana: sem.chave, mes };
  }
  const pad = periodoPadrao();
  const state = Object.assign({
    view: "painel", nivel: "Semanal", semana: pad.semana, mes: pad.mes, fonte: "real", regras: "corrigidas",
    intercambiavel: [], clusterFiltro: "todos", metrica: "ocupCorrigida", mostrarTodas: false,
  }, store.get("estado", {}));
  let cenario = Object.assign(E.cenarioVazio(), store.get("cenario", {}));
  if (!periodos.semanas.some((s) => s.chave === state.semana)) state.semana = pad.semana;
  const hashView = (location.hash || "").slice(1);
  if (VIEWS.some((v) => v[0] === hashView)) state.view = hashView;

  const salvarEstado = () => store.set("estado", { nivel: state.nivel, semana: state.semana, mes: state.mes, fonte: state.fonte,
    regras: state.regras, intercambiavel: state.intercambiavel, clusterFiltro: state.clusterFiltro, metrica: state.metrica });

  // ── cálculo com cache ───────────────────────────────────────────────────
  const regras = () => (state.regras === "legado" ? E.REGRAS_LEGADO : E.REGRAS_CORRIGIDAS);
  function periodoObj(nivel = state.nivel, id) {
    if (nivel === "Semanal") {
      const s = periodos.semanas.find((x) => x.chave === (id ?? state.semana));
      return { ...s, fonte: state.fonte === "projecao" ? "projecao" : "real" };
    }
    return periodos.meses.find((m) => m.mes === (id ?? state.mes));
  }
  let cache = new Map();
  function calc(p = periodoObj(), cen = null) {
    const k = [p.nivel, p.id, p.fonte, state.regras, cen ? JSON.stringify(cen) : ""].join("|");
    if (!cache.has(k)) cache.set(k, E.calcular(ctx, p, { regras: regras(), cenario: cen || undefined }));
    return cache.get(k);
  }
  function listaNivel() {
    return state.nivel === "Semanal" ? periodos.semanas.map((s) => periodoObj("Semanal", s.chave)) : periodos.meses;
  }
  const idx = (p) => listaNivel().findIndex((x) => x.id === p.id);
  const anteriores = (n) => { const l = listaNivel(); const i = idx(periodoObj()); return l.slice(Math.max(0, i - n), i); };
  const cenarioAtivo = () => JSON.stringify(cenario) !== JSON.stringify(E.cenarioVazio());

  // ── textos do período ───────────────────────────────────────────────────
  function rotuloPeriodo(p = periodoObj()) {
    if (p.nivel === "Semanal") return `Semana ${String(p.semana).padStart(2, "0")} · ${E.ANO}`;
    return p.mes.replace("/", " · ");
  }
  function intervaloPeriodo(p = periodoObj()) {
    if (p.nivel === "Semanal") return `${dataBR(p.inicio)} a ${dataBR(p.fim)}`;
    return p.mes;
  }
  const fonteTxt = (r) => (r.fonte === "real" ? "quarentena real" : r.fonte === "projecao" ? "projeção do plano mensal" : "plano mensal");

  // ── shell ───────────────────────────────────────────────────────────────
  function montarNav() {
    $("#nav").innerHTML = VIEWS.map(([id, nome, ic]) =>
      `<button type="button" data-view="${id}" ${state.view === id ? 'aria-current="page"' : ""}>
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ic}</svg>${nome}</button>`).join("");
  }
  function montarSeletor() {
    const sel = $("#period-select");
    if (state.nivel === "Semanal") {
      sel.innerHTML = periodos.semanas.map((s) => `<option value="${s.chave}" ${s.chave === state.semana ? "selected" : ""}>S${String(s.semana).padStart(2, "0")} · ${dataBR(s.inicio)}–${dataBR(s.fim)}${s.temReal ? " · real" : ""}</option>`).join("");
    } else {
      sel.innerHTML = periodos.meses.map((m) => `<option value="${m.mes}" ${m.mes === state.mes ? "selected" : ""}>${m.mes}</option>`).join("");
    }
    $$(".seg [data-nivel]").forEach((b) => b.setAttribute("aria-pressed", String(b.dataset.nivel === state.nivel)));
    const ult = store.get("historico", [])[0];
    $("#last-import").textContent = ult ? `${ult.arquivo.slice(0, 26)} · ${ult.data}` : `${dataset.meta?.gerado_em || "—"} · base inicial`;
  }

  function renderSignal(r) {
    const k = r.kpi;
    const criticas = r.mo.filter((l) => l.CARREGAMENTO > 1).sort((a, b) => b.GAP_HC - a.GAP_HC).slice(0, 3).map((l) => l.FAMILIA);
    const acao = k.severidade === "danger" ? `Ação: reforçar ou remanejar HC para ${criticas.join(", ") || "as famílias críticas"}.`
      : k.severidade === "warning" ? "Ação: acompanhar diariamente e evitar novas ausências." : "Capacidade confortável para o período.";
    const flags = [];
    if (r.fonte === "projecao") flags.push('<span class="scenario-flag">Demanda projetada</span>');
    if (state.regras === "legado") flags.push('<span class="scenario-flag">Regras legado</span>');
    $("#signal").innerHTML = `<div class="signal-row">
      <span class="dot lg halo s-${k.severidade}" aria-hidden="true"></span>
      <div class="signal-text">
        <div class="signal-verdict"><span>${esc(k.veredito)}</span><span class="micro">Semáforo · ${SEV_TXT[k.severidade]}</span>${flags.join("")}</div>
        <div class="signal-sub">Ocupação corrigida de ${pct(k.ocupCorrigida)} descontando ${num(r.ausencias.total)} HC de ausência (${r.ausencias.origem}). ${esc(acao)}</div>
      </div>
      <div class="signal-nums">
        <div><div class="micro">Ocupação</div><div class="v num s-${k.severidade}">${pct(k.ocupCorrigida)}</div></div>
        <div><div class="micro">Gap HC</div><div class="v num">${sinal(k.gapHCEfetivo, 0)}</div></div>
      </div></div>
      <div class="signal-strip" aria-hidden="true"><span></span><span></span><span></span><span></span></div>`;
  }

  // ── componentes ─────────────────────────────────────────────────────────
  function card(titulo, sub, corpo, extra = "", cls = "") {
    return `<section class="card ${cls}"><div class="card-h"><div><h2>${titulo}</h2>${sub ? `<p class="micro">${sub}</p>` : ""}</div>${extra}</div><div class="card-b">${corpo}</div></section>`;
  }
  const legenda = `<div class="legend"><span>${dot("danger")}&gt;100%</span><span>${dot("warning")}80–100%</span><span>${dot("success")}&lt;80%</span></div>`;
  function barras(linhas, { max, valor = (l) => l.CARREGAMENTO, rotulo, cor = (l) => sev(valor(l)) } = {}) {
    const escala = max || Math.max(1.2, ...linhas.map((l) => (Number.isFinite(valor(l)) ? valor(l) : 0)));
    return `<div class="bars">${linhas.map((l, i) => {
      const v = valor(l), w = Math.max(0, Math.min(100, ((Number.isFinite(v) ? v : escala) / escala) * 100));
      return `<div class="bar-row" title="${esc(l.FAMILIA)}: ${esc(rotulo ? rotulo(l).replace(/<[^>]+>/g, "") : pct(v))}">
        <div class="bar-lbl"><span class="name">${esc(l.FAMILIA)}</span><span class="val">${rotulo ? rotulo(l) : pct(v)}</span></div>
        <div class="track"><div class="fill s-${cor(l)}" style="width:${w}%;animation-delay:${i * 30}ms"></div>
        ${escala >= 1 && !max ? `<div class="mark100" style="left:${(1 / escala) * 100}%"></div>` : ""}</div></div>`;
    }).join("")}</div>`;
  }
  const semDados = (txt) => `<p class="muted">${txt}</p>`;

  // ── insights e resumo ───────────────────────────────────────────────────
  function oportunidades(r) {
    const inter = new Set(state.intercambiavel);
    const doadoras = r.mo.filter((l) => inter.has(l.FAMILIA) && l.CARREGAMENTO < 0.8).map((l) => ({ ...l, folga: l.RECURSO - l.HC_NECESSARIO }));
    const receptoras = r.mo.filter((l) => l.CARREGAMENTO > 1).sort((a, b) => b.GAP_HC - a.GAP_HC);
    return { doadoras, receptoras };
  }
  function cronico(nivelLista, i) {
    // família de MO acima de 100% por mais períodos seguidos até o período selecionado
    const seq = new Map();
    for (let j = i; j >= 0; j--) {
      const r = calc(nivelLista[j]);
      const acima = new Set(r.mo.filter((l) => l.CARREGAMENTO > 1).map((l) => l.FAMILIA));
      if (j === i) for (const f of acima) seq.set(f, 1);
      else { let algum = false; for (const [f, n] of seq) if (n === i - j && acima.has(f)) { seq.set(f, n + 1); algum = true; } if (!algum) break; }
    }
    let melhor = null;
    for (const [f, n] of seq) if (!melhor || n > melhor.n) melhor = { familia: f, n };
    return melhor;
  }
  function insights(r) {
    const out = [];
    const k = r.kpi;
    const criticas = r.mo.filter((l) => l.CARREGAMENTO > 1).sort((a, b) => b.GAP_HC - a.GAP_HC);
    if (criticas.length) {
      const c = criticas[0];
      out.push(["danger", `${criticas.length} família(s) acima de 100%. A maior falta é ${c.FAMILIA}: ${pct(c.CARREGAMENTO)} de carregamento, faltam ${num(c.GAP_HC)} HC por turno.`, "mapa"]);
    }
    const { doadoras, receptoras } = oportunidades(r);
    if (!state.intercambiavel.length) out.push(["neutral", "Nenhuma família marcada como intercambiável. Marque-as no Mapa para ver oportunidades de remanejamento.", "mapa"]);
    else if (doadoras.length && receptoras.length) {
      const d = doadoras.sort((a, b) => b.folga - a.folga)[0], rc = receptoras[0];
      out.push(["success", `${d.FAMILIA} tem folga de ${num(d.folga)} HC/turno e pode reforçar ${rc.FAMILIA} (${num(Math.min(d.folga, rc.GAP_HC))} HC).`, "gargalos"]);
    }
    const porCluster = new Map();
    for (const l of r.mo) porCluster.set(l.CLUSTER, (porCluster.get(l.CLUSTER) || 0) + l.CM_HRS);
    const totCM = [...porCluster.values()].reduce((a, b) => a + b, 0);
    const topC = [...porCluster.entries()].sort((a, b) => b[1] - a[1])[0];
    if (topC && totCM) out.push(["neutral", `${topC[0]} concentra ${pct(topC[1] / totCM, 0)} das horas de mão de obra (${num(topC[1], 0)} h).`, "gargalos"]);
    if (r.ausencias.origem === "semana" && r.ausencias.mediaAno) {
      const s = r.ausencias.total > r.ausencias.mediaAno ? "warning" : "success";
      out.push([s, `Ausências de ${num(r.ausencias.total)} HC na semana, contra ${num(r.ausencias.mediaAno)} HC na média do ano.`, "mao"]);
    }
    const ant = anteriores(4).map((p) => calc(p).complexidade.pct34).filter((v) => v !== null);
    if (r.complexidade.pct34 !== null && ant.length) {
      const m = ant.reduce((a, b) => a + b, 0) / ant.length, d = r.complexidade.pct34 - m;
      out.push([d > 0.05 ? "warning" : "neutral", `Complexidade 3+4 em ${pct(r.complexidade.pct34)}, ${sinal(d * 100)} p.p. contra a média dos ${ant.length} períodos anteriores.`, "complexidade"]);
    }
    const cmax = capMax(periodoObj());
    if (cmax && cmax.fator !== null && r.dias) {
      const sust = cmax.lotes / r.dias;
      out.push([k.lotesDia > sust ? "danger" : "success", `Ritmo necessário de ${num(k.lotesDia)} lotes/dia. Com o quadro atual, sem remanejar, ${cmax.gargalo} limita o laboratório a ${num(sust)} lotes/dia.`, "simulador"]);
    }
    const lista = listaNivel(), i = idx(periodoObj());
    const cr = cronico(lista, i);
    if (cr && cr.n >= 3) out.push(["danger", `Gargalo crônico: ${cr.familia} está acima de 100% há ${cr.n} ${state.nivel === "Semanal" ? "semanas seguidas" : "meses seguidos"}.`, "tendencias"]);
    const t = r.porTurno, medias = t.map((_, j) => (t.reduce((a, b) => a + b, 0) - t[j]) / 2);
    const j = t.findIndex((v, jj) => v < medias[jj] * 0.9);
    if (j >= 0) out.push(["warning", `${j + 1}º turno com ${num(t[j])} HC, ${num(medias[j] - t[j])} abaixo da média dos outros turnos.`, "mao"]);
    if (k.equipSobrecarga) {
      const pe = r.eq.slice().sort((a, b) => b.CARREGAMENTO - a.CARREGAMENTO)[0];
      out.push(["warning", `${k.equipSobrecarga} equipamento(s) acima de 100%; o mais carregado é ${pe.FAMILIA} (${pct(pe.CARREGAMENTO)}).`, "mapa"]);
    }
    const ordem = { danger: 0, warning: 1, success: 2, neutral: 3 };
    return out.sort((a, b) => ordem[a[0]] - ordem[b[0]]).slice(0, 6);
  }
  const capCache = new Map();
  function capMax(p, cen = null) {
    const k = [p.nivel, p.id, p.fonte, state.regras, cen ? JSON.stringify(cen) : ""].join("|");
    if (!capCache.has(k)) capCache.set(k, E.capacidadeMaxima(ctx, p, { regras: regras(), cenario: cen || undefined }));
    return capCache.get(k);
  }
  function resumoTexto(r) {
    const k = r.kpi, p = r.periodo;
    const criticas = r.mo.filter((l) => l.CARREGAMENTO > 1).sort((a, b) => b.CARREGAMENTO - a.CARREGAMENTO).slice(0, 3);
    const cx = r.complexidade;
    return `${rotuloPeriodo(p)} (${intervaloPeriodo(p)}): ${k.veredito.toLowerCase()}. ` +
      `São ${num(k.lotes, 0)} lotes (${fonteTxt(r)})${cx.pct34 !== null ? `, ${pct(cx.pct34)} de complexidade 3+4` : ""}. ` +
      `A mão de obra precisa de ${num(k.hcNecTotal, 0)} HC nos ${k.nTurnos} turnos e tem ${num(k.hcRealTotal, 0)} HC (${num(k.hcEfetivo, 0)} efetivos após ausências). ` +
      `Ocupação global de ${pct(k.ocupGlobal)} e corrigida de ${pct(k.ocupCorrigida)}. ` +
      (criticas.length ? `Famílias mais críticas: ${criticas.map((l) => `${l.FAMILIA} (${pct(l.CARREGAMENTO, 0)})`).join(", ")}. ` : "Nenhuma família acima de 100%. ") +
      `É preciso liberar ${num(k.lotesDia)} lotes por dia útil (${num(r.dias, 2)} dias úteis).`;
  }

  // ── TELA 1: painel ──────────────────────────────────────────────────────
  function viewPainel(r) {
    const k = r.kpi, cx = r.complexidade;
    const kpis = [
      ["Ocupação global", pct(k.ocupGlobal), sev(k.ocupGlobal), k.ocupGlobal > 1 ? "acima da capacidade" : "da capacidade"],
      ["Ocupação corrigida", pct(k.ocupCorrigida), sev(k.ocupCorrigida), "descontando ausências"],
      ["HC necessário", num(k.hcNecTotal, 0), sev(k.ocupGlobal), `disponível ${num(k.hcRealTotal, 0)} HC`],
      ["Gap de HC", sinal(k.gapHC, 0), k.gapHC > 0 ? "danger" : "success", k.gapHC > 0 ? "déficit estimado" : "folga estimada"],
      ["Total de lotes", num(k.lotes, 0), "neutral", r.fonte === "real" ? "programáveis na quarentena" : "programáveis no plano"],
      ["Lotes/dia", num(k.lotesDia), "neutral", `${num(r.dias, 2)} dias úteis`],
      ["Ausência", num(r.ausencias.total), r.ausencias.total > r.ausencias.mediaAno ? "warning" : "neutral", `média anual ${num(r.ausencias.mediaAno)}`],
      ["Complexidade 3+4", pct(cx.pct34), cx.pct34 > 0.7 ? "warning" : "neutral", `${num(cx.lotes34, 0)} lotes`],
    ];
    const ordenadas = r.mo.slice().sort((a, b) => b.CARREGAMENTO - a.CARREGAMENTO);
    const visiveis = state.mostrarTodas ? ordenadas : ordenadas.slice(0, 8);
    return `
      <div class="kpis">${kpis.map(([t, v, s, d]) => `<section class="card kpi"><div class="kpi-top"><span class="micro">${t}</span>${dot(s)}</div>
        <div class="kpi-v">${v}</div><div class="kpi-d">${d}</div></section>`).join("")}</div>
      <div class="grid-3">
        ${card("Carregamento por família", `Mão de obra · ${esc(rotuloPeriodo())} · ordenado do maior para o menor`,
          barras(visiveis, { rotulo: (l) => `${num(l.CM_HRS, 0)} h · <b>${pct(l.CARREGAMENTO)}</b>` }) +
          (ordenadas.length > 8 ? `<p><button class="link-btn" data-acao="todas">${state.mostrarTodas ? "Mostrar as 8 mais críticas" : `Mostrar todas as ${ordenadas.length} famílias`}</button></p>` : ""),
          legenda, "span-2")}
        <div class="stack">
          ${card("Resumo executivo", "gerado a partir do período", `<p class="summary" id="resumo">${esc(resumoTexto(r))}</p>`,
            `<button class="btn btn-sm no-print" data-acao="copiar-resumo">Copiar</button>`)}
          ${card("Insights acionáveis", "calculados a partir dos dados",
            `<ul class="insights">${insights(r).map(([s, t, v]) => `<li>${dot(s)}<div>${esc(t)}<button class="link-btn no-print" data-go="${v}">Ver detalhe</button></div></li>`).join("")}</ul>`)}
        </div>
      </div>`;
  }

  // ── TELA 2: mapa ────────────────────────────────────────────────────────
  const GRUPOS = [["todos", "Todos"], ["BANCADA", "Bancada"], ["EQUIPAMENTO", "Equipamento"], ["MO EQUIPAMENTO", "MO equipamento"], ["REVISÃO FINAL", "Revisão final"], ["HPLC", "HPLC / CG"]];
  const grupoDe = (l) => (l.PROCESSO.startsWith("EQUIPAMENTO ") ? "HPLC" : l.CLUSTER || "—");
  function linhasMapa(r) {
    return r.linhas.filter((l) => state.clusterFiltro === "todos" || grupoDe(l) === state.clusterFiltro);
  }
  function viewMapa(r) {
    const inter = new Set(state.intercambiavel);
    const linhas = linhasMapa(r);
    const grupos = new Map();
    for (const l of linhas) { const g = grupoDe(l); if (!grupos.has(g)) grupos.set(g, []); grupos.get(g).push(l); }
    const nomeGrupo = (g) => (g === "HPLC" ? "Equipamentos HPLC / CG (tempo de máquina)" : g);
    const corpo = [...grupos.entries()].map(([g, ls]) => `<tr class="group"><td colspan="12">${esc(nomeGrupo(g))} · ${ls.length}</td></tr>` +
      ls.sort((a, b) => b.CARREGAMENTO - a.CARREGAMENTO).map((l) => `<tr>
        <td class="name">${esc(l.FAMILIA)}${l.PROCESSO !== l.FAMILIA && !l.PROCESSO.startsWith("EQUIP") ? `<div class="micro">${esc(l.PROCESSO)}</div>` : ""}</td>
        <td class="l">${esc(l.CLUSTER || "—")}</td><td>${num(l.CM_HRS, 1)}</td><td>${num(l.RECURSO, 2)}</td><td>${num(l.HC_NECESSARIO, 2)}</td>
        <td><span class="pill">${dot(sev(l.CARREGAMENTO))}<span class="s-${sev(l.CARREGAMENTO)}">${pct(l.CARREGAMENTO)}</span></span></td>
        <td>${num(l.TC_MEDIO, 2)}</td><td>${num(l.TAKT, 2)}</td><td>${sinal(l.GAP_HC, 2)}</td><td>${num(l.DEMANDA, 1)}</td><td>${num(l.DEMANDA_MAX, 1)}</td>
        <td class="l">${l.CATEGORIA === "Mão de Obra" ? `<label class="toggle"><input type="checkbox" data-inter="${esc(l.FAMILIA)}" ${inter.has(l.FAMILIA) ? "checked" : ""}>${inter.has(l.FAMILIA) ? "Sim" : "Não"}</label>` : '<span class="muted">—</span>'}</td>
      </tr>`).join("")).join("");
    const semRec = r.semRecurso.length ? `<div class="notice warn">${ICON.alert}<div>Carga sem recurso cadastrado na aba FAMÍLIAS (fica fora da ocupação): ${r.semRecurso.map((s) => `${esc(s.FAMILIA)} (${num(s.CM_HRS, 0)} h)`).join(", ")}.</div></div>` : "";
    return `${semRec}${card("Detalhamento por família", `Cálculos de ${esc(rotuloPeriodo())} · ${fonteTxt(r)} · tempos em horas, demanda em lotes`,
      `<div class="table-wrap"><table class="data"><thead><tr><th>Família</th><th>Cluster</th><th>CM (h)</th><th>Recurso</th><th>HC necessário</th><th>Carregamento</th><th>TC médio (h)</th><th>Takt (h)</th><th>Gap HC</th><th>Demanda</th><th>Demanda máx.</th><th>Intercambiável</th></tr></thead>
      <tbody>${corpo || `<tr><td colspan="12">Nenhuma linha neste filtro.</td></tr>`}</tbody></table></div>
      <p class="micro" style="margin-top:10px">Recurso e HC por turno. Demanda máx. = lotes que o recurso atende com carregamento 100%. Nos equipamentos, "recurso" é o nº de equipamentos.</p>`,
      `<div class="filters" role="group" aria-label="Filtrar por cluster">${GRUPOS.map(([id, n]) => `<button class="chip" data-cluster="${id}" aria-pressed="${state.clusterFiltro === id}">${n}</button>`).join("")}</div>`)}`;
  }

  // ── TELA 3: gargalos ────────────────────────────────────────────────────
  function viewGargalos(r) {
    const { doadoras, receptoras } = oportunidades(r);
    let matriz;
    if (!state.intercambiavel.length) matriz = semDados("Marque as famílias intercambiáveis no Mapa de capacidade para montar a matriz.");
    else if (!doadoras.length || !receptoras.length) matriz = semDados(!receptoras.length ? "Nenhuma família acima de 100% neste período." : "Nenhuma família intercambiável abaixo de 80% neste período.");
    else matriz = `<div class="table-wrap"><table class="matrix"><thead><tr><th class="l">Doadora ↓ · Receptora →</th>${receptoras.map((x) => `<th>${esc(x.FAMILIA)}<div class="micro">falta ${num(x.GAP_HC)}</div></th>`).join("")}</tr></thead>
      <tbody>${doadoras.map((d) => `<tr><th class="l">${esc(d.FAMILIA)}<div class="micro">folga ${num(d.folga)}</div></th>${receptoras.map((x) => { const v = Math.min(d.folga, x.GAP_HC); return `<td class="${v > 0 ? "pos" : ""}">${v > 0 ? num(v) : "—"}</td>`; }).join("")}</tr>`).join("")}</tbody></table></div>
      <p class="micro" style="margin-top:8px">HC por turno potencialmente remanejável: limitado pela folga da doadora e pela falta da receptora.</p>`;
    const gaps = r.mo.filter((l) => l.GAP_HC > 0).sort((a, b) => b.GAP_HC - a.GAP_HC);
    const maxGap = Math.max(1, ...gaps.map((l) => l.GAP_HC));
    const ranking = gaps.length ? barras(gaps, { max: maxGap, valor: (l) => l.GAP_HC, cor: (l) => sev(l.CARREGAMENTO), rotulo: (l) => `<b>${sinal(l.GAP_HC)}</b> HC · ${pct(l.CARREGAMENTO, 0)}` }) : semDados("Nenhuma família com falta de HC.");
    const porCM = r.mo.slice().sort((a, b) => b.CM_HRS - a.CM_HRS);
    const tot = porCM.reduce((a, l) => a + l.CM_HRS, 0);
    let acc = 0, n80 = 0;
    for (const l of porCM) { if (acc / tot < 0.8) n80++; acc += l.CM_HRS; }
    const acc80 = porCM.slice(0, n80).reduce((a, l) => a + l.CM_HRS, 0) / tot;
    const pareto = `<p>${n80} de ${porCM.length} famílias concentram ${pct(acc80)} da carga de mão de obra (${num(tot, 0)} h).</p>
      <div class="pareto" role="img" aria-label="Pareto da carga">${porCM.map((l, i) => `<span class="${i < n80 ? "top" : "rest"}" style="flex:${l.CM_HRS}" title="${esc(l.FAMILIA)}: ${num(l.CM_HRS, 0)} h (${pct(l.CM_HRS / tot)})"></span>`).join("")}</div>
      <div class="legend" style="margin-top:8px">${porCM.slice(0, n80).map((l) => `<span>${esc(l.FAMILIA)} ${pct(l.CM_HRS / tot, 0)}</span>`).join("")}</div>`;
    const eqs = r.eq.filter((l) => l.CARREGAMENTO > 1).sort((a, b) => b.CARREGAMENTO - a.CARREGAMENTO);
    return `<div class="grid-2">
      ${card("Matriz doadoras × receptoras", "HC por turno · famílias intercambiáveis", matriz)}
      ${card("Ranking por gap de HC", "falta de HC por turno", ranking)}
    </div>
    ${card("Pareto da carga", "mão de obra · horas do período", pareto)}
    ${card("Equipamentos acima de 100%", `tempo de máquina · OEE real de cada equipamento`, eqs.length ? barras(eqs, { rotulo: (l) => `${num(l.CM_HRS, 0)} h · <b>${pct(l.CARREGAMENTO)}</b> · OEE ${pct(l.OEE, 0)}` }) : semDados("Nenhum equipamento acima de 100%."), legenda)}`;
  }

  // ── TELA 4: mão de obra ─────────────────────────────────────────────────
  function viewMao(r) {
    const t = r.porTurno, maxT = Math.max(...t, 1);
    const medias = t.map((_, j) => (t.reduce((a, b) => a + b, 0) - t[j]) / 2);
    const j = t.findIndex((v, jj) => v < medias[jj] * 0.9);
    const turnos = `<div class="bars">${t.map((v, i) => `<div class="bar-row"><div class="bar-lbl"><span class="name">${i + 1}º turno</span><span class="val"><b>${num(v)}</b> HC</span></div>
      <div class="track"><div class="fill s-primary" style="width:${(v / maxT) * 100}%"></div></div></div>`).join("")}</div>
      <p class="micro" style="margin-top:12px">Balanceamento: menor turno com ${pct(Math.min(...t) / maxT, 0)} do maior.</p>
      ${j >= 0 ? `<div class="notice warn">${ICON.alert}<div>${j + 1}º turno está ${num(medias[j] - t[j])} HC abaixo da média dos outros turnos.</div></div>` : ""}`;
    const a = r.ausencias;
    const aus = `<div class="kv"><div><span class="micro">Tipo</span><span class="micro">${a.origem === "semana" ? `Semana ${a.semanaRef}` : "Período"} · média ano</span></div>
      ${a.tipos.map((x) => `<div><span>${esc(x.tipo)}</span><span class="num">${num(x.valor)} · <span class="muted">${num(x.ano)}</span></span></div>`).join("")}
      <div><span><b>Total de ausências</b></span><span class="num"><b>${num(a.total)}</b> · <span class="muted">${num(a.mediaAno)}</span></span></div></div>
      <div class="effective"><span>HC efetivo</span><span class="num">${num(r.kpi.hcEfetivo)}</span></div>
      <p class="micro" style="margin-top:8px">HC real ${num(r.kpi.hcRealTotal)} menos ausências. ${a.origem === "semana" ? "Valores da semana informada na aba FAMÍLIAS." : "Sem apontamento para este período: usa a média do ano."}</p>`;
    const fams = r.mo.slice().sort((x, y) => (x.PROCESSO + x.FAMILIA).localeCompare(y.PROCESSO + y.FAMILIA));
    const tabela = `<div class="table-wrap"><table class="data" style="min-width:680px"><thead><tr><th>Família</th><th>1º turno</th><th>2º turno</th><th>3º turno</th><th>Média/turno</th><th>HC necessário</th><th>Gap HC</th></tr></thead><tbody>
      ${fams.map((l) => `<tr><td class="name">${esc(l.FAMILIA)}</td>${l.QTDE.map((q) => `<td>${num(q, 2)}</td>`).join("")}<td>${num(l.RECURSO, 2)}</td><td>${num(l.HC_NECESSARIO, 2)}</td><td class="s-${l.GAP_HC > 0 ? "danger" : "success"}">${sinal(l.GAP_HC, 2)}</td></tr>`).join("")}</tbody></table></div>`;
    return `<div class="grid-2">${card("HC por turno", "mão de obra · soma das famílias", turnos)}${card("Ausências", "HC fora do quadro no período", aus)}</div>
      ${card("Quadro por família e turno", "HC por turno", tabela)}`;
  }

  // ── TELA 5: complexidade ────────────────────────────────────────────────
  function colunas(itens, { alturaMax, rot = (x) => x.rotulo, val = (x) => x.valor, fmt = (v) => num(v, 0), cor = () => "primary", sel = () => false, pior = () => false, ref = null, titulo = (x) => "" }) {
    const mx = alturaMax || Math.max(...itens.map((x) => (Number.isFinite(val(x)) ? val(x) : 0)), ref || 0, 1e-9);
    return `<div class="cols-wrap"><div class="cols" role="img">${itens.map((x, i) => {
      const v = val(x), h = Number.isFinite(v) ? Math.max(0, (v / mx) * 170) : 0;
      return `<button type="button" class="col ${sel(x) ? "sel" : ""} ${pior(x) ? "worst" : ""}" data-col="${i}" title="${esc(titulo(x) || `${rot(x)}: ${fmt(v)}`)}">
        <span class="cv">${fmt(v)}</span><span class="bar s-${cor(x)}" style="height:${h}px;animation-delay:${i * 12}ms"></span><span class="cl">${esc(rot(x))}</span></button>`;
    }).join("")}${ref !== null ? `<div class="ref100" style="bottom:${(ref / mx) * 170 + 22}px"><span>100%</span></div>` : ""}</div></div>`;
  }
  function viewComplexidade(r) {
    const cx = r.complexidade;
    const ant = anteriores(4).map((p) => calc(p).complexidade.pct34).filter((v) => v !== null);
    const m4 = ant.length ? ant.reduce((a, b) => a + b, 0) / ant.length : null;
    const itens = [1, 2, 3, 4].map((n) => ({ rotulo: `Nível ${n}`, valor: cx.cont[n], n }));
    if (cx.cont.sem > 0.5) itens.push({ rotulo: "Sem nível", valor: cx.cont.sem, n: 0 });
    const cor = (x) => (x.n === 0 ? "neutral" : x.n <= 2 ? "primary" : x.n === 3 ? "warning" : "danger");
    return `<div class="kpis" style="grid-template-columns:repeat(3,minmax(0,1fr))">
      <section class="card kpi"><span class="micro">Índice ponderado (1 a 4)</span><div class="kpi-v">${num(cx.indice, 2)}</div><div class="kpi-d">${num(cx.classificados, 0)} lotes classificados</div></section>
      <section class="card kpi"><span class="micro">Complexidade 3+4</span><div class="kpi-v">${pct(cx.pct34)}</div><div class="kpi-d">${num(cx.lotes34, 0)} lotes</div></section>
      <section class="card kpi"><span class="micro">Média dos ${ant.length || 4} períodos anteriores</span><div class="kpi-v">${pct(m4)}</div><div class="kpi-d">${m4 !== null && cx.pct34 !== null ? `${sinal((cx.pct34 - m4) * 100)} p.p. no período` : "sem histórico"}</div></section>
    </div>
    ${card("Lotes por nível de complexidade", `${esc(rotuloPeriodo())} · ${fonteTxt(r)}`, colunas(itens, { fmt: (v) => num(v, 0), cor }),
      `<div class="legend"><span>${dot("warning")}nível 3</span><span>${dot("danger")}nível 4</span></div>`)}
    ${r.fonte !== "real" ? `<div class="notice">${ICON.info}<div>No plano mensal a complexidade vem da coluna "complexidade" da aba DEMANDA, por produto. Na quarentena real ela vem de cada lote.</div></div>` : ""}`;
  }

  // ── TELA 6: tendências ──────────────────────────────────────────────────
  const METRICAS = [
    ["ocupCorrigida", "Ocupação corrigida", (t) => t.ocupCorrigida, (v) => pct(v, 0), true],
    ["ocupGlobal", "Ocupação global", (t) => t.ocupGlobal, (v) => pct(v, 0), true],
    ["gapHC", "Gap de HC", (t) => t.gapHC, (v) => sinal(v, 0), false],
    ["lotes", "Lotes", (t) => t.lotes, (v) => num(v, 0), false],
    ["pct34", "Complexidade 3+4", (t) => t.pct34, (v) => pct(v, 0), false],
    ["ausencias", "Ausências", (t) => t.ausencias, (v) => num(v, 1), false],
  ];
  function serie() {
    return listaNivel().map((p) => {
      const r = calc(p);
      const piorMO = r.mo.reduce((a, l) => (!a || l.CARREGAMENTO > a.CARREGAMENTO ? l : a), null);
      return { p, r, ocupCorrigida: r.kpi.ocupCorrigida, ocupGlobal: r.kpi.ocupGlobal, gapHC: r.kpi.gapHC, lotes: r.kpi.lotes,
        pct34: r.complexidade.pct34, indice: r.complexidade.indice, ausencias: r.ausencias.total, piorMO };
    });
  }
  function viewTendencias() {
    const s = serie();
    const [mid, mnome, get, fmt, ehOcup] = METRICAS.find((m) => m[0] === state.metrica) || METRICAS[0];
    const vals = s.map(get).filter((v) => Number.isFinite(v));
    const piorV = Math.max(...vals);
    const cur = s.findIndex((x) => x.p.id === periodoObj().id);
    const rot = (x) => (x.p.nivel === "Semanal" ? `S${x.p.semana}` : x.p.mes.slice(0, 3));
    const grafico = colunas(s, { rot, val: get, fmt, ref: ehOcup ? 1 : null,
      cor: (x) => (ehOcup ? sev(get(x)) : "primary"), sel: (x) => x.p.id === periodoObj().id, pior: (x) => get(x) === piorV,
      titulo: (x) => `${rotuloPeriodo(x.p)} · ${mnome}: ${fmt(get(x))}${x.r.fonte === "projecao" ? " (projeção)" : ""}` });
    const atual = s[cur], ant = s[cur - 1];
    const cr = cronico(listaNivel(), cur);
    const idxAnt = s.slice(Math.max(0, cur - 4), cur).map((x) => x.indice).filter((v) => v !== null);
    const mIdx = idxAnt.length ? idxAnt.reduce((a, b) => a + b, 0) / idxAnt.length : null;
    const piorP = s.find((x) => get(x) === piorV);
    // mapa de calor família × período
    const fams = [...new Set(s.flatMap((x) => x.r.mo.map((l) => l.FAMILIA)))];
    const heatCor = (v) => (!Number.isFinite(v) ? "var(--danger)" : v > 1 ? `color-mix(in srgb, var(--danger) ${Math.min(90, 35 + (v - 1) * 50)}%, var(--panel))`
      : v >= 0.8 ? "color-mix(in srgb, var(--warning) 40%, var(--panel))" : `color-mix(in srgb, var(--success) ${15 + v * 25}%, var(--panel))`);
    const heat = `<div class="table-wrap"><table class="heat"><thead><tr><th class="l">Família</th>${s.map((x) => `<th>${rot(x)}</th>`).join("")}<th>Pior</th></tr></thead><tbody>
      ${fams.map((f) => {
        const vs = s.map((x) => (x.r.mo.find((l) => l.FAMILIA === f) || {}).CARREGAMENTO);
        const mx = Math.max(...vs.filter((v) => v !== undefined));
        const iw = vs.indexOf(mx);
        return `<tr><td class="name">${esc(f)}</td>${vs.map((v, i) => `<td class="${i === iw ? "worst" : ""}" style="background:${v === undefined ? "transparent" : heatCor(v)}" title="${esc(f)} · ${rotuloPeriodo(s[i].p)}: ${pct(v)}">${v === undefined ? "" : num(v * 100, 0)}</td>`).join("")}<td>${iw >= 0 ? rot(s[iw]) : "—"}</td></tr>`;
      }).join("")}</tbody></table></div><p class="micro" style="margin-top:8px">Carregamento em % por período, calculado período a período (sem média). Contorno = pior período da família.</p>`;
    return `${card(`Evolução · ${mnome}`, `${state.nivel === "Semanal" ? "semana a semana" : "mês a mês"} · cada período calculado separadamente · pior período em destaque: ${piorP ? rotuloPeriodo(piorP.p) : "—"}`, grafico,
        `<div class="filters" role="group" aria-label="Métrica">${METRICAS.map((m) => `<button class="chip" data-metrica="${m[0]}" aria-pressed="${m[0] === mid}">${m[1]}</button>`).join("")}</div>`)}
      <div class="kpis" style="grid-template-columns:repeat(3,minmax(0,1fr))">
        <section class="card kpi"><span class="micro">Gap de HC</span><div class="kpi-v">${sinal(atual.gapHC, 0)}</div><div class="kpi-d">${ant ? `${sinal(atual.gapHC - ant.gapHC, 1)} contra o período anterior` : "primeiro período"}</div></section>
        <section class="card kpi"><span class="micro">Gargalo crônico</span><div class="kpi-v" style="font-size:18px">${cr ? esc(cr.familia) : "—"}</div><div class="kpi-d">${cr ? `${cr.n} ${state.nivel === "Semanal" ? "semana(s) seguida(s)" : "mês(es) seguido(s)"} acima de 100%` : "nenhuma família acima de 100%"}</div></section>
        <section class="card kpi"><span class="micro">Índice do mix</span><div class="kpi-v">${num(atual.indice, 2)}</div><div class="kpi-d">${mIdx !== null ? `${sinal(atual.indice - mIdx, 2)} contra a média recente` : "sem histórico"}</div></section>
      </div>
      ${card("Carregamento por família e período", "mão de obra", heat)}
      ${state.nivel === "Semanal" && state.fonte !== "projecao" ? `<div class="notice">${ICON.info}<div>Só há quarentena real para ${periodos.semanas.filter((x) => x.temReal).map((x) => "S" + x.semana).join(", ")}. As demais semanas usam a projeção do plano mensal rateada pelos dias úteis.</div></div>` : ""}`;
  }

  // ── TELA 7: simulador ───────────────────────────────────────────────────
  function famMO() { return ctx.t.familias.filter((f) => f.CATEGORIA === "Mão de Obra").map((f) => f.FAMILIA); }
  function viewSimulador() {
    const base = calc();
    const fams = famMO();
    const oeeBase = E.toNum((ctx.t.familias.find((f) => f.CATEGORIA === "Mão de Obra") || {})["OEE REAL"]) || 0.8;
    const opt = (sel) => fams.map((f) => `<option ${f === sel ? "selected" : ""}>${esc(f)}</option>`).join("");
    const oeeMO = cenario.oeeMO ?? oeeBase;
    return `<div class="sim">
      ${card("Ajustes do cenário", "não altera os dados originais", `
        <div class="field"><label for="sim-oee">OEE da mão de obra</label>
          <div class="field-row"><input type="range" id="sim-oee" min="0.4" max="1" step="0.01" value="${oeeMO}"><span class="out" id="sim-oee-v">${pct(oeeMO, 0)}</span></div>
          <span class="micro">Base ${pct(oeeBase, 0)}. Reduza para ver o efeito de uma eficiência menor.</span></div>
        <div class="field"><label for="sim-oee-eq">OEE dos equipamentos</label>
          <div class="field-row"><input type="range" id="sim-oee-eq" min="0.5" max="1.5" step="0.05" value="${cenario.oeeEquip ?? 1}"><span class="out" id="sim-oee-eq-v">${pct(cenario.oeeEquip ?? 1, 0)}</span></div>
          <span class="micro">Percentual do OEE real de cada equipamento.</span></div>
        <div class="field"><label for="sim-dem">Demanda do período</label>
          <div class="field-row"><input type="range" id="sim-dem" min="0.5" max="1.5" step="0.01" value="${cenario.fatorDemanda}"><span class="out" id="sim-dem-v">${pct(cenario.fatorDemanda, 0)}</span></div></div>
        <div class="field"><span class="lbl">Demanda por família</span>
          <div class="field-row"><label class="sr-only" for="sim-fd-fam">Família</label><select id="sim-fd-fam">${[...new Set(ctx.tc.map((r) => r[E.C.FAM]).filter(Boolean))].sort().map((f) => `<option>${esc(f)}</option>`).join("")}</select>
          <label class="sr-only" for="sim-fd-v">Percentual</label><input type="number" id="sim-fd-v" min="0" max="300" step="5" value="120" style="width:74px"><span class="micro">%</span>
          <button class="btn btn-sm" data-acao="add-fd">Aplicar</button></div>
          <div class="moves">${Object.entries(cenario.fatorFamilia).map(([f, v]) => `<div class="move"><span>${esc(f)} · ${pct(v, 0)}</span><button class="btn btn-sm btn-ghost" data-del-fd="${esc(f)}" aria-label="Remover ajuste de ${esc(f)}">Remover</button></div>`).join("")}</div></div>
        <div class="field"><label for="sim-adiados">Lotes adiados para o período seguinte</label>
          <div class="field-row"><input type="number" id="sim-adiados" min="0" step="1" value="${cenario.lotesAdiados || 0}" style="width:110px"><span class="micro">de ${num(base.kpi.lotes, 0)} lotes</span></div></div>
        <div class="field"><label for="sim-aus">Ausências previstas (HC)</label>
          <div class="field-row"><input type="number" id="sim-aus" min="0" step="0.5" placeholder="${num(base.ausencias.total)}" value="${cenario.ausencias ?? ""}" style="width:110px"><span class="micro">vazio = ${num(base.ausencias.total)} (${base.ausencias.origem})</span></div></div>
        <div class="field"><span class="lbl">Remanejar HC entre famílias e turnos</span>
          <div class="field-row"><label class="sr-only" for="sim-de">De</label><select id="sim-de">${opt(fams[0])}</select><span class="micro">→</span><label class="sr-only" for="sim-para">Para</label><select id="sim-para">${opt(fams[1])}</select></div>
          <div class="field-row"><label class="sr-only" for="sim-turno">Turno</label><select id="sim-turno"><option value="todos">Todos os turnos</option><option value="0">1º turno</option><option value="1">2º turno</option><option value="2">3º turno</option></select>
          <label class="sr-only" for="sim-hc">HC</label><input type="number" id="sim-hc" min="0" step="0.5" value="1" style="width:74px"><span class="micro">HC</span><button class="btn btn-sm" data-acao="add-mv">Adicionar</button></div>
          <div class="moves">${cenario.remanejamentos.map((m, i) => `<div class="move"><span>${num(m.hc)} HC · ${esc(m.de)} → ${esc(m.para)} · ${m.turno === "todos" ? "todos os turnos" : `${Number(m.turno) + 1}º turno`}</span><button class="btn btn-sm btn-ghost" data-del-mv="${i}" aria-label="Remover remanejamento">Remover</button></div>`).join("")}</div></div>
        <div class="field"><span class="lbl">Quadro de uma família</span>
          <div class="field-row"><label class="sr-only" for="sim-q-fam">Família</label><select id="sim-q-fam">${opt()}</select></div>
          <div class="field-row" id="sim-q-inputs"></div>
          <div class="moves">${Object.entries(cenario.turnos).map(([f, q]) => `<div class="move"><span>${esc(f)} · ${q.map((x) => num(x, 1)).join(" / ")}</span><button class="btn btn-sm btn-ghost" data-del-q="${esc(f)}" aria-label="Remover ajuste de quadro">Remover</button></div>`).join("")}</div></div>
        <div class="field"><button class="btn" data-acao="limpar-cenario">Limpar cenário</button></div>`)}
      <div class="stack" id="sim-result"></div></div>`;
  }
  function renderSimResult() {
    const alvo = $("#sim-result"); if (!alvo) return;
    const base = calc(), sim = calc(periodoObj(), cenario);
    const cb = capMax(periodoObj()), cs = capMax(periodoObj(), cenario);
    const crit = (r) => r.mo.filter((l) => l.CARREGAMENTO > 1).sort((a, b) => b.CARREGAMENTO - a.CARREGAMENTO);
    const bloco = (t, r) => `<div><span class="micro">${t}</span><div class="big s-${r.kpi.severidade}">${pct(r.kpi.ocupCorrigida)}</div>
      <div class="pill">${dot(r.kpi.severidade)}<b>${esc(r.kpi.veredito)}</b></div><span class="micro">ocupação corrigida · global ${pct(r.kpi.ocupGlobal)}</span></div>`;
    const linha = (n, a, b, f, bomSeMenor = true) => {
      const d = b - a, s = Math.abs(d) < 1e-9 ? "neutral" : (d < 0) === bomSeMenor ? "success" : "danger";
      return `<div><span>${n}</span><span class="num muted">${f(a)}</span><span class="num s-${s}">${f(b)}</span></div>`;
    };
    const cB = crit(base), cS = crit(sim);
    alvo.innerHTML = card("Antes × depois", `${esc(rotuloPeriodo())} · recalculado em tempo real`, `
      <div class="compare">${bloco("Original", base)}${bloco("Simulado", sim)}</div>
      <div class="delta-list" style="margin-top:14px">
        <div><span class="micro">Indicador</span><span class="micro">Original</span><span class="micro">Simulado</span></div>
        ${linha("HC necessário", base.kpi.hcNecTotal, sim.kpi.hcNecTotal, (v) => num(v, 1))}
        ${linha("HC efetivo", base.kpi.hcEfetivo, sim.kpi.hcEfetivo, (v) => num(v, 1), false)}
        ${linha("Gap de HC (efetivo)", base.kpi.gapHCEfetivo, sim.kpi.gapHCEfetivo, (v) => sinal(v, 1))}
        ${linha("Lotes", base.kpi.lotes, sim.kpi.lotes, (v) => num(v, 0), false)}
        ${linha("Lotes/dia", base.kpi.lotesDia, sim.kpi.lotesDia, (v) => num(v, 1), false)}
        ${linha("Famílias acima de 100%", cB.length, cS.length, (v) => num(v, 0))}
        ${linha("Equipamentos acima de 100%", base.kpi.equipSobrecarga, sim.kpi.equipSobrecarga, (v) => num(v, 0))}
      </div>`) +
      card("Quanto consigo atender", "demanda máxima com carregamento ≤ 100% em todas as famílias de mão de obra", `
      <div class="compare">
        <div><span class="micro">Original</span><div class="big">${num(cb.lotes, 0)}</div><span class="micro">lotes · ${pct(cb.fator, 0)} da demanda · gargalo ${esc(cb.gargalo || "—")}</span></div>
        <div><span class="micro">Simulado</span><div class="big">${num(cs.lotes, 0)}</div><span class="micro">lotes · ${pct(cs.lotesAtuais ? cs.lotes / cs.lotesAtuais : null, 0)} da demanda simulada · gargalo ${esc(cs.gargalo || "—")}</span></div>
      </div>`) +
      card("Famílias críticas no cenário", "carregamento simulado", cS.length ? barras(cS.slice(0, 10), { rotulo: (l) => { const b = base.mo.find((x) => x.FAMILIA === l.FAMILIA); return `${b ? pct(b.CARREGAMENTO, 0) + " → " : ""}<b>${pct(l.CARREGAMENTO)}</b>`; } }) : semDados("Nenhuma família acima de 100% no cenário."), legenda) +
      `<p class="foot-note">${ICON.swap}Cenário isolado: os dados originais permanecem intactos.</p>`;
  }
  function quadroInputs() {
    const f = $("#sim-q-fam")?.value; const alvo = $("#sim-q-inputs"); if (!f || !alvo) return;
    const row = ctx.t.familias.find((x) => x.FAMILIA === f);
    const q = cenario.turnos[f] || ["QTDE 1° TURNO", "QTDE 2° TURNO", "QTDE 3° TURNO"].map((c) => E.toNum(row[c]));
    alvo.innerHTML = q.map((v, i) => `<label class="micro" for="sim-q${i}">${i + 1}º</label><input type="number" id="sim-q${i}" min="0" step="0.5" value="${Math.round(v * 100) / 100}" style="width:70px">`).join("") + '<button class="btn btn-sm" data-acao="add-q">Aplicar</button>';
  }
  let simTimer = null;
  function mudouCenario(recriarControles = false) {
    store.set("cenario", cenario);
    if (recriarControles) render(); else { clearTimeout(simTimer); simTimer = setTimeout(renderSimResult, 90); }
  }

  // ── TELA 8: importação ──────────────────────────────────────────────────
  let importacao = null; // { arquivo, resultados, blocos, dataset, confirmar }
  function alertasCadastro() {
    const out = [];
    const sem = new Map();
    for (const r of ctx.tc) {
      const temCorrida = E.toNum(r[E.C.ET12_N]) + E.toNum(r[E.C.ET13_N]) > 0;
      const eq = r[E.C.EQUIP];
      if (temCorrida && !(typeof eq === "string" && /^CR(LQ|GS)/.test(eq))) sem.set(r[E.C.MPR], (sem.get(r[E.C.MPR]) || 0) + 1);
    }
    const dAno = E.linhasDemanda(ctx, { nivel: "Mensal", meses: E.MESES });
    const comDem = new Set(dAno.linhas.filter((l) => l.qtde > 0).map((l) => l.mpr));
    const semEq = [...sem.keys()].filter((m) => comDem.has(m));
    if (semEq.length) out.push(`${semEq.length} MPR(s) com tempo de HPLC/CG e sem Equipamento Proposto (carga de máquina ignorada): ${semEq.join(", ")}.`);
    const mnf = dAno.linhas.filter((l) => l.semTempos).reduce((a, l) => a + l.qtde, 0);
    if (mnf) out.push(`${num(mnf, 0)} lotes no ano de códigos com "Método não encontrado" (sem MPR, fora do cálculo).`);
    const mprsTc = new Set(ctx.tc.map((r) => r[E.C.MPR]));
    const semTc = [...comDem].filter((m) => m && m !== "Método não encontrado" && !mprsTc.has(m));
    if (semTc.length) out.push(`${semTc.length} MPR(s) com demanda e sem tempos na BASE TC: ${semTc.slice(0, 12).join(", ")}${semTc.length > 12 ? "…" : ""}.`);
    const semCl = ctx.t.familias.filter((f) => !f.CLUSTER).map((f) => f.FAMILIA);
    if (semCl.length) out.push(`Recurso sem CLUSTER na aba FAMÍLIAS: ${semCl.join(", ")}.`);
    return out;
  }
  function viewImportacao() {
    const hist = store.get("historico", []);
    const imp = importacao;
    const lista = imp ? `<div class="status-list">${imp.resultados.map((x) => {
      const s = x.erros.length ? "danger" : x.avisos.length ? "warning" : "success";
      return `<div>${dot(s)}<div><b>${esc(x.rotulo)}</b><div class="micro">aba ${esc(x.aba)}</div></div><span class="num">${x.linhas ? num(x.linhas, 0) + " linhas" : "—"}</span>
        ${x.erros.length || x.avisos.length ? `<div class="errs">${x.erros.map((e) => `<span class="s-danger">${esc(e)}</span>`).join("")}${x.avisos.slice(0, 6).map((e) => `<span>${esc(e)}</span>`).join("")}${x.avisos.length > 6 ? `<span>… e mais ${x.avisos.length - 6} avisos.</span>` : ""}</div>` : ""}</div>`;
    }).join("")}
      <div>${dot(imp.blocos.ausencias ? "success" : "warning")}<div><b>Ausências e dias úteis</b><div class="micro">blocos da aba FAMÍLIAS</div></div><span class="num">${imp.blocos.ausencias ? imp.blocos.ausencias.rows.length + " tipos" : "—"}</span>
      ${imp.blocos.avisos.length ? `<div class="errs">${imp.blocos.avisos.map((e) => `<span>${esc(e)}</span>`).join("")}</div>` : ""}</div></div>` : "";
    const podeSalvar = imp && imp.dataset;
    const confirmar = imp && imp.confirmar ? `<div class="confirm"><span>Já existe uma carga salva (${esc(hist[0]?.arquivo || "")}, ${esc(hist[0]?.data || "")}). Substituir? A anterior fica no histórico.</span>
      <div class="field-row"><button class="btn btn-primary" data-acao="confirmar-carga">Substituir carga</button><button class="btn" data-acao="cancelar-carga">Cancelar</button></div></div>` : "";
    const calRows = state.nivel === "Semanal"
      ? periodos.semanas.filter((s) => Math.abs(s.chave - state.semana) <= 2 || s.chave === state.semana).map((s) => [`S${s.semana} · ${dataBR(s.inicio)}–${dataBR(s.fim)}`, E.diasUteis(ctx, s)])
      : periodos.meses.map((m) => [m.mes, E.diasUteis(ctx, m)]);
    const alertas = alertasCadastro();
    return `<div class="grid-3">
      <div class="stack span-2">
        ${card("Carregue a planilha MFV", "o app lê as tabelas direto do arquivo .xlsb/.xlsx", `
          <label class="drop" id="drop" for="arquivo">${ICON.upload}<b>Carregue as tabelas da planilha</b>
            <span class="muted">Arraste o arquivo aqui ou clique para escolher. Lidos: BASE TC, DEMANDA, quarentena semanal, FAMÍLIAS (recursos, calendário e ausências), CAMPANHA, SKIP e % de campanha.</span>
            <input type="file" id="arquivo" accept=".xlsb,.xlsx,.xlsm" class="sr-only"></label>
          ${imp ? `<p class="micro" style="margin-top:12px">Arquivo: ${esc(imp.arquivo)}</p>${lista}` : ""}
          ${imp && imp.lendo ? `<p class="muted">Lendo o arquivo…</p>` : ""}
          ${imp && imp.falha ? `<div class="notice warn">${ICON.alert}<div>${esc(imp.falha)}</div></div>` : ""}
          <div class="field-row" style="margin-top:14px"><button class="btn btn-primary" data-acao="salvar-carga" ${podeSalvar ? "" : "disabled"}>Validar e salvar carga</button>
          ${store.get("dataset", null) ? '<button class="btn" data-acao="restaurar">Voltar à base inicial</button>' : ""}</div>${confirmar}`)}
        ${card("Alertas de cadastro", "itens que ficam fora do cálculo", alertas.length ? `<ul class="insights">${alertas.map((a) => `<li>${dot("warning")}<div>${esc(a)}</div></li>`).join("")}</ul>` : semDados("Nenhum alerta."))}
        ${hist.length ? card("Histórico de cargas", "neste navegador", `<div class="kv">${hist.map((h) => `<div><span>${esc(h.arquivo)}</span><span class="num">${esc(h.data)} · ${num(h.linhas, 0)} linhas</span></div>`).join("")}</div>`) : ""}
      </div>
      <div class="stack">
        ${card("Parâmetros de cálculo", "valem para todas as telas", `
          <div class="field"><span class="lbl">Regras</span><div class="seg" role="group" aria-label="Regras de cálculo">
            <button type="button" data-regras="corrigidas" aria-pressed="${state.regras === "corrigidas"}">Corrigidas</button><button type="button" data-regras="legado" aria-pressed="${state.regras === "legado"}">Legado (Excel)</button></div>
            <span class="micro">Corrigidas: 1º lote de cada campanha com tempo cheio; skip só no teste de impureza de MPR da tabela; lotes/dia pelos dias úteis do calendário.</span></div>
          <div class="field"><span class="lbl">Demanda semanal</span><div class="seg" role="group" aria-label="Fonte da demanda semanal">
            <button type="button" data-fonte="real" aria-pressed="${state.fonte !== "projecao"}">Quarentena real</button><button type="button" data-fonte="projecao" aria-pressed="${state.fonte === "projecao"}">Projeção do plano</button></div>
            <span class="micro">Semanas sem quarentena usam a projeção do plano mensal rateada pelos dias úteis (seg ${num(ctx.pesoDia[0], 2)} · ter–sex 1 · sáb ${num(ctx.pesoDia[5], 2)}).</span></div>
          <div class="field"><span class="lbl">Eficiência (OEE)</span><span class="muted">Mão de obra ${pct(E.toNum((ctx.t.familias.find((f) => f.CATEGORIA === "Mão de Obra") || {})["OEE REAL"]), 0)} (fixo). Equipamentos: OEE real de cada um. Para testar outros valores use o Simulador.</span></div>`)}
        ${card("Calendário industrial", "dias úteis da aba FAMÍLIAS (feriados e férias já descontados)", `<div class="kv">${calRows.map(([n, d]) => `<div><span>${esc(n)}</span><span class="num">${num(d, 2)}</span></div>`).join("")}</div>
          <p class="micro" style="margin-top:8px">Para mudar feriados, edite o calendário na planilha e importe de novo.</p>`)}
      </div></div>`;
  }
  let xlsxPromise = null;
  function carregarXLSX() {
    if (window.XLSX) return Promise.resolve(window.XLSX);
    if (!xlsxPromise) xlsxPromise = new Promise((res, rej) => {
      const s = document.createElement("script");
      s.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
      s.onload = () => res(window.XLSX); s.onerror = () => { xlsxPromise = null; rej(new Error("Não foi possível carregar o leitor de planilhas. Verifique a conexão e tente de novo.")); };
      document.head.appendChild(s);
    });
    return xlsxPromise;
  }
  async function lerArquivo(file) {
    importacao = { arquivo: file.name, resultados: [], blocos: { avisos: [] }, lendo: true };
    render();
    try {
      const XLSX = await carregarXLSX();
      const buf = new Uint8Array(await file.arrayBuffer());
      const r = I.importarPlanilha(XLSX, buf, E.COLUNAS, file.name);
      importacao = { arquivo: file.name, ...r };
      if (!r.dataset) importacao.falha = "Há tabelas obrigatórias com erro. Corrija na planilha e carregue de novo; nada foi alterado.";
    } catch (e) {
      importacao = { arquivo: file.name, resultados: [], blocos: { avisos: [] }, falha: `Não foi possível ler o arquivo: ${e.message}` };
    }
    render();
  }
  function aplicarDataset(ds, arquivo) {
    const linhas = Object.values(ds.tabelas).reduce((a, t) => a + (t.rows ? t.rows.length : t.length), 0);
    dataset = ds; ctx = E.preparar(ds); periodos = E.listarPeriodos(ctx); cache = new Map(); capCache.clear();
    const salvo = store.set("dataset", ds);
    const hist = store.get("historico", []);
    hist.unshift({ arquivo, data: new Date().toLocaleString("pt-BR", { dateStyle: "short", timeStyle: "short" }), linhas });
    store.set("historico", hist.slice(0, 12));
    if (!periodos.semanas.some((s) => s.chave === state.semana)) state.semana = periodoPadrao().semana;
    importacao = null;
    toast(salvo ? "Carga salva. Todas as telas foram recalculadas." : "Carga aplicada nesta sessão. O navegador não permitiu salvar; ao recarregar volta a base anterior.");
    montarSeletor(); render();
  }

  // ── exportação ──────────────────────────────────────────────────────────
  function tabelaExport() {
    const r = calc();
    const nome = `flow-lab-${E.ANO}-${state.nivel === "Semanal" ? "S" + String(periodoObj().semana).padStart(2, "0") : periodoObj().mes.slice(0, 3)}`;
    const cols = [["Processo", "PROCESSO"], ["Família", "FAMILIA"], ["Cluster", "CLUSTER"], ["Categoria", "CATEGORIA"], ["CM (h)", "CM_HRS"], ["Recurso", "RECURSO"],
      ["HC necessário", "HC_NECESSARIO"], ["Carregamento", "CARREGAMENTO"], ["TC médio (h)", "TC_MEDIO"], ["Takt (h)", "TAKT"], ["Gap HC", "GAP_HC"],
      ["Demanda (lotes)", "DEMANDA"], ["Demanda máx. (lotes)", "DEMANDA_MAX"], ["Lotes/dia", "LOTES_DIA"], ["OEE", "OEE"], ["Horas turno", "HRS_TURNO"], ["Total recurso (h)", "TOTAL_RECURSO"]];
    const linhas = (state.view === "mapa" ? linhasMapa(r) : r.linhas).map((l) => cols.map(([, k]) => (typeof l[k] === "number" && !Number.isFinite(l[k]) ? null : l[k] ?? null)));
    return { nome, cabecalho: cols.map((c) => c[0]), linhas, r };
  }
  const podeImprimir = (() => { try { return window.self === window.top; } catch (e) { return false; } })();
  async function exportar(tipo) {
    const t = tabelaExport();
    if (tipo === "copiar") {
      const tsv = [t.cabecalho, ...t.linhas].map((l) => l.map((v) => (typeof v === "number" ? String(v).replace(".", ",") : v ?? "")).join("\t")).join("\n");
      try { await navigator.clipboard.writeText(tsv); toast("Tabela copiada. Cole no Excel."); } catch (e) { toast("O navegador bloqueou a cópia."); }
    } else if (tipo === "xlsx") {
      try {
        const XLSX = await carregarXLSX();
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([t.cabecalho, ...t.linhas]), "Capacidade");
        const k = t.r.kpi;
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["Indicador", "Valor"], ["Período", rotuloPeriodo()], ["Fonte da demanda", fonteTxt(t.r)],
          ["Ocupação global", k.ocupGlobal], ["Ocupação corrigida", k.ocupCorrigida], ["HC necessário", k.hcNecTotal], ["HC real", k.hcRealTotal],
          ["Ausências", t.r.ausencias.total], ["Gap HC", k.gapHC], ["Lotes", k.lotes], ["Lotes/dia", k.lotesDia], ["Veredito", k.veredito]]), "Resumo");
        XLSX.writeFile(wb, t.nome + ".xlsx");
        toast("Excel gerado. Se o download não começar, use Copiar tabela.");
      } catch (e) { toast(e.message); }
    } else if (tipo === "imprimir") {
      window.print();
    }
  }
  function menuExportar() {
    const ex = $("#export-menu");
    if (ex) { ex.remove(); return; }
    const m = document.createElement("div");
    m.id = "export-menu"; m.className = "card no-print";
    m.style.cssText = "position:absolute;right:20px;top:60px;z-index:40;padding:8px;display:grid;gap:4px;min-width:220px";
    m.innerHTML = `<button class="btn btn-ghost" data-exp="copiar">Copiar tabela (colar no Excel)</button>
      ${podeImprimir ? '<button class="btn btn-ghost" data-exp="xlsx">Baixar Excel (.xlsx)</button><button class="btn btn-ghost" data-exp="imprimir">Imprimir / salvar PDF</button>'
        : '<span class="micro" style="padding:6px 12px">Excel e PDF: abra o arquivo flow-lab.html direto no navegador.</span>'}`;
    $(".topbar").appendChild(m);
  }

  // ── render ──────────────────────────────────────────────────────────────
  const ICON = {
    alert: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M12 3l10 18H2z"/><path d="M12 10v5M12 18v.5"/></svg>',
    info: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><circle cx="12" cy="12" r="9"/><path d="M12 11v6M12 7.5v.5"/></svg>',
    swap: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M7 4L3 8l4 4M3 8h14M17 20l4-4-4-4M21 16H7"/></svg>',
    upload: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" aria-hidden="true"><path d="M12 16V4m0 0L8 8m4-4l4 4M5 20h14"/></svg>',
  };
  function toast(msg) {
    const t = $("#toast"); t.textContent = msg; t.hidden = false;
    clearTimeout(toast._t); toast._t = setTimeout(() => { t.hidden = true; }, 3200);
  }
  function render() {
    const r = calc();
    const nome = VIEWS.find((v) => v[0] === state.view)[1];
    $("#area-label").textContent = nome;
    $("#period-title").textContent = `${rotuloPeriodo()} · ${intervaloPeriodo()}${r.fonte === "projecao" ? " · projeção" : ""}`;
    document.title = `FLOW LAB · ${nome}`;
    renderSignal(r);
    const views = { painel: viewPainel, mapa: viewMapa, gargalos: viewGargalos, mao: viewMao, complexidade: viewComplexidade, tendencias: viewTendencias, simulador: viewSimulador, importacao: viewImportacao };
    $("#content").innerHTML = views[state.view](r);
    if (state.view === "simulador") { renderSimResult(); quadroInputs(); }
    montarNav();
  }
  function irPara(view) {
    state.view = view;
    try { history.replaceState(null, "", "#" + view); } catch (e) { /* sem histórico */ }
    $("#shell").classList.remove("nav-open");
    render(); window.scrollTo(0, 0);
  }

  // ── eventos ─────────────────────────────────────────────────────────────
  document.addEventListener("click", (ev) => {
    const b = ev.target.closest("button, [data-exp]");
    if (!b) { if (!ev.target.closest("#export-menu") && $("#export-menu") && !ev.target.closest("#btn-export")) $("#export-menu").remove(); return; }
    const d = b.dataset;
    if (d.view) return irPara(d.view);
    if (d.go) return irPara(d.go);
    if (d.nivel) { state.nivel = d.nivel; salvarEstado(); montarSeletor(); return render(); }
    if (d.cluster) { state.clusterFiltro = d.cluster; salvarEstado(); return render(); }
    if (d.metrica) { state.metrica = d.metrica; salvarEstado(); return render(); }
    if (d.regras) { state.regras = d.regras; salvarEstado(); capCache.clear(); return render(); }
    if (d.fonte) { state.fonte = d.fonte; salvarEstado(); capCache.clear(); return render(); }
    if (d.exp) { $("#export-menu")?.remove(); return exportar(d.exp); }
    if (d.col !== undefined && state.view === "tendencias") {
      const p = listaNivel()[Number(d.col)];
      if (state.nivel === "Semanal") state.semana = p.chave; else state.mes = p.mes;
      salvarEstado(); montarSeletor(); return render();
    }
    if (d.delMv !== undefined) { cenario.remanejamentos.splice(Number(d.delMv), 1); return mudouCenario(true); }
    if (d.delFd !== undefined) { delete cenario.fatorFamilia[d.delFd]; return mudouCenario(true); }
    if (d.delQ !== undefined) { delete cenario.turnos[d.delQ]; return mudouCenario(true); }
    switch (d.acao) {
      case "todas": state.mostrarTodas = !state.mostrarTodas; return render();
      case "copiar-resumo": {
        const txt = $("#resumo").textContent;
        navigator.clipboard.writeText(txt).then(() => { b.textContent = "Copiado"; setTimeout(() => { b.textContent = "Copiar"; }, 1500); })
          .catch(() => { const sel = getSelection(); const rg = document.createRange(); rg.selectNodeContents($("#resumo")); sel.removeAllRanges(); sel.addRange(rg); toast("Texto selecionado: use Ctrl+C."); });
        return;
      }
      case "add-mv": {
        const de = $("#sim-de").value, para = $("#sim-para").value, hc = Number($("#sim-hc").value);
        if (de === para || !(hc > 0)) return toast("Escolha famílias diferentes e um HC maior que zero.");
        cenario.remanejamentos.push({ de, para, turno: $("#sim-turno").value, hc }); return mudouCenario(true);
      }
      case "add-fd": { const v = Number($("#sim-fd-v").value); if (!(v >= 0)) return; cenario.fatorFamilia[$("#sim-fd-fam").value] = v / 100; return mudouCenario(true); }
      case "add-q": { const f = $("#sim-q-fam").value; cenario.turnos[f] = [0, 1, 2].map((i) => Math.max(0, Number($("#sim-q" + i).value) || 0)); return mudouCenario(true); }
      case "limpar-cenario": cenario = E.cenarioVazio(); return mudouCenario(true);
      case "salvar-carga":
        if (!importacao?.dataset) return;
        if (store.get("historico", []).length && !importacao.confirmar) { importacao.confirmar = true; return render(); }
        return aplicarDataset(importacao.dataset, importacao.arquivo);
      case "confirmar-carga": return aplicarDataset(importacao.dataset, importacao.arquivo);
      case "cancelar-carga": importacao.confirmar = false; return render();
      case "restaurar":
        store.del("dataset"); dataset = embutido; ctx = E.preparar(dataset); periodos = E.listarPeriodos(ctx); cache = new Map(); capCache.clear();
        toast("Base inicial restaurada."); montarSeletor(); return render();
    }
    if (b.id === "btn-export") return menuExportar();
    if (b.id === "btn-sim") return irPara("simulador");
    if (b.id === "nav-open") return $("#shell").classList.add("nav-open");
    if (b.id === "nav-close") return $("#shell").classList.remove("nav-open");
  });
  $("#backdrop").addEventListener("click", () => $("#shell").classList.remove("nav-open"));
  document.addEventListener("keydown", (ev) => { if (ev.key === "Escape") { $("#shell").classList.remove("nav-open"); $("#export-menu")?.remove(); } });

  document.addEventListener("change", (ev) => {
    const el = ev.target;
    if (el.id === "period-select") {
      if (state.nivel === "Semanal") state.semana = Number(el.value); else state.mes = el.value;
      salvarEstado(); return render();
    }
    if (el.dataset.inter !== undefined) {
      const f = el.dataset.inter, s = new Set(state.intercambiavel);
      el.checked ? s.add(f) : s.delete(f); state.intercambiavel = [...s]; salvarEstado(); return render();
    }
    if (el.id === "arquivo" && el.files[0]) return lerArquivo(el.files[0]);
    if (el.id === "sim-q-fam") return quadroInputs();
    if (el.id === "sim-adiados") { cenario.lotesAdiados = Math.max(0, Number(el.value) || 0); return mudouCenario(); }
    if (el.id === "sim-aus") { cenario.ausencias = el.value === "" ? null : Math.max(0, Number(el.value)); return mudouCenario(); }
  });
  document.addEventListener("input", (ev) => {
    const el = ev.target, v = Number(el.value);
    if (el.id === "sim-oee") { cenario.oeeMO = v; $("#sim-oee-v").textContent = pct(v, 0); return mudouCenario(); }
    if (el.id === "sim-oee-eq") { cenario.oeeEquip = v === 1 ? null : v; $("#sim-oee-eq-v").textContent = pct(v, 0); return mudouCenario(); }
    if (el.id === "sim-dem") { cenario.fatorDemanda = v; $("#sim-dem-v").textContent = pct(v, 0); return mudouCenario(); }
  });
  // arrastar e soltar o arquivo
  document.addEventListener("dragover", (ev) => { const d = ev.target.closest && ev.target.closest("#drop"); if (d) { ev.preventDefault(); d.classList.add("over"); } });
  document.addEventListener("dragleave", (ev) => { const d = ev.target.closest && ev.target.closest("#drop"); if (d) d.classList.remove("over"); });
  document.addEventListener("drop", (ev) => { const d = ev.target.closest && ev.target.closest("#drop"); if (d) { ev.preventDefault(); d.classList.remove("over"); if (ev.dataTransfer.files[0]) lerArquivo(ev.dataTransfer.files[0]); } });

  montarSeletor();
  render();
})();
