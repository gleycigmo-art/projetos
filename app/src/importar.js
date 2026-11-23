/*
 * FLOW LAB — leitura da planilha MFV (.xlsb/.xlsx) no navegador com SheetJS.
 * Localiza cada tabela pela aba e pela linha de cabeçalho, valida colunas, tipos e
 * duplicidades, e devolve o mesmo formato de dados que o build gera a partir dos CSVs.
 */
(function (root) {
  "use strict";

  // Onde está cada tabela. `ancora` identifica o cabeçalho; `paraNoVazio` encerra a
  // leitura na primeira linha sem valor na âncora (tabelas com blocos abaixo delas);
  // `exige` descarta linhas sem valor nessa coluna (ex.: totais abaixo da DEMANDA).
  const FONTES = [
    { nome: "base_tc", rotulo: "Base de tempos (BASE TC)", aba: "BASE TC", ancora: "Método e versão (DocNix)", obrigatoria: true },
    { nome: "demanda_mensal", rotulo: "Demanda mensal", aba: "DEMANDA", ancora: "CONSIDERAR?", exige: "CÓD", obrigatoria: true },
    { nome: "demanda_semanal", rotulo: "Lotes da quarentena (semana)", aba: "Tabela_DemandaSemanal", ancora: "Chave semana", obrigatoria: false },
    { nome: "familias", rotulo: "Mão de obra e recursos (FAMÍLIAS)", aba: "FAMÍLIAS", ancora: "PROCESSO", paraNoVazio: true, obrigatoria: true },
    { nome: "calendario_mes", rotulo: "Calendário mensal", aba: "FAMÍLIAS", ancora: "Mês Nº", paraNoVazio: true, obrigatoria: true },
    { nome: "calendario_semana", rotulo: "Calendário semanal", aba: "FAMÍLIAS", ancora: "Semana Chave", paraNoVazio: true, obrigatoria: true },
    { nome: "campanha", rotulo: "Lotes por campanha", aba: "CAMPANHA", ancora: "QTD LOTES CAMPANHA", obrigatoria: true, chave: "CÓD" },
    { nome: "skip", rotulo: "Skip de impureza", aba: "SKIP_IMPUREZA", ancora: "Qtde Lotes Skip", obrigatoria: true, chave: "MPR" },
    { nome: "etapas_campanha", rotulo: "% de tempo em campanha", aba: "ETAPAS_CAMPANHAS", ancora: "% tempo", obrigatoria: true },
  ];

  // Colunas que precisam ser numéricas (o texto "S" nas etapas de leitura é aceito)
  const NUMERICAS = {
    familias: ["TURNOS", "QTDE 1° TURNO", "QTDE 2° TURNO", "QTDE 3° TURNO", "OEE REAL"],
    calendario_mes: ["Mês Nº", "Dias Úteis", "MO - 3", "EQP - 3"],
    calendario_semana: ["Semana Chave", "Dias Úteis", "MO - 3", "EQP - 3"],
    campanha: ["QTD LOTES CAMPANHA"], skip: ["Qtde Lotes Skip"], etapas_campanha: ["% tempo"],
    demanda_semanal: ["Chave semana", "Lote semana"],
  };

  const norm = (v) => (typeof v === "string" ? v.trim() : v);
  const vazio = (v) => v === null || v === undefined || v === "";
  const colLetra = (c) => { let s = ""; c += 1; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; };

  function grade(XLSX, ws) {
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
  }

  function lerTabela(XLSX, wb, fonte, colunas) {
    const res = { nome: fonte.nome, rotulo: fonte.rotulo, aba: fonte.aba, erros: [], avisos: [], linhas: 0, dados: null };
    const ws = wb.Sheets[fonte.aba];
    if (!ws) { res.erros.push(`Aba "${fonte.aba}" não encontrada.`); return res; }
    const g = grade(XLSX, ws);

    // linha de cabeçalho = a que contém a âncora e o maior número de colunas esperadas
    let hdr = -1, melhor = -1;
    for (let r = 0; r < Math.min(g.length, 40); r++) {
      const row = (g[r] || []).map(norm);
      if (!row.includes(fonte.ancora)) continue;
      const n = colunas.filter((c) => row.includes(c)).length;
      if (n > melhor) { melhor = n; hdr = r; }
    }
    if (hdr < 0) { res.erros.push(`Cabeçalho "${fonte.ancora}" não encontrado na aba "${fonte.aba}".`); return res; }
    const cab = (g[hdr] || []).map(norm);
    const cAnc = cab.indexOf(fonte.ancora);
    // Há tabelas lado a lado (ex.: FAMÍLIAS): procura primeiro dentro do bloco contíguo
    // de cabeçalho em volta da âncora e só depois na linha inteira.
    let ini = cAnc, fim = cAnc;
    while (ini > 0 && !vazio(cab[ini - 1])) ini--;
    while (fim < cab.length - 1 && !vazio(cab[fim + 1])) fim++;
    const idx = {};
    const faltando = [];
    for (const c of colunas) {
      let best = -1;
      for (let i = ini; i <= fim && best < 0; i++) if (cab[i] === c) best = i;
      if (best < 0) cab.forEach((h, i) => { if (h === c && (best < 0 || Math.abs(i - cAnc) < Math.abs(best - cAnc))) best = i; });
      if (best < 0) faltando.push(c); else idx[c] = best;
    }
    if (faltando.length) { res.erros.push(`Colunas obrigatórias ausentes: ${faltando.join(", ")}.`); return res; }

    const rows = [];
    let vaziasSeguidas = 0;
    for (let r = hdr + 1; r < g.length; r++) {
      const row = g[r] || [];
      const anc = norm(row[cAnc]);
      if (vazio(anc)) {
        if (fonte.paraNoVazio) break;
        if (++vaziasSeguidas > 200) break;
        // linhas sem âncora são mantidas (a BASE TC tem linhas sem MPR) mas só se tiverem algum dado
        if (!colunas.some((c) => !vazio(row[idx[c]]))) continue;
      } else vaziasSeguidas = 0;
      if (fonte.exige && vazio(norm(row[idx[fonte.exige]]))) continue;
      rows.push({ excelLinha: r + 1, vals: colunas.map((c) => { const v = row[idx[c]]; return typeof v === "string" ? v.trim() : v; }) });
    }
    if (!rows.length) { res.erros.push("A tabela está vazia."); return res; }

    // tipos
    for (const c of NUMERICAS[fonte.nome] || []) {
      const i = colunas.indexOf(c);
      const ruins = rows.filter((x) => !vazio(x.vals[i]) && typeof x.vals[i] !== "number" && !Number.isFinite(Number(x.vals[i])));
      ruins.slice(0, 4).forEach((x) => res.avisos.push(`Linha ${x.excelLinha}, coluna ${colLetra(idx[c])} (${c}): "${x.vals[i]}" não é número; será tratado como 0.`));
      if (ruins.length > 4) res.avisos.push(`… e mais ${ruins.length - 4} valores não numéricos em ${c}.`);
    }
    // duplicidades de chave
    if (fonte.chave) {
      const i = colunas.indexOf(fonte.chave), vistos = new Map();
      for (const x of rows) {
        const k = x.vals[i]; if (vazio(k)) continue;
        if (vistos.has(k)) res.avisos.push(`${fonte.chave} "${k}" repetido nas linhas ${vistos.get(k)} e ${x.excelLinha}; vale a primeira.`);
        else vistos.set(k, x.excelLinha);
      }
    }
    res.linhas = rows.length;
    res.dados = { cols: colunas, rows: rows.map((x) => x.vals) };
    return res;
  }

  /** Blocos soltos da aba FAMÍLIAS: ausências (P/D … Total Ausências) e peso dos dias úteis. */
  function lerBlocosFamilias(XLSX, wb) {
    const out = { ausencias: null, dia_util: null, avisos: [] };
    const ws = wb.Sheets["FAMÍLIAS"]; if (!ws) return out;
    const g = grade(XLSX, ws);
    const achar = (txt) => { for (let r = 0; r < g.length; r++) { const row = g[r] || []; for (let c = 0; c < row.length; c++) if (norm(row[c]) === txt) return [r, c]; } return null; };
    const pd = achar("P/D");
    if (pd) {
      const [r0, c0] = pd;
      const cabSem = String((g[r0 - 1] || [])[c0 + 1] || "");
      const m = cabSem.match(/(\d+)/);
      const rows = [];
      for (let r = r0; r < g.length; r++) {
        const t = norm((g[r] || [])[c0]); if (typeof t !== "string" || !t) break;
        rows.push([t.replace(/:$/, ""), (g[r] || [])[c0 + 1] ?? null, (g[r] || [])[c0 + 2] ?? null, m ? Number(m[1]) : null]);
      }
      out.ausencias = { cols: ["TIPO", "QTDE_SEMANA", "QTDE_MEDIA_ANO", "SEMANA_REF"], rows };
    } else out.avisos.push("Bloco de ausências (P/D) não encontrado na aba FAMÍLIAS; ausências ficam zeradas.");
    const du = achar("Dia útil");
    if (du) {
      const [r0, c0] = du, rows = [];
      for (let r = r0 + 1; r < r0 + 8; r++) { const d = norm((g[r] || [])[c0 - 1]); if (typeof d === "string" && d) rows.push([d, (g[r] || [])[c0]]); }
      out.dia_util = { cols: ["DIA", "PESO"], rows };
    }
    return out;
  }

  /** Lê o arquivo (ArrayBuffer) e devolve { resultados, dataset|null }. */
  function importarPlanilha(XLSX, buffer, colunasEngine, nomeArquivo) {
    const wb = XLSX.read(buffer, { type: "array", cellFormula: false, cellHTML: false, cellText: false, cellDates: false });
    const resultados = FONTES.map((f) => lerTabela(XLSX, wb, f, colunasEngine[f.nome]));
    const blocos = lerBlocosFamilias(XLSX, wb);
    const ok = resultados.every((r) => !r.erros.length || !FONTES.find((f) => f.nome === r.nome).obrigatoria);
    let dataset = null;
    if (ok) {
      const tabelas = {};
      for (const r of resultados) tabelas[r.nome] = r.dados || { cols: colunasEngine[r.nome], rows: [] };
      tabelas.ausencias = blocos.ausencias || { cols: colunasEngine.ausencias, rows: [] };
      tabelas.dia_util = blocos.dia_util || { cols: colunasEngine.dia_util, rows: [] };
      dataset = { meta: { fonte: nomeArquivo, gerado_em: new Date().toISOString().slice(0, 10) }, tabelas };
    }
    return { resultados, blocos, dataset };
  }

  const api = { FONTES, importarPlanilha };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.FlowImport = api;
})(typeof globalThis !== "undefined" ? globalThis : this);
