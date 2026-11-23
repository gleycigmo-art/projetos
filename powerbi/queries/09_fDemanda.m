// Lotes por período e produto.
//   Mensal: plano da aba DEMANDA (CONSIDERAR? = SIM).
//   Semanal: quarentena real quando existe; nas demais semanas, o plano mensal rateado
//   pelos dias da semana que caem em cada mês (pesos de dia útil da aba FAMÍLIAS).
let
    ANO = 2026,
    NomesMes = {"Jan/2026", "Fev/2026", "Mar/2026", "Abr/2026", "Mai/2026", "Jun/2026",
                "Jul/2026", "Ago/2026", "Set/2026", "Out/2026", "Nov/2026", "Dez/2026"},
    Pad2 = (n) => Text.PadStart(Text.From(n), 2, "0"),
    Ids = {"CÓD", "DESCRIÇÃO", "MÉTODO DOCNIX", "complexidade"},

    // ── plano mensal ─────────────────────────────────────────────────────────
    Dem = fnTabela("Tabela_DEMANDA"),
    Considerar = Table.SelectRows(Dem, each [#"CONSIDERAR?"] = "SIM"),
    Sel = Table.SelectColumns(Considerar, Ids & NomesMes, MissingField.UseNull),
    SemErro = Table.ReplaceErrorValues(Sel, List.Transform(Ids & NomesMes, each {_, null})),
    Unpivot = Table.UnpivotOtherColumns(SemErro, Ids, "Mes", "Qtde"),   // descarta meses vazios
    MesNum = Table.AddColumn(Unpivot, "MesNum", each List.PositionOf(NomesMes, [Mes]) + 1, Int64.Type),
    PlanoMes = Table.FromRecords(List.Transform(Table.ToRecords(MesNum), (r) =>
        [Periodo_ID = Text.From(ANO) & "-M" & Pad2(r[MesNum]), MesNum = r[MesNum], COD = Text.From(r[#"CÓD"]),
         Descricao = r[#"DESCRIÇÃO"], MPR = r[#"MÉTODO DOCNIX"], Lotes = fnNum(r[Qtde]),
         Complexidade = fnNum(r[complexidade]), SemTempos = r[#"MÉTODO DOCNIX"] = "Método não encontrado"])),

    // ── quarentena real (semanal) ────────────────────────────────────────────
    Sem = Table.ReplaceErrorValues(fnTabela("Tabela_DemandaSemanal"),
              List.Transform(Table.ColumnNames(fnTabela("Tabela_DemandaSemanal")), each {_, null})),
    SemValidas = Table.SelectRows(Sem, each fnNum([Chave semana]) > 0),
    SemanaDe = (chave) => Number.From(Text.Middle(Text.From(fnNum(chave)), 4)),
    Real = Table.FromRecords(List.Transform(Table.ToRecords(SemValidas), (r) =>
        [Periodo_ID = Text.From(ANO) & "-S" & Pad2(SemanaDe(r[Chave semana])), COD = Text.From(r[Cod]),
         Descricao = null, MPR = r[#"MÉTODO DOCNIX"], Lotes = fnNum(r[Lote semana]),
         Complexidade = fnNum(r[Complexidade]), SemTempos = r[Status2] = "Sem tempos"])),
    SemanasReais = List.Buffer(List.Distinct(Real[Periodo_ID])),

    // ── projeção semanal: fração de cada mês contida em cada semana ISO ──────
    Peso = List.Buffer(stg_PesoDia),
    PesoDia = (d) => Peso{Date.DayOfWeek(d, Day.Monday)},
    PesoMes = List.Buffer(List.Transform({1 .. 12}, (m) =>
        List.Sum(List.Transform(List.Dates(#date(ANO, m, 1), Date.Day(Date.EndOfMonth(#date(ANO, m, 1))), #duration(1, 0, 0, 0)), PesoDia)))),
    Jan4 = #date(ANO, 1, 4),
    Seg1 = Date.AddDays(Jan4, -Date.DayOfWeek(Jan4, Day.Monday)),
    NumSemanas = List.Transform(fnTabela("CALENDARIO_SEMANA")[#"Semana Nº"], each Number.From(fnNum(_))),
    Fracoes = Table.FromRecords(List.Combine(List.Transform(NumSemanas, (n) =>
        let
            id = Text.From(ANO) & "-S" & Pad2(n),
            dias = List.Select(List.Dates(Date.AddDays(Seg1, 7 * (n - 1)), 7, #duration(1, 0, 0, 0)), each Date.Year(_) = ANO),
            porMes = List.Distinct(List.Transform(dias, Date.Month))
        in
            if List.Contains(SemanasReais, id) then {}
            else List.Transform(porMes, (m) =>
                [Periodo_ID = id, MesNum = m,
                 Frac = List.Sum(List.Transform(List.Select(dias, each Date.Month(_) = m), PesoDia)) / PesoMes{m - 1}])))),
    JFrac = Table.NestedJoin(Fracoes, {"MesNum"}, Table.RemoveColumns(PlanoMes, {"Periodo_ID"}), {"MesNum"}, "P", JoinKind.Inner),
    EFrac = Table.ExpandTableColumn(JFrac, "P", {"COD", "Descricao", "MPR", "Lotes", "Complexidade", "SemTempos"}),
    Projecao = Table.Group(EFrac, {"Periodo_ID", "COD", "Descricao", "MPR", "Complexidade", "SemTempos"},
                   {{"Lotes", each List.Sum(List.Transform(Table.ToRecords(_), (x) => x[Lotes] * x[Frac])), type number}}),

    Tudo = Table.Combine({Table.RemoveColumns(PlanoMes, {"MesNum"}), Real, Projecao}),
    Tipado = Table.TransformColumnTypes(Tudo, {
                 {"Periodo_ID", type text}, {"COD", type text}, {"Descricao", type text}, {"MPR", type text},
                 {"Lotes", type number}, {"Complexidade", Int64.Type}, {"SemTempos", type logical}})
in
    Tipado
