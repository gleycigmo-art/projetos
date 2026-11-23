// Dimensão de períodos: 12 meses e as semanas ISO de 2026, com dias úteis e horas
// disponíveis por posto (MO - n / EQP - n = horas para n turnos) do calendário da aba FAMÍLIAS.
let
    ANO = 2026,
    NomesMes = {"Jan/2026", "Fev/2026", "Mar/2026", "Abr/2026", "Mai/2026", "Jun/2026",
                "Jul/2026", "Ago/2026", "Set/2026", "Out/2026", "Nov/2026", "Dez/2026"},
    ColsHoras = {"MO - 1", "MO - 2", "MO - 3", "EQP - 1", "EQP - 2", "EQP - 3"},
    Pad2 = (n) => Text.PadStart(Text.From(n), 2, "0"),

    CalMes = fnTabela("CALENDARIO_MES"),
    Meses = Table.FromRecords(List.Transform(Table.ToRecords(CalMes), (r) =>
        let n = fnNum(r[#"Mês Nº"]) in
        [Periodo_ID = Text.From(ANO) & "-M" & Pad2(n), Nivel = "Mensal", Numero = n,
         Rotulo = NomesMes{n - 1}, Ordem = n,
         Inicio = #date(ANO, n, 1), Fim = Date.EndOfMonth(#date(ANO, n, 1)),
         DiasUteis = fnNum(r[#"Dias Úteis"])]
        & Record.FromList(List.Transform(ColsHoras, (c) => fnNum(Record.Field(r, c))), ColsHoras))),

    Jan4 = #date(ANO, 1, 4),
    Seg1 = Date.AddDays(Jan4, -Date.DayOfWeek(Jan4, Day.Monday)),
    Reais = List.Buffer(List.Distinct(List.Transform(fnTabela("Tabela_DemandaSemanal")[Chave semana], each fnNum(_)))),
    CalSem = fnTabela("CALENDARIO_SEMANA"),
    Semanas = Table.FromRecords(List.Transform(Table.ToRecords(CalSem), (r) =>
        let
            n = fnNum(r[#"Semana Nº"]),
            ini = Date.AddDays(Seg1, 7 * (n - 1)),
            fim = Date.AddDays(ini, 6),
            dm = (d) => Pad2(Date.Day(d)) & "/" & Pad2(Date.Month(d)),
            real = List.Contains(Reais, fnNum(r[Semana Chave]))
        in
        [Periodo_ID = Text.From(ANO) & "-S" & Pad2(n), Nivel = "Semanal", Numero = n,
         Rotulo = "S" & Pad2(n) & " · " & dm(ini) & "–" & dm(fim) & (if real then " · real" else ""),
         Ordem = 100 + n, Inicio = ini, Fim = fim, DiasUteis = fnNum(r[#"Dias Úteis"])]
        & Record.FromList(List.Transform(ColsHoras, (c) => fnNum(Record.Field(r, c))), ColsHoras))),

    Todos = Table.Combine({Meses, Semanas}),
    ComFonte = Table.AddColumn(Todos, "Fonte", each
                   if [Nivel] = "Mensal" then "Plano mensal"
                   else if List.Contains(Reais, fnNum(Text.From(ANO) & Text.From([Numero]))) then "Quarentena real"
                   else "Projeção do plano", type text),
    Tipado = Table.TransformColumnTypes(ComFonte, {
                 {"Periodo_ID", type text}, {"Nivel", type text}, {"Numero", Int64.Type}, {"Rotulo", type text},
                 {"Ordem", Int64.Type}, {"Inicio", type date}, {"Fim", type date}, {"DiasUteis", type number},
                 {"MO - 1", type number}, {"MO - 2", type number}, {"MO - 3", type number},
                 {"EQP - 1", type number}, {"EQP - 2", type number}, {"EQP - 3", type number}})
in
    Tipado
