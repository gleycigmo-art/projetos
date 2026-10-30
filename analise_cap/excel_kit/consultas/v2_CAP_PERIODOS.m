// CAP de cada mês (Jan a Dez) e de cada semana disponível, numa tabela só.
// Intervalos de meses são calculados no PAINEL como média simples dos meses.
let
    Meses      = {"Jan/2026", "Fev/2026", "Mar/2026", "Abr/2026", "Mai/2026", "Jun/2026",
         "Jul/2026", "Ago/2026", "Set/2026", "Out/2026", "Nov/2026", "Dez/2026"},
    FiltrosMes = List.Transform({0..11}, (i) =>
                     [NivelDeVisao = "Mensal", MesesAtivos = {Meses{i}}, SemanaChave = null,
                      Tipo = "Mensal", Ordem = i + 1, Periodo = Meses{i}]),
    FiltrosSem = List.Transform(Table.ToRecords(v2_SEMANAS), (s) =>
                     [NivelDeVisao = "Semanal", MesesAtivos = {}, SemanaChave = s[Ordem],
                      Tipo = "Semanal", Ordem = s[Ordem], Periodo = s[Periodo]]),
    Tabelas    = List.Transform(FiltrosMes & FiltrosSem, (f) =>
                     let
                         Cap = fnCAP(f),
                         T1  = Table.AddColumn(Cap, "Tipo", each f[Tipo], type text),
                         T2  = Table.AddColumn(T1, "Periodo", each f[Periodo], type text),
                         T3  = Table.AddColumn(T2, "Ordem", each f[Ordem], Int64.Type)
                     in
                         T3),
    Combinado  = Table.Combine(Tabelas),
    Ordenado   = Table.ReorderColumns(Combinado,
                     {"Tipo", "Periodo", "Ordem"} & List.RemoveItems(Table.ColumnNames(Combinado), {"Tipo", "Periodo", "Ordem"}))
in
    Ordenado
