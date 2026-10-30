// Linhas do calendário do período: a semana escolhida ou os meses ativos.
(Filtro as record) as table =>
let
    CalMes     = v2_FONTES_CAL[CalMes],
    CalSem     = v2_FONTES_CAL[CalSem],
    Meses      = {"Jan/2026", "Fev/2026", "Mar/2026", "Abr/2026", "Mai/2026", "Jun/2026",
         "Jul/2026", "Ago/2026", "Set/2026", "Out/2026", "Nov/2026", "Dez/2026"},
    NumsAtivos = List.Transform(Filtro[MesesAtivos], each List.PositionOf(Meses, _) + 1),
    Resultado  = if Filtro[NivelDeVisao] = "Semanal"
                 then Table.SelectRows(CalSem, each [Semana Chave] = Filtro[SemanaChave])
                 else Table.SelectRows(CalMes, each List.Contains(NumsAtivos, [#"Mês Nº"]))
in
    Resultado
