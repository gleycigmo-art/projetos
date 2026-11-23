// Peso de cada dia da semana (seg..dom) do bloco "Dia útil" da aba FAMÍLIAS.
// Usado para ratear o plano mensal nas semanas sem quarentena real. (não carregar)
let
    Folha = Table.SelectRows(Pasta, each [Kind] = "Sheet" and [Item] = "FAMÍLIAS"){0}[Data],
    Linhas = List.Buffer(Table.ToRows(Folha)),
    Txt = (x) => try Text.Trim(Text.From(x)) otherwise null,
    iLinha = List.PositionOf(List.Transform(Linhas, each List.Contains(List.Transform(_, Txt), "Dia útil")), true),
    iCol = if iLinha < 0 then -1 else List.PositionOf(List.Transform(Linhas{iLinha}, Txt), "Dia útil"),
    Dias = {"seg", "ter", "qua", "qui", "sex", "sáb", "dom"},
    Padrao = {1, 1, 1, 1, 1, 0, 0},
    Lidos = if iLinha < 0 then {} else
        List.Transform({1 .. 7}, each
            let r = if iLinha + _ < List.Count(Linhas) then Linhas{iLinha + _} else {}
            in [Dia = try Text.Lower(Txt(r{iCol - 1})) otherwise null, Peso = try fnNum(r{iCol}) otherwise null]),
    Pesos = List.Transform({0 .. 6}, (i) =>
        let achado = List.Select(Lidos, each [Dia] = Dias{i})
        in if List.IsEmpty(achado) then Padrao{i} else achado{0}[Peso])
in
    Pesos
