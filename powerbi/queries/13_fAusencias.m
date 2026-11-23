// Bloco de ausências da aba FAMÍLIAS (P/D, Férias, Licença, ... Total Ausências),
// com a semana apontada e a média do ano, em HC.
let
    Folha = Table.SelectRows(Pasta, each [Kind] = "Sheet" and [Item] = "FAMÍLIAS"){0}[Data],
    Linhas = List.Buffer(Table.ToRows(Folha)),
    Txt = (x) => try Text.Trim(Text.From(x)) otherwise null,
    iLinha = List.PositionOf(List.Transform(Linhas, each List.Contains(List.Transform(_, Txt), "P/D")), true),
    iCol = List.PositionOf(List.Transform(Linhas{iLinha}, Txt), "P/D"),
    Cab = Txt(Linhas{iLinha - 1}{iCol + 1}),
    SemanaRef = try Number.From(Text.Select(Cab, {"0" .. "9"})) otherwise null,
    Bloco = List.Generate(
                () => iLinha,
                each _ < List.Count(Linhas) and Txt(Linhas{_}{iCol}) <> null and Txt(Linhas{_}{iCol}) <> ""
                     and not Value.Is(Linhas{_}{iCol}, type number),
                each _ + 1,
                each [TIPO = Text.TrimEnd(Txt(Linhas{_}{iCol}), ":"),
                      QTDE_SEMANA = fnNum(Linhas{_}{iCol + 1}),
                      QTDE_MEDIA_ANO = fnNum(Linhas{_}{iCol + 2}),
                      SEMANA_REF = SemanaRef]),
    Tabela = if iLinha < 0 then #table(type table [TIPO = text, QTDE_SEMANA = number, QTDE_MEDIA_ANO = number, SEMANA_REF = number], {})
             else Table.FromRecords(Bloco),
    Tipado = Table.TransformColumnTypes(Tabela, {{"TIPO", type text}, {"QTDE_SEMANA", type number},
                                                 {"QTDE_MEDIA_ANO", type number}, {"SEMANA_REF", Int64.Type}})
in
    Tipado
