// Lê uma tabela do Excel pelo nome (Tabela_BASE_TC, Tabela_DEMANDA, ...).
(nome as text) as table =>
    let
        Linha = Table.SelectRows(Pasta, each [Kind] = "Table" and [Item] = nome),
        Dados = if Table.IsEmpty(Linha)
                then error Error.Record("Tabela não encontrada", "A pasta de trabalho não tem a tabela " & nome, nome)
                else Linha{0}[Data]
    in
        Dados
