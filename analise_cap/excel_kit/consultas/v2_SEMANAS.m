// Semanas disponíveis na base de quarentena (SharePoint), com o rótulo usado no PAINEL.
let
    Chaves  = List.Sort(List.Distinct(List.RemoveNulls(Table.Column(Tabela_DemandaSemanal, "Chave semana")))),
    Tabela  = Table.FromList(Chaves, Splitter.SplitByNothing(), {"Ordem"}),
    Tipada  = Table.TransformColumnTypes(Tabela, {{"Ordem", Int64.Type}}),
    Rotulo  = Table.AddColumn(Tipada, "Periodo",
                  each "Sem " & Text.PadStart(Text.Middle(Text.From([Ordem]), 4), 2, "0")
                       & "/" & Text.Start(Text.From([Ordem]), 4),
                  type text)
in
    Table.ReorderColumns(Rotulo, {"Periodo", "Ordem"})
