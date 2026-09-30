shared SRC_DEMANDA = let
    Fonte = Excel.CurrentWorkbook(){[Name="Tabela_DEMANDA"]}[Content]
in
    Fonte;
