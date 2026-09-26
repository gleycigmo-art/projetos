shared FAM_PROPOSTA = let
    Fonte = Excel.CurrentWorkbook(){[Name="Tabela_FAM"]}[Content],
    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte,{
        {"PROCESSO", type text}, {"TURNOS", Int64.Type},
        {"QTDE 1° TURNO", type number}, {"QTDE 2° TURNO", type number},
        {"QTDE 3° TURNO", type number}, {"RECURSO MÉDIO ATUAL", type number},
        {"HRS TURNO", Int64.Type}, {"TOTAL RECURSO", type number}, {"OEE REAL", type number}})
in
    #"Tipo Alterado";
