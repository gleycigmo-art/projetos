shared Tabela_DEMANDA = let
    Fonte = Excel.CurrentWorkbook(){[Name="Tabela_DEMANDA"]}[Content],
    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte,{{"CONSIDERAR?", type text}, {"TIPO", type text}, {"SITE", type text}, {"CHECK DOC", type text}, {"CÓD", type text}, {"DESCRIÇÃO", type text}, {"MÉTODO DOCNIX", type text}, {"Jan/2026", type any}, {"Fev/2026", type any}, {"Mar/2026", type any}, {"Abr/2026", type any}, {"Mai/2026", type any}, {"Jun/2026", type any}, {"Jul/2026", Int64.Type}, {"Ago/2026", Int64.Type}, {"Set/2026", type number}, {"Out/2026", type number}, {"Nov/2026", type number}, {"Dez/2026", type any}, {"Soma", type number}, {"Média", type number}, {"TC levantados?", type text}, {"OBS", type any}, {"Local de Análise", type text}})
in
    #"Tipo Alterado";
