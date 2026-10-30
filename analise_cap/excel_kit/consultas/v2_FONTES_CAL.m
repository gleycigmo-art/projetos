// Tabelas da planilha usadas pelas funções v2 (lidas só aqui, como em FONTES_TC).
let
    CalMes = Excel.CurrentWorkbook(){[Name="CALENDARIO_MES"]}[Content],
    CalSem = Excel.CurrentWorkbook(){[Name="CALENDARIO_SEMANA"]}[Content],
    Fam    = Excel.CurrentWorkbook(){[Name="Tabela_FAM"]}[Content]
in
    [
        CalMes = CalMes,
        CalSem = CalSem,
        Fam    = Fam
    ]
