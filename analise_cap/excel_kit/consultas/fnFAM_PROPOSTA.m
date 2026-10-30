// Tabela_FAM com as horas do período. Mesma regra da coluna HRS TURNO da planilha:
// média, nos períodos ativos, das horas acumuladas para N turnos (Mão de Obra: MO-N; Equipamento: EQP-N).
// Depois, as mesmas conversões de tipo da consulta FAM_PROPOSTA (HRS TURNO vira inteiro).
(Filtro as record) as table =>
let
    Fonte    = v2_FONTES_CAL[Fam],
    Semanal  = Filtro[NivelDeVisao] = "Semanal",
    CalAtivo = fnCALENDARIO_ATIVO(Filtro),
    ColsMO   = if Semanal then {"MO - 1", "MO - 2", "MO - 3"} else {"MO - 1", "MO - 2", "MO - 3", "ADM"},
    ColsEQ   = if Semanal then {"EQP - 1", "EQP - 2", "EQP - 3"} else {"EQP - 1", "EQP - 2", "EQP - 3", "ADM2"},
    HrsMedia = (categoria as any, turnos as any) as any =>
                   let
                       Cols = if categoria = "Mão de Obra" then ColsMO else ColsEQ,
                       Col  = Cols{Int64.From(turnos) - 1}
                   in
                       List.Average(Table.Column(CalAtivo, Col)),
    SemHrs   = Table.RemoveColumns(Fonte, {"HRS TURNO", "TOTAL RECURSO"}),
    ComHrs   = Table.AddColumn(SemHrs, "HRS TURNO", each HrsMedia([CATEGORIA], [TURNOS]), type number),
    ComTotal = Table.AddColumn(ComHrs, "TOTAL RECURSO", each [#"RECURSO MÉDIO ATUAL"] * [HRS TURNO], type number),
    #"Tipo Alterado" = Table.TransformColumnTypes(ComTotal,{
        {"PROCESSO", type text}, {"TURNOS", Int64.Type},
        {"QTDE 1° TURNO", type number}, {"QTDE 2° TURNO", type number},
        {"QTDE 3° TURNO", type number}, {"RECURSO MÉDIO ATUAL", type number},
        {"HRS TURNO", Int64.Type}, {"TOTAL RECURSO", type number}, {"OEE REAL", type number}})
in
    #"Tipo Alterado"
