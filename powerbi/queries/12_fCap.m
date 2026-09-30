// Mapa de capacidade por período (equivale à CAP_PROPOSTA, com as regras corrigidas):
// carga (CM) de cada família/processo × recursos (pessoas por turno, horas do período, OEE).
// Os indicadores (HC necessário, carregamento, ocupação) são medidas DAX, para permitir
// simular OEE e demanda no relatório.
let
    Comp = Table.Buffer(stg_Componentes),
    Periodos = Table.Buffer(dPeriodo),
    DemMPR = Table.Buffer(Table.Group(fDemanda, {"Periodo_ID", "MPR"}, {{"L", each List.Sum([Lotes]), type number}})),
    Up = (x) => Number.RoundUp(x - 0.000000000001, 0),

    // ── carga de cada componente com os lotes do MPR no período ──────────────
    J = Table.NestedJoin(Comp, {"MPR"}, DemMPR, {"MPR"}, "D", JoinKind.Inner),
    E = Table.ExpandTableColumn(J, "D", {"Periodo_ID", "L"}),
    CM = Table.AddColumn(E, "CM", each
            let
                L = [L], tc = [TC], lxc = if [LXC] = null then 0 else [LXC],
                nc = if [CAMPANHA] and lxc > 0 then Up(L / lxc) else 0,
                prim = List.Min({nc, L})
            in
                if [NSKIP] <> null then tc * Up(L / [NSKIP])
                else if nc = 0 or [PCT] = null then tc * L
                else prim * tc + (L - prim) * tc * [PCT], type number),
    Carga = Table.Group(CM, {"Periodo_ID", "DESTINO", "PROCESSO"}, {{"CM_MIN", each List.Sum([CM]), type number}}),

    // todas as combinações recurso × período (recursos sem demanda aparecem com carga 0)
    Grupos = Table.Distinct(Table.SelectColumns(Comp, {"DESTINO", "PROCESSO"})),
    Cruz = Table.ExpandListColumn(Table.AddColumn(Grupos, "Periodo_ID", each Periodos[Periodo_ID]), "Periodo_ID"),
    JC = Table.NestedJoin(Cruz, {"Periodo_ID", "DESTINO", "PROCESSO"}, Carga, {"Periodo_ID", "DESTINO", "PROCESSO"}, "C", JoinKind.LeftOuter),
    EC = Table.ExpandTableColumn(JC, "C", {"CM_MIN"}),

    // ── recurso: família+processo; se não houver (ou sem pessoas), só família ─
    Rec = Table.Buffer(dRecurso),
    RecFP = Table.Distinct(Rec, {"FAMILIA", "PROCESSO"}),
    RecF = Table.Distinct(Rec, {"FAMILIA"}),
    JFP = Table.NestedJoin(EC, {"DESTINO", "PROCESSO"}, RecFP, {"FAMILIA", "PROCESSO"}, "RFP", JoinKind.LeftOuter),
    JF = Table.NestedJoin(JFP, {"DESTINO"}, RecF, {"FAMILIA"}, "RF", JoinKind.LeftOuter),
    Escolha = Table.AddColumn(JF, "R", each
                let a = if Table.IsEmpty([RFP]) then null else [RFP]{0},
                    b = if Table.IsEmpty([RF]) then null else [RF]{0}
                in if a <> null and a[RECURSO] > 0 then a else if b <> null and b[RECURSO] > 0 then b else a),
    ComR = Table.SelectRows(Table.RemoveColumns(Escolha, {"RFP", "RF"}), each [R] <> null),
    ER = Table.ExpandRecordColumn(ComR, "R", {"CLUSTER", "CATEGORIA", "TURNOS", "QTDE_1T", "QTDE_2T", "QTDE_3T", "RECURSO", "OEE"}),

    // horas do período para o nº de turnos do recurso (MO - n ou EQP - n)
    JP = Table.NestedJoin(ER, {"Periodo_ID"}, Periodos, {"Periodo_ID"}, "P", JoinKind.Inner),
    Horas = Table.AddColumn(JP, "HRS_TURNO", each
                let p = [P]{0} in Record.Field(p, (if [CATEGORIA] = "Equipamento" then "EQP - " else "MO - ") & Text.From([TURNOS])),
                type number),
    Validas = Table.SelectRows(Table.RemoveColumns(Horas, {"P"}), each [HRS_TURNO] > 0 and [RECURSO] > 0),

    // ── demanda do recurso (lotes): família → processo → equipamento ─────────
    DemRec = Table.Buffer(Table.Group(
                Table.ExpandTableColumn(
                    Table.NestedJoin(stg_MprRecurso, {"MPR"}, DemMPR, {"MPR"}, "D", JoinKind.Inner), "D", {"Periodo_ID", "L"}),
                {"Periodo_ID", "TIPO", "CHAVE"}, {{"DEM", each List.Sum([L]), type number}})),
    Busca = (tipo as text) => Table.SelectColumns(Table.SelectRows(DemRec, each [TIPO] = tipo), {"Periodo_ID", "CHAVE", "DEM"}),
    D1 = Table.ExpandTableColumn(Table.NestedJoin(Validas, {"Periodo_ID", "DESTINO"}, Busca("FAM"), {"Periodo_ID", "CHAVE"}, "X", JoinKind.LeftOuter), "X", {"DEM"}, {"DEM_FAM"}),
    D2 = Table.ExpandTableColumn(Table.NestedJoin(D1, {"Periodo_ID", "PROCESSO"}, Busca("PROC"), {"Periodo_ID", "CHAVE"}, "X", JoinKind.LeftOuter), "X", {"DEM"}, {"DEM_PROC"}),
    D3 = Table.ExpandTableColumn(Table.NestedJoin(D2, {"Periodo_ID", "DESTINO"}, Busca("EQUIP"), {"Periodo_ID", "CHAVE"}, "X", JoinKind.LeftOuter), "X", {"DEM"}, {"DEM_EQUIP"}),
    Demanda = Table.AddColumn(D3, "DEMANDA", each
                if [DEM_FAM] <> null then [DEM_FAM] else if [DEM_PROC] <> null then [DEM_PROC]
                else if [DEM_EQUIP] <> null then [DEM_EQUIP] else 0, type number),

    Final = Table.AddColumn(Demanda, "CM_HRS", each (if [CM_MIN] = null then 0 else [CM_MIN]) / 60, type number),
    Selecao = Table.SelectColumns(Table.RenameColumns(Final, {{"DESTINO", "FAMILIA"}}), {
                  "Periodo_ID", "PROCESSO", "FAMILIA", "CLUSTER", "CATEGORIA", "CM_HRS", "RECURSO",
                  "QTDE_1T", "QTDE_2T", "QTDE_3T", "TURNOS", "HRS_TURNO", "OEE", "DEMANDA"}),
    Tipado = Table.TransformColumnTypes(Selecao, {
                 {"Periodo_ID", type text}, {"PROCESSO", type text}, {"FAMILIA", type text}, {"CLUSTER", type text},
                 {"CATEGORIA", type text}, {"CM_HRS", type number}, {"RECURSO", type number}, {"QTDE_1T", type number},
                 {"QTDE_2T", type number}, {"QTDE_3T", type number}, {"TURNOS", Int64.Type}, {"HRS_TURNO", type number},
                 {"OEE", type number}, {"DEMANDA", type number}})
in
    Tipado
