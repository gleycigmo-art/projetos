shared CAP_PROPOSTA = let

    // ── Fonte principal ───────────────────────────────────────────────────────
    Fonte = TC_COM_CAMPANHA,

    FonteTexto = Table.TransformColumnTypes(Fonte, {
                     {"FAM proposta", type text},
                     {"PROCESSO",     type text},
                     {"CM_Min",       type number},
                     {"Demanda_MPR",  type number}}),

    Agrupado = Table.Group(FonteTexto,
                   {"FAM proposta", "PROCESSO"}, {
                   {"CM_Min_Total", each List.Sum([CM_Min]),
                    type number}}),

    AgrupRen = Table.RenameColumns(Agrupado,
                   {{"PROCESSO", "PROCESSO_KEY"}}),

    // ── Tabelas pequenas de lookup, bufferizadas uma única vez ─────────────────
    FAM_PROPOSTA_Buf = Table.Buffer(FAM_PROPOSTA),

    Mesclado = Table.NestedJoin(AgrupRen,
                   {"FAM proposta", "PROCESSO_KEY"},
                   FAM_PROPOSTA_Buf,
                   {"FAMILIA", "PROCESSO"},
                   "Map", JoinKind.LeftOuter),
    Expandido = Table.ExpandTableColumn(Mesclado, "Map", {
                    "FAMILIA", "CLUSTER",
                    "RECURSO MÉDIO ATUAL",
                    "HRS TURNO",
                    "TOTAL RECURSO",
                    "OEE REAL"}),

    ComProcesso = Table.RenameColumns(Expandido,
                      {{"PROCESSO_KEY", "PROCESSO"}}),

    ComTipos = Table.TransformColumnTypes(ComProcesso, {
                   {"FAM proposta", type text},
                   {"PROCESSO",     type text},
                   {"FAMILIA",      type text},
                   {"CLUSTER",      type text}}),

    ComFamilia = Table.AddColumn(ComTipos, "FAMILIA_FINAL",
                     each if [FAMILIA] <> null
                          then [FAMILIA]
                          else [#"FAM proposta"],
                     type text),

    ComMatch = Table.SelectRows(ComFamilia,
                   each [TOTAL RECURSO] <> null),
    SemMatch = Table.SelectRows(ComFamilia,
                   each [TOTAL RECURSO] = null),

    SemMatch_Join = Table.NestedJoin(SemMatch,
                        {"FAMILIA_FINAL"},
                        FAM_PROPOSTA_Buf,
                        {"FAMILIA"},
                        "Map2", JoinKind.LeftOuter),
    SemMatch_Exp  = Table.ExpandTableColumn(SemMatch_Join, "Map2", {
                        "CLUSTER",
                        "RECURSO MÉDIO ATUAL",
                        "HRS TURNO",
                        "TOTAL RECURSO",
                        "OEE REAL"},
                        {"CLUSTER_2",
                         "RECURSO MÉDIO ATUAL_2",
                         "HRS TURNO_2",
                         "TOTAL RECURSO_2",
                         "OEE REAL_2"}),

    SemMatch_Ren = Table.RenameColumns(SemMatch_Exp, {
                       {"CLUSTER",             "CLUSTER_OLD"},
                       {"RECURSO MÉDIO ATUAL", "RECURSO MÉDIO ATUAL_OLD"},
                       {"HRS TURNO",           "HRS TURNO_OLD"},
                       {"TOTAL RECURSO",       "TOTAL RECURSO_OLD"},
                       {"OEE REAL",            "OEE REAL_OLD"}}),

    SemMatch_ComRecord = Table.AddColumn(SemMatch_Ren, "Map3",
                              each [
                                  CLUSTER = [CLUSTER_2],
                                  #"RECURSO MÉDIO ATUAL" = [#"RECURSO MÉDIO ATUAL_2"],
                                  #"HRS TURNO" = [#"HRS TURNO_2"],
                                  #"TOTAL RECURSO" = [#"TOTAL RECURSO_2"],
                                  #"OEE REAL" = [#"OEE REAL_2"]
                              ]),
    SemMatch_ExpRec = Table.ExpandRecordColumn(SemMatch_ComRecord, "Map3",
                              {"CLUSTER", "RECURSO MÉDIO ATUAL", "HRS TURNO",
                               "TOTAL RECURSO", "OEE REAL"}),
    SemMatch_Final = Table.SelectColumns(SemMatch_ExpRec,
                          {"FAM proposta", "PROCESSO", "CM_Min_Total",
                           "FAMILIA_FINAL", "CLUSTER",
                           "RECURSO MÉDIO ATUAL", "HRS TURNO",
                           "TOTAL RECURSO", "OEE REAL"}),

    ComMatch_Sel = Table.SelectColumns(ComMatch, {
                       "FAM proposta", "PROCESSO", "CM_Min_Total",
                       "FAMILIA_FINAL", "CLUSTER",
                       "RECURSO MÉDIO ATUAL", "HRS TURNO",
                       "TOTAL RECURSO", "OEE REAL"}),

    Reunido = Table.Combine({ComMatch_Sel, SemMatch_Final}),

    ReunidoLimpo = Table.TransformColumnTypes(Reunido, {
                       {"FAM proposta",        type text},
                       {"PROCESSO",            type text},
                       {"FAMILIA_FINAL",       type text},
                       {"CLUSTER",             type text},
                       {"CM_Min_Total",        type number},
                       {"RECURSO MÉDIO ATUAL", type number},
                       {"HRS TURNO",           type number},
                       {"TOTAL RECURSO",       type number},
                       {"OEE REAL",            type number}}),

    Filtrado = Table.SelectRows(ReunidoLimpo,
                   each [TOTAL RECURSO] <> null
                     and [TOTAL RECURSO] <> 0),

    // ── Métricas de capacidade ────────────────────────────────────────────────
    CM_HRS = Table.AddColumn(Filtrado, "CM (HRS)",
                 each [CM_Min_Total] / 60,
                 type number),

    OEE_NEC = Table.AddColumn(CM_HRS, "OEE NEC",
                  each if [TOTAL RECURSO] <> null and [TOTAL RECURSO] <> 0
                       then [#"CM (HRS)"] / [TOTAL RECURSO]
                       else null,
                  Percentage.Type),

    CARREG = Table.AddColumn(OEE_NEC, "CARREGAMENTO",
                 each if [OEE REAL] <> null and [OEE REAL] <> 0
                      then [OEE NEC] / [OEE REAL]
                      else null,
                 Percentage.Type),

    HC_NEC = Table.AddColumn(CARREG, "HC NECESSARIO",
                 each if [HRS TURNO] <> null and [HRS TURNO] <> 0
                       and [OEE REAL] <> null and [OEE REAL] <> 0
                      then [#"CM (HRS)"] / ([HRS TURNO] * [OEE REAL])
                      else null,
                 type number),

    // ── Demanda por família ───────────────────────────────────────────────────
    ColsBase       = Table.ColumnNames(BASE_TC_CAMP),
    TC_Orig_Global = FONTES_TC[TC_Orig],

    ColsEquipTags = List.Distinct(
                        List.Select(TC_Orig_Global[Equipamento Proposto],
                            each _ <> null and _ <> ""
                              and (Text.StartsWith(_, "CRLQ") or Text.StartsWith(_, "CRGS")))),

    EqNInj12      = "12 - Nº INJEÇÕES PADRÃO - LEITURA",
    EqTCor12      = "12 - TEMPO PADRÃO (CORRIDA) - LEITURA",
    EqNInj13      = "13 - Nº INJEÇÕES AMOSTRA/ DILUENTE/ PLACEBO - LEITURA",
    EqTCor13      = "13 - TEMPO AMOSTRA/ DILUENTE/ PLACEBO (CORRIDA) - LEITURA",
    EqCols        = List.Intersect({
                        {EqNInj12, EqTCor12, EqNInj13, EqTCor13},
                        ColsBase}),

    // ── BASE_TC_CAMP + Demanda_MPR pré-unidos e bufferizados UMA VEZ ───────────
    BaseTC_JoinDem = Table.NestedJoin(BASE_TC_CAMP,
                          {"Método e versão (DocNix)"}, DEMANDA_MPR,
                          {"MÉTODO DOCNIX"}, "Dem", JoinKind.LeftOuter),
    BaseTC_ExpDem  = Table.ExpandTableColumn(BaseTC_JoinDem, "Dem",
                          {"Demanda_MPR"}, {"Demanda_MPR"}),
    BaseTC_ComDemanda = Table.Buffer(BaseTC_ExpDem),

    BaseDem_Fam = Table.Distinct(
                      Table.SelectColumns(BaseTC_ComDemanda,
                          {"Método e versão (DocNix)", "FAM proposta", "Demanda_MPR"}),
                      {"Método e versão (DocNix)", "FAM proposta"}),

    DemPorFam   = Table.SelectRows(
                      Table.Group(BaseDem_Fam,
                          {"FAM proposta"}, {
                          {"Demanda_Familia", each List.Sum([Demanda_MPR]),
                           type number}}),
                      each [#"FAM proposta"] <> null
                        and [#"FAM proposta"] <> "Terceirizado"),

    // ── Funções auxiliares para demanda ─────────────────────────────────────────
    SomaDemMPR = (colunas as list) as number =>
        let
            ColsExist = List.Intersect({colunas, ColsBase}),
            Resultado =
                if List.IsEmpty(ColsExist) then 0
                else let
                    B    = Table.SelectColumns(BaseTC_ComDemanda,
                               {"Método e versão (DocNix)", "Demanda_MPR"} & ColsExist),
                    UP   = Table.UnpivotOtherColumns(B,
                               {"Método e versão (DocNix)", "Demanda_MPR"},
                               "Etapa", "TC_Val"),
                    F    = Table.SelectRows(UP,
                               each let v = [TC_Val]
                                    in v <> null and v <> 0 and v <> ""),
                    MPRs = Table.Distinct(
                               Table.SelectColumns(F,
                                   {"Método e versão (DocNix)", "Demanda_MPR"}),
                               {"Método e versão (DocNix)"})
                in
                    List.Sum(MPRs[Demanda_MPR])
        in
            Resultado,

    SomaDemMPRFam = (coluna as text, filtroFam as text,
                     operador as text) as number =>
        let
            Existe = List.Contains(ColsBase, coluna),
            Resultado =
                if not Existe then 0
                else let
                    B    = Table.SelectColumns(BaseTC_ComDemanda,
                               {"Método e versão (DocNix)", "FAM proposta",
                                "Demanda_MPR", coluna}),
                    F    = Table.SelectRows(B,
                               each let v = Record.Field(_, coluna)
                                    in v <> null and v <> 0 and v <> ""),
                    FF   = if operador = "<>" then
                               Table.SelectRows(F,
                                   each [#"FAM proposta"] <> filtroFam)
                           else
                               Table.SelectRows(F,
                                   each [#"FAM proposta"] = filtroFam),
                    MPRs = Table.Distinct(
                               Table.SelectColumns(FF,
                                   {"Método e versão (DocNix)", "Demanda_MPR"}),
                               {"Método e versão (DocNix)"})
                in
                    List.Sum(MPRs[Demanda_MPR])
        in
            Resultado,

    SomaDemEquipIndiv = (nomeEquip as text) as number =>
        let
            F    = Table.SelectRows(TC_Orig_Global,
                       each [Equipamento Proposto] = nomeEquip),
            MPRs = Table.Distinct(
                       Table.SelectColumns(F, {"Método e versão (DocNix)"}),
                       {"Método e versão (DocNix)"}),
            JD   = Table.NestedJoin(MPRs,
                       {"Método e versão (DocNix)"}, DEMANDA_MPR,
                       {"MÉTODO DOCNIX"}, "Dem", JoinKind.LeftOuter),
            ED   = Table.ExpandTableColumn(JD, "Dem",
                       {"Demanda_MPR"}, {"Demanda_MPR"}),
            Soma = List.Sum(ED[Demanda_MPR])
        in
            if Soma = null then 0 else Soma,

    // ── Etapas para cálculo de demanda ────────────────────────────────────────
    CONF_Etapa = "7 - CONFERÊNCIA DO PREPARO - BANCADA",
    Et8        = "8 - TEMPO DO DISSOLUTOR - BANCADA",
    Et9        = "9 - MONTAGEM DO EQUIPAMENTO - LEITURA",
    Et10       = "10 - CONFERÊNCIA EQUIPAMENTO - LEITURA",
    Et14       = "14 - PROCESSAMENTO DA CORRIDA - LEITURA",
    Et15       = "15 - CONFERÊNCIA PROCESSAMENTO DA CORRIDA",
    Et16       = "16 - REVISÃO LAUDO - REVISÃO",
    ColsEsp    = {"SPCT0071 - AA", "ICP", "SPCT - IV", "SPCT - UV"},

    BaseDem_Conf = #table(
                       {"PROCESSO", "Demanda_Familia"},
                       {
                           {"MO CONFERÊNCIA BANCADA",
                            SomaDemMPRFam(CONF_Etapa, "Fam PI/TF", "<>")},
                           {"MO CONFERÊNCIA PI/TF",
                            SomaDemMPRFam(CONF_Etapa, "Fam PI/TF", "=")},
                           {"MO CONFERÊNCIA EQUIPAMENTO",
                            SomaDemMPR({Et10, Et15})},
                           {"MO EQUIPAMENTO BANCADA",
                            SomaDemMPR(ColsEsp)},
                           {"REVISÃO FINAL",
                            SomaDemMPR({Et16})},
                           {"DISSOLUTOR",
                            SomaDemMPR({Et8})},
                           {"MO EQUIPAMENTO",
                            SomaDemMPR({Et9, Et14})}
                       }),

    LinhasEquip  = List.Transform(ColsEquipTags,
                       each {_, SomaDemEquipIndiv(_)}),
    EquipDemTab  = #table(
                       {"FAM proposta", "Demanda_Familia"},
                       LinhasEquip),

    // ── Join com demanda por família ──────────────────────────────────────────
    ComDemanda = Table.NestedJoin(HC_NEC,
                     {"FAM proposta"}, DemPorFam,
                     {"FAM proposta"}, "Dem",
                     JoinKind.LeftOuter),
    DemExp     = Table.ExpandTableColumn(ComDemanda, "Dem",
                     {"Demanda_Familia"}, {"Soma Demanda Media"}),

    ComDemConf = Table.NestedJoin(DemExp,
                     {"PROCESSO"}, BaseDem_Conf,
                     {"PROCESSO"}, "DemConf",
                     JoinKind.LeftOuter),
    DemConfExp = Table.ExpandTableColumn(ComDemConf, "DemConf",
                     {"Demanda_Familia"}, {"Dem_Conf"}),

    ComDemEquip = Table.NestedJoin(DemConfExp,
                      {"FAM proposta"}, EquipDemTab,
                      {"FAM proposta"}, "DemEquip",
                      JoinKind.LeftOuter),
    DemEquipExp = Table.ExpandTableColumn(ComDemEquip, "DemEquip",
                      {"Demanda_Familia"}, {"Dem_Equip"}),

    DemFinal = Table.AddColumn(DemEquipExp, "Soma Demanda Media Final",
                  each
                      let
                          T1 = try [Soma Demanda Media],
                          T2 = try [Dem_Conf],
                          T3 = try [Dem_Equip],
                          valor =
                              if not T1[HasError] and T1[Value] <> null then T1[Value]
                              else if not T2[HasError] and T2[Value] <> null then T2[Value]
                              else if not T3[HasError] and T3[Value] <> null then T3[Value]
                              else 0
                      in
                          valor,
                  type number),
    DemFinalSel = Table.RemoveColumns(
                      Table.RenameColumns(DemFinal,
                          {{"Soma Demanda Media", "Soma Demanda Media Old"}}),
                      {"Soma Demanda Media Old", "Dem_Conf", "Dem_Equip"}),
    DemFinalRen = Table.RenameColumns(DemFinalSel,
                      {{"Soma Demanda Media Final", "Soma Demanda Media"}}),

    // ── TAKT, TC MEDIO, FIFO ──────────────────────────────────────────────────
    TAKT = Table.AddColumn(DemFinalRen, "TAKT",
               each if [Soma Demanda Media] <> null
                     and [Soma Demanda Media] <> 0
                    then [TOTAL RECURSO] / [Soma Demanda Media]
                    else null,
               type number),

    TC_MED = Table.AddColumn(TAKT, "TC MEDIO",
                 each if [Soma Demanda Media] <> null
                       and [Soma Demanda Media] <> 0
                      then [#"CM (HRS)"] / [Soma Demanda Media]
                      else null,
               type number),

    FIFO = Table.AddColumn(TC_MED, "FIFO",
               each if [TAKT] <> null and [TAKT] <> 0
                    then [TOTAL RECURSO] / [TAKT] / 22.97
                    else null,
               type number),

    // ── Arredondar ────────────────────────────────────────────────────────────
    Arredondado = Table.TransformColumns(FIFO, {
                      {"CM_Min_Total",        each if _ = null then null else Number.Round(_, 2), type number},
                      {"CM (HRS)",            each if _ = null then null else Number.Round(_, 2), type number},
                      {"RECURSO MÉDIO ATUAL", each if _ = null then null else Number.Round(_, 2), type number},
                      {"TOTAL RECURSO",       each if _ = null then null else Number.Round(_, 2), type number},
                      {"HC NECESSARIO",       each if _ = null then null else Number.Round(_, 2), type number},
                      {"Soma Demanda Media",  each if _ = null then null else Number.Round(_, 2), type number},
                      {"TAKT",               each if _ = null then null else Number.Round(_, 2), type number},
                      {"TC MEDIO",           each if _ = null then null else Number.Round(_, 2), type number},
                      {"FIFO",               each if _ = null then null else Number.Round(_, 2), type number}}),

    Renomeado  = Table.RenameColumns(Arredondado,
                     {{"FAMILIA_FINAL", "FAMILIA"}}),

    Reordenado = Table.ReorderColumns(
                     Table.SelectRows(Renomeado,
                         each [PROCESSO] <> null),
                     {"PROCESSO", "FAMILIA", "CLUSTER",
                      "CM (HRS)", "RECURSO MÉDIO ATUAL",
                      "TOTAL RECURSO", "HRS TURNO",
                      "OEE REAL", "OEE NEC", "HC NECESSARIO",
                      "CARREGAMENTO", "Soma Demanda Media",
                      "TAKT", "TC MEDIO", "FIFO"}),

    Final = Table.RemoveColumns(Reordenado,
                {"FAM proposta", "CM_Min_Total"}),

    Classificado = Table.Sort(Final,
                       {{"CLUSTER", Order.Ascending},
                        {"PROCESSO", Order.Ascending}})

in
    Classificado;
