shared TC_COM_CAMPANHA = let

    // ── Fontes — carregadas uma única vez no início ───────────────────────────
    TC_Fonte          = BASE_TC_CAMP,
    Lotes_Camp        = FONTES_TC[Lotes_Camp],
    FamEquip          = FONTES_TC[FamEquip],
    Camp_Dedup_Global = FONTES_TC[Camp_Dedup],
    Skip_Dedup_Global = FONTES_TC[Skip_Dedup],
    TC_Orig_Global    = FONTES_TC[TC_Orig],

    // ── Função auxiliar: converte qualquer valor para número ──────────────────
    ToNum = (v as any) as number =>
        if v = null then 0
        else if v = "" then 0
        else if Value.Is(v, type number) then v
        else let n = try Number.From(v) otherwise 0 in n,

    // ── Função de cálculo de campanha ─────────────────────────────────────────
    CalcCamp = (tc as number, lotes as number, lotesXcamp as number,
                pct as number, numMeses as number) as number =>
        let
            nCamp     = Number.RoundUp(lotes / lotesXcamp, 0),
            campComp  = nCamp * tc,
            campResto = (lotes - nCamp) * tc * pct,
            tcMedio   = if lotes > 0
                        then (campComp + campResto) / lotes
                        else 0,
            resultado = if numMeses > 0
                        then tcMedio * lotes / numMeses
                        else 0
        in
            resultado,

    // ── Função de cálculo de skip ─────────────────────────────────────────────
    CalcSkip = (tc as number, lotes as number, nSkip as number,
                numMeses as number) as number =>
        let
            nSkip_Camp = if nSkip > 0
                         then Number.RoundUp(lotes / nSkip, 0)
                         else 0,
            resultado  = if numMeses > 0
                         then tc * nSkip_Camp / numMeses
                         else 0
        in
            resultado,

    // ── Colunas de identificação ──────────────────────────────────────────────
    ColsID = {"Método e versão (DocNix)", "FAM proposta",
              "Teste original", "PI"},

    // ── Colunas CRLQ/CRGS ────────────────────────────────────────────────────
    ColsNaTabela  = Table.ColumnNames(TC_Fonte),
    ColsEquipTags = List.Select(ColsNaTabela,
                       each Text.StartsWith(_, "CRLQ")
                         or Text.StartsWith(_, "CRGS")),

    // ── Grupos de etapas ──────────────────────────────────────────────────────
    EtapasBancada = {
        "1 - PREPARO (PADRÃO E SOLUÇÕES) - BANCADA",
        "2 - PREPARO AMOSTRA - BANCADA",
        "3 - ANÁLISE FQ - BANCADA",
        "6 - FINALIZAÇÃO DOS DADOS - BANCADA",
        "7 - CONFERÊNCIA DO PREPARO - BANCADA",
        "8 - TEMPO DO DISSOLUTOR - BANCADA"
    },
    EtapasEquip = {
        "9 - MONTAGEM DO EQUIPAMENTO - LEITURA",
        "10 - CONFERÊNCIA EQUIPAMENTO - LEITURA",
        "11 - TEMPO CONDICIONAMENTO - LEITURA",
        "12 - Nº INJEÇÕES PADRÃO - LEITURA",
        "12 - TEMPO PADRÃO (CORRIDA) - LEITURA",
        "13 - Nº INJEÇÕES AMOSTRA/ DILUENTE/ PLACEBO - LEITURA",
        "13 - TEMPO AMOSTRA/ DILUENTE/ PLACEBO (CORRIDA) - LEITURA",
        "14 - PROCESSAMENTO DA CORRIDA - LEITURA",
        "15 - CONFERÊNCIA PROCESSAMENTO DA CORRIDA",
        "16 - REVISÃO LAUDO - REVISÃO"
    },
    ColsEspectro = {"SPCT0071 - AA", "ICP", "SPCT - IV", "SPCT - UV"},
    EqNInj12     = "12 - Nº INJEÇÕES PADRÃO - LEITURA",
    EqTCor12     = "12 - TEMPO PADRÃO (CORRIDA) - LEITURA",
    EqNInj13     = "13 - Nº INJEÇÕES AMOSTRA/ DILUENTE/ PLACEBO - LEITURA",
    EqTCor13     = "13 - TEMPO AMOSTRA/ DILUENTE/ PLACEBO (CORRIDA) - LEITURA",

    // ── Colunas existentes por grupo ──────────────────────────────────────────
    ColsIDExist   = List.Intersect({ColsID, ColsNaTabela}),
    EtapasBExist  = List.Intersect({EtapasBancada, ColsNaTabela}),
    EtapasEExist  = List.Intersect({EtapasEquip, ColsNaTabela}),
    ColsEspExist  = List.Intersect({ColsEspectro, ColsNaTabela}),
    EqCols        = List.Intersect({
                        {EqNInj12,EqTCor12,EqNInj13,EqTCor13},
                        ColsNaTabela}),
    FlagsExist    = {"N_Camp","N_Skip","EhSkip"},

    // ── Bases limpas separadas ────────────────────────────────────────────────
    TC_Banc_Flag = Table.SelectRows(
                       Table.SelectColumns(TC_Fonte,
                           List.Distinct(
                               ColsIDExist & FlagsExist & EtapasBExist)),
                       each [#"Método e versão (DocNix)"] <> null),

    TC_Equip_Flag = Table.SelectRows(
                        Table.SelectColumns(TC_Fonte,
                            List.Intersect({
                                List.Distinct(
                                    ColsIDExist & FlagsExist & EtapasEExist),
                                ColsNaTabela})),
                        each [#"Método e versão (DocNix)"] <> null),

    TC_Esp_Flag  = if List.IsEmpty(ColsEspExist) then
                       #table(ColsIDExist & FlagsExist, {})
                   else
                       Table.SelectRows(
                           Table.SelectColumns(TC_Fonte,
                               List.Distinct(
                                   ColsIDExist & FlagsExist & ColsEspExist)),
                           each [#"Método e versão (DocNix)"] <> null),

    TC_EqTag_Flag = if List.IsEmpty(ColsEquipTags) or List.IsEmpty(EqCols) then
                        #table(ColsIDExist & FlagsExist, {})
                    else
                        Table.SelectRows(
                            Table.SelectColumns(TC_Fonte,
                                List.Distinct(
                                    ColsIDExist & FlagsExist
                                    & EqCols & ColsEquipTags)),
                            each [#"Método e versão (DocNix)"] <> null),

    // ── Funções auxiliares — usando fontes pré-carregadas ─────────────────────
    JoinDem = (tbl as table) =>
        let
            J = Table.NestedJoin(tbl,
                    {"Método e versão (DocNix)"}, DEMANDA_MPR,
                    {"MÉTODO DOCNIX"}, "Dem", JoinKind.LeftOuter),
            E = Table.ExpandTableColumn(J, "Dem",
                    {"Demanda_Total", "Demanda_MPR", "Num_Meses"},
                    {"Demanda_Total", "Demanda_MPR", "Num_Meses"})
        in
            E,

    JoinCamp = (tbl as table) =>
        let
            J = Table.NestedJoin(tbl,
                    {"Método e versão (DocNix)"},
                    Camp_Dedup_Global,
                    {"CÓD"}, "Camp2", JoinKind.LeftOuter),
            E = Table.ExpandTableColumn(J, "Camp2",
                    {"QTD LOTES CAMPANHA"}, {"LotePorCampanha"})
        in
            E,

    GetPct = (etapaKey as text) as any =>
        let r = Table.SelectRows(Lotes_Camp,
                    each [Etapa] = etapaKey & "_CAMPANHA")
        in if Table.RowCount(r) > 0
           then r{0}[#"% tempo"] else null,

    Pct_Et1  = GetPct("1 - PREPARO (PADRÃO E SOLUÇÕES) - BANCADA"),
    Pct_Et6  = GetPct("6 - FINALIZAÇÃO DOS DADOS - BANCADA"),
    Pct_Et7  = GetPct("7 - CONFERÊNCIA DO PREPARO - BANCADA"),
    Pct_Et9  = GetPct("9 - MONTAGEM DO EQUIPAMENTO - LEITURA"),
    Pct_Et10 = GetPct("10 - CONFERÊNCIA EQUIPAMENTO - LEITURA"),
    Pct_Et12 = GetPct("12 - TEMPO PADRÃO (CORRIDA) - LEITURA"),

    // ══════════════════════════════════════════════════════════════════════════
    // BLOCO A — MO BANCADA (etapas 1+2+3+6)
    // ══════════════════════════════════════════════════════════════════════════
    EtapasMOB = List.Intersect({
                    {"1 - PREPARO (PADRÃO E SOLUÇÕES) - BANCADA",
                     "2 - PREPARO AMOSTRA - BANCADA",
                     "3 - ANÁLISE FQ - BANCADA",
                     "6 - FINALIZAÇÃO DOS DADOS - BANCADA"},
                    ColsNaTabela}),

    MOB_Base  = Table.SelectColumns(TC_Banc_Flag,
                    ColsIDExist & FlagsExist & EtapasMOB),
    MOB_JoinD = JoinDem(MOB_Base),
    MOB_JoinC = JoinCamp(MOB_JoinD),

    MOB_CM = Table.AddColumn(MOB_JoinC, "CM_Min",
                 each
                   let
                       lotes      = ToNum([Demanda_Total]),
                       lotesXcamp = Number.RoundUp(ToNum([LotePorCampanha]), 0),
                       numMeses   = ToNum([Num_Meses]),
                       nSkip      = ToNum([N_Skip]),
                       ehSkip     = [EhSkip] = true,
                       nCamp      = if lotesXcamp > 0
                                    then Number.RoundUp(lotes / lotesXcamp, 0)
                                    else 0,
                       tc1 = ToNum(if List.Contains(EtapasMOB,
                                 "1 - PREPARO (PADRÃO E SOLUÇÕES) - BANCADA")
                             then Record.Field(_,
                                 "1 - PREPARO (PADRÃO E SOLUÇÕES) - BANCADA")
                             else null),
                       tc2 = ToNum(if List.Contains(EtapasMOB,
                                 "2 - PREPARO AMOSTRA - BANCADA")
                             then Record.Field(_,
                                 "2 - PREPARO AMOSTRA - BANCADA")
                             else null),
                       tc3 = ToNum(if List.Contains(EtapasMOB,
                                 "3 - ANÁLISE FQ - BANCADA")
                             then Record.Field(_,
                                 "3 - ANÁLISE FQ - BANCADA")
                             else null),
                       tc6 = ToNum(if List.Contains(EtapasMOB,
                                 "6 - FINALIZAÇÃO DOS DADOS - BANCADA")
                             then Record.Field(_,
                                 "6 - FINALIZAÇÃO DOS DADOS - BANCADA")
                             else null),
                       somaMOB = tc1 + tc2 + tc3 + tc6,
                       resultado =
                           if ehSkip then
                               CalcSkip(somaMOB, lotes, nSkip, numMeses)
                           else if nCamp = 0 then
                               if numMeses > 0
                               then somaMOB * lotes / numMeses
                               else 0
                           else if nCamp = 1 then
                               if numMeses > 0 and lotes > 0
                               then (somaMOB / lotes) * lotes / numMeses
                               else 0
                           else
                               let
                                   cm1 = if Pct_Et1 <> null then
                                             CalcCamp(tc1, lotes, lotesXcamp,
                                                      ToNum(Pct_Et1), numMeses)
                                         else if numMeses > 0 then
                                             tc1 * lotes / numMeses
                                         else 0,
                                   cm2 = if numMeses > 0
                                         then tc2 * lotes / numMeses
                                         else 0,
                                   cm3 = if numMeses > 0
                                         then tc3 * lotes / numMeses
                                         else 0,
                                   cm6 = if Pct_Et6 <> null then
                                             CalcCamp(tc6, lotes, lotesXcamp,
                                                      ToNum(Pct_Et6), numMeses)
                                         else if numMeses > 0 then
                                             tc6 * lotes / numMeses
                                         else 0
                               in
                                   cm1 + cm2 + cm3 + cm6
                   in
                       resultado,
                 type number),

    MOB_Grp   = Table.Group(MOB_CM,
                    {"Método e versão (DocNix)", "FAM proposta"}, {
                    {"CM_Min",      each List.Sum([CM_Min]),          type number},
                    {"Demanda_MPR", each List.Average([Demanda_MPR]), type number}}),
    MOB_Final = Table.AddColumn(
                    Table.TransformColumnTypes(MOB_Grp,
                        {{"FAM proposta", type text}}),
                    "PROCESSO", each "MO BANCADA", type text),

    // ══════════════════════════════════════════════════════════════════════════
    // BLOCO B — MO CONFERÊNCIA BANCADA e PI/TF (etapa 7)
    // ══════════════════════════════════════════════════════════════════════════
    CONF_Etapa  = "7 - CONFERÊNCIA DO PREPARO - BANCADA",
    CONF_Existe = List.Contains(ColsNaTabela, CONF_Etapa),

    CONF_Final =
        if not CONF_Existe then
            #table({"FAM proposta","PROCESSO","CM_Min","Demanda_MPR"}, {})
        else let
            B_Banc    = Table.SelectColumns(TC_Banc_Flag,
                            ColsIDExist & FlagsExist & {CONF_Etapa}),
            F_Banc    = Table.SelectRows(B_Banc,
                            each [#"FAM proposta"] <> "Fam PI/TF"
                              and ToNum(Record.Field(_, CONF_Etapa)) <> 0),
            JD_Banc   = JoinDem(F_Banc),
            JC_Banc   = JoinCamp(JD_Banc),
            CM_Banc   = Table.AddColumn(JC_Banc, "CM_Min",
                            each
                              let
                                  tc         = ToNum(Record.Field(_, CONF_Etapa)),
                                  lotes      = ToNum([Demanda_Total]),
                                  lotesXcamp = Number.RoundUp(
                                                   ToNum([LotePorCampanha]), 0),
                                  numMeses   = ToNum([Num_Meses]),
                                  nSkip      = ToNum([N_Skip]),
                                  ehSkip     = [EhSkip] = true,
                                  nCamp      = if lotesXcamp > 0
                                               then Number.RoundUp(
                                                        lotes / lotesXcamp, 0)
                                               else 0
                              in
                                  if ehSkip then
                                      CalcSkip(tc, lotes, nSkip, numMeses)
                                  else if nCamp = 0 then
                                      if numMeses > 0
                                      then tc * lotes / numMeses
                                      else 0
                                  else if nCamp = 1 then
                                      if numMeses > 0 and lotes > 0
                                      then (tc / lotes) * lotes / numMeses
                                      else 0
                                  else
                                      if Pct_Et7 <> null then
                                          CalcCamp(tc, lotes, lotesXcamp,
                                                   ToNum(Pct_Et7), numMeses)
                                      else if numMeses > 0 then
                                          tc * lotes / numMeses
                                      else 0,
                            type number),
            GRP_Banc  = Table.Group(CM_Banc,
                            {"Método e versão (DocNix)"}, {
                            {"CM_Min",      each List.Sum([CM_Min]),          type number},
                            {"Demanda_MPR", each List.Average([Demanda_MPR]), type number}}),
            FAM_Banc  = Table.AddColumn(GRP_Banc, "FAM proposta",
                            each "MO CONFERÊNCIA BANCADA", type text),
            PROC_Banc = Table.AddColumn(FAM_Banc, "PROCESSO",
                            each "MO CONFERÊNCIA BANCADA", type text),

            B_PITF    = Table.SelectColumns(TC_Banc_Flag,
                            ColsIDExist & FlagsExist & {CONF_Etapa}),
            F_PITF    = Table.SelectRows(B_PITF,
                            each [#"FAM proposta"] = "Fam PI/TF"
                              and ToNum(Record.Field(_, CONF_Etapa)) <> 0),
            JD_PITF   = JoinDem(F_PITF),
            JC_PITF   = JoinCamp(JD_PITF),
            CM_PITF   = Table.AddColumn(JC_PITF, "CM_Min",
                            each
                              let
                                  tc         = ToNum(Record.Field(_, CONF_Etapa)),
                                  lotes      = ToNum([Demanda_Total]),
                                  lotesXcamp = Number.RoundUp(
                                                   ToNum([LotePorCampanha]), 0),
                                  numMeses   = ToNum([Num_Meses]),
                                  nSkip      = ToNum([N_Skip]),
                                  ehSkip     = [EhSkip] = true,
                                  nCamp      = if lotesXcamp > 0
                                               then Number.RoundUp(
                                                        lotes / lotesXcamp, 0)
                                               else 0
                              in
                                  if ehSkip then
                                      CalcSkip(tc, lotes, nSkip, numMeses)
                                  else if nCamp = 0 then
                                      if numMeses > 0
                                      then tc * lotes / numMeses
                                      else 0
                                  else if nCamp = 1 then
                                      if numMeses > 0 and lotes > 0
                                      then (tc / lotes) * lotes / numMeses
                                      else 0
                                  else
                                      if Pct_Et7 <> null then
                                          CalcCamp(tc, lotes, lotesXcamp,
                                                   ToNum(Pct_Et7), numMeses)
                                      else if numMeses > 0 then
                                          tc * lotes / numMeses
                                      else 0,
                            type number),
            GRP_PITF  = Table.Group(CM_PITF,
                            {"Método e versão (DocNix)"}, {
                            {"CM_Min",      each List.Sum([CM_Min]),          type number},
                            {"Demanda_MPR", each List.Average([Demanda_MPR]), type number}}),
            FAM_PITF  = Table.AddColumn(GRP_PITF, "FAM proposta",
                            each "MO CONFERÊNCIA PI/TF", type text),
            PROC_PITF = Table.AddColumn(FAM_PITF, "PROCESSO",
                            each "MO CONFERÊNCIA PI/TF", type text)
        in
            Table.Combine({PROC_Banc, PROC_PITF}),

    // ══════════════════════════════════════════════════════════════════════════
    // BLOCO C — MO EQUIPAMENTO BANCADA (espectro)
    // ══════════════════════════════════════════════════════════════════════════
    ESPB_Final =
        if List.IsEmpty(ColsEspExist) then
            #table({"FAM proposta","PROCESSO","CM_Min","Demanda_MPR"}, {})
        else let
            B    = Table.AddColumn(TC_Esp_Flag, "TC_Unitario",
                       each
                           ToNum(if List.Contains(ColsEspExist, "SPCT0071 - AA")
                                 then Record.Field(_, "SPCT0071 - AA") else null)
                           +
                           ToNum(if List.Contains(ColsEspExist, "ICP")
                                 then Record.Field(_, "ICP") else null)
                           +
                           ToNum(if List.Contains(ColsEspExist, "SPCT - IV")
                                 then Record.Field(_, "SPCT - IV") else null)
                           +
                           ToNum(if List.Contains(ColsEspExist, "SPCT - UV")
                                 then Record.Field(_, "SPCT - UV") else null),
                       type number),
            F    = Table.SelectRows(B,
                       each [TC_Unitario] <> null and [TC_Unitario] <> 0),
            JD   = JoinDem(F),
            CM   = Table.AddColumn(JD, "CM_Min",
                       each
                         let
                             tc       = [TC_Unitario],
                             lotes    = ToNum([Demanda_Total]),
                             numMeses = ToNum([Num_Meses]),
                             nSkip    = ToNum([N_Skip]),
                             ehSkip   = [EhSkip] = true
                         in
                         if ehSkip then
                             CalcSkip(tc, lotes, nSkip, numMeses)
                         else if numMeses > 0
                         then tc * lotes / numMeses
                         else 0,
                       type number),
            GRP  = Table.Group(CM,
                       {"Método e versão (DocNix)"}, {
                       {"CM_Min",      each List.Sum([CM_Min]),          type number},
                       {"Demanda_MPR", each List.Average([Demanda_MPR]), type number}}),
            FAM  = Table.AddColumn(GRP, "FAM proposta",
                       each "MO EQUIPAMENTO BANCADA", type text),
            PROC = Table.AddColumn(FAM, "PROCESSO",
                       each "MO EQUIPAMENTO BANCADA", type text)
        in
            PROC,

    // ══════════════════════════════════════════════════════════════════════════
    // BLOCO D — MO EQUIPAMENTO (etapas 9+14)
    // ══════════════════════════════════════════════════════════════════════════
    Et9         = "9 - MONTAGEM DO EQUIPAMENTO - LEITURA",
    Et14        = "14 - PROCESSAMENTO DA CORRIDA - LEITURA",
    EtsMOEExist = List.Intersect({{Et9, Et14}, ColsNaTabela}),

    MOE_Final =
        if List.IsEmpty(EtsMOEExist) then
            #table({"FAM proposta","PROCESSO","CM_Min","Demanda_MPR"}, {})
        else let
            B    = Table.SelectColumns(TC_Equip_Flag,
                       ColsIDExist & FlagsExist & EtsMOEExist),
            UP   = Table.UnpivotOtherColumns(B,
                       ColsIDExist & FlagsExist,
                       "Etapa", "TC_Unitario"),
            F    = Table.SelectRows(UP,
                       each ToNum([TC_Unitario]) <> 0
                         and Text.From([TC_Unitario]) <> "S"),
            JD   = JoinDem(F),
            JC   = JoinCamp(JD),
            CM   = Table.AddColumn(JC, "CM_Min",
                       each
                         let
                             tc         = ToNum([TC_Unitario]),
                             lotes      = ToNum([Demanda_Total]),
                             lotesXcamp = Number.RoundUp(
                                              ToNum([LotePorCampanha]), 0),
                             numMeses   = ToNum([Num_Meses]),
                             nSkip      = ToNum([N_Skip]),
                             ehSkip     = [EhSkip] = true,
                             nCamp      = if lotesXcamp > 0
                                          then Number.RoundUp(
                                                   lotes / lotesXcamp, 0)
                                          else 0,
                             pct        = if [Etapa] = Et9
                                          then ToNum(Pct_Et9)
                                          else 0
                         in
                             if ehSkip then
                                 CalcSkip(tc, lotes, nSkip, numMeses)
                             else if nCamp = 0 then
                                 if numMeses > 0
                                 then tc * lotes / numMeses
                                 else 0
                             else if nCamp = 1 then
                                 if numMeses > 0 and lotes > 0
                                 then (tc / lotes) * lotes / numMeses
                                 else 0
                             else
                                 if pct > 0 then
                                     CalcCamp(tc, lotes, lotesXcamp,
                                              pct, numMeses)
                                 else if numMeses > 0 then
                                     tc * lotes / numMeses
                                 else 0,
                       type number),
            GRP  = Table.Group(CM,
                       {"Método e versão (DocNix)"}, {
                       {"CM_Min",      each List.Sum([CM_Min]),          type number},
                       {"Demanda_MPR", each List.Average([Demanda_MPR]), type number}}),
            FAM  = Table.AddColumn(GRP, "FAM proposta",
                       each "MO EQUIPAMENTO", type text),
            PROC = Table.AddColumn(FAM, "PROCESSO",
                       each "MO EQUIPAMENTO", type text)
        in
            PROC,

    // ══════════════════════════════════════════════════════════════════════════
    // BLOCO E — MO CONFERÊNCIA EQUIPAMENTO (etapas 10+15)
    // ══════════════════════════════════════════════════════════════════════════
    Et10          = "10 - CONFERÊNCIA EQUIPAMENTO - LEITURA",
    Et15          = "15 - CONFERÊNCIA PROCESSAMENTO DA CORRIDA",
    EtsCONFEExist = List.Intersect({{Et10, Et15}, ColsNaTabela}),

    CONFE_Final =
        if List.IsEmpty(EtsCONFEExist) then
            #table({"FAM proposta","PROCESSO","CM_Min","Demanda_MPR"}, {})
        else let
            B    = Table.SelectColumns(TC_Equip_Flag,
                       ColsIDExist & FlagsExist & EtsCONFEExist),
            UP   = Table.UnpivotOtherColumns(B,
                       ColsIDExist & FlagsExist,
                       "Etapa", "TC_Unitario"),
            F    = Table.SelectRows(UP,
                       each ToNum([TC_Unitario]) <> 0
                         and Text.From([TC_Unitario]) <> "S"),
            JD   = JoinDem(F),
            JC   = JoinCamp(JD),
            CM   = Table.AddColumn(JC, "CM_Min",
                       each
                         let
                             tc         = ToNum([TC_Unitario]),
                             lotes      = ToNum([Demanda_Total]),
                             lotesXcamp = Number.RoundUp(
                                              ToNum([LotePorCampanha]), 0),
                             numMeses   = ToNum([Num_Meses]),
                             nSkip      = ToNum([N_Skip]),
                             ehSkip     = [EhSkip] = true,
                             nCamp      = if lotesXcamp > 0
                                          then Number.RoundUp(
                                                   lotes / lotesXcamp, 0)
                                          else 0,
                             pct        = if [Etapa] = Et10
                                          then ToNum(Pct_Et10)
                                          else 0
                         in
                             if ehSkip then
                                 CalcSkip(tc, lotes, nSkip, numMeses)
                             else if nCamp = 0 then
                                 if numMeses > 0
                                 then tc * lotes / numMeses
                                 else 0
                             else if nCamp = 1 then
                                 if numMeses > 0 and lotes > 0
                                 then (tc / lotes) * lotes / numMeses
                                 else 0
                             else
                                 if pct > 0 then
                                     CalcCamp(tc, lotes, lotesXcamp,
                                              pct, numMeses)
                                 else if numMeses > 0 then
                                     tc * lotes / numMeses
                                 else 0,
                       type number),
            GRP  = Table.Group(CM,
                       {"Método e versão (DocNix)"}, {
                       {"CM_Min",      each List.Sum([CM_Min]),          type number},
                       {"Demanda_MPR", each List.Average([Demanda_MPR]), type number}}),
            FAM  = Table.AddColumn(GRP, "FAM proposta",
                       each "MO CONFERÊNCIA EQUIPAMENTO", type text),
            PROC = Table.AddColumn(FAM, "PROCESSO",
                       each "MO CONFERÊNCIA EQUIPAMENTO", type text)
        in
            PROC,

    // ══════════════════════════════════════════════════════════════════════════
    // BLOCO F — REVISÃO FINAL (etapa 16)
    // ══════════════════════════════════════════════════════════════════════════
    Et16 = "16 - REVISÃO LAUDO - REVISÃO",

    REV_Final =
        if not List.Contains(ColsNaTabela, Et16) then
            #table({"FAM proposta","PROCESSO","CM_Min","Demanda_MPR"}, {})
        else let
            B    = Table.SelectColumns(TC_Equip_Flag,
                       ColsIDExist & FlagsExist & {Et16}),
            F    = Table.SelectRows(B,
                       each ToNum(Record.Field(_, Et16)) <> 0),
            JD   = JoinDem(F),
            CM   = Table.AddColumn(JD, "CM_Min",
                       each
                         let
                             tc       = ToNum(Record.Field(_, Et16)),
                             lotes    = ToNum([Demanda_Total]),
                             numMeses = ToNum([Num_Meses]),
                             nSkip    = ToNum([N_Skip]),
                             ehSkip   = [EhSkip] = true
                         in
                         if ehSkip then
                             CalcSkip(tc, lotes, nSkip, numMeses)
                         else if numMeses > 0
                         then tc * lotes / numMeses
                         else 0,
                       type number),
            GRP  = Table.Group(CM,
                       {"Método e versão (DocNix)"}, {
                       {"CM_Min",      each List.Sum([CM_Min]),          type number},
                       {"Demanda_MPR", each List.Average([Demanda_MPR]), type number}}),
            FAM  = Table.AddColumn(GRP, "FAM proposta",
                       each "REVISÃO FINAL", type text),
            PROC = Table.AddColumn(FAM, "PROCESSO",
                       each "REVISÃO FINAL", type text)
        in
            PROC,

    // ══════════════════════════════════════════════════════════════════════════
    // BLOCO G — DISSOLUTOR (etapa 8)
    // ══════════════════════════════════════════════════════════════════════════
    Et8 = "8 - TEMPO DO DISSOLUTOR - BANCADA",

    DIS_Final =
        if not List.Contains(ColsNaTabela, Et8) then
            #table({"FAM proposta","PROCESSO","CM_Min","Demanda_MPR"}, {})
        else let
            B    = Table.SelectColumns(TC_Banc_Flag,
                       ColsIDExist & FlagsExist & {Et8}),
            F    = Table.SelectRows(B,
                       each ToNum(Record.Field(_, Et8)) <> 0),
            JD   = JoinDem(F),
            CM   = Table.AddColumn(JD, "CM_Min",
                       each
                         let
                             tc       = ToNum(Record.Field(_, Et8)),
                             lotes    = ToNum([Demanda_Total]),
                             numMeses = ToNum([Num_Meses]),
                             nSkip    = ToNum([N_Skip]),
                             ehSkip   = [EhSkip] = true
                         in
                         if ehSkip then
                             CalcSkip(tc, lotes, nSkip, numMeses)
                         else if numMeses > 0
                         then tc * lotes / numMeses
                         else 0,
                       type number),
            GRP  = Table.Group(CM,
                       {"Método e versão (DocNix)"}, {
                       {"CM_Min",      each List.Sum([CM_Min]),          type number},
                       {"Demanda_MPR", each List.Average([Demanda_MPR]), type number}}),
            FAM  = Table.AddColumn(GRP, "FAM proposta",
                       each "DISSOLUTOR", type text),
            PROC = Table.AddColumn(FAM, "PROCESSO",
                       each "DISSOLUTOR", type text)
        in
            PROC,

    // ══════════════════════════════════════════════════════════════════════════
    // BLOCO H — EQUIPAMENTO HPLC/CG
    // ══════════════════════════════════════════════════════════════════════════
    EQ_Final =
        if List.IsEmpty(EqCols) then
            #table({"FAM proposta","PROCESSO","CM_Min","Demanda_MPR"}, {})
        else let
            ColsOrig   = Table.ColumnNames(TC_Orig_Global),
            EqColsOrig = List.Intersect({EqCols, ColsOrig}),

            TC_Eq = Table.SelectRows(
                        Table.SelectColumns(TC_Orig_Global,
                            {"Método e versão (DocNix)",
                             "Equipamento Proposto",
                             "FAM proposta"} & EqColsOrig),
                        each let ep = [Equipamento Proposto]
                             in ep <> null and ep <> ""
                               and (Text.StartsWith(ep, "CRLQ")
                                 or Text.StartsWith(ep, "CRGS"))),

            JSkip = Table.NestedJoin(TC_Eq,
                        {"Método e versão (DocNix)"}, Skip_Dedup_Global,
                        {"MPR"}, "Skip", JoinKind.LeftOuter),
            ESkip = Table.ExpandTableColumn(JSkip, "Skip",
                        {"Qtde Lotes Skip"}, {"Qtde Lotes Skip"}),

            ComFlags = Table.AddColumn(
                           Table.AddColumn(ESkip,
                               "N_Skip",
                               each let v = [#"Qtde Lotes Skip"]
                                    in if v = null or v < 1 then 1 else v,
                               type number),
                               "EhSkip",
                               each [#"FAM proposta"] = "Impureza Skip Test"
                                 and [#"Qtde Lotes Skip"] <> null
                                 and [#"Qtde Lotes Skip"] >= 1,
                               type logical),

            // Buscar LotePorCampanha usando Camp_Dedup_Global
            JCamp = Table.NestedJoin(ComFlags,
                        {"Método e versão (DocNix)"}, Camp_Dedup_Global,
                        {"CÓD"}, "Camp3", JoinKind.LeftOuter),
            ECamp = Table.ExpandTableColumn(JCamp, "Camp3",
                        {"QTD LOTES CAMPANHA"}, {"LotePorCampanha"}),

            // Join com demanda direto — SEM usar JoinCamp para evitar duplicação
            JD = Table.ExpandTableColumn(
                     Table.NestedJoin(ECamp,
                         {"Método e versão (DocNix)"}, DEMANDA_MPR,
                         {"MÉTODO DOCNIX"}, "Dem", JoinKind.LeftOuter),
                     "Dem",
                     {"Demanda_Total", "Demanda_MPR", "Num_Meses"},
                     {"Demanda_Total", "Demanda_MPR", "Num_Meses"}),

            CM = Table.AddColumn(JD, "CM_Min",
                     each
                       let
                           nI12 = ToNum(if List.Contains(EqColsOrig, EqNInj12)
                                        then Record.Field(_, EqNInj12)
                                        else null),
                           tC12 = ToNum(if List.Contains(EqColsOrig, EqTCor12)
                                        then Record.Field(_, EqTCor12)
                                        else null),
                           tc12       = nI12 * tC12,
                           nI13 = ToNum(if List.Contains(EqColsOrig, EqNInj13)
                                        then Record.Field(_, EqNInj13)
                                        else null),
                           tC13 = ToNum(if List.Contains(EqColsOrig, EqTCor13)
                                        then Record.Field(_, EqTCor13)
                                        else null),
                           tc13       = nI13 * tC13,
                           somaEq     = tc12 + tc13,
                           lotes      = ToNum([Demanda_Total]),
                           lotesXcamp = Number.RoundUp(
                                            ToNum([LotePorCampanha]), 0),
                           numMeses   = ToNum([Num_Meses]),
                           nSkip      = ToNum([N_Skip]),
                           ehSkip     = [EhSkip] = true,
                           nCamp      = if lotesXcamp > 0
                                        then Number.RoundUp(
                                                 lotes / lotesXcamp, 0)
                                        else 0,
                           somaEqCamp =
                               if ehSkip then
                                   let nSkipCamp = if nSkip > 0
                                                   then Number.RoundUp(
                                                            lotes / nSkip, 0)
                                                   else 0
                                   in if lotes > 0
                                      then somaEq * nSkipCamp / lotes
                                      else 0
                               else if nCamp = 0 then somaEq
                               else if nCamp = 1 then
                                   if lotes > 0 then somaEq / lotes else 0
                               else
                                   let
                                       cc  = nCamp * tc12,
                                       cr  = (lotes - nCamp) * tc12
                                             * ToNum(Pct_Et12),
                                       t12 = if lotes > 0
                                             then (cc + cr) / lotes
                                             else 0
                                   in t12 + tc13,
                           resultado = if numMeses > 0
                                       then somaEqCamp * lotes / numMeses
                                       else 0
                       in
                           resultado,
                     type number),

            GRP  = Table.Group(CM,
                       {"Equipamento Proposto"}, {
                       {"CM_Min",      each List.Sum([CM_Min]),          type number},
                       {"Demanda_MPR", each List.Average([Demanda_MPR]), type number}}),
            PROC = Table.AddColumn(
                       Table.TransformColumnTypes(GRP,
                           {{"Equipamento Proposto", type text}}),
                       "PROCESSO",
                       each if Text.StartsWith([Equipamento Proposto], "CRLQ")
                            then "EQUIPAMENTO HPLC"
                            else if Text.StartsWith([Equipamento Proposto], "CRGS")
                            then "EQUIPAMENTO CG"
                            else "EQUIPAMENTO",
                       type text),
            REN  = Table.RenameColumns(PROC,
                       {{"Equipamento Proposto", "FAM proposta"}})
        in
            REN,

    // ══════════════════════════════════════════════════════════════════════════
    // COMBINAR TODOS OS BLOCOS
    // ══════════════════════════════════════════════════════════════════════════
    ColsSaida = {"FAM proposta", "PROCESSO", "CM_Min", "Demanda_MPR"},

    Padrao = (tbl as table) =>
        let
            Tipado      = Table.TransformColumnTypes(tbl, {
                              {"FAM proposta", type text},
                              {"PROCESSO",     type text},
                              {"CM_Min",       type number},
                              {"Demanda_MPR",  type number}}),
            Selecionado = Table.SelectColumns(Tipado, ColsSaida)
        in
            Selecionado,

    Final = Table.Combine({
                Padrao(MOB_Final),
                Padrao(CONF_Final),
                Padrao(ESPB_Final),
                Padrao(MOE_Final),
                Padrao(CONFE_Final),
                Padrao(REV_Final),
                Padrao(DIS_Final),
                Padrao(EQ_Final)
            })

in
    Final;
