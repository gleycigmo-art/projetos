section Section1;

shared FAM_PROPOSTA = let
    Fonte = Excel.CurrentWorkbook(){[Name="Tabela_FAM"]}[Content],
    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte,{
        {"PROCESSO", type text}, {"TURNOS", Int64.Type},
        {"QTDE 1° TURNO", type number}, {"QTDE 2° TURNO", type number},
        {"QTDE 3° TURNO", type number}, {"RECURSO MÉDIO ATUAL", type number},
        {"HRS TURNO", Int64.Type}, {"TOTAL RECURSO", type number}, {"OEE REAL", type number}})
in
    #"Tipo Alterado";

shared FILTRO_PERIODO = let
    Tabela = Excel.CurrentWorkbook(){[Name="Tabela_FILTRO"]}[Content],

    // Sem "as text" no retorno — deixa o valor no tipo que vier da célula
    ObterValor = (nomeParametro as text) as any =>
        Table.SelectRows(Tabela, each [Parametro] = nomeParametro){0}[Valor],

    NivelDeVisao = Text.From(ObterValor("Nível de Visão")),
    MesInicio    = Text.From(ObterValor("Mês Início")),
    MesFim       = Text.From(ObterValor("Mês Fim")),
    SemanaChave  = Number.From(ObterValor("Semana (Chave)")),

    NomesColunas = {"Jan/2026","Fev/2026","Mar/2026","Abr/2026","Mai/2026","Jun/2026",
                    "Jul/2026","Ago/2026","Set/2026","Out/2026","Nov/2026","Dez/2026"},

    IdxInicio = List.PositionOf(NomesColunas, MesInicio),
    IdxFim    = List.PositionOf(NomesColunas, MesFim),

    MesesAtivos = List.Range(NomesColunas, IdxInicio, IdxFim - IdxInicio + 1),

    Resultado = [
        NivelDeVisao = NivelDeVisao,
        MesesAtivos  = MesesAtivos,
        SemanaChave  = SemanaChave
    ]
in
    Resultado;

shared SRC_DEMANDA = let
    Fonte = Excel.CurrentWorkbook(){[Name="Tabela_DEMANDA"]}[Content]
in
    Fonte;

shared DEMANDA_MPR = let
    NivelDeVisao = FILTRO_PERIODO[NivelDeVisao],
    // ===================== RAMO MENSAL =====================
    ResultadoMensal =
        let
            Fonte       = SRC_DEMANDA,
            MesesAtivos = FILTRO_PERIODO[MesesAtivos],
            NumPeriodos = List.Count(MesesAtivos),
            ColsFixas   = {"CONSIDERAR?", "MÉTODO DOCNIX", "CÓD"},
            Filtrado    = Table.SelectRows(Fonte, each [#"CONSIDERAR?"] = "SIM"),
            ColsSel     = Table.SelectColumns(Filtrado, ColsFixas & MesesAtivos),
            Unpivot     = Table.UnpivotOtherColumns(ColsSel, ColsFixas, "Mês", "Qtde"),
            UnpivotNum  = Table.TransformColumnTypes(Unpivot, {{"Qtde", type number}}),
            GrpSKU      = Table.Group(UnpivotNum, {"MÉTODO DOCNIX", "CÓD"},
                              {{"Qtde_SKU", each List.Sum([Qtde]), type number}}),
            GrpMPR      = Table.Group(GrpSKU, {"MÉTODO DOCNIX"},
                              {{"Demanda_Total", each List.Sum([Qtde_SKU]), type number}}),
            ComMedia    = Table.AddColumn(GrpMPR, "Demanda_MPR", each [Demanda_Total] / NumPeriodos, type number),
            ComPeriodos = Table.AddColumn(ComMedia, "Num_Meses", each NumPeriodos, type number)
        in
            ComPeriodos,
    // ===================== RAMO SEMANAL (inalterado) =====================
    ResultadoSemanal =
        let
            SemanaChave = FILTRO_PERIODO[SemanaChave],
            FonteTyped  = Table.TransformColumnTypes(Tabela_DemandaSemanal, {{"Chave semana", type number}}),
            Filtrado    = Table.SelectRows(FonteTyped, each [#"Chave semana"] = SemanaChave),
            FiltradoNum = Table.TransformColumnTypes(Filtrado, {{"Lote semana", type number}}),
            GrpSKU      = Table.Group(FiltradoNum, {"MÉTODO DOCNIX", "Cod"},
                              {{"Qtde_SKU", each List.Sum([Lote semana]), type number}}),
            GrpMPR      = Table.Group(GrpSKU, {"MÉTODO DOCNIX"},
                              {{"Demanda_Total", each List.Sum([Qtde_SKU]), type number}}),
            ComMedia    = Table.AddColumn(GrpMPR, "Demanda_MPR", each [Demanda_Total], type number),
            ComPeriodos = Table.AddColumn(ComMedia, "Num_Meses", each 1, type number)
        in
            ComPeriodos,
    Resultado = if NivelDeVisao = "Semanal" then ResultadoSemanal else ResultadoMensal
in
    Resultado;

shared BASE_TC_CAMP = let
    TC    = Excel.CurrentWorkbook(){[Name="Tabela_BASE_TC"]}[Content],
    Camp  = Excel.CurrentWorkbook(){[Name="Tabela_BD_CAMPANHA"]}[Content],
    Skip  = Excel.CurrentWorkbook(){[Name="Tabela_SKIP_IMP"]}[Content],

    J1    = Table.NestedJoin(TC,
                {"Método e versão (DocNix)"}, Camp,
                {"CÓD"}, "Camp", JoinKind.LeftOuter),
    E1    = Table.ExpandTableColumn(J1, "Camp",
                {"QTD LOTES CAMPANHA"}, {"Lotes_Campanha"}),
    J2    = Table.NestedJoin(E1,
                {"Método e versão (DocNix)"}, Skip,
                {"MPR"}, "Skip", JoinKind.LeftOuter),
    E2    = Table.ExpandTableColumn(J2, "Skip",
                {"Qtde Lotes Skip"}, {"Lotes_Skip"}),
    N1    = Table.AddColumn(E2, "N_Camp",
                each if [Lotes_Campanha] = null or [Lotes_Campanha] < 1
                     then 1 else [Lotes_Campanha], type number),
    N2    = Table.AddColumn(N1, "N_Skip",
                each if [Lotes_Skip] = null or [Lotes_Skip] < 1
                     then 1 else [Lotes_Skip], type number),
    N3    = Table.AddColumn(N2, "EhSkip",
                each [#"FAM proposta"] = "Impureza Skip Test",
                type logical)
in
    N3;

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

shared METODOS_EQUIPAMENTO = let

    // ── Fontes via FONTES_TC — sem Excel.CurrentWorkbook() aqui ──────────────
    TC_Orig    = FONTES_TC[TC_Orig],
    Skip_Dedup = FONTES_TC[Skip_Dedup],
    Camp_Dedup = FONTES_TC[Camp_Dedup],
    Lotes_Camp = FONTES_TC[Lotes_Camp],

    ToNum = (v as any) as number =>
        if v = null then 0
        else if v = "" then 0
        else if Value.Is(v, type number) then v
        else let n = try Number.From(v) otherwise 0 in n,

    EqNInj12 = "12 - Nº INJEÇÕES PADRÃO - LEITURA",
    EqTCor12 = "12 - TEMPO PADRÃO (CORRIDA) - LEITURA",
    EqNInj13 = "13 - Nº INJEÇÕES AMOSTRA/ DILUENTE/ PLACEBO - LEITURA",
    EqTCor13 = "13 - TEMPO AMOSTRA/ DILUENTE/ PLACEBO (CORRIDA) - LEITURA",
    EqCols   = {EqNInj12, EqTCor12, EqNInj13, EqTCor13},

    ColsOrig   = Table.ColumnNames(TC_Orig),
    EqColsOrig = List.Intersect({EqCols, ColsOrig}),

    // ── Selecionar linhas com Equipamento Proposto CRLQ ou CRGS ──────────────
    TC_Eq = Table.SelectRows(
                Table.SelectColumns(TC_Orig,
                    {"Método e versão (DocNix)",
                     "Equipamento Proposto",
                     "FAM proposta"} & EqColsOrig),
                each let ep = [Equipamento Proposto]
                     in ep <> null and ep <> ""
                       and (Text.StartsWith(ep, "CRLQ")
                         or Text.StartsWith(ep, "CRGS"))),

    // ── Flags skip ────────────────────────────────────────────────────────────
    JSkip = Table.NestedJoin(TC_Eq,
                {"Método e versão (DocNix)"}, Skip_Dedup,
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

    // ── LotePorCampanha ───────────────────────────────────────────────────────
    JCamp = Table.NestedJoin(ComFlags,
                {"Método e versão (DocNix)"}, Camp_Dedup,
                {"CÓD"}, "Camp3", JoinKind.LeftOuter),
    ECamp = Table.ExpandTableColumn(JCamp, "Camp3",
                {"QTD LOTES CAMPANHA"}, {"LotePorCampanha"}),

    // ── Demanda ───────────────────────────────────────────────────────────────
    JD = Table.NestedJoin(ECamp,
             {"Método e versão (DocNix)"}, DEMANDA_MPR,
             {"MÉTODO DOCNIX"}, "Dem", JoinKind.LeftOuter),
    ED = Table.ExpandTableColumn(JD, "Dem",
             {"Demanda_Total", "Demanda_MPR", "Num_Meses"},
             {"Demanda_Total", "Demanda_MPR", "Num_Meses"}),

    // ── Pct etapa 12 ──────────────────────────────────────────────────────────
    Pct_Et12 = let r = Table.SelectRows(Lotes_Camp,
                           each [Etapa] = "12 - TEMPO PADRÃO (CORRIDA) - LEITURA_CAMPANHA")
               in if Table.RowCount(r) > 0
                  then r{0}[#"% tempo"] else 0,

    // ── CM_Min por linha — mesmo racional do bloco H ──────────────────────────
    ComCM = Table.AddColumn(ED, "CM_Min",
                each
                  let
                      nI12 = ToNum(if List.Contains(EqColsOrig, EqNInj12)
                                   then Record.Field(_, EqNInj12) else null),
                      tC12 = ToNum(if List.Contains(EqColsOrig, EqTCor12)
                                   then Record.Field(_, EqTCor12) else null),
                      tc12       = nI12 * tC12,
                      nI13 = ToNum(if List.Contains(EqColsOrig, EqNInj13)
                                   then Record.Field(_, EqNInj13) else null),
                      tC13 = ToNum(if List.Contains(EqColsOrig, EqTCor13)
                                   then Record.Field(_, EqTCor13) else null),
                      tc13       = nI13 * tC13,
                      somaEq     = tc12 + tc13,
                      lotes      = ToNum([Demanda_Total]),
                      lotesXcamp = Number.RoundUp(ToNum([LotePorCampanha]), 0),
                      numMeses   = ToNum([Num_Meses]),
                      nSkip      = ToNum([N_Skip]),
                      ehSkip     = [EhSkip] = true,
                      nCamp      = if lotesXcamp > 0
                                   then Number.RoundUp(lotes / lotesXcamp, 0)
                                   else 0,
                      somaEqCamp =
                          if ehSkip then
                              let nSkipCamp = if nSkip > 0
                                              then Number.RoundUp(lotes / nSkip, 0)
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
                                  cr  = (lotes - nCamp) * tc12 * ToNum(Pct_Et12),
                                  t12 = if lotes > 0 then (cc + cr) / lotes else 0
                              in t12 + tc13
                  in
                      if numMeses > 0 then somaEqCamp * lotes / numMeses else 0,
                type number)

in
    // Expor nível de método — sem agrupar
    Table.SelectColumns(ComCM, {
        "Método e versão (DocNix)",
        "FAM proposta",
        "Equipamento Proposto",
        "CM_Min"});

shared CAPACIDADE_EQUIPAMENTO = let
    CAPACIDADE_EQUIPAMENTO = let

    // Buscar dados de capacidade da FAM_PROPOSTA
    Base = FAM_PROPOSTA,

    // Filtrar apenas equipamentos CRLQ e CRGS
    Filt = Table.SelectRows(Base,
               each Text.StartsWith([FAMILIA], "CRLQ")
                 or Text.StartsWith([FAMILIA], "CRGS")),

    // Usar RECURSO DISP (MIN) diretamente como Capacidade_Min
    ComCap = Table.RenameColumns(Filt,
                 {{"RECURSO DISP (MIN)", "Capacidade_Min"}}),

    // Adicionar tipo de equipamento
    ComTipo = Table.AddColumn(ComCap, "Tipo",
                  each if Text.StartsWith([FAMILIA], "CRLQ")
                       then "HPLC"
                       else "CG",
                  type text),

    // Selecionar colunas finais
    Final = Table.SelectColumns(ComTipo, {
                "FAMILIA", "PROCESSO", "CLUSTER", "Tipo",
                "RECURSO MÉDIO ATUAL", "HRS TURNO",
                "TOTAL RECURSO", "OEE REAL",
                "Capacidade_Min"}),

    // Renomear para facilitar uso no Lovable
    Renomeado = Table.RenameColumns(Final, {
                    {"FAMILIA",             "Equipamento"},
                    {"PROCESSO",            "Processo"},
                    {"RECURSO MÉDIO ATUAL", "Recursos"},
                    {"HRS TURNO",           "Hrs_Turno"},
                    {"TOTAL RECURSO",       "Total_Recurso_Hrs"},
                    {"OEE REAL",            "OEE_Real"}})
in
    Table.Sort(Renomeado, {{"Equipamento", Order.Ascending}}),
    #"Tipo Alterado" = Table.TransformColumnTypes(CAPACIDADE_EQUIPAMENTO,{{"Capacidade_Min", type number}})
in
    #"Tipo Alterado";

shared FONTES_TC = let
    Camp_Dedup = Table.Distinct(
                     Excel.CurrentWorkbook(){[Name="Tabela_BD_CAMPANHA"]}[Content],
                     {"CÓD"}),
    Skip_Dedup = Table.Distinct(
                     Excel.CurrentWorkbook(){[Name="Tabela_SKIP_IMP"]}[Content],
                     {"MPR"}),
    Lotes_Camp = Excel.CurrentWorkbook(){[Name="Tabela_Lotes_Campanha"]}[Content],
    TC_Orig    = Excel.CurrentWorkbook(){[Name="Tabela_BASE_TC"]}[Content],
    FamEquip   = Excel.CurrentWorkbook(){[Name="FAMILIAS_EQUIPAMENTO"]}[Content]
in
    [
        Camp_Dedup = Camp_Dedup,
        Skip_Dedup = Skip_Dedup,
        Lotes_Camp = Lotes_Camp,
        TC_Orig    = TC_Orig,
        FamEquip   = FamEquip
    ];

shared Tabela_DemandaSemanal = let
    URLSite      = "https://hypermarcas.sharepoint.com/sites/ControledeQualidade-PA",
    CaminhoPasta = URLSite & "/Documentos Compartilhados/MFV/",
    NomeArquivo  = "Base Quarentena tratada.xlsx",
    NomeAba      = "Tbela_DemandaSemanal_futura",

    Conteudo =
        SharePoint.Files(URLSite, [ApiVersion = 15])
            {[Name = NomeArquivo, #"Folder Path" = CaminhoPasta]}[Content],

    Objetos = Excel.Workbook(Binary.Buffer(Conteudo), null, false),

    Filtrado =
        Table.SelectRows(
            Objetos,
            each [Kind] = "Sheet"
                 and Comparer.OrdinalIgnoreCase(Text.Trim([Item]), NomeAba) = 0
        ),

    Aba =
        if Table.IsEmpty(Filtrado)
        then error "Aba '" & NomeAba & "' não encontrada em " & NomeArquivo
        else Filtrado{0}[Data],

    Promovido = Table.PromoteHeaders(Aba, [PromoteAllScalars = true]),

    #"Linhas Válidas" =
        Table.SelectRows(
            Promovido,
            each [Cod] <> null and [Cod] <> "" and [#"Chave semana"] <> null
        ),

    #"Tipo Alterado" =
        Table.TransformColumnTypes(
            #"Linhas Válidas",
            {{"Cod",             type text},
             {"Chave semana",    Int64.Type},
             {"STATUS",          type text},
             {"Lote semana",     Int64.Type},
             {"MÉTODO DOCNIX",   type text},
             {"TC levantados?",  type text}}
        ),

    #"Status2 Adicionado" =
        Table.AddColumn(
            #"Tipo Alterado",
            "Status2",
            each
                if [MÉTODO DOCNIX] = "Método não encontrado" then "Sem tempos"
                else if [MÉTODO DOCNIX] = null then "Sem tempos"
                else if [#"TC levantados?"] = "NÃO" then "Sem tempos"
                else "Tempos OK",
            type text
        ),

    #"Colunas Selecionadas" =
        Table.SelectColumns(
            #"Status2 Adicionado",
            {"Cod", "Chave semana", "STATUS", "Lote semana",
             "MÉTODO DOCNIX", "TC levantados?", "Status2"}
        )
in
    #"Colunas Selecionadas";

shared Tabela_DEMANDA = let
    Fonte = Excel.CurrentWorkbook(){[Name="Tabela_DEMANDA"]}[Content],
    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte,{{"CONSIDERAR?", type text}, {"TIPO", type text}, {"SITE", type text}, {"CHECK DOC", type text}, {"CÓD", type text}, {"DESCRIÇÃO", type text}, {"MÉTODO DOCNIX", type text}, {"Jan/2026", type any}, {"Fev/2026", type any}, {"Mar/2026", type any}, {"Abr/2026", type any}, {"Mai/2026", type any}, {"Jun/2026", type any}, {"Jul/2026", Int64.Type}, {"Ago/2026", Int64.Type}, {"Set/2026", type number}, {"Out/2026", type number}, {"Nov/2026", type number}, {"Dez/2026", type any}, {"Soma", type number}, {"Média", type number}, {"TC levantados?", type text}, {"OBS", type any}, {"Local de Análise", type text}})
in
    #"Tipo Alterado";

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

shared #"SOMA TEMPOS TOTAIS" = let
    Fonte = Excel.CurrentWorkbook(){[Name="Tabela_BASE_TC"]}[Content],
    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte,{{"ID", Int64.Type}, {"Método e versão (DocNix)", type text}, {"Versão (DocNix)", type any}, {"Ativo", type text}, {"Teste original", type text}, {"Descrição Metodo", type text}, {"Teste Grupo", type text}, {"Sigla teste", type text}, {"Equipamento", type text}, {"Ativos", type any}, {"FAM proposta", type text}, {"SPCT0071 - AA", type any}, {"ICP", type any}, {"SPCT - IV", Int64.Type}, {"SPCT - UV", Int64.Type}, {"1 - PREPARO (PADRÃO E SOLUÇÕES) - BANCADA", Int64.Type}, {"2 - PREPARO AMOSTRA - BANCADA", Int64.Type}, {"3 - ANÁLISE FQ - BANCADA", Int64.Type}, {"4 - TEMPO DE REPOUSO - BANCADA", Int64.Type}, {"5 - TEMPOS DE INTRUMENTOS - BANCADA", Int64.Type}, {"6 - FINALIZAÇÃO DOS DADOS - BANCADA", Int64.Type}, {"7 - CONFERÊNCIA DO PREPARO - BANCADA", Int64.Type}, {"8 - TEMPO DO DISSOLUTOR - BANCADA", Int64.Type}, {"Equipamento Principal Original", type text}, {"CRLQ0198", type text}, {"CRLQ0200", type text}, {"CRLQ0201", type text}, {"CRLQ0202", type text}, {"CRLQ0204", type text}, {"CRLQ0210", type text}, {"CRLQ0150", type text}, {"CRLQ0254", type text}, {"CRLQ0256", type any}, {"CRLQ0143", type text}, {"CRLQ0144", type text}, {"CRLQ0145", type text}, {"CRLQ0024", type text}, {"CRLQ0043", type text}, {"CRLQ0055", type text}, {"CRLQ0064", type text}, {"CRLQ0079", type text}, {"CRLQ0147", type text}, {"CRLQ0148", type text}, {"CRLQ0151", type text}, {"CRLQ0188", type text}, {"CRLQ0081", type text}, {"CRLQ0085", type text}, {"CRLQ0086", type text}, {"CRLQ0250", type text}, {"CRLQ0251", type text}, {"CRLQ0252", type text}, {"CRLQ0033", type text}, {"CRLQ0034", type text}, {"CRLQ0066", type text}, {"CRLQ0093", type text}, {"CRLQ0094", type text}, {"CRLQ0127", type text}, {"CRLQ0132", type text}, {"CRLQ0131", type text}, {"CRLQ0072", type text}, {"CRGS0043", type any}, {"Equipamento Proposto", type text}, {"Equipe Equipamento", type text}, {"Check_Equip", Int64.Type}, {"Check_Equip_P", Int64.Type}, {"9 - MONTAGEM DO EQUIPAMENTO - LEITURA", Int64.Type}, {"10 - CONFERÊNCIA EQUIPAMENTO - LEITURA", Int64.Type}, {"11 - TEMPO CONDICIONAMENTO - LEITURA", Int64.Type}, {"12 - Nº INJEÇÕES PADRÃO - LEITURA", Int64.Type}, {"12 - TEMPO PADRÃO (CORRIDA) - LEITURA", type number}, {"13 - Nº INJEÇÕES AMOSTRA/ DILUENTE/ PLACEBO - LEITURA", Int64.Type}, {"13 - TEMPO AMOSTRA/ DILUENTE/ PLACEBO (CORRIDA) - LEITURA", type number}, {"14 - PROCESSAMENTO DA CORRIDA - LEITURA", Int64.Type}, {"15 - CONFERÊNCIA PROCESSAMENTO DA CORRIDA", Int64.Type}, {"16 - REVISÃO LAUDO - REVISÃO", Int64.Type}, {"obs", type text}, {"LOTES PROGRAMADOS", type number}, {"LOTE POR CAMPANHA", Int64.Type}, {"Qtde Campanha", Int64.Type}, {"SOMA MO BANCADA", Int64.Type}, {"1 - PREPARO (PADRÃO E SOLUÇÕES) - BANCADA_CAMPANHA", type number}, {"6 - FINALIZAÇÃO DOS DADOS - BANCADA_CAMPANHA", type number}, {"SOMA MO BANCADA_CAMPANHA", type number}, {"SOMA MO BANCADA_TCxDEM", type number}, {"SOMA MO CONFERÊNCIA BANCADA", Int64.Type}, {"SOMA MO CONFERÊNCIA PI/TF", Int64.Type}, {"7 - CONFERÊNCIA DO PREPARO - BANCADA_CAMPANHA", type number}, {"SOMA MO CONFERÊNCIA BANCADA_CAMPANHA", type number}, {"SOMA MO CONFERÊNCIA PI/TF_CAMPANHA", type number}, {"SOMA MO CONFERÊNCIA BANCADA_TCxDEM", type number}, {"SOMA MO CONFERÊNCIA PI/TF_TCxDEM", type number}, {"SOMA MO CONFERÊNCIA EQUIPAMENTO", Int64.Type}, {"10 - CONFERÊNCIA EQUIPAMENTO - LEITURA_CAMPANHA", Int64.Type}, {"SOMA MO CONFERÊNCIA EQUIPAMENTO_CAMPANHA", Int64.Type}, {"SOMA MO CONFERÊNCIA EQUIPAMENTO_TCxDEM", type number}, {"SOMA REVISÃO FINAL_TCxDEM", type number}, {"SOMA MO EQUIPAMENTO BANCADA_TCxDEM", Int64.Type}, {"SOMA DISSOLUTOR_TCxDEM", type number}, {"SOMA EQUIPAMENTO", type number}, {"12 - TEMPO PADRÃO (CORRIDA) - LEITURA_CAMPANHA", type number}, {"SOMA EQUIPAMENTO_CAMPANHA", type number}, {"SOMA EQUIPAMENTO_TCxDEM", type number}, {"SOMA MO EQUIPAMENTO", Int64.Type}, {"9 - MONTAGEM DO EQUIPAMENTO - LEITURA_CAMPANHA", type number}, {"SOMA MO EQUIPAMENTO_CAMPANHA", type number}, {"SOMA MO EQUIPAMENTO_TCxDEM", type number}}),

    #"Personalização Adicionada" = Table.AddColumn(#"Tipo Alterado", "12 - TEMPO PADRÃO", each [#"12 - Nº INJEÇÕES PADRÃO - LEITURA"]*[#"12 - TEMPO PADRÃO (CORRIDA) - LEITURA"]),

    #"Outras Colunas Removidas" = Table.SelectColumns(#"Personalização Adicionada",{"Método e versão (DocNix)", "Descrição Metodo", "Teste original", "SPCT0071 - AA", "ICP", "SPCT - IV", "SPCT - UV", "1 - PREPARO (PADRÃO E SOLUÇÕES) - BANCADA", "2 - PREPARO AMOSTRA - BANCADA", "3 - ANÁLISE FQ - BANCADA", "4 - TEMPO DE REPOUSO - BANCADA", "5 - TEMPOS DE INTRUMENTOS - BANCADA", "6 - FINALIZAÇÃO DOS DADOS - BANCADA", "7 - CONFERÊNCIA DO PREPARO - BANCADA", "8 - TEMPO DO DISSOLUTOR - BANCADA", "9 - MONTAGEM DO EQUIPAMENTO - LEITURA", "10 - CONFERÊNCIA EQUIPAMENTO - LEITURA", "11 - TEMPO CONDICIONAMENTO - LEITURA", "12 - Nº INJEÇÕES PADRÃO - LEITURA", "12 - TEMPO PADRÃO (CORRIDA) - LEITURA", "13 - Nº INJEÇÕES AMOSTRA/ DILUENTE/ PLACEBO - LEITURA", "13 - TEMPO AMOSTRA/ DILUENTE/ PLACEBO (CORRIDA) - LEITURA", "14 - PROCESSAMENTO DA CORRIDA - LEITURA", "15 - CONFERÊNCIA PROCESSAMENTO DA CORRIDA", "16 - REVISÃO LAUDO - REVISÃO", "12 - TEMPO PADRÃO"}),

    // ── guarda a base antes do filtro, para não perder métodos ────────────────
    BaseBuf = Table.Buffer(#"Outras Colunas Removidas"),

    // ── descrição vem da base COMPLETA, não do agrupamento ───────────────────
    TodosMetodos =
        Table.Group(
            Table.SelectColumns(BaseBuf, {"Método e versão (DocNix)", "Descrição Metodo"}),
            {"Método e versão (DocNix)"},
            {
                {"Descrição Metodo",
                 each List.First(List.RemoveNulls([Descrição Metodo]), null),
                 type text}
            }
        ),

    MetodosCheck =
        List.Buffer(
            List.Distinct(
                Table.SelectRows(
                    BaseBuf,
                    each [Teste original] <> null
                         and Text.StartsWith([Teste original], "Check")
                )[#"Método e versão (DocNix)"]
            )
        ),

    #"Linhas Filtradas" = Table.SelectRows(BaseBuf, each [Teste original] = null or not Text.StartsWith([Teste original], "Check")),

    #"Personalização Adicionada1" = Table.AddColumn(#"Linhas Filtradas", "13 - TEMPO AMOSTRA", each [#"13 - Nº INJEÇÕES AMOSTRA/ DILUENTE/ PLACEBO - LEITURA"]*[#"13 - TEMPO AMOSTRA/ DILUENTE/ PLACEBO (CORRIDA) - LEITURA"]),

    #"Colunas Removidas" = Table.RemoveColumns(#"Personalização Adicionada1",{"12 - Nº INJEÇÕES PADRÃO - LEITURA", "12 - TEMPO PADRÃO (CORRIDA) - LEITURA", "13 - Nº INJEÇÕES AMOSTRA/ DILUENTE/ PLACEBO - LEITURA", "13 - TEMPO AMOSTRA/ DILUENTE/ PLACEBO (CORRIDA) - LEITURA"}),

    #"Tipo Alterado2" = Table.TransformColumnTypes(#"Colunas Removidas",{{"SPCT0071 - AA", Int64.Type}, {"ICP", Int64.Type}, {"13 - TEMPO AMOSTRA", Int64.Type}, {"12 - TEMPO PADRÃO", Int64.Type}}),

    #"Personalização Adicionada2" = Table.AddColumn(#"Tipo Alterado2", "SOMA BANCADA", each List.Sum({
    [#"SPCT0071 - AA"],
    [ICP],
    [#"SPCT - IV"],
    [#"SPCT - UV"],
    [#"1 - PREPARO (PADRÃO E SOLUÇÕES) - BANCADA"],
    [#"2 - PREPARO AMOSTRA - BANCADA"],
    [#"3 - ANÁLISE FQ - BANCADA"],
    [#"5 - TEMPOS DE INTRUMENTOS - BANCADA"],
    [#"6 - FINALIZAÇÃO DOS DADOS - BANCADA"],
    [#"7 - CONFERÊNCIA DO PREPARO - BANCADA"],
    [#"8 - TEMPO DO DISSOLUTOR - BANCADA"]
}), type number),

    #"Personalização Adicionada3" = Table.AddColumn(#"Personalização Adicionada2", "SOMA EQUIPAMENTO", each List.Sum({
    [#"9 - MONTAGEM DO EQUIPAMENTO - LEITURA"],
    [#"10 - CONFERÊNCIA EQUIPAMENTO - LEITURA"],
    [#"11 - TEMPO CONDICIONAMENTO - LEITURA"],
    [#"14 - PROCESSAMENTO DA CORRIDA - LEITURA"],
    [#"15 - CONFERÊNCIA PROCESSAMENTO DA CORRIDA"],
    [#"12 - TEMPO PADRÃO"],
    [#"13 - TEMPO AMOSTRA"]
}), type number),

    #"Personalização Adicionada4" = Table.AddColumn(#"Personalização Adicionada3", "TEMPO TOTAL", each List.Sum({
    [#"SOMA BANCADA"],
    [#"SOMA EQUIPAMENTO"],
    [#"4 - TEMPO DE REPOUSO - BANCADA"],
    [#"16 - REVISÃO LAUDO - REVISÃO"]
}), type number),

    #"Linhas Agrupadas" = Table.Group(
    #"Personalização Adicionada4",
    {"Método e versão (DocNix)"},
    {
        {"TOTAL SOMA BANCADA", each List.Sum([SOMA BANCADA]), type number},
        {"TOTAL SOMA EQUIPAMENTO", each List.Sum([SOMA EQUIPAMENTO]), type number},
        {"TOTAL GERAL", each List.Sum([TEMPO TOTAL]), type number}
    }
),

    // ── reincorpora métodos que só tinham linhas Check ────────────────────────
    ComTodos =
        Table.ExpandTableColumn(
            Table.NestedJoin(
                TodosMetodos,
                {"Método e versão (DocNix)"},
                #"Linhas Agrupadas",
                {"Método e versão (DocNix)"},
                "Tot",
                JoinKind.LeftOuter
            ),
            "Tot",
            {"TOTAL SOMA BANCADA", "TOTAL SOMA EQUIPAMENTO", "TOTAL GERAL"},
            {"TOTAL SOMA BANCADA", "TOTAL SOMA EQUIPAMENTO", "TOTAL GERAL"}
        ),

    // ── marca os métodos Check ────────────────────────────────────────────────
    ComCheck =
        Table.AddColumn(ComTodos, "EhCheck",
            each List.Contains(MetodosCheck, [#"Método e versão (DocNix)"]),
            type logical),

    // ── zera os tempos dos métodos Check ──────────────────────────────────────
    Zerado =
        Table.ReplaceValue(
            ComCheck,
            each [EhCheck],
            each 0,
            (valorAtual, ehCheck, novo) => if ehCheck then novo else valorAtual,
            {"TOTAL SOMA BANCADA", "TOTAL SOMA EQUIPAMENTO", "TOTAL GERAL"}
        ),

    // ── coluna sinalizadora Sim/Não ───────────────────────────────────────────
    ComFlag =
        Table.AddColumn(Zerado, "Check doc",
            each if [EhCheck] then "Sim" else "Não",
            type text),

    Final =
        Table.SelectColumns(
            ComFlag,
            {"Método e versão (DocNix)", "Descrição Metodo", "Check doc", "TOTAL SOMA BANCADA", "TOTAL SOMA EQUIPAMENTO", "TOTAL GERAL"}
        ),

    TipoFinal =
        Table.TransformColumnTypes(
            Final,
            {{"TOTAL SOMA BANCADA", type number}, {"TOTAL SOMA EQUIPAMENTO", type number}, {"TOTAL GERAL", type number}}
        )
in
    TipoFinal;