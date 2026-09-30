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
