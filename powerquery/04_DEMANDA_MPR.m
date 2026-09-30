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
