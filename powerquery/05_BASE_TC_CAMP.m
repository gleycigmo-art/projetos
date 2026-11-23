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
