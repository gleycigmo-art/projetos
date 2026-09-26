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
