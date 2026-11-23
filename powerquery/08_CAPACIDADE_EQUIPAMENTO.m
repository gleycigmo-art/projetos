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
