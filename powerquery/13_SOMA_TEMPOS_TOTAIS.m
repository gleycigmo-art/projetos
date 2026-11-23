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
