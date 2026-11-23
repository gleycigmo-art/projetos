// Base de tempos: só as colunas usadas no cálculo (não carregar).
let
    Fonte = fnTabela("Tabela_BASE_TC"),
    Colunas = {
        "Método e versão (DocNix)", "FAM proposta", "Equipamento Proposto",
        "SPCT0071 - AA", "ICP", "SPCT - IV", "SPCT - UV",
        "1 - PREPARO (PADRÃO E SOLUÇÕES) - BANCADA", "2 - PREPARO AMOSTRA - BANCADA",
        "3 - ANÁLISE FQ - BANCADA", "6 - FINALIZAÇÃO DOS DADOS - BANCADA",
        "7 - CONFERÊNCIA DO PREPARO - BANCADA", "8 - TEMPO DO DISSOLUTOR - BANCADA",
        "9 - MONTAGEM DO EQUIPAMENTO - LEITURA", "10 - CONFERÊNCIA EQUIPAMENTO - LEITURA",
        "12 - Nº INJEÇÕES PADRÃO - LEITURA", "12 - TEMPO PADRÃO (CORRIDA) - LEITURA",
        "13 - Nº INJEÇÕES AMOSTRA/ DILUENTE/ PLACEBO - LEITURA",
        "13 - TEMPO AMOSTRA/ DILUENTE/ PLACEBO (CORRIDA) - LEITURA",
        "14 - PROCESSAMENTO DA CORRIDA - LEITURA", "15 - CONFERÊNCIA PROCESSAMENTO DA CORRIDA",
        "16 - REVISÃO LAUDO - REVISÃO"
    },
    Selecionado = Table.SelectColumns(Fonte, Colunas, MissingField.UseNull),
    SemErros = Table.ReplaceErrorValues(Selecionado, List.Transform(Colunas, each {_, null}))
in
    SemErros
