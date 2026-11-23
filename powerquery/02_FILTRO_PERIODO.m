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
