// Campanha (lotes por campanha), skip e % de tempo das etapas em campanha (não carregar).
let
    Campanha = Table.Distinct(fnTabela("Tabela_BD_CAMPANHA"), {"CÓD"}),
    LXC = Table.AddColumn(Campanha, "LXC",
            each Number.RoundUp(fnNum([QTD LOTES CAMPANHA]) - 0.000000000001, 0), type number),
    TabCampanha = Table.SelectColumns(LXC, {"CÓD", "LXC"}),

    Skip = Table.Distinct(fnTabela("Tabela_SKIP_IMP"), {"MPR"}),
    SkipValido = Table.SelectRows(Skip, each [Qtde Lotes Skip] <> null and fnNum([Qtde Lotes Skip]) >= 1),
    TabSkip = Table.SelectColumns(
                Table.AddColumn(SkipValido, "NSKIP", each fnNum([Qtde Lotes Skip]), type number),
                {"MPR", "NSKIP"}),

    Pct = fnTabela("Tabela_Lotes_Campanha"),
    PctDe = (etapa as text) =>
        let r = Table.SelectRows(Pct, each [Etapa] = etapa & "_CAMPANHA")
        in if Table.IsEmpty(r) or r{0}[#"% tempo"] = null then null else fnNum(r{0}[#"% tempo"]),

    Resultado = [
        Campanha = TabCampanha,
        Skip = TabSkip,
        P1  = PctDe("1 - PREPARO (PADRÃO E SOLUÇÕES) - BANCADA"),
        P6  = PctDe("6 - FINALIZAÇÃO DOS DADOS - BANCADA"),
        P7  = PctDe("7 - CONFERÊNCIA DO PREPARO - BANCADA"),
        P9  = PctDe("9 - MONTAGEM DO EQUIPAMENTO - LEITURA"),
        P10 = PctDe("10 - CONFERÊNCIA EQUIPAMENTO - LEITURA"),
        P12 = PctDe("12 - TEMPO PADRÃO (CORRIDA) - LEITURA")
    ]
in
    Resultado
