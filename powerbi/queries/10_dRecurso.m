// Recursos da aba FAMÍLIAS (Tabela_FAM): pessoas por turno e OEE por família/processo.
let
    Fam = fnTabela("Tabela_FAM"),
    Linhas = Table.SelectRows(Fam, each [FAMILIA] <> null and [FAMILIA] <> ""),
    Registros = List.Transform(Table.ToRecords(Linhas), (r) =>
        let
            t = Number.RoundAwayFromZero(fnNum(r[TURNOS])),
            n = if t = 0 then 3 else t,
            q = {fnNum(r[#"QTDE 1° TURNO"]), fnNum(r[#"QTDE 2° TURNO"]), fnNum(r[#"QTDE 3° TURNO"])}
        in
        [FAMILIA = r[FAMILIA], PROCESSO = r[PROCESSO], CLUSTER = r[CLUSTER], CATEGORIA = r[CATEGORIA],
         TURNOS = n, QTDE_1T = q{0}, QTDE_2T = q{1}, QTDE_3T = q{2},
         RECURSO = List.Sum(List.FirstN(q, n)) / n, OEE = fnNum(r[OEE REAL])]),
    Tabela = Table.FromRecords(Registros),
    Tipado = Table.TransformColumnTypes(Tabela, {
                 {"FAMILIA", type text}, {"PROCESSO", type text}, {"CLUSTER", type text}, {"CATEGORIA", type text},
                 {"TURNOS", Int64.Type}, {"QTDE_1T", type number}, {"QTDE_2T", type number}, {"QTDE_3T", type number},
                 {"RECURSO", type number}, {"OEE", type number}})
in
    Tipado
