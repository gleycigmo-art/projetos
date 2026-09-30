// Cada linha MPR × teste vira "componentes" de carga: tempo (min) de uma etapa, o
// recurso que ela ocupa (DESTINO + PROCESSO), o % de campanha e se aplica campanha.
// Regras (corrigidas):
//   • 1º lote de cada campanha com tempo cheio, demais com o % da etapa
//   • skip só no teste de impureza ("Impureza Skip Test") de MPR que está na Tabela_SKIP_IMP
// (não carregar)
let
    P = stg_Parametros,
    Pos = (p) => if p <> null and p > 0 then p else null,
    Preenchido = (v) => not (v = null or v = "" or (Value.Is(v, type number) and v = 0)),
    Valido = (v) => v <> null and fnNum(v) <> 0 and Text.From(v) <> "S",

    COL_MPR = "Método e versão (DocNix)",
    E1 = "1 - PREPARO (PADRÃO E SOLUÇÕES) - BANCADA", E2 = "2 - PREPARO AMOSTRA - BANCADA",
    E3 = "3 - ANÁLISE FQ - BANCADA", E6 = "6 - FINALIZAÇÃO DOS DADOS - BANCADA",
    E7 = "7 - CONFERÊNCIA DO PREPARO - BANCADA", E8 = "8 - TEMPO DO DISSOLUTOR - BANCADA",
    E9 = "9 - MONTAGEM DO EQUIPAMENTO - LEITURA", E10 = "10 - CONFERÊNCIA EQUIPAMENTO - LEITURA",
    E12N = "12 - Nº INJEÇÕES PADRÃO - LEITURA", E12T = "12 - TEMPO PADRÃO (CORRIDA) - LEITURA",
    E13N = "13 - Nº INJEÇÕES AMOSTRA/ DILUENTE/ PLACEBO - LEITURA",
    E13T = "13 - TEMPO AMOSTRA/ DILUENTE/ PLACEBO (CORRIDA) - LEITURA",
    E14 = "14 - PROCESSAMENTO DA CORRIDA - LEITURA", E15 = "15 - CONFERÊNCIA PROCESSAMENTO DA CORRIDA",
    E16 = "16 - REVISÃO LAUDO - REVISÃO",
    ESPECTRO = {"SPCT0071 - AA", "ICP", "SPCT - IV", "SPCT - UV"},

    Base = stg_BaseTC,
    ComMPR = Table.SelectRows(Base, each Record.Field(_, COL_MPR) <> null and Record.Field(_, COL_MPR) <> ""),

    CompMO = (r as record) as list =>
        let
            mpr = Record.Field(r, COL_MPR),
            f = r[FAM proposta],
            v = (c) => Record.Field(r, c),
            n = (c) => fnNum(Record.Field(r, c)),
            c = (dest, proc, tc, pct, camp) =>
                [MPR = mpr, FAM = f, DESTINO = dest, PROCESSO = proc, TC = tc, PCT = pct, CAMPANHA = camp],
            conf = if f = "Fam PI/TF" then "MO CONFERÊNCIA PI/TF" else "MO CONFERÊNCIA BANCADA",
            esp = List.Sum(List.Transform(ESPECTRO, each n(_))),
            A = {
                c(f, "MO BANCADA", n(E1), P[P1], true),
                c(f, "MO BANCADA", n(E2), null, true),
                c(f, "MO BANCADA", n(E3), null, true),
                c(f, "MO BANCADA", n(E6), P[P6], true)
            },
            B = if Preenchido(v(E7)) and n(E7) <> 0 then {c(conf, conf, n(E7), P[P7], true)} else {},
            Cesp = if esp <> 0 then {c("MO EQUIPAMENTO BANCADA", "MO EQUIPAMENTO BANCADA", esp, null, false)} else {},
            D = (if Valido(v(E9)) then {c("MO EQUIPAMENTO", "MO EQUIPAMENTO", n(E9), Pos(P[P9]), true)} else {})
              & (if Valido(v(E14)) then {c("MO EQUIPAMENTO", "MO EQUIPAMENTO", n(E14), null, true)} else {}),
            Ecf = (if Valido(v(E10)) then {c("MO CONFERÊNCIA EQUIPAMENTO", "MO CONFERÊNCIA EQUIPAMENTO", n(E10), Pos(P[P10]), true)} else {})
              & (if Valido(v(E15)) then {c("MO CONFERÊNCIA EQUIPAMENTO", "MO CONFERÊNCIA EQUIPAMENTO", n(E15), null, true)} else {}),
            F = if n(E16) <> 0 then {c("REVISÃO FINAL", "REVISÃO FINAL", n(E16), null, false)} else {},
            G = if n(E8) <> 0 then {c("DISSOLUTOR", "DISSOLUTOR", n(E8), null, false)} else {}
        in
            A & B & Cesp & D & Ecf & F & G,

    // Tempo de máquina HPLC (CRLQ) / CG (CRGS): corrida do padrão com campanha, amostras por lote
    ComEquip = Table.SelectRows(Base, each let e = [Equipamento Proposto] in
                   e <> null and Value.Is(e, type text) and (Text.StartsWith(e, "CRLQ") or Text.StartsWith(e, "CRGS"))),
    CompEq = (r as record) as list =>
        let
            e = r[Equipamento Proposto],
            proc = if Text.StartsWith(e, "CRLQ") then "EQUIPAMENTO HPLC" else "EQUIPAMENTO CG",
            base = [MPR = Record.Field(r, COL_MPR), FAM = r[FAM proposta], DESTINO = e, PROCESSO = proc, CAMPANHA = true]
        in
            {
                base & [TC = fnNum(Record.Field(r, E12N)) * fnNum(Record.Field(r, E12T)), PCT = if P[P12] = null then 0 else P[P12]],
                base & [TC = fnNum(Record.Field(r, E13N)) * fnNum(Record.Field(r, E13T)), PCT = null]
            },

    Lista = List.Combine(List.Transform(Table.ToRecords(ComMPR), CompMO))
          & List.Combine(List.Transform(Table.ToRecords(ComEquip), CompEq)),
    Tabela = Table.FromRecords(Lista,
                type table [MPR = nullable text, FAM = nullable text, DESTINO = nullable text, PROCESSO = text,
                            TC = number, PCT = nullable number, CAMPANHA = logical],
                MissingField.UseNull),

    // skip: N do MPR quando o componente é do teste de impureza e o MPR está na tabela
    JSkip = Table.NestedJoin(Tabela, {"MPR"}, P[Skip], {"MPR"}, "S", JoinKind.LeftOuter),
    ESkip = Table.ExpandTableColumn(JSkip, "S", {"NSKIP"}, {"NSKIP_MPR"}),
    NSkip = Table.RemoveColumns(
                Table.AddColumn(ESkip, "NSKIP", each if [FAM] = "Impureza Skip Test" then [NSKIP_MPR] else null, type nullable number),
                {"NSKIP_MPR"}),
    JCamp = Table.NestedJoin(NSkip, {"MPR"}, P[Campanha], {"CÓD"}, "C", JoinKind.LeftOuter),
    ECamp = Table.ExpandTableColumn(JCamp, "C", {"LXC"}),

    // soma os tempos de componentes com a mesma regra (a carga é linear no tempo)
    Agrupado = Table.Group(ECamp, {"MPR", "DESTINO", "PROCESSO", "PCT", "CAMPANHA", "NSKIP", "LXC"},
                   {{"TC", each List.Sum([TC]), type number}})
in
    Agrupado
