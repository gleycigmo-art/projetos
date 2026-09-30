// Quais MPRs passam por cada recurso — base da "Soma Demanda Media" (lotes do recurso).
// TIPO = FAM (família proposta), PROC (processo de conferência/equipamento/revisão)
// ou EQUIP (HPLC/CG). (não carregar)
let
    COL_MPR = "Método e versão (DocNix)",
    Cheio = (r, cols) => List.AnyTrue(List.Transform(cols, (c) =>
                let v = Record.Field(r, c) in not (v = null or v = "" or (Value.Is(v, type number) and v = 0)))),
    Base = Table.ToRecords(Table.SelectRows(stg_BaseTC, each Record.Field(_, COL_MPR) <> null and Record.Field(_, COL_MPR) <> "")),
    Pares = List.Combine(List.Transform(Base, (r) =>
        let
            m = Record.Field(r, COL_MPR), f = r[FAM proposta],
            p = (tipo, chave) => [TIPO = tipo, CHAVE = chave, MPR = m]
        in
            (if f <> null and f <> "Terceirizado" then {p("FAM", f)} else {})
          & (if Cheio(r, {"7 - CONFERÊNCIA DO PREPARO - BANCADA"})
                then {p("PROC", if f = "Fam PI/TF" then "MO CONFERÊNCIA PI/TF" else "MO CONFERÊNCIA BANCADA")} else {})
          & (if Cheio(r, {"10 - CONFERÊNCIA EQUIPAMENTO - LEITURA", "15 - CONFERÊNCIA PROCESSAMENTO DA CORRIDA"})
                then {p("PROC", "MO CONFERÊNCIA EQUIPAMENTO")} else {})
          & (if Cheio(r, {"SPCT0071 - AA", "ICP", "SPCT - IV", "SPCT - UV"}) then {p("PROC", "MO EQUIPAMENTO BANCADA")} else {})
          & (if Cheio(r, {"16 - REVISÃO LAUDO - REVISÃO"}) then {p("PROC", "REVISÃO FINAL")} else {})
          & (if Cheio(r, {"8 - TEMPO DO DISSOLUTOR - BANCADA"}) then {p("PROC", "DISSOLUTOR")} else {})
          & (if Cheio(r, {"9 - MONTAGEM DO EQUIPAMENTO - LEITURA", "14 - PROCESSAMENTO DA CORRIDA - LEITURA"})
                then {p("PROC", "MO EQUIPAMENTO")} else {}))),
    Equip = List.Transform(
                Table.ToRecords(Table.SelectRows(stg_BaseTC, each let e = [Equipamento Proposto] in
                    e <> null and Value.Is(e, type text) and (Text.StartsWith(e, "CRLQ") or Text.StartsWith(e, "CRGS")))),
                (r) => [TIPO = "EQUIP", CHAVE = r[Equipamento Proposto], MPR = Record.Field(r, COL_MPR)]),
    Tabela = Table.Distinct(Table.FromRecords(Pares & Equip, type table [TIPO = text, CHAVE = text, MPR = nullable text]))
in
    Tabela
