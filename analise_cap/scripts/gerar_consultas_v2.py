"""Gera as consultas Power Query novas (v2) a partir do texto ORIGINAL das consultas do arquivo.

As consultas atuais não são alteradas. As novas reaproveitam o corpo original, sem reescrever a lógica:
- fnDEMANDA_MPR(Filtro)        = corpo de DEMANDA_MPR, com FILTRO_PERIODO trocado pelo parâmetro.
- fnTC_COM_CAMPANHA(DEMANDA_MPR) = corpo de TC_COM_CAMPANHA, recebendo a demanda como parâmetro.
- fnCALENDARIO_ATIVO(Filtro)   = linhas do calendário (semana ou meses) do período.
- fnFAM_PROPOSTA(Filtro)       = Tabela_FAM com HRS TURNO / TOTAL RECURSO do período (mesma fórmula da planilha)
                                 + as mesmas conversões de tipo de FAM_PROPOSTA.
- fnCAP(Filtro)                = corpo de CAP_PROPOSTA com 2 mudanças: coluna Categoria (de Tabela_FAM)
                                 e FIFO dividido pelos dias úteis do período (calendário).
- v2_SEMANAS, v2_CAP_PERIODOS  = lista de semanas e a CAP de cada mês e de cada semana, numa tabela só.
"""
import re, sys, os

SRC = sys.argv[1]
OUT = sys.argv[2]
os.makedirs(OUT, exist_ok=True)
txt = open(SRC, encoding='utf-8').read()


def corpo(nome):
    """Texto entre 'shared <nome> =' e o ';' final do membro."""
    m = re.search(r'^shared ' + re.escape(nome) + r' = (.*?);\s*(?=^shared |\Z)', txt, re.S | re.M)
    assert m, nome
    return m.group(1).rstrip()


def troca(s, velho, novo):
    n = s.count(velho)
    assert n == 1, f'esperava 1 ocorrência, achei {n}: {velho[:60]!r}'
    return s.replace(velho, novo)


def indenta(s, n):
    return '\n'.join((' ' * n + l) if l.strip() else l for l in s.splitlines())


MESES = '{"Jan/2026", "Fev/2026", "Mar/2026", "Abr/2026", "Mai/2026", "Jun/2026",\n         "Jul/2026", "Ago/2026", "Set/2026", "Out/2026", "Nov/2026", "Dez/2026"}'

consultas = {}

# ── fnDEMANDA_MPR ─────────────────────────────────────────────────────────────
consultas['fnDEMANDA_MPR'] = f"""// Demanda por método para um período. Corpo idêntico à consulta DEMANDA_MPR.
(Filtro as record) as table =>
let
    FILTRO_PERIODO = Filtro,
    Resultado =
{indenta(corpo('DEMANDA_MPR'), 8)}
in
    Resultado"""

# ── fnTC_COM_CAMPANHA ─────────────────────────────────────────────────────────
consultas['fnTC_COM_CAMPANHA'] = f"""// Carga em minutos por bloco para uma demanda. Corpo idêntico à consulta TC_COM_CAMPANHA.
(DEMANDA_MPR as table) as table =>
let
    Resultado =
{indenta(corpo('TC_COM_CAMPANHA'), 8)}
in
    Resultado"""

# ── v2_FONTES_CAL ──────────────────────────────────────────────────────────────
# Mesmo padrão de FONTES_TC: só esta consulta lê a planilha; as funções apenas a referenciam
# (evita o erro de "Formula.Firewall" ao combinar fontes).
consultas['v2_FONTES_CAL'] = """// Tabelas da planilha usadas pelas funções v2 (lidas só aqui, como em FONTES_TC).
let
    CalMes = Excel.CurrentWorkbook(){[Name="CALENDARIO_MES"]}[Content],
    CalSem = Excel.CurrentWorkbook(){[Name="CALENDARIO_SEMANA"]}[Content],
    Fam    = Excel.CurrentWorkbook(){[Name="Tabela_FAM"]}[Content]
in
    [
        CalMes = CalMes,
        CalSem = CalSem,
        Fam    = Fam
    ]"""

# ── fnCALENDARIO_ATIVO ────────────────────────────────────────────────────────
consultas['fnCALENDARIO_ATIVO'] = f"""// Linhas do calendário do período: a semana escolhida ou os meses ativos.
(Filtro as record) as table =>
let
    CalMes     = v2_FONTES_CAL[CalMes],
    CalSem     = v2_FONTES_CAL[CalSem],
    Meses      = {MESES},
    NumsAtivos = List.Transform(Filtro[MesesAtivos], each List.PositionOf(Meses, _) + 1),
    Resultado  = if Filtro[NivelDeVisao] = "Semanal"
                 then Table.SelectRows(CalSem, each [Semana Chave] = Filtro[SemanaChave])
                 else Table.SelectRows(CalMes, each List.Contains(NumsAtivos, [#"Mês Nº"]))
in
    Resultado"""

# ── fnFAM_PROPOSTA ────────────────────────────────────────────────────────────
fam = corpo('FAM_PROPOSTA')
fam_tipos = re.search(r'(#"Tipo Alterado" = Table\.TransformColumnTypes\(Fonte,\{.*?\}\}\))', fam, re.S).group(1)
fam_tipos = troca(fam_tipos, 'TransformColumnTypes(Fonte,', 'TransformColumnTypes(ComTotal,')
consultas['fnFAM_PROPOSTA'] = f"""// Tabela_FAM com as horas do período. Mesma regra da coluna HRS TURNO da planilha:
// média, nos períodos ativos, das horas acumuladas para N turnos (Mão de Obra: MO-N; Equipamento: EQP-N).
// Depois, as mesmas conversões de tipo da consulta FAM_PROPOSTA (HRS TURNO vira inteiro).
(Filtro as record) as table =>
let
    Fonte    = v2_FONTES_CAL[Fam],
    Semanal  = Filtro[NivelDeVisao] = "Semanal",
    CalAtivo = fnCALENDARIO_ATIVO(Filtro),
    ColsMO   = if Semanal then {{"MO - 1", "MO - 2", "MO - 3"}} else {{"MO - 1", "MO - 2", "MO - 3", "ADM"}},
    ColsEQ   = if Semanal then {{"EQP - 1", "EQP - 2", "EQP - 3"}} else {{"EQP - 1", "EQP - 2", "EQP - 3", "ADM2"}},
    HrsMedia = (categoria as any, turnos as any) as any =>
                   let
                       Cols = if categoria = "Mão de Obra" then ColsMO else ColsEQ,
                       Col  = Cols{{Int64.From(turnos) - 1}}
                   in
                       List.Average(Table.Column(CalAtivo, Col)),
    SemHrs   = Table.RemoveColumns(Fonte, {{"HRS TURNO", "TOTAL RECURSO"}}),
    ComHrs   = Table.AddColumn(SemHrs, "HRS TURNO", each HrsMedia([CATEGORIA], [TURNOS]), type number),
    ComTotal = Table.AddColumn(ComHrs, "TOTAL RECURSO", each [#"RECURSO MÉDIO ATUAL"] * [HRS TURNO], type number),
    {fam_tipos}
in
    #"Tipo Alterado\""""

# ── fnCAP ─────────────────────────────────────────────────────────────────────
cap = corpo('CAP_PROPOSTA')
cap = troca(cap, '''                    "FAMILIA", "CLUSTER",
                    "RECURSO MÉDIO ATUAL",''', '''                    "FAMILIA", "CLUSTER", "CATEGORIA",
                    "RECURSO MÉDIO ATUAL",''')
cap = troca(cap, '''SemMatch_Exp  = Table.ExpandTableColumn(SemMatch_Join, "Map2", {
                        "CLUSTER",''', '''SemMatch_Exp  = Table.ExpandTableColumn(SemMatch_Join, "Map2", {
                        "CLUSTER", "CATEGORIA",''')
cap = troca(cap, '''                        {"CLUSTER_2",''', '''                        {"CLUSTER_2", "CATEGORIA_2",''')
cap = troca(cap, '''                       {"CLUSTER",             "CLUSTER_OLD"},''',
            '''                       {"CLUSTER",             "CLUSTER_OLD"},
                       {"CATEGORIA",           "CATEGORIA_OLD"},''')
cap = troca(cap, '''                                  CLUSTER = [CLUSTER_2],''', '''                                  CLUSTER = [CLUSTER_2],
                                  CATEGORIA = [CATEGORIA_2],''')
cap = troca(cap, '''                              {"CLUSTER", "RECURSO MÉDIO ATUAL", "HRS TURNO",''',
            '''                              {"CLUSTER", "CATEGORIA", "RECURSO MÉDIO ATUAL", "HRS TURNO",''')
cap = troca(cap, '''                          {"FAM proposta", "PROCESSO", "CM_Min_Total",
                           "FAMILIA_FINAL", "CLUSTER",''', '''                          {"FAM proposta", "PROCESSO", "CM_Min_Total",
                           "FAMILIA_FINAL", "CLUSTER", "CATEGORIA",''')
cap = troca(cap, '''                       "FAM proposta", "PROCESSO", "CM_Min_Total",
                       "FAMILIA_FINAL", "CLUSTER",''', '''                       "FAM proposta", "PROCESSO", "CM_Min_Total",
                       "FAMILIA_FINAL", "CLUSTER", "CATEGORIA",''')
cap = troca(cap, '''                       {"CLUSTER",             type text},
                       {"CM_Min_Total",''', '''                       {"CLUSTER",             type text},
                       {"CATEGORIA",           type text},
                       {"CM_Min_Total",''')
cap = troca(cap, 'then [TOTAL RECURSO] / [TAKT] / 22.97', 'then [TOTAL RECURSO] / [TAKT] / DiasUteisPeriodo')
cap = troca(cap, '''                     {{"FAMILIA_FINAL", "FAMILIA"}}),''',
            '''                     {{"FAMILIA_FINAL", "FAMILIA"}, {"CATEGORIA", "Categoria"}}),''')
cap = troca(cap, '''                     {"PROCESSO", "FAMILIA", "CLUSTER",
                      "CM (HRS)", "RECURSO MÉDIO ATUAL",''', '''                     {"PROCESSO", "FAMILIA", "CLUSTER", "Categoria",
                      "CM (HRS)", "RECURSO MÉDIO ATUAL",''')
consultas['fnCAP'] = f"""// CAP de um período. Corpo da consulta CAP_PROPOSTA com duas mudanças:
//   1) coluna Categoria vinda da Tabela_FAM (antes era digitada na planilha);
//   2) FIFO = TOTAL RECURSO / TAKT / dias úteis do período no calendário (antes: 22,97 fixo).
(Filtro as record) as table =>
let
    FILTRO_PERIODO   = Filtro,
    DEMANDA_MPR      = Table.Buffer(fnDEMANDA_MPR(Filtro)),
    TC_COM_CAMPANHA  = fnTC_COM_CAMPANHA(DEMANDA_MPR),
    FAM_PROPOSTA     = fnFAM_PROPOSTA(Filtro),
    DiasUteisPeriodo = List.Average(Table.Column(fnCALENDARIO_ATIVO(Filtro), "Dias Úteis")),
    Resultado =
{indenta(cap, 8)}
in
    Resultado"""

# ── v2_SEMANAS ────────────────────────────────────────────────────────────────
consultas['v2_SEMANAS'] = """// Semanas disponíveis na base de quarentena (SharePoint), com o rótulo usado no PAINEL.
let
    Chaves  = List.Sort(List.Distinct(List.RemoveNulls(Table.Column(Tabela_DemandaSemanal, "Chave semana")))),
    Tabela  = Table.FromList(Chaves, Splitter.SplitByNothing(), {"Ordem"}),
    Tipada  = Table.TransformColumnTypes(Tabela, {{"Ordem", Int64.Type}}),
    Rotulo  = Table.AddColumn(Tipada, "Periodo",
                  each "Sem " & Text.PadStart(Text.Middle(Text.From([Ordem]), 4), 2, "0")
                       & "/" & Text.Start(Text.From([Ordem]), 4),
                  type text)
in
    Table.ReorderColumns(Rotulo, {"Periodo", "Ordem"})"""

# ── v2_CAP_PERIODOS ───────────────────────────────────────────────────────────
consultas['v2_CAP_PERIODOS'] = f"""// CAP de cada mês (Jan a Dez) e de cada semana disponível, numa tabela só.
// Intervalos de meses são calculados no PAINEL como média simples dos meses.
let
    Meses      = {MESES},
    FiltrosMes = List.Transform({{0..11}}, (i) =>
                     [NivelDeVisao = "Mensal", MesesAtivos = {{Meses{{i}}}}, SemanaChave = null,
                      Tipo = "Mensal", Ordem = i + 1, Periodo = Meses{{i}}]),
    FiltrosSem = List.Transform(Table.ToRecords(v2_SEMANAS), (s) =>
                     [NivelDeVisao = "Semanal", MesesAtivos = {{}}, SemanaChave = s[Ordem],
                      Tipo = "Semanal", Ordem = s[Ordem], Periodo = s[Periodo]]),
    Tabelas    = List.Transform(FiltrosMes & FiltrosSem, (f) =>
                     let
                         Cap = fnCAP(f),
                         T1  = Table.AddColumn(Cap, "Tipo", each f[Tipo], type text),
                         T2  = Table.AddColumn(T1, "Periodo", each f[Periodo], type text),
                         T3  = Table.AddColumn(T2, "Ordem", each f[Ordem], Int64.Type)
                     in
                         T3),
    Combinado  = Table.Combine(Tabelas),
    Ordenado   = Table.ReorderColumns(Combinado,
                     {{"Tipo", "Periodo", "Ordem"}} & List.RemoveItems(Table.ColumnNames(Combinado), {{"Tipo", "Periodo", "Ordem"}}))
in
    Ordenado"""

ORDEM = ['v2_FONTES_CAL', 'fnCALENDARIO_ATIVO', 'fnFAM_PROPOSTA', 'fnDEMANDA_MPR', 'fnTC_COM_CAMPANHA', 'fnCAP', 'v2_SEMANAS', 'v2_CAP_PERIODOS']
for n in ORDEM:
    open(os.path.join(OUT, n + '.m'), 'w', encoding='utf-8').write(consultas[n] + '\n')
    s = '\n'.join(l for l in consultas[n].splitlines() if not l.lstrip().startswith('//'))  # ignora comentários
    assert s.count('(') == s.count(')') and s.count('{') == s.count('}') and s.count('[') == s.count(']'), n
    print(n, len(s.splitlines()), 'linhas')
