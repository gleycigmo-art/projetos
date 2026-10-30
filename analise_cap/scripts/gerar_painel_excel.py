"""Define a aba PAINEL (fórmulas, textos e formatos) num lugar só e gera:
  - a macro VBA que monta a aba no arquivo do usuário (MontarPainel.bas, em cp1252 para o editor VBA);
  - um arquivo de teste .xlsx com as mesmas fórmulas e tabelas de dados, para recalcular no LibreOffice.

As fórmulas usam nomes de funções em inglês e vírgula (padrão de Range.Formula no VBA).
"""
import json, pickle, sys

MESES = ["Jan/2026", "Fev/2026", "Mar/2026", "Abr/2026", "Mai/2026", "Jun/2026",
         "Jul/2026", "Ago/2026", "Set/2026", "Out/2026", "Nov/2026", "Dez/2026"]

R_MES0, R_SEM = 26, 38          # linhas do quadro mês a mês (26..37) e da semana (38)
R_CAPH, R_CAP0, N_CAP = 42, 43, 60   # cabeçalho e linhas da CAP do período
CAP = 'CAP_PERIODOS'
DEM = 'Tabela_DEMANDA'
DS = 'Tabela_DemandaSemanal'
MESCOLS = f'{DEM}[[Jan/2026]:[Dez/2026]]'

# indicador: (rótulo, coluna no quadro mês a mês, formato, explicação)
INDICADORES = [
    ('HC por turno real', 'I', '0.00', 'Soma do recurso médio atual (Mão de Obra)'),
    ('HC real total', 'K', '0.00', 'HC por turno real x 3 turnos'),
    ('HC por turno necessário', 'J', '0.00', 'Soma do HC NECESSARIO (Mão de Obra)'),
    ('HC necessário total', 'L', '0.00', 'HC por turno necessário x 3 turnos'),
    ('% Ocupação', 'M', '0.0%', 'HC necessário total / HC real total'),
    ('% Lotes com TC', 'G', '0.0%', 'Lotes com tempo levantado / lotes'),
    ('HC por turno corrigido', 'N', '0.00', 'HC por turno necessário / % lotes com TC'),
    ('HC total corrigido', 'O', '0.00', 'HC por turno corrigido x 3 turnos'),
    ('% Ocupação corrigida', 'P', '0.0%', 'HC total corrigido / HC real total'),
    ('Lotes no período', 'Q', '#,##0', 'Semana: lotes com TC. Mês: demanda total CQ P30'),
    ('Dias úteis', 'H', '0.00', 'Calendário da semana ou do mês (aba FAMÍLIAS)'),
    ('Lotes por dia', 'R', '0.0', 'Lotes no período / dias úteis'),
    ('Lotes por dia a 100%', 'S', '0.0', 'Lotes por dia / % ocupação corrigida'),
]
QUADRO = [  # coluna, cabeçalho, formato
    ('B', 'Período', '@'), ('C', 'Nº', '0'), ('D', 'No filtro', '0'), ('E', 'Lotes', '#,##0.0'),
    ('F', 'Lotes c/ TC', '#,##0.0'), ('G', '% c/ TC', '0.0%'), ('H', 'Dias úteis', '0.00'),
    ('I', 'HC turno real', '0.00'), ('J', 'HC turno nec.', '0.00'), ('K', 'HC real total', '0.00'),
    ('L', 'HC nec. total', '0.00'), ('M', '% Ocupação', '0.0%'), ('N', 'HC turno corrig.', '0.00'),
    ('O', 'HC total corrig.', '0.00'), ('P', '% Ocup. corrig.', '0.0%'), ('Q', 'Lotes período', '#,##0.0'),
    ('R', 'Lotes/dia', '0.0'), ('S', 'Lotes/dia 100%', '0.0'), ('T', 'Nível 1', '#,##0.0'),
    ('U', 'Nível 2', '#,##0.0'), ('V', 'Nível 3', '#,##0.0'), ('W', 'Nível 4', '#,##0.0'),
    ('X', 'Nível 0', '#,##0.0'), ('Y', 'Sem classif.', '#,##0.0'),
]
CAPCOLS = [  # coluna PAINEL, coluna CAP_PERIODOS, formato
    ('B', 'PROCESSO', '@'), ('C', 'FAMILIA', '@'), ('D', 'CLUSTER', '@'), ('E', 'Categoria', '@'),
    ('F', 'CM (HRS)', '#,##0.00'), ('G', 'RECURSO MÉDIO ATUAL', '0.00'), ('H', 'TOTAL RECURSO', '#,##0.00'),
    ('I', 'HRS TURNO', '0.0'), ('J', 'OEE REAL', '0.0%'), ('K', 'OEE NEC', '0.0%'), ('L', 'HC NECESSARIO', '0.00'),
    ('M', 'CARREGAMENTO', '0.0%'), ('N', 'Soma Demanda Media', '#,##0.00'), ('O', 'TAKT', '0.00'),
    ('P', 'TC MEDIO', '0.00'), ('Q', 'FIFO', '0.00'),
]


def spec():
    """Lista de (célula, conteúdo, formato) da aba PAINEL. Conteúdo começando com '=' é fórmula."""
    c = []
    add = lambda ref, v, f=None: c.append((ref, v, f))
    add('B1', 'MAPA DE CAPACIDADE — CQPA')
    add('B2', '="Período: "&$F$6')
    # filtros
    for r, lab, val in ((3, 'Visão', 'Semanal'), (4, 'Semana', None), (5, 'Mês início', 'Set/2026'), (6, 'Mês fim', 'Set/2026')):
        add(f'B{r}', lab)
        if val:
            add(f'C{r}', val)
    add('C4', '=IFERROR(INDEX(v2_SEMANAS[Periodo],1),"")')  # valor inicial; a macro troca pelo texto
    add('E3', 'Chave da semana'); add('F3', '=IFERROR(INDEX(v2_SEMANAS[Ordem],MATCH($C$4,v2_SEMANAS[Periodo],0)),0)')
    add('E4', 'Nº mês início'); add('F4', '=MATCH($C$5,PAINEL_AUX!$A$2:$A$13,0)')
    add('E5', 'Nº mês fim'); add('F5', '=MAX($F$4,MATCH($C$6,PAINEL_AUX!$A$2:$A$13,0))')
    add('E6', 'Período'); add('F6', '=IF($C$3="Semanal",$C$4,IF($F$4=$F$5,$C$5,$C$5&" a "&INDEX(PAINEL_AUX!$A$2:$A$13,$F$5)&" (média dos meses)"))')
    add('E7', 'Linhas da CAP'); add('F7', f'=COUNTIFS({CAP}[Tipo],"Mensal",{CAP}[Ordem],1)')
    # indicadores
    add('B8', 'INDICADORES DO PERÍODO'); add('D8', 'Como é calculado')
    for i, (lab, col, fmt, exp) in enumerate(INDICADORES):
        r = 9 + i
        add(f'B{r}', lab)
        add(f'C{r}', f'=IF($C$3="Semanal",{col}{R_SEM},AVERAGEIFS({col}{R_MES0}:{col}{R_MES0 + 11},$D${R_MES0}:$D${R_MES0 + 11},1))', fmt)
        add(f'D{r}', exp)
    # complexidade
    add('F8', 'COMPLEXIDADE DOS LOTES')
    add('F9', 'Nível'); add('G9', 'Lotes'); add('H9', '% (1 a 4)'); add('I9', 'Faixa (tempo total do método)')
    niveis = [('Nível 1', 'T', '≤ 500 min'), ('Nível 2', 'U', '501 a 1.800 min'), ('Nível 3', 'V', '1.801 a 3.500 min'),
              ('Nível 4', 'W', '> 3.500 min'), ('Nível 0', 'X', 'só Check doc'), ('Sem classificação', 'Y', 'método não encontrado')]
    for i, (lab, col, faixa) in enumerate(niveis):
        r = 10 + i
        add(f'F{r}', lab)
        add(f'G{r}', f'=IF($C$3="Semanal",{col}{R_SEM},AVERAGEIFS({col}{R_MES0}:{col}{R_MES0 + 11},$D${R_MES0}:$D${R_MES0 + 11},1))', '#,##0.0')
        if i < 4:
            add(f'H{r}', f'=IFERROR(G{r}/SUM($G$10:$G$13),0)', '0.0%')
        add(f'I{r}', faixa)
    # conferência com a CAP antiga
    add('F17', 'CONFERÊNCIA COM A CAP_PROPOSTA ANTIGA')
    add('F18', 'Vale quando o PAINEL mostra o mesmo período do FILTRO_DATA. Esperado: 0,00.')
    add('F19', 'Soma |diferença| CM (HRS)'); add('H19', f'=SUMIF(S{R_CAP0}:S{R_CAP0 + N_CAP - 1},">0")-SUMIF(S{R_CAP0}:S{R_CAP0 + N_CAP - 1},"<0")', '0.0000')
    add('F20', 'Soma |diferença| HC NECESSARIO'); add('H20', f'=SUMIF(T{R_CAP0}:T{R_CAP0 + N_CAP - 1},">0")-SUMIF(T{R_CAP0}:T{R_CAP0 + N_CAP - 1},"<0")', '0.0000')
    add('F21', 'Soma |diferença| CARREGAMENTO'); add('H21', f'=SUMIF(U{R_CAP0}:U{R_CAP0 + N_CAP - 1},">0")-SUMIF(U{R_CAP0}:U{R_CAP0 + N_CAP - 1},"<0")', '0.0000')

    # quadro mês a mês
    add(f'B{R_MES0 - 2}', 'INDICADORES MÊS A MÊS (base dos números acima; intervalo = média dos meses marcados com 1)')
    for col, cab, _ in QUADRO:
        add(f'{col}{R_MES0 - 1}', cab)
    idx = f'INDEX({MESCOLS},0,$C{{r}})'
    for m in range(12):
        r = R_MES0 + m
        ix = idx.format(r=r)
        f = {
            'B': f'=INDEX(PAINEL_AUX!$A$2:$A$13,$C{r})', 'C': m + 1,
            'D': f'=IF(AND($C$3="Mensal",$C{r}>=$F$4,$C{r}<=$F$5),1,0)',
            'E': f'=SUMIFS({ix},{DEM}[Local de Análise],"CQ P30")',
            'F': f'=SUMIFS({ix},{DEM}[Local de Análise],"CQ P30",{DEM}[TC levantados?],"SIM")',
            'H': f'=INDEX(CALENDARIO_MES[Dias Úteis],MATCH($C{r},CALENDARIO_MES[Mês Nº],0))',
            'I': f'=SUMIFS({CAP}[RECURSO MÉDIO ATUAL],{CAP}[Tipo],"Mensal",{CAP}[Ordem],$C{r},{CAP}[Categoria],"Mão de Obra")',
            'J': f'=SUMIFS({CAP}[HC NECESSARIO],{CAP}[Tipo],"Mensal",{CAP}[Ordem],$C{r},{CAP}[Categoria],"Mão de Obra")',
            'Q': f'=E{r}',
        }
        for k, col in zip(range(5), 'TUVWX'):
            nivel = [1, 2, 3, 4, 0][k]
            f[col] = f'=SUMIFS({ix},{DEM}[Local de Análise],"CQ P30",{DEM}[complexidade],{nivel})'
        f['Y'] = f'=E{r}-SUM(T{r}:X{r})'
        for col, v in f.items():
            add(f'{col}{r}', v)
    r = R_SEM
    f = {
        'B': '=$C$4', 'C': '=$F$3', 'D': '=IF($C$3="Semanal",1,0)',
        'E': f'=SUMIFS({DS}[Lote semana],{DS}[Chave semana],$F$3)',
        'F': f'=SUMIFS({DS}[Lote semana],{DS}[Chave semana],$F$3,{DS}[Status2],"Tempos OK")',
        'H': f'=IFERROR(INDEX(CALENDARIO_SEMANA[Dias Úteis],MATCH($F$3,CALENDARIO_SEMANA[Semana Chave],0)),0)',
        'I': f'=SUMIFS({CAP}[RECURSO MÉDIO ATUAL],{CAP}[Tipo],"Semanal",{CAP}[Ordem],$F$3,{CAP}[Categoria],"Mão de Obra")',
        'J': f'=SUMIFS({CAP}[HC NECESSARIO],{CAP}[Tipo],"Semanal",{CAP}[Ordem],$F$3,{CAP}[Categoria],"Mão de Obra")',
        'Q': f'=F{r}',
    }
    for k, col in zip(range(5), 'TUVWX'):
        f[col] = f'=SUMIFS({DS}[Lote semana],{DS}[Chave semana],$F$3,{DS}[Complexidade],{[1, 2, 3, 4, 0][k]})'
    f['Y'] = f'=E{r}-SUM(T{r}:X{r})'
    for col, v in f.items():
        add(f'{col}{r}', v)
    for r in list(range(R_MES0, R_MES0 + 12)) + [R_SEM]:
        add(f'G{r}', f'=IFERROR(F{r}/E{r},0)')
        add(f'K{r}', f'=I{r}*3'); add(f'L{r}', f'=J{r}*3')
        add(f'M{r}', f'=IFERROR(L{r}/K{r},0)'); add(f'N{r}', f'=IFERROR(J{r}/G{r},0)')
        add(f'O{r}', f'=N{r}*3'); add(f'P{r}', f'=IFERROR(O{r}/K{r},0)')
        add(f'R{r}', f'=IFERROR(Q{r}/H{r},0)'); add(f'S{r}', f'=IFERROR(R{r}/P{r},0)')

    # CAP do período
    add(f'B{R_CAPH - 1}', 'CAP DO PERÍODO (filtre pela linha de cabeçalho; intervalo = média dos meses)')
    for col, nome, _ in CAPCOLS:
        add(f'{col}{R_CAPH}', nome)
    add(f'R{R_CAPH}', 'Status'); add(f'S{R_CAPH}', 'Dif. CM x CAP antiga')
    add(f'T{R_CAPH}', 'Dif. HC x CAP antiga'); add(f'U{R_CAPH}', 'Dif. Carreg. x CAP antiga')
    for i in range(N_CAP):
        r = R_CAP0 + i
        n = r - R_CAPH
        for col, nome, fmt in CAPCOLS[:4]:
            add(f'{col}{r}', f'=IF({n}<=$F$7,INDEX({CAP}[{nome}],{n}),"")', fmt)
        for col, nome, fmt in CAPCOLS[4:]:
            sem = f'AVERAGEIFS({CAP}[{nome}],{CAP}[Tipo],"Semanal",{CAP}[Ordem],$F$3,{CAP}[PROCESSO],$B{r},{CAP}[FAMILIA],$C{r})'
            mes = (f'AVERAGEIFS({CAP}[{nome}],{CAP}[Tipo],"Mensal",{CAP}[Ordem],">="&$F$4,{CAP}[Ordem],"<="&$F$5,'
                   f'{CAP}[PROCESSO],$B{r},{CAP}[FAMILIA],$C{r})')
            add(f'{col}{r}', f'=IF($C{r}="","",IFERROR(IF($C$3="Semanal",{sem},{mes}),""))', fmt)
        add(f'R{r}', f'=IF(M{r}="","",IF(M{r}>1,"Sobrecarregado",IF(M{r}>=0.8,"Adequado","Sub-utilizado")))')
        for col, nome, src in (('S', 'CM (HRS)', 'F'), ('T', 'HC NECESSARIO', 'L'), ('U', 'CARREGAMENTO', 'M')):
            add(f'{col}{r}', f'=IF($C{r}="","",IFERROR({src}{r}-SUMIFS(CAP_PROPOSTA[{nome}],CAP_PROPOSTA[PROCESSO],$B{r},CAP_PROPOSTA[FAMILIA],$C{r}),""))', '0.0000')
    # formatos do quadro mês a mês
    fmts = {col: fmt for col, _, fmt in QUADRO}
    linhas_q = set(range(R_MES0, R_MES0 + 12)) | {R_SEM}
    out = []
    for ref, v, fmt in c:
        col, row = ''.join(ch for ch in ref if ch.isalpha()), int(''.join(ch for ch in ref if ch.isdigit()))
        if fmt is None and row in linhas_q and col in fmts and col != 'B':
            fmt = fmts[col]
        out.append((ref, v, fmt))
    return out


# ───────────────────────────── VBA ─────────────────────────────
def vba_str(s):
    """Literal VBA (aspas dobradas); caracteres fora do cp1252 viram ChrW()."""
    out, buf = [], ''
    for ch in s:
        try:
            ch.encode('cp1252')
            buf += '""' if ch == '"' else ch
        except UnicodeEncodeError:
            if buf:
                out.append(f'"{buf}"'); buf = ''
            out.append(f'ChrW({ord(ch)})')
    if buf or not out:
        out.append(f'"{buf}"')
    return ' & '.join(out)


def gerar_vba(cells):
    linhas = []
    for ref, v, fmt in cells:
        if isinstance(v, str) and v.startswith('='):
            linhas.append(f'    ws.Range("{ref}").Formula = {vba_str(v)}')
        elif isinstance(v, str):
            linhas.append(f'    ws.Range("{ref}").Value = {vba_str(v)}')
        else:
            linhas.append(f'    ws.Range("{ref}").Value = {v}')
        if fmt:
            linhas.append(f'    ws.Range("{ref}").NumberFormat = {vba_str(fmt)}')
    # VBA tem limite de tamanho por procedimento: divide em blocos
    blocos, atual = [], []
    for l in linhas:
        atual.append(l)
        if len(atual) >= 150:
            blocos.append(atual); atual = []
    if atual:
        blocos.append(atual)
    subs = []
    for i, b in enumerate(blocos):
        subs.append(f'Private Sub EscreverPainel{i + 1}(ws As Worksheet)\n' + '\n'.join(b) + '\nEnd Sub\n')
    chamadas = '\n'.join(f'    EscreverPainel{i + 1} ws' for i in range(len(blocos)))
    return '\n'.join(subs), chamadas


# ───────────────────────────── teste no LibreOffice ─────────────────────────────
def gerar_teste(cells, tabelas_pkl, periodos_json, saida, visao, ini, fim):
    import openpyxl
    from openpyxl.worksheet.table import Table
    t = pickle.load(open(tabelas_pkl, 'rb'))
    per = json.load(open(periodos_json))
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    def tabela(sheet, nome, header, rows, cell='A1'):
        ws = wb[sheet] if sheet in wb.sheetnames else wb.create_sheet(sheet)
        c0 = openpyxl.utils.column_index_from_string(''.join(ch for ch in cell if ch.isalpha()))
        r0 = int(''.join(ch for ch in cell if ch.isdigit()))
        for j, h in enumerate(header):
            ws.cell(r0, c0 + j, h)
        for i, row in enumerate(rows):
            for j, v in enumerate(row):
                ws.cell(r0 + 1 + i, c0 + j, v if not (isinstance(v, str) and v.startswith('0x')) else None)
        ref = f'{openpyxl.utils.get_column_letter(c0)}{r0}:{openpyxl.utils.get_column_letter(c0 + len(header) - 1)}{r0 + len(rows)}'
        ws.add_table(Table(displayName=nome, ref=ref))

    cols = ['Tipo', 'Periodo', 'Ordem', 'PROCESSO', 'FAMILIA', 'CLUSTER', 'Categoria'] + [n for _, n, _ in CAPCOLS[4:]]
    rows = []
    ordem = [p for p in per if p['tipo'] == 'Mensal'] + [p for p in per if p['tipo'] == 'Semanal']
    for p in ordem:
        for r in p['cap']:
            rows.append([p['tipo'], p['rotulo'] if p['tipo'] == 'Mensal' else 'Sem 45/2026', p['ordem']] +
                        [r[c] for c in cols[3:]])
    tabela('CAP_PERIODOS', 'CAP_PERIODOS', cols, rows)
    for nome, sheet in (('Tabela_DEMANDA', 'DEMANDA'), ('Tabela_DemandaSemanal', 'Tabela_DemandaSemanal'),
                        ('CALENDARIO_MES', 'FAMILIAS'), ('CALENDARIO_SEMANA', 'FAMILIAS2'), ('CAP_PROPOSTA', 'CAP_PROPOSTA')):
        tabela(sheet, nome, t[nome]['header'], t[nome]['rows'])
    aux = wb.create_sheet('PAINEL_AUX')
    aux['A1'] = 'Mês'
    for i, m in enumerate(MESES):
        aux.cell(2 + i, 1, m)
    tabela('PAINEL_AUX', 'v2_SEMANAS', ['Periodo', 'Ordem'], [['Sem 45/2026', 202645]], 'C1')
    ws = wb.create_sheet('PAINEL')
    for ref, v, _ in cells:
        ws[ref] = v
    ws['C3'], ws['C4'], ws['C5'], ws['C6'] = visao, 'Sem 45/2026', ini, fim
    wb.save(saida)


if __name__ == '__main__':
    cells = spec()
    if sys.argv[1] == 'vba':
        subs, chamadas = gerar_vba(cells)
        tpl = open(sys.argv[2], encoding='utf-8').read()
        out = tpl.replace("'__SUBS__", subs).replace("'__CHAMADAS__", chamadas)
        out.encode('cp1252')  # garante que o editor VBA lê os acentos
        open(sys.argv[3], 'w', encoding='cp1252', newline='\r\n').write(out)
        print('ok', len(cells), 'células')
    else:  # teste
        gerar_teste(cells, sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6], sys.argv[7])
        print('ok')
