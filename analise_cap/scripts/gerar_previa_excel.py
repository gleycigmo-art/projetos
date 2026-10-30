"""Gera um .xlsx independente com a aba PAINEL funcionando sobre os dados do arquivo de 30/10
(retrato: as tabelas são valores, não consultas). Usa as mesmas fórmulas do kit.

Uso: python3 gerar_previa_excel.py tabelas.pkl periodos.json saida.xlsx
"""
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from gerar_painel_excel import spec, gerar_teste, R_MES0, R_SEM, R_CAPH, R_CAP0, N_CAP

TAB, PER, OUT = sys.argv[1:4]
cells = spec()
gerar_teste(cells, TAB, PER, OUT, 'Semanal', 'Set/2026', 'Set/2026')

wb = openpyxl.load_workbook(OUT)
ws = wb['PAINEL']
for ref, _, fmt in cells:
    if fmt:
        ws[ref].number_format = fmt

AZUL, CINZA, BRANCO = '1F4E78', '808080', 'FFFFFF'
cab_fill = PatternFill('solid', fgColor=AZUL)
cab_font = Font(bold=True, color=BRANCO)
fina = Side(style='thin', color='D9D9D9')


def cab(rng):
    for row in ws[rng]:
        for c in row:
            c.fill, c.font = cab_fill, cab_font
            c.alignment = Alignment(wrap_text=True, vertical='center')


ws.sheet_view.showGridLines = False
for col, w in {'A': 2, 'B': 30, 'C': 30, 'D': 40, 'I': 24}.items():
    ws.column_dimensions[col].width = w
for col in 'EFGHJKLMNOPQRSTUVWXY':
    ws.column_dimensions[col].width = 13
ws['B1'].font = Font(size=16, bold=True, color=AZUL)
ws['B2'].font = Font(size=12)

# filtros
for r in range(3, 7):
    ws[f'B{r}'].font = Font(bold=True)
    c = ws[f'C{r}']
    c.fill = PatternFill('solid', fgColor='FFF2CC')
    c.font = Font(bold=True, size=11)
    c.border = Border(left=fina, right=fina, top=fina, bottom=fina)
aux = wb['PAINEL_AUX']
aux['E1'], aux['E2'], aux['E3'] = 'Visão', 'Semanal', 'Mensal'
for ref, src in (('C3', 'PAINEL_AUX!$E$2:$E$3'), ('C4', 'PAINEL_AUX!$C$2:$C$2'),
                 ('C5', 'PAINEL_AUX!$A$2:$A$13'), ('C6', 'PAINEL_AUX!$A$2:$A$13')):
    dv = DataValidation(type='list', formula1=f'={src}', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(ref)
for r in range(3, 8):
    ws[f'E{r}'].font = ws[f'F{r}'].font = Font(color=CINZA)

# indicadores
cab('B8:D8')
for r in range(9, 22):
    ws[f'C{r}'].font = Font(bold=True, size=14 if r in (13, 17, 21) else 12)
    ws[f'C{r}'].alignment = Alignment(horizontal='right')
    ws[f'B{r}'].font = Font(bold=r in (13, 17, 21), size=14 if r in (13, 17, 21) else 10)
    ws[f'D{r}'].font = Font(italic=True, color=CINZA)
    for col in 'BCD':
        ws[f'{col}{r}'].border = Border(bottom=fina)
for ref in ('C13', 'C17'):
    ws.conditional_formatting.add(ref, FormulaRule(formula=[f'{ref}>1'], font=Font(color='C00000', bold=True), stopIfTrue=True))
    ws.conditional_formatting.add(ref, FormulaRule(formula=[f'{ref}>=0.8'], font=Font(color='008000', bold=True), stopIfTrue=True))
    ws.conditional_formatting.add(ref, FormulaRule(formula=[f'{ref}<0.8'], font=Font(color='1F70D6', bold=True)))

# complexidade e conferência
cab('F8:J8')
for col in 'FGHIJ':
    ws[f'{col}9'].font = Font(bold=True)
for r in range(10, 14):
    ws[f'J{r}'] = f'=REPT("█",ROUND(H{r}*40,0))'
    ws[f'J{r}'].font = Font(color=AZUL)
for r in (14, 15):
    for col in 'FGHI':
        ws[f'{col}{r}'].font = Font(color=CINZA)
cab('F17:J17')
ws['F18'].font = Font(italic=True, color=CINZA)
for r in (19, 20, 21):
    ws[f'H{r}'].font = Font(bold=True)

# quadro mês a mês
ws[f'B{R_MES0 - 2}'].font = Font(bold=True, color=AZUL)
cab(f'B{R_MES0 - 1}:Y{R_MES0 - 1}')
ws.row_dimensions[R_MES0 - 1].height = 30
for r in range(R_MES0, R_SEM + 1):
    for col in 'BCDEFGHIJKLMNOPQRSTUVWXY':
        ws[f'{col}{r}'].border = Border(left=fina, right=fina, top=fina, bottom=fina)
        if r == R_SEM:
            ws[f'{col}{r}'].fill = PatternFill('solid', fgColor='F2F2F2')
ws.conditional_formatting.add(f'D{R_MES0}:D{R_SEM}', FormulaRule(formula=[f'D{R_MES0}=1'], fill=PatternFill('solid', bgColor='FFF2CC')))

# CAP do período
ws[f'B{R_CAPH - 1}'].font = Font(bold=True, color=AZUL)
cab(f'B{R_CAPH}:U{R_CAPH}')
ws.row_dimensions[R_CAPH].height = 30
last = R_CAP0 + N_CAP - 1
for r in range(R_CAP0, last + 1):
    for col in 'STU':
        ws[f'{col}{r}'].font = Font(color=CINZA)
    for col in 'MR':
        ws[f'{col}{r}'].font = Font(bold=True)
ws.auto_filter.ref = f'B{R_CAPH}:U{last}'
for txt, cor in (('Sobrecarregado', 'C00000'), ('Adequado', '008000'), ('Sub-utilizado', '1F70D6')):
    ws.conditional_formatting.add(f'M{R_CAP0}:M{last} R{R_CAP0}:R{last}',
                                  FormulaRule(formula=[f'$R{R_CAP0}="{txt}"'], font=Font(color=cor, bold=True)))
ws.freeze_panes = 'A3'

# aba de leitura
info = wb.create_sheet('LEIA-ME', 0)
linhas = [
    'PRÉVIA DO PAINEL DA CAP · dados do arquivo MFV_CQPA_2026_Oficial de 30/10/2026',
    '',
    'Este arquivo é um retrato: as abas de dados (CAP_PERIODOS, DEMANDA, calendários) são valores, não consultas.',
    'Serve para ver e testar o PAINEL. Ele não atualiza quando as bases mudam.',
    '',
    'Como usar: aba PAINEL, células C3 a C6 (Visão, Semana, Mês início, Mês fim). Os números mudam na hora.',
    'Mês único = CAP daquele mês. Intervalo = média simples dos meses. Dias úteis = calendário.',
    '',
    'Para ter o PAINEL no arquivo oficial, ligado às consultas: use o kit (MontarPainel.bas + pasta consultas),',
    'conforme o LEIA-ME.md do kit. As fórmulas da aba PAINEL são as mesmas deste arquivo.',
    '',
    'Semana disponível: só a 45 (é a única na base de quarentena do SharePoint hoje).',
]
for i, t in enumerate(linhas):
    info.cell(1 + i, 2, t)
info['B1'].font = Font(bold=True, size=14, color=AZUL)
info.column_dimensions['A'].width = 2
info.column_dimensions['B'].width = 110
info.sheet_view.showGridLines = False

wb.move_sheet('PAINEL', offset=-(len(wb.sheetnames) - 2))
wb.active = 1
wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print('ok', wb.sheetnames)
