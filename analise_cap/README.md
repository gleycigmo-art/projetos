# Análise da CAP proposta (MFV_CQPA_2026)

Material de apoio para entender e simplificar o arquivo `MFV_CQPA_2026_Oficial. 30-10-26.xlsb`.

| Arquivo | Conteúdo |
|---|---|
| `painel_cap.html` | Protótipo do painel: filtro Semanal/Mensal, indicadores do rodapé em cartões, complexidade, CAP com status de carregamento, racional do cálculo, achados, sugestões e validação. Abre direto no navegador. |
| `excel_kit/` | **Kit para o Excel**: macro + consultas novas que criam a aba PAINEL com filtro instantâneo. Ver `excel_kit/LEIA-ME.md`. |
| `consultas_power_query.m` | As 13 consultas Power Query extraídas do arquivo (texto original, sem alteração). |
| `scripts/cap_engine.py` | Réplica em Python das consultas `FILTRO_PERIODO → DEMANDA_MPR → BASE_TC_CAMP → TC_COM_CAMPANHA → CAP_PROPOSTA` e da fórmula de HRS TURNO da `Tabela_FAM`. |
| `scripts/load_tables.py` | Lê as tabelas do `.xlsb` com os valores que o Excel salvou (via `pyxlsb`). |
| `scripts/validate.py` | Compara a CAP recalculada com a CAP gravada no arquivo, célula a célula. |
| `scripts/painel_data.py` | Recalcula o rodapé e a complexidade e confere contra o arquivo; gera os dados de todos os períodos. |
| `scripts/build_painel.py` | Monta o `painel_cap.html` a partir do template. |
| `scripts/gerar_consultas_v2.py` | Gera as consultas do kit a partir do texto original das consultas (troca exata de trechos, conferida). |
| `scripts/gerar_painel_excel.py` | Define a aba PAINEL num lugar só e gera a macro e os arquivos de teste. |
| `scripts/testar_painel_excel.py` | Compara a aba PAINEL recalculada no LibreOffice com o motor Python. |

## Resultado da validação (arquivo de 30/10, semana 45)

- CAP: 55 linhas × 14 colunas = 770 valores, **0 diferenças** (tolerância 1e-9).
- Rodapé (13 indicadores), bloco "Mês Setembro" (7) e complexidade semanal/setembro (8): **28 de 28 iguais**.
- Fórmulas da aba PAINEL do kit, recalculadas no LibreOffice (semana 45, setembro, julho a setembro): **2.367 valores, 0 diferenças** contra o motor Python.
- Não há CAP mensal salva no arquivo, e as consultas novas só rodam no Excel: a aba PAINEL tem uma conferência automática contra a CAP_PROPOSTA atual (ver `excel_kit/LEIA-ME.md`).

## Regras decididas em 30/10/2026

Mês filtrado = CAP daquele mês; intervalo = média simples dos meses. Dias úteis do calendário (semana e mês) para lotes/dia e FIFO. Espectro e CRGS0043 seguem fora da CAP; HRS TURNO segue inteiro no HC; 3 turnos.

## Como rodar de novo (quando o arquivo mudar)

```bash
pip install pyxlsb openpyxl
```

```bash
cd analise_cap/scripts
```

```bash
python3 load_tables.py "../../MFV_CQPA_2026_Oficial. 30-10-26.xlsb" tabelas.pkl
```

```bash
python3 validate.py tabelas.pkl
```

```bash
python3 painel_data.py tabelas.pkl periodos.json
```

```bash
python3 gerar_consultas_v2.py ../consultas_power_query.m ../excel_kit/consultas
```

```bash
python3 gerar_painel_excel.py vba MontarPainel_template.bas ../excel_kit/MontarPainel.bas
```

Teste das fórmulas do PAINEL (gera os 3 cenários, recalcula no LibreOffice e compara):

```bash
for c in "Semanal Set/2026 Set/2026 sem" "Mensal Set/2026 Set/2026 set" "Mensal Jul/2026 Set/2026 jul_set"; do set -- $c; python3 gerar_painel_excel.py teste tabelas.pkl periodos.json teste_$4.xlsx $1 $2 $3; done
```

```bash
soffice --headless --convert-to xlsx --outdir recalc teste_sem.xlsx teste_set.xlsx teste_jul_set.xlsx
```

```bash
for f in recalc/teste_*.xlsx; do mv "$f" "recalc/painel_${f#recalc/teste_}"; done
```

```bash
python3 testar_painel_excel.py periodos.json recalc teste_excel.json
```

```bash
python3 build_painel.py tabelas.pkl painel_template.html ../painel_cap.html teste_excel.json
```

`load_tables.py` usa os endereços das tabelas deste arquivo (ex.: `CAP_PROPOSTA` em `A4:P59`). Se uma tabela crescer, ajuste o dicionário `REFS`. `build_painel.py` interrompe com erro se a CAP da semana deixar de bater com o arquivo.
