# Análise da CAP proposta (MFV_CQPA_2026)

Material de apoio para entender e simplificar o arquivo `MFV_CQPA_2026_Oficial. 30-10-26.xlsb`.

| Arquivo | Conteúdo |
|---|---|
| `painel_cap.html` | Protótipo do painel: filtro Semanal/Mensal, indicadores do rodapé em cartões, complexidade, CAP com status de carregamento, racional do cálculo, achados, sugestões e validação. Abre direto no navegador. |
| `consultas_power_query.m` | As 13 consultas Power Query extraídas do arquivo (texto original, sem alteração). |
| `scripts/cap_engine.py` | Réplica em Python das consultas `FILTRO_PERIODO → DEMANDA_MPR → BASE_TC_CAMP → TC_COM_CAMPANHA → CAP_PROPOSTA` e da fórmula de HRS TURNO da `Tabela_FAM`. |
| `scripts/load_tables.py` | Lê as tabelas do `.xlsb` com os valores que o Excel salvou (via `pyxlsb`). |
| `scripts/validate.py` | Compara a CAP recalculada com a CAP gravada no arquivo, célula a célula. |
| `scripts/painel_data.py` | Recalcula o rodapé e a complexidade e confere contra o arquivo; gera os dados de todos os períodos. |
| `scripts/build_painel.py` | Monta o `painel_cap.html` a partir do template. |

## Resultado da validação (arquivo de 30/10, semana 45)

- CAP: 55 linhas × 14 colunas = 770 valores, **0 diferenças** (tolerância 1e-9).
- Rodapé (13 indicadores), bloco "Mês Setembro" (7) e complexidade semanal/setembro (8): **28 de 28 iguais**.
- A visão mensal usa o mesmo código; não há CAP mensal salva no arquivo atual para comparar. Para conferir, troque o filtro para Mensal Set/2026, atualize e compare com a tabela da aba *Validação* do painel.

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
python3 painel_data.py tabelas.pkl painel.json
```

```bash
python3 build_painel.py tabelas.pkl painel_template.html ../painel_cap.html
```

`load_tables.py` usa os endereços das tabelas deste arquivo (ex.: `CAP_PROPOSTA` em `A4:P59`). Se uma tabela crescer, ajuste o dicionário `REFS`. `build_painel.py` interrompe com erro se a CAP da semana deixar de bater com o arquivo.
