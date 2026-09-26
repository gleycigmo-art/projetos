# Capacidade do laboratório de CQ

Réplica em Python das queries Power Query da planilha `MFV_CQPA_2026_Oficial`, com as
correções de regra combinadas, visão mês a mês, pior mês e a pergunta ao contrário
("quanto consigo atender com os recursos que tenho").

- Lógica e decisões: [`docs/logica-capacidade.md`](docs/logica-capacidade.md)
- Queries originais: [`powerquery/`](powerquery)

## Como rodar

```bash
pip install -r requirements.txt

# 1. Exportar as tabelas da planilha para dados/*.csv
python -m capacidade.extrair_excel "MFV_CQPA_2026_Oficial - 25-09-26.xlsb" dados

# 2. Gerar o relatório (usa o período da aba FILTRO_DATA)
python -m capacidade.relatorio dados resultados/relatorio_capacidade.xlsx
```

## Uso no Python (simulação)

```python
from capacidade.motor import Dados, Periodo, Regras, Cenario, cap_proposta, capacidade_maxima

d = Dados.carregar("dados")
base = cap_proposta(d)                                      # regras corrigidas
legado = cap_proposta(d, regras=Regras.legado())            # igual ao Excel atual
sim = cap_proposta(d, cenario=Cenario(
    fator_demanda=1.10,                                     # +10% de demanda
    pessoas={"Sólidos - Neolefrin": 3.0}))                  # 3 pessoas/turno
capacidade_maxima(d)                                        # quanto a MO aguenta
```

## FLOW LAB (app web)

Painel de capacidade semanal e mensal em um único arquivo HTML: `app/dist/flow-lab.html`.
Abra direto no navegador; não precisa de servidor.

| Tela | O que mostra |
|---|---|
| Painel executivo | 8 KPIs, carregamento por família, resumo copiável, insights |
| Mapa de capacidade | tabela por família/processo, filtro por cluster, marcação de intercambiáveis |
| Gargalos | matriz doadoras × receptoras, ranking de gap de HC, Pareto, equipamentos > 100% |
| Mão de obra | HC por turno e balanceamento, ausências, HC efetivo, quadro por família |
| Complexidade | índice ponderado, % 3+4, distribuição por nível |
| Tendências | evolução período a período, pior período, gargalo crônico, mapa de calor |
| Simulador | OEE (MO e equipamentos), demanda, remanejamento, quadro, ausências, lotes adiados; antes × depois e "quanto consigo atender" |
| Importação | lê a planilha MFV (.xlsb) no navegador, valida e recalcula tudo; parâmetros e alertas de cadastro |

- **Semana × mês:** seletor no topo. A semana 38 usa a quarentena real; as demais
  semanas usam a projeção do plano mensal, rateada pelos pesos de dia útil da aba FAMÍLIAS.
- **Regras:** corrigidas por padrão; em Importação dá para alternar para "Legado (Excel)".
- **Motor:** `app/src/engine.js` é a porta de `capacidade/motor.py`, e o teste de
  paridade garante os mesmos números.

```bash
python app/build.py            # gera app/dist/flow-lab.html e o gabarito de testes
node --test app/test/          # paridade JS × Python e regras do simulador
```
