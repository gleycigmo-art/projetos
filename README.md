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
