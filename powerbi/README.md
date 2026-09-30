# FLOW LAB no Power BI

Cards HTML de capacidade do laboratório: tudo o que é indicador é calculado **dentro da
própria medida** (em `var`), sem medidas auxiliares. O Power Query só prepara as tabelas
com a carga e a capacidade de cada mês e de cada semana.

```
Planilha MFV (.xlsx) ──► Power Query (powerbi/queries) ──► dPeriodo · fCap · fDemanda · dRecurso · fAusencias
                                                              │
                         Tabelas de simulação (DAX) ──────────┤
                                                              ▼
                         card_kpis_capacidade  ·  card_carregamento_familias  (visual HTML Content)
```

## 1. Fonte

Salve a planilha como **.xlsx**. O conector de Excel do Power BI nem sempre lê .xlsb;
se o seu ler, pode manter .xlsb. As tabelas nomeadas usadas são:
`Tabela_BASE_TC`, `Tabela_DEMANDA`, `Tabela_DemandaSemanal`, `Tabela_FAM`,
`CALENDARIO_MES`, `CALENDARIO_SEMANA`, `Tabela_BD_CAMPANHA`, `Tabela_SKIP_IMP` e
`Tabela_Lotes_Campanha`. Também são lidos os blocos de ausências (P/D … Total Ausências)
e de peso do dia útil, da aba `FAMÍLIAS`.

## 2. Power Query

Em *Transformar dados*, crie uma **Consulta Nula** para cada arquivo de `queries/`, na
ordem, com o **mesmo nome do arquivo sem o número** (ex.: `04_stg_BaseTC.m` → `stg_BaseTC`).
Cole o conteúdo no *Editor Avançado*.

| Consulta | Carregar no modelo? | Conteúdo |
|---|---|---|
| `CaminhoArquivo` | não (parâmetro) | caminho da planilha: **ajuste para o seu** |
| `Pasta`, `fnNum`, `fnTabela` | não | leitura e conversão |
| `stg_BaseTC`, `stg_Parametros`, `stg_Componentes`, `stg_PesoDia`, `stg_MprRecurso` | não | etapas intermediárias |
| **`dPeriodo`** | sim | 12 meses + 53 semanas ISO, dias úteis, horas do calendário, fonte da demanda |
| **`fDemanda`** | sim | lotes por período e produto, com complexidade |
| **`dRecurso`** | sim | pessoas por turno e OEE por família |
| **`fCap`** | sim | carga (CM em horas) × recurso por período, família e processo |
| **`fAusencias`** | sim | ausências da semana apontada e média do ano |

Para as consultas que não vão para o modelo: botão direito > desmarcar *Habilitar carga*.

Regras aplicadas (as mesmas do app, validadas): 1º lote de cada campanha com tempo
cheio e os demais com o % da etapa; skip só no teste de impureza de MPR da
`Tabela_SKIP_IMP`; semanas sem quarentena real usam o plano mensal rateado pelos dias
úteis.

## 3. Modelo

- Relacionamentos (1 → muitos, filtro simples):
  `dPeriodo[Periodo_ID]` → `fCap[Periodo_ID]` e `dPeriodo[Periodo_ID]` → `fDemanda[Periodo_ID]`.
- `dRecurso` e `fAusencias` ficam **sem** relacionamento.
- Classifique `dPeriodo[Rotulo]` pela coluna `dPeriodo[Ordem]`.
- Crie as tabelas de `dax/tabelas_simulacao.dax` (*Modelagem > Nova tabela*), cada uma
  com um segmentador de **seleção única**. Se não quiser simulação, apague as linhas
  `SELECTEDVALUE('Sim …')` das medidas e use `1` / `BLANK()` / `0` no lugar.

## 4. Cards

1. Instale o visual **HTML Content** (AppSource, de Daniel Marsh-Patrick).
2. Crie as medidas `dax/card_kpis_capacidade.dax` e `dax/card_carregamento_familias.dax`.
3. Coloque cada medida no campo *Values* de um visual HTML Content.
4. Adicione segmentadores de `dPeriodo[Nivel]` (Mensal/Semanal) e `dPeriodo[Rotulo]`
   (seleção única). Os cards pedem **um período por vez**, porque o carregamento nunca é
   média de vários períodos.

Para o card de equipamentos, duplique `card_carregamento_familias` e troque
`var categoria = "Mão de Obra"` por `"Equipamento"`.

## 5. Conferência

Valores esperados sem nenhuma simulação selecionada (mesmos números do app FLOW LAB):

| Indicador | Semana 38 (quarentena real) | Set/2026 |
|---|---|---|
| Ocupação global | 151,2% | 110,8% |
| Ocupação corrigida | 165,2% | 140,5% |
| HC necessário | 226 | 166 |
| HC disponível | 149 | 149 |
| Gap de HC | +76 | +16 |
| Ausência (HC) | 12,7 | 31,6 (média do ano) |
| Total de lotes | 461 | 1.716 |
| Lotes/dia | 82,6 (5,58 dias) | 68,7 (25,00 dias) |
| Complexidade 3+4 | 72,6% | 68,8% |
| Famílias de MO acima de 100% | 11 | 9 |
| 1ª família no card de barras | Sólidos - Neolefrin 379,3% | Sólidos - Neolefrin 234,3% |
| Ocupação corrigida com OEE MO = 0,60 | 220,3% | 187,4% |

## Limitações

- **Simulação de demanda no Power BI é linear**: o fator multiplica a carga já calculada.
  O app recalcula as campanhas com a nova quantidade de lotes, então ali a resposta é
  exata; no Power BI é uma aproximação (fica mais perto quanto menor o ajuste).
- **Remanejamento de pessoas** e **"quanto consigo atender"** ficaram só no app; em DAX
  exigiriam iteração.
- O código M foi validado quanto à sintaxe e a lógica foi conferida período a período
  contra o motor do app (65 períodos iguais), mas não foi executado dentro do Power BI.
  Se alguma consulta der erro ao carregar, a mensagem do Power Query indica a etapa.
