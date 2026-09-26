# Lógica do cálculo de capacidade — MFV CQPA 2026

Engenharia reversa das 13 queries Power Query da planilha
`MFV_CQPA_2026_Oficial - 25-09-26.xlsb`. O código M original de cada query está em
[`/powerquery`](../powerquery).

---

## 1. Visão geral do fluxo

```
 ENTRADAS (abas / tabelas)                      QUERIES                          SAÍDA
 ─────────────────────────                      ───────                          ─────
 FILTRO_DATA  (Tabela_FILTRO) ───────────► FILTRO_PERIODO ─┐
 DEMANDA      (Tabela_DEMANDA, mensal) ──► SRC_DEMANDA ────┼─► DEMANDA_MPR ─┐
 SharePoint "Base Quarentena tratada" ───► Tabela_DemandaSemanal ┘            │
                                                                              │
 BASE TC      (Tabela_BASE_TC) ──┐                                            │
 CAMPANHA     (Tabela_BD_CAMPANHA)├─► BASE_TC_CAMP ─┐                         │
 SKIP_IMPUREZA(Tabela_SKIP_IMP) ──┘                 ├─► TC_COM_CAMPANHA ◄─────┘
 ETAPAS_CAMPANHAS (Tabela_Lotes_Campanha) ─► FONTES_TC ┘   (carga em min/mês
                                                           por família/processo)
                                                                 │
 FAMÍLIAS     (Tabela_FAM: turnos, pessoas, OEE) ─► FAM_PROPOSTA ┴─► CAP_PROPOSTA
                                                                     (mapa de capacidade)
```

Auxiliares: `METODOS_EQUIPAMENTO` (carga de HPLC/CG por método, sem agrupar),
`CAPACIDADE_EQUIPAMENTO` (capacidade dos HPLC/CG) e `SOMA TEMPOS TOTAIS`
(tempo total por MPR, com os métodos "Check" zerados).

---

## 2. Entradas

| Tabela | O que contém | Chave |
|---|---|---|
| **Tabela_DEMANDA** | Lotes previstos por **código de produto** e mês (Jan–Dez/2026), com o MPR de cada código e a flag `CONSIDERAR?` | `CÓD` → `MÉTODO DOCNIX` (MPR) |
| **Tabela_DemandaSemanal** (SharePoint) | Lotes por código e semana (`Chave semana` = AAAASS) | `Cod` → `MÉTODO DOCNIX` |
| **Tabela_BASE_TC** | Uma linha por **MPR × teste**, com o tempo de ciclo (min) de cada etapa 1–16, a família proposta, o equipamento proposto (CRLQ/CRGS) e o nº de injeções | `Método e versão (DocNix)` |
| **Tabela_BD_CAMPANHA** | Lotes analisados juntos por campanha, por MPR | `CÓD` (= MPR) |
| **Tabela_SKIP_IMP** | Skip test: 1 lote testado a cada N | `MPR` |
| **Tabela_Lotes_Campanha** | % do tempo da etapa que se repete nos lotes seguintes da campanha | etapa |
| **Tabela_FAM** (aba FAMÍLIAS) | Por família/processo: turnos, pessoas por turno, horas por turno, OEE | `FAMILIA` + `PROCESSO` |
| **Tabela_FILTRO** | Nível de visão (Mensal/Semanal), mês início/fim, semana | — |

Valores atuais de `% tempo` para campanha:

| Etapa | % repetido nos lotes seguintes |
|---|---|
| 1 – Preparo (padrão e soluções) | 0 % (só o 1º lote paga) |
| 6 – Finalização dos dados | 50 % |
| 7 – Conferência do preparo | 50 % |
| 9 – Montagem do equipamento | 10 % |
| 10 – Conferência do equipamento | 10 % |
| 12 – Tempo de corrida do padrão | 10 % |

---

## 3. Passo a passo do cálculo

### 3.1 Demanda por MPR — `DEMANDA_MPR`

A demanda nasce por **código de produto** e é somada por **MPR**, porque os tempos
estão no MPR e um MPR atende vários códigos.

- **Mensal:** filtra `CONSIDERAR? = "SIM"`, soma os meses do intervalo escolhido e
  - `Demanda_Total` = lotes somados no período
  - `Num_Meses` = nº de meses do intervalo
  - `Demanda_MPR` = `Demanda_Total / Num_Meses` (média mensal)
- **Semanal:** pega só a semana escolhida da base do SharePoint; `Num_Meses = 1`.

### 3.2 Carga de trabalho — `TC_COM_CAMPANHA`

Para cada linha MPR × teste, a query calcula `CM_Min` (**carga média em minutos por
período**). A regra é a mesma em todos os blocos:

```
lotes      = Demanda_Total do MPR
lotesXcamp = ARREDONDAR.PARA.CIMA(lotes por campanha)
nCamp      = ARREDONDAR.PARA.CIMA(lotes / lotesXcamp)       (0 se não há campanha)

se é Skip Test   → CM = tc × ARRED.CIMA(lotes / N_skip) / meses
senão se nCamp=0 → CM = tc × lotes / meses                  (sem campanha)
senão se nCamp=1 → CM = tc / meses                          (tudo em 1 campanha)
senão            → CM = [nCamp × tc + (lotes − nCamp) × tc × %etapa] / meses
```

A campanha funciona assim: o **primeiro lote** de cada campanha paga o tempo cheio da
etapa, e os demais pagam só o `%` da tabela de campanha. Nas etapas sem % cadastrado,
todos os lotes pagam o tempo cheio.

As etapas são agrupadas em **processos** (os "recursos" do laboratório):

| Bloco | Processo (recurso) | Etapas somadas | Agrupado por |
|---|---|---|---|
| A | **MO BANCADA** | 1 + 2 + 3 + 6 (campanha nas 1 e 6) | Família proposta |
| B | **MO CONFERÊNCIA BANCADA** / **MO CONFERÊNCIA PI/TF** | 7 (separada entre Fam PI/TF e as demais) | — |
| C | **MO EQUIPAMENTO BANCADA** | espectro: AA + ICP + IV + UV (sem campanha) | — |
| D | **MO EQUIPAMENTO** | 9 (com campanha) + 14 | — |
| E | **MO CONFERÊNCIA EQUIPAMENTO** | 10 (com campanha) + 15 | — |
| F | **REVISÃO FINAL** | 16 | — |
| G | **DISSOLUTOR** | 8 | — |
| H | **EQUIPAMENTO HPLC / CG** | (nº inj. padrão × t. padrão) + (nº inj. amostra × t. amostra); campanha só na parte do padrão | Equipamento proposto (CRLQ/CRGS) |

As etapas **4 (repouso)**, **5 (instrumentos)** e **11 (condicionamento)** não entram
na capacidade. Só aparecem na query `SOMA TEMPOS TOTAIS`.

### 3.3 Capacidade disponível — aba FAMÍLIAS (`Tabela_FAM`)

Estas colunas são fórmulas de planilha, não Power Query. Os valores conferem com:

```
RECURSO MÉDIO ATUAL = média(QTDE 1º, 2º, 3º TURNO)             → pessoas médias por turno
HRS TURNO           = média das horas/mês do calendário do período
                      (coluna MO-3 para mão de obra, EQP-3 para equipamento)
TOTAL RECURSO (h)   = RECURSO MÉDIO ATUAL × HRS TURNO           → horas-recurso/mês
RECURSO DISP (min)  = TOTAL RECURSO × 60 × OEE REAL
```

Exemplo, Dermo/Semissólidos: (1,76 + 1,79 + 1,31)/3 = 1,62 pessoas × 524,9 h =
**850,2 h**. Com OEE de 0,8, sobram 40.810 min disponíveis.

O calendário desconta dias úteis, feriados (turno e ADM) e férias, e tem versão
mensal e semanal. Mão de obra usa OEE fixo de 0,80. Cada HPLC usa o seu OEE real
medido, de 0,21 a 0,50.

### 3.4 Mapa de capacidade — `CAP_PROPOSTA`

Junta a carga (3.2) com a capacidade (3.3) por **Família + Processo**:

| Indicador | Fórmula | Leitura |
|---|---|---|
| **CM (HRS)** | Σ CM_Min / 60 | horas de trabalho necessárias no período |
| **OEE NEC** | CM (HRS) / TOTAL RECURSO | eficiência que seria necessária |
| **CARREGAMENTO** | OEE NEC / OEE REAL | < 0,80 sub-utilizado · 0,80–1,00 adequado · > 1,00 sobrecarregado |
| **HC NECESSÁRIO** | CM (HRS) / (HRS TURNO × OEE REAL) | pessoas (ou equipamentos) necessários por turno |
| **Soma Demanda Média** | Σ Demanda_MPR dos MPRs da família/processo | lotes/mês |
| **TAKT** | TOTAL RECURSO / Demanda | horas disponíveis por lote |
| **TC MÉDIO** | CM (HRS) / Demanda | horas gastas por lote |
| **FIFO** | Demanda / 22,97 | lotes/dia (ritmo diário, usado no HEIJUNKA) |

Nota: `CARREGAMENTO = HC NECESSÁRIO / RECURSO MÉDIO ATUAL`, ou seja, pessoas
necessárias divididas pelas pessoas que existem.

---

## 4. Pontos de atenção encontrados

Estes pontos precisam ser **validados com o negócio antes** de levar a lógica para o
sistema:

1. **Salto quando `nCamp = 1`.** Se todos os lotes cabem numa campanha, a carga
   inteira do MPR vira `tc / meses`: o tempo é cobrado uma vez só, inclusive nas
   etapas 2 e 3 (preparo da amostra e análise FQ) e na injeção de amostra do HPLC,
   que normalmente são por lote. Com um lote a mais, `nCamp` vira 2 e essas etapas
   voltam a ser cobradas lote a lote. Resultado: 12 lotes podem custar bem menos que
   13. É intencional?
2. **Constante 22,97 no FIFO.** Os dias por mês estão fixos no código. Na visão
   **semanal** o FIFO fica errado, porque deveria dividir por cerca de 5,6 dias. Na
   mensal o valor também não bate com os dias úteis do calendário (média de 24,75).
3. **Critério de Skip diferente entre blocos.** Nos blocos A–G, "é skip" significa
   só família = `Impureza Skip Test`. No bloco H (HPLC), também exige que o MPR esteja
   na tabela de skip.
4. **Campanha calculada por linha de teste.** A mesma campanha é aplicada a cada teste
   do MPR separadamente. Se os testes de um MPR compartilham o preparo, a carga de
   preparo pode estar sendo contada em dobro.
5. **Mão de obra usa OEE fixo de 0,80.** Os equipamentos usam o OEE real. Vale ter o
   OEE real também para as equipes.
6. **Horas por turno em média.** No modo mensal, as horas são a média dos meses do
   intervalo. Um mês com muitos feriados (nov/dez) fica diluído, e esconde picos.

---

## 5. Como isso vira o sistema

O modelo de dados sai quase pronto das tabelas atuais:

```
Produto (CÓD) ──N:1──► MPR ──1:N──► Teste(MPR) ──► tempos das etapas 1–16
                         │                         equipamento proposto
                         ├── lotes por campanha
                         └── N skip
Família ──► processo ──► pessoas por turno, turnos, OEE
Calendário (dia) ──► horas disponíveis por turno (feriados, férias)
Demanda (CÓD × dia/semana/mês) ──► lotes
```

O motor de cálculo é uma função pura:
`carga(demanda, tempos, campanha, skip) ÷ capacidade(pessoas, calendário, OEE)`.
Assim, **demanda e pessoas viram parâmetros de simulação**, sem precisar de
"Atualizar Tudo".

Para ter a **visão diária**: guardar o calendário por dia (horas × turnos) e a demanda
por dia. A demanda diária pode vir da base de quarentena, ou ser distribuída a partir
da semanal pelos dias úteis. Com isso, mês e semana passam a ser só agregações do dia,
e o FIFO/HEIJUNKA sai direto do cálculo, sem a constante 22,97.
