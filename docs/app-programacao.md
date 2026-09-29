# Da planilha "Status + Programação V0" para um app de programação do CQ

Documento de trabalho: como a planilha funciona hoje, o que um sistema precisa
reproduzir, o que ele pode melhorar e uma proposta de telas, dados e roadmap.
Números tirados da versão carregada em 29/09/2026, via
`python -m programacao.extrair_programacao "Status + Programação V0.xlsx" dados_programacao`.

---

## 1. Como a planilha funciona hoje

### 1.1 Abas

| Aba | Papel | O que tem |
|---|---|---|
| **Fam 1 … Fam 8, Flavonid, Dipirona, FQ** (12 abas) | Fila e status de cada família | Blocos por produto → lotes → testes com status e turno programado |
| **Lista Analista Fam** | Escala do dia | Analistas por turno (1º, 2º, 3º) e onde estão: família, Curinga, Espectro, Dissolução, Férias, Atestado, Treinamento |
| **Equipamentos** | Quadro dos HPLC/UPLC/CG | Equipamento, família dona, produto/teste rodando, horário de término, próximo produto, calibração, responsáveis por turno |
| **Dissolução** | Programação da dissolução | Prioridades + lista ordenada de lotes por turno (1º, 2º, 3º e escala 6x2), com cluster e o que preparar (Amostra, PD, FM, UV) |
| Planilha1/2/3 | Vazias | — |

### 1.2 Estrutura de uma aba de família

```
Linhas 1–9   Cabeçalho livre: "PRIORIDADES" (J1:R8) e recados por turno (S1, S4, S7)
Linha 10     Sobre cada teste: "Prog." | "Backup"
Bloco produto 1
  Cabeçalho  Entrada CQ | Código | Produto | Lote PA | Lote PI | L T CQ | Data Lib CQ | Cluster Inicial | Cluster Atual | Teste 1 | … | Teste N | Observação
  "Padrão"   status padrão de cada teste
  Lote SA    lote do semiacabado/granel (Lote PI), onde sai a maior parte dos testes físico-químicos
  Lote PA    lotes do acabado ligados ao SA pelo Lote PI (normalmente só Descrição, ou "FEITO NO PI")
Bloco produto 2 …
```

- **257 blocos de produto** e **790 linhas de lote** (86 SA e 704 PA).
- Cada teste ocupa **3 colunas**: `Status | Prog. | Backup`.
  - `Prog.` = **turno programado** (1 amarelo, 2 azul, 3 laranja pela formatação condicional). Aparece junto com EM PREPARO / AG. ANÁLISE. Depois de concluído, vira `s`.
  - `Backup` = turno de cobertura ou confirmação de backup dos dados; também vira `s` quando o teste é concluído.
- `L T CQ` = `HOJE() − Entrada CQ` (lead time em dias).
- `Data Lib CQ`, `Cluster Inicial` e `Cluster Atual` vêm por **XLOOKUP de outra pasta de trabalho** (`[1]QUARENTENA`). Clusters: **Crítico DDE**, **Crítico PV**, **Estq Interno Abaixo**, **Ok**.
- Colunas A–C (ANL. 1º/2º/3º T) guardam o analista do lote por turno, mas quase não são usadas (28 registros).
- Passagem de turno por **comentário na célula** ("Pendência aos cuidados do 2º turno").

### 1.3 Ciclo de vida de um teste (status encontrados)

```
AG. ANÁLISE ──► EM PREPARO ──► HPLC / CG (em corrida) ──► AG. LEITURA ──► AG. CONFERÊNCIA ──► APROVADO
                                       │
                                       └─► FDE FASE 1A / 1B ─► INVESTIGAÇÃO / ROA / RO com GQ
Atalhos que encerram o teste: FEITO NO PI - OK · FEITO NA VALIDAÇÃO · PEGAR DO COA · N/A · ANÁLISE REDUZIDA
Outros: FIFO · Sem Amostra · AG. CONF. PROC (E2/FDE/OK) · AG. APROVAÇÃO PD · AG. CONF. MONTAGEM · LAUDOS DA AUDITORIA · "?"
```

Retrato atual (4.162 análises com status):

| Status | Qtde |
|---|---:|
| AG. ANÁLISE | 2.400 |
| APROVADO | 1.310 |
| EM PREPARO | 79 |
| HPLC | 75 |
| FEITO NO PI - OK | 63 |
| AG. CONFERÊNCIA | 55 |
| FDE FASE 1A + 1B | 56 |
| INVESTIGAÇÃO | 19 |
| Outros | 105 |

- **2.768 análises pendentes**, só **158 com turno programado**.
- Só **159 dos 728 lotes** estão com todos os testes concluídos: lote aprovado continua na aba e se mistura com a fila.
- Pendências por família: Doralgina 462 · Estéreis 459 · Dipirona 436 · Flavonid 337 · Torsilax 305 · Neolefrin 265 · FQ 160 · Líquidos 158 · Dermo 104 · Losartana 93.

### 1.4 O que a planilha resolve e onde ela sofre

| Funciona bem (manter no app) | Dor (o app resolve) |
|---|---|
| Visão "lote × teste" colorida, fácil de ler | Cada produto tem cabeçalho próprio; incluir um produto = copiar bloco à mão |
| Ligação SA → PA pelo Lote PI | Status em texto livre: `aprovado`/`APROVADO`, `S`/`s`, `2-`, `?`, nomes de teste soltos |
| Turno programado com cor | Não guarda histórico: quem mudou, quando, quanto tempo ficou em cada status |
| Prioridade por cluster (DDE/PV) | Link externo com QUARENTENA quebra fácil e não atualiza sozinho |
| Um lugar só para os três turnos | Várias pessoas editando o mesmo .xlsx; conflito de versão |
| | Prioridades e recados em texto livre no topo; passagem de turno em comentário |
| | Programação desconectada da capacidade: não sabe se cabe no turno, se o equipamento está livre, se tem analista |
| | Lotes aprovados continuam na aba (lotes de 2025 com 347 dias ainda aparecem) |

---

## 2. Modelo de dados do app

```
Familia 1───* Produto 1───* ProdutoTeste (catálogo: testes que o produto exige, técnica, TC em h, equipamento compatível)
                  │
                  └──* Lote (tipo SA|PA, lote_pa, lote_pi, entrada_cq, data_lib_prevista, cluster, prioridade)
                          │   PA ──(lote_pi)──► SA  (herança: teste feito no PI encerra o PA)
                          └──* Analise (lote × teste: status, turno/data programados, analista, equipamento, obs)
                                   └──* Evento (histórico: status de→para, quem, quando, comentário)

Analista (turno, família padrão, habilidades/técnicas) ──* Escala (data, turno, alocação: família|curinga|dissolução|ausência)
Equipamento (tipo, família dona, calibração, restrições) ──* Reserva (análises agrupadas em corrida, início, fim previsto)
Padrao (produto/teste → status do padrão de referência; bloqueia a análise)
Recado (turno, família, texto, lote/análise opcional, resolvido?)  ← substitui o topo da aba e os comentários
```

O extrator já gera este modelo a partir da planilha (`lotes.csv`, `analises.csv`,
`padroes.csv`, `escala.csv`, `equipamentos.csv`, `programacao_dissolucao.csv`). Esses
CSVs servem para a **carga inicial** do sistema.

Regras que o sistema passa a garantir:
- **Status fechado** (lista acima) com transições válidas; "?" deixa de existir e vira "sem definição" com alerta.
- **Status do lote** = derivado dos testes (todos concluídos → pronto para liberação; qualquer FDE/Investigação → bloqueado).
- **Herança SA → PA**: aprovar no SA marca os PAs ligados como "FEITO NO PI - OK".
- **Lead time** e **aging por status** calculados a partir do histórico (hoje só existe o total).

---

## 3. Telas

| # | Tela | Para quem | Equivale hoje a |
|---|---|---|---|
| 1 | **Programação do turno**: colunas 1º/2º/3º turno × analistas; arrastar análises pendentes para turno/analista/equipamento; barra de horas usadas × disponíveis | Líder/programador | Coluna "Prog." + Lista Analista |
| 2 | **Fila da família**: grid lote × teste com chips coloridos (mesmo visual da aba), filtros por cluster, status e produto; ações em massa ("programar 1º turno", "EM PREPARO") | Analistas e líder | Abas Fam |
| 3 | **Meu turno** (celular/tablet na bancada): o que está programado para mim, botão para avançar status, recado para o próximo turno | Analista | Olhar a aba e mudar a célula |
| 4 | **Lote**: todos os testes, SA e PAs ligados, linha do tempo, previsão de liberação × data pedida | Líder, PCP | Procurar o lote nas abas |
| 5 | **Equipamentos**: quadro de ocupação (Gantt do dia), término previsto, próximo, calibração vencendo, restrições ("amostrador não resfria") | Líder, analistas | Aba Equipamentos |
| 6 | **Dissolução**: fila ordenada por turno com o que preparar (Amostra/PD/FM/UV) | Equipe de dissolução | Aba Dissolução |
| 7 | **Passagem de turno**: pendências abertas, recados, o que foi programado e não iniciou | Todos | Comentários + topo da aba |
| 8 | **Painel**: WIP por família/status, aging, lotes críticos DDE/PV, programado × realizado, ritmo (Heijunka) e capacidade (FLOW LAB) | Coordenação | Heijunka + FLOW LAB |

---

## 4. Sugestão automática de programação

A programação continua sendo decisão do líder. O app monta uma **proposta** para cada
turno, que ele ajusta arrastando itens.

1. **Fila elegível**: análises pendentes com amostra disponível, padrão OK e sem bloqueio (FDE/Investigação).
2. **Prioridade** (pontuação configurável):
   - cluster: Crítico DDE > Crítico PV > Estq Interno Abaixo > Ok;
   - atraso em relação à data de liberação pedida (QUARENTENA);
   - aging (dias no CQ) e FIFO como desempate;
   - SA antes de PA (o SA destrava vários PAs).
3. **Agrupamento em corridas**: mesmo produto + mesmo teste/método entram juntos no HPLC (1 sequência, 1 padrão).
4. **Encaixe na capacidade** do turno:
   - horas por analista, usando o TC por teste da planilha MFV (`BASE TC`) já tratado no FLOW LAB;
   - equipamento compatível, livre (término previsto) e com calibração válida;
   - analista da família no turno, ou curinga.
5. **Resultado**: lista por turno/analista/equipamento, com o que ficou de fora e o motivo ("sem padrão", "sem HPLC livre", "sem analista").

Isso liga a programação diária ao trabalho de capacidade que já existe
(`claude/lab-capacity-analysis-system-tvegi3`: motor em Python/JS, TC e HC por família).

---

## 5. Opções de tecnologia

| Opção | Prós | Contras | Quando escolher |
|---|---|---|---|
| **A. Power Apps + SharePoint/Dataverse** | Já está no Microsoft 365; login corporativo; TI costuma aprovar mais rápido; Power BI para o painel | Grid lote × teste e arrastar-e-soltar são limitados; regra de sugestão mais difícil; Dataverse tem custo de licença | TI exige ficar no ecossistema Microsoft |
| **B. App web (React + API + PostgreSQL)** | Tela sob medida (grid, Gantt, arrastar); reaproveita o `engine.js` do FLOW LAB; tempo real entre turnos; histórico completo | Precisa de hospedagem e de alguém para manter; validação com TI | Liberdade de hospedagem (servidor interno ou nuvem aprovada) |
| **C. Evoluir o FLOW LAB (HTML estático)** | Já existe e funciona offline | Sem banco nem multiusuário: não serve para 3 turnos editando juntos | Só como protótipo de telas |

**Recomendação:** usar a opção **C para prototipar** as telas 1 e 2 com os dados reais
(dá para fazer rápido e mostrar para os líderes de turno) e construir o sistema em
**B**, ou em **A** se a TI exigir Microsoft. O modelo de dados da seção 2 é o mesmo
nos dois casos.

**BPF:** o app é ferramenta de **planejamento**. O registro GMP do resultado continua
no LIMS/Empower/caderno. Mesmo assim, histórico de alterações, login individual e
backup deixam o app pronto para uma validação (GAMP 5 cat. 4/5) se ele virar
referência oficial de status.

---

## 6. Roadmap

| Fase | Entrega | Critério de pronto |
|---|---|---|
| 0 | Extrator + modelo de dados (**feito**: `programacao/extrair_programacao.py`) | CSVs batem com a planilha |
| 1 | Protótipo navegável: Fila da família + Programação do turno com dados reais | Líderes dos 3 turnos validam o fluxo |
| 2 | MVP multiusuário: login, status com histórico, programação por turno, passagem de turno, importação da QUARENTENA | Um turno/família usa em paralelo com a planilha por 2 semanas |
| 3 | Equipamentos + Dissolução + sugestão automática de programação | Proposta automática aceita com poucos ajustes |
| 4 | Painel (WIP, aging, programado × realizado) integrado ao FLOW LAB | Planilha aposentada |

---

## 7. Perguntas para validar com o time

1. **Prog. e Backup**: `Prog.` = turno que vai executar, e `Backup` = turno de cobertura ou "backup dos dados feito"? O `s` quer dizer "concluído"?
2. **Linha "Padrão"** em cada bloco: é o status do **padrão de referência** do teste ou uma **linha modelo** copiada quando entra um lote novo?
3. **QUARENTENA**: de onde vem (SAP, outro Excel)? Tem exportação automática?
4. Existe **LIMS** (LabWare, STARLIMS…) ou Empower com status que o app possa ler, em vez de digitar de novo?
5. Quem programa: um líder por turno, um programador central ou cada família?
6. A TI permite app web próprio, ou precisa ser Power Platform?
7. Os TCs por teste da MFV (`BASE TC`) valem para programar o dia, ou é preciso um TC por método/equipamento?
