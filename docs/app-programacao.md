# Da planilha "Status + Programação V0" para um app de programação do CQ (Power Apps)

Documento de trabalho: como a planilha funciona hoje, o modelo de dados, as telas e a
arquitetura em Power Platform. Números da versão carregada em 29/09/2026, gerados por:

```bash
python -m programacao.extrair_programacao "Status + Programação V0.xlsx" dados_programacao
```

O comando também gera `dados_programacao/carga_powerapps.xlsx`, com uma tabela do Excel
por entidade, pronta para a carga inicial no Dataverse.

---

## 1. Como a planilha funciona hoje

### 1.1 Abas

| Aba | Papel | O que tem |
|---|---|---|
| **Fam 1 … Fam 8, Flavonid, Dipirona, FQ** (12 abas) | Fila, status e programação de cada família | Blocos por grupo de produto → lotes → testes |
| **Lista Analista Fam** | Escala do dia | Analistas por turno (1º, 2º, 3º) e alocação: família, Curinga, Espectro, Dissolução, Férias, Atestado, Treinamento |
| **Equipamentos** | Quadro dos HPLC/UPLC/CG | Equipamento, família dona, produto/teste rodando, término, próximo, calibração, responsáveis por turno |
| **Dissolução** | Programação da dissolução | Prioridades + fila ordenada por turno (1º, 2º, 3º e escala 6x2) com cluster e o que preparar (Amostra, PD, FM, UV) |
| Planilha1/2/3 | Vazias | — |

### 1.2 Estrutura de uma aba de família

```
Linhas 1–9   Texto livre: "PRIORIDADES" (J1:R8) e recados por turno (S1, S4, S7)
Linha 10     Sobre cada teste: "Prog." | "Backup"
Bloco (grupo de produto)
  Cabeçalho  Entrada CQ | Código | Produto | Lote PA | Lote PI | L T CQ | Data Lib CQ | Cluster Inicial | Cluster Atual | Teste 1 … Teste N | Observação
  "Padrão"   preparo de PADRÃO e FASE MÓVEL de cada teste do grupo (também é programado)
  Lote SA    semiacabado/granel (Lote PI): sai a maior parte dos testes físico-químicos
  Lote PA    acabado, ligado ao SA pelo Lote PI (normalmente Descrição, ou "FEITO NO PI")
Próximo bloco …
```

Cada teste ocupa **3 colunas**:

| Coluna | Quem preenche | Valores |
|---|---|---|
| **Status** | Analista/líder | AG. ANÁLISE, EM PREPARO, HPLC, AG. CONFERÊNCIA, APROVADO… |
| **Prog.** | Quem programa | **1** = programado para o turno |
| **Backup** | Quem executou (retorno) | **1** = feito · **C** = continuidade no próximo turno · **N** = não feito |

- O que tem `1` "na horizontal" (na linha do lote ou na linha Padrão) é o que está programado.
- `s` nas duas colunas é a marcação antiga dos testes já concluídos (1.123 células).
- `2-` na coluna Prog. (45 células, Dipirona/Doralgina/Flavonid/Neolefrin) parece "programado para o 2º turno". **A confirmar.**
- `L T CQ` = `HOJE() − Entrada CQ`. `Data Lib CQ` e os clusters vêm por **XLOOKUP de outro Excel (QUARENTENA)**. Clusters: Crítico DDE, Crítico PV, Estq Interno Abaixo, Ok.
- Passagem de turno por **comentário na célula** ("Pendência aos cuidados do 2º turno").

### 1.3 Ciclo de vida de um teste

```
AG. ANÁLISE ─► EM PREPARO ─► HPLC / CG ─► AG. LEITURA ─► AG. CONFERÊNCIA ─► APROVADO
                                  └─► FDE FASE 1A / 1B ─► INVESTIGAÇÃO / ROA / RO com GQ
Encerram o teste: FEITO NO PI - OK · FEITO NA VALIDAÇÃO · PEGAR DO COA · N/A · ANÁLISE REDUZIDA
Outros: FIFO · Sem Amostra · AG. CONF. PROC (E2/FDE/OK) · AG. APROVAÇÃO PD · AG. CONF. MONTAGEM · "?"
```

**Achado importante:** a programação é **por etapa e turno**, não pelo teste inteiro.
Das 160 análises programadas hoje, 74 já têm retorno "Feito", mas o status continua
EM PREPARO: o turno fez o preparo, e a corrida e a conferência serão programadas depois.
Um teste passa por vários "itens de programação" até ser aprovado, e a planilha só
guarda o último.

### 1.4 Retrato atual

| Indicador | Valor |
|---|---:|
| Grupos de produto (blocos) | 257 (132 com lotes) |
| Linhas de lote | 790 (86 SA, 704 PA) |
| Testes distintos no catálogo | 372 nomes em 834 combinações grupo × teste |
| Análises com status | 4.162 |
| Análises pendentes | 2.768 |
| Programadas agora (Prog. = 1) | 160 |
| Com retorno (Backup) | 74 feito · 0 continuidade · 0 não feito |
| Preparos de padrão/FM programados | 38 |
| Lotes com todos os testes concluídos (ainda na aba) | 159 de 728 |

Status das 4.162 análises: AG. ANÁLISE 2.400 · APROVADO 1.310 · EM PREPARO 79 ·
HPLC 75 · FEITO NO PI - OK 63 · AG. CONFERÊNCIA 55 · FDE 1A/1B 56 · INVESTIGAÇÃO 19 · outros 105.

### 1.5 O que manter e o que resolver

| Funciona bem (manter no app) | Dor (o app resolve) |
|---|---|
| Visão lote × teste colorida | Cada grupo tem cabeçalho próprio; produto novo = copiar bloco à mão |
| Programar marcando `1` e o executor dar retorno 1/C/N | O retorno é sobrescrito no dia seguinte: não existe histórico de aderência |
| Ligação SA → PA pelo Lote PI | Status em texto livre (`aprovado`/`APROVADO`, `?`, nomes de teste soltos) |
| Preparo de padrão/FM como item programável | Nada liga o preparo aos lotes que dependem dele |
| Prioridade por cluster (DDE/PV) | Vínculo externo com a QUARENTENA quebra e não atualiza sozinho |
| | Vários usuários editando o mesmo .xlsx |
| | "C" (continuidade) depende de o próximo turno ler a planilha |
| | Lotes aprovados continuam na fila |

---

## 2. Modelo de dados

A tabela central é o **Item de Programação**: um registro por coisa programada em um
turno de um dia. Ela substitui as colunas Prog./Backup e guarda o histórico.

```
Familia 1──* GrupoProduto 1──* TesteGrupo (catálogo: testes do grupo, técnica, TC)
                  │                 │
                  │                 └──* PreparoPadrao (status do padrão/FM do teste no grupo)
                  └──* Lote (SA|PA, lote_pa, lote_pi → Lote SA, entrada, data lib, cluster)
                          └──* Analise (lote × teste, status atual)

ItemProgramacao (data, turno, tipo = Análise | Preparo padrão/FM,
                 → Analise ou → PreparoPadrao, analista, equipamento,
                 programado_por, retorno = Feito | Continuidade | Não feito, motivo, obs)

Analista ──* Escala (data, turno, alocação)        Equipamento ──* ItemProgramacao
HistoricoStatus (análise, de → para, quem, quando)  Recado (turno, família, texto, resolvido)
Quarentena (espelho do Excel: lote, código, data lib, cluster inicial/atual)
```

| Tabela (Dataverse) | Campos principais | Origem da carga |
|---|---|---|
| `cq_familia` | nome, cluster de capacidade | abas Fam |
| `cq_grupoproduto` | nome, família, códigos | `grupos` |
| `cq_testegrupo` | grupo, teste, ordem, técnica, equipamento compatível, TC (h) | `catalogo_testes` + BASE TC da MFV |
| `cq_lote` | nº lote (chave), tipo, código, produto, lote PI → Lote SA, entrada CQ, data lib, cluster | `lotes` + QUARENTENA |
| `cq_analise` | lote, teste, status (escolha), concluído | `analises` |
| `cq_preparopadrao` | grupo, teste, status | `preparo_padrao` |
| `cq_itemprogramacao` | data, turno, tipo, análise/preparo, analista, equipamento, retorno, motivo | começa vazia (a carga traz os 160 + 38 atuais) |
| `cq_analista` / `cq_escala` | nome, turno padrão / data, turno, alocação | `escala` |
| `cq_equipamento` | tag, característica, família, calibração, restrições | `equipamentos` |
| `cq_recado` | data, turno, família, lote, texto, resolvido | novo |
| `cq_quarentena` | lote, código, data lib, cluster | fluxo de dados do Excel QUARENTENA |

Regras garantidas pelo sistema:
- **Status** como coluna de escolha (lista fechada, cores fixas). Mudança de status grava `HistoricoStatus`.
- **Retorno C** gera automaticamente um item para o próximo turno ("continuidade") e aparece no topo da fila dele.
- **Retorno N** exige motivo (sem amostra, sem padrão, equipamento, sem analista, prioridade mudou): vira indicador.
- **Herança SA → PA:** aprovar no SA marca as análises dos PAs ligados como FEITO NO PI - OK.
- **Dependência do padrão:** análise cujo preparo de padrão/FM não está pronto aparece como "aguardando padrão".
- **Lote concluído** (todos os testes encerrados) sai da fila e vai para "pronto para liberação".

---

## 3. Arquitetura em Power Platform

| Peça | Uso |
|---|---|
| **Dataverse** (ou Dataverse for Teams) | Banco com relacionamentos, auditoria nativa, segurança por papel, sem problema de delegação para 4 mil análises e histórico crescente |
| **Canvas app** (tablet/PC) | Programação do turno, Fila da família, Meu turno, Lote, Equipamentos, Dissolução, Passagem de turno |
| **Model-driven app** | Cadastros (grupos, catálogo de testes, equipamentos, analistas) sem desenhar tela |
| **Fluxo de dados (Power Query)** | Importa o Excel QUARENTENA do SharePoint/OneDrive para `cq_quarentena` em horário agendado. É a mesma linguagem das queries da MFV, então o time já domina. |
| **Power Automate** | Fim de turno: itens sem retorno → alerta ao líder; C → cria item no próximo turno; SA aprovado → propaga para os PAs |
| **Power BI** | WIP por família/status, aging, aderência (programado × feito), motivos de "não feito", ritmo Heijunka e capacidade (MFV/FLOW LAB) |

### Dataverse × listas do SharePoint

| | Dataverse | Dataverse for Teams | Listas do SharePoint |
|---|---|---|---|
| Licença | Power Apps Premium por usuário | Incluída no M365; app roda dentro do Teams | Incluída no M365 |
| Relacionamentos e regras | Nativos | Nativos | Colunas de pesquisa (limitadas) |
| Delegação / volume | Sem dor | Até 2 GB por ambiente, sobra para este caso | Cuidado: 2.000 itens por consulta não delegável, 5.000 no limite de exibição |
| Auditoria (quem mudou o quê) | Nativa | Limitada | Histórico de versões |
| Fluxo de dados do Excel | Sim | Sim | Não (só por Power Automate) |

**Recomendação:** começar em **Dataverse for Teams**, sem custo extra de licença e já com
tabelas relacionais e fluxo de dados. Se precisar de auditoria completa, ou de uso fora
do Teams, migrar para Dataverse Premium (mesmo modelo, a migração é direta). Listas do
SharePoint servem para um piloto pequeno, mas o volume de análises e o histórico vão
bater nos limites de delegação.

**BPF:** o app é ferramenta de planejamento. O resultado analítico continua no
Empower/caderno/laudo. Login individual e histórico de status deixam o app pronto para
validação (GAMP 5) se o status dele passar a ser referência oficial.

---

## 4. Telas do canvas app

| # | Tela | Para quem | Substitui | Como funciona no Power Apps |
|---|---|---|---|---|
| 1 | **Fila da família** | Líder, analistas | Abas Fam | Galeria de lotes (filtro por família, cluster, status) com galeria horizontal de testes dentro; cada teste é um "chip" colorido pelo status. Caixa de seleção nos chips. |
| 2 | **Programação do turno** | Quem programa | Coluna Prog. | Escolhe data + turno; seleciona chips na Fila (e preparos de padrão/FM) → botão **Programar** → escolhe analista e equipamento. Barra de horas programadas × disponíveis (TC × HC do turno). |
| 3 | **Meu turno** | Analista (tablet na bancada) | Coluna Backup | Lista do que foi programado para mim hoje, com botões **Feito · Continuidade · Não feito** (motivo obrigatório) e **Avançar status**. |
| 4 | **Passagem de turno** | Todos | Comentários + topo da aba | Continuidades abertas, "não feitos" com motivo, recados, prioridades do dia. |
| 5 | **Lote** | Líder, PCP | Procurar nas abas | Testes do lote, SA e PAs ligados, histórico de status e de programação, data lib × previsão. |
| 6 | **Equipamentos** | Líder | Aba Equipamentos | Quadro com o que está rodando, término previsto, próximo, calibração vencendo (alerta ≤ 30 dias). |
| 7 | **Dissolução** | Equipe de dissolução | Aba Dissolução | Mesma lógica de programação, com o campo "o que preparar" (Amostra/PD/FM/UV). |
| 8 | **Painel** | Coordenação | Heijunka | Power BI incorporado. |

Trechos de Power Fx (ideia da implementação):

```powerfx
// Programar os testes selecionados (Tela 2)
ForAll(
    colSelecionados As s,
    Patch('Itens de Programação', Defaults('Itens de Programação'), {
        Data: dpData.SelectedDate,
        Turno: ddTurno.Selected.Value,
        Tipo: 'Tipo (Itens de Programação)'.Análise,
        Análise: LookUp('Análises', 'Análise' = s.'Análise'),
        Analista: ddAnalista.Selected,
        Equipamento: ddEquipamento.Selected,
        'Programado por': User().FullName
    })
);
Clear(colSelecionados)

// Meu turno (Tela 3)
Filter('Itens de Programação',
    Data = Today() && Turno = varTurno && Analista.Email = User().Email)

// Retorno "Continuidade": grava e cria o item do próximo turno
Patch('Itens de Programação', ThisItem, { Retorno: 'Retorno'.Continuidade });
Patch('Itens de Programação', Defaults('Itens de Programação'), {
    Data: If(varTurno = 3, DateAdd(Today(), 1), Today()),
    Turno: Mod(varTurno, 3) + 1,
    Tipo: ThisItem.Tipo, Análise: ThisItem.Análise,
    Observação: "Continuidade do " & varTurno & "º turno"
})

// Cor do chip de status (Fila da família)
Switch(ThisItem.Status,
    "APROVADO", RGBA(0,176,80,1),  "EM PREPARO", RGBA(255,192,0,1),
    "HPLC", RGBA(204,204,0,1),     "AG. CONFERÊNCIA", RGBA(112,48,160,1),
    "FDE FASE 1A", RGBA(255,0,0,1), "FDE FASE 1B", RGBA(255,0,0,1),
    RGBA(0,51,102,1))
```

---

## 5. Sugestão automática de programação (fase 3)

A programação continua sendo do líder. O app monta uma **proposta** para o turno, que
ele confirma ou ajusta:

1. **Elegíveis:** análises pendentes com amostra, sem bloqueio (FDE/Investigação), mais as continuidades (C) do turno anterior, sempre no topo.
2. **Prioridade:** Crítico DDE > Crítico PV > Estq Interno Abaixo > Ok; depois atraso sobre a data lib, aging e FIFO; SA antes de PA.
3. **Agrupar por padrão/FM:** lotes do mesmo grupo e teste entram juntos. Se o preparo do padrão/FM não está pronto, ele é programado primeiro ou no mesmo turno.
4. **Encaixar na capacidade:** horas do turno (analistas da família + curingas na escala) × TC dos testes (BASE TC da MFV); equipamento compatível, livre e calibrado.
5. **O que sobrou** aparece com motivo (sem analista, sem HPLC, padrão pendente).

Pode ser feito em Power Fx, para regras simples, ou num fluxo Power Automate que chama
uma função. O motor do FLOW LAB (`claude/lab-capacity-analysis-system-tvegi3`) serve de
referência para TC e HC.

---

## 6. Roadmap

| Fase | Entrega | Critério de pronto |
|---|---|---|
| 0 | Extrator + modelo + arquivo de carga (**feito**) | Tabelas batem com a planilha |
| 1 | Ambiente Dataverse for Teams, tabelas da seção 2, carga inicial, fluxo de dados da QUARENTENA | Dados de uma família conferidos com a planilha |
| 2 | Telas 1, 2, 3 e 4 (Fila, Programação, Meu turno, Passagem) para **uma família piloto** | Os 3 turnos da família usam por 2 semanas em paralelo com a planilha |
| 3 | Todas as famílias + Equipamentos + Dissolução + cadastros no model-driven | Planilha só para consulta |
| 4 | Power BI (aderência, aging, WIP) + sugestão automática | Planilha aposentada |

---

## 7. Perguntas ainda abertas

1. **`2-` na coluna Prog.**: é "programado para o 2º turno"? A programação é sempre para o turno seguinte ou pode ser para vários turnos à frente?
2. **Quem programa**: um líder por turno para todas as famílias, ou cada família programa a sua?
3. **QUARENTENA**: onde fica o Excel (SharePoint, OneDrive, rede)? Tem tabela formatada? Com que frequência é atualizado? (Define se o fluxo de dados roda de hora em hora ou por turno.)
4. **Licença**: vocês têm Power Apps Premium, ou só as licenças do M365? E usam Teams? (Define Dataverse × Dataverse for Teams.)
5. **Família piloto**: qual família começar? Sugestão: uma com volume médio e bastante programação hoje (Losartana ou Torsilax).
