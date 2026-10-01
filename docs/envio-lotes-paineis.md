# Envio automático dos lotes para os painéis (sem fórmulas)

Arquivo: `Programação_MP_Famílias_V.2.xlsx`
Macro: [`macros/modEnvioPaineis.bas`](../macros/modEnvioPaineis.bas)

## Por que trocar a fórmula

Os painéis (`PAINEL - BANCADA`, `PAINEL - CG`, `PAINEL - HPLC`) eram ao mesmo tempo
**calculados** (fórmula `FILTRO`) e **editados à mão** (a analista arrasta os lotes
para mudar a prioridade). Ao arrastar, a fórmula vai junto e passa a apontar para
outra data e outra família. No arquivo atual isso já aconteceu, por exemplo:

| Célula | O que a fórmula faz hoje |
|---|---|
| `PAINEL - BANCADA!FW74` | lê a data de `EU$4` e a família de `$B82` (outra coluna e outra linha) |
| `PAINEL - CG!BA5` | filtra a coluna `L` de ROTAS E FAMILIAS, que hoje é **FAM IFA** e não Teste CG |
| `PAINEL - CG!EZ14` | filtra a coluna `N`, que é **Teste HPLC** |

## Como a macro funciona

Um botão **"Enviar lotes"** roda `EnviarLotesParaPaineis`, que:

1. Lê a tabela `Tabela3_1` da aba **ROTAS E FAMILIAS**, localizando as colunas
   pelo **nome do cabeçalho** (incluir ou remover colunas não quebra nada).
2. Para cada lote, e para cada família marcada na linha dele:
   - famílias de bancada → `PAINEL - BANCADA`, data da coluna **D+8 (-2 dias)**;
   - **Teste CG** → `PAINEL - CG` e **Teste HPLC** → `PAINEL - HPLC`, data da coluna **D+8 (-3 dias)**;
   - **Check doc** não tem painel e é ignorado.
3. Procura, no painel, a coluna daquela data (linha 4) e o bloco da família (coluna B).
4. Escreve o concatenado **como texto** na **primeira célula vazia** do bloco
   (3º turno preparo 1, 2, 3 → 1º turno → 2º turno).

Regras:

- **Nunca apaga nem move nada.** Só escreve em células vazias.
- **Não duplica.** Se o lote já está em qualquer data do bloco daquela família, ele
  não é enviado de novo. A comparação usa o trecho antes de `| D+8`
  (código + descrição + lotes), então a analista pode arrastar o lote para outro
  dia e o próximo envio não traz uma cópia de volta. Isso vale mesmo se a data
  mudar em ROTAS E FAMILIAS.
- **Pinta de cinza claro** (cor da legenda "Prog. em andamento", célula D2 do
  painel) cada célula preenchida pela macro. Assim a analista vê o que veio
  automático e troca a cor conforme ajusta. Desligue com `PINTAR_ENVIADOS = False`.
- **Ignora datas passadas** (configurável em `DIAS_RETROATIVOS`).
- **Bloco cheio:** se as 9 vagas da data estão ocupadas, o lote não entra e aparece
  no log como "Sem vaga", para a analista decidir onde colocar.
- Ao final, mostra um resumo e grava o detalhe na aba **LOG ENVIO**: o que foi
  enviado (com a célula) e o que não entrou (e por quê).
- `DesfazerUltimoEnvio` desfaz **um envio por clique**, do mais recente para o
  mais antigo. Apaga as células que aquele envio escreveu e devolve a cor
  anterior. Lotes que a analista já moveu ou editou são mantidos e marcados no
  log como "Não desfeito". O log **acumula** os últimos 30 envios (o mais
  recente no topo, coluna **Envio** = data/hora do clique). Um clique em
  "Enviar lotes" que não envia nada não apaga o histórico.
- O `PAINEL - BANCADA` está protegido **com senha**. Na primeira vez que precisar
  escrever nele, a macro **pede a senha** (uma vez por clique), escreve e protege
  de novo com a mesma senha e as mesmas permissões (formatar, classificar,
  filtrar). Se a senha não for informada, o painel é pulado e os lotes aparecem
  no log como "Painel protegido". Para não precisar digitar, preencha
  `SENHA_PAINEIS` no topo do módulo; nesse caso, qualquer pessoa que abrir o
  código verá a senha.

## Instalação (uma vez)

1. Abra o arquivo no **Excel para desktop (Windows)**.
2. `Arquivo > Salvar como` → tipo **Pasta de Trabalho Habilitada para Macro do Excel (*.xlsm)**.
3. `Alt + F11` → `Arquivo > Importar arquivo...` → selecione `modEnvioPaineis.bas`. Feche o editor.
4. **Congelar as fórmulas antigas:** `Alt + F8` → `ConverterFormulasEmValores` → Executar.
   Troca cada fórmula dos painéis pelo texto que ela mostra hoje. Assim o que já
   está programado fica parado no lugar.
5. **Botão:** na aba ROTAS E FAMILIAS, `Inserir > Formas` → desenhe um retângulo
   e escreva "Enviar lotes" → botão direito → **Atribuir macro...** →
   `EnviarLotesParaPaineis`. Se quiser, crie também um botão "Desfazer envio" →
   `DesfazerUltimoEnvio`.
6. Salve.

## Uso no dia a dia

1. Atualize a tabela de ROTAS E FAMILIAS como hoje (botão direito → Atualizar).
   Se preferir que o botão já faça isso, mude `ATUALIZAR_ROTAS_ANTES` para `True`.
2. Clique em **Enviar lotes**.
3. Confira o resumo e, se houver "Sem vaga", a aba **LOG ENVIO**.
4. A analista arruma a prioridade **arrastando livremente** nos painéis. Como não
   há fórmula, não há o que quebrar.

## Configuração (topo do módulo)

| Constante | Padrão | Para quê |
|---|---|---|
| `DIAS_RETROATIVOS` | `0` | `0` = envia de hoje em diante; `2` = aceita até 2 dias atrás |
| `ATUALIZAR_ROTAS_ANTES` | `False` | atualiza a consulta antes de enviar |
| `SENHA_PAINEIS` | `""` | senha dos painéis; vazio = a macro pergunta na hora. **Não versionar a senha no repositório**: preencha só na cópia importada no Excel |
| `PINTAR_ENVIADOS` | `True` | pinta as células enviadas com a cor da legenda `TEXTO_LEGENDA_AUTO` |
| `MapaFamilias()` | 9 famílias | família → painel → coluna de data. Família nova = uma linha nova aqui (o nome tem de ser igual ao da coluna B do painel) |

## Simulação com o arquivo atual (01/10/2026)

A lógica foi executada em Python contra o arquivo do repositório. As fórmulas
antigas foram tratadas como vazias, pois o arquivo não guarda o valor calculado:

| Resultado | Qtde |
|---|---|
| Enviados | 137 |
| Já estavam no painel (não duplicados) | 40 |
| Data passada (ignorados) | 63 |
| Sem vaga | 10 (Teste HPLC em 03/10; IR/NIR em 03/10 e 12/10) |

Todas as 9 famílias foram encontradas nos painéis e todas as datas de cabeçalho
foram reconhecidas.

## Ajuste necessário na planilha

No `PAINEL - BANCADA`, os cabeçalhos **CH4 = 30/07/2027** e **CI4 = 31/07/2027**
estão com o ano errado (deveriam ser 2026). Lotes dessas datas iriam para o log
como "Data não existe na linha 4 do painel".

## Limitações

- Macros **não rodam no Excel Online** (navegador). Quem clica no botão precisa
  usar o Excel desktop. Se a equipe trabalha só no navegador, a mesma lógica pode
  ser feita como **Office Script** (Automatizar > Novo script) com botão na planilha.
- `Ctrl + Z` não desfaz macro. Para isso existe o `DesfazerUltimoEnvio`.
