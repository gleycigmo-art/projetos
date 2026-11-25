# Diagnóstico – Status  Programação V0 - LIMPA.xlsx

## Causa principal: formatação condicional duplicada
Cada vez que linhas são copiadas/coladas ou arrastadas, o Excel duplica as regras de
formatação condicional. Com o tempo o arquivo acumulou:

| Item | Original | Otimizado |
|---|---|---|
| Regras de formatação condicional (total) | ~359.000 | ~9.100 |
| Regras só na aba "Fam - Torsilax" | 318.040 | 4.776 |
| Estilos de formatação condicional (dxf) | 359.595 | 37 |
| styles.xml (interno) | 39 MB | 0,2 MB |
| Arquivo descompactado (o que o Excel carrega) | ~90 MB | ~8,6 MB |
| Arquivo .xlsx | 5,3 MB | 1,0 MB |

A aba "Fam - Torsilax" tinha 40 MB de XML (as outras ~1 MB). Isso explica a lentidão para
abrir, travamentos, fechamento sozinho e perda de alterações na coautoria/salvamento.

Na versão OTIMIZADA as regras foram consolidadas preservando a ordem de prioridade por
célula. Foi verificado célula a célula (todas as abas, 34 valores de status testados) que
as cores resultantes são idênticas, e que valores e fórmulas não mudaram.

## Outros problemas encontrados
1. **Recuperação automática desativada** neste arquivo (`autoRecover="0"`). Se o Excel
   fecha sozinho, nada é recuperado. **Corrigido.**
2. **~11.300 fórmulas PROCX para outro arquivo** (`G:\Controle de Qualidade - Quarentena\Quarentena CQPA V1.xlsx`)
   usando colunas inteiras (`QUARENTENA!$I:$I`), com atualização de vínculos sempre ao
   abrir. **Ainda não alterado** – ver "Próximos passos".
3. **HOJE() repetido em ~4.000 células** (coluna "L T CQ"). **Corrigido:** a data fica em
   uma única célula (aba oculta `Parametros`, célula B1, nome definido `DataHoje`) e a
   coluna "L T CQ" usa `DataHoje`. Não apagar essa aba. O ganho de velocidade é pequeno:
   as células que dependem da data continuam recalculando a cada edição.
4. **Fórmulas com #REF!** – **corrigidas** (detalhe abaixo).
5. 4 nomes definidos quebrados (`_xleta.*` = #NAME?). **Removidos.**
6. Aba oculta vazia "Planilha2" (a "Planilha1" virou `Parametros`).

## Fórmulas #REF! corrigidas
Nas abas visíveis, nenhuma busca de dados estava quebrada: as 93 fórmulas com #REF!
eram sobras coladas em linhas de **cabeçalho** e em linhas **"Padrão"**, onde não deveria
haver fórmula.

| Aba | Tipo de linha | Células | O que foi feito |
|---|---|---|---|
| Fam - Torsilax | cabeçalho | J295 | texto do cabeçalho restaurado |
| Fam  - Flavonid | linha Padrão | L12, L25, L96, L140, L184, L208, L220, L246, L268, L288, L400, L431, L442 | fórmula removida (a célula fica vazia, como nas outras linhas "Padrão") |
| Fam - Doralgina | linha Padrão | L77, L93, L150, J203, L245, L345, J363, J381, J453, L503 | fórmula removida (a célula fica vazia, como nas outras linhas "Padrão") |
| Fam - Doralgina | cabeçalho | L149 | texto do cabeçalho restaurado |
| Fam  - Dipirona | linha Padrão | L12, L36, L57, J201, K201, L201, K234, L234, L260, L272, L332, L367, L390, L405, L420, L436, L447 | fórmula removida (a célula fica vazia, como nas outras linhas "Padrão") |
| Fam  - Dipirona | cabeçalho | K200, L200, J233, L233 | texto do cabeçalho restaurado |
| Fam - Neolefrin | linha Padrão | J23, L37, J70, J107, L135, J215, J236, J259, J306, J324, L356, L366, L378, L394, L408, L426, L434, L449, L459, L485 | fórmula removida (a célula fica vazia, como nas outras linhas "Padrão") |
| Fam - Neolefrin | cabeçalho | J152, J188 | texto do cabeçalho restaurado |
| Fam - Losartana | linha Padrão | J12, L31, J127, L154, L197, J211, J235, J249, J291, J309, J394, J411, J427, J441, J468, J478, J491, J502, J514 | fórmula removida (a célula fica vazia, como nas outras linhas "Padrão") |
| Fam - Losartana | cabeçalho | K290, L290, K308, L308, K467, L467 | texto do cabeçalho restaurado |

Além disso:
- **Fam - Neolefrin, linha 88:** as fórmulas estavam deslocadas uma coluna para a direita
  (lead time na coluna J, buscas usando o Lote PI). Realinhadas para o padrão das outras
  linhas (I88:L88).
- **Aba oculta "Fam 8 - Imp":** as 831 buscas (J, K e L) estavam todas com #REF! e
  receberam a fórmula padrão das outras abas. Hoje não retornam nada porque a aba não
  tem lotes preenchidos. Se ela não é mais usada, o melhor é excluí-la.

Na primeira abertura o Excel recalcula tudo uma vez (alguns segundos) e pede para salvar.

## Próximos passos sugeridos
- **Buscas da Quarentena via Power Query:** criar uma consulta que traz só as 5 colunas
  usadas (Código, Lote PA, Cluster Disp. Inicial, Cluster Disp. Atual, Previsão DU) para
  uma aba oculta e apontar os PROCX para essa tabela. Elimina o vínculo externo e as
  colunas inteiras, mantendo o formato das abas.
- **Padronizar a formatação condicional:** a legenda de cores é a mesma em todas as abas
  (33 condições; só o "N" aparece com duas cores). Dá para ter um único conjunto de
  ~35 regras por aba, em vez de centenas/milhares, e um script que reaplica esse padrão.

## Sobre o erro ao compartilhar na pasta de rede
A mensagem vem do recurso antigo "Compartilhar Pasta de Trabalho (Legado)", que não aceita
Tabelas do Excel (a aba "Equipamentos" tem a Tabela "Table1"). Esse modo legado é
descontinuado pela Microsoft e é conhecido por corromper arquivos e perder alterações.
**Não recomendado.**

Recomendado: salvar o arquivo no SharePoint / Teams / OneDrive e usar **coautoria**
(várias pessoas editando ao mesmo tempo, com salvamento automático e histórico de
versões). Com o arquivo enxuto, a coautoria fica estável.
