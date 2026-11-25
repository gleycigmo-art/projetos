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
   fecha sozinho, nada é recuperado. **Corrigido** na versão OTIMIZADA.
2. **~11.300 fórmulas XLOOKUP para outro arquivo** (`G:\Controle de Qualidade - Quarentena\Quarentena CQPA V1.xlsx`)
   usando colunas inteiras (`QUARENTENA!$I:$I`), e o arquivo está configurado para
   atualizar vínculos sempre ao abrir. Isso deixa a abertura lenta e não funciona no
   Excel Online / coautoria. **Não alterado** (exige decisão do processo).
3. **~1.000 células com HOJE()** (função volátil: recalcula tudo a cada edição).
4. **Fórmulas com #REF!** (vínculo quebrado), ex.: Fam - Dipirona J12, J36, J57;
   Fam - Neolefrin J23, J37, G70; Fam - Losartana J12, G31; Fam - Flavonid G12, G25;
   Fam - Doralgina J77, J93; Fam - Torsilax J295. A aba oculta "Fam 8 - Imp" tem 831.
5. 4 nomes definidos quebrados (`_xleta.*` = #NAME?). **Removidos** na versão OTIMIZADA.
6. Abas ocultas vazias "Planilha1" e "Planilha2".

## Sobre o erro ao compartilhar na pasta de rede
A mensagem vem do recurso antigo "Compartilhar Pasta de Trabalho (Legado)", que não aceita
Tabelas do Excel (a aba "Equipamentos" tem a Tabela "Table1"). Esse modo legado é
descontinuado pela Microsoft e é conhecido por corromper arquivos e perder alterações.
**Não recomendado.**

Recomendado: salvar o arquivo no SharePoint / Teams / OneDrive e usar **coautoria**
(várias pessoas editando ao mesmo tempo, com salvamento automático e histórico de
versões). Com o arquivo enxuto, a coautoria fica estável.
