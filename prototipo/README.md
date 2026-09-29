# Protótipo do app de programação do CQ

`programacao-cq.html` abre direto no navegador (arquivo único, com os dados da planilha
`Status + Programação V0.xlsx`). Simula as telas previstas para o Power Apps:

| Tela | O que faz |
|---|---|
| Programar | Fila por família (lote × teste, Padrão/FM), seleção de testes, programação para data/turno com analista e equipamento |
| Quadro dos turnos | O que foi programado em cada turno do dia e o retorno (Feito, Continuidade, Não feito) |
| Meu turno | Tela da bancada: retorno 1/C/N, motivo obrigatório no "não feito", atualização de status |
| Passagem de turno | Continuidades recebidas (com atribuição), não feitos, itens sem retorno, recados |
| Equipamentos | Quadro dos cromatógrafos com calibração vencendo |
| Painel | Pendências por família, aderência do dia, motivos de "não feito" |

As alterações ficam salvas só no navegador (localStorage). No app real, ficam no Dataverse.

Para regerar com uma planilha nova:

```bash
python -m prototipo.build_prototipo "Status + Programação V0.xlsx"
```
