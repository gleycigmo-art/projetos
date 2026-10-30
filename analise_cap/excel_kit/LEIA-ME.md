# Kit PAINEL da CAP (Excel)

Adiciona ao MFV_CQPA uma aba **PAINEL** com filtro instantâneo de semana, mês e intervalo de meses. A CAP é calculada para todos os períodos de uma vez: trocar o período não exige "Atualizar Tudo"; atualize só quando uma base mudar.

**Nada existente é alterado ou apagado.** O macro só adiciona consultas e abas novas.

## Conteúdo

| Arquivo | O que é |
|---|---|
| `MontarPainel.bas` | Macro que instala tudo (importar no editor VBA) |
| `consultas/*.m` | As 8 consultas Power Query novas, lidas pelo macro |

Consultas novas:

- `fnDEMANDA_MPR`, `fnTC_COM_CAMPANHA` e `fnCAP`: corpo das consultas atuais, copiado sem reescrever a lógica, em forma de função do período.
- `fnFAM_PROPOSTA`, `fnCALENDARIO_ATIVO` e `v2_FONTES_CAL`: horas do turno do período, com a mesma regra da coluna HRS TURNO.
- `v2_SEMANAS` e `v2_CAP_PERIODOS`: as semanas disponíveis e a CAP de cada mês e de cada semana.

## Regras (decididas em 30/10/2026)

1. **Mês filtrado:** CAP e indicadores daquele mês. **Intervalo filtrado:** média simples dos meses.
2. **Dias úteis:** calendário de cada mês e de cada semana (aba FAMÍLIAS). Vale para lotes/dia e FIFO.
3. **Espectro e CRGS0043:** continuam fora da CAP (sem recurso cadastrado).
4. **HRS TURNO:** continua arredondado para inteiro no HC necessário.
5. **Turnos:** × 3 para todas as famílias de mão de obra.

Únicas diferenças na CAP em si: **FIFO** (÷ dias úteis do período em vez de ÷ 22,97) e a **Categoria**, que passa a vir da consulta, com os mesmos valores.

## Como instalar

1. Abra uma **cópia** do `MFV_CQPA_2026_Oficial`.
2. Alt+F11 → Arquivo → Importar arquivo → `MontarPainel.bas`.
3. Alt+F8 → `MontarPainelCAP` → Executar → escolha a pasta `consultas` deste kit.
4. Aguarde. O cálculo de 12 meses mais as semanas pode levar alguns minutos; a barra de status mostra a etapa.
5. Confira os valores abaixo. Se baterem, apague o módulo `MontarPainel` (o PAINEL não depende dele) e salve.

Se o Excel pedir credenciais do SharePoint, use as mesmas da consulta `Tabela_DemandaSemanal`.

## O que conferir depois de instalar

**Primeiro, a conferência automática** (PAINEL, células H19:H21). Com Visão = Semanal e a mesma semana do FILTRO_DATA (hoje, Sem 45), as três diferenças contra a CAP_PROPOSTA atual devem ser **0,0000**.

**Depois, os números do painel:**

| Período no PAINEL | HC por turno necessário | % Ocupação | % Ocupação corrigida | Lotes | Dias úteis | Lotes por dia | Lotes por dia a 100% |
|---|---|---|---|---|---|---|---|
| Semanal · Sem 45/2026 | 59,48 | 116,7% | 117,2% | 414 | 4,58 | 90,4 | 77,1 |
| Mensal · Set/2026 | 48,68 | 95,5% | 96,4% | 1.691 | 25,00 | 67,6 | 70,2 |
| Mensal · Jul a Set/2026 | 45,70 | 89,6% | 90,5% | 1.656,3 | 25,67 | 64,5 | 71,4 |

**FIFO:** Fam Dissolutor na semana 45 = 50,43.

**Visão mensal contra a CAP antiga:** no FILTRO_DATA escolha Mensal, Set/2026 a Set/2026 e Atualizar Tudo; no PAINEL escolha Mensal, Set/2026. A conferência H19:H21 também deve dar 0,0000.

## Se der erro

O macro para e avisa sem apagar nada. Para tentar de novo:

- remova as abas PAINEL, PAINEL_AUX e CAP_PERIODOS criadas;
- as consultas são atualizadas no lugar, não precisa removê-las.

Se o erro for **"Formula.Firewall"**, vá em Dados → Obter Dados → Opções de Consulta → Privacidade e use a mesma configuração que já funciona para a CAP_PROPOSTA.

## Como foi testado

- **Consultas:** a lógica das consultas novas foi reproduzida em Python e conferida contra a CAP gravada no arquivo (semana 45: 770 valores, 0 diferenças).
- **Fórmulas do PAINEL:** recalculadas no LibreOffice nos três períodos da tabela acima (2.367 valores, 0 diferenças contra o Python).
- **Não testado:** o macro e as consultas não puderam ser executados num Excel aqui. Por isso existe a conferência automática H19:H21.
