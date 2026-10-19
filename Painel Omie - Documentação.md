# Painel Omie — Desafios 1, 2 e 3 no Power BI

Projeto `Painel Omie.pbip` (modelo em TMDL, relatório em PBIR), gerado por `omie-case-contadores/powerbi/build_pbip.py`. Abordagem **descritiva** (sem modelo estatístico), com os mesmos números dos Excel `output/Desafio2_Propensao.xlsx` e `output/Desafio3_Capacity.xlsx`, e visual no design system Omie do deck.

**Versão online (sem instalar nada):** [Painel Omie no Power BI](https://app.powerbi.com/view?r=eyJrIjoiYjRiNTRkMjItNjA1ZS00YjA5LThlM2QtYWQ5NDRmZWM1ODliIiwidCI6IjMyOGFmM2M3LTUyYjUtNDk0Yy05YWFhLWQyYTNiYTkxYWVkZiJ9).

## 1. Como abrir e atualizar

1. Pré-requisito: no SQL Server, rode uma vez `sql/sqlserver/anexo_analises_complementares.sql` (cria as views `dbo.vw_tiers_prospects`, `dbo.vw_desafio1a_mercado` e `dbo.vw_desafio1b_conversao`, com o mesmo texto das consultas do Desafio 1, e `dbo.vw_propensao_d2`, com o score do Desafio 2).
2. Abra `Painel Omie.pbip` no Power BI Desktop e clique em **Atualizar** (conexão **Windows** em `localhost`, banco `OMIE`).
3. Outro servidor/banco: **Transformar dados → Editar parâmetros** (`ServidorSQL`, `BancoSQL`).

## 2. Tabelas

| Tabela | Origem | Grão | Papel |
|---|---|---|---|
| `mercado_contadores` | `raw.mercado_contadores` | 1 linha por contador (600) | Principal: `converteu_omie` e características |
| `leads_pipeline` | `raw.leads_pipeline` | 1 linha por lead (350) | Resultado comercial real (Ganho/Perdido) |
| `franquias` | `raw.franquias` | 1 linha por franquia (15) | Meta e realizado de nMRR (impacto) |
| `tiers_1c` (oculta) | `dbo.vw_tiers_prospects` | 1 linha por endereçável (303) | Fit, engajamento, score e tier do Desafio 1c |
| `d1a_mercado` | `dbo.vw_desafio1a_mercado` | região × porte + totais | Resposta do Desafio 1a |
| `d1b_conversao` | `dbo.vw_desafio1b_conversao` | canal e franquia | Resposta do Desafio 1b |
| `propensao_d2` (oculta) | `dbo.vw_propensao_d2` | 1 linha por contador (600) | Score e faixa de propensão do Desafio 2 |
| `Param_Variaveis` | Tabela fixa | 9 características | Lista para o ranking dos drivers |
| `Param_Avaliacao` | Tabela fixa | 6 métricas | Plano de avaliação em duas etapas (Desafio 2) |
| `Param_Cenarios` | Tabela fixa | 3 cenários | Premissas do impacto |
| `clientes_ativos` | `raw.clientes_ativos` | 1 linha por contrato (73) | Ticket por porte e churn (Desafio 3) |
| `Param_Ondas` | Tabela fixa | 5 ondas | Quem entra, o que o EC faz e por quê (Desafio 3b) |
| `Param_Territorio` | Tabela fixa | 1 linha por estado com franquia | Franquia do estado (MT, MA, RO e MS não têm franquia) |
| `execucao_comercial` (oculta) | `raw.execucao_comercial` | franquia × mês × EC | Conversão do funil e dispersão por EC (Desafio 3) |
| `_Medidas` | — | — | Tabela de medidas |

## 3. Relacionamentos

| De (muitos) | Para (um) | Direção |
|---|---|---|
| `leads_pipeline[id_contador]` | `mercado_contadores[id_contador]` | Única (mercado filtra leads) |
| `leads_pipeline[id_franquia]` | `franquias[id_franquia]` | Única |
| `clientes_ativos[id_contador]` | `mercado_contadores[id_contador]` | Única |
| `execucao_comercial[id_franquia]` | `franquias[id_franquia]` | Única |

`tiers_1c`, `propensao_d2` e `Param_*` não se relacionam: o tier e o score entram em `mercado_contadores` por `LOOKUPVALUE` e as tabelas de parâmetros são lidas pelas medidas.

## 4. Colunas calculadas

### mercado_contadores

**`sem_erp`** — 1 = usa planilha, nenhum sistema ou ContaAzul (concorrente direto)

```dax
IF(mercado_contadores[erp_atual] IN {"Nenhum", "Planilha", "ContaAzul"}, 1, 0)
```

**`uso_erp`** — ERP de mercado (Domínio, Questor, Alterdata, Totvs, SAP) × sem ERP (planilha, nenhum, ContaAzul)

```dax
IF(mercado_contadores[sem_erp] = 1, "Sem ERP de mercado", "ERP de mercado")
```

**`icp`** — Perfil ideal: 50+ clientes, 6+ funcionários e ERP de mercado

```dax
IF(mercado_contadores[num_clientes_ativos] >= 50 && mercado_contadores[num_funcionarios] >= 6 && mercado_contadores[sem_erp] = 0, "Sim", "Não")
```

**`enderecavel`** — Regra do Desafio 1a: não é Cliente Ativo e ainda não converteu (303 contadores)

```dax
IF(mercado_contadores[status_omie] <> "Cliente Ativo" && mercado_contadores[converteu_omie] = 0, "Sim", "Não")
```

**`faixa_carteira`** — Faixa do nº de clientes

```dax
SWITCH(TRUE(), mercado_contadores[num_clientes_ativos] < 20, "1. < 20", mercado_contadores[num_clientes_ativos] < 50, "2. 20 a 49",
    mercado_contadores[num_clientes_ativos] < 150, "3. 50 a 149", "4. 150+")
```

**`faixa_funcionarios`** — Faixa do nº de funcionários

```dax
SWITCH(TRUE(), mercado_contadores[num_funcionarios] <= 2, "1. 1 a 2", mercado_contadores[num_funcionarios] <= 5, "2. 3 a 5",
    mercado_contadores[num_funcionarios] <= 15, "3. 6 a 15", "4. 16+")
```

**`faixa_anos`** — Faixa de anos de mercado

```dax
SWITCH(TRUE(), mercado_contadores[anos_mercado] < 8, "1. < 8 anos", mercado_contadores[anos_mercado] < 16, "2. 8 a 15 anos", "3. 16+ anos")
```

**`score_prioridade`** — Score do Desafio 1c: nível do fit (1 a 3) + nível do engajamento (0 a 2); vazio para quem não é endereçável

```dax
LOOKUPVALUE(tiers_1c[score_prioridade], tiers_1c[id_contador], mercado_contadores[id_contador])
```

**`fit_ordem`** — Ordem do fit (3 = Alto)

```dax
LOOKUPVALUE(tiers_1c[nivel_fit], tiers_1c[id_contador], mercado_contadores[id_contador])
```

**`fit_carteira`** — Fit do 1c pela carteira: Alto > 24 clientes, Médio 15 a 24, Baixo < 15

```dax
LOOKUPVALUE(tiers_1c[fit_carteira], tiers_1c[id_contador], mercado_contadores[id_contador])
```

**`engajamento`** — Engajamento do 1c: Quente = Reunião/Proposta, Morno = Novo/Em Contato (lead aberto até 90 dias), Frio

```dax
LOOKUPVALUE(tiers_1c[engajamento], tiers_1c[id_contador], mercado_contadores[id_contador])
```

**`tier_ordem`** — Ordem do tier

```dax
VAR t = LOOKUPVALUE(tiers_1c[tier_prioridade], tiers_1c[id_contador], mercado_contadores[id_contador])
RETURN SWITCH(t, "Alta", 1, "Média", 2, "Baixa", 3)
```

**`tier_prioridade`** — Tier do Desafio 1c pela matriz Fit x Engajamento: score >= 4 Alta, 3 Média, <= 2 Baixa

```dax
SWITCH(mercado_contadores[tier_ordem], 1, "Alta", 2, "Média", 3, "Baixa")
```

**`propensao`** — Score de propensão do Desafio 2: média das taxas de conversão do perfil do contador nos 5 drivers do top 5

```dax
LOOKUPVALUE(propensao_d2[propensao], propensao_d2[id_contador], mercado_contadores[id_contador])
```

**`faixa_prop_ordem`** — Ordem da faixa de propensão (1 = Alta)

```dax
LOOKUPVALUE(propensao_d2[faixa_ordem], propensao_d2[id_contador], mercado_contadores[id_contador])
```

**`faixa_propensao`** — Faixa de propensão: Alta = top 30% do score, Média = do percentil 30 ao 70, Baixa = demais

```dax
LOOKUPVALUE(propensao_d2[faixa_propensao], propensao_d2[id_contador], mercado_contadores[id_contador])
```

**`situacao_pipeline`** — Estágio do lead ativo (até 90 dias) do contador

```dax
COALESCE(LOOKUPVALUE(tiers_1c[estagio_aberto], tiers_1c[id_contador], mercado_contadores[id_contador]), "—")
```

**`franquia_responsavel`** — Franquia do estado do contador (SP: cidade de São Paulo = SP Capital, demais = SP Interior); vazio = sem franquia

```dax
IF(mercado_contadores[estado] = "SP",
    IF(mercado_contadores[cidade] = "São Paulo", "FRQ001", "FRQ002"),
    LOOKUPVALUE(Param_Territorio[id_franquia], Param_Territorio[estado], mercado_contadores[estado]))
```

**`franquia_nome`** — Nome da franquia responsável pelo contador

```dax
COALESCE(SUBSTITUTE(LOOKUPVALUE(franquias[nome_franquia], franquias[id_franquia], mercado_contadores[franquia_responsavel]), "Franquia ", ""), "Sem franquia")
```

**`cobertura`** — Própria (estado com franquia) ou Sem franquia

```dax
IF(ISBLANK(mercado_contadores[franquia_responsavel]), "Sem franquia", "Própria")
```

**`ticket_porte`** — Ticket MRR médio do porte (mrr_inicial em clientes_ativos)

```dax
VAR p = mercado_contadores[porte]
RETURN ROUND(AVERAGEX(FILTER(ALL(clientes_ativos), clientes_ativos[porte] = p), clientes_ativos[mrr_inicial]), 0)
```

**`mrr_esperado`** — Ticket do porte × (1 − churn): MRR esperado de um contador endereçável

```dax
IF(mercado_contadores[enderecavel] = "Sim", mercado_contadores[ticket_porte] * (1 - DIVIDE(COUNTROWS(FILTER(ALL(clientes_ativos), clientes_ativos[status_cliente] = "Churned")), COUNTROWS(ALL(clientes_ativos)))))
```

**`estagio_ordem`** — Estágio do lead ativo em número (4 = proposta)

```dax
SWITCH(mercado_contadores[situacao_pipeline], "Proposta Enviada", 4, "Reunião Agendada", 3, "Em Contato", 2, "Novo", 1, 0)
```

**`lead_parado`** — 1 = tem lead aberto há mais de 90 dias

```dax
IF(COUNTROWS(FILTER(RELATEDTABLE(leads_pipeline),
    NOT leads_pipeline[status_lead] IN {"Ganho", "Perdido"} && leads_pipeline[tempo_ciclo_dias] > 90)) > 0, 1, 0)
```

**`onda`** — Onda de priorização (1 a 5)

```dax
IF(mercado_contadores[enderecavel] = "Sim",
    SWITCH(TRUE(), mercado_contadores[estagio_ordem] >= 3, 1, mercado_contadores[tier_prioridade] = "Alta", 2, mercado_contadores[lead_parado] = 1, 3,
        mercado_contadores[tier_prioridade] = "Média", 4, 5))
```

**`onda_nome`** — Onda de priorização do Desafio 3b

```dax
SWITCH(mercado_contadores[onda], 1, "1. Fechar (Reunião/Proposta)", 2, "2. Prospectar tier Alta",
    3, "3. Requalificar lead parado", 4, "4. Nutrir (marketing)", 5, "5. Régua automatizada")
```

### leads_pipeline

**`ganho`** — 1 = lead Ganho

```dax
IF(leads_pipeline[status_lead] = "Ganho", 1, 0)
```

**`com_desfecho`** — 1 = lead encerrado (Ganho ou Perdido)

```dax
IF(leads_pipeline[status_lead] IN {"Ganho", "Perdido"}, 1, 0)
```

### franquias

**`atingimento`** — Realizado ÷ meta de nMRR

```dax
DIVIDE(franquias[realizado_nmrr_mes], franquias[meta_nmrr_mes])
```

**`gap`** — Meta − realizado (R$/mês), zero para quem bateu a meta

```dax
MAX(0, franquias[meta_nmrr_mes] - franquias[realizado_nmrr_mes])
```

**`realizado_por_ec`** — Realizado de nMRR por EC

```dax
DIVIDE(franquias[realizado_nmrr_mes], franquias[num_executivos])
```

**`situacao`** — Na meta ou abaixo da meta

```dax
IF(franquias[atingimento] >= 1, "Na meta", "Abaixo da meta")
```

**`ordem_atingimento`** — Ordem do menor para o maior atingimento (desempate pelo id)

```dax
RANKX(ALLNOBLANKROW(franquias), franquias[atingimento] + VALUE(RIGHT(franquias[id_franquia], 3)) / 1000000, , ASC, DENSE)
```

**`rotulo`** — Franquia, meses de operação e atingimento, ordenada pelo atingimento

```dax
SUBSTITUTE(franquias[nome_franquia], "Franquia ", "") & " · " & franquias[meses_operando] & "m · " & FORMAT(franquias[atingimento], "0%")
```

### clientes_ativos

**`porte`** — Porte do contador (para o ticket por porte)

```dax
RELATED(mercado_contadores[porte])
```

## 5. Medidas

### 1. Base

**[Contadores]** — Nº de contadores no contexto · formato `#,0`

```dax
COUNTROWS(mercado_contadores)
```

**[Convertidos]** — Contadores com converteu_omie = 1 · formato `#,0`

```dax
SUM(mercado_contadores[converteu_omie])
```

**[Taxa de Conversão]** — Convertidos ÷ contadores · formato `0.0%`

```dax
DIVIDE([Convertidos], [Contadores])
```

**[Taxa Média]** — Taxa dos 600 contadores · formato `0.0%`

```dax
CALCULATE([Taxa de Conversão], REMOVEFILTERS(mercado_contadores))
```

**[Lift vs Média]** — Taxa do grupo ÷ taxa média · formato `0.0"x"`

```dax
DIVIDE([Taxa de Conversão], [Taxa Média])
```

### 2. Drivers

**[Amplitude (p.p.)]** — Maior − menor taxa de conversão entre as categorias (30+ contadores): quanto a característica separa quem converte · formato `0`

```dax
SWITCH(SELECTEDVALUE(Param_Variaveis[variavel]),
    "Porte",
        VAR t = FILTER(ADDCOLUMNS(ALL(mercado_contadores[porte]), "@n", CALCULATE(COUNTROWS(mercado_contadores)), "@tx", CALCULATE([Taxa de Conversão])), [@n] >= 30)
        RETURN (MAXX(t, [@tx]) - MINX(t, [@tx])) * 100,
    "Nº de funcionários",
        VAR t = FILTER(ADDCOLUMNS(ALL(mercado_contadores[faixa_funcionarios]), "@n", CALCULATE(COUNTROWS(mercado_contadores)), "@tx", CALCULATE([Taxa de Conversão])), [@n] >= 30)
        RETURN (MAXX(t, [@tx]) - MINX(t, [@tx])) * 100,
    "Nº de clientes da carteira",
        VAR t = FILTER(ADDCOLUMNS(ALL(mercado_contadores[faixa_carteira]), "@n", CALCULATE(COUNTROWS(mercado_contadores)), "@tx", CALCULATE([Taxa de Conversão])), [@n] >= 30)
        RETURN (MAXX(t, [@tx]) - MINX(t, [@tx])) * 100,
    "ERP atual",
        VAR t = FILTER(ADDCOLUMNS(ALL(mercado_contadores[erp_atual]), "@n", CALCULATE(COUNTROWS(mercado_contadores)), "@tx", CALCULATE([Taxa de Conversão])), [@n] >= 30)
        RETURN (MAXX(t, [@tx]) - MINX(t, [@tx])) * 100,
    "Anos de mercado",
        VAR t = FILTER(ADDCOLUMNS(ALL(mercado_contadores[faixa_anos]), "@n", CALCULATE(COUNTROWS(mercado_contadores)), "@tx", CALCULATE([Taxa de Conversão])), [@n] >= 30)
        RETURN (MAXX(t, [@tx]) - MINX(t, [@tx])) * 100,
    "Região",
        VAR t = FILTER(ADDCOLUMNS(ALL(mercado_contadores[regiao]), "@n", CALCULATE(COUNTROWS(mercado_contadores)), "@tx", CALCULATE([Taxa de Conversão])), [@n] >= 30)
        RETURN (MAXX(t, [@tx]) - MINX(t, [@tx])) * 100,
    "Segmento",
        VAR t = FILTER(ADDCOLUMNS(ALL(mercado_contadores[segmento_principal]), "@n", CALCULATE(COUNTROWS(mercado_contadores)), "@tx", CALCULATE([Taxa de Conversão])), [@n] >= 30)
        RETURN (MAXX(t, [@tx]) - MINX(t, [@tx])) * 100,
    "Tem site",
        VAR t = FILTER(ADDCOLUMNS(ALL(mercado_contadores[tem_site]), "@n", CALCULATE(COUNTROWS(mercado_contadores)), "@tx", CALCULATE([Taxa de Conversão])), [@n] >= 30)
        RETURN (MAXX(t, [@tx]) - MINX(t, [@tx])) * 100,
    "LinkedIn",
        VAR t = FILTER(ADDCOLUMNS(ALL(mercado_contadores[presenca_linkedin]), "@n", CALCULATE(COUNTROWS(mercado_contadores)), "@tx", CALCULATE([Taxa de Conversão])), [@n] >= 30)
        RETURN (MAXX(t, [@tx]) - MINX(t, [@tx])) * 100)
```

**[Ranking do Driver]** — 1 = driver mais forte · formato `0`

```dax
RANKX(ALL(Param_Variaveis), [Amplitude (p.p.)])
```

**[Taxa de Conversão ICP]** — Taxa do ICP · formato `0.0%`

```dax
CALCULATE([Taxa de Conversão], mercado_contadores[icp] = "Sim")
```

**[Lift do ICP]** — Taxa do ICP ÷ taxa média · formato `0.0"x"`

```dax
DIVIDE([Taxa de Conversão ICP], [Taxa Média])
```

**[Conversão Top 30%]** — Taxa de conversão (target) da faixa Alta de propensão · formato `0%`

```dax
CALCULATE([Taxa de Conversão], REMOVEFILTERS(mercado_contadores), mercado_contadores[faixa_propensao] = "Alta")
```

**[Lift Top 30%]** — Conversão do top 30% ÷ média · formato `0.0"x"`

```dax
DIVIDE([Conversão Top 30%], [Taxa Média])
```

**[Captura Top 30%]** — % dos convertidos que estão no top 30% do score · formato `0%`

```dax
DIVIDE(CALCULATE([Convertidos], REMOVEFILTERS(mercado_contadores), mercado_contadores[faixa_propensao] = "Alta"), CALCULATE([Convertidos], REMOVEFILTERS(mercado_contadores)))
```

### 3. Resultado Real

**[Leads]** — Leads no pipeline · formato `#,0`

```dax
COUNTROWS(leads_pipeline)
```

**[Leads com Desfecho]** — Leads Ganho + Perdido · formato `#,0`

```dax
SUM(leads_pipeline[com_desfecho])
```

**[Ganhos (Real)]** — Leads Ganho · formato `#,0`

```dax
SUM(leads_pipeline[ganho])
```

**[Win Rate Real]** — Ganho ÷ (Ganho + Perdido): o fechamento comercial de verdade · formato `0.0%`

```dax
DIVIDE([Ganhos (Real)], [Leads com Desfecho])
```

**[Win Rate Real ICP]** — Win rate real do ICP · formato `0.0%`

```dax
CALCULATE([Win Rate Real], mercado_contadores[icp] = "Sim")
```

**[Win Rate Real Fora do ICP]** — Win rate real de quem não é ICP · formato `0.0%`

```dax
CALCULATE([Win Rate Real], mercado_contadores[icp] = "Não")
```

**[Win Rate Real Faixa Alta]** — Win rate real do top 30% do score · formato `0%`

```dax
CALCULATE([Win Rate Real], REMOVEFILTERS(mercado_contadores), mercado_contadores[faixa_propensao] = "Alta")
```

**[Win Rate Real Demais Faixas]** — Win rate real das faixas Média e Baixa · formato `0%`

```dax
CALCULATE([Win Rate Real], REMOVEFILTERS(mercado_contadores), mercado_contadores[faixa_propensao] <> "Alta")
```

### 4. Priorização

**[Endereçáveis]** — Contadores endereçáveis (regra do Desafio 1a) · formato `#,0`

```dax
CALCULATE([Contadores], mercado_contadores[enderecavel] = "Sim")
```

**[Tier Alta]** — Endereçáveis no tier Alta do Desafio 1c · formato `#,0`

```dax
CALCULATE([Endereçáveis], mercado_contadores[tier_prioridade] = "Alta")
```

**[ICP Endereçáveis]** — Endereçáveis com perfil ICP · formato `#,0`

```dax
CALCULATE([Endereçáveis], mercado_contadores[icp] = "Sim")
```

**[Tier Média]** — Endereçáveis no tier Média do Desafio 1c · formato `#,0`

```dax
CALCULATE([Endereçáveis], mercado_contadores[tier_prioridade] = "Média")
```

**[Tier Baixa]** — Endereçáveis no tier Baixa do Desafio 1c · formato `#,0`

```dax
CALCULATE([Endereçáveis], mercado_contadores[tier_prioridade] = "Baixa")
```

**[Quente]** — Endereçáveis com lead em Reunião/Proposta nos últimos 90 dias · formato `#,0`

```dax
CALCULATE([Endereçáveis], mercado_contadores[engajamento] = "Quente")
```

**[Morno]** — Endereçáveis com lead Novo/Em Contato nos últimos 90 dias · formato `#,0`

```dax
CALCULATE([Endereçáveis], mercado_contadores[engajamento] = "Morno")
```

**[Frio]** — Endereçáveis sem lead ativo (inclui parados há mais de 90 dias) · formato `#,0`

```dax
CALCULATE([Endereçáveis], mercado_contadores[engajamento] = "Frio")
```

**[MRR Esperado]** — Soma do MRR esperado dos endereçáveis · formato `#,0`

```dax
SUM(mercado_contadores[mrr_esperado])
```

**[Contadores Ondas 1 e 2]** — Endereçáveis nas ondas 1 (fechar) e 2 (prospectar tier Alta) · formato `#,0`

```dax
CALCULATE([Endereçáveis], mercado_contadores[onda] <= 2)
```

**[Melhor ROI de Esforço]** — Canal com menos tentativas de contato por fechamento (Desafio 1b) · texto

```dax
MAXX(FILTER(d1b_conversao, d1b_conversao[ranking_roi_esforco] = 1), d1b_conversao[canal_origem])
```

**[Carteira Endereçável]** — Clientes finais nas carteiras dos endereçáveis · formato `#,0`

```dax
CALCULATE(SUM(mercado_contadores[num_clientes_ativos]), mercado_contadores[enderecavel] = "Sim")
```

### 5. Impacto

**[Realizado Franquias Abaixo da Meta]** — nMRR oficial mensal (R$) das franquias abaixo da meta · formato `#,0`

```dax
SUMX(FILTER(ALLNOBLANKROW(franquias), franquias[realizado_nmrr_mes] < franquias[meta_nmrr_mes]), franquias[realizado_nmrr_mes])
```

**[% Capacidade Redirecionada]** — Parte da capacidade dos ECs levada ao topo da propensão · formato `0%`

```dax
SELECTEDVALUE(Param_Cenarios[pct_redirecionado])
```

**[Fator de Realização]** — Fração do lift que vira Ganho real (o target diverge do real) · formato `0%`

```dax
SELECTEDVALUE(Param_Cenarios[fator_realizacao])
```

**[Uplift nMRR]** — % capacidade × (lift do top 30% do score − 1) × fator de realização · formato `+0.0%`

```dax
[% Capacidade Redirecionada] * ([Lift Top 30%] - 1) * [Fator de Realização]
```

**[nMRR Incremental (Franquias Abaixo da Meta)]** — Uplift aplicado ao nMRR oficial dessas franquias (R$/mês) · formato `#,0`

```dax
[Uplift nMRR] * [Realizado Franquias Abaixo da Meta]
```

### 6. Desafio 3

**[Franquias Abaixo da Meta]** — Nº de franquias abaixo da meta · formato `0`

```dax
COUNTROWS(FILTER(ALLNOBLANKROW(franquias), franquias[realizado_nmrr_mes] < franquias[meta_nmrr_mes]))
```

**[Correlação Maturidade × Atingimento]** — Correlação de Pearson entre meses de operação e atingimento (≈ 0,65) · formato `0.00`

```dax
VAR t = ALLNOBLANKROW(franquias)
VAR mx = AVERAGEX(t, franquias[meses_operando])
VAR my = AVERAGEX(t, franquias[atingimento])
VAR cov = SUMX(t, (franquias[meses_operando] - mx) * (franquias[atingimento] - my))
VAR sx = SQRT(SUMX(t, (franquias[meses_operando] - mx) ^ 2))
VAR sy = SQRT(SUMX(t, (franquias[atingimento] - my) ^ 2))
RETURN DIVIDE(cov, sx * sy)
```

**[Gap (R$)]** — Gap de meta no contexto (R$/mês) · formato `#,0`

```dax
SUM(franquias[gap])
```

**[Atingimento (abaixo da meta)]** — Atingimento quando abaixo de 100% · formato `0%`

```dax
VAR a = DIVIDE(SUM(franquias[realizado_nmrr_mes]), SUM(franquias[meta_nmrr_mes])) RETURN IF(a < 1, a)
```

**[Atingimento (na meta)]** — Atingimento quando 100% ou mais · formato `0%`

```dax
VAR a = DIVIDE(SUM(franquias[realizado_nmrr_mes]), SUM(franquias[meta_nmrr_mes])) RETURN IF(a >= 1, a)
```

**[ECs]** — Nº de ECs · formato `0`

```dax
SUM(franquias[num_executivos])
```

**[Prospects no Território]** — Endereçáveis do território da franquia · formato `0`

```dax
COUNTROWS(FILTER(ALL(mercado_contadores), mercado_contadores[franquia_responsavel] = SELECTEDVALUE(franquias[id_franquia]) && mercado_contadores[enderecavel] = "Sim"))
```

**[Tier Alta no Território]** — Endereçáveis do tier Alta no território · formato `0`

```dax
COUNTROWS(FILTER(ALL(mercado_contadores), mercado_contadores[franquia_responsavel] = SELECTEDVALUE(franquias[id_franquia]) && mercado_contadores[tier_prioridade] = "Alta"))
```

**[Prospects por EC]** — Endereçáveis do território ÷ ECs · formato `0.0`

```dax
DIVIDE([Prospects no Território], [ECs])
```

**[Realizado por EC]** — nMRR oficial por EC · formato `#,0`

```dax
DIVIDE(SUM(franquias[realizado_nmrr_mes]), [ECs])
```

**[Meta por EC]** — Meta de nMRR por EC · formato `#,0`

```dax
DIVIDE(SUM(franquias[meta_nmrr_mes]), [ECs])
```

**[Meta (R$)]** — Meta de nMRR (R$/mês) · formato `#,0`

```dax
SUM(franquias[meta_nmrr_mes])
```

**[Realizado até a Meta]** — Parte do realizado que preenche a meta (barra de progresso) · formato `#,0`

```dax
IF(NOT ISBLANK([Meta (R$)]), MIN([Realizado (R$)], [Meta (R$)]))
```

**[Falta para a Meta]** — O que falta para bater a meta · formato `#,0`

```dax
IF(NOT ISBLANK([Meta (R$)]), MAX([Meta (R$)] - [Realizado (R$)], 0))
```

**[Acima da Meta]** — O que passou da meta · formato `#,0`

```dax
IF(NOT ISBLANK([Meta (R$)]), MAX([Realizado (R$)] - [Meta (R$)], 0))
```

**[Atingimento]** — Realizado ÷ meta no contexto (franquia ou região) · formato `0%`

```dax
DIVIDE([Realizado (R$)], [Meta (R$)])
```

**[Realizado por EC vs Rede]** — Realizado por EC ÷ realizado por EC da rede · formato `0%`

```dax
DIVIDE([Realizado por EC], DIVIDE(SUMX(ALLNOBLANKROW(franquias), franquias[realizado_nmrr_mes]), SUMX(ALLNOBLANKROW(franquias), franquias[num_executivos])))
```

**[Fechamentos por EC-mês]** — Média de fechamentos por EC por mês (execucao_comercial) · formato `0.00`

```dax
AVERAGE(execucao_comercial[fechamentos])
```

**[% EC-mês sem Fechamento]** — % dos EC-mês sem nenhum fechamento (dispersão entre ECs) · formato `0%`

```dax
DIVIDE(COUNTROWS(FILTER(execucao_comercial, execucao_comercial[fechamentos] = 0)), COUNTROWS(execucao_comercial))
```

**[Realizado (R$)]** — Realizado de nMRR (R$/mês) · formato `#,0`

```dax
SUM(franquias[realizado_nmrr_mes])
```

**[Conversão do Funil]** — Fechamentos ÷ leads trabalhados (execucao_comercial, 2024) · formato `0.0%`

```dax
DIVIDE(SUM(execucao_comercial[fechamentos]), SUM(execucao_comercial[leads_trabalhados]))
```

**[Fechamentos Potenciais por EC]** — Endereçáveis por EC × conversão do funil da franquia · formato `0.00`

```dax
[Prospects por EC] * [Conversão do Funil]
```

**[Ranking EC Adicional]** — 1 = franquia com mais fechamentos potenciais por EC (onde alocar o EC adicional) · formato `0`

```dax
IF(HASONEVALUE(franquias[id_franquia]), RANKX(ALLNOBLANKROW(franquias), [Fechamentos Potenciais por EC]))
```

**[Franquia Recomendada]** — Franquia com mais fechamentos potenciais por EC · texto

```dax
SUBSTITUTE(MAXX(TOPN(1, ALLNOBLANKROW(franquias), [Fechamentos Potenciais por EC]), franquias[nome_franquia]), "Franquia ", "")
```

## 6. Páginas (uma por pergunta do case: pergunta → resposta → evidência)

- **D1 · Mercado, conversão e tiers** — 1a (endereçáveis por região e porte), 1b (conversão e ROI de esforço por canal e franquia) e 1c (matriz Fit × Engajamento e tiers), lidos das views do SQL.
- **D2 · Lógica, top 5 e propensão** — variáveis e porquê, ranking dos drivers e os 303 endereçáveis ordenados pelo score de propensão.
- **D2 · Como avaliar** — plano em duas etapas (base e piloto): tabela de métricas com hoje, meta e etapa, e o target × win rate real por faixa de propensão.
- **D2 · Impacto em nMRR** — uplift por cenário e premissas.
- **D3a · Gap e risco** — meta × realizado por franquia (barra de progresso), gap das franquias abaixo da meta e atingimento por região, e a análise por EC equivalente.
- **D3b · Priorização** — as 5 ondas, ondas 1 e 2 por franquia e a lista de quem atacar primeiro.
- **D3c · EC adicional** — franquia recomendada e ranking por fechamentos potenciais por EC (endereçáveis por EC × conversão do funil).

## 7. Design system

Cores, fontes e componentes de `deck/design_system/tokens.json`: tema `Omie.json` (paleta deep/cyan/teal, Century Gothic nos títulos), logo Omie no topo de cada página, cabeçalho com kicker em caixa alta e título em duas linhas com a frase-chave em teal, cartões gelo (número teal) e o cartão principal escuro (deep + cyan), bordas arredondadas sem sombra.

## 8. Valores de referência (devem bater com o Excel)

| Medida | Esperado |
|---|---|
| Taxa Média | 40,0% |
| Taxa de Conversão ICP / Lift do ICP | 83,8% / 2,1x |
| Win Rate Real ICP / Fora do ICP | 58,6% / 65,9% |
| Conversão por faixa (Alta · Média · Baixa) / lift e captura do top 30% | 91% · 29% · 3% / 2,3x e 69,2% |
| Ciclo médio do Ganho: faixa Alta × demais | 36 × 48 dias |
| Win rate real: faixa Alta / demais | 60,0% / 65,5% |
| Amplitude: Porte / Funcionários / Carteira / ERP / Anos | 93 / 89 / 88 / 59 / 34 |
| Endereçáveis / Tier Alta / Média / Baixa / ICP Endereçáveis | 303 / 21 / 96 / 186 / 24 |
| Matriz 1c (Alto, Médio, Baixo × Quente, Morno, Frio) | 3-13-80 / 5-9-103 / 7-10-73 |
| Melhor ROI de esforço (1b) | Evento (24,8 tentativas por ganho) |
| Uplift nMRR (Conservador / Base / Otimista) | +9,5% / +25,4% / +45,6% |
| Gap de meta / franquias abaixo da meta | R$ 19.500 / 9 |
| Correlação maturidade × atingimento | 0,65 |
| Gap das franquias abaixo da meta por região: Nordeste / Norte / Centro-Oeste / Sudeste / Sul | 6.500 / 4.000 / 4.000 / 4.000 / 1.000 |
| Atingimento da região: Norte / Nordeste / Centro-Oeste / Sudeste / Sul | 74% / 82% / 89% / 107% / 108% |
| Ondas 1 a 5 | 15 / 13 / 54 / 66 / 155 |
| EC adicional: 1º / 2º (fechamentos potenciais por EC) | BA 0,56 / MG 0,50 |
| Endereçáveis sem franquia (MT, MA, RO, MS) | 27 |
