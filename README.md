# Case Omie · Revenue Ops — Canal de Contadores

**[▶ Abrir o painel online (Power BI)](https://app.powerbi.com/view?r=eyJrIjoiYjRiNTRkMjItNjA1ZS00YjA5LThlM2QtYWQ5NDRmZWM1ODliIiwidCI6IjMyOGFmM2M3LTUyYjUtNDk0Yy05YWFhLWQyYTNiYTkxYWVkZiJ9)**
· 7 páginas, uma por pergunta do case (pergunta → resposta → evidência); abre no navegador, sem login.

Case técnico para a vaga de **Analista de Revenue Ops Sênior** (Parcerias / canal de contadores).
Tudo foi calculado a partir das 5 tabelas fornecidas (`base_case_contador_omie.xlsx`), sem premissas externas:
a única exceção são os parâmetros de cenário do impacto em nMRR (Desafio 2, extra), que o próprio enunciado pede
e que estão declarados como premissa.

## Resumo executivo

- **Mercado:** 303 dos 600 escritórios ainda são endereçáveis (penetração de 49,5%); 99% são pequenos ou autônomos,
  e o Sudeste concentra 51%.
- **Conversão:** o Evento tem o melhor ROI de esforço (24,8 tentativas de contato por ganho, amostra pequena), à frente
  do Inbound (25,6).
- **Propensão:** porte, funcionários, carteira, ERP e anos de mercado explicam a conversão do target, mas o fechamento
  real não confirma o perfil (win rate do topo 60% contra 65% das demais faixas). O perfil é hipótese a validar em piloto.
- **Capacity:** gap de R$ 19,5 mil/mês nas 9 franquias abaixo da meta (a rede como um todo bate a meta). O gap acompanha
  os meses de operação (correlação 0,65), não o funil (0,15).
- **EC adicional:** BA, a franquia com mais fechamentos potenciais por EC (12,5 endereçáveis por EC × 4,5% de conversão).
- **Governança:** o mesmo nMRR vale R$ 299 mil, R$ 17 mil ou R$ 3 mil por mês conforme a tabela; o target
  `converteu_omie` vaza do `score_potencial_latente` e diverge dos contratos.

## Entregáveis

| Desafio | Entregável | Onde |
|---|---|---|
| 1 · SQL | Consultas das perguntas a, b e c (T-SQL) | `omie-case-contadores/sql/sqlserver/desafio1*.sql` |
| 1 · apoio | Views do painel e análises de governança | `omie-case-contadores/sql/sqlserver/anexo_analises_complementares.sql` |
| 2 · Propensão | Planilha com fórmulas vivas (aba **Respostas** primeiro) | `omie-case-contadores/output/Desafio2_Propensao.xlsx` |
| 3 · Capacity | Planilha com fórmulas vivas | `omie-case-contadores/output/Desafio3_Capacity.xlsx` |
| 4 · Storytelling | Apresentação de 6 slides (com notas do apresentador) | `omie-case-contadores/output/Apresentacao_Omie_Canal_Contadores.pptx` |
| 1 a 3 · Painel | Power BI (7 páginas, uma por pergunta: pergunta → resposta → evidência) | `Painel Omie.pbip` |
| Documentação do painel | Modelo, medidas DAX e valores de referência | `Painel Omie - Documentação.md` |

## Respostas

### Desafio 1 — SQL

**a) Mercado endereçável por região e porte.** Endereçável = `status_omie <> 'Cliente Ativo'` **e**
`converteu_omie = 0`: 303 contadores. Sudeste 156 (51,5%), Nordeste 53 (17,5%), Sul 40 (13,2%),
Centro-Oeste 37 (12,2%), Norte 17 (5,6%). Médio e Grande já foram convertidos: só 3 seguem endereçáveis.

**b) Conversão e ROI de esforço.** Conversão = Ganhos ÷ total de leads, por canal (15,5% a 26,7%) e por franquia.
ROI de esforço = tentativas de contato de todos os leads ÷ Ganhos: Evento 24,8 (amostra pequena, 20 leads),
Inbound 25,6, Outreach 28,0, Franquia 44,6 e Indicação 49,4.

**c) Tiers de prioridade.** Matriz **Fit × Engajamento** (padrão de lead scoring do mercado), sem pesos arbitrários:

| Critério | Regra |
|---|---|
| Fit da carteira | `num_clientes_ativos` nos tercis dos 303: Alto > 24, Médio 15 a 24, Baixo < 15 |
| Estágio do funil | Quente = Reunião/Proposta, Morno = Novo/Em Contato, Frio = sem lead |
| Recência | lead aberto há mais de 90 dias conta como Frio (nenhum Ganho da base fechou depois de 89 dias) |

Score = nível do fit (1 a 3) + nível do engajamento (0 a 2): ≥ 4 Alta, 3 Média, ≤ 2 Baixa →
**Alta 21, Média 96, Baixa 186**. A consulta usa `mercado_contadores`, `leads_pipeline`, `clientes_ativos`
(marca os 35 prospects que já têm contrato ativo) e `franquias`.

### Desafio 2 — Propensão à conversão (abordagem descritiva)

- **Lógica e variáveis:** taxa de conversão (`converteu_omie`) por categoria de 9 características; entram as 5 de
  maior amplitude (maior taxa − menor taxa, categorias com 30+ contadores). Fora: `score_potencial_latente`
  (regra do case) e `status_omie` (sem sinal: conversão de 36% a 43% em qualquer status).
- **Propensão de cada contador:** média das taxas de conversão das categorias dele nos 5 drivers. Faixa Alta = top 30%.
- **Top 5:** porte (93 p.p.), nº de funcionários (89), nº de clientes da carteira (88), ERP atual (59),
  anos de mercado (34). Os três primeiros medem tamanho; com o ERP, descrevem o perfil do topo
  (ICP: 50+ clientes, 6+ funcionários e ERP de mercado).
- **Como avaliar (duas etapas, só métricas descritivas):**
  1. Base: conversão por faixa Alta · Média · Baixa = 91% · 29% · 3%; o top 30% capta 69% dos convertidos;
     win rate real da faixa Alta 60% contra 65% das demais; ciclo do Ganho 36 contra 48 dias.
  2. Piloto de 90 dias com grupo de controle: win rate, nMRR por EC e ciclo do piloto melhores que os do controle.
- **Impacto (extra):** +25% de nMRR no cenário base (10% a 46%), ≈ R$ 31 mil/mês nas 9 franquias abaixo da meta.
  Premissas declaradas: % da capacidade redirecionada ao topo (30% / 50% / 60%) e fator de realização do lift
  (25% / 40% / 60%). Com fator 0, que é o que o pipeline real sugere hoje, o uplift é 0: o número só se confirma
  com o piloto.

### Desafio 3 — Capacity e foco dos ECs

- **a) Franquias e ECs:** 9 de 15 franquias abaixo da meta; gap de R$ 19.500/mês. Norte, Nordeste e Centro-Oeste
  estão abaixo da meta como região (74%, 82% e 89%); Sudeste e Sul batem a meta (107% e 108%) e o gap vem só do RJ e
  do RS. Por EC equivalente (a base não identifica o EC individual: `id_executivo` se repete na franquia), os ECs de
  PE, PA, BA e RJ entregam 74% a 82% da média da rede (R$ 6.795 por EC); o CE realiza acima da rede, mas tem a maior
  meta por EC. Em 49% dos EC-mês não há nenhum fechamento.
- **b) Framework de priorização (5 ondas):** 1. Fechar 15 contadores com lead em Reunião/Proposta → 2. Prospectar
  13 do tier Alta → 3. Requalificar 54 com lead parado → 4. Nutrir 66 (Média) → 5. Régua automatizada para 155 (Baixa).
  Dentro da onda, maior MRR esperado (ticket médio do porte × (1 − churn)). O Fit do 1c é a carteira, que acompanha a
  propensão do Desafio 2 (correlação 0,8 nos 303). Território = franquia do estado do contador.
- **c) EC adicional: BA.** Critério: fechamentos potenciais por EC = endereçáveis do território por EC × conversão
  lead → fechamento da franquia. BA 0,56 (12,5 endereçáveis por EC contra 6,6 da rede; 4,5% de conversão contra 4,0%),
  MG 0,50 (já acima da meta). MT, MA, RO e MS somam 27 endereçáveis, mas não têm franquia nem histórico de execução.

### Desafio 4 — Apresentação (6 slides)

1. Tamanho da oportunidade · 2. Onde está o dinheiro · 3. Propensão e impacto · 4. Capacity e EC adicional ·
5. Três ações imediatas · 6. KPI de 90 dias: levar as 9 franquias abaixo da meta de 86% para 100% da meta delas.

## Escolhas metodológicas

| Tema | Escolha |
|---|---|
| Endereçável | `mercado_contadores` sem `Cliente Ativo` e com `converteu_omie = 0` (303); `converteu_omie = 1` conta como já convertido |
| Win rate | Ganho ÷ (Ganho + Perdido); a conversão do 1b segue o enunciado (Ganho ÷ total de leads) |
| Lead parado | aberto há mais de 90 dias; nenhum Ganho da base passou de 89 dias |
| Propensão | descritiva, sem modelo estatístico e sem métricas de modelo (AUC, precision, recall) |
| Território | franquia do estado; SP: cidade de São Paulo = SP Capital, demais cidades = SP Interior |
| Nível de EC | EC equivalente (franquia ÷ nº de ECs) e dispersão mensal da execução |
| Receita | o uplift % é o número principal, porque o nMRR aparece em três escalas diferentes |

## Achados de qualidade de dados (governança)

- O target vaza: `converteu_omie = 1` exatamente quando `score_potencial_latente` ≥ ~0,399.
- O target não reflete o real: 52 dos 73 contratos de `clientes_ativos` pertencem a contadores com `converteu_omie = 0`.
- `status_omie` diverge dos contratos: 48 contadores com contrato ativo não têm status de cliente, e 87 com status de
  cliente não têm contrato ativo; 35 dos 303 endereçáveis já têm contrato ativo.
- `id_executivo` não identifica o EC (1 ID por franquia).
- nMRR em três escalas: franquias R$ 299 mil/mês, execução R$ 17 mil/mês, CRM R$ 3 mil/mês.
- Menores: 184 linhas com `usa_erp = 'Não'` e ERP de mercado preenchido; 3 contadores duplicados em
  `clientes_ativos`; 20 leads abertos de contadores que já são clientes ativos; `meses_operando` não bate com
  `data_inicio`.

## Como reproduzir

A base do case (`base_case_contador_omie.xlsx`) não está neste repositório: ela é material do processo seletivo.
Para reproduzir, coloque-a em `omie-case-contadores/data/raw/` e carregue as 5 abas no SQL Server.

Pré-requisitos: SQL Server local (banco `OMIE`, tabelas no esquema `raw`), Python 3 com `pandas`, `openpyxl` e
`duckdb`, Microsoft Excel (para recalcular as fórmulas), Node.js (deck) e Power BI Desktop.

```bash
cd omie-case-contadores
# 1. SQL Server: rodar sql/sqlserver/anexo_analises_complementares.sql (cria as views do painel)
# 2. Planilhas (DuckDB local + recálculo pelo Excel)
python src/load.py
python src/build_excel_propensao.py && python src/recalc_excel.py output/Desafio2_Propensao.xlsx
python src/build_excel_capacity.py  && python src/recalc_excel.py output/Desafio3_Capacity.xlsx
# 3. Deck (npm install em deck/; PPTX_SKILL_DIR aponta para a pasta que contém scripts/apply_theme.js)
python src/deck_data.py
NODE_PATH=deck/node_modules PPTX_SKILL_DIR=<pasta> node deck/build_deck.js
# 4. Power BI (com o Power BI Desktop fechado); depois abrir Painel Omie.pbip e clicar em Atualizar
python powerbi/build_pbip.py
```

## Estrutura

```
├── Painel Omie.pbip                     # painel Power BI (modelo TMDL + relatório PBIR)
├── Painel Omie - Documentação.md
└── omie-case-contadores/
    ├── sql/sqlserver/                   # respostas do Desafio 1 e anexo (T-SQL)
    ├── sql/00_staging.sql, 01c_*.sql    # mesma lógica em DuckDB, usada para gerar as planilhas
    ├── src/                             # geradores das planilhas e dos dados do deck
    ├── deck/                            # gerador do deck e design system Omie
    ├── powerbi/build_pbip.py            # gerador do painel
    └── output/                          # planilhas e apresentação
```
