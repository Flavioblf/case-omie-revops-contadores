/* =====================================================================
   DESAFIO 1c — Tiers de prioridade (Alta / Média / Baixa) para contadores prospect
   Pergunta: classificar os prospects em 3 tiers com base em pelo menos 3 critérios
             de potencial de conversão, apresentando o critério de score.

   Universo: mesma regra do 1a (303 contadores). score_potencial_latente e status_omie NÃO entram.

   Método: matriz Fit x Engajamento, o padrão de lead scoring do mercado
   (Salesforce Account Engagement / HubSpot): o perfil (fit) e o interesse (engajamento)
   são avaliados separadamente e o tier sai do cruzamento dos dois, sem pesos arbitrários.

   Critérios:
     1. Fit da carteira       num_clientes_ativos, cortado nos tercis da própria base:
                              Alto > 24 | Médio 15 a 24 | Baixo < 15
                              (programas de parceria para contadores, como QuickBooks e Xero,
                              classificam o escritório pelo nº de clientes que ele leva à plataforma)
     2. Estágio do funil      status_lead do lead aberto mais avançado:
                              Quente = Reunião Agendada ou Proposta Enviada
                              Morno  = Novo ou Em Contato
                              Frio   = sem lead aberto ou Sem Resposta
     3. Recência              lead aberto há mais de 90 dias (tempo_ciclo_dias, ~2x o ciclo médio
                              de Ganho) perde o engajamento e conta como Frio

   Score = nível do fit (Baixo 1 | Médio 2 | Alto 3) + nível do engajamento (Frio 0 | Morno 1 | Quente 2)
   Tier pela matriz:          Quente     Morno      Frio
            Fit Alto          Alta (5)   Alta (4)   Média (3)
            Fit Médio         Alta (4)   Média (3)  Baixa (2)
            Fit Baixo         Média (3)  Baixa (2)  Baixa (1)
   Score >= 4 = Alta | 3 = Média | <= 2 = Baixa

   Joins: mercado_contadores + leads_pipeline (engajamento) + clientes_ativos (status_cliente = 'Ativo':
          prospects que já têm contrato, divergência a corrigir na origem).
   T-SQL (SQL Server) · banco OMIE · esquema raw
   ===================================================================== */
USE OMIE;
GO

WITH leads AS (
    SELECT
        id_contador,
        -- lead aberto mais avançado dentro de 90 dias: 2 = Quente, 1 = Morno, 0 = Frio
        MAX(CASE WHEN tempo_ciclo_dias <= 90 AND status_lead IN (N'Reunião Agendada', 'Proposta Enviada') THEN 2
                 WHEN tempo_ciclo_dias <= 90 AND status_lead IN ('Novo', 'Em Contato')                    THEN 1
                 ELSE 0 END) AS nivel_engajamento
    FROM raw.leads_pipeline
    GROUP BY id_contador
),
contratos AS (
    SELECT DISTINCT id_contador, status_cliente FROM raw.clientes_ativos WHERE status_cliente = 'Ativo'
),
criterios AS (
    SELECT
        m.id_contador, m.estado, m.regiao, m.porte, m.num_clientes_ativos,
        c.status_cliente,
        CASE WHEN m.num_clientes_ativos > 24  THEN 3
             WHEN m.num_clientes_ativos >= 15 THEN 2
             ELSE 1 END                          AS nivel_fit,
        COALESCE(l.nivel_engajamento, 0)         AS nivel_engajamento
    FROM raw.mercado_contadores m
    LEFT JOIN leads l     ON l.id_contador = m.id_contador
    LEFT JOIN contratos c ON c.id_contador = m.id_contador
    WHERE m.status_omie <> 'Cliente Ativo'
      AND m.converteu_omie = 0
)
SELECT
    id_contador, estado, regiao, porte, num_clientes_ativos, status_cliente,
    CASE nivel_fit WHEN 3 THEN 'Alto' WHEN 2 THEN N'Médio' ELSE 'Baixo' END         AS fit_carteira,
    CASE nivel_engajamento WHEN 2 THEN 'Quente' WHEN 1 THEN 'Morno' ELSE 'Frio' END AS engajamento,
    nivel_fit + nivel_engajamento                                                   AS score_prioridade,
    CASE WHEN nivel_fit + nivel_engajamento >= 4 THEN 'Alta'
         WHEN nivel_fit + nivel_engajamento = 3  THEN N'Média'
         ELSE 'Baixa' END                                                           AS tier_prioridade
FROM criterios
ORDER BY score_prioridade DESC, nivel_engajamento DESC, num_clientes_ativos DESC;
