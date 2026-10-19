-- Desafio 1c · Tiers de prioridade dos prospects (matriz Fit x Engajamento)
-- Universo: o mesmo do 1a (303); score_potencial_latente e status_omie não entram
-- Fit: num_clientes_ativos em tercis (> 24 Alto = 3 | 15-24 Médio = 2 | < 15 Baixo = 1)
-- Engajamento: lead aberto até 90 dias em Reunião/Proposta = 2, Novo/Em Contato = 1, demais = 0
-- Tier: fit + engajamento -> >= 4 Alta | 3 Média | <= 2 Baixa
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
