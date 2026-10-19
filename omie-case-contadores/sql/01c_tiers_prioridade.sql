-- =====================================================================
-- Desafio 1c — Tiers de prioridade (Alta / Média / Baixa) para contadores prospect
-- Mesma lógica de sql/sqlserver/desafio1c_tiers_prioridade.sql (fonte única do tier).
-- Universo: mesma regra do 1a — status_omie <> 'Cliente Ativo' E converteu_omie = 0 (303 contadores).
-- score_potencial_latente e status_omie NÃO entram no score.
--
-- Método: matriz Fit x Engajamento (padrão de lead scoring: Salesforce Account Engagement / HubSpot).
--   1. Fit da carteira   num_clientes_ativos nos tercis da base: Alto > 24 | Médio 15 a 24 | Baixo < 15
--   2. Engajamento       lead aberto mais avançado: Quente = Reunião/Proposta | Morno = Novo/Em Contato | Frio
--   3. Recência          lead aberto há mais de 90 dias conta como Frio
-- Score = nível do fit (1 a 3) + nível do engajamento (0 a 2): >= 4 Alta | 3 Média | <= 2 Baixa
-- =====================================================================

CREATE OR REPLACE VIEW vw_tiers_prospects AS
WITH niveis AS (
    SELECT
        *,
        CASE WHEN num_clientes_ativos > 24  THEN 3
             WHEN num_clientes_ativos >= 15 THEN 2
             ELSE 1 END                                       AS nivel_fit,
        -- estagio_aberto_ordem já ignora leads parados (> 90 dias) e 'Sem Resposta'
        CASE WHEN estagio_aberto_ordem >= 3 THEN 2
             WHEN estagio_aberto_ordem >= 1 THEN 1
             ELSE 0 END                                       AS nivel_engajamento
    FROM stg_contador
    WHERE status_omie_origem <> 'Cliente Ativo'
      AND converteu_omie = 0
)
SELECT
    *,
    CASE nivel_fit WHEN 3 THEN 'Alto' WHEN 2 THEN 'Médio' ELSE 'Baixo' END         AS fit_carteira,
    CASE nivel_engajamento WHEN 2 THEN 'Quente' WHEN 1 THEN 'Morno' ELSE 'Frio' END AS engajamento,
    nivel_fit + nivel_engajamento                                                   AS score_prioridade,
    CASE WHEN nivel_fit + nivel_engajamento >= 4 THEN 'Alta'
         WHEN nivel_fit + nivel_engajamento = 3  THEN 'Média'
         ELSE 'Baixa' END                                                           AS tier_prioridade
FROM niveis;

-- @name: 1c_tiers_prospects
SELECT
    id_contador, estado, regiao, porte, num_clientes_ativos, status_canonico,
    fit_carteira, engajamento, score_prioridade, tier_prioridade
FROM vw_tiers_prospects
ORDER BY score_prioridade DESC, nivel_engajamento DESC, num_clientes_ativos DESC;

-- @name: 1c_tiers_resumo_validacao
-- Matriz Fit x Engajamento com as contagens
SELECT
    fit_carteira,
    SUM(CASE WHEN engajamento = 'Quente' THEN 1 ELSE 0 END) AS quente,
    SUM(CASE WHEN engajamento = 'Morno'  THEN 1 ELSE 0 END) AS morno,
    SUM(CASE WHEN engajamento = 'Frio'   THEN 1 ELSE 0 END) AS frio
FROM vw_tiers_prospects
GROUP BY fit_carteira, nivel_fit
ORDER BY nivel_fit DESC;

-- @name: 1c_tiers_por_regiao
SELECT
    regiao,
    SUM(CASE WHEN tier_prioridade = 'Alta'  THEN 1 ELSE 0 END) AS alta,
    SUM(CASE WHEN tier_prioridade = 'Média' THEN 1 ELSE 0 END) AS media,
    SUM(CASE WHEN tier_prioridade = 'Baixa' THEN 1 ELSE 0 END) AS baixa,
    ROUND(100.0 * SUM(CASE WHEN tier_prioridade = 'Alta' THEN 1 ELSE 0 END) / COUNT(*), 1) AS pct_alta
FROM vw_tiers_prospects
GROUP BY regiao
ORDER BY alta DESC;
