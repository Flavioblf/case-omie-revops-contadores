-- Desafio 1b · Conversão lead -> Ganho por canal e por franquia; ROI de esforço
-- Conversão: Ganhos / total de leads | ROI de esforço: tentativas de todos os leads / Ganhos (menor = melhor)
-- nome_franquia vazio = total do canal; só o canal 'Franquia' tem id_franquia
USE OMIE;
GO

WITH por_canal AS (
    SELECT
        1                                                         AS ordem_bloco,
        canal_origem,
        CAST(NULL AS NVARCHAR(200))                               AS nome_franquia,
        COUNT(*)                                                  AS num_leads,
        SUM(CASE WHEN status_lead = 'Ganho' THEN 1 ELSE 0 END)    AS num_ganhos,
        SUM(num_tentativas_contato)                               AS num_tentativas_contato
    FROM raw.leads_pipeline
    GROUP BY canal_origem
),
por_franquia AS (
    SELECT
        2                                                         AS ordem_bloco,
        l.canal_origem,
        f.nome_franquia,
        COUNT(*)                                                  AS num_leads,
        SUM(CASE WHEN l.status_lead = 'Ganho' THEN 1 ELSE 0 END)  AS num_ganhos,
        SUM(l.num_tentativas_contato)                             AS num_tentativas_contato
    FROM raw.leads_pipeline l
    JOIN raw.franquias f ON f.id_franquia = l.id_franquia
    GROUP BY l.canal_origem, f.nome_franquia
),
metricas AS (
    SELECT
        *,
        CAST(100.0 * num_ganhos / num_leads                       AS DECIMAL(5,1)) AS taxa_conv_lead_ganho_pct,
        CAST(1.0 * num_tentativas_contato / NULLIF(num_ganhos, 0) AS DECIMAL(6,1)) AS tentativas_contato_por_ganho
    FROM (SELECT * FROM por_canal UNION ALL SELECT * FROM por_franquia) AS u
),
ranqueado AS (
    SELECT
        *,
        CASE WHEN ordem_bloco = 1
             THEN RANK() OVER (PARTITION BY ordem_bloco ORDER BY tentativas_contato_por_ganho)
        END AS ranking_roi_esforco
    FROM metricas
)
SELECT
    canal_origem,
    nome_franquia,
    num_leads,
    num_ganhos,
    taxa_conv_lead_ganho_pct,
    num_tentativas_contato,
    tentativas_contato_por_ganho,
    ranking_roi_esforco,
    CASE WHEN num_leads < CASE WHEN ordem_bloco = 1 THEN 30 ELSE 10 END
         THEN 'Amostra pequena' ELSE '' END AS alerta_amostra
FROM ranqueado
ORDER BY ordem_bloco, ranking_roi_esforco, taxa_conv_lead_ganho_pct DESC;
