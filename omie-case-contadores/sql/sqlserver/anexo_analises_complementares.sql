-- Anexo · views do Power BI e análises de apoio (governança)
-- A0 views | A1 penetração por porte | A2 definições de cliente | A3 leads parados
-- A4 validação dos tiers | A5 tiers por região | A6/A7 reconciliação de receita
USE OMIE;
GO

-- A0 Views (rodar uma vez)
CREATE OR ALTER VIEW dbo.stg_lead AS
SELECT
    l.*,
    CASE WHEN l.status_lead = 'Ganho'   THEN 1 ELSE 0 END AS is_ganho,
    CASE WHEN l.status_lead = 'Perdido' THEN 1 ELSE 0 END AS is_perdido,
    CASE WHEN l.status_lead IN ('Ganho', 'Perdido') THEN 0 ELSE 1 END AS is_aberto,
    CASE WHEN l.status_lead NOT IN ('Ganho', 'Perdido') AND l.tempo_ciclo_dias > 90 THEN 1 ELSE 0 END AS is_parado,
    CASE l.status_lead
        WHEN 'Novo'              THEN 1
        WHEN 'Em Contato'        THEN 2
        WHEN N'Reunião Agendada' THEN 3
        WHEN 'Proposta Enviada'  THEN 4
        ELSE 0 END AS estagio_ordem
FROM raw.leads_pipeline l;
GO

CREATE OR ALTER VIEW dbo.stg_contador AS
WITH cli AS (
    SELECT
        id_contador,
        MAX(CASE WHEN status_cliente = 'Ativo' THEN 1 ELSE 0 END)          AS is_cliente_ativo,
        COUNT(*)                                                           AS n_contratos,
        SUM(CASE WHEN status_cliente = 'Ativo' THEN mrr_atual ELSE 0 END) AS mrr_atual
    FROM raw.clientes_ativos
    GROUP BY id_contador
),
lds AS (
    SELECT
        id_contador,
        COUNT(*)                                                    AS n_leads,
        SUM(num_tentativas_contato)                                 AS tentativas_total,
        MAX(CASE WHEN estagio_ordem > 0 THEN 1 ELSE 0 END)          AS tem_lead_aberto,
        MAX(CASE WHEN is_parado = 0 THEN estagio_ordem ELSE 0 END)  AS estagio_aberto_ordem,
        MAX(is_parado)                                              AS tem_lead_parado,
        MAX(is_perdido)                                             AS ja_perdido
    FROM dbo.stg_lead
    GROUP BY id_contador
),
frq AS (
    SELECT DISTINCT estado FROM raw.franquias
)
SELECT
    m.id_contador, m.estado, m.regiao, m.cidade, m.porte, m.segmento_principal,
    m.anos_mercado, m.num_clientes_ativos, m.num_funcionarios, m.erp_atual,
    CASE WHEN m.erp_atual IN ('Nenhum', 'Planilha') THEN N'Não' ELSE 'Sim' END AS usa_erp,
    m.usa_erp                                   AS usa_erp_origem,
    m.tem_site, m.presenca_linkedin,
    m.status_omie                               AS status_omie_origem,
    m.converteu_omie,
    m.score_potencial_latente,
    -- endereçável (regra do 1a)
    CASE WHEN m.status_omie <> 'Cliente Ativo' AND m.converteu_omie = 0 THEN 1 ELSE 0 END AS is_enderecavel,
    -- contrato ativo em clientes_ativos
    COALESCE(c.is_cliente_ativo, 0)             AS is_cliente_ativo,
    COALESCE(c.mrr_atual, 0)                    AS mrr_atual,
    COALESCE(l.n_leads, 0)                      AS n_leads,
    COALESCE(l.tentativas_total, 0)             AS tentativas_total,
    COALESCE(l.tem_lead_aberto, 0)              AS tem_lead_aberto,
    COALESCE(l.estagio_aberto_ordem, 0)         AS estagio_aberto_ordem,
    CASE COALESCE(l.estagio_aberto_ordem, 0)
        WHEN 4 THEN 'Proposta Enviada'
        WHEN 3 THEN N'Reunião Agendada'
        WHEN 2 THEN 'Em Contato'
        WHEN 1 THEN 'Novo'
        ELSE 'Sem lead ativo'
    END                                         AS estagio_aberto,
    COALESCE(l.tem_lead_parado, 0)              AS tem_lead_parado,
    COALESCE(l.ja_perdido, 0)                   AS ja_perdido,
    CASE WHEN f.estado IS NOT NULL THEN 1 ELSE 0 END AS tem_franquia_no_estado
FROM raw.mercado_contadores m
LEFT JOIN cli c ON c.id_contador = m.id_contador
LEFT JOIN lds l ON l.id_contador = m.id_contador
LEFT JOIN frq f ON f.estado = m.estado;
GO

CREATE OR ALTER VIEW dbo.stg_execucao_franquia_mes AS
SELECT
    e.id_franquia, f.nome_franquia, e.estado, e.regiao, e.mes_ref,
    COUNT(*)                   AS registros_ec,
    f.num_executivos,
    SUM(e.leads_trabalhados)   AS leads_trabalhados,
    SUM(e.reunioes_realizadas) AS reunioes_realizadas,
    SUM(e.propostas_enviadas)  AS propostas_enviadas,
    SUM(e.fechamentos)         AS fechamentos,
    SUM(e.nmrr_gerado)         AS nmrr_gerado
FROM raw.execucao_comercial e
JOIN raw.franquias f ON f.id_franquia = e.id_franquia
GROUP BY e.id_franquia, f.nome_franquia, e.estado, e.regiao, e.mes_ref, f.num_executivos;
GO

CREATE OR ALTER VIEW dbo.vw_tiers_prospects AS
-- Tier do 1c (mesma regra de desafio1c_tiers_prioridade.sql)
WITH niveis AS (
    SELECT
        c.*,
        CASE WHEN num_clientes_ativos > 24  THEN 3
             WHEN num_clientes_ativos >= 15 THEN 2
             ELSE 1 END                                            AS nivel_fit,
        -- estagio_aberto_ordem já exclui leads parados (> 90 dias) e 'Sem Resposta'
        CASE WHEN estagio_aberto_ordem >= 3 THEN 2
             WHEN estagio_aberto_ordem >= 1 THEN 1
             ELSE 0 END                                            AS nivel_engajamento
    FROM dbo.stg_contador c
    WHERE is_enderecavel = 1
)
SELECT
    *,
    CASE nivel_fit WHEN 3 THEN 'Alto' WHEN 2 THEN N'Médio' ELSE 'Baixo' END         AS fit_carteira,
    CASE nivel_engajamento WHEN 2 THEN 'Quente' WHEN 1 THEN 'Morno' ELSE 'Frio' END AS engajamento,
    nivel_fit + nivel_engajamento                                                   AS score_prioridade,
    CASE WHEN nivel_fit + nivel_engajamento >= 4 THEN 'Alta'
         WHEN nivel_fit + nivel_engajamento = 3  THEN N'Média'
         ELSE 'Baixa' END                                                           AS tier_prioridade
FROM niveis;
GO

-- Desafio 2: propensão = média das taxas de conversão das categorias do contador nos 5 drivers
-- (fórmula da aba Base do Excel); Alta >= P70, Média >= P30
CREATE OR ALTER VIEW dbo.vw_propensao_d2 AS
WITH b AS (
    SELECT
        id_contador, converteu_omie, porte, erp_atual,
        CASE WHEN anos_mercado < 8 THEN 1 WHEN anos_mercado < 16 THEN 2 ELSE 3 END              AS f_anos,
        CASE WHEN num_clientes_ativos < 20 THEN 1 WHEN num_clientes_ativos < 50 THEN 2
             WHEN num_clientes_ativos < 150 THEN 3 ELSE 4 END                                    AS f_clientes,
        CASE WHEN num_funcionarios <= 2 THEN 1 WHEN num_funcionarios <= 5 THEN 2
             WHEN num_funcionarios <= 15 THEN 3 ELSE 4 END                                       AS f_funcionarios
    FROM raw.mercado_contadores
),
p AS (
    SELECT
        id_contador,
        converteu_omie,
        -- 10 casas: evita empate de ponto flutuante no percentil
        ROUND(( AVG(CAST(converteu_omie AS FLOAT)) OVER (PARTITION BY porte)
              + AVG(CAST(converteu_omie AS FLOAT)) OVER (PARTITION BY f_funcionarios)
              + AVG(CAST(converteu_omie AS FLOAT)) OVER (PARTITION BY f_clientes)
              + AVG(CAST(converteu_omie AS FLOAT)) OVER (PARTITION BY erp_atual)
              + AVG(CAST(converteu_omie AS FLOAT)) OVER (PARTITION BY f_anos) ) / 5, 10)        AS propensao
    FROM b
),
c AS (
    SELECT
        *,
        PERCENTILE_CONT(0.7) WITHIN GROUP (ORDER BY propensao) OVER () AS p70,
        PERCENTILE_CONT(0.3) WITHIN GROUP (ORDER BY propensao) OVER () AS p30
    FROM p
)
SELECT
    id_contador,
    propensao,
    CASE WHEN propensao >= p70 THEN 'Alta' WHEN propensao >= p30 THEN N'Média' ELSE 'Baixa' END AS faixa_propensao,
    CASE WHEN propensao >= p70 THEN 1 WHEN propensao >= p30 THEN 2 ELSE 3 END                    AS faixa_ordem
FROM c;
GO

-- Desafio 1a para o Power BI (ordem_linha substitui o ORDER BY)
CREATE OR ALTER VIEW dbo.vw_desafio1a_mercado AS
WITH totais AS (
    SELECT
        regiao,
        porte,
        COUNT(id_contador) AS num_contadores,
        GROUPING(regiao)   AS e_total_brasil,
        GROUPING(porte)    AS e_total_regiao
    FROM raw.mercado_contadores
    WHERE status_omie <> 'Cliente Ativo'
      AND converteu_omie = 0
    GROUP BY GROUPING SETS ((regiao, porte), (regiao), ())
),
bases AS (
    SELECT
        *,
        SUM(CASE WHEN e_total_regiao = 1 THEN num_contadores END) OVER (PARTITION BY regiao, e_total_brasil) AS total_da_regiao,
        SUM(CASE WHEN e_total_brasil = 1 THEN num_contadores END) OVER ()                                    AS total_brasil
    FROM totais
)
SELECT
    CASE WHEN e_total_brasil = 1 THEN 'Total Brasil' ELSE regiao END             AS regiao,
    CASE WHEN e_total_brasil = 1 THEN 'Todos os portes'
         WHEN e_total_regiao = 1 THEN N'Total da região' ELSE porte END          AS porte,
    num_contadores,
    CAST(100.0 * num_contadores / total_da_regiao AS DECIMAL(5,1))                AS participacao_regiao_pct,
    CAST(100.0 * num_contadores / total_brasil    AS DECIMAL(5,1))                AS participacao_total_pct,
    e_total_brasil,
    e_total_regiao,
    ROW_NUMBER() OVER (ORDER BY e_total_brasil, total_da_regiao DESC, e_total_regiao,
                       CASE porte WHEN N'Grande Escritório'  THEN 1
                                  WHEN N'Médio Escritório'   THEN 2
                                  WHEN N'Pequeno Escritório' THEN 3
                                  ELSE 4 END)                                     AS ordem_linha
FROM bases;
GO

-- Desafio 1b para o Power BI
CREATE OR ALTER VIEW dbo.vw_desafio1b_conversao AS
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
)
SELECT
    *,
    CASE WHEN ordem_bloco = 1
         THEN RANK() OVER (PARTITION BY ordem_bloco ORDER BY tentativas_contato_por_ganho)
    END AS ranking_roi_esforco,
    CASE WHEN num_leads < CASE WHEN ordem_bloco = 1 THEN 30 ELSE 10 END
         THEN 'Amostra pequena' ELSE '' END AS alerta_amostra
FROM metricas;
GO

-- Conferência: 350 / 600 / 180 / 303 (tier 1c: Alta 21 | Média 96 | Baixa 186)
SELECT (SELECT COUNT(*) FROM dbo.stg_lead)                  AS stg_lead,
       (SELECT COUNT(*) FROM dbo.stg_contador)              AS stg_contador,
       (SELECT COUNT(*) FROM dbo.stg_execucao_franquia_mes) AS stg_execucao_franquia_mes,
       (SELECT COUNT(*) FROM dbo.vw_tiers_prospects)        AS vw_tiers_prospects;
GO

-- A1 Penetração e carteira endereçável por porte
SELECT
    porte,
    COUNT(*)                                                               AS universo_contadores,
    COUNT(*) - SUM(is_enderecavel)                                         AS clientes_ou_convertidos,
    CAST(100.0 * (COUNT(*) - SUM(is_enderecavel)) / COUNT(*) AS DECIMAL(5,1)) AS penetracao_pct,
    SUM(is_enderecavel)                                                    AS contadores_enderecaveis,
    SUM(CASE WHEN is_enderecavel = 1 THEN num_clientes_ativos ELSE 0 END)  AS carteira_enderecavel,
    CAST(100.0 * SUM(CASE WHEN is_enderecavel = 1 THEN num_clientes_ativos ELSE 0 END)
         / SUM(SUM(CASE WHEN is_enderecavel = 1 THEN num_clientes_ativos ELSE 0 END)) OVER () AS DECIMAL(5,1)) AS pct_carteira_enderecavel
FROM dbo.stg_contador
GROUP BY porte
ORDER BY carteira_enderecavel DESC;
GO

-- A2 Endereçáveis em cada definição de "cliente"
-- A2.1 Endereçáveis por região em cada régua
SELECT
    COALESCE(regiao, 'Total')                                                       AS regiao,
    SUM(CASE WHEN is_cliente_ativo = 0 THEN 1 ELSE 0 END)                           AS enderecavel_por_contratos,
    SUM(CASE WHEN status_omie_origem <> 'Cliente Ativo' THEN 1 ELSE 0 END)          AS enderecavel_por_status_omie,
    SUM(is_enderecavel)                                                             AS enderecavel_regra_adotada,
    SUM(CASE WHEN is_enderecavel = 1 AND is_cliente_ativo = 1 THEN 1 ELSE 0 END)    AS adotada_mas_com_contrato_ativo
FROM dbo.stg_contador
GROUP BY GROUPING SETS ((regiao), ())
ORDER BY GROUPING(regiao), enderecavel_regra_adotada DESC;
GO

-- A2.2 Onde as fontes se contradizem (contagens únicas)
SELECT
    SUM(CASE WHEN is_cliente_ativo = 1 AND status_omie_origem <> 'Cliente Ativo' THEN 1 ELSE 0 END) AS contrato_ativo_sem_status_cliente,
    SUM(CASE WHEN is_cliente_ativo = 0 AND status_omie_origem = 'Cliente Ativo' THEN 1 ELSE 0 END)  AS status_cliente_sem_contrato_ativo,
    SUM(CASE WHEN status_omie_origem = 'Cliente Ativo' AND converteu_omie = 0 THEN 1 ELSE 0 END)    AS status_cliente_com_converteu_0,
    SUM(CASE WHEN status_omie_origem = 'Sem Contato' AND converteu_omie = 1 THEN 1 ELSE 0 END)      AS sem_contato_com_converteu_1
FROM dbo.stg_contador;
GO

-- A3 Leads abertos parados há mais de 90 dias
SELECT
    status_lead,
    COUNT(*)                                                   AS leads_abertos,
    CAST(AVG(1.0 * tempo_ciclo_dias) AS DECIMAL(6,0))          AS idade_media_dias,
    SUM(is_parado)                                             AS parados_90d,
    CAST(100.0 * SUM(is_parado) / COUNT(*) AS DECIMAL(5,1))    AS pct_parado
FROM dbo.stg_lead
WHERE is_aberto = 1
GROUP BY status_lead, estagio_ordem
ORDER BY estagio_ordem DESC;
GO

-- A4 Validação dos tiers do 1c (score_potencial_latente só valida, não pontua)
SELECT
    fit_carteira,
    SUM(CASE WHEN engajamento = 'Quente' THEN 1 ELSE 0 END) AS quente,
    SUM(CASE WHEN engajamento = 'Morno'  THEN 1 ELSE 0 END) AS morno,
    SUM(CASE WHEN engajamento = 'Frio'   THEN 1 ELSE 0 END) AS frio
FROM dbo.vw_tiers_prospects
GROUP BY fit_carteira, nivel_fit
ORDER BY nivel_fit DESC;
GO

SELECT
    tier_prioridade,
    COUNT(*)                                                    AS contadores,
    CAST(AVG(1.0 * num_clientes_ativos) AS DECIMAL(6,0))        AS carteira_media,
    SUM(CASE WHEN engajamento <> 'Frio' THEN 1 ELSE 0 END)      AS com_lead_ativo,
    SUM(tem_lead_parado)                                        AS com_lead_parado,
    CAST(AVG(score_potencial_latente) AS DECIMAL(5,3))          AS score_latente_medio
FROM dbo.vw_tiers_prospects
GROUP BY tier_prioridade
ORDER BY CASE tier_prioridade WHEN 'Alta' THEN 1 WHEN N'Média' THEN 2 ELSE 3 END;
GO

-- A5 Tiers por região
SELECT
    regiao,
    SUM(CASE WHEN tier_prioridade = 'Alta'   THEN 1 ELSE 0 END) AS alta,
    SUM(CASE WHEN tier_prioridade = N'Média' THEN 1 ELSE 0 END) AS media,
    SUM(CASE WHEN tier_prioridade = 'Baixa'  THEN 1 ELSE 0 END) AS baixa,
    CAST(100.0 * SUM(CASE WHEN tier_prioridade = 'Alta' THEN 1 ELSE 0 END) / COUNT(*) AS DECIMAL(5,1)) AS pct_alta
FROM dbo.vw_tiers_prospects
GROUP BY regiao
ORDER BY alta DESC;
GO

-- A6/A7 Reconciliação do nMRR entre as fontes
-- A6 Rede
SELECT 'franquias.realizado_nmrr_mes' AS fonte, 'Snapshot mensal (oficial de meta)' AS granularidade,
       CAST(SUM(1.0 * realizado_nmrr_mes) AS DECIMAL(12,0)) AS nmrr_mensal, CAST(NULL AS DECIMAL(8,1)) AS fechamentos_mes
FROM raw.franquias
UNION ALL
SELECT 'execucao_comercial.nmrr_gerado', N'Média mensal 2024 (funil dos ECs)',
       CAST(SUM(nmrr_gerado) / COUNT(DISTINCT mes_ref) AS DECIMAL(12,0)),
       CAST(1.0 * SUM(fechamentos) / COUNT(DISTINCT mes_ref) AS DECIMAL(8,1))
FROM raw.execucao_comercial
UNION ALL
SELECT 'leads_pipeline.mrr_inicial (Ganho)', N'Média mensal 2024-2025 (CRM)',
       CAST(SUM(mrr_inicial) / 24 AS DECIMAL(12,0)),
       CAST(SUM(is_ganho) / 24.0 AS DECIMAL(8,1))
FROM dbo.stg_lead
WHERE is_ganho = 1;
GO

-- A7 Por franquia
WITH exe AS (
    SELECT id_franquia, SUM(nmrr_gerado) / 12.0 AS exec_nmrr_mes
    FROM raw.execucao_comercial
    GROUP BY id_franquia
),
crm AS (
    SELECT id_franquia, SUM(mrr_inicial) / 24.0 AS crm_nmrr_mes, SUM(is_ganho) AS ganhos_crm
    FROM dbo.stg_lead
    WHERE id_franquia IS NOT NULL AND is_ganho = 1
    GROUP BY id_franquia
)
SELECT
    f.nome_franquia,
    f.realizado_nmrr_mes                                                     AS oficial_nmrr_mes,
    CAST(e.exec_nmrr_mes AS DECIMAL(12,0))                                   AS execucao_nmrr_mes,
    CAST(COALESCE(c.crm_nmrr_mes, 0) AS DECIMAL(12,0))                       AS crm_nmrr_mes,
    CAST(f.realizado_nmrr_mes / NULLIF(e.exec_nmrr_mes, 0) AS DECIMAL(6,1))  AS oficial_sobre_execucao,
    COALESCE(c.ganhos_crm, 0)                                                AS ganhos_no_crm_24m
FROM raw.franquias f
LEFT JOIN exe e ON e.id_franquia = f.id_franquia
LEFT JOIN crm c ON c.id_franquia = f.id_franquia
ORDER BY oficial_sobre_execucao DESC;
GO
