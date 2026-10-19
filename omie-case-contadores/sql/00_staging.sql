-- =====================================================================
-- 00_staging.sql — camada canônica sobre as tabelas brutas
-- Dialeto ANSI (roda em DuckDB local e em Databricks SQL)
-- =====================================================================

-- ---------------------------------------------------------------------
-- stg_lead: 1 linha por lead, com a situação canônica do funil.
--   win rate     = Ganho / (Ganho + Perdido)        (definição de mercado)
--   Sem Resposta = lead estagnado, NÃO é encerrado (fica fora do win rate)
--   lead parado  = aberto há mais de 90 dias (tempo_ciclo_dias = idade do lead aberto;
--                  ~2x o ciclo médio dos Ganhos, de ~45 dias)
-- ---------------------------------------------------------------------
CREATE OR REPLACE VIEW stg_lead AS
SELECT
    l.*,
    CASE WHEN status_lead = 'Ganho'   THEN 1 ELSE 0 END AS is_ganho,
    CASE WHEN status_lead = 'Perdido' THEN 1 ELSE 0 END AS is_perdido,
    CASE WHEN status_lead IN ('Ganho', 'Perdido') THEN 0 ELSE 1 END AS is_aberto,
    CASE WHEN status_lead NOT IN ('Ganho', 'Perdido') AND tempo_ciclo_dias > 90 THEN 1 ELSE 0 END AS is_parado,
    CASE status_lead
        WHEN 'Novo'             THEN 1
        WHEN 'Em Contato'       THEN 2
        WHEN 'Reunião Agendada' THEN 3
        WHEN 'Proposta Enviada' THEN 4
        ELSE 0 END AS estagio_ordem
FROM leads_pipeline l;

-- ---------------------------------------------------------------------
-- stg_contador: 1 linha por contador, com o status canônico.
-- Regra: "cliente" = existe em clientes_ativos com status_cliente = 'Ativo'.
-- status_omie NÃO é usado como fonte da verdade (não bate com clientes_ativos).
-- usa_erp é recalculado a partir de erp_atual (184 linhas inconsistentes na origem).
-- ---------------------------------------------------------------------
CREATE OR REPLACE VIEW stg_contador AS
WITH cli AS (
    SELECT
        id_contador,
        MAX(CASE WHEN status_cliente = 'Ativo' THEN 1 ELSE 0 END) AS is_cliente_ativo,
        COUNT(*)                                                  AS n_contratos,
        SUM(CASE WHEN status_cliente = 'Ativo' THEN mrr_atual ELSE 0 END) AS mrr_atual
    FROM clientes_ativos
    GROUP BY id_contador
),
lds AS (
    SELECT
        id_contador,
        COUNT(*) AS n_leads,
        SUM(num_tentativas_contato) AS tentativas_total,
        MAX(CASE WHEN estagio_ordem > 0 THEN 1 ELSE 0 END) AS tem_lead_aberto,
        -- estágio mais avançado entre os leads abertos e ATIVOS (até 90 dias); 0 = nenhum
        MAX(CASE WHEN is_parado = 0 THEN estagio_ordem ELSE 0 END) AS estagio_aberto_ordem,
        MAX(is_parado)  AS tem_lead_parado,
        MAX(is_perdido) AS ja_perdido
    FROM stg_lead
    GROUP BY id_contador
),
frq AS (
    SELECT DISTINCT estado FROM franquias
)
SELECT
    m.id_contador,
    m.estado,
    m.regiao,
    m.cidade,
    m.porte,
    m.segmento_principal,
    m.anos_mercado,
    m.num_clientes_ativos,
    m.num_funcionarios,
    m.erp_atual,
    CASE WHEN m.erp_atual IN ('Nenhum', 'Planilha') THEN 'Não' ELSE 'Sim' END AS usa_erp,
    m.usa_erp                                   AS usa_erp_origem,
    m.tem_site,
    m.presenca_linkedin,
    m.status_omie                               AS status_omie_origem,
    m.converteu_omie,
    m.score_potencial_latente,                  -- só para validação; nunca como feature
    COALESCE(c.is_cliente_ativo, 0)             AS is_cliente_ativo,
    CASE WHEN c.id_contador IS NOT NULL AND c.is_cliente_ativo = 0 THEN 1 ELSE 0 END AS is_churned,
    CASE
        WHEN c.is_cliente_ativo = 1     THEN 'Cliente Ativo'
        WHEN c.id_contador IS NOT NULL  THEN 'Churned (win-back)'
        WHEN l.id_contador IS NOT NULL  THEN 'Prospect em pipeline'
        ELSE 'Prospect sem contato'
    END                                         AS status_canonico,
    COALESCE(c.mrr_atual, 0)                    AS mrr_atual,
    COALESCE(l.n_leads, 0)                      AS n_leads,
    COALESCE(l.tentativas_total, 0)             AS tentativas_total,
    COALESCE(l.tem_lead_aberto, 0)              AS tem_lead_aberto,
    COALESCE(l.estagio_aberto_ordem, 0)         AS estagio_aberto_ordem,
    CASE COALESCE(l.estagio_aberto_ordem, 0)
        WHEN 4 THEN 'Proposta Enviada'
        WHEN 3 THEN 'Reunião Agendada'
        WHEN 2 THEN 'Em Contato'
        WHEN 1 THEN 'Novo'
        ELSE 'Sem lead ativo'
    END                                         AS estagio_aberto,
    COALESCE(l.tem_lead_parado, 0)              AS tem_lead_parado,
    COALESCE(l.ja_perdido, 0)                   AS ja_perdido,
    CASE WHEN f.estado IS NOT NULL THEN 1 ELSE 0 END AS tem_franquia_no_estado
FROM mercado_contadores m
LEFT JOIN cli c ON c.id_contador = m.id_contador
LEFT JOIN lds l ON l.id_contador = m.id_contador
LEFT JOIN frq f ON f.estado = m.estado;

-- ---------------------------------------------------------------------
-- stg_execucao_franquia_mes: execução agregada por franquia x mês.
-- id_executivo se repete num_executivos vezes por mês (1 ID por franquia),
-- então o grão confiável é franquia x mês; métricas "por EC" = total / num_executivos.
-- ---------------------------------------------------------------------
CREATE OR REPLACE VIEW stg_execucao_franquia_mes AS
SELECT
    e.id_franquia,
    f.nome_franquia,
    e.estado,
    e.regiao,
    e.mes_ref,
    COUNT(*)                     AS registros_ec,
    f.num_executivos,
    SUM(e.leads_trabalhados)     AS leads_trabalhados,
    SUM(e.reunioes_realizadas)   AS reunioes_realizadas,
    SUM(e.propostas_enviadas)    AS propostas_enviadas,
    SUM(e.fechamentos)           AS fechamentos,
    SUM(e.nmrr_gerado)           AS nmrr_gerado
FROM execucao_comercial e
JOIN franquias f ON f.id_franquia = e.id_franquia
GROUP BY e.id_franquia, f.nome_franquia, e.estado, e.regiao, e.mes_ref, f.num_executivos;
