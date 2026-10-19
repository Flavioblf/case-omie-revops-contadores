/* =====================================================================
   DESAFIO 1a — Mercado endereçável por região e porte
   Pergunta: tamanho do mercado endereçável (contadores ainda não clientes da Omie)
             por região, segmentado por porte, com volumetria e % do total por região.

   Endereçável = status_omie diferente de 'Cliente Ativo' E converteu_omie = 0
   (premissa: converteu_omie = 1 já é convertido).
   Colunas: regiao e porte vêm de mercado_contadores;
            num_contadores = contagem de id_contador;
            participacao_regiao_pct = peso do porte dentro da região;
            participacao_total_pct  = peso no Brasil (na linha 'Total da região' = participação da região).
   T-SQL (SQL Server) · banco OMIE · esquema raw
   ===================================================================== */
USE OMIE;
GO

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
    CAST(100.0 * num_contadores / total_brasil    AS DECIMAL(5,1))                AS participacao_total_pct
FROM bases
ORDER BY
    e_total_brasil,
    total_da_regiao DESC,
    e_total_regiao,
    CASE porte WHEN N'Grande Escritório'  THEN 1
               WHEN N'Médio Escritório'   THEN 2
               WHEN N'Pequeno Escritório' THEN 3
               ELSE 4 END;
