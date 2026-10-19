"""Extrai os números do deck direto das fontes (DuckDB + planilhas recalculadas) para deck/data.json.

Assim todo número do deck é rastreável às queries e planilhas. Uso: python src/deck_data.py
"""
import json

import openpyxl

from load import ROOT, connect

con = connect(read_only=True)


def rows(sql):
    cur = con.execute(sql)
    cols = [d[0] for d in cur.description]
    return [dict(zip(cols, r)) for r in cur.fetchall()]


def one(sql):
    return con.execute(sql).fetchone()


d = {}
# Mercado endereçável (pergunta 1a e slides 1-2): só mercado_contadores, sem clientes e sem convertidos
# (premissa do usuário: converteu_omie = 1 = já convertido).
# A lista de prospecção (Desafios 2-3, slides 3-5) segue os contratos ativos de clientes_ativos.
CLI = "(status_omie = 'Cliente Ativo' OR converteu_omie = 1)"
uni, cli, end, cart, churn_end = one(f"""
    SELECT COUNT(*), SUM(CASE WHEN {CLI} THEN 1 ELSE 0 END), SUM(CASE WHEN NOT {CLI} THEN 1 ELSE 0 END),
           SUM(CASE WHEN NOT {CLI} THEN num_clientes_ativos ELSE 0 END),
           SUM(CASE WHEN NOT {CLI} AND status_omie = 'Churned' THEN 1 ELSE 0 END)
    FROM mercado_contadores""")
d["mercado"] = {"universo": uni, "clientes": int(cli), "enderecaveis": int(end), "carteira_enderecavel": int(cart),
                "penetracao": cli / uni, "ex_clientes": int(churn_end)}
d["porte"] = rows(f"""
    SELECT porte, SUM(CASE WHEN NOT {CLI} THEN 1 ELSE 0 END) AS enderecaveis,
           SUM(CASE WHEN {CLI} THEN 1 ELSE 0 END) * 1.0 / COUNT(*) AS penetracao,
           SUM(CASE WHEN NOT {CLI} THEN num_clientes_ativos ELSE 0 END) AS carteira
    FROM mercado_contadores GROUP BY porte ORDER BY carteira DESC""")
d["regiao"] = rows(f"""
    SELECT regiao, SUM(CASE WHEN NOT {CLI} THEN 1 ELSE 0 END) AS enderecaveis,
           SUM(CASE WHEN NOT {CLI} THEN num_clientes_ativos ELSE 0 END) AS carteira
    FROM mercado_contadores GROUP BY regiao ORDER BY enderecaveis DESC""")
icp_n, icp_cart, mtms = one(f"""
    SELECT SUM(CASE WHEN icp THEN 1 ELSE 0 END), SUM(CASE WHEN icp THEN num_clientes_ativos ELSE 0 END),
           SUM(CASE WHEN estado IN ('MT', 'MS') THEN 1 ELSE 0 END)
    FROM (SELECT *, num_clientes_ativos >= 50 AND num_funcionarios >= 6
                    AND erp_atual NOT IN ('Nenhum', 'Planilha', 'ContaAzul') AS icp
          FROM mercado_contadores WHERE NOT {CLI})""")
icp_total = one("""SELECT COUNT(*) FROM mercado_contadores WHERE num_clientes_ativos >= 50 AND num_funcionarios >= 6
                   AND erp_atual NOT IN ('Nenhum', 'Planilha', 'ContaAzul')""")[0]
d["icp_mercado"] = {"enderecaveis": int(icp_n), "pct_carteira": icp_cart / cart, "mt_ms": int(mtms), "total": int(icp_total)}
d["converteu_vs_status"] = dict(zip(["cliente_ativo_com_0", "sem_contato_com_1"], one("""
    SELECT SUM(CASE WHEN status_omie = 'Cliente Ativo' AND converteu_omie = 0 THEN 1 ELSE 0 END),
           SUM(CASE WHEN status_omie = 'Sem Contato' AND converteu_omie = 1 THEN 1 ELSE 0 END)
    FROM mercado_contadores""")))
d["divergencia_cliente"] = dict(zip(["contrato_sem_status", "status_sem_contrato"], one("""
    SELECT SUM(CASE WHEN is_cliente_ativo = 1 AND status_omie_origem <> 'Cliente Ativo' THEN 1 ELSE 0 END),
           SUM(CASE WHEN is_cliente_ativo = 0 AND status_omie_origem = 'Cliente Ativo' THEN 1 ELSE 0 END)
    FROM stg_contador""")))
d["prospeccao"] = dict(zip(["lista", "com_contrato_ativo"], one("""
    SELECT COUNT(*), SUM(is_cliente_ativo) FROM stg_contador
    WHERE status_omie_origem <> 'Cliente Ativo' AND converteu_omie = 0""")))
d["pipeline"] = dict(zip(["abertos", "parados", "win_rate_rede"], one(
    "SELECT SUM(is_aberto), SUM(is_parado), SUM(is_ganho) * 1.0 / (SUM(is_ganho) + SUM(is_perdido)) FROM stg_lead")))
d["abaixo_meta"] = dict(zip(["n", "realizado", "meta"], one(
    "SELECT COUNT(*), SUM(realizado_nmrr_mes), SUM(meta_nmrr_mes) FROM franquias WHERE realizado_nmrr_mes < meta_nmrr_mes")))
d["receita_fontes"] = [r["nmrr_mensal"] for r in rows("""
    SELECT ROUND(SUM(realizado_nmrr_mes), 0) AS nmrr_mensal FROM franquias
    UNION ALL SELECT ROUND(SUM(nmrr_gerado) / 12, 0) FROM execucao_comercial
    UNION ALL SELECT ROUND(SUM(mrr_inicial) / 24, 0) FROM stg_lead WHERE is_ganho = 1""")]
d["win_rate_canal"] = rows("""
    SELECT canal_origem, SUM(is_ganho) * 1.0 / NULLIF(SUM(is_ganho) + SUM(is_perdido), 0) AS win_rate,
           SUM(num_tentativas_contato) * 1.0 / SUM(is_ganho) AS tentativas_por_fechamento
    FROM stg_lead GROUP BY canal_origem ORDER BY tentativas_por_fechamento""")
d["conv_canal"] = rows("""
    SELECT canal_origem, SUM(is_ganho) * 1.0 / COUNT(*) AS conversao FROM stg_lead GROUP BY canal_origem""")
d["tiers"] = dict(zip(["alta", "media", "baixa", "alta_quente"], one("""
    SELECT SUM(CASE WHEN tier_prioridade = 'Alta' THEN 1 ELSE 0 END), SUM(CASE WHEN tier_prioridade = 'Média' THEN 1 ELSE 0 END),
           SUM(CASE WHEN tier_prioridade = 'Baixa' THEN 1 ELSE 0 END),
           SUM(CASE WHEN tier_prioridade = 'Alta' AND engajamento = 'Quente' THEN 1 ELSE 0 END)
    FROM vw_tiers_prospects""")))
con.close()

# --- Desafio 2
wb2 = openpyxl.load_workbook(ROOT / "output" / "Desafio2_Propensao.xlsx", data_only=True)
imp, dr, cr = wb2["Impacto_nMRR"], wb2["Drivers"], wb2["Como_Avaliar"]
d["drivers"] = {  # Desafio 2 descritivo (sem modelo estatístico)
    "taxa_media": dr["B3"].value,
    "ranking": sorted(({"variavel": dr.cell(row=r, column=8).value, "amplitude": dr.cell(row=r, column=9).value,
                        "rank": dr.cell(row=r, column=10).value} for r in range(6, 15)), key=lambda x: x["rank"]),
    "porte": [{"porte": cr.cell(row=r, column=1).value, "target": cr.cell(row=r, column=3).value,
               "win_real": cr.cell(row=r, column=6).value} for r in range(26, 30)],
    "icp_conv": dr["J29"].value, "fora_icp_conv": dr["J30"].value,
    "icp_lift": dr["J31"].value, "icp_lift_fora": dr["J32"].value,
    "conv_faixas": cr["C5"].value, "top30_conv": cr["C19"].value, "top30_lift": cr["D19"].value,
    "captura": cr["C6"].value, "win_alta": cr["G19"].value, "win_demais": cr["G22"].value,
    "n_alta_desfecho": cr["E19"].value, "ciclo_alta": cr["H19"].value, "ciclo_demais": cr["H22"].value,
    "alta_enderecaveis": wb2["Respostas"]["B7"].value,
    "premissas": {"ecs": imp["C5"].value, "leads_ec": imp["C6"].value, "ticket": imp["C11"].value,
                  "churn": imp["C12"].value},
    "uplift": [imp[f"{c}23"].value for c in "BCD"],
    "nmrr_oficial": [imp[f"{c}24"].value for c in "BCD"],
}

# --- Desafio 3
wb3 = openpyxl.load_workbook(ROOT / "output" / "Desafio3_Capacity.xlsx", data_only=True)
g, oc, ec, sim, reg = (wb3["Gap_Meta"], wb3["Ondas_Capacidade"], wb3["EC_Adicional"], wb3["Simulador_Funil"],
                       wb3["Regioes"])
franq = []
for r in range(5, 20):
    franq.append({"nome": g[f"B{r}"].value.replace("Franquia ", ""), "regiao": g[f"D{r}"].value,
                  "meses": g[f"E{r}"].value, "atingimento": g[f"I{r}"].value, "gap": g[f"J{r}"].value,
                  "real_ec": g[f"L{r}"].value, "meta_ec": g[f"K{r}"].value, "diag": g[f"AA{r}"].value,
                  "tendencia": g[f"Y{r}"].value, "meta": g[f"G{r}"].value, "realizado": g[f"H{r}"].value,
                  "ecs": g[f"F{r}"].value, "sem_fech": g[f"Z{r}"].value})
d["franquias"] = sorted(franq, key=lambda f: f["atingimento"])
d["gap_total"] = g["J20"].value
d["real_ec_rede"] = g["L20"].value
d["sem_fech_rede"] = g["Z21"].value
d["corr_maturidade"] = g["D26"].value
d["corr_ecs"] = g["D27"].value
d["corr_data_inicio"] = g["D31"].value
d["corr_funil"] = g["D30"].value
d["bench_pf"] = g["T21"].value
d["regioes_gap"] = [{"regiao": reg[f"A{r}"].value, "gap": reg[f"I{r}"].value, "pct_abaixo": reg[f"D{r}"].value,
                     "atingimento": reg[f"H{r}"].value}
                    for r in range(4, 9)]
d["ondas"] = [{"onda": oc[f"A{r}"].value, "n": oc[f"F{r}"].value, "ve": oc[f"G{r}"].value} for r in range(5, 10)]
d["sim_pf"] = {"pp": sim["C7"].value, "uplift": sim["B11"].value, "cobertura": sim["B14"].value}
pz = wb3["Priorizacao"]
cart_prop = [(r[5], r[23]) for r in pz.iter_rows(min_row=2, values_only=True) if r[23] is not None]
mc_, mp_ = (sum(x) / len(cart_prop) for x in zip(*cart_prop))
cov = sum((c - mc_) * (p - mp_) for c, p in cart_prop)
d["corr_prop_carteira"] = cov / (sum((c - mc_) ** 2 for c, _ in cart_prop) * sum((p - mp_) ** 2 for _, p in cart_prop)) ** 0.5
rank = {ec[f"L{r}"].value: r for r in range(5, 20)}
r1 = rank[1]
d["ec_adicional"] = {
    "franquia": ec[f"B{r1}"].value.replace("Franquia ", ""), "meses": ec[f"C{r1}"].value,
    "demanda_ec": ec[f"F{r1}"].value, "demanda_rede": sum(ec[f"E{r}"].value for r in range(5, 20)) / sum(ec[f"D{r}"].value for r in range(5, 20)),
    "carga_ec": ec[f"O{r1}"].value, "carga_rede": g["N20"].value / sum(1 for _ in wb3["Execucao"].iter_rows(min_row=2)),
    "pot": ec[f"H{r1}"].value, "pot_2": ec[f"H{rank[2]}"].value, "conv_2": ec[f"G{rank[2]}"].value,
    "conv": ec[f"G{r1}"].value, "conv_rede": g["Q20"].value / g["N20"].value, "gap": ec[f"I{r1}"].value,
    "rank_demanda": ec[f"M{r1}"].value, "rank_conv": ec[f"N{r1}"].value,
    "segunda": ec[f"B{rank[2]}"].value.replace("Franquia ", ""),
    "sem_franquia": sum(1 for r in wb3["Priorizacao"].iter_rows(min_row=2, values_only=True) if r[14] == "Sem franquia"),
}

out = ROOT / "deck" / "data.json"
out.write_text(json.dumps(d, ensure_ascii=False, indent=1, default=float), encoding="utf-8")
print("Gerado:", out)
