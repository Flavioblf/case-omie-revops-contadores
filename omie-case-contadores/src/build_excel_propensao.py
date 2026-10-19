"""Gera output/Desafio2_Propensao.xlsx: análise DESCRITIVA de propensão à conversão (sem modelo estatístico).

Responde ao Desafio 2 com: drivers de conversão (taxa por característica), top 5, ICP, contraste com o
resultado real, como avaliar a abordagem e o impacto em nMRR (extra) — tudo em fórmulas Excel.
Abas: Respostas (as 4 perguntas), Drivers, Como_Avaliar, Impacto_nMRR e as bases de apoio (Base, Leads, Clientes).
Uso: python src/build_excel_propensao.py   (depois: python src/recalc_excel.py output/Desafio2_Propensao.xlsx)
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from load import ROOT, connect
from xlsx_style import (BORDA, BRL, F_ALERTA, F_SUB, F_TITULO, FILL_DESTAQUE, FILL_INPUT, PCT, PCT1, WRAP, header,
                        larguras, py)

OUT = ROOT / "output" / "Desafio2_Propensao.xlsx"
SQL_TIERS = ROOT / "sql" / "01c_tiers_prioridade.sql"

# ---------------------------------------------------------------- dados
con = connect()
# view do tier 1c (mesma regra da pergunta C); recria a partir do SQL versionado
preambulo = SQL_TIERS.read_text(encoding="utf-8").split("-- @name:")[0]
con.execute(preambulo)
base = con.execute("""
    SELECT c.id_contador, c.estado, c.regiao, c.porte, c.segmento_principal, c.anos_mercado, c.num_clientes_ativos,
           c.num_funcionarios, c.erp_atual, c.tem_site, c.presenca_linkedin, c.status_omie_origem AS status_omie,
           c.converteu_omie, t.score_prioridade, t.tier_prioridade
    FROM stg_contador c
    LEFT JOIN vw_tiers_prospects t USING (id_contador)
    ORDER BY c.id_contador""").df()
leads = con.execute("""
    SELECT id_lead, id_contador, canal_origem, status_lead, mrr_inicial, num_tentativas_contato, tempo_ciclo_dias,
           data_conversao
    FROM leads_pipeline ORDER BY id_lead""").df()
clientes = con.execute("""
    SELECT id_cliente, id_contador, status_cliente, mrr_inicial, mrr_atual
    FROM clientes_ativos ORDER BY id_cliente""").df()
n_ecs = int(con.execute("SELECT SUM(num_executivos) FROM franquias").fetchone()[0])
leads_por_ec, conv_execucao, ticket_execucao = con.execute(
    """SELECT AVG(leads_trabalhados), SUM(fechamentos) / SUM(leads_trabalhados), SUM(nmrr_gerado) / SUM(fechamentos)
       FROM execucao_comercial""").fetchone()
f_leads, f_reun, f_prop, f_fech = con.execute(
    """SELECT SUM(leads_trabalhados), SUM(reunioes_realizadas), SUM(propostas_enviadas), SUM(fechamentos)
       FROM execucao_comercial""").fetchone()
realizado_rede = con.execute("SELECT SUM(realizado_nmrr_mes) FROM franquias").fetchone()[0]
realizado_abaixo, gap_abaixo, n_abaixo = con.execute(
    """SELECT SUM(realizado_nmrr_mes), SUM(meta_nmrr_mes - realizado_nmrr_mes), COUNT(*)
       FROM franquias WHERE realizado_nmrr_mes < meta_nmrr_mes""").fetchone()
con.close()

N = len(base)
L = N + 1


def rng(col):
    return f"Base!${col}$2:${col}${L}"


wb = Workbook()

# ---------------------------------------------------------------- Respostas (preenchida no fim)
ws_resp = wb.active
ws_resp.title = "Respostas"

# ---------------------------------------------------------------- Base
ws = wb.create_sheet("Base")
cols_base = ["id_contador", "estado", "regiao", "porte", "segmento_principal", "anos_mercado", "num_clientes_ativos",
             "num_funcionarios", "erp_atual", "tem_site", "presenca_linkedin", "status_omie", "converteu_omie",
             "faixa_anos", "faixa_clientes", "faixa_funcionarios", "sem_erp", "icp", "enderecavel (regra 1a)",
             "score Fit + Engajamento (1c, 1 a 5)", "tier de prioridade (1c)",
             "taxa do porte", "taxa da faixa de funcionários", "taxa da faixa de clientes", "taxa do ERP",
             "taxa da faixa de anos", "PROPENSÃO (média das 5 taxas)", "faixa de propensão"]
header(ws, 1, cols_base)
campos = ["id_contador", "estado", "regiao", "porte", "segmento_principal", "anos_mercado", "num_clientes_ativos",
          "num_funcionarios", "erp_atual", "tem_site", "presenca_linkedin", "status_omie", "converteu_omie"]
for i, row in enumerate(base.itertuples(index=False), start=2):
    for j, campo in enumerate(campos, start=1):
        ws.cell(row=i, column=j, value=py(getattr(row, campo)))
    if row.tier_prioridade is not None and row.tier_prioridade == row.tier_prioridade:
        ws.cell(row=i, column=20, value=py(row.score_prioridade))
        ws.cell(row=i, column=21, value=row.tier_prioridade)
    ws[f"N{i}"] = f'=IF(F{i}<8,"1. < 8 anos",IF(F{i}<16,"2. 8 a 15 anos","3. 16+ anos"))'
    ws[f"O{i}"] = f'=IF(G{i}<20,"1. < 20",IF(G{i}<50,"2. 20 a 49",IF(G{i}<150,"3. 50 a 149","4. 150+")))'
    ws[f"P{i}"] = f'=IF(H{i}<=2,"1. 1 a 2",IF(H{i}<=5,"2. 3 a 5",IF(H{i}<=15,"3. 6 a 15","4. 16+")))'
    ws[f"Q{i}"] = f'=IF(OR(I{i}="Nenhum",I{i}="Planilha",I{i}="ContaAzul"),1,0)'
    ws[f"R{i}"] = f"=IF(AND(G{i}>=50,H{i}>=6,Q{i}=0),1,0)"
    ws[f"S{i}"] = f'=IF(AND(L{i}<>"Cliente Ativo",M{i}=0),1,0)'
    # propensão = taxa histórica de conversão de quem tem o mesmo perfil nos 5 drivers do top 5 (sem pesos)
    for col, src in zip("VWXYZ", "DPOIN"):
        ws[f"{col}{i}"] = f"=AVERAGEIF(${src}$2:${src}${L},{src}{i},$M$2:$M${L})"
    ws[f"AA{i}"] = f"=ROUND(AVERAGE(V{i}:Z{i}),10)"
    ws[f"AB{i}"] = f'=IF(AA{i}>=Como_Avaliar!$B$15,"Alta",IF(AA{i}>=Como_Avaliar!$B$16,"Média","Baixa"))'
    for col in "VWXYZ":
        ws[f"{col}{i}"].number_format = PCT
    ws[f"AA{i}"].number_format, ws[f"AA{i}"].fill = PCT1, FILL_DESTAQUE
ws.freeze_panes = "B2"
ws.auto_filter.ref = f"A1:AB{L}"
ws.row_dimensions[1].height = 45
for j in range(1, len(cols_base) + 1):
    ws.column_dimensions[get_column_letter(j)].width = 13
ws.column_dimensions["E"].width = 17

# ---------------------------------------------------------------- Drivers
ws = wb.create_sheet("Drivers")
ws["A1"], ws["A1"].font = "Drivers de conversão — taxa de conversão por característica (600 contadores)", F_TITULO
ws["A2"] = ("Amplitude = maior taxa − menor taxa entre as categorias com 30+ contadores. "
            "Quanto maior, mais a característica separa quem converte de quem não converte.")
ws["A3"], ws["B3"] = "Taxa média geral", f"=AVERAGE({rng('M')})"
ws["B3"].number_format, ws["B3"].fill = PCT1, FILL_DESTAQUE
header(ws, 4, ["Variável", "Categoria", "Contadores", "Convertidos", "Taxa", "Lift vs média"])
header(ws, 5, ["Variável", "Amplitude (p.p.)", "Ranking"], start_col=8)
ws["H4"], ws["H4"].font = "Ranking das variáveis", F_SUB
drivers = [("Nº de clientes da carteira", "O"), ("Nº de funcionários", "P"), ("Porte", "D"), ("ERP atual", "I"),
           ("Anos de mercado", "N"), ("Região", "C"), ("Segmento", "E"), ("Tem site", "J"), ("LinkedIn", "K")]
cats_src = {"D": base.porte, "I": base.erp_atual, "C": base.regiao, "E": base.segmento_principal,
            "J": base.tem_site, "K": base.presenca_linkedin}
faixas = {"O": ["1. < 20", "2. 20 a 49", "3. 50 a 149", "4. 150+"],
          "P": ["1. 1 a 2", "2. 3 a 5", "3. 6 a 15", "4. 16+"], "N": ["1. < 8 anos", "2. 8 a 15 anos", "3. 16+ anos"]}
r = 5
for k, (var, col) in enumerate(drivers):
    if var == "Região":
        ws.cell(row=r, column=1, value="Fora do top 5 (efeito pequeno)").font = F_SUB
        r += 1
    cats = faixas.get(col) or sorted(cats_src[col].unique())
    ini = r
    for cat in cats:
        ws.cell(row=r, column=1, value=var)
        ws.cell(row=r, column=2, value=cat)
        ws.cell(row=r, column=3, value=f"=COUNTIF({rng(col)},B{r})")
        ws.cell(row=r, column=4, value=f"=COUNTIFS({rng(col)},B{r},{rng('M')},1)")
        ws.cell(row=r, column=5, value=f"=IF(C{r}>0,D{r}/C{r},0)").number_format = PCT1
        ws.cell(row=r, column=6, value=f"=E{r}/$B$3").number_format = '0.00"x"'
        for c in range(1, 7):
            ws.cell(row=r, column=c).border = BORDA
        r += 1
    rr = 6 + k
    taxas, ns = f"E{ini}:E{r - 1}", f"C{ini}:C{r - 1}"
    ws.cell(row=rr, column=8, value=var)
    ws.cell(row=rr, column=9, value=(f'=(_xlfn.MAXIFS({taxas},{ns},">=30")'
                                     f'-_xlfn.MINIFS({taxas},{ns},">=30"))*100')).number_format = "0"
    ws.cell(row=rr, column=10, value=f"=RANK(I{rr},$I$6:$I${5 + len(drivers)})")
    r += 1
ULT = r
ws.conditional_formatting.add(f"E5:E{ULT}", ColorScaleRule(start_type="num", start_value=0, start_color="F8696B",
                                                           mid_type="num", mid_value=0.4, mid_color="FFFFFF",
                                                           end_type="num", end_value=1, end_color="63BE7B"))
ws.conditional_formatting.add(f"I6:I{5 + len(drivers)}", ColorScaleRule(start_type="min", start_color="FFFFFF",
                                                                        end_type="max", end_color="63BE7B"))
# ICP: os drivers viram regra de ação
ws["H26"], ws["H26"].font = "ICP do canal — perfil ideal de contador (drivers → regra de ação)", F_SUB
ws["H27"] = "Regra: carteira com 50+ clientes  E  6+ funcionários  E  ERP de mercado (não Nenhum/Planilha/ContaAzul)."
header(ws, 28, ["Grupo", "Contadores", "% converteu_omie", "Cliente Ativo (status_omie)", "Endereçáveis (regra 1a)",
                "Carteira endereçável"], start_col=8)
for k, (nome, v) in enumerate([("ICP", 1), ("Fora do ICP", 0)]):
    r = 29 + k
    ws.cell(row=r, column=8, value=nome)
    ws.cell(row=r, column=9, value=f'=COUNTIF({rng("R")},{v})')
    ws.cell(row=r, column=10, value=f'=AVERAGEIF({rng("R")},{v},{rng("M")})').number_format = PCT
    ws.cell(row=r, column=11, value=f'=COUNTIFS({rng("R")},{v},{rng("L")},"Cliente Ativo")')
    ws.cell(row=r, column=12, value=f'=COUNTIFS({rng("R")},{v},{rng("S")},1)')
    ws.cell(row=r, column=13, value=f'=SUMIFS({rng("G")},{rng("R")},{v},{rng("S")},1)').number_format = "#,##0"
ws["H31"], ws["J31"] = "Lift do ICP vs média (taxa do ICP ÷ média)", "=J29/$B$3"
ws["H32"], ws["J32"] = "Lift do ICP vs fora do ICP", "=J29/J30"
ws["H33"], ws["J33"] = "% da carteira endereçável no ICP", "=M29/(M29+M30)"
for a, fmt in (("J31", '0.0"x"'), ("J32", '0.0"x"'), ("J33", PCT)):
    ws[a].number_format, ws[a].fill = fmt, FILL_DESTAQUE
ws["H34"] = ("Atenção: no pipeline real o win rate do ICP NÃO é maior (aba Como_Avaliar). "
             "O ICP é uma hipótese a validar com piloto + grupo de controle.")
ws["H34"].font = F_ALERTA

chart = BarChart()
chart.type = "bar"
chart.title = "Amplitude da taxa de conversão por variável (p.p.)"
chart.add_data(Reference(ws, min_col=9, min_row=5, max_row=5 + len(drivers)), titles_from_data=True)
chart.set_categories(Reference(ws, min_col=8, min_row=6, max_row=5 + len(drivers)))
chart.legend = None
chart.height, chart.width = 7.5, 14
ws.add_chart(chart, "O4")
larguras(ws, {"A": 26, "B": 20, "C": 11, "D": 11, "E": 9, "F": 11, "G": 3, "H": 32, "I": 15, "J": 14,
              "K": 14, "L": 14, "M": 14})

# ---------------------------------------------------------------- Leads e Clientes (apoio)
ws_l = wb.create_sheet("Leads")
header(ws_l, 1, ["id_lead", "id_contador", "canal_origem", "status_lead", "mrr_inicial", "tentativas",
                 "porte", "regiao", "converteu_omie", "com desfecho (Ganho/Perdido)", "ganho", "tier (1c)", "icp",
                 "faixa de propensão (D2)", "tempo_ciclo_dias", "data_conversao"])
for i, row in enumerate(leads.itertuples(index=False), start=2):
    for j, v in enumerate(row[:6], start=1):
        val = py(v)
        ws_l.cell(row=i, column=j, value=None if (isinstance(val, float) and val != val) else val)
    ws_l[f"O{i}"] = py(row[6])
    if row[7] == row[7] and row[7] is not None:  # só Ganho tem data_conversao
        ws_l[f"P{i}"] = row[7].to_pydatetime() if hasattr(row[7], "to_pydatetime") else row[7]
        ws_l[f"P{i}"].number_format = "dd/mm/yyyy"
    m = f"MATCH(B{i},Base!$A$2:$A${L},0)"
    ws_l[f"G{i}"] = f"=INDEX(Base!$D$2:$D${L},{m})"
    ws_l[f"H{i}"] = f"=INDEX(Base!$C$2:$C${L},{m})"
    ws_l[f"I{i}"] = f"=INDEX(Base!$M$2:$M${L},{m})"
    ws_l[f"J{i}"] = f'=IF(OR(D{i}="Ganho",D{i}="Perdido"),1,0)'
    ws_l[f"K{i}"] = f'=IF(D{i}="Ganho",1,0)'
    ws_l[f"L{i}"] = f'=IF(INDEX(Base!$U$2:$U${L},{m})="","",INDEX(Base!$U$2:$U${L},{m}))'
    ws_l[f"M{i}"] = f"=INDEX(Base!$R$2:$R${L},{m})"
    ws_l[f"N{i}"] = f"=INDEX(Base!$AB$2:$AB${L},{m})"
LL_ = len(leads) + 1
ws_l.freeze_panes = "A2"
ws_l.auto_filter.ref = f"A1:P{LL_}"

ws_c = wb.create_sheet("Clientes")
header(ws_c, 1, ["id_cliente", "id_contador", "status_cliente", "mrr_inicial", "mrr_atual", "porte", "converteu_omie"])
for i, row in enumerate(clientes.itertuples(index=False), start=2):
    for j, v in enumerate(row, start=1):
        ws_c.cell(row=i, column=j, value=py(v))
    m = f"MATCH(B{i},Base!$A$2:$A${L},0)"
    ws_c[f"F{i}"] = f"=INDEX(Base!$D$2:$D${L},{m})"
    ws_c[f"G{i}"] = f"=INDEX(Base!$M$2:$M${L},{m})"
LC = len(clientes) + 1
ws_c.freeze_panes = "A2"


def lr(col):
    return f"Leads!${col}$2:${col}${LL_}"


def cr(col):
    return f"Clientes!${col}$2:${col}${LC}"


# ---------------------------------------------------------------- Como avaliar (plano de avaliação em 2 etapas)
ws = wb.create_sheet("Como_Avaliar")
ws["A1"], ws["A1"].font = "Como avaliar se a abordagem está funcionando", F_TITULO
ws["A2"] = ("Três perguntas: (1) o score separa quem tem perfil para converter? (target); (2) separa quem fecha de verdade? "
            "(pipeline real); (3) priorizar o topo gera mais receita? (piloto de 90 dias com grupo de controle)")
header(ws, 4, ["Pergunta", "Métrica", "Hoje", "Resposta", "Etapa"])
M_, AB_ = rng("M"), rng("AB")
win_inbound = (f'COUNTIFS({lr("C")},"Inbound",{lr("D")},"Ganho")/'
               f'(COUNTIFS({lr("C")},"Inbound",{lr("D")},"Ganho")+COUNTIFS({lr("C")},"Inbound",{lr("D")},"Perdido"))')
perguntas = [
    ("1. O score separa quem tem perfil para converter?", "Conversão no target (converteu_omie) por faixa: Alta · Média · Baixa",
     '=FIXED(C19*100,0)&"% · "&FIXED(C20*100,0)&"% · "&FIXED(C21*100,0)&"%"',
     '="Sim: o top 30% capta "&FIXED(B9*100,0)&"% dos convertidos"', "1. Base"),
    ("2. O score separa quem fecha de verdade?", "Win rate real (Ganho ÷ Ganho + Perdido) por faixa: Alta · Média · Baixa",
     '=FIXED(G19*100,0)&"% · "&FIXED(G20*100,0)&"% · "&FIXED(G21*100,0)&"%"',
     '="Não: a faixa Alta fecha "&FIXED(G19*100,0)&"%, contra "&FIXED(G22*100,0)&"% das demais"', "1. Base"),
    ("3. Priorizar o topo gera mais receita?", "Piloto de 90 dias × grupo de controle: win rate e nMRR por EC",
     "a medir no piloto",
     f'="Meta: win rate ≥ "&FIXED(({win_inbound})*100,1)&"% (Inbound, melhor canal com 30+ leads encerrados) e acima do controle"',
     "2. Piloto"),
]
for k, (perg, metr, hoje, resp, etapa) in enumerate(perguntas):
    r = 5 + k
    ws.cell(row=r, column=1, value=perg).alignment = WRAP
    ws.cell(row=r, column=2, value=metr).alignment = WRAP
    c = ws.cell(row=r, column=3, value=hoje)
    if hoje.startswith("="):
        c.fill = FILL_DESTAQUE
    ws.cell(row=r, column=4, value=resp).alignment = WRAP
    ws.cell(row=r, column=5, value=etapa)
ws["A9"] = "Captura: % dos convertidos (target) que estão no top 30% do score"
ws["B9"] = f'=SUMIF({AB_},"Alta",{M_})/SUM({M_})'
ws["B9"].number_format = PCT
ws["A10"] = "Nota: ciclo médio do Ganho, faixa Alta × demais (o topo fecha mais rápido, mas não fecha mais)"
ws["B10"] = '=FIXED(H19,0)&" × "&FIXED(H22,0)&" dias"'
ws["A12"] = ("Regra de decisão: se a pergunta 3 der sim (piloto acima do controle), a priorização pelo score vira processo; "
             "senão, o perfil é revisto.")
ws["A12"].font = F_SUB

ws["A14"], ws["A14"].font = "Cortes do score (faixas de propensão dos 600 contadores)", F_SUB
ws["A15"], ws["B15"] = "Faixa Alta: score a partir do percentil 70 (top 30%)", f"=_xlfn.PERCENTILE.INC({rng('AA')},0.7)"
ws["A16"], ws["B16"] = "Faixa Média: a partir do percentil 30", f"=_xlfn.PERCENTILE.INC({rng('AA')},0.3)"
for a in ("B15", "B16"):
    ws[a].number_format = PCT1

header(ws, 18, ["Faixa de propensão", "Contadores", "% converteu_omie (target)", "Lift vs média",
                "Leads com desfecho (Ganho+Perdido)", "Leads Ganhos", "Win rate real", "Ciclo médio do Ganho (dias)"])
for k, fx in enumerate(["Alta", "Média", "Baixa"]):
    r = 19 + k
    ws.cell(row=r, column=1, value=fx)
    ws.cell(row=r, column=2, value=f'=COUNTIF({AB_},A{r})')
    ws.cell(row=r, column=3, value=f'=AVERAGEIF({AB_},A{r},{M_})').number_format = PCT
    ws.cell(row=r, column=4, value=f"=C{r}/Drivers!$B$3").number_format = '0.0"x"'
    ws.cell(row=r, column=5, value=f'=COUNTIFS({lr("N")},A{r},{lr("J")},1)')
    ws.cell(row=r, column=6, value=f'=COUNTIFS({lr("N")},A{r},{lr("K")},1)')
    ws.cell(row=r, column=7, value=f"=IF(E{r}>0,F{r}/E{r},0)").number_format = PCT
    ws.cell(row=r, column=8, value=f'=AVERAGEIFS({lr("O")},{lr("N")},A{r},{lr("K")},1)').number_format = "0"
ws["A22"] = "Demais (Média + Baixa)"
ws["B22"], ws["C22"] = "=B20+B21", "=(C20*B20+C21*B21)/B22"
ws["D22"], ws["E22"], ws["F22"] = "=C22/Drivers!$B$3", "=E20+E21", "=F20+F21"
ws["G22"] = "=IF(E22>0,F22/E22,0)"
ws["H22"] = f'=AVERAGEIFS({lr("O")},{lr("N")},"<>Alta",{lr("K")},1)'
ws["C22"].number_format, ws["D22"].number_format, ws["G22"].number_format, ws["H22"].number_format = PCT, '0.0"x"', PCT, "0"
for c in range(1, 9):
    ws.cell(row=22, column=c).font = Font(bold=True)

ws["A24"], ws["A24"].font = "O target reflete o fechamento real? (por porte)", F_SUB
cab = ["Porte", "Contadores", "% converteu_omie (target)", "Leads com desfecho (Ganho+Perdido)",
       "Leads Ganhos", "Win rate real", "Contratos em clientes_ativos"]
header(ws, 25, cab)
portes = ["Autônomo", "Pequeno Escritório", "Médio Escritório", "Grande Escritório"]
for k, p_ in enumerate(portes + ["Total"]):
    r = 26 + k
    ws.cell(row=r, column=1, value=p_)
    if p_ == "Total":
        for c, f in ((2, "=SUM(B26:B29)"), (3, f"=AVERAGE({M_})"), (4, "=SUM(D26:D29)"), (5, "=SUM(E26:E29)"),
                     (7, "=SUM(G26:G29)")):
            ws.cell(row=r, column=c, value=f)
        for c in range(1, 8):
            ws.cell(row=r, column=c).font = Font(bold=True)
    else:
        ws.cell(row=r, column=2, value=f'=COUNTIF({rng("D")},A{r})')
        ws.cell(row=r, column=3, value=f'=AVERAGEIF({rng("D")},A{r},{M_})')
        ws.cell(row=r, column=4, value=f'=COUNTIFS({lr("G")},A{r},{lr("J")},1)')
        ws.cell(row=r, column=5, value=f'=COUNTIFS({lr("G")},A{r},{lr("K")},1)')
        ws.cell(row=r, column=7, value=f'=COUNTIF({cr("F")},A{r})')
    ws.cell(row=r, column=6, value=f"=IF(D{r}>0,E{r}/D{r},0)")
    for c in (3, 6):
        ws.cell(row=r, column=c).number_format = PCT
ws.row_dimensions[4].height = ws.row_dimensions[18].height = ws.row_dimensions[25].height = 45
ws["A32"], ws["A32"].font = "Leitura", F_SUB
for k, t in enumerate([
    "• Etapa 1: pelo target o score ordena muito bem (a faixa Alta converte ~30x a Baixa), mas o target vaza do "
    "score_potencial_latente, dominado por tamanho e ERP: a separação é alta por construção.",
    "• No pipeline real o topo fecha mais rápido, mas não fecha mais: só a etapa 2 (piloto com controle) prova se "
    "priorizar o score gera receita.",
]):
    ws.cell(row=33 + k, column=1, value=t)
larguras(ws, {"A": 46, "B": 44, "C": 20, "D": 22, "E": 18, "F": 12, "G": 13, "H": 16})

# ---------------------------------------------------------------- Impacto em nMRR
ws = wb.create_sheet("Impacto_nMRR")
ws["A1"], ws["A1"].font = "Extra — nMRR incremental ao priorizar quem tem maior propensão", F_TITULO
ws["A2"] = ("Lógica: parte da capacidade mensal dos ECs vai para o topo da propensão (top 30% do score). O ganho de "
            "conversão vem do lift do topo (taxa do top 30% ÷ média), descontado pelo fator de realização. "
            "O resultado principal é o UPLIFT %.")
header(ws, 4, ["Premissa", "Conservador", "Base", "Otimista", "Fonte / racional"])
prem = [
    ("ECs na rede", n_ecs, None, None, "Soma de num_executivos (aba franquias)"),
    ("Leads trabalhados por EC/mês", round(leads_por_ec, 1), None, None, "Média de leads_trabalhados (execucao_comercial)"),
    ("Taxa atual lead trabalhado → fechamento", round(float(conv_execucao), 4), None, None,
     "Σ fechamentos / Σ leads_trabalhados (execucao_comercial)"),
    ("% da capacidade redirecionada ao topo da propensão", 0.3, 0.5, 0.6, "Decisão de gestão (premissa)"),
    ("Lift do topo da propensão (top 30% ÷ média)", "=Como_Avaliar!$D$19", None, None,
     "Aba Como_Avaliar — descritivo, sem modelo"),
    ("Fator de realização do lift", 0.25, 0.4, 0.6,
     "Fração do lift que vira Ganho real (o target diverge do real — aba Como_Avaliar)"),
    ("Ticket MRR por novo cliente (R$)", round(float(ticket_execucao), 0), f"=AVERAGE({cr('D')})", f"=AVERAGE({cr('D')})",
     "Conservador: nMRR por fechamento da execução. Base/Otimista: mrr_inicial médio (clientes_ativos)"),
    ("Haircut de churn", f'=COUNTIF({cr("C")},"Churned")/COUNTA({cr("A")})', None, None,
     "Contratos Churned / total (vida do contrato) — desconto fixo, conservador"),
    (f"nMRR oficial das {n_abaixo} franquias abaixo da meta (R$/mês)", int(realizado_abaixo), None, None,
     "franquias.realizado_nmrr_mes onde realizado < meta"),
]
FMT = {7: PCT1, 8: PCT, 9: '0.00"x"', 10: PCT, 11: BRL, 12: PCT1, 13: BRL, 14: BRL}
for k, (nome, c1, c2, c3, fonte) in enumerate(prem):
    r = 5 + k
    ws.cell(row=r, column=1, value=nome)
    ws.cell(row=r, column=2, value=c1)
    ws.cell(row=r, column=3, value=c2 if c2 is not None else f"=B{r}")
    ws.cell(row=r, column=4, value=c3 if c3 is not None else f"=B{r}")
    ws.cell(row=r, column=5, value=fonte)
    for c in (2, 3, 4):
        cell = ws.cell(row=r, column=c)
        if r in (8, 10) or (r == 11 and c == 2):
            cell.fill = FILL_INPUT
        cell.number_format = FMT.get(r, "0.0")
header(ws, 16, ["Cálculo", "Conservador", "Base", "Otimista", "Fórmula"])
calc = [
    ("Leads trabalhados/mês (rede)", "={c}5*{c}6", "0", "ECs × leads por EC"),
    ("Leads redirecionados ao topo da propensão", "={c}17*{c}8", "0", "capacidade × % redirecionada"),
    ("Conversão nova nesses leads", "={c}7*(1+({c}9-1)*{c}10)", PCT1, "taxa atual × (1 + (lift − 1) × fator)"),
    ("Fechamentos incrementais/mês", "={c}18*({c}19-{c}7)", "0.0", "leads redirecionados × ganho de conversão"),
    ("nMRR incremental/mês — escala da execução", "={c}20*{c}11*(1-{c}12)", BRL, "fechamentos × ticket × (1 − churn)"),
    ("nMRR atual/mês — mesma escala", "={c}17*{c}7*{c}11*(1-{c}12)", BRL, "leads × conversão × ticket × (1 − churn)"),
    ("UPLIFT no nMRR do canal", "={c}21/{c}22", "+0%", "incremental ÷ atual (independe da escala de receita)"),
    ("nMRR incremental/mês — escala oficial (franquias abaixo da meta)", "={c}23*{c}13", BRL,
     "uplift × realizado oficial dessas franquias"),
]
for k, (nome, f, fmt, expl) in enumerate(calc):
    r = 17 + k
    ws.cell(row=r, column=1, value=nome).alignment = WRAP
    for j, c in enumerate("BCD"):
        cell = ws.cell(row=r, column=2 + j, value=f.format(c=c))
        cell.number_format = fmt
        if r in (23, 24):
            cell.fill = FILL_DESTAQUE
    ws.cell(row=r, column=5, value=expl)
for k, (t, fonte) in enumerate([
    ("Premissas assumidas", F_SUB),
    ("• Capacidade dos ECs constante; o redirecionamento não muda o ticket nem o ciclo; sem custo adicional de aquisição.", None),
    ("• O lift do topo foi medido no target converteu_omie; só uma fração (fator de realização) vira Ganho real.", None),
    ("• Teste de estresse: com fator de realização = 0 (como o pipeline real sugere hoje), o uplift é zero — por isso o piloto.", None),
    ("• O nMRR da execução é 10 a 27x menor que o oficial das franquias; por isso o uplift % é o número principal "
     "(franquias.realizado_nmrr_mes contra execucao_comercial.nmrr_gerado).", None),
]):
    c = ws.cell(row=27 + k, column=1, value=t)
    if fonte:
        c.font = fonte
ws.row_dimensions[24].height = 30
larguras(ws, {"A": 44, "B": 15, "C": 15, "D": 15, "E": 70})

# ---------------------------------------------------------------- Previsibilidade de receita (próximo trimestre)
ws = wb.create_sheet("Previsao_Receita")
ws["A1"], ws["A1"].font = "Previsibilidade de receita: quanto o pipeline atual deve gerar no próximo trimestre", F_TITULO
ws["A2"] = ("Dois métodos independentes, só com dados da base: (1) pipeline aberto ponderado pela chance de fechar de cada "
            "estágio; (2) run-rate histórico de Ganhos do CRM. Janela de 90 dias: nenhum Ganho da base fechou depois de 89 dias.")
ws["A4"], ws["A4"].font = "1. Chance de fechar por estágio (funil de execucao_comercial, 2024)", F_SUB
header(ws, 5, ["Etapa do funil", "Volume", "Taxa da etapa"])
for k, (nome, val, taxa) in enumerate([("Leads trabalhados", f_leads, None), ("Reuniões realizadas", f_reun, "Lead → reunião"),
                                       ("Propostas enviadas", f_prop, "Reunião → proposta"),
                                       ("Fechamentos", f_fech, "Proposta → fechamento")]):
    r = 6 + k
    ws.cell(row=r, column=1, value=nome)
    ws.cell(row=r, column=2, value=int(val)).number_format = "#,##0"
    if taxa:
        ws.cell(row=r, column=3, value=f"=B{r}/B{r - 1}").number_format = PCT1
        ws.cell(row=r, column=4, value=taxa)

ws["A11"], ws["A11"].font = "2. Método 1: pipeline aberto ponderado (leads abertos há até 90 dias)", F_SUB
header(ws, 12, ["Estágio do lead", "Leads abertos até 90 dias", "Chance de fechar", "Ganhos esperados",
                "Leads parados (> 90 dias): chance 0", "De onde vem a chance"])
estagios = [("Proposta Enviada", "=C9", "Proposta → fechamento"),
            ("Reunião Agendada", "=C8*C9", "Reunião → proposta → fechamento"),
            ("Em Contato", "=C7*C8*C9", "Lead → reunião → proposta → fechamento"),
            ("Novo", "=C7*C8*C9", "Lead → reunião → proposta → fechamento"),
            ("Sem Resposta", 0, "Fora do funil: o contador não respondeu")]
for k, (est, chance, origem) in enumerate(estagios):
    r = 13 + k
    ws.cell(row=r, column=1, value=est)
    ws.cell(row=r, column=2, value=f'=COUNTIFS({lr("D")},A{r},{lr("O")},"<=90")')
    ws.cell(row=r, column=3, value=chance).number_format = PCT1
    ws.cell(row=r, column=4, value=f"=B{r}*C{r}").number_format = "0.0"
    ws.cell(row=r, column=5, value=f'=COUNTIFS({lr("D")},A{r},{lr("O")},">90")')
    ws.cell(row=r, column=6, value=origem)
ws["A18"], ws["A18"].font = "Total", F_SUB
for col in "BDE":
    ws[f"{col}18"] = f"=SUM({col}13:{col}17)"
ws["D18"].number_format = "0.0"
ws["D18"].fill = FILL_DESTAQUE
for k, (nome, f, fmt) in enumerate([
    ("Ticket médio do Ganho (R$)", f'=AVERAGEIF({lr("D")},"Ganho",{lr("E")})', BRL),
    ("Churn dos contratos", f'=COUNTIF({cr("C")},"Churned")/COUNTA({cr("A")})', PCT1),
    ("nMRR esperado do pipeline no trimestre (R$, escala do CRM)", "=D18*B20*(1-B21)", BRL)]):
    r = 20 + k
    ws.cell(row=r, column=1, value=nome)
    c = ws.cell(row=r, column=2, value=f)
    c.number_format = fmt
ws["B22"].fill = FILL_DESTAQUE

ws["A24"], ws["A24"].font = "3. Método 2: run-rate histórico de Ganhos (CRM)", F_SUB
for k, (nome, f, fmt) in enumerate([
    ("Primeiro Ganho", f"=MIN({lr('P')})", "dd/mm/yyyy"),
    ("Último Ganho", f"=MAX({lr('P')})", "dd/mm/yyyy"),
    ("Meses de histórico", "=(YEAR(B26)-YEAR(B25))*12+MONTH(B26)-MONTH(B25)+1", "0"),
    ("Ganhos no período", f'=COUNTIF({lr("D")},"Ganho")', "0"),
    ("Ganhos por trimestre (run-rate)", "=B28/B27*3", "0.0"),
    ("Menor e maior trimestre fechado", '=MIN(D33:D40)&" a "&MAX(D33:D40)', None)]):
    r = 25 + k
    ws.cell(row=r, column=1, value=nome)
    c = ws.cell(row=r, column=2, value=f)
    if fmt:
        c.number_format = fmt
ws["B29"].fill = FILL_DESTAQUE
header(ws, 32, ["Trimestre", "Início", "Fim", "Ganhos"])
tri = [(a, t) for a in (2024, 2025) for t in (1, 2, 3, 4)] + [(2026, 1)]
for k, (ano, t) in enumerate(tri):
    r = 33 + k
    ws.cell(row=r, column=1, value=f"{ano} T{t}" + (" (parcial, até 07/03)" if (ano, t) == (2026, 1) else ""))
    ws.cell(row=r, column=2, value=f"=DATE({ano},{3 * t - 2},1)").number_format = "dd/mm/yyyy"
    ws.cell(row=r, column=3, value=f"=EOMONTH(B{r},2)").number_format = "dd/mm/yyyy"
    ws.cell(row=r, column=4, value=f'=COUNTIFS({lr("D")},"Ganho",{lr("P")},">="&B{r},{lr("P")},"<="&C{r})')

ws["A44"], ws["A44"].font = "4. Lacuna de previsibilidade e alavanca", F_SUB
for k, (nome, f, fmt) in enumerate([
    ("nMRR de Ganhos do CRM por mês (R$)", f'=SUMIF({lr("D")},"Ganho",{lr("E")})/B27', BRL),
    ("nMRR oficial da rede por mês (R$, aba franquias)", int(realizado_rede), BRL),
    ("CRM ÷ oficial", "=B45/B46", PCT1),
    ("Ganhos previstos no trimestre com a priorização (cenário base do Impacto_nMRR)", "=D18*(1+Impacto_nMRR!C23)", "0.0")]):
    r = 45 + k
    ws.cell(row=r, column=1, value=nome)
    c = ws.cell(row=r, column=2, value=f)
    c.number_format = fmt
ws["B47"].fill = FILL_DESTAQUE
ws["A50"], ws["A50"].font = "Leitura", F_SUB
for k, t in enumerate([
    "• Os dois métodos convergem (pipeline ponderado × run-rate): o pipeline do CRM é previsível no curto prazo.",
    "• Os leads parados há mais de 90 dias entram com chance zero: o pipeline nominal superestima o que vai fechar.",
    "• A chance de fechar por estágio vem de execucao_comercial (outra fonte) aplicada aos leads do CRM: é a melhor evidência "
    "da base, e a convergência com o run-rate mostra que funciona.",
    "• Lacuna: os Ganhos do CRM explicam uma fração mínima do nMRR oficial. Sem uma fonte única de nMRR, dá para prever o "
    "pipeline do CRM, mas não a receita oficial do canal.",
]):
    ws.cell(row=51 + k, column=1, value=t)
larguras(ws, {"A": 62, "B": 22, "C": 16, "D": 16, "E": 22, "F": 40})

# ---------------------------------------------------------------- Respostas (as 4 perguntas do Desafio 2)
ws = ws_resp
RK = "Drivers!$J$6:$J$14"
top5 = {1: "Médio e Grande convertem 94–100%; Autônomo 7%",
        2: "de ~9% (1 a 2 funcionários) a ~97% (16+)",
        3: "de ~10% (< 20 clientes) a ~98% (150+); é também o Fit do 1c e do Desafio 3",
        4: "ERP de mercado converte muito mais que planilha, nenhum ou ContaAzul",
        5: "16+ anos convertem mais que o dobro dos com < 8 anos"}
linhas = [
    ("Desafio 2 — Propensão à conversão: respostas", F_TITULO, None),
    ("Abordagem descritiva (o case aceita): a propensão de cada contador é a taxa histórica de conversão de quem tem "
     "o mesmo perfil — média das taxas das categorias dele nos 5 drivers do top 5. Sem modelo e sem pesos; "
     "a lista completa está na aba Base (coluna AA).", None, None),
    ("", None, None),
    ("1. Lógica — quais variáveis e por quê", F_SUB, None),
    ("• Testadas 9 características. Entram as 5 que mais separam quem converte de quem não converte "
     "(amplitude = maior taxa − menor taxa, categorias com 30+ contadores). Região, segmento, site e LinkedIn "
     "quase não separam.", None, None),
    ("• Fora: score_potencial_latente (regra do case) e status_omie (não tem sinal e diverge dos contratos).", None, None),
    ("Endereçáveis (303) na faixa Alta de propensão", None, (f'=COUNTIFS(Base!$AB$2:$AB${L},"Alta",Base!$S$2:$S${L},1)', "0")),
    ("• O topo da propensão já foi quase todo convertido: entre os 303 endereçáveis, poucos estão na faixa Alta.",
     None, None),
    ("", None, None),
    ("2. Como avaliar se está funcionando: três perguntas", F_SUB, None),
    ("Pergunta 1 · O score separa quem tem perfil para converter? Conversão no target, Alta · Média · Baixa", None, ("=Como_Avaliar!C5", None)),
    ("Pergunta 2 · O score separa quem fecha de verdade? Win rate real, Alta · Média · Baixa", None, ("=Como_Avaliar!C6", None)),
    ("Pergunta 3 · Priorizar o topo gera mais receita? Piloto de 90 dias × grupo de controle", None, ("=Como_Avaliar!D7", None)),
    ("Nota: ciclo médio do Ganho, faixa Alta × demais (fecha mais rápido, não mais)", None, ("=Como_Avaliar!B10", None)),
    ("• Sim na pergunta 1, não na pergunta 2: o score mede potencial, não fechamento. A pergunta 3 (piloto) decide: "
     "se superar o controle, a priorização vira processo; senão, o perfil é revisto.", None, None),
    ("", None, None),
    ("3. Top 5 características (amplitude da taxa de conversão)", F_SUB, None),
]
for k in range(1, 6):
    linhas.append((None, None, ("top5", k)))
linhas += [
    ("• Porte, funcionários e carteira medem o mesmo conceito (tamanho). Com o ERP, descrevem o perfil do topo "
     "(ICP: 50+ clientes, 6+ funcionários e ERP de mercado).", None, None),
    ("", None, None),
    ("4. Extra — impacto em nMRR ao priorizar quem tem maior propensão", F_SUB, None),
    ("Uplift no nMRR do canal — cenário base", None, ("=Impacto_nMRR!C23", "+0%")),
    ("Faixa (conservador a otimista)", None,
     ('=FIXED(Impacto_nMRR!B23*100,0)&"% a "&FIXED(Impacto_nMRR!D23*100,0)&"%"', None)),
    ("nMRR incremental/mês nas franquias abaixo da meta (escala oficial, cenário base)", None, ("=Impacto_nMRR!C24", BRL)),
    ("• Premissas: 50% da capacidade dos ECs vai para o top 30% do score; 40% do lift do topo vira Ganho real "
     "(fator de realização, hipótese); ticket = mrr_inicial médio; desconto de churn de 20,5%.", None, None),
    ("• Teste de estresse: com fator de realização 0, o uplift é 0. O número só se confirma com o piloto.", None, None),
    ("", None, None),
    ("5. Previsibilidade de receita: o que o pipeline atual deve gerar no próximo trimestre", F_SUB, None),
    ("Método 1: pipeline aberto ponderado pela chance de fechar de cada estágio (Ganhos)", None, ("=Previsao_Receita!D18", "0.0")),
    ("Método 2: run-rate histórico do CRM (Ganhos por trimestre)", None, ("=Previsao_Receita!B29", "0.0")),
    ("nMRR esperado do pipeline no trimestre (R$, escala do CRM)", None, ("=Previsao_Receita!B22", BRL)),
    ("Ganhos do CRM ÷ nMRR oficial do canal", None, ("=Previsao_Receita!B47", PCT1)),
    ("• Os dois métodos convergem: o pipeline do CRM é previsível. Mas ele explica uma fração mínima do nMRR oficial: "
     "sem fonte única de nMRR, não dá para prever a receita oficial do canal.", None, None),
]
for i, (txt, fonte, val) in enumerate(linhas, start=1):
    if val and val[0] == "top5":
        k = val[1]
        m_ = f"MATCH({k},{RK},0)"
        ws.cell(row=i, column=1, value=(f'="{k}. "&INDEX(Drivers!$H$6:$H$14,{m_})&" — "&'
                                        f'FIXED(INDEX(Drivers!$I$6:$I$14,{m_}),0)&" p.p.: {top5[k]}"'))
        continue
    c = ws.cell(row=i, column=1, value=txt)
    c.alignment = WRAP
    if fonte:
        c.font = fonte
    if val:
        cv = ws.cell(row=i, column=2, value=val[0])
        cv.fill = FILL_DESTAQUE
        if val[1]:
            cv.number_format = val[1]
larguras(ws, {"A": 120, "B": 16})

wb.move_sheet("Previsao_Receita", offset=-(len(wb.sheetnames) - 1 - wb.sheetnames.index("Impacto_nMRR")) + 1)
wb.move_sheet("Base", offset=10)
wb.move_sheet("Leads", offset=10)
wb.move_sheet("Clientes", offset=10)
OUT.parent.mkdir(exist_ok=True)
wb.save(OUT)
print("Gerado:", OUT)
