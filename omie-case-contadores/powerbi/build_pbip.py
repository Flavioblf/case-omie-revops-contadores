"""Gera o modelo semântico (TMDL), o relatório (PBIR) e a documentação do 'Painel Omie.pbip' - Desafios 1, 2 e 3.

Uma página por pergunta do case: pergunta → resposta (números de deck/data.json) → evidência.
Fonte: SQL Server local, banco OMIE (tabelas no esquema raw; views dbo.vw_* criadas por
sql/sqlserver/anexo_analises_complementares.sql). Visual no design system Omie (deck/design_system).

Uso (com o Power BI Desktop FECHADO): python powerbi/build_pbip.py
"""
import hashlib
import json
import os
import shutil
import subprocess
import time
import uuid
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]           # omie-case-contadores
PASTA = ROOT.parent                                   # Omie_Antigravity (onde está o .pbip)
SM = PASTA / "Painel Omie.SemanticModel" / "definition"
RP = PASTA / "Painel Omie.Report"
DOC = PASTA / "Painel Omie - Documentação.md"
DS = ROOT / "deck" / "design_system"
TOKENS = json.loads((DS / "tokens.json").read_text(encoding="utf-8"))
HEX = {k: "#" + v["hex"] for g in ("marca", "neutras", "feedback") for k, v in TOKENS["cores"][g].items()}
FONTE_T = TOKENS["tipografia"]["titulos"]["fonte"]   # Century Gothic
LOGO = "logo_omie_escuro.png"

guid = lambda: str(uuid.uuid4())  # noqa: E731


def vid(chave: str | None = None) -> str:
    """Nome de página/visual (20 hex). Com chave é determinístico: regerar sobrescreve em vez de recriar pastas."""
    return hashlib.md5(chave.encode("utf-8")).hexdigest()[:20] if chave else uuid.uuid4().hex[:20]


def remover(item: Path, tentativas: int = 3) -> bool:
    """Apaga arquivo/pasta. Pastas do OneDrive às vezes recusam o rmtree do Python: cai para o Remove-Item."""
    for _ in range(tentativas):
        try:
            shutil.rmtree(item) if item.is_dir() else item.unlink()
            return True
        except (PermissionError, OSError):
            subprocess.run(["powershell", "-NoProfile", "-Command",
                            f"Remove-Item -LiteralPath '{item}' -Recurse -Force -ErrorAction SilentlyContinue"],
                           capture_output=True)
            if not item.exists():
                return True
            time.sleep(1.5)
    print(f"AVISO: não consegui apagar {item} (OneDrive). Apague manualmente.")
    return False


def limpar(pasta: Path, manter: set[str] = frozenset()) -> None:
    pasta.mkdir(parents=True, exist_ok=True)
    for item in pasta.iterdir():
        if item.name not in manter:
            remover(item)


def q(nome: str) -> str:
    """Nome TMDL: entre aspas simples quando tem espaço ou caractere especial."""
    return nome if nome.replace("_", "").isalnum() and nome.isascii() else "'" + nome.replace("'", "''") + "'"


# =============================================================== modelo semântico
M_RAW = """let
    Fonte = Sql.Database(ServidorSQL, BancoSQL),
    Dados = Fonte{{[Schema="raw", Item="{t}"]}}[Data],
    SemLinhaOrigem = Table.RemoveColumns(Dados, {{"linha_origem"}})
in
    SemLinhaOrigem"""
# mercado_contadores ganha a ordem do porte já no Power Query: ordenar 'porte' por uma coluna DAX derivada
# dele mesmo gera dependência circular no modelo
M_MERCADO = """let
    Fonte = Sql.Database(ServidorSQL, BancoSQL),
    Dados = Fonte{[Schema="raw", Item="mercado_contadores"]}[Data],
    SemLinhaOrigem = Table.RemoveColumns(Dados, {"linha_origem"}),
    OrdemPorte = Table.AddColumn(SemLinhaOrigem, "porte_ordem", each
        if [porte] = "Autônomo" then 1 else if [porte] = "Pequeno Escritório" then 2
        else if [porte] = "Médio Escritório" then 3 else 4, Int64.Type)
in
    OrdemPorte"""
M_TIERS = """let
    Fonte = Sql.Database(ServidorSQL, BancoSQL),
    Dados = Fonte{[Schema="dbo", Item="vw_tiers_prospects"]}[Data],
    Colunas = Table.SelectColumns(Dados, {"id_contador", "nivel_fit", "fit_carteira", "nivel_engajamento",
        "engajamento", "score_prioridade", "tier_prioridade", "estagio_aberto"})
in
    Colunas"""
# Desafio 1a e 1b: as views do anexo SQL têm o mesmo texto das consultas da resposta
M_VIEW = """let
    Fonte = Sql.Database(ServidorSQL, BancoSQL),
    Dados = Fonte{{[Schema="dbo", Item="{v}"]}}[Data]
in
    Dados"""

S, I, D, DT = "string", "int64", "double", "dateTime"
COLS = {
    "mercado_contadores": [("id_contador", S), ("estado", S), ("regiao", S), ("cidade", S), ("porte", S),
                           ("segmento_principal", S), ("anos_mercado", I), ("num_clientes_ativos", I),
                           ("num_funcionarios", I), ("usa_erp", S), ("erp_atual", S), ("tem_site", S),
                           ("presenca_linkedin", S), ("status_omie", S), ("converteu_omie", I),
                           ("score_potencial_latente", D), ("porte_ordem", I)],
    "leads_pipeline": [("id_lead", S), ("id_contador", S), ("canal_origem", S), ("id_franquia", S),
                       ("data_entrada", DT), ("status_lead", S), ("data_conversao", DT), ("mrr_inicial", D),
                       ("motivo_perda", S), ("num_tentativas_contato", I), ("tempo_ciclo_dias", I)],
    "franquias": [("id_franquia", S), ("nome_franquia", S), ("estado", S), ("regiao", S), ("meses_operando", I),
                  ("meta_nmrr_mes", I), ("realizado_nmrr_mes", I), ("num_executivos", I), ("atingimento_pct", D),
                  ("data_inicio", DT)],
    "tiers_1c": [("id_contador", S), ("nivel_fit", I), ("fit_carteira", S), ("nivel_engajamento", I),
                 ("engajamento", S), ("score_prioridade", I), ("tier_prioridade", S), ("estagio_aberto", S)],
    "propensao_d2": [("id_contador", S), ("propensao", D), ("faixa_propensao", S), ("faixa_ordem", I)],
    "d1a_mercado": [("ordem_linha", I), ("regiao", S), ("porte", S), ("num_contadores", I),
                    ("participacao_regiao_pct", D), ("participacao_total_pct", D), ("e_total_brasil", I),
                    ("e_total_regiao", I)],
    "d1b_conversao": [("ordem_bloco", I), ("canal_origem", S), ("nome_franquia", S), ("num_leads", I),
                      ("num_ganhos", I), ("num_tentativas_contato", I), ("taxa_conv_lead_ganho_pct", D),
                      ("tentativas_contato_por_ganho", D), ("ranking_roi_esforco", I), ("alerta_amostra", S)],
}
OCULTAS = {"score_potencial_latente", "porte_ordem", "nivel_fit", "nivel_engajamento", "e_total_brasil",
           "e_total_regiao", "ordem_bloco", "faixa_ordem"}  # latente: proibida pelo case; as demais são técnicas

MC = "mercado_contadores"
CALC_MERCADO = [  # (nome, dataType, formatString, DAX, sortBy, oculta, descrição)
    ("sem_erp", I, "0", f'IF({MC}[erp_atual] IN {{"Nenhum", "Planilha", "ContaAzul"}}, 1, 0)', None, True,
     "1 = usa planilha, nenhum sistema ou ContaAzul (concorrente direto)"),
    ("uso_erp", S, None, f'IF({MC}[sem_erp] = 1, "Sem ERP de mercado", "ERP de mercado")', None, False,
     "ERP de mercado (Domínio, Questor, Alterdata, Totvs, SAP) × sem ERP (planilha, nenhum, ContaAzul)"),
    ("icp", S, None, f'IF({MC}[num_clientes_ativos] >= 50 && {MC}[num_funcionarios] >= 6 && {MC}[sem_erp] = 0, "Sim", "Não")',
     None, False, "Perfil ideal: 50+ clientes, 6+ funcionários e ERP de mercado"),
    ("enderecavel", S, None, f'IF({MC}[status_omie] <> "Cliente Ativo" && {MC}[converteu_omie] = 0, "Sim", "Não")',
     None, False, "Regra do Desafio 1a: não é Cliente Ativo e ainda não converteu (303 contadores)"),
    ("faixa_carteira", S, None, "\n".join([
        f'SWITCH(TRUE(), {MC}[num_clientes_ativos] < 20, "1. < 20", {MC}[num_clientes_ativos] < 50, "2. 20 a 49",',
        f'    {MC}[num_clientes_ativos] < 150, "3. 50 a 149", "4. 150+")']), None, False, "Faixa do nº de clientes"),
    ("faixa_funcionarios", S, None, "\n".join([
        f'SWITCH(TRUE(), {MC}[num_funcionarios] <= 2, "1. 1 a 2", {MC}[num_funcionarios] <= 5, "2. 3 a 5",',
        f'    {MC}[num_funcionarios] <= 15, "3. 6 a 15", "4. 16+")']), None, False, "Faixa do nº de funcionários"),
    ("faixa_anos", S, None,
     f'SWITCH(TRUE(), {MC}[anos_mercado] < 8, "1. < 8 anos", {MC}[anos_mercado] < 16, "2. 8 a 15 anos", "3. 16+ anos")',
     None, False, "Faixa de anos de mercado"),
    ("score_prioridade", I, "0",
     f"LOOKUPVALUE(tiers_1c[score_prioridade], tiers_1c[id_contador], {MC}[id_contador])", None, False,
     "Score do Desafio 1c: nível do fit (1 a 3) + nível do engajamento (0 a 2); vazio para quem não é endereçável"),
    ("fit_ordem", I, "0",
     f"LOOKUPVALUE(tiers_1c[nivel_fit], tiers_1c[id_contador], {MC}[id_contador])", None, True,
     "Ordem do fit (3 = Alto)"),
    ("fit_carteira", S, None,
     f"LOOKUPVALUE(tiers_1c[fit_carteira], tiers_1c[id_contador], {MC}[id_contador])", "fit_ordem", False,
     "Fit do 1c pela carteira: Alto > 24 clientes, Médio 15 a 24, Baixo < 15"),
    ("engajamento", S, None,
     f"LOOKUPVALUE(tiers_1c[engajamento], tiers_1c[id_contador], {MC}[id_contador])", None, False,
     "Engajamento do 1c: Quente = Reunião/Proposta, Morno = Novo/Em Contato (lead aberto até 90 dias), Frio"),
    ("tier_ordem", I, "0", "\n".join([
        f"VAR t = LOOKUPVALUE(tiers_1c[tier_prioridade], tiers_1c[id_contador], {MC}[id_contador])",
        'RETURN SWITCH(t, "Alta", 1, "Média", 2, "Baixa", 3)']), None, True, "Ordem do tier"),
    ("tier_prioridade", S, None, f'SWITCH({MC}[tier_ordem], 1, "Alta", 2, "Média", 3, "Baixa")', "tier_ordem", False,
     "Tier do Desafio 1c pela matriz Fit x Engajamento: score >= 4 Alta, 3 Média, <= 2 Baixa"),
    ("propensao", D, "0.0%",
     f"LOOKUPVALUE(propensao_d2[propensao], propensao_d2[id_contador], {MC}[id_contador])", None, False,
     "Score de propensão do Desafio 2: média das taxas de conversão do perfil do contador nos 5 drivers do top 5"),
    ("faixa_prop_ordem", I, "0",
     f"LOOKUPVALUE(propensao_d2[faixa_ordem], propensao_d2[id_contador], {MC}[id_contador])", None, True,
     "Ordem da faixa de propensão (1 = Alta)"),
    ("faixa_propensao", S, None,
     f"LOOKUPVALUE(propensao_d2[faixa_propensao], propensao_d2[id_contador], {MC}[id_contador])", "faixa_prop_ordem", False,
     "Faixa de propensão: Alta = top 30% do score, Média = do percentil 30 ao 70, Baixa = demais"),
    ("situacao_pipeline", S, None,
     f'COALESCE(LOOKUPVALUE(tiers_1c[estagio_aberto], tiers_1c[id_contador], {MC}[id_contador]), "-")', None, False,
     "Estágio do lead ativo (até 90 dias) do contador"),
]
CALC_LEADS = [
    ("ganho", I, "0", 'IF(leads_pipeline[status_lead] = "Ganho", 1, 0)', None, True, "1 = lead Ganho"),
    ("trimestre_ganho", S, None,
     'IF(leads_pipeline[status_lead] = "Ganho", FORMAT(YEAR(leads_pipeline[data_conversao]), "0") & " T" & '
     'FORMAT(QUARTER(leads_pipeline[data_conversao]), "0"))', None, False, "Trimestre do Ganho (data_conversao)"),
    ("com_desfecho", I, "0", 'IF(leads_pipeline[status_lead] IN {"Ganho", "Perdido"}, 1, 0)', None, True,
     "1 = lead encerrado (Ganho ou Perdido)"),
]
VARIAVEIS = [  # (ordem, nome exibido, coluna de mercado_contadores)
    (1, "Porte", "porte"), (2, "Nº de funcionários", "faixa_funcionarios"), (3, "Nº de clientes da carteira", "faixa_carteira"),
    (4, "ERP atual", "erp_atual"), (5, "Anos de mercado", "faixa_anos"), (6, "Região", "regiao"),
    (7, "Segmento", "segmento_principal"), (8, "Tem site", "tem_site"), (9, "LinkedIn", "presenca_linkedin")]
VAR_M = "{" + ", ".join("{" + f'{o}, "{n}"' + "}" for o, n, _ in VARIAVEIS) + "}"
CEN_M = '{{1, "Conservador", 0.3, 0.25}, {2, "Base", 0.5, 0.4}, {3, "Otimista", 0.6, 0.6}}'
PARAM_TABLES = {
    "Param_Variaveis": (
        [("ordem", I), ("variavel", S)],
        f"let\n    Fonte = #table(type table [ordem = Int64.Type, variavel = text],\n        {VAR_M})\nin\n    Fonte",
        "Características analisadas como drivers de conversão", {"variavel": "ordem"}),
    "Param_Cenarios": (
        [("ordem", I), ("cenario", S), ("pct_redirecionado", D), ("fator_realizacao", D)],
        "let\n    Fonte = #table(\n        type table [ordem = Int64.Type, cenario = text, pct_redirecionado = number, fator_realizacao = number],\n"
        f"        {CEN_M})\nin\n    Fonte",
        "Premissas dos cenários de impacto (iguais à aba Impacto_nMRR do Excel)", {"cenario": "ordem"}),
    "Param_Estagios": (
        [("ordem", I), ("estagio", S), ("etapas_ate_fechar", I)],
        "let\n    Fonte = #table(type table [ordem = Int64.Type, estagio = text, etapas_ate_fechar = Int64.Type],\n"
        '        {{1, "Proposta Enviada", 1}, {2, "Reunião Agendada", 2}, {3, "Em Contato", 3}, {4, "Novo", 3}, '
        '{5, "Sem Resposta", 0}})\nin\n    Fonte',
        "Estágios do lead aberto e quantas etapas do funil faltam para fechar (Sem Resposta: fora do funil)",
        {"estagio": "ordem"}),
}


# =============================================================== Desafio 3 (página 4): capacity e EC adicional
# Mesmas regras de src/build_excel_capacity.py: território, MRR esperado, ondas e ranking do EC.
COLS["clientes_ativos"] = [("id_cliente", S), ("id_lead", S), ("id_contador", S), ("data_ativacao", DT), ("plano", S),
                           ("mrr_inicial", D), ("mrr_atual", D), ("status_cliente", S), ("data_churn", DT),
                           ("meses_ativo", I), ("nps_score", I), ("upsell_realizado", S)]
COLS["execucao_comercial"] = [("id_executivo", S), ("id_franquia", S), ("estado", S), ("regiao", S), ("mes_ref", S),
                              ("leads_trabalhados", I), ("reunioes_realizadas", I), ("propostas_enviadas", I),
                              ("fechamentos", I), ("nmrr_gerado", D), ("taxa_conv_lead_reuniao", D),
                              ("taxa_conv_reuniao_proposta", D), ("taxa_conv_proposta_fechamento", D)]
CALC_CLIENTES = [
    ("porte", S, None, f"RELATED({MC}[porte])", None, False, "Porte do contador (para o ticket por porte)"),
]
CHURN_DAX = ('DIVIDE(COUNTROWS(FILTER(ALL(clientes_ativos), clientes_ativos[status_cliente] = "Churned")), '
             "COUNTROWS(ALL(clientes_ativos)))")
EXEC_TAXA = ("CALCULATE(DIVIDE(SUM(execucao_comercial[{a}]), SUM(execucao_comercial[{b}])), "
             "REMOVEFILTERS(execucao_comercial))")  # taxa de uma etapa do funil da rede
CALC_MERCADO += [
    ("franquia_responsavel", S, None, "\n".join([
        f'IF({MC}[estado] = "SP",',
        f'    IF({MC}[cidade] = "São Paulo", "FRQ001", "FRQ002"),',
        f"    LOOKUPVALUE(Param_Territorio[id_franquia], Param_Territorio[estado], {MC}[estado]))"]), None, False,
     "Franquia do estado do contador (SP: cidade de São Paulo = SP Capital, demais = SP Interior); vazio = sem franquia"),
    ("franquia_nome", S, None,
     f'COALESCE(SUBSTITUTE(LOOKUPVALUE(franquias[nome_franquia], franquias[id_franquia], {MC}[franquia_responsavel]), "Franquia ", ""), "Sem franquia")',
     None, False, "Nome da franquia responsável pelo contador"),
    ("cobertura", S, None, f'IF(ISBLANK({MC}[franquia_responsavel]), "Sem franquia", "Própria")',
     None, False, "Própria (estado com franquia) ou Sem franquia"),
    ("ticket_porte", D, "#,0", "\n".join([
        f"VAR p = {MC}[porte]",
        "RETURN ROUND(AVERAGEX(FILTER(ALL(clientes_ativos), clientes_ativos[porte] = p), clientes_ativos[mrr_inicial]), 0)"]),
     None, True, "Ticket MRR médio do porte (mrr_inicial em clientes_ativos)"),
    ("mrr_esperado", D, "#,0",
     f'IF({MC}[enderecavel] = "Sim", {MC}[ticket_porte] * (1 - {CHURN_DAX}))',
     None, False, "Ticket do porte × (1 − churn): MRR esperado de um contador endereçável"),
    ("estagio_ordem", I, "0",
     f'SWITCH({MC}[situacao_pipeline], "Proposta Enviada", 4, "Reunião Agendada", 3, "Em Contato", 2, "Novo", 1, 0)',
     None, True, "Estágio do lead ativo em número (4 = proposta)"),
    ("lead_parado", I, "0", "\n".join([
        "IF(COUNTROWS(FILTER(RELATEDTABLE(leads_pipeline),",
        '    NOT leads_pipeline[status_lead] IN {"Ganho", "Perdido"} && leads_pipeline[tempo_ciclo_dias] > 90)) > 0, 1, 0)']),
     None, True, "1 = tem lead aberto há mais de 90 dias"),
    ("onda", I, "0", "\n".join([
        f'IF({MC}[enderecavel] = "Sim",',
        f'    SWITCH(TRUE(), {MC}[estagio_ordem] >= 3, 1, {MC}[tier_prioridade] = "Alta", 2, {MC}[lead_parado] = 1, 3,',
        f'        {MC}[tier_prioridade] = "Média", 4, 5))']), None, True, "Onda de priorização (1 a 5)"),
    ("onda_nome", S, None, "\n".join([
        f'SWITCH({MC}[onda], 1, "1. Fechar (Reunião/Proposta)", 2, "2. Prospectar tier Alta",',
        '    3, "3. Requalificar lead parado", 4, "4. Nutrir (marketing)", 5, "5. Régua automatizada")']),
     "onda", False, "Onda de priorização do Desafio 3b"),
]
CALC_FRANQUIAS = [
    ("atingimento", D, "0%", "DIVIDE(franquias[realizado_nmrr_mes], franquias[meta_nmrr_mes])", None, False,
     "Realizado ÷ meta de nMRR"),
    ("gap", D, "#,0", "MAX(0, franquias[meta_nmrr_mes] - franquias[realizado_nmrr_mes])", None, False,
     "Meta − realizado (R$/mês), zero para quem bateu a meta"),
    ("realizado_por_ec", D, "#,0", "DIVIDE(franquias[realizado_nmrr_mes], franquias[num_executivos])", None, False,
     "Realizado de nMRR por EC"),
    ("situacao", S, None, 'IF(franquias[atingimento] >= 1, "Na meta", "Abaixo da meta")', None, False,
     "Na meta ou abaixo da meta"),
    ("ordem_atingimento", I, "0",
     "RANKX(ALLNOBLANKROW(franquias), franquias[atingimento] + VALUE(RIGHT(franquias[id_franquia], 3)) / 1000000, , ASC, DENSE)",
     None, True, "Ordem do menor para o maior atingimento (desempate pelo id)"),
    ("rotulo", S, None,
     'SUBSTITUTE(franquias[nome_franquia], "Franquia ", "") & " · " & franquias[num_executivos] & " EC · " & FORMAT(franquias[atingimento], "0%")',
     "ordem_atingimento", False, "Franquia, nº de ECs e atingimento, ordenada pelo atingimento"),
]
TERRITORIO = {}
FRQ_ESTADO = {"RJ": "FRQ003", "MG": "FRQ004", "PR": "FRQ005", "SC": "FRQ006", "RS": "FRQ007", "DF": "FRQ008",
              "GO": "FRQ009", "BA": "FRQ010", "PE": "FRQ011", "CE": "FRQ012", "AM": "FRQ013", "PA": "FRQ014", "ES": "FRQ015"}
TERRITORIO.update({uf: (f, "Própria") for uf, f in FRQ_ESTADO.items()})
TER_M = "{" + ", ".join("{" + f'"{uf}", "{f}", "{tp}"' + "}" for uf, (f, tp) in sorted(TERRITORIO.items())) + "}"
_D = json.loads((ROOT / "deck" / "data.json").read_text(encoding="utf-8"))["drivers"]
_pct = lambda v: f"{v * 100:.0f}%"  # noqa: E731
AVALIACAO = [
    (1, "Conversão por faixa (Alta · Média · Baixa)", "Se o score ordena quem converte (target)",
     _D["conv_faixas"], "Alta > Média > Baixa", "1. Base"),
    (2, "% dos convertidos no top 30%", "Quanto da conversão está no topo", _pct(_D["captura"]), "Acima de 30%", "1. Base"),
    (3, "Win rate real: Alta × demais", "Se o topo fecha mais de verdade",
     f"{_pct(_D['win_alta'])} × {_pct(_D['win_demais'])}", "Alta acima das demais", "1. Base e 2. Piloto"),
    (4, "Win rate: piloto × controle", "Se priorizar o topo gera mais Ganhos", "a medir", "Piloto acima do controle", "2. Piloto"),
    (5, "nMRR por EC: piloto × controle", "Se o ganho vira receita", "a medir", "Piloto acima do controle", "2. Piloto"),
    (6, "Ciclo do Ganho: Alta × demais", "Se priorizar não alonga a venda",
     f"{_D['ciclo_alta']:.0f} × {_D['ciclo_demais']:.0f} dias", "Alta ≤ demais", "1. Base e 2. Piloto"),
]
AVA_M = "{" + ", ".join("{" + ", ".join([str(o)] + [json.dumps(x, ensure_ascii=False) for x in r]) + "}"
                        for o, *r in AVALIACAO) + "}"
PARAM_TABLES["Param_Avaliacao"] = (
    [("ordem", I), ("metrica", S), ("o_que_mede", S), ("hoje", S), ("meta", S), ("etapa", S)],
    "let\n    Fonte = #table(type table [ordem = Int64.Type, metrica = text, o_que_mede = text, hoje = text, meta = text, "
    f"etapa = text],\n        {AVA_M})\nin\n    Fonte",
    "Plano de avaliação do Desafio 2 (igual à aba Como_Avaliar do Excel)", {"metrica": "ordem"})
_ON = json.loads((ROOT / "deck" / "data.json").read_text(encoding="utf-8"))["ondas"]
ONDAS = [
    ("1. Fechar", "Lead aberto há até 90 dias em Reunião ou Proposta", "Follow-up da proposta e negociação",
     "Único sinal real de compra; nenhum Ganho da base fechou depois de 89 dias"),
    ("2. Prospectar tier Alta", "Tier Alta do 1c fora da onda 1", "Outbound do EC: diagnóstico e reunião",
     "Carteira > 24 clientes (driver do D2) e lead em andamento"),
    ("3. Requalificar parado", "Lead aberto há mais de 90 dias", "Requalificar ou encerrar como Perdido",
     "Lead com mais de 89 dias nunca virou Ganho: limpar libera tempo de EC"),
    ("4. Nutrir", "Tier Média fora das ondas 1 a 3", "Marketing; EC só se o contador levantar a mão",
     "Fit ou engajamento parcial: ainda não justifica tempo de EC"),
    ("5. Régua automatizada", "Tier Baixa fora das ondas 1 a 3", "Comunicação automatizada", "Sem fit e sem engajamento"),
]
OND_M = "{" + ", ".join("{" + ", ".join([json.dumps(o, ensure_ascii=False), json.dumps(q, ensure_ascii=False),
                                         str(_ON[k]["n"]), json.dumps(a, ensure_ascii=False), json.dumps(w, ensure_ascii=False)])
                        + "}" for k, (o, q, a, w) in enumerate(ONDAS)) + "}"
PARAM_TABLES["Param_Ondas"] = (
    [("onda", S), ("quem_entra", S), ("contadores", I), ("acao", S), ("por_que", S)],
    "let\n    Fonte = #table(type table [onda = text, quem_entra = text, contadores = Int64.Type, acao = text, "
    f"por_que = text],\n        {OND_M})\nin\n    Fonte",
    "Framework de priorização do Desafio 3b (igual à aba Ondas_Capacidade do Excel)", {})
PARAM_TABLES["Param_Territorio"] = (
    [("estado", S), ("id_franquia", S), ("tipo", S)],
    f"let\n    Fonte = #table(type table [estado = text, id_franquia = text, tipo = text],\n        {TER_M})\nin\n    Fonte",
    "Franquia de cada estado com franquia (SP é tratado por cidade); estados sem franquia não entram", {})


def amplitude_dax() -> str:
    """Amplitude da taxa de conversão (maior − menor, categorias com 30+ contadores) da variável da linha."""
    linhas = ["SWITCH(SELECTEDVALUE(Param_Variaveis[variavel]),"]
    for k, (_, nome, col) in enumerate(VARIAVEIS):
        fim = "," if k < len(VARIAVEIS) - 1 else ")"
        linhas += [f'    "{nome}",',
                   f"        VAR t = FILTER(ADDCOLUMNS(ALL({MC}[{col}]), \"@n\", CALCULATE(COUNTROWS({MC})), "
                   f"\"@tx\", CALCULATE([Taxa de Conversão])), [@n] >= 30)",
                   f"        RETURN (MAXX(t, [@tx]) - MINX(t, [@tx])) * 100{fim}"]
    return "\n".join(linhas)


T = MC
MEDIDAS = [  # (nome, pasta, formato, DAX, descrição)
    ("Contadores", "1. Base", "#,0", f"COUNTROWS({T})", "Nº de contadores no contexto"),
    ("Convertidos", "1. Base", "#,0", f"SUM({T}[converteu_omie])", "Contadores com converteu_omie = 1"),
    ("Taxa de Conversão", "1. Base", "0.0%", "DIVIDE([Convertidos], [Contadores])", "Convertidos ÷ contadores"),
    ("Taxa Média", "1. Base", "0.0%", f"CALCULATE([Taxa de Conversão], REMOVEFILTERS({T}))", "Taxa dos 600 contadores"),
    ("Lift vs Média", "1. Base", '0.0"x"', "DIVIDE([Taxa de Conversão], [Taxa Média])", "Taxa do grupo ÷ taxa média"),
    ("Amplitude (p.p.)", "2. Drivers", "0", amplitude_dax(),
     "Maior − menor taxa de conversão entre as categorias (30+ contadores): quanto a característica separa quem converte"),
    ("Ranking do Driver", "2. Drivers", "0", "RANKX(ALL(Param_Variaveis), [Amplitude (p.p.)])", "1 = driver mais forte"),
    ("Taxa de Conversão ICP", "2. Drivers", "0.0%", f'CALCULATE([Taxa de Conversão], {T}[icp] = "Sim")', "Taxa do ICP"),
    ("Lift do ICP", "2. Drivers", '0.0"x"', "DIVIDE([Taxa de Conversão ICP], [Taxa Média])", "Taxa do ICP ÷ taxa média"),
    ("Conversão Top 30%", "2. Drivers", "0%", f'CALCULATE([Taxa de Conversão], REMOVEFILTERS({T}), {T}[faixa_propensao] = "Alta")',
     "Taxa de conversão (target) da faixa Alta de propensão"),
    ("Lift Top 30%", "2. Drivers", '0.0"x"', "DIVIDE([Conversão Top 30%], [Taxa Média])", "Conversão do top 30% ÷ média"),
    ("Captura Top 30%", "2. Drivers", "0%",
     f'DIVIDE(CALCULATE([Convertidos], REMOVEFILTERS({T}), {T}[faixa_propensao] = "Alta"), CALCULATE([Convertidos], REMOVEFILTERS({T})))',
     "% dos convertidos que estão no top 30% do score"),
    ("Leads", "3. Resultado Real", "#,0", "COUNTROWS(leads_pipeline)", "Leads no pipeline"),
    ("Leads com Desfecho", "3. Resultado Real", "#,0", "SUM(leads_pipeline[com_desfecho])", "Leads Ganho + Perdido"),
    ("Ganhos (Real)", "3. Resultado Real", "#,0", "SUM(leads_pipeline[ganho])", "Leads Ganho"),
    ("Win Rate Real", "3. Resultado Real", "0.0%", "DIVIDE([Ganhos (Real)], [Leads com Desfecho])",
     "Ganho ÷ (Ganho + Perdido): o fechamento comercial de verdade"),
    ("Win Rate Real ICP", "3. Resultado Real", "0.0%", f'CALCULATE([Win Rate Real], {T}[icp] = "Sim")', "Win rate real do ICP"),
    ("Win Rate Real Fora do ICP", "3. Resultado Real", "0.0%", f'CALCULATE([Win Rate Real], {T}[icp] = "Não")',
     "Win rate real de quem não é ICP"),
    ("Win Rate Real Faixa Alta", "3. Resultado Real", "0%",
     f'CALCULATE([Win Rate Real], REMOVEFILTERS({T}), {T}[faixa_propensao] = "Alta")', "Win rate real do top 30% do score"),
    ("Win Rate Real Demais Faixas", "3. Resultado Real", "0%",
     f'CALCULATE([Win Rate Real], REMOVEFILTERS({T}), {T}[faixa_propensao] <> "Alta")', "Win rate real das faixas Média e Baixa"),
    ("Endereçáveis", "4. Priorização", "#,0", f'CALCULATE([Contadores], {T}[enderecavel] = "Sim")',
     "Contadores endereçáveis (regra do Desafio 1a)"),
    ("Tier Alta", "4. Priorização", "#,0", f'CALCULATE([Endereçáveis], {T}[tier_prioridade] = "Alta")',
     "Endereçáveis no tier Alta do Desafio 1c"),
    ("ICP Endereçáveis", "4. Priorização", "#,0", f'CALCULATE([Endereçáveis], {T}[icp] = "Sim")', "Endereçáveis com perfil ICP"),
    ("Tier Média", "4. Priorização", "#,0", f'CALCULATE([Endereçáveis], {T}[tier_prioridade] = "Média")',
     "Endereçáveis no tier Média do Desafio 1c"),
    ("Tier Baixa", "4. Priorização", "#,0", f'CALCULATE([Endereçáveis], {T}[tier_prioridade] = "Baixa")',
     "Endereçáveis no tier Baixa do Desafio 1c"),
    ("Quente", "4. Priorização", "#,0", f'CALCULATE([Endereçáveis], {T}[engajamento] = "Quente")',
     "Endereçáveis com lead em Reunião/Proposta nos últimos 90 dias"),
    ("Morno", "4. Priorização", "#,0", f'CALCULATE([Endereçáveis], {T}[engajamento] = "Morno")',
     "Endereçáveis com lead Novo/Em Contato nos últimos 90 dias"),
    ("Frio", "4. Priorização", "#,0", f'CALCULATE([Endereçáveis], {T}[engajamento] = "Frio")',
     "Endereçáveis sem lead ativo (inclui parados há mais de 90 dias)"),
    ("MRR Esperado", "4. Priorização", "#,0", f"SUM({T}[mrr_esperado])", "Soma do MRR esperado dos endereçáveis"),
    ("Contadores Ondas 1 e 2", "4. Priorização", "#,0", f"CALCULATE([Endereçáveis], {T}[onda] <= 2)",
     "Endereçáveis nas ondas 1 (fechar) e 2 (prospectar tier Alta)"),
    ("Melhor ROI de Esforço", "4. Priorização", None,
     'MAXX(FILTER(d1b_conversao, d1b_conversao[ranking_roi_esforco] = 1), d1b_conversao[canal_origem])',
     "Canal com menos tentativas de contato por fechamento (Desafio 1b)"),
    ("Carteira Endereçável", "4. Priorização", "#,0", f'CALCULATE(SUM({T}[num_clientes_ativos]), {T}[enderecavel] = "Sim")',
     "Clientes finais nas carteiras dos endereçáveis"),
    ("Realizado Franquias Abaixo da Meta", "5. Impacto", "#,0",
     "SUMX(FILTER(ALLNOBLANKROW(franquias), franquias[realizado_nmrr_mes] < franquias[meta_nmrr_mes]), franquias[realizado_nmrr_mes])",
     "nMRR oficial mensal (R$) das franquias abaixo da meta"),
    ("% Capacidade Redirecionada", "5. Impacto", "0%", "SELECTEDVALUE(Param_Cenarios[pct_redirecionado])",
     "Parte da capacidade dos ECs levada ao topo da propensão"),
    ("Fator de Realização", "5. Impacto", "0%", "SELECTEDVALUE(Param_Cenarios[fator_realizacao])",
     "Fração do lift que vira Ganho real (o target diverge do real)"),
    ("Uplift nMRR", "5. Impacto", "+0.0%", "[% Capacidade Redirecionada] * ([Lift Top 30%] - 1) * [Fator de Realização]",
     "% capacidade × (lift do top 30% do score − 1) × fator de realização"),
    ("nMRR Incremental (Franquias Abaixo da Meta)", "5. Impacto", "#,0",
     "[Uplift nMRR] * [Realizado Franquias Abaixo da Meta]", "Uplift aplicado ao nMRR oficial dessas franquias (R$/mês)"),
    ("Chance de Fechar", "5. Impacto", "0.0%", "\n".join([
        "VAR lr = " + EXEC_TAXA.format(a="reunioes_realizadas", b="leads_trabalhados"),
        "VAR rp = " + EXEC_TAXA.format(a="propostas_enviadas", b="reunioes_realizadas"),
        "VAR pf = " + EXEC_TAXA.format(a="fechamentos", b="propostas_enviadas"),
        "RETURN SWITCH(SELECTEDVALUE(Param_Estagios[etapas_ate_fechar]), 1, pf, 2, rp * pf, 3, lr * rp * pf, 0)"]),
     "Chance de o lead fechar a partir do estágio atual: produto das taxas do funil que faltam (execucao_comercial)"),
    ("Leads Abertos até 90 Dias", "5. Impacto", "0",
     "CALCULATE(COUNTROWS(leads_pipeline), leads_pipeline[status_lead] = SELECTEDVALUE(Param_Estagios[estagio]), "
     "leads_pipeline[tempo_ciclo_dias] <= 90)", "Leads abertos no estágio, com até 90 dias"),
    ("Leads Parados (> 90 dias)", "5. Impacto", "0",
     "CALCULATE(COUNTROWS(leads_pipeline), leads_pipeline[status_lead] = SELECTEDVALUE(Param_Estagios[estagio]), "
     "leads_pipeline[tempo_ciclo_dias] > 90)", "Leads abertos há mais de 90 dias: nenhum Ganho da base fechou depois de 89 dias"),
    ("Ganhos Esperados", "5. Impacto", "0.0", "[Leads Abertos até 90 Dias] * [Chance de Fechar]",
     "Leads abertos até 90 dias × chance de fechar"),
    ("Ganhos Previstos (Pipeline)", "5. Impacto", "0.0", "SUMX(Param_Estagios, [Ganhos Esperados])",
     "Método 1: soma dos Ganhos esperados do pipeline aberto (próximo trimestre)"),
    ("Ticket Médio do Ganho", "5. Impacto", "#,0",
     'CALCULATE(AVERAGE(leads_pipeline[mrr_inicial]), leads_pipeline[status_lead] = "Ganho")', "mrr_inicial médio dos Ganhos"),
    ("nMRR Previsto (Pipeline)", "5. Impacto", "#,0",
     f"[Ganhos Previstos (Pipeline)] * [Ticket Médio do Ganho] * (1 - {CHURN_DAX})",
     "Ganhos previstos × ticket médio × (1 − churn), escala do CRM"),
    ("Meses de Histórico de Ganhos", "5. Impacto", "0", "\n".join([
        'VAR d1 = CALCULATE(MIN(leads_pipeline[data_conversao]), leads_pipeline[status_lead] = "Ganho")',
        'VAR d2 = CALCULATE(MAX(leads_pipeline[data_conversao]), leads_pipeline[status_lead] = "Ganho")',
        "RETURN (YEAR(d2) - YEAR(d1)) * 12 + MONTH(d2) - MONTH(d1) + 1"]), "Meses entre o primeiro e o último Ganho"),
    ("Ganhos por Trimestre (Run-rate)", "5. Impacto", "0.0",
     'DIVIDE(CALCULATE(COUNTROWS(leads_pipeline), leads_pipeline[status_lead] = "Ganho"), [Meses de Histórico de Ganhos]) * 3',
     "Método 2: Ganhos do CRM por mês × 3"),
    ("CRM ÷ nMRR Oficial", "5. Impacto", "0.0%",
     'DIVIDE(DIVIDE(CALCULATE(SUM(leads_pipeline[mrr_inicial]), leads_pipeline[status_lead] = "Ganho"), '
     "[Meses de Histórico de Ganhos]), SUM(franquias[realizado_nmrr_mes]))",
     "nMRR mensal dos Ganhos do CRM ÷ nMRR oficial da rede (tabela franquias)"),
]


F_ = "SELECTEDVALUE(franquias[id_franquia])"
MEDIDAS += [
    ("Franquias Abaixo da Meta", "6. Desafio 3", "0",
     "COUNTROWS(FILTER(ALLNOBLANKROW(franquias), franquias[realizado_nmrr_mes] < franquias[meta_nmrr_mes]))", "Nº de franquias abaixo da meta"),
    ("Correlação Nº de ECs × Atingimento", "6. Desafio 3", "0.00", "\n".join([
        "VAR t = ALLNOBLANKROW(franquias)",
        "VAR mx = AVERAGEX(t, franquias[num_executivos])",
        "VAR my = AVERAGEX(t, franquias[atingimento])",
        "VAR cov = SUMX(t, (franquias[num_executivos] - mx) * (franquias[atingimento] - my))",
        "VAR sx = SQRT(SUMX(t, (franquias[num_executivos] - mx) ^ 2))",
        "VAR sy = SQRT(SUMX(t, (franquias[atingimento] - my) ^ 2))",
        "RETURN DIVIDE(cov, sx * sy)"]), "Correlação de Pearson entre nº de ECs da franquia e atingimento (≈ 0,64)"),
    ("Gap (R$)", "6. Desafio 3", "#,0", "SUM(franquias[gap])", "Gap de meta no contexto (R$/mês)"),
    ("Atingimento (abaixo da meta)", "6. Desafio 3", "0%",
     "VAR a = DIVIDE(SUM(franquias[realizado_nmrr_mes]), SUM(franquias[meta_nmrr_mes])) RETURN IF(a < 1, a)",
     "Atingimento quando abaixo de 100%"),
    ("Atingimento (na meta)", "6. Desafio 3", "0%",
     "VAR a = DIVIDE(SUM(franquias[realizado_nmrr_mes]), SUM(franquias[meta_nmrr_mes])) RETURN IF(a >= 1, a)",
     "Atingimento quando 100% ou mais"),
    ("ECs", "6. Desafio 3", "0", "SUM(franquias[num_executivos])", "Nº de ECs"),
    ("Prospects no Território", "6. Desafio 3", "0",
     f'COUNTROWS(FILTER(ALL({MC}), {MC}[franquia_responsavel] = {F_} && {MC}[enderecavel] = "Sim"))',
     "Endereçáveis do território da franquia"),
    ("Tier Alta no Território", "6. Desafio 3", "0",
     f'COUNTROWS(FILTER(ALL({MC}), {MC}[franquia_responsavel] = {F_} && {MC}[tier_prioridade] = "Alta"))',
     "Endereçáveis do tier Alta no território"),
    ("Prospects por EC", "6. Desafio 3", "0.0", "DIVIDE([Prospects no Território], [ECs])", "Endereçáveis do território ÷ ECs"),
    ("Realizado por EC", "6. Desafio 3", "#,0", "DIVIDE(SUM(franquias[realizado_nmrr_mes]), [ECs])", "nMRR oficial por EC"),
    ("Meta por EC", "6. Desafio 3", "#,0", "DIVIDE(SUM(franquias[meta_nmrr_mes]), [ECs])", "Meta de nMRR por EC"),
    ("Meta (R$)", "6. Desafio 3", "#,0", "SUM(franquias[meta_nmrr_mes])", "Meta de nMRR (R$/mês)"),
    ("Realizado até a Meta", "6. Desafio 3", "#,0", "IF(NOT ISBLANK([Meta (R$)]), MIN([Realizado (R$)], [Meta (R$)]))",
     "Parte do realizado que preenche a meta (barra de progresso)"),
    ("Falta para a Meta", "6. Desafio 3", "#,0", "IF(NOT ISBLANK([Meta (R$)]), MAX([Meta (R$)] - [Realizado (R$)], 0))",
     "O que falta para bater a meta"),
    ("Acima da Meta", "6. Desafio 3", "#,0", "IF(NOT ISBLANK([Meta (R$)]), MAX([Realizado (R$)] - [Meta (R$)], 0))",
     "O que passou da meta"),
    ("Atingimento", "6. Desafio 3", "0%", "DIVIDE([Realizado (R$)], [Meta (R$)])", "Realizado ÷ meta no contexto (franquia ou região)"),
    ("Realizado por EC vs Rede", "6. Desafio 3", "0%",
     "DIVIDE([Realizado por EC], DIVIDE(SUMX(ALLNOBLANKROW(franquias), franquias[realizado_nmrr_mes]), SUMX(ALLNOBLANKROW(franquias), franquias[num_executivos])))",
     "Realizado por EC ÷ realizado por EC da rede"),
    ("Fechamentos por EC-mês", "6. Desafio 3", "0.00", "AVERAGE(execucao_comercial[fechamentos])",
     "Média de fechamentos por EC por mês (execucao_comercial)"),
    ("Leads por EC-mês", "6. Desafio 3", "0.0", "AVERAGE(execucao_comercial[leads_trabalhados])",
     "Média de leads trabalhados por EC por mês: a carga do EC (execucao_comercial)"),
    ("% EC-mês sem Fechamento", "6. Desafio 3", "0%",
     "DIVIDE(COUNTROWS(FILTER(execucao_comercial, execucao_comercial[fechamentos] = 0)), COUNTROWS(execucao_comercial))",
     "% dos EC-mês sem nenhum fechamento (dispersão entre ECs)"),
    ("Realizado (R$)", "6. Desafio 3", "#,0", "SUM(franquias[realizado_nmrr_mes])", "Realizado de nMRR (R$/mês)"),
    ("Conversão do Funil", "6. Desafio 3", "0.0%",
     "DIVIDE(SUM(execucao_comercial[fechamentos]), SUM(execucao_comercial[leads_trabalhados]))",
     "Fechamentos ÷ leads trabalhados (execucao_comercial, 2024)"),
    ("Fechamentos Potenciais por EC", "6. Desafio 3", "0.00", "[Prospects por EC] * [Conversão do Funil]",
     "Endereçáveis por EC × conversão do funil da franquia"),
    ("Ranking EC Adicional", "6. Desafio 3", "0",
     "IF(HASONEVALUE(franquias[id_franquia]), RANKX(ALLNOBLANKROW(franquias), [Fechamentos Potenciais por EC]))",
     "1 = franquia com mais fechamentos potenciais por EC (onde alocar o EC adicional)"),
    ("Franquia Recomendada", "6. Desafio 3", None,
     'SUBSTITUTE(MAXX(TOPN(1, ALLNOBLANKROW(franquias), [Fechamentos Potenciais por EC]), franquias[nome_franquia]), "Franquia ", "")',
     "Franquia com mais fechamentos potenciais por EC"),
]


def bloco_expr(expr: str, nivel: int) -> list[str]:
    return ["\t" * nivel + linha for linha in expr.splitlines()]


def tabela(nome, colunas, m_expr, calcs=(), medidas=(), oculta=False, sort_by=None, descricao=None):
    sort_by = sort_by or {}
    L = [f"/// {descricao}"] if descricao else []
    L += [f"table {q(nome)}", f"\tlineageTag: {guid()}"] + (["\tisHidden"] if oculta else []) + [""]
    for nm, pasta, fmt, dax, desc in medidas:
        L += [f"\t/// {desc}", f"\tmeasure {q(nm)} ="] + bloco_expr(dax, 3)
        L += ([f"\t\tformatString: {fmt}"] if fmt else []) + [f"\t\tdisplayFolder: {pasta}", f"\t\tlineageTag: {guid()}", ""]
    for nm, tipo in colunas:
        L += [f"\tcolumn {q(nm)}", f"\t\tdataType: {tipo}"]
        if nm in OCULTAS or nome == "_Medidas":
            L.append("\t\tisHidden")
        if tipo == DT:
            L.append("\t\tformatString: dd/mm/yyyy")
        if tipo == D and nome.startswith("d1"):
            L.append("\t\tformatString: 0.0")
        L += [f"\t\tlineageTag: {guid()}", "\t\tsummarizeBy: none", f"\t\tsourceColumn: {nm}"]
        if nm in sort_by:
            L.append(f"\t\tsortByColumn: {sort_by[nm]}")
        L += ["", "\t\tannotation SummarizationSetBy = User", ""]
    for nm, tipo, fmt, dax, sb, oc, desc in calcs:
        L += [f"\t/// {desc}", f"\tcolumn {q(nm)} ="] + bloco_expr(dax, 3) + [f"\t\tdataType: {tipo}"]
        if fmt:
            L.append(f"\t\tformatString: {fmt}")
        L += [f"\t\tlineageTag: {guid()}", "\t\tsummarizeBy: none"]
        if sb:
            L.append(f"\t\tsortByColumn: {sb}")
        if oc:
            L.append("\t\tisHidden")
        L += ["", "\t\tannotation SummarizationSetBy = User", ""]
    L += [f"\tpartition {q(nome)} = m", "\t\tmode: import", "\t\tsource ="] + bloco_expr(m_expr, 4)
    L += ["", "\tannotation PBI_ResultType = Table", ""]
    return "\n".join(L)


def escrever_modelo():
    tdir = SM / "tables"
    arquivos = {
        "_Medidas": tabela("_Medidas", [("Coluna", S)], 'let\n    Fonte = #table(type table [Coluna = text], {})\nin\n    Fonte',
                           medidas=MEDIDAS, descricao="Tabela de medidas do Desafio 2."),
        "mercado_contadores": tabela(MC, COLS[MC], M_MERCADO, calcs=CALC_MERCADO, sort_by={"porte": "porte_ordem"},
                                     descricao="Base de mercado: 1 linha por contador (alvo converteu_omie e características)."),
        "leads_pipeline": tabela("leads_pipeline", COLS["leads_pipeline"], M_RAW.format(t="leads_pipeline"), calcs=CALC_LEADS,
                                 descricao="Pipeline comercial: resultado real (Ganho/Perdido) para contrastar com o target."),
        "franquias": tabela("franquias", COLS["franquias"], M_RAW.format(t="franquias"), calcs=CALC_FRANQUIAS,
                            descricao="Franquias: meta e realizado de nMRR (impacto e Desafio 3)."),
        "execucao_comercial": tabela("execucao_comercial", COLS["execucao_comercial"],
                                     M_RAW.format(t="execucao_comercial"), oculta=True,
                                     descricao="Funil mensal por franquia (execucao_comercial): conversão lead → fechamento."),
        "clientes_ativos": tabela("clientes_ativos", COLS["clientes_ativos"], M_RAW.format(t="clientes_ativos"),
                                  calcs=CALC_CLIENTES,
                                  descricao="Contratos de clientes: ticket por porte e churn (Desafio 3)."),
        "tiers_1c": tabela("tiers_1c", COLS["tiers_1c"], M_TIERS, oculta=True,
                           sort_by={"fit_carteira": "nivel_fit", "engajamento": "nivel_engajamento"},
                           descricao="Tier do Desafio 1c (view dbo.vw_tiers_prospects, matriz Fit x Engajamento): 303 endereçáveis."),
        "propensao_d2": tabela("propensao_d2", COLS["propensao_d2"], M_VIEW.format(v="vw_propensao_d2"), oculta=True,
                               sort_by={"faixa_propensao": "faixa_ordem"},
                               descricao="Score de propensão do Desafio 2 por contador (view dbo.vw_propensao_d2): 600 contadores."),
        "d1a_mercado": tabela("d1a_mercado", COLS["d1a_mercado"], M_VIEW.format(v="vw_desafio1a_mercado"),
                              descricao="Desafio 1a: mercado endereçável por região e porte (view dbo.vw_desafio1a_mercado)."),
        "d1b_conversao": tabela("d1b_conversao", COLS["d1b_conversao"], M_VIEW.format(v="vw_desafio1b_conversao"),
                                descricao="Desafio 1b: conversão e ROI de esforço por canal e franquia (view dbo.vw_desafio1b_conversao)."),
    }
    for nome, (cols, mexpr, desc, sb) in PARAM_TABLES.items():
        arquivos[nome] = tabela(nome, cols, mexpr, sort_by=sb, descricao=desc)
    limpar(tdir, manter={f"{n}.tmdl" for n in arquivos})
    for nome, txt in arquivos.items():
        (tdir / f"{nome}.tmdl").write_text(txt, encoding="utf-8")
    (SM / "expressions.tmdl").write_text("\n".join([
        "/// Servidor do SQL Server com a base do case",
        'expression ServidorSQL = "localhost" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]',
        f"\tlineageTag: {guid()}", "", "\tannotation PBI_ResultType = Text", "",
        "/// Banco com as tabelas do case (esquema raw) e as views do Desafio 1 (dbo.vw_*)",
        'expression BancoSQL = "OMIE" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]',
        f"\tlineageTag: {guid()}", "", "\tannotation PBI_ResultType = Text", ""]), encoding="utf-8")
    (SM / "relationships.tmdl").write_text("\n".join([
        f"relationship {guid()}", "\tfromColumn: leads_pipeline.id_contador", "\ttoColumn: mercado_contadores.id_contador", "",
        f"relationship {guid()}", "\tfromColumn: leads_pipeline.id_franquia", "\ttoColumn: franquias.id_franquia", "",
        f"relationship {guid()}", "\tfromColumn: clientes_ativos.id_contador", "\ttoColumn: mercado_contadores.id_contador", "",
        f"relationship {guid()}", "\tfromColumn: execucao_comercial.id_franquia", "\ttoColumn: franquias.id_franquia", ""]),
        encoding="utf-8")
    ordem = list(arquivos)
    (SM / "model.tmdl").write_text("\n".join([
        "model Model", "\tculture: pt-BR", "\tdefaultPowerBIDataSourceVersion: powerBI_V3", "\tsourceQueryCulture: pt-BR",
        "\tvalueFilterBehavior: independent", "\tdataAccessOptions", "\t\tlegacyRedirects", "\t\treturnErrorValuesAsNull", "",
        "annotation __PBI_TimeIntelligenceEnabled = 0", "", 'annotation PBI_ProTooling = ["DevMode"]', "",
        'annotation PBI_QueryOrder = ["ServidorSQL","BancoSQL",' + ",".join(json.dumps(t) for t in ordem) + "]", ""]
        + [f"ref table {q(t)}" for t in ordem] + ["", "ref cultureInfo pt-BR", ""]), encoding="utf-8")


# =============================================================== relatório (PBIR) no design system Omie
VC = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.4.0/schema.json"
PG = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.1.0/schema.json"


def lit(v):
    return {"expr": {"Literal": {"Value": v}}}


def cor(h):
    return {"solid": {"color": lit(f"'{h}'")}}


def medida(nome):
    return {"Measure": {"Expression": {"SourceRef": {"Entity": "_Medidas"}}, "Property": nome}}


def coluna(tab, nome):
    return {"Column": {"Expression": {"SourceRef": {"Entity": tab}}, "Property": nome}}


def proj(campo, nome=None):
    chave = campo.get("Measure") or campo.get("Column")
    ent, prop = chave["Expression"]["SourceRef"]["Entity"], chave["Property"]
    p = {"field": campo, "queryRef": f"{ent}.{prop}", "nativeQueryRef": prop}
    if nome:
        p["displayName"] = nome
    return p


def moldura(fundo=HEX["branco"], borda=True, titulo=None):
    """Formatação do container no padrão do design system: cartão arredondado, sem sombra."""
    o = {"background": [{"properties": {"show": lit("true"), "color": cor(fundo), "transparency": lit("0D")}}],
         "border": [{"properties": {"show": lit("true" if borda else "false"), "color": cor(HEX["linha"]),
                                    "radius": lit("14D")}}],
         "dropShadow": [{"properties": {"show": lit("false")}}],
         "subTitle": [{"properties": {"show": lit("false")}}],
         "padding": [{"properties": {"top": lit("10D"), "bottom": lit("10D"), "left": lit("12D"), "right": lit("12D")}}]}
    if titulo:
        o["title"] = [{"properties": {"show": lit("true"), "text": lit("'" + titulo.replace("'", "''") + "'"),
                                      "fontFamily": lit(f"'{FONTE_T}'"), "fontSize": lit("14D"), "bold": lit("true"),
                                      "fontColor": cor(HEX["deep"])}}]
    else:
        o["title"] = [{"properties": {"show": lit("false")}}]
    return o


# O layout é desenhado numa grade de 1920×1080 e convertido para a página padrão do Power BI (1280×720):
# assim as fontes em pt ficam 1,5x maiores na tela, legíveis com a página ajustada à janela.
ESCALA = 1280 / 1920
# tamanhos das caixas de texto na página de 1280×720 (corpo ≥ 12 pt)
FONTE_TEXTO = {10: 12, 11: 12, 14: 15, 22: 20}


def container(x, y, w, h, visual, filtro=None):
    pos = {k: round(v * ESCALA) for k, v in (("x", x), ("y", y), ("width", w), ("height", h))}
    v = {"$schema": VC, "name": vid(), "position": {"x": pos["x"], "y": pos["y"], "z": 0, "width": pos["width"],
                                                    "height": pos["height"], "tabOrder": 0},
         "visual": visual}
    if filtro:
        v["filterConfig"] = {"filters": [filtro]}
    return v


def caixa_texto(x, y, w, h, paragrafos, fundo=None):
    """paragrafos: lista de listas de (texto, tamanho_pt, negrito, cor)."""
    pars = [{"textRuns": [{"value": t, "textStyle": {"fontFamily": FONTE_T if tam >= 14 else "Segoe UI",
                                                     "fontSize": f"{FONTE_TEXTO.get(tam, tam)}pt",
                                                     "fontWeight": "bold" if neg else "normal",
                                                     "color": c}} for t, tam, neg, c in p]} for p in paragrafos]
    v = {"visualType": "textbox", "objects": {"general": [{"properties": {"paragraphs": pars}}]},
         "drillFilterOtherVisuals": True}
    if fundo:
        v["visualContainerObjects"] = moldura(fundo, borda=False)
    return container(x, y, w, h, v)


def logo(x=1650, y=28, w=240, h=74):
    return container(x, y, w, h, {
        "visualType": "image",
        "objects": {"general": [{"properties": {"imageUrl": {"expr": {"ResourcePackageItem": {
            "PackageName": "RegisteredResources", "PackageType": 1, "ItemName": LOGO}}}}}],
                    "imageScaling": [{"properties": {"imageScalingType": lit("'Fit'")}}]},
        "drillFilterOtherVisuals": True})


def cabecalho(kicker, texto, destaque):
    """Padrão do deck: kicker em caixa alta teal, título em 2 linhas com a frase-chave em teal, logo à direita."""
    return [caixa_texto(24, 8, 1580, 156, [[(kicker.upper(), 11, True, HEX["teal"])],
                                            [(texto, 22, True, HEX["deep"])], [(destaque, 22, True, HEX["teal"])]]),
            logo()]


def cartao(x, y, w, h, nome_medida, rotulo, escuro=False):
    """Número de destaque: cartão gelo com número teal; versão escura (deep + cyan) para o número principal."""
    num, lab, fundo = (HEX["cyan"], HEX["branco"], HEX["deep"]) if escuro else (HEX["teal"], HEX["texto_2"], HEX["gelo"])
    return container(x, y, w, h, {
        "visualType": "card",
        "query": {"queryState": {"Values": {"projections": [proj(medida(nome_medida), rotulo)]}}},
        "objects": {"labels": [{"properties": {"color": cor(num), "fontFamily": lit(f"'{FONTE_T}'"), "fontSize": lit("28D"),
                                               "labelDisplayUnits": lit("1D")}}],
                    "categoryLabels": [{"properties": {"color": cor(lab), "fontSize": lit("12D")}}]},
        "visualContainerObjects": moldura(fundo, borda=False), "drillFilterOtherVisuals": True})


def grafico(tipo, x, y, w, h, categoria, medidas_, tit, ordenar=True, cores=None, filtro=None, rotulos=True):
    v = {"visualType": tipo,
         "query": {"queryState": {"Category": {"projections": [proj(categoria)]},
                                  "Y": {"projections": [proj(medida(m), n) for m, n in medidas_]}}},
         "objects": {"labels": [{"properties": {"show": lit("true" if rotulos else "false"), "fontSize": lit("13D"),
                                                "color": cor(HEX["deep"]),
                                                "labelDisplayUnits": lit("1D")}}],
                     "legend": [{"properties": {"show": lit("true" if len(medidas_) > 1 else "false"),
                                                "position": lit("'Top'"), "fontSize": lit("12D")}}],
                     "valueAxis": [{"properties": {"show": lit("false"), "gridlineShow": lit("false")}}],
                     "categoryAxis": [{"properties": {"fontSize": lit("12D"), "labelColor": cor(HEX["deep"])}}]},
         "visualContainerObjects": moldura(titulo=tit), "drillFilterOtherVisuals": True}
    if cores:
        v["objects"]["dataPoint"] = [{"properties": {"fill": cor(c)}, "selector": {"metadata": f"_Medidas.{m}"}}
                                     for (m, _), c in zip(medidas_, cores)]
    if ordenar:
        v["query"]["sortDefinition"] = {"sort": [{"field": categoria, "direction": "Ascending"}], "isDefaultSort": True}
    return container(x, y, w, h, v, filtro)


def grafico_combo(x, y, w, h, categoria, coluna_m, linha_m, tit):
    """Colunas (eixo da esquerda) + linha (eixo da direita), com rótulos nas duas séries."""
    v = {"visualType": "lineClusteredColumnComboChart",
         "query": {"queryState": {"Category": {"projections": [proj(categoria)]},
                                  "Y": {"projections": [proj(medida(coluna_m[0]), coluna_m[1])]},
                                  "Y2": {"projections": [proj(medida(linha_m[0]), linha_m[1])]}},
                   "sortDefinition": {"sort": [{"field": categoria, "direction": "Ascending"}], "isDefaultSort": True}},
         "objects": {"labels": [{"properties": {"show": lit("true"), "fontSize": lit("12D"), "color": cor(HEX["deep"]),
                                                "labelDisplayUnits": lit("1D")}}],
                     "legend": [{"properties": {"show": lit("true"), "position": lit("'Top'"), "fontSize": lit("12D")}}],
                     "valueAxis": [{"properties": {"show": lit("false"), "secShow": lit("true"), "gridlineShow": lit("false"),
                                                    "secLabelDisplayUnits": lit("1D")}}],
                     "categoryAxis": [{"properties": {"fontSize": lit("12D"), "labelColor": cor(HEX["deep"])}}],
                     "dataPoint": [{"properties": {"fill": cor(HEX["teal"])}, "selector": {"metadata": f"_Medidas.{coluna_m[0]}"}},
                                   {"properties": {"fill": cor(HEX["alerta"])}, "selector": {"metadata": f"_Medidas.{linha_m[0]}"}}]},
         "visualContainerObjects": moldura(titulo=tit), "drillFilterOtherVisuals": True}
    return container(x, y, w, h, v)


def tabela_visual(x, y, w, h, campos, tit, ordem=None, filtro=None, quebrar=False):
    v = {"visualType": "tableEx",
         "query": {"queryState": {"Values": {"projections": [proj(c, n) for c, n in campos]}}},
         "objects": {"columnHeaders": [{"properties": {"backColor": cor(HEX["teal"]), "fontColor": cor(HEX["branco"]),
                                                       "fontFamily": lit(f"'{FONTE_T}'"), "fontSize": lit("12D")}}],
                     "values": [{"properties": {"fontSize": lit("12D"), "fontColorPrimary": cor(HEX["deep"]),
                                                "backColorSecondary": cor(HEX["gelo"])}}],
                     "total": [{"properties": {"totals": lit("false")}}]},
         "visualContainerObjects": moldura(titulo=tit), "drillFilterOtherVisuals": True}
    if quebrar:  # texto longo: quebra de linha nas células e nos cabeçalhos
        v["objects"]["values"][0]["properties"]["wordWrap"] = lit("true")
        v["objects"]["columnHeaders"][0]["properties"]["wordWrap"] = lit("true")
    if ordem:
        v["query"]["sortDefinition"] = {"sort": [{"field": ordem[0], "direction": ordem[1]}], "isDefaultSort": True}
    return container(x, y, w, h, v, filtro)


def segmentador(x, y, w, h, campo, tit):
    return container(x, y, w, h, {
        "visualType": "slicer", "query": {"queryState": {"Values": {"projections": [proj(campo)]}}},
        "objects": {"data": [{"properties": {"mode": lit("'Dropdown'")}}],
                    "header": [{"properties": {"show": lit("false")}}],
                    "items": [{"properties": {"textSize": lit("12D")}}]},
        "visualContainerObjects": moldura(HEX["gelo"], borda=False, titulo=tit), "drillFilterOtherVisuals": True})


def filtro_coluna(tab, col, valor):
    valores = valor if isinstance(valor, list) else [valor]
    lits = [[{"Literal": {"Value": f"{x}L" if isinstance(x, int) else "'" + x.replace("'", "''") + "'"}}] for x in valores]
    return {"name": vid(), "field": coluna(tab, col), "type": "Categorical",
            "filter": {"Version": 2, "From": [{"Name": "t", "Entity": tab, "Type": 0}],
                       "Where": [{"Condition": {"In": {
                           "Expressions": [{"Column": {"Expression": {"SourceRef": {"Source": "t"}}, "Property": col}}],
                           "Values": lits}}}]}}


def br(v, casas=0):
    """Número no formato brasileiro (1.234,5)."""
    return f"{v:,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def pc(v, casas=0):
    return br(v * 100, casas) + "%"


def bloco_qa(x, y, w, h, pergunta, respostas):
    """Pergunta do enunciado em teal (subtítulo) e, logo abaixo, a resposta como afirmação direta, em cartão gelo."""
    pars = [[(pergunta, 11, True, HEX["teal"])]]
    for t in respostas:
        pars.append([(t, 11, False, HEX["deep"])])
    return caixa_texto(x, y, w, h, pars, fundo=HEX["gelo"])


def paginas():
    """Uma página por pergunta do case: pergunta → resposta (números do deck/data.json) → evidência.
    Layout desenhado numa grade 1920×1080 (convertido para 1280×720 em container())."""
    D = json.loads((ROOT / "deck" / "data.json").read_text(encoding="utf-8"))
    mc = lambda c: coluna(T, c)  # noqa: E731
    fq = lambda c: coluna("franquias", c)  # noqa: E731
    d1a = lambda c: coluna("d1a_mercado", c)  # noqa: E731
    d1b = lambda c: coluna("d1b_conversao", c)  # noqa: E731
    Y_QA, Y1 = 172, 398
    ps = []

    # ---------------------------------------------------------------- 1. Desafio 1 (1a, 1b, 1c)
    end = D["mercado"]["enderecaveis"]
    reg = {r["regiao"]: r["enderecaveis"] / end for r in D["regiao"]}
    roi = sorted(D["win_rate_canal"], key=lambda r: r["tentativas_por_fechamento"])
    conv = [r["conversao"] for r in D["conv_canal"]]
    tr = D["tiers"]
    v = cabecalho("Desafio 1 · SQL", "Mercado endereçável, conversão por canal e prioridade -",
                  f"{end} endereçáveis · {roi[0]['canal_origem']} lidera o ROI · {tr['alta']} contadores na Alta")
    v.append(bloco_qa(24, Y_QA, 600, 210, "1a · Tamanho do mercado endereçável por região e porte (volume e %)", [
        f"{end} escritórios. Sudeste {pc(reg['Sudeste'])}, Nordeste {pc(reg['Nordeste'])}, Sul {pc(reg['Sul'])}, "
        f"Centro-Oeste {pc(reg['Centro-Oeste'])}, Norte {pc(reg['Norte'])}. 99% são pequenos ou autônomos."]))
    v.append(tabela_visual(24, Y1, 600, 662,
                           [(d1a("ordem_linha"), "#"), (d1a("regiao"), "Região"), (d1a("porte"), "Porte"),
                            (d1a("num_contadores"), "Qtd."), (d1a("participacao_regiao_pct"), "% reg."),
                            (d1a("participacao_total_pct"), "% total")],
                           "Endereçáveis por região e porte", ordem=(d1a("ordem_linha"), "Ascending")))
    v.append(bloco_qa(640, Y_QA, 800, 210, "1b · Conversão por canal e franquia; melhor ROI de esforço", [
        f"{roi[0]['canal_origem']}: {br(roi[0]['tentativas_por_fechamento'], 1)} tentativas por ganho (amostra "
        f"pequena), à frente do {roi[1]['canal_origem']} ({br(roi[1]['tentativas_por_fechamento'], 1)}). "
        f"Conversão por canal de {pc(min(conv), 1)} a {pc(max(conv), 1)}."]))
    v.append(tabela_visual(640, Y1, 800, 370,
                           [(d1b("canal_origem"), "Canal"), (d1b("num_leads"), "Leads"), (d1b("num_ganhos"), "Ganhos"),
                            (d1b("taxa_conv_lead_ganho_pct"), "Conv. %"), (d1b("num_tentativas_contato"), "Tentativas"),
                            (d1b("tentativas_contato_por_ganho"), "Tent./ganho"), (d1b("ranking_roi_esforco"), "ROI #")],
                           "Por canal", ordem=(d1b("ranking_roi_esforco"), "Ascending"),
                           filtro=filtro_coluna("d1b_conversao", "ordem_bloco", 1)))
    v.append(tabela_visual(640, 784, 800, 276,
                           [(d1b("nome_franquia"), "Franquia"), (d1b("num_leads"), "Leads"), (d1b("num_ganhos"), "Ganhos"),
                            (d1b("taxa_conv_lead_ganho_pct"), "Conv. %"), (d1b("alerta_amostra"), "Alerta")],
                           "Por franquia", ordem=(d1b("taxa_conv_lead_ganho_pct"), "Descending"),
                           filtro=filtro_coluna("d1b_conversao", "ordem_bloco", 2)))
    v.append(bloco_qa(1456, Y_QA, 440, 210, "1c · Tiers com 3+ critérios", [
        f"Matriz Fit × Engajamento, sem pesos: Alta {tr['alta']} · Média {tr['media']} · Baixa {tr['baixa']}."]))
    v.append(tabela_visual(1456, Y1, 440, 270,
                           [(mc("fit_carteira"), "Fit"), (medida("Quente"), "Quente"),
                            (medida("Morno"), "Morno"), (medida("Frio"), "Frio")],
                           "Matriz Fit × Engajamento", ordem=(mc("fit_carteira"), "Descending"),
                           filtro=filtro_coluna(T, "enderecavel", "Sim")))
    for k, (m, r, esc) in enumerate([("Tier Alta", "Alta", True), ("Tier Média", "Média", False),
                                     ("Tier Baixa", "Baixa", False)]):
        v.append(cartao(1456 + k * 150, 684, 140, 150, m, r, escuro=esc))
    v.append(caixa_texto(1456, 850, 440, 210, [
        [("Fit = carteira (> 24 · 15 a 24 · < 15)", 10, False, HEX["deep"])],
        [("Engajamento = estágio + recência", 10, False, HEX["deep"])],
        [("Alta = Alto + Quente/Morno ou Médio + Quente", 10, True, HEX["teal"])]], fundo=HEX["gelo"]))
    ps.append(("D1 · Mercado, conversão e tiers", v))

    # ---------------------------------------------------------------- 2. Desafio 2: lógica, top 5 e quais contadores
    dv = D["drivers"]
    rk = sorted(dv["ranking"], key=lambda r: r["rank"])[:5]
    v = cabecalho("Desafio 2 · Propensão: lógica e top 5", "Converte mais quem é maior, estruturado e usa ERP -",
                  "e a propensão de cada contador sai desses 5 drivers")
    v.append(bloco_qa(24, Y_QA, 1872, 210,
                      "Quais variáveis você considerou relevantes e por quê? Quais contadores têm maior propensão? "
                      "Quais as top 5 características?", [
                          "Testadas 9 características; o top 5 é o que mais separa quem converte (amplitude da taxa): "
                          + ", ".join(f"{r['variavel'].lower()} {br(r['amplitude'])} p.p." for r in rk) + ".",
                          "Propensão do contador = média das taxas de conversão do perfil dele nesses 5 drivers (sem modelo "
                          "e sem pesos). Fora: score_potencial_latente (regra do case) e status_omie (sem sinal). "
                          f"Só {dv['alta_enderecaveis']} dos {end} endereçáveis estão na faixa Alta: o topo já foi convertido."]))
    v.append(tabela_visual(24, Y1, 520, 330,
                           [(coluna("Param_Variaveis", "variavel"), "Característica"),
                            (medida("Amplitude (p.p.)"), "p.p."), (medida("Ranking do Driver"), "#")],
                           "Top 5 (ranking dos drivers)", ordem=(medida("Ranking do Driver"), "Ascending")))
    v.append(grafico("clusteredColumnChart", 24, 744, 520, 316, mc("porte"), [("Taxa de Conversão", "Conversão")],
                     "Conversão por porte (driver nº 1)"))
    v.append(tabela_visual(560, Y1, 1336, 662,
                           [(mc("id_contador"), "Contador"), (mc("estado"), "UF"), (mc("porte"), "Porte"),
                            (mc("num_clientes_ativos"), "Clientes"), (mc("num_funcionarios"), "Func."),
                            (mc("erp_atual"), "ERP"), (mc("anos_mercado"), "Anos"), (mc("propensao"), "Propensão"),
                            (mc("faixa_propensao"), "Faixa"), (mc("tier_prioridade"), "Tier 1c")],
                           "Quais contadores: os 303 endereçáveis ordenados pela propensão",
                           ordem=(mc("propensao"), "Descending"), filtro=filtro_coluna(T, "enderecavel", "Sim")))
    ps.append(("D2 · Lógica, top 5 e propensão", v))

    # ---------------------------------------------------------------- 3. Desafio 2: como avaliar
    v = cabecalho("Desafio 2 · Como avaliar", "Avaliação em duas etapas -",
                  "o score ordena na base; o piloto prova se gera receita")
    v.append(bloco_qa(24, Y_QA, 1872, 210,
                      "Como você avaliaria se a abordagem está funcionando? Que métricas usaria?", [
                          f"Etapa 1, validação na base: o score ordena quem converte (Alta · Média · Baixa = "
                          f"{dv['conv_faixas']}) e o top 30% capta {pc(dv['captura'])} dos convertidos; no pipeline real, "
                          f"porém, a faixa Alta fecha {pc(dv['win_alta'])} contra {pc(dv['win_demais'])} das demais "
                          f"(mais rápido: {br(dv['ciclo_alta'])} × {br(dv['ciclo_demais'])} dias).",
                          "Etapa 2, piloto de 90 dias com grupo de controle: win rate, nMRR por EC e ciclo do piloto "
                          "melhores que os do controle. Se for, a priorização vira processo; senão, o perfil é revisto."]))
    pa = lambda c: coluna("Param_Avaliacao", c)  # noqa: E731
    v.append(tabela_visual(24, Y1, 1872, 400,
                           [(pa("metrica"), "Métrica"), (pa("o_que_mede"), "O que mede"), (pa("hoje"), "Hoje"),
                            (pa("meta"), "Meta"), (pa("etapa"), "Etapa")],
                           "Métricas de avaliação", ordem=(pa("metrica"), "Ascending"), quebrar=True))
    v.append(grafico("clusteredColumnChart", 24, 814, 1180, 246, mc("faixa_propensao"),
                     [("Taxa de Conversão", "Target"), ("Win Rate Real", "Win rate real")],
                     "Faixa de propensão: target × real", cores=[HEX["teal"], HEX["atencao"]]))
    v.append(caixa_texto(1220, 814, 676, 246, [
        [("Regra de decisão", 14, True, HEX["deep"])],
        [("Se o piloto superar o controle, a priorização pelo score vira processo; senão, o perfil é revisto.",
          10, False, HEX["deep"])],
        [("O target ordena; o real ainda não confirma.", 10, True, HEX["teal"])]], fundo=HEX["gelo"]))
    ps.append(("D2 · Como avaliar", v))

    # ---------------------------------------------------------------- 4. Desafio 2: previsibilidade e impacto (extra)
    pr, pv = dv["premissas"], D["previsao"]
    v = cabecalho("Desafio 2 · Previsibilidade e impacto em nMRR (extra)",
                  f"O pipeline prevê {br(pv['ganhos_pipeline'], 0)} a {br(pv['run_rate'], 0)} Ganhos no próximo trimestre -",
                  f"mas os Ganhos do CRM explicam só {pc(pv['crm_pct'], 1)} do nMRR oficial")
    v.append(bloco_qa(24, Y_QA, 1872, 210,
                      "Extra: qual o nMRR incremental se o time comercial priorizar os contadores com maior score de "
                      "propensão? Que premissas você assumiu?", [
                          f"Previsão (só dados): pipeline ponderado = {br(pv['ganhos_pipeline'], 1)} Ganhos no próximo trimestre "
                          f"(≈ R$ {br(pv['nmrr_pipeline'] / 1000, 1)} mil de nMRR); run-rate do CRM = {br(pv['run_rate'], 1)} "
                          f"({pv['faixa_trimestre']} por trimestre). Os métodos convergem; os {pv['parados']} leads parados entram com zero.",
                          f"Priorizar o topo: +{pc(dv['uplift'][1])} no cenário base ({pc(dv['uplift'][0])} a {pc(dv['uplift'][2])}) "
                          f"≈ {br(pv['ganhos_priorizacao'], 1)} Ganhos no trimestre e R$ {br(dv['nmrr_oficial'][1] / 1000, 1)} mil/mês "
                          "nas 9 franquias abaixo da meta, se o piloto confirmar (premissas na tabela).",
                          f"Lacuna: os Ganhos do CRM são {pc(pv['crm_pct'], 1)} do nMRR oficial: sem fonte única, prevê-se o "
                          "pipeline, não a receita oficial."]))
    pe = lambda c: coluna("Param_Estagios", c)  # noqa: E731
    v.append(tabela_visual(24, Y1, 1100, 360,
                           [(pe("estagio"), "Estágio"), (medida("Leads Abertos até 90 Dias"), "Abertos ≤ 90d"),
                            (medida("Chance de Fechar"), "Chance (funil)"), (medida("Ganhos Esperados"), "Ganhos esp."),
                            (medida("Leads Parados (> 90 dias)"), "Parados > 90d")],
                           "Método 1: pipeline ponderado pela chance de fechar",
                           ordem=(pe("ordem"), "Ascending")))
    trimestres = [t["trimestre"].split(" (")[0] for t in pv["trimestres"]]
    v.append(grafico("columnChart", 24, 774, 1100, 286, coluna("leads_pipeline", "trimestre_ganho"),
                     [("Ganhos (Real)", "Ganhos")], "Método 2: Ganhos por trimestre no CRM (2026 T1 parcial, até 07/03)",
                     cores=[HEX["teal"]], filtro=filtro_coluna("leads_pipeline", "trimestre_ganho", trimestres)))
    for k, (m_, rot, escuro) in enumerate([("Ganhos Previstos (Pipeline)", "Ganhos previstos (pipeline)", True),
                                          ("Ganhos por Trimestre (Run-rate)", "Ganhos/trimestre (run-rate)", False),
                                          ("nMRR Previsto (Pipeline)", "nMRR previsto (R$, CRM)", False),
                                          ("CRM ÷ nMRR Oficial", "CRM ÷ nMRR oficial", False)]):
        v.append(cartao(1140 + (k % 2) * 386, Y1 + (k // 2) * 188, 370, 172, m_, rot, escuro=escuro))
    v.append(tabela_visual(1140, 774, 756, 286,
                           [(coluna("Param_Cenarios", "cenario"), "Cenário"),
                            (medida("% Capacidade Redirecionada"), "% capacidade"), (medida("Fator de Realização"), "Fator"),
                            (medida("Uplift nMRR"), "Uplift"), (medida("nMRR Incremental (Franquias Abaixo da Meta)"), "R$/mês")],
                           "Alavanca: priorizar o topo da propensão", ordem=(coluna("Param_Cenarios", "cenario"), "Ascending")))
    ps.append(("D2 · Previsibilidade e impacto", v))

    # ---------------------------------------------------------------- 5. Desafio 3a
    fr = D["franquias"]
    regs = {}
    for f_ in fr:
        r_ = regs.setdefault(f_["regiao"], {"meta": 0, "real": 0, "gap": 0, "abaixo": []})
        r_["meta"] += f_["meta"]
        r_["real"] += f_["realizado"]
        r_["gap"] += f_["gap"]
        if f_["atingimento"] < 1:
            r_["abaixo"].append(f_["nome"])
    at = {k: r_["real"] / r_["meta"] for k, r_ in regs.items()}
    risco = sorted([k for k in regs if at[k] < 1], key=lambda k: at[k])
    ok_ = sorted([k for k in regs if at[k] >= 1], key=lambda k: at[k])
    reg_txt = (f"{', '.join(risco[:-1])} e {risco[-1]} estão abaixo da meta como região "
               f"({', '.join(pc(at[k]) for k in risco)}); {' e '.join(ok_)} batem a meta "
               f"({' e '.join(pc(at[k]) for k in ok_)}) e o gap vem só de "
               f"{' e '.join('/'.join(regs[k]['abaixo']) for k in ok_)}")
    rede = D["real_ec_rede"]
    baixos = sorted([f_ for f_ in fr if f_["real_ec"] / rede < 0.85], key=lambda f_: f_["real_ec"])
    ce = next(f_ for f_ in fr if f_["nome"] == "CE")
    sem = sorted(fr, key=lambda f_: f_["sem_fech"])
    v = cabecalho("Desafio 3a · Franquias e ECs vs meta", f"Gap de R$ {br(D['gap_total'] / 1000, 1)} mil/mês nas "
                  f"{D['abaixo_meta']['n']} franquias abaixo da meta -", "risco em Norte, Nordeste e Centro-Oeste")
    v.append(bloco_qa(24, Y_QA, 1872, 210,
                      "Quais franquias e ECs estão subperformando vs meta de nMRR? Qual o gap total e quais regiões "
                      "concentram o maior risco?", [
                          f"Franquias: {D['abaixo_meta']['n']} de 15 abaixo da meta, com gap de R$ {br(D['gap_total'])}/mês "
                          "(a rede como um todo bate a meta). "
                          f"Por região: {reg_txt}. O gap está nas franquias pequenas (correlação nº de ECs × atingimento "
                          f"{br(D['corr_ecs'], 2)}), não no funil ({br(D['corr_funil'], 2)}).",
                          f"ECs (equivalentes: a base não identifica o EC individual): os ECs de "
                          f"{', '.join(f_['nome'] for f_ in baixos[:4])} entregam {pc(baixos[0]['real_ec'] / rede)} a "
                          f"{pc(baixos[3]['real_ec'] / rede)} da média da rede (R$ {br(rede)} por EC); o CE entrega "
                          f"R$ {br(ce['real_ec'])} por EC, acima da rede, mas tem a maior meta por EC (R$ {br(ce['meta_ec'])}). "
                          f"Em {pc(D['sem_fech_rede'])} dos EC-mês da rede não há nenhum fechamento "
                          f"({pc(sem[0]['sem_fech'])} no {sem[0]['nome']} a {pc(sem[-1]['sem_fech'])} na {sem[-1]['nome']})."]))
    v.append(grafico("barChart", 24, Y1, 760, 662, fq("rotulo"),
                     [("Realizado até a Meta", "Realizado"), ("Falta para a Meta", "Falta para a meta"),
                      ("Acima da Meta", "Acima da meta")],
                     "Meta × realizado: a barra é a meta, o preenchimento é o realizado",
                     cores=[HEX["teal"], HEX["linha"], HEX["cyan"]], rotulos=False))
    v.append(grafico_combo(800, Y1, 1096, 300, fq("regiao"), ("Gap (R$)", "Gap das franquias abaixo da meta (R$)"),
                           ("Atingimento", "Atingimento da região"), "Por região: gap das franquias abaixo da meta e atingimento da região"))
    v.append(tabela_visual(800, 714, 1096, 346,
                           [(fq("nome_franquia"), "Franquia"), (medida("ECs"), "ECs"), (medida("Meta por EC"), "Meta/EC"),
                            (medida("Realizado por EC"), "Realizado/EC"), (medida("Realizado por EC vs Rede"), "vs rede"),
                            (medida("Fechamentos por EC-mês"), "Fech./mês"),
                            (medida("% EC-mês sem Fechamento"), "% sem fech.")],
                           "Por EC (equivalente: franquia ÷ nº de ECs)", ordem=(medida("Realizado por EC"), "Ascending")))
    ps.append(("D3a · Gap e risco", v))

    # ---------------------------------------------------------------- 6. Desafio 3b
    on = D["ondas"]
    v = cabecalho("Desafio 3b · Framework de priorização", "Primeiro quem já está comprando, depois o perfil -",
                  f"5 ondas: {on[0]['n']} para fechar, {on[1]['n']} para prospectar, {on[2]['n']} para limpar")
    v.append(bloco_qa(24, Y_QA, 1872, 210,
                      "Considerando o score de propensão do Desafio 2 e a distribuição geográfica, quais contadores atacar "
                      "primeiro, em qual ordem e por quê?", [
                          "Ordem: intenção real (lead em Reunião/Proposta) → tier Alta do 1c → lead parado → tier Média → "
                          "tier Baixa; dentro de cada onda, maior MRR esperado. O Fit do 1c é a carteira, que acompanha a "
                          f"propensão do D2 (correlação {br(D['corr_prop_carteira'], 1)} nos 303).",
                          "Geografia: cada contador fica com a franquia do seu estado (SP dividido em Capital e Interior). "
                          f"{D['ec_adicional']['sem_franquia']} endereçáveis em MT, MA, RO e MS não têm franquia: ficam "
                          "sem dono até haver cobertura."]))
    po = lambda c: coluna("Param_Ondas", c)  # noqa: E731
    v.append(tabela_visual(24, Y1, 1872, 360,
                           [(po("onda"), "Onda"), (po("quem_entra"), "Quem entra"), (po("contadores"), "Contadores"),
                            (po("acao"), "O que o EC faz"), (po("por_que"), "Por quê")],
                           "O framework", ordem=(po("onda"), "Ascending"), quebrar=True))
    v.append(tabela_visual(24, 774, 1872, 286,
                           [(mc("onda_nome"), "Onda"), (mc("id_contador"), "Contador"), (mc("estado"), "UF"),
                            (mc("franquia_nome"), "Franquia"), (mc("porte"), "Porte"),
                            (mc("num_clientes_ativos"), "Clientes"), (mc("situacao_pipeline"), "Pipeline"),
                            (mc("propensao"), "Propensão"), (mc("mrr_esperado"), "MRR esperado")],
                           "Quem atacar primeiro: ondas 1 e 2", ordem=(mc("onda_nome"), "Ascending"),
                           filtro=filtro_coluna(T, "onda_nome", ["1. Fechar (Reunião/Proposta)", "2. Prospectar tier Alta"])))
    ps.append(("D3b · Priorização", v))

    # ---------------------------------------------------------------- 7. Desafio 3c
    ec = D["ec_adicional"]
    v = cabecalho("Desafio 3c · EC adicional", "Se só houver 1 EC adicional -", f"ele vai para a {ec['franquia']}")
    v.append(bloco_qa(24, Y_QA, 1872, 210,
                      "Se você pudesse contratar apenas 1 EC adicional, em qual franquia/estado ele deveria ser alocado? "
                      "Justifique com dados.", [
                          f"{ec['franquia']}: mais fechamentos potenciais por EC da rede = {br(ec['demanda_ec'], 1)} "
                          f"endereçáveis por EC (rede: {br(ec['demanda_rede'], 1)}) × conversão do funil de "
                          f"{br(ec['conv'] * 100, 1)}% (rede: {br(ec['conv_rede'] * 100, 1)}%), com os ECs mais carregados da rede "
                          f"({br(ec['carga_ec'], 1)} leads por mês; rede: {br(ec['carga_rede'], 1)}). Está abaixo da meta (gap de "
                          f"R$ {br(ec['gap'])}) no Nordeste, a região de maior gap. 2ª opção, por pouco: {ec['segunda']} "
                          f"({br(ec['pot_2'], 2)} contra {br(ec['pot'], 2)}), já acima da meta.",
                          "Critério só com dados da base, sem premissas de custo ou território. "
                          f"MT, MA, RO e MS ({ec['sem_franquia']} endereçáveis) não têm franquia nem histórico de execução "
                          "para avaliar."]))
    v.append(cartao(24, Y1, 362, 150, "Franquia Recomendada", "EC adicional vai para", escuro=True))
    v.append(tabela_visual(24, 564, 1872, 496,
                           [(fq("nome_franquia"), "Franquia"), (medida("ECs"), "ECs"), (medida("Leads por EC-mês"), "Leads/EC-mês"),
                            (medida("Prospects no Território"), "Endereçáveis"), (medida("Prospects por EC"), "Endereç./EC"),
                            (medida("Conversão do Funil"), "Conversão do funil"),
                            (medida("Fechamentos Potenciais por EC"), "Fech. potenciais/EC"),
                            (medida("Gap (R$)"), "Gap"), (fq("atingimento"), "Atingimento"),
                            (medida("Ranking EC Adicional"), "#")],
                           "Onde alocar 1 EC: fechamentos potenciais por EC = endereçáveis por EC × conversão do funil",
                           ordem=(medida("Ranking EC Adicional"), "Ascending")))
    ps.append(("D3c · EC adicional", v))
    return ps


def escrever_relatorio():
    pdir = RP / "definition" / "pages"
    pdir.mkdir(parents=True, exist_ok=True)
    ids = []
    for nome, visuais in paginas():
        pid = vid(nome)
        ids.append(pid)
        (pdir / pid).mkdir(exist_ok=True)
        (pdir / pid / "page.json").write_text(json.dumps({
            "$schema": PG, "name": pid, "displayName": nome, "displayOption": "FitToPage", "height": 720, "width": 1280,
            "objects": {"background": [{"properties": {"color": cor(HEX["branco"]), "transparency": lit("0D")}}]}},
            ensure_ascii=False, indent=2), encoding="utf-8")
        vnomes = set()
        for k, vis in enumerate(visuais):
            vis["name"] = vid(f"{nome}|{k}")
            vis["position"]["tabOrder"] = k
            for n, flt in enumerate(vis.get("filterConfig", {}).get("filters", [])):
                flt["name"] = vid(f"{nome}|{k}|filtro{n}")
            vnomes.add(vis["name"])
            vdir = pdir / pid / "visuals" / vis["name"]
            vdir.mkdir(parents=True, exist_ok=True)
            (vdir / "visual.json").write_text(json.dumps(vis, ensure_ascii=False, indent=2), encoding="utf-8")
        limpar(pdir / pid / "visuals", manter=vnomes)
    limpar(pdir, manter=set(ids) | {"pages.json"})
    (pdir / "pages.json").write_text(json.dumps({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.1.0/schema.json",
        "pageOrder": ids, "activePageName": ids[int(os.environ.get("PBI_PAGINA_ATIVA", "0"))]},
        ensure_ascii=False, indent=2), encoding="utf-8")

    # tema Omie + logo como recursos registrados do relatório
    rr = RP / "StaticResources" / "RegisteredResources"
    rr.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(DS / LOGO, rr / LOGO)
    tema = {
        "name": "Omie",
        "dataColors": [HEX["teal"], HEX["atencao"], HEX["cyan"], HEX["alerta"], HEX["petrol"], HEX["sky"], HEX["texto_2"], HEX["deep"]],
        "background": HEX["branco"], "foreground": HEX["deep"], "tableAccent": HEX["teal"],
        "good": HEX["teal"], "neutral": HEX["atencao"], "bad": HEX["alerta"],
        "textClasses": {"title": {"fontFace": FONTE_T, "color": HEX["deep"], "fontSize": 14},
                        "callout": {"fontFace": FONTE_T, "color": HEX["teal"], "fontSize": 30},
                        "header": {"fontFace": FONTE_T, "color": HEX["deep"], "fontSize": 13},
                        "label": {"fontFace": "Segoe UI", "color": HEX["deep"], "fontSize": 12}},
        "visualStyles": {"*": {"*": {"dropShadow": [{"show": False}],
                                     "border": [{"show": True, "color": {"solid": {"color": HEX["linha"]}}, "radius": 14}]}}},
    }
    (rr / "Omie.json").write_text(json.dumps(tema, ensure_ascii=False, indent=2), encoding="utf-8")
    rj = json.loads((RP / "definition" / "report.json").read_text(encoding="utf-8"))
    base_tema = rj["themeCollection"]["baseTheme"]
    rj["themeCollection"]["customTheme"] = {"name": "Omie.json", "reportVersionAtImport": base_tema["reportVersionAtImport"],
                                            "type": "RegisteredResources"}
    rj["resourcePackages"] = [p for p in rj["resourcePackages"] if p["name"] != "RegisteredResources"] + [
        {"name": "RegisteredResources", "type": "RegisteredResources",
         "items": [{"name": "Omie.json", "path": "Omie.json", "type": "CustomTheme"},
                   {"name": LOGO, "path": LOGO, "type": "Image"}]}]
    (RP / "definition" / "report.json").write_text(json.dumps(rj, ensure_ascii=False, indent=2), encoding="utf-8")
    limpar(rr, manter={"Omie.json", LOGO})


# =============================================================== documentação
def escrever_doc():
    L = ["# Painel Omie - Desafios 1, 2 e 3 no Power BI", "",
         "Projeto `Painel Omie.pbip` (modelo em TMDL, relatório em PBIR), gerado por "
         "`omie-case-contadores/powerbi/build_pbip.py`. Abordagem **descritiva** (sem modelo estatístico), com os mesmos "
         "números dos Excel `output/Desafio2_Propensao.xlsx` e `output/Desafio3_Capacity.xlsx`, e visual no design system Omie do deck.", "",
         "**Versão online (sem instalar nada):** "
         "[Painel Omie no Power BI](https://app.powerbi.com/view?r=eyJrIjoiYjRiNTRkMjItNjA1ZS00YjA5LThlM2QtYWQ5NDRmZWM1ODliIiwidCI6IjMyOGFmM2M3LTUyYjUtNDk0Yy05YWFhLWQyYTNiYTkxYWVkZiJ9).", "",
         "## 1. Como abrir e atualizar", "",
         "1. Pré-requisito: no SQL Server, rode uma vez `sql/sqlserver/anexo_analises_complementares.sql` (cria as views "
         "`dbo.vw_tiers_prospects`, `dbo.vw_desafio1a_mercado` e `dbo.vw_desafio1b_conversao`, com o mesmo texto das "
         "consultas do Desafio 1, e `dbo.vw_propensao_d2`, com o score do Desafio 2).",
         "2. Abra `Painel Omie.pbip` no Power BI Desktop e clique em **Atualizar** (conexão **Windows** em `localhost`, banco `OMIE`).",
         "3. Outro servidor/banco: **Transformar dados → Editar parâmetros** (`ServidorSQL`, `BancoSQL`).", "",
         "## 2. Tabelas", "",
         "| Tabela | Origem | Grão | Papel |", "|---|---|---|---|",
         "| `mercado_contadores` | `raw.mercado_contadores` | 1 linha por contador (600) | Principal: `converteu_omie` e características |",
         "| `leads_pipeline` | `raw.leads_pipeline` | 1 linha por lead (350) | Resultado comercial real (Ganho/Perdido) |",
         "| `franquias` | `raw.franquias` | 1 linha por franquia (15) | Meta e realizado de nMRR (impacto) |",
         "| `tiers_1c` (oculta) | `dbo.vw_tiers_prospects` | 1 linha por endereçável (303) | Fit, engajamento, score e tier do Desafio 1c |",
         "| `d1a_mercado` | `dbo.vw_desafio1a_mercado` | região × porte + totais | Resposta do Desafio 1a |",
         "| `d1b_conversao` | `dbo.vw_desafio1b_conversao` | canal e franquia | Resposta do Desafio 1b |",
         "| `propensao_d2` (oculta) | `dbo.vw_propensao_d2` | 1 linha por contador (600) | Score e faixa de propensão do Desafio 2 |",
         "| `Param_Variaveis` | Tabela fixa | 9 características | Lista para o ranking dos drivers |",
         "| `Param_Avaliacao` | Tabela fixa | 6 métricas | Plano de avaliação em duas etapas (Desafio 2) |",
         "| `Param_Estagios` | Tabela fixa | 5 estágios | Estágios do lead aberto para o pipeline ponderado (Desafio 2, extra) |",
         "| `Param_Cenarios` | Tabela fixa | 3 cenários | Premissas do impacto |",
         "| `clientes_ativos` | `raw.clientes_ativos` | 1 linha por contrato (73) | Ticket por porte e churn (Desafio 3) |",
         "| `Param_Ondas` | Tabela fixa | 5 ondas | Quem entra, o que o EC faz e por quê (Desafio 3b) |",
         "| `Param_Territorio` | Tabela fixa | 1 linha por estado com franquia | Franquia do estado (MT, MA, RO e MS não têm franquia) |",
         "| `execucao_comercial` (oculta) | `raw.execucao_comercial` | franquia × mês × EC | Conversão do funil e dispersão por EC (Desafio 3) |",
         "| `_Medidas` | - | - | Tabela de medidas |", "",
         "## 3. Relacionamentos", "",
         "| De (muitos) | Para (um) | Direção |", "|---|---|---|",
         "| `leads_pipeline[id_contador]` | `mercado_contadores[id_contador]` | Única (mercado filtra leads) |",
         "| `leads_pipeline[id_franquia]` | `franquias[id_franquia]` | Única |",
         "| `clientes_ativos[id_contador]` | `mercado_contadores[id_contador]` | Única |",
         "| `execucao_comercial[id_franquia]` | `franquias[id_franquia]` | Única |", "",
         "`tiers_1c`, `propensao_d2` e `Param_*` não se relacionam: o tier e o score entram em `mercado_contadores` por "
         "`LOOKUPVALUE` e as tabelas de "
         "parâmetros são lidas pelas medidas.", "", "## 4. Colunas calculadas", ""]
    for tab, calcs in ((MC, CALC_MERCADO), ("leads_pipeline", CALC_LEADS), ("franquias", CALC_FRANQUIAS),
                       ("clientes_ativos", CALC_CLIENTES)):
        L += [f"### {tab}", ""]
        for nm, _, _, dax, _, _, desc in calcs:
            L += [f"**`{nm}`** - {desc}", "", "```dax", dax, "```", ""]
    L += ["## 5. Medidas", ""]
    atual = None
    for nm, pasta, fmt, dax, desc in MEDIDAS:
        if pasta != atual:
            L += [f"### {pasta}", ""]
            atual = pasta
        L += [f"**[{nm}]** - {desc}" + (f" · formato `{fmt}`" if fmt else " · texto"), "", "```dax", dax, "```", ""]
    L += ["## 6. Páginas (uma por pergunta do case: pergunta → resposta → evidência)", "",
          "- **D1 · Mercado, conversão e tiers** - 1a (endereçáveis por região e porte), 1b (conversão e ROI de esforço por "
          "canal e franquia) e 1c (matriz Fit × Engajamento e tiers), lidos das views do SQL.",
          "- **D2 · Lógica, top 5 e propensão** - variáveis e porquê, ranking dos drivers e os 303 endereçáveis ordenados pelo "
          "score de propensão.",
          "- **D2 · Como avaliar** - plano em duas etapas (base e piloto): tabela de métricas com hoje, meta e etapa, e o "
          "target × win rate real por faixa de propensão.",
          "- **D2 · Previsibilidade e impacto** - previsão do próximo trimestre por dois métodos (pipeline ponderado e "
          "run-rate do CRM), lacuna do CRM frente ao nMRR oficial e a priorização como alavanca (cenários).",
          "- **D3a · Gap e risco** - meta × realizado por franquia (barra de progresso), gap das franquias abaixo da meta e "
          "atingimento por região, e a análise por EC equivalente.",
          "- **D3b · Priorização** - as 5 ondas, ondas 1 e 2 por franquia e a lista de quem atacar primeiro.",
          "- **D3c · EC adicional** - franquia recomendada e ranking por fechamentos potenciais por EC (endereçáveis por EC × "
          "conversão do funil).", "",
          "## 7. Design system", "",
          "Cores, fontes e componentes de `deck/design_system/tokens.json`: tema `Omie.json` (paleta deep/cyan/teal, "
          "Century Gothic nos títulos), logo Omie no topo de cada página, cabeçalho com kicker em caixa alta e título em "
          "duas linhas com a frase-chave em teal, cartões gelo (número teal) e o cartão principal escuro (deep + cyan), "
          "bordas arredondadas sem sombra.", "",
          "## 8. Valores de referência (devem bater com o Excel)", "",
          "| Medida | Esperado |", "|---|---|",
          "| Taxa Média | 40,0% |", "| Taxa de Conversão ICP / Lift do ICP | 83,8% / 2,1x |",
          "| Win Rate Real ICP / Fora do ICP | 58,6% / 65,9% |",
          "| Conversão por faixa (Alta · Média · Baixa) / lift e captura do top 30% | 91% · 29% · 3% / 2,3x e 69,2% |",
          "| Ciclo médio do Ganho: faixa Alta × demais | 36 × 48 dias |",
          "| Win rate real: faixa Alta / demais | 60,0% / 65,5% |",
          "| Amplitude: Porte / Funcionários / Carteira / ERP / Anos | 93 / 89 / 88 / 59 / 34 |",
          "| Endereçáveis / Tier Alta / Média / Baixa / ICP Endereçáveis | 303 / 21 / 96 / 186 / 24 |",
          "| Matriz 1c (Alto, Médio, Baixo × Quente, Morno, Frio) | 3-13-80 / 5-9-103 / 7-10-73 |",
          "| Melhor ROI de esforço (1b) | Evento (24,8 tentativas por ganho) |",
          "| Uplift nMRR (Conservador / Base / Otimista) | +9,5% / +25,4% / +45,6% |",
          "| Ganhos previstos no trimestre: pipeline ponderado / run-rate do CRM | 6,9 / 8,1 |",
          "| nMRR previsto do pipeline (R$, escala do CRM) / CRM ÷ nMRR oficial | 5.597 / 0,9% |",
          "| Gap de meta / franquias abaixo da meta | R$ 19.500 / 9 |",
          "| Correlação nº de ECs × atingimento (meses_operando dá 0,65, mas contradiz data_inicio) | 0,64 |",
          "| Gap das franquias abaixo da meta por região: Nordeste / Norte / Centro-Oeste / Sudeste / Sul | 6.500 / 4.000 / 4.000 / 4.000 / 1.000 |",
          "| Atingimento da região: Norte / Nordeste / Centro-Oeste / Sudeste / Sul | 74% / 82% / 89% / 107% / 108% |",
          "| Ondas 1 a 5 | 15 / 13 / 54 / 66 / 155 |",
          "| EC adicional: 1º / 2º (fechamentos potenciais por EC) | BA 0,56 / MG 0,50 |",
          "| Endereçáveis sem franquia (MT, MA, RO, MS) | 27 |", ""]
    DOC.write_text("\n".join(L), encoding="utf-8")


if __name__ == "__main__":
    escrever_modelo()
    escrever_relatorio()
    escrever_doc()
    print("Modelo, relatório e documentação gerados em", PASTA)
