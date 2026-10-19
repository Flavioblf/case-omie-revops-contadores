// Gera output/Apresentacao_Omie_Canal_Contadores.pptx (6 slides + notas de fala) no design system Omie (deck/ds.js).
// Todos os números vêm de deck/data.json (gerado por src/deck_data.py a partir das queries e planilhas).
// Uso: node deck/build_deck.js
const path = require("path");
const fs = require("fs");
const { applyTheme } = require(process.env.PPTX_SKILL_DIR + "/scripts/apply_theme.js");
const { HEX, THEME, ESC, M, CW, createPres } = require("./ds");

const D = JSON.parse(fs.readFileSync(path.join(__dirname, "data.json"), "utf8"));
const OUT = path.join(__dirname, "..", "output", "Apresentacao_Omie_Canal_Contadores.pptx");

// ---------------------------------------------------------------- formatação pt-BR
const num = (x, dec = 0) => x.toLocaleString("pt-BR", { minimumFractionDigits: dec, maximumFractionDigits: dec });
const pct = (x, dec = 0) => num(x * 100, dec) + "%";
// espaço não separável: "R$ 31 mil" nunca quebra entre número e unidade
const mil = (x, dec = 1) => "R$ " + num(x / 1000, x % 1000 === 0 ? 0 : dec) + " mil";

const { pres, S, ui } = createPres();
pres.title = "Canal de Contadores - oportunidade, perfil de conversão e plano de 90 dias";
pres.subject = "Case técnico - Analista de Revenue Ops Sênior";

// =================================================================== SLIDE 1 - oportunidade
pres.addSection({ title: "Oportunidade" });
{
  const s = pres.addSlide({ masterName: "ESCURO", sectionTitle: "Oportunidade" });
  const mk = D.mercado;
  ui.kicker(s, "Case técnico · Revenue Ops · Canal de Contadores");
  ui.titulo(s, `${pct(1 - mk.penetracao)} do canal ainda está fora da Omie -`, "e o que sobra são escritórios pequenos e autônomos", true);

  const stats = [
    [pct(mk.penetracao, 1), `penetração: ${mk.clientes} de ${mk.universo} escritórios já são clientes ou já converteram`],
    [num(mk.enderecaveis), `escritórios endereçáveis: nem clientes, nem convertidos (${mk.ex_clientes} são ex-clientes)`],
    [num(mk.carteira_enderecavel / 1000, 1) + " mil", "clientes finais nas carteiras desses escritórios: o que o contador pode indicar"],
  ];
  stats.forEach(([big, lbl], i) => {
    const y = 2.1 + i * 1.5;
    ui.card(s, `Stat ${i + 1}`, M, y, 5.5, 1.3, HEX.superficie);
    ui.numero(s, `Stat ${i + 1} número`, big, M + 0.25, y, 2.45, 1.3, S.cyan, 36);
    s.addText(lbl, { x: M + 2.75, y, w: 2.6, h: 1.3, fontSize: ESC.corpo, color: S.branco, valign: "middle",
      margin: 0, isTextBox: true, objectName: `Stat ${i + 1} rótulo` });
  });

  const curto = { "Grande Escritório": "Grande", "Médio Escritório": "Médio", "Pequeno Escritório": "Pequeno", "Autônomo": "Autônomo" };
  const pt = D.porte;
  s.addChart(pres.charts.BAR, [{
    name: "Clientes finais", labels: pt.map(p => `${curto[p.porte]} · penetração ${pct(p.penetracao, 1)}`),
    values: pt.map(p => p.carteira),
  }], {
    x: 6.45, y: 2.0, w: 6.28, h: 4.05, barDir: "bar", catAxisOrientation: "maxMin", ...ui.chart(true),
    title: "Clientes finais nas carteiras endereçáveis, por porte", chartColors: [HEX.cyan],
    dataLabelFormatCode: "#,##0", valAxisLabelFormatCode: "#,##0", valAxisHidden: true, barGapWidthPct: 45,
    objectName: "Gráfico carteira por porte",
  });
  const pq = pt.filter(p => p.porte === "Autônomo" || p.porte === "Pequeno Escritório");
  const pqEnd = pq.reduce((a, p) => a + p.enderecaveis, 0), pqCart = pq.reduce((a, p) => a + p.carteira, 0);
  const mgEnd = mk.enderecaveis - pqEnd;
  const cv = D.converteu_vs_status;
  s.addText([
    { text: `Pequeno + Autônomo: ${pct(pqEnd / mk.enderecaveis)} dos endereçáveis e ${pct(pqCart / mk.carteira_enderecavel)} da carteira.`, options: { bold: true, color: S.cyan, breakLine: true } },
    { text: `Médios e grandes já foram convertidos: só ${mgEnd} seguem endereçáveis.`, options: { color: S.branco } },
  ], { x: 6.45, y: 6.1, w: 6.28, h: 0.75, fontSize: ESC.corpo, margin: 0, valign: "top", isTextBox: true, objectName: "Destaque porte" });
  ui.fonte(s, "Fonte: mercado_contadores (endereçável = status_omie ≠ Cliente Ativo e converteu_omie = 0) · sql/sqlserver/desafio1a_mercado_enderecavel.sql");
  s.addNotes(
    `Abro com o tamanho da oportunidade. Na base de mercado, ${mk.clientes} dos ${mk.universo} escritórios já são clientes ou já converteram: penetração de ${pct(mk.penetracao, 1)}. ` +
    `Sobram ${mk.enderecaveis} escritórios endereçáveis, com cerca de ${num(mk.carteira_enderecavel / 1000, 0)} mil clientes finais nas carteiras: é isso que o contador ainda pode indicar ou revender.\n\n` +
    `O ponto principal está no gráfico: o perfil do que sobra. Médios e grandes já foram praticamente todos convertidos, só ${mgEnd} continuam endereçáveis. ` +
    `O mercado restante é ${pct(pqEnd / mk.enderecaveis)} de pequenos escritórios e autônomos, com carteiras menores. Isso muda o modelo de atendimento: menos EC por conta e mais escala, com inbound, franquia e automação.\n\n` +
    `Uma nota de método: usei só a base de mercado, tirando quem está como Cliente Ativo e quem já converteu, e essa mesma régua vale para todo o case. ` +
    `A coluna de conversão, porém, não bate com os contratos: ${cv.cliente_ativo_com_0} clientes ativos têm valor 0, ${cv.sem_contato_com_1} contadores nunca contatados têm valor 1, ` +
    `e ${D.prospeccao.com_contrato_ativo} dos ${mk.enderecaveis} endereçáveis já têm contrato ativo na tabela de contratos. Volto a esse tema de governança no final.`);
}

// =================================================================== SLIDE 2 - onde está o dinheiro
pres.addSection({ title: "Onde está o dinheiro" });
{
  const s = pres.addSlide({ masterName: "CLARO", sectionTitle: "Onde está o dinheiro" });
  ui.kicker(s, "2 · Onde está o dinheiro");
  ui.titulo(s, "O que sobra: Sudeste e escritórios pequenos -", "e o perfil ideal (ICP) já foi quase todo convertido", false);
  const rg = D.regiao;
  s.addChart(pres.charts.BAR, [{ name: "Clientes finais", labels: rg.map(r => r.regiao), values: rg.map(r => r.carteira) }], {
    x: M, y: 2.0, w: 6.5, h: 4.75, barDir: "col", ...ui.chart(false),
    title: "Clientes finais nas carteiras endereçáveis, por região", chartColors: [HEX.teal],
    dataLabelFormatCode: "#,##0", valAxisLabelFormatCode: "#,##0", valAxisHidden: true, barGapWidthPct: 50,
    dataLabelPosition: "outEnd", objectName: "Gráfico carteira por região",
  });
  const se = rg.find(r => r.regiao === "Sudeste");
  const im = D.icp_mercado, mk = D.mercado;
  const peq = D.porte.find(p => p.porte === "Pequeno Escritório");
  const cards = [
    [pct(se.enderecaveis / mk.enderecaveis), "Sudeste",
      `${se.enderecaveis} escritórios endereçáveis e ${num(se.carteira / 1000, 1)} mil clientes finais: metade do mercado restante.`],
    [pct(peq.carteira / mk.carteira_enderecavel), "Pequenos escritórios",
      `${peq.enderecaveis} escritórios com ${num(peq.carteira / 1000, 1)} mil clientes finais: a maior fatia da carteira restante.`],
    [num(im.enderecaveis), "Escritórios ICP endereçáveis",
      `De ${im.total} no mercado: o perfil ideal (50+ clientes, 6+ funcionários, ERP) já foi quase todo convertido.`],
  ];
  cards.forEach(([big, head, txt], i) => {
    const y = 2.0 + i * 1.62;
    ui.card(s, `Card ${i + 1}`, 7.55, y, 5.18, 1.45, HEX.gelo);
    ui.numero(s, `Card ${i + 1} número`, big, 7.75, y, 1.5, 1.45, S.teal, 30);
    s.addText([
      { text: head, options: { bold: true, color: S.deep, breakLine: true } },
      { text: txt, options: { color: S.deep } },
    ], { x: 9.3, y: y + 0.08, w: 3.3, h: 1.3, fontSize: ESC.corpo, valign: "middle", margin: 0, isTextBox: true,
      objectName: `Card ${i + 1} texto` });
  });
  ui.fonte(s, "Fonte: mercado_contadores (endereçável = não cliente e converteu_omie = 0; carteira = num_clientes_ativos) · ICP: Desafio 2");
  s.addNotes(
    `Onde está o que sobra. Geografia: o Sudeste tem ${se.enderecaveis} dos ${mk.enderecaveis} escritórios endereçáveis e ${num(se.carteira / 1000, 1)} mil clientes finais, metade do mercado restante.\n\n` +
    `Perfil: os pequenos escritórios são ${peq.enderecaveis} endereçáveis e concentram ${pct(peq.carteira / mk.carteira_enderecavel)} da carteira restante; os autônomos são mais numerosos, mas com carteiras bem menores.\n\n` +
    `E um achado importante: a análise do Desafio 2 aponta um ICP simples de operar, escritório com 50 ou mais clientes, 6 ou mais funcionários e ERP de mercado. ` +
    `Dos ${im.total} escritórios com esse perfil, só ${im.enderecaveis} ainda são endereçáveis: pela base de mercado, o perfil ideal já foi quase todo convertido. ` +
    `MT e MS, sem franquia, somam ${im.mt_ms} endereçáveis. O próximo slide mostra por que essa leitura de conversão precisa ser validada.`);
}

// =================================================================== SLIDE 3 - drivers (Desafio 2, descritivo)
pres.addSection({ title: "Drivers de conversão" });
{
  const s = pres.addSlide({ masterName: "CLARO", sectionTitle: "Drivers de conversão" });
  const dv = D.drivers;
  ui.kicker(s, "3 · Propensão à conversão");
  ui.titulo(s, "Converte mais quem é maior, estruturado e usa ERP -", "mas o fechamento real ainda não confirma esse perfil", false);
  const curto = { "Autônomo": "Autônomo", "Pequeno Escritório": "Pequeno", "Médio Escritório": "Médio", "Grande Escritório": "Grande" };
  s.addChart(pres.charts.BAR, [
    { name: "Target (converteu_omie)", labels: dv.porte.map(p => curto[p.porte]), values: dv.porte.map(p => p.target) },
    { name: "Win rate real (pipeline)", labels: dv.porte.map(p => curto[p.porte]), values: dv.porte.map(p => p.win_real) },
  ], {
    x: M, y: 2.0, w: 5.6, h: 4.0, barDir: "col", ...ui.chart(false),
    title: "Conversão por porte: base × pipeline real", chartColors: [HEX.teal, HEX.atencao],
    dataLabelFormatCode: "0%", valAxisLabelFormatCode: "0%", valAxisHidden: true, valAxisMaxVal: 1.15,
    dataLabelPosition: "outEnd", dataLabelFontSize: 11, barGapWidthPct: 40,
    showLegend: true, legendPos: "b", legendFontSize: 11, legendFontFace: "+mn-lt", legendColor: HEX.deep,
    objectName: "Gráfico target x real por porte",
  });
  s.addText([
    { text: `Score de propensão: o top 30% converte ${pct(dv.top30_conv)} (${num(dv.top30_lift, 1)}x a média)`, options: { bold: true, color: S.deep, breakLine: true } },
    { text: `Média das taxas de conversão do perfil do contador nos 5 drivers · capta ${pct(dv.captura)} dos convertidos`, options: { color: HEX.texto2 } },
  ], { x: M, y: 6.1, w: 5.6, h: 0.65, fontSize: ESC.corpo, margin: 0, valign: "top", isTextBox: true, objectName: "Resumo ICP" });

  s.addText("Top 5 drivers de conversão", { x: 6.55, y: 2.0, w: 3.1, h: 0.4, fontSize: 16, bold: true, color: S.deep,
    fontFace: "+mj-lt", margin: 0, isTextBox: true, objectName: "Drivers título" });
  const leitura = { "Porte": "Médio e Grande quase sempre", "Nº de funcionários": "estrutura para adotar software",
    "Nº de clientes da carteira": "é o Fit da priorização (1c)", "ERP atual": "planilha/nenhum convertem pouco",
    "Anos de mercado": "16+ anos convertem 2x mais" };
  dv.ranking.slice(0, 5).forEach((r, i) => {
    const y = 2.55 + i * 0.82;
    ui.badge(s, `Driver ${i + 1}`, String(i + 1), 6.55, y + 0.06, 0.42, HEX.cyan, HEX.deep);
    s.addText([{ text: r.variavel, options: { bold: true, color: S.deep, breakLine: true } },
      { text: `${num(r.amplitude, 0)} p.p. entre categorias · ${leitura[r.variavel] || ""}`, options: { color: HEX.texto2, fontSize: 12 } }],
    { x: 7.1, y, w: 2.6, h: 0.72, fontSize: ESC.corpo, margin: 0, valign: "middle", isTextBox: true, objectName: `Driver ${i + 1} texto` });
  });

  ui.card(s, "Alerta target", 9.85, 2.0, 2.88, 2.2, HEX.alertaFundo);
  s.addText([
    { text: "Alerta de dados", options: { bold: true, color: S.alerta, breakLine: true } },
    { text: `Win rate real do top 30%: ${pct(dv.win_alta)}, contra ${pct(dv.win_demais)} nas demais faixas (${dv.n_alta_desfecho} leads). O score mede potencial; o perfil é hipótese a validar.`, options: { color: S.deep } },
  ], { x: 10.05, y: 2.0, w: 2.5, h: 2.2, fontSize: ESC.corpo, margin: 0, valign: "middle", isTextBox: true, objectName: "Alerta target texto" });
  ui.card(s, "Impacto", 9.85, 4.4, 2.88, 2.3, HEX.deep);
  s.addText([
    { text: `${num(D.previsao.ganhos_pipeline, 0)} a ${num(D.previsao.run_rate, 0)} Ganhos`, options: { bold: true, color: S.cyan, fontSize: 22, fontFace: "+mj-lt", breakLine: true } },
    { text: `previstos no próximo trimestre: pipeline ponderado e run-rate convergem. Priorizar o topo: +${pct(dv.uplift[1])}, se o piloto confirmar.`, options: { color: S.branco } },
  ], { x: 10.05, y: 4.4, w: 2.5, h: 2.3, fontSize: ESC.corpo, margin: 0, valign: "middle", isTextBox: true, objectName: "Impacto texto" });
  ui.fonte(s, "Análise descritiva: taxa de conversão por característica (600 contadores) · Desafio2_Propensao.xlsx · Painel Omie (Power BI), págs. 2 a 4 (D2)");
  const [d1, d2, d3, d4, d5] = dv.ranking.slice(0, 5).map(r => r.variavel.toLowerCase());
  s.addNotes(
    `Propensão à conversão. Fiz uma análise descritiva: para cada característica do contador, calculei a taxa de conversão de cada categoria. ` +
    `A característica que mais separa quem converte de quem não converte é a que tem maior diferença entre a melhor e a pior categoria. ` +
    `Em ordem: ${d1}, ${d2}, ${d3}, ${d4} e ${d5}. A propensão de cada contador é a média das taxas de conversão do perfil dele nessas cinco características: sem modelo e sem pesos. ` +
    `O score ordena bem: a faixa Alta, Média e Baixa convertem ${dv.conv_faixas}; o top 30% converte ${num(dv.top30_lift, 1)} vezes a média e capta ${pct(dv.captura)} dos convertidos.\n\n` +
    `O ponto mais importante está no gráfico: pela base, Médio e Grande convertem quase sempre e o Autônomo quase nunca. Mas no pipeline real o fechamento não segue esse padrão: o Autônomo fecha 67% e o Médio, 46%. ` +
    `E o top 30% do score fecha ${pct(dv.win_alta)} no pipeline real, contra ${pct(dv.win_demais)} nas demais faixas. Ou seja, a coluna de conversão mede potencial, não fechamento.\n\n` +
    `Previsibilidade: o pipeline aberto, ponderado pela chance de fechar de cada estágio, prevê ${num(D.previsao.ganhos_pipeline, 1)} Ganhos no próximo trimestre; o run-rate histórico do CRM dá ${num(D.previsao.run_rate, 1)}. Os dois métodos convergem. ` +
    `Os ${D.previsao.parados} leads parados há mais de 90 dias entram com zero. E um alerta: os Ganhos do CRM somam ${pct(D.previsao.crm_pct, 1)} do nMRR oficial, então dá para prever o pipeline, mas não a receita oficial sem uma fonte única.\n\n` +
    `Por isso trato o impacto da priorização como hipótese: +${pct(dv.uplift[1])} de nMRR no cenário base, cerca de ${mil(dv.nmrr_oficial[1], 0)} por mês nas franquias abaixo da meta. ` +
    `Para saber se funciona, a métrica é simples: num piloto de 90 dias, o grupo que prioriza o topo do score fecha mais do que o grupo de controle?`);
}

// =================================================================== SLIDE 4 - capacity
pres.addSection({ title: "Capacity e EC adicional" });
{
  const s = pres.addSlide({ masterName: "CLARO", sectionTitle: "Capacity e EC adicional" });
  const ec = D.ec_adicional;
  ui.kicker(s, "4 · Capacity & foco dos ECs");
  ui.titulo(s, `${mil(D.gap_total)}/mês de gap, concentrado nas franquias pequenas -`, `e o próximo EC vai para a ${ec.franquia}`, false);
  const fr = D.franquias;
  s.addChart(pres.charts.BAR, [
    { name: "Abaixo da meta", labels: fr.map(f => `${f.nome} · ${f.ecs} EC`), values: fr.map(f => f.atingimento < 1 ? f.atingimento : 0) },
    { name: "Na meta", labels: fr.map(f => `${f.nome} · ${f.ecs} EC`), values: fr.map(f => f.atingimento >= 1 ? f.atingimento : 0) },
  ], {
    x: M, y: 1.95, w: 5.7, h: 4.9, barDir: "bar", barGrouping: "stacked", catAxisOrientation: "maxMin",
    ...ui.chart(false), title: "Atingimento da meta · franquia e nº de ECs",
    chartColors: [HEX.alerta, HEX.positivo], dataLabelFormatCode: '0%;;;', dataLabelPosition: "inEnd",
    dataLabelColor: HEX.branco, dataLabelFontSize: 10, catAxisLabelFontSize: 11, valAxisHidden: true, barGapWidthPct: 30,
    showLegend: true, legendPos: "b", legendFontSize: 11, legendFontFace: "+mn-lt", legendColor: HEX.deep,
    objectName: "Gráfico atingimento por franquia",
  });

  const mini = [[num(D.corr_ecs, 2), "Tamanho", "nº de ECs × atingimento"],
    [num(D.corr_funil, 2), "Funil", "proposta → fechamento × atingimento"]];
  mini.forEach(([big, head, lbl], i) => {
    const x = 6.75 + i * 3.04;
    ui.card(s, `Correlação ${i + 1}`, x, 1.95, 2.94, 1.05, HEX.gelo);
    ui.numero(s, `Correlação ${i + 1} número`, big, x + 0.15, 1.95, 1.0, 1.05, i ? HEX.texto2 : S.teal, 26);
    s.addText([{ text: `Correlação: ${head}`, options: { bold: true, color: S.deep, breakLine: true } },
      { text: lbl, options: { color: S.deep } }],
    { x: x + 1.15, y: 1.95, w: 1.72, h: 1.05, fontSize: 12, valign: "middle", margin: 0, isTextBox: true,
      objectName: `Correlação ${i + 1} rótulo` });
  });

  // gap por região (só dados da base): onde está o risco
  const porReg = (rs) => fr.filter(f => rs.includes(f.regiao) && f.atingimento < 1);
  const soma = (g) => g.reduce((a, f) => a + f.gap, 0);
  const nomes = (g) => g.map(f => f.nome).join(", ");
  const ne = porReg(["Nordeste"]), noco = porReg(["Norte", "Centro-Oeste"]), suse = porReg(["Sudeste", "Sul"]);
  const ating = Object.fromEntries(D.regioes_gap.map(r => [r.regiao, r.atingimento]));
  const diag = [
    [mil(soma(ne)), "Nordeste", `região a ${pct(ating["Nordeste"])} da meta; ${nomes(ne)} abaixo, com 1 a 2 ECs`, HEX.alerta, HEX.branco],
    [mil(soma(noco)), "Norte e Centro-Oeste", `regiões a ${pct(ating["Norte"])} e ${pct(ating["Centro-Oeste"])} da meta`, HEX.atencao, HEX.deep],
    [mil(soma(suse)), "Sudeste e Sul", `batem a meta (${pct(ating["Sudeste"])} e ${pct(ating["Sul"])}); gap só no ${nomes(suse).replace(", ", " e no ")}`, HEX.texto2, HEX.branco],
  ];
  diag.forEach(([v, h, t, fill, color], i) => {
    const y = 3.2 + i * 0.72;
    ui.pill(s, `Diagnóstico ${i + 1}`, v, 6.75, y + 0.04, 1.45, 0.5, { fill, color, fontSize: 14 });
    s.addText([{ text: h + ": ", options: { bold: true, color: S.deep } }, { text: t, options: { color: S.deep } }],
      { x: 8.35, y, w: 4.38, h: 0.58, fontSize: 13, valign: "middle", margin: 0, isTextBox: true, objectName: `Diagnóstico ${i + 1} texto` });
  });

  ui.card(s, "EC adicional", 6.75, 5.42, 5.98, 1.43, HEX.deep);
  s.addText([
    { text: `+1 EC na ${ec.franquia}`, options: { bold: true, color: S.cyan, fontSize: 17, fontFace: "+mj-lt", breakLine: true } },
    { text: `Mais fechamentos potenciais por EC: ${num(ec.demanda_ec, 1)} endereçáveis por EC (rede: ${num(ec.demanda_rede, 1)}), conversão de ${num(ec.conv * 100, 1)}% (rede: ${num(ec.conv_rede * 100, 1)}%) e os ECs mais carregados (${num(ec.carga_ec, 1)} leads/mês; rede: ${num(ec.carga_rede, 1)}). 2ª opção, por pouco: ${ec.segunda}.`, options: { color: S.branco } },
  ], { x: 6.95, y: 5.47, w: 5.6, h: 1.33, fontSize: 13, margin: 0, valign: "middle", isTextBox: true, objectName: "EC adicional texto" });
  ui.fonte(s, "Meta e realizado: tabela franquias (oficial) · funil: execucao_comercial 2024 · Desafio3_Capacity.xlsx · Painel Omie (Power BI), págs. 5 a 7 (D3)");
  s.addNotes(
    `Capacity. São ${D.abaixo_meta.n} franquias abaixo da meta, com gap de ${mil(D.gap_total)} por mês. A primeira hipótese seria execução, mas os dados não sustentam: ` +
    `as taxas do funil são praticamente iguais entre franquias e a correlação da etapa proposta → fechamento com o atingimento é de só ${num(D.corr_funil, 2)}. ` +
    `O gap está nas franquias pequenas: das 7 com 1 a 2 ECs, 6 estão abaixo da meta, e a correlação entre nº de ECs e atingimento é de ${num(D.corr_ecs, 2)}. ` +
    `Não chamo isso de maturidade: a coluna meses_operando diz o mesmo, mas contradiz a data_inicio da própria tabela, e pela data de início a correlação é negativa (${num(D.corr_data_inicio, 2)}). É mais um ponto de governança.\n\n` +
    `A rede como um todo bate a meta; os ${mil(D.gap_total)} são o gap das ${D.abaixo_meta.n} franquias abaixo dela. Por região: Norte, Nordeste e Centro-Oeste estão abaixo da meta como região (${pct(ating["Norte"])}, ${pct(ating["Nordeste"])} e ${pct(ating["Centro-Oeste"])}); Sudeste e Sul batem a meta (${pct(ating["Sudeste"])} e ${pct(ating["Sul"])}) e o gap vem só do RJ, que tem o maior gap da rede, e do RS. ` +
    `O CE, por exemplo, tem a maior meta por EC da rede e realiza acima da média por EC: parte do gap é meta, não execução.\n\n` +
    `Sobre o EC adicional: usei só dados da base. O EC vai para onde há mais fechamentos potenciais por EC, que é a demanda do território por EC vezes a conversão do funil da própria franquia. ` +
    `A ${ec.franquia} lidera: ${num(ec.demanda_ec, 1)} endereçáveis por EC, contra ${num(ec.demanda_rede, 1)} da rede, e conversão de ${num(ec.conv * 100, 1)}%, contra ${num(ec.conv_rede * 100, 1)}% da rede (${ec.rank_demanda}ª em demanda e ${ec.rank_conv}ª em conversão). ` +
    `E os ECs dela já são os mais carregados da rede: ${num(ec.carga_ec, 1)} leads trabalhados por mês, contra ${num(ec.carga_rede, 1)}. ` +
    `A margem é pequena: ${num(ec.pot, 2)} contra ${num(ec.pot_2, 2)} fechamentos potenciais por EC da ${ec.segunda}, que converte parecido (${num(ec.conv_2 * 100, 1)}%) e já está acima da meta. A diferença vem da demanda por EC. ` +
    `Os ${ec.sem_franquia} endereçáveis de MT, MA, RO e MS não têm franquia nem histórico de execução, então não dá para avaliá-los com dado.`);
}

// =================================================================== SLIDE 5 - ações
pres.addSection({ title: "Ações imediatas" });
{
  const s = pres.addSlide({ masterName: "ESCURO", sectionTitle: "Ações imediatas" });
  ui.kicker(s, "5 · Ações imediatas");
  ui.titulo(s, "Três ações imediatas", "para o time de campo", true);
  const on = D.ondas, pl = D.pipeline;
  const acoes = [
    ["Fechar o que já está quente",
      `${on[0].n} contadores com lead ativo em Reunião ou Proposta: fechar antes dos 90 dias (nenhum Ganho da base passou de 89). Requalificar ou encerrar os parados de outros ${on[2].n}.`,
      `Hoje, ${pct(pl.parados / pl.abertos)} do pipeline aberto está parado há mais de 90 dias`],
    ["Piloto do tier Alta com grupo de controle",
      `${on[1].n} contadores do tier Alta ainda fora de negociação (os outros ${D.tiers.alta_quente} já estão na ação 1). Metade dos ECs prioriza essa lista por 90 dias; a outra metade segue o processo atual.`,
      `Potencial de +${pct(D.drivers.uplift[1])} de nMRR, a ser provado`],
    ["Playbook de proposta → fechamento",
      `É a etapa de menor conversão (${pct(D.bench_pf)}). Proposta padrão por porte, decisor mapeado e follow-up.`,
      `+${num(D.sim_pf.pp * 100)} p.p. nessa etapa = +${pct(D.sim_pf.uplift)} de nMRR`],
  ];
  acoes.forEach(([h, t, imp], i) => {
    const x = M + i * 4.13;
    ui.card(s, `Ação ${i + 1}`, x, 1.95, 3.88, 4.0, HEX.superficie);
    ui.badge(s, `Ação ${i + 1} número`, String(i + 1), x + 0.3, 2.2, 0.55, HEX.cyan, HEX.deep);
    s.addText(h, { x: x + 0.3, y: 2.95, w: 3.3, h: 0.8, fontSize: 18, bold: true, color: S.branco, fontFace: "+mj-lt",
      margin: 0, valign: "top", isTextBox: true, objectName: `Ação ${i + 1} título` });
    s.addText(t, { x: x + 0.3, y: 3.8, w: 3.3, h: 1.4, fontSize: ESC.corpo, color: HEX.texto2Escuro, margin: 0, valign: "top",
      isTextBox: true, objectName: `Ação ${i + 1} texto` });
    s.addText(imp, { x: x + 0.3, y: 5.2, w: 3.3, h: 0.65, fontSize: ESC.corpo, bold: true, color: S.cyan, margin: 0,
      valign: "top", isTextBox: true, objectName: `Ação ${i + 1} impacto` });
  });
  s.addText("Decisões de gestão:", { x: M, y: 6.2, w: 2.0, h: 0.45, fontSize: ESC.corpo, bold: true, color: S.branco,
    margin: 0, valign: "middle", isTextBox: true, objectName: "Decisões de gestão rótulo" });
  [["Metas das franquias pequenas", 3.0], ["Plano de ação no RJ", 3.0], [`+1 EC na ${D.ec_adicional.franquia}`, 3.0]]
    .forEach(([t, w], i) => ui.pill(s, `Decisão ${i + 1}`, t, 2.65 + i * 3.2, 6.2, w, 0.45,
      { line: HEX.branco, color: HEX.branco, fontSize: 12 }));
  ui.fonte(s, "Lista priorizada por franquia: Desafio3_Capacity.xlsx (Priorizacao e Ondas_Capacidade) · Painel Omie (Power BI), págs. 6 e 7");
  s.addNotes(
    `As três ações para o campo, em ordem de prioridade.\n\n` +
    `Primeira: fechar o que já está quente. São ${on[0].n} contadores com lead ativo em reunião ou proposta, o único sinal real de compra que temos, independente de qualquer perfil. ` +
    `Junto com isso, limpar o pipeline: ${pct(pl.parados / pl.abertos)} dos leads abertos estão parados há mais de 90 dias, o dobro do ciclo médio de um ganho. Isso libera tempo de EC e torna a cobertura confiável.\n\n` +
    `Segunda: o piloto do tier Alta, a lista priorizada da pergunta C: carteira acima de 24 clientes com lead em andamento, ou de 15 a 24 com lead em reunião ou proposta. São ${D.tiers.alta} contadores; ${D.tiers.alta_quente} já estão em negociação, então o piloto ataca os outros ${on[1].n}. Metade dos ECs ataca essa lista, a outra metade é o grupo de controle. Em 90 dias eu sei se o perfil gera receita de verdade.\n\n` +
    `Terceira: proposta → fechamento é a etapa de menor conversão, ${pct(D.bench_pf)}. Cinco pontos percentuais a mais nela valem ${pct(D.sim_pf.uplift)} de nMRR.\n\n` +
    `Embaixo estão as três decisões de gestão que sustentam isso: meta revista nas franquias pequenas, plano para o RJ e o EC adicional em ${D.ec_adicional.franquia}.`);
}

// =================================================================== SLIDE 6 - KPI
pres.addSection({ title: "KPI de 90 dias" });
{
  const s = pres.addSlide({ masterName: "CLARO", sectionTitle: "KPI de 90 dias" });
  const ab = D.abaixo_meta;
  ui.kicker(s, "6 · KPI de 90 dias");
  ui.titulo(s, "KPI norte em 90 dias:", `levar as ${ab.n} franquias abaixo da meta à meta`, false);
  ui.card(s, "KPI norte", M, 1.95, 4.1, 4.75, HEX.deep);
  s.addText([
    { text: "KPI NORTE", options: { bold: true, color: S.cyan, fontSize: ESC.kicker, charSpacing: 2, breakLine: true } },
    { text: `${pct(ab.realizado / ab.meta)} → 100%`, options: { bold: true, color: S.branco, fontSize: 36, fontFace: "+mj-lt", breakLine: true } },
    { text: `atingimento de nMRR das ${ab.n} franquias abaixo da meta (${mil(ab.realizado, 0)} de ${mil(ab.meta)} por mês)`, options: { color: HEX.texto2Escuro, fontSize: ESC.corpo, breakLine: true } },
    { text: " ", options: { fontSize: 10, breakLine: true } },
    { text: "Acompanhamento por franquia no Painel Omie (Power BI), a mesma base do case.", options: { color: S.branco, fontSize: ESC.corpo } },
  ], { x: M + 0.3, y: 2.15, w: 3.5, h: 3.0, margin: 0, valign: "top", isTextBox: true, objectName: "KPI norte texto" });

  // barra de progresso: atingimento atual vs a meta das próprias franquias (100%)
  const atual = ab.realizado / ab.meta, bx = M + 0.3, bw = 3.5, by = 5.5;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: bx, y: by, w: bw, h: 0.3, fill: { color: HEX.petrol },
    line: { color: HEX.petrol }, rectRadius: 0.15, objectName: "Progresso fundo" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: bx, y: by, w: bw * atual, h: 0.3, fill: { color: HEX.cyan },
    line: { color: HEX.cyan }, rectRadius: 0.15, objectName: "Progresso atual" });
  s.addShape(pres.shapes.RECTANGLE, { x: bx + bw - 0.04, y: by - 0.1, w: 0.04, h: 0.5, fill: { color: HEX.branco },
    line: { color: HEX.branco }, objectName: "Marcador meta" });
  s.addText(`hoje ${pct(atual)}`, { x: bx, y: by + 0.42, w: 1.6, h: 0.3, fontSize: 12, color: S.cyan, bold: true,
    margin: 0, isTextBox: true, objectName: "Progresso rótulo atual" });
  s.addText("meta 100%", { x: bx + bw - 1.6, y: by + 0.42, w: 1.6, h: 0.3, fontSize: 12, color: S.branco,
    align: "right", margin: 0, isTextBox: true, objectName: "Progresso rótulo alvo" });

  const Hd = (t) => ({ text: t, options: { bold: true, color: HEX.branco, fill: { color: HEX.teal } } });
  const rowsT = [
    [Hd("KPI de suporte"), Hd("Hoje"), Hd("Sinal de sucesso"), Hd("Por quê")],
    ["Win rate do tier Alta (piloto vs controle)", `${pct(D.pipeline.win_rate_rede)} (rede)`, "Piloto acima do controle", "Valida o perfil"],
    ["Conversão proposta → fechamento", pct(D.bench_pf), "Subir", "Maior alavanca do funil"],
    ["Leads abertos parados há +90 dias", pct(D.pipeline.parados / D.pipeline.abertos), "Cair", "Pipeline confiável"],
    ["Tier Alta com lead em Reunião/Proposta", `${D.tiers.alta_quente} de ${D.tiers.alta}`, "Todos", "Cobertura do tier Alta"],
  ];
  s.addTable(rowsT, { x: 5.0, y: 1.95, w: 7.73, colW: [2.95, 1.35, 1.65, 1.78], fontSize: 13, color: HEX.deep,
    fontFace: THEME.bodyFontFace, border: { type: "solid", pt: 0.5, color: HEX.linha }, rowH: 0.52, valign: "middle",
    margin: [0.04, 0.08, 0.04, 0.08], objectName: "Tabela KPIs de suporte" });

  const rf = D.receita_fontes;
  ui.card(s, "Governança", 5.0, 4.9, 7.73, 1.8, HEX.gelo);
  s.addText([
    { text: "Pré-requisito: uma fonte única de nMRR", options: { bold: true, color: S.teal, breakLine: true } },
    { text: `Hoje o mesmo KPI vale ${mil(rf[0], 0)}, ${mil(rf[1], 0)} ou ${mil(rf[2], 0)} por mês, conforme a tabela (franquias, execução ou CRM). `, options: { color: S.deep } },
    { text: "Junto: ID único de EC no CRM e definição canônica de conversão, para que o piloto meça a mesma coisa em todas as franquias.", options: { color: S.deep } },
  ], { x: 5.25, y: 5.0, w: 7.3, h: 1.6, fontSize: ESC.corpo, margin: 0, valign: "middle", isTextBox: true, objectName: "Governança texto" });
  ui.fonte(s, "Hoje: franquias, leads_pipeline e Desafio 3 · Painel Omie (Power BI) · sinais de sucesso a validar com a liderança");
  s.addNotes(
    `Para fechar: como medir o sucesso em 90 dias. Proponho um único KPI norte, o atingimento de nMRR das ${ab.n} franquias que hoje estão abaixo da meta. ` +
    `Elas estão em ${pct(ab.realizado / ab.meta)} e a meta é chegar à meta delas, 100%, acompanhado no painel de Power BI que montei sobre a mesma base do case: drivers, contraste com o real, priorização e capacity, tudo lendo direto do SQL Server.\n\n` +
    `Os quatro indicadores de suporte medem se as ações estão funcionando: o win rate do piloto contra o grupo de controle, a conversão de proposta para fechamento, ` +
    `o percentual de pipeline parado e a cobertura dos contadores do tier Alta.\n\n` +
    `E o pré-requisito, que é o que eu faria primeiro: hoje o nMRR do canal vale ${mil(rf[0], 0)}, ${mil(rf[1], 0)} ou ${mil(rf[2], 0)} por mês, dependendo da tabela que se consulta. ` +
    `Sem uma fonte única, não dá para medir o piloto, calibrar meta nem decidir headcount com segurança. É exatamente o papel da camada semântica de Parcerias.`);
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("Gerado:", OUT);
})();
