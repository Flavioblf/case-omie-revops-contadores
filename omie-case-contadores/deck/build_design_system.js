// Gera output/Design_System_Omie.pptx: referência visual do design system (cores, tipografia, componentes).
// Mesma fonte de verdade do deck: design_system/tokens.json + ds.js.  Uso: node deck/build_design_system.js
const path = require("path");
const { applyTheme } = require(process.env.PPTX_SKILL_DIR + "/scripts/apply_theme.js");
const { TK, HEX, THEME, ESC, HEAD, M, DIR, LOGO_RATIO, createPres } = require("./ds");

const OUT = path.join(__dirname, "..", "output", "Design_System_Omie.pptx");
const { pres, S, ui } = createPres();
pres.title = "Omie · Design System da apresentação";

const rotulo = (s, txt, x, y, w, color = HEX.texto2) => s.addText(txt, { x, y, w, h: 0.3, fontSize: 11, color, margin: 0,
  isTextBox: true, objectName: "Rótulo " + txt });

// =================================================================== SLIDE 1 — fundamentos
{
  const s = pres.addSlide({ masterName: "CLARO" });
  ui.kicker(s, "Design system · fundamentos");
  ui.titulo(s, "Cores, tipografia e logo", "extraídos do site da Omie", false);

  s.addText("Cores da marca", { x: M, y: 2.0, w: 6, h: 0.35, fontSize: 16, bold: true, fontFace: HEAD, color: S.deep,
    margin: 0, isTextBox: true, objectName: "Seção cores" });
  const marca = Object.entries(TK.cores.marca);
  marca.forEach(([nome, c], i) => {
    const x = M + i * 1.27;
    ui.card(s, `Swatch ${nome}`, x, 2.45, 1.12, 1.0, c.hex);
    rotulo(s, nome, x, 3.5, 1.12, HEX.deep);
    rotulo(s, "#" + c.hex, x, 3.75, 1.12);
  });
  s.addText("Neutras e feedback", { x: M, y: 4.25, w: 6, h: 0.35, fontSize: 16, bold: true, fontFace: HEAD, color: S.deep,
    margin: 0, isTextBox: true, objectName: "Seção neutras" });
  const outras = [["gelo", TK.cores.neutras.gelo.hex], ["superfície", TK.cores.neutras.superficie.hex],
    ["texto 2", TK.cores.neutras.texto_2.hex], ["linha", TK.cores.neutras.linha.hex],
    ["alerta", TK.cores.feedback.alerta.hex], ["atenção", TK.cores.feedback.atencao.hex]];
  outras.forEach(([nome, hex], i) => {
    const x = M + i * 1.06;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 4.7, w: 0.92, h: 0.7, fill: { color: hex },
      line: { color: HEX.linha, width: 0.75 }, rectRadius: 0.1, objectName: `Swatch ${nome}` });
    rotulo(s, nome, x, 5.45, 1.0, HEX.deep);
    rotulo(s, "#" + hex, x, 5.7, 1.0);
  });
  s.addText("Regra de contraste: cyan só sobre fundo escuro; em fundo claro, o destaque é teal.",
    { x: M, y: 6.15, w: 6.4, h: 0.5, fontSize: 12, color: HEX.texto2, margin: 0, isTextBox: true, objectName: "Regra contraste" });

  // tipografia
  const tx = 7.5;
  s.addText("Tipografia", { x: tx, y: 2.0, w: 5, h: 0.35, fontSize: 16, bold: true, fontFace: HEAD, color: S.deep,
    margin: 0, isTextBox: true, objectName: "Seção tipografia" });
  s.addText("KICKER · 12 PT", { x: tx, y: 2.45, w: 5.2, h: 0.3, fontSize: ESC.kicker, bold: true, charSpacing: 2,
    color: S.teal, margin: 0, isTextBox: true, objectName: "Amostra kicker" });
  s.addText([{ text: "Título 28 pt,", options: { color: S.deep, breakLine: true } }, { text: "destaque em teal", options: { color: S.teal } }],
    { x: tx, y: 2.8, w: 5.2, h: 0.95, fontSize: ESC.titulo, bold: true, fontFace: HEAD, margin: 0, isTextBox: true, objectName: "Amostra título" });
  s.addText("40%", { x: tx, y: 3.85, w: 1.6, h: 0.65, fontSize: ESC.numero_destaque, bold: true, fontFace: HEAD,
    color: S.teal, margin: 0, isTextBox: true, objectName: "Amostra número" });
  s.addText("Número destaque 40 pt · Century Gothic", { x: tx + 1.7, y: 3.85, w: 3.5, h: 0.65, fontSize: 12, color: HEX.texto2,
    valign: "middle", margin: 0, isTextBox: true, objectName: "Amostra número rótulo" });
  s.addText("Corpo 14 pt em Calibri: legível em cards e tabelas densos.", { x: tx, y: 4.6, w: 5.2, h: 0.35,
    fontSize: ESC.corpo, color: S.deep, margin: 0, isTextBox: true, objectName: "Amostra corpo" });
  s.addText(`Marca: ${TK.tipografia.marca}. Na apresentação: Century Gothic (títulos) e Calibri (corpo), que acompanham o Office.`,
    { x: tx, y: 5.0, w: 5.2, h: 0.6, fontSize: 11, color: HEX.texto2, margin: 0, isTextBox: true, objectName: "Nota fonte" });

  // logo
  ui.card(s, "Logo fundo escuro", tx, 5.75, 2.5, 0.95, HEX.deep);
  s.addImage({ path: path.join(DIR, "logo_omie_ciano.png"), x: tx + 0.55, y: 5.75 + (0.95 - 1.4 / LOGO_RATIO) / 2, w: 1.4, h: 1.4 / LOGO_RATIO,
    altText: "Logo Omie em cyan" });
  ui.card(s, "Logo fundo claro", tx + 2.7, 5.75, 2.5, 0.95, HEX.gelo);
  s.addImage({ path: path.join(DIR, "logo_omie_escuro.png"), x: tx + 3.25, y: 5.75 + (0.95 - 1.4 / LOGO_RATIO) / 2, w: 1.4, h: 1.4 / LOGO_RATIO,
    altText: "Logo Omie escuro" });
  ui.fonte(s, "Cores amostradas do print da home omie.com.br · tokens em deck/design_system/tokens.json");
}

// =================================================================== SLIDE 2 — componentes
{
  const s = pres.addSlide({ masterName: "ESCURO" });
  ui.kicker(s, "Design system · componentes");
  ui.titulo(s, "Componentes da apresentação", "inspirados na home da Omie", true);
  const col = [M, 4.7, 8.8];
  const sec = (txt, x, y) => s.addText(txt, { x, y, w: 3.8, h: 0.35, fontSize: 14, bold: true, fontFace: HEAD,
    color: S.branco, margin: 0, isTextBox: true, objectName: "Seção " + txt });

  sec("Botões e tags (pill)", col[0], 2.0);
  ui.pill(s, "Botão primário", "Quero conhecer", col[0], 2.5, 2.4, 0.5, { fill: HEX.cyan, color: HEX.deep, fontSize: 14 });
  ui.pill(s, "Botão secundário", "Teste grátis", col[0], 3.15, 2.4, 0.5, { line: HEX.branco, color: HEX.branco, fontSize: 14 });
  ui.pill(s, "Tag", "INDÚSTRIA", col[0], 3.8, 1.5, 0.34, { line: HEX.cyan, color: HEX.cyan, fontSize: 10 });
  ui.pill(s, "Pill de status", "R$ 7,5 mil", col[0] + 1.65, 3.8, 1.5, 0.34, { fill: HEX.atencao, color: HEX.deep, fontSize: 11 });

  sec("Número destaque", col[1], 2.0);
  ui.card(s, "Card número", col[1], 2.5, 3.8, 1.25, HEX.superficie);
  ui.numero(s, "Número exemplo", "9,3%", col[1] + 0.25, 2.5, 1.7, 1.25, S.cyan, 36);
  s.addText("rótulo curto ao lado, corpo 14 pt", { x: col[1] + 2.0, y: 2.5, w: 1.65, h: 1.25, fontSize: 13, color: S.branco,
    valign: "middle", margin: 0, isTextBox: true, objectName: "Número exemplo rótulo" });

  sec("Badge numerado", col[2], 2.0);
  [1, 2, 3].forEach((n, i) => ui.badge(s, `Badge ${n}`, String(n), col[2] + i * 0.75, 2.55, 0.55, HEX.cyan, HEX.deep));
  s.addText("Ordem de ações e drivers", { x: col[2], y: 3.25, w: 3.8, h: 0.3, fontSize: 12, color: HEX.texto2Escuro,
    margin: 0, isTextBox: true, objectName: "Badge legenda" });

  sec("Gráfico", col[0], 4.55);
  s.addChart(pres.charts.BAR, [{ name: "Série", labels: ["A", "B", "C", "D"], values: [21.8, 12.7, 9.6, 3.2] }], {
    x: col[0], y: 4.95, w: 3.8, h: 1.85, barDir: "bar", catAxisOrientation: "maxMin", ...ui.chart(true), showTitle: false,
    chartColors: [HEX.cyan], valAxisHidden: true, dataLabelFontSize: 11, objectName: "Gráfico exemplo" });

  sec("Cards", col[1], 4.55);
  ui.card(s, "Card superfície", col[1], 4.95, 3.8, 0.85, HEX.superficie);
  s.addText("Card sobre fundo escuro: superfície #0B3440", { x: col[1] + 0.2, y: 4.95, w: 3.4, h: 0.85, fontSize: 12,
    color: S.branco, valign: "middle", margin: 0, isTextBox: true, objectName: "Card superfície texto" });
  ui.card(s, "Card gelo", col[1], 5.95, 3.8, 0.85, HEX.gelo);
  s.addText("Card sobre fundo claro: gelo #EAF8FA", { x: col[1] + 0.2, y: 5.95, w: 3.4, h: 0.85, fontSize: 12,
    color: S.deep, valign: "middle", margin: 0, isTextBox: true, objectName: "Card gelo texto" });

  sec("Símbolo play", col[2], 4.55);
  s.addText("Círculo petrol com o triângulo vazado do hero do site, sangrando no canto dos slides escuros. Decorativo e sutil: nunca sobre dados críticos.",
    { x: col[2], y: 4.95, w: 2.5, h: 1.85, fontSize: 12, color: HEX.texto2Escuro, margin: 0, valign: "top", isTextBox: true,
      objectName: "Símbolo descrição" });
  ui.fonte(s, "Componentes implementados em deck/ds.js (ui.pill, ui.card, ui.numero, ui.badge, ui.titulo, ui.chart)");
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("Gerado:", OUT);
})();
