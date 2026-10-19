// Design system Omie para pptxgenjs: tema, layouts e componentes, a partir de design_system/tokens.json.
// Usado por build_deck.js (apresentação) e build_design_system.js (página de referência do design system).
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");

const DIR = path.join(__dirname, "design_system");
const TK = JSON.parse(fs.readFileSync(path.join(DIR, "tokens.json"), "utf8"));
const cor = (grupo, nome) => TK.cores[grupo][nome].hex;

// Paleta em hex (para opções que só aceitam hex: gráficos, linhas, sombras)
const HEX = {
  deep: cor("marca", "deep"), cyan: cor("marca", "cyan"), sky: cor("marca", "sky"), teal: cor("marca", "teal"),
  petrol: cor("marca", "petrol"), branco: cor("neutras", "branco"), gelo: cor("neutras", "gelo"),
  superficie: cor("neutras", "superficie"), texto2: cor("neutras", "texto_2"), texto2Escuro: cor("neutras", "texto_2_escuro"),
  linha: cor("neutras", "linha"), linhaEscuro: cor("neutras", "linha_escuro"), alerta: cor("feedback", "alerta"),
  atencao: cor("feedback", "atencao"), alertaFundo: cor("feedback", "alerta_fundo"),
  positivo: TK.graficos.positivo,
};

const THEME = {
  name: "Omie Design System",
  headFontFace: TK.tipografia.titulos.fonte,
  bodyFontFace: TK.tipografia.corpo.fonte,
  colors: {
    dk1: HEX.deep, lt1: HEX.branco, dk2: HEX.petrol, lt2: HEX.gelo,
    accent1: HEX.cyan, accent2: HEX.teal, accent3: HEX.alerta, accent4: HEX.superficie,
    accent5: HEX.sky, accent6: HEX.atencao, hlink: HEX.teal, folHlink: HEX.petrol,
  },
};
const ESC = TK.tipografia.escala_pt;
const HEAD = "+mj-lt"; // fonte de título do tema (Century Gothic)
const W = 13.333, H = 7.5, M = TK.forma.margem_pol, CW = W - 2 * M;
const RAIO = TK.forma.raio_card_pol;
const LOGO_RATIO = 1028 / 316;

function createPres() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;
  // nomes semânticos -> cores do esquema (seguem o tema se ele mudar)
  const S = {
    deep: C.text1, branco: C.background1, gelo: C.background2, petrol: C.text2, cyan: C.accent1, teal: C.accent2,
    alerta: C.accent3, superficie: C.accent4, sky: C.accent5, atencao: C.accent6,
  };

  const kicker = { x: M, y: 0.4, w: 9.5, h: 0.3, fontSize: ESC.kicker, bold: true, charSpacing: 2, margin: 0, valign: "middle" };
  const titulo = { x: M, y: 0.8, w: CW, h: 1.05, fontSize: ESC.titulo, bold: true, valign: "top", align: "left", margin: 0 };
  const fonte = { x: M, y: 6.95, w: CW - 1, h: 0.3, fontSize: ESC.legenda, margin: 0, valign: "middle" };
  const logoW = 1.1, logo = { x: W - M - logoW, y: 0.36, w: logoW, h: logoW / LOGO_RATIO };

  pres.defineSlideMaster({
    title: "ESCURO",
    background: { color: S.deep },
    objects: [
      { image: { x: 10.5, y: 4.4, w: 4.6, h: 4.6, path: path.join(DIR, "simbolo_play.png") } },
      { image: { ...logo, path: path.join(DIR, "logo_omie_ciano.png") } },
      { placeholder: { options: { name: "kicker", type: "body", ...kicker, color: S.cyan }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", ...titulo, color: S.branco }, text: "" } },
      { placeholder: { options: { name: "fonte", type: "body", ...fonte, color: HEX.texto2Escuro }, text: "" } },
    ],
    slideNumber: { x: W - M - 0.5, y: 6.95, w: 0.5, h: 0.3, fontSize: ESC.legenda, color: HEX.texto2Escuro, align: "right" },
  });
  pres.defineSlideMaster({
    title: "CLARO",
    background: { color: S.branco },
    objects: [
      { image: { ...logo, path: path.join(DIR, "logo_omie_escuro.png") } },
      { placeholder: { options: { name: "kicker", type: "body", ...kicker, color: S.teal }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", ...titulo, color: S.deep }, text: "" } },
      { placeholder: { options: { name: "fonte", type: "body", ...fonte, color: HEX.texto2 }, text: "" } },
    ],
    slideNumber: { x: W - M - 0.5, y: 6.95, w: 0.5, h: 0.3, fontSize: ESC.legenda, color: HEX.texto2, align: "right" },
  });

  // ------------------------------------------------------------ componentes
  const ui = {
    // Título em duas cores e duas linhas: mensagem na 1ª, frase-chave em cyan (escuro) ou teal (claro) na 2ª,
    // como no hero do site ("Sistema de gestão... / tudo em um só lugar.")
    titulo(slide, texto, destaque, escuro) {
      slide.addText([{ text: texto, options: { breakLine: true } },
        { text: destaque, options: { color: escuro ? S.cyan : S.teal } }], { placeholder: "title" });
    },
    kicker(slide, texto) { slide.addText(texto.toUpperCase(), { placeholder: "kicker" }); },
    fonte(slide, texto) { slide.addText(texto, { placeholder: "fonte" }); },

    card(slide, nome, x, y, w, h, fill) {
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: fill },
        rectRadius: RAIO, objectName: nome });
    },
    // Pill: botão/tag da marca. variante: "primario" (cyan), "contorno", ou cor de fundo em hex
    pill(slide, nome, texto, x, y, w, h, { fill, line, color, fontSize = 12, bold = true } = {}) {
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: h / 2,
        fill: fill ? { color: fill } : { type: "none" }, line: { color: line || fill, width: 1.25 }, objectName: nome });
      slide.addText(texto, { x, y, w, h, align: "center", valign: "middle", fontSize, bold, color, margin: 0,
        isTextBox: true, objectName: nome + " texto" });
    },
    badge(slide, nome, texto, x, y, d, fill, color) {
      slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill }, objectName: nome });
      slide.addText(texto, { x, y, w: d, h: d, align: "center", valign: "middle", fontSize: 14, bold: true, color,
        fontFace: HEAD, margin: 0, isTextBox: true, objectName: nome + " texto" });
    },
    numero(slide, nome, texto, x, y, w, h, color, fontSize = ESC.numero_destaque) {
      slide.addText(texto, { x, y, w, h, fontSize, bold: true, color, fontFace: HEAD, valign: "middle", margin: 0,
        isTextBox: true, fit: "none", objectName: nome });
    },
    chart(escuro) {
      const label = escuro ? HEX.texto2Escuro : HEX.deep;
      return {
        catAxisLabelColor: label, valAxisLabelColor: label, catAxisLabelFontSize: 12, valAxisLabelFontSize: 11,
        catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt",
        titleFontFace: "+mn-lt", titleFontSize: 14, titleColor: escuro ? HEX.branco : HEX.deep, showTitle: true,
        valGridLine: { color: escuro ? HEX.linhaEscuro : HEX.linha, size: 0.5 }, catGridLine: { style: "none" },
        catAxisLineColor: escuro ? HEX.linhaEscuro : HEX.linha,
        showValue: true, dataLabelFontSize: 12, dataLabelColor: label, showLegend: false,
      };
    },
  };
  return { pres, S, ui };
}

module.exports = { TK, HEX, THEME, ESC, HEAD, W, H, M, CW, RAIO, LOGO_RATIO, DIR, createPres };
