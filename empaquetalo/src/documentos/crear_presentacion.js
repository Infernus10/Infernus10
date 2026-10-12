// Genera entregables/Avance2_Empaquetalo.pptx con el estilo del deck del Avance 1
// (azul marino, crema y café, motivo de puntos) y los resultados reales del pipeline.
// Uso: node src/documentos/crear_presentacion.js
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const { GiDeliveryDrone } = require("react-icons/gi");
const { MdDeliveryDining } = require("react-icons/md");
const {
  FaTruck, FaDatabase, FaShieldAlt, FaCommentDots, FaKeyboard, FaFileCsv, FaSitemap,
  FaBrain, FaDna, FaDiceThree, FaCalculator, FaCheck, FaCogs, FaSearch,
} = require("react-icons/fa");

const SKILL = process.env.PPTX_SKILL_DIR;
const { applyTheme } = require(path.join(SKILL, "scripts", "apply_theme.js"));

const RAIZ = path.resolve(__dirname, "..", "..");
const bit = JSON.parse(fs.readFileSync(path.join(RAIZ, "data/procesado/bitacora_preprocesamiento.json")));
const R = JSON.parse(fs.readFileSync(path.join(RAIZ, "data/procesado/resultados_viabilidad.json"))).resultados;
const miles = (x) => x.toLocaleString("en-US");
const f2 = (x) => x.toFixed(2);

const HEX = {
  marino: "293D60", oscuro: "1E2235", crema: "F1E7DD", cafe: "B38B5D", naranja: "C8622B",
  blanco: "FFFFFF", gris: "5B6475", dron: "3567B5", moto: "C9702E", carro: "7A5195",
};
const THEME = {
  name: "Empaquetalo",
  headFontFace: "Arial",
  bodyFontFace: "Calibri",
  colors: {
    dk1: HEX.oscuro, lt1: HEX.blanco, dk2: HEX.marino, lt2: HEX.crema,
    accent1: HEX.cafe, accent2: HEX.naranja, accent3: HEX.dron, accent4: HEX.moto,
    accent5: HEX.carro, accent6: HEX.gris, hlink: HEX.dron, folHlink: HEX.carro,
  },
};

async function icono(Componente, color, px = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Componente, { color: `#${color}`, size: px }));
  const png = await sharp(Buffer.from(svg)).resize(px, px).png().toBuffer();
  return "image/png;base64," + png.toString("base64");
}

async function main() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9"; // 10 x 5.625 in
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  pres.title = "Empaquétalo – Avance 2";
  pres.author = "Equipo Empaquétalo";
  const C = pres.SchemeColor;

  const I = {
    dron: await icono(GiDeliveryDrone, HEX.blanco), moto: await icono(MdDeliveryDining, HEX.blanco),
    carro: await icono(FaTruck, HEX.blanco), db: await icono(FaDatabase, HEX.marino), dbB: await icono(FaDatabase, HEX.blanco),
    escudo: await icono(FaShieldAlt, HEX.blanco), coment: await icono(FaCommentDots, HEX.blanco),
    teclado: await icono(FaKeyboard, HEX.blanco), csv: await icono(FaFileCsv, HEX.marino),
    arbol: await icono(FaSitemap, HEX.blanco), arbolM: await icono(FaSitemap, HEX.marino),
    cerebro: await icono(FaBrain, HEX.blanco), adn: await icono(FaDna, HEX.blanco),
    dado: await icono(FaDiceThree, HEX.blanco), calc: await icono(FaCalculator, HEX.blanco),
    check: await icono(FaCheck, HEX.blanco), cogs: await icono(FaCogs, HEX.blanco),
    lupa: await icono(FaSearch, HEX.marino),
  };

  // ---------- layouts ----------
  const titulo = (color) => ({
    placeholder: {
      options: { name: "title", type: "title", x: 0.5, y: 0.32, w: 9, h: 0.62, fontFace: "Arial", fontSize: 26,
        bold: true, color, charSpacing: 5, align: "left", valign: "middle", margin: 0 },
      text: "TÍTULO",
    },
  });
  const pie = (color) => ({ text: { text: "Empaquétalo · Avance 2", options: { x: 0.5, y: 5.25, w: 4, h: 0.25,
    fontSize: 9, color, fontFace: "Calibri", margin: 0 } } });
  pres.defineSlideMaster({
    title: "PORTADA", background: { color: HEX.marino },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 1.55, w: 6.2, h: 0.9, fontFace: "Arial", fontSize: 40,
        bold: true, color: HEX.blanco, charSpacing: 12, margin: 0, valign: "middle", align: "left" }, text: "TÍTULO" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 2.5, w: 6.0, h: 0.9, fontFace: "Arial", fontSize: 13,
        color: HEX.crema, charSpacing: 3, margin: 0, valign: "top", align: "left" }, text: "Subtítulo" } },
    ],
  });
  pres.defineSlideMaster({
    title: "OSCURO", background: { color: HEX.marino },
    objects: [titulo(HEX.blanco), pie(HEX.crema)],
    slideNumber: { x: 9.1, y: 5.22, w: 0.4, h: 0.28, fontSize: 9, color: HEX.crema, align: "right" },
  });
  pres.defineSlideMaster({
    title: "CLARO", background: { color: HEX.crema },
    objects: [titulo(HEX.marino), pie(HEX.gris)],
    slideNumber: { x: 9.1, y: 5.22, w: 0.4, h: 0.28, fontSize: 9, color: HEX.gris, align: "right" },
  });

  // ---------- utilidades de composición ----------
  const puntos = (s, x, y, n, color = C.accent1) => {
    for (let i = 0; i < n; i++) {
      s.addShape(pres.shapes.OVAL, { x: x + i * 0.22, y, w: 0.11, h: 0.11, fill: { color }, line: { type: "none" }, objectName: `punto ${i + 1}` });
    }
  };
  const circuloIcono = (s, img, x, y, d, fondo, nombre) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fondo }, line: { type: "none" }, objectName: `fondo ${nombre}` });
    s.addImage({ data: img, x: x + d * 0.24, y: y + d * 0.24, w: d * 0.52, h: d * 0.52, objectName: `icono ${nombre}`, altText: nombre });
  };
  const tarjeta = (s, x, y, w, h, color, nombre) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, rectRadius: 0.08, fill: { color }, line: { type: "none" }, objectName: nombre,
    shadow: { type: "outer", color: "000000", opacity: 0.12, blur: 6, offset: 2, angle: 90 },
  });
  const texto = (s, t, op) => s.addText(t, { isTextBox: true, fontFace: "Calibri", margin: 0, valign: "top", ...op });

  // ===== 1. Portada =====
  pres.addSection({ title: "Introducción" });
  let s = pres.addSlide({ masterName: "PORTADA", sectionTitle: "Introducción" });
  s.addText("EMPAQUÉTALO", { placeholder: "title" });
  s.addText("AVANCE 2 · RECOPILACIÓN DE DATOS Y ELECCIÓN DEL PARADIGMA", { placeholder: "body" });
  puntos(s, 0.62, 3.45, 10);
  s.addShape(pres.shapes.OVAL, { x: 7.5, y: -0.9, w: 3.4, h: 3.4, fill: { color: C.background2 }, line: { type: "none" }, objectName: "círculo crema" });
  s.addShape(pres.shapes.OVAL, { x: 7.1, y: 1.65, w: 1.0, h: 1.0, fill: { color: C.accent1 }, line: { type: "none" }, objectName: "círculo café" });
  s.addShape(pres.shapes.OVAL, { x: 8.6, y: 3.6, w: 2.0, h: 2.0, fill: { color: C.accent1 }, line: { type: "none" }, objectName: "círculo café grande" });
  texto(s, "Instituto Tecnológico de Morelia · Inteligencia Artificial · Semana 8", { x: 0.6, y: 4.55, w: 6.5, h: 0.3, fontSize: 11, color: C.background2 });
  texto(s, "Prof. Jonathan Zavala Díaz", { x: 0.6, y: 4.85, w: 6.5, h: 0.3, fontSize: 11, color: C.background2 });
  s.addNotes("Presentamos el Avance 2 de Empaquétalo: de dónde salen los datos, cómo los preparamos, qué paradigma de IA elegimos y qué modelo proponemos.");

  // ===== 2. Ruta =====
  s = pres.addSlide({ masterName: "CLARO", sectionTitle: "Introducción" });
  s.addText("RUTA DEL AVANCE", { placeholder: "title" });
  puntos(s, 7.5, 0.58, 8);
  const pasos = [
    ["Problema", "Elegir dron, moto o carro/camión para cada paquete", I.teclado],
    ["Datos", `${miles(bit.filas_limpias)} envíos simulados, limpios y etiquetados`, I.dbB],
    ["Paradigma", "Simbólico: reglas aprendidas con árboles de decisión", I.arbol],
    ["Modelo", "Random Forest + filtro de seguridad + explicación", I.cogs],
  ];
  pasos.forEach(([t, d, img], i) => {
    const x = 0.5 + i * 2.32;
    tarjeta(s, x, 1.5, 2.05, 2.75, C.background1, `tarjeta ruta ${t}`);
    s.addShape(pres.shapes.OVAL, { x: x + 0.65, y: 1.75, w: 0.75, h: 0.75, fill: { color: C.text2 }, line: { type: "none" }, objectName: `fondo icono ${t}` });
    s.addImage({ data: img, x: x + 0.83, y: 1.93, w: 0.39, h: 0.39, objectName: `icono ${t}` });
    texto(s, `${i + 1}`, { x: x + 0.2, y: 2.62, w: 1.65, h: 0.35, fontSize: 14, bold: true, color: C.accent2, align: "center" });
    texto(s, t, { x: x + 0.15, y: 2.95, w: 1.75, h: 0.35, fontSize: 16, bold: true, color: C.text2, align: "center" });
    texto(s, d, { x: x + 0.15, y: 3.35, w: 1.75, h: 0.8, fontSize: 11.5, color: C.text1, align: "center" });
    if (i < 3) s.addShape(pres.shapes.OVAL, { x: x + 2.12, y: 2.83, w: 0.09, h: 0.09, fill: { color: C.accent1 }, line: { type: "none" }, objectName: `conector ${i + 1}` });
  });
  texto(s, "Pregunta de la revisión: ¿los datos permiten resolver el problema y el paradigma elegido es adecuado?",
    { x: 0.5, y: 4.5, w: 9, h: 0.4, fontSize: 13, italic: true, color: C.text2 });

  // ===== 3. Fuente y tipo =====
  pres.addSection({ title: "Datos" });
  s = pres.addSlide({ masterName: "OSCURO", sectionTitle: "Datos" });
  s.addText("FUENTE Y TIPO DE DATOS", { placeholder: "title" });
  tarjeta(s, 0.5, 1.25, 4.3, 3.65, C.background2, "tarjeta fuente");
  s.addImage({ data: I.db, x: 0.75, y: 1.5, w: 0.5, h: 0.5, objectName: "icono base de datos" });
  texto(s, "Fuente: dataset sintético", { x: 1.4, y: 1.55, w: 3.3, h: 0.4, fontSize: 16, bold: true, color: C.text2, valign: "middle" });
  texto(s, [
    { text: "La paquetería no tiene historial de envíos y el alcance excluye datos reales.", options: { bullet: true, breakLine: true } },
    { text: "Simulador propio en Python con límites investigados en el Avance 1.", options: { bullet: true, breakLine: true } },
    { text: "Etiquetas asignadas por una política de costo, tiempo y seguridad.", options: { bullet: true, breakLine: true } },
    { text: "Semilla fija (42): cualquiera puede regenerar los mismos datos.", options: { bullet: true } },
  ], { x: 0.75, y: 2.2, w: 3.85, h: 2.5, fontSize: 13, color: C.text1, paraSpaceAfter: 6 });
  const tipos = [
    ["Tabulares", "Una fila por envío, una columna por variable"],
    ["Numéricos continuos", "Peso, dimensiones, distancia, viento"],
    ["Categórico nominal", "Clima: despejado, nublado, lluvia, tormenta"],
    ["Etiquetados · 3 clases", "Aprendizaje supervisado multiclase"],
  ];
  tipos.forEach(([t, d], i) => {
    const y = 1.25 + i * 0.93;
    tarjeta(s, 5.1, y, 4.4, 0.78, C.text1, `tarjeta tipo ${t}`);
    texto(s, t, { x: 5.3, y: y + 0.1, w: 4.0, h: 0.3, fontSize: 14, bold: true, color: C.accent2 });
    texto(s, d, { x: 5.3, y: y + 0.42, w: 4.0, h: 0.3, fontSize: 12, color: C.background1 });
  });

  // ===== 4. Cantidad =====
  s = pres.addSlide({ masterName: "CLARO", sectionTitle: "Datos" });
  s.addText("CANTIDAD", { placeholder: "title" });
  puntos(s, 7.5, 0.58, 8);
  const stats = [
    [miles(bit.filas_crudas), "registros crudos"],
    [miles(bit.filas_limpias), "envíos válidos"],
    [`${miles(bit.entrenamiento)} / ${miles(bit.prueba)}`, "entrenamiento / prueba (80/20)"],
  ];
  stats.forEach(([n, l], i) => {
    const y = 1.2 + i * 1.25;
    texto(s, n, { x: 0.5, y, w: 3.6, h: 0.7, fontSize: i === 2 ? 30 : 40, bold: true, color: i === 1 ? C.accent2 : C.text2, fontFace: "Arial", valign: "bottom" });
    texto(s, l, { x: 0.5, y: y + 0.72, w: 3.6, h: 0.3, fontSize: 13, color: C.text1 });
  });
  tarjeta(s, 4.4, 1.15, 5.1, 3.8, C.background1, "tarjeta gráfica clases");
  const cl = bit.clases;
  s.addChart(pres.charts.BAR, [{ name: "Envíos", labels: ["Dron", "Moto", "Carro/camión"], values: [cl.dron, cl.moto, cl["carro/camion"]] }], {
    x: 4.55, y: 1.25, w: 4.8, h: 3.0, barDir: "bar", chartColors: [HEX.dron, HEX.moto, HEX.carro],
    showTitle: true, title: "Envíos por medio de transporte", titleFontSize: 13, titleColor: HEX.marino, titleFontFace: "+mn-lt",
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 11, dataLabelColor: HEX.oscuro, dataLabelFontFace: "+mn-lt",
    catAxisLabelColor: HEX.gris, catAxisLabelFontSize: 12, catAxisLabelFontFace: "+mn-lt", catAxisOrientation: "maxMin",
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false, barGapWidthPct: 45,
    valAxisMaxVal: 2800,
  });
  texto(s, "El dron es la clase minoritaria (16%), como en la realidad. Por eso la métrica principal es el F1-macro.",
    { x: 4.65, y: 4.25, w: 4.6, h: 0.6, fontSize: 12, color: C.text1, italic: true });

  // ===== 5. Variables =====
  s = pres.addSlide({ masterName: "OSCURO", sectionTitle: "Datos" });
  s.addText("VARIABLES Y CARACTERÍSTICAS", { placeholder: "title" });
  const enc = (t) => ({ text: t, options: { bold: true, color: HEX.blanco, fill: { color: HEX.oscuro } } });
  const fila = (arr) => arr.map((t) => ({ text: t, options: { color: HEX.oscuro, fill: { color: HEX.crema } } }));
  s.addTable([
    [enc("Variable"), enc("Tipo"), enc("Unidad · rango")],
    fila(["Peso", "Numérica", "kg · 0.05 – 600"]),
    fila(["Largo, ancho, alto", "Numéricas", "cm · 0.5 – 180"]),
    fila(["Distancia", "Numérica", "km · 0.5 – 120"]),
    fila(["Clima", "Categórica", "4 categorías"]),
    fila(["Viento", "Numérica", "km/h · 0 – 88"]),
    [{ text: "Medio (objetivo)", options: { bold: true, color: HEX.blanco, fill: { color: HEX.naranja } } },
      { text: "Categórica", options: { color: HEX.blanco, fill: { color: HEX.naranja } } },
      { text: "dron · moto · carro", options: { color: HEX.blanco, fill: { color: HEX.naranja } } }],
  ], { x: 0.5, y: 1.25, w: 4.9, colW: [1.75, 1.25, 1.9], fontSize: 12, fontFace: "Calibri", rowH: 0.42,
    border: { type: "solid", pt: 1, color: HEX.marino }, valign: "middle", margin: [0, 0.08, 0, 0.08] });
  tarjeta(s, 5.7, 1.25, 3.8, 3.65, C.text1, "tarjeta características derivadas");
  texto(s, "CARACTERÍSTICAS DERIVADAS", { x: 5.95, y: 1.45, w: 3.4, h: 0.3, fontSize: 12, bold: true, color: C.accent2, charSpacing: 2 });
  texto(s, [
    { text: "Dimensiones ordenadas", options: { bold: true, breakLine: true } },
    { text: "mayor, media y menor: el paquete cabe sin importar cómo se capturó", options: { breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "Volumen en litros", options: { bold: true, breakLine: true } },
    { text: "resume el tamaño en un número", options: { breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "Clima one-hot", options: { bold: true, breakLine: true } },
    { text: "4 columnas binarias", options: {} },
  ], { x: 5.95, y: 1.85, w: 3.35, h: 2.2, fontSize: 12.5, color: C.background1 });
  texto(s, "11 características de entrada", { x: 5.95, y: 4.3, w: 3.4, h: 0.4, fontSize: 16, bold: true, color: C.accent1 });

  // ===== 6. Ejemplos =====
  s = pres.addSlide({ masterName: "CLARO", sectionTitle: "Datos" });
  s.addText("EJEMPLOS DE DATOS", { placeholder: "title" });
  puntos(s, 7.5, 0.58, 8);
  const ej = [
    ["ENV-03972", "0.28", "21 × 23 × 3", "13.0", "despejado", "0.7", "Dron", "Sobre ligero, sin viento", HEX.dron],
    ["ENV-01703", "2.29", "32 × 29 × 6", "8.4", "nublado", "32.9", "Moto", "Viento ≥ 30 descarta el dron", HEX.moto],
    ["ENV-00228", "2.78", "24 × 25 × 12", "8.2", "despejado", "11.3", "Moto", "Excede 2.5 kg por 280 g", HEX.moto],
    ["ENV-02270", "12.00", "144 × 84 × 67", "47.2", "despejado", "0.6", "Carro/camión", "Volumen de 805 L", HEX.carro],
    ["ENV-04130", "22.42", "69 × 63 × 69", "4.5", "lluvia", "23.2", "Carro/camión", "Pesado y con lluvia", HEX.carro],
  ];
  const encC = (t) => ({ text: t, options: { bold: true, color: HEX.blanco, fill: { color: HEX.marino } } });
  s.addTable([
    ["ID", "Peso kg", "Dimensiones cm", "Dist. km", "Clima", "Viento", "Medio", "Por qué"].map(encC),
    ...ej.map((r) => [...r.slice(0, 6).map((t) => ({ text: t, options: { color: HEX.oscuro, fill: { color: HEX.blanco } } })),
      { text: r[6], options: { bold: true, color: HEX.blanco, fill: { color: r[8] } } },
      { text: r[7], options: { color: HEX.oscuro, fill: { color: HEX.blanco } } }]),
  ], { x: 0.5, y: 1.2, w: 9.0, colW: [1.0, 0.75, 1.35, 0.75, 0.95, 0.7, 1.15, 2.35], fontSize: 11.5, fontFace: "Calibri",
    rowH: 0.42, border: { type: "solid", pt: 1, color: HEX.crema }, valign: "middle", margin: [0, 0.06, 0, 0.06] });
  texto(s, "ENV-01703 y ENV-00228 son paquetes pequeños a 8 km que una regla simple mandaría en dron. Uno lo impide el viento y el otro 280 gramos.",
    { x: 0.5, y: 4.0, w: 9.0, h: 0.6, fontSize: 13, color: C.text2, italic: true });

  // ===== 7. Formato y adquisición =====
  s = pres.addSlide({ masterName: "OSCURO", sectionTitle: "Datos" });
  s.addText("FORMATO Y ADQUISICIÓN", { placeholder: "title" });
  tarjeta(s, 0.5, 1.25, 2.9, 3.65, C.background2, "tarjeta formato");
  s.addImage({ data: I.csv, x: 0.75, y: 1.45, w: 0.5, h: 0.5, objectName: "icono CSV" });
  texto(s, "Formato", { x: 1.4, y: 1.5, w: 1.8, h: 0.4, fontSize: 16, bold: true, color: C.text2, valign: "middle" });
  texto(s, [
    { text: "CSV, UTF-8, con encabezado", options: { bullet: true, breakLine: true } },
    { text: "Una fila por envío", options: { bullet: true, breakLine: true } },
    { text: "Archivo crudo, limpio y conjuntos de entrenamiento y prueba", options: { bullet: true } },
  ], { x: 0.75, y: 2.15, w: 2.5, h: 2.4, fontSize: 13, color: C.text1, paraSpaceAfter: 8 });
  const adq = [
    ["1", "Generar paquetes", "Tipo (sobre, pequeño, mediano, grande) → peso y dimensiones. Distancia log-normal (mediana 9 km). Viento según el clima."],
    ["2", "Etiquetar", "Se descartan los medios que violan un límite y se elige el de menor costo-tiempo-riesgo, más ruido (σ = 7)."],
    ["3", "Agregar errores de captura", "Clima mal escrito, faltantes, milímetros en vez de centímetros, pesos imposibles y duplicados."],
  ];
  adq.forEach(([n, t, d], i) => {
    const y = 1.25 + i * 1.25;
    s.addShape(pres.shapes.OVAL, { x: 3.75, y: y + 0.1, w: 0.55, h: 0.55, fill: { color: C.accent1 }, line: { type: "none" }, objectName: `paso ${n}` });
    texto(s, n, { x: 3.75, y: y + 0.1, w: 0.55, h: 0.55, fontSize: 18, bold: true, color: C.background1, align: "center", valign: "middle" });
    texto(s, t, { x: 4.5, y: y + 0.05, w: 5.0, h: 0.35, fontSize: 15, bold: true, color: C.accent2 });
    texto(s, d, { x: 4.5, y: y + 0.42, w: 5.0, h: 0.7, fontSize: 12, color: C.background1 });
  });

  // ===== 8. Preprocesamiento =====
  s = pres.addSlide({ masterName: "CLARO", sectionTitle: "Datos" });
  s.addText("PREPROCESAMIENTO INICIAL", { placeholder: "title" });
  const pre = [
    [`${bit.duplicados_eliminados}`, "duplicados eliminados"],
    [`${bit.variantes_clima_antes} → ${bit.variantes_clima_despues}`, "variantes de clima unificadas"],
    [`${bit.eliminados_sin_peso_o_clima}`, "sin peso o clima, eliminados"],
    [`${bit.eliminados_peso_invalido}`, "pesos ≤ 0, eliminados"],
    [`${bit.largo_corregido_mm_a_cm}`, "largos en mm corregidos"],
    [`${bit.viento_imputado}`, "vientos imputados con la mediana del clima"],
  ];
  pre.forEach(([n, l], i) => {
    const x = 0.5 + (i % 3) * 3.05, y = 1.2 + Math.floor(i / 3) * 1.3;
    tarjeta(s, x, y, 2.85, 1.1, C.background1, `tarjeta limpieza ${i + 1}`);
    texto(s, n, { x: x + 0.2, y: y + 0.12, w: 2.5, h: 0.5, fontSize: 26, bold: true, color: C.accent2, fontFace: "Arial" });
    texto(s, l, { x: x + 0.2, y: y + 0.65, w: 2.5, h: 0.35, fontSize: 12, color: C.text1 });
  });
  tarjeta(s, 0.5, 3.85, 9.0, 1.05, C.text2, "tarjeta transformación");
  texto(s, [
    { text: "Después: ", options: { bold: true, color: HEX.cafe } },
    { text: "dimensiones ordenadas + volumen + one-hot del clima → división 80/20 estratificada → StandardScaler ajustado sólo con entrenamiento (para KNN y la red neuronal; los árboles no lo necesitan).", options: { color: HEX.blanco } },
  ], { x: 0.75, y: 3.98, w: 8.5, h: 0.8, fontSize: 13, valign: "middle" });

  // ===== 9. Exploración =====
  s = pres.addSlide({ masterName: "OSCURO", sectionTitle: "Datos" });
  s.addText("LO QUE DICEN LOS DATOS", { placeholder: "title" });
  tarjeta(s, 0.5, 1.2, 5.6, 3.75, C.background1, "marco figura");
  s.addImage({ path: path.join(RAIZ, "figuras", "peso_vs_distancia.png"), x: 0.6, y: 1.28, w: 5.4, h: 3.59,
    objectName: "figura peso contra distancia", altText: "Dispersión de peso contra distancia coloreada por medio" });
  const hallazgos = [
    ["Tormenta", "100% carro/camión"],
    ["Lluvia", "nunca dron"],
    ["Buen clima", "ningún medio domina"],
  ];
  hallazgos.forEach(([t, d], i) => {
    const y = 1.2 + i * 0.85;
    tarjeta(s, 6.4, y, 3.1, 0.7, C.text1, `hallazgo ${t}`);
    texto(s, t, { x: 6.6, y: y + 0.08, w: 2.8, h: 0.28, fontSize: 13, bold: true, color: C.accent2 });
    texto(s, d, { x: 6.6, y: y + 0.36, w: 2.8, h: 0.28, fontSize: 12, color: C.background1 });
  });
  texto(s, "Bajo 2.5 kg conviven dron y moto, y entre 2.5 y 20 kg conviven moto y carro. La frontera depende del viento, el volumen y la distancia a la vez.",
    { x: 6.4, y: 3.85, w: 3.1, h: 1.1, fontSize: 12, color: C.background2, italic: true });

  // ===== 10. Paradigmas =====
  pres.addSection({ title: "Paradigma" });
  s = pres.addSlide({ masterName: "CLARO", sectionTitle: "Paradigma" });
  s.addText("ELECCIÓN DEL PARADIGMA", { placeholder: "title" });
  const par = [
    ["Simbólico", "Reglas si-entonces aprendidas con árboles", "ELEGIDO", I.arbol, true],
    ["Conexionista", `Red neuronal: F1 ${f2(R["Red neuronal (MLP)"].f1_prueba)}, pero es caja negra`, "Comparación", I.cerebro, false],
    ["Bioinspirado", "Optimiza; no clasifica directamente", "Descartado", I.adn, false],
    ["Probabilístico", "Naive Bayes ignora las interacciones", "Descartado", I.dado, false],
    ["Computacional", `KNN: F1 ${f2(R["KNN (k=5)"].f1_prueba)}, sensible a la escala`, "Comparación", I.calc, false],
  ];
  par.forEach(([t, d, v, img, elegido], i) => {
    const x = 0.5 + i * 1.82, w = 1.65;
    tarjeta(s, x, 1.2, w, 3.4, elegido ? C.text2 : C.background1, `tarjeta paradigma ${t}`);
    circuloIcono(s, img, x + 0.45, 1.42, 0.75, elegido ? C.accent2 : C.text2, t);
    texto(s, t, { x: x + 0.1, y: 2.3, w: w - 0.2, h: 0.35, fontSize: 13.5, bold: true, color: elegido ? C.background1 : C.text2, align: "center" });
    texto(s, d, { x: x + 0.12, y: 2.7, w: w - 0.24, h: 1.15, fontSize: 11.5, color: elegido ? C.background2 : C.text1, align: "center" });
    texto(s, v, { x: x + 0.1, y: 4.05, w: w - 0.2, h: 0.35, fontSize: 12, bold: true, color: elegido ? C.accent1 : C.accent6, align: "center", charSpacing: elegido ? 3 : 0 });
  });

  // ===== 11. Por qué simbólico =====
  s = pres.addSlide({ masterName: "OSCURO", sectionTitle: "Paradigma" });
  s.addText("¿POR QUÉ SIMBÓLICO?", { placeholder: "title" });
  const razones = [
    ["Coincide con el problema", "Los límites son umbrales (2.5 kg, 30 km/h, 15 km) y cada nodo de un árbol es un umbral."],
    ["Permite explicar", "Las reglas que sigue un árbol se pueden leer; los pesos de una red neuronal, no."],
    ["Habla el idioma del filtro", "El filtro de seguridad también son reglas: modelo y filtro se integran."],
    ["Lo respalda la evidencia", `Los árboles quedaron primero sin ajuste: Random Forest ${f2(R["Random Forest"].f1_prueba)}, árbol ${f2(R["Árbol de decisión"].f1_prueba)}.`],
  ];
  razones.forEach(([t, d], i) => {
    const x = 0.5 + (i % 2) * 4.6, y = 1.2 + Math.floor(i / 2) * 1.55;
    tarjeta(s, x, y, 4.4, 1.35, C.background2, `razón ${i + 1}`);
    texto(s, `${i + 1}`, { x: x + 0.2, y: y + 0.15, w: 0.5, h: 0.5, fontSize: 28, bold: true, color: C.accent2, fontFace: "Arial" });
    texto(s, t, { x: x + 0.8, y: y + 0.17, w: 3.45, h: 0.35, fontSize: 15, bold: true, color: C.text2 });
    texto(s, d, { x: x + 0.8, y: y + 0.55, w: 3.45, h: 0.75, fontSize: 12, color: C.text1 });
  });
  texto(s, "No son reglas escritas a mano: el modelo induce cientos de reglas y sus umbrales a partir de 4,087 ejemplos.",
    { x: 0.5, y: 4.4, w: 9.0, h: 0.45, fontSize: 13, italic: true, color: C.accent1 });

  // ===== 12. Modelo =====
  pres.addSection({ title: "Modelo" });
  s = pres.addSlide({ masterName: "CLARO", sectionTitle: "Modelo" });
  s.addText("MODELO PROPUESTO", { placeholder: "title" });
  puntos(s, 7.5, 0.58, 8);
  const flujo = [
    ["Entrada", "peso, medidas, distancia, clima, viento", I.teclado],
    ["Preproceso", "ordenar medidas, volumen, one-hot", I.cogs],
    ["Random Forest", "vota y da la probabilidad de cada medio", I.arbol],
    ["Filtro", "descarta medios que violan un límite", I.escudo],
    ["Salida", "medio + confianza + explicación", I.coment],
  ];
  flujo.forEach(([t, d, img], i) => {
    const x = 0.5 + i * 1.82;
    circuloIcono(s, img, x + 0.42, 1.2, 0.8, i === 2 ? C.accent2 : C.text2, t);
    texto(s, t, { x, y: 2.1, w: 1.65, h: 0.33, fontSize: 13.5, bold: true, color: C.text2, align: "center" });
    texto(s, d, { x: x + 0.05, y: 2.45, w: 1.55, h: 0.7, fontSize: 11, color: C.text1, align: "center" });
    if (i < 4) puntos(s, x + 1.33, 1.55, 2, C.accent1);
  });
  tarjeta(s, 0.5, 3.35, 9.0, 1.55, C.text1, "tarjeta ejemplo");
  texto(s, "EJEMPLO REAL DEL MODELO PRELIMINAR", { x: 0.75, y: 3.48, w: 8.5, h: 0.28, fontSize: 11, bold: true, color: C.accent1, charSpacing: 2 });
  texto(s, [
    { text: "1.8 kg · 30×20×15 cm · 7.5 km · lluvia · 25 km/h  →  ", options: { color: HEX.blanco } },
    { text: "Moto (0.88)", options: { bold: true, color: HEX.moto } },
  ], { x: 0.75, y: 3.8, w: 8.5, h: 0.32, fontSize: 14 });
  texto(s, "«La lluvia descarta el dron; paquete pequeño y distancia media favorecen la moto.»",
    { x: 0.75, y: 4.13, w: 8.5, h: 0.3, fontSize: 12, italic: true, color: C.background2 });
  texto(s, [
    { text: "Mismo paquete con cielo despejado y 10 km/h  →  ", options: { color: HEX.blanco } },
    { text: "Dron (0.60)", options: { bold: true, color: "8FB0E8" } },
  ], { x: 0.75, y: 4.48, w: 8.5, h: 0.3, fontSize: 13 });

  // ===== 13. Viabilidad =====
  pres.addSection({ title: "Revisión" });
  s = pres.addSlide({ masterName: "OSCURO", sectionTitle: "Revisión" });
  s.addText("PRUEBA DE VIABILIDAD", { placeholder: "title" });
  tarjeta(s, 0.5, 1.2, 5.9, 3.75, C.background1, "tarjeta gráfica F1");
  const nombres = ["Sistema base (reglas)", "Árbol de decisión", "KNN (k=5)", "Random Forest", "Red neuronal (MLP)"];
  s.addChart(pres.charts.BAR, [{ name: "F1-macro", labels: nombres, values: nombres.map((n) => +R[n].f1_prueba.toFixed(2)) }], {
    x: 0.6, y: 1.3, w: 5.7, h: 3.2, barDir: "bar",
    chartColors: ["9AA3B2", HEX.marino, HEX.marino, HEX.naranja, HEX.marino], invertedColors: ["9AA3B2", HEX.marino, HEX.marino, HEX.naranja, HEX.marino],
    showTitle: true, title: "F1-macro en prueba (sin ajustar)", titleFontSize: 13, titleColor: HEX.marino, titleFontFace: "+mn-lt",
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.00", dataLabelFontSize: 11, dataLabelColor: HEX.oscuro, dataLabelFontFace: "+mn-lt",
    catAxisLabelColor: HEX.gris, catAxisLabelFontSize: 11, catAxisLabelFontFace: "+mn-lt", catAxisOrientation: "maxMin",
    valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 1.1, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    showLegend: false, barGapWidthPct: 50,
  });
  texto(s, "Meta del Avance 1: F1-macro ≥ 0.85 y +10 puntos sobre las reglas", { x: 0.7, y: 4.5, w: 5.5, h: 0.3, fontSize: 11, italic: true, color: C.accent6 });
  const mejora = Math.round((R["Random Forest"].f1_prueba - R["Sistema base (reglas)"].f1_prueba) * 100);
  texto(s, `+${mejora}`, { x: 6.7, y: 1.2, w: 2.8, h: 0.95, fontSize: 60, bold: true, color: C.accent2, fontFace: "Arial", valign: "bottom" });
  texto(s, "puntos de F1-macro sobre el sistema de reglas simples", { x: 6.7, y: 2.2, w: 2.8, h: 0.6, fontSize: 13, color: C.background1 });
  texto(s, `${f2(R["Random Forest"].f1_cv)} ± ${f2(R["Random Forest"].f1_cv_std)}`, { x: 6.7, y: 3.05, w: 2.8, h: 0.6, fontSize: 28, bold: true, color: C.accent1, fontFace: "Arial", valign: "bottom" });
  texto(s, "F1-macro con validación cruzada de 5 pliegues: resultado estable", { x: 6.7, y: 3.7, w: 2.8, h: 0.6, fontSize: 13, color: C.background1 });

  // ===== 14. Respuesta =====
  s = pres.addSlide({ masterName: "CLARO", sectionTitle: "Revisión" });
  s.addText("RESPUESTA A LA REVISIÓN", { placeholder: "title" });
  const resp = [
    ["¿Los datos permiten resolver el problema?", `Sí. F1-macro de ${f2(R["Random Forest"].f1_prueba)} sin ajuste, arriba de la meta de 0.85.`],
    ["¿El paradigma es adecuado?", "Sí. Los árboles quedan primero y además son explicables."],
  ];
  resp.forEach(([q, a], i) => {
    const y = 1.2 + i * 1.05;
    tarjeta(s, 0.5, y, 9.0, 0.88, C.background1, `respuesta ${i + 1}`);
    circuloIcono(s, I.check, 0.7, y + 0.14, 0.6, C.accent2, `sí ${i + 1}`);
    texto(s, q, { x: 1.5, y: y + 0.1, w: 7.8, h: 0.32, fontSize: 15, bold: true, color: C.text2 });
    texto(s, a, { x: 1.5, y: y + 0.45, w: 7.8, h: 0.32, fontSize: 13, color: C.text1 });
  });
  tarjeta(s, 0.5, 3.35, 4.4, 1.55, C.text2, "tarjeta atención");
  texto(s, "Por atender", { x: 0.75, y: 3.48, w: 4.0, h: 0.3, fontSize: 14, bold: true, color: C.accent1 });
  texto(s, `El dron es la clase más difícil (F1 ${f2(R["Random Forest"].f1_dron)}). Se probará class_weight = "balanced" y el filtro bloquea cualquier dron inseguro.`,
    { x: 0.75, y: 3.82, w: 3.95, h: 1.0, fontSize: 12, color: C.background1 });
  tarjeta(s, 5.1, 3.35, 4.4, 1.55, C.text2, "tarjeta limitación");
  texto(s, "Limitación", { x: 5.35, y: 3.48, w: 4.0, h: 0.3, fontSize: 14, bold: true, color: C.accent1 });
  texto(s, "Los datos son sintéticos: el modelo aprende la política del equipo. El ruido evita que memorice reglas (nadie llega a 1.0).",
    { x: 5.35, y: 3.82, w: 3.95, h: 1.0, fontSize: 12, color: C.background1 });

  // ===== 15. Producto =====
  s = pres.addSlide({ masterName: "OSCURO", sectionTitle: "Revisión" });
  s.addText("PRODUCTO DE LA REVISIÓN", { placeholder: "title" });
  const prod = [
    ["Dataset", `${miles(bit.filas_limpias)} envíos limpios, reproducibles`],
    ["Ejemplos", "Envíos comentados de las 3 clases"],
    ["Preprocesamiento", "9 pasos con bitácora de registros"],
    ["Paradigma", "Simbólico, contra los otros 4"],
    ["Modelo", "Random Forest + filtro + explicación"],
  ];
  prod.forEach(([t, d], i) => {
    const y = 1.2 + i * 0.75;
    circuloIcono(s, I.check, 0.6, y + 0.04, 0.5, C.accent1, `entregado ${t}`);
    texto(s, t, { x: 1.3, y: y + 0.02, w: 2.4, h: 0.55, fontSize: 16, bold: true, color: C.background1, valign: "middle" });
    texto(s, d, { x: 3.7, y: y + 0.02, w: 5.8, h: 0.55, fontSize: 14, color: C.background2, valign: "middle" });
  });

  // ===== 16. Cierre =====
  s = pres.addSlide({ masterName: "PORTADA", sectionTitle: "Revisión" });
  s.addText("GRACIAS", { placeholder: "title" });
  s.addText("SIGUIENTE: ENTRENAMIENTO Y AJUSTE DE MODELOS (SEMANAS 9 A 11)", { placeholder: "body" });
  puntos(s, 0.62, 3.45, 10);
  s.addShape(pres.shapes.OVAL, { x: 7.5, y: -0.9, w: 3.4, h: 3.4, fill: { color: C.background2 }, line: { type: "none" }, objectName: "círculo crema" });
  s.addShape(pres.shapes.OVAL, { x: 7.1, y: 1.65, w: 1.0, h: 1.0, fill: { color: C.accent1 }, line: { type: "none" }, objectName: "círculo café" });
  s.addShape(pres.shapes.OVAL, { x: 8.6, y: 3.6, w: 2.0, h: 2.0, fill: { color: C.accent1 }, line: { type: "none" }, objectName: "círculo café grande" });

  const salida = path.join(RAIZ, "entregables", "Avance2_Empaquetalo.pptx");
  await pres.writeFile({ fileName: salida });
  await applyTheme(salida, THEME);
  console.log("Escrito", salida);
}

main().catch((e) => { console.error(e); process.exit(1); });
