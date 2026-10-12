// Genera entregables/Avance2_Empaquetalo.docx a partir de los resultados reales
// del pipeline (data/procesado/*.json) y de las figuras en figuras/.
// Uso: node src/documentos/crear_word.js
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  AlignmentType, HeadingLevel, WidthType, ShadingType, BorderStyle, LevelFormat,
  PageBreak, TableOfContents, Footer, PageNumber, TabStopType,
} = require("docx");

const RAIZ = path.resolve(__dirname, "..", "..");
const FIG = (f) => fs.readFileSync(path.join(RAIZ, "figuras", f));
const bit = JSON.parse(fs.readFileSync(path.join(RAIZ, "data/procesado/bitacora_preprocesamiento.json")));
const res = JSON.parse(fs.readFileSync(path.join(RAIZ, "data/procesado/resultados_viabilidad.json"))).resultados;

const f2 = (x) => x.toFixed(2);
const miles = (x) => x.toLocaleString("en-US");
const pct = (a, b) => `${Math.round((100 * a) / b)}%`;

const AZUL = "1F3864";
const GRIS_TABLA = "D9E2F3";
const ANCHO = 9360; // carta con márgenes de 1"
const FUENTE = "Arial";

// ---------- utilidades de texto ----------
function runs(texto, base = {}) {
  // **negritas** dentro del texto.
  return texto.split(/(\*\*[^*]+\*\*)/).filter(Boolean).map((t) =>
    t.startsWith("**") ? new TextRun({ ...base, text: t.slice(2, -2), bold: true })
      : new TextRun({ ...base, text: t }));
}
const P = (texto, op = {}) => new Paragraph({
  alignment: op.align ?? AlignmentType.JUSTIFIED,
  spacing: { after: op.after ?? 160, line: 300 },
  children: runs(texto, op.run ?? {}),
  ...(op.extra ?? {}),
});
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(t)] });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(t)] });
const H3 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(t)] });
const VIN = (t, nivel = 0) => new Paragraph({
  numbering: { reference: "vinetas", level: nivel }, alignment: AlignmentType.JUSTIFIED,
  spacing: { after: 80, line: 290 }, children: runs(t),
});
const NUM = (t, ref = "numeros") => new Paragraph({
  numbering: { reference: ref, level: 0 }, alignment: AlignmentType.JUSTIFIED,
  spacing: { after: 80, line: 290 }, children: runs(t),
});
const CENTRO = (t, op = {}) => new Paragraph({
  alignment: AlignmentType.CENTER, spacing: { after: op.after ?? 60 },
  children: [new TextRun({ text: t, bold: op.bold, size: op.size ?? 22, font: FUENTE })],
});
const VACIO = (n = 1) => Array.from({ length: n }, () => new Paragraph({ children: [] }));

function figura(archivo, anchoPx, altoPx, pie) {
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER, spacing: { before: 120, after: 60 },
      children: [new ImageRun({
        type: "png", data: FIG(archivo), transformation: { width: anchoPx, height: altoPx },
        altText: { title: pie, description: pie, name: archivo },
      })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER, spacing: { after: 220 },
      children: [new TextRun({ text: pie, italics: true, size: 18, color: "595959" })],
    }),
  ];
}

// ---------- tablas ----------
const borde = { style: BorderStyle.SINGLE, size: 4, color: "8EA0BF" };
const bordes = { top: borde, bottom: borde, left: borde, right: borde };

function tabla(anchos, encabezado, filas, op = {}) {
  const celda = (texto, w, esEnc, i) => new TableCell({
    width: { size: w, type: WidthType.DXA }, borders: bordes,
    shading: esEnc ? { fill: AZUL, type: ShadingType.CLEAR, color: "auto" }
      : (op.cebra && i % 2 === 1 ? { fill: "F2F5FA", type: ShadingType.CLEAR, color: "auto" } : undefined),
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: String(texto).split("\n").map((linea) => new Paragraph({
      alignment: op.centrar && !esEnc ? AlignmentType.CENTER : AlignmentType.LEFT,
      children: runs(linea, { size: op.tam ?? 18, color: esEnc ? "FFFFFF" : "000000", bold: esEnc || undefined }),
    })),
  });
  return new Table({
    width: { size: anchos.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: anchos,
    rows: [
      new TableRow({ tableHeader: true, children: encabezado.map((t, j) => celda(t, anchos[j], true, 0)) }),
      ...filas.map((f, i) => new TableRow({ children: f.map((t, j) => celda(t, anchos[j], false, i)) })),
    ],
  });
}
const PIE_TABLA = (t) => new Paragraph({
  alignment: AlignmentType.CENTER, spacing: { before: 60, after: 220 },
  children: [new TextRun({ text: t, italics: true, size: 18, color: "595959" })],
});

// ---------- portada ----------
const integrantes = [
  "22121299 Figueroa Merino Cesar",
  "22121355 Romero García Emiliano",
  "22121336 Rodríguez Reyna Kelly Merali",
  "22121283 Murillo Rivera Annatar Armando",
  "22121330 Sandoval Mercado Paulo Aquiles",
  "22121338 Parra Fernández Bruno Leonardo",
];
const portada = [
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new ImageRun({
      type: "png", data: FIG("encabezado_logos.png"), transformation: { width: 600, height: 55 },
      altText: { title: "Logos institucionales", description: "SEP, TecNM e Instituto Tecnológico de Morelia", name: "logos" },
    })],
  }),
  ...VACIO(2),
  CENTRO("INSTITUTO TECNOLÓGICO DE MORELIA", { bold: true, size: 28, after: 360 }),
  CENTRO("INTELIGENCIA ARTIFICIAL", { size: 24, after: 480 }),
  CENTRO("ENTREGA 2", { bold: true, size: 28, after: 120 }),
  CENTRO("Semana 8: Avance 2", { bold: true, size: 26, after: 120 }),
  CENTRO("Recopilación de datos y elección del paradigma", { bold: true, size: 26, after: 480 }),
  CENTRO("Proyecto Empaquétalo: clasificación inteligente del medio de envío", { size: 22, after: 480 }),
  CENTRO("INGENIERÍA EN SISTEMAS COMPUTACIONALES", { bold: true, size: 22, after: 480 }),
  CENTRO("PRESENTA:", { size: 20, after: 120 }),
  ...integrantes.map((n) => CENTRO(n, { size: 20 })),
  ...VACIO(1),
  CENTRO("Profesor:", { size: 20, after: 60 }),
  CENTRO("Jonathan Zavala Díaz", { size: 20, after: 600 }),
  new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: ANCHO }],
    children: [
      new TextRun({ text: "MORELIA, MICHOACÁN", size: 20 }),
      new TextRun({ text: "\t12 de octubre de 2026", size: 20, bold: true }),
    ],
  }),
  new Paragraph({ children: [new PageBreak()] }),
];

// ---------- índice ----------
const indice = [
  new Paragraph({ spacing: { after: 240 }, children: [new TextRun({ text: "Índice", size: 32, font: FUENTE })] }),
  new TableOfContents("Índice", { hyperlink: true, headingStyleRange: "1-2" }),
  new Paragraph({ children: [new PageBreak()] }),
];

// ---------- contenido ----------
const R = res;
const base = R["Sistema base (reglas)"];
const rf = R["Random Forest"];
const mejora = Math.round((rf.f1_prueba - base.f1_prueba) * 100);
const total = bit.filas_limpias;
const cl = bit.clases;

const cuerpo = [
  new Paragraph({
    spacing: { after: 200 },
    children: [new TextRun({ text: "Proyecto Empaquétalo: clasificación inteligente del medio de envío de paquetes", bold: true, size: 24 })],
  }),

  H1("Resumen del avance"),
  P("En el Avance 1 se definió el problema: dadas las características de un paquete y las condiciones del envío, clasificar automáticamente el medio de transporte más adecuado (dron, moto o carro/camión). En este segundo avance se recopilaron los datos, se preprocesaron, se eligió el paradigma de Inteligencia Artificial y se propuso el modelo. Este documento sigue la ruta que pide la revisión:"),
  P("**Problema → características de los datos → paradigma → modelo propuesto**", { align: AlignmentType.CENTER, run: { color: AZUL } }),
  tabla([2600, 6760], ["Elemento", "Resultado del Avance 2"], [
    ["Dataset", `${miles(bit.filas_crudas)} registros crudos generados por simulación; ${miles(total)} registros válidos después de la limpieza.`],
    ["Tipo de problema", "Clasificación supervisada multiclase sobre datos tabulares (3 clases)."],
    ["Variables", "5 entradas originales (peso, largo/ancho/alto, distancia, clima, viento) y 11 características finales."],
    ["Paradigma elegido", "Enfoque simbólico: aprendizaje inductivo de reglas mediante árboles de decisión."],
    ["Modelo propuesto", "Random Forest (ensamble de árboles) + filtro de seguridad basado en reglas + módulo de explicación."],
    ["Prueba de viabilidad", `F1-macro de ${f2(rf.f1_prueba)} con Random Forest sin ajustar contra ${f2(base.f1_prueba)} del sistema de reglas simples (+${mejora} puntos).`],
  ], { cebra: true }),
  PIE_TABLA("Tabla 1. Resumen de resultados del Avance 2."),

  // ===== 1. Recopilación de datos =====
  H1("1. Recopilación de datos"),

  H2("1.1 Fuente de los datos"),
  P("La paquetería del caso de estudio no cuenta con registros históricos de envíos y el alcance del proyecto excluye el uso de datos reales de clientes. Por ello, la fuente de datos es un **dataset sintético** generado por el equipo con un simulador escrito en Python (src/generar_dataset.py)."),
  P("El simulador no produce valores al azar sin sentido. Sus distribuciones y límites se basan en la investigación de capacidades realizada en el Avance 1 (carga útil de drones de reparto, cajas de moto y capacidad de camionetas). La etiqueta de cada envío la asigna una política de decisión documentada, construida con criterios de costo, tiempo y seguridad. Esa política representa al despachador experto que el modelo debe aprender a imitar."),
  P("El generador usa una semilla fija (42), por lo que cualquier integrante o el profesor puede regenerar exactamente el mismo dataset."),

  H2("1.2 Tipo de datos"),
  VIN("**Estructurados y tabulares:** una fila por envío y una columna por variable."),
  VIN("**Numéricos continuos:** peso, dimensiones, distancia y velocidad del viento."),
  VIN("**Categórico nominal:** condición climática (despejado, nublado, lluvia, tormenta)."),
  VIN("**Etiquetados:** cada envío trae el medio correcto, así que el problema es de aprendizaje supervisado."),
  VIN("**Variable objetivo categórica de 3 clases:** dron, moto y carro/camión (clasificación multiclase)."),

  H2("1.3 Cantidad"),
  P(`Se simularon 5,200 envíos. Al agregar los errores de captura (sección 1.7) el archivo crudo quedó con ${miles(bit.filas_crudas)} filas. Después de la limpieza quedaron **${miles(total)} envíos válidos**, lo que cumple el objetivo específico 1 del Avance 1 (al menos 5,000 envíos). La sección 4 confirma que esta cantidad basta: los resultados de la validación cruzada varían menos de ±0.015 entre pliegues.`),
  tabla([2340, 2340, 2340, 2340], ["Clase", "Envíos", "Porcentaje", "Comentario"], [
    ["Dron", miles(cl.dron), pct(cl.dron, total), "Clase minoritaria"],
    ["Moto", miles(cl.moto), pct(cl.moto, total), "—"],
    ["Carro/camión", miles(cl["carro/camion"]), pct(cl["carro/camion"], total), "Clase mayoritaria"],
    ["**Total**", `**${miles(total)}**`, "**100%**", ""],
  ], { centrar: true }),
  PIE_TABLA("Tabla 2. Distribución de clases en el dataset limpio."),
  P(`La clase dron es minoritaria (${pct(cl.dron, total)}), igual que en la realidad: sólo aplica a paquetes ligeros, a distancias cortas y con buen clima. Por eso la métrica principal es el F1-macro, que da el mismo peso a las tres clases. La división entrenamiento/prueba se hace estratificada para conservar estas proporciones (${miles(bit.entrenamiento)} envíos para entrenamiento y ${miles(bit.prueba)} para prueba).`),
  ...figura("distribucion_clases.png", 470, 204, "Figura 1. Número de envíos por medio de transporte."),

  H2("1.4 Variables y características"),
  tabla([1900, 1500, 1300, 1600, 3060], ["Variable", "Tipo", "Unidad", "Rango", "Descripción"], [
    ["peso_kg", "Numérica continua", "kg", "0.05 – 600", "Peso del paquete."],
    ["largo_cm, ancho_cm, alto_cm", "Numéricas continuas", "cm", "0.5 – 180", "Dimensiones del paquete capturadas en ventanilla."],
    ["distancia_km", "Numérica continua", "km", "0.5 – 120", "Distancia del origen al destino."],
    ["clima", "Categórica nominal", "—", "4 categorías", "Despejado, nublado, lluvia o tormenta."],
    ["viento_kmh", "Numérica continua", "km/h", "0 – 88", "Velocidad del viento al momento del envío."],
    ["**medio** (objetivo)", "Categórica nominal", "—", "3 clases", "Dron, moto o carro/camión."],
  ], { cebra: true }),
  PIE_TABLA("Tabla 3. Variables originales del dataset."),
  P("A partir de las variables originales se construyen **características derivadas** que facilitan el aprendizaje:"),
  VIN("**dim_mayor_cm, dim_media_cm, dim_menor_cm:** las tres dimensiones ordenadas de mayor a menor. Un paquete cabe en la caja de una moto sin importar cómo se haya capturado (largo × ancho o ancho × largo), así que ordenar las dimensiones elimina esa ambigüedad."),
  VIN("**volumen_l:** largo × ancho × alto / 1000, en litros. Resume el tamaño en un solo número."),
  VIN("**clima_despejado, clima_nublado, clima_lluvia, clima_tormenta:** codificación one-hot del clima, porque los algoritmos trabajan con números."),
  P("En total el modelo recibe **11 características**: 7 numéricas (peso, tres dimensiones ordenadas, volumen, distancia y viento) y 4 binarias del clima."),

  H2("1.5 Ejemplos de datos"),
  P("La Tabla 4 muestra envíos reales del dataset limpio, dos por clase. Los casos se eligieron porque muestran las interacciones que justifican usar IA."),
  tabla([1200, 760, 1500, 900, 1100, 900, 1300, 1700], ["ID", "Peso (kg)", "Dimensiones (cm)", "Dist. (km)", "Clima", "Viento (km/h)", "Medio", "Por qué"], [
    ["ENV-03972", "0.28", "21.3 × 23.0 × 3.1", "12.98", "despejado", "0.7", "Dron", "Sobre ligero, sin viento"],
    ["ENV-03981", "0.05", "29.0 × 19.0 × 3.8", "10.48", "nublado", "21.7", "Dron", "Nublado aún permite volar"],
    ["ENV-01703", "2.29", "31.9 × 28.5 × 5.9", "8.43", "nublado", "32.9", "Moto", "Viento ≥ 30 km/h descarta el dron"],
    ["ENV-00228", "2.78", "23.6 × 25.3 × 11.6", "8.20", "despejado", "11.3", "Moto", "Excede 2.5 kg por muy poco"],
    ["ENV-02270", "12.00", "143.6 × 83.5 × 67.1", "47.23", "despejado", "0.6", "Carro/camión", "Volumen de 805 L"],
    ["ENV-04130", "22.42", "68.9 × 62.6 × 68.7", "4.52", "lluvia", "23.2", "Carro/camión", "Pesado y con lluvia"],
  ], { tam: 16, cebra: true }),
  PIE_TABLA("Tabla 4. Ejemplos del dataset limpio (envios_limpio.csv)."),
  P("Los envíos ENV-01703 y ENV-00228 son buenos ejemplos de por qué una regla simple falla. Ambos son paquetes pequeños a 8 km, que a primera vista irían en dron. Uno se descarta por el viento y el otro porque pesa 280 g más que el límite."),

  H2("1.6 Formato"),
  VIN("**Formato de archivo:** CSV (valores separados por comas), codificación UTF-8, con encabezado en la primera fila."),
  VIN("**Estructura:** una fila por envío; columnas id_envio, peso_kg, largo_cm, ancho_cm, alto_cm, distancia_km, clima, viento_kmh y medio."),
  VIN("**Decimales:** punto decimal; pesos y distancias con 2 decimales, dimensiones y viento con 1."),
  tabla([3600, 5760], ["Archivo", "Contenido"], [
    ["data/envios_crudo.csv", `${miles(bit.filas_crudas)} filas con errores de captura, tal como “llegan”.`],
    ["data/envios_limpio.csv", `${miles(total)} filas limpias con las características derivadas.`],
    ["data/procesado/entrenamiento.csv y prueba.csv", "División 80/20 estratificada, sin escalar (para árboles y Random Forest)."],
    ["data/procesado/*_escalado.csv", "Misma división con variables numéricas estandarizadas (para KNN y la red neuronal)."],
    ["data/procesado/escalador.joblib", "Escalador ajustado, para aplicar la misma transformación en el prototipo."],
  ], { cebra: true }),
  PIE_TABLA("Tabla 5. Archivos del dataset."),

  H2("1.7 Método de adquisición"),
  P("Los datos se obtienen por **simulación** en tres pasos:"),
  NUM("**Generación de paquetes.** Primero se sortea el tipo de paquete: sobre (22%), pequeño (38%), mediano (25%) o grande (15%). Cada tipo tiene sus propios rangos de peso y dimensiones, de modo que no aparecen combinaciones absurdas como un sobre de 80 kg. La distancia sigue una distribución log-normal con mediana de 9 km: hay muchos envíos urbanos cortos y pocos foráneos. El clima se sortea con probabilidades de 45% despejado, 25% nublado, 20% lluvia y 10% tormenta, y el viento depende del clima (en promedio 10 km/h despejado y 48 km/h en tormenta).", "pasos"),
  NUM("**Etiquetado con la política del equipo.** Para cada envío se descartan los medios que violan algún límite físico o climático del Avance 1. Entre los medios restantes se calcula un puntaje de costo, tiempo y riesgo con interacciones no lineales. Por ejemplo, el viento castiga al dron de forma cuadrática, la lluvia castiga más a la moto cuanto más pesado es el paquete y el carro sólo compensa en distancias largas o con paquetes grandes. Se elige el medio de menor puntaje, al que se suma ruido gaussiano (σ = 7) para simular que dos despachadores no siempre deciden igual.", "pasos"),
  NUM("**Errores de captura.** Para que el dataset se parezca a datos reales capturados en ventanilla, se agregan errores en porcentajes controlados: clima escrito de 16 formas distintas (6%), viento faltante (2.5%), clima o peso faltante (1.4%), largo capturado en milímetros (0.6%), pesos imposibles (0.4%) y registros duplicados (1.2%).", "pasos"),
  tabla([2300, 1350, 1350, 1350, 1350, 1660], ["Medio", "Peso máx.", "Tamaño máx. (cm)", "Distancia máx.", "Viento máx.", "Clima permitido"], [
    ["Dron", "2.5 kg", "30 × 30 × 20", "15 km", "< 30 km/h", "Despejado, nublado"],
    ["Moto", "20 kg", "50 × 40 × 40", "40 km", "—", "Todo menos tormenta"],
    ["Carro/camión", "800 kg", "Sin límite práctico", "Sin límite", "—", "Cualquiera"],
  ], { cebra: true }),
  PIE_TABLA("Tabla 6. Límites usados por el etiquetado y por el filtro de seguridad (supuestos del Avance 1)."),

  H2("1.8 Preprocesamiento necesario"),
  P("El script src/preprocesar.py aplica los siguientes pasos en orden. La Tabla 7 resume cuántos registros afectó cada uno."),
  tabla([600, 2900, 4060, 1800], ["#", "Paso", "Decisión y motivo", "Registros"], [
    ["1", "Eliminar duplicados", "Mismo id_envio registrado dos veces; se conserva el primero.", `${bit.duplicados_eliminados} eliminados`],
    ["2", "Normalizar el clima", "Minúsculas, sin acentos ni espacios extra y sinónimos unificados (soleado → despejado, lluvioso → lluvia).", `${bit.variantes_clima_antes} variantes → ${bit.variantes_clima_despues}`],
    ["3", "Registros sin peso o clima", "Se eliminan: son las variables que más pesan en la decisión e imputarlas inventaría el caso.", `${bit.eliminados_sin_peso_o_clima} eliminados`],
    ["4", "Pesos imposibles", "Peso menor o igual a 0 kg; se eliminan.", `${bit.eliminados_peso_invalido} eliminados`],
    ["5", "Largo capturado en mm", "Largo mayor a 300 cm o más de 6 veces la otra dimensión mayor; se divide entre 10.", `${bit.largo_corregido_mm_a_cm} corregidos`],
    ["6", "Viento faltante", "Se imputa con la mediana del viento del mismo clima (no con la mediana global).", `${bit.viento_imputado} imputados`],
    ["7", "Características derivadas", "Dimensiones ordenadas, volumen en litros y one-hot del clima.", "Todas"],
    ["8", "División entrenamiento/prueba", "80/20 estratificada por clase, semilla 42.", `${miles(bit.entrenamiento)} / ${miles(bit.prueba)}`],
    ["9", "Estandarización", "StandardScaler (media 0, desviación 1) sólo para KNN y la red neuronal. Se ajusta únicamente con entrenamiento para no filtrar información del conjunto de prueba.", "7 numéricas"],
  ], { tam: 17, cebra: true }),
  PIE_TABLA("Tabla 7. Bitácora del preprocesamiento inicial (data/procesado/bitacora_preprocesamiento.json)."),
  P("Los árboles de decisión y Random Forest no requieren escalado porque sus divisiones dependen del orden de los valores, no de su magnitud. Por eso se guardan dos versiones del conjunto de datos."),

  H2("1.9 Análisis exploratorio inicial"),
  ...figura("clima_vs_medio.png", 520, 228, "Figura 2. Proporción de cada medio según el clima."),
  P("El clima por sí solo ya separa casos: con tormenta el 100% de los envíos va en carro/camión y con lluvia nunca se usa el dron. Sin embargo, con buen clima ningún medio domina, así que la decisión depende de las demás variables."),
  ...figura("peso_vs_distancia.png", 520, 341, "Figura 3. Peso contra distancia, coloreado por medio (escala logarítmica)."),
  P("La Figura 3 muestra que debajo de 2.5 kg conviven dron y moto, y que entre 2.5 y 20 kg conviven moto y carro. La frontera depende del viento, del volumen y de la distancia al mismo tiempo, por lo que no basta con cortes simples por peso. Ésta es la evidencia de que el problema necesita un modelo que aprenda interacciones."),

  // ===== 2. Paradigma =====
  H1("2. Elección del paradigma de IA"),
  H2("2.1 Comparación de los paradigmas"),
  P("Se evaluó cada paradigma con base en las características de los datos: son tabulares, están etiquetados, mezclan variables numéricas y categóricas, sus fronteras son umbrales (2.5 kg, 30 km/h, 15 km) y el proyecto exige explicar cada decisión."),
  tabla([1700, 2700, 3460, 1500], ["Paradigma", "Idea central", "Ajuste a Empaquétalo", "Decisión"], [
    ["**Simbólico**", "El conocimiento se representa con símbolos y reglas si-entonces. En aprendizaje automático, los árboles de decisión inducen esas reglas a partir de ejemplos.", "Muy alto. Los límites de cada medio son umbrales y un árbol aprende justo eso (“si viento ≥ 30 y peso ≤ 2.5 → moto”). Las reglas aprendidas se pueden leer y sirven para explicar cada decisión.", "**Elegido**"],
    ["Conexionista", "Redes de neuronas artificiales que aprenden pesos con retropropagación.", `Bueno en desempeño (F1 ${f2(R["Red neuronal (MLP)"].f1_prueba)}), pero es una caja negra, requiere escalado y en datos tabulares pequeños no supera a los árboles.`, "Competidor"],
    ["Bioinspirado", "Algoritmos genéticos, enjambres de partículas y colonias de hormigas.", "Bajo. Son métodos de búsqueda y optimización, no clasificadores; sólo servirían para ajustar hiperparámetros, lo cual se resuelve con búsqueda en malla.", "Descartado"],
    ["Probabilístico", "Razonamiento con probabilidades: Naive Bayes y redes bayesianas.", "Medio. Naive Bayes supone independencia entre variables y aquí la esencia es la interacción (lluvia × peso). Sí se aprovecha la idea de entregar una probabilidad como nivel de confianza.", "Descartado como núcleo"],
    ["Computacional", "Resolver por cálculo, búsqueda o similitud (por ejemplo, KNN o lógica difusa).", `Medio-bajo. KNN depende de la escala y tuvo el menor desempeño (F1 ${f2(R["KNN (k=5)"].f1_prueba)}, F1 dron ${f2(R["KNN (k=5)"].f1_dron)}).`, "Competidor"],
  ], { tam: 17 }),
  PIE_TABLA("Tabla 8. Evaluación de los cinco paradigmas para el problema."),

  H2("2.2 Paradigma elegido: enfoque simbólico"),
  P("El proyecto usa el **enfoque simbólico en su variante de aprendizaje automático inductivo**: el conocimiento no lo escribe el programador regla por regla, sino que un árbol de decisión lo induce de los datos en forma de reglas si-entonces. En la clasificación de Domingos (2015), los árboles de decisión pertenecen a la tribu de los “simbolistas”. Se eligió por cuatro razones:"),
  NUM("**Coincide con la forma del problema.** Las restricciones de cada medio son umbrales sobre variables (peso, viento, distancia) y cada nodo de un árbol es justamente un umbral sobre una variable.", "razones"),
  NUM("**Hace posible la explicación.** El objetivo específico 5 exige mostrar el motivo de cada decisión. Las reglas que sigue un árbol son legibles, a diferencia de los pesos de una red neuronal.", "razones"),
  NUM("**Se integra con el filtro de seguridad.** El filtro también es simbólico: son las reglas de la Tabla 6. Modelo y filtro hablan el mismo lenguaje de reglas.", "razones"),
  NUM(`**La evidencia lo respalda.** En la prueba de viabilidad, los modelos basados en árboles obtuvieron el mejor F1-macro (Random Forest ${f2(rf.f1_prueba)}, árbol individual ${f2(R["Árbol de decisión"].f1_prueba)}) sin ningún ajuste.`, "razones"),
  P("**¿En qué se distingue del sistema de reglas simples?** El sistema base son reglas escritas a mano por una persona, con pocos umbrales fijos. El enfoque elegido aprende cientos de reglas y sus umbrales directamente de 4,000 ejemplos, incluidas las interacciones entre viento, volumen, peso y distancia que las reglas simples no contemplan. La diferencia de rendimiento se cuantifica en la sección 4."),
  P("La red neuronal (conexionista) y KNN se conservan como **modelos de comparación**, tal como lo establece el objetivo específico 2."),

  // ===== 3. Modelo =====
  H1("3. Algoritmo y modelo propuesto"),
  H2("3.1 Arquitectura de la solución"),
  tabla([1870, 1870, 1870, 1870, 1880], ["1. Entrada", "2. Preprocesamiento", "3. Random Forest", "4. Filtro de seguridad", "5. Salida"], [
    ["Peso, dimensiones, distancia, clima y viento capturados por el usuario.", "Normalizar clima, ordenar dimensiones, calcular volumen y one-hot.", "Vota entre 3 clases y entrega la probabilidad de cada una.", "Descarta medios que violan la Tabla 6 y toma el siguiente más probable.", "Medio, confianza (0 a 1) y explicación en lenguaje claro."],
  ], { tam: 17 }),
  PIE_TABLA("Tabla 9. Flujo del modelo propuesto."),

  H2("3.2 Random Forest"),
  P("Random Forest (Breiman, 2001) entrena muchos árboles de decisión. Cada árbol ve una muestra distinta de los envíos (bootstrap) y en cada nodo sólo considera un subconjunto aleatorio de variables. Para clasificar un envío, cada árbol vota y la probabilidad de cada medio es la proporción de votos. Esa proporción es el **nivel de confianza** que se mostrará al usuario."),
  P("Cada árbol elige sus cortes minimizando la impureza de Gini, G = 1 − Σ pᵢ², donde pᵢ es la proporción de cada medio en el nodo. Un nodo puro (todos los envíos van en moto) tiene G = 0."),
  P("Se prefirió el ensamble a un solo árbol porque un árbol individual se sobreajusta con facilidad. En la prueba de viabilidad, el árbol individual tuvo 3 puntos menos de F1-macro."),
  tabla([3000, 3600, 2760], ["Hiperparámetro", "Valores a probar (semanas 9 a 11)", "Efecto"], [
    ["n_estimators", "100, 200, 400", "Número de árboles; más árboles dan más estabilidad."],
    ["max_depth", "Sin límite, 10, 20", "Controla el sobreajuste."],
    ["min_samples_leaf", "1, 3, 5", "Hojas más grandes absorben el ruido del etiquetado."],
    ["class_weight", "Ninguno, balanced", "Da más peso a la clase dron, que es minoritaria."],
  ], { cebra: true }),
  PIE_TABLA("Tabla 10. Malla de hiperparámetros (búsqueda en malla con validación cruzada de 5 pliegues y F1-macro)."),

  H2("3.3 Filtro de seguridad y explicación"),
  P("El modelo aprende de ejemplos y siempre puede equivocarse. Por eso, antes de dar la respuesta, el filtro verifica la Tabla 6. Si el medio predicho viola algún límite (por ejemplo, dron con 2.6 kg), se descarta y se toma el siguiente medio más probable. Con esto se garantiza la meta de 0% de clasificaciones inseguras."),
  P("La explicación se arma con dos fuentes: (a) los medios que el filtro descartó y por qué, y (b) las variables que más influyeron en la predicción. Éste es el resultado real del modelo preliminar para un paquete de 1.8 kg, 30 × 20 × 15 cm, a 7.5 km, con lluvia y viento de 25 km/h:"),
  P("**Moto (confianza 0.88).** «La lluvia descarta el dron; el paquete es pequeño y la distancia es media, por lo que la moto es más barata que el carro.»", { align: AlignmentType.CENTER, run: { color: AZUL } }),
  P("Si sólo se cambia el clima a despejado con viento de 10 km/h, el mismo modelo responde **Dron (confianza 0.60)**. Cambiar una sola variable cambia el medio; ésa es la interacción que el modelo aprendió."),

  // ===== 4. Revisión =====
  H1("4. Revisión"),
  P("**Pregunta central: ¿los datos disponibles permiten resolver el problema y el paradigma elegido es adecuado?**", { run: { color: AZUL } }),
  H2("4.1 Prueba de viabilidad"),
  P("Para responder con evidencia y no sólo con argumentos, se entrenaron los cuatro algoritmos con hiperparámetros por defecto (sin ajuste) y se compararon contra el sistema base de reglas simples. Se usó validación cruzada de 5 pliegues sobre entrenamiento y una evaluación final sobre el 20% de prueba (src/prueba_viabilidad.py)."),
  tabla([3060, 2100, 2100, 2100], ["Modelo", "F1-macro (CV 5 pliegues)", "F1-macro (prueba)", "F1 clase dron (prueba)"],
    Object.entries(R).map(([n, v]) => [
      n === "Random Forest" ? `**${n}**` : n,
      v.f1_cv == null ? "—" : `${f2(v.f1_cv)} ± ${f2(v.f1_cv_std)}`,
      n === "Random Forest" ? `**${f2(v.f1_prueba)}**` : f2(v.f1_prueba),
      f2(v.f1_dron),
    ]), { centrar: true, cebra: true }),
  PIE_TABLA("Tabla 11. Resultados preliminares (modelos sin ajustar). Meta del Avance 1: F1-macro ≥ 0.85."),
  ...figura("comparacion_f1.png", 500, 234, "Figura 4. F1-macro en el conjunto de prueba contra la meta de 0.85."),
  ...figura("matrices_confusion.png", 560, 250, "Figura 5. Matrices de confusión del sistema base y de Random Forest en prueba."),

  H2("4.2 Respuesta"),
  P(`**Sí, los datos permiten resolver el problema.** Random Forest alcanza un F1-macro de ${f2(rf.f1_prueba)} sin ningún ajuste, por encima de la meta de 0.85, y supera al sistema de reglas simples por ${mejora} puntos porcentuales (la meta pedía 10). La validación cruzada da ${f2(rf.f1_cv)} ± ${f2(rf.f1_cv_std)}, así que el resultado es estable y no depende de una división afortunada.`),
  P("**Sí, el paradigma es adecuado.** Los dos modelos basados en árboles quedan en los primeros lugares y además son explicables. La matriz de confusión del sistema base muestra su falla típica: manda 147 motos y 186 carros a un medio incorrecto porque ignora el viento, las dimensiones y el costo. Random Forest reduce esos errores a 39 y 15, respectivamente."),
  P("**Hallazgos que se atenderán en las siguientes semanas:**"),
  VIN(`La clase dron es la más difícil (F1 ${f2(rf.f1_dron)}): 25 drones reales se clasificaron como moto y 39 motos como dron. Se probará class_weight = "balanced" y se medirá la precisión de dron por separado, porque asignar dron de más es el error más grave.`),
  VIN("En el conjunto de prueba, ninguna predicción de Random Forest violó un límite físico (0 de 1,022). Aun así, el filtro de seguridad se mantiene para cubrir casos extremos fuera de la distribución, como 2.6 kg o 16 km."),
  VIN("**Limitación:** como el dataset es sintético, el modelo aprende la política del equipo y no el comportamiento real de una paquetería. El ruido del etiquetado impide que el modelo memorice reglas (ningún modelo llega a 1.0) y el sistema base demuestra que las reglas simples no bastan. Esta limitación se reportará en el informe final."),

  // ===== 5. Producto =====
  H1("5. Producto de la revisión"),
  tabla([2900, 4360, 2100], ["Producto solicitado", "Entregado", "Ubicación"], [
    ["Dataset real o muestra significativa", `Dataset sintético de ${miles(total)} envíos limpios (${miles(bit.filas_crudas)} crudos), reproducible con semilla 42.`, "data/"],
    ["Ejemplos de datos", "Seis envíos limpios comentados y errores típicos del archivo crudo.", "Sección 1.5"],
    ["Preprocesamiento inicial", "Nueve pasos con bitácora de registros afectados y conjuntos listos para entrenar.", "Sección 1.8, src/preprocesar.py"],
    ["Paradigma elegido", "Enfoque simbólico (aprendizaje inductivo de árboles de decisión), justificado contra los otros cuatro.", "Sección 2"],
    ["Algoritmo/modelo propuesto", "Random Forest + filtro de seguridad + explicación, con prueba de viabilidad.", "Secciones 3 y 4"],
  ], { cebra: true }),
  PIE_TABLA("Tabla 12. Relación entre lo que pide la revisión y lo entregado."),

  H1("6. Siguientes pasos (semanas 9 a 11)"),
  VIN("**Semana 9:** implementar el sistema base como clase reutilizable y entrenar árbol de decisión y KNN con validación cruzada."),
  VIN("**Semana 10:** entrenar Random Forest y la red neuronal, y hacer la búsqueda de hiperparámetros de la Tabla 10."),
  VIN("**Semana 11:** comparar los modelos, elegir el final e implementar el filtro de seguridad y el módulo de explicación."),

  H1("Referencias"),
  P("Breiman, L. (2001). Random forests. Machine Learning, 45(1), 5–32."),
  P("Domingos, P. (2015). The Master Algorithm: How the Quest for the Ultimate Learning Machine Will Remake Our World. Basic Books."),
  P("Pedregosa, F. et al. (2011). Scikit-learn: Machine learning in Python. Journal of Machine Learning Research, 12, 2825–2830."),
  P("Russell, S. y Norvig, P. (2021). Artificial Intelligence: A Modern Approach (4.ª ed.). Pearson."),

  H1("Anexo: reproducir los resultados"),
  P("Desde la carpeta empaquetalo/ con Python 3 y las bibliotecas numpy, pandas, scikit-learn, matplotlib y joblib:"),
  ...["python src/generar_dataset.py", "python src/preprocesar.py", "python src/prueba_viabilidad.py"].map((c) =>
    new Paragraph({ spacing: { after: 40 }, indent: { left: 720 }, children: [new TextRun({ text: c, font: "Courier New", size: 19 })] })),
];

const doc = new Document({
  features: { updateFields: true },
  creator: "Equipo Empaquétalo",
  title: "Avance 2 – Empaquétalo",
  styles: {
    default: { document: { run: { font: FUENTE, size: 21 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 30, font: "Calibri Light", color: AZUL },
        paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0, keepNext: true } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 23, bold: true, font: FUENTE, color: "000000" },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1, keepNext: true } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 21, bold: true, font: FUENTE },
        paragraph: { spacing: { before: 180, after: 100 }, outlineLevel: 2, keepNext: true } },
    ],
  },
  numbering: {
    config: [
      { reference: "vinetas", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      ...["numeros", "pasos", "razones"].map((reference) => ({ reference, levels: [{ level: 0, format: LevelFormat.DECIMAL,
        text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] })),
    ],
  },
  sections: [
    {
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      children: [...portada, ...indice],
    },
    {
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }, pageNumbers: { start: 1 } } },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT,
        children: [new TextRun({ children: [PageNumber.CURRENT], size: 18, color: "595959" })] })] }) },
      children: cuerpo,
    },
  ],
});

const salida = path.join(RAIZ, "entregables", "Avance2_Empaquetalo.docx");
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(salida, buf); console.log("Escrito", salida); });
