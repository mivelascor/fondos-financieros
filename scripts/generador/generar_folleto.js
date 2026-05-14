const pptxgen = require("pptxgenjs");
const fs = require("fs");

const C = {
  negro:   "000000",
  blanco:  "FFFFFF",
  grisOsc: "333333",
  grisMed: "666666",
  grisClr: "999999",
  grisLin: "CCCCCC",
  grisF:   "F7F7F7",
};

const W = 8.27;
const H = 11.69;
const IZQ_X = 0.25;
const IZQ_W = 2.45;
const DER_X = 2.95;
const DER_W = 5.07;

function ln(slide, y, x=IZQ_X, w=IZQ_W) {
  slide.addShape("line", { x, y, w, h:0, line:{ color:C.grisLin, width:0.4 } });
}

function sec(slide, txt, y, x=IZQ_X, w=IZQ_W) {
  slide.addText(txt, { x, y, w, h:0.19, fontSize:8.5, bold:true, color:C.negro, fontFace:"Calibri", margin:0 });
}

// ─── SLIDE 1 ─────────────────────────────────────────────────────────────────
function slide1(pres, d) {
  const slide = pres.addSlide();
  slide.background = { color: C.blanco };

  // Franja negra
  slide.addShape("rect", { x:0, y:0, w:2.78, h:2.0, fill:{ color:C.negro } });

  // Nombre fondo
  const words = d.nombre_corto.split(" ");
  const l1 = words.length > 2 ? words.slice(0,-1).join(" ") : words[0];
  const l2 = words.length > 2 ? words[words.length-1] : words.slice(1).join(" ");
  slide.addText([
    {text: l1, options:{breakLine:true}},
    {text: l2}
  ], {
    x:0.2, y:0.18, w:2.4, h:1.45,
    fontSize:24, bold:true, color:C.blanco, fontFace:"Calibri", valign:"top", margin:0
  });

  slide.addText("F O N D O", {
    x:0.2, y:1.68, w:2.4, h:0.2,
    fontSize:7, color:"AAAAAA", fontFace:"Calibri", charSpacing:5, margin:0
  });

  // Separador vertical
  slide.addShape("line", { x:2.8, y:0, w:0, h:H, line:{ color:C.grisLin, width:0.5 } });

  // Logo derecha
  slide.addText("VanTrust Capital", {
    x:DER_X, y:0.15, w:DER_W, h:0.2,
    fontSize:7.5, color:C.grisMed, fontFace:"Calibri", align:"right", margin:0
  });

  // ─── COL IZQUIERDA ────────────────────────────────────────────────────────
  let y = 2.1;

  sec(slide, "Información General", y); y += 0.22;
  ln(slide, y); y += 0.08;

  const info = [
    ["Administradora",   d.administradora],
    ["RUT Fondo",        d.rut],
    ["Moneda",           d.moneda],
    ["Tipo de Fondo",    d.tipo],
    ["Fecha Inicio",     d.fecha_inicio],
    ["Benchmark",        d.benchmark],
    ["Fondo Rescatable", "Sí"],
    ["Plazo Rescate",    d.plazo_rescate],
    ["Riesgos",          "Mercado- Crédito - Liquidez - Tasa de interés - Derivados"],
    ["Remuneración",     d.remuneracion || "0,295% IVA Incluido"],
    ["Custodio",         "Vantrust Capital C. de Bolsa"],
  ];

  for (const [k,v] of info) {
    const alto = (k==="Riesgos"||k==="Remuneración") ? 0.28 : 0.18;
    slide.addText(k, { x:IZQ_X, y, w:1.0, h:alto, fontSize:6.8, bold:true, color:C.negro, fontFace:"Calibri", valign:"top", margin:0, wrap:true });
    slide.addText(v, { x:IZQ_X+1.02, y, w:1.42, h:alto, fontSize:6.8, color:C.grisOsc, fontFace:"Calibri", valign:"top", margin:0, wrap:true });
    y += alto + 0.02;
  }

  y += 0.06; ln(slide, y); y += 0.1;

  sec(slide, "Objetivo", y); y += 0.2;
  slide.addText(d.objetivo, { x:IZQ_X, y, w:IZQ_W, h:0.72, fontSize:7, color:C.grisOsc, fontFace:"Calibri", valign:"top", margin:0, wrap:true });
  y += 0.82; ln(slide, y); y += 0.1;

  sec(slide, "Rentabilidad", y); y += 0.2;
  slide.addText(d.rentabilidad_texto || `La rentabilidad esperada del ${d.nombre_fondo}, es la tasa de política monetaria promedio del Banco Central de Chile.`, {
    x:IZQ_X, y, w:IZQ_W, h:0.62, fontSize:7, color:C.grisOsc, fontFace:"Calibri", valign:"top", margin:0, wrap:true
  });
  y += 0.72; ln(slide, y); y += 0.1;

  sec(slide, "Inversionistas", y); y += 0.2;
  slide.addText(d.inversionistas, { x:IZQ_X, y, w:IZQ_W, h:0.55, fontSize:7, color:C.grisOsc, fontFace:"Calibri", valign:"top", margin:0, wrap:true });

  // ─── COL DERECHA ──────────────────────────────────────────────────────────
  let yr = 0.15;

  slide.addText("Comentario Portafolio Manager", {
    x:DER_X, y:yr, w:DER_W, h:0.27,
    fontSize:13, bold:true, color:C.negro, fontFace:"Calibri", margin:0
  });
  yr += 0.32;

  // Línea bajo título comentario
  slide.addShape("line", { x:DER_X, y:yr, w:DER_W, h:0, line:{ color:C.grisLin, width:0.4 } });
  yr += 0.1;

  slide.addText(d.comentario, {
    x:DER_X, y:yr, w:DER_W, h:1.75,
    fontSize:8, color:C.grisOsc, fontFace:"Calibri", valign:"top", margin:0, wrap:true
  });
  yr += 1.85;

  // Gráfico
  if (d.grafico_labels && d.grafico_labels.length > 0) {
    slide.addText(d.nombre_corto, {
      x:DER_X, y:yr, w:DER_W, h:0.18,
      fontSize:8, bold:true, color:C.negro, fontFace:"Calibri", align:"center", margin:0
    });
    yr += 0.2;

    const series = [];
    if (d.grafico_fondo?.length) series.push({ name:`FIP ${d.nombre_corto}`, labels:d.grafico_labels, values:d.grafico_fondo });
    if (d.grafico_icp?.length)   series.push({ name:"ICP Nom",              labels:d.grafico_labels, values:d.grafico_icp });
    if (d.grafico_comp?.length)  series.push({ name:"Competencia Relevante",labels:d.grafico_labels, values:d.grafico_comp });

    if (series.length) {
      slide.addChart("line", series, {
        x:DER_X, y:yr, w:DER_W, h:2.0,
        chartColors:["000000","999999","BBBBBB"],
        lineSize:1.5, lineSmooth:false,
        showLegend:true, legendPos:"b", legendFontSize:7,
        catAxisLabelColor:C.grisMed, valAxisLabelColor:C.grisMed,
        catAxisLabelFontSize:6, valAxisLabelFontSize:7,
        catAxisLabelFrequency: Math.max(1, Math.floor(d.grafico_labels.length/8)),
        valGridLine:{ color:"EEEEEE", size:0.5 },
        catGridLine:{ style:"none" },
        chartArea:{ fill:{ color:C.blanco } },
        plotArea:{ fill:{ color:C.blanco } },
        showTitle:false,
        valAxisMinVal: Math.floor(Math.min(...series.flatMap(s=>s.values)) * 0.95),
      });
      yr += 2.18;
    }
  }

  // Evolución Rentabilidad
  slide.addText("Evolución Rentabilidad", {
    x:DER_X, y:yr, w:DER_W, h:0.25,
    fontSize:13, bold:true, color:C.negro, fontFace:"Calibri", margin:0
  });
  yr += 0.3;

  // Tabla resumen
  if (d.tabla_rentab?.length) {
    const colW = [1.45, 0.65, 0.72, 0.72, 0.65, 0.68];
    const hdrs = ["Rentabilidad","Mensual","Trimestral","Semestral","Anual",`Acum\n${d.anio_acum||"2026"} (*)`];

    const rows = [
      hdrs.map(h => ({ text:h, options:{ bold:true, fontSize:7.5, fontFace:"Calibri", align:"center", color:C.negro, fill:{ color:C.blanco }, border:[{pt:0.5,color:C.grisLin},{pt:0.5,color:C.grisLin},{pt:1,color:C.negro},{pt:0.5,color:C.grisLin}] } })),
      ...d.tabla_rentab.map((r,i) => [
        { text:r.nombre,     options:{ bold:r.es_fondo||i===2, fontSize:7.5, fontFace:"Calibri", align:"left",   color:C.negro, fill:{ color: i%2===0?"F9F9F9":C.blanco } } },
        { text:r.mensual,    options:{ fontSize:7.5, fontFace:"Calibri", align:"center", color:C.grisOsc, fill:{ color: i%2===0?"F9F9F9":C.blanco } } },
        { text:r.trimestral, options:{ fontSize:7.5, fontFace:"Calibri", align:"center", color:C.grisOsc, fill:{ color: i%2===0?"F9F9F9":C.blanco } } },
        { text:r.semestral,  options:{ fontSize:7.5, fontFace:"Calibri", align:"center", color:C.grisOsc, fill:{ color: i%2===0?"F9F9F9":C.blanco } } },
        { text:r.anual,      options:{ fontSize:7.5, fontFace:"Calibri", align:"center", color:C.grisOsc, fill:{ color: i%2===0?"F9F9F9":C.blanco } } },
        { text:r.ytd,        options:{ fontSize:7.5, fontFace:"Calibri", align:"center", color:C.grisOsc, fill:{ color: i%2===0?"F9F9F9":C.blanco } } },
      ])
    ];

    slide.addTable(rows, { x:DER_X, y:yr, w:DER_W, colW, rowH:0.21, border:{ pt:0.3, color:C.grisLin } });
    yr += rows.length * 0.21 + 0.05;
  }

  slide.addText("*Valores correspondientes a la rentabilidad anualizada, no acumulada", {
    x:DER_X, y:yr, w:DER_W, h:0.15,
    fontSize:6, color:C.grisClr, fontFace:"Calibri", italic:true, margin:0
  });
  yr += 0.2;

  // Tabla histórica
  if (d.tabla_historica?.length) {
    ln(slide, yr, DER_X, DER_W); yr += 0.07;

    const meses = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic","Total\nAño"];
    const cw = [0.22, 0.82, 0.31,0.31,0.31,0.31,0.31,0.31,0.31,0.31,0.31,0.31,0.31,0.31, 0.33];

    const mkOpt = (txt, opts={}) => ({ text:String(txt||""), options:{ fontSize:5.8, fontFace:"Calibri", color:C.grisOsc, ...opts } });

    const hRow = [
      mkOpt("Año", {bold:true, align:"left", color:C.negro}),
      mkOpt("Fondo", {bold:true, align:"left", color:C.negro}),
      ...meses.map(m => mkOpt(m, {bold:true, align:"center", color:C.negro}))
    ];

    const tableRows = [hRow];
    for (const anioData of d.tabla_historica) {
      for (let si=0; si<anioData.series.length; si++) {
        const serie = anioData.series[si];
        const esF = serie.es_fondo;
        const fila = [
          mkOpt(si===0 ? anioData.anio : "", {bold:true, color:C.negro, align:"left"}),
          mkOpt(serie.nombre, {bold:esF, color: esF ? C.negro : C.grisMed, align:"left"}),
          ...Array(13).fill(0).map((_,i) => mkOpt(serie.valores[i]||"", {align:"center", color: esF ? C.negro : C.grisMed}))
        ];
        tableRows.push(fila);
      }
    }

    slide.addTable(tableRows, {
      x:DER_X, y:yr, w:DER_W,
      colW:cw, rowH:0.135,
      border:{ pt:0.2, color:"EEEEEE" }
    });
  }
}

// ─── SLIDE 2 ─────────────────────────────────────────────────────────────────
function slide2(pres, d) {
  const slide = pres.addSlide();
  slide.background = { color: C.blanco };

  slide.addShape("rect", { x:0, y:0, w:2.78, h:1.45, fill:{ color:C.negro } });
  slide.addText("F O N D O", { x:0.2, y:1.12, w:2.4, h:0.2, fontSize:7, color:"AAAAAA", fontFace:"Calibri", charSpacing:5, margin:0 });
  slide.addShape("line", { x:2.8, y:0, w:0, h:H, line:{ color:C.grisLin, width:0.5 } });
  slide.addText("VanTrust Capital", { x:DER_X, y:0.15, w:DER_W, h:0.2, fontSize:7.5, color:C.grisMed, fontFace:"Calibri", align:"right", margin:0 });

  // ─── COL IZQ ─────────────────────────────────────────────────────────────
  let y = 1.6;

  sec(slide, "Composición por Moneda", y); y += 0.22;
  ln(slide, y); y += 0.1;

  if (d.comp_moneda?.length) {
    for (const [mon, pct] of d.comp_moneda) {
      slide.addText(mon, { x:IZQ_X, y, w:1.4, h:0.2, fontSize:8, fontFace:"Calibri", margin:0 });
      slide.addText(pct, { x:IZQ_X+1.4, y, w:1.05, h:0.2, fontSize:8, fontFace:"Calibri", align:"right", margin:0, bold:true });
      y += 0.22;
    }
  }
  y += 0.12; ln(slide, y); y += 0.1;

  sec(slide, "Composición General por Instrumentos Financieros", y); y += 0.22;

  // Composición por Instrumento (si existe)
  if (d.comp_instrumento && d.comp_instrumento.length > 0) {
    for (const [instr, pct] of d.comp_instrumento) {
      slide.addText(instr, { x:IZQ_X, y, w:1.55, h:0.18, fontSize:7.5, fontFace:"Calibri", margin:0 });
      slide.addText(pct,   { x:IZQ_X+1.55, y, w:0.9,  h:0.18, fontSize:7.5, fontFace:"Calibri", align:"right", margin:0 });
      y += 0.2;
    }
    y += 0.05;
  }

  // Tabla límites
  const limites = d.limites || [
    ["Bonos","0%","25%"],["Depósitos a Plazo","5%","100%"],["Facturas","0%","40%"],
    ["Financiamientos\n(Inmobiliario, Mezzanine, etc.)","0%","100%"],
    ["Efectos de Comercio","0%","40%"],["Pagaré Banco Central","0%","100%"],["Caja","5%","25%"]
  ];

  const limRows = [
    [ {text:"", options:{fontSize:7}},
      {text:"Límites de Inversión", options:{bold:true, fontSize:7, align:"center", colspan:2, fontFace:"Calibri"}} ],
    [ {text:"", options:{fontSize:7}},
      {text:"Mínimo", options:{bold:true, fontSize:7, align:"center", fontFace:"Calibri"}},
      {text:"Máximo", options:{bold:true, fontSize:7, align:"center", fontFace:"Calibri"}} ],
    ...limites.map(([inst,min,max]) => [
      {text:inst, options:{fontSize:7, fontFace:"Calibri", wrap:true}},
      {text:min,  options:{fontSize:7, fontFace:"Calibri", align:"center"}},
      {text:max,  options:{fontSize:7, fontFace:"Calibri", align:"center"}}
    ])
  ];

  const limH = limRows.length * 0.19;
  slide.addTable(limRows, {
    x:IZQ_X, y, w:IZQ_W,
    colW:[1.32, 0.55, 0.58], rowH:0.19,
    border:{ pt:0.3, color:C.grisLin }
  });
  y += limH + 0.22;
  ln(slide, y); y += 0.1;

  sec(slide, "Composición por Duración", y); y += 0.25;

  if (d.comp_duracion?.length) {
    for (const [tramo, pct] of d.comp_duracion) {
      slide.addText(tramo, { x:IZQ_X, y, w:1.55, h:0.22, fontSize:8, fontFace:"Calibri", margin:0 });
      slide.addText(pct,   { x:IZQ_X+1.55, y, w:0.9,  h:0.22, fontSize:8, fontFace:"Calibri", align:"right", margin:0, bold:true });
      y += 0.24;
    }
  }

  // ─── COL DER ─────────────────────────────────────────────────────────────
  let yr = 0.15;

  slide.addText("Glosario", {
    x:DER_X, y:yr, w:DER_W, h:0.27,
    fontSize:13, bold:true, color:C.negro, fontFace:"Calibri", margin:0
  });
  yr += 0.35;

  const glosario = [
    ["Riesgo de Mercado","Este es el riesgo de una variación adversa en el precio o tasa de mercado en los instrumentos en que invierte el Fondo"],
    ["Riesgo de Crédito","Es la posible pérdida que se asume como consecuencia del incumplimiento de las obligaciones directas, indirectas o de derivados que conllevan al no pago parcial, total o falta de oportunidad del pago de los emisores de los instrumentos en que invierte el Fondo y que pudieran ocasionar una pérdida financiera."],
    ["Riesgo de Liquidez","Riesgo asociado a la capacidad de generación de recursos del Fondo para cumplir con sus obligaciones de rescate o vencimiento del mismo."],
    ["Riesgo Profundidad de Mercado","Corresponde a la posibilidad de comprar o vender un activo financiero al valor de mercado en un período de tiempo acorde a las características del instrumento. Períodos de alta volatilidad de mercado afectan negativamente estos tiempos."],
    ["Riesgo de Tasa de interés","Exposición a pérdidas ocasionadas por cambios adversos en las tasas de interés de mercado y que afecten el valor de los instrumentos, contratos y demás operaciones registradas en el balance."],
    ["Riesgo de Moneda","Es la exposición a pérdidas ocasionadas por cambios adversos en el valor en moneda nacional de las monedas extranjeras que están expresados a los instrumentos, contratos y demás operaciones del balance del Fondo."],
    ["Riesgo Sectorial","Este riesgo está asociado a condiciones de mercado adversas que pueden afectar a un sector industrial en particular y por ende la rentabilidad del Fondo."],
    ["Gastos del fondo","Corresponden a los gastos directos e indirectos necesarios para el correcto funcionamiento del fondo los que están detallados en el Reglamento Interno"],
    ["Forma de Ingreso y Pago del Fondo","La moneda en que el inversionista entra al fondo es aportando pesos chilenos, y al rescate de las cuotas, el fondo le entrega pesos chilenos."],
  ];

  for (const [tit, desc] of glosario) {
    slide.addText([
      { text:tit,  options:{ bold:true, breakLine:true } },
      { text:desc }
    ], {
      x:DER_X, y:yr, w:DER_W, h:0.44,
      fontSize:7, color:C.grisOsc, fontFace:"Calibri", valign:"top", margin:0, wrap:true
    });
    yr += 0.47;
  }

  yr += 0.08;
  ln(slide, yr, DER_X, DER_W); yr += 0.14;

  slide.addText("Disclaimer", {
    x:DER_X, y:yr, w:DER_W, h:0.25,
    fontSize:13, bold:true, color:C.negro, fontFace:"Calibri", margin:0
  });
  yr += 0.3;

  slide.addText("Conforme a la Ley Única de Fondos, las administradoras de fondos de inversión privados están sujetas a las obligaciones de información establecidas por la Comisión para el Mercado Financiero. Tales fondos no están sometidos a fiscalización de la Comisión y no hacemos oferta pública de sus cuotas.", {
    x:DER_X, y:yr, w:DER_W, h:0.7,
    fontSize:7.5, color:C.grisOsc, fontFace:"Calibri", valign:"top", margin:0, wrap:true
  });
}

// ─── Main ────────────────────────────────────────────────────────────────────
async function main() {
  const ai = process.argv.indexOf("--data");
  const oi = process.argv.indexOf("--out");
  if (ai===-1||oi===-1) { console.error("Uso: node generar_folleto.js --data datos.json --out output.pptx"); process.exit(1); }

  const d = JSON.parse(fs.readFileSync(process.argv[ai+1], "utf8"));
  const pres = new pptxgen();
  pres.defineLayout({ name:"A4P", width:8.27, height:11.69 });
  pres.layout = "A4P";

  slide1(pres, d);
  slide2(pres, d);

  await pres.writeFile({ fileName: process.argv[oi+1] });
  console.log("OK:" + process.argv[oi+1]);
}

main().catch(e => { console.error(e); process.exit(1); });
