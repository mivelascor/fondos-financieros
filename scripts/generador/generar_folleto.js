/**
 * generar_folleto.js — Genera folleto comercial fiel al diseño HTML de referencia.
 * Uso: node generar_folleto.js --data datos.json --out output.pptx
 */
const pptxgen = require("pptxgenjs");
const fs      = require("fs");

// ── Colores ──────────────────────────────────────────────────────────────────
const C = {
  negro:   "000000", blanco:  "FFFFFF",
  grisOsc: "333333", grisMed: "666666",
  grisClr: "999999", grisLin: "DDDDDD",
  grisF:   "F7F7F7",
};

// ── Layout A4 portrait ────────────────────────────────────────────────────────
const W  = 8.27;   // ancho total inches
const H  = 11.69;  // alto total inches

// Columnas (iguales al HTML: col-izq 68mm, col-der flex)
// 68mm / 25.4 * 1 inch ≈ 2.68"
const IZQ_W  = 2.68;
const IZQ_X  = 0.0;
const DER_X  = 2.85;
const DER_W  = W - DER_X - 0.15;
const PAD    = 0.13;  // padding interno

// ── Helpers ───────────────────────────────────────────────────────────────────
function pct(v, d=2) {
  if (v === null || v === undefined || isNaN(v)) return "—";
  return (v * 100).toFixed(d) + "%";
}

function lin(slide, y, x=DER_X, w=DER_W, color=C.grisLin, pt=0.5) {
  slide.addShape("line", { x, y, w, h:0, line:{ color, width:pt } });
}

function sec(slide, txt, y, x=DER_X) {
  slide.addText(txt, {
    x, y, w: DER_W, h:0.26,
    fontSize:13, bold:true, color:C.negro, fontFace:"Calibri",
    margin:0,
  });
}

// ── SLIDE 1 ───────────────────────────────────────────────────────────────────
function slide1(pres, d) {
  const slide = pres.addSlide();
  slide.background = { color: C.blanco };

  // ── COLUMNA IZQUIERDA ────────────────────────────────────────────────────
  // Franja negra con nombre del fondo
  slide.addShape("rect", { x:0, y:0, w:IZQ_W, h:H, fill:{ color: C.blanco } });
  slide.addShape("rect", { x:0, y:0, w:IZQ_W, h:1.9,  fill:{ color: C.negro } });

  // Nombre del fondo
  const words = d.nombre_corto.split(" ");
  const l1 = words.length > 1 ? words.slice(0,-1).join(" ") : words[0];
  const l2 = words.length > 1 ? words[words.length-1] : "";
  const titulo = l2 ? `${l1}\n${l2}` : l1;
  slide.addText(titulo, {
    x:PAD, y:0.18, w:IZQ_W-PAD*2, h:1.3,
    fontSize:22, bold:true, color:C.blanco, fontFace:"Calibri",
    valign:"top", margin:0, wrap:true,
  });
  slide.addText("F O N D O", {
    x:PAD, y:1.55, w:IZQ_W-PAD*2, h:0.2,
    fontSize:7, color:"AAAAAA", fontFace:"Calibri",
    charSpacing:4, margin:0,
  });

  // Separador derecha de col izq
  slide.addShape("line", {
    x:IZQ_W, y:0, w:0, h:H,
    line:{ color:"EEEEEE", width:0.5 },
  });

  // Contenido col izq
  let y = 2.0;
  const LP = PAD;
  const LW = IZQ_W - PAD*2;

  // Información General
  slide.addText("Información General", {
    x:LP, y, w:LW, h:0.18,
    fontSize:8, bold:true, color:C.negro, fontFace:"Calibri", margin:0,
  });
  y += 0.2;
  slide.addShape("line", { x:LP, y, w:LW, h:0, line:{ color:C.negro, width:1 } });
  y += 0.07;

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

  for (const [k, v] of info) {
    const alto = (k==="Riesgos"||k==="Remuneración") ? 0.27 : 0.17;
    slide.addText(k, {
      x:LP, y, w:1.0, h:alto,
      fontSize:6.5, bold:true, color:C.negro, fontFace:"Calibri",
      valign:"top", margin:0, wrap:true,
    });
    slide.addText(v, {
      x:LP+1.02, y, w:LW-1.02, h:alto,
      fontSize:6.5, color:"555555", fontFace:"Calibri",
      valign:"top", margin:0, wrap:true,
    });
    y += alto + 0.02;
  }

  y += 0.06;
  slide.addShape("line", { x:LP, y, w:LW, h:0, line:{ color:C.grisLin, width:0.5 } });
  y += 0.1;

  // Objetivo
  slide.addText("Objetivo", { x:LP, y, w:LW, h:0.18, fontSize:8, bold:true, color:C.negro, fontFace:"Calibri", margin:0 });
  y += 0.2;
  slide.addShape("line", { x:LP, y, w:LW, h:0, line:{ color:C.negro, width:1 } });
  y += 0.07;
  slide.addText(d.objetivo, {
    x:LP, y, w:LW, h:0.75,
    fontSize:7, color:"555555", fontFace:"Calibri", valign:"top", margin:0, wrap:true,
  });
  y += 0.82;

  slide.addShape("line", { x:LP, y, w:LW, h:0, line:{ color:C.grisLin, width:0.5 } });
  y += 0.1;

  // Rentabilidad
  slide.addText("Rentabilidad", { x:LP, y, w:LW, h:0.18, fontSize:8, bold:true, color:C.negro, fontFace:"Calibri", margin:0 });
  y += 0.2;
  slide.addShape("line", { x:LP, y, w:LW, h:0, line:{ color:C.negro, width:1 } });
  y += 0.07;
  slide.addText(d.rentabilidad_texto, {
    x:LP, y, w:LW, h:0.65,
    fontSize:7, color:"555555", fontFace:"Calibri", valign:"top", margin:0, wrap:true,
  });
  y += 0.72;

  slide.addShape("line", { x:LP, y, w:LW, h:0, line:{ color:C.grisLin, width:0.5 } });
  y += 0.1;

  // Inversionistas
  slide.addText("Inversionistas", { x:LP, y, w:LW, h:0.18, fontSize:8, bold:true, color:C.negro, fontFace:"Calibri", margin:0 });
  y += 0.2;
  slide.addShape("line", { x:LP, y, w:LW, h:0, line:{ color:C.negro, width:1 } });
  y += 0.07;
  slide.addText(d.inversionistas, {
    x:LP, y, w:LW, h:0.55,
    fontSize:7, color:"555555", fontFace:"Calibri", valign:"top", margin:0, wrap:true,
  });

  // ── COLUMNA DERECHA ──────────────────────────────────────────────────────
  let yr = 0.15;

  // Comentario PM
  slide.addText("Comentario Portafolio Manager", {
    x:DER_X, y:yr, w:DER_W, h:0.27,
    fontSize:13, bold:true, color:C.negro, fontFace:"Calibri", margin:0,
  });
  yr += 0.3;
  lin(slide, yr, DER_X, DER_W, C.negro, 1.5);
  yr += 0.1;

  slide.addText(d.comentario, {
    x:DER_X, y:yr, w:DER_W, h:1.8,
    fontSize:8, color:C.grisOsc, fontFace:"Calibri",
    valign:"top", margin:0, wrap:true, align:"justify",
  });
  yr += 1.9;

  // Gráfico
  if (d.grafico_labels && d.grafico_labels.length > 0) {
    slide.addText(d.nombre_corto, {
      x:DER_X, y:yr, w:DER_W, h:0.17,
      fontSize:7.5, bold:true, color:C.negro, fontFace:"Calibri",
      align:"center", margin:0,
    });
    yr += 0.19;

    const series = [];
    if (d.grafico_fondo?.length) series.push({ name:`FIP ${d.nombre_corto}`,       labels:d.grafico_labels, values:d.grafico_fondo });
    if (d.grafico_icp?.length)   series.push({ name:"ICP Nom.",                    labels:d.grafico_labels, values:d.grafico_icp   });
    if (d.grafico_comp?.length)  series.push({ name:"Competencia Relevante (*)",   labels:d.grafico_labels, values:d.grafico_comp  });

    if (series.length) {
      slide.addChart("line", series, {
        x:DER_X, y:yr, w:DER_W, h:1.9,
        chartColors:["000000","999999","555555"],
        lineSize:1.5, lineSmooth:false,
        showLegend:true, legendPos:"r", legendFontSize:6,
        catAxisLabelFontSize:6, valAxisLabelFontSize:6,
        catAxisLabelColor:C.grisMed, valAxisLabelColor:C.grisMed,
        catAxisLabelFrequency: Math.max(1, Math.floor(d.grafico_labels.length/10)),
        valGridLine:{ color:"EEEEEE", size:0.3 },
        catGridLine:{ style:"none" },
        chartArea:{ fill:{ color:C.blanco } },
        plotArea:{ fill:{ color:C.blanco } },
        showTitle:false,
        valAxisMinVal: Math.floor(Math.min(...series.flatMap(s=>s.values).filter(v=>v!==null)) * 0.95),
      });
      yr += 2.08;
    }
  }

  // Evolución Rentabilidad
  slide.addText("Evolución Rentabilidad", {
    x:DER_X, y:yr, w:DER_W, h:0.25,
    fontSize:13, bold:true, color:C.negro, fontFace:"Calibri", margin:0,
  });
  yr += 0.28;
  lin(slide, yr, DER_X, DER_W, C.negro, 1.5);
  yr += 0.08;

  // Tabla resumen
  if (d.tabla_rentab?.length) {
    const anioLabel = `Acum\n${d.anio_acum||"2026"} (*)`;
    const colW = [1.42, 0.62, 0.72, 0.72, 0.62, 0.65];
    const hdrs = ["Rentabilidad","Mensual","Trimestral","Semestral","Anual", anioLabel];

    const mkH = (txt) => ({
      text:txt,
      options:{ bold:true, fontSize:7.5, fontFace:"Calibri", align:"center",
                color:C.negro, fill:{ color:C.blanco },
                border:[{pt:1.5,color:C.negro},{pt:1.5,color:C.negro},{pt:1.5,color:C.negro},{pt:1.5,color:C.negro}] }
    });
    const mkH1 = (txt) => ({ ...mkH(txt), options:{ ...mkH(txt).options, align:"left" } });

    const rows = [
      [mkH1(hdrs[0]), ...hdrs.slice(1).map(mkH)],
      ...d.tabla_rentab.map((r, i) => {
        const bg = i%2===0 ? "F9F9F9" : C.blanco;
        const mkD = (txt, left=false) => ({
          text: String(txt||"—"),
          options:{ fontSize:7.5, fontFace:"Calibri",
                    align: left?"left":"center",
                    color: left ? C.negro : C.grisOsc,
                    bold: left,
                    fill:{ color: bg },
                    border:[{pt:0.3,color:C.grisLin},{pt:0.3,color:C.grisLin},{pt:0.3,color:C.grisLin},{pt:0.3,color:C.grisLin}] }
        });
        return [mkD(r.nombre,true), mkD(r.mensual), mkD(r.trimestral), mkD(r.semestral), mkD(r.anual), mkD(r.ytd)];
      })
    ];

    slide.addTable(rows, {
      x:DER_X, y:yr, w:DER_W, colW, rowH:0.2,
    });
    yr += rows.length * 0.2 + 0.05;
  }

  // Nota
  slide.addText("*Valores correspondientes a la rentabilidad anualizada, no acumulada", {
    x:DER_X, y:yr, w:DER_W, h:0.14,
    fontSize:6, color:C.grisClr, fontFace:"Calibri", italic:true, margin:0,
  });
  yr += 0.18;

  // Tabla histórica
  if (d.tabla_historica?.length) {
    lin(slide, yr, DER_X, DER_W, C.grisLin, 0.5); yr += 0.07;

    const meses = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic","Total\nAño"];
    const cw    = [0.24, 0.94, 0.29,0.29,0.29,0.29,0.29,0.29,0.28,0.29,0.28,0.29,0.28,0.28, 0.30];

    const border = [{pt:0.2,color:"EEEEEE"},{pt:0.2,color:"EEEEEE"},{pt:0.2,color:"EEEEEE"},{pt:0.2,color:"EEEEEE"}];
    const borderTop = [{pt:1.5,color:C.negro},{pt:0.2,color:"EEEEEE"},{pt:1.5,color:C.negro},{pt:0.2,color:"EEEEEE"}];
    const borderBot = [{pt:0.2,color:"EEEEEE"},{pt:0.2,color:"EEEEEE"},{pt:1.5,color:C.negro},{pt:0.2,color:"EEEEEE"}];

    const mkTH = (txt) => ({
      text:txt,
      options:{ fontSize:6, bold:true, fontFace:"Calibri", color:C.negro, align:"center",
                border:borderTop }
    });
    const mkTH1 = (txt) => ({ ...mkTH(txt), options:{ ...mkTH(txt).options, align:"left" } });

    const tableRows = [[mkTH1("Año"), mkTH1("Fondo"), ...meses.map(mkTH)]];

    for (const anioData of d.tabla_historica) {
      let isFirst = true;
      for (const serie of anioData.series) {
        const es_fondo = serie.es_fondo;
        const color    = es_fondo ? C.negro : (serie.nombre==="ICP"?"666666":"999999");

        const mkD = (txt, left=false) => ({
          text: String(txt||""),
          options:{ fontSize:5.8, fontFace:"Calibri", color, align:left?"left":"center", border }
        });

        const vals = serie.valores || [];  // floats
        const total = serie.total;

        const fila = [
          { text: isFirst ? String(anioData.anio) : "",
            options:{ fontSize:5.8, bold:true, fontFace:"Calibri", color:C.negro, align:"left", border } },
          mkD(serie.nombre, true),
          ...vals.map(v => ({
            text: (v!==null&&v!==undefined) ? pct(v) : "",
            options:{ fontSize:5.8, fontFace:"Calibri", color, align:"center", border }
          })),
          { text: total!==null&&total!==undefined ? pct(total) : "",
            options:{ fontSize:5.8, bold:true, fontFace:"Calibri", color, align:"center", border } }
        ];
        tableRows.push(fila);
        isFirst = false;
      }
    }

    slide.addTable(tableRows, {
      x:DER_X, y:yr, w:DER_W, colW:cw, rowH:0.14,
    });
  }
}

// ── SLIDE 2 ───────────────────────────────────────────────────────────────────
function slide2(pres, d) {
  const slide = pres.addSlide();
  slide.background = { color: C.blanco };

  // Col izq
  slide.addShape("rect", { x:0, y:0, w:IZQ_W, h:H, fill:{ color:C.blanco } });
  slide.addShape("rect", { x:0, y:0, w:IZQ_W, h:1.45, fill:{ color:C.negro } });
  slide.addText("F O N D O", {
    x:PAD, y:1.12, w:IZQ_W-PAD*2, h:0.2,
    fontSize:7, color:"AAAAAA", fontFace:"Calibri", charSpacing:4, margin:0,
  });
  slide.addShape("line", { x:IZQ_W, y:0, w:0, h:H, line:{ color:"EEEEEE", width:0.5 } });

  const LP = PAD;
  const LW = IZQ_W - PAD*2;
  let y = 1.6;

  const secL = (txt, yy) => {
    slide.addText(txt, { x:LP, y:yy, w:LW, h:0.18, fontSize:8, bold:true, color:C.negro, fontFace:"Calibri", margin:0 });
  };

  // Composición por Moneda
  secL("Composición por Moneda", y); y += 0.2;
  slide.addShape("line", { x:LP, y, w:LW, h:0, line:{ color:C.negro, width:1 } });
  y += 0.08;

  if (d.comp_moneda?.length) {
    for (const [mon, pcstr] of d.comp_moneda) {
      slide.addText(mon, { x:LP, y, w:1.3, h:0.19, fontSize:7.5, fontFace:"Calibri", margin:0 });
      slide.addText(pcstr, { x:LP+1.3, y, w:LW-1.3, h:0.19, fontSize:7.5, fontFace:"Calibri", align:"right", bold:true, margin:0 });
      y += 0.21;
    }
  }
  y += 0.1;
  slide.addShape("line", { x:LP, y, w:LW, h:0, line:{ color:C.grisLin, width:0.5 } });
  y += 0.1;

  // Composición por Instrumento
  secL("Composición por Instrumento", y); y += 0.2;
  slide.addShape("line", { x:LP, y, w:LW, h:0, line:{ color:C.negro, width:1 } });
  y += 0.08;

  if (d.comp_instrumento?.length) {
    for (const [inst, pcstr] of d.comp_instrumento) {
      slide.addText(inst, { x:LP, y, w:1.5, h:0.18, fontSize:7, fontFace:"Calibri", margin:0, wrap:true });
      slide.addText(pcstr, { x:LP+1.5, y, w:LW-1.5, h:0.18, fontSize:7, fontFace:"Calibri", align:"right", bold:true, margin:0 });
      y += 0.19;
    }
  }
  y += 0.1;
  slide.addShape("line", { x:LP, y, w:LW, h:0, line:{ color:C.grisLin, width:0.5 } });
  y += 0.1;

  // Composición por Duración
  secL("Composición por Duración", y); y += 0.2;
  slide.addShape("line", { x:LP, y, w:LW, h:0, line:{ color:C.negro, width:1 } });
  y += 0.08;

  if (d.comp_duracion?.length) {
    for (const [tramo, pcstr] of d.comp_duracion) {
      slide.addText(tramo, { x:LP, y, w:1.5, h:0.2, fontSize:7.5, fontFace:"Calibri", margin:0 });
      slide.addText(pcstr, { x:LP+1.5, y, w:LW-1.5, h:0.2, fontSize:7.5, fontFace:"Calibri", align:"right", bold:true, margin:0 });
      y += 0.22;
    }
  }

  // ── COL DER ──────────────────────────────────────────────────────────────
  let yr = 0.15;

  slide.addText("Glosario", {
    x:DER_X, y:yr, w:DER_W, h:0.27,
    fontSize:13, bold:true, color:C.negro, fontFace:"Calibri", margin:0,
  });
  yr += 0.3;
  slide.addShape("line", { x:DER_X, y:yr, w:DER_W, h:0, line:{ color:C.negro, width:1.5 } });
  yr += 0.1;

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
      x:DER_X, y:yr, w:DER_W, h:0.43,
      fontSize:7, color:C.grisOsc, fontFace:"Calibri",
      valign:"top", margin:0, wrap:true,
    });
    yr += 0.46;
  }

  yr += 0.08;
  lin(slide, yr, DER_X, DER_W, C.grisLin, 0.5); yr += 0.14;

  slide.addText("Disclaimer", {
    x:DER_X, y:yr, w:DER_W, h:0.25,
    fontSize:13, bold:true, color:C.negro, fontFace:"Calibri", margin:0,
  });
  yr += 0.28;
  lin(slide, yr, DER_X, DER_W, C.negro, 1.5); yr += 0.1;

  slide.addText("Conforme a la Ley Única de Fondos, las administradoras de fondos de inversión privados están sujetas a las obligaciones de información establecidas por la Comisión para el Mercado Financiero. Tales fondos no están sometidos a fiscalización de la Comisión y no hacemos oferta pública de sus cuotas.", {
    x:DER_X, y:yr, w:DER_W, h:0.7,
    fontSize:7.5, color:C.grisOsc, fontFace:"Calibri",
    valign:"top", margin:0, wrap:true,
  });
}

// ── Main ──────────────────────────────────────────────────────────────────────
async function main() {
  const ai = process.argv.indexOf("--data");
  const oi = process.argv.indexOf("--out");
  if (ai===-1||oi===-1) {
    console.error("Uso: node generar_folleto.js --data datos.json --out output.pptx");
    process.exit(1);
  }

  const d    = JSON.parse(fs.readFileSync(process.argv[ai+1], "utf8"));
  const pres = new pptxgen();
  pres.defineLayout({ name:"A4P", width:8.27, height:11.69 });
  pres.layout = "A4P";

  slide1(pres, d);
  slide2(pres, d);

  await pres.writeFile({ fileName: process.argv[oi+1] });
  console.log("OK:" + process.argv[oi+1]);
}

main().catch(e => { console.error(e); process.exit(1); });
