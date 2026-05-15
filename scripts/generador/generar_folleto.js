/**
 * generar_folleto.js  — v3
 *
 * Replica el formato exacto de los folletos comerciales VanTrust (PDF de referencia):
 *   - Slide 1 : Panel negro izq | Logo + nombre | Comentario PM | Gráfico índices |
 *               Tabla resumen (ICP/Comp/FIP) | Tabla histórica completa
 *   - Slide 2 : Panel negro izq | Composición (Moneda / Instrumentos / Duración) |
 *               Glosario | Disclaimer
 *
 * Layout: Portrait A4 en PPT  (7.5" × 10.0")
 *
 * Uso:
 *   node generar_folleto.js <datos.json> <config_fondo.json> [output.pptx]
 *
 * donde datos.json   = salida de extraer_datos.py
 *       config.json  = datos de información general del fondo (RUT, inicio, etc.)
 */

'use strict';

const pptxgen = require('pptxgenjs');
const fs      = require('fs');
const path    = require('path');

// ─── Dimensiones (portrait A4 en pulgadas) ───────────────────────────────────
const SW = 7.5;   // slide width
const SH = 10.0;  // slide height

// Panel izquierdo negro
const PW  = 2.55;  // panel width
const PAD = 0.18;  // padding interno panel

// Área de contenido derecha
const CX  = PW + 0.22;
const CW  = SW - CX - 0.18;

// ─── Colores ──────────────────────────────────────────────────────────────────
const BK = '000000';      // negro puro (panel izq)
const WH = 'FFFFFF';      // blanco
const GR = 'F0F0F0';      // gris muy claro (fila ICP)
const MG = 'DCDCDC';      // gris medio (fila Competencia)
const LG = '808080';      // gris texto secundario
const BL = '003366';      // azul encabezados tabla
const GN = '2B7D2B';      // verde barras chart

// Para tabla histórica: formato compacto sin "%", 2 decimales con coma
const pctCompact = (v) => v != null ? (v * 100).toFixed(2).replace('.', ',') : '';
// Para tabla resumen: con "%"
const pct = (v, d = 2) => v != null ? (v * 100).toFixed(d).replace('.', ',') + '%' : '';

function makeShadow() {
  return { type: 'outer', color: '000000', blur: 4, offset: 2, angle: 135, opacity: 0.10 };
}

// ─── SLIDE 1 ─────────────────────────────────────────────────────────────────
function buildSlide1(pres, datos, cfg) {
  const slide = pres.addSlide();
  slide.background = { color: WH };

  const { nombre_fondo, comentario, resumen, historico } = datos;

  /* ── Panel izquierdo negro ── */
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: PW, h: SH,
    fill: { color: BK }, line: { color: BK, width: 0 },
  });

  /* Logo "VanTrust Capital" arriba izquierdo */
  slide.addText('VanTrust Capital', {
    x: PAD, y: 0.22, w: PW - PAD * 2, h: 0.22,
    color: WH, fontSize: 8, fontFace: 'Calibri', bold: false, margin: 0,
  });

  /* Logo VNT (caja blanca) */
  slide.addShape(pres.shapes.RECTANGLE, {
    x: PW - 0.58, y: 0.16, w: 0.44, h: 0.30,
    fill: { color: WH }, line: { color: WH, width: 0 },
  });
  slide.addText('VNT', {
    x: PW - 0.56, y: 0.18, w: 0.40, h: 0.22,
    color: BK, fontSize: 8, fontFace: 'Calibri', bold: true, align: 'center', margin: 0,
  });

  /* Nombre del fondo (grande, blanco) */
  const nombreParts = nombre_fondo.replace('FIP ', '').split(' ');
  // Dividir en líneas de 1-2 palabras para el estilo visual
  const lineas = [];
  for (let i = 0; i < nombreParts.length; i += 2) {
    lineas.push(nombreParts.slice(i, i + 2).join(' '));
  }
  const nombreTexto = lineas.map((l, i) => ({
    text: l,
    options: { bold: true, breakLine: i < lineas.length - 1, fontSize: 22, color: WH },
  }));

  slide.addText(nombreTexto, {
    x: PAD, y: 0.65, w: PW - PAD * 2, h: 1.2,
    fontFace: 'Georgia',
    valign: 'top',
    margin: 0,
  });

  /* "F O N D O" subtítulo */
  slide.addText('F O N D O', {
    x: PAD, y: 1.90, w: PW - PAD * 2, h: 0.22,
    color: WH, fontSize: 7.5, fontFace: 'Calibri', bold: false,
    align: 'center', charSpacing: 3, margin: 0,
  });

  /* Línea separadora */
  slide.addShape(pres.shapes.LINE, {
    x: PAD, y: 2.15, w: PW - PAD * 2, h: 0,
    line: { color: WH, width: 0.5 },
  });

  /* Información General */
  const infoY = 2.25;
  slide.addText('Información General', {
    x: PAD, y: infoY, w: PW - PAD * 2, h: 0.2,
    color: WH, fontSize: 8, fontFace: 'Calibri', bold: true, margin: 0,
  });
  slide.addShape(pres.shapes.LINE, {
    x: PAD, y: infoY + 0.21, w: PW - PAD * 2, h: 0,
    line: { color: WH, width: 0.3 },
  });

  const infoFields = [
    ['Administradora',   'Vantrust Gestion Patrimonial S.A.'],
    ['RUT Fondo',        cfg.rut || ''],
    ['Moneda',           cfg.moneda || ''],
    ['Tipo de Fondo',    cfg.tipo || ''],
    ['Fecha Inicio',     cfg.inicio || ''],
    ['Benchmark',        cfg.benchmark || ''],
    ['Fondo Rescatable', cfg.rescatable || ''],
    ['Plazo Rescate',    cfg.plazo || ''],
    ['Riesgos',          cfg.riesgos || ''],
    ['Remuneración',     cfg.remuneracion || ''],
    ['Custodio',         cfg.custodio || ''],
  ];

  const fieldY0 = infoY + 0.24;
  const fieldH  = 0.185;
  const lblW    = 0.85;
  infoFields.forEach(([lbl, val], i) => {
    const fy = fieldY0 + i * fieldH;
    slide.addText(lbl, {
      x: PAD, y: fy, w: lblW, h: fieldH,
      color: WH, fontSize: 6, fontFace: 'Calibri', bold: true,
      valign: 'top', margin: 0,
    });
    slide.addText(val, {
      x: PAD + lblW, y: fy, w: PW - PAD * 2 - lblW, h: fieldH,
      color: WH, fontSize: 6, fontFace: 'Calibri', bold: false,
      valign: 'top', margin: 0, wrap: true,
    });
  });

  /* Línea separadora */
  const sepY1 = fieldY0 + infoFields.length * fieldH + 0.06;
  slide.addShape(pres.shapes.LINE, {
    x: PAD, y: sepY1, w: PW - PAD * 2, h: 0,
    line: { color: WH, width: 0.3 },
  });

  /* Objetivo */
  const objY = sepY1 + 0.08;
  slide.addText('Objetivo', {
    x: PAD, y: objY, w: PW - PAD * 2, h: 0.18,
    color: WH, fontSize: 8, fontFace: 'Calibri', bold: true, margin: 0,
  });
  slide.addText(cfg.objetivo || '', {
    x: PAD, y: objY + 0.2, w: PW - PAD * 2, h: 0.55,
    color: WH, fontSize: 6, fontFace: 'Calibri', bold: false,
    valign: 'top', margin: 0, wrap: true,
  });

  /* Línea */
  const sepY2 = objY + 0.78;
  slide.addShape(pres.shapes.LINE, {
    x: PAD, y: sepY2, w: PW - PAD * 2, h: 0,
    line: { color: WH, width: 0.3 },
  });

  /* Rentabilidad (descripción) */
  const rentY = sepY2 + 0.08;
  slide.addText('Rentabilidad', {
    x: PAD, y: rentY, w: PW - PAD * 2, h: 0.18,
    color: WH, fontSize: 8, fontFace: 'Calibri', bold: true, margin: 0,
  });
  slide.addText(cfg.rentabilidad_desc || '', {
    x: PAD, y: rentY + 0.2, w: PW - PAD * 2, h: 0.45,
    color: WH, fontSize: 6, fontFace: 'Calibri', bold: false,
    valign: 'top', margin: 0, wrap: true,
  });

  /* Línea */
  const sepY3 = rentY + 0.68;
  slide.addShape(pres.shapes.LINE, {
    x: PAD, y: sepY3, w: PW - PAD * 2, h: 0,
    line: { color: WH, width: 0.3 },
  });

  /* Inversionistas */
  const invY = sepY3 + 0.08;
  slide.addText('Inversionistas', {
    x: PAD, y: invY, w: PW - PAD * 2, h: 0.18,
    color: WH, fontSize: 8, fontFace: 'Calibri', bold: true, margin: 0,
  });
  slide.addText(cfg.inversionistas || '', {
    x: PAD, y: invY + 0.2, w: PW - PAD * 2, h: 0.5,
    color: WH, fontSize: 6, fontFace: 'Calibri', bold: false,
    valign: 'top', margin: 0, wrap: true,
  });

  /* ── Contenido derecho — posiciones calibradas contra folleto PDF ── */

  // Coordenadas calibradas — cadena sin superposición verificada
  const COM_TITULO_Y  = 0.32;
  const COM_TEXTO_Y   = 0.64;
  const COM_TEXTO_H   = 1.06;   // fin = 1.70
  const CHART_TITLE_Y = 1.80;
  const CHART_Y       = 1.86;   // gap 0.16" sobre fin comentario
  const CHART_H       = 1.30;   // fin = 3.16
  const EVOL_Y        = 3.28;   // gap 0.12"
  const EVOL_LINE_Y   = 3.56;
  const RES_Y         = 3.60;
  const RES_ROW_H     = 0.20;   // 4 filas × 0.20 = 0.80 → fin tabla = 4.40
  const RES_N_ROWS    = 4;
  const NOTA_Y        = 4.48;   // gap 0.08" bajo tabla
  const HIST_LINE_Y   = 4.68;
  const HIST_Y        = 4.74;

  /* Comentario Portafolio Manager */
  slide.addText('Comentario Portafolio Manager', {
    x: CX, y: COM_TITULO_Y, w: CW, h: 0.28,
    color: BK, fontSize: 13, fontFace: 'Georgia', bold: true, margin: 0,
  });

  slide.addText(comentario || '', {
    x: CX, y: COM_TEXTO_Y, w: CW, h: COM_TEXTO_H,
    color: '333333', fontSize: 7, fontFace: 'Calibri',
    valign: 'top', margin: 0, wrap: true,
  });

  /* Gráfico de evolución */
  const chartData = buildChartData(historico, nombre_fondo);

  // Título del gráfico como título nativo del chart (no texto flotante)
  // así no interfiere con el comentario
  const fondoSinFIP = nombre_fondo.replace(/^FIP\s+/, '');

  // Gráfico: líneas finas (0.5-0.75pt), sin marcadores, colores del PDF
  slide.addChart(pres.charts.LINE, chartData, {
    x: CX, y: CHART_Y, w: CW, h: CHART_H,
    lineSize: 0.75,
    chartColors: ['AAAAAA', '000000', '888888'],
    showLegend: true,
    legendPos: 'b',
    legendFontSize: 5.5,
    legendColor: '444444',
    catAxisLabelColor: '888888',
    catAxisLabelFontSize: 5,
    valAxisLabelColor: '888888',
    valAxisLabelFontSize: 5,
    valGridLine: { color: 'EEEEEE', size: 0.25 },
    catGridLine: { style: 'none' },
    plotArea: { fill: { color: WH } },
    chartArea: { fill: { color: WH } },
    showTitle: true,
    title: fondoSinFIP,
    titleFontSize: 7,
    titleColor: BK,
    lineDataSymbol: 'none',
  });

  /* Título "Evolución Rentabilidad" */
  slide.addText('Evolución Rentabilidad', {
    x: CX, y: EVOL_Y, w: CW, h: 0.26,
    color: BK, fontSize: 12, fontFace: 'Georgia', bold: true, margin: 0,
  });
  slide.addShape(pres.shapes.LINE, {
    x: CX, y: EVOL_LINE_Y, w: CW, h: 0,
    line: { color: 'CCCCCC', width: 0.5 },
  });

  /* Tabla resumen */
  buildTablaResumen(pres, slide, resumen, nombre_fondo, CX, RES_Y, CW, RES_ROW_H);

  /* Nota anualizada */
  slide.addText('*Valores correspondientes a la rentabilidad anualizada, no acumulada', {
    x: CX, y: NOTA_Y, w: CW, h: 0.16,
    color: LG, fontSize: 5.5, fontFace: 'Calibri', italic: true, margin: 0,
  });

  /* Línea separadora */
  slide.addShape(pres.shapes.LINE, {
    x: CX, y: HIST_LINE_Y, w: CW, h: 0,
    line: { color: 'CCCCCC', width: 0.5 },
  });

  /* Tabla histórica */
  buildTablaHistorica(pres, slide, historico, nombre_fondo, CX, HIST_Y, CW);
}

// ─── Chart data ──────────────────────────────────────────────────────────────
function buildChartData(historico, nombre_fondo) {
  // Construir índice base 100 desde la primera observación disponible
  const porFondo = {};
  historico.forEach(r => {
    if (!porFondo[r.fondo]) porFondo[r.fondo] = [];
    r.meses.forEach((v, mi) => {
      if (v !== null && v !== undefined) {
        porFondo[r.fondo].push({ año: r.año, mes: mi, v });
      }
    });
  });

  const MESES_ABREV = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
  const labels = [];
  const normalize = (serie) => {
    let idx = 100;
    return serie.map(pt => {
      idx = idx * (1 + pt.v);
      return idx;
    });
  };

  // Tomar primeros 3 fondos disponibles
  const fondoKeys = Object.keys(porFondo).slice(0, 3);

  // Generar labels desde la primera observación del primer fondo
  if (fondoKeys.length > 0) {
    porFondo[fondoKeys[0]].forEach(pt => {
      labels.push(MESES_ABREV[pt.mes] + ' ' + String(pt.año).slice(2));
    });
  }

  return fondoKeys.map(nombre => ({
    name: nombre === nombre_fondo ? nombre_fondo : nombre,
    labels,
    values: normalize(porFondo[nombre]),
  }));
}

// ─── Tabla resumen (addText posicionado — altura exacta garantizada) ──────────
function buildTablaResumen(pres, slide, resumen, nombre_fondo, x, y, w, rowH = 0.20) {
  const COLS = ['Mensual', 'Trimestral', 'Semestral', 'Anual', 'Acum\n2026 (*)'];
  const colN = w * 0.30;
  const colV = (w - colN) / 5;

  const cell = (t, xp, yp, wp, opts = {}) =>
    slide.addText(t, { x: xp, y: yp, w: wp, h: rowH, fontSize: 7.5, fontFace: 'Calibri',
      color: BK, valign: 'middle', margin: 0, ...opts });

  // — Header —
  cell('Rentabilidad', x, y, colN, { align: 'left' });
  COLS.forEach((c, i) => cell(c, x + colN + i * colV, y, colV, { align: 'center', fontSize: 7 }));
  slide.addShape(pres.shapes.LINE, { x, y: y + rowH, w, h: 0, line: { color: 'CCCCCC', width: 0.5 } });

  // — Filas de datos —
  Object.entries(resumen).forEach(([nombre, v], ri) => {
    const esFip = nombre.toUpperCase().includes('FIP') || nombre === nombre_fondo;
    const ry    = y + rowH * (ri + 1);
    const vals  = [v.mensual, v.trimestral, v.semestral, v.anual, v.acum_ytd];
    cell(nombre, x, ry, colN, { align: 'left', bold: esFip });
    vals.forEach((vv, i) =>
      cell(pct(vv), x + colN + i * colV, ry, colV, { align: 'center', bold: esFip }));
    slide.addShape(pres.shapes.LINE,
      { x, y: ry + rowH, w, h: 0, line: { color: 'EEEEEE', width: 0.3 } });
  });
}

// ─── Tabla histórica (usando addText posicionado para evitar word-wrap) ───────
function buildTablaHistorica(pres, slide, historico, nombre_fondo, x, y, w) {
  const MESES = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];

  // Agrupar por año
  const porAño = {};
  historico.forEach(r => {
    if (!porAño[r.año]) porAño[r.año] = {};
    porAño[r.año][r.fondo] = r;
  });
  const años = Object.keys(porAño).map(Number).sort((a, b) => a - b);

  // Anchos fijos de columnas
  const colAño   = 0.30;
  const colFondo = 1.28;
  const colTotal = 0.50;
  const colMes   = (w - colAño - colFondo - colTotal) / 12;

  // Posiciones X de cada columna
  const colX = [x];
  colX.push(colX[0] + colAño);
  MESES.forEach((_, i) => colX.push(colX[1] + colFondo + i * colMes));
  colX.push(colX[1] + colFondo + 12 * colMes);  // Total

  const ROW_H = 0.128;
  const FONT  = 5.5;
  const FONT_H = 5.5;

  // ── Fila header ──
  const hdrOpts = (txt, xp, wp, al = 'center') => ({
    text: txt, x: xp, y, w: wp, h: ROW_H,
    color: BK, fontSize: FONT_H, fontFace: 'Calibri', bold: false, align: al,
    valign: 'middle', margin: 0,
  });

  slide.addText('Año',       { ...hdrOpts('Año',       colX[0],  colAño,   'left')  });
  slide.addText('Fondo',     { ...hdrOpts('Fondo',     colX[1],  colFondo, 'left')  });
  MESES.forEach((m, i) => {
    slide.addText(m, { ...hdrOpts(m, colX[2 + i], colMes) });
  });
  slide.addText('Total\nAño', { ...hdrOpts('Total\nAño', colX[colX.length - 1], colTotal) });

  // Línea bajo header
  slide.addShape(pres.shapes.LINE, {
    x, y: y + ROW_H, w, h: 0,
    line: { color: 'AAAAAA', width: 0.5 },
  });

  let rowY = y + ROW_H + 0.01;

  años.forEach(año => {
    const grupo  = porAño[año];
    const fondos = Object.keys(grupo);

    fondos.forEach((fondoNombre, idx) => {
      const r      = grupo[fondoNombre];
      const esFip  = fondoNombre === nombre_fondo;
      const color  = BK;
      const bold   = esFip;

      const cellOpts = (txt, xp, wp, al = 'center') => ({
        text: txt, x: xp, y: rowY, w: wp, h: ROW_H,
        color, fontSize: FONT, fontFace: 'Calibri', bold, align: al,
        valign: 'middle', margin: 0,
      });

      // Año (solo primera fila del grupo)
      if (idx === 0) {
        slide.addText(String(año), {
          ...cellOpts(String(año), colX[0], colAño, 'left'), bold: false,
        });
      }

      // Nombre fondo
      slide.addText(fondoNombre, { ...cellOpts(fondoNombre, colX[1], colFondo, 'left') });

      // Meses
      r.meses.forEach((v, mi) => {
        const txt = v != null ? (v * 100).toFixed(2).replace('.', ',') : '';
        slide.addText(txt, { ...cellOpts(txt, colX[2 + mi], colMes) });
      });

      // Total año
      const tot = r.total_año != null ? (r.total_año * 100).toFixed(2).replace('.', ',') : '';
      slide.addText(tot, { ...cellOpts(tot, colX[colX.length - 1], colTotal) });

      // Línea separadora ligera entre fondos del mismo año
      if (idx < fondos.length - 1) {
        slide.addShape(pres.shapes.LINE, {
          x, y: rowY + ROW_H, w, h: 0,
          line: { color: 'EEEEEE', width: 0.3 },
        });
      }

      rowY += ROW_H;
    });

    // Línea entre años
    slide.addShape(pres.shapes.LINE, {
      x, y: rowY + 0.005, w, h: 0,
      line: { color: 'CCCCCC', width: 0.5 },
    });
    rowY += 0.01;
  });
}

// ─── SLIDE 2 ─────────────────────────────────────────────────────────────────
function buildSlide2(pres, datos, cfg) {
  const slide = pres.addSlide();
  slide.background = { color: WH };

  const { nombre_fondo, composicion } = datos;
  const { duracion, instrumentos, moneda } = composicion;

  /* Panel negro izquierdo (sin texto) */
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: PW, h: SH,
    fill: { color: BK }, line: { color: BK, width: 0 },
  });

  /* "F O N D O" al fondo del panel */
  slide.addText('F O N D O', {
    x: PAD, y: SH - 0.35, w: PW - PAD * 2, h: 0.22,
    color: WH, fontSize: 7.5, fontFace: 'Calibri', align: 'center', charSpacing: 3, margin: 0,
  });
  slide.addShape(pres.shapes.LINE, {
    x: PAD, y: SH - 0.38, w: PW - PAD * 2, h: 0,
    line: { color: WH, width: 0.3 },
  });

  /* Composición por Moneda */
  let leftY = 0.22;
  slide.addText('Composición por Moneda', {
    x: PAD, y: leftY, w: PW - PAD * 2, h: 0.2,
    color: WH, fontSize: 8, fontFace: 'Calibri', bold: true, margin: 0,
  });
  slide.addShape(pres.shapes.LINE, {
    x: PAD, y: leftY + 0.22, w: PW - PAD * 2, h: 0,
    line: { color: WH, width: 0.3 },
  });
  leftY += 0.28;
  Object.entries(moneda).filter(([, v]) => v > 0).forEach(([k, v]) => {
    slide.addText(`${k}`, { x: PAD, y: leftY, w: PW - PAD * 2 - 0.7, h: 0.18, color: WH, fontSize: 7, fontFace: 'Calibri', margin: 0 });
    slide.addText(pct(v), { x: PW - PAD - 0.68, y: leftY, w: 0.65, h: 0.18, color: WH, fontSize: 7, fontFace: 'Calibri', align: 'right', margin: 0 });
    leftY += 0.19;
  });

  /* Línea */
  leftY += 0.12;
  slide.addShape(pres.shapes.LINE, { x: PAD, y: leftY, w: PW - PAD * 2, h: 0, line: { color: WH, width: 0.3 } });
  leftY += 0.10;

  /* Composición General por Instrumentos */
  slide.addText('Composición General por Instrumentos Financieros', {
    x: PAD, y: leftY, w: PW - PAD * 2, h: 0.30,
    color: WH, fontSize: 8, fontFace: 'Calibri', bold: true, margin: 0, wrap: true,
  });
  slide.addShape(pres.shapes.LINE, { x: PAD, y: leftY + 0.32, w: PW - PAD * 2, h: 0, line: { color: WH, width: 0.3 } });
  leftY += 0.38;

  /* Header límites */
  slide.addText('Límites de Inversión', { x: PAD, y: leftY, w: PW - PAD * 2, h: 0.16, color: WH, fontSize: 6.5, fontFace: 'Calibri', bold: true, align: 'center', margin: 0 });
  leftY += 0.17;
  slide.addText('Mínimo', { x: PAD + (PW - PAD*2) * 0.55, y: leftY, w: (PW-PAD*2)*0.22, h: 0.15, color: WH, fontSize: 6, fontFace: 'Calibri', bold: true, align: 'center', margin: 0 });
  slide.addText('Máximo', { x: PAD + (PW - PAD*2) * 0.78, y: leftY, w: (PW-PAD*2)*0.22, h: 0.15, color: WH, fontSize: 6, fontFace: 'Calibri', bold: true, align: 'center', margin: 0 });
  leftY += 0.16;

  const LIMITES = [
    ['Bonos',                      '0%',  '25%'],
    ['Depósitos a Plazo',          '5%',  '100%'],
    ['Facturas',                   '0%',  '40%'],
    ['Financiamientos',            '0%',  '100%'],
    ['Efectos de Comercio',        '0%',  '40%'],
    ['Pagaré Banco Central',       '0%',  '100%'],
    ['Caja',                       '5%',  '25%'],
  ];
  LIMITES.forEach(([nombre, minv, maxv]) => {
    const lw = PW - PAD * 2;
    slide.addText(nombre,  { x: PAD,               y: leftY, w: lw*0.53, h: 0.165, color: WH, fontSize: 6, fontFace: 'Calibri', margin: 0 });
    slide.addText(minv,    { x: PAD + lw*0.55,     y: leftY, w: lw*0.22, h: 0.165, color: WH, fontSize: 6, fontFace: 'Calibri', align: 'center', margin: 0 });
    slide.addText(maxv,    { x: PAD + lw*0.78,     y: leftY, w: lw*0.22, h: 0.165, color: WH, fontSize: 6, fontFace: 'Calibri', align: 'center', margin: 0 });
    leftY += 0.165;
  });

  /* Línea */
  leftY += 0.10;
  slide.addShape(pres.shapes.LINE, { x: PAD, y: leftY, w: PW - PAD * 2, h: 0, line: { color: WH, width: 0.3 } });
  leftY += 0.10;

  /* Composición por Duración */
  slide.addText('Composición por Duración', {
    x: PAD, y: leftY, w: PW - PAD * 2, h: 0.20,
    color: WH, fontSize: 8, fontFace: 'Calibri', bold: true, margin: 0,
  });
  slide.addShape(pres.shapes.LINE, { x: PAD, y: leftY + 0.22, w: PW - PAD * 2, h: 0, line: { color: WH, width: 0.3 } });
  leftY += 0.27;

  const durItems = Object.entries(duracion).filter(([, v]) => v >= 0);
  durItems.forEach(([k, v]) => {
    slide.addText(k,      { x: PAD, y: leftY, w: PW - PAD*2 - 0.7, h: 0.18, color: WH, fontSize: 7, fontFace: 'Calibri', margin: 0 });
    slide.addText(pct(v), { x: PW - PAD - 0.68, y: leftY, w: 0.65, h: 0.18, color: WH, fontSize: 7, fontFace: 'Calibri', align: 'right', bold: true, margin: 0 });
    leftY += 0.19;
  });

  /* ── Contenido derecho ── */

  /* Glosario */
  slide.addText('Glosario', {
    x: CX, y: 0.22, w: CW, h: 0.26,
    color: BK, fontSize: 12, fontFace: 'Georgia', bold: true, margin: 0,
  });

  const GLOSARIO = [
    ['Riesgo de Mercado',
     'Este es el riesgo de una variación adversa en el precio o tasa de mercado en los instrumentos en que invierte el Fondo'],
    ['Riesgo de Crédito',
     'Es la posible pérdida que se asume como consecuencia del incumplimiento de las obligaciones directas, indirectas o de derivados que conllevan al no pago parcial, total o falta de oportunidad del pago de los emisores de los instrumentos en que invierte el Fondo y que pudieran ocasionar una pérdida financiera.'],
    ['Riesgo de Liquidez',
     'Riesgo asociado a la capacidad de generación de recursos del Fondo para cumplir con sus obligaciones de rescate o vencimiento del mismo.'],
    ['Riesgo Profundidad de Mercado',
     'Corresponde a la posibilidad de comprar o vender un activo financiero al valor de mercado en un período de tiempo acorde a las características del instrumento. Períodos de alta volatilidad de mercado afectan negativamente estos tiempos.'],
    ['Riesgo de Tasa de interés',
     'Exposición a pérdidas ocasionadas por cambios adversos en las tasas de interés de mercado y que afecten el valor de los instrumentos, contratos y demás operaciones registradas en el balance.'],
    ['Riesgo de Moneda',
     'Es la exposición a pérdidas ocasionadas por cambios adversos en el valor en moneda nacional de las monedas extranjeras que están expresados a los instrumentos, contratos y demás operaciones del balance del Fondo.'],
    ['Riesgo Sectorial',
     'Este riesgo está asociado a condiciones de mercado adversas que pueden afectar a un sector industrial en particular y por ende la rentabilidad del Fondo.'],
    ['Gastos del fondo',
     'Corresponden a los gastos directos e indirectos necesarios para el correcto funcionamiento del fondo los que están detallados en el Reglamento Interno'],
    ['Forma de Ingreso y Pago del Fondo',
     'La moneda en que el inversionista entra al fondo es aportando pesos chilenos, y al rescate de las cuotas, el fondo le entrega pesos chilenos.'],
  ];

  let glosY = 0.52;
  GLOSARIO.forEach(([term, def]) => {
    slide.addText(term, {
      x: CX, y: glosY, w: CW, h: 0.155,
      color: BK, fontSize: 7, fontFace: 'Calibri', bold: true, margin: 0,
    });
    glosY += 0.155;
    slide.addText(def, {
      x: CX, y: glosY, w: CW, h: 0.38,
      color: '333333', fontSize: 6.5, fontFace: 'Calibri', bold: false,
      valign: 'top', margin: 0, wrap: true,
    });
    glosY += 0.40;
  });

  /* Línea */
  slide.addShape(pres.shapes.LINE, {
    x: CX, y: glosY + 0.05, w: CW, h: 0,
    line: { color: 'CCCCCC', width: 0.5 },
  });
  glosY += 0.15;

  /* Disclaimer */
  slide.addText('Disclaimer', {
    x: CX, y: glosY, w: CW, h: 0.24,
    color: BK, fontSize: 12, fontFace: 'Georgia', bold: true, margin: 0,
  });
  glosY += 0.28;
  slide.addText(
    'Conforme a la Ley Única de Fondos, las administradoras de fondos de inversión privados están ' +
    'sujetas a las obligaciones de información establecidas por la Comisión para el Mercado ' +
    'Financiero. Tales fondos no están sometidos a fiscalización de la Comisión y no hacemos oferta ' +
    'pública de sus cuotas',
    {
      x: CX, y: glosY, w: CW, h: 0.55,
      color: '333333', fontSize: 7, fontFace: 'Calibri', valign: 'top', margin: 0, wrap: true,
    }
  );
}

// ─── Main ──────────────────────────────────────────────────────────────────────
async function main() {
  const args = process.argv.slice(2);
  if (args.length < 2) {
    console.error('Uso: node generar_folleto.js <datos.json> <config_fondo.json> [output.pptx]');
    process.exit(1);
  }

  const datos = JSON.parse(fs.readFileSync(args[0], 'utf8'));
  const cfg   = JSON.parse(fs.readFileSync(args[1], 'utf8'));

  const nombreArchivo = args[2]
    || ('folleto_' + datos.nombre_fondo.replace(/\s+/g, '_') + '.pptx');

  const pres = new pptxgen();
  pres.layout  = 'LAYOUT_4x3';    // 10×7.5 — lo sobreescribimos manualmente
  // Portrait A4: 7.5" × 10"
  pres.defineLayout({ name: 'PORTRAIT', width: SW, height: SH });
  pres.layout  = 'PORTRAIT';
  pres.title   = datos.nombre_fondo;
  pres.author  = 'VanTrust Asset Management';

  buildSlide1(pres, datos, cfg);
  buildSlide2(pres, datos, cfg);

  await pres.writeFile({ fileName: nombreArchivo });
  console.log('✓ ' + nombreArchivo);
}

main().catch(err => { console.error(err); process.exit(1); });
