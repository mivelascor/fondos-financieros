"""
generador/pptx_builder.py — Pasa datos al generar_folleto.js.
"""
import subprocess, json
import numpy as np
from pathlib import Path

JS_DIR    = Path(__file__).parent
JS_SCRIPT = JS_DIR / "generar_folleto.js"


def _s(v):
    if v is None: return None
    if isinstance(v, float) and (np.isnan(v) or np.isinf(v)): return None
    return float(v)


def generar_pptx(nombre_fondo, periodo_str, comentario_pm,
                 datos_template, comp_cartera, info_fondo, out_path):
    """
    datos_template: resultado de template_reader.leer_template()
    comp_cartera:   resultado de excel_reader.get_cartera_composicion()
    """
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    nombre_corto = nombre_fondo.replace("FIP VANTRUST ", "").title()
    anio_actual  = periodo_str.split(" ")[-1] if " " in periodo_str else "2026"

    # Resumen: serializar floats
    resumen_json = []
    for r in (datos_template.get("resumen") or []):
        resumen_json.append({
            "nombre":   r.get("nombre",""),
            "m":        _s(r.get("m")),
            "t":        _s(r.get("t")),
            "s":        _s(r.get("s")),
            "a":        _s(r.get("a")),
            "ac":       _s(r.get("ac")),
            "es_icp":   bool(r.get("es_icp")),
            "es_comp":  bool(r.get("es_comp")),
            "es_fip":   bool(r.get("es_fip")),
        })

    # Histórico: serializar floats
    hist_json = []
    for año_data in (datos_template.get("historico") or []):
        filas_json = []
        for fila in año_data.get("filas", []):
            filas_json.append({
                "nombre": fila.get("nombre",""),
                "meses":  [_s(v) for v in (fila.get("meses") or [None]*12)],
                "total":  _s(fila.get("total")),
            })
        hist_json.append({"año": año_data["año"], "filas": filas_json})

    # Gráfico
    graf = datos_template.get("grafico") or {}
    grafico_json = {
        "labels": graf.get("labels") or [],
        "icp":    [_s(v) for v in (graf.get("icp")  or [])],
        "comp":   [_s(v) for v in (graf.get("comp") or [])],
        "fip":    [_s(v) for v in (graf.get("fip")  or [])],
    }

    datos = {
        "nombre_fondo":      nombre_fondo,
        "nombre_corto":      nombre_corto,
        "nombre_fip":        datos_template.get("nombre_fip", f"FIP {nombre_corto}"),
        "periodo":           periodo_str,
        "anio_acum":         anio_actual,
        "acum_label":        datos_template.get("acum_label", f"Acum. {anio_actual} (*)"),
        "comentario":        comentario_pm,
        # Info general
        "administradora":    info_fondo.get("administradora","Vantrust Gestion Patrimonial S.A."),
        "rut":               info_fondo.get("rut",""),
        "moneda":            info_fondo.get("moneda","CLP"),
        "tipo":              info_fondo.get("tipo","Fondo de Inversión Privado"),
        "fecha_inicio":      info_fondo.get("fecha_inicio",""),
        "benchmark":         info_fondo.get("benchmark","Índice Cámara Promedio (ICP)"),
        "plazo_rescate":     info_fondo.get("plazo_rescate","A más tardar 15 días corridos"),
        "remuneracion":      info_fondo.get("remuneracion","0,295% IVA Incluido"),
        "objetivo":          info_fondo.get("objetivo",""),
        "rentabilidad_texto":info_fondo.get("rentabilidad_texto",""),
        "inversionistas":    info_fondo.get("inversionistas",""),
        # Datos del template
        "resumen":           resumen_json,
        "historico":         hist_json,
        "grafico":           grafico_json,
        # Composición de cartera
        "comp_moneda":       comp_cartera.get("moneda",      []),
        "comp_instrumentos": comp_cartera.get("instrumentos",[]),
        "comp_duracion":     comp_cartera.get("duracion",    []),
    }

    json_path = out_path.parent / f"_tmp_{out_path.stem}.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(datos, f, ensure_ascii=False)

    try:
        r = subprocess.run(
            ["node", str(JS_SCRIPT), "--data", str(json_path), "--out", str(out_path)],
            capture_output=True, text=True, timeout=60, cwd=str(JS_DIR)
        )
        if r.returncode != 0:
            raise RuntimeError(f"JS error:\n{r.stderr}\n{r.stdout}")
        if not out_path.exists():
            raise FileNotFoundError(f"PPTX no creado: {out_path}")
    finally:
        if json_path.exists():
            json_path.unlink()

    return out_path
