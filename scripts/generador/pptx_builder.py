"""
generador/pptx_builder.py — Genera PPTX usando generar_folleto.js (pptxgenjs).
"""
import subprocess
import json
import numpy as np
import pandas as pd
from pathlib import Path

JS_DIR    = Path(__file__).parent
JS_SCRIPT = JS_DIR / "generar_folleto.js"


def _fmt(v) -> str:
    if v is None or (isinstance(v, float) and (np.isnan(v) or np.isinf(v))):
        return ""
    return f"{v*100:.2f}%".replace(".", ",")


def generar_pptx(
    nombre_fondo:    str,
    periodo_str:     str,
    comentario_pm:   str,
    b1000_fondo:     pd.Series,
    b1000_comp:      pd.Series,
    b1000_icp:       pd.Series,
    tabla_resumen:   pd.DataFrame,
    tabla_historica: list,
    comp_cartera:    dict,
    info_fondo:      dict,
    out_path:        Path,
) -> Path:
    """
    Genera el PPTX de un fondo usando generar_folleto.js.
    """
    out_path.parent.mkdir(parents=True, exist_ok=True)

    nombre_corto = nombre_fondo.replace("FIP VANTRUST ", "").title()
    anio_actual  = int(periodo_str.split(" ")[-1]) if " " in periodo_str else 2026

    # ── Gráfico (base 1000) ────────────────────────────────────────────────────
    grafico_labels, grafico_fondo, grafico_icp, grafico_comp = [], [], [], []
    for f in sorted(b1000_fondo.index):
        vf = float(b1000_fondo.get(f, np.nan))
        if not np.isnan(vf):
            grafico_labels.append(pd.Timestamp(f).strftime("%b %Y"))
            grafico_fondo.append(round(vf, 2))
            vi = float(b1000_icp.get(f, np.nan)) if not b1000_icp.empty and f in b1000_icp.index else None
            vc = float(b1000_comp.get(f, np.nan)) if not b1000_comp.empty and f in b1000_comp.index else None
            grafico_icp.append(round(vi, 2) if vi and not np.isnan(vi) else None)
            grafico_comp.append(round(vc, 2) if vc and not np.isnan(vc) else None)

    # ── Tabla resumen ──────────────────────────────────────────────────────────
    tabla_rentab = []
    if tabla_resumen is not None and not tabla_resumen.empty:
        for _, row in tabla_resumen.iterrows():
            tabla_rentab.append({
                "nombre":      row.get("nombre", ""),
                "mensual":     row.get("mensual", ""),
                "trimestral":  row.get("trimestral", ""),
                "semestral":   row.get("semestral", ""),
                "anual":       row.get("anual", ""),
                "ytd":         row.get("ytd", ""),
                "es_fondo":    bool(row.get("es_fondo", False)),
            })

    # ── JSON de datos ──────────────────────────────────────────────────────────
    datos = {
        "nombre_fondo":      nombre_fondo,
        "nombre_corto":      nombre_corto,
        "administradora":    info_fondo.get("administradora", "Vantrust Gestion Patrimonial S.A."),
        "rut":               info_fondo.get("rut", "76,637,334-8"),
        "moneda":            info_fondo.get("moneda", "CLP"),
        "tipo":              info_fondo.get("tipo", "Fondo de Inversión Privado"),
        "fecha_inicio":      info_fondo.get("fecha_inicio", ""),
        "benchmark":         info_fondo.get("benchmark", "Indice Camara Promedio (ICP)"),
        "plazo_rescate":     info_fondo.get("plazo_rescate", "A mas tardar 15 dias corridos"),
        "remuneracion":      info_fondo.get("remuneracion", "0,295% IVA Incluido"),
        "objetivo":          "Invertir los recursos del fondo en instrumentos de deuda de corto y mediano plazo, en una cartera diversificada, obteniendo una rentabilidad igual o superior al ICP.",
        "rentabilidad_texto": f"La rentabilidad esperada del {nombre_fondo}, es la tasa de politica monetaria promedio del Banco Central de Chile.",
        "inversionistas":    "Dirigida a empresas y personas que buscan invertir sus excedentes de caja con una rentabilidad de corto plazo y baja tolerancia al riesgo.",
        "comentario":        comentario_pm,
        "anio_acum":         str(anio_actual),
        "grafico_labels":    grafico_labels,
        "grafico_fondo":     grafico_fondo,
        "grafico_icp":       [x for x in grafico_icp  if x is not None],
        "grafico_comp":      [x for x in grafico_comp if x is not None],
        "tabla_rentab":      tabla_rentab,
        "tabla_historica":   tabla_historica,
        "comp_moneda":       comp_cartera.get("moneda", []),
        "comp_duracion":     comp_cartera.get("duracion", []),
        "comp_instrumento":  comp_cartera.get("instrumento", []),
    }

    json_path = out_path.parent / f"_tmp_{out_path.stem}.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(datos, f, ensure_ascii=False)

    try:
        result = subprocess.run(
            ["node", str(JS_SCRIPT), "--data", str(json_path), "--out", str(out_path)],
            capture_output=True, text=True, timeout=60,
            cwd=str(JS_DIR)
        )
        if result.returncode != 0:
            raise RuntimeError(f"generar_folleto.js fallo:\n{result.stderr}\n{result.stdout}")
        if not out_path.exists():
            raise FileNotFoundError(f"PPTX no generado: {out_path}")
    finally:
        if json_path.exists():
            json_path.unlink()

    return out_path
