"""
generador/pptx_builder.py — Genera PPTX usando el generador JS (pptxgenjs).
Fiel al diseño de referencia: layout A4, 2 columnas, franja negra, gráfico nativo.
"""
import subprocess
import json
import os
import numpy as np
import pandas as pd
from pathlib import Path

# Ruta al script JS (en la misma carpeta que este archivo)
JS_DIR    = Path(__file__).parent
JS_SCRIPT = JS_DIR / "generar_folleto.js"


def _fmt_pct(v) -> str:
    """Formatea un float decimal como porcentaje con coma decimal."""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return ""
    return f"{v*100:+.2f}%".replace(".", ",")


def _fmt_pct_pos(v) -> str:
    """Igual pero sin signo +."""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return ""
    return f"{v*100:.2f}%".replace(".", ",")


def construir_datos_json(
    nombre_fondo:   str,
    periodo_str:    str,
    comentario_pm:  str,
    b1000_fondo:    pd.Series,
    b1000_comp:     pd.Series,
    b1000_icp:      pd.Series,
    metricas:       pd.DataFrame,
    df_rentm:       pd.DataFrame,
    df_cartera:     pd.DataFrame,
    info_fondo:     dict,
) -> dict:
    """
    Construye el dict JSON que espera generar_folleto.js
    """
    nombre_corto = nombre_fondo.replace("FIP VANTRUST ", "").title()
    anio_actual  = periodo_str.split(" ")[-1] if " " in periodo_str else "2026"

    # Gráfico: índices en formato "Mes YYYY" cada 2 meses
    grafico_labels, grafico_fondo, grafico_icp, grafico_comp = [], [], [], []
    fechas_comunes = sorted(set(b1000_fondo.index) | set(b1000_icp.index))
    for i, f in enumerate(fechas_comunes):
        label = pd.Timestamp(f).strftime("%b %Y")
        vf = float(b1000_fondo.get(f, np.nan)) if f in b1000_fondo.index else None
        vi = float(b1000_icp.get(f,   np.nan)) if f in b1000_icp.index   else None
        vc = float(b1000_comp.get(f,  np.nan)) if not b1000_comp.empty and f in b1000_comp.index else None
        if vf is not None and not np.isnan(vf):
            grafico_labels.append(label)
            grafico_fondo.append(round(vf, 2))
            grafico_icp.append(round(vi, 2) if vi and not np.isnan(vi) else None)
            grafico_comp.append(round(vc, 2) if vc and not np.isnan(vc) else None)

    # Tabla resumen rentabilidades
    tabla_rentab = []
    for _, row in metricas.iterrows():
        tabla_rentab.append({
            "nombre":      row["nombre"],
            "mensual":     _fmt_pct_pos(row.get("mensual")),
            "trimestral":  _fmt_pct_pos(row.get("trimestral")),
            "semestral":   _fmt_pct_pos(row.get("semestral")),
            "anual":       _fmt_pct_pos(row.get("anual")),
            "ytd":         _fmt_pct_pos(row.get("ytd")),
            "es_fondo":    "Fondo" in row["nombre"],
        })

    # Tabla histórica: agrupar por año
    tabla_historica = []
    if not df_rentm.empty:
        df_rentm = df_rentm.copy()
        df_rentm["fecha"] = pd.to_datetime(df_rentm["mes"], format="%b %Y", errors="coerce")
        df_rentm["anio"]  = df_rentm["fecha"].dt.year
        df_rentm["mes_n"] = df_rentm["fecha"].dt.month

        for anio, grp in df_rentm.groupby("anio"):
            row_icp  = [""] * 13
            row_comp = [""] * 13
            row_f    = [""] * 13
            for _, r in grp.iterrows():
                m = int(r["mes_n"]) - 1  # 0-indexed
                row_icp[m]  = _fmt_pct_pos(r.get("rent_icp"))
                row_comp[m] = _fmt_pct_pos(r.get("rent_comp"))
                row_f[m]    = _fmt_pct_pos(r.get("rent_fondo"))

            tabla_historica.append({
                "anio": anio,
                "series": [
                    {"nombre": "ICP",            "es_fondo": False, "valores": row_icp},
                    {"nombre": "Competencia",    "es_fondo": False, "valores": row_comp},
                    {"nombre": nombre_fondo.replace("FIP VANTRUST ","FIP "), "es_fondo": True, "valores": row_f},
                ]
            })

    # Composición por moneda y duración desde cartera
    comp_moneda   = []
    comp_duracion = []
    if not df_cartera.empty:
        por_moneda = df_cartera.groupby("moneda")["pct"].sum()
        for mon, pct in sorted(por_moneda.items(), key=lambda x: -x[1]):
            comp_moneda.append([mon, f"{pct*100:.2f}%".replace(".",",")])

        por_tramo = df_cartera.groupby("duracion")["pct"].sum()
        orden_tramos = ["Menos de 1 mes","1-3 meses","3-4 meses","Más de 4 meses"]
        for tramo in orden_tramos:
            pct = por_tramo.get(tramo, 0)
            comp_duracion.append([tramo, f"{pct*100:.2f}%".replace(".",",")])

    return {
        "nombre_fondo":      nombre_fondo,
        "nombre_corto":      nombre_corto,
        "administradora":    info_fondo.get("administradora", "Vantrust Gestion Patrimonial S.A."),
        "rut":               info_fondo.get("rut", "76,637,334-8"),
        "moneda":            info_fondo.get("moneda", "CLP"),
        "tipo":              info_fondo.get("tipo", "Fondo de Inversión Privado"),
        "fecha_inicio":      info_fondo.get("fecha_inicio", ""),
        "benchmark":         info_fondo.get("benchmark", "Índice Cámara Promedio (ICP)"),
        "plazo_rescate":     info_fondo.get("plazo_rescate", "A más tardar 15 días corridos"),
        "remuneracion":      info_fondo.get("remuneracion", "0,295% IVA Incluido"),
        "objetivo":          "Invertir los recursos del fondo en instrumentos de deuda de corto y mediano plazo, en una cartera diversificada, obteniendo una rentabilidad igual o superior al ICP.",
        "rentabilidad_texto":f"La rentabilidad esperada del {nombre_fondo}, es la tasa de política monetaria promedio del Banco Central de Chile.",
        "inversionistas":    "Dirigida a empresas y personas que buscan invertir sus excedentes de caja con una rentabilidad de corto plazo y baja tolerancia al riesgo.",
        "comentario":        comentario_pm,
        "anio_acum":         anio_actual,
        "grafico_labels":    grafico_labels,
        "grafico_fondo":     grafico_fondo,
        "grafico_icp":       [x for x in grafico_icp   if x is not None] if any(grafico_icp)   else [],
        "grafico_comp":      [x for x in grafico_comp  if x is not None] if any(grafico_comp)  else [],
        "tabla_rentab":      tabla_rentab,
        "tabla_historica":   tabla_historica,
        "comp_moneda":       comp_moneda,
        "comp_duracion":     comp_duracion,
    }


def generar_pptx(
    nombre_fondo:    str,
    periodo_str:     str,
    comentario_pm:   str,
    b1000_fondo:     pd.Series,
    b1000_comp:      pd.Series,
    b1000_icp:       pd.Series,
    metricas:        pd.DataFrame,
    df_rentm:        pd.DataFrame,
    df_cartera:      pd.DataFrame,
    info_fondo:      dict,
    out_path:        Path,
) -> Path:
    """
    Genera el PPTX de un fondo usando pptxgenjs.
    Retorna la ruta del archivo generado.
    """
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # Construir JSON de datos
    datos = construir_datos_json(
        nombre_fondo, periodo_str, comentario_pm,
        b1000_fondo, b1000_comp, b1000_icp,
        metricas, df_rentm, df_cartera, info_fondo
    )

    # Guardar JSON temporal
    json_path = out_path.parent / f"_tmp_{out_path.stem}.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(datos, f, ensure_ascii=False)

    try:
        # Llamar al generador JS
        result = subprocess.run(
            ["node", str(JS_SCRIPT), "--data", str(json_path), "--out", str(out_path)],
            capture_output=True, text=True, timeout=60,
            cwd=str(JS_DIR)
        )
        if result.returncode != 0:
            raise RuntimeError(f"generar_folleto.js falló:\n{result.stderr}\n{result.stdout}")
        if not out_path.exists():
            raise FileNotFoundError(f"PPTX no generado: {out_path}")
    finally:
        if json_path.exists():
            json_path.unlink()

    return out_path
