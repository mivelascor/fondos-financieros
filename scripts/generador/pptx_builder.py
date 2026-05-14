"""generador/pptx_builder.py — Genera PPTX con datos calculados automáticamente."""
import subprocess, json
import numpy as np
import pandas as pd
from pathlib import Path
from etl.excel_reader import get_cartera_composicion
from calculos.rentabilidades import (
    normalizar_b1000, construir_tabla_rentabilidades, construir_tabla_historica
)

JS_DIR    = Path(__file__).parent
JS_SCRIPT = JS_DIR / "generar_folleto.js"


def generar_pptx(
    nombre_fondo, periodo_str, comentario_pm,
    serie_vc_fondo, serie_vc_comp, serie_icp,
    info_fondo, out_path
):
    out_path.parent.mkdir(parents=True, exist_ok=True)

    nombre_corto = nombre_fondo.replace("FIP VANTRUST ", "").title()
    nombre_serie = nombre_fondo.replace("FIP VANTRUST ", "FIP ")
    anio_actual  = int(periodo_str.split(" ")[-1]) if " " in periodo_str else 2026
    fd_ts        = pd.Timestamp(f"{anio_actual}-{int(periodo_str.split()[0][:2]) if False else '04'}-30")

    # Usar fecha_fin real desde periodo_str
    from datetime import date, timedelta
    hoy = date.today()
    fd = hoy.replace(day=1) - timedelta(days=1)
    fd_ts = pd.Timestamp(fd)

    # Gráfico (base 1000)
    fecha_inicio_fondo = pd.Timestamp(serie_vc_fondo.index[0]) if not serie_vc_fondo.empty else pd.Timestamp("2020-01-01")
    b1000_f = normalizar_b1000(serie_vc_fondo, fecha_inicio_fondo)
    b1000_c = normalizar_b1000(serie_vc_comp,  fecha_inicio_fondo) if not serie_vc_comp.empty else pd.Series(dtype=float)
    b1000_i = normalizar_b1000(serie_icp,      fecha_inicio_fondo) if not serie_icp.empty     else pd.Series(dtype=float)

    grafico_labels, grafico_fondo, grafico_icp, grafico_comp = [], [], [], []
    for f in sorted(b1000_f.index):
        vf = float(b1000_f.get(f, np.nan))
        if not np.isnan(vf):
            grafico_labels.append(pd.Timestamp(f).strftime("%b %Y"))
            grafico_fondo.append(round(vf, 2))
            vi = float(b1000_i.get(f, np.nan)) if not b1000_i.empty and f in b1000_i.index else None
            vc = float(b1000_c.get(f, np.nan)) if not b1000_c.empty and f in b1000_c.index else None
            grafico_icp.append(round(vi, 2) if vi and not np.isnan(vi) else None)
            grafico_comp.append(round(vc, 2) if vc and not np.isnan(vc) else None)

    # Tabla resumen
    df_rent = construir_tabla_rentabilidades(
        serie_vc_fondo, serie_vc_comp, serie_icp, nombre_fondo, fd_ts
    )

    # Tabla histórica
    tabla_hist = construir_tabla_historica(
        serie_vc_fondo, serie_vc_comp, serie_icp, nombre_fondo, fd_ts
    )

    # Composición desde cartera.xlsx
    comp = get_cartera_composicion(nombre_fondo)

    datos = {
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
        "rentabilidad_texto": f"La rentabilidad esperada del {nombre_fondo}, es la tasa de política monetaria promedio del Banco Central de Chile.",
        "inversionistas":    "Dirigida a empresas y personas que buscan invertir sus excedentes de caja con una rentabilidad de corto plazo y baja tolerancia al riesgo.",
        "comentario":        comentario_pm,
        "anio_acum":         str(anio_actual),
        "grafico_labels":    grafico_labels,
        "grafico_fondo":     grafico_fondo,
        "grafico_icp":       [x for x in grafico_icp  if x is not None],
        "grafico_comp":      [x for x in grafico_comp if x is not None],
        "tabla_rentab":      df_rent.to_dict("records"),
        "tabla_historica":   tabla_hist,
        "comp_moneda":       comp["moneda"],
        "comp_duracion":     comp["duracion"],
        "comp_instrumento":  comp["instrumento"],
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
            raise RuntimeError(f"JS fallo:\n{r.stderr}\n{r.stdout}")
        if not out_path.exists():
            raise FileNotFoundError(f"PPTX no generado: {out_path}")
    finally:
        if json_path.exists():
            json_path.unlink()

    return out_path
