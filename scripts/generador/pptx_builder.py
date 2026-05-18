"""generador/pptx_builder.py — Pasa datos al generar_folleto.js."""
import subprocess, json, numpy as np, pandas as pd
from pathlib import Path

JS_DIR    = Path(__file__).parent
JS_SCRIPT = JS_DIR / "generar_folleto.js"


def _s(v):
    if v is None: return None
    if isinstance(v, float) and (np.isnan(v) or np.isinf(v)): return None
    return float(v)


def generar_pptx(nombre_fondo, periodo_str, comentario_pm,
                 tabla_resumen, tabla_historica,
                 b1000_fondo, b1000_icp, b1000_comp,
                 comp_cartera, info_fondo, out_path):

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    nombre_corto = nombre_fondo.replace("FIP VANTRUST LIQUIDEZ ","").replace("FIP VANTRUST ","").title()
    anio_actual  = periodo_str.split(" ")[-1] if " " in periodo_str else "2026"

    # Resumen: extraer los valores string de la tabla
    resumen = {"icp":{}, "comp":{}, "fip":{}}
    if tabla_resumen is not None and not tabla_resumen.empty:
        for _, row in tabla_resumen.iterrows():
            nombre = row.get("nombre","")
            vals   = {"m": row.get("mensual","—"), "t": row.get("trimestral","—"),
                      "s": row.get("semestral","—"), "a": row.get("anual","—"),
                      "ac": row.get("ytd","—")}
            if "ICP" in nombre:
                resumen["icp"]  = vals
            elif "Comp" in nombre:
                resumen["comp"] = vals
            else:
                resumen["fip"]  = vals

    # Nombre del FIP (tomado de la tabla resumen)
    nombre_fip = "FIP"
    if tabla_resumen is not None and not tabla_resumen.empty:
        fip_row = tabla_resumen[tabla_resumen.get("es_fondo", pd.Series(False, index=tabla_resumen.index))]
        if not fip_row.empty:
            nombre_fip = str(fip_row.iloc[0]["nombre"])

    # Histórico
    historico = []
    for h in (tabla_historica or []):
        historico.append({
            "año":   h["año"],
            "icp":   [_s(v) for v in (h.get("icp")  or [None]*12)],
            "icpT":  _s(h.get("icpT")),
            "comp":  [_s(v) for v in (h.get("comp") or [])] if h.get("comp") is not None else None,
            "compT": _s(h.get("compT")),
            "fip":   [_s(v) for v in (h.get("fip")  or [])] if h.get("fip")  is not None else None,
            "fipT":  _s(h.get("fipT")),
        })

    # Gráfico: datos b1000 reducidos a EOM
    def serie_to_list(serie):
        if serie is None or serie.empty: return []
        s = serie.sort_index()
        s = s.resample("ME").last().dropna()
        return [{"fecha": str(i.date()), "val": round(float(v),2)} for i,v in s.items()]

    datos = {
        "nombre_fondo":   nombre_fondo,
        "nombre_corto":   nombre_corto,
        "nombre_fip":     nombre_fip,
        "periodo":        periodo_str,
        "anio_acum":      anio_actual,
        "comentario":     comentario_pm,
        # Info general
        "administradora": info_fondo.get("administradora","Vantrust Gestion Patrimonial S.A."),
        "rut":            info_fondo.get("rut",""),
        "moneda":         info_fondo.get("moneda","CLP"),
        "tipo":           info_fondo.get("tipo","Fondo de Inversión Privado"),
        "fecha_inicio":   info_fondo.get("fecha_inicio",""),
        "benchmark":      info_fondo.get("benchmark","Índice Cámara Promedio (ICP)"),
        "plazo_rescate":  info_fondo.get("plazo_rescate","A más tardar 15 días corridos"),
        "remuneracion":   info_fondo.get("remuneracion","0,295% IVA Incluido"),
        "objetivo":       info_fondo.get("objetivo",""),
        "rentabilidad_texto": info_fondo.get("rentabilidad_texto",""),
        "inversionistas": info_fondo.get("inversionistas",""),
        # Rentabilidades
        "resumen":     resumen,
        "historico":   historico,
        # Gráfico
        "b1000_fondo": serie_to_list(b1000_fondo),
        "b1000_icp":   serie_to_list(b1000_icp),
        "b1000_comp":  serie_to_list(b1000_comp),
        # Composición
        "comp_moneda":       comp_cartera.get("moneda",[]),
        "comp_duracion":     comp_cartera.get("duracion",[]),
        "comp_instrumentos": comp_cartera.get("instrumentos",[]),
    }

    json_path = out_path.parent / f"_tmp_{out_path.stem}.json"
    with open(json_path,"w",encoding="utf-8") as f:
        json.dump(datos, f, ensure_ascii=False)

    try:
        r = subprocess.run(
            ["node", str(JS_SCRIPT), "--data", str(json_path), "--out", str(out_path)],
            capture_output=True, text=True, timeout=60, cwd=str(JS_DIR))
        if r.returncode != 0:
            raise RuntimeError(f"JS error:\n{r.stderr}\n{r.stdout}")
        if not out_path.exists():
            raise FileNotFoundError(f"PPTX no creado: {out_path}")
    finally:
        if json_path.exists(): json_path.unlink()

    return out_path
