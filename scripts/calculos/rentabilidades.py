"""
calculos/rentabilidades.py — Rentabilidades con ajuste por dividendos.
Replica exactamente la fórmula del Excel template de cada fondo.
"""
import pandas as pd
import numpy as np

DIVIDENDOS = {
    "FIP VANTRUST LIQUIDEZ ACTIVA":         0.0,
    "FIP VANTRUST LIQUIDEZ ALTO APORTE":    0.0,
    "FIP VANTRUST LIQUIDEZ ALTO CAPITAL":  26.8912,
    "FIP VANTRUST LIQUIDEZ ALTO MONTO":     0.0,
    "FIP VANTRUST LIQUIDEZ CAJA":           6.6493,
    "FIP VANTRUST LIQUIDEZ CONTINUA":       0.0,
    "FIP VANTRUST LIQUIDEZ CORRIENTE":      0.0,
    "FIP VANTRUST LIQUIDEZ CORTO PLAZO":    0.0,
    "FIP VANTRUST LIQUIDEZ DISPONIBLE I":   0.0,
    "FIP VANTRUST LIQUIDEZ DOLAR":          0.0,
    "FIP VANTRUST LIQUIDEZ DOLAR CAJA":     0.0,
    "FIP VANTRUST LIQUIDEZ EFECTIVO":       0.0,
    "FIP VANTRUST LIQUIDEZ FLEXIBLE":       0.0,
    "FIP VANTRUST LIQUIDEZ I":             79.3866,
    "FIP VANTRUST LIQUIDEZ LOCAL":          0.0,
    "FIP VANTRUST LIQUIDEZ MONETARIO I":    0.0,
    "FIP VANTRUST LIQUIDEZ PERMANENTE":     2.0800,
    "FIP VANTRUST LIQUIDEZ PLUS":          29.4500,
    "FIP VANTRUST LIQUIDEZ PRESENTE":       2.0800,
    "FIP VANTRUST LIQUIDEZ RECURRENTE":    60.5472,
    "FIP VANTRUST LIQUIDEZ RENDIMIENTO":    0.0,
    "FIP VANTRUST LIQUIDEZ RESERVA DOLAR":  0.0,
    "FIP VANTRUST LIQUIDEZ SENCILLO":       6.6493,
    "FIP VANTRUST LIQUIDEZ TEMPORAL":       2.0735,
}


def _eom_serie(serie_diaria: pd.Series) -> pd.Series:
    """
    Serie mensual con la convención del template:
    - Todos los meses: último dato disponible del mes
    - Diciembre: primer dato disponible de enero siguiente
    El índice retornado es un Period (año-mes) para evitar duplicados.
    """
    if serie_diaria.empty:
        return pd.Series(dtype=float)

    s = serie_diaria.sort_index()
    result = {}

    df = pd.DataFrame({"vc": s.values}, index=s.index)
    df["period"] = df.index.to_period("M")

    for period, grupo in df.groupby("period"):
        anio, mes = period.year, period.month
        if mes == 12:
            siguiente = pd.Timestamp(f"{anio+1}-01-01")
            enero_data = s[s.index >= siguiente]
            if not enero_data.empty:
                result[period] = enero_data.iloc[0]
                continue
        result[period] = grupo["vc"].iloc[-1]

    if not result:
        return pd.Series(dtype=float)

    return pd.Series(result, dtype=float).sort_index()


def calcular_rent_mensual(serie_diaria: pd.Series, dividendos: float = 0.0) -> pd.Series:
    """Rentabilidades mensuales con ajuste de dividendos."""
    eom = _eom_serie(serie_diaria)
    if eom.empty or len(eom) < 2:
        return pd.Series(dtype=float)
    H = eom + dividendos
    return H.pct_change().dropna()


def normalizar_b1000(serie_diaria: pd.Series, fecha_inicio: pd.Timestamp) -> pd.Series:
    """Base 1000 desde fecha_inicio para gráfico de evolución."""
    if serie_diaria.empty:
        return pd.Series(dtype=float)
    s = serie_diaria.sort_index()
    mask = s.index >= fecha_inicio
    if not mask.any():
        return pd.Series(dtype=float)
    v0 = s[mask].iloc[0]
    return (s / v0) * 1000 if v0 != 0 else pd.Series(dtype=float)


def _fmt(v) -> str:
    if v is None or (isinstance(v, float) and (np.isnan(v) or np.isinf(v))):
        return ""
    return f"{v*100:.2f}%".replace(".", ",")


def construir_tabla_rentabilidades(
    serie_vc_fondo, serie_vc_comp, serie_icp,
    nombre_fondo, fecha_fin
) -> pd.DataFrame:
    div = DIVIDENDOS.get(nombre_fondo, 0.0)
    nombre_serie = nombre_fondo.replace("FIP VANTRUST ", "FIP ")

    def calc(serie, d, n):
        r = calcular_rent_mensual(serie, d)
        if r.empty:
            return None
        r = r[r.index <= fecha_fin.to_period("M")]
        if len(r) < n:
            return None
        acc = 1.0
        for v in r.iloc[-n:]:
            acc *= (1 + v)
        return acc - 1

    def ytd(serie, d):
        r = calcular_rent_mensual(serie, d)
        if r.empty:
            return None
        r = r[r.index <= fecha_fin.to_period("M")]
        r_anio = r[r.index.year == fecha_fin.year]
        if r_anio.empty:
            return None
        acc = 1.0
        for v in r_anio:
            acc *= (1 + v)
        return acc - 1

    rows = []
    for nombre, serie, d in [
        ("ICP (Benchmark)", serie_icp,      0.0),
        ("Competencia",     serie_vc_comp,  0.0),
        (nombre_serie,      serie_vc_fondo, div),
    ]:
        rows.append({
            "nombre":      nombre,
            "mensual":     _fmt(calc(serie, d, 1)),
            "trimestral":  _fmt(calc(serie, d, 3)),
            "semestral":   _fmt(calc(serie, d, 6)),
            "anual":       _fmt(calc(serie, d, 12)),
            "ytd":         _fmt(ytd(serie, d)),
            "es_fondo":    nombre == nombre_serie,
        })
    return pd.DataFrame(rows)


def construir_tabla_historica(
    serie_vc_fondo, serie_vc_comp, serie_icp,
    nombre_fondo, fecha_fin
) -> list:
    div = DIVIDENDOS.get(nombre_fondo, 0.0)
    nombre_serie = nombre_fondo.replace("FIP VANTRUST ", "FIP ")

    def to_dict(serie, d):
        r = calcular_rent_mensual(serie, d)
        if r.empty:
            return {}
        r = r[r.index <= fecha_fin.to_period("M")]
        return {(p.year, p.month): v for p, v in r.items()}

    rf = to_dict(serie_vc_fondo, div)
    rc = to_dict(serie_vc_comp,  0.0)
    ri = to_dict(serie_icp,      0.0)

    all_keys = set(rf) | set(rc) | set(ri)
    if not all_keys:
        return []

    anio_ini = min(k[0] for k in all_keys)
    anio_fin = fecha_fin.year

    def total(rents, anio):
        acc, found = 1.0, False
        for m in range(1, 13):
            v = rents.get((anio, m))
            if v is not None:
                acc *= (1 + v)
                found = True
        return (acc - 1) if found else None

    tabla = []
    for anio in range(anio_ini, anio_fin + 1):
        series_out = []
        for nombre, rents, es_f in [
            ("ICP", ri, False), ("Competencia", rc, False), (nombre_serie, rf, True)
        ]:
            vals = [_fmt(rents.get((anio, m))) for m in range(1, 13)]
            vals.append(_fmt(total(rents, anio)))
            series_out.append({"nombre": nombre, "es_fondo": es_f, "valores": vals})
        if any(v for s in series_out for v in s["valores"] if v):
            tabla.append({"anio": anio, "series": series_out})

    return tabla
