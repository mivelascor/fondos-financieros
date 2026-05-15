"""
calculos/rentabilidades.py

Fórmulas exactas replicando el Excel template:

  H_t = VC_t + dividendos_historicos
  Rent_mes = H_t / H_{t-1} - 1

  Mensual    = último mes disponible
  Trimestral = product(últimos 3 meses) - 1
  Semestral  = product(últimos 6 meses) - 1
  Anual      = product(últimos 12 meses) - 1
  Acum YTD   = (H_actual / H_ene_año - 1) / n_meses * 12  [anualizado]
  Total Año  = product(todos los meses del año) - 1

Convención EOM: para diciembre se usa el primer día hábil de enero siguiente
(replica CONTAB_PATRIMONIO que publica con T+1 el 31 dic).
"""
import pandas as pd
import numpy as np

# ── Dividendos históricos ─────────────────────────────────────────────────────
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
    EOM con convención del template:
    - Todos los meses: último dato disponible del mes
    - Diciembre: primer dato disponible de enero siguiente
    Índice: Period mensual (para evitar duplicados).
    """
    if serie_diaria.empty:
        return pd.Series(dtype=float)

    s  = serie_diaria.sort_index()
    df = pd.DataFrame({"vc": s.values}, index=s.index)
    df["period"] = df.index.to_period("M")
    result = {}

    for period, grupo in df.groupby("period"):
        anio, mes = period.year, period.month
        if mes == 12:
            siguiente = pd.Timestamp(f"{anio+1}-01-01")
            enero     = s[s.index >= siguiente]
            result[period] = enero.iloc[0] if not enero.empty else grupo["vc"].iloc[-1]
        else:
            result[period] = grupo["vc"].iloc[-1]

    return pd.Series(result, dtype=float).sort_index()


def calcular_rent_mensual(serie_diaria: pd.Series,
                          dividendos: float = 0.0) -> pd.Series:
    """Rentabilidades mensuales con ajuste de dividendos. Índice: Period."""
    eom = _eom_serie(serie_diaria)
    if eom.empty or len(eom) < 2:
        return pd.Series(dtype=float)
    H = eom + dividendos
    return H.pct_change().dropna()


def normalizar_b1000(serie_diaria: pd.Series,
                     fecha_inicio: pd.Timestamp) -> pd.Series:
    """Base 1000 desde fecha_inicio (para gráfico de evolución)."""
    if serie_diaria.empty:
        return pd.Series(dtype=float)
    s    = serie_diaria.sort_index()
    mask = s.index >= fecha_inicio
    if not mask.any():
        return pd.Series(dtype=float)
    v0   = s[mask].iloc[0]
    return (s / v0) * 1000 if v0 != 0 else pd.Series(dtype=float)


def _fmt(v) -> str:
    if v is None or (isinstance(v, float) and (np.isnan(v) or np.isinf(v))):
        return ""
    return f"{v*100:.2f}%".replace(".", ",")


def construir_tabla_rentabilidades(
    serie_vc_fondo: pd.Series,
    serie_vc_comp:  pd.Series,
    serie_icp:      pd.Series,
    nombre_fondo:   str,
    fecha_fin:      pd.Timestamp,
) -> pd.DataFrame:
    """
    Tabla resumen de rentabilidades.
    Anual    = product(últimos 12 meses) - 1
    Acum YTD = (H_actual/H_ene - 1) / n_meses * 12  [anualizado]
    """
    div          = DIVIDENDOS.get(nombre_fondo, 0.0)
    nombre_serie = nombre_fondo.replace("FIP VANTRUST ", "FIP ")
    periodo_fin  = fecha_fin.to_period("M")

    def calc_product(rent: pd.Series, n: int) -> float | None:
        r = rent[rent.index <= periodo_fin]
        if len(r) < n:
            return None
        acc = 1.0
        for v in r.iloc[-n:]:
            acc *= (1 + v)
        return acc - 1

    def calc_anual_ytd(rent: pd.Series) -> float | None:
        """
        Acum YTD anualizado = (H_ultimo / H_ene - 1) / n_meses * 12
        H_ene = valor del índice en enero del año en curso
        """
        r = rent[rent.index <= periodo_fin]
        anio = fecha_fin.year
        r_anio = r[r.index.year == anio]
        if r_anio.empty:
            return None
        n = len(r_anio)
        # Acum = product de los meses del año - 1 (no anualizado para el product)
        acc = 1.0
        for v in r_anio:
            acc *= (1 + v)
        acum = acc - 1
        # Anualizar: acum / n * 12
        return acum / n * 12

    rows = []
    for nombre, serie, d in [
        ("ICP (Benchmark)", serie_icp,      0.0),
        ("Competencia",     serie_vc_comp,  0.0),
        (nombre_serie,      serie_vc_fondo, div),
    ]:
        rent = calcular_rent_mensual(serie, d)
        rows.append({
            "nombre":      nombre,
            "mensual":     _fmt(calc_product(rent, 1)),
            "trimestral":  _fmt(calc_product(rent, 3)),
            "semestral":   _fmt(calc_product(rent, 6)),
            "anual":       _fmt(calc_product(rent, 12)),
            "ytd":         _fmt(calc_anual_ytd(rent)),
            "es_fondo":    nombre == nombre_serie,
        })
    return pd.DataFrame(rows)


def construir_tabla_historica(
    serie_vc_fondo: pd.Series,
    serie_vc_comp:  pd.Series,
    serie_icp:      pd.Series,
    nombre_fondo:   str,
    fecha_fin:      pd.Timestamp,
) -> list:
    """
    Tabla histórica mensual por año.
    Total Año = product(meses disponibles del año) - 1
    """
    div          = DIVIDENDOS.get(nombre_fondo, 0.0)
    nombre_serie = nombre_fondo.replace("FIP VANTRUST ", "FIP ")
    periodo_fin  = fecha_fin.to_period("M")

    def to_dict(serie, d):
        r = calcular_rent_mensual(serie, d)
        if r.empty:
            return {}
        r = r[r.index <= periodo_fin]
        return {(p.year, p.month): v for p, v in r.items()}

    rf = to_dict(serie_vc_fondo, div)
    rc = to_dict(serie_vc_comp,  0.0)
    ri = to_dict(serie_icp,      0.0)

    all_keys = set(rf) | set(rc) | set(ri)
    if not all_keys:
        return []

    anio_ini = min(k[0] for k in all_keys)
    anio_fin = fecha_fin.year

    def total_anio(rents, anio):
        acc, found = 1.0, False
        for m in range(1, 13):
            v = rents.get((anio, m))
            if v is not None:
                acc  *= (1 + v)
                found = True
        return (acc - 1) if found else None

    tabla = []
    for anio in range(anio_ini, anio_fin + 1):
        series_out = []
        for nombre, rents, es_f in [
            ("ICP",         ri, False),
            ("Competencia", rc, False),
            (nombre_serie,  rf, True),
        ]:
            vals = [rents.get((anio, m)) for m in range(1, 13)]
            total = total_anio(rents, anio)
            series_out.append({
                "nombre":   nombre,
                "es_fondo": es_f,
                "valores":  vals,         # floats o None (NO strings)
                "total":    total,        # float o None
            })

        if any(v for s in series_out for v in s["valores"] if v is not None):
            tabla.append({"anio": anio, "series": series_out})

    return tabla
