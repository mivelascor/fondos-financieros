"""
calculos/rentabilidades.py
Fórmulas verificadas contra TEMPLATE FONDO LIQUIDEZ UNO.xlsx (abril 2026).

Nivel FIP = VC_EOM + dividendos_historicos
Nivel ICP = nivel acumulado desde BCCh/mindicador (ya viene como nivel)
Nivel Comp = VC_EOM de la competencia (ya viene como nivel)

Mensual    = nivel_t / nivel_{t-1} - 1
Trimestral = nivel_t / nivel_{t-3} - 1
Semestral  = nivel_t / nivel_{t-6} - 1
Anual      = nivel_t / nivel_{t-12} - 1
Acum YTD (ICP/Comp) = (nivel_t / nivel_ene_año - 1) / n_meses * 12
Acum YTD (FIP)      = (nivel_t / nivel_dic_prev  - 1) / n_meses * 12
Total año histórico  = product(rent_mensuales_del_año) - 1
"""
import pandas as pd
import numpy as np
from datetime import date

DIVIDENDOS = {
    "FIP VANTRUST LIQUIDEZ ACTIVA":         0.0,
    "FIP VANTRUST LIQUIDEZ ALTO APORTE":   66.6742,
    "FIP VANTRUST LIQUIDEZ ALTO CAPITAL": 154.4727,
    "FIP VANTRUST LIQUIDEZ ALTO MONTO":     0.0,
    "FIP VANTRUST LIQUIDEZ CAJA":         113.0267,
    "FIP VANTRUST LIQUIDEZ CONTINUA":       0.0,
    "FIP VANTRUST LIQUIDEZ CORRIENTE":      0.0,
    "FIP VANTRUST LIQUIDEZ CORTO PLAZO":   23.6148,
    "FIP VANTRUST LIQUIDEZ DISPONIBLE I":   0.0,
    "FIP VANTRUST LIQUIDEZ DOLAR":          0.0239,
    "FIP VANTRUST LIQUIDEZ DOLAR CAJA":     0.0046,
    "FIP VANTRUST LIQUIDEZ EFECTIVO":     292.1144,
    "FIP VANTRUST LIQUIDEZ FLEXIBLE":       0.0,
    "FIP VANTRUST LIQUIDEZ I":            153.6105,
    "FIP VANTRUST LIQUIDEZ LOCAL":         84.7916,
    "FIP VANTRUST LIQUIDEZ MONETARIO I":    0.0,
    "FIP VANTRUST LIQUIDEZ PERMANENTE":     0.0,
    "FIP VANTRUST LIQUIDEZ PLUS":          29.4450,
    "FIP VANTRUST LIQUIDEZ PRESENTE":       0.0,
    "FIP VANTRUST LIQUIDEZ RECURRENTE":   177.5114,
    "FIP VANTRUST LIQUIDEZ RENDIMIENTO":   70.4139,
    "FIP VANTRUST LIQUIDEZ RESERVA DOLAR":  0.0239,
    "FIP VANTRUST LIQUIDEZ SENCILLO":      71.9851,
    "FIP VANTRUST LIQUIDEZ TEMPORAL":       0.0173,
}


def _eom_niveles(serie_diaria: pd.Series, dividendos: float = 0.0) -> pd.Series:
    """
    Convierte serie diaria de valor cuota a niveles EOM.
    Nivel = VC_EOM + dividendos.
    Índice: DatetimeIndex con fechas EOM.
    Convención diciembre: usa el primer dato de enero siguiente (T+1).
    """
    if serie_diaria.empty:
        return pd.Series(dtype=float)
    s = serie_diaria.sort_index()
    result = {}
    for period, grupo in s.groupby(s.index.to_period("M")):
        if period.month == 12:
            sig = pd.Timestamp(f"{period.year+1}-01-01")
            ene = s[s.index >= sig]
            vc = ene.iloc[0] if not ene.empty else grupo.iloc[-1]
        else:
            vc = grupo.iloc[-1]
        fecha_eom = period.to_timestamp("M")
        result[fecha_eom] = vc + dividendos
    return pd.Series(result).sort_index()


def _resumen(niveles: pd.Series, fecha_fin: pd.Timestamp,
             es_fip: bool = False) -> dict:
    """
    Calcula los 5 indicadores del resumen para la fecha_fin dada.
    es_fip=True usa dic_anterior como base del acum YTD.
    es_fip=False usa enero del año en curso como base del acum YTD.
    """
    n = niveles.sort_index()
    # Filtrar hasta fecha_fin
    n = n[n.index <= fecha_fin]
    if n.empty or len(n) < 2:
        return {"m": None, "t": None, "s": None, "a": None, "ac": None}

    nivel_t = n.iloc[-1]

    def ratio(offset):
        idx = len(n) - 1 - offset
        if idx < 0:
            return None
        return nivel_t / n.iloc[idx] - 1

    mensual    = ratio(1)
    trimestral = ratio(3)
    semestral  = ratio(6)
    anual      = ratio(12)

    # Acum YTD
    anio = fecha_fin.year
    if es_fip:
        # Base = diciembre del año anterior
        dic_prev = pd.Timestamp(f"{anio-1}-12-31")
        base_candidates = n[n.index <= dic_prev]
        if base_candidates.empty:
            # No hay datos de dic anterior → usar el primer dato disponible del año
            base_candidates = n[n.index.year == anio]
            nivel_base = base_candidates.iloc[0] if not base_candidates.empty else None
        else:
            nivel_base = base_candidates.iloc[-1]
    else:
        # Base = enero del año en curso
        ene_año = n[(n.index.year == anio) & (n.index.month == 1)]
        nivel_base = ene_año.iloc[0] if not ene_año.empty else None

    n_meses = len(n[n.index.year == anio])

    if nivel_base and n_meses > 0 and nivel_base != 0:
        acum_ytd = (nivel_t / nivel_base - 1) / n_meses * 12
    else:
        acum_ytd = None

    return {"m": mensual, "t": trimestral, "s": semestral,
            "a": anual, "ac": acum_ytd}


def _fmt(v) -> str:
    if v is None or (isinstance(v, float) and (np.isnan(v) or np.isinf(v))):
        return "—"
    return f"{v*100:.2f}%".replace(".", ",")


def calcular_resumen(serie_vc_fondo:  pd.Series,
                     serie_nivel_icp:  pd.Series,
                     serie_vc_comp:    pd.Series,
                     nombre_fondo:     str,
                     fecha_fin:        pd.Timestamp) -> pd.DataFrame:
    """
    Tabla resumen: 3 filas (ICP, Competencia, FIP) × 5 columnas.
    Los valores son strings formateados como "X,XX%".
    """
    div          = DIVIDENDOS.get(nombre_fondo, 0.0)
    nombre_serie = nombre_fondo.replace("FIP VANTRUST ", "FIP ")

    nivel_fip  = _eom_niveles(serie_vc_fondo, div)
    nivel_icp  = serie_nivel_icp  # ya viene como nivel
    nivel_comp = serie_vc_comp    # ya viene como nivel (valor cuota EOM)

    r_icp  = _resumen(nivel_icp,  fecha_fin, es_fip=False)
    r_comp = _resumen(nivel_comp, fecha_fin, es_fip=False)
    r_fip  = _resumen(nivel_fip,  fecha_fin, es_fip=True)

    rows = []
    for nombre, r in [("ICP (Benchmark)", r_icp),
                      ("Competencia",     r_comp),
                      (nombre_serie,      r_fip)]:
        rows.append({
            "nombre":     nombre,
            "mensual":    _fmt(r["m"]),
            "trimestral": _fmt(r["t"]),
            "semestral":  _fmt(r["s"]),
            "anual":      _fmt(r["a"]),
            "ytd":        _fmt(r["ac"]),
            "es_fondo":   nombre == nombre_serie,
        })
    return pd.DataFrame(rows)


def calcular_historico(serie_vc_fondo:  pd.Series,
                       serie_nivel_icp:  pd.Series,
                       serie_vc_comp:    pd.Series,
                       nombre_fondo:     str,
                       fecha_fin:        pd.Timestamp) -> list:
    """
    Tabla histórica: list de dicts por año.
    Cada dict: {año, icp:[12], icpT, comp:[12], compT, fip:[12], fipT}
    Valores son floats (None si no hay dato).
    Total año = product(meses del año) - 1.
    """
    div          = DIVIDENDOS.get(nombre_fondo, 0.0)
    nombre_serie = nombre_fondo.replace("FIP VANTRUST ", "FIP ")

    nivel_fip  = _eom_niveles(serie_vc_fondo, div)
    nivel_icp  = serie_nivel_icp
    nivel_comp = serie_vc_comp

    # Rentabilidades mensuales de cada serie
    def rent_mensual(niveles: pd.Series) -> dict:
        n = niveles.sort_index()
        n = n[n.index <= fecha_fin]
        if n.empty or len(n) < 2:
            return {}
        r = n.pct_change().dropna()
        return {(ts.year, ts.month): float(v) for ts, v in r.items()}

    ri = rent_mensual(nivel_icp)
    rc = rent_mensual(nivel_comp)
    rf = rent_mensual(nivel_fip)

    all_years = sorted(set(k[0] for k in list(ri) + list(rc) + list(rf)))
    anio_fin  = fecha_fin.year

    resultado = []
    for anio in all_years:
        if anio > anio_fin:
            continue

        icp_meses  = [ri.get((anio, m)) for m in range(1, 13)]
        comp_meses = [rc.get((anio, m)) for m in range(1, 13)]
        fip_meses  = [rf.get((anio, m)) for m in range(1, 13)]

        def total(meses):
            acc, hay = 1.0, False
            for v in meses:
                if v is not None:
                    acc *= (1 + v)
                    hay = True
            return acc - 1 if hay else None

        hay_datos = any(v is not None for v in icp_meses + comp_meses + fip_meses)
        if not hay_datos:
            continue

        resultado.append({
            "año":   anio,
            "icp":   icp_meses,
            "icpT":  total(icp_meses),
            "comp":  comp_meses  if any(v is not None for v in comp_meses)  else None,
            "compT": total(comp_meses) if any(v is not None for v in comp_meses) else None,
            "fip":   fip_meses   if any(v is not None for v in fip_meses)   else None,
            "fipT":  total(fip_meses) if any(v is not None for v in fip_meses) else None,
        })

    return resultado


def normalizar_b1000(serie_diaria: pd.Series, fecha_inicio: pd.Timestamp,
                     dividendos: float = 0.0) -> pd.Series:
    """Base 1000 desde fecha_inicio para el gráfico."""
    if serie_diaria.empty:
        return pd.Series(dtype=float)
    nivel = _eom_niveles(serie_diaria, dividendos)
    nivel = nivel.sort_index()
    mask  = nivel.index >= fecha_inicio
    if not mask.any():
        return pd.Series(dtype=float)
    v0 = nivel[mask].iloc[0]
    return (nivel / v0) * 1000 if v0 else pd.Series(dtype=float)
