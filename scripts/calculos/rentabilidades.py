"""
calculos/rentabilidades.py
Verificado contra templates Alto Aporte y Liquidez Uno (abril 2026).

FÓRMULAS (todas verificadas):
  nivel_FIP = VC_EOM + dividendos_historicos_acumulados
  Mensual    = nivel_t   / nivel_{t-1}  - 1
  Trimestral = nivel_t   / nivel_{t-3}  - 1
  Semestral  = nivel_t   / nivel_{t-6}  - 1
  Anual      = nivel_t   / nivel_{t-12} - 1
  Acum YTD   = (nivel_t  / nivel_ene_año - 1) / n_meses * 12
               ← base = ENERO del año en curso (igual para ICP, Comp y FIP)
  Total año  = product(rent_mensuales_del_año) - 1
"""
import pandas as pd
import numpy as np

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
    Nivel_t = VC_EOM_t + dividendos.
    Convención diciembre: usa el primer dato de enero siguiente (T+1).
    Retorna Serie con DatetimeIndex (fechas EOM).
    """
    if serie_diaria.empty:
        return pd.Series(dtype=float)
    s = serie_diaria.sort_index()
    result = {}
    for period, grupo in s.groupby(s.index.to_period("M")):
        if period.month == 12:
            sig = pd.Timestamp(f"{period.year + 1}-01-01")
            ene = s[s.index >= sig]
            vc = ene.iloc[0] if not ene.empty else grupo.iloc[-1]
        else:
            vc = grupo.iloc[-1]
        result[period.to_timestamp("M")] = vc + dividendos
    return pd.Series(result).sort_index()


def _calcular_5_indicadores(niveles: pd.Series,
                             fecha_fin: pd.Timestamp) -> dict:
    """
    Calcula los 5 indicadores para fecha_fin dada.
    Acum YTD base = ENERO del año en curso (igual para ICP, Comp y FIP).
    """
    n = niveles.sort_index()
    n = n[n.index <= fecha_fin]
    if n.empty or len(n) < 2:
        return {"m": None, "t": None, "s": None, "a": None, "ac": None}

    nivel_t = n.iloc[-1]

    def ratio(offset):
        idx = len(n) - 1 - offset
        return nivel_t / n.iloc[idx] - 1 if idx >= 0 else None

    mensual    = ratio(1)
    trimestral = ratio(3)
    semestral  = ratio(6)
    anual      = ratio(12)

    # Acum YTD = (nivel_t / nivel_ene_año - 1) / n_meses * 12
    anio     = fecha_fin.year
    ene_año  = n[(n.index.year == anio) & (n.index.month == 1)]
    n_meses  = len(n[n.index.year == anio])

    if not ene_año.empty and n_meses > 0:
        nivel_base = ene_año.iloc[0]
        acum_ytd   = (nivel_t / nivel_base - 1) / n_meses * 12
    else:
        acum_ytd = None

    return {"m": mensual, "t": trimestral, "s": semestral,
            "a": anual, "ac": acum_ytd}


def _fmt(v) -> str:
    if v is None or (isinstance(v, float) and (np.isnan(v) or np.isinf(v))):
        return "—"
    return f"{v * 100:.2f}%".replace(".", ",")


def calcular_resumen(serie_vc_fondo:  pd.Series,
                     serie_nivel_icp:  pd.Series,
                     serie_vc_comp:    pd.Series,
                     nombre_fondo:     str,
                     fecha_fin:        pd.Timestamp) -> pd.DataFrame:
    """
    Tabla resumen: 3 filas (ICP, Competencia, FIP) × 5 columnas.
    Valores formateados como strings "X,XX%".
    """
    div          = DIVIDENDOS.get(nombre_fondo, 0.0)
    nombre_serie = nombre_fondo.replace("FIP VANTRUST ", "FIP ")

    nivel_fip  = _eom_niveles(serie_vc_fondo, div)
    nivel_icp  = serie_nivel_icp   # ya es nivel acumulado EOM
    nivel_comp = serie_vc_comp     # ya es valor cuota EOM

    r_icp  = _calcular_5_indicadores(nivel_icp,  fecha_fin)
    r_comp = _calcular_5_indicadores(nivel_comp, fecha_fin)
    r_fip  = _calcular_5_indicadores(nivel_fip,  fecha_fin)

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
    Tabla histórica por año, desde el año de inicio del FIP.
    Retorna lista de dicts: {año, icp:[12 floats/None], icpT, comp, compT, fip, fipT}
    Total año = product(meses_año) - 1.
    """
    div          = DIVIDENDOS.get(nombre_fondo, 0.0)
    nombre_serie = nombre_fondo.replace("FIP VANTRUST ", "FIP ")

    nivel_fip  = _eom_niveles(serie_vc_fondo, div)
    nivel_icp  = serie_nivel_icp
    nivel_comp = serie_vc_comp

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

    # Solo mostrar desde el año de inicio del FIP
    if not rf:
        return []
    anio_inicio_fip = min(k[0] for k in rf)
    anio_fin        = fecha_fin.year

    resultado = []
    for anio in range(anio_inicio_fip, anio_fin + 1):
        icp_m  = [ri.get((anio, m)) for m in range(1, 13)]
        comp_m = [rc.get((anio, m)) for m in range(1, 13)]
        fip_m  = [rf.get((anio, m)) for m in range(1, 13)]

        def total(meses):
            acc, hay = 1.0, False
            for v in meses:
                if v is not None:
                    acc *= (1 + v)
                    hay  = True
            return acc - 1 if hay else None

        hay = any(v is not None for v in icp_m + comp_m + fip_m)
        if not hay:
            continue

        resultado.append({
            "año":   anio,
            "icp":   icp_m,
            "icpT":  total(icp_m),
            "comp":  comp_m  if any(v is not None for v in comp_m)  else None,
            "compT": total(comp_m) if any(v is not None for v in comp_m) else None,
            "fip":   fip_m   if any(v is not None for v in fip_m)   else None,
            "fipT":  total(fip_m)  if any(v is not None for v in fip_m)  else None,
        })

    return resultado


def normalizar_b1000(serie_diaria: pd.Series,
                     fecha_inicio:  pd.Timestamp,
                     dividendos:    float = 0.0) -> pd.Series:
    """Base 1000 desde fecha_inicio para el gráfico."""
    if serie_diaria.empty:
        return pd.Series(dtype=float)
    nivel = _eom_niveles(serie_diaria, dividendos).sort_index()
    mask  = nivel.index >= fecha_inicio
    if not mask.any():
        return pd.Series(dtype=float)
    v0 = nivel[mask].iloc[0]
    return (nivel / v0) * 1000 if v0 else pd.Series(dtype=float)
