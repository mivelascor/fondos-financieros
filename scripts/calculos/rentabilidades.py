import pandas as pd
import numpy as np


def normalizar_b1000(serie: pd.Series, fecha_inicio: pd.Timestamp) -> pd.Series:
    if serie.empty:
        return pd.Series(dtype=float)
    serie = serie.sort_index()
    mask = serie.index >= fecha_inicio
    if not mask.any():
        return pd.Series(dtype=float)
    vc_base = serie[mask].iloc[0]
    if vc_base == 0:
        return pd.Series(dtype=float)
    return (serie / vc_base) * 1000


def _ultimo_valor(serie: pd.Series, fecha: pd.Timestamp):
    if serie.empty:
        return np.nan
    idx = pd.DatetimeIndex(serie.index)
    mask = idx <= fecha
    if not mask.any():
        return np.nan
    return serie.iloc[np.where(mask)[0][-1]]


def rent_periodo(serie_b1000: pd.Series, fecha_fin: pd.Timestamp, meses: int) -> float:
    if serie_b1000.empty:
        return np.nan
    fecha_ini = fecha_fin - pd.DateOffset(months=meses)
    v_ini = _ultimo_valor(serie_b1000, fecha_ini)
    v_fin = _ultimo_valor(serie_b1000, fecha_fin)
    if pd.isna(v_ini) or pd.isna(v_fin) or v_ini == 0:
        return np.nan
    return (v_fin / v_ini) - 1


def rent_ytd(serie_b1000: pd.Series, fecha_fin: pd.Timestamp) -> float:
    if serie_b1000.empty:
        return np.nan
    if fecha_fin.month == 1:
        fecha_base = pd.Timestamp(fecha_fin.year - 1, 12, 31)
    else:
        fecha_base = pd.Timestamp(fecha_fin.year, 1, 31)
    v_base = _ultimo_valor(serie_b1000, fecha_base)
    v_fin  = _ultimo_valor(serie_b1000, fecha_fin)
    if pd.isna(v_base) or pd.isna(v_fin) or v_base == 0:
        return np.nan
    return (v_fin / v_base) - 1


def construir_tabla_rentabilidades(
    b1000_fondo: pd.Series,
    b1000_comp:  pd.Series,
    b1000_icp:   pd.Series,
    fecha_fin:   pd.Timestamp,
) -> pd.DataFrame:
    series = [
        ("Fondo",       b1000_fondo),
        ("Competencia", b1000_comp),
        ("ICP",         b1000_icp),
    ]
    rows = []
    for nombre, s in series:
        rows.append({
            "nombre":     nombre,
            "mensual":    rent_periodo(s, fecha_fin, 1),
            "trimestral": rent_periodo(s, fecha_fin, 3),
            "semestral":  rent_periodo(s, fecha_fin, 6),
            "anual":      rent_periodo(s, fecha_fin, 12),
            "ytd":        rent_ytd(s, fecha_fin),
        })
    return pd.DataFrame(rows)


def construir_rentabilidades_mensuales(
    b1000_fondo: pd.Series,
    b1000_comp:  pd.Series,
    b1000_icp:   pd.Series,
    fecha_fin:   pd.Timestamp,
    n_meses:     int = 12,
) -> pd.DataFrame:
    rows = []
    for i in range(n_meses, 0, -1):
        f2 = fecha_fin - pd.DateOffset(months=i - 1)
        rows.append({
            "mes":        f2.strftime("%b %Y"),
            "rent_fondo": rent_periodo(b1000_fondo, f2, 1),
            "rent_comp":  rent_periodo(b1000_comp,  f2, 1),
            "rent_icp":   rent_periodo(b1000_icp,   f2, 1),
        })
    return pd.DataFrame(rows)
