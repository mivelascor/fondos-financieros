"""
etl/icp_bcch.py — Descarga la TIB diaria desde la API del BCCh
y calcula el ICP según la fórmula oficial:

    ICP_i = ICP_{i-1} × (1 + TIB_{i-1}/100 × Ndias/360)

Requiere credenciales del BDE del BCCh (registro gratuito en si3.bcentral.cl).
Configurar como Secrets en GitHub:
  - BCCH_USER: tu email registrado en el BDE
  - BCCH_PASS: tu contraseña del BDE
"""
import os
import requests
import pandas as pd
import numpy as np
from datetime import date, timedelta


# Código de serie de la TIB diaria en el BDE
SERIE_TIB = "F022.TIB.INC.D001.NO.Z.D"
BCCH_API  = "https://si3.bcentral.cl/SieteRestWS/SieteRestWS.ashx"


def _get_tib_diaria(desde: str, hasta: str, user: str, pwd: str) -> pd.DataFrame:
    """
    Descarga la TIB diaria desde la API del BCCh.
    desde/hasta: formato YYYY-MM-DD
    Retorna DataFrame: fecha, tib
    """
    params = {
        "user":       user,
        "pass":       pwd,
        "function":   "GetSeries",
        "timeseries": SERIE_TIB,
        "firstdate":  desde,
        "lastdate":   hasta,
    }
    r = requests.get(BCCH_API, params=params, timeout=30)
    r.raise_for_status()
    data = r.json()

    if data.get("Codigo") != 0:
        raise ValueError(f"BCCh API error: {data.get('Descripcion')}")

    obs = data["Series"]["Obs"]
    rows = []
    for o in obs:
        val = o.get("value", "")
        if val and val not in ("", "NaN", "ND"):
            try:
                rows.append({
                    "fecha": pd.to_datetime(o["indexDateString"], dayfirst=True),
                    "tib":   float(val),
                })
            except Exception:
                continue

    df = pd.DataFrame(rows).sort_values("fecha").reset_index(drop=True)
    return df


def _calcular_icp(df_tib: pd.DataFrame, icp_inicial: float = 10000.0) -> pd.DataFrame:
    """
    Calcula el ICP diario a partir de la TIB diaria.
    Fórmula: ICP_i = ICP_{i-1} × (1 + TIB_{i-1}/100 × Ndias/360)

    - TIB_{i-1}: TIB del día hábil bancario anterior
    - Ndias: número de días corridos entre el día hábil anterior y el día i
    - icp_inicial: valor del ICP en la primera fecha disponible

    El resultado se aproxima al segundo decimal.
    """
    df = df_tib.copy().sort_values("fecha").reset_index(drop=True)

    icp_vals = [icp_inicial]
    for i in range(1, len(df)):
        fecha_actual   = df.loc[i,   "fecha"]
        fecha_anterior = df.loc[i-1, "fecha"]
        tib_anterior   = df.loc[i-1, "tib"]
        ndias          = (fecha_actual - fecha_anterior).days

        icp_prev = icp_vals[-1]
        icp_i    = icp_prev * (1 + tib_anterior / 100 * ndias / 360)
        icp_vals.append(round(icp_i, 2))

    df["icp"] = icp_vals
    return df[["fecha", "tib", "icp"]]


def get_icp_eom() -> pd.DataFrame:
    """
    Descarga la TIB, calcula el ICP y retorna el valor de fin de mes.
    Lee credenciales desde variables de entorno BCCH_USER y BCCH_PASS.
    Retorna: fecha, icp
    """
    user = os.environ.get("BCCH_USER", "")
    pwd  = os.environ.get("BCCH_PASS", "")

    if not user or not pwd:
        raise EnvironmentError(
            "Faltan credenciales BCCh. Configura los Secrets BCCH_USER y BCCH_PASS en GitHub."
        )

    # Descargar últimos 15 meses de TIB para tener historia suficiente
    hasta  = date.today().strftime("%Y-%m-%d")
    desde  = (date.today() - timedelta(days=450)).strftime("%Y-%m-%d")

    print(f"    Descargando TIB desde BCCh ({desde} → {hasta})...")
    df_tib = _get_tib_diaria(desde, hasta, user, pwd)
    print(f"    {len(df_tib)} observaciones de TIB descargadas")

    # Calcular ICP
    df_icp = _calcular_icp(df_tib)

    # Filtrar último día disponible de cada mes
    df_icp["anio_mes"] = df_icp["fecha"].dt.to_period("M")
    df_eom = (
        df_icp.sort_values("fecha")
              .groupby("anio_mes")
              .last()
              .reset_index()
              .drop(columns=["anio_mes"])
    )
    return df_eom[["fecha", "icp"]]
