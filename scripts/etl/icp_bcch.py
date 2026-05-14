"""
etl/icp_bcch.py — Calcula el ICP automáticamente usando la TPM
de mindicador.cl (fuente pública, sin credenciales).

La TIB sigue exactamente la Tasa de Política Monetaria (TPM) del BCCh.
mindicador.cl publica la TPM diaria gratuitamente desde 2013.

Fórmula oficial ICP:
  ICP_i = ICP_{i-1} x (1 + TIB_{i-1}/100 x Ndias/360)
"""
import requests
import pandas as pd
from datetime import date, timedelta

ICP_INICIAL = 10000.0


def _get_tpm_anio(anio: int) -> pd.DataFrame:
    """Descarga la TPM diaria de un año desde mindicador.cl"""
    url = f"https://mindicador.cl/api/tpm/{anio}"
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    data = r.json()
    rows = [
        {"fecha": item["fecha"][:10], "tib": float(item["valor"])}
        for item in data.get("serie", [])
        if item.get("valor") is not None
    ]
    df = pd.DataFrame(rows)
    if not df.empty:
        df["fecha"] = pd.to_datetime(df["fecha"])
    return df


def get_icp_eom() -> pd.DataFrame:
    """
    Descarga TPM de los últimos 2 años, calcula ICP acumulado
    y retorna el valor de fin de mes.
    Retorna: fecha, icp
    """
    print("    Descargando TPM desde mindicador.cl (público, sin credenciales)...")

    hoy  = date.today()
    anio_actual   = hoy.year
    anio_anterior = hoy.year - 1
    anio_dos_atras = hoy.year - 2

    frames = []
    for anio in [anio_dos_atras, anio_anterior, anio_actual]:
        try:
            df = _get_tpm_anio(anio)
            if not df.empty:
                frames.append(df)
        except Exception as e:
            print(f"      [warn] No se pudo obtener TPM {anio}: {e}")

    if not frames:
        raise ValueError("No se pudo obtener la TPM de mindicador.cl")

    df_tpm = (
        pd.concat(frames, ignore_index=True)
          .sort_values("fecha")
          .drop_duplicates("fecha")
          .reset_index(drop=True)
    )

    print(f"      {len(df_tpm)} días de TPM ({df_tpm['fecha'].min().date()} → {df_tpm['fecha'].max().date()})")

    # Calcular ICP acumulado con la fórmula oficial
    # ICP_i = ICP_{i-1} x (1 + TIB_{i-1}/100 x Ndias/360)
    icp_vals = [ICP_INICIAL]
    for i in range(1, len(df_tpm)):
        tib_ant   = df_tpm.loc[i-1, "tib"]
        ndias     = (df_tpm.loc[i, "fecha"] - df_tpm.loc[i-1, "fecha"]).days
        icp_nuevo = icp_vals[-1] * (1 + tib_ant / 100 * ndias / 360)
        icp_vals.append(round(icp_nuevo, 2))

    df_tpm["icp"] = icp_vals

    # Filtrar último día disponible de cada mes
    df_tpm["anio_mes"] = df_tpm["fecha"].dt.to_period("M")
    df_eom = (
        df_tpm.sort_values("fecha")
              .groupby("anio_mes")
              .last()
              .reset_index()
              .drop(columns=["anio_mes"])
    )
    print(f"      {len(df_eom)} meses de ICP calculados")
    return df_eom[["fecha", "icp"]]
