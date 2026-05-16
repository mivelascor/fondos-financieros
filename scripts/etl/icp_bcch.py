"""
etl/icp_bcch.py — Descarga TPM histórica desde mindicador.cl (público, sin credenciales)
y calcula la rentabilidad mensual del ICP como TPM_promedio_mes / 1200.

Para precisión máxima se pueden usar credenciales BCCh vía si3.bcentral.cl,
pero la aproximación TPM/1200 es suficientemente precisa para los folletos.
"""
import os
import requests
import pandas as pd
from datetime import date

BCCH_API  = "https://si3.bcentral.cl/SieteRestWS/SieteRestWS.ashx"
SERIE_TIB = "F022.TIB.INC.D001.NO.Z.D"
MINDICADOR = "https://mindicador.cl/api/tpm"


def _icp_desde_bcch(user: str, pwd: str) -> pd.DataFrame | None:
    """Intenta descargar ICP diario desde BCCh si hay credenciales."""
    try:
        desde = "2009-01-01"
        hasta = date.today().strftime("%Y-%m-%d")
        params = {"user": user, "pass": pwd, "function": "GetSeries",
                  "timeseries": SERIE_TIB, "firstdate": desde, "lastdate": hasta}
        r = requests.get(BCCH_API, params=params, timeout=30)
        data = r.json()
        if data.get("Codigo") != 0:
            return None
        obs = data["Series"]["Obs"]
        rows = []
        for o in obs:
            val = o.get("value", "")
            if val and val not in ("", "NaN", "ND"):
                try:
                    rows.append({
                        "fecha": pd.to_datetime(o["indexDateString"], dayfirst=True),
                        "tib": float(val),
                    })
                except Exception:
                    continue
        df = pd.DataFrame(rows).sort_values("fecha").reset_index(drop=True)
        if df.empty:
            return None
        # Calcular ICP mensual
        df["anio_mes"] = df["fecha"].dt.to_period("M")
        monthly = df.groupby("anio_mes")["tib"].mean().reset_index()
        monthly["icp"] = monthly["tib"] / 1200
        monthly["fecha"] = monthly["anio_mes"].dt.to_timestamp("M")
        return monthly[["fecha", "icp"]].copy()
    except Exception as e:
        print(f"    [WARN] BCCh no disponible: {e}")
        return None


def _icp_desde_mindicador() -> pd.DataFrame:
    """Descarga TPM histórica desde mindicador.cl (sin credenciales)."""
    anio_inicio = 2009
    anio_fin    = date.today().year
    all_data    = []

    for anio in range(anio_inicio, anio_fin + 1):
        try:
            r = requests.get(f"{MINDICADOR}/{anio}", timeout=15)
            for item in r.json().get("serie", []):
                try:
                    all_data.append({
                        "fecha": pd.to_datetime(item["fecha"]),
                        "tpm": float(item["valor"]),
                    })
                except Exception:
                    continue
        except Exception:
            continue

    df = pd.DataFrame(all_data).sort_values("fecha").reset_index(drop=True)
    df["anio_mes"] = df["fecha"].dt.to_period("M")
    monthly = df.groupby("anio_mes")["tpm"].mean().reset_index()
    monthly["icp"] = monthly["tpm"] / 1200
    monthly["fecha"] = monthly["anio_mes"].dt.to_timestamp("M")
    return monthly[["fecha", "icp"]].copy()


def get_icp_eom() -> pd.DataFrame:
    """
    Retorna DataFrame con columnas: fecha (EOM), icp (rentabilidad mensual decimal).
    Datos desde 2009.
    """
    print("    Descargando TPM desde mindicador.cl...")

    user = os.environ.get("BCCH_USER", "")
    pwd  = os.environ.get("BCCH_PASS", "")

    df = None
    if user and pwd:
        df = _icp_desde_bcch(user, pwd)
        if df is not None:
            print(f"      {len(df)} meses de ICP (BCCh)")

    if df is None:
        df = _icp_desde_mindicador()
        print(f"      {len(df)} meses de ICP (mindicador.cl)")

    df["fecha"] = pd.to_datetime(df["fecha"])
    return df.sort_values("fecha").reset_index(drop=True)
