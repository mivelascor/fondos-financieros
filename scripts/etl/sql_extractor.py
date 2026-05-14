import requests
import pandas as pd
from datetime import date, timedelta

API_URL = "https://claudeods.vantrustcapital.cl/query"
HEADERS = {"Content-Type": "application/json"}


def _query(sql: str) -> list:
    resp = requests.post(API_URL, json={"Sql": sql}, headers=HEADERS, timeout=60)
    resp.raise_for_status()
    data = resp.json()
    if isinstance(data, list):
        return data
    return data.get("rows", data.get("data", data.get("results", [])))


def get_valores_cuota_eom() -> pd.DataFrame:
    desde = (date.today() - timedelta(days=455)).strftime("%Y-%m-%d")

    sql = f"""
        SELECT
            FECHA_CIERRE                AS fecha,
            RTRIM(LTRIM(EMPRESA))       AS fondo,
            VALOR_CUOTA                 AS valor_cuota
        FROM ODS.VALORES_CUOTA_GPI
        WHERE FECHA_CIERRE >= '{desde}'
          AND VALOR_CUOTA > 0
        ORDER BY FECHA_CIERRE ASC
    """
    print("    Consultando ODS.VALORES_CUOTA_GPI...")
    rows = _query(sql)
    if not rows:
        raise ValueError("La API no retorno datos de valores cuota.")

    df = pd.DataFrame(rows)
    df["fecha"]       = pd.to_datetime(df["fecha"])
    df["fondo"]       = df["fondo"].str.strip()
    df["valor_cuota"] = pd.to_numeric(df["valor_cuota"], errors="coerce")
    df = df.dropna(subset=["fecha", "valor_cuota", "fondo"])

    df["anio_mes"] = df["fecha"].dt.to_period("M")
    df_eom = (
        df.sort_values("fecha")
          .groupby(["fondo", "anio_mes"])
          .last()
          .reset_index()
          .drop(columns=["anio_mes"])
    )
    print(f"      {df_eom['fondo'].nunique()} fondos, {len(df_eom)} registros EOM")
    return df_eom[["fecha", "fondo", "valor_cuota"]]
