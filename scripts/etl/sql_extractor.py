"""
etl/sql_extractor.py — Obtiene valores cuota desde la API REST interna.

API: POST https://claudeods.vantrustcapital.cl/query
     Body: {"Sql": "SELECT ..."}

Tabla: PUBLICADOR_PRECIO (base GPIVANTRUST_CB)
  FECHA       → fecha del valor cuota
  NEMOTECNICO → nombre del fondo
  COD_MONEDA  → moneda
  PRECIO      → valor cuota
"""
import requests
import pandas as pd
from datetime import date, timedelta

API_URL = "https://claudeods.vantrustcapital.cl/query"
HEADERS = {"Content-Type": "application/json"}

FONDOS_QUERY = (
    "'FIP VANTRUST LIQUIDEZ ACTIVA',"
    "'FIP VANTRUST LIQUIDEZ ALTO APORTE',"
    "'FIP VANTRUST LIQUIDEZ ALTO CAPITAL',"
    "'FIP VANTRUST LIQUIDEZ ALTO MONTO',"
    "'FIP VANTRUST LIQUIDEZ CAJA',"
    "'FIP VANTRUST LIQUIDEZ CONTINUA',"
    "'FIP VANTRUST LIQUIDEZ CORRIENTE',"
    "'FIP VANTRUST LIQUIDEZ CORTO PLAZO',"
    "'FIP VANTRUST LIQUIDEZ DISPONIBLE I',"
    "'FIP VANTRUST LIQUIDEZ DOLAR',"
    "'FIP VANTRUST LIQUIDEZ DOLAR CAJA',"
    "'FIP VANTRUST LIQUIDEZ EFECTIVO',"
    "'FIP VANTRUST LIQUIDEZ FLEXIBLE',"
    "'FIP VANTRUST LIQUIDEZ I',"
    "'FIP VANTRUST LIQUIDEZ LOCAL',"
    "'FIP VANTRUST LIQUIDEZ MONETARIO I',"
    "'FIP VANTRUST LIQUIDEZ PERMANENTE',"
    "'FIP VANTRUST LIQUIDEZ PLUS',"
    "'FIP VANTRUST LIQUIDEZ PRESENTE',"
    "'FIP VANTRUST LIQUIDEZ RECURRENTE',"
    "'FIP VANTRUST LIQUIDEZ RENDIMIENTO',"
    "'FIP VANTRUST LIQUIDEZ RESERVA DOLAR',"
    "'FIP VANTRUST LIQUIDEZ SENCILLO',"
    "'FIP VANTRUST LIQUIDEZ TEMPORAL'"
)


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
            FECHA       AS fecha,
            NEMOTECNICO AS fondo,
            COD_MONEDA  AS moneda,
            PRECIO      AS valor_cuota
        FROM PUBLICADOR_PRECIO
        WHERE FECHA >= '{desde}'
          AND PRECIO > 0
          AND NEMOTECNICO IN ({FONDOS_QUERY})
        ORDER BY FECHA ASC
    """
    print("    Consultando PUBLICADOR_PRECIO...")
    rows = _query(sql)
    if not rows:
        raise ValueError("La API no retornó datos de valores cuota.")

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
