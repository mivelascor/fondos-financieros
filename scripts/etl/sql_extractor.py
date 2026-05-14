"""
etl/sql_extractor.py — Obtiene valores cuota desde la API REST interna.

API:
  GET  https://claudeods.vantrustcapital.cl/schema  → schema
  POST https://claudeods.vantrustcapital.cl/query   → {"Sql": "SELECT ..."}

Tabla usada: VALORES_CUOTA_GPI
  - EMPRESA      → nombre del fondo
  - FECHA_CIERRE → fecha del valor cuota
  - VALOR_CUOTA  → valor cuota del día

No requiere credenciales. Solo acepta SELECT.
"""
import requests
import pandas as pd
from datetime import date, timedelta

API_URL = "https://claudeods.vantrustcapital.cl/query"
HEADERS = {"Content-Type": "application/json"}


def _query(sql: str) -> list[dict]:
    """Ejecuta un SELECT en la API y retorna lista de filas."""
    resp = requests.post(API_URL, json={"Sql": sql}, headers=HEADERS, timeout=60)
    resp.raise_for_status()
    data = resp.json()
    # La API retorna lista de dicts o dict con key 'data'/'results'
    if isinstance(data, list):
        return data
    for key in ("data", "results", "rows", "value"):
        if key in data:
            return data[key]
    return data


def get_valores_cuota_eom() -> pd.DataFrame:
    """
    Descarga valores cuota diarios de los últimos 15 meses
    y filtra el último día disponible de cada mes por fondo.

    Retorna: fecha, fondo, valor_cuota
    """
    # Fecha inicio: 15 meses atrás
    desde = (date.today() - timedelta(days=455)).strftime("%Y-%m-%d")

    sql = f"""
        SELECT
            FECHA_CIERRE   AS fecha,
            EMPRESA        AS fondo,
            VALOR_CUOTA    AS valor_cuota
        FROM VALORES_CUOTA_GPI
        WHERE FECHA_CIERRE >= '{desde}'
          AND VALOR_CUOTA  > 0
        ORDER BY FECHA_CIERRE ASC
    """
    print("    Consultando VALORES_CUOTA_GPI...")
    rows = _query(sql)
    if not rows:
        raise ValueError("La API no retornó datos de valores cuota.")

    df = pd.DataFrame(rows)
    df["fecha"]       = pd.to_datetime(df["fecha"])
    df["fondo"]       = df["fondo"].str.strip()
    df["valor_cuota"] = pd.to_numeric(df["valor_cuota"], errors="coerce")
    df = df.dropna(subset=["fecha", "valor_cuota", "fondo"])

    # Filtrar último día disponible de cada mes
    df["anio_mes"] = df["fecha"].dt.to_period("M")
    df_eom = (
        df.sort_values("fecha")
          .groupby(["fondo", "anio_mes"])
          .last()
          .reset_index()
          .drop(columns=["anio_mes"])
    )
    print(f"    {df_eom['fondo'].nunique()} fondos, {len(df_eom)} registros EOM")
    return df_eom[["fecha", "fondo", "valor_cuota"]]
