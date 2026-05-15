"""
etl/sql_extractor.py — Valor cuota desde API REST interna ODS.VALORES_CUOTA_GPI.
Retorna DataFrame con columnas en minúsculas: fecha, fondo, valor_cuota
"""
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
    # Intentar distintas claves que puede usar la API
    for key in ("rows", "data", "results", "Records", "value"):
        if key in data:
            return data[key]
    return []


def get_valores_cuota_eom() -> pd.DataFrame:
    """
    Obtiene valores cuota diarios desde ODS.VALORES_CUOTA_GPI.
    Retorna DataFrame con columnas: fecha, fondo, valor_cuota
    """
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

    # Normalizar nombres de columnas a minúsculas
    df.columns = [c.lower() for c in df.columns]

    # Asegurar que existen las columnas esperadas
    # La API puede retornar las columnas con el alias o con el nombre original
    col_map = {}
    for col in df.columns:
        col_lower = col.lower()
        if col_lower in ("fecha_cierre", "fecha"):
            col_map[col] = "fecha"
        elif col_lower in ("empresa", "fondo", "nemotecnico"):
            col_map[col] = "fondo"
        elif col_lower in ("valor_cuota", "precio", "valor"):
            col_map[col] = "valor_cuota"

    if col_map:
        df = df.rename(columns=col_map)

    # Verificar columnas mínimas
    for col in ("fecha", "fondo", "valor_cuota"):
        if col not in df.columns:
            raise ValueError(
                f"Columna '{col}' no encontrada en respuesta API. "
                f"Columnas disponibles: {list(df.columns)}"
            )

    df["fecha"]       = pd.to_datetime(df["fecha"])
    df["fondo"]       = df["fondo"].astype(str).str.strip()
    df["valor_cuota"] = pd.to_numeric(df["valor_cuota"], errors="coerce")
    df = df.dropna(subset=["fecha", "valor_cuota", "fondo"])
    df = df[df["valor_cuota"] > 0]

    n_fondos = df["fondo"].nunique()
    print(f"      {n_fondos} fondos, {len(df)} registros")
    return df[["fecha", "fondo", "valor_cuota"]].sort_values("fecha").reset_index(drop=True)
