"""etl/sql_extractor.py — Valor cuota desde API REST interna."""
import requests, pandas as pd
from datetime import date, timedelta

API_URL = "https://claudeods.vantrustcapital.cl/query"
HEADERS = {"Content-Type": "application/json"}

def _query(sql):
    resp = requests.post(API_URL, json={"Sql": sql}, headers=HEADERS, timeout=60)
    resp.raise_for_status()
    data = resp.json()
    if isinstance(data, list): return data
    return data.get("rows", data.get("data", data.get("results", [])))

def get_valores_cuota_eom():
    desde = (date.today() - timedelta(days=455)).strftime("%Y-%m-%d")
    sql = f"""
        SELECT FECHA_CIERRE AS fecha, RTRIM(LTRIM(EMPRESA)) AS fondo, VALOR_CUOTA AS valor_cuota
        FROM ODS.VALORES_CUOTA_GPI
        WHERE FECHA_CIERRE >= '{desde}' AND VALOR_CUOTA > 0
        ORDER BY FECHA_CIERRE ASC
    """
    print("    Consultando ODS.VALORES_CUOTA_GPI...")
    rows = _query(sql)
    if not rows:
        raise ValueError("Sin datos de valores cuota.")
    df = pd.DataFrame(rows)
    df["fecha"] = pd.to_datetime(df["fecha"])
    df["fondo"] = df["fondo"].str.strip()
    df["valor_cuota"] = pd.to_numeric(df["valor_cuota"], errors="coerce")
    df = df.dropna(subset=["fecha","valor_cuota","fondo"])
    # Retornar serie diaria completa (sin agrupar EOM - lo hace rentabilidades.py)
    print(f"      {df['fondo'].nunique()} fondos, {len(df)} registros")
    return df
