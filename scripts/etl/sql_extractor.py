"""
scripts/etl/sql_extractor.py

Obtiene valores cuota desde la API REST interna de VanTrust.

API:  POST https://claudeods.vantrustcapital.cl/query
      Body: {"Sql": "SELECT ..."}
Tabla: ODS.VALORES_CUOTA_GPI
  EMPRESA      → nombre del fondo
  FECHA_CIERRE → fecha del valor cuota
  VALOR_CUOTA  → precio de la cuota
"""

import requests
import pandas as pd
from datetime import date, timedelta

API_URL = "https://claudeods.vantrustcapital.cl/query"
HEADERS = {"Content-Type": "application/json"}

# Nombres exactos en la base de datos → nombre del FIP en el folleto
FONDOS_MAP = {
    "FIP VANTRUST LIQUIDEZ":            "FIP Liquidez Sencillo",
    "FIP VANTRUST LIQUIDEZ ACTIVA":     "FIP Liquidez Activa",
    "FIP VANTRUST LIQUIDEZ ALTO APORTE":"FIP Alto Aporte",
    "FIP VANTRUST LIQUIDEZ ALTO CAPITAL":"FIP Alto Capital",
    "FIP VANTRUST LIQUIDEZ ALTO MONTO": "FIP Liquidez Alto Monto",
    "FIP VANTRUST LIQUIDEZ ALTO PATRIMONIO": "FIP Liquidez Gran Patrimonio",
    "FIP VANTRUST LIQUIDEZ CAJA":       "FIP Liquidez Caja",
    "FIP VANTRUST LIQUIDEZ CONTINUA":   "FIP Liquidez Continua",
    "FIP VANTRUST LIQUIDEZ CORRIENTE":  "FIP Liquidez Corriente",
    "FIP VANTRUST LIQUIDEZ CORTO PLAZO":"FIP Liquidez Corto Plazo",
    "FIP VANTRUST LIQUIDEZ DISPONIBLE": "FIP Liquidez Disponible",
    "FIP VANTRUST LIQUIDEZ DISPONIBLE I":"FIP Liquidez Disponible I",
    "FIP VANTRUST LIQUIDEZ DOLAR":      "FIP Liquidez Dólar",
    "FIP VANTRUST LIQUIDEZ DOLAR CAJA": "FIP Liquidez Dólar Caja",
    "FIP VANTRUST LIQUIDEZ EFECTIVO":   "FIP Liquidez Efectivo",
    "FIP VANTRUST LIQUIDEZ FLEXIBLE":   "FIP Liquidez Flexible",
    "FIP VANTRUST LIQUIDEZ FLEXIBLE DÓLAR": "FIP Liquidez Flexible Dólar",
    "FIP VANTRUST LIQUIDEZ I":          "FIP Liquidez Uno",
    "FIP VANTRUST LIQUIDEZ LOCAL":      "FIP Liquidez Local",
    "FIP VANTRUST LIQUIDEZ MONETARIO":  "FIP Liquidez Monetario",
    "FIP VANTRUST LIQUIDEZ MONETARIO I":"FIP Liquidez Monetario I",
    "FIP VANTRUST LIQUIDEZ PERMANENTE": "FIP Liquidez Permanente",
    "FIP VANTRUST LIQUIDEZ PLUS":       "FIP Liquidez Plus",
    "FIP VANTRUST LIQUIDEZ PRESENTE":   "FIP Liquidez Presente",
    "FIP VANTRUST LIQUIDEZ RECURRENTE": "FIP Liquidez Sencillo",
    "FIP VANTRUST LIQUIDEZ RENDIMIENTO":"FIP Liquidez Rendimiento",
    "FIP VANTRUST LIQUIDEZ RESERVA DÓLAR": "FIP Liquidez Reserva Dólar",
    "FIP VANTRUST LIQUIDEZ SENCILLO":   "FIP Liquidez Sencillo",
    "FIP VANTRUST LIQUIDEZ TEMPORAL":   "FIP Liquidez Sencillo",
    "FIP VANTRUST FACTURA DOLAR":       "FIP Factura Dólar",
}


def _query(sql: str) -> list:
    """Ejecuta un SELECT en la API interna y retorna lista de filas."""
    resp = requests.post(API_URL, json={"Sql": sql}, headers=HEADERS, timeout=60)
    resp.raise_for_status()
    data = resp.json()
    if isinstance(data, list):
        return data
    return data.get("rows", data.get("data", data.get("results", [])))


def get_valores_cuota_eom() -> dict:
    """
    Descarga valores cuota de todos los fondos desde el inicio (toda la historia)
    y retorna el último valor cuota de cada mes (EOM) por fondo.

    Retorna:
        { 'FIP Liquidez Uno': { (2020, 4): 1234.56, (2020, 5): 1238.90, ... }, ... }
    """
    # Buscar desde el inicio de todos los fondos (2017 es cuando comenzó el primero)
    desde = "2017-01-01"

    nombres_sql = ", ".join(f"'{n}'" for n in FONDOS_MAP.keys())

    sql = f"""
        SELECT
            EMPRESA      AS fondo,
            FECHA_CIERRE AS fecha,
            VALOR_CUOTA  AS valor_cuota
        FROM ODS.VALORES_CUOTA_GPI
        WHERE FECHA_CIERRE >= '{desde}'
          AND VALOR_CUOTA > 0
          AND EMPRESA IN ({nombres_sql})
        ORDER BY EMPRESA, FECHA_CIERRE ASC
    """

    print("  [SQL] Consultando ODS.VALORES_CUOTA_GPI...", flush=True)
    rows = _query(sql)

    if not rows:
        raise ValueError(
            "La API no retornó datos de valores cuota. "
            "Verifica que claudeods.vantrustcapital.cl esté accesible."
        )

    df = pd.DataFrame(rows)
    df["fecha"]       = pd.to_datetime(df["fecha"])
    df["valor_cuota"] = pd.to_numeric(df["valor_cuota"], errors="coerce")
    df["fondo"]       = df["fondo"].str.strip()
    df = df.dropna(subset=["fecha", "valor_cuota"])

    # Tomar el último dato disponible de cada mes (EOM)
    df["año"] = df["fecha"].dt.year
    df["mes"] = df["fecha"].dt.month

    resultado = {}
    for nombre_bd, grupo in df.groupby("fondo"):
        nombre_fip = FONDOS_MAP.get(nombre_bd)
        if not nombre_fip:
            continue
        eom = (
            grupo.sort_values("fecha")
                 .groupby(["año", "mes"])["valor_cuota"]
                 .last()
        )
        eom_dict = {(int(a), int(m)): float(v) for (a, m), v in eom.items()}

        # Si el mismo FIP tiene múltiples nombres en BD, combinar
        if nombre_fip not in resultado:
            resultado[nombre_fip] = eom_dict
        else:
            resultado[nombre_fip].update(eom_dict)

    print(f"  [SQL] {len(resultado)} fondos, "
          f"{sum(len(v) for v in resultado.values())} observaciones EOM", flush=True)
    return resultado
