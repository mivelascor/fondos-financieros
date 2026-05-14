"""
etl/excel_reader.py — Lee los 2 Excels de input manuales.

Archivos en inputs/:
  - valor_cuota.xlsx  → hoja 'Consulta1 (2)': FECHA, NEMOTECNICO, COD_MONEDA, PRECIO
  - cartera.xlsx      → columnas: fondos, SUB CLASE, COD_MONEDA, TRAMO,
                        VALOR_PRESENTE_MERCADO_MON_CTA, total
"""
import pandas as pd
from config import ARCHIVO_VALOR_CUOTA, ARCHIVO_CARTERA


def get_valores_cuota_eom() -> pd.DataFrame:
    """Lee valor_cuota.xlsx y retorna el último valor de cada mes por fondo."""
    df = pd.read_excel(
        ARCHIVO_VALOR_CUOTA,
        sheet_name="Consulta1 (2)",
        usecols=["FECHA", "NEMOTECNICO", "COD_MONEDA", "PRECIO"],
    )
    df.columns = ["fecha", "fondo", "moneda", "valor_cuota"]
    df["fecha"] = pd.to_datetime(df["fecha"])
    df["fondo"] = df["fondo"].str.strip()
    df = df.dropna(subset=["fecha", "valor_cuota", "fondo"])

    df["anio_mes"] = df["fecha"].dt.to_period("M")
    df_eom = (
        df.sort_values("fecha")
          .groupby(["fondo", "anio_mes"])
          .last()
          .reset_index()
          .drop(columns=["anio_mes"])
    )
    return df_eom[["fecha", "fondo", "moneda", "valor_cuota"]]


def get_cartera() -> pd.DataFrame:
    """Lee cartera.xlsx y calcula el % de cada instrumento por fondo."""
    df = pd.read_excel(ARCHIVO_CARTERA)

    df = df[["fondos", "SUB CLASE", "COD_MONEDA", "TRAMO",
             "VALOR_PRESENTE_MERCADO_MON_CTA", "total"]].copy()
    df.columns = ["fondo", "instrumento", "moneda", "duracion", "monto", "total"]
    df["fondo"] = df["fondo"].astype(str).str.strip()
    df = df.dropna(subset=["fondo", "monto", "total"])
    df = df[df["monto"] > 0]
    df["pct"] = df["monto"] / df["total"]
    return df[["fondo", "instrumento", "moneda", "duracion", "monto", "pct"]]
