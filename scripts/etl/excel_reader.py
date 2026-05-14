"""
etl/excel_reader.py — Lee el Excel de cartera subido manualmente via admin.html.

El archivo cartera.xlsx se sube al repo en inputs/cartera.xlsx
antes de ejecutar el workflow.

Columnas esperadas:
  fondos, SUB CLASE, COD_MONEDA, TRAMO,
  VALOR_PRESENTE_MERCADO_MON_CTA, total
"""
import pandas as pd
from config import ARCHIVO_CARTERA


def get_cartera() -> pd.DataFrame:
    """
    Lee cartera.xlsx y calcula el % de cada instrumento por fondo.
    Retorna: fondo, instrumento, moneda, duracion, monto, pct
    """
    df = pd.read_excel(ARCHIVO_CARTERA)

    # Columnas requeridas
    cols = {
        "fondos":                           "fondo",
        "SUB CLASE":                        "instrumento",
        "COD_MONEDA":                       "moneda",
        "TRAMO":                            "duracion",
        "VALOR_PRESENTE_MERCADO_MON_CTA":   "monto",
        "total":                            "total",
    }
    df = df[list(cols.keys())].rename(columns=cols)
    df["fondo"] = df["fondo"].astype(str).str.strip()
    df = df.dropna(subset=["fondo", "monto", "total"])
    df = df[df["monto"] > 0]
    df["pct"] = df["monto"] / df["total"]
    return df[["fondo", "instrumento", "moneda", "duracion", "monto", "pct"]]
