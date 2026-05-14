"""
etl/sql_extractor.py — Extrae datos desde SQL Server.
Lee valores cuota diarios y composición de cartera.
"""
import pandas as pd
from sqlalchemy import create_engine, text
from config import SQL_CONN, TABLA_VALORES_CUOTA, TABLA_CARTERA


def get_engine():
    return create_engine(SQL_CONN)


def get_valores_cuota_eom() -> pd.DataFrame:
    """
    Lee valores cuota diarios y filtra el último día de cada mes.
    Retorna: fecha, fondo, moneda, valor_cuota
    """
    engine = get_engine()
    query = f"""
        SELECT
            FECHA,
            NEMOTECNICO  AS fondo,
            COD_MONEDA   AS moneda,
            PRECIO       AS valor_cuota
        FROM {TABLA_VALORES_CUOTA}
        WHERE FECHA >= DATEADD(MONTH, -14, GETDATE())
        ORDER BY FECHA ASC
    """
    df = pd.read_sql(query, engine, parse_dates=["FECHA"])
    df.rename(columns={"FECHA": "fecha"}, inplace=True)

    # Filtrar último día disponible de cada mes por fondo
    df["anio_mes"] = df["fecha"].dt.to_period("M")
    df_eom = (
        df.sort_values("fecha")
          .groupby(["fondo", "anio_mes"])
          .last()
          .reset_index()
          .drop(columns=["anio_mes"])
    )
    df_eom["fondo"] = df_eom["fondo"].str.strip()
    return df_eom[["fecha", "fondo", "moneda", "valor_cuota"]]


def get_cartera(fecha_cierre: str) -> pd.DataFrame:
    """
    Lee composición de cartera para la fecha de cierre dada (YYYY-MM-DD).
    Retorna: fondo, instrumento, moneda, duracion, monto, pct
    """
    engine = get_engine()
    query = f"""
        SELECT
            DSC_CUENTA                          AS fondo,
            SUB_CLASE                           AS instrumento,
            COD_MONEDA                          AS moneda,
            TRAMO                               AS duracion,
            VALOR_PRESENTE_MERCADO_MON_CTA      AS monto
        FROM {TABLA_CARTERA}
        WHERE FECHA_CIERRE = '{fecha_cierre}'
    """
    df = pd.read_sql(query, engine)
    df["fondo"] = df["fondo"].str.strip()

    # Calcular porcentaje sobre el total de cada fondo
    totales = df.groupby("fondo")["monto"].sum().rename("total")
    df = df.merge(totales, on="fondo")
    df["pct"] = df["monto"] / df["total"]
    return df[["fondo", "instrumento", "moneda", "duracion", "monto", "pct"]]
