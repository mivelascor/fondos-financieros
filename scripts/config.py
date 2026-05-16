"""config.py — Configuración central."""
import os
from pathlib import Path

BASE_DIR        = Path(__file__).parent
INPUTS_DIR      = BASE_DIR.parent / "inputs"
OUTPUT_DIR      = BASE_DIR.parent / "folletos"
ARCHIVO_CARTERA = INPUTS_DIR / "cartera.xlsx"

GITHUB_TOKEN  = os.environ.get("GH_TOKEN", "")
GITHUB_REPO   = os.environ.get("GH_REPO", "mivelascor/fondos-financieros")
GITHUB_BRANCH = "main"

LIBREOFFICE_PATH = "/usr/bin/libreoffice"

CMF_COMP_CLP = {"nombre": "FONDO MUTUO SANTANDER MONEY MARKET",   "rut": "8057"}
CMF_COMP_USD = {"nombre": "FONDO MUTUO BANCHILE CORPORATE DOLLAR", "rut": "8248"}

# Meses en español (para fechas de inicio de fondos)
MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}

FONDOS_CON_FOLLETO = [
    "FIP VANTRUST LIQUIDEZ ACTIVA",
    "FIP VANTRUST LIQUIDEZ ALTO APORTE",
    "FIP VANTRUST LIQUIDEZ ALTO CAPITAL",
    "FIP VANTRUST LIQUIDEZ ALTO MONTO",
    "FIP VANTRUST LIQUIDEZ CAJA",
    "FIP VANTRUST LIQUIDEZ CONTINUA",
    "FIP VANTRUST LIQUIDEZ CORRIENTE",
    "FIP VANTRUST LIQUIDEZ CORTO PLAZO",
    "FIP VANTRUST LIQUIDEZ DISPONIBLE I",
    "FIP VANTRUST LIQUIDEZ DOLAR",
    "FIP VANTRUST LIQUIDEZ DOLAR CAJA",
    "FIP VANTRUST LIQUIDEZ EFECTIVO",
    "FIP VANTRUST LIQUIDEZ FLEXIBLE",
    "FIP VANTRUST LIQUIDEZ I",
    "FIP VANTRUST LIQUIDEZ LOCAL",
    "FIP VANTRUST LIQUIDEZ MONETARIO I",
    "FIP VANTRUST LIQUIDEZ PERMANENTE",
    "FIP VANTRUST LIQUIDEZ PLUS",
    "FIP VANTRUST LIQUIDEZ PRESENTE",
    "FIP VANTRUST LIQUIDEZ RECURRENTE",
    "FIP VANTRUST LIQUIDEZ RENDIMIENTO",
    "FIP VANTRUST LIQUIDEZ RESERVA DOLAR",
    "FIP VANTRUST LIQUIDEZ SENCILLO",
    "FIP VANTRUST LIQUIDEZ TEMPORAL",
]

INFO_POR_FONDO = {
    "FIP VANTRUST LIQUIDEZ I": {
        "rut": "77.155.267-6",
        "remuneracion": "0,295% IVA Incluido",
        "fecha_inicio_fija": "Abril 2020",
    },
    "FIP VANTRUST LIQUIDEZ RESERVA DOLAR": {
        "rut": "76.637.335-6",
        "remuneracion": "0,50% TPM + IVA",
        "fecha_inicio_fija": "Febrero 2025",
    },
}

INFO_DEFAULT = {
    "administradora": "Vantrust Gestion Patrimonial S.A.",
    "rut":            "76,637,334-8",
    "tipo":           "Fondo de Inversión Privado",
    "benchmark":      "Índice Cámara Promedio (ICP)",
    "plazo_rescate":  "A más tardar 15 días corridos",
    "remuneracion":   "0,295% IVA Incluido",
}

import pandas as pd

def fecha_inicio_es(ts: pd.Timestamp) -> str:
    """Retorna la fecha de inicio en español: 'Julio 2025'"""
    return f"{MESES_ES[ts.month]} {ts.year}"

def get_info_fondo(nombre: str, moneda: str, fecha_inicio_ts: pd.Timestamp) -> dict:
    info = dict(INFO_DEFAULT)
    info.update(INFO_POR_FONDO.get(nombre, {}))
    info["moneda"] = moneda
    # Usar fecha fija si está definida, si no calcularla en español
    if "fecha_inicio_fija" in info:
        info["fecha_inicio"] = info.pop("fecha_inicio_fija")
    else:
        info["fecha_inicio"] = fecha_inicio_es(fecha_inicio_ts)
    return info
