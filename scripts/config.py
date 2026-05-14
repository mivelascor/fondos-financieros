"""config.py — Configuración central."""
import os
from pathlib import Path

BASE_DIR      = Path(__file__).parent
INPUTS_DIR    = BASE_DIR.parent / "inputs"
OUTPUT_DIR    = BASE_DIR.parent / "folletos"
TEMPLATE_PPTX = BASE_DIR / "templates" / "folleto_template_clean.pptx"
ARCHIVO_CARTERA = INPUTS_DIR / "cartera.xlsx"

GITHUB_TOKEN  = os.environ.get("GH_TOKEN", "")
GITHUB_REPO   = os.environ.get("GH_REPO", "mivelascor/fondos-financieros")
GITHUB_BRANCH = "main"

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
        "rut": "77.155.267-6", "remuneracion": "0,295% IVA Incluido",
        "fecha_inicio": "Abril 2020",
    },
    "FIP VANTRUST LIQUIDEZ RESERVA DOLAR": {
        "rut": "76.637.335-6", "remuneracion": "0,50% TPM + IVA",
        "fecha_inicio": "Febrero 2025",
    },
}

INFO_DEFAULT = {
    "administradora": "Vantrust Gestion Patrimonial S.A.",
    "rut":            "76,637,334-8",
    "tipo":           "Fondo de Inversión Privado",
    "benchmark":      "Índice Cámara Promedio (ICP)",
    "plazo_rescate":  "A más tardar 15 días corridos",
    "remuneracion":   "0,295% IVA Incluido",
    "fecha_inicio":   "",
}

def get_info_fondo(nombre: str, moneda: str, fecha_inicio_str: str) -> dict:
    info = dict(INFO_DEFAULT)
    info.update(INFO_POR_FONDO.get(nombre, {}))
    info["moneda"] = moneda
    info["fecha_inicio"] = fecha_inicio_str or info.get("fecha_inicio", "")
    return info
