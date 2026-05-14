"""
config.py — Configuración central.

Fuentes de datos:
  - Valor cuota : API REST interna (automático)
  - ICP         : API BCCh (automático, requiere BCCH_USER / BCCH_PASS)
  - Cartera     : inputs/cartera.xlsx subido via admin.html (manual)

Input del PM via admin.html:
  - cartera.xlsx
  - Comentario CLP
  - Comentario USD
"""
import os
from pathlib import Path

# ── Rutas ─────────────────────────────────────────────────────────────────────
BASE_DIR      = Path(__file__).parent
INPUTS_DIR    = BASE_DIR.parent / "inputs"
OUTPUT_DIR    = BASE_DIR.parent / "folletos"
TEMPLATE_PPTX = BASE_DIR / "templates" / "folleto_template_clean.pptx"

# Solo la cartera es input manual
ARCHIVO_CARTERA = INPUTS_DIR / "cartera.xlsx"

# ── API SQL interna ───────────────────────────────────────────────────────────
SQL_API_URL = "https://claudeods.vantrustcapital.cl/query"

# ── BCCh API ──────────────────────────────────────────────────────────────────
BCCH_USER = os.environ.get("BCCH_USER", "")
BCCH_PASS = os.environ.get("BCCH_PASS", "")

# ── GitHub ────────────────────────────────────────────────────────────────────
GITHUB_TOKEN  = os.environ.get("GH_TOKEN", "")
GITHUB_REPO   = os.environ.get("GH_REPO", "mivelascor/fondos-financieros")
GITHUB_BRANCH = "main"

# ── LibreOffice ───────────────────────────────────────────────────────────────
LIBREOFFICE_PATH = "/usr/bin/libreoffice"

# ── Competencia CMF ───────────────────────────────────────────────────────────
CMF_COMP_CLP = {"nombre": "FONDO MUTUO SANTANDER MONEY MARKET",
                "rut": "8057", "row": "AAAw cAAhAAAACcAAs"}
CMF_COMP_USD = {"nombre": "FONDO MUTUO BANCHILE CORPORATE DOLLAR",
                "rut": "8248", "row": ""}

# ── Posiciones OLE (EMU) ──────────────────────────────────────────────────────
OLE_POSITIONS = {
    "tabla_rentabilidad":  {"left": 3762590, "top": 1462653, "width": 4743450, "height": 1533525},
    "grafico_evolucion":   {"left": 3806246, "top": 3174018, "width": 4656137, "height": 2860675},
    "grafico_composicion": {"left": 3927283, "top": 6762063, "width": 4267200, "height":  885825},
    "tabla_comparacion":   {"left": 3616035, "top": 8713003, "width": 5042553, "height": 1898057},
}

# ── Fondos con folleto ────────────────────────────────────────────────────────
# Nombre exacto como aparece en VALORES_CUOTA_GPI.EMPRESA
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

INFO_FONDO_DEFAULT = {
    "administradora": "Vantrust Gestion Patrimonial S.A.",
    "rut":            "76,637,334-8",
    "tipo":           "Fondo de Inversión Privado",
    "benchmark":      "Índice Cámara Promedio (ICP)",
    "plazo_rescate":  "A más tardar 15 días corridos",
}

def get_info_fondo(nombre: str, moneda: str) -> dict:
    info = dict(INFO_FONDO_DEFAULT)
    info["moneda"] = moneda
    return info
