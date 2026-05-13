"""
config.py — Configuración central del sistema de folletos.
Todas las variables de entorno y parámetros globales van aquí.
"""
import os
from pathlib import Path

# ── Rutas ────────────────────────────────────────────────────────────────────
BASE_DIR     = Path(__file__).parent                        # scripts/
TEMPLATE_PPTX = BASE_DIR / "templates" / "folleto_template_clean.pptx"
OUTPUT_DIR   = BASE_DIR.parent / "folletos"                 # folletos/
DATA_DIR     = BASE_DIR / "data"

# ── SQL ──────────────────────────────────────────────────────────────────────
# Se lee desde variable de entorno (GitHub Secret: SQL_CONN)
SQL_CONN = os.environ.get("SQL_CONN", "")

# Nombres de tablas — ajusta si difieren en tu base de datos
TABLA_VALORES_CUOTA = "dbo.valores_cuota"   # columnas: FECHA, NEMOTECNICO, COD_MONEDA, PRECIO
TABLA_CARTERA       = "dbo.cartera_fondos"  # columnas: DSC_CUENTA, SUB_CLASE, COD_MONEDA, TRAMO, VALOR_PRESENTE_MERCADO_MON_CTA, FECHA_CIERRE

# ── BCCh (ICP) ───────────────────────────────────────────────────────────────
BCCH_USER = os.environ.get("BCCH_USER", "")
BCCH_PASS = os.environ.get("BCCH_PASS", "")
BCCH_SERIE_ICP = "F022.ICP.IND.N.M"   # Serie del ICP en el BCCh — verifica el código exacto

# ── GitHub ───────────────────────────────────────────────────────────────────
GITHUB_TOKEN  = os.environ.get("GH_TOKEN", "")
GITHUB_REPO   = os.environ.get("GH_REPO", "mivelascor/fondos-financieros")
GITHUB_BRANCH = "main"

# ── LibreOffice ───────────────────────────────────────────────────────────────
LIBREOFFICE_PATH = "/usr/bin/libreoffice"   # Ubuntu (GitHub Actions)
# Si corres en Windows localmente cambia a:
# LIBREOFFICE_PATH = r"C:\Program Files\LibreOffice\program\soffice.exe"

# ── Competencia CMF ──────────────────────────────────────────────────────────
# Fondos que NO contienen "DOLAR" ni "USD" en el nombre → usan competencia CLP
CMF_COMP_CLP = {
    "nombre": "FONDO MUTUO SANTANDER MONEY MARKET",
    "rut":    "8057",
    "row":    "AAAw cAAhAAAACcAAs",
}
# Fondos que contienen "DOLAR" o "USD" → usan competencia USD
CMF_COMP_USD = {
    "nombre": "FONDO MUTUO BANCHILE CORPORATE DOLLAR",
    "rut":    "8248",
    "row":    "",
}

# ── Posiciones OLE (en EMU) ───────────────────────────────────────────────────
# Estas coordenadas vienen de correr preparar_template.py sobre tu PPTX original.
# Los valores de abajo son los que leímos del archivo FIP_Liquidez_Activa_AUTO.pptx.
# Si el diseño cambia, vuelve a correr preparar_template.py para actualizarlos.
OLE_POSITIONS = {
    "tabla_rentabilidad":  {"left": 3762590, "top": 1462653, "width": 4743450, "height": 1533525},
    "grafico_evolucion":   {"left": 3806246, "top": 3174018, "width": 4656137, "height": 2860675},
    "grafico_composicion": {"left": 3927283, "top": 6762063, "width": 4267200, "height":  885825},
    "tabla_comparacion":   {"left": 3616035, "top": 8713003, "width": 5042553, "height": 1898057},
}

# ── Lista de fondos con folleto ───────────────────────────────────────────────
# Nombre exacto como aparece en la columna NEMOTECNICO del SQL.
# Agrega o quita fondos según corresponda.
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

# Información estática por fondo (lo que va en la sección "Información General")
# Si un fondo no está aquí se usan los valores por defecto.
INFO_FONDOS = {
    "FIP VANTRUST LIQUIDEZ ACTIVA": {
        "administradora": "Vantrust Gestion Patrimonial S.A.",
        "rut":            "76,637,334-8",
        "tipo":           "Fondo de Inversión Privado",
        "benchmark":      "Índice Cámara Promedio (ICP)",
        "plazo_rescate":  "A más tardar 15 días corridos",
        "objetivo":       "Invertir los recursos del fondo en instrumentos de deuda de corto y mediano plazo, en una cartera diversificada, obteniendo una rentabilidad igual o superior al ICP.",
        "inversionistas": "Dirigida a empresas y personas que buscan invertir sus excedentes de caja con una rentabilidad de corto plazo y baja tolerancia al riesgo.",
    },
    # Los demás fondos heredan los mismos valores por defecto.
    # Agrega entradas específicas si algún fondo tiene datos distintos.
}

INFO_FONDO_DEFAULT = {
    "administradora": "Vantrust Gestion Patrimonial S.A.",
    "rut":            "76,637,334-8",
    "tipo":           "Fondo de Inversión Privado",
    "benchmark":      "Índice Cámara Promedio (ICP)",
    "plazo_rescate":  "A más tardar 15 días corridos",
    "objetivo":       "Invertir los recursos del fondo en instrumentos de deuda de corto y mediano plazo, en una cartera diversificada, obteniendo una rentabilidad igual o superior al ICP.",
    "inversionistas": "Dirigida a empresas y personas que buscan invertir sus excedentes de caja.",
}

def get_info_fondo(nombre: str, moneda: str) -> dict:
    info = dict(INFO_FONDOS.get(nombre, INFO_FONDO_DEFAULT))
    info["moneda"] = moneda
    return info
