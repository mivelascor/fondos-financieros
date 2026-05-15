"""
scripts/main.py
Orquestador del pipeline de folletos VanTrust.

Uso:
    python main.py "Comentario CLP..." "Comentario USD..."

El script:
1. Lee templates TEMPLATE_FONDO_*.xlsx desde scripts/data/
2. Extrae datos con extraer_datos.py
3. Inyecta comentarios del PM
4. Genera .pptx con generar_folleto.js  (soporta flags y posicional)
5. Convierte a PDF con LibreOffice
6. Guarda en folletos/YYYY-MM/ y crea latest.zip
"""

import sys
import os
import json
import subprocess
import zipfile
from pathlib import Path
from datetime import datetime

sys.path.insert(0, str(Path(__file__).parent))
from extraer_datos import extraer_datos

# ─── Rutas ─────────────────────────────────────────────────────────────────────
SCRIPT_DIR   = Path(__file__).parent
REPO_ROOT    = SCRIPT_DIR.parent
DATA_DIR     = SCRIPT_DIR / 'data'
FOLLETOS_DIR = REPO_ROOT  / 'folletos'
GENJS        = SCRIPT_DIR / 'generador' / 'generar_folleto.js'
CFG_DIR      = SCRIPT_DIR / 'configs'

# ─── Fondos a procesar ─────────────────────────────────────────────────────────
FONDOS_CONFIG = [
    # (template_xlsx,                              config_json,                        moneda)
    ('TEMPLATE_FONDO_LIQUIDEZ.xlsx',               'FIP_Liquidez_Sencillo.json',       'clp'),
    ('TEMPLATE_FONDO_LIQUIDEZ_UNO.xlsx',           'FIP_Liquidez_Uno.json',            'clp'),
    ('TEMPLATE_FONDO_LIQUIDEZ_LOCAL.xlsx',         'FIP_Liquidez_Local.json',          'clp'),
    ('TEMPLATE_FONDO_LIQUIDEZ_ACTIVA.xlsx',        'FIP_Liquidez_Activa.json',         'clp'),
    ('TEMPLATE_FONDO_LIQUIDEZ_ALTO_MONTO.xlsx',    'FIP_Liquidez_Alto_Monto.json',     'clp'),
    ('TEMPLATE_FONDO_LIQUIDEZ_CAJA.xlsx',          'FIP_Liquidez_Caja.json',           'clp'),
    ('TEMPLATE_FONDO_LIQUIDEZ_CONTINUA.xlsx',      'FIP_Liquidez_Continua.json',       'clp'),
    ('TEMPLATE_FONDO_LIQUIDEZ_CORRIENTE.xlsx',     'FIP_Liquidez_Corriente.json',      'clp'),
    ('TEMPLATE_FONDO_LIQUIDEZ_CORTO_PLAZO.xlsx',   'FIP_Liquidez_Corto_Plazo.json',    'clp'),
    ('TEMPLATE_FONDO_LIQUIDEZ_DISPONIBLE.xlsx',    'FIP_Liquidez_Disponible.json',     'clp'),
    ('TEMPLATE_FONDO_LIQUIDEZ_EFECTIVO.xlsx',      'FIP_Liquidez_Efectivo.json',       'clp'),
    ('TEMPLATE_FONDO_LIQUIDEZ_FLEXIBLE.xlsx',      'FIP_Liquidez_Flexible.json',       'clp'),
    ('TEMPLATE_FONDO_LIQUIDEZ_MONETARIO.xlsx',     'FIP_Liquidez_Monetario.json',      'clp'),
    ('TEMPLATE_FONDO_LIQUIDEZ_PLUS.xlsx',          'FIP_Liquidez_Plus.json',           'clp'),
    ('TEMPLATE_FONDO_LIQUIDEZ_RENDIMIENTO.xlsx',   'FIP_Liquidez_Rendimiento.json',    'clp'),
    ('TEMPLATE_FONDO_LIQUIDEZ_RESERVA_DOLAR.xlsx', 'FIP_Liquidez_Reserva_Dolar.json',  'usd'),
    ('TEMPLATE_FONDO_LIQUIDEZ_DOLAR.xlsx',         'FIP_Liquidez_Dolar.json',          'usd'),
    ('TEMPLATE_FONDO_LIQUIDEZ_DOLAR_CAJA.xlsx',    'FIP_Liquidez_Dolar_Caja.json',     'usd'),
    ('TEMPLATE_FONDO_ALTO_CAPITAL.xlsx',           'FIP_Alto_Capital.json',            'clp'),
    ('TEMPLATE_FONDO_ALTO_APORTE.xlsx',            'FIP_Alto_Aporte.json',             'clp'),
    ('TEMPLATE_FONDO_ALTO_PATRIMONIO.xlsx',        'FIP_Alto_Patrimonio.json',         'clp'),
    ('TEMPLATE_FONDO_EXTRA.xlsx',                  'FIP_Extra.json',                   'clp'),
    ('TEMPLATE_FONDO_USD_MONEY_MARKET.xlsx',       'FIP_USD_Money_Market.json',        'usd'),
    ('TEMPLATE_FONDO_FACTURA_DOLAR.xlsx',          'FIP_Factura_Dolar.json',           'usd'),
    # Fondos con nombres especiales en el xlsx
    ('TEMPLATE_FONDO_LIQUIDEZ_Monetario_I.xlsx',   'FIP_Liquidez_Monetario_I.json',    'clp'),
    ('TEMPLATE_FONDO_LIQUIDEZ_Disponible_I.xlsx',  'FIP_Liquidez_Disponible_I.json',   'clp'),
    ('TEMPLATE_FONDO_LIQUIDEZ_Presente.xlsx',      'FIP_Liquidez_Presente.json',       'clp'),
    ('TEMPLATE_FONDO_LIQUIDEZ_Permanente.xlsx',    'FIP_Liquidez_Permanente.json',     'clp'),
]


def log(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


def generar_pptx(datos_json: Path, cfg_json: Path, output_pptx: Path):
    """Llama a generar_folleto.js. Soporta argumentos posicionales."""
    cmd = ['node', str(GENJS), str(datos_json), str(cfg_json), str(output_pptx)]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or result.stdout.strip())
    return output_pptx


def convertir_pdf(pptx_path: Path, output_dir: Path) -> Path:
    """Convierte .pptx → .pdf con LibreOffice headless."""
    result = subprocess.run(
        ['soffice', '--headless', '--convert-to', 'pdf',
         '--outdir', str(output_dir), str(pptx_path)],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip())
    pdf = output_dir / (pptx_path.stem + '.pdf')
    if not pdf.exists():
        raise FileNotFoundError(f"PDF no generado: {pdf}")
    return pdf


def crear_zip(carpeta_mes: Path) -> Path:
    zip_path = FOLLETOS_DIR / 'latest.zip'
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for pdf in sorted(carpeta_mes.glob('*.pdf')):
            zf.write(pdf, pdf.name)
    log(f"ZIP: {zip_path.name}  ({zip_path.stat().st_size // 1024} KB, {len(list(carpeta_mes.glob('*.pdf')))} PDFs)")
    return zip_path


def main():
    if len(sys.argv) < 3:
        print("Uso: python main.py '<comentario_clp>' '<comentario_usd>'")
        sys.exit(1)

    comentario_clp = sys.argv[1].strip()
    comentario_usd = sys.argv[2].strip()

    mes_str      = datetime.now().strftime('%Y-%m')
    carpeta_mes  = FOLLETOS_DIR / mes_str
    carpeta_mes.mkdir(parents=True, exist_ok=True)

    tmp_dir = Path('/tmp/folletos_vantrust')
    tmp_dir.mkdir(exist_ok=True)

    generados = 0
    errores   = []

    for template_file, cfg_file, moneda in FONDOS_CONFIG:
        template_path = DATA_DIR / template_file
        cfg_path      = CFG_DIR  / cfg_file

        if not template_path.exists():
            log(f"⚠  Sin template: {template_file}")
            continue
        if not cfg_path.exists():
            log(f"⚠  Sin config:   {cfg_file}")
            continue

        try:
            log(f"→  {template_file}")

            # 1. Extraer datos del Excel
            datos = extraer_datos(str(template_path))

            # 2. Inyectar comentario del PM
            datos['comentario'] = comentario_usd if moneda == 'usd' else comentario_clp

            # 3. Guardar JSON temporal
            slug = datos['nombre_fondo'].replace(' ', '_').replace('/', '-')
            datos_json  = tmp_dir / f"{slug}.json"
            output_pptx = tmp_dir / f"{slug}.pptx"

            with open(datos_json, 'w', encoding='utf-8') as f:
                json.dump(datos, f, ensure_ascii=False, default=str)

            # 4. Generar PPTX
            generar_pptx(datos_json, cfg_path, output_pptx)

            # 5. Convertir a PDF
            pdf = convertir_pdf(output_pptx, carpeta_mes)
            log(f"✓  {pdf.name}")
            generados += 1

        except Exception as e:
            log(f"✗  Error en {template_file}: {e}")
            errores.append((template_file, str(e)))

    # 6. ZIP
    if generados > 0:
        crear_zip(carpeta_mes)
    else:
        log("ERROR: Sin folletos generados. Abortando.")
        sys.exit(1)

    if errores:
        log(f"\n⚠  {len(errores)} error(es):")
        for f, e in errores:
            log(f"   {f}: {e}")
        sys.exit(1)

    log(f"\n✅ {generados} folletos en folletos/{mes_str}/")


if __name__ == '__main__':
    main()
