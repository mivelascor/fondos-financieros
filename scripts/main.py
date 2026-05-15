"""
scripts/main.py — v6 (definitivo)

AUTOMÁTICO (sin intervención manual):
  - Valor cuota de cada fondo  → API claudeods.vantrustcapital.cl
  - ICP benchmark               → mindicador.cl/api/tpm
  - Competencia CLP             → CMF (Santander Money Market UNIVE)
  - Competencia USD             → CMF (BanChile Corporate Dollar A)

MANUAL (via admin.html):
  - Comentario CLP              → argumento 1
  - Comentario USD              → argumento 2
  - cartera.xlsx                → scripts/inputs/cartera.xlsx
                                  (subido via admin.html antes de correr)

Uso:
    python main.py "<comentario_clp>" "<comentario_usd>"
"""

import sys
import json
import math
import subprocess
import zipfile
from pathlib import Path
from datetime import datetime, date
from dateutil.relativedelta import relativedelta

import pandas as pd
import requests

# ─── Rutas ─────────────────────────────────────────────────────────────────────
SCRIPT_DIR   = Path(__file__).parent
REPO_ROOT    = SCRIPT_DIR.parent
INPUTS_DIR   = SCRIPT_DIR / 'inputs'
FOLLETOS_DIR = REPO_ROOT  / 'folletos'
GENJS        = SCRIPT_DIR / 'generador' / 'generar_folleto.js'
CFG_DIR      = SCRIPT_DIR / 'configs'

# ─── Fondos a generar ──────────────────────────────────────────────────────────
# (config_json, moneda_comentario)
FONDOS = [
    ('FIP_Liquidez_Sencillo.json',      'clp'),
    ('FIP_Liquidez_Uno.json',           'clp'),
    ('FIP_Liquidez_Local.json',         'clp'),
    ('FIP_Liquidez_Activa.json',        'clp'),
    ('FIP_Liquidez_Alto_Monto.json',    'clp'),
    ('FIP_Liquidez_Caja.json',          'clp'),
    ('FIP_Liquidez_Continua.json',      'clp'),
    ('FIP_Liquidez_Corriente.json',     'clp'),
    ('FIP_Liquidez_Corto_Plazo.json',   'clp'),
    ('FIP_Liquidez_Disponible.json',    'clp'),
    ('FIP_Liquidez_Efectivo.json',      'clp'),
    ('FIP_Liquidez_Flexible.json',      'clp'),
    ('FIP_Liquidez_Monetario.json',     'clp'),
    ('FIP_Liquidez_Plus.json',          'clp'),
    ('FIP_Liquidez_Rendimiento.json',   'clp'),
    ('FIP_Liquidez_Monetario_I.json',   'clp'),
    ('FIP_Liquidez_Disponible_I.json',  'clp'),
    ('FIP_Liquidez_Presente.json',      'clp'),
    ('FIP_Liquidez_Permanente.json',    'clp'),
    ('FIP_Alto_Capital.json',           'clp'),
    ('FIP_Alto_Aporte.json',            'clp'),
    ('FIP_Alto_Patrimonio.json',        'clp'),
    ('FIP_Liquidez_Reserva_Dolar.json', 'usd'),
    ('FIP_Liquidez_Dolar.json',         'usd'),
    ('FIP_Liquidez_Dolar_Caja.json',    'usd'),
    ('FIP_Factura_Dolar.json',          'usd'),
    ('FIP_USD_Money_Market.json',       'usd'),
]

SESSION = requests.Session()
SESSION.headers['User-Agent'] = 'Mozilla/5.0 (VanTrust-FolletoBot/1.0)'


def log(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


# ─── 1. Valor cuota desde la API interna ──────────────────────────────────────

def get_valor_cuota() -> dict:
    """Importa y ejecuta sql_extractor.py para obtener EOM por fondo."""
    sys.path.insert(0, str(SCRIPT_DIR / 'etl'))
    from sql_extractor import get_valores_cuota_eom
    return get_valores_cuota_eom()


# ─── 2. ICP desde mindicador.cl ───────────────────────────────────────────────

def get_icp() -> dict:
    """
    Descarga TPM histórica y calcula ICP mensual.
    Retorna { (año, mes): icp_decimal }.
    """
    log("Descargando TPM desde mindicador.cl...")
    try:
        resp = SESSION.get('https://mindicador.cl/api/tpm', timeout=20)
        resp.raise_for_status()
        serie = resp.json().get('serie', [])
    except Exception as e:
        log(f"⚠  Error TPM: {e}")
        return {}

    icp = {}
    for item in serie:
        try:
            f   = datetime.strptime(item['fecha'][:10], '%Y-%m-%d').date()
            tpm = float(item['valor']) / 100
            a, m = f.year, f.month
            # días del mes
            dias = (date(a, m, 1) + relativedelta(months=1) - date(a, m, 1)).days
            icp_mes = (1 + tpm / 365) ** dias - 1
            if (a, m) not in icp or f > icp[(a, m)][0]:
                icp[(a, m)] = (f, icp_mes)
        except (KeyError, ValueError):
            continue

    result = {k: v[1] for k, v in icp.items()}
    log(f"ICP: {len(result)} meses")
    return result


# ─── 3. Competencia desde CMF ─────────────────────────────────────────────────

def _scrape_cmf_vc(url: str, nombre: str) -> dict:
    """
    Obtiene EOM de valor cuota de un fondo mutuo de la CMF.
    Retorna { (año, mes): vc }.
    """
    try:
        resp = SESSION.get(url, timeout=30)
        resp.raise_for_status()
        tablas = pd.read_html(resp.text)
        for t in tablas:
            cols_lower = [str(c).lower() for c in t.columns]
            has_fecha = any('fecha' in c or 'date' in c for c in cols_lower)
            has_valor = any(k in c for c in cols_lower for k in ('valor', 'cuota', 'precio'))
            if not (has_fecha and has_valor):
                continue
            col_f = next(c for c in t.columns if any(
                k in str(c).lower() for k in ('fecha', 'date')))
            col_v = next(c for c in t.columns if any(
                k in str(c).lower() for k in ('valor', 'cuota', 'precio')))
            t = t.copy()
            t['_f'] = pd.to_datetime(t[col_f], errors='coerce', dayfirst=True)
            t['_v'] = pd.to_numeric(
                t[col_v].astype(str).str.replace(',', '.').str.replace(' ', ''),
                errors='coerce')
            t = t.dropna(subset=['_f', '_v'])
            if t.empty:
                continue
            t['_a'] = t['_f'].dt.year
            t['_m'] = t['_f'].dt.month
            eom = t.groupby(['_a', '_m'])['_v'].last()
            return {(int(a), int(m)): float(v) for (a, m), v in eom.items()}
    except Exception as e:
        log(f"⚠  Scraping {nombre}: {e}")
    return {}


def get_competencia() -> tuple:
    """Retorna (eom_clp, eom_usd)."""
    URL_CLP = (
        'https://www.cmfchile.cl/institucional/mercados/entidad.php'
        '?mercado=V&rut=8057&grupo=&tipoentidad=RGFMU'
        '&row=AAAw%20cAAhAAAACcAAs&vig=VI&control=svs&pestania=7&'
    )
    URL_USD = (
        'https://www.cmfchile.cl/institucional/mercados/entidad.php'
        '?mercado=V&rut=8248&grupo=&tipoentidad=RGFMU'
        '&row=AAAw%20cAAhAAAACfAAj&vig=VI&control=svs&pestania=7'
    )
    log("Descargando competencia CLP (Santander Money Market)...")
    clp = _scrape_cmf_vc(URL_CLP, 'Santander MM')
    log(f"  CLP: {len(clp)} meses")

    log("Descargando competencia USD (BanChile Corporate Dollar)...")
    usd = _scrape_cmf_vc(URL_USD, 'BanChile USD')
    log(f"  USD: {len(usd)} meses")

    return clp, usd


# ─── 4. Leer cartera.xlsx ─────────────────────────────────────────────────────

def get_cartera() -> dict:
    """Lee scripts/inputs/cartera.xlsx (subido via admin.html)."""
    path = INPUTS_DIR / 'cartera.xlsx'
    if not path.exists():
        log("⚠  cartera.xlsx no encontrada en scripts/inputs/ — composición vacía")
        return {'duracion': {}, 'instrumentos': {}, 'moneda': {}}

    duracion = {}
    inst     = {}
    moneda   = {}

    DUR  = {'Menos de 1 mes', '1-3 meses', '3-4 meses', 'Más de 4 meses'}
    MON  = {'Pesos', 'UF', 'USD', 'Dólar'}
    INST = ('Depósito', 'Bono', 'Factura', 'Pagaré', 'Efectos',
            'Financi', 'Simultán', 'Caja', 'Garantía', 'Nota', 'Factoring')

    try:
        xl = pd.ExcelFile(path)
        for hoja in xl.sheet_names:
            df = xl.parse(hoja, header=None)
            for _, row in df.iterrows():
                for col in range(len(row) - 1):
                    label = str(row.iloc[col]).strip() if pd.notna(row.iloc[col]) else ''
                    try:
                        val = float(row.iloc[col + 1])
                    except (TypeError, ValueError):
                        continue
                    if label in DUR:
                        duracion[label] = val
                    elif label in MON:
                        moneda[label] = val
                    elif any(k in label for k in INST) and 0 <= val <= 1:
                        inst[label] = val
    except Exception as e:
        log(f"⚠  Error leyendo cartera.xlsx: {e}")

    log(f"Cartera: dur={len(duracion)}, inst={len(inst)}, moneda={len(moneda)}")
    return {'duracion': duracion, 'instrumentos': inst, 'moneda': moneda}


# ─── 5. Calcular rentabilidades ───────────────────────────────────────────────

def eom_a_hist(eom: dict, nombre: str) -> list:
    """Convierte {(año,mes): vc} a lista de registros con rentabilidades."""
    if not eom:
        return []
    años = sorted(set(a for a, _ in eom.keys()))
    recs = []
    for año in años:
        meses_r = []
        for mes in range(1, 13):
            ini = eom.get((año - 1, 12) if mes == 1 else (año, mes - 1))
            fin = eom.get((año, mes))
            meses_r.append(fin / ini - 1 if ini and fin and ini > 0 else None)

        vals = [v for v in meses_r if v is not None]
        if not vals:
            continue
        acc = 1.0
        for v in vals:
            acc *= (1 + v)
        recs.append({'año': año, 'fondo': nombre,
                     'meses': meses_r, 'total_año': acc - 1})
    return recs


def icp_a_hist(icp: dict) -> list:
    """Convierte {(año,mes): icp} a lista de registros."""
    años = sorted(set(a for a, _ in icp.keys()))
    recs = []
    for año in años:
        meses_r = [icp.get((año, m)) for m in range(1, 13)]
        vals    = [v for v in meses_r if v is not None]
        if not vals:
            continue
        acc = 1.0
        for v in vals:
            acc *= (1 + v)
        recs.append({'año': año, 'fondo': 'ICP',
                     'meses': meses_r, 'total_año': acc - 1})
    return recs


def resumen_fondo(hist: list, nombre: str) -> dict:
    """Calcula mensual, trimestral, semestral, anual, acum YTD del último mes cerrado."""
    hoy   = date.today()
    mes_c = hoy.month - 1 or 12
    año_c = hoy.year if hoy.month > 1 else hoy.year - 1

    reg = next((r for r in hist if r['año'] == año_c), None)
    if not reg:
        return {'mensual': None, 'trimestral': None,
                'semestral': None, 'anual': None, 'acum_ytd': None}

    meses = reg['meses']

    def prod_n(n):
        vs = [meses[mes_c - 1 - i] for i in range(n)
              if 0 <= mes_c - 1 - i < 12 and meses[mes_c - 1 - i] is not None]
        if not vs: return None
        acc = 1.0
        for v in vs: acc *= (1 + v)
        return acc - 1

    # Anual: producto de los últimos 12 meses cruzando año anterior
    vs12 = []
    for i in range(12):
        m = mes_c - i; a = año_c
        if m <= 0: m += 12; a -= 1
        r = next((r for r in hist if r['año'] == a), None)
        if r and r['meses'][m - 1] is not None:
            vs12.append(r['meses'][m - 1])
    acc12 = 1.0
    for v in vs12: acc12 *= (1 + v)

    return {
        'mensual':    prod_n(1),
        'trimestral': prod_n(3),
        'semestral':  prod_n(6),
        'anual':      acc12 - 1 if vs12 else None,
        'acum_ytd':   reg['total_año'],
    }


def armar_datos(cfg, eom_fip, icp_dict, comp_eom, cartera) -> dict:
    """Ensambla el dict completo para un fondo."""
    nombre  = cfg['nombre_fondo']
    es_usd  = cfg.get('moneda', 'CLP').upper() == 'USD'
    eom_c   = comp_eom[1] if es_usd else comp_eom[0]

    h_fip  = eom_a_hist(eom_fip,  nombre)
    h_icp  = icp_a_hist(icp_dict)
    h_comp = eom_a_hist(eom_c,    'Competencia')

    años = sorted(set(
        [r['año'] for r in h_icp] +
        [r['año'] for r in h_comp] +
        [r['año'] for r in h_fip]
    ))

    historico = []
    for año in años:
        for h, fn in [(h_icp, 'ICP'), (h_comp, 'Competencia'), (h_fip, nombre)]:
            r = next((x for x in h if x['año'] == año), None)
            if r:
                historico.append(r)

    resumen = {
        'ICP (Benchmark)': resumen_fondo(h_icp,  'ICP'),
        'Competencia':      resumen_fondo(h_comp, 'Competencia'),
        nombre:             resumen_fondo(h_fip,  nombre),
    }

    comp = dict(cartera) if cartera else {}
    if not comp.get('moneda'):
        comp['moneda'] = {'USD': 1.0} if es_usd else {'Pesos': 1.0}

    return {
        'nombre_fondo': nombre,
        'tipo_moneda':  cfg.get('moneda', 'CLP'),
        'comentario':   '',
        'resumen':      resumen,
        'historico':    historico,
        'composicion':  comp,
    }


# ─── 6. Generar PPT y PDF ──────────────────────────────────────────────────────

def generar_pptx(datos_json, cfg_json, output_pptx):
    r = subprocess.run(
        ['node', str(GENJS), str(datos_json), str(cfg_json), str(output_pptx)],
        capture_output=True, text=True
    )
    if r.returncode != 0:
        raise RuntimeError(r.stderr.strip() or r.stdout.strip())


def convertir_pdf(pptx_path, output_dir):
    r = subprocess.run(
        ['soffice', '--headless', '--convert-to', 'pdf',
         '--outdir', str(output_dir), str(pptx_path)],
        capture_output=True, text=True
    )
    if r.returncode != 0:
        raise RuntimeError(r.stderr.strip())
    pdf = output_dir / (pptx_path.stem + '.pdf')
    if not pdf.exists():
        raise FileNotFoundError(f"PDF no generado: {pdf}")
    return pdf


def crear_zip(carpeta_mes):
    zip_path = FOLLETOS_DIR / 'latest.zip'
    pdfs = sorted(carpeta_mes.glob('*.pdf'))
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for pdf in pdfs:
            zf.write(pdf, pdf.name)
    log(f"ZIP: {len(pdfs)} PDFs — {zip_path.stat().st_size // 1024} KB")
    return zip_path


# ─── Main ──────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 3:
        print("Uso: python main.py '<comentario_clp>' '<comentario_usd>'")
        sys.exit(1)

    com_clp = sys.argv[1].strip()
    com_usd = sys.argv[2].strip()

    mes_str     = datetime.now().strftime('%Y-%m')
    carpeta_mes = FOLLETOS_DIR / mes_str
    carpeta_mes.mkdir(parents=True, exist_ok=True)
    tmp = Path('/tmp/folletos_vantrust')
    tmp.mkdir(exist_ok=True)

    # ── Cargar todas las fuentes de datos ──────────────────────────────────────
    log("=" * 55)
    log("CARGANDO FUENTES DE DATOS")
    log("=" * 55)

    log("1/4 Valor cuota (API interna)...")
    vc_todos = get_valor_cuota()

    log("2/4 ICP (mindicador.cl)...")
    icp_dict = get_icp()

    log("3/4 Competencia (CMF)...")
    comp_eom = get_competencia()

    log("4/4 Cartera (scripts/inputs/cartera.xlsx)...")
    cartera = get_cartera()

    log("=" * 55)
    log(f"GENERANDO FOLLETOS — {datetime.now().strftime('%B %Y').upper()}")
    log("=" * 55)

    generados = 0
    errores   = []
    vistos    = set()

    for cfg_file, moneda in FONDOS:
        if cfg_file in vistos:
            continue
        vistos.add(cfg_file)

        cfg_path = CFG_DIR / cfg_file
        if not cfg_path.exists():
            log(f"⚠  Sin config: {cfg_file}")
            continue

        with open(cfg_path, encoding='utf-8') as f:
            cfg = json.load(f)

        if cfg.get('rut', '').startswith('COMPLETAR'):
            log(f"⚠  Sin RUT: {cfg['nombre_fondo']} — saltando")
            continue

        nombre = cfg['nombre_fondo']

        # Buscar valor cuota para este fondo
        eom_fip = vc_todos.get(nombre)
        if not eom_fip:
            # Búsqueda flexible por palabras clave
            partes = [p for p in nombre.upper().split()
                      if len(p) > 3 and p not in ('FONDO', 'VANTRUST', 'FIP')]
            for key, val in vc_todos.items():
                key_up = key.upper()
                if sum(1 for p in partes if p in key_up) >= max(1, len(partes) - 1):
                    eom_fip = val
                    log(f"  Mapeado '{nombre}' → '{key}'")
                    break

        if not eom_fip:
            log(f"⚠  Sin valor cuota: {nombre} — saltando")
            continue

        log(f"→  {nombre} ({len(eom_fip)} meses EOM)")

        try:
            datos = armar_datos(cfg, eom_fip, icp_dict, comp_eom, cartera)
            datos['comentario'] = com_usd if moneda == 'usd' else com_clp

            slug        = nombre.replace(' ', '_').replace('/', '-').replace('ó', 'o').replace('é', 'e')
            datos_json  = tmp / f"{slug}.json"
            output_pptx = tmp / f"{slug}.pptx"

            with open(datos_json, 'w', encoding='utf-8') as f:
                json.dump(datos, f, ensure_ascii=False, default=str)

            generar_pptx(datos_json, cfg_path, output_pptx)
            pdf = convertir_pdf(output_pptx, carpeta_mes)
            log(f"✓  {pdf.name}")
            generados += 1

        except Exception as e:
            log(f"✗  {nombre}: {e}")
            errores.append((nombre, str(e)))

    log("=" * 55)
    if generados == 0:
        log("ERROR: Sin folletos generados. Abortando.")
        sys.exit(1)

    crear_zip(carpeta_mes)

    if errores:
        log(f"⚠  {len(errores)} error(es):")
        for n, e in errores:
            log(f"   {n}: {e}")
    else:
        log(f"✅ {generados} folletos generados en folletos/{mes_str}/")


if __name__ == '__main__':
    main()
