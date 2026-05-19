"""
main.py — Genera los 24 folletos mensuales de Vantrust.
"""
import sys, os, zipfile, base64, requests
from datetime import date, timedelta
from pathlib import Path
import pandas as pd

from config import (FONDOS_CON_FOLLETO, OUTPUT_DIR, GITHUB_TOKEN,
                    GITHUB_REPO, GITHUB_BRANCH, get_info_fondo, MESES_ES)
from etl.datos_manager   import (cargar_historico, calcular_datos_fondo,
                                  get_icp_nivel_serie,
                                  get_competencia_clp, get_competencia_usd)
from etl.excel_reader    import get_cartera_composicion
from generador.pptx_builder import generar_pptx
from generador.pdf_exporter import pptx_a_pdf


def _fecha_ref():
    hoy = date.today()
    return hoy.replace(day=1) - timedelta(days=1)

def _periodo_es(fd):
    return f"{MESES_ES[fd.month]} {fd.year}"

def _gh_put(path, repo_path, msg):
    h = {"Authorization": f"Bearer {GITHUB_TOKEN}",
         "Accept": "application/vnd.github+json"}
    api = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{repo_path}"
    r   = requests.get(api, headers=h, timeout=30)
    sha = r.json().get("sha") if r.status_code == 200 else None
    with open(path, "rb") as f:
        content = base64.b64encode(f.read()).decode()
    payload = {"message": msg, "content": content, "branch": GITHUB_BRANCH}
    if sha: payload["sha"] = sha
    r = requests.put(api, headers=h, json=payload, timeout=60)
    print(f"  {'[OK]' if r.status_code in (200,201) else '[WARN]'} {repo_path}")


def run(comentario_clp: str, comentario_usd: str):
    fd      = _fecha_ref()
    mes_str = fd.strftime("%Y-%m")
    periodo = _periodo_es(fd)
    fd_ts   = pd.Timestamp(fd)

    print(f"\n{'='*60}\n FOLLETOS {periodo} — {fd}\n{'='*60}\n")

    print("[1/4] Cargando datos globales...")
    historico = cargar_historico()
    print(f"  Histórico: {len(historico)} fondos cargados del JSON")

    # Obtener ICP y competencia (solo se usan si la API tiene un mes más nuevo)
    print("  ICP (mindicador.cl)...")
    try:
        icp_serie = get_icp_nivel_serie()
        print(f"    {len(icp_serie)} meses ({icp_serie.index[0].year}-{icp_serie.index[-1].year})")
    except Exception as e:
        print(f"    [WARN] ICP no disponible: {e}")
        icp_serie = pd.Series(dtype=float)

    print("  Competencia CLP (CMF)...")
    try:
        comp_clp = get_competencia_clp()
    except Exception as e:
        print(f"    [WARN] Comp CLP no disponible: {e}")
        comp_clp = pd.Series(dtype=float)

    print("  Competencia USD (CMF)...")
    try:
        comp_usd = get_competencia_usd()
    except Exception as e:
        print(f"    [WARN] Comp USD no disponible: {e}")
        comp_usd = pd.Series(dtype=float)

    pptx_dir = OUTPUT_DIR / mes_str / "pptx"
    pdf_dir  = OUTPUT_DIR / mes_str
    pptx_dir.mkdir(parents=True, exist_ok=True)

    print(f"\n[2/4] Generando {len(FONDOS_CON_FOLLETO)} folletos...\n")
    pdf_paths, errores = [], []

    for nombre_fondo in FONDOS_CON_FOLLETO:
        try:
            es_usd     = any(x in nombre_fondo.upper() for x in ("DOLAR","USD"))
            moneda     = "USD" if es_usd else "CLP"
            comentario = comentario_usd if es_usd else comentario_clp
            comp_serie = comp_usd if es_usd else comp_clp

            datos = calcular_datos_fondo(
                nombre_fondo = nombre_fondo,
                historico    = historico,
                icp_serie    = icp_serie,
                comp_serie   = comp_serie,
                fecha_fin    = fd_ts,
            )

            comp_cartera = get_cartera_composicion(nombre_fondo)
            info         = get_info_fondo(nombre_fondo, moneda, fd_ts)

            pptx_path = pptx_dir / f"{nombre_fondo.replace(' ','_')}.pptx"
            generar_pptx(
                nombre_fondo   = nombre_fondo,
                periodo_str    = periodo,
                comentario_pm  = comentario,
                datos_template = datos,
                comp_cartera   = comp_cartera,
                info_fondo     = info,
                out_path       = pptx_path,
            )
            pdf_path = pptx_a_pdf(pptx_path, pdf_dir)
            pdf_paths.append(pdf_path)
            print(f"  [OK] {nombre_fondo}")

        except Exception as e:
            print(f"  [ERROR] {nombre_fondo}: {e}")
            errores.append((nombre_fondo, str(e)))

    if not pdf_paths:
        print("\nERROR: Sin folletos. Abortando.")
        sys.exit(1)

    print(f"\n[3/4] ZIP ({len(pdf_paths)} PDFs)...")
    zip_mes    = pdf_dir    / f"folletos_{mes_str}.zip"
    zip_latest = OUTPUT_DIR / "latest.zip"
    for zp in (zip_mes, zip_latest):
        with zipfile.ZipFile(zp, "w", zipfile.ZIP_DEFLATED) as zf:
            for pdf in pdf_paths:
                zf.write(pdf, arcname=pdf.name)

    if GITHUB_TOKEN:
        print("[4/4] Subiendo a GitHub...")
        _gh_put(zip_latest, "folletos/latest.zip",       f"latest.zip {mes_str}")
        _gh_put(zip_mes, f"folletos/{mes_str}/folletos_{mes_str}.zip", f"ZIP {mes_str}")
        for pdf in pdf_paths:
            _gh_put(pdf, f"folletos/{mes_str}/{pdf.name}", f"{pdf.stem} {mes_str}")

    print(f"\n{'='*60} {periodo}: {len(pdf_paths)} OK", end="")
    if errores:
        print(f", {len(errores)} errores:")
        for n, e in errores: print(f"   ✗ {n}: {e}")
    else:
        print(" ✓")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print('Uso: python main.py "Comentario CLP" "Comentario USD"')
        sys.exit(1)
    run(sys.argv[1], sys.argv[2])
