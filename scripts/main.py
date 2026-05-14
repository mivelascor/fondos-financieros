"""main.py — Orquestador principal."""
import sys, os, zipfile, base64, requests
from datetime import date, timedelta
from pathlib import Path
import pandas as pd

from config import FONDOS_CON_FOLLETO, OUTPUT_DIR, GITHUB_TOKEN, GITHUB_REPO, GITHUB_BRANCH, get_info_fondo
from etl.sql_extractor import get_valores_cuota_eom
from etl.icp_bcch      import get_icp_eom
from etl.cmf_scraper   import get_competencia_clp, get_competencia_usd
from generador.pptx_builder import generar_pptx
from generador.pdf_exporter import pptx_a_pdf


def _fecha_referencia():
    hoy = date.today()
    return hoy.replace(day=1) - timedelta(days=1)


def _github_put(file_path, repo_path, msg):
    h = {"Authorization": f"Bearer {GITHUB_TOKEN}", "Accept": "application/vnd.github+json"}
    api = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{repo_path}"
    r = requests.get(api, headers=h, timeout=30)
    sha = r.json().get("sha") if r.status_code == 200 else None
    with open(file_path, "rb") as f:
        content = base64.b64encode(f.read()).decode()
    payload = {"message": msg, "content": content, "branch": GITHUB_BRANCH}
    if sha:
        payload["sha"] = sha
    r = requests.put(api, headers=h, json=payload, timeout=60)
    print(f"  {'[OK]' if r.status_code in (200,201) else '[WARN]'} {repo_path}")


def run(comentario_clp: str, comentario_usd: str):
    fd = _fecha_referencia()
    mes_str = fd.strftime("%Y-%m")
    periodo = fd.strftime("%B %Y").capitalize()

    print(f"\n{'='*60}")
    print(f" GENERANDO FOLLETOS — {periodo}")
    print(f" Fecha referencia: {fd}")
    print(f"{'='*60}\n")

    # ── 1. Datos ──────────────────────────────────────────────────────────────
    print("[1/5] Obteniendo datos...")

    print("  → Valor cuota (API SQL)")
    df_vc = get_valores_cuota_eom()

    print("  → ICP (mindicador.cl)")
    df_icp = get_icp_eom()
    print(f"      {len(df_icp)} meses")

    print("  → Competencia CMF")
    df_comp_clp = get_competencia_clp()
    df_comp_usd = get_competencia_usd()
    print(f"      CLP:{len(df_comp_clp)} USD:{len(df_comp_usd)}")

    # Series de ICP y competencia (diarias para cálculo)
    icp_serie = pd.Series(
        df_icp["icp"].values,
        index=pd.DatetimeIndex(df_icp["fecha"])
    ).sort_index()

    comp_clp = (pd.Series(df_comp_clp["valor_cuota"].values,
                          index=pd.DatetimeIndex(df_comp_clp["fecha"])).sort_index()
                if not df_comp_clp.empty else pd.Series(dtype=float))
    comp_usd = (pd.Series(df_comp_usd["valor_cuota"].values,
                          index=pd.DatetimeIndex(df_comp_usd["fecha"])).sort_index()
                if not df_comp_usd.empty else pd.Series(dtype=float))

    # ── 2. Generar folletos ───────────────────────────────────────────────────
    pptx_dir = OUTPUT_DIR / mes_str / "pptx"
    pdf_dir  = OUTPUT_DIR / mes_str
    pptx_dir.mkdir(parents=True, exist_ok=True)
    pdf_dir.mkdir(parents=True, exist_ok=True)

    print(f"\n[2/5] Generando {len(FONDOS_CON_FOLLETO)} folletos...")
    pdf_paths, errores = [], []

    for nombre_fondo in FONDOS_CON_FOLLETO:
        try:
            print(f"\n  ▶ {nombre_fondo}")
            es_usd     = any(x in nombre_fondo.upper() for x in ("DOLAR","USD"))
            moneda     = "USD" if es_usd else "CLP"
            comentario = comentario_usd if es_usd else comentario_clp
            comp_serie = comp_usd if es_usd else comp_clp

            # Serie diaria del fondo
            df_f = df_vc[df_vc["fondo"] == nombre_fondo].sort_values("fecha")
            if df_f.empty:
                raise ValueError("Sin datos en VALORES_CUOTA_GPI")

            serie_vc = pd.Series(
                df_f["valor_cuota"].values,
                index=pd.DatetimeIndex(df_f["fecha"])
            ).sort_index()

            fecha_inicio = pd.Timestamp(serie_vc.index[0])
            info = get_info_fondo(nombre_fondo, moneda, fecha_inicio.strftime("%B %Y").capitalize())

            pptx_path = pptx_dir / f"{nombre_fondo.replace(' ','_')}.pptx"

            generar_pptx(
                nombre_fondo  = nombre_fondo,
                periodo_str   = periodo,
                comentario_pm = comentario,
                serie_vc_fondo = serie_vc,
                serie_vc_comp  = comp_serie,
                serie_icp      = icp_serie,
                info_fondo     = info,
                out_path       = pptx_path,
            )
            print("    ✓ PPTX")

            pdf_path = pptx_a_pdf(pptx_path, pdf_dir)
            pdf_paths.append(pdf_path)
            print(f"    ✓ PDF: {pdf_path.name}")

        except Exception as e:
            print(f"    ✗ ERROR: {e}")
            errores.append((nombre_fondo, str(e)))

    if not pdf_paths:
        print("\n❌ Sin folletos generados. Abortando.")
        sys.exit(1)

    # ── 3. ZIP ────────────────────────────────────────────────────────────────
    print(f"\n[3/5] ZIP ({len(pdf_paths)} PDFs)...")
    zip_mes    = pdf_dir    / f"folletos_{mes_str}.zip"
    zip_latest = OUTPUT_DIR / "latest.zip"
    for zp in (zip_mes, zip_latest):
        with zipfile.ZipFile(zp, "w", zipfile.ZIP_DEFLATED) as zf:
            for pdf in pdf_paths:
                zf.write(pdf, arcname=pdf.name)

    # ── 4. GitHub ─────────────────────────────────────────────────────────────
    if GITHUB_TOKEN:
        print(f"\n[4/5] Subiendo a GitHub...")
        _github_put(zip_latest, "folletos/latest.zip", f"latest.zip {mes_str}")
        _github_put(zip_mes, f"folletos/{mes_str}/folletos_{mes_str}.zip", f"ZIP {mes_str}")
        for pdf in pdf_paths:
            _github_put(pdf, f"folletos/{mes_str}/{pdf.name}", f"{pdf.stem} {mes_str}")

    # ── Resumen ───────────────────────────────────────────────────────────────
    print(f"\n{'='*60}")
    print(f" {periodo}: {len(pdf_paths)} folletos OK", end="")
    if errores:
        print(f", {len(errores)} errores:")
        for n, e in errores:
            print(f"   ✗ {n}: {e}")
    else:
        print(" ✓")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print('Uso: python main.py "Comentario CLP" "Comentario USD"')
        sys.exit(1)
    run(sys.argv[1], sys.argv[2])
