"""
main.py — Orquestador principal del sistema de folletos Vantrust.
"""
import sys, os, zipfile, base64, requests
from datetime import date, timedelta
from pathlib import Path
import pandas as pd

from config import (FONDOS_CON_FOLLETO, OUTPUT_DIR, GITHUB_TOKEN,
                    GITHUB_REPO, GITHUB_BRANCH, get_info_fondo)
from etl.sql_extractor  import get_valores_cuota_eom
from etl.excel_reader   import get_cartera_composicion
from etl.icp_bcch       import get_icp_eom
from etl.cmf_scraper    import get_competencia_clp, get_competencia_usd, update_historico
from calculos.rentabilidades import (calcular_rent_mensual, normalizar_b1000,
                                     construir_tabla_rentabilidades,
                                     construir_tabla_historica, DIVIDENDOS)
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
    fd      = _fecha_referencia()
    mes_str = fd.strftime("%Y-%m")
    periodo = fd.strftime("%B %Y").capitalize()
    fd_ts   = pd.Timestamp(fd)

    print(f"\n{'='*60}")
    print(f" GENERANDO FOLLETOS - {periodo}")
    print(f" Fecha referencia: {fd}")
    print(f"{'='*60}\n")

    print("[1/5] Obteniendo datos...")
    print("  -> Valor cuota (API SQL)")
    df_vc = get_valores_cuota_eom()

    print("  -> ICP (mindicador.cl)")
    df_icp = get_icp_eom()
    icp_serie = pd.Series(df_icp["icp"].values,
                          index=pd.DatetimeIndex(df_icp["fecha"])).sort_index()
    print(f"      {len(df_icp)} meses de ICP")

    print("  -> Competencia CMF")
    df_comp_clp, val_clp_nuevo, fecha_clp_nueva = get_competencia_clp()
    df_comp_usd, val_usd_nuevo, fecha_usd_nueva = get_competencia_usd()
    print(f"      CLP: {len(df_comp_clp)} | USD: {len(df_comp_usd)}")

    comp_clp_serie = (pd.Series(df_comp_clp["valor_cuota"].values,
                                index=pd.DatetimeIndex(df_comp_clp["fecha"])).sort_index()
                      if not df_comp_clp.empty else pd.Series(dtype=float))
    comp_usd_serie = (pd.Series(df_comp_usd["valor_cuota"].values,
                                index=pd.DatetimeIndex(df_comp_usd["fecha"])).sort_index()
                      if not df_comp_usd.empty else pd.Series(dtype=float))

    if val_clp_nuevo and val_usd_nuevo:
        update_historico(val_clp_nuevo, fecha_clp_nueva, val_usd_nuevo, fecha_usd_nueva)

    pptx_dir = OUTPUT_DIR / mes_str / "pptx"
    pdf_dir  = OUTPUT_DIR / mes_str
    pptx_dir.mkdir(parents=True, exist_ok=True)
    pdf_dir.mkdir(parents=True, exist_ok=True)

    print(f"\n[2/5] Generando {len(FONDOS_CON_FOLLETO)} folletos...")
    pdf_paths, errores = [], []

    for nombre_fondo in FONDOS_CON_FOLLETO:
        try:
            print(f"\n  -> {nombre_fondo}")
            es_usd     = any(x in nombre_fondo.upper() for x in ("DOLAR", "USD"))
            moneda     = "USD" if es_usd else "CLP"
            comentario = comentario_usd if es_usd else comentario_clp
            comp_serie = comp_usd_serie if es_usd else comp_clp_serie

            df_f = df_vc[df_vc["fondo"] == nombre_fondo].sort_values("fecha").copy()
            if df_f.empty:
                raise ValueError("Sin datos en VALORES_CUOTA_GPI")

            serie_vc     = pd.Series(df_f["valor_cuota"].values,
                                     index=pd.DatetimeIndex(df_f["fecha"])).sort_index()
            fecha_inicio = pd.Timestamp(serie_vc.index[0])
            div          = DIVIDENDOS.get(nombre_fondo, 0.0)

            b1000_f = normalizar_b1000(serie_vc,   fecha_inicio)
            b1000_c = (normalizar_b1000(comp_serie, fecha_inicio)
                       if not comp_serie.empty else pd.Series(dtype=float))
            b1000_i = normalizar_b1000(icp_serie,  fecha_inicio)

            tabla_resumen   = construir_tabla_rentabilidades(
                serie_vc, comp_serie, icp_serie, nombre_fondo, fd_ts)
            tabla_historica = construir_tabla_historica(
                serie_vc, comp_serie, icp_serie, nombre_fondo, fd_ts)

            comp_cartera = get_cartera_composicion(nombre_fondo)
            info         = get_info_fondo(nombre_fondo, moneda,
                                          fecha_inicio.strftime("%B %Y").capitalize())

            pptx_path = pptx_dir / f"{nombre_fondo.replace(' ', '_')}.pptx"
            generar_pptx(
                nombre_fondo    = nombre_fondo,
                periodo_str     = periodo,
                comentario_pm   = comentario,
                b1000_fondo     = b1000_f,
                b1000_comp      = b1000_c,
                b1000_icp       = b1000_i,
                tabla_resumen   = tabla_resumen,
                tabla_historica = tabla_historica,
                comp_cartera    = comp_cartera,
                info_fondo      = info,
                out_path        = pptx_path,
            )

            pdf_path = pptx_a_pdf(pptx_path, pdf_dir)
            pdf_paths.append(pdf_path)
            print(f"     OK: {pdf_path.name}")

        except Exception as e:
            print(f"     ERROR: {e}")
            errores.append((nombre_fondo, str(e)))

    if not pdf_paths:
        print("\nERROR: Sin folletos generados. Abortando.")
        sys.exit(1)

    print(f"\n[3/5] Creando ZIP ({len(pdf_paths)} PDFs)...")
    zip_mes    = pdf_dir    / f"folletos_{mes_str}.zip"
    zip_latest = OUTPUT_DIR / "latest.zip"
    for zp in (zip_mes, zip_latest):
        with zipfile.ZipFile(zp, "w", zipfile.ZIP_DEFLATED) as zf:
            for pdf in pdf_paths:
                zf.write(pdf, arcname=pdf.name)

    if GITHUB_TOKEN:
        print(f"\n[4/5] Subiendo a GitHub...")
        _github_put(zip_latest, "folletos/latest.zip", f"latest.zip {mes_str}")
        _github_put(zip_mes, f"folletos/{mes_str}/folletos_{mes_str}.zip", f"ZIP {mes_str}")
        for pdf in pdf_paths:
            _github_put(pdf, f"folletos/{mes_str}/{pdf.name}", f"{pdf.stem} {mes_str}")
        if val_clp_nuevo and val_usd_nuevo:
            scraper_path = Path(__file__).parent / "etl" / "cmf_scraper.py"
            if scraper_path.exists():
                _github_put(scraper_path, "scripts/etl/cmf_scraper.py", f"CMF {mes_str}")

    print(f"\n{'='*60}")
    print(f" {periodo}: {len(pdf_paths)} folletos OK", end="")
    if errores:
        print(f", {len(errores)} errores:")
        for n, e in errores: print(f"   X {n}: {e}")
    else:
        print(" OK")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print('Uso: python main.py "Comentario CLP" "Comentario USD"')
        sys.exit(1)
    run(sys.argv[1], sys.argv[2])
