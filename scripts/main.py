"""
main.py — Orquestador principal del sistema de folletos.

Fuentes automáticas:
  - Valor cuota : API REST interna (claudeods.vantrustcapital.cl)
  - ICP         : API BCCh (BCCH_USER / BCCH_PASS en GitHub Secrets)
  - Competencia : scraping CMF

Input manual (subido via admin.html antes de ejecutar):
  - inputs/cartera.xlsx
  - Comentario CLP y USD

La fecha de referencia SIEMPRE es el último día del mes anterior,
independiente del día en que se ejecute.

Uso: python main.py "Comentario CLP" "Comentario USD"
"""
import sys, os, zipfile, base64, requests
from datetime import date, timedelta
from pathlib import Path
import pandas as pd

from config import (FONDOS_CON_FOLLETO, OUTPUT_DIR, GITHUB_TOKEN,
                    GITHUB_REPO, GITHUB_BRANCH, get_info_fondo)
from etl.sql_extractor import get_valores_cuota_eom
from etl.excel_reader  import get_cartera
from etl.icp_bcch      import get_icp_eom
from etl.cmf_scraper   import get_competencia_clp, get_competencia_usd
from calculos.rentabilidades import (normalizar_b1000,
    construir_tabla_rentabilidades, construir_rentabilidades_mensuales)
from generador.graficos import (grafico_evolucion, tabla_rentabilidades_img,
    grafico_composicion_duracion, tabla_comparacion_mensual)
from generador.pptx_builder import generar_pptx
from generador.pdf_exporter import pptx_a_pdf


def _fecha_referencia() -> date:
    """
    Retorna siempre el último día del mes anterior,
    independiente del día de ejecución.
    Ej: si hoy es 3-abr o 25-abr → retorna 31-mar.
    """
    hoy = date.today()
    return hoy.replace(day=1) - timedelta(days=1)


def _github_put(file_path: Path, repo_path: str, msg: str):
    """Sube o actualiza un archivo en GitHub via API."""
    headers = {
        "Authorization": f"Bearer {GITHUB_TOKEN}",
        "Accept": "application/vnd.github+json",
    }
    api_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{repo_path}"
    r = requests.get(api_url, headers=headers, timeout=30)
    sha = r.json().get("sha") if r.status_code == 200 else None
    with open(file_path, "rb") as f:
        content = base64.b64encode(f.read()).decode()
    payload = {"message": msg, "content": content, "branch": GITHUB_BRANCH}
    if sha:
        payload["sha"] = sha
    r = requests.put(api_url, headers=headers, json=payload, timeout=60)
    if r.status_code in (200, 201):
        print(f"  [OK] {repo_path}")
    else:
        print(f"  [WARN] {repo_path}: {r.status_code} {r.text[:100]}")


def run(comentario_clp: str, comentario_usd: str):
    # ── Fecha de referencia: siempre último día del mes anterior ─────────────
    fecha_fin = _fecha_referencia()
    mes_str   = fecha_fin.strftime("%Y-%m")
    periodo   = fecha_fin.strftime("%B %Y").capitalize()
    fd_ts     = pd.Timestamp(fecha_fin)

    print(f"\n{'='*60}")
    print(f" GENERANDO FOLLETOS — {periodo}")
    print(f" Fecha referencia: {fecha_fin} (último día mes anterior)")
    print(f"{'='*60}\n")

    # ── 1. Obtener datos ──────────────────────────────────────────────────────
    print("[1/6] Obteniendo datos...")

    print("  → Valor cuota (API SQL interna)")
    df_vc = get_valores_cuota_eom()

    print("  → Cartera (inputs/cartera.xlsx)")
    df_cartera = get_cartera()
    print(f"      {df_cartera['fondo'].nunique()} fondos, {len(df_cartera)} instrumentos")

    print("  → ICP (API BCCh)")
    df_icp = get_icp_eom()
    print(f"      {len(df_icp)} meses de ICP")

    print("  → Competencia CMF")
    df_comp_clp = get_competencia_clp()
    df_comp_usd = get_competencia_usd()
    print(f"      CLP: {len(df_comp_clp)} | USD: {len(df_comp_usd)}")

    # ── 2. Preparar carpetas ──────────────────────────────────────────────────
    pptx_dir = OUTPUT_DIR / mes_str / "pptx"
    pdf_dir  = OUTPUT_DIR / mes_str
    pptx_dir.mkdir(parents=True, exist_ok=True)
    pdf_dir.mkdir(parents=True, exist_ok=True)

    icp_serie = pd.Series(df_icp["icp"].values,
                          index=pd.DatetimeIndex(df_icp["fecha"])).sort_index()
    comp_clp_serie = (pd.Series(df_comp_clp["valor_cuota"].values,
                                index=pd.DatetimeIndex(df_comp_clp["fecha"])).sort_index()
                      if not df_comp_clp.empty else pd.Series(dtype=float))
    comp_usd_serie = (pd.Series(df_comp_usd["valor_cuota"].values,
                                index=pd.DatetimeIndex(df_comp_usd["fecha"])).sort_index()
                      if not df_comp_usd.empty else pd.Series(dtype=float))

    # ── 3. Generar folletos ───────────────────────────────────────────────────
    print(f"\n[2/6] Generando {len(FONDOS_CON_FOLLETO)} folletos...")
    pdf_paths, errores = [], []

    for nombre_fondo in FONDOS_CON_FOLLETO:
        try:
            print(f"\n  ▶ {nombre_fondo}")
            es_usd     = any(x in nombre_fondo.upper() for x in ("DOLAR", "USD"))
            moneda     = "USD" if es_usd else "CLP"
            comentario = comentario_usd if es_usd else comentario_clp
            comp_serie = comp_usd_serie if es_usd else comp_clp_serie

            # Serie del fondo
            df_f = df_vc[df_vc["fondo"] == nombre_fondo].sort_values("fecha").copy()
            if df_f.empty:
                raise ValueError("Sin datos en VALORES_CUOTA_GPI para este fondo")

            fecha_inicio = pd.Timestamp(df_f["fecha"].min())
            fondo_serie  = pd.Series(df_f["valor_cuota"].values,
                                     index=pd.DatetimeIndex(df_f["fecha"])).sort_index()

            # Normalizar base 1000
            b1000_f = normalizar_b1000(fondo_serie, fecha_inicio)
            b1000_c = (normalizar_b1000(comp_serie, fecha_inicio)
                       if not comp_serie.empty else pd.Series(dtype=float))
            b1000_i = normalizar_b1000(icp_serie,  fecha_inicio)

            # Métricas
            metricas = construir_tabla_rentabilidades(b1000_f, b1000_c, b1000_i, fd_ts)
            df_rentm = construir_rentabilidades_mensuales(b1000_f, b1000_c, b1000_i, fd_ts, 12)

            df_evo = (
                pd.DataFrame({"fecha": b1000_f.index, "b1000_fondo": b1000_f.values})
                .merge(pd.DataFrame({"fecha": b1000_c.index, "b1000_comp": b1000_c.values}),
                       on="fecha", how="outer")
                .merge(pd.DataFrame({"fecha": b1000_i.index, "b1000_icp": b1000_i.values}),
                       on="fecha", how="outer")
                .sort_values("fecha").reset_index(drop=True)
            )

            cartera_f = df_cartera[df_cartera["fondo"] == nombre_fondo].copy()

            # Imágenes
            img_evo  = grafico_evolucion(df_evo)
            img_rent = tabla_rentabilidades_img(metricas)
            img_comp = grafico_composicion_duracion(cartera_f)
            img_tabm = tabla_comparacion_mensual(df_rentm)

            info = get_info_fondo(nombre_fondo, moneda)
            info["fecha_inicio"] = fecha_inicio.strftime("%B %Y").capitalize()

            # PPTX
            nombre_archivo = nombre_fondo.replace(" ", "_")
            pptx_path = pptx_dir / f"{nombre_archivo}.pptx"
            generar_pptx(nombre_fondo, periodo, comentario,
                         img_evo, img_rent, img_comp, img_tabm, info, pptx_path)
            print(f"    ✓ PPTX")

            # PDF
            pdf_path = pptx_a_pdf(pptx_path, pdf_dir)
            pdf_paths.append(pdf_path)
            print(f"    ✓ PDF: {pdf_path.name}")

        except Exception as e:
            print(f"    ✗ ERROR: {e}")
            errores.append((nombre_fondo, str(e)))

    if not pdf_paths:
        print("\n❌ No se generó ningún folleto. Abortando.")
        sys.exit(1)

    # ── 4. ZIP ────────────────────────────────────────────────────────────────
    print(f"\n[3/6] Creando ZIP ({len(pdf_paths)} PDFs)...")
    zip_mes    = pdf_dir    / f"folletos_{mes_str}.zip"
    zip_latest = OUTPUT_DIR / "latest.zip"

    # Solo comprimir los PDFs que se generaron exitosamente
    for zp in (zip_mes, zip_latest):
        with zipfile.ZipFile(zp, "w", zipfile.ZIP_DEFLATED) as zf:
            for pdf in pdf_paths:
                zf.write(pdf, arcname=pdf.name)
    print(f"    ✓ folletos_{mes_str}.zip  ({zip_mes.stat().st_size // 1024} KB)")

    # ── 5. Subir a GitHub automáticamente ─────────────────────────────────────
    if GITHUB_TOKEN:
        print(f"\n[4/6] Subiendo folletos a GitHub (folletos/{mes_str}/)...")

        # ZIP del mes
        _github_put(zip_mes,    f"folletos/{mes_str}/folletos_{mes_str}.zip",
                    f"ZIP folletos {mes_str}")
        # latest.zip (siempre el más reciente)
        _github_put(zip_latest, "folletos/latest.zip",
                    f"latest.zip → {mes_str}")

        # Cada PDF individual (solo los exitosos)
        print(f"    Subiendo {len(pdf_paths)} PDFs individuales...")
        for pdf in pdf_paths:
            _github_put(pdf, f"folletos/{mes_str}/{pdf.name}",
                        f"{pdf.stem} {mes_str}")
    else:
        print(f"\n[4/6] Sin GH_TOKEN — archivos locales en: {pdf_dir}")

    # ── Resumen ───────────────────────────────────────────────────────────────
    print(f"\n{'='*60}")
    print(f" {periodo}: {len(pdf_paths)} folletos OK", end="")
    if errores:
        print(f", {len(errores)} con error:")
        for n, e in errores:
            print(f"   ✗ {n}: {e}")
        print("\n  Los folletos con error NO fueron subidos al repositorio.")
    else:
        print(" ✓ — todos subidos a folletos/" + mes_str + "/")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print('Uso: python main.py "Comentario CLP" "Comentario USD"')
        sys.exit(1)
    run(sys.argv[1], sys.argv[2])
