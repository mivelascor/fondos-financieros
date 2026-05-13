"""
main.py — Orquestador principal del sistema de folletos.

Uso:
    python main.py "Comentario CLP..." "Comentario USD..."

El script:
1. Extrae datos de SQL, BCCh y CMF
2. Calcula rentabilidades (base 1000, mensual, trimestral, semestral, anual, YTD)
3. Genera 24 PPTX con gráficos e imágenes
4. Exporta a PDF con LibreOffice
5. Crea un ZIP con todos los PDFs
6. Sube los PDFs y el ZIP a GitHub
"""
import sys
import os
import zipfile
import base64
import requests
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

# ── Importar módulos propios ──────────────────────────────────────────────────
from config import (
    FONDOS_CON_FOLLETO, OUTPUT_DIR, DATA_DIR,
    GITHUB_TOKEN, GITHUB_REPO, GITHUB_BRANCH,
    get_info_fondo,
)
from etl.sql_extractor   import get_valores_cuota_eom, get_cartera
from etl.icp_extractor   import get_icp_eom
from etl.cmf_scraper     import get_competencia_clp, get_competencia_usd
from calculos.rentabilidades import (
    normalizar_b1000,
    construir_tabla_rentabilidades,
    construir_rentabilidades_mensuales,
)
from generador.graficos import (
    grafico_evolucion,
    tabla_rentabilidades_img,
    grafico_composicion_duracion,
    tabla_comparacion_mensual,
)
from generador.pptx_builder import generar_pptx
from generador.pdf_exporter import pptx_a_pdf


# ── GitHub helpers ────────────────────────────────────────────────────────────
def _github_put(file_path: Path, repo_path: str, mensaje: str):
    """Sube o actualiza un archivo en GitHub via API."""
    headers = {
        "Authorization": f"Bearer {GITHUB_TOKEN}",
        "Accept": "application/vnd.github+json",
    }
    api_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{repo_path}"

    # Obtener SHA si el archivo ya existe
    r = requests.get(api_url, headers=headers, timeout=30)
    sha = r.json().get("sha") if r.status_code == 200 else None

    with open(file_path, "rb") as f:
        content_b64 = base64.b64encode(f.read()).decode()

    payload = {
        "message": mensaje,
        "content": content_b64,
        "branch":  GITHUB_BRANCH,
    }
    if sha:
        payload["sha"] = sha

    r = requests.put(api_url, headers=headers, json=payload, timeout=60)
    if r.status_code not in (200, 201):
        print(f"  [WARN] GitHub PUT falló para {repo_path}: {r.status_code} {r.text[:200]}")
    else:
        print(f"  [OK] Subido: {repo_path}")


# ── Main ──────────────────────────────────────────────────────────────────────
def run(comentario_clp: str, comentario_usd: str):
    hoy       = date.today()
    fecha_fin = hoy.replace(day=1) - timedelta(days=1)   # último día del mes anterior
    mes_str   = fecha_fin.strftime("%Y-%m")
    periodo   = fecha_fin.strftime("%B %Y").capitalize()
    desde_str = (fecha_fin - pd.DateOffset(months=14)).strftime("%Y-%m-%d")
    hasta_str = fecha_fin.strftime("%Y-%m-%d")

    print(f"\n{'='*60}")
    print(f" GENERANDO FOLLETOS — {periodo}")
    print(f"{'='*60}\n")

    # ── Paso 1: Extracción de datos ───────────────────────────────────────────
    print("[1/6] Extrayendo datos...")

    print("  → SQL: valores cuota")
    df_vc = get_valores_cuota_eom()
    print(f"      {len(df_vc)} registros, {df_vc['fondo'].nunique()} fondos")

    print("  → SQL: composición de cartera")
    df_cartera = get_cartera(hasta_str)
    print(f"      {len(df_cartera)} instrumentos")

    print("  → BCCh: ICP")
    df_icp = get_icp_eom(desde_str, hasta_str)
    print(f"      {len(df_icp)} meses de ICP")

    print("  → CMF: competencia CLP")
    df_comp_clp = get_competencia_clp()
    print(f"      {len(df_comp_clp)} registros")

    print("  → CMF: competencia USD")
    df_comp_usd = get_competencia_usd()
    print(f"      {len(df_comp_usd)} registros")

    # Guardar Parquet para trazabilidad
    DATA_DIR.mkdir(exist_ok=True)
    df_vc.to_parquet(DATA_DIR / "valores_cuota.parquet")
    df_icp.to_parquet(DATA_DIR / "icp.parquet")
    df_comp_clp.to_parquet(DATA_DIR / "comp_clp.parquet")
    df_comp_usd.to_parquet(DATA_DIR / "comp_usd.parquet")
    df_cartera.to_parquet(DATA_DIR / "cartera.parquet")

    # ── Paso 2: Generar folletos ──────────────────────────────────────────────
    print(f"\n[2/6] Generando {len(FONDOS_CON_FOLLETO)} folletos...")

    pptx_dir  = OUTPUT_DIR / mes_str / "pptx"
    pdf_dir   = OUTPUT_DIR / mes_str
    pptx_dir.mkdir(parents=True, exist_ok=True)
    pdf_dir.mkdir(parents=True, exist_ok=True)

    pdf_paths  = []
    errores    = []
    fd_ts      = pd.Timestamp(fecha_fin)

    # Preparar series ICP indexadas por fecha
    icp_serie = pd.Series(
        df_icp["icp"].values,
        index=pd.DatetimeIndex(df_icp["fecha"])
    ).sort_index()

    # Series de competencia
    comp_clp_serie = pd.Series(
        df_comp_clp["valor_cuota"].values,
        index=pd.DatetimeIndex(df_comp_clp["fecha"])
    ).sort_index() if not df_comp_clp.empty else pd.Series(dtype=float)

    comp_usd_serie = pd.Series(
        df_comp_usd["valor_cuota"].values,
        index=pd.DatetimeIndex(df_comp_usd["fecha"])
    ).sort_index() if not df_comp_usd.empty else pd.Series(dtype=float)

    for nombre_fondo in FONDOS_CON_FOLLETO:
        try:
            print(f"\n  ▶ {nombre_fondo}")

            # Determinar moneda y comentario
            es_usd = any(x in nombre_fondo.upper() for x in ("DOLAR", "USD"))
            moneda     = "USD" if es_usd else "CLP"
            comentario = comentario_usd if es_usd else comentario_clp
            comp_serie = comp_usd_serie if es_usd else comp_clp_serie

            # Serie del fondo
            df_f = df_vc[df_vc["fondo"] == nombre_fondo].copy()
            if df_f.empty:
                raise ValueError(f"Sin datos de valor cuota en SQL")

            df_f = df_f.sort_values("fecha")
            fecha_inicio = pd.Timestamp(df_f["fecha"].min())
            fondo_serie  = pd.Series(
                df_f["valor_cuota"].values,
                index=pd.DatetimeIndex(df_f["fecha"])
            ).sort_index()

            # Normalizar base 1000 desde fecha de inicio del fondo
            b1000_f = normalizar_b1000(fondo_serie, fecha_inicio)
            b1000_c = normalizar_b1000(comp_serie,  fecha_inicio) if not comp_serie.empty else pd.Series(dtype=float)
            b1000_i = normalizar_b1000(icp_serie,   fecha_inicio)

            # Tabla de rentabilidades
            metricas = construir_tabla_rentabilidades(b1000_f, b1000_c, b1000_i, fd_ts)

            # Rentabilidades mensuales (últimos 12 meses)
            df_rentm = construir_rentabilidades_mensuales(b1000_f, b1000_c, b1000_i, fd_ts, n_meses=12)

            # DataFrame para gráfico de evolución
            df_evo = pd.DataFrame({
                "fecha":        b1000_f.index,
                "b1000_fondo":  b1000_f.values,
            }).merge(
                pd.DataFrame({"fecha": b1000_c.index, "b1000_comp": b1000_c.values}),
                on="fecha", how="outer"
            ).merge(
                pd.DataFrame({"fecha": b1000_i.index, "b1000_icp": b1000_i.values}),
                on="fecha", how="outer"
            ).sort_values("fecha").reset_index(drop=True)

            # Cartera del fondo
            cartera_f = df_cartera[df_cartera["fondo"] == nombre_fondo].copy()

            # ── Generar imágenes ──────────────────────────────────────────────
            img_evo   = grafico_evolucion(df_evo)
            img_rent  = tabla_rentabilidades_img(metricas)
            img_comp  = grafico_composicion_duracion(cartera_f)
            img_tabm  = tabla_comparacion_mensual(df_rentm)

            # Info del fondo
            info = get_info_fondo(nombre_fondo, moneda)
            info["fecha_inicio"] = fecha_inicio.strftime("%B %Y").capitalize()

            # ── Generar PPTX ──────────────────────────────────────────────────
            nombre_archivo = nombre_fondo.replace(" ", "_")
            pptx_path = pptx_dir / f"{nombre_archivo}.pptx"

            generar_pptx(
                nombre_fondo   = nombre_fondo,
                periodo_str    = periodo,
                comentario_pm  = comentario,
                img_evolucion  = img_evo,
                img_tabla_rent = img_rent,
                img_composicion= img_comp,
                img_tabla_comp = img_tabm,
                info_fondo     = info,
                out_path       = pptx_path,
            )
            print(f"    ✓ PPTX generado")

            # ── Exportar a PDF ────────────────────────────────────────────────
            pdf_path = pptx_a_pdf(pptx_path, pdf_dir)
            pdf_paths.append(pdf_path)
            print(f"    ✓ PDF: {pdf_path.name}")

        except Exception as e:
            print(f"    ✗ ERROR: {e}")
            errores.append((nombre_fondo, str(e)))

    # ── Paso 3: Crear ZIP ─────────────────────────────────────────────────────
    print(f"\n[3/6] Creando ZIP con {len(pdf_paths)} PDFs...")
    zip_mes_path    = pdf_dir / f"folletos_{mes_str}.zip"
    zip_latest_path = OUTPUT_DIR / "latest.zip"

    for zip_path in (zip_mes_path, zip_latest_path):
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for pdf in pdf_paths:
                zf.write(pdf, arcname=pdf.name)
    print(f"    ✓ folletos_{mes_str}.zip  ({zip_mes_path.stat().st_size // 1024} KB)")

    # ── Paso 4: Subir a GitHub ────────────────────────────────────────────────
    if GITHUB_TOKEN:
        print(f"\n[4/6] Subiendo a GitHub ({GITHUB_REPO})...")

        # latest.zip
        _github_put(zip_latest_path, "folletos/latest.zip",
                    f"Folletos {mes_str} - latest.zip")

        # ZIP del mes
        _github_put(zip_mes_path, f"folletos/{mes_str}/folletos_{mes_str}.zip",
                    f"Folletos {mes_str} - ZIP")

        # PDFs individuales
        for pdf in pdf_paths:
            _github_put(pdf, f"folletos/{mes_str}/{pdf.name}",
                        f"Folleto {pdf.stem} - {mes_str}")
    else:
        print("\n[4/6] GITHUB_TOKEN no configurado — saltando subida a GitHub")
        print(f"      Archivos disponibles en: {pdf_dir}")

    # ── Resumen ───────────────────────────────────────────────────────────────
    print(f"\n{'='*60}")
    print(f" RESUMEN — {periodo}")
    print(f"  ✓ {len(pdf_paths)} folletos generados")
    if errores:
        print(f"  ✗ {len(errores)} errores:")
        for nombre, err in errores:
            print(f"      - {nombre}: {err}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Uso: python main.py \"Comentario CLP\" \"Comentario USD\"")
        sys.exit(1)

    comentario_clp = sys.argv[1]
    comentario_usd = sys.argv[2]
    run(comentario_clp, comentario_usd)
