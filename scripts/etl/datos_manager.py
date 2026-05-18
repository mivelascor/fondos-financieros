"""
etl/datos_manager.py

Gestiona la fusión de datos históricos (JSON) con datos actuales (APIs).

Flujo mensual:
  1. Cargar historico_fondos.json (histórico desde inicio de cada fondo)
  2. Obtener datos nuevos: API SQL (FIP), mindicador (ICP), CMF scraping (Comp)
  3. Extender el histórico con los meses nuevos
  4. Calcular rentabilidades mensuales y los 5 indicadores del resumen
  5. Guardar el JSON actualizado en el repo (para el mes siguiente)

El JSON nunca necesita actualizarse manualmente: se actualiza solo cada mes.
"""
import json
import os
import requests
import pandas as pd
import numpy as np
from datetime import date, timedelta
from pathlib import Path

from calculos.rentabilidades import (DIVIDENDOS, _eom_niveles,
                                      calcular_5_indicadores, _fmt)

HISTORICO_PATH = Path(__file__).parent.parent.parent / "inputs" / "historico_fondos.json"
MINDICADOR     = "https://mindicador.cl/api/tpm"
API_SQL        = "https://claudeods.vantrustcapital.cl/query"
HEADERS_SQL    = {"Content-Type": "application/json"}


# ── CMF: histórico de competencia CLP y USD ────────────────────────────────
# Estos valores se usan como fallback si el scraping falla
VC_CLP_HIST = {
    "2018-01-31":5000.0,"2018-12-31":5150.0,
    "2019-01-31":5155.0,"2019-12-31":5300.0,
    "2020-01-31":5305.0,"2020-12-31":5350.0,
    "2021-01-31":5352.0,"2021-12-31":5380.0,
    "2022-01-31":5384.0,"2022-12-31":5780.0,
    "2023-01-31":5838.0,"2023-12-31":6226.0,
    "2024-01-31":6275.0,"2024-12-31":6861.4923,
    "2025-01-31":6087.6762,"2025-02-28":6104.4912,
    "2025-03-31":6123.1882,"2025-04-30":6141.4105,
    "2025-05-31":6160.3860,"2025-06-30":6178.8490,
    "2025-07-31":6197.0800,"2025-08-31":6215.6746,
    "2025-09-30":6232.8399,"2025-10-31":6250.6982,
    "2025-11-30":6268.1690,"2025-12-31":6285.9566,
    "2026-01-31":6303.0823,"2026-02-28":6318.4947,
    "2026-03-31":6335.4774,"2026-04-30":6351.9305,
}

VC_USD_HIST = {
    "2024-12-31":22.3823,
    "2025-01-31":22.5132,"2025-02-28":22.5845,
    "2025-03-31":22.6604,"2025-04-30":22.7370,
    "2025-05-31":22.8205,"2025-06-30":22.9101,
    "2025-07-31":23.0004,"2025-08-31":23.0921,
    "2025-09-30":23.1889,"2025-10-31":23.2843,
    "2025-11-30":23.3831,"2025-12-31":23.5091,
    "2026-01-31":23.5903,"2026-02-28":23.6693,
    "2026-03-31":23.7633,"2026-04-30":23.8567,
}

CMF_URL_CLP = ("https://www.cmfchile.cl/institucional/mercados/entidad.php"
               "?mercado=V&rut=8057&grupo=&tipoentidad=RGFMU"
               "&row=AAAw%20cAAhAAAACcAAs&vig=VI&control=svs&pestania=7")
CMF_URL_USD = ("https://www.cmfchile.cl/institucional/mercados/entidad.php"
               "?mercado=V&rut=8248&grupo=&tipoentidad=RGFMU"
               "&row=AAAw%20cAAhAAAACfAAj&vig=VI&control=svs&pestania=7")


def _scrape_cmf(url: str) -> tuple[float | None, str | None]:
    """Scrape el último valor cuota de la CMF. Retorna (valor, fecha_str) o (None, None)."""
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page    = browser.new_page()
            page.goto(url, timeout=30000)
            page.wait_for_timeout(3000)
            # Buscar la tabla de valores cuota
            rows = page.query_selector_all("table tr")
            ultimo_val, ultima_fecha = None, None
            for row in rows:
                cells = row.query_selector_all("td")
                if len(cells) >= 2:
                    try:
                        fecha = cells[0].inner_text().strip()
                        valor = float(cells[1].inner_text().strip().replace(",", "."))
                        if valor > 0:
                            ultimo_val, ultima_fecha = valor, fecha
                    except Exception:
                        continue
            browser.close()
            return ultimo_val, ultima_fecha
    except Exception as e:
        print(f"    [WARN] CMF scraping fallido: {e}")
        return None, None


def _serie_desde_dict(d: dict) -> pd.Series:
    """Crea una pd.Series con DatetimeIndex garantizado desde un dict fecha→valor."""
    s = pd.Series(d, dtype=float)
    s.index = pd.to_datetime(s.index)
    s.index = pd.DatetimeIndex(s.index)
    return s.sort_index()


def get_competencia_clp() -> pd.Series:
    """Retorna serie de valores cuota EOM de la competencia CLP."""
    val, fecha = _scrape_cmf(CMF_URL_CLP)
    hist = dict(VC_CLP_HIST)
    if val and fecha:
        try:
            f   = pd.to_datetime(fecha, dayfirst=True)
            key = f.strftime("%Y-%m-%d")
            hist[key] = val
            print(f"    CMF CLP: {key} = {val:.4f}")
        except Exception:
            pass
    return _serie_desde_dict(hist)


def get_competencia_usd() -> pd.Series:
    """Retorna serie de valores cuota EOM de la competencia USD."""
    val, fecha = _scrape_cmf(CMF_URL_USD)
    hist = dict(VC_USD_HIST)
    if val and fecha:
        try:
            f   = pd.to_datetime(fecha, dayfirst=True)
            key = f.strftime("%Y-%m-%d")
            hist[key] = val
            print(f"    CMF USD: {key} = {val:.4f}")
        except Exception:
            pass
    return _serie_desde_dict(hist)


def get_icp_nivel_serie() -> pd.Series:
    """
    Descarga TPM desde mindicador.cl y construye el nivel acumulado del ICP.
    Retorna serie con DatetimeIndex EOM y nivel acumulado (base 10000 desde 2009).
    """
    all_data = []
    for y in range(2009, date.today().year + 1):
        try:
            r = requests.get(f"{MINDICADOR}/{y}", timeout=15)
            for item in r.json().get("serie", []):
                all_data.append({
                    "fecha": pd.to_datetime(item["fecha"]).tz_localize(None) if pd.to_datetime(item["fecha"]).tzinfo is None else pd.to_datetime(item["fecha"]).tz_convert(None),
                    "tpm":   float(item["valor"])
                })
        except Exception:
            continue

    if not all_data:
        return pd.Series(dtype=float)

    df = pd.DataFrame(all_data).sort_values("fecha").reset_index(drop=True)
    df["period"] = df["fecha"].dt.to_period("M")
    m = df.groupby("period")["tpm"].mean().reset_index()
    m["rent"] = m["tpm"] / 1200
    nivel = [10000.0]
    for i in range(1, len(m)):
        nivel.append(nivel[-1] * (1 + m.loc[i, "rent"]))
    m["nivel"] = nivel
    m["fecha"] = m["period"].dt.to_timestamp("M")
    return pd.Series(m["nivel"].values,
                     index=pd.DatetimeIndex(m["fecha"])).sort_index()


def get_vc_fondo(nombre_fondo: str, dias: int = 455) -> pd.Series:
    """
    Obtiene valores cuota diarios del fondo desde la API SQL.
    Retorna serie con DatetimeIndex.
    """
    desde = (date.today() - timedelta(days=dias)).strftime("%Y-%m-%d")
    sql = f"""
        SELECT FECHA_CIERRE AS fecha,
               RTRIM(LTRIM(EMPRESA)) AS fondo,
               VALOR_CUOTA AS valor_cuota
        FROM ODS.VALORES_CUOTA_GPI
        WHERE FECHA_CIERRE >= '{desde}'
          AND RTRIM(LTRIM(EMPRESA)) = '{nombre_fondo}'
          AND VALOR_CUOTA > 0
        ORDER BY FECHA_CIERRE ASC
    """
    try:
        r = requests.post(API_SQL, json={"Sql": sql}, headers=HEADERS_SQL, timeout=60)
        r.raise_for_status()
        data = r.json()
        rows = data if isinstance(data, list) else data.get("rows", data.get("data", []))
        if not rows:
            return pd.Series(dtype=float)
        df = pd.DataFrame(rows)
        df.columns = [c.lower() for c in df.columns]
        for col in ("fecha_cierre","fecha"):
            if col in df.columns:
                df = df.rename(columns={col: "fecha"})
                break
        for col in ("valor_cuota","precio","valor"):
            if col in df.columns:
                df = df.rename(columns={col: "valor_cuota"})
                break
        df["fecha"]       = pd.to_datetime(df["fecha"])
        df["valor_cuota"] = pd.to_numeric(df["valor_cuota"], errors="coerce")
        df = df.dropna(subset=["fecha","valor_cuota"])
        return pd.Series(df["valor_cuota"].values,
                         index=pd.DatetimeIndex(df["fecha"])).sort_index()
    except Exception as e:
        print(f"    [WARN] API SQL ({nombre_fondo}): {e}")
        return pd.Series(dtype=float)


def cargar_historico() -> dict:
    """Carga el JSON histórico desde el repo."""
    if HISTORICO_PATH.exists():
        with open(HISTORICO_PATH, encoding="utf-8") as f:
            return json.load(f)
    return {}


def guardar_historico(historico: dict):
    """Guarda el JSON histórico actualizado en el repo."""
    HISTORICO_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(HISTORICO_PATH, "w", encoding="utf-8") as f:
        json.dump(historico, f, ensure_ascii=False, separators=(",",":"))


def actualizar_y_calcular(
    nombre_fondo:   str,
    historico:      dict,
    icp_serie:      pd.Series,
    comp_serie:     pd.Series,
    fecha_fin:      pd.Timestamp,
) -> dict:
    """
    Combina el histórico JSON con los datos nuevos de las APIs.
    Retorna un dict con todo lo necesario para el folleto:
      resumen:   lista de 3 dicts {nombre, m, t, s, a, ac, es_fip, ...}
      historico: lista de {año, filas:[{nombre, meses, total}]}
      grafico:   {labels, icp, comp, fip}
      nombre_fip, acum_label
    """
    div          = DIVIDENDOS.get(nombre_fondo, 0.0)
    hist_fondo   = historico.get(nombre_fondo, {})
    nombre_fip   = hist_fondo.get("nombre_fip", nombre_fondo.replace("FIP VANTRUST ","FIP "))
    acum_label   = hist_fondo.get("acum_label", f"Acum. {fecha_fin.year} (*)")

    # ── Construir series de niveles EOM para los 3 indicadores ─────────────
    # Garantizar DatetimeIndex en todas las series de entrada
    def _ensure_dti(s: pd.Series) -> pd.Series:
        if s.empty: return s
        s = s.copy()
        s.index = pd.DatetimeIndex(pd.to_datetime(s.index))
        return s.sort_index()

    # ICP: ya viene como niveles desde get_icp_nivel_serie()
    nivel_icp  = _ensure_dti(icp_serie)
    nivel_icp  = nivel_icp[nivel_icp.index <= fecha_fin]

    # Competencia: nivel = VC_EOM (valores cuota ya son mensuales EOM)
    nivel_comp = _eom_niveles(_ensure_dti(comp_serie), 0.0)
    nivel_comp = nivel_comp[nivel_comp.index <= fecha_fin]

    # FIP: nivel = VC_EOM + dividendos. Obtener VC del API SQL.
    serie_vc_fip = get_vc_fondo(nombre_fondo)
    nivel_fip    = _eom_niveles(_ensure_dti(serie_vc_fip), div)
    nivel_fip    = nivel_fip[nivel_fip.index <= fecha_fin]

    # Si la API SQL no tiene datos del FIP, intentar reconstruir desde los niveles del JSON
    if nivel_fip.empty:
        niv_json = hist_fondo.get("niveles", [])
        if niv_json:
            fip_niv = [(pd.Timestamp(n["fecha"]), n["fip"])
                       for n in niv_json if n.get("fip")]
            if fip_niv:
                serie_niv = pd.Series(
                    [v for _, v in fip_niv],
                    index=pd.DatetimeIndex([f for f, _ in fip_niv])
                )
                nivel_fip = _eom_niveles(serie_niv - div, div)
                nivel_fip = nivel_fip[nivel_fip.index <= fecha_fin]

    # ── Calcular resumen ───────────────────────────────────────────────────
    es_usd   = any(x in nombre_fondo.upper() for x in ("DOLAR", "USD"))
    r_icp    = calcular_5_indicadores(nivel_icp,  fecha_fin) if not es_usd else None
    r_comp   = calcular_5_indicadores(nivel_comp, fecha_fin)
    r_fip    = calcular_5_indicadores(nivel_fip,  fecha_fin)

    resumen = []
    if r_icp:
        resumen.append({"nombre":"ICP (Benchmark)", **r_icp, "es_icp":True, "es_comp":False, "es_fip":False})
    resumen.append({"nombre":"Competencia", **r_comp, "es_icp":False, "es_comp":True, "es_fip":False})
    resumen.append({"nombre":nombre_fip, **r_fip, "es_icp":False, "es_comp":False, "es_fip":True})

    # ── Construir tabla histórica ──────────────────────────────────────────
    # Rentabilidades mensuales desde los niveles
    def rent_monthly(niveles: pd.Series) -> dict:
        """Retorna dict {(año,mes): rent_float}"""
        if niveles.empty or len(niveles) < 2:
            return {}
        r = niveles.pct_change().dropna()
        return {(ts.year, ts.month): float(v) for ts, v in r.items()}

    ri = rent_monthly(nivel_icp)  if not es_usd else {}
    rc = rent_monthly(nivel_comp)
    rf = rent_monthly(nivel_fip)

    # Combinar con histórico del JSON (para años no cubiertos por las APIs)
    hist_json = hist_fondo.get("historico", {})

    def completar(rent_api: dict, tipo: str) -> dict:
        """Añade meses del JSON que no están en la API."""
        result = dict(rent_api)
        for año_str, data in hist_json.items():
            año = int(año_str)
            if tipo in data:
                meses_json = data[tipo].get("meses", [])
                for m_idx, v in enumerate(meses_json):
                    key = (año, m_idx + 1)
                    if key not in result and v is not None:
                        result[key] = float(v) / 100  # el JSON guarda en %
        return result

    ri = completar(ri, "icp")
    rc = completar(rc, "comp")
    rf = completar(rf, "fip")

    # Construir lista de años
    todos_años = sorted(set(k[0] for k in list(ri) + list(rc) + list(rf)))
    anio_inicio_fip = min((k[0] for k in rf), default=fecha_fin.year)

    historico_list = []
    for año in todos_años:
        if año < anio_inicio_fip:
            continue
        if año > fecha_fin.year:
            continue

        def serie_año(rents, nombre_serie, es_icp_=False, es_comp_=False):
            meses = [rents.get((año, m)) for m in range(1, 13)]
            if not any(v is not None for v in meses):
                return None
            # Total año = product(meses) - 1
            acc, hay = 1.0, False
            for v in meses:
                if v is not None:
                    acc *= (1 + v); hay = True
            total = acc - 1 if hay else None

            # Para el año actual, el total es el acum YTD del JSON
            if año == fecha_fin.year:
                # Usar el total calculado desde la API
                total_calc = acc - 1 if hay else None
                # Si está en el JSON del resumen, usarlo
                for row_res in hist_fondo.get("resumen", []):
                    chk = row_res.get("es_icp") if es_icp_ else row_res.get("es_comp") if es_comp_ else row_res.get("es_fip")
                    if chk and row_res.get("ac") is not None:
                        total = row_res["ac"]
                        break
                else:
                    total = total_calc

            return {"nombre": nombre_serie, "meses": meses, "total": total}

        filas = []
        icp_f   = serie_año(ri, "ICP",          es_icp_=True)
        comp_f  = serie_año(rc, "Competencia",   es_comp_=True)
        fip_f   = serie_año(rf, nombre_fip,      es_icp_=False, es_comp_=False)

        if icp_f:  filas.append(icp_f)
        if comp_f: filas.append(comp_f)
        if fip_f:  filas.append(fip_f)

        if filas:
            historico_list.append({"año": año, "filas": filas})

    # ── Gráfico base 100 ──────────────────────────────────────────────────
    # Usar los niveles para construir el gráfico desde inicio del FIP
    fecha_inicio_fip = nivel_fip.index[0] if not nivel_fip.empty else fecha_fin
    grafico = {"labels": [], "icp": [], "comp": [], "fip": []}

    # Combinar todas las fechas disponibles
    fechas_all = sorted(set(nivel_icp.index) | set(nivel_comp.index) | set(nivel_fip.index))
    fechas_all = [f for f in fechas_all if f >= fecha_inicio_fip]

    if fechas_all:
        v0_icp  = nivel_icp.get(fechas_all[0])  or (nivel_icp[nivel_icp.index >= fechas_all[0]].iloc[0] if not nivel_icp.empty else None)
        v0_comp = nivel_comp.get(fechas_all[0]) or (nivel_comp[nivel_comp.index >= fechas_all[0]].iloc[0] if not nivel_comp.empty else None)
        v0_fip  = nivel_fip.get(fechas_all[0])  or (nivel_fip[nivel_fip.index >= fechas_all[0]].iloc[0] if not nivel_fip.empty else None)

        for f in fechas_all:
            # Solo EOM (filtrar fechas que están en el índice de al menos una serie)
            icp_v  = nivel_icp.get(f)  if f in nivel_icp.index  else None
            comp_v = nivel_comp.get(f) if f in nivel_comp.index else None
            fip_v  = nivel_fip.get(f)  if f in nivel_fip.index  else None

            grafico["labels"].append(f.strftime("%b %Y"))
            grafico["icp"].append(round(icp_v / v0_icp * 100, 2)   if icp_v  and v0_icp  else None)
            grafico["comp"].append(round(comp_v / v0_comp * 100, 2) if comp_v and v0_comp else None)
            grafico["fip"].append(round(fip_v  / v0_fip  * 100, 2) if fip_v  and v0_fip  else None)

    return {
        "acum_label":  acum_label,
        "nombre_fip":  nombre_fip,
        "resumen":     resumen,
        "historico":   historico_list,
        "grafico":     grafico,
    }
