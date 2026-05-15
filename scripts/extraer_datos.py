"""
extraer_datos.py
Extrae datos de un template Excel de fondo (TEMPLATE_FONDO_*.xlsx)
y genera un dict/JSON con toda la info para generar el folleto PPT.

Uso:
    python extraer_datos.py TEMPLATE_FONDO_LIQUIDEZ.xlsx
    python extraer_datos.py TEMPLATE_FONDO_LIQUIDEZ.xlsx --output datos.json
"""

import sys
import json
import math
import pandas as pd
import numpy as np
from pathlib import Path


def safe_float(v):
    """Convierte a float o None."""
    if v is None:
        return None
    if isinstance(v, (int, float, np.floating, np.integer)):
        f = float(v)
        return None if math.isnan(f) else f
    return None


def parse_comentario(wb_path):
    df = pd.read_excel(wb_path, sheet_name='Comentario', header=None)
    texto = ''
    for _, row in df.iterrows():
        for cell in row:
            if isinstance(cell, str) and len(cell) > 20:
                texto = cell.strip()
                break
        if texto:
            break
    return texto


def parse_resumen(wb_path):
    """Lee la tabla resumen: ICP, Competencia, FIP con Mensual/Trimestral/Semestral/Anual/Acum."""
    df = pd.read_excel(wb_path, sheet_name='rentabilidad', header=None)
    
    headers = None
    resumen = {}
    
    for i, row in df.iterrows():
        c18 = row.iloc[18] if len(row) > 18 else None
        if isinstance(c18, str) and 'Rentabilidad' in c18:
            headers = [str(row.iloc[j]) if pd.notna(row.iloc[j]) else '' for j in range(18, 24)]
            continue
        if headers and isinstance(c18, str) and len(c18) > 2:
            vals = {}
            # Mensual, Trimestral, Semestral, Anual, Acum
            labels = ['mensual', 'trimestral', 'semestral', 'anual', 'acum_ytd']
            for k, lbl in enumerate(labels):
                idx = 19 + k
                v = safe_float(row.iloc[idx]) if len(row) > idx else None
                vals[lbl] = v
            nombre = c18.strip()
            resumen[nombre] = vals
    return resumen


def parse_historico(wb_path):
    """
    Lee la tabla histórica mes a mes.
    Retorna lista de dicts:
      { año, fondo, meses: [v1..v12], total_año }
    fondo puede ser 'ICP', 'Competencia', o nombre del FIP.
    """
    df = pd.read_excel(wb_path, sheet_name='rentabilidad', header=None)
    
    registros = []
    año_actual = None
    
    for i, row in df.iterrows():
        c2 = row.iloc[2] if len(row) > 2 else None
        c3 = row.iloc[3] if len(row) > 3 else None
        
        if isinstance(c2, str) and c2 == 'Año':
            continue  # fila de headers
        
        # Año nuevo: fila ICP con año en col 2
        if isinstance(c2, (int, float, np.integer, np.floating)) and not pd.isna(c2):
            año_actual = int(c2)
        
        # Si hay nombre de fondo en col 3
        if isinstance(c3, str) and c3 not in ('', 'Fondo') and año_actual is not None:
            meses = []
            for col in range(4, 16):
                v = safe_float(row.iloc[col]) if len(row) > col else None
                meses.append(v)
            total_año = safe_float(row.iloc[16]) if len(row) > 16 else None
            
            registros.append({
                'año': año_actual,
                'fondo': c3.strip(),
                'meses': meses,
                'total_año': total_año,
            })
    
    return registros


def parse_composicion(wb_path):
    """Lee duración, instrumentos, moneda."""
    df = pd.read_excel(wb_path, sheet_name='composición', header=None)
    
    duracion = {}
    instrumentos = {}
    moneda = {}
    
    dur_labels = ['Menos de 1 mes', '1-3 meses', '3-4 meses', 'Más de 4 meses']
    
    for _, row in df.iterrows():
        for col in range(len(row) - 1):
            cell = row.iloc[col]
            next_cell = row.iloc[col + 1]
            if isinstance(cell, str):
                label = cell.strip()
                val = safe_float(next_cell)
                if label in dur_labels and val is not None:
                    duracion[label] = val
                elif label in ('Bono', 'Depósitos', 'Factoring', 'Financiamiento Inmobiliario',
                               'Financiamiento Mesa Moneda Extranjera', 'Garantías',
                               'Nota Estructurada', 'Financiamiento Mesa Renta Fija',
                               'Simultánea', 'Efectos de Comercio', 'Pagaré', 'Facturas') and val is not None:
                    instrumentos[label] = val
                elif label in ('Pesos', 'UF', 'USD', 'Dólar') and val is not None:
                    moneda[label] = val
    
    return {'duracion': duracion, 'instrumentos': instrumentos, 'moneda': moneda}


def inferir_nombre_fondo(wb_path):
    """Intenta leer el nombre del FIP de la hoja Datos ICP (2)."""
    try:
        df = pd.read_excel(wb_path, sheet_name='Datos ICP (2)', header=None)
        # El nombre del fondo suele estar en las primeras filas, col 9 o 10
        for i in range(min(5, len(df))):
            for col in range(8, 15):
                val = df.iloc[i, col] if len(df.columns) > col else None
                if isinstance(val, str) and 'FIP' in val.upper() and len(val) > 5:
                    return val.strip()
                if isinstance(val, str) and 'FONDO' in val.upper() and len(val) > 5:
                    return val.strip()
    except Exception:
        pass
    # Fallback: extraer del nombre del archivo
    fname = Path(wb_path).stem  # e.g. TEMPLATE_FONDO_LIQUIDEZ_SENCILLO
    return fname.replace('TEMPLATE_', '').replace('_', ' ').title()


def determinar_tipo(nombre_fondo):
    """Determina si el fondo es en Pesos o Dólares."""
    n = nombre_fondo.upper()
    if any(x in n for x in ('DOLAR', 'USD', 'DOLLAR')):
        return 'USD'
    return 'CLP'


def extraer_datos(wb_path):
    """Función principal: extrae todos los datos del template."""
    wb_path = str(wb_path)
    
    nombre_fondo = inferir_nombre_fondo(wb_path)
    tipo_moneda = determinar_tipo(nombre_fondo)
    comentario = parse_comentario(wb_path)
    resumen = parse_resumen(wb_path)
    historico = parse_historico(wb_path)
    composicion = parse_composicion(wb_path)
    
    # Detectar nombre FIP real desde resumen
    nombre_fip = None
    for k in resumen:
        if 'FIP' in k.upper() or 'FONDO' in k.upper():
            nombre_fip = k
            break
    
    return {
        'nombre_fondo': nombre_fip or nombre_fondo,
        'tipo_moneda': tipo_moneda,
        'comentario': comentario,
        'resumen': resumen,
        'historico': historico,
        'composicion': composicion,
    }


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Uso: python extraer_datos.py <archivo.xlsx> [--output <salida.json>]")
        sys.exit(1)
    
    wb_path = sys.argv[1]
    
    output_path = None
    if '--output' in sys.argv:
        idx = sys.argv.index('--output')
        output_path = sys.argv[idx + 1]
    
    datos = extraer_datos(wb_path)
    
    json_str = json.dumps(datos, ensure_ascii=False, indent=2, default=str)
    
    if output_path:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(json_str)
        print(f"Datos escritos en {output_path}")
    else:
        print(json_str)
