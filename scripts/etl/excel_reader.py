"""etl/excel_reader.py — Lee cartera.xlsx subida via admin.html."""
import openpyxl
from config import ARCHIVO_CARTERA

TRAMOS_ORDEN = ["Hasta 30 días", "Entre 31 y 90 días", "Entre 91 y 120 días", "Entre 1 y 2 años"]

def get_cartera_composicion(nombre_fondo: str) -> dict:
    wb = openpyxl.load_workbook(str(ARCHIVO_CARTERA), read_only=True, data_only=True)
    result = {"moneda": [], "duracion": [], "instrumento": []}
    nombre_upper = nombre_fondo[:20].upper()

    def find_col(hdr, nu):
        return next((j for j, h in enumerate(hdr) if h and nu in str(h).upper()), -1)

    # Moneda
    ws_m = wb['moneda']
    rows_m = list(ws_m.iter_rows(values_only=True))
    hdr = rows_m[1]; col_f = find_col(hdr, nombre_upper)
    col_t = next((j for j,h in enumerate(hdr) if h and str(h).upper()=="TOTAL"), -1)
    if col_f >= 0 and col_t >= 0:
        for row in rows_m[2:]:
            mon, monto, total = row[1], row[col_f], row[col_t]
            if mon and mon not in ("","total") and isinstance(monto,(int,float)) and isinstance(total,(int,float)) and total > 0:
                pct = monto / total
                if pct > 0.001:
                    result["moneda"].append((str(mon), f"{pct*100:.2f}%".replace(".",",")))

    # Duración (segundo bloque con %)
    ws_d = wb['duracion']
    rows_d = list(ws_d.iter_rows(values_only=True))
    hdr_d_idx = next((i for i,r in enumerate(rows_d) if r[1]=="TRAMO" and i>5), None)
    if hdr_d_idx is not None:
        hdr_d = rows_d[hdr_d_idx]; col_fd = find_col(hdr_d, nombre_upper)
        if col_fd >= 0:
            tv = {str(r[1]): r[col_fd] for r in rows_d[hdr_d_idx+1:]
                  if r[1] and r[1] not in (0,"total") and isinstance(r[col_fd],(int,float))}
            for tramo in TRAMOS_ORDEN:
                result["duracion"].append((tramo, f"{tv.get(tramo,0)*100:.2f}%".replace(".",",")))

    # Instrumento (segundo bloque con %)
    ws_i = wb['instrumento']
    rows_i = list(ws_i.iter_rows(values_only=True))
    hdr_i_idx = next((i for i,r in enumerate(rows_i) if r[1]=="CLASIFICACION_OBS" and i>5), None)
    if hdr_i_idx is not None:
        hdr_i = rows_i[hdr_i_idx]; col_fi = find_col(hdr_i, nombre_upper)
        if col_fi >= 0:
            for row in rows_i[hdr_i_idx+1:]:
                instr, pct = row[1], row[col_fi]
                if instr and instr!="total" and isinstance(pct,(int,float)) and pct>0.001:
                    result["instrumento"].append((str(instr), f"{pct*100:.2f}%".replace(".",",")))

    wb.close()
    return result
