"""Reporte de fin de día AVSL: Excel + resumen.json para el panel.
Uso: python3 informe.py actual.json historial.json cambios.json rutas.txt carpeta_salida
rutas.txt: una fila por ruta site|ruta|conductores|ciclo|salida_plan|salida_real|ultima_parada|paradas|hechas|entregados|devueltos|restantes|intentos|entregados_tras_intento
Genera Reporte_fin_dia_AAAA-MM-DD.xlsx y resumen.json en carpeta_salida.
"""
import json, os, sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule

ORDER = ["VES2", "VCN1", "VAD4", "DCT4", "DCT7", "DCT9", "OCN2", "DMZ4", "DQA4"]
VOLTA = {"VES2": ["A", "B", "C"], "VCN1": ["A", "B", "C"], "VAD4": ["A", "E", "B", "C"]}
RZ = {"CUSTOMER_UNAVAILABLE": "Cliente no disponible", "BUSINESS_CLOSED": "Negocio cerrado", "RESCHEDULED_BY_CUSTOMER": "Reprogramado por el cliente",
      "NO_LOCKER_AVAILABLE": "Sin taquilla disponible", "OBJECT_MISSING": "Paquete no encontrado", "INACCESSIBLE_DELIVERY_LOCATION": "Lugar de entrega inaccesible",
      "BAD_WEATHER": "Mal tiempo", "UNSAFE_DUE_TO_DOG": "Perro, no es seguro", "NO_SECURE_LOCATION": "Sin lugar seguro", "LOCKER_ISSUE": "Incidencia en taquilla",
      "ADDRESS_NOT_FOUND": "Dirección no encontrada", "LOCKER_SPACE_INSUFFICIENT": "Taquilla sin espacio", "OTP_NOT_AVAILABLE": "Cliente sin código OTP",
      "MISSING_OR_INCORRECT_ACCESS_CODE": "Código de acceso erróneo", "SHIPMENT_NOT_READY": "Envío no preparado", "COLLECTED_BY_CUSTOMER": "Ya recogido por el cliente",
      "INACCESSIBLE_PICKUP_LOCATION": "Recogida inaccesible", "PICKED_BY_CUSTOMER": "Ya recogido por el cliente", "NONE": "Sin motivo", "OTRO": "Otro motivo"}
TIPO = {"rescate": "Rescate", "salida": "Salida", "fin": "Rutas terminadas", "ciclo": "Ciclo terminado", "mejora": "Se recupera", "empeora": "Pasa a retrasada",
        "parado": "Parado", "activo": "Vuelve a actividad", "jornada": "Fuera de jornada", "salida_tarde": "Salida tarde", "ritmo": "Ritmo"}


def rzn(k):
    return RZ.get(k, str(k).replace("_", " ").capitalize())


def cargar(p, d=None):
    if not p or not os.path.exists(p):
        return d
    x = json.load(open(p, encoding="utf-8"))
    return x.get("data", x) if isinstance(x, dict) and "data" in x and "f" not in x else x


actual, hist, camb = cargar(sys.argv[1]), cargar(sys.argv[2], {}), cargar(sys.argv[3], {})
out = sys.argv[5]
filas = []
for ln in open(sys.argv[4], encoding="utf-8").read().splitlines():
    p = ln.split("|")
    if len(p) < 14:
        continue
    filas.append({"c": p[0], "ru": p[1], "dr": p[2], "cy": p[3], "sp": p[4], "sr": p[5], "ul": p[6],
                  "p": int(p[7] or 0), "pc": int(p[8] or 0), "e": int(p[9] or 0), "d": int(p[10] or 0), "q": int(p[11] or 0),
                  "at": int(p[12] or 0), "rec": int(p[13] or 0)})
filas.sort(key=lambda r: (ORDER.index(r["c"]) if r["c"] in ORDER else 99, r["cy"], r["ru"]))
fecha = actual["f"]
S = {s["c"]: s for s in actual["s"] if not s.get("err")}
sites = [c for c in ORDER if c in S]
ev = camb.get("ev", []) if camb.get("f") == fecha else []

# ---------- estilo ----------
F = "Arial"
H = Font(name=F, bold=True, color="FFFFFF")
HF = PatternFill("solid", fgColor="232F3E")
TF = PatternFill("solid", fgColor="EDEFF2")
BAND = PatternFill("solid", fgColor="FFF1D6")
B = Font(name=F, bold=True)
N = Font(name=F)
TIT = Font(name=F, bold=True, size=14)
thin = Side(style="thin", color="D5D9DE")
BOR = Border(bottom=thin)
PCT = "0.0%"
wb = Workbook()


def hoja(nombre, titulo, cab, anchos):
    ws = wb.active if wb.active.title == "Sheet" else wb.create_sheet()
    ws.title = nombre
    ws["A1"] = titulo
    ws["A1"].font = TIT
    for i, c in enumerate(cab, 1):
        x = ws.cell(3, i, c)
        x.font, x.fill = H, HF
        x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[3].height = 30
    for i, w in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"
    return ws


def fila(ws, r, vals, fmt=None, font=N, fill=None):
    for i, v in enumerate(vals, 1):
        x = ws.cell(r, i, v)
        x.font, x.border = font, BOR
        if fill:
            x.fill = fill
        if fmt and i in fmt:
            x.number_format = fmt[i]


def semaforo(ws, rango, bajo=0.97, alto=0.99):
    ws.conditional_formatting.add(rango, CellIsRule(operator="lessThan", formula=[str(bajo)], font=Font(name=F, color="C0392B", bold=True)))
    ws.conditional_formatting.add(rango, CellIsRule(operator="greaterThanOrEqual", formula=[str(alto)], font=Font(name=F, color="1E8449", bold=True)))


# ---------- Rutas (base de datos de las demás hojas) ----------
wr = None
cabR = ["Site", "Tipo", "Ciclo", "Ruta", "Conductor", "Salida prevista", "Salida real", "Última parada", "Paradas", "Paradas hechas", "% paradas",
        "Entregados", "Devueltos", "Restantes", "Paquetes", "% entrega", "Intentos fallidos", "Entregados tras intento"]
ultima = {}
for r in filas:
    if r["ul"] and r["ul"] > ultima.get(r["c"], ("", ""))[0]:
        ultima[r["c"]] = (r["ul"], r["ru"])

# ---------- Resumen ----------
ws = hoja("Resumen", f"Reporte de fin de día AVSL · {fecha}", ["Site", "Tipo", "Rutas", "Paradas", "Paradas hechas", "% paradas", "Entregados", "Devueltos",
          "Paquetes", "% entrega", "Última parada", "Ruta que cierra", "Rescates", "Alertas de parado", "Fuera de jornada", "Salida tarde",
          "Intentos fallidos", "Entregados tras intento", "Sin entregar tras intento", "2+ intentos", "Negocio cerrado 13–16 h"],
          [8, 10, 8, 10, 10, 10, 11, 10, 10, 10, 10, 14, 9, 10, 10, 9, 10, 11, 11, 9, 11])
n0 = len(filas) + 3  # última fila de la hoja Rutas
row = 4
cnt = lambda c, t: sum(1 for e in ev if e.get("c") == c and e.get("t") == t)
for c in sites:
    q = S[c].get("q2") or {}
    rng = lambda col: f"Rutas!${col}$4:${col}${n0}"
    k = f'Rutas!$A$4:$A${n0},$A{row}'
    fila(ws, row, [c, "VOLTA" if c in VOLTA else "Next Day",
                   f"=COUNTIFS({k})", f"=SUMIFS({rng('I')},{k})", f"=SUMIFS({rng('J')},{k})", f"=IF(D{row}=0,0,E{row}/D{row})",
                   f"=SUMIFS({rng('L')},{k})", f"=SUMIFS({rng('M')},{k})", f"=SUMIFS({rng('O')},{k})", f"=IF(I{row}=0,0,G{row}/I{row})",
                   ultima.get(c, ("", ""))[0], ultima.get(c, ("", ""))[1], cnt(c, "rescate"), cnt(c, "parado"), cnt(c, "jornada"), cnt(c, "salida_tarde"),
                   q.get("at", 0), q.get("rec", 0), q.get("pend", 0), q.get("dbl", 0), sum((q.get("hb") or {}).get(str(h), (q.get("hb") or {}).get(h, 0)) for h in (13, 14, 15))],
         fmt={6: PCT, 10: PCT})
    row += 1
fila(ws, row, ["Total", "", f"=SUM(C4:C{row-1})", f"=SUM(D4:D{row-1})", f"=SUM(E4:E{row-1})", f"=IF(D{row}=0,0,E{row}/D{row})", f"=SUM(G4:G{row-1})",
               f"=SUM(H4:H{row-1})", f"=SUM(I4:I{row-1})", f"=IF(I{row}=0,0,G{row}/I{row})", "", ""] +
     [f"=SUM({get_column_letter(i)}4:{get_column_letter(i)}{row-1})" for i in range(13, 22)], fmt={6: PCT, 10: PCT}, font=B, fill=TF)
semaforo(ws, f"J4:J{row}")
ws.cell(row + 2, 1, "% entrega = entregados / paquetes; paquetes = entregados + devueltos + restantes que no se devolvieron (mismo criterio que el reporte diario anterior).").font = Font(name=F, italic=True, color="666666")
ws.cell(row + 3, 1, "Rescates, parado, fuera de jornada y salida tarde: veces que se detectó en las lecturas horarias del panel. Intentos y calidad: detalle de cada parada en Entrega.").font = Font(name=F, italic=True, color="666666")
ws.cell(row + 4, 1, "Entregados tras intento: paquetes con un intento fallido que se entregaron más tarde; ya no cuentan como pendientes, pero sí como reintento.").font = Font(name=F, italic=True, color="666666")

# ---------- VOLTA por ciclo ----------
wv = hoja("VOLTA - % por ciclo", "VOLTA · resultado por ciclo", ["Site", "Ciclo", "Rutas", "Paradas", "Paradas hechas", "% paradas", "Entregados", "Devueltos",
          "Paquetes", "% entrega", "Última parada"], [8, 7, 8, 10, 10, 10, 11, 10, 10, 10, 11])
row = 4
for c in sites:
    if c not in VOLTA:
        continue
    for cy in VOLTA[c]:
        rs = [r for r in filas if r["c"] == c and r["cy"] == cy]
        if not rs:
            continue
        k = f'Rutas!$A$4:$A${n0},$A{row},Rutas!$C$4:$C${n0},$B{row}'
        rng = lambda col: f"Rutas!${col}$4:${col}${n0}"
        fila(wv, row, [c, cy, f"=COUNTIFS({k})", f"=SUMIFS({rng('I')},{k})", f"=SUMIFS({rng('J')},{k})", f"=IF(D{row}=0,0,E{row}/D{row})",
                       f"=SUMIFS({rng('L')},{k})", f"=SUMIFS({rng('M')},{k})", f"=SUMIFS({rng('O')},{k})", f"=IF(I{row}=0,0,G{row}/I{row})",
                       max((r["ul"] for r in rs if r["ul"]), default="")], fmt={6: PCT, 10: PCT})
        row += 1
semaforo(wv, f"J4:J{max(row-1,4)}")

# ---------- Rutas ----------
wr = hoja("Rutas", "Detalle por ruta", cabR, [7, 9, 6, 13, 26, 9, 9, 9, 8, 8, 9, 10, 9, 9, 9, 9, 9, 10])
for i, r in enumerate(filas):
    x = i + 4
    fila(wr, x, [r["c"], "VOLTA" if r["c"] in VOLTA else "Next Day", r["cy"] or "-", r["ru"], r["dr"], r["sp"], r["sr"], r["ul"], r["p"], r["pc"],
                 f"=IF(I{x}=0,0,J{x}/I{x})", r["e"], r["d"], r["q"], f"=L{x}+M{x}+MAX(0,N{x}-M{x})", f"=IF(O{x}=0,0,L{x}/O{x})", r["at"], r["rec"]],
         fmt={11: PCT, 16: PCT})
wr.auto_filter.ref = f"A3:R{max(n0,4)}"
semaforo(wr, f"P4:P{max(n0,4)}")

# ---------- Por conductor ----------
wc = hoja("Por conductor", "Resultado por conductor", ["Site", "Conductor", "Rutas", "Paradas", "Paradas hechas", "% paradas", "Entregados", "Paquetes", "% entrega",
          "Intentos fallidos"], [8, 28, 8, 10, 10, 10, 11, 10, 10, 10])
cond = {}
for r in filas:
    cond.setdefault((r["c"], r["dr"]), []).append(r)
row = 4
for (c, dr), rs in sorted(cond.items(), key=lambda kv: (ORDER.index(kv[0][0]) if kv[0][0] in ORDER else 99, kv[0][1])):
    k = f'Rutas!$A$4:$A${n0},$A{row},Rutas!$E$4:$E${n0},$B{row}'
    rng = lambda col: f"Rutas!${col}$4:${col}${n0}"
    fila(wc, row, [c, dr, f"=COUNTIFS({k})", f"=SUMIFS({rng('I')},{k})", f"=SUMIFS({rng('J')},{k})", f"=IF(D{row}=0,0,E{row}/D{row})",
                   f"=SUMIFS({rng('L')},{k})", f"=SUMIFS({rng('O')},{k})", f"=IF(H{row}=0,0,G{row}/H{row})", f"=SUMIFS({rng('Q')},{k})"], fmt={6: PCT, 9: PCT})
    row += 1
wc.auto_filter.ref = f"A3:J{max(row-1,4)}"
semaforo(wc, f"I4:I{max(row-1,4)}")

# ---------- Calidad ----------
wq = hoja("Calidad", "Calidad · motivo de los intentos fallidos", ["Motivo"] + sites + ["Total"], [30] + [8] * len(sites) + [9])
motivos = sorted({k for c in sites for k in ((S[c].get("q2") or {}).get("rz") or {})},
                 key=lambda k: -sum(((S[c].get("q2") or {}).get("rz") or {}).get(k, 0) for c in sites))
row = 4
for m in motivos:
    fila(wq, row, [rzn(m)] + [((S[c].get("q2") or {}).get("rz") or {}).get(m, 0) for c in sites] + [f"=SUM(B{row}:{get_column_letter(len(sites)+1)}{row})"],
         fill=BAND if m == "BUSINESS_CLOSED" else None)
    row += 1
L = get_column_letter(len(sites) + 2)
fila(wq, row, ["Total intentos fallidos"] + [f"=SUM({get_column_letter(i)}4:{get_column_letter(i)}{row-1})" for i in range(2, len(sites) + 3)], font=B, fill=TF)
tot_row = row
row += 1
for et, key in (("Entregados tras intento", "rec"), ("Siguen sin entregar", "pend"), ("Paquetes con 2+ intentos", "dbl")):
    fila(wq, row, [et] + [(S[c].get("q2") or {}).get(key, 0) for c in sites] + [f"=SUM(B{row}:{get_column_letter(len(sites)+1)}{row})"], font=B)
    row += 1
fila(wq, row, ["% recuperado tras intento"] + [f"=IF({get_column_letter(i)}{tot_row}=0,0,{get_column_letter(i)}{tot_row+1}/{get_column_letter(i)}{tot_row})" for i in range(2, len(sites) + 3)],
     fmt={i: PCT for i in range(2, len(sites) + 3)}, font=B)
row += 2

wq.cell(row, 1, "Intentos fallidos por hora (franja 13–16 h resaltada)").font = TIT
row += 1
for i, cab in enumerate(["Hora"] + sites + ["Total", "Negocio cerrado"], 1):
    x = wq.cell(row, i, cab)
    x.font, x.fill = H, HF
wq.column_dimensions[get_column_letter(len(sites) + 3)].width = 12
row += 1
h0 = row
g = lambda q, k, h: (q.get(k) or {}).get(str(h), (q.get(k) or {}).get(h, 0))
for h in range(7, 24):
    vals = [g(S[c].get("q2") or {}, "hr", h) for c in sites]
    bc = sum(g(S[c].get("q2") or {}, "hb", h) for c in sites)
    if not any(vals) and h not in (13, 14, 15):
        continue
    fila(wq, row, [f"{h:02d}:00–{h+1:02d}:00"] + vals + [f"=SUM(B{row}:{get_column_letter(len(sites)+1)}{row})", bc], fill=BAND if 13 <= h < 16 else None)
    row += 1
wq.conditional_formatting.add(f"B{h0}:{get_column_letter(len(sites)+1)}{row-1}", ColorScaleRule(start_type="min", start_color="FFFFFF", end_type="max", end_color="F4A259"))
row += 1

wq.cell(row, 1, "Negocios cerrados entre 13 y 16 h").font = TIT
row += 1
for i, cab in enumerate(["Site", "Negocio cerrado 13–16 h", "Intentos fallidos 13–16 h", "% de la franja"], 1):
    x = wq.cell(row, i, cab)
    x.font, x.fill = H, HF
row += 1
for c in sites:
    q = S[c].get("q2") or {}
    fila(wq, row, [c, sum(g(q, "hb", h) for h in (13, 14, 15)), sum(g(q, "hr", h) for h in (13, 14, 15)), f"=IF(C{row}=0,0,B{row}/C{row})"], fmt={4: PCT})
    row += 1
row += 1

wq.cell(row, 1, "Recogidas fallidas").font = TIT
row += 1
for i, cab in enumerate(["Motivo"] + sites + ["Total"], 1):
    x = wq.cell(row, i, cab)
    x.font, x.fill = H, HF
row += 1
pkl = {}
for c in sites:
    for k, v in ((S[c].get("q2") or {}).get("pk") or {}).items():
        pkl.setdefault(rzn(k), {}).setdefault(c, 0)
        pkl[rzn(k)][c] += v
pks = sorted(pkl, key=lambda m: -sum(pkl[m].values()))
for m in pks:
    fila(wq, row, [m] + [pkl[m].get(c, 0) for c in sites] + [f"=SUM(B{row}:{get_column_letter(len(sites)+1)}{row})"])
    row += 1
if not pks:
    wq.cell(row, 1, "Sin recogidas fallidas").font = N
wq.freeze_panes = "B4"

# ---------- Incidencias ----------
wi = hoja("Incidencias", "Lo que pasó durante el día (lecturas horarias del panel)", ["Hora", "Site", "Tipo", "Ruta", "Conductor", "Detalle"], [7, 7, 18, 14, 26, 70])
row = 4
for e in sorted([e for e in ev if e.get("t") not in ("ritmo", "fin", "salida")], key=lambda e: (e["h"], ORDER.index(e["c"]) if e.get("c") in ORDER else 99)):
    fila(wi, row, [e["h"], e.get("c", ""), TIPO.get(e["t"], e["t"]), e.get("ru", ""), e.get("dr", ""), e.get("x", "")])
    row += 1
wi.auto_filter.ref = f"A3:F{max(row-1,4)}"

# ---------- Evolución ----------
we = hoja("Evolución", "Evolución del día · % de paradas hechas en cada lectura", ["Hora"] + sites, [8] + [9] * len(sites))
row = 4
for pnt in (hist or {}).get("puntos", []):
    vals = []
    for c in sites:
        v = (pnt.get("v") or {}).get(c)
        vals.append(round(v[0] / v[1], 4) if v and v[1] else None)
    fila(we, row, [pnt["h"]] + vals, fmt={i: PCT for i in range(2, len(sites) + 2)})
    row += 1

os.makedirs(out, exist_ok=True)
xlsx = os.path.join(out, f"Reporte_fin_dia_{fecha}.xlsx")
wb.save(xlsx)

# ---------- resumen.json para el panel ----------
res = {"f": fecha, "h": actual.get("h", ""), "s": []}
for c in sites:
    rs = [r for r in filas if r["c"] == c]
    res["s"].append({"c": c, "r": len(rs), "p": sum(r["p"] for r in rs), "pc": sum(r["pc"] for r in rs), "e": sum(r["e"] for r in rs),
                     "d": sum(r["d"] for r in rs), "pq": sum(r["e"] + r["d"] + max(0, r["q"] - r["d"]) for r in rs), "ul": ultima.get(c, ("", ""))[0]})
json.dump(res, open(os.path.join(out, "resumen.json"), "w", encoding="utf-8"), ensure_ascii=False)
print(xlsx, len(filas), "rutas")
