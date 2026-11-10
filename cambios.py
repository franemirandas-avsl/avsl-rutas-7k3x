"""Compara la lectura anterior con la nueva y añade los cambios al registro del día.
Uso: python3 cambios.py anterior.json nueva.json cambios_previos.json salida.json
(anterior.json y cambios_previos.json pueden no existir.)
Cada evento: {h, c (site), t (tipo), ru, dr, x (texto)}.
"""
import json, os, sys

VOLTA = {"VES2": ["A", "B", "C"], "VCN1": ["A", "B", "C"], "VAD4": ["A", "E", "B", "C"]}
NOMBRE = {"B": "Retrasada", "R": "En riesgo", "O": "A tiempo", "A": "Adelantada", "C": "Completada", "N": "Sin estado"}


def cargar(p, defecto=None):
    if not p or not os.path.exists(p):
        return defecto
    try:
        d = json.load(open(p, encoding="utf-8"))
    except Exception:
        return defecto
    return d.get("data", d) if isinstance(d, dict) and "data" in d and "f" not in d else d


def rutas(s):
    out = {}
    for t in (s.get("rt") or "").split(";"):
        if not t:
            continue
        p = t.split("|") + [""] * 8
        out[p[0]] = {"st": p[1], "pc": int(p[2] or 0), "p": int(p[3] or 0), "dl": int(p[4] or 0), "dr": p[5], "cy": p[6], "sal": p[7]}
    return out


def ciclo_actual(s, rn):
    c = s["c"]
    if c not in VOLTA:
        return None
    if s.get("cc"):
        return s["cc"]
    orden = VOLTA[c]
    en_marcha = [orden.index(r["cy"]) for r in rn.values() if r["cy"] in orden and (r["pc"] > 0 or r["st"] == "C")]
    return orden[max(en_marcha)] if en_marcha else orden[0]


def mins(h):
    a, b = h.split(":")
    return int(a) * 60 + int(b)


prev, new = cargar(sys.argv[1]), cargar(sys.argv[2])
reg = cargar(sys.argv[3], None)
if not new:
    sys.exit("sin lectura nueva")
if not reg or reg.get("f") != new["f"]:
    reg = {"f": new["f"], "ev": []}
if not prev or prev.get("f") != new["f"] or prev.get("h") == new["h"]:
    json.dump(reg, open(sys.argv[4], "w", encoding="utf-8"), ensure_ascii=False)
    print("sin lectura anterior de hoy: 0 cambios")
    sys.exit(0)

h = new["h"]
ev = []
P = {s["c"]: s for s in prev.get("s", []) if not s.get("err")}
ritmo = []
for s in new.get("s", []):
    if s.get("err") or s["c"] not in P:
        continue
    c, ps = s["c"], P[s["c"]]
    rn, rp = rutas(s), rutas(ps)
    cc = ciclo_actual(s, rn)
    if cc:
        orden = VOLTA[c]
        viejo = {ru for ru, r in {**rp, **rn}.items() if r["cy"] in orden and orden.index(r["cy"]) < orden.index(cc)}
        rn = {k: v for k, v in rn.items() if k not in viejo}
        rp = {k: v for k, v in rp.items() if k not in viejo}
    # rutas nuevas: rescates y salidas
    nuevas = [ru for ru in rn if ru not in rp]
    salen = {}
    for ru in nuevas:
        r = rn[ru]
        if ru.startswith("RDM_"):
            orig = ru[4:]
            ev.append({"h": h, "c": c, "t": "rescate", "ru": ru, "dr": r["dr"], "x": (f"rescate nuevo de {orig}" if orig else "rescate nuevo") + f", {r['p']} paradas"})
        else:
            salen[r["cy"]] = salen.get(r["cy"], 0) + 1
    for cy, n in salen.items():
        ev.append({"h": h, "c": c, "t": "salida", "ru": "", "dr": "", "x": f"salen {n} rutas del ciclo {cy}" if cy else f"{n} rutas nuevas"})
    # cambios de estado
    fin = 0
    for ru, r in rn.items():
        if ru not in rp:
            continue
        a, b = rp[ru]["st"], r["st"]
        if a != "C" and b == "C":
            fin += 1
        elif a == "B" and b in ("O", "A"):
            ev.append({"h": h, "c": c, "t": "mejora", "ru": ru, "dr": r["dr"], "x": f"se recupera: de Retrasada a {NOMBRE[b]} ({r['dl']:+d})"})
        elif a in ("O", "A") and b == "B" and r["dl"] <= -3:
            ev.append({"h": h, "c": c, "t": "empeora", "ru": ru, "dr": r["dr"], "x": f"pasa a Retrasada ({r['dl']:+d})"})
    if fin:
        ev.append({"h": h, "c": c, "t": "fin", "ru": "", "dr": "", "x": f"{fin} ruta{'s' if fin > 1 else ''} terminada{'s' if fin > 1 else ''}"})
    if c in VOLTA:
        for cy in VOLTA[c]:
            tn = [r for r in rutas(s).values() if r["cy"] == cy]
            tp = [r for r in rutas(ps).values() if r["cy"] == cy]
            if tn and all(r["st"] == "C" for r in tn) and not (tp and all(r["st"] == "C" for r in tp)):
                ev.append({"h": h, "c": c, "t": "ciclo", "ru": "", "dr": "", "x": f"ciclo {cy} terminado"})
    # alertas nuevas y resueltas
    an, ap = s.get("al") or {}, ps.get("al") or {}
    if cc:
        an = {k: [x for x in v if x.get("ru") not in viejo] for k, v in an.items()}
        ap = {k: [x for x in v if x.get("ru") not in viejo] for k, v in ap.items()}
    pa_n = {x["ru"]: x for x in an.get("pa", [])}
    pa_p = {x["ru"]: x for x in ap.get("pa", [])}
    for ru, x in pa_n.items():
        if ru not in pa_p:
            ev.append({"h": h, "c": c, "t": "parado", "ru": ru, "dr": x.get("dr", ""), "x": f"{x['m']} min sin actividad, {x['pe']} paradas pendientes"})
    for ru, x in pa_p.items():
        if ru not in pa_n and ru in rn:
            ev.append({"h": h, "c": c, "t": "activo", "ru": ru, "dr": x.get("dr", ""), "x": "vuelve a tener actividad" if rn[ru]["st"] != "C" else "ruta cerrada"})
    jo_p = {x["ru"] for x in ap.get("jo", []) if x.get("ya")}
    for x in an.get("jo", []):
        if x.get("ya") and x["ru"] not in jo_p:
            ev.append({"h": h, "c": c, "t": "jornada", "ru": x["ru"], "dr": x.get("dr", ""), "x": f"fuera de jornada desde las {x['fin']} con {x['pe']} paradas"})
    sa_p = {x["ru"] for x in ap.get("sa", [])}
    for x in an.get("sa", []):
        if x["ru"] not in sa_p:
            ev.append({"h": h, "c": c, "t": "salida_tarde", "ru": x["ru"], "dr": x.get("dr", ""), "x": "sin salir" if x.get("ns") else f"salió tarde (prevista {x['sal']})"})
    ritmo.append((c, s["pc"] - ps["pc"]))

dt = mins(h) - mins(prev["h"])
if dt >= 20 and ritmo:
    ritmo.sort(key=lambda x: -x[1])
    ev.append({"h": h, "c": "", "t": "ritmo", "ru": "", "dr": "", "x": f"Paradas hechas desde las {prev['h']}: " + " · ".join(f"{c} {n:+d}" for c, n in ritmo)})

reg["ev"] = (reg.get("ev", []) + ev)[-400:]
json.dump(reg, open(sys.argv[4], "w", encoding="utf-8"), ensure_ascii=False)
print(f"{len(ev)} cambios nuevos, {len(reg['ev'])} en el día")
