"""Convierte los datos del panel (actual.json + historial.json) en data.json para la web pública.
Solo deja el nombre de pila de cada conductor y quita el detalle de las anomalías de horas.
Uso: python3 publicar.py ruta/actual.json ruta/historial.json
"""
import json, sys, os

def cargar(p):
    d = json.load(open(p, encoding="utf-8"))
    return d.get("data", d) if isinstance(d, dict) and "data" in d and "f" not in d else d

def pila(n):
    if not n:
        return n
    return ", ".join(x.strip().split(" ")[0].capitalize() for x in n.split(",") if x.strip())

actual, hist = cargar(sys.argv[1]), cargar(sys.argv[2])
for s in actual.get("s", []):
    for lista in (s.get("al") or {}).values():
        for x in lista:
            x["dr"] = pila(x.get("dr", ""))
    s["wa"] = [{"dr": pila(x.get("dr", "")), "x": "Anomalía de horas registrada por Amazon"} for x in s.get("wa", [])]
    filas = []
    for t in (s.get("rt") or "").split(";"):
        if not t:
            continue
        p = t.split("|")
        if len(p) > 5:
            p[5] = pila(p[5])
        filas.append("|".join(p))
    s["rt"] = ";".join(filas)

dst = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data.json")
json.dump({"actual": actual, "historial": hist}, open(dst, "w", encoding="utf-8"), ensure_ascii=False)
print("data.json:", actual.get("f"), actual.get("h"), len(actual.get("s", [])), "sites")
