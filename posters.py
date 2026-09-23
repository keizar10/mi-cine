#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Recoge de TMDB la carátula (en español si existe) y el fondo de cada
película/serie y los guarda en posters.csv (imdb_id, poster, fondo).
Solo consulta las que aún no tiene, así la primera vez tarda unos minutos
y las siguientes, segundos.

    export TMDB_API_KEY="tu_clave"
    python3 posters.py
"""
import csv
import os

from comun import IDS_CSV, IDS_SERIES_CSV, Tmdb, cargar_ids

SALIDA = "posters.csv"


def main():
    tmdb = Tmdb(pausa=0.05)
    hechos = {}
    if os.path.exists(SALIDA):
        with open(SALIDA, encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                hechos[r["imdb_id"]] = (r.get("poster", ""), r.get("fondo", ""))

    pendientes = []
    for ruta, tipo in ((IDS_CSV, "movie"), (IDS_SERIES_CSV, "tv")):
        for fl in cargar_ids(ruta):
            if fl["imdb_id"] and fl["tmdb_id"] and fl["imdb_id"] not in hechos:
                pendientes.append((fl["imdb_id"], fl["tmdb_id"], tipo))
    print(f"Carátulas ya guardadas: {len(hechos)}; por consultar: {len(pendientes)}")

    nuevos = 0
    for i, (tt, tid, tipo) in enumerate(pendientes, 1):
        d = tmdb.get(f"/{tipo}/{tid}", language="es-ES")
        if d is None:                      # error de red: lo intentará el mes que viene
            continue
        hechos[tt] = (d.get("poster_path") or "", d.get("backdrop_path") or "")
        nuevos += 1
        if i % 250 == 0:
            print(f"  {i}/{len(pendientes)}…")

    with open(SALIDA, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["imdb_id", "poster", "fondo"])
        for tt, (po, bd) in sorted(hechos.items()):
            w.writerow([tt, po, bd])
    print(f"Guardadas {nuevos} nuevas; total {len(hechos)} en {SALIDA}")


if __name__ == "__main__":
    main()
