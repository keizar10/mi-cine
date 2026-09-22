#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Actualiza la columna VOD (plataformas de suscripción/gratis EN ESPAÑA, según
TMDB/JustWatch) en las hojas de películas y series.

Usa el identificador IMDb de cada fila (columna "IMDb ID") y la caché
ids.csv / ids_series.csv para saber el id de TMDB; si falta, lo resuelve
con /find y lo guarda.

    export TMDB_API_KEY="tu_clave"
    python3 actualizar_vod.py
"""
from datetime import date

import openpyxl
from openpyxl.styles import Alignment, Font

from comun import (EXCEL, HOJA_PELIS, HOJA_SERIES, IDS_CSV, IDS_SERIES_CSV, P, S,
                   Tmdb, cargar_ids, guardar_ids)


def procesar(ws, tipo, col_tt, col_vod, ids, tmdb):
    idx = {f["imdb_id"]: f for f in ids if f["imdb_id"]}
    con = sin = sin_id = 0
    for i in range(2, ws.max_row + 1):
        if not ws.cell(i, 1).value:
            continue
        tt = ws.cell(i, col_tt).value
        if not tt:
            sin_id += 1
            continue
        f = idx.get(tt)
        tid = (f or {}).get("tmdb_id") or ""
        if not tid:
            t, tid = tmdb.find_por_imdb(tt)
            if tid and t == tipo:
                if f is None:
                    f = dict(titulo=ws.cell(i, 1).value, anio=ws.cell(i, 2).value, imdb_id=tt, tmdb_id="")
                    ids.append(f); idx[tt] = f
                f["tmdb_id"] = str(tid)
            else:
                sin_id += 1
                continue
        valor = tmdb.proveedores(tipo, tid)
        if valor is None:                       # error de red: no tocar la celda
            continue
        c = ws.cell(i, col_vod)
        c.value = valor if valor else None
        c.font = Font(name="Calibri", size=11)
        c.alignment = Alignment(wrap_text=False, vertical="center")
        if valor:
            con += 1
        else:
            sin += 1
        if (con + sin) % 250 == 0:
            print(f"  {ws.title}: procesadas {con + sin}…")
    ws.cell(1, col_vod).value = f"VOD ({date.today():%d/%m/%Y})"
    return con, sin, sin_id


def main():
    tmdb = Tmdb(pausa=0.06)
    wb = openpyxl.load_workbook(EXCEL)
    ids_p = cargar_ids(IDS_CSV)
    ids_s = cargar_ids(IDS_SERIES_CSV)

    rp = procesar(wb[HOJA_PELIS], "movie", P["tt"], P["vod"], ids_p, tmdb)
    rs = procesar(wb[HOJA_SERIES], "tv", S["tt"], S["vod"], ids_s, tmdb)

    wb.save(EXCEL)
    guardar_ids(IDS_CSV, ids_p)
    guardar_ids(IDS_SERIES_CSV, ids_s)
    print(f"\nPelículas: {rp[0]} con plataforma, {rp[1]} sin, {rp[2]} sin identificador")
    print(f"Series:    {rs[0]} con plataforma, {rs[1]} sin, {rs[2]} sin identificador")
    print(f"Guardado: {EXCEL}")


if __name__ == "__main__":
    main()
