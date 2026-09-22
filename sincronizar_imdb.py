#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Lee la lista PÚBLICA de puntuaciones de IMDb del usuario y la vuelca al Excel:
  - actualiza "Mi nota" de las películas/series que ya están
  - añade las que faltan (datos de TMDB en español)
  - refresca nota IMDb y votos de las puntuadas
  - resuelve identificadores IMDb/TMDB que falten (series sobre todo)

VARIABLES DE ENTORNO:
    IMDB_USER_ID    identificador del usuario (lo que va tras /user/ en la URL)
    TMDB_API_KEY    clave v3 de TMDB
"""
import copy
import json
import os
import re
import sys
import time
from datetime import date

import requests
import openpyxl
from openpyxl.styles import Font

from comun import (EXCEL, HOJA_PELIS, HOJA_SERIES, IDS_CSV, IDS_SERIES_CSV, P, S,
                   Tmdb, cargar_ids, guardar_ids, decada, clasificar, paises,
                   sinopsis_corta)

USER = os.environ.get("IMDB_USER_ID", "").strip()
if not USER:
    sys.exit("ERROR: define IMDB_USER_ID (p. ej. ur12345678 o p.xxxxxxxx).")

HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"),
    "Accept-Language": "es-ES,es;q=0.9,en;q=0.8",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}
LOG = []


def log(msg):
    print(msg)
    LOG.append(msg)


# ------------------------------------------------------------- IMDb (lectura)
def descargar(url):
    for intento in range(3):
        try:
            r = requests.get(url, headers=HEADERS, timeout=30)
        except requests.RequestException as e:
            time.sleep(3); continue
        if r.status_code == 200:
            return r.text
        log(f"  IMDb respondió {r.status_code} en {url}")
        time.sleep(3)
    return None


def next_data(html):
    m = re.search(r'<script id="__NEXT_DATA__" type="application/json">(.*?)</script>', html, re.S)
    return json.loads(m.group(1)) if m else None


def es_tt(v):
    return isinstance(v, str) and re.fullmatch(r"tt\d{5,10}", v) is not None


def buscar_items(obj, items, camino=()):
    """Recorre el JSON y recoge los nodos que describen un título con nota de usuario."""
    if isinstance(obj, dict):
        if es_tt(obj.get("id")) and ("titleText" in obj or "originalTitleText" in obj):
            items.append(obj)
            return
        for k, v in obj.items():
            buscar_items(v, items, camino + (k,))
    elif isinstance(obj, list):
        for v in obj:
            buscar_items(v, items, camino)


def valor(d, *ruta, default=None):
    for k in ruta:
        if isinstance(d, dict):
            d = d.get(k)
        else:
            return default
    return d if d is not None else default


def nota_usuario(node, raiz):
    """Busca la nota del usuario dentro del nodo o en su contenedor."""
    for cand in (valor(node, "userRating", "value"), valor(node, "userRating", "rating"),
                 valor(raiz, "userRating", "value"), valor(raiz, "rating", "value"),
                 valor(raiz, "userRating"), valor(raiz, "rating")):
        if isinstance(cand, (int, float)) and 1 <= cand <= 10:
            return int(cand)
    return None


def fecha_usuario(node, raiz):
    for cand in (valor(node, "userRating", "date"), valor(raiz, "userRating", "date"),
                 valor(raiz, "createdDate"), valor(raiz, "ratingDate"), valor(raiz, "date")):
        if isinstance(cand, str) and re.match(r"\d{4}-\d{2}-\d{2}", cand):
            return cand[:10]
    return None


def extraer(data):
    """Devuelve lista de dicts: tt, titulo, original, anio, tipo, dur, imdb, votos, nota, fecha."""
    out = []
    contenedores = []

    def walk(obj):
        if isinstance(obj, dict):
            hijo = None
            for k in ("listItem", "title", "node", "item"):
                if isinstance(obj.get(k), dict) and es_tt(obj[k].get("id")):
                    hijo = obj[k]
            if hijo is not None:
                contenedores.append((hijo, obj))
            elif es_tt(obj.get("id")) and "titleText" in obj:
                contenedores.append((obj, obj))
            for v in obj.values():
                walk(v)
        elif isinstance(obj, list):
            for v in obj:
                walk(v)

    walk(data)
    vistos = set()
    for node, raiz in contenedores:
        tt = node["id"]
        if tt in vistos:
            continue
        nota = nota_usuario(node, raiz)
        if nota is None:
            continue
        vistos.add(tt)
        seg = valor(node, "runtime", "seconds")
        out.append(dict(
            tt=tt,
            titulo=valor(node, "titleText", "text"),
            original=valor(node, "originalTitleText", "text"),
            anio=valor(node, "releaseYear", "year"),
            fin=valor(node, "releaseYear", "endYear"),
            tipo=valor(node, "titleType", "id") or valor(node, "titleType", "text"),
            dur=round(seg / 60) if seg else None,
            imdb=valor(node, "ratingsSummary", "aggregateRating"),
            votos=valor(node, "ratingsSummary", "voteCount"),
            nota=nota, fecha=fecha_usuario(node, raiz),
        ))
    return out


def leer_votos_imdb():
    base = f"https://www.imdb.com/user/{USER}/ratings/"
    todos, vistos = [], set()
    total = None
    for pagina in range(1, 60):
        url = base + f"?sort=date_added%2Cdesc&page={pagina}"
        html = descargar(url)
        if not html:
            break
        data = next_data(html)
        if not data:
            log("  no encuentro __NEXT_DATA__ en la página; IMDb ha cambiado el formato.")
            with open("imdb_debug.html", "w", encoding="utf-8") as f:
                f.write(html[:200000])
            break
        if total is None:
            m = re.search(r'"total"\s*:\s*(\d+)', json.dumps(data))
            total = int(m.group(1)) if m else None
        items = extraer(data)
        if pagina == 1 and not items:
            with open("imdb_debug.json", "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False)
            log("  la página no contiene títulos con nota; guardo imdb_debug.json para revisar.")
            break
        nuevos = [i for i in items if i["tt"] not in vistos]
        if not nuevos:
            break
        for i in nuevos:
            vistos.add(i["tt"]); todos.append(i)
        log(f"  página {pagina}: {len(nuevos)} títulos (acumulado {len(todos)}" +
            (f" de {total}" if total else "") + ")")
        if total and len(todos) >= total:
            break
        if len(items) < 25:
            break
        time.sleep(1.0)
    return todos


# ------------------------------------------------------------- Excel (escritura)
def fila_vacia(ws):
    for i in range(ws.max_row, 1, -1):
        if ws.cell(i, 1).value not in (None, ""):
            return i + 1
    return 2


def copiar_estilo(ws, origen, destino, ncols):
    for c in range(1, ncols + 1):
        ws.cell(destino, c)._style = copy.copy(ws.cell(origen, c)._style)
    ws.row_dimensions[destino].height = ws.row_dimensions[origen].height


def anadir_pelicula(ws, tmdb, tt, v, fila_modelo):
    tipo, tid = tmdb.find_por_imdb(tt)
    det = tmdb.detalles("movie", tid) if tipo == "movie" and tid else None
    i = fila_vacia(ws)
    copiar_estilo(ws, fila_modelo, i, P["tt"])
    anio = v["anio"] or (int(det["release_date"][:4]) if det and det.get("release_date") else None)
    directores = ", ".join(c["name"] for c in (det or {}).get("credits", {}).get("crew", [])
                           if c.get("job") == "Director")[:60] or None
    gen, sub = clasificar((det or {}).get("genres", []))
    ws.cell(i, P["titulo"], (det or {}).get("title") or v["titulo"])
    ws.cell(i, P["anio"], anio)
    ws.cell(i, P["decada"], decada(anio))
    ws.cell(i, P["dur"], v["dur"] or (det or {}).get("runtime") or None)
    ws.cell(i, P["director"], directores)
    ws.cell(i, P["imdb"], v["imdb"])
    ws.cell(i, P["mi"], v["nota"])
    ws.cell(i, P["votos"], v["votos"])
    ws.cell(i, P["pais"], paises((det or {}).get("production_countries")))
    ws.cell(i, P["genero"], gen)
    ws.cell(i, P["subgenero"], sub)
    ws.cell(i, P["sinopsis"], sinopsis_corta((det or {}).get("overview")))
    for k in ("oscars", "oscar_peli", "oscar_dir", "nom", "fest", "globo"):
        ws.cell(i, P[k], 0)
    f = ws.cell(fila_modelo, P["puntos"]).value
    if isinstance(f, str) and f.startswith("="):
        ws.cell(i, P["puntos"], re.sub(r"([A-Z]+)%d\b" % fila_modelo, lambda m: m.group(1) + str(i), f))
    ws.cell(i, P["tt"], tt)
    return i, tid


def anadir_serie(ws, tmdb, tt, v, fila_modelo):
    tipo, tid = tmdb.find_por_imdb(tt)
    det = tmdb.detalles("tv", tid) if tipo == "tv" and tid else None
    i = fila_vacia(ws)
    copiar_estilo(ws, fila_modelo, i, S["tt"])
    anio = v["anio"] or (int(det["first_air_date"][:4]) if det and det.get("first_air_date") else None)
    fin = v["fin"] or (int(det["last_air_date"][:4]) if det and det.get("last_air_date") and not det.get("in_production") else None)
    gen, sub = clasificar((det or {}).get("genres", []))
    creadores = ", ".join(c["name"] for c in (det or {}).get("created_by", []))[:60] or None
    mini = (v["tipo"] or "").lower().find("mini") >= 0 or ((det or {}).get("type") == "Miniseries")
    ws.cell(i, S["titulo"], (det or {}).get("name") or v["titulo"])
    ws.cell(i, S["anio"], anio)
    ws.cell(i, S["fin"], fin)
    ws.cell(i, S["decada"], decada(anio))
    ws.cell(i, S["tipo"], "Miniserie" if mini else "Serie")
    ws.cell(i, S["creador"], creadores)
    ws.cell(i, S["imdb"], v["imdb"])
    ws.cell(i, S["mi"], v["nota"])
    ws.cell(i, S["votos"], v["votos"])
    ws.cell(i, S["pais"], paises((det or {}).get("production_countries") or
                                 [{"iso_3166_1": c} for c in (det or {}).get("origin_country", [])]))
    ws.cell(i, S["genero"], gen)
    ws.cell(i, S["subgenero"], sub)
    ws.cell(i, S["sinopsis"], sinopsis_corta((det or {}).get("overview")))
    for k in ("emmys", "emmy_serie", "nom_emmy"):
        ws.cell(i, S[k], 0)
    ws.cell(i, S["tt"], tt)
    return i, tid


def main():
    log(f"== Sincronización IMDb {date.today():%d/%m/%Y} ==")
    votos = leer_votos_imdb()
    log(f"Votos leídos de IMDb: {len(votos)}")
    if not votos:
        sys.exit("No se ha podido leer ningún voto; no toco el Excel.")
    if os.environ.get("SOLO_LEER"):
        for v in votos[:8]:
            log(f"  {v['fecha']}  {v['nota']}/10  {v['titulo']} ({v['anio']}) {v['tt']} [{v['tipo']}] IMDb {v['imdb']} {v['votos']} votos")
        log("Modo SOLO_LEER: no toco el Excel.")
        with open("sincronizacion.log", "a", encoding="utf-8") as f:
            f.write("\n".join(LOG) + "\n\n")
        return

    tmdb = Tmdb()
    wb = openpyxl.load_workbook(EXCEL)
    wp, wsr = wb[HOJA_PELIS], wb[HOJA_SERIES]
    ids_p = cargar_ids(IDS_CSV)
    ids_s = cargar_ids(IDS_SERIES_CSV)
    idx_p = {f["imdb_id"]: f for f in ids_p if f["imdb_id"]}
    idx_s = {f["imdb_id"]: f for f in ids_s if f["imdb_id"]}

    fila_p = {wp.cell(i, P["tt"]).value: i for i in range(2, wp.max_row + 1) if wp.cell(i, P["tt"]).value}
    fila_s = {wsr.cell(i, S["tt"]).value: i for i in range(2, wsr.max_row + 1) if wsr.cell(i, S["tt"]).value}
    modelo_p = max(fila_p.values()) if fila_p else 2
    modelo_s = max(fila_s.values()) if fila_s else 2

    # 1) resolver ids que falten en la hoja de series (título -> TMDB -> IMDb)
    resueltas = 0
    for i in range(2, wsr.max_row + 1):
        t = wsr.cell(i, S["titulo"]).value
        if not t or wsr.cell(i, S["tt"]).value:
            continue
        base = re.sub(r"\s*\(.*\)$", "", str(t))
        tid = tmdb.buscar_tv(base, wsr.cell(i, S["anio"]).value) or tmdb.buscar_tv(base)
        tt = tmdb.externos("tv", tid) if tid else None
        if tt:
            wsr.cell(i, S["tt"], tt); fila_s[tt] = i; resueltas += 1
            ids_s.append(dict(titulo=t, anio=wsr.cell(i, S["anio"]).value, imdb_id=tt, tmdb_id=str(tid)))
            idx_s[tt] = ids_s[-1]
    if resueltas:
        log(f"Series con identificador resuelto vía TMDB: {resueltas}")

    # 2) volcar votos
    nuevas_p, nuevas_s, cambiadas, refrescadas = [], [], [], 0
    for v in votos:
        tt = v["tt"]
        es_serie = (v["tipo"] or "").lower() in ("tvseries", "tvminiseries", "serie de tv", "miniserie de tv", "tv series", "tv mini series")
        if tt in fila_p:
            ws, i, col_mi, col_imdb, col_votos = wp, fila_p[tt], P["mi"], P["imdb"], P["votos"]
        elif tt in fila_s:
            ws, i, col_mi, col_imdb, col_votos = wsr, fila_s[tt], S["mi"], S["imdb"], S["votos"]
        else:
            ws = None
        if ws is None:
            if es_serie:
                i, tid = anadir_serie(wsr, tmdb, tt, v, modelo_s)
                fila_s[tt] = i; nuevas_s.append(wsr.cell(i, 1).value)
                ids_s.append(dict(titulo=wsr.cell(i, 1).value, anio=wsr.cell(i, 2).value, imdb_id=tt, tmdb_id=str(tid or "")))
            else:
                i, tid = anadir_pelicula(wp, tmdb, tt, v, modelo_p)
                fila_p[tt] = i; nuevas_p.append(wp.cell(i, 1).value)
                ids_p.append(dict(titulo=wp.cell(i, 1).value, anio=wp.cell(i, 2).value, imdb_id=tt, tmdb_id=str(tid or "")))
            continue
        actual = ws.cell(i, col_mi).value
        if actual != v["nota"]:
            ws.cell(i, col_mi, v["nota"])
            cambiadas.append(f"{ws.cell(i, 1).value} ({actual} → {v['nota']})")
        if v["imdb"] and ws.cell(i, col_imdb).value != v["imdb"]:
            ws.cell(i, col_imdb, v["imdb"]); refrescadas += 1
        if v["votos"]:
            ws.cell(i, col_votos, v["votos"])

    wb.save(EXCEL)
    guardar_ids(IDS_CSV, ids_p)
    guardar_ids(IDS_SERIES_CSV, ids_s)

    log(f"Películas nuevas añadidas: {len(nuevas_p)}" + (": " + "; ".join(nuevas_p[:15]) if nuevas_p else ""))
    log(f"Series nuevas añadidas: {len(nuevas_s)}" + (": " + "; ".join(nuevas_s[:15]) if nuevas_s else ""))
    log(f"Notas cambiadas: {len(cambiadas)}" + (": " + "; ".join(cambiadas[:15]) if cambiadas else ""))
    log(f"Notas IMDb refrescadas: {refrescadas}")
    with open("sincronizacion.log", "a", encoding="utf-8") as f:
        f.write("\n".join(LOG) + "\n\n")


if __name__ == "__main__":
    main()
