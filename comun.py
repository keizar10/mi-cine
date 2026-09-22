#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Funciones compartidas por los scripts de mi-cine: acceso al Excel,
caché de identificadores (IMDb <-> TMDB) y cliente de TMDB.
"""
import csv
import os
import re
import sys
import time
import unicodedata

import requests

EXCEL = "Movie Data Base by Keizar.xlsx"
HOJA_PELIS = "Movie Data Base by Keizar"
HOJA_SERIES = "Series"
IDS_CSV = "ids.csv"            # titulo, anio, imdb_id, tmdb_id  (películas)
IDS_SERIES_CSV = "ids_series.csv"

# columnas (1 = A) en la hoja de películas
P = dict(titulo=1, anio=2, decada=3, dur=4, director=5, imdb=6, mi=7, votos=8,
         pais=9, genero=10, subgenero=11, vod=12, sinopsis=13, oscars=14,
         oscar_peli=15, oscar_dir=16, nom=17, fest=18, globo=19, ss=20,
         puntos=21, hist=22, jw=23, top150=24, tt=25)
# columnas en la hoja de series
S = dict(titulo=1, anio=2, fin=3, decada=4, tipo=5, creador=6, imdb=7, mi=8,
         votos=9, pais=10, genero=11, subgenero=12, vod=13, sinopsis=14,
         emmys=15, emmy_serie=16, nom_emmy=17, hist=18, tt=19)

REGION = "ES"
TMDB = "https://api.themoviedb.org/3"


def norm(s):
    if not s:
        return ""
    s = str(s).lower()
    s = "".join(c for c in unicodedata.normalize("NFD", s)
                if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9]", "", s)


def decada(anio):
    try:
        return f"{int(anio) // 10 * 10}s"
    except (TypeError, ValueError):
        return None


# ------------------------------------------------------------------ caché ids
def cargar_ids(ruta):
    """Devuelve lista de dicts y un índice imdb_id -> dict."""
    filas = []
    if os.path.exists(ruta):
        with open(ruta, encoding="utf-8-sig") as f:
            filas = list(csv.DictReader(f))
    for fl in filas:
        fl.setdefault("imdb_id", ""); fl.setdefault("tmdb_id", "")
    return filas


def guardar_ids(ruta, filas):
    with open(ruta, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["titulo", "anio", "imdb_id", "tmdb_id"])
        w.writeheader()
        for fl in filas:
            w.writerow({k: fl.get(k, "") for k in w.fieldnames})


# ------------------------------------------------------------------ TMDB
class Tmdb:
    def __init__(self, pausa=0.05):
        self.key = os.environ.get("TMDB_API_KEY", "").strip()
        if not self.key:
            sys.exit("ERROR: define la variable TMDB_API_KEY.")
        self.s = requests.Session()
        self.pausa = pausa

    def get(self, ruta, **params):
        params["api_key"] = self.key
        for intento in range(3):
            try:
                r = self.s.get(TMDB + ruta, params=params, timeout=20)
            except requests.RequestException:
                time.sleep(1.5); continue
            if r.status_code == 429:
                time.sleep(2); continue
            time.sleep(self.pausa)
            if r.status_code == 200:
                return r.json()
            if r.status_code == 404:
                return {}
            time.sleep(1)
        return None

    def find_por_imdb(self, tt):
        """-> ('movie'|'tv', tmdb_id) o (None, None)"""
        d = self.get(f"/find/{tt}", external_source="imdb_id")
        if not d:
            return None, None
        if d.get("movie_results"):
            return "movie", d["movie_results"][0]["id"]
        if d.get("tv_results"):
            return "tv", d["tv_results"][0]["id"]
        return None, None

    def buscar_tv(self, titulo, anio=None):
        d = self.get("/search/tv", query=titulo, language="es-ES",
                     **({"first_air_date_year": anio} if anio else {}))
        if not d or not d.get("results"):
            return None
        return d["results"][0]["id"]

    def externos(self, tipo, tmdb_id):
        d = self.get(f"/{tipo}/{tmdb_id}/external_ids")
        return (d or {}).get("imdb_id")

    def detalles(self, tipo, tmdb_id):
        return self.get(f"/{tipo}/{tmdb_id}", language="es-ES",
                        append_to_response="credits")

    def proveedores(self, tipo, tmdb_id, tipos=("flatrate", "free", "ads")):
        """Cadena de plataformas en España, '' si ninguna, None si error."""
        d = self.get(f"/{tipo}/{tmdb_id}/watch/providers")
        if d is None:
            return None
        datos = d.get("results", {}).get(REGION)
        if not datos:
            return ""
        nombres = []
        for t in tipos:
            for p in datos.get(t, []):
                n = ABREVIA.get(p.get("provider_name", ""), p.get("provider_name", ""))
                if n and n not in nombres:
                    nombres.append(n)
        return ", ".join(nombres)


ABREVIA = {
    "Netflix": "Netflix", "Netflix Standard with Ads": "Netflix",
    "Amazon Prime Video": "Prime", "Amazon Prime Video with Ads": "Prime", "Prime Video": "Prime",
    "Disney Plus": "Disney+", "Movistar Plus": "Movistar", "Movistar Plus+": "Movistar",
    "Apple TV+": "Apple TV+", "Apple TV Plus": "Apple TV+", "Filmin": "Filmin",
    "HBO Max": "HBO Max", "Max": "HBO Max", "SkyShowtime": "SkyShowtime",
    "Rakuten TV": "Rakuten", "Atres Player": "atresplayer", "RTVE Play": "RTVE Play",
    "Pluto TV": "Pluto TV", "Crunchyroll": "Crunchyroll", "FlixOlé": "FlixOlé",
    "Acontra Plus": "Acontra", "Tivify": "Tivify", "Runtime": "Runtime",
}

# géneros TMDB -> taxonomía del Excel (principal, subgénero)
GENERO_TMDB = {
    "Animation": ("Animación", None), "Documentary": ("Documental", None),
    "Western": ("Western", None), "War": ("Bélico", None), "Horror": ("Terror", None),
    "Science Fiction": ("Ciencia ficción", "Ciencia ficción"), "Fantasy": ("Fantasía", "Fantasía"),
    "Music": ("Musical", "Musical"), "Action": ("Acción", "Acción"),
    "Crime": ("Thriller", "Crimen"), "Mystery": ("Thriller", "Misterio"),
    "Thriller": ("Thriller", "Thriller"), "Comedy": ("Comedia", "Comedia"),
    "Romance": ("Romance", "Romance"), "Adventure": ("Aventura", "Aventura"),
    "Drama": ("Drama", "Drama"), "History": ("Drama", "Histórico"),
    "Family": ("Comedia", "Familiar"), "TV Movie": ("Drama", None),
    # series
    "Action & Adventure": ("Acción", "Aventura"), "Sci-Fi & Fantasy": ("Ciencia ficción", "Fantasía"),
    "War & Politics": ("Bélico", None), "Kids": ("Animación", "Familiar"),
    "Reality": ("Documental", None), "Talk": ("Documental", None), "News": ("Documental", None),
    "Soap": ("Drama", "Romance"),
}
PRIORIDAD = ["Animación", "Documental", "Western", "Bélico", "Terror", "Ciencia ficción",
             "Fantasía", "Musical", "Acción", "Thriller", "Comedia", "Romance", "Aventura", "Drama"]


def clasificar(generos_tmdb):
    """Lista de nombres de género TMDB -> (Género, Subgénero)."""
    pares = [GENERO_TMDB.get(g["name"] if isinstance(g, dict) else g) for g in generos_tmdb]
    pares = [p for p in pares if p]
    if not pares:
        return None, None
    principal = min((p[0] for p in pares), key=lambda g: PRIORIDAD.index(g) if g in PRIORIDAD else 99)
    subs = [p[1] for p in pares if p[1] and p[1] != principal]
    return principal, (subs[0] if subs else None)


PAIS = {
    "US": "EE.UU.", "GB": "Reino Unido", "FR": "Francia", "ES": "España", "IT": "Italia",
    "DE": "Alemania", "JP": "Japón", "KR": "Corea del Sur", "CN": "China", "HK": "Hong Kong",
    "IN": "India", "MX": "México", "AR": "Argentina", "BR": "Brasil", "CA": "Canadá",
    "AU": "Australia", "NZ": "Nueva Zelanda", "SE": "Suecia", "DK": "Dinamarca", "NO": "Noruega",
    "FI": "Finlandia", "IR": "Irán", "RU": "Rusia", "PL": "Polonia", "BE": "Bélgica",
    "NL": "Países Bajos", "IE": "Irlanda", "AT": "Austria", "CH": "Suiza", "TW": "Taiwán",
    "TH": "Tailandia", "TR": "Turquía", "IL": "Israel", "CL": "Chile", "PT": "Portugal",
    "CZ": "Chequia", "HU": "Hungría", "GR": "Grecia", "ZA": "Sudáfrica", "SU": "URSS",
}


def paises(lista):
    out = []
    for p in lista or []:
        c = p.get("iso_3166_1") if isinstance(p, dict) else p
        out.append(PAIS.get(c, (p.get("name") if isinstance(p, dict) else c) or c))
    return ", ".join(out[:3]) or None


def sinopsis_corta(texto, maximo=160):
    """Primera frase del resumen, recortada a `maximo` caracteres."""
    if not texto:
        return None
    t = re.sub(r"\s+", " ", texto).strip()
    m = re.match(r"(.+?[.!?])(\s|$)", t)
    frase = m.group(1) if m else t
    if len(frase) > maximo:
        frase = frase[:maximo - 1].rsplit(" ", 1)[0] + "…"
    return frase
