"""
Procesa el shapefile de provincias de Ecuador (formato R .RDS, objeto
`SpatialPolygonsDataFrame` de CONALI) SIN necesitar R, rpy2 ni pyreadr,
y genera:

  1. provincias_ecuador.geojson  -> polígonos de las 25 provincias en
     latitud/longitud (WGS84), listos para dash/plotly (go.Choropleth).
  2. Una figura de ejemplo en Plotly coloreando las provincias que
     aparecen en EJE1_INSTITUCIONES, igual a como se usa en app.py.

Cómo ejecutarlo:
    python3 procesar_mapa_provincias.py /ruta/al/PROVINCIA_CONALI_EDIT.RDS

Si no se indica ruta, intenta usar "PROVINCIA_CONALI_EDIT.RDS" en la
carpeta actual.

No requiere ninguna librería externa de geoespacial (no usa pyproj,
shapely, geopandas, sf ni R): incluye su propio lector binario del
formato de serialización RDS de R y su propia conversión UTM -> lat/lon
(fórmulas estándar de Snyder para el elipsoide WGS84).
"""
import gzip
import math
import struct
import sys
import json
import os

# --------------------------------------------------------------------------
# 1) Lector del formato binario RDS de R (XDR, versión 2/3), sin dependencias.
# --------------------------------------------------------------------------
NILSXP = 0
SYMSXP = 1
LISTSXP = 2
LANGSXP = 6
CHARSXP = 9
LGLSXP = 10
INTSXP = 13
REALSXP = 14
CPLXSXP = 15
STRSXP = 16
VECSXP = 19
RAWSXP = 24
S4SXP = 25
REFSXP = 255
NILVALUE_SXP = 254

OBJECT_BIT = 1 << 8
ATTR_BIT = 1 << 9
TAG_BIT = 1 << 10


class RNull:
    def __repr__(self):
        return "RNull"


R_NULL = RNull()


class RObject:
    __slots__ = ("value", "attrs")

    def __init__(self, value, attrs=None):
        self.value = value
        self.attrs = attrs or {}


def _pairlist_to_dict(robj):
    if robj is R_NULL:
        return {}
    if isinstance(robj, RObject) and isinstance(robj.value, tuple) and robj.value[0] == "pairlist":
        return {tag: val for tag, val in robj.value[1]}
    return {}


class RDSReader:
    def __init__(self, data: bytes):
        self.data = data
        self.pos = 0
        # Only symbols (SYMSXP) are registered in R's shared reference
        # table -- ordinary CHARSXP string elements inside a STRSXP are
        # NOT, even if the same string repeats many times.
        self.ref_table = []

    def read(self, n):
        chunk = self.data[self.pos:self.pos + n]
        self.pos += n
        if len(chunk) != n:
            raise EOFError(f"wanted {n} bytes at {self.pos - n}, got {len(chunk)}")
        return chunk

    def read_i32(self):
        return struct.unpack(">i", self.read(4))[0]

    def read_f64(self, n=1):
        return struct.unpack(f">{n}d", self.read(8 * n)) if n else ()

    def read_i32n(self, n):
        return struct.unpack(f">{n}i", self.read(4 * n)) if n else ()

    def read_header(self):
        magic = self.read(2)
        if magic != b"X\n":
            raise ValueError(f"RDS no soportado, magic inesperado {magic!r}")
        version = self.read_i32()
        self.read_i32()
        self.read_i32()
        if version == 3:
            enc_len = self.read_i32()
            self.read(enc_len)
        return version

    def read_object(self):
        flags = self.read_i32()
        return self._read_with_flags(flags)

    def _read_with_flags(self, flags):
        sxp_type = flags & 0xFF
        has_attr = bool(flags & ATTR_BIT)
        has_tag = bool(flags & TAG_BIT)

        if sxp_type == REFSXP:
            idx = flags >> 8
            if idx == 0:
                idx = self.read_i32()
            return self.ref_table[idx - 1]

        if sxp_type in (NILSXP, NILVALUE_SXP):
            return R_NULL

        if sxp_type == CHARSXP:
            return self._read_charsxp_raw()

        if sxp_type == SYMSXP:
            name_flags = self.read_i32()
            name_type = name_flags & 0xFF
            if name_type == REFSXP:
                idx = name_flags >> 8
                if idx == 0:
                    idx = self.read_i32()
                name = self.ref_table[idx - 1]
            elif name_type == CHARSXP:
                name = self._read_charsxp_raw()
            else:
                raise NotImplementedError(f"tipo inesperado de nombre de símbolo {name_type}")
            sym = RObject(("symbol", name))
            self.ref_table.append(sym)
            return sym

        if sxp_type in (LISTSXP, LANGSXP):
            attrs = _pairlist_to_dict(self.read_object()) if has_attr else {}
            tag = None
            if has_tag:
                tag_obj = self.read_object()
                tag = (tag_obj.value[1] if isinstance(tag_obj, RObject)
                       and isinstance(tag_obj.value, tuple) and tag_obj.value[0] == "symbol"
                       else tag_obj)
            car = self.read_object()
            cdr = self.read_object()
            items = [(tag, car)]
            if isinstance(cdr, RObject) and isinstance(cdr.value, tuple) and cdr.value[0] == "pairlist":
                items.extend(cdr.value[1])
            return RObject(("pairlist", items), attrs)

        # Para cualquier otro tipo, R escribe (flags, contenido..., luego
        # los atributos si has_attr) -- los atributos van DESPUÉS del
        # contenido (al revés que en LISTSXP/LANGSXP de arriba).
        if sxp_type == STRSXP:
            n = self.read_i32()
            items = [self.read_object() for _ in range(n)]
            attrs = _pairlist_to_dict(self.read_object()) if has_attr else {}
            return RObject(items, attrs)
        if sxp_type == VECSXP:
            n = self.read_i32()
            items = [self.read_object() for _ in range(n)]
            attrs = _pairlist_to_dict(self.read_object()) if has_attr else {}
            return RObject(items, attrs)
        if sxp_type == INTSXP:
            n = self.read_i32()
            vals = list(self.read_i32n(n))
            NA_INTEGER = -2147483648
            vals = [None if v == NA_INTEGER else v for v in vals]
            attrs = _pairlist_to_dict(self.read_object()) if has_attr else {}
            return RObject(vals, attrs)
        if sxp_type == LGLSXP:
            n = self.read_i32()
            vals = list(self.read_i32n(n))
            vals = [None if v == -2147483648 else bool(v) for v in vals]
            attrs = _pairlist_to_dict(self.read_object()) if has_attr else {}
            return RObject(vals, attrs)
        if sxp_type == REALSXP:
            n = self.read_i32()
            vals = list(self.read_f64(n))
            attrs = _pairlist_to_dict(self.read_object()) if has_attr else {}
            return RObject(vals, attrs)
        if sxp_type == RAWSXP:
            n = self.read_i32()
            raw = self.read(n)
            attrs = _pairlist_to_dict(self.read_object()) if has_attr else {}
            return RObject(raw, attrs)
        if sxp_type == CPLXSXP:
            n = self.read_i32()
            raw = self.read_f64(2 * n)
            vals = [complex(raw[2 * i], raw[2 * i + 1]) for i in range(n)]
            attrs = _pairlist_to_dict(self.read_object()) if has_attr else {}
            return RObject(vals, attrs)
        if sxp_type == S4SXP:
            attrs = _pairlist_to_dict(self.read_object()) if has_attr else {}
            return RObject(None, attrs)
        raise NotImplementedError(f"tipo SEXP {sxp_type} no soportado (byte {self.pos})")

    def _read_charsxp_raw(self):
        length = self.read_i32()
        if length == -1:
            return None
        raw = self.read(length)
        return raw.decode("utf-8", errors="replace")


def load_rds(path):
    with open(path, "rb") as f:
        raw = f.read()
    try:
        raw = gzip.decompress(raw)
    except OSError:
        pass
    reader = RDSReader(raw)
    reader.read_header()
    return reader.read_object()


# --------------------------------------------------------------------------
# 2) Conversión UTM -> latitud/longitud (elipsoide WGS84), sin pyproj.
#    Fórmulas estándar de Snyder (las mismas que usa el paquete "utm").
# --------------------------------------------------------------------------
_K0 = 0.9996
_E = 0.00669438
_E2 = _E * _E
_E3 = _E2 * _E
_E_P2 = _E / (1 - _E)

_SQRT_E = math.sqrt(1 - _E)
_e1 = (1 - _SQRT_E) / (1 + _SQRT_E)
_e1_2, _e1_3, _e1_4, _e1_5 = _e1 ** 2, _e1 ** 3, _e1 ** 4, _e1 ** 5

_M1 = 1 - _E / 4 - 3 * _E2 / 64 - 5 * _E3 / 256
_P2 = 3. / 2 * _e1 - 27. / 32 * _e1_3 + 269. / 512 * _e1_5
_P3 = 21. / 16 * _e1_2 - 55. / 32 * _e1_4
_P4 = 151. / 96 * _e1_3 - 417. / 128 * _e1_5
_P5 = 1097. / 512 * _e1_4

_R = 6378137.0  # radio ecuatorial WGS84


def utm_to_lonlat(easting, northing, zone_number, southern=True):
    """Convierte un punto UTM a (lon, lat) grados decimales WGS84."""
    x = easting - 500000.0
    y = northing - 10000000.0 if southern else northing

    m = y / _K0
    mu = m / (_R * _M1)

    p_rad = (mu + _P2 * math.sin(2 * mu) + _P3 * math.sin(4 * mu)
             + _P4 * math.sin(6 * mu) + _P5 * math.sin(8 * mu))

    p_sin = math.sin(p_rad)
    p_sin2 = p_sin * p_sin
    p_cos = math.cos(p_rad)
    p_tan = p_sin / p_cos
    p_tan2 = p_tan * p_tan
    p_tan4 = p_tan2 * p_tan2

    ep_sin = 1 - _E * p_sin2
    ep_sin_sqrt = math.sqrt(1 - _E * p_sin2)

    n = _R / ep_sin_sqrt
    r = (1 - _E) / ep_sin

    c = _E_P2 * p_cos ** 2
    c2 = c * c

    d = x / (n * _K0)
    d2, d3, d4, d5, d6 = d ** 2, d ** 3, d ** 4, d ** 5, d ** 6

    latitude = (p_rad - (p_tan / r) * (
        d2 / 2
        - d4 / 24 * (5 + 3 * p_tan2 + 10 * c - 4 * c2 - 9 * _E_P2)
        + d6 / 720 * (61 + 90 * p_tan2 + 298 * c + 45 * p_tan4 - 252 * _E_P2 - 3 * c2)
    ))

    longitude = (d
                 - d3 / 6 * (1 + 2 * p_tan2 + c)
                 + d5 / 120 * (5 - 2 * c + 28 * p_tan2 - 3 * c2 + 8 * _E_P2 + 24 * p_tan4)
                 ) / p_cos

    central_meridian = -183.0 + 6.0 * zone_number
    return math.degrees(longitude) + central_meridian, math.degrees(latitude)


# --------------------------------------------------------------------------
# 3) Extraer provincias + anillos, convertir a GeoJSON.
# --------------------------------------------------------------------------
def extraer_provincias(spdf):
    """spdf: el RObject raíz (SpatialPolygonsDataFrame). Devuelve una
    lista de dicts {codigo, nombre, rings: [[(lon,lat),...], ...], holes:[bool,...]}."""
    data = spdf.attrs["data"]
    nombres_col = [c for c in data.attrs["names"].value]
    col_codigo = data.value[nombres_col.index("DPA_PROVIN")].value
    col_nombre = data.value[nombres_col.index("DPA_DESPRO")].value

    proj = spdf.attrs["proj4string"].attrs["projargs"].value[0] or ""
    zona = 17
    if "+zone=" in proj:
        try:
            zona = int(proj.split("+zone=")[1].split()[0])
        except (ValueError, IndexError):
            pass
    sur = "+south" in proj

    polygons_list = spdf.attrs["polygons"].value
    provincias = []
    for i, prov_obj in enumerate(polygons_list):
        rings_raw = prov_obj.attrs["Polygons"].value
        anillos = []
        for ring in rings_raw:
            coords_flat = ring.attrs["coords"].value
            es_hueco = bool(ring.attrs["hole"].value[0])
            n_pts = len(coords_flat) // 2
            # coords viene column-major (dim = [n_pts, 2]): primero todas
            # las X, luego todas las Y.
            xs = coords_flat[:n_pts]
            ys = coords_flat[n_pts:]
            anillo_latlon = [utm_to_lonlat(x, y, zona, sur) for x, y in zip(xs, ys)]
            anillos.append({"coords": anillo_latlon, "hole": es_hueco})
        # "labpt" es el punto que el propio shapefile define para rotular
        # la provincia (su centroide de mayor área) -- mucho más confiable
        # que promediar manualmente todos los vértices.
        labpt_x, labpt_y = prov_obj.attrs["labpt"].value
        lon_c, lat_c = utm_to_lonlat(labpt_x, labpt_y, zona, sur)
        provincias.append({
            "codigo": col_codigo[i],
            "nombre": col_nombre[i],
            "anillos": anillos,
            "centroide": (lon_c, lat_c),
        })
    return provincias


def agrupar_en_poligonos(anillos):
    """Agrupa anillos sueltos en polígonos GeoJSON: cada anillo no-hueco
    abre un nuevo polígono; los huecos que le siguen son sus interiores."""
    poligonos = []
    actual = None
    for anillo in anillos:
        if not anillo["hole"]:
            actual = [anillo["coords"]]
            poligonos.append(actual)
        else:
            if actual is None:
                actual = [anillo["coords"]]
                poligonos.append(actual)
            else:
                actual.append(anillo["coords"])
    return poligonos


def construir_geojson(provincias):
    features = []
    for p in provincias:
        poligonos = agrupar_en_poligonos(p["anillos"])
        geojson_coords = [
            [[[round(lon, 6), round(lat, 6)] for lon, lat in anillo] for anillo in poligono]
            for poligono in poligonos
        ]
        features.append({
            "type": "Feature",
            "id": p["nombre"],
            "properties": {
                "codigo": p["codigo"],
                "nombre": p["nombre"],
                "centroide_lon": round(p["centroide"][0], 5),
                "centroide_lat": round(p["centroide"][1], 5),
            },
            "geometry": {"type": "MultiPolygon", "coordinates": geojson_coords},
        })
    return {"type": "FeatureCollection", "features": features}


# --------------------------------------------------------------------------
# 4) Punto de entrada: genera el GeoJSON y una figura Plotly de muestra.
# --------------------------------------------------------------------------
def main():
    ruta_rds = sys.argv[1] if len(sys.argv) > 1 else "PROVINCIA_CONALI_EDIT.RDS"
    if not os.path.exists(ruta_rds):
        print(f"No se encontró el archivo: {ruta_rds}")
        sys.exit(1)

    print(f"Leyendo {ruta_rds} ...")
    spdf = load_rds(ruta_rds)
    provincias = extraer_provincias(spdf)
    print(f"  {len(provincias)} provincias extraídas.")

    geojson = construir_geojson(provincias)
    # El dashboard busca este archivo junto a app.py y también en assets.
    # Se genera junto al RDS; si el RDS está en otra carpeta, copie el archivo
    # resultante a la raíz del proyecto o a Dashboard_CGTGEv2/assets/.
    salida_geojson = os.path.join(os.path.dirname(os.path.abspath(ruta_rds)) or ".",
                                   "provincias_ecuador.geojson")
    with open(salida_geojson, "w", encoding="utf-8") as f:
        json.dump(geojson, f, ensure_ascii=False)
    print(f"  GeoJSON escrito en: {salida_geojson}")

    # --- Figura de ejemplo con Plotly, coloreando las provincias del Eje 1 ---
    try:
        import plotly.graph_objects as go
    except ImportError:
        print("  (plotly no está instalado; se omite la figura de ejemplo)")
        return

    provincias_resaltadas = {
        # Inauguraciones de infraestructura.
        "SANTO DOMINGO DE LOS TSÁCHILAS": {"valor": 2, "grupo": 2},
        "MANABÍ": {"valor": 3, "grupo": 1},
        "MORONA SANTIAGO": {"valor": 1, "grupo": 1},
        "NAPO": {"valor": 1, "grupo": 2},
    }
    nombres_geojson = [feat["properties"]["nombre"] for feat in geojson["features"]]
    z = [provincias_resaltadas.get(n, {"grupo": 0})["grupo"] for n in nombres_geojson]

    etiquetas = {
        "SANTO DOMINGO DE LOS TSÁCHILAS": "Santo Domingo<br>de los Tsáchilas · 2",
        "MANABÍ": "Manabí · 3",
        "MORONA SANTIAGO": "Morona Santiago · 1",
        "NAPO": "Napo · 1",
    }
    centros = {
        feat["properties"]["nombre"]: (
            feat["properties"]["centroide_lon"],
            feat["properties"]["centroide_lat"],
        )
        for feat in geojson["features"]
    }
    provincias_etiquetadas = list(etiquetas)

    fig = go.Figure(go.Choropleth(
        geojson=geojson,
        locations=nombres_geojson,
        featureidkey="properties.nombre",
        z=z,
        zmin=0, zmax=2,
        # 0 = blanco, 1 = amarillo institucional y 2 = violeta institucional.
        colorscale=[
            [0.00, "#ffffff"], [0.24, "#ffffff"],
            [0.25, "#f8bd20"], [0.74, "#f8bd20"],
            [0.75, "#4d3a94"], [1.00, "#4d3a94"],
        ],
        showscale=False,
        marker_line_color="#25325f",
        marker_line_width=1.15,
        hovertemplate="%{location}<extra></extra>",
    ))
    fig.add_trace(go.Scattergeo(
        lon=[centros[n][0] for n in provincias_etiquetadas],
        lat=[centros[n][1] for n in provincias_etiquetadas],
        mode="markers",
        marker=dict(size=5, color="#17245b", line=dict(width=1, color="#ffffff")),
        hoverinfo="skip",
        showlegend=False,
    ))
    fig.update_geos(
        scope="south america",
        lataxis_range=[-5.6, 1.8],
        lonaxis_range=[-82.6, -74.2],
        showcountries=True, countrycolor="#c8cddd",
        showland=True, landcolor="#ffffff",
        showocean=True, oceancolor="#ffffff",
        resolution=50,
        bgcolor="#ffffff",
    )
    fig.update_layout(margin=dict(l=18, r=18, t=12, b=12), height=420,
                       paper_bgcolor="#ffffff", plot_bgcolor="#ffffff")

    salida_html = os.path.join(os.path.dirname(os.path.abspath(ruta_rds)) or ".",
                                "mapa_provincias_preview.html")
    fig.write_html(salida_html)
    print(f"  Vista previa del mapa guardada en: {salida_html}")


if __name__ == "__main__":
    main()
