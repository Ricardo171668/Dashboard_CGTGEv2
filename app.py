"""Panel Integrado de Gestión MINEDEC - aplicación Dash."""
from __future__ import annotations

from datetime import datetime, timedelta
from email.utils import parsedate_to_datetime
from io import BytesIO
from pathlib import Path
import hashlib
import json
import os
import re
import unicodedata
from urllib.parse import quote, urlencode
from urllib.request import Request, urlopen
import pandas as pd
import plotly.graph_objects as go
import requests
from dash import ALL, MATCH, Dash, Input, Output, State, ctx, dash_table, dcc, html, no_update
from flask import abort, send_file

BASE_DIR = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()
NOMBRES_ENTREGABLES = {
    "E-03_Manual_de_usuario.docx",
    "E-04_Diccionario_de_datos_e_indicadores_Anexo5.xlsx",
    "E-05_Fichas_metodologicas.docx",
}


def directorios_entregables():
    """Ubicaciones admitidas para ejecutar el proyecto desde Windows o ZIP.

    En el equipo institucional, ``app.py`` está dentro de Dashboard_CGTGEv2,
    pero Quipux se guarda en la carpeta superior Dashboard_CGTGE. El segundo
    candidato permite también distribuir el proyecto como una sola carpeta.
    """
    candidatos = [
        BASE_DIR.parent / "Quipux" / "Entregables05102026",
        BASE_DIR / "Quipux" / "Entregables05102026",
        BASE_DIR / "Entregables05102026",
        BASE_DIR,
        Path.cwd().parent / "Quipux" / "Entregables05102026",
        Path.cwd() / "Quipux" / "Entregables05102026",
        Path.cwd() / "Entregables05102026",
        Path.cwd(),
    ]
    resultado = []
    for directorio in candidatos:
        try:
            directorio = directorio.resolve()
        except OSError:
            continue
        if directorio not in resultado:
            resultado.append(directorio)
    return resultado
SECCIONES = [
    ("vision", "Visión Ejecutiva"),
    ("pnd", "Plan Nacional de Desarrollo"),
    ("kpi-estrategicos", "KPI´s Estratégicos"),
    ("kpi-institucionales", "KPI´s Institucionales"),
    ("presupuesto", "Ejecución Presupuestaria - Inversión"),
    ("inventario", "Inventario y Recurso de Información"),
    ("documentacion", "Documentación"),
]

SECCIONES_INDICADORES = [
    ("pnd", "Plan Nacional de Desarrollo"),
    ("kpi-estrategicos", "KPI´s Estratégicos"),
    ("kpi-institucionales", "KPI´s Institucionales"),
]

# Al entrar a Visión Ejecutiva se presenta un menú de los 5 viceministerios
# (igual al prototipo); solo "gestion-educativa" (bases de Alimentación/
# Uniformes/Textos/Mobiliario) tiene datos reales hoy — el resto se muestra
# "en construcción" hasta tener sus fuentes.
VISION_TABS = [
    ("gestion-educativa", "Viceministerio de Gestión Educativa"),
    ("educacion-superior", "Viceministerio de Educación Superior"),
    ("educacion", "Viceministerio de Educación"),
    ("deporte", "Viceministerio del Deporte"),
    ("cultura", "Viceministerio de Cultura"),
]

# Las 4 bases de Gestión Educativa, cada una con su propia estructura real
# (no siguen el formato PND/KPI): tablas por Zona/Provincia/Cantón con
# beneficiarios y, en Mobiliario y Transporte, estudiantes por institución.
# Solo se usan conteos — nunca los montos de inversión (pedido explícito:
# "sin mostrar dinero").
BASES_GESTION_EDUCATIVA = [
    # El prefijo debe ser el INICIO real del nombre de archivo que llega de
    # OneDrive (buscar_archivo solo ignora tildes/mayúsculas, no cambia
    # espacios ni corrige variantes). Se usa una sola palabra distintiva en
    # vez del nombre completo para no fallar por una tilde o un plural que
    # cambie en una próxima actualización del archivo:
    #   Alimentación Escolar.xlsx
    #   Uniformes Ecolares.xlsx        (tal como llega, sin la "s" de "Escolares")
    #   Textos Escolares.xlsx
    #   INVERSION MOBILIARIO Y TRANSPORTE.xlsx  (empieza con "Inversión", no con "Mobiliario")
    ("alimentacion", "Alimentación Escolar"),
    ("uniformes", "Uniformes Escolares"),
    ("textos escolares", "Textos Escolares"),
    ("inversion mobiliario", "Mobiliario y Transporte Escolar"),
]

SECCIONES_PRINCIPALES = [
    ("vision", "Visión Ejecutiva"),
    ("indicadores", "Indicadores"),
    ("presupuesto", "Ejecución Presupuestaria - Inversión"),
    ("inventario", "Inventario y Recurso de Información"),
    ("documentacion", "Documentación"),
]

# Iconografía lineal, monocromática y minimalista según el manual técnico.
ICONOS = {
    "vision": "◎",
    "pnd": "⌖",
    "kpi-estrategicos": "◇",
    "kpi-institucionales": "◫",
    "presupuesto": "▥",
    "inventario": "▤",
    "documentacion": "▧",
}

MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
         "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]


def buscar_archivo(prefijo):
    """Busca el .xlsx más reciente cuyo nombre empiece con el prefijo dado,
    sin distinguir mayúsculas, tildes ni carpeta de publicación."""
    def sin_tildes(texto):
        texto = str(texto)

        # Plotly Cloud codifica algunos caracteres Unicode de los nombres de
        # archivo como texto literal. Por ejemplo, "é" puede convertirse en
        # "#U00e9". Se reconstruye el carácter antes de normalizar el nombre.
        def decodificar_plotly(coincidencia):
            try:
                return chr(int(coincidencia.group(1), 16))
            except (TypeError, ValueError):
                return coincidencia.group(0)

        texto = re.sub(r"#U([0-9A-Fa-f]{4,8})", decodificar_plotly,
                       texto, flags=re.IGNORECASE)
        normalizado = unicodedata.normalize("NFKD", texto.lower())
        return "".join(c for c in normalizado if not unicodedata.combining(c))

    prefijo_normalizado = sin_tildes(prefijo)

    # Plotly Cloud puede iniciar el proceso desde una carpeta diferente a la
    # que contiene app.py. Primero se prueban ubicaciones directas y luego se
    # recorren únicamente carpetas razonables del paquete publicado. El filtro
    # por suffix permite también extensiones .XLSX en mayúsculas.
    raices = []
    for raiz in (BASE_DIR, Path.cwd(), BASE_DIR.parent, Path.cwd().parent):
        try:
            raiz = raiz.resolve()
        except OSError:
            continue
        if raiz.exists() and raiz not in raices:
            raices.append(raiz)

    archivos = []
    vistos = set()
    for raiz in raices:
        try:
            encontrados = (p for p in raiz.rglob("*")
                            if p.is_file() and p.suffix.lower() == ".xlsx")
            for p in encontrados:
                clave = str(p.resolve())
                if clave not in vistos:
                    vistos.add(clave)
                    archivos.append(p)
        except (OSError, PermissionError):
            # Una raíz sin permisos no debe impedir revisar las demás.
            continue

    candidatos = [p for p in archivos
                  if sin_tildes(p.stem).startswith(prefijo_normalizado)]
    # Si existe el archivo oficial sin sufijos como (1), (2), se utiliza ese.
    # Así una copia antigua descargada por Windows no reemplaza accidentalmente
    # a la matriz que el usuario está actualizando.
    exactos = [p for p in candidatos if sin_tildes(p.stem) == prefijo_normalizado]
    if exactos:
        return max(exactos, key=lambda p: p.stat().st_mtime)
    return max(candidatos, key=lambda p: p.stat().st_mtime) if candidatos else None


def diagnostico_excel():
    """Informa qué archivos de Excel son visibles para el servidor publicado."""
    rutas = []
    for raiz in (BASE_DIR, Path.cwd(), BASE_DIR.parent, Path.cwd().parent):
        try:
            raiz = raiz.resolve()
            if not raiz.exists():
                continue
            for p in raiz.rglob("*"):
                if p.is_file() and p.suffix.lower() == ".xlsx":
                    visible = str(p.resolve())
                    if visible not in rutas:
                        rutas.append(visible)
        except (OSError, PermissionError):
            continue
    return "; ".join(rutas) if rutas else "ningún archivo .xlsx visible"


def normalizar_periodo(valor):
    """Conserva años y años lectivos como etiquetas categóricas."""
    if pd.isna(valor):
        return pd.NA
    if isinstance(valor, (int, float)) and float(valor).is_integer():
        return str(int(valor))
    texto = str(valor).strip().replace("–", "-").replace("—", "-")
    if re.fullmatch(r"\d{4}\.0", texto):
        return texto[:-2]
    coincidencia = re.fullmatch(r"(\d{4})\s*-\s*(\d{4})", texto)
    return f"{coincidencia.group(1)}-{coincidencia.group(2)}" if coincidencia else texto


def orden_periodo(valor):
    """Orden cronológico: 2024 antes de 2024-2025 y luego 2025."""
    texto = normalizar_periodo(valor)
    coincidencia = re.fullmatch(r"(\d{4})(?:-(\d{4}))?", str(texto))
    if not coincidencia:
        return 99999999
    inicio = int(coincidencia.group(1))
    fin = int(coincidencia.group(2) or inicio)
    return inicio * 10000 + fin


def normalizar_mes(valor):
    """Conserva meses y rangos semestrales como categorías del gráfico."""
    if pd.isna(valor) or not str(valor).strip():
        return pd.NA
    texto = str(valor).strip().replace("–", "-").replace("—", "-")
    texto = re.sub(r"\s*-\s*", "-", texto)
    return texto


def orden_mes(valor):
    """Ordena meses simples y rangos como Enero-Junio o Julio-Diciembre."""
    texto = normalizar_mes(valor)
    if pd.isna(texto):
        return 0
    inicio = str(texto).split("-", 1)[0].strip().lower()
    if inicio in MES_ORDEN:
        return MES_ORDEN[inicio]
    if inicio in {"primer semestre", "primer sem", "semestre 1"}:
        return 1
    if inicio in {"segundo semestre", "segundo sem", "semestre 2"}:
        return 7
    return 99


def etiqueta_mes(valor):
    """Abrevia meses simples y mantiene completos los rangos semestrales."""
    texto = normalizar_mes(valor)
    if pd.isna(texto):
        return pd.NA
    if str(texto).lower() in MES_ORDEN:
        return str(texto).capitalize()[:3]
    return str(texto)


# ---------------------------------------------------------------------------
# Carga de datos · Plan Nacional de Desarrollo
# ---------------------------------------------------------------------------
def cargar_base_pnd():
    """Lee el Excel del PND aunque Windows añada (1), (2), etc. al nombre."""
    # La matriz vigente indicada por el usuario es "Indicadores_PND version 1".
    # Se conserva el nombre anterior únicamente como respaldo para despliegues
    # donde todavía no se haya reemplazado el archivo.
    archivo = (buscar_archivo("Indicadores_PND version 1")
               or buscar_archivo("Indicadores_PND"))
    if not archivo:
        return pd.DataFrame(), "No se encontró 'Indicadores_PND version 1.xlsx' junto a app.py."
    try:
        df = pd.read_excel(archivo, engine="openpyxl")
        df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")].copy()
        # También elimina espacios internos duplicados en los encabezados.
        df.columns = [" ".join(str(c).strip().split()) for c in df.columns]
        requeridas = {
            "NOMBRE DEL INDICADOR", "VICEMINISTERIO", "Definición", "Unidad de medida",
            "Fuente de datos", "Periodicidad", "Año", "Línea base", "Meta",
            "Estimador", "Numerador", "Denominador", "Alerta", "Observación",
        }
        faltan = sorted(requeridas - set(df.columns))
        if faltan:
            return pd.DataFrame(), "Faltan columnas: " + ", ".join(faltan)
        # No convertir a número: el libro también contiene años lectivos como
        # 2024-2025, que deben conservarse literalmente en el gráfico.
        df["Año"] = df["Año"].map(normalizar_periodo).astype("string")
        df["_orden_periodo"] = df["Año"].map(orden_periodo)
        for col in ["Línea base", "Meta", "Estimador", "Numerador", "Denominador"]:
            df[col] = pd.to_numeric(df[col], errors="coerce")
        for col in ["NOMBRE DEL INDICADOR", "VICEMINISTERIO"]:
            df[col] = df[col].astype("string").str.strip()
        # "Fórmula de Cálculo" es opcional: los Excel del PND anteriores no la traían.
        if "Fórmula de Cálculo" in df.columns:
            df["Fórmula de Cálculo"] = df["Fórmula de Cálculo"].astype("string")
        df = df.dropna(subset=["NOMBRE DEL INDICADOR", "VICEMINISTERIO", "Año"])

        # Respaldo para libros cuyas fórmulas no fueron recalculadas al guardarse.
        alerta_vacia = df["Alerta"].isna() | df["Alerta"].astype(str).str.strip().eq("")
        comparables = df["Meta"].notna() & df["Estimador"].notna()
        df.loc[alerta_vacia & comparables & (df["Estimador"] > df["Meta"]), "Alerta"] = "Cumplimiento sobre la meta"
        df.loc[alerta_vacia & comparables & (df["Estimador"] == df["Meta"]), "Alerta"] = "Cumplimiento igual a la meta"
        df.loc[alerta_vacia & comparables & (df["Estimador"] < df["Meta"]), "Alerta"] = "Incumplimiento de meta"
        return df.sort_values(["VICEMINISTERIO", "NOMBRE DEL INDICADOR", "_orden_periodo"]), None
    except Exception as exc:
        return pd.DataFrame(), f"No fue posible leer la base del PND: {exc}"


def version_pnd():
    archivo = (buscar_archivo("Indicadores_PND version 1")
               or buscar_archivo("Indicadores_PND version 1"))
    return archivo.stat().st_mtime_ns if archivo else 0


# ---------------------------------------------------------------------------
# Carga de datos · Excel con estructura tipo PND (VICEMINISTERIO, NOMBRE DEL
# INDICADOR, Línea base, Meta, Estimador, Alerta, Observación), usada tanto por
# KPI's Estratégicos como por KPI's Institucionales, con "Mes" opcional
# (vacío para indicadores anuales, con valor para los mensuales/semestrales).
# ---------------------------------------------------------------------------
MES_ORDEN = {mes.lower(): i for i, mes in enumerate(MESES, start=1)}


def cargar_base_tipo_pnd(prefijo, nombre_legible):
    archivo = buscar_archivo(prefijo)
    if not archivo:
        return pd.DataFrame(), (
            f"No se encontró '{nombre_legible}'. "
            f"Carpeta de app.py: {BASE_DIR}. "
            f"Carpeta de ejecución: {Path.cwd()}. "
            f"Excel visibles: {diagnostico_excel()}."
        )
    try:
        df = pd.read_excel(archivo, engine="openpyxl")
        df.columns = [" ".join(str(c).strip().split()) if not str(c).startswith("Unnamed") else c
                      for c in df.columns]

        # Unifica las variantes de encabezados que llegan desde las matrices.
        # Así, nuevas actualizaciones pueden usar mayúsculas/minúsculas o
        # "cálculo/calculo" sin romper la ficha del indicador.
        equivalencias = {
            "Fórmula de calculo": "Fórmula de Cálculo",
            "Fórmula de cálculo": "Fórmula de Cálculo",
            "Formula de calculo": "Fórmula de Cálculo",
            "Formula de cálculo": "Fórmula de Cálculo",
            "Fecha de corte": "Fecha de Corte",
            "Fecha de transferencia": "Fecha de Transferencia",
        }
        df = df.rename(columns={c: equivalencias.get(c, c) for c in df.columns})

        # Compatibilidad con archivos donde la columna de Alerta llega sin
        # encabezado (queda justo antes de "Observación").
        cols = list(df.columns)
        if "Observación" in cols:
            idx_obs = cols.index("Observación")
            if idx_obs > 0 and str(cols[idx_obs - 1]).startswith("Unnamed"):
                cols[idx_obs - 1] = "Alerta"
                df.columns = cols

        df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")].copy()

        requeridas = {
            "NOMBRE DEL INDICADOR", "VICEMINISTERIO", "Unidad de medida", "Fuente de datos",
            "Periodicidad", "Año", "Línea base", "Meta", "Estimador",
            "Numerador", "Denominador", "Alerta", "Observación",
        }
        faltan = sorted(requeridas - set(df.columns))
        if faltan:
            return pd.DataFrame(), "Faltan columnas: " + ", ".join(faltan)

        # Conserva tanto años calendario (2026) como años lectivos
        # (2026-2027). Al convertir esta columna a número, los años lectivos
        # se volvían NaN y sus filas desaparecían del dashboard.
        df["Año"] = df["Año"].map(normalizar_periodo).astype("string")
        for col in ["Línea base", "Meta", "Estimador", "Numerador", "Denominador"]:
            df[col] = pd.to_numeric(df[col], errors="coerce")
        for col in ["NOMBRE DEL INDICADOR", "VICEMINISTERIO"]:
            df[col] = df[col].astype("string").str.strip()
        if "Mes" in df.columns:
            df["Mes"] = df["Mes"].map(normalizar_mes).astype("string")
        else:
            df["Mes"] = pd.array([pd.NA] * len(df), dtype="string")
        # Si toda la columna Alerta viene vacía, pandas la infiere como float64
        # y luego falla al intentar escribir texto ahí; se fuerza a texto.
        df["Alerta"] = df["Alerta"].astype("string")
        for col in ("ENCARGADO", "TIPO INDICADOR"):
            if col in df.columns:
                df[col] = df[col].astype("string").str.strip()
        if "Fórmula de Cálculo" in df.columns:
            df["Fórmula de Cálculo"] = df["Fórmula de Cálculo"].astype("string")
        df = df.dropna(subset=["NOMBRE DEL INDICADOR", "VICEMINISTERIO", "Año"])

        # Algunos indicadores de "Porcentaje" se registran como fracción (0,90 = 90%)
        # y otros ya en escala 0-100 (97,02), mezclados en el mismo archivo. Se
        # normaliza celda por celda: solo se multiplica por 100 cuando el valor
        # es ≤ 1, para no alterar los que ya vienen correctos.
        es_porcentaje = df["Unidad de medida"].astype(str).str.contains("orcentaje", case=False, na=False)
        for col in ["Línea base", "Meta", "Estimador"]:
            parece_fraccion = es_porcentaje & df[col].notna() & (df[col] <= 1)
            df.loc[parece_fraccion, col] = df.loc[parece_fraccion, col] * 100

        # Orden cronológico real y etiqueta de período: "Ene 2026" si hay mes
        # registrado, o el año/año lectivo tal como consta en el Excel.
        df["_mes_num"] = df["Mes"].map(orden_mes)
        df["_orden_periodo"] = (
            df["Año"].map(orden_periodo) * 100
            + df["_mes_num"].fillna(0).astype(int)
        )
        tiene_mes = df["Mes"].notna()
        df["Periodo"] = df["Año"].astype("string")
        df.loc[tiene_mes, "Periodo"] = (df.loc[tiene_mes, "Mes"].map(etiqueta_mes)
                                         + " " + df.loc[tiene_mes, "Año"].astype("string"))

        # Respaldo para filas cuya Alerta no fue registrada.
        alerta_vacia = df["Alerta"].isna() | df["Alerta"].astype(str).str.strip().eq("")
        comparables = df["Meta"].notna() & df["Estimador"].notna()
        df.loc[alerta_vacia & comparables & (df["Estimador"] > df["Meta"]), "Alerta"] = "Cumplimiento sobre la meta"
        df.loc[alerta_vacia & comparables & (df["Estimador"] == df["Meta"]), "Alerta"] = "Cumplimiento igual a la meta"
        df.loc[alerta_vacia & comparables & (df["Estimador"] < df["Meta"]), "Alerta"] = "Incumplimiento de meta"

        return df.sort_values(["VICEMINISTERIO", "NOMBRE DEL INDICADOR", "_orden_periodo"]), None
    except Exception as exc:
        return pd.DataFrame(), f"No fue posible leer '{nombre_legible}': {exc}"


# ---------------------------------------------------------------------------
# Carga de datos · KPI's Estratégicos
# ---------------------------------------------------------------------------
def cargar_base_kpi():
    return cargar_base_tipo_pnd("kpi_estratégicos", "kpi_estratégicos.xlsx")


def version_kpi():
    archivo = buscar_archivo("kpi_estratégicos")
    return archivo.stat().st_mtime_ns if archivo else 0


# ---------------------------------------------------------------------------
# Carga de datos · KPI's Institucionales
# ---------------------------------------------------------------------------
def cargar_base_kpi_inst():
    return cargar_base_tipo_pnd("kpi_institucional", "KPI_INSTITUCIONALES.xlsx")


def version_kpi_inst():
    archivo = buscar_archivo("kpi_institucional")
    return archivo.stat().st_mtime_ns if archivo else 0


# ---------------------------------------------------------------------------
# Carga de datos · Gestión Educativa (Alimentación Escolar, Uniformes, Textos
# Escolares y Mobiliario y Transporte). A diferencia del PND/KPI, cada Excel
# es una tabla de beneficiarios por Zona/Provincia/Cantón (o por institución,
# en Mobiliario), sin la estructura de indicador con Meta/Línea base. Aquí
# solo se agregan conteos (beneficiarios, instituciones, estudiantes); el
# monto de inversión de cada base NUNCA se usa para la tarjeta de resumen
# (pedido explícito del usuario: "sin mostrar dinero").
# ---------------------------------------------------------------------------
def _localizar_hoja_con_columna(ruta, columna_clave, filas_busqueda=30):
    """Devuelve el primer DataFrame, de cualquier hoja del libro, cuya fila de
    encabezado contenga `columna_clave` (ya normalizada). Algunas matrices
    (p. ej. Textos Escolares) traen una fila de título antes del encabezado
    real, así que no se asume que el encabezado está en la fila 1."""
    libro = pd.ExcelFile(ruta, engine="openpyxl")
    for hoja in libro.sheet_names:
        vista = pd.read_excel(libro, sheet_name=hoja, header=None,
                               nrows=filas_busqueda, engine="openpyxl")
        for fila_num, fila in vista.iterrows():
            etiquetas = {_normalizar_encabezado_presupuesto(x)
                         for x in fila.dropna() if str(x).strip()}
            if columna_clave in etiquetas:
                return pd.read_excel(libro, sheet_name=hoja, header=int(fila_num),
                                      engine="openpyxl")
    return None


PROVINCIAS_ECUADOR = {
    "AZUAY", "BOLIVAR", "CANAR", "CARCHI", "CHIMBORAZO", "COTOPAXI",
    "EL ORO", "ESMERALDAS", "GALAPAGOS", "GUAYAS", "IMBABURA", "LOJA",
    "LOS RIOS", "MANABI", "MORONA SANTIAGO", "NAPO", "ORELLANA", "PASTAZA",
    "PICHINCHA", "SANTA ELENA", "SANTO DOMINGO DE LOS TSACHILAS", "SUCUMBIOS",
    "TUNGURAHUA", "ZAMORA CHINCHIPE",
}


def _provincia_homologada(valor):
    """Homologa variantes de escritura y descarta valores que no sean una de
    las 24 provincias oficiales del Ecuador."""
    clave = _normalizar_encabezado_presupuesto(valor)
    clave = re.sub(r"^(PROVINCIA_DE_|PROVINCIA_)", "", clave).replace("_", " ").strip()
    equivalencias = {
        "SANTO DOMINGO": "SANTO DOMINGO DE LOS TSACHILAS",
        "SANTO DOMINGO DE LOS TSACHILAS": "SANTO DOMINGO DE LOS TSACHILAS",
        "GALAPAGO": "GALAPAGOS",
        "CANAR": "CANAR",
    }
    clave = equivalencias.get(clave, clave)
    return clave if clave in PROVINCIAS_ECUADOR else None


def _columna_por_claves(mapa, claves):
    for clave in claves:
        if clave in mapa:
            return mapa[clave]
    return None


def _metricas_gestion(df, mapa):
    dato = {}
    col_benef = mapa.get("BENEFICIARIOS")
    col_ie = _columna_por_claves(mapa, ("NRO_IE", "NO_IE", "NO_I_E"))
    col_amie = mapa.get("AMIE")
    col_est = mapa.get("TOTAL_ESTUDIANTES")
    col_prov = mapa.get("PROVINCIA")
    col_sost = _columna_por_claves(mapa, ("SOSTENIMIENTO", "TIPO_SOSTENIMIENTO"))
    col_area = _columna_por_claves(mapa, ("AREA", "AREA_GEOGRAFICA", "URBANO_RURAL"))

    if col_benef:
        dato["beneficiarios"] = float(pd.to_numeric(df[col_benef], errors="coerce").sum())
    if col_ie:
        dato["instituciones"] = float(pd.to_numeric(df[col_ie], errors="coerce").sum())
    elif col_amie:
        dato["instituciones"] = float(df[col_amie].dropna().astype(str).str.strip().nunique())
    if col_est:
        dato["estudiantes"] = float(pd.to_numeric(df[col_est], errors="coerce").sum())
    if col_prov:
        provincias = df[col_prov].map(_provincia_homologada).dropna()
        dato["provincias"] = min(24, int(provincias.nunique()))
    if col_sost:
        dato["sostenimientos"] = int(df[col_sost].dropna().astype(str).str.strip().replace("", pd.NA).nunique())
    if col_area:
        dato["areas"] = int(df[col_area].dropna().astype(str).str.strip().replace("", pd.NA).nunique())
    return dato


def _resumen_base_gestion_educativa(archivo, nombre_legible):
    """Lee una base, homologa provincias y precalcula sus filtros reales."""
    df = _localizar_hoja_con_columna(archivo, "BENEFICIARIOS")
    if df is None:
        df = _localizar_hoja_con_columna(archivo, "TOTAL_ESTUDIANTES")
    if df is None:
        raise ValueError(
            f"'{nombre_legible}' no tiene una columna de Beneficiarios ni "
            "de Total Estudiantes reconocible."
        )
    df.columns = [" ".join(str(c).strip().split()) for c in df.columns]
    mapa = {_normalizar_encabezado_presupuesto(c): c for c in df.columns}
    dato = _metricas_gestion(df, mapa)
    col_regimen = next((c for k, c in mapa.items() if "REGIMEN" in k), None)
    col_clasificacion = next((c for k, c in mapa.items() if "CLASIFICACION" in k), None)
    regimenes = sorted(df[col_regimen].dropna().astype(str).str.strip().replace("", pd.NA).dropna().unique()) \
        if col_regimen else []
    clasificaciones = sorted(df[col_clasificacion].dropna().astype(str).str.strip().replace("", pd.NA).dropna().unique()) \
        if col_clasificacion else []
    dato["regimenes"] = list(regimenes)
    dato["clasificaciones"] = list(clasificaciones)
    segmentos = {}
    for regimen in ["Todos"] + list(regimenes):
        base_regimen = df if regimen == "Todos" else df.loc[df[col_regimen].astype(str).str.strip() == regimen]
        for clasificacion in ["Todas"] + list(clasificaciones):
            filtrado = base_regimen
            if col_clasificacion and clasificacion != "Todas":
                filtrado = filtrado.loc[filtrado[col_clasificacion].astype(str).str.strip() == clasificacion]
            segmentos[f"{regimen}||{clasificacion}"] = _metricas_gestion(filtrado, mapa)
    dato["segmentos"] = segmentos
    # La fuente conjunta de Mobiliario y Transporte se presenta en dos
    # tarjetas independientes. La clasificación se usa internamente y deja
    # de exponerse como filtro al usuario.
    if col_clasificacion:
        clasificacion_normalizada = df[col_clasificacion].map(
            _normalizar_encabezado_presupuesto
        )
        dato["por_recurso"] = {
            "mobiliario": _metricas_gestion(
                df.loc[clasificacion_normalizada.str.contains("MOBILIARIO", na=False)], mapa
            ),
            "transporte": _metricas_gestion(
                df.loc[clasificacion_normalizada.str.contains("TRANSPORTE", na=False)], mapa
            ),
        }
    return dato


def _obtener_origen_gestion_educativa(prefijo, nombre_legible):
    """Da preferencia al vínculo de OneDrive (la fuente que sí se sigue
    actualizando); si no hay uno configurado para esta base, cae de vuelta a
    buscar el Excel junto a app.py. Devuelve (origen_para_pandas,
    descripcion_para_mensajes, version_para_detectar_cambios) o levanta
    una excepción con el motivo exacto si no se pudo obtener ninguno."""
    url = GESTION_EDUCATIVA_DOWNLOAD_URLS.get(nombre_legible, "")
    if url:
        origen, version, _fecha = _descargar_excel_onedrive(url, f"{nombre_legible}.xlsx")
        return origen, f"OneDrive ({nombre_legible}.xlsx)", version
    archivo = buscar_archivo(prefijo)
    if not archivo:
        # Diagnóstico explícito: evita el error silencioso de la vez pasada
        # (prefijo mal escrito) — dice exactamente qué .xlsx sí ve el
        # servidor, para distinguir "no está el archivo" de "el nombre no
        # coincide con el prefijo buscado".
        raise FileNotFoundError(
            f"no se encontró un .xlsx que empiece con '{prefijo}' junto a app.py "
            f"(y no hay vínculo de OneDrive configurado para esta base); "
            f"Excel visibles: {diagnostico_excel()}"
        )
    return archivo, archivo.name, str(archivo.stat().st_mtime_ns)


def cargar_resumen_gestion_educativa():
    """Intenta leer las 4 bases (de OneDrive si hay vínculo configurado, o
    junto a app.py si no); cada una que falte o no se pueda interpretar se
    reporta por separado, sin bloquear a las demás (igual que el resto del
    panel: se muestra lo que sí hay disponible)."""
    resumen = {}
    pendientes = []
    for prefijo, nombre_legible in BASES_GESTION_EDUCATIVA:
        try:
            origen, descripcion, _version = _obtener_origen_gestion_educativa(prefijo, nombre_legible)
        except Exception as exc:
            pendientes.append(f"{nombre_legible} ({exc})")
            continue
        try:
            resumen[nombre_legible] = _resumen_base_gestion_educativa(origen, nombre_legible)
        except Exception as exc:
            pendientes.append(f"{nombre_legible} (se leyó '{descripcion}' pero no se pudo interpretar: {exc})")
    return resumen, pendientes


def version_gestion_educativa(resumen=None, pendientes=None):
    """No hay forma barata de saber si un Excel en OneDrive cambió sin
    descargarlo de nuevo (a diferencia de un archivo local, que sí expone su
    fecha de modificación), así que la versión es un hash del propio
    resultado leído: cambia solo cuando el contenido realmente cambió."""
    resumen = RESUMEN_GESTION_EDUCATIVA if resumen is None else resumen
    pendientes = PENDIENTES_GESTION_EDUCATIVA if pendientes is None else pendientes
    firma = repr((sorted(resumen.items()), sorted(pendientes)))
    return hashlib.sha256(firma.encode()).hexdigest()


# ---------------------------------------------------------------------------
# Carga de datos · Visión Ejecutiva
# ---------------------------------------------------------------------------
def cargar_base_vision():
    """Lee la base normalizada que alimenta las tablas de Visión Ejecutiva."""
    archivo = buscar_archivo("vision_ejecutiva")
    if not archivo:
        return pd.DataFrame(), "No se encontró 'vision_ejecutiva.xlsx' junto a app.py."
    try:
        df = pd.read_excel(archivo, sheet_name="Datos", engine="openpyxl")
        df.columns = [" ".join(str(c).strip().split()) for c in df.columns]
        requeridas = {
            "Sección", "Orden sección", "Tabla", "Orden tabla", "Orden fila",
            "Etiqueta fila", "Indicador", "Orden indicador", "Valor", "Fuente",
            "Fecha de corte", "Nota",
        }
        faltan = sorted(requeridas - set(df.columns))
        if faltan:
            return pd.DataFrame(), "Faltan columnas en Visión Ejecutiva: " + ", ".join(faltan)
        for col in ["Orden sección", "Orden tabla", "Orden fila", "Orden indicador", "Valor"]:
            df[col] = pd.to_numeric(df[col], errors="coerce")
        for col in ["Sección", "Tabla", "Etiqueta fila", "Indicador", "Fuente", "Fecha de corte", "Nota"]:
            df[col] = df[col].fillna("").astype(str).str.strip()
        df = df.dropna(subset=["Orden sección", "Orden tabla", "Orden fila", "Orden indicador", "Valor"])
        df = df[(df["Sección"] != "") & (df["Tabla"] != "") & (df["Indicador"] != "")]
        return df.sort_values(["Orden sección", "Orden tabla", "Orden fila", "Orden indicador"]), None
    except Exception as exc:
        return pd.DataFrame(), f"No fue posible leer la base de Visión Ejecutiva: {exc}"


def version_vision():
    archivo = buscar_archivo("vision_ejecutiva")
    return archivo.stat().st_mtime_ns if archivo else 0


# ---------------------------------------------------------------------------
# Carga de datos · Ejecución Presupuestaria (último ESIGEF disponible)
# ---------------------------------------------------------------------------
PRESUPUESTO_MONETARIAS = [
    "ASIGNADO", "CODIFICADO", "RESERVADO_NEGATIVO", "PRECOMPROMISO",
    "COMPROMISO", "DEVENGADO", "PAGADO", "SALDO_DISPONIBLE",
]
PRESUPUESTO_META = {"archivo": "", "fecha": "", "origen": ""}


ESIGEF_DOWNLOAD_URL = os.getenv(
    "ESIGEF_DOWNLOAD_URL",
    "https://educacionec-my.sharepoint.com/:x:/g/personal/"
    "ricardo_castellanos_educacion_gob_ec/"
    "IQCIs7P5gRv3RbMPaTt_tCufAeiAcJdopSUXr5tr1kiTToA"
    "?e=0KSa78&download=1",
).strip()

# Vínculos "Copiar vínculo" de OneDrive para las 4 bases de Gestión Educativa
# (mismo mecanismo que ESIGEF_actual.xlsx: deben terminar en "&download=1").
# Mientras no se defina un vínculo para una base, se sigue buscando su Excel
# junto a app.py (BASES_GESTION_EDUCATIVA), igual que antes.
GESTION_EDUCATIVA_DOWNLOAD_URLS = {
    "Alimentación Escolar": os.getenv(
        "ALIMENTACION_ESCOLAR_DOWNLOAD_URL",
        "https://educacionec-my.sharepoint.com/:x:/g/personal/daei_educacion_gob_ec/"
        "IQBrz0MHostwQIW5R61Mz3tnATQK3DrSUL7YqpgcgXcYvpY?e=2YwCXT&download=1",
    ).strip(),
    "Uniformes Escolares": os.getenv(
        "UNIFORMES_ESCOLARES_DOWNLOAD_URL",
        "https://educacionec-my.sharepoint.com/:x:/g/personal/daei_educacion_gob_ec/"
        "IQC9L2hLRK4SQoTVUakfR24WAUVeGIAfMNOEv3l0kx7IgJA?e=6fou2z&download=1",
    ).strip(),
    "Textos Escolares": os.getenv(
        "TEXTOS_ESCOLARES_DOWNLOAD_URL",
        "https://educacionec-my.sharepoint.com/:x:/g/personal/daei_educacion_gob_ec/"
        "IQCG5JN7n6U5SYEW3w-5KfxMAepbyAr8GplbqLBlEd5s1DA?e=4wgwGk&download=1",
    ).strip(),
    "Mobiliario y Transporte Escolar": os.getenv(
        "MOBILIARIO_TRANSPORTE_DOWNLOAD_URL",
        "https://educacionec-my.sharepoint.com/:x:/g/personal/daei_educacion_gob_ec/"
        "IQDc3cxCQrPZSqLnDHiYf7IfAdYqyIj2Y374Az_Xg4nHTfI?e=2Ck9L8&download=1",
    ).strip(),
}


def _descargar_excel_onedrive(url, nombre_archivo):
    """Descarga un Excel desde un vínculo de OneDrive/SharePoint que termine en
    '&download=1'. Reproduce la estrategia usada en Shiny/httr: sigue
    redirecciones y evita reutilizar una copia almacenada en caché. Devuelve
    (contenido_en_memoria, version_hash, fecha_ultima_modificacion)."""
    respuesta = requests.get(
        url,
        headers={
            "User-Agent": "Mozilla/5.0",
            "Cache-Control": "no-cache, no-store",
            "Pragma": "no-cache",
        },
        timeout=120,
        allow_redirects=True,
    )
    respuesta.raise_for_status()
    contenido = respuesta.content
    ultima_modificacion = respuesta.headers.get("Last-Modified", "")

    # XLSX es un contenedor ZIP y, por tanto, empieza con la firma PK.
    # Esta comprobación evita que una página HTML de inicio de sesión sea
    # enviada por error a openpyxl.
    if not contenido.startswith(b"PK"):
        raise ValueError(
            f"El enlace de {nombre_archivo} no devolvió un archivo Excel. "
            "Revise que el vínculo de OneDrive permita descargar el archivo "
            "(debe terminar en '&download=1')."
        )
    version = hashlib.sha256(contenido).hexdigest()
    return BytesIO(contenido), version, ultima_modificacion


def _descargar_esigef_actual():
    """Descarga el archivo presupuestario vigente desde OneDrive."""
    origen, version, ultima_modificacion = _descargar_excel_onedrive(
        ESIGEF_DOWNLOAD_URL, "ESIGEF_actual.xlsx"
    )
    return origen, {
        "archivo": "ESIGEF_actual.xlsx",
        "fecha": ultima_modificacion or datetime.now().strftime("%d/%m/%Y %H:%M"),
        "origen": "OneDrive",
        "version": version,
    }


def _normalizar_encabezado_presupuesto(valor):
    """Normaliza tildes, saltos, espacios y guiones de los encabezados ESIGEF."""
    texto = unicodedata.normalize("NFKD", str(valor).strip())
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = re.sub(r"[^A-Za-z0-9]+", "_", texto.upper()).strip("_")
    return texto


def _leer_excel_presupuesto(origen):
    """Localiza automáticamente la hoja y la fila real del encabezado.

    Los reportes ESIGEF pueden incluir una portada, hojas auxiliares o filas
    previas al encabezado. Por eso no se asume la primera hoja ni la fila 1.
    """
    if hasattr(origen, "seek"):
        origen.seek(0)
    libro = pd.ExcelFile(origen, engine="openpyxl")
    revisadas = []

    for hoja in libro.sheet_names:
        vista = pd.read_excel(libro, sheet_name=hoja, header=None,
                             nrows=150, engine="openpyxl")
        revisadas.append(hoja)
        for fila_num, fila in vista.iterrows():
            etiquetas = {
                _normalizar_encabezado_presupuesto(x)
                for x in fila.dropna()
                if str(x).strip()
            }
            requeridas = {"CODIFICADO", "DEVENGADO", "VICEMINISTERIO"}
            if requeridas.issubset(etiquetas):
                return pd.read_excel(
                    libro, sheet_name=hoja, header=int(fila_num),
                    engine="openpyxl"
                )

    raise ValueError(
        "No se encontró una hoja con los encabezados CODIFICADO, DEVENGADO "
        "y Viceministerio. Hojas revisadas: " + ", ".join(revisadas)
    )


def cargar_base_presupuesto():
    global PRESUPUESTO_META
    try:
        try:
            origen, meta = _descargar_esigef_actual()
        except Exception as error_remoto:
            locales = [p for p in BASE_DIR.rglob("*.xlsx")
                       if p.name.lower() == "esigef_actual.xlsx"]
            if not locales:
                return pd.DataFrame(), (
                    "No fue posible descargar ESIGEF_actual.xlsx y tampoco existe una copia local. "
                    f"Detalle: {error_remoto}"
                )
            archivo = max(locales, key=lambda p: p.stat().st_mtime_ns)
            origen = archivo
            meta = {"archivo": archivo.name,
                    "fecha": datetime.fromtimestamp(archivo.stat().st_mtime).strftime("%d/%m/%Y"),
                    "origen": "Copia local de ESIGEF_actual.xlsx",
                    "version": str(archivo.stat().st_mtime_ns)}
        df = _leer_excel_presupuesto(origen)
        df.columns = [" ".join(str(c).replace("\n", " ").strip().split()) for c in df.columns]
        mapa = {_normalizar_encabezado_presupuesto(c): c for c in df.columns}
        if "VICEMINISTERIO" not in mapa:
            raise ValueError("La base no contiene la variable Viceministerio.")
        df = df.rename(columns={mapa["VICEMINISTERIO"]: "Viceministerio"})
        for col in PRESUPUESTO_MONETARIAS:
            original = mapa.get(col)
            if original is None:
                df[col] = 0.0
            else:
                def numero_esigef(valor):
                    if pd.isna(valor):
                        return 0.0
                    if isinstance(valor, (int, float)):
                        return float(valor)
                    texto = str(valor).strip().replace("$", "").replace(" ", "")
                    if "," in texto:
                        texto = texto.replace(".", "").replace(",", ".")
                    return pd.to_numeric(texto, errors="coerce")
                df[col] = df[original].map(numero_esigef).fillna(0.0)
        df["Viceministerio"] = df["Viceministerio"].fillna("Sin clasificación").astype(str).str.strip()
        df = df[df["Viceministerio"].ne("")]
        PRESUPUESTO_META = meta
        return df, None
    except Exception as exc:
        return pd.DataFrame(), f"No fue posible cargar la ejecución presupuestaria: {exc}"


def version_presupuesto():
    return PRESUPUESTO_META.get("version", "")


DATA_PND, ERROR_PND = cargar_base_pnd()
DATA_KPI, ERROR_KPI = cargar_base_kpi()
DATA_KPI_INST, ERROR_KPI_INST = cargar_base_kpi_inst()
RESUMEN_GESTION_EDUCATIVA, PENDIENTES_GESTION_EDUCATIVA = cargar_resumen_gestion_educativa()
# Visión Ejecutiva ahora es una infografía institucional estática y ya no
# depende del archivo vision_ejecutiva.xlsx.
DATA_VISION, ERROR_VISION = pd.DataFrame(), None
DATA_PRESUPUESTO, ERROR_PRESUPUESTO = cargar_base_presupuesto()
VERSION_PND = version_pnd()
VERSION_KPI = version_kpi()
VERSION_KPI_INST = version_kpi_inst()
VERSION_GESTION_EDUCATIVA = version_gestion_educativa()
VERSION_VISION = "infografia-minedec-2026"
VERSION_PRESUPUESTO = version_presupuesto()

# Metadatos de las secciones que sí tienen datos tabulares (menú lateral con acordeón).
SECCIONES_CON_DATOS = {
    "pnd": {"vice_col": "VICEMINISTERIO", "indicador_col": "NOMBRE DEL INDICADOR"},
    "kpi-estrategicos": {"vice_col": "VICEMINISTERIO", "indicador_col": "NOMBRE DEL INDICADOR"},
    "kpi-institucionales": {"vice_col": "VICEMINISTERIO", "indicador_col": "NOMBRE DEL INDICADOR"},
    "presupuesto": {"vice_col": "Viceministerio", "indicador_col": "Viceministerio"},
}

# Secciones cuyos indicadores son pocos y se listan directamente en el menú
# lateral, sin agruparlos primero por viceministerio.
SECCIONES_INDICADORES_PLANOS = {"kpi-estrategicos", "kpi-institucionales"}


def obtener_datos(seccion):
    if seccion == "vision":
        return pd.DataFrame(), None
    if seccion == "pnd":
        return DATA_PND, ERROR_PND
    if seccion == "kpi-estrategicos":
        return DATA_KPI, ERROR_KPI
    if seccion == "kpi-institucionales":
        return DATA_KPI_INST, ERROR_KPI_INST
    if seccion == "presupuesto":
        return DATA_PRESUPUESTO, ERROR_PRESUPUESTO
    return pd.DataFrame(), None


app = Dash(__name__, title="Panel Integrado de Gestión MINEDEC", update_title=None,
           suppress_callback_exceptions=True)
server = app.server


@server.route("/descargar-entregable/<path:nombre_archivo>")
def descargar_entregable(nombre_archivo):
    """Entrega los anexos oficiales desde la carpeta Quipux del proyecto."""
    # Solo se exponen los tres entregables definidos por la aplicación.
    if Path(nombre_archivo).name != nombre_archivo or nombre_archivo not in NOMBRES_ENTREGABLES:
        abort(404)

    for directorio in directorios_entregables():
        archivo = directorio / nombre_archivo
        if archivo.is_file():
            return send_file(archivo, as_attachment=True, download_name=nombre_archivo)

    # En GitHub/Posit Cloud la carpeta puede quedar en un nivel diferente al
    # usado en Windows. Como los nombres están protegidos por la lista blanca,
    # se permite una búsqueda final dentro del proyecto desplegado.
    try:
        coincidencias = [
            ruta for ruta in BASE_DIR.rglob(nombre_archivo)
            if ruta.is_file() and ".git" not in ruta.parts
        ]
    except OSError:
        coincidencias = []
    if coincidencias:
        return send_file(coincidencias[0], as_attachment=True,
                         download_name=nombre_archivo)
    abort(404, description="El entregable no se encuentra en la carpeta Quipux/Entregables05102026.")

# MathJax renderiza las fórmulas en LaTeX (delimitadas con $...$, $$...$$,
# \(...\) o \[...\]) que
# se escriban en la celda "Fórmula de cálculo" del Excel. Se configura antes de
# cargar el script para que reconozca esos delimitadores.
app.index_string = """<!DOCTYPE html>
<html>
    <head>
        {%metas%}
        <title>{%title%}</title>
        {%favicon%}
        {%css%}
        <script>
        window.MathJax = {
          tex: {
            inlineMath: [['$', '$'], ['\\\\(', '\\\\)']],
            displayMath: [['$$', '$$'], ['\\\\[', '\\\\]']],
            processEscapes: true
          },
          svg: {fontCache: 'global'}
        };
        </script>
        <script type="text/javascript" id="MathJax-script" async
          src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js">
        </script>
    </head>
    <body>
        {%app_entry%}
        <footer>
            {%config%}
            {%scripts%}
            {%renderer%}
        </footer>
    </body>
</html>"""


def barra_superior():
    return html.Header([
        html.Button("⌂", id="home-button", className="home-button", title="Inicio"),
        html.Div([
            html.Strong("PANEL INTEGRADO DE GESTIÓN MINEDEC"),
            html.Span("Panel institucional"),
        ], className="header-title"),
        html.Div([
            html.Div([html.Span("Fecha de corte"), html.Strong(datetime.now().strftime("%d/%m/%Y"))],
                     className="cutoff-date"),
            html.Button("⎙", id="print-button", className="print-button", title="Imprimir"),
        ], className="header-actions"),
    ], className="header-bar")


def portada():
    # Iconos lineales, monocromáticos y minimalistas conforme a la guía visual.
    iconos_portada = {
        "vision": """<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64' fill='none' stroke='white' stroke-width='2.8' stroke-linecap='round' stroke-linejoin='round'><rect x='8' y='8' width='48' height='39' rx='4'/><path d='M15 39V27h8v12M28 39V20h8v19M41 39V15h8v24M18 55h28M32 47v8'/></svg>""",
        "indicadores": """<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64' fill='none' stroke='white' stroke-width='2.8' stroke-linecap='round' stroke-linejoin='round'><rect x='10' y='8' width='44' height='48' rx='5'/><path d='M20 21l3 3 6-7M34 21h11M20 35l3 3 6-7M34 35h11M20 49l3 3 6-7M34 49h11'/></svg>""",
        "presupuesto": """<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64' fill='none' stroke='white' stroke-width='2.8' stroke-linecap='round' stroke-linejoin='round'><path d='M9 53h46M14 48V29M25 48V37M36 48V23M47 48V14'/><path d='M14 21l11-7 11 3 14-9M43 8h7v7'/><circle cx='14' cy='21' r='2.5'/><circle cx='25' cy='14' r='2.5'/><circle cx='36' cy='17' r='2.5'/></svg>""",
        "inventario": """<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64' fill='none' stroke='white' stroke-width='2.8' stroke-linecap='round' stroke-linejoin='round'><path d='M7 20h20l5 6h25v27H7z'/><path d='M7 20v-8h19l5 6h20v8M17 35h12M17 43h22'/><circle cx='48' cy='43' r='7'/><path d='M53 48l5 5'/></svg>""",
        "documentacion": """<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64' fill='none' stroke='white' stroke-width='2.8' stroke-linecap='round' stroke-linejoin='round'><path d='M16 7h24l10 10v40H16z'/><path d='M40 7v12h12M24 30h18M24 39h18M24 48h12'/></svg>""",
    }

    def icono(codigo):
        return html.Img(
            src="data:image/svg+xml;utf8," + quote(iconos_portada[codigo]),
            className="portada-boton-icono-svg",
            alt="",
            **{"aria-hidden": "true"},
        )

    def boton(codigo, nombre):
        return html.Button(
            [
                html.Span(icono(codigo), className="portada-boton-icono"),
                html.Span(nombre, className="portada-boton-texto"),
                html.Span("›", className="portada-boton-flecha"),
            ],
            id={"type": "home-card", "index": codigo},
            className="portada-boton",
            n_clicks=0,
            type="button",
        )

    return html.Main(
        className="portada-pagina",
        children=[
            html.Section(
                className="portada-contenido",
                children=[
                    html.Div(
                        className="portada-ilustracion-contenedor",
                        children=html.Img(
                            src=app.get_asset_url("portada-minedec.png"),
                            className="portada-ilustracion",
                            alt="Educación, deporte y cultura en el Ecuador",
                        ),
                    ),
                    html.Img(
                        src=app.get_asset_url("logo-minedec.png"),
                        className="portada-logo",
                        alt="Ministerio de Educación, Deporte y Cultura",
                    ),
                    html.Nav(
                        [boton(codigo, nombre)
                         for codigo, nombre in SECCIONES_PRINCIPALES],
                        className="portada-menu",
                        **{"aria-label": "Secciones principales"},
                    ),
                ],
            ),
            html.Div(className="portada-onda-inferior"),
        ],
    )


def portada_indicadores():
    iconos = {
        "pnd": "⌖",
        "kpi-estrategicos": "◇",
        "kpi-institucionales": "◫",
    }

    def opcion(codigo, nombre):
        return html.Button(
            [
                html.Span(iconos[codigo], className="portada-boton-icono"),
                html.Span(nombre, className="portada-boton-texto"),
                html.Span("›", className="portada-boton-flecha"),
            ],
            id={"type": "indicator-home-card", "index": codigo},
            className="portada-boton indicador-portada-boton",
            n_clicks=0,
            type="button",
        )

    return html.Main(
        className="indicadores-pagina",
        children=[
            html.Section(
                className="indicadores-cabecera",
                children=[
                    html.Button(
                        "← Volver a la portada",
                        id="btn-volver-portada",
                        className="btn-volver-portada",
                        n_clicks=0,
                    ),
                    html.Div([
                        html.Span("INDICADORES", className="indicadores-etiqueta"),
                        html.H2("Seguimiento de indicadores institucionales"),
                        html.P("Seleccione el grupo de indicadores que desea consultar."),
                    ]),
                    html.Img(
                        src=app.get_asset_url("logo-minedec.png"),
                        className="indicadores-logo",
                        alt="MINEDEC",
                    ),
                ],
            ),
            html.Section(
                [opcion(codigo, nombre)
                 for codigo, nombre in SECCIONES_INDICADORES],
                className="indicadores-opciones",
            ),
        ],
    )


def menu_lateral(seccion_activa=None, vice_activo=None, indicador_activo=None):
    items = []
    # Cada ambiente conserva únicamente su propia navegación. Al consultar
    # indicadores no se muestran Visión Ejecutiva, Presupuesto ni Inventario.
    secciones_menu = (SECCIONES_INDICADORES if seccion_activa in
                      {codigo for codigo, _ in SECCIONES_INDICADORES}
                      else [(seccion_activa, dict(SECCIONES).get(
                          seccion_activa, "Módulo"))])
    for codigo, nombre in secciones_menu:
        activo = codigo == seccion_activa
        items.append(html.Button([
            html.Span(ICONOS[codigo], className="side-icon"), html.Span(nombre),
            html.Span("⌄" if activo else "", className="side-chevron")
        ], id={"type": "side-section", "index": codigo}, n_clicks=0,
           className="side-section active" if activo else "side-section"))
        if activo and codigo in SECCIONES_CON_DATOS:
            df, _ = obtener_datos(codigo)
            cfg = SECCIONES_CON_DATOS[codigo]
            if not df.empty and codigo in SECCIONES_INDICADORES_PLANOS:
                # Pocos indicadores: se listan directo, sin el paso intermedio
                # de escoger primero un viceministerio.
                indicadores = df[cfg["indicador_col"]].dropna().unique()
                items.append(html.Div([
                    html.Button([
                            html.Span(f"{j+1:02d}", className="indicator-index"), html.Span(indicador)
                        ], id={"type": "indicator-button", "index": indicador}, n_clicks=0, title=indicador,
                           className="indicator-button selected" if indicador == indicador_activo else "indicator-button")
                    for j, indicador in enumerate(indicadores)
                ], className="indicator-group open"))
            elif not df.empty:
                if codigo == "vision":
                    opciones_vice = (df[[cfg["vice_col"], "Orden sección"]].drop_duplicates()
                                     .sort_values("Orden sección")[cfg["vice_col"]].tolist())
                else:
                    opciones_vice = sorted(df[cfg["vice_col"]].dropna().unique())
                    if codigo == "presupuesto":
                        opciones_vice = [
                            x for x in opciones_vice
                            if _normalizar_encabezado_presupuesto(x)
                            not in {"SIN_CLASIFICACION", "SIN_CLASIFICAR", "NO_APLICA"}
                        ]
                        opciones_vice = sorted(
                            opciones_vice,
                            key=lambda x: (
                                str(x).strip().lower() == "coordinaciones generales",
                                str(x).lower(),
                            ),
                        )
                        opciones_vice = ["General"] + opciones_vice
                for vice in opciones_vice:
                    abierto = vice == vice_activo
                    items.append(html.Button([
                        html.Span("▾" if abierto else "▸"), html.Span(vice)
                    ], id={"type": "vice-button", "index": vice}, n_clicks=0,
                       className="vice-button selected" if abierto else "vice-button", title=vice))
                    if codigo in {"vision", "presupuesto"}:
                        continue
                    indicadores = df.loc[df[cfg["vice_col"]] == vice, cfg["indicador_col"]].unique()
                    items.append(html.Div([
                        html.Button([
                                html.Span(f"{j+1:02d}", className="indicator-index"), html.Span(indicador)
                            ], id={"type": "indicator-button", "index": indicador}, n_clicks=0, title=indicador,
                               className="indicator-button selected" if indicador == indicador_activo else "indicator-button")
                        for j, indicador in enumerate(indicadores)
                    ], id={"type": "indicator-group", "index": vice},
                       className="indicator-group open" if abierto else "indicator-group"))
    return html.Aside([
        html.Img(src=app.get_asset_url("logo-minedec.png"), className="sidebar-logo"),
        html.P("PANEL INSTITUCIONAL", className="sidebar-kicker"),
        html.Button([html.Span("←"), html.Span("Volver al panel principal")],
                    id="side-back", className="side-back", n_clicks=0),
        html.Div(items, className="sidebar-menu"),
    ], className="module-sidebar")


def primer_texto(grupo, columna, defecto="No registrado"):
    valores = grupo[columna].dropna()
    if valores.empty:
        return defecto
    texto = str(valores.iloc[0]).strip()
    return texto if texto and texto.lower() != "nan" else defecto


def formato_valor(valor):
    """Formato general: enteros sin decimales y cifras decimales con dos."""
    if pd.isna(valor):
        return "No registrado"
    numero = float(valor)
    if numero.is_integer():
        return f"{int(numero):,}".replace(",", ".")
    return f"{numero:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def formato_conteo(valor):
    """Corrige conteos como 1.693 cuando representan 1.693 unidades."""
    if pd.isna(valor):
        return "No registrado"
    numero = float(valor)
    if 0 < abs(numero) < 10 and not numero.is_integer():
        numero *= 1000
    return f"{int(round(numero)):,}".replace(",", ".")


def fecha_ficha(grupo, columna):
    """Muestra una fecha del Excel como dd/mm/aaaa, incluso si aún está vacía."""
    if columna not in grupo.columns:
        return "No registrada"
    valores = grupo[columna].dropna()
    valores = valores[valores.astype(str).str.strip().ne("")]
    if valores.empty:
        return "No registrada"
    valor = valores.iloc[0]
    try:
        if isinstance(valor, (int, float)) and 20000 < float(valor) < 80000:
            fecha = pd.to_datetime(valor, unit="D", origin="1899-12-30")
        else:
            fecha = pd.to_datetime(valor, errors="coerce", dayfirst=True)
        if pd.notna(fecha):
            return fecha.strftime("%d/%m/%Y")
    except (TypeError, ValueError, OverflowError):
        pass
    return str(valor).strip() or "No registrada"


def etiqueta_periodo(valor):
    """Muestra el período tal cual viene en el Excel: año simple (2025) o
    rango de año lectivo (2024-2025), sin forzar una conversión numérica."""
    if pd.isna(valor):
        return "No registrado"
    if isinstance(valor, (int, float)) and float(valor).is_integer():
        return str(int(valor))
    return str(valor).strip()


def ficha(etiqueta, valor, clase=""):
    return html.Div([html.Span(etiqueta), html.Strong(valor)], className=f"detail-item {clase}".strip())


def ficha_con_periodo(etiqueta, valor, periodo):
    """Muestra el año de referencia junto a numeradores y denominadores."""
    contenido = [html.Span(etiqueta), html.Strong(valor)]
    if pd.notna(periodo) and str(periodo).strip():
        contenido.append(
            html.Small(
                f"Año de referencia: {etiqueta_periodo(periodo)}",
                className="metric-year",
            )
        )
    return html.Div(contenido, className="detail-item count-card")


def normalizar_latex(texto, ecuacion_principal=False):
    """Adapta el LaTeX del Excel al formato que interpreta dcc.Markdown."""
    texto = str(texto)

    # MathJax no implementa entornos de documento como itemize. Los convertimos
    # a viñetas Markdown y conservamos el LaTeX en línea de cada definición.
    texto = re.sub(r"\\begin\s*\{itemize\}", "", texto, flags=re.IGNORECASE)
    texto = re.sub(r"\\end\s*\{itemize\}", "", texto, flags=re.IGNORECASE)
    texto = re.sub(r"(?m)^\s*\\item\s*", "- ", texto)

    # dcc.Markdown procesa de forma consistente los delimitadores con dólares.
    # Los reemplazos deben ignorar expresiones como ``\\[9pt]``: en LaTeX eso
    # es un salto de fila con separación vertical, no el inicio de una fórmula.
    texto = re.sub(r"(?<!\\)\\\(", "$ ", texto)
    texto = re.sub(r"(?<!\\)\\\)", " $", texto)
    texto = re.sub(r"(?<!\\)\\\[", "$$", texto)
    texto = re.sub(r"(?<!\\)\\\]", "$$", texto)

    # Corrige comandos escritos en mayúsculas en algunas celdas del Excel.
    comandos = {
        "sum": "sum", "frac": "frac", "times": "times", "cdot": "cdot",
        "sqrt": "sqrt", "left": "left", "right": "right", "big": "Big",
        "text": "text", "mathrm": "mathrm", "operatorname": "operatorname",
    }

    def corregir_comando(coincidencia):
        comando = coincidencia.group(1)
        corregido = comandos.get(comando.lower())
        return rf"\{corregido}" if corregido else coincidencia.group(0)

    texto = re.sub(r"\\([A-Za-z]+)", corregir_comando, texto)

    # No se reescriben siglas ni identificadores. Las fórmulas de los tres
    # Excel ya son LaTeX válido y MathJax debe recibirlas tal como fueron
    # registradas. Transformaciones generales sobre textos como NOPP_t,
    # Pe4Egb_{S,M} o RU_{\text{previos},t} terminaban actuando también dentro
    # de \mathrm, \operatorname y \text, duplicando comandos y superponiendo
    # letras, subíndices y superíndices.

    # En el apartado Fórmula, cada expresión ocupa una línea independiente.
    if ecuacion_principal:
        texto = re.sub(
            r"(?<!\$)\$(?!\$)(.+?)(?<!\$)\$(?!\$)",
            r"$$\1$$",
            texto,
            flags=re.DOTALL,
        )
    return texto


def ficha_multilinea(etiqueta, texto, clase="full"):
    """Pestaña desplegable, estilo documento, para fórmulas LaTeX extensas."""
    if texto is None or (isinstance(texto, float) and pd.isna(texto)) or not str(texto).strip():
        return html.Div(
            [html.Span(etiqueta), html.Strong("No registrada")],
            className=f"detail-item {clase}".strip(),
        )

    texto_latex = str(texto).strip()

    # Presentación editorial para las medianas de comportamiento sedentario.
    # En la matriz vienen como dos expresiones separadas y muy cargadas de
    # subíndices. Se unifican en una sola función por casos, como aparecería en
    # un documento técnico o PDF, sin modificar las definiciones posteriores.
    edad_mediana = re.search(r"Med_\{cs\}\^\{(\d+\s*-\s*\d+)\}", texto_latex, re.IGNORECASE)
    donde_mediana = re.search(r"\\textbf\{\s*(?:Dónde|Donde)\s*:\s*\}", texto_latex,
                              flags=re.IGNORECASE)
    if edad_mediana and donde_mediana:
        edad = edad_mediana.group(1).replace(" ", "")
        definiciones = texto_latex[donde_mediana.start():]
        ecuacion = rf"""\textbf{{Fórmula:}}

$$
\operatorname{{Med}}_{{cs}}^{{{edad}}}=
\begin{{cases}}
\dfrac{{\mathrm{{cs}}_{{n/2}}^{{{edad}}}+\mathrm{{cs}}_{{(n/2)+1}}^{{{edad}}}}}{{2}},
& \text{{si }} n \text{{ es par}},\\[9pt]
\mathrm{{cs}}_{{(n+1)/2}}^{{{edad}}},
& \text{{si }} n \text{{ es impar}}.
\end{{cases}}
$$"""
        texto_latex = ecuacion + "\n\n" + definiciones
    bloques = []
    en_formula = False
    for bloque in texto_latex.split("\n\n"):
        bloque = bloque.strip()
        if not bloque:
            continue
        titulo = re.fullmatch(r"\\textbf\{\s*(Fórmula|Formula)\s*:\s*\}", bloque,
                              flags=re.IGNORECASE)
        donde = re.fullmatch(r"\\textbf\{\s*(Dónde|Donde)\s*:\s*\}", bloque,
                             flags=re.IGNORECASE)
        if titulo:
            bloques.append(html.Strong("Fórmula:", className="formula-section-title"))
            en_formula = True
        elif donde:
            bloques.append(html.Strong("Donde:", className="formula-section-title"))
            en_formula = False
        else:
            bloques.append(dcc.Markdown(
                normalizar_latex(bloque, ecuacion_principal=en_formula),
                mathjax=True,
                className="formula-text formula-equation" if en_formula else "formula-text",
                dangerously_allow_html=False,
            ))

    return html.Details([
        html.Summary([
            html.Span("∑", className="formula-accordion-icon"),
            html.Strong(etiqueta),
            html.Span("Ver fórmula", className="formula-accordion-help"),
            html.Span("⌄", className="formula-accordion-chevron"),
        ], className="formula-static-header"),
        html.Div(bloques, className="formula-accordion-body formula-content"),
    ], className=f"detail-item {clase} formula-static formula-accordion".strip())


def ficha_metrica(etiqueta, valor, año, unidad):
    """Presenta cifra y período; la unidad ya consta en su ficha independiente."""
    unidad_texto = str(unidad or "").strip()
    sufijo = "%" if "porcentaje" in unidad_texto.lower() else ""
    return html.Div([
        html.Span(etiqueta),
        html.Div([html.Strong(formato_valor(valor)), html.B(sufijo)], className="metric-value"),
        html.Small(f"Período de referencia: {etiqueta_periodo(año)}", className="metric-year"),
    ], className="detail-item highlight metric-card")


def ficha_metrica_periodo(etiqueta, valor, periodo, unidad):
    """Variante para KPI: el período puede ser un año simple o un año lectivo (2024-2025)."""
    unidad_texto = str(unidad or "").strip()
    sufijo = "%" if "porcentaje" in unidad_texto.lower() else ""
    return html.Div([
        html.Span(etiqueta),
        html.Div([html.Strong(formato_valor(valor)), html.B(sufijo)], className="metric-value"),
        html.Small(f"Período: {etiqueta_periodo(periodo)}", className="metric-year"),
    ], className="detail-item highlight metric-card")


def clase_alerta(alerta):
    if pd.isna(alerta):
        return "status-empty"
    texto = str(alerta).lower()
    if "sobre" in texto: return "status-over"
    if "igual" in texto: return "status-equal"
    if "incumpl" in texto: return "status-under"
    return "status-empty"


def formato_texto_barra(valor, unidad):
    """Etiqueta de una barra: agrega '%' cuando la unidad es Porcentaje, para
    no mostrar '90' en un gráfico que representa 90%."""
    if pd.isna(valor):
        return None
    texto = formato_valor(valor)
    if "porcentaje" in str(unidad or "").lower():
        texto = f"{texto}%"
    return f"<b>{texto}</b>"


# ---------------------------------------------------------------------------
# Plan Nacional de Desarrollo · gráfico y ficha
# ---------------------------------------------------------------------------
def crear_grafico_pnd(grupo):
    grupo = grupo.sort_values("_orden_periodo")
    años = list(grupo["Año"].map(etiqueta_periodo))
    unidad = primer_texto(grupo, "Unidad de medida", "")

    # Posiciones reales dentro de cada año. Una barra sola queda centrada y
    # ancha; cuando coinciden varias, se separan sin amontonar sus etiquetas.
    posiciones = list(range(len(grupo)))

    fig = go.Figure()
    fig.add_bar(x=posiciones, y=grupo["Línea base"], name="Línea base",
                marker=dict(color="#80BBBF", line=dict(color="#446381", width=1.4)),
                text=[formato_texto_barra(v, unidad) for v in grupo["Línea base"]], textposition="outside",
                textfont=dict(size=10, color="#232D5A", family="Arial"), constraintext="none",
                width=.34, offset=-.17, cliponaxis=False, customdata=años,
                hovertemplate="Período %{customdata}<br>Línea base: %{y:,.2f}<extra></extra>")
    fig.add_bar(x=posiciones, y=grupo["Meta"], name="Meta",
                marker=dict(color="#F1B620", line=dict(color="#C88F00", width=1.4)),
                text=[formato_texto_barra(v, unidad) for v in grupo["Meta"]], textposition="outside",
                textfont=dict(size=10, color="#232D5A", family="Arial"), constraintext="none",
                width=.34, offset=-.36, cliponaxis=False, customdata=años,
                hovertemplate="Período %{customdata}<br>Meta: %{y:,.2f}<extra></extra>")
    fig.add_bar(x=posiciones, y=grupo["Estimador"], name="Estimador",
                marker=dict(color="#4F449A", line=dict(color="#232D5A", width=1.4)),
                text=[formato_texto_barra(v, unidad) for v in grupo["Estimador"]],
                textposition="outside",
                textfont=dict(size=10, color="#232D5A", family="Arial"), constraintext="none",
                width=.34, offset=.02, cliponaxis=False, customdata=años,
                hovertemplate="Período %{customdata}<br>Estimador: %{y:,.2f}<extra></extra>")
    candidatos = pd.concat([grupo["Línea base"], grupo["Meta"], grupo["Estimador"]]).dropna()
    techo = candidatos.max() * 1.18 if not candidatos.empty else None
    fig.update_layout(
        barmode="overlay", height=380,
        margin=dict(l=55, r=25, t=60, b=45), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="#fff",
        font=dict(family="Arial", color="#232D5A", size=13),
        title=dict(text="Evolución del indicador", x=.02, font=dict(size=20)),
        legend=dict(orientation="h", y=1.13, x=1, xanchor="right"),
        xaxis=dict(title="Año / período lectivo", showgrid=False, fixedrange=True,
                   tickmode="array", tickvals=posiciones, ticktext=años,
                   range=[-.55, len(años) - .45]),
        yaxis=dict(title=unidad or "Valor", gridcolor="#E9EBF3",
                   zeroline=False, fixedrange=True, range=[0, techo] if techo else None),
    )
    return fig


def detalle_indicador_pnd(indicador):
    grupo = DATA_PND.loc[DATA_PND["NOMBRE DEL INDICADOR"] == indicador].sort_values("_orden_periodo").copy()
    unidad = primer_texto(grupo, "Unidad de medida", "")
    linea = grupo.dropna(subset=["Línea base"])
    metas = grupo.dropna(subset=["Meta"])

    campos = [ficha("Definición", primer_texto(grupo, "Definición"), "full")]
    formula_documento = None
    if "Fórmula de Cálculo" in grupo.columns:
        formula = primer_texto(grupo, "Fórmula de Cálculo", "")
        if formula:
            formula_documento = ficha_multilinea("Fórmula de cálculo", formula, "full")
    metricas = []
    if linea.empty:
        metricas.append(ficha("Línea base", "No registrada", "highlight"))
    else:
        metricas.append(ficha_metrica("Línea base", linea.iloc[0]["Línea base"], linea.iloc[0]["Año"], unidad))
    if metas.empty:
        metricas.append(ficha("Meta final", "No registrada", "highlight"))
    else:
        metricas.append(ficha_metrica("Meta final", metas.iloc[-1]["Meta"], metas.iloc[-1]["Año"], unidad))
    campos += [
        ficha("Periodicidad", primer_texto(grupo, "Periodicidad")),
        ficha("Unidad de medida", primer_texto(grupo, "Unidad de medida")),
        ficha("Fecha de reporte", fecha_ficha(grupo, "Fecha de Transferencia")),
    ]
    campos.append(ficha("Fuente de datos", primer_texto(grupo, "Fuente de datos"), "full source"))
    columnas_ficha = max(1, len(campos) - 2)

    estados = []
    for _, fila in grupo.iterrows():
        año = etiqueta_periodo(fila["Año"])
        clase = clase_alerta(fila["Alerta"])
        alerta = str(fila["Alerta"]) if pd.notna(fila["Alerta"]) else "Sin información registrada"
        estados.append(html.Div([
            html.Span(str(año), className="status-year-label"),
            html.Button("i", id={"type": "status-year", "index": año},
                        className=f"status-mini {clase}", n_clicks=0,
                        title=f"{año}: {alerta}", **{"aria-label": f"Consultar información de {año}"})
        ], className="status-year-item"))

    return html.Div([
        html.Div([html.P("PLAN NACIONAL DE DESARROLLO", className="content-kicker"),
                  html.H1(indicador), html.P(primer_texto(grupo, "VICEMINISTERIO"), className="content-subtitle")],
                 className="content-heading"),
        html.Div([
            html.Section([html.H2("Ficha del indicador"),
                          html.Div(campos, className=f"details-grid cols-{columnas_ficha}")],
                         className="indicator-detail-card", style={"marginTop": "18px"}),
            html.Section([
                html.Div(metricas, className="metric-summary-row"),
                dcc.Graph(figure=crear_grafico_pnd(grupo), config={"displayModeBar": False, "responsive": True},
                          className="indicator-chart", style={"width": "100%", "height": "380px"}),
                html.P("Seleccione el botón de un período para consultar su cumplimiento y observación.",
                       className="status-help"),
                html.Div(estados, className="status-row",
                         style={"gridTemplateColumns": f"repeat({len(estados)}, minmax(0, 1fr))"}),
                html.Div("Seleccione un período para consultar su observación.", id="observation-area",
                         className="observation-area"),
            ], className="chart-card"),
        ], className="indicator-workspace"),
        html.Section(formula_documento, className="formula-document formula-document-bottom")
        if formula_documento else None,
    ], className="pnd-content")


def contenido_pnd(vice=None, indicador=None):
    if ERROR_PND:
        return html.Div([html.H2("Base no disponible"), html.P(ERROR_PND)], className="load-error")
    elif indicador:
        return detalle_indicador_pnd(indicador)
    elif vice:
        return html.Div([html.H1(vice), html.P("Seleccione uno de sus indicadores en el menú lateral.")],
                        className="module-welcome")
    return html.Div([html.P("PANEL INSTITUCIONAL", className="content-kicker"),
                     html.H1("Plan Nacional de Desarrollo"),
                     html.P("Seleccione un viceministerio y luego el indicador que desea consultar.")],
                    className="module-welcome")


# ---------------------------------------------------------------------------
# KPI's Estratégicos y KPI's Institucionales · gráfico y ficha compartidos
# (misma estructura tipo PND: Línea base / Meta / Estimador por período,
# con eje de dos niveles: mes arriba, año agrupado debajo).
# ---------------------------------------------------------------------------
def preparar_periodos_visuales(grupo):
    """Ordena y abrevia períodos para evitar etiquetas inclinadas o saturadas."""
    grupo = grupo.copy()
    # Un indicador debe tener una sola categoría por año y mes/semestre.
    # Se conserva la primera fila registrada y se evita repetir barras y botones.
    grupo = grupo.drop_duplicates(subset=["Año", "Mes"], keep="first").copy()
    grupo["_orden_visual"] = grupo["_orden_periodo"]
    grupo["Periodo_corto"] = (grupo["Periodo"].astype(str)
                                .str.replace("Enero-Junio", "Ene-Jun", regex=False)
                                .str.replace("Julio-Diciembre", "Jul-Dic", regex=False))
    grupo["Periodo_grafico"] = grupo["Periodo_corto"]

    es_semestral = grupo["Mes"].dropna().astype(str).str.contains("-", regex=False).any()
    if es_semestral:
        años_con_semestre = set(grupo.loc[grupo["Mes"].notna(), "Año"].astype(str))
        es_cierre = grupo["Mes"].isna() & grupo["Año"].astype(str).isin(años_con_semestre)
        # El cierre anual se ubica después de Julio-Diciembre, no antes de los semestres.
        grupo.loc[es_cierre, "_orden_visual"] = (
            grupo.loc[es_cierre, "Año"].map(orden_periodo) * 100 + 13
        )
        grupo.loc[es_cierre, "Periodo_corto"] = "Línea base " + grupo.loc[es_cierre, "Año"].astype(str)

    grupo = grupo.sort_values("_orden_visual", kind="stable")

    # Evita que registros repetidos se dibujen en la misma coordenada. Plotly
    # necesita una clave X única; la etiqueta visible puede seguir siendo amigable.
    grupo["_ocurrencia_periodo"] = grupo.groupby("Periodo", sort=False).cumcount() + 1
    grupo["_total_periodo"] = grupo.groupby("Periodo")["Periodo"].transform("size")
    repetido = grupo["_total_periodo"] > 1
    grupo.loc[repetido, "Periodo_corto"] = (
        grupo.loc[repetido, "Periodo_corto"]
        + " · " + grupo.loc[repetido, "_ocurrencia_periodo"].astype(str)
    )
    grupo["Periodo_id"] = (
        grupo["Periodo"].astype(str) + "__" + grupo["_ocurrencia_periodo"].astype(str)
    )

    # Plotly interpreta <br> como salto de línea y mantiene horizontales las etiquetas.
    grupo["Periodo_grafico"] = grupo["Periodo_corto"].str.replace(
        r"^(Ene-Jun|Jul-Dic|[A-ZÁÉÍÓÚ][a-záéíóú]{2}|Línea base)\s+(.+)$",
        r"\1<br>\2", regex=True
    )
    return grupo


def crear_grafico_periodo(grupo):
    grupo = preparar_periodos_visuales(grupo)
    unidad = primer_texto(grupo, "Unidad de medida", "")
    # En series mensuales hay hasta 13 categorías. Una escala tipográfica
    # ligeramente menor evita cruces en portátiles sin sacrificar legibilidad.
    etiqueta_size = 9 if len(grupo) >= 10 else 10
    claves_periodo = list(grupo["Periodo_id"])
    periodos = list(grupo["Periodo_grafico"])
    periodos_completos = list(grupo["Periodo"])

    fig = go.Figure()
    fig.add_bar(x=claves_periodo, y=grupo["Línea base"], name="Línea base",
                marker=dict(color="#80BBBF", line=dict(color="#446381", width=1.4)),
                text=[formato_texto_barra(v, unidad) for v in grupo["Línea base"]],
                textposition="outside", textfont=dict(size=etiqueta_size, color="#232D5A", family="Arial"), constraintext="none",
                width=.34, offset=-.17, cliponaxis=False,
                customdata=periodos_completos,
                hovertemplate="%{customdata}<br>Línea base: %{y:,.2f}<extra></extra>")
    fig.add_bar(x=claves_periodo, y=grupo["Meta"], name="Meta",
                marker=dict(color="#F1B620", line=dict(color="#C88F00", width=1.4)),
                text=[formato_texto_barra(v, unidad) for v in grupo["Meta"]],
                textposition="outside", textfont=dict(size=etiqueta_size, color="#232D5A", family="Arial"),
                constraintext="none", width=.34, offset=-.36, cliponaxis=False,
                customdata=periodos_completos,
                hovertemplate="%{customdata}<br>Meta: %{y:,.2f}<extra></extra>")
    fig.add_bar(x=claves_periodo, y=grupo["Estimador"], name="Ejecutado",
                marker=dict(color="#4F449A", line=dict(color="#232D5A", width=1.4)),
                text=[formato_texto_barra(v, unidad) for v in grupo["Estimador"]],
                textposition="outside", textfont=dict(size=etiqueta_size, color="#232D5A", family="Arial"),
                constraintext="none", width=.34, offset=.02, cliponaxis=False,
                customdata=periodos_completos,
                hovertemplate="%{customdata}<br>Ejecutado: %{y:,.2f}<extra></extra>")

    candidatos = pd.concat([grupo["Línea base"], grupo["Meta"], grupo["Estimador"]]).dropna()
    techo = candidatos.max() * 1.22 if not candidatos.empty else None

    fig.update_layout(
        barmode="overlay", bargap=.20, height=380, autosize=True,
        margin=dict(l=55, r=15, t=65, b=70), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="#fff",
        font=dict(family="Arial", color="#232D5A", size=13),
        title=dict(text="Evolución del indicador", x=.02, font=dict(size=20)),
        legend=dict(orientation="h", y=1.13, x=1, xanchor="right"),
        # Categoría simple con orden forzado explícitamente: es la forma
        # confiable de garantizar el orden cronológico. El eje "multicategory"
        # ignoraba tanto categoryorder="trace" como categoryarray.
        xaxis=dict(title="Período", showgrid=False, type="category", categoryorder="array",
                   categoryarray=claves_periodo, tickmode="array",
                   tickvals=claves_periodo, ticktext=periodos,
                   fixedrange=True, tickangle=0,
                   automargin=True, tickfont=dict(size=etiqueta_size)),
        yaxis=dict(title=unidad or "Valor", gridcolor="#E9EBF3",
                   zeroline=False, fixedrange=True, range=[0, techo] if techo else None),
    )
    return fig


def detalle_indicador_periodo(df_fuente, indicador, kicker, obs_area_id, status_type):
    grupo = df_fuente.loc[df_fuente["NOMBRE DEL INDICADOR"] == indicador].sort_values("_orden_periodo").copy()
    grupo = preparar_periodos_visuales(grupo)
    unidad = primer_texto(grupo, "Unidad de medida", "")
    linea = grupo.dropna(subset=["Línea base"])
    metas = grupo.dropna(subset=["Meta"])
    periodicidad = primer_texto(grupo, "Periodicidad")
    # Si la matriz registra rangos como Enero-Junio y Julio-Diciembre,
    # cada rango representa una categoría semestral completa.
    if grupo["Mes"].dropna().astype(str).str.contains("-", regex=False).any():
        periodicidad = "Semestral"

    campos = [ficha("Definición", primer_texto(grupo, "Definición"), "full")]
    formula_documento = None
    if "Fórmula de Cálculo" in grupo.columns:
        formula = primer_texto(grupo, "Fórmula de Cálculo", "")
        if formula:
            formula_documento = ficha_multilinea("Fórmula de cálculo", formula, "full")
    metricas = []
    if linea.empty:
        metricas.append(ficha("Línea base", "No registrada", "highlight"))
    else:
        metricas.append(ficha_metrica_periodo("Línea base", linea.iloc[0]["Línea base"],
                                               linea.iloc[0]["Periodo"], unidad))
    if metas.empty:
        metricas.append(ficha("Meta", "No registrada", "highlight"))
    else:
        metricas.append(ficha_metrica_periodo("Meta", metas.iloc[-1]["Meta"], metas.iloc[-1]["Periodo"], unidad))
    campos += [
        ficha("Periodicidad", periodicidad),
        ficha("Unidad de medida", primer_texto(grupo, "Unidad de medida")),
        ficha("Fecha de reporte", fecha_ficha(grupo, "Fecha de Transferencia")),
    ]
    if "TIPO INDICADOR" in grupo.columns:
        campos.append(ficha("Tipo de indicador", primer_texto(grupo, "TIPO INDICADOR")))
    if "ENCARGADO" in grupo.columns:
        campos.append(ficha("Encargado", primer_texto(grupo, "ENCARGADO")))
    campos.append(ficha("Fuente de datos", primer_texto(grupo, "Fuente de datos"), "full source"))
    columnas_ficha = max(1, len(campos) - 2)

    estados = []
    for _, fila in grupo.iterrows():
        periodo = fila["Periodo"]
        periodo_corto = fila["Periodo_corto"]
        periodo_id = fila["Periodo_id"]
        clase = clase_alerta(fila["Alerta"])
        alerta = str(fila["Alerta"]) if pd.notna(fila["Alerta"]) else "Sin información registrada"
        estados.append(html.Div([
            html.Span(periodo_corto, className="status-year-label"),
            html.Button("i", id={"type": status_type, "index": periodo_id},
                        className=f"status-mini {clase}", n_clicks=0,
                        title=f"{periodo}: {alerta}", **{"aria-label": f"Consultar información de {periodo}"})
        ], className="status-year-item"))

    return html.Div([
        html.Div([html.P(kicker, className="content-kicker"),
                  html.H1(indicador), html.P(primer_texto(grupo, "VICEMINISTERIO"), className="content-subtitle")],
                 className="content-heading"),
        html.Div([
            html.Section([html.H2("Ficha del indicador"),
                          html.Div(campos, className=f"details-grid cols-{columnas_ficha}")],
                         className="indicator-detail-card", style={"marginTop": "18px"}),
            html.Section([
                html.Div(metricas, className="metric-summary-row"),
                dcc.Graph(figure=crear_grafico_periodo(grupo), config={"displayModeBar": False, "responsive": True},
                          className="indicator-chart", style={"width": "100%", "height": "380px"}),
                html.P("Seleccione el botón de un período para consultar su cumplimiento y observación.",
                       className="status-help"),
                html.Div(estados, className="status-row",
                         style={"gridTemplateColumns": f"repeat({len(estados)}, minmax(0, 1fr))"}),
                html.Div("Seleccione un período para consultar su observación.", id=obs_area_id,
                         className="observation-area"),
            ], className="chart-card"),
        ], className="indicator-workspace"),
        html.Section(formula_documento, className="formula-document formula-document-bottom")
        if formula_documento else None,
    ], className="pnd-content")


def detalle_indicador_kpi(indicador):
    return detalle_indicador_periodo(DATA_KPI, indicador, "KPI´S ESTRATÉGICOS",
                                      "observation-area-kpi", "status-periodo-kpi")


def contenido_kpi(vice=None, indicador=None):
    if ERROR_KPI:
        return html.Div([html.H2("Base no disponible"), html.P(ERROR_KPI)], className="load-error")
    elif indicador:
        return detalle_indicador_kpi(indicador)
    elif vice:
        return html.Div([html.H1(vice), html.P("Seleccione uno de sus indicadores en el menú lateral.")],
                        className="module-welcome")
    return html.Div([html.P("PANEL INSTITUCIONAL", className="content-kicker"),
                     html.H1("KPI´s Estratégicos"),
                     html.P("Seleccione el indicador que desea consultar en el menú lateral.")],
                    className="module-welcome")


def detalle_indicador_kpi_inst(indicador):
    return detalle_indicador_periodo(DATA_KPI_INST, indicador, "KPI´S INSTITUCIONALES",
                                      "observation-area-kpi-inst", "status-periodo-inst")


def contenido_kpi_inst(vice=None, indicador=None):
    if ERROR_KPI_INST:
        return html.Div([html.H2("Base no disponible"), html.P(ERROR_KPI_INST)], className="load-error")
    elif indicador:
        return detalle_indicador_kpi_inst(indicador)
    elif vice:
        return html.Div([html.H1(vice), html.P("Seleccione uno de sus indicadores en el menú lateral.")],
                        className="module-welcome")
    return html.Div([html.P("PANEL INSTITUCIONAL", className="content-kicker"),
                     html.H1("KPI´s Institucionales"),
                     html.P("Seleccione el indicador que desea consultar en el menú lateral.")],
                    className="module-welcome")


# ---------------------------------------------------------------------------
# Visión Ejecutiva · menú de viceministerios (prototipo "MINEDEC · Prototipo
# Ejecutivo"). Al entrar se presentan los 5 botones; "Gestión Educativa" arma
# las tarjetas de cobertura con las 4 bases reales, sin montos; el resto de
# viceministerios no tiene fuente todavía y se muestra "en construcción".
# ---------------------------------------------------------------------------
def vision_tabs_nav(activo):
    """Menú de los 5 viceministerios. En la pantalla de entrada (sin
    viceministerio elegido) se ve como lista vertical de botones; una vez
    elegido uno, se compacta en una franja superior tipo pestañas."""
    compacto = activo is not None
    estilo_nav = None if compacto else {
        "display": "grid", "gridTemplateColumns": "repeat(5, minmax(170px, 1fr))",
        "gap": "9px", "width": "100%", "maxWidth": "none", "overflowX": "auto",
        "paddingBottom": "3px",
    }
    estilo_fila = None if compacto else {
        "minWidth": "170px", "padding": "4px", "display": "flex",
        "border": "1px solid #e1e4ed", "borderRadius": "12px", "background": "#fff",
        "boxShadow": "0 6px 16px rgba(24,37,87,.05)",
    }
    estilo_boton = None if compacto else {
        "width": "100%", "minHeight": "54px", "padding": "10px 12px",
        "border": "none", "borderRadius": "999px", "background": "#e9ecfb",
        "color": "#3b3f91", "fontSize": "14px", "fontWeight": "800",
        "textAlign": "center", "cursor": "pointer",
    }
    return html.Nav(
        [html.Div(
            html.Button(nombre, id={"type": "vision-tab", "index": codigo}, n_clicks=0,
                        className="vision-vice-pill active" if codigo == activo else "vision-vice-pill",
                        type="button", style=estilo_boton),
            className="vision-vice-row", style=estilo_fila,
        ) for codigo, nombre in VISION_TABS],
        className="vision-vice-menu vision-vice-menu-compacta" if compacto else "vision-vice-menu",
        style=estilo_nav,
        **{"aria-label": "Viceministerios de Visión Ejecutiva"},
    )


def tarjeta_gestion_educativa(codigo, titulo, subtitulo, valor, detalles=None,
                              unidad=None, filtros=None, contenido_id=None, valor_id=None):
    return html.Article([
        html.Span(className="vision-card-icon-mark"),
        html.Span(titulo, className="vision-card-label"),
        html.Div([
            html.Strong(
                valor,
                className="vision-card-total",
                **({"id": valor_id} if valor_id is not None else {}),
            ),
            html.Span(unidad, className="vision-card-unit") if unidad else None,
        ], className="vision-card-value-row"),
        html.Div(
            detalles,
            className="vision-card-details",
            **({"id": contenido_id} if contenido_id is not None else {}),
        ) if detalles else None,
        html.P(subtitulo) if subtitulo else None,
        html.Div(filtros, className="vision-card-filters") if filtros else None,
    ], className=f"vision-exec-card {codigo}")


# Eje 2 · Educación Superior. La primera sección se alimenta de la matriz
# "Mapeo - Rendicion de cuentas.xlsx"; la segunda recoge la agenda entregada
# en las diapositivas y se presenta separadamente como cartera futura.
EJE2_PROGRAMAS_RESPALDO = [
    {"direccion": "Dirección de Admisión", "programa": "Cupos Aceptados en Educación Superior",
     "beneficiarios": 194000, "tipo": "Estudiantes"},
    {"direccion": "Dirección de Diseño y Evaluación de Política Pública de Fortalecimiento del Talento Humano",
     "programa": "BECAUSE HE IS NICE", "beneficiarios": 400000, "tipo": "Ciudadanos"},
    {"direccion": "Dirección de Gestión Académica de Tercer y Cuarto Nivel",
     "programa": "ValidaTec (Validación de Trayectoria para obtener tercer nivel)",
     "beneficiarios": 120, "tipo": "Docentes"},
    {"direccion": "Dirección de Cooperación y Asuntos Internacionales",
     "programa": "Becas Técnicas y tecnológicas financiadas por aliados",
     "beneficiarios": 5000, "tipo": "Estudiantes"},
]

# Totales nacionales de Educación Superior mostrados en la portada
# institucional MINEDEC 2026. Se mantienen separados de los programas y
# servicios porque describen el tamaño del sistema, no sus beneficiarios.
DATOS_GENERALES_EDUCACION_SUPERIOR = {
    "estudiantes": 1004175,
    "estudiantes_itts": 138613,
    "estudiantes_uep": 865562,
    "docentes": 47145,
    "docentes_itts": 9320,
    "docentes_uep": 37825,
    "instituciones": 257,
    "instituciones_itts": 193,
    "instituciones_uep": 64,
}


def tarjetas_datos_educacion_superior():
    """Resumen nacional de estudiantes, docentes e instituciones superiores."""
    d = DATOS_GENERALES_EDUCACION_SUPERIOR

    def desglose(valor_itts, valor_uep):
        return [
            html.Div([
                html.Span("ITTS"),
                html.Strong(formato_valor(valor_itts)),
            ], className="vision-mini-stat"),
            html.Div([
                html.Span("UEP"),
                html.Strong(formato_valor(valor_uep)),
            ], className="vision-mini-stat secondary"),
        ]

    subtitulo = None
    return [
        tarjeta_gestion_educativa(
            "superior-instituciones", "Instituciones de Educación Superior", subtitulo,
            formato_valor(d["instituciones"]),
            detalles=desglose(d["instituciones_itts"], d["instituciones_uep"]),
        ),
        tarjeta_gestion_educativa(
            "superior-estudiantes", "Estudiantes", subtitulo,
            formato_valor(d["estudiantes"]),
            detalles=desglose(d["estudiantes_itts"], d["estudiantes_uep"]),
        ),
        tarjeta_gestion_educativa(
            "superior-docentes", "Docentes", subtitulo,
            formato_valor(d["docentes"]),
            detalles=desglose(d["docentes_itts"], d["docentes_uep"]),
        ),
    ]

EJE2_PROYECTOS_FUTUROS = [
    {"mes": "Octubre", "titulo": "Entrega de 595 becas",
     "descripcion": "Becas de especialización y posgrado para ampliar las oportunidades de formación.",
     "beneficiarios": 595, "inversion": 1785778.00},
    {"mes": "Octubre", "titulo": "Reconocimiento de títulos Ecuador–España",
     "descripcion": "Convenio de homologación y reconocimiento bilateral de títulos académicos.",
     "beneficiarios": None, "inversion": None},
    {"mes": "Octubre", "titulo": "Inicio de construcción de Casa U",
     "descripcion": "Espacios en Azuay, El Oro y Chimborazo junto con la Universidad de Cuenca, UTMACH y UNACH.",
     "beneficiarios": 1000, "inversion": 20136155.55},
    {"mes": "Octubre", "titulo": "Universidad Pública de Santo Domingo de los Tsáchilas",
     "descripcion": "Inicio de la construcción de la primera fase de la nueva universidad pública.",
     "beneficiarios": 243, "inversion": 5638691.00},
    {"mes": "Noviembre", "titulo": "Entrega de becas técnicas y tecnológicas",
     "descripcion": "Nuevas oportunidades de formación técnica y tecnológica para jóvenes.",
     "beneficiarios": 2550, "inversion": 5236916.00},
]


def cargar_programas_educacion_superior():
    """Lee únicamente las filas de Educación Superior de la matriz oficial."""
    archivo = buscar_archivo("Mapeo - Rendicion de cuentas")
    if archivo is None:
        return EJE2_PROGRAMAS_RESPALDO, "Se muestra el último corte disponible; publique el Excel para actualizarlo."
    try:
        df = pd.read_excel(archivo)
        columnas = {_sin_tildes(c): c for c in df.columns}
        campos = {"vice": columnas.get("viceministerio"), "direccion": columnas.get("direccion"),
                  "programa": columnas.get("programa o servicio"),
                  "beneficiarios": columnas.get("beneficiarios"),
                  "tipo": columnas.get("tipo beneficiario")}
        if any(v is None for v in campos.values()):
            raise ValueError("la matriz no conserva las cinco columnas esperadas")
        mascara = df[campos["vice"]].map(_sin_tildes).str.contains("educacion superior", na=False)
        programas = []
        for _, fila in df.loc[mascara].iterrows():
            beneficiarios = pd.to_numeric(fila[campos["beneficiarios"]], errors="coerce")
            programas.append({"direccion": str(fila[campos["direccion"]]).strip(),
                              "programa": str(fila[campos["programa"]]).strip(),
                              "beneficiarios": int(beneficiarios) if pd.notna(beneficiarios) else 0,
                              "tipo": str(fila[campos["tipo"]]).strip()})
        if not programas:
            raise ValueError("no se encontraron filas de Educación Superior")
        return programas, None
    except Exception as exc:
        return EJE2_PROGRAMAS_RESPALDO, f"Se muestra el último corte disponible ({exc})."


def cargar_programas_viceministerio(nombre_viceministerio):
    """Obtiene de la misma matriz los programas del viceministerio solicitado."""
    archivo = buscar_archivo("Mapeo - Rendicion de cuentas")
    if archivo is None:
        return [], "No se encontró la matriz de rendición de cuentas."
    try:
        df = pd.read_excel(archivo)
        columnas = {_sin_tildes(c): c for c in df.columns}
        campos = {"vice": columnas.get("viceministerio"), "direccion": columnas.get("direccion"),
                  "programa": columnas.get("programa o servicio"),
                  "beneficiarios": columnas.get("beneficiarios"),
                  "tipo": columnas.get("tipo beneficiario")}
        if any(v is None for v in campos.values()):
            raise ValueError("la matriz no conserva las cinco columnas esperadas")
        objetivo = _sin_tildes(nombre_viceministerio)
        mascara = df[campos["vice"]].map(_sin_tildes).eq(objetivo)
        datos = []
        for _, fila in df.loc[mascara].iterrows():
            beneficiarios = pd.to_numeric(fila[campos["beneficiarios"]], errors="coerce")
            datos.append({"direccion": str(fila[campos["direccion"]]).strip(),
                          "programa": str(fila[campos["programa"]]).strip(),
                          "beneficiarios": int(beneficiarios) if pd.notna(beneficiarios) else 0,
                          "tipo": str(fila[campos["tipo"]]).strip()})
        return datos, None if datos else f"No existen registros para {nombre_viceministerio}."
    except Exception as exc:
        return [], f"No se pudo leer el consolidado ({exc})."


def _stats_eje2(datos):
    resumen = [(formato_valor(len(datos)), "Programas y servicios"),
               (formato_valor(sum(d["beneficiarios"] for d in datos)), "Beneficiarios registrados"),
               (formato_valor(len({d["direccion"] for d in datos})), "Direcciones responsables"),
               (formato_valor(len({d["tipo"] for d in datos})), "Perfiles beneficiarios")]
    return html.Div([html.Div([html.Strong(v), html.Span(e)], className="eje1-stat") for v, e in resumen],
                    className="eje1-stats-grid eje2-stats-grid")


def _tarjetas_programas_eje2(datos):
    """Presenta cada programa con el mismo lenguaje visual de las cifras."""
    tarjetas = []
    colores = ["azul", "verde", "amarillo", "morado"]
    for i, dato in enumerate(datos):
        clave_programa = _sin_tildes(dato["programa"])
        es_dece = clave_programa.startswith("profesionales dece")
        es_comunidades = clave_programa.startswith("comunidades seguras")
        titulo = ("Profesionales DECE (Apoyo Psicológico)"
                  if es_dece else dato["programa"])
        if es_comunidades:
            unidad = "Comunidad Educativa"
        elif es_dece:
            unidad = "DECEs contratados"
        else:
            unidad = str(dato["tipo"]).strip()
        contenido = [
            html.Span(className="vision-card-icon-mark"),
            html.Span(titulo, className="vision-card-label"),
            html.Div([
                html.Strong(formato_valor(dato["beneficiarios"])),
                html.Span(unidad, className="unidad-sin-mayusculas" if es_dece else None),
                ], className="vision-card-value-row programa-card-resultado"),
        ]
        if es_dece:
            contenido.append(html.P("La brecha nacional se ha reducido al 45%",
                                    className="programa-card-destacado"))
        tarjetas.append(html.Article(
            contenido,
            className=f"vision-exec-card programa-resumen-card programa-{colores[i % len(colores)]}",
        ))
    return html.Div(tarjetas,
                    className="vision-exec-grid gestion-educativa-grid programas-servicios-grid")


def _cartera_futura_eje2(proyectos):
    stats = [(formato_valor(len(proyectos)), "Acciones programadas"),
             (formato_valor(sum(p["beneficiarios"] or 0 for p in proyectos)), "Beneficiarios directos"),
             (_moneda_corta(sum(p["inversion"] or 0 for p in proyectos)), "Inversión asociada")]
    tarjetas = []
    for i, p in enumerate(proyectos, 1):
        meta = []
        if p["beneficiarios"] is not None:
            meta.append(html.Span(
                f"{formato_valor(p['beneficiarios'])} beneficiarios",
                style={"padding": "6px 9px", "borderRadius": "7px", "color": "#473286",
                       "background": "#efedf8", "fontSize": "12px", "fontWeight": "900"}))
        if p["inversion"] is not None:
            meta.append(html.Span(
                _moneda_corta(p["inversion"]),
                style={"padding": "6px 9px", "borderRadius": "7px", "color": "#795b11",
                       "background": "#fff0c2", "fontSize": "12px", "fontWeight": "900"}))
        es_noviembre = p["mes"] == "Noviembre"
        tarjetas.append(html.Article([
            html.Div([
                html.Span(p["mes"], className="eje2-project-month",
                          style={"padding": "5px 10px", "borderRadius": "999px",
                                 "color": "#241259" if es_noviembre else "#ffffff",
                                 "background": "#f4b91e" if es_noviembre else "#503a98",
                                 "fontSize": "9px", "fontWeight": "900",
                                 "letterSpacing": ".08em", "textTransform": "uppercase"}),
                html.Span(f"{i:02d}", className="eje2-project-number",
                          style={"color": "#c9cde0", "fontSize": "29px", "fontWeight": "900"})
            ], className="eje2-project-top",
               style={"display": "flex", "alignItems": "center", "justifyContent": "space-between"}),
            html.H3(p["titulo"], style={"margin": "11px 0 6px", "color": "#17245b",
                                        "fontSize": "13px", "lineHeight": "1.2"}),
            html.P(p["descripcion"], style={"minHeight": "42px", "margin": "0 0 13px",
                                            "color": "#59627d", "fontSize": "11px",
                                            "lineHeight": "1.45"}),
            html.Div(meta, className="eje2-project-meta",
                     style={"display": "flex", "flexWrap": "wrap", "gap": "6px"}) if meta else
            html.Div("Hito normativo", className="eje2-project-meta eje2-project-meta-single",
                     style={"display": "inline-flex", "padding": "6px 9px", "borderRadius": "7px",
                            "color": "#473286", "background": "#efedf8", "fontSize": "10px",
                            "fontWeight": "800"}),
        ], className="eje2-project-card",
           style={"minWidth": "0", "padding": "13px", "border": "1px solid #e1e4ed",
                  "borderTop": f"4px solid {'#f4b91e' if es_noviembre else '#503a98'}",
                  "borderRadius": "13px", "background": "#fffdf7" if es_noviembre else "#ffffff",
                  "boxShadow": "0 7px 18px rgba(24, 37, 87, .07)"}))
    return html.Section([
        html.Div([
            html.Div([
                html.Strong(v, style={
                    "fontSize": "clamp(24px, 2vw, 34px)",
                    "whiteSpace": "nowrap", "letterSpacing": "-.02em"
                }),
                html.Span(e)
            ], className="eje1-stat") for i, (v, e) in enumerate(stats)
        ],
                 className="eje1-stats-grid eje2-future-stats"),
        html.Div(tarjetas, className="eje2-project-grid",
                 style={"display": "grid", "gridTemplateColumns": "repeat(5, minmax(230px, 1fr))",
                        "gap": "10px", "alignItems": "stretch", "overflowX": "auto",
                        "paddingBottom": "6px"}),
        html.P("Cifras consolidadas de la agenda presentada para octubre y noviembre. El hito de "
               "reconocimiento de títulos no registra beneficiarios ni inversión en la fuente.",
               className="eje2-source-note")], className="eje2-future-section")


def pagina_educacion_superior():
    programas, aviso = cargar_programas_educacion_superior()
    return html.Section(className="vision-exec-page educacion-superior-page", children=[
        html.Div(className="vision-exec-header vision-exec-header-compacta", children=[
            html.Div(className="vision-title-mark"),
            html.Div([html.H1(
                          "EJE 2 • BECAS Y OPORTUNIDADES PARA JÓVENES",
                          className="gestion-educativa-titulo-eje",
                          style={
                              "fontSize": "clamp(15px, 1.15vw, 20px)",
                              "lineHeight": "1.2",
                              "whiteSpace": "normal",
                              "maxWidth": "1240px",
                          },
                      ),
                      html.P("Programas vigentes y agenda priorizada para ampliar el acceso, reconocer trayectorias "
                             "académicas y fortalecer la formación técnica, tecnológica y universitaria.")])]),
        html.H2("Educación Superior en cifras", className="vision-eje-banner"),
        html.Div(
            className="vision-exec-grid gestion-educativa-grid datos-generales-grid "
                      "educacion-superior-cifras-grid",
            children=tarjetas_datos_educacion_superior(),
        ),
        html.H2("Programas y servicios de Educación Superior", className="vision-eje-banner"),
        _tarjetas_programas_eje2(programas),
        html.P(aviso, className="eje2-source-note") if aviso else None,
        html.H2("Proyectos a futuro", className="vision-eje-banner"),
        html.P("Agenda de acciones estratégicas prevista para octubre y noviembre.", className="eje1-subtitulo"),
        _cartera_futura_eje2(EJE2_PROYECTOS_FUTUROS)])


EJES_VICEMINISTERIALES = {
    "educacion": {
        "numero": 3, "vice_excel": "Educación", "titulo": "Educación para el Nuevo Ecuador",
        "introduccion": "Aprendizajes, convivencia, transformación digital y servicios que fortalecen la educación.",
        "resumen": [("8", "Líneas de acción consolidadas"), ("Octubre · Noviembre", "Agenda prevista"),
                    ("USD 3.864.343", "Inversión y autogestión")],
        "acciones": [
            {"icono": "✦", "mes": "Octubre", "titulo": "Plan Nacional para el Fortalecimiento de los Aprendizajes",
             "descripcion": "Plan nacional enfocado en Lectura y Matemática, Ciencias y Pensamiento Computacional.",
             "detalle": "Quito · modalidad de autogestión"},
            {"icono": "▤", "mes": "Octubre", "titulo": "Guía «Que no te cuenten cuentos»",
             "descripcion": "1.000 guías orientadas a la cultura de paz.", "detalle": "Guayaquil · USD 7.500 aprox."},
            {"icono": "⌘", "mes": "Octubre", "titulo": "Transformación Digital para la Educación",
             "descripcion": "Alianza con Google para fortalecer capacidades desde educación básica hasta superior.",
             "detalle": "Primera fase 2026: USD 9,6 millones · alcance: USD 2,8 millones"},
            {"icono": "⚙", "mes": "Octubre–Noviembre", "titulo": "Robótica Educativa para reducir la brecha digital",
             "descripcion": "Fase 2: entrega de kits, capacitación docente y 69 clubes en 47 cantones.",
             "detalle": "Inversión total: USD 159.843"},
            {"icono": "▥", "mes": "Noviembre", "titulo": "Entrega de 222 ambientes de lectura",
             "descripcion": "Ambientes implementados en instituciones educativas rurales de 58 distritos.",
             "detalle": "La Concordia · 222 instituciones · USD 697.000"},
            {"icono": "⌾", "mes": "Octubre–Noviembre", "titulo": "Plan ESCUDO / Comunidades Educativas",
             "descripcion": "Prevención, protección y seguridad escolar en 102 cantones y 14 provincias priorizadas.",
             "detalle": "1.264.863 estudiantes · 5.293 instituciones"},
            {"icono": "▣", "mes": "Octubre–Noviembre", "titulo": "Concursos Nacionales de Comprensión Lectora",
             "descripcion": "Concursos de alcance nacional mediante alianzas estratégicas.",
             "detalle": "1.000 estudiantes"},
            {"icono": "◉", "mes": "Octubre–Noviembre", "titulo": "Fortalecimiento de los DECE",
             "descripcion": "Avance en la contratación de profesionales para Sierra y Amazonía.",
             "detalle": "Meta: 1.000 profesionales · inversión aprox. USD 3 millones"},
        ],
    },
    "deporte": {
        "numero": 4, "vice_excel": "Deporte", "titulo": "Deporte para el Nuevo Ecuador",
        "introduccion": "Programas de recreación, actividad física e infraestructura deportiva con alcance territorial.",
        "resumen": [("4", "Programas estratégicos"), ("98.500", "Beneficiarios directos"),
                    ("USD 2.120.195,10", "Inversión ejecutada")],
        "acciones": [
            {"icono": "▧", "mes": "Octubre–Noviembre", "titulo": "Programa «Pinta tu Cancha»",
             "descripcion": "Pintado de canchas comunitarias e institucionales mediante autogestión.",
             "detalle": "8.300 beneficiarios · 6 provincias · 9 intervenciones"},
            {"icono": "⌂", "mes": "Octubre–Noviembre", "titulo": "Infraestructura deportiva",
             "descripcion": "Rehabilitación, adecuación y mantenimiento de escenarios deportivos emblemáticos.",
             "detalle": "85.720 beneficiarios · USD 2.088.465,10 · 14 inauguraciones"},
            {"icono": "●", "mes": "Octubre–Noviembre", "titulo": "Programa «Actívate»",
             "descripcion": "Eventos y festivales orientados a promover actividad física en adultos y adultos mayores.",
             "detalle": "2.930 beneficiarios · USD 24.692 · 6 actividades"},
            {"icono": "★", "mes": "Octubre–Noviembre", "titulo": "Programa «Vamos a la Cancha»",
             "descripcion": "Práctica deportiva para niñas, niños y adolescentes de 5 a 17 años.",
             "detalle": "1.550 beneficiarios · USD 7.038 · 6 actividades"},
        ],
    },
    "cultura": {
        "numero": 5, "vice_excel": "Cultura", "titulo": "Cultura para el Nuevo Ecuador",
        "introduccion": "Circulación artística, memoria social, lectura y fortalecimiento de capacidades culturales.",
        "resumen": [("11", "Actividades totales"), ("181.915", "Beneficiarios directos"),
                    ("USD 1.297.399,25", "Inversión total ejecutada")],
        "acciones": [
            {"icono": "✦", "mes": "Octubre–Noviembre", "titulo": "Arte en mi Ciudad",
             "descripcion": "Dos ediciones mensuales de actividades artísticas abiertas a la comunidad.",
             "detalle": "3.000 beneficiarios"},
            {"icono": "◈", "mes": "Octubre–Noviembre", "titulo": "Arte en mi Escuela",
             "descripcion": "Sensibilización, mediación y formación artística en instituciones educativas.",
             "detalle": "8 provincias"},
            {"icono": "▤", "mes": "Octubre", "titulo": "Feria Académica Elige Crear (3.ª edición)",
             "descripcion": "Encuentro académico y cultural desarrollado en Azuay.", "detalle": "2.000 beneficiarios"},
            {"icono": "▥", "mes": "Octubre", "titulo": "Feria Internacional del Libro Quito",
             "descripcion": "Promoción del libro, la lectura y la circulación editorial.",
             "detalle": "45.000 beneficiarios · USD 400.000"},
            {"icono": "⌂", "mes": "Octubre", "titulo": "Ludobiblioteca del Complejo Ingapirca",
             "descripcion": "Entrega de un espacio cultural y educativo en Cañar.",
             "detalle": "5.525 beneficiarios · USD 57.699,25"},
            {"icono": "♨", "mes": "Octubre", "titulo": "IV Encuentro de Cocinas Iberoamericanas",
             "descripcion": "Encuentro para la puesta en valor del patrimonio alimentario.",
             "detalle": "322.925 beneficiarios · USD 30.000"},
            {"icono": "▶", "mes": "Octubre", "titulo": "Cine al Río MAAC",
             "descripcion": "Programación cinematográfica con una edición mensual.", "detalle": "700 beneficiarios"},
            {"icono": "▣", "mes": "Noviembre", "titulo": "Reapertura Showroom MUNA",
             "descripcion": "Reapertura del espacio expositivo del Museo Nacional.",
             "detalle": "Pichincha · 1.000 beneficiarios · USD 10.000"},
            {"icono": "◎", "mes": "Noviembre", "titulo": "Fortalecimiento de capacidades del REMAB",
             "descripcion": "Capacitación para fortalecer la gestión de la red.",
             "detalle": "Pichincha · 20 beneficiarios capacitados"},
            {"icono": "♫", "mes": "Noviembre", "titulo": "Festival Internacional de Artes Vivas de Loja",
             "descripcion": "Programación nacional e internacional de artes vivas.",
             "detalle": "160.000 beneficiarios · USD 829.700"},
            {"icono": "◇", "mes": "Noviembre", "titulo": "Remodelación integral del Museo de Ibarra",
             "descripcion": "Primera piedra para la intervención integral del museo.",
             "detalle": "187.536 beneficiarios potenciales"},
        ],
    },
}


def _tarjetas_acciones_eje(acciones):
    tarjetas = []
    for i, accion in enumerate(acciones, 1):
        tarjetas.append(html.Article([
            html.Div([
                html.Span(accion["icono"], className="eje-card-icon"),
                html.Span(accion["mes"], className="eje2-project-month"),
                html.Span(f"{i:02d}", className="eje2-project-number"),
            ], className="eje-card-top"),
            html.H3(accion["titulo"]), html.P(accion["descripcion"]),
            html.Div(accion["detalle"], className="eje-card-detail"),
        ], className="eje-action-card"))
    return html.Div(tarjetas, className="eje-actions-grid eje-actions-grid-4")


def pagina_eje_viceministerial(codigo):
    eje = EJES_VICEMINISTERIALES[codigo]
    programas, aviso = cargar_programas_viceministerio(eje["vice_excel"])
    return html.Section(className="vision-exec-page eje-viceministerial-page", children=[
        html.Div(className="vision-exec-header vision-exec-header-compacta", children=[
            html.Div(className="vision-title-mark"),
            html.Div([html.H1(
                          f"EJE {eje['numero']} • {eje['titulo'].upper()}",
                          className="gestion-educativa-titulo-eje",
                          style={
                              "fontSize": "clamp(15px, 1.15vw, 20px)",
                              "lineHeight": "1.2",
                              "whiteSpace": "normal",
                              "maxWidth": "1240px",
                          },
                      ),
                      html.P(eje["introduccion"])])]),
        html.H2(f"Programas y servicios de {eje['vice_excel']}", className="vision-eje-banner"),
        html.Div(
            _tarjetas_programas_eje2(programas),
            className="eje-programas-fila-unica",
        ) if programas else None,
        html.P(aviso, className="eje2-source-note") if aviso else None,
        html.H2("Acciones estratégicas", className="vision-eje-banner"),
        html.P("Resumen ejecutivo de actividades y acciones previstas para octubre y noviembre.",
               className="eje1-subtitulo"),
        html.Div([html.Div([html.Strong(v), html.Span(e)], className="eje1-stat")
                  for v, e in eje["resumen"] if "agenda" not in e.lower()],
                 className="eje1-stats-grid eje-summary-grid"),
        _tarjetas_acciones_eje(eje["acciones"]),
        html.P("Las cifras de esta sección corresponden a las diapositivas institucionales entregadas.",
               className="eje2-source-note"),
    ])


# Cifras generales del sistema educativo nacional (año lectivo 2025-2026),
# tal como las proporcionó la Viceministra — no provienen de las 4 bases de
# abajo (que son por programa), sino del portal de Datos Abiertos del
# Ministerio. Son un valor fijo hasta que se conecte una base propia.
DATOS_GENERALES_GESTION_EDUCATIVA = {
    "anio_lectivo": "2025-2026",
    "instituciones": 16215,
    "estudiantes_mujeres": 2005491,
    "estudiantes_hombres": 2034159,
    "docentes_total": 217693,
    "docentes_mujeres": 157629,
    "docentes_hombres": 60064,
}

DATOS_ABIERTOS_MINEDEC_URL = "https://educacion.gob.ec/datos-abiertos-minedec/"


def tarjetas_datos_generales(epja=None):
    """Cifras de Educación Media y, cuando existe, la fila EPJA del Excel."""
    d = DATOS_GENERALES_GESTION_EDUCATIVA
    subtitulo = f"Año lectivo {d['anio_lectivo']}"
    tarjetas = [
        tarjeta_gestion_educativa(
            "generales", "Instituciones Educativas (IE)", subtitulo,
            formato_valor(d["instituciones"]),
        ),
        tarjeta_gestion_educativa(
            "generales", "Estudiantes", subtitulo,
            formato_valor(d["estudiantes_mujeres"] + d["estudiantes_hombres"]),
            detalles=[
                html.Div([html.Span("Mujeres"), html.Strong(formato_valor(d["estudiantes_mujeres"]))],
                         className="vision-mini-stat"),
                html.Div([html.Span("Hombres"), html.Strong(formato_valor(d["estudiantes_hombres"]))],
                         className="vision-mini-stat secondary"),
            ],
        ),
        tarjeta_gestion_educativa(
            "generales", "Docentes", subtitulo,
            formato_valor(d["docentes_total"]),
            detalles=[
                html.Div([html.Span("Mujeres"), html.Strong(formato_valor(d["docentes_mujeres"]))],
                         className="vision-mini-stat"),
                html.Div([html.Span("Hombres"), html.Strong(formato_valor(d["docentes_hombres"]))],
                         className="vision-mini-stat secondary"),
            ],
        ),
    ]
    if epja is not None:
        tarjetas.append(tarjeta_gestion_educativa(
            "epja", "Educación para Jóvenes y Adultos (EPJA)",
            "Programa de Gestión Educativa",
            formato_valor(epja["beneficiarios"]),
            unidad=str(epja["tipo"]).strip(),
        ))
    return tarjetas


# ---------------------------------------------------------------------------
# Eje 1 · Infraestructura educativa (Fortalecimiento de la Infraestructura,
# Equipamiento y Alimentación Escolar). No proviene de ninguna de las 4 bases
# conectadas — son los registros puntuales de infraestructura que la
# Viceministra pasó a mano; se dejan aquí como fuente única hasta que exista
# una base propia. A diferencia de Cobertura, aquí SÍ se muestra la inversión
# (pedido explícito: esto es para presumir metas/logros, no cobertura de
# programas). Cada registro trae lat/lon aproximados de su provincia para el
# mapa (sin necesidad de un archivo geográfico externo).
EJE1_INSTITUCIONES = [
    {"institucion": "Unidad Educativa Puerto Limón", "beneficiarios": 1510,
     "inversion": 647584.59, "provincia": "Santo Domingo de los Tsáchilas",
     "canton": "Santo Domingo de los Tsáchilas", "lat": -0.2530, "lon": -79.1719},
    {"institucion": "Unidad Educativa Jaime del Hierro", "beneficiarios": 733,
     "inversion": 308154.49, "provincia": "Santo Domingo de los Tsáchilas",
     "canton": "Santo Domingo de los Tsáchilas", "lat": -0.2530, "lon": -79.1719},
    {"institucion": "Unidad Educativa \"Velasco Ibarra\"", "beneficiarios": 639,
     "inversion": 395258.71, "provincia": "Manabí", "canton": "Portoviejo",
     "lat": -1.0546, "lon": -80.4525},
    {"institucion": "Unidad Educativa Provincia de Manabí", "beneficiarios": 823,
     "inversion": 216434.73, "provincia": "Manabí", "canton": "Puerto López",
     "lat": -1.0546, "lon": -80.4525},
    {"institucion": "Unidad Educativa Las Mercedes", "beneficiarios": 631,
     "inversion": 509222.24, "provincia": "Manabí", "canton": "24 de Mayo",
     "lat": -1.0546, "lon": -80.4525},
    {"institucion": "Unidad Educativa Bosco Wisuma", "beneficiarios": 700,
     "inversion": 43472.62, "provincia": "Morona Santiago", "canton": "Morona",
     "lat": -2.3086, "lon": -78.1114},
    {"institucion": "Unidad Educativa 2 de Octubre", "beneficiarios": 129,
     "inversion": 199281.21, "provincia": "Napo", "canton": "Tena",
     "lat": -1.0021, "lon": -77.8140},
]

# Programa de Alimentación Escolar: resumen provincial entregado para la
# presentación ejecutiva. Se mantiene separado de las inauguraciones porque
# su unidad de análisis es la provincia y no la institución individual.
PMA_PROVINCIAS = [
    {"provincia": "Esmeraldas", "instituciones": 6, "beneficiarios": 5654,
     "inversion": 930280.85},
    {"provincia": "Pichincha", "instituciones": 9, "beneficiarios": 13999,
     "inversion": 2094137.13},
    {"provincia": "Tungurahua", "instituciones": 3, "beneficiarios": 1206,
     "inversion": 203108.06},
    {"provincia": "Santo Domingo de los Tsáchilas", "instituciones": 2,
     "beneficiarios": 3978, "inversion": 674679.66},
    {"provincia": "Los Ríos", "instituciones": 2, "beneficiarios": 3653,
     "inversion": 598117.78},
    {"provincia": "Santa Elena", "instituciones": 15, "beneficiarios": 6220,
     "inversion": 1053251.35},
    {"provincia": "Azuay", "instituciones": 1, "beneficiarios": 975,
     "inversion": 145942.94},
    {"provincia": "El Oro", "instituciones": 3, "beneficiarios": 3509,
     "inversion": 630482.80},
    {"provincia": "Guayas", "instituciones": 2, "beneficiarios": 4226,
     "inversion": 669398.40},
]


def _moneda_corta(valor):
    """Formato de moneda consistente con el resto del panel (ver _moneda,
    más abajo, para Ejecución Presupuestaria)."""
    return "$ " + f"{float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def franja_stats_eje1(registros):
    total_instituciones = len(registros)
    total_beneficiarios = sum(r["beneficiarios"] for r in registros)
    total_inversion = sum(r["inversion"] for r in registros)
    total_provincias = len({r["provincia"] for r in registros})
    datos = [
        (formato_valor(total_instituciones), "Instituciones Educativas"),
        (formato_valor(total_beneficiarios), "Estudiantes Beneficiados"),
        (_moneda_corta(total_inversion), "Inversión Total"),
        (formato_valor(total_provincias), "Provincias"),
    ]
    return html.Div([
        html.Div([html.Strong(valor), html.Span(etiqueta)], className="eje1-stat")
        for valor, etiqueta in datos
    ], className="eje1-stats-grid")


def tabla_eje1(registros):
    """La tabla va detrás de un <details>/<summary> (igual que las fórmulas
    del panel): colapsada por defecto, se despliega con un clic — pedido
    explícito del usuario, no mostrar la tabla de entrada."""
    filas = [
        html.Tr([
            html.Td(r["institucion"]), html.Td(formato_valor(r["beneficiarios"])),
            html.Td(_moneda_corta(r["inversion"])), html.Td(r["provincia"]), html.Td(r["canton"]),
        ]) for r in registros
    ]
    tabla = html.Table([
        html.Thead(html.Tr([html.Th("Institución"), html.Th("Beneficiarios"), html.Th("Inversión"),
                             html.Th("Provincia"), html.Th("Cantón")])),
        html.Tbody(filas),
    ], className="eje1-tabla")
    return html.Details([
        html.Summary([
            html.Span("Ver tabla de instituciones"),
            html.Span("⌄", className="eje1-tabla-chevron"),
        ], className="eje1-tabla-header"),
        html.Div(tabla, className="eje1-tabla-body"),
    ], className="eje1-tabla-accordion")


def _sin_tildes(texto):
    normalizado = unicodedata.normalize("NFKD", str(texto).strip().lower())
    return "".join(c for c in normalizado if not unicodedata.combining(c))


_GEOJSON_PROVINCIAS_CACHE = None


def _cargar_geojson_provincias():
    """Carga provincias_ecuador.geojson (generado por
    procesar_mapa_provincias.py a partir del shapefile oficial de CONALI).
    Se cachea en memoria; si el archivo no existe todavía, devuelve None y
    el mapa cae de vuelta al modo de burbujas por coordenadas."""
    global _GEOJSON_PROVINCIAS_CACHE
    if _GEOJSON_PROVINCIAS_CACHE is not None:
        return _GEOJSON_PROVINCIAS_CACHE
    # El archivo puede estar junto a app.py o dentro de assets. No se guarda
    # un fallo en caché: si el procesador genera el GeoJSON mientras la app
    # está abierta, una actualización posterior podrá encontrarlo.
    raices = (BASE_DIR, BASE_DIR / "assets", Path.cwd(), Path.cwd() / "assets")
    candidatos = [raiz / "provincias_ecuador.geojson" for raiz in raices]

    # GitHub/Posit Cloud puede conservar el archivo dentro de una subcarpeta
    # diferente o con otra combinación de mayúsculas y minúsculas.
    vistos = set()
    for raiz in (BASE_DIR, Path.cwd()):
        try:
            for ruta in raiz.rglob("*"):
                if (ruta.is_file()
                        and ruta.name.casefold() == "provincias_ecuador.geojson"
                        and ".git" not in ruta.parts):
                    candidatos.append(ruta)
        except OSError:
            continue

    for ruta in candidatos:
        try:
            ruta_resuelta = ruta.resolve()
        except OSError:
            continue
        if ruta_resuelta in vistos or not ruta_resuelta.is_file():
            continue
        vistos.add(ruta_resuelta)
        try:
            with open(ruta_resuelta, "r", encoding="utf-8-sig") as f:
                contenido = json.load(f)
            if contenido.get("type") == "FeatureCollection" and contenido.get("features"):
                _GEOJSON_PROVINCIAS_CACHE = contenido
                return _GEOJSON_PROVINCIAS_CACHE
        except (OSError, UnicodeError, json.JSONDecodeError, AttributeError):
            continue
    return None


def mapa_eje1(registros):
    """Mapa coroplético de provincias: colorea con los límites reales de
    Ecuador (shapefile oficial de CONALI, procesado por
    procesar_mapa_provincias.py) las provincias donde hubo intervenciones.
    Si todavía no se generó el GeoJSON, usa un mapa de burbujas como
    respaldo para que la página nunca se rompa."""
    agregados = {}
    for r in registros:
        prov = agregados.setdefault(r["provincia"], {
            "lat": r["lat"], "lon": r["lon"], "instituciones": 0, "beneficiarios": 0,
        })
        prov["instituciones"] += 1
        prov["beneficiarios"] += r["beneficiarios"]

    geojson = _cargar_geojson_provincias()

    if geojson:
        texto_hover = {}
        resaltadas = []  # (nombre, lon, lat, instituciones) de las provincias con intervención
        nombres_resaltados = set()
        for feat in geojson["features"]:
            nombre = feat["properties"]["nombre"]
            match = None
            for clave_prov, datos in agregados.items():
                if _sin_tildes(clave_prov) == _sin_tildes(nombre):
                    match = datos
                    break
            if match:
                nombres_resaltados.add(nombre)
                texto_hover[nombre] = (
                    f"<b>{nombre.title()}</b><br>{match['instituciones']} institución(es)<br>"
                    f"{formato_valor(match['beneficiarios'])} estudiantes"
                )
                resaltadas.append((
                    nombre.title(),
                    feat["properties"]["centroide_lon"],
                    feat["properties"]["centroide_lat"],
                    match["instituciones"],
                ))
            else:
                texto_hover[nombre] = f"<b>{nombre.title()}</b><br>Sin intervenciones registradas"

        # Tres capas sólidas: provincias sin intervención en blanco y las
        # cuatro provincias destacadas repartidas entre amarillo y azul.
        feats_base = geojson["features"]
        # Distribución tomada del diseño original:
        # amarillo = Manabí y Morona Santiago;
        # violeta = Santo Domingo de los Tsáchilas y Napo.
        amarillas = {"manabi", "morona santiago"}
        azules = {"santo domingo de los tsachilas", "napo"}
        feats_amarillas = [
            f for f in feats_base
            if _sin_tildes(f["properties"]["nombre"]) in amarillas
        ]
        feats_azules = [
            f for f in feats_base
            if _sin_tildes(f["properties"]["nombre"]) in azules
        ]
        feats_normales = [
            f for f in feats_base
            if _sin_tildes(f["properties"]["nombre"]) not in amarillas | azules
        ]

        def _trazo(features, color):
            nombres = [f["properties"]["nombre"] for f in features]
            return go.Choropleth(
                geojson={"type": "FeatureCollection", "features": features},
                locations=nombres,
                featureidkey="properties.nombre",
                z=[1] * len(nombres),
                zmin=0, zmax=1,
                colorscale=[[0, color], [1, color]],
                showscale=False,
                marker_line_color="#28345f",
                marker_line_width=1.0,
                text=[texto_hover[n] for n in nombres],
                hoverinfo="text",
            )

        fig = go.Figure()
        if feats_normales:
            fig.add_trace(_trazo(feats_normales, "#ffffff"))
        if feats_amarillas:
            fig.add_trace(_trazo(feats_amarillas, "#f8bd20"))
        if feats_azules:
            fig.add_trace(_trazo(feats_azules, "#4d3a94"))

        if resaltadas:
            fig.add_trace(go.Scattergeo(
                lon=[r[1] for r in resaltadas],
                lat=[r[2] for r in resaltadas],
                mode="markers",
                marker=dict(size=5, color="#17245b", line=dict(width=1, color="#ffffff")),
                hoverinfo="skip",
                showlegend=False,
            ))
        fig.update_geos(
            scope="south america",
            lataxis_range=[-5.6, 2.0],
            lonaxis_range=[-82.3, -74.2],
            showcountries=False,
            showland=False,
            showocean=True, oceancolor="#ffffff",
            resolution=50,
            bgcolor="#ffffff",
            fitbounds=False,
            projection_scale=1,
        )
        fig.update_layout(margin=dict(l=28, r=28, t=10, b=10), height=400,
                           paper_bgcolor="#ffffff", plot_bgcolor="#ffffff",
                           autosize=True, showlegend=False)
        return dcc.Graph(figure=fig, config={"displayModeBar": False, "responsive": True},
                          className="eje1-mapa", style={"width": "100%", "height": "400px"})

    # --- Respaldo: mapa de burbujas por coordenadas (sin GeoJSON) ---
    provincias = list(agregados)
    etiqueta_pin = [
        f"{p}<br>{agregados[p]['instituciones']} "
        f"{'institución' if agregados[p]['instituciones'] == 1 else 'instituciones'}"
        for p in provincias
    ]
    fig = go.Figure(go.Scattergeo(
        lat=[agregados[p]["lat"] for p in provincias],
        lon=[agregados[p]["lon"] for p in provincias],
        text=etiqueta_pin,
        textposition="top center",
        textfont=dict(size=12, color="#28304f", family="Arial Black, Arial"),
        hovertext=[f"<b>{p}</b><br>{agregados[p]['instituciones']} institución(es)<br>"
                   f"{formato_valor(agregados[p]['beneficiarios'])} estudiantes" for p in provincias],
        mode="markers+text",
        marker=dict(
            size=[22 + agregados[p]["instituciones"] * 7 for p in provincias],
            color="#f8bd20",
            opacity=1,
            line=dict(width=2.5, color="#503a98"),
            symbol="circle",
        ),
        hoverinfo="text",
    ))
    fig.update_geos(
        scope="south america",
        lataxis_range=[-5.6, 1.8],
        lonaxis_range=[-82, -74.8],
        showcountries=True, countrycolor="#9aa2bd",
        showland=True, landcolor="#f5f6fa",
        showocean=True, oceancolor="#eef1fb",
        showsubunits=True, subunitcolor="#c7cbe0",
        countrywidth=1.4,
        resolution=50,
        bgcolor="rgba(0,0,0,0)",
    )
    fig.update_layout(margin=dict(l=10, r=10, t=30, b=10), height=420,
                       paper_bgcolor="rgba(0,0,0,0)")
    return dcc.Graph(figure=fig, config={"displayModeBar": False}, className="eje1-mapa")


def seccion_eje1_infraestructura():
    registros = EJE1_INSTITUCIONES
    return html.Div([
        html.P("Infraestructura educativa", className="eje1-subtitulo"),
        franja_stats_eje1(registros),
        html.Div([
            html.Div(tabla_eje1(registros), className="eje1-col-tabla"),
            html.Div([mapa_eje1(registros)], className="eje1-mapa-wrap eje1-col-mapa"),
        ], className="eje1-fila-detalle"),
    ], className="eje1-seccion")


def franja_stats_pma(registros):
    datos = [
        (formato_valor(sum(r["instituciones"] for r in registros)), "Instituciones Educativas"),
        (formato_valor(sum(r["beneficiarios"] for r in registros)), "Estudiantes Beneficiarios"),
        (_moneda_corta(sum(r["inversion"] for r in registros)), "Inversión Total"),
        (formato_valor(len(registros)), "Provincias"),
    ]
    return html.Div([
        html.Div([html.Strong(valor), html.Span(etiqueta)], className="eje1-stat")
        for valor, etiqueta in datos
    ], className="eje1-stats-grid pma-stats-grid")


def tabla_pma(registros):
    filas = [
        html.Tr([
            html.Td(r["provincia"]),
            html.Td(formato_valor(r["instituciones"])),
            html.Td(formato_valor(r["beneficiarios"])),
            html.Td(_moneda_corta(r["inversion"])),
        ]) for r in registros
    ]
    tabla = html.Table([
        html.Thead(html.Tr([
            html.Th("Provincia"), html.Th("N.° IE"),
            html.Th("Beneficiarios"), html.Th("Inversión PMA"),
        ])),
        html.Tbody(filas),
    ], className="eje1-tabla pma-tabla")
    return html.Details([
        html.Summary([
            html.Span("Ver tabla por provincias"),
            html.Span("⌄", className="eje1-tabla-chevron"),
        ], className="eje1-tabla-header"),
        html.Div(tabla, className="eje1-tabla-body"),
    ], className="eje1-tabla-accordion")


def mapa_pma(registros):
    geojson = _cargar_geojson_provincias()
    if not geojson:
        # Respaldo visible cuando el despliegue no incluyó el GeoJSON. Evita
        # dejar un cuadro vacío y mantiene las nueve provincias identificadas.
        centros = {
            "esmeraldas": (0.73, -79.15), "pichincha": (-0.18, -78.47),
            "tungurahua": (-1.25, -78.62),
            "santo domingo de los tsachilas": (-0.25, -79.17),
            "los rios": (-1.42, -79.47), "santa elena": (-2.23, -80.86),
            "azuay": (-2.90, -79.01), "el oro": (-3.26, -79.96),
            "guayas": (-2.20, -79.89),
        }
        amarillas = {"esmeraldas", "pichincha", "los rios", "santa elena", "azuay"}
        claves = [_sin_tildes(r["provincia"]) for r in registros]
        fig = go.Figure(go.Scattergeo(
            lat=[centros[c][0] for c in claves],
            lon=[centros[c][1] for c in claves],
            text=[f"<b>{r['provincia']}</b><br>{r['instituciones']} IE" for r in registros],
            hovertext=[
                f"<b>{r['provincia']}</b><br>{r['instituciones']} instituciones"
                f"<br>{formato_valor(r['beneficiarios'])} beneficiarios"
                f"<br>{_moneda_corta(r['inversion'])}" for r in registros
            ],
            mode="markers+text", textposition="top center",
            textfont=dict(size=10, color="#17245b"), hoverinfo="text",
            marker=dict(
                size=[14 + min(r["instituciones"], 10) * 2 for r in registros],
                color=["#f8bd20" if c in amarillas else "#4d3a94" for c in claves],
                line=dict(width=2, color="#ffffff"), opacity=1,
            ),
        ))
        fig.update_geos(
            scope="south america", lataxis_range=[-5.6, 2.0],
            lonaxis_range=[-82.3, -74.2], showcountries=True,
            countrycolor="#9aa2bd", showland=True, landcolor="#f5f6fa",
            showocean=True, oceancolor="#ffffff", resolution=50,
            bgcolor="#ffffff",
        )
        fig.update_layout(
            margin=dict(l=10, r=10, t=20, b=10), height=420,
            paper_bgcolor="#ffffff", plot_bgcolor="#ffffff", showlegend=False,
        )
        return dcc.Graph(
            figure=fig, config={"displayModeBar": False, "responsive": True},
            className="eje1-mapa", style={"width": "100%", "height": "420px"},
        )

    por_provincia = {_sin_tildes(r["provincia"]): r for r in registros}
    # Distribución tomada de la lámina original del PMA.
    amarillas = {"esmeraldas", "pichincha", "los rios", "santa elena", "azuay"}
    azules = {
        "santo domingo de los tsachilas", "tungurahua", "guayas", "el oro",
    }
    seleccionadas = amarillas | azules

    normales, feats_amarillas, feats_azules = [], [], []
    etiquetas = []
    hover = {}
    for feat in geojson["features"]:
        nombre = feat["properties"]["nombre"]
        clave = _sin_tildes(nombre)
        dato = por_provincia.get(clave)
        if dato:
            hover[nombre] = (
                f"<b>{dato['provincia']}</b><br>{dato['instituciones']} instituciones"
                f"<br>{formato_valor(dato['beneficiarios'])} beneficiarios"
                f"<br>{_moneda_corta(dato['inversion'])}"
            )
            etiqueta = "Santo Domingo" if clave == "santo domingo de los tsachilas" else dato["provincia"]
            etiquetas.append((
                feat["properties"]["centroide_lon"],
                feat["properties"]["centroide_lat"],
                f"<b>{etiqueta}</b><br>{dato['instituciones']} IE",
            ))
        else:
            hover[nombre] = f"<b>{nombre.title()}</b><br>Sin cobertura priorizada"

        if clave in amarillas:
            feats_amarillas.append(feat)
        elif clave in azules:
            feats_azules.append(feat)
        elif clave not in seleccionadas:
            normales.append(feat)

    def _capa(features, color):
        nombres = [f["properties"]["nombre"] for f in features]
        return go.Choropleth(
            geojson={"type": "FeatureCollection", "features": features},
            locations=nombres, featureidkey="properties.nombre",
            z=[1] * len(nombres), zmin=0, zmax=1,
            colorscale=[[0, color], [1, color]], showscale=False,
            marker_line_color="#4d4f83", marker_line_width=.85,
            text=[hover[n] for n in nombres], hoverinfo="text",
        )

    fig = go.Figure()
    fig.add_trace(_capa(normales, "#ffffff"))
    fig.add_trace(_capa(feats_amarillas, "#f8bd20"))
    fig.add_trace(_capa(feats_azules, "#4d3a94"))
    fig.add_trace(go.Scattergeo(
        lon=[e[0] for e in etiquetas], lat=[e[1] for e in etiquetas],
        mode="markers", text=[e[2] for e in etiquetas],
        marker=dict(size=5, color="#17245b", line=dict(width=1, color="#ffffff")),
        hovertemplate="%{text}<extra></extra>", showlegend=False,
    ))
    fig.update_geos(
        scope="south america", lataxis_range=[-5.6, 2.0], lonaxis_range=[-82.3, -74.2],
        showcountries=False, showland=False, showocean=True, oceancolor="#ffffff",
        resolution=50, bgcolor="#ffffff", fitbounds=False,
    )
    fig.update_layout(
        margin=dict(l=20, r=20, t=8, b=8), height=420,
        paper_bgcolor="#ffffff", plot_bgcolor="#ffffff", showlegend=False,
    )
    return dcc.Graph(
        figure=fig, config={"displayModeBar": False, "responsive": True},
        className="eje1-mapa", style={"width": "100%", "height": "420px"},
    )


def seccion_pma():
    registros = PMA_PROVINCIAS
    return html.Div([
        html.H2(
            "Alimentación escolar: inversión que llega al territorio",
            className="vision-eje-banner",
        ),
        html.P(
            "Por provincia priorizada un plato de comida con inversión",
            className="eje1-subtitulo pma-subtitulo",
        ),
        franja_stats_pma(registros),
        html.Div([
            html.Div(tabla_pma(registros), className="eje1-col-tabla"),
            html.Div([mapa_pma(registros)], className="eje1-mapa-wrap eje1-col-mapa"),
        ], className="eje1-fila-detalle pma-fila-detalle"),
    ], className="eje1-seccion pma-seccion")


def seccion_proximamente(titulo, descripcion):
    """Placeholder visual para secciones que todavía no tienen contenido ni
    fuente de datos, pero ya reservan su lugar en la página."""
    return html.Div([
        html.H2(titulo),
        html.P(descripcion),
    ], className="vision-proximamente")


# (código CSS, fuente, título, subtítulo, clave principal, unidad, subgrupo)
TARJETAS_GESTION_EDUCATIVA = [
    ("alimentacion", "Alimentación Escolar", "Alimentación Escolar",
     "Raciones entregadas por cantón", "beneficiarios", "Estudiantes", None),
    ("uniformes", "Uniformes Escolares", "Uniformes Escolares",
     "Entrega de uniformes por cantón", "beneficiarios", "Estudiantes", None),
    ("textos", "Textos Escolares", "Textos Escolares",
     "Entrega de textos por cantón", "beneficiarios", "Estudiantes", None),
    ("mobiliario", "Mobiliario y Transporte Escolar", "Mobiliario Escolar",
     "Instituciones atendidas con mobiliario", "estudiantes", "Estudiantes", "mobiliario"),
    ("transporte", "Mobiliario y Transporte Escolar", "Transporte Escolar",
     "Estudiantes atendidos con transporte", "estudiantes", "Estudiantes", "transporte"),
]


def _detalles_tarjeta_gestion(dato):
    detalles = [html.Div([
        html.Span("Instituciones"),
        html.Strong(formato_valor(dato.get("instituciones", 0)))
    ], className="vision-mini-stat secondary")]
    return detalles


def pagina_gestion_educativa():
    """Una tarjeta por cada una de las 4 bases reales, mostrando exactamente
    lo que cada una tiene (beneficiarios/instituciones y su cobertura por
    provincia) — nunca el monto de inversión. No se inventan tarjetas de
    Matrícula, Permanencia o Riesgos de continuidad: esas requieren la base
    de matrícula estudiantil, que todavía no está conectada aquí."""
    resumen = RESUMEN_GESTION_EDUCATIVA
    pendientes = PENDIENTES_GESTION_EDUCATIVA
    programas_gestion, aviso_programas = cargar_programas_viceministerio("Gestión Educativa")
    epja = next((p for p in programas_gestion
                 if "epja" in _sin_tildes(p.get("programa", ""))), None)
    programas_gestion = [p for p in programas_gestion
                         if "epja" not in _sin_tildes(p.get("programa", ""))]

    def dato(nombre_legible, clave):
        return resumen.get(nombre_legible, {}).get(clave)

    tarjetas = []
    for (codigo, nombre_legible, titulo, subtitulo, clave_principal,
         etiqueta_principal, subgrupo) in TARJETAS_GESTION_EDUCATIVA:
        fuente_completa = resumen.get(nombre_legible, {})
        fuente = (fuente_completa.get("por_recurso", {}).get(subgrupo, {})
                  if subgrupo else fuente_completa)
        valor_principal = fuente.get(clave_principal)
        if valor_principal is None and clave_principal == "estudiantes":
            valor_principal = fuente.get("beneficiarios")
        detalles = _detalles_tarjeta_gestion(fuente)
        tarjetas.append(tarjeta_gestion_educativa(
            codigo, titulo, subtitulo,
            formato_valor(valor_principal) if valor_principal is not None else "Sin fuente",
            detalles=detalles, unidad=etiqueta_principal,
        ))

    pie = None
    if pendientes:
        pie = html.Div([
            html.Strong("Fuentes por complementar: "),
            html.Span("aún no se pudieron leer las bases de " + "; ".join(pendientes) + "."),
        ], className="vision-exec-footer-aviso")

    return html.Section(className="vision-exec-page gestion-educativa-page", children=[
        html.Div(className="vision-exec-header vision-exec-header-compacta", children=[
            html.Div(className="vision-title-mark"),
            html.Div([
                html.H1(
                    "EJE 1 • FORTALECIMIENTO DE LA INFRAESTRUCTURA, "
                    "EQUIPAMIENTO Y ALIMENTACIÓN ESCOLAR",
                    className="gestion-educativa-titulo-eje",
                    style={
                        "fontSize": "clamp(15px, 1.15vw, 20px)",
                        "lineHeight": "1.2",
                        "whiteSpace": "normal",
                        "maxWidth": "1240px",
                    },
                ),
                html.P("Totales nacionales del sistema educativo y, por separado, la cobertura de Alimentación, "
                       "Uniformes, Textos Escolares, Mobiliario y Transporte Escolar. No "
                       "incluye montos de inversión. Matrícula, permanencia y riesgos de continuidad requieren "
                       "la base de matrícula estudiantil, todavía no conectada a esta vista."),
            ]),
        ]),

        html.H2("Información de Educación Media", className="vision-eje-banner"),
        html.Div(
                 className="vision-exec-grid gestion-educativa-grid datos-generales-grid "
                           "educacion-media-cifras-grid",
                 children=tarjetas_datos_generales(epja)),

        html.H2("Programas y servicios de Gestión Educativa", className="vision-eje-banner"),
        html.P("Cifras consolidadas de los programas y servicios institucionales.",
               className="eje1-subtitulo"),
        _tarjetas_programas_eje2(programas_gestion) if programas_gestion else None,
        html.P(aviso_programas, className="eje2-source-note") if aviso_programas else None,

        html.H2("Recursos educativos y complementarios",
                className="vision-eje-banner"),
        html.Div(className="vision-exec-grid gestion-educativa-grid", children=tarjetas),
        pie,

        html.H2("Proyectos que transforman la educación",
                className="vision-eje-banner"),
        seccion_eje1_infraestructura(),
        seccion_pma(),
    ])


# ---------------------------------------------------------------------------
# Visión Ejecutiva · infografía institucional estática
# ---------------------------------------------------------------------------
def tabla_vision(grupo):
    """Convierte el formato largo del Excel en una tabla HTML ordenada."""
    indicadores = (grupo[["Indicador", "Orden indicador"]].drop_duplicates()
                    .sort_values("Orden indicador")["Indicador"].tolist())
    filas = (grupo[["Etiqueta fila", "Orden fila"]].drop_duplicates()
             .sort_values("Orden fila"))
    omitir_etiqueta = len(filas) == 1 and filas.iloc[0]["Etiqueta fila"].strip().lower() == "nacional"

    encabezados = ([] if omitir_etiqueta else [html.Th("Categoría")])
    encabezados += [html.Th(indicador) for indicador in indicadores]
    cuerpo = []
    for _, fila in filas.iterrows():
        etiqueta = fila["Etiqueta fila"]
        celdas = [] if omitir_etiqueta else [html.Td(etiqueta, className="vision-row-label")]
        for indicador in indicadores:
            valor = grupo.loc[(grupo["Etiqueta fila"] == etiqueta)
                              & (grupo["Indicador"] == indicador), "Valor"]
            celdas.append(html.Td(formato_valor(valor.iloc[0]) if not valor.empty else "—"))
        clase_fila = "vision-national-row" if str(etiqueta).strip().lower() == "nacional" else ""
        cuerpo.append(html.Tr(celdas, className=clase_fila))
    return html.Div(html.Table([
        html.Thead(html.Tr(encabezados)), html.Tbody(cuerpo)
    ], className="vision-table"), className="vision-table-wrap")


def tarjeta_tabla_vision(grupo):
    titulo = grupo["Tabla"].iloc[0]
    seccion_tabla = primer_texto(grupo, "Sección", "")
    es_informacion_general = str(titulo).strip().lower() == "educación y gestión"
    if es_informacion_general:
        titulo = "Información General"
    if str(titulo).strip().lower() == "deportistas identificados":
        titulo = "Deportistas"
    titulo_normalizado = str(titulo).strip().lower()
    es_instituto_tecnico = "institutos técnicos y tecnológicos superiores" in titulo_normalizado
    fuente = primer_texto(grupo, "Fuente", "No registrada")
    nota = primer_texto(grupo, "Nota", "")
    pie = [html.Div([
        html.Strong(f"Fuente: {fuente}"),
        html.Strong("Año lectivo: 2025–2026") if es_informacion_general else None,
        html.Strong(f"Año: {'2025' if es_instituto_tecnico else '2024'}")
        if seccion_tabla == "Educación Superior" else None,
    ], className="vision-source-main")]
    if nota:
        pie.append(html.P([html.Strong("Nota: "), nota]))
    return html.Article([
        html.H3(titulo), tabla_vision(grupo), html.Div(pie, className="vision-source")
    ], className="vision-table-card")


def contenido_resumen_minedec():
    """Muestra la infografía institucional (imagen) sin recortes ni deformación."""
    return html.Section(
        className="vision-image-page",
        children=[
            html.Div(
                className="vision-image-frame",
                children=html.Img(
                    src=app.get_asset_url("vision-ejecutiva-moderna.png"),
                    className="vision-image-original",
                    alt="Infografía MINEDEC 2026",
                    style={"width": "100%", "height": "auto", "display": "block"},
                ),
            ),
        ],
    )


def contenido_vision(tab=None):
    """Al entrar a Visión Ejecutiva se presenta el menú de los 5 viceministerios
    y nada más, hasta que se elige uno; hoy solo "Gestión Educativa" tiene
    contenido real, el resto está en construcción."""
    tab = tab if tab in dict(VISION_TABS) else None
    menu = vision_tabs_nav(tab)
    if tab is None:
        return html.Div([
            html.H1(
                "MINISTERIO DE EDUCACIÓN, DEPORTE Y CULTURA",
                className="vision-landing-title",
                style={
                    "margin": "0",
                    "color": "#4e3cab",
                    "fontSize": "clamp(25px, 2.35vw, 40px)",
                    "fontWeight": "900",
                    "lineHeight": "1.08",
                    "letterSpacing": ".01em",
                    "textAlign": "center",
                },
            ),
            html.Div(menu, className="vision-landing-menu", style={"width": "100%", "marginTop": "0"}),
            html.Div([
                html.Img(
                    src=app.get_asset_url("vision-ejecutiva-portada.png"),
                    className="vision-landing-image",
                    alt="Resumen ejecutivo MINEDEC 2026",
                    style={"display": "block", "width": "100%", "height": "auto",
                           "objectFit": "contain",
                           "borderRadius": "14px", "background": "#fff",
                           "boxShadow": "0 12px 28px rgba(24,37,87,.10)"},
                ),
            ], className="vision-landing-visual",
               style={"width": "100%", "minWidth": "0", "display": "flex",
                      "alignItems": "flex-start", "justifyContent": "center",
                      "position": "relative", "aspectRatio": "1909 / 1079",
                      "overflow": "hidden",
                      "flex": "1 1 auto", "borderRadius": "14px"}),
        ], className="vision-landing",
           style={"minHeight": "calc(100vh - 92px)", "padding": "18px 28px 26px",
                  "display": "flex", "flexDirection": "column", "gap": "14px",
                  "alignItems": "stretch", "boxSizing": "border-box", "background": "#f5f6fa"})
    if tab == "gestion-educativa":
        cuerpo = pagina_gestion_educativa()
    elif tab == "educacion-superior":
        cuerpo = pagina_educacion_superior()
    elif tab in EJES_VICEMINISTERIALES:
        cuerpo = pagina_eje_viceministerial(tab)
    else:
        cuerpo = html.Div([
            html.H1(dict(VISION_TABS)[tab]),
            html.P("Esta sección está en construcción: aún no se ha definido ni cargado su fuente de datos."),
        ], className="module-welcome")
    return html.Div([menu, cuerpo])


# ---------------------------------------------------------------------------
# Ejecución Presupuestaria · resumen y detalle por viceministerio
# ---------------------------------------------------------------------------
def _moneda(valor):
    return "$ " + f"{float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _formatear_fecha_actualizacion():
    """Devuelve únicamente la fecha de corte; la hora HTTP no es informativa."""
    crudo = (PRESUPUESTO_META.get("fecha") or "").strip()
    if not crudo:
        return None
    try:
        fecha = parsedate_to_datetime(crudo)  # Encabezado HTTP, con zona horaria.
        fecha = fecha.astimezone().replace(tzinfo=None) if fecha.tzinfo else fecha
        # El servidor de descarga informa en GMT; Ecuador es GMT-5.
        fecha = fecha - timedelta(hours=5)
        return fecha.strftime("%d/%m/%Y")
    except (TypeError, ValueError):
        pass
    for patron in ("%d/%m/%Y %H:%M", "%d/%m/%Y"):
        try:
            fecha = datetime.strptime(crudo, patron)
            return fecha.strftime("%d/%m/%Y")
        except ValueError:
            continue
    return crudo


def _columna_agrupacion_presupuesto(df, agrupar_por="gasto"):
    """Escoge la mejor columna disponible para agrupar el detalle: por grupo de
    gasto (por defecto) o por proyecto, según lo que exista en el Excel ESIGEF."""
    if agrupar_por == "proyecto":
        candidatas = ["NOM_PROYECTO", "PROYECTO", "NOMBRE_PROYECTO", "DESCRIPCION_PROYECTO"]
    else:
        candidatas = [
            "NOM_GRUPO", "GRUPO", "GRUPO_GASTO", "NOMBRE_GRUPO",
            "NOM_ITEM", "ITEM", "NOM_PROGRAMA", "TIPO",
        ]
    return next((col for col in candidatas if col in df.columns), "Viceministerio")


def _filtrar_alcance_presupuesto(df, alcance="Total"):
    """Filtra Corriente/Inversión tolerando variantes de escritura."""
    if df.empty or alcance in (None, "Total") or "TIPO" not in df.columns:
        return df.copy()
    patron = "INVERSION" if alcance == "Inversión" else "CORRIENTE"
    tipo_normalizado = (df["TIPO"].fillna("").astype(str)
                        .map(_normalizar_encabezado_presupuesto))
    return df.loc[tipo_normalizado.str.contains(patron, na=False)].copy()


def _etiqueta_columna_agrupacion(agrupar_por="gasto"):
    return "Proyecto" if agrupar_por == "proyecto" else "Grupo de gasto"


def _resumen_grupos_presupuesto(df, agrupar_por="gasto"):
    """Consolida partidas individuales y calcula ejecución y semáforo, agrupando
    por grupo de gasto o por proyecto según `agrupar_por`."""
    etiqueta_columna = _etiqueta_columna_agrupacion(agrupar_por)
    if df.empty:
        return pd.DataFrame(columns=[
            "POA/PAI", etiqueta_columna, "CODIFICADO", "COMPROMISO",
            "DEVENGADO", "SALDO_DISPONIBLE", "EJECUCION", "SEMAFORO",
        ])
    columna = _columna_agrupacion_presupuesto(df, agrupar_por)
    trabajo = df.copy()
    trabajo[etiqueta_columna] = (trabajo[columna].fillna("No especificado")
                                  .astype(str).str.strip().replace("", "No especificado"))
    trabajo["POA/PAI"] = (trabajo["TIPO"].fillna("No especificado").astype(str).str.strip()
                           if "TIPO" in trabajo.columns else "No especificado")
    resumen = (trabajo.groupby(["POA/PAI", etiqueta_columna], dropna=False)[
        ["CODIFICADO", "COMPROMISO", "DEVENGADO", "SALDO_DISPONIBLE"]
    ].sum().reset_index())
    resumen = resumen.loc[(resumen[["CODIFICADO", "COMPROMISO", "DEVENGADO"]]
                           .abs().sum(axis=1) > 0)].copy()
    resumen["EJECUCION"] = resumen.apply(
        lambda r: (r["DEVENGADO"] / r["CODIFICADO"] * 100) if r["CODIFICADO"] else 0.0,
        axis=1,
    )
    resumen["SEMAFORO"] = resumen["EJECUCION"].map(
        lambda v: "Verde" if v >= 70 else "Amarillo" if v >= 40 else "Rojo"
    )
    return resumen.sort_values(["POA/PAI", "CODIFICADO"], ascending=[True, False])


_PUNTO_SEMAFORO = {"Verde": "🟢", "Amarillo": "🟡", "Rojo": "🔴"}


def _tabla_resumen_presupuesto(df, agrupar_por="gasto"):
    resumen = _resumen_grupos_presupuesto(df, agrupar_por)
    salida = resumen.copy()
    for col in ["CODIFICADO", "COMPROMISO", "DEVENGADO", "SALDO_DISPONIBLE"]:
        salida[col] = salida[col].map(_moneda)
    salida["EJECUCION"] = salida["EJECUCION"].map(
        lambda valor: f"{valor:.2f}%".replace(".", ",")
    )
    # Solo un punto de color (sin el nombre "Rojo"/"Amarillo"/"Verde" como texto).
    salida["SEMAFORO"] = salida["SEMAFORO"].map(lambda v: _PUNTO_SEMAFORO.get(v, "⚪"))
    salida = salida.rename(columns={
        "CODIFICADO": "Codificado",
        "COMPROMISO": "Comprometido",
        "DEVENGADO": "Devengado",
        "SALDO_DISPONIBLE": "Saldo disponible",
        "EJECUCION": "Ejecución",
        "SEMAFORO": "Semáforo",
    })
    return salida


def _fila_total_presupuesto(df, agrupar_por="gasto"):
    """Construye la fila TOTAL como un registro más de la tabla (mismas columnas)."""
    etiqueta_columna = _etiqueta_columna_agrupacion(agrupar_por)
    codificado = float(df["CODIFICADO"].sum()) if not df.empty else 0.0
    compromiso = float(df["COMPROMISO"].sum()) if not df.empty else 0.0
    devengado = float(df["DEVENGADO"].sum()) if not df.empty else 0.0
    saldo = float(df["SALDO_DISPONIBLE"].sum()) if not df.empty else 0.0
    ejecucion = (devengado / codificado * 100) if codificado else 0.0
    semaforo = "Verde" if ejecucion >= 70 else "Amarillo" if ejecucion >= 40 else "Rojo"
    return {
        "POA/PAI": "TOTAL",
        etiqueta_columna: "",
        "Codificado": _moneda(codificado),
        "Comprometido": _moneda(compromiso),
        "Devengado": _moneda(devengado),
        "Saldo disponible": _moneda(saldo),
        "Ejecución": f"{ejecucion:.2f}%".replace(".", ","),
        "Semáforo": _PUNTO_SEMAFORO.get(semaforo, "⚪"),
    }


def _tabla_con_total(df, agrupar_por="gasto"):
    """Detalle por grupo/proyecto + una fila TOTAL final, lista para el DataTable."""
    detalle = _tabla_resumen_presupuesto(df, agrupar_por)
    registros = detalle.to_dict("records")
    registros.append(_fila_total_presupuesto(df, agrupar_por))
    return detalle, registros


def _figura_ejecucion_general(df):
    """Anillo con el porcentaje devengado respecto del codificado."""
    codificado = float(df["CODIFICADO"].sum()) if not df.empty else 0.0
    devengado = float(df["DEVENGADO"].sum()) if not df.empty else 0.0
    porcentaje = min((devengado / codificado * 100) if codificado else 0.0, 100.0)
    figura = go.Figure(go.Pie(
        labels=["Ejecutado (devengado)", "Pendiente por ejecutar"],
        values=[porcentaje, max(100 - porcentaje, 0)], hole=.68,
        marker={"colors": ["#5442a3", "#e8eaf2"]},
        textinfo="none", hovertemplate="%{label}: %{value:.2f}%<extra></extra>",
    ))
    figura.add_annotation(
        text=f"<b>{porcentaje:.1f}%</b><br><span style='font-size:11px'>ejecutado</span>",
        x=.5, y=.5, showarrow=False, font={"size": 24, "color": "#17245b"},
    )
    figura.update_layout(
        title={"text": "Avance de ejecución presupuestaria", "x": .5,
               "font": {"size": 16, "color": "#17245b"}},
        height=300, margin=dict(l=20, r=20, t=55, b=48),
        legend={"orientation": "h", "x": .5, "xanchor": "center", "y": -.05,
                "font": {"size": 10}}, paper_bgcolor="white",
        font={"family": "Arial", "color": "#17245b"},
    )
    return figura


def _millones(valor):
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _figura_montos_presupuesto(df):
    """Barra con los montos consolidados (en millones de USD): Codificado,
    Certificado (PRECOMPROMISO), Comprometido, Devengado y Saldo disponible."""
    # Paleta validada (dataviz: banda de luminosidad, piso de croma, separación
    # CVD y contraste vs. fondo) — reutiliza el morado del pastel de ejecución
    # (#5442a3 = "Devengado") para mantener coherencia visual entre ambos gráficos.
    metricas = [
        ("Codificado", "CODIFICADO", "#2f6fed"),
        ("Certificado", "PRECOMPROMISO", "#1f9e6f"),
        ("Comprometido", "COMPROMISO", "#c9860a"),
        ("Devengado", "DEVENGADO", "#5442a3"),
        ("Saldo disponible", "SALDO_DISPONIBLE", "#de232d"),
    ]
    etiquetas = [nombre for nombre, _, _ in metricas]
    colores = [color for _, _, color in metricas]
    valores = [
        (float(df[col].sum()) / 1_000_000) if (not df.empty and col in df.columns) else 0.0
        for _, col, _ in metricas
    ]
    montos = [v * 1_000_000 for v in valores]
    figura = go.Figure(go.Bar(
        x=etiquetas, y=valores, marker={"color": colores},
        text=[f"{_millones(v)} M" for v in valores], textposition="outside",
        customdata=[_moneda(m) for m in montos],
        hovertemplate="<b>%{x}</b><br>%{customdata}<extra></extra>",
    ))
    figura.update_layout(
        title={"text": "Montos presupuestarios (millones de USD) · Inversión", "x": .5,
               "font": {"size": 16, "color": "#17245b"}},
        height=300, margin=dict(l=55, r=25, t=55, b=45),
        paper_bgcolor="white", plot_bgcolor="white", showlegend=False,
        font={"family": "Arial", "color": "#17245b"},
    )
    figura.update_xaxes(showgrid=False, tickfont={"size": 10.5})
    # Se agrega un 20% de margen superior sobre el valor máximo para que la
    # etiqueta "outside" de la barra más alta no quede recortada por el borde
    # del gráfico.
    tope = (max(valores) * 1.2) if valores and max(valores) > 0 else 1
    figura.update_yaxes(title="Millones de USD", gridcolor="#e4e7f0",
                        tickformat=",.1f", tickfont={"size": 10},
                        range=[0, tope])
    return figura


def contenido_presupuesto(vice=None):
    actualizado = _formatear_fecha_actualizacion()
    insignia_actualizacion = (
        html.Span([html.Span("Corte de datos: ", className="budget-updated-label"), actualizado],
                  className="budget-updated-badge")
        if actualizado else None
    )
    if ERROR_PRESUPUESTO or DATA_PRESUPUESTO.empty:
        return html.Div([
            html.Div([html.P("EJECUCIÓN PRESUPUESTARIA", className="content-kicker"),
                      html.H1("Seguimiento presupuestario"),
                      html.P("Información diaria del reporte ESIGEF.")], className="content-heading"),
            html.Div([html.H2("Conexión pendiente"),
                      html.P(ERROR_PRESUPUESTO or "No existen registros disponibles."),
                      html.P("Verifique que el vínculo de OneDrive permita descargar el archivo sin iniciar sesión.")],
                     className="load-error")
        ], className="indicator-content")
    if not vice:
        return html.Div([
            html.Div([html.P("EJECUCIÓN PRESUPUESTARIA", className="content-kicker"),
                      html.H1("Seguimiento presupuestario"),
                      html.P("Información consolidada de la ejecución presupuestaria."),
                      insignia_actualizacion],
                     className="content-heading"),
            html.Div([html.H2("Seleccione un viceministerio"),
                      html.P("Use el menú lateral para consultar su ejecución, composición y detalle presupuestario.")],
                     className="module-welcome")
        ], className="indicator-content")

    es_general = vice == "General"
    if es_general:
        vice_normalizado = (DATA_PRESUPUESTO["Viceministerio"].fillna("").astype(str)
                            .map(_normalizar_encabezado_presupuesto))
        base = DATA_PRESUPUESTO.loc[
            ~vice_normalizado.isin({"SIN_CLASIFICACION", "SIN_CLASIFICAR", "NO_APLICA", ""})
        ].copy()
    else:
        base = DATA_PRESUPUESTO.loc[DATA_PRESUPUESTO["Viceministerio"] == vice].copy()
    # Todo el módulo trabaja únicamente sobre Inversión (ya no existe la vista
    # "Corriente"): tanto en General como en cada viceministerio.
    grupo = _filtrar_alcance_presupuesto(base, "Inversión")
    totales = grupo[PRESUPUESTO_MONETARIAS].sum()
    codificado = float(totales["CODIFICADO"])
    devengado = float(totales["DEVENGADO"])
    ejecucion = (devengado / codificado * 100) if codificado else 0.0

    certificado = float(totales["PRECOMPROMISO"]) if "PRECOMPROMISO" in totales else 0.0
    compromiso = float(totales["COMPROMISO"])
    saldo = float(totales["SALDO_DISPONIBLE"])
    kpis = [
        ("codificado", "Codificado", _moneda(codificado), "Presupuesto vigente"),
        ("certificado", "Certificado", _moneda(certificado), "Precompromiso"),
        ("comprometido", "Comprometido", _moneda(compromiso), "Obligaciones registradas"),
        ("devengado", "Devengado", _moneda(devengado), "Monto ejecutado"),
        ("saldo", "Saldo disponible", _moneda(saldo), "Recursos por utilizar"),
        ("ejecucion", "% de ejecución", f"{ejecucion:.2f}%".replace(".", ","),
         "Devengado / codificado"),
    ]

    iconos_kpi = {
        "asignado": "<path d='M8 18h16M10 18V8h12v10M13 13h2m3 0h2M7 22h18'/>",
        "codificado": "<path d='M9 5h10l4 4v14H9zM19 5v5h4M13 14h6m-6 4h6'/>",
        "certificado": "<path d='M15 4l2.6 5.3 5.8.9-4.2 4.1 1 5.8-5.2-2.8-5.2 2.8 1-5.8-4.2-4.1 5.8-.9z'/><path d='M12 21l1.5 4M18 21l-1.5 4'/>",
        "comprometido": "<path d='M6 13l5 5L22 7M5 4h20v20H5z'/>",
        "devengado": "<circle cx='15' cy='15' r='10'/><path d='M11 15l3 3 6-7'/>",
        "saldo": "<path d='M5 10h20v13H5zM8 10V7h14v3M9 16h8m4 0h1'/>",
        "ejecucion": "<path d='M8 22L22 8M10 8h.01M20 22h.01'/><circle cx='10' cy='8' r='3'/><circle cx='20' cy='22' r='3'/>",
    }

    def icono_kpi(nombre):
        svg = ("<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 30 30' "
               "fill='none' stroke='#4d3a94' stroke-width='1.8' stroke-linecap='round' "
               f"stroke-linejoin='round'>{iconos_kpi[nombre]}</svg>")
        # Ancho/alto fijos: un <img> de SVG sin esos atributos puede caer al
        # tamaño por defecto del navegador (300x150) y tapar el texto vecino.
        return html.Img(src="data:image/svg+xml;utf8," + quote(svg), alt="",
                        style={"width": "14px", "height": "14px", "display": "block"},
                        **{"aria-hidden": "true"})

    def punto(color):
        return html.Span(style={
            "display": "inline-block", "width": "6px", "height": "6px",
            "borderRadius": "50%", "background": color, "marginRight": "3px",
        })

    leyenda_semaforo = html.Div([
        html.Span("Semáforo:", style={"fontWeight": "700", "marginRight": "3px"}),
        html.Span([punto("#df3e5b"), "menor a 40%"], style={"marginRight": "8px"}),
        html.Span([punto("#efb629"), "40% a 69,99%"], style={"marginRight": "8px"}),
        html.Span([punto("#2fbd83"), "70% o más"]),
    ], style={
        "textAlign": "center", "color": "#9aa1b3", "fontSize": "9px",
        "marginTop": "8px", "display": "flex", "alignItems": "center",
        "justifyContent": "center", "flexWrap": "wrap",
    })

    agrupar_por_inicial = "proyecto"
    detalle_sin_total, registros_tabla = _tabla_con_total(grupo, agrupar_por_inicial)
    columnas = list(detalle_sin_total.columns)
    etiqueta_columna_inicial = _etiqueta_columna_agrupacion(agrupar_por_inicial)
    titulo_pagina = "Visión general de la gestión" if es_general else vice
    subtitulo_pagina = ("Resumen consolidado de la ejecución presupuestaria de Inversión."
                        if es_general else "Ejecución presupuestaria de Inversión.")

    texto_explicativo_style = {"fontSize": "13px", "color": "#4a5170", "margin": "0"}

    # Misma visualización en General y en cada viceministerio: anillo de
    # ejecución + montos consolidados (en millones), siempre solo Inversión.
    visuales = html.Section([
        html.Div([
            html.H2("Resumen de ejecución"),
            html.P("El porcentaje de ejecución corresponde a Devengado / Codificado × 100. "
                   "Cifras de Inversión.", style=texto_explicativo_style),
        ], className="budget-visual-header"),
        html.Div([
            html.Div(dcc.Graph(figure=_figura_ejecucion_general(grupo),
                               config={"displayModeBar": False, "responsive": False},
                               style={"height": "300px", "width": "100%"}),
                     className="budget-panel", style={"height": "300px"}),
            html.Div(dcc.Graph(figure=_figura_montos_presupuesto(grupo),
                               config={"displayModeBar": False, "responsive": False},
                               style={"height": "300px", "width": "100%"}),
                     className="budget-panel", style={"height": "300px"}),
        ], className="budget-general-chart-grid", style={
            "display": "grid", "gridTemplateColumns": "minmax(280px, .65fr) minmax(0, 1.35fr)",
            "gap": "14px",
        }),
        leyenda_semaforo,
    ], className="budget-visual-section")

    selector_tabla = html.Div([
        html.Span("Solo Inversión", className="budget-investment-badge"),
        dcc.RadioItems(
            id="budget-table-scope",
            options=[{"label": "Grupo de gasto", "value": "gasto"},
                     {"label": "Proyecto", "value": "proyecto"}],
            value=agrupar_por_inicial, inline=True, className="budget-table-filter",
        ),
    ], style={"display": "flex", "alignItems": "center", "gap": "10px", "flexWrap": "wrap"})

    return html.Div([
        html.Div([html.P("EJECUCIÓN PRESUPUESTARIA", className="content-kicker"),
                  html.H1(titulo_pagina),
                  html.P(subtitulo_pagina),
                  insignia_actualizacion],
                 className="content-heading"),
        html.Div([
            html.Div([
                html.Span(icono_kpi(codigo), className="budget-kpi-icon", style={
                    "display": "flex", "alignItems": "center", "justifyContent": "center",
                    "width": "24px", "height": "24px", "flex": "0 0 24px",
                    "borderRadius": "7px", "background": "rgba(241, 182, 32, .25)",
                }),
                html.Div([
                    html.Span(titulo, style={
                        "display": "block", "fontSize": "8.5px", "fontWeight": "800",
                        "textTransform": "uppercase", "letterSpacing": ".02em",
                        "color": "#7a6008", "whiteSpace": "normal",
                    }),
                    html.Strong(valor, style={
                        "display": "block", "fontSize": "clamp(11px, .95vw, 14px)",
                        "color": "#232d5a", "margin": "2px 0 1px", "overflowWrap": "anywhere",
                    }),
                    html.Small(nota, style={
                        "display": "block", "fontSize": "8px", "color": "#806000",
                    }),
                ], style={"minWidth": "0"})
            ], className="budget-kpi-card", style={
                "display": "grid", "gridTemplateColumns": "24px minmax(0, 1fr)",
                "alignItems": "center", "gap": "8px",
                "minHeight": "64px", "padding": "8px 10px", "minWidth": "0",
            })
            for codigo, titulo, valor, nota in kpis
        ], className="budget-kpi-grid", style={
            "display": "grid",
            "gridTemplateColumns": f"repeat({len(kpis)}, minmax(0, 1fr))",
            "gap": "8px",
        }),
        visuales,
        html.Section([
            html.Div([html.H2("Detalle presupuestario"),
                      html.Div([
                          selector_tabla,
                          html.Span(f"{len(detalle_sin_total):,} grupos".replace(",", "."),
                                    id="budget-table-count"),
                      ], className="budget-table-actions")],
                     className="budget-table-title"),
            dash_table.DataTable(
                id="budget-detail-table",
                data=registros_tabla,
                columns=[{"name": c, "id": c} for c in columnas],
                page_size=20, sort_action="none", filter_action="none",
                fixed_rows={"headers": True},
                style_table={"overflowX": "auto", "maxHeight": "620px"},
                style_cell={"fontFamily": "Arial", "fontSize": "12px", "padding": "9px",
                            "minWidth": "110px", "maxWidth": "300px", "whiteSpace": "normal",
                            "textAlign": "right"},
                style_cell_conditional=[
                    {"if": {"column_id": "POA/PAI"}, "textAlign": "left", "fontWeight": "700"},
                    {"if": {"column_id": ["Grupo de gasto", "Proyecto"]}, "textAlign": "left",
                     "minWidth": "230px"},
                    {"if": {"column_id": "Semáforo"}, "textAlign": "center", "fontSize": "15px",
                     "minWidth": "60px", "maxWidth": "60px", "padding": "0"},
                ],
                style_header={"backgroundColor": "#17245b", "color": "white", "fontWeight": "800",
                              "border": "1px solid #384273", "textAlign": "center"},
                style_data_conditional=[
                    {"if": {"row_index": "odd"}, "backgroundColor": "#f7f7fb"},
                    {"if": {"filter_query": "{Ejecución} contains '0,00%'", "column_id": "Ejecución"},
                     "color": "#b82043", "fontWeight": "800"},
                    {"if": {"filter_query": '{POA/PAI} = "TOTAL"'}, "backgroundColor": "#17245b",
                     "color": "#ffffff", "fontWeight": "800", "border": "1px solid #384273"},
                ],
            )
        ], className="budget-table-card budget-table-full"),
        html.P(f"Fuente: ESIGEF · {PRESUPUESTO_META.get('origen', '')}", className="budget-source")
    ], className="indicator-content budget-content")


# ---------------------------------------------------------------------------
# Enrutamiento genérico de módulos
# ---------------------------------------------------------------------------
def contenido_inventario():
    """Enlaces directos a los portales de datos abiertos por ámbito."""
    enlaces = [
        ("Educación Media", "https://educacion.gob.ec/datos-abiertos-minedec/"),
        ("Educación Superior", "https://siau.senescyt.gob.ec/portal-de-indicadores-de-educacion-superior/"),
        ("Cultura", "https://siic.culturaypatrimonio.gob.ec/que-es-la-cultura-en-cifras/"),
    ]
    return html.Div([
        html.Div([html.P("INVENTARIO Y RECURSO DE INFORMACIÓN", className="content-kicker"),
                  html.H1("Recursos de información abiertos"),
                  html.P("Seleccione un ámbito para abrir su portal de datos en una pestaña nueva.")],
                 className="content-heading"),
        html.Div([
            html.A(nombre, href=url, target="_blank", rel="noopener noreferrer", style={
                "display": "flex", "alignItems": "center", "justifyContent": "center",
                "flex": "1 1 220px", "minWidth": "220px", "minHeight": "90px",
                "background": "#4f449a", "color": "#fff", "fontWeight": "800",
                "fontSize": "15px", "borderRadius": "12px", "textDecoration": "none",
                "boxShadow": "0 8px 20px rgba(35,45,90,.18)", "textAlign": "center",
                "padding": "18px", "cursor": "pointer",
            })
            for nombre, url in enlaces
        ], style={"display": "flex", "gap": "16px", "flexWrap": "wrap", "marginTop": "20px"}),
    ], className="indicator-content")


def contenido_documentacion():
    """Repositorio documental exigido para el entregable E-02 (DP-SI-020)."""
    icono_documento = (
        "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 48 48' "
        "fill='none' stroke='#503a98' stroke-width='2.6' "
        "stroke-linecap='round' stroke-linejoin='round'>"
        "<path d='M13 6h15l8 8v28H13z'/><path d='M28 6v9h9'/>"
        "<path d='M19 23h11M19 29h11M19 35h8'/></svg>"
    )

    def tarjeta_entregable(codigo, titulo, descripcion, formato, nombre_archivo):
        return html.Article([
            html.Div([
                html.Span(
                    html.Img(
                        src="data:image/svg+xml;utf8," + quote(icono_documento),
                        className="document-file-icon-svg", alt="",
                    ),
                    className="document-file-icon", **{"aria-hidden": "true"},
                ),
            ], className="document-card-top"),
            html.H3(titulo),
            html.P(descripcion),
            html.Div([html.Span("Formato: "), html.Strong(formato)], className="document-format"),
            html.Div([
                html.Span("• Disponible", className="document-status available"),
                html.A(
                    [html.Span("Descargar archivo"), html.Span("↓", **{"aria-hidden": "true"})],
                    href=app.get_relative_path(
                        "/descargar-entregable/" + quote(nombre_archivo)
                    ),
                    className="document-action",
                    download=nombre_archivo,
                ),
            ], className="document-card-actions"),
        ], className="document-deliverable-card")

    entregables = [
        tarjeta_entregable(
            "E-03", "Manual de usuario",
            "Cómo navegar el portal, leer sus tableros, usar los filtros y descargar gráficos y tablas.",
            "DOCX", "E-03_Manual_de_usuario.docx",
        ),
        tarjeta_entregable(
            "E-04", "Diccionario de datos e indicadores",
            "Variables de cada archivo de datos del portal y definición de cada indicador: fórmula, unidad, fuente y desagregaciones.",
            "XLSX", "E-04_Diccionario_de_datos_e_indicadores_Anexo5.xlsx",
        ),
        tarjeta_entregable(
            "E-05", "Fichas metodológicas de los indicadores",
            "Archivo único con las fichas metodológicas de todos los indicadores publicados en el portal.",
            "DOCX", "E-05_Fichas_metodologicas.docx",
        ),
    ]
    return html.Div([
        html.Section([
            html.Div([
                html.H1("Documentación"),
                html.P("Manual de usuario, diccionario de datos e indicadores y fichas metodológicas de la solución de información."),
            ]),
        ], className="documentation-hero"),
        html.Section([
            html.Div([
                html.H2("DOCUMENTOS DISPONIBLES"),
            ], className="documentation-section-heading"),
            html.Div(entregables, className="document-deliverables-grid"),
        ], className="documentation-section"),
    ], className="documentation-page")


def contenido_seccion(seccion, vice=None, indicador=None):
    if seccion == "vision":
        return contenido_vision(vice)
    if seccion == "pnd":
        return contenido_pnd(vice, indicador)
    if seccion == "kpi-estrategicos":
        return contenido_kpi(vice, indicador)
    if seccion == "kpi-institucionales":
        return contenido_kpi_inst(vice, indicador)
    if seccion == "presupuesto":
        return contenido_presupuesto(vice)
    if seccion == "inventario":
        return contenido_inventario()
    if seccion == "documentacion":
        return contenido_documentacion()
    return html.Div([html.H1(dict(SECCIONES).get(seccion, "Módulo")),
                      html.P("Módulo preparado para la siguiente etapa.")],
                     className="module-welcome")


def modulo_seccion(seccion, vice=None, indicador=None):
    return html.Div([menu_lateral(seccion, vice, indicador),
                     html.Main(contenido_seccion(seccion, vice, indicador), id="module-detail",
                               className="module-main")],
                    className="module-layout")


app.layout = html.Div([
    dcc.Store(id="active-section", data="home"), dcc.Store(id="selected-vice"),
    dcc.Store(id="selected-indicator"),
    dcc.Store(id="data-version", data={"pnd": VERSION_PND, "kpi-estrategicos": VERSION_KPI,
                                        "kpi-institucionales": VERSION_KPI_INST,
                                        "gestion-educativa": VERSION_GESTION_EDUCATIVA,
                                        "vision": VERSION_VISION,
                                        "presupuesto": VERSION_PRESUPUESTO}),
    dcc.Interval(id="excel-watcher", interval=15000, n_intervals=0),
    barra_superior(),
    html.Div(id="main-content", children=portada()),
    html.Div(id="mathjax-trigger", style={"display": "none"}),
], className="app-shell")


@app.callback(Output("active-section", "data"), Output("selected-vice", "data", allow_duplicate=True),
              Output("selected-indicator", "data", allow_duplicate=True),
              Input("home-button", "n_clicks"),
              Input({"type": "home-card", "index": ALL}, "n_clicks"),
              Input({"type": "indicator-home-card", "index": ALL}, "n_clicks"),
              Input({"type": "side-section", "index": ALL}, "n_clicks"),
              State("active-section", "data"), prevent_initial_call=True)
def cambiar_seccion(_home, _cards, _indicator_cards, _side, actual):
    disparador = ctx.triggered_id
    if disparador == "home-button":
        return ("home", None, None) if _home else (actual, no_update, no_update)
    if isinstance(disparador, dict):
        tipo = disparador.get("type")
        if tipo == "home-card":
            valores = _cards
        elif tipo == "indicator-home-card":
            valores = _indicator_cards
        else:
            valores = _side
        # La creación dinámica de botones produce eventos con cero clics.
        # Se ignoran para que el usuario nunca sea expulsado de la pantalla actual.
        if not valores or not any((v or 0) > 0 for v in valores):
            return actual, no_update, no_update
        nueva_seccion = disparador["index"]
        if nueva_seccion == actual:
            return actual, no_update, no_update
        # Al cambiar de sección se limpia la selección de vice/indicador anterior,
        # para no arrastrar un viceministerio que no existe en la nueva sección.
        return nueva_seccion, None, None
    return actual, no_update, no_update


@app.callback(Output("active-section", "data", allow_duplicate=True),
              Input("btn-volver-portada", "n_clicks"), prevent_initial_call=True)
def volver_desde_indicadores(n_clicks):
    return "home" if n_clicks else no_update


@app.callback(Output("active-section", "data", allow_duplicate=True),
              Input("side-back", "n_clicks"), prevent_initial_call=True)
def volver_al_panel(n_clicks):
    return "home" if n_clicks else no_update


@app.callback(Output("selected-vice", "data"), Output("selected-indicator", "data", allow_duplicate=True),
              Input({"type": "vice-button", "index": ALL}, "n_clicks"),
              State("selected-vice", "data"), State("active-section", "data"), prevent_initial_call=True)
def seleccionar_vice(_clicks, actual, seccion):
    cfg = SECCIONES_CON_DATOS.get(seccion)
    df, _ = obtener_datos(seccion)
    if (not isinstance(ctx.triggered_id, dict) or not cfg or df.empty or not _clicks
            or not any((v or 0) > 0 for v in _clicks)):
        return actual, no_update
    vice = ctx.triggered_id["index"]
    if not (seccion == "presupuesto" and vice == "General") and \
            vice not in set(df[cfg["vice_col"]].dropna().unique()):
        return actual, no_update
    # Segundo clic sobre el mismo viceministerio: recoge el acordeón y limpia la ficha.
    if vice == actual:
        return None, None
    if seccion == "presupuesto":
        return vice, None
    indicadores = df.loc[df[cfg["vice_col"]] == vice, cfg["indicador_col"]].dropna().unique()
    primer_indicador = indicadores[0] if len(indicadores) else None
    return vice, primer_indicador


@app.callback(Output("selected-vice", "data", allow_duplicate=True),
              Input({"type": "vision-tab", "index": ALL}, "n_clicks"),
              State("active-section", "data"), prevent_initial_call=True)
def seleccionar_pestana_vision(_clicks, seccion):
    """Pestañas de Visión Ejecutiva (reutiliza el Store 'selected-vice', que
    en esta sección no se usa para viceministerios sino para la pestaña
    activa)."""
    if (seccion != "vision" or not isinstance(ctx.triggered_id, dict)
            or not _clicks or not any((v or 0) > 0 for v in _clicks)):
        return no_update
    return ctx.triggered_id["index"]


@app.callback(Output("selected-indicator", "data"),
              Input({"type": "indicator-button", "index": ALL}, "n_clicks"),
              State("selected-vice", "data"), State("selected-indicator", "data"),
              State("active-section", "data"), prevent_initial_call=True)
def seleccionar_indicador(_clicks, vice, actual, seccion):
    cfg = SECCIONES_CON_DATOS.get(seccion)
    df, _ = obtener_datos(seccion)
    plano = seccion in SECCIONES_INDICADORES_PLANOS
    if (not isinstance(ctx.triggered_id, dict) or not cfg or (not vice and not plano) or not _clicks
            or not any((v or 0) > 0 for v in _clicks)):
        return actual
    indicador = ctx.triggered_id["index"]
    if plano:
        indicadores_validos = set(df[cfg["indicador_col"]].dropna().unique())
    else:
        indicadores_validos = set(df.loc[df[cfg["vice_col"]] == vice, cfg["indicador_col"]].dropna().unique())
    return indicador if indicador in indicadores_validos else actual


@app.callback(Output("main-content", "children"), Input("active-section", "data"),
              Input("selected-vice", "data"), Input("selected-indicator", "data"),
              Input("data-version", "data"))
def mostrar_pantalla(seccion, vice, indicador, _version):
    """Reconstruye la sección completa (menú lateral + contenido) en un solo callback.
    Antes existían callbacks separados para el contenido y para las clases del acordeón,
    y ambos se disparaban al cambiar de sección; como cada sección tiene un número distinto
    de viceministerios, competían por actualizar los mismos botones con tamaños distintos
    y Dash lanzaba 'Invalid number of output values'. Un único callback evita la carrera."""
    if seccion == "home":
        return portada()
    if seccion == "indicadores":
        return portada_indicadores()
    return modulo_seccion(seccion, vice, indicador)


@app.callback(
    Output("budget-detail-table", "data"),
    Output("budget-detail-table", "columns"),
    Output("budget-table-count", "children"),
    Input("budget-table-scope", "value"),
    State("selected-vice", "data"),
    prevent_initial_call=True,
)
def filtrar_tabla_presupuesto(agrupar_por, vice):
    """Reagrupa la tabla por Grupo de gasto o por Proyecto; siempre solo
    Inversión. La fila TOTAL viaja dentro de los mismos datos."""
    agrupar_por = agrupar_por or "proyecto"
    if not vice or DATA_PRESUPUESTO.empty:
        detalle_vacio, registros_vacios = _tabla_con_total(pd.DataFrame(), agrupar_por)
        columnas_vacias = [{"name": c, "id": c} for c in detalle_vacio.columns]
        return registros_vacios, columnas_vacias, "0 grupos"
    if vice == "General":
        vice_normalizado = (DATA_PRESUPUESTO["Viceministerio"].fillna("").astype(str)
                            .map(_normalizar_encabezado_presupuesto))
        base = DATA_PRESUPUESTO.loc[
            ~vice_normalizado.isin({"SIN_CLASIFICACION", "SIN_CLASIFICAR", "NO_APLICA", ""})
        ].copy()
    else:
        base = DATA_PRESUPUESTO.loc[DATA_PRESUPUESTO["Viceministerio"] == vice].copy()
    filtrado = _filtrar_alcance_presupuesto(base, "Inversión")
    detalle, registros = _tabla_con_total(filtrado, agrupar_por)
    columnas = [{"name": c, "id": c} for c in detalle.columns]
    cantidad = f"{len(detalle):,} grupos".replace(",", ".")
    return registros, columnas, cantidad


@app.callback(Output("observation-area", "children"),
              Input({"type": "status-year", "index": ALL}, "n_clicks"),
              State("selected-indicator", "data"), prevent_initial_call=True)
def mostrar_observacion(_clicks, indicador):
    """Solo aplica al módulo PND, que registra alertas por año."""
    if (not isinstance(ctx.triggered_id, dict) or not indicador or not _clicks
            or not any((v or 0) > 0 for v in _clicks)):
        return no_update
    año = ctx.triggered_id["index"]
    fila = DATA_PND.loc[(DATA_PND["NOMBRE DEL INDICADOR"] == indicador)
                        & (DATA_PND["Año"].astype(str) == str(año))]
    if fila.empty:
        return "No existe información para el período seleccionado."
    return [html.Strong(f"Observación {año} · {primer_texto(fila, 'Alerta', 'Sin clasificación')}"),
            html.P(primer_texto(fila, "Observación", "Sin observación registrada."))]


@app.callback(Output("observation-area-kpi", "children"),
              Input({"type": "status-periodo-kpi", "index": ALL}, "n_clicks"),
              State("selected-indicator", "data"), prevent_initial_call=True)
def mostrar_observacion_kpi(_clicks, indicador):
    """Semáforo por período de KPI's Estratégicos (misma lógica que el PND, por Año+Mes)."""
    if (not isinstance(ctx.triggered_id, dict) or not indicador or not _clicks
            or not any((v or 0) > 0 for v in _clicks)):
        return no_update
    periodo_id = ctx.triggered_id["index"]
    grupo = preparar_periodos_visuales(
        DATA_KPI.loc[DATA_KPI["NOMBRE DEL INDICADOR"] == indicador]
    )
    fila = grupo.loc[grupo["Periodo_id"] == periodo_id]
    if fila.empty:
        return "No existe información para el período seleccionado."
    periodo = primer_texto(fila, "Periodo", "Período seleccionado")
    return [html.Strong(f"Observación {periodo} · {primer_texto(fila, 'Alerta', 'Sin clasificación')}"),
            html.P(primer_texto(fila, "Observación", "Sin observación registrada."))]


@app.callback(Output("observation-area-kpi-inst", "children"),
              Input({"type": "status-periodo-inst", "index": ALL}, "n_clicks"),
              State("selected-indicator", "data"), prevent_initial_call=True)
def mostrar_observacion_kpi_inst(_clicks, indicador):
    """Igual que el semáforo del PND, pero por período (Mes + Año) en vez de solo Año."""
    if (not isinstance(ctx.triggered_id, dict) or not indicador or not _clicks
            or not any((v or 0) > 0 for v in _clicks)):
        return no_update
    periodo_id = ctx.triggered_id["index"]
    grupo = preparar_periodos_visuales(
        DATA_KPI_INST.loc[DATA_KPI_INST["NOMBRE DEL INDICADOR"] == indicador]
    )
    fila = grupo.loc[grupo["Periodo_id"] == periodo_id]
    if fila.empty:
        return "No existe información para el período seleccionado."
    periodo = primer_texto(fila, "Periodo", "Período seleccionado")
    return [html.Strong(f"Observación {periodo} · {primer_texto(fila, 'Alerta', 'Sin clasificación')}"),
            html.P(primer_texto(fila, "Observación", "Sin observación registrada."))]


@app.callback(Output("data-version", "data"), Input("excel-watcher", "n_intervals"),
              State("data-version", "data"), prevent_initial_call=True)
def actualizar_excel(_intervalo, version_actual):
    """Recarga únicamente el Excel que cambió; no interrumpe los clics del usuario."""
    global DATA_PND, ERROR_PND, VERSION_PND, DATA_KPI, ERROR_KPI, VERSION_KPI
    global DATA_KPI_INST, ERROR_KPI_INST, VERSION_KPI_INST
    global RESUMEN_GESTION_EDUCATIVA, PENDIENTES_GESTION_EDUCATIVA, VERSION_GESTION_EDUCATIVA
    global DATA_PRESUPUESTO, ERROR_PRESUPUESTO, VERSION_PRESUPUESTO
    version_actual = dict(version_actual or {})
    cambio = False

    nueva_version_pnd = version_pnd()
    if nueva_version_pnd and nueva_version_pnd != version_actual.get("pnd"):
        nueva_data, nuevo_error = cargar_base_pnd()
        if not nuevo_error and not nueva_data.empty:
            DATA_PND, ERROR_PND, VERSION_PND = nueva_data, None, nueva_version_pnd
            version_actual["pnd"] = nueva_version_pnd
            cambio = True

    nueva_version_kpi = version_kpi()
    if nueva_version_kpi and nueva_version_kpi != version_actual.get("kpi-estrategicos"):
        nueva_data, nuevo_error = cargar_base_kpi()
        if not nuevo_error and not nueva_data.empty:
            DATA_KPI, ERROR_KPI, VERSION_KPI = nueva_data, None, nueva_version_kpi
            version_actual["kpi-estrategicos"] = nueva_version_kpi
            cambio = True

    nueva_version_kpi_inst = version_kpi_inst()
    if nueva_version_kpi_inst and nueva_version_kpi_inst != version_actual.get("kpi-institucionales"):
        nueva_data, nuevo_error = cargar_base_kpi_inst()
        if not nuevo_error and not nueva_data.empty:
            DATA_KPI_INST, ERROR_KPI_INST, VERSION_KPI_INST = nueva_data, None, nueva_version_kpi_inst
            version_actual["kpi-institucionales"] = nueva_version_kpi_inst
            cambio = True

    # Las bases de Gestión Educativa (algunas en OneDrive) se vuelven a leer
    # cada 15 minutos, igual que el presupuesto: no hay forma barata de saber
    # si cambiaron sin descargarlas, así que se relee y se compara después.
    if _intervalo % 60 == 0:
        nuevo_resumen, nuevos_pendientes = cargar_resumen_gestion_educativa()
        nueva_version_gestion_educ = version_gestion_educativa(nuevo_resumen, nuevos_pendientes)
        if nueva_version_gestion_educ != version_actual.get("gestion-educativa"):
            RESUMEN_GESTION_EDUCATIVA, PENDIENTES_GESTION_EDUCATIVA = nuevo_resumen, nuevos_pendientes
            VERSION_GESTION_EDUCATIVA = nueva_version_gestion_educ
            version_actual["gestion-educativa"] = nueva_version_gestion_educ
            cambio = True

    # SharePoint se consulta cada 15 minutos para evitar solicitudes innecesarias.
    if _intervalo % 60 == 0:
        nueva_data, nuevo_error = cargar_base_presupuesto()
        nueva_version = version_presupuesto()
        if (not nuevo_error and not nueva_data.empty
                and nueva_version != version_actual.get("presupuesto")):
            DATA_PRESUPUESTO, ERROR_PRESUPUESTO = nueva_data, None
            VERSION_PRESUPUESTO = nueva_version
            version_actual["presupuesto"] = nueva_version
            cambio = True

    return version_actual if cambio else no_update


app.clientside_callback("function(n){if(n)window.print();return window.dash_clientside.no_update;}",
                        Output("print-button", "title"), Input("print-button", "n_clicks"),
                        prevent_initial_call=True)


app.clientside_callback(
    """
    function(children) {
        function ampliarEcuaciones() {
            document.querySelectorAll(
                '.formula-equation mjx-container, ' +
                '.formula-markdown mjx-container[display="true"]'
            ).forEach(function(ecuacion) {
                ecuacion.style.setProperty('font-size', '16px', 'important');
                ecuacion.style.setProperty('color', '#000', 'important');
                ecuacion.style.setProperty('opacity', '1', 'important');
            });
        }
        if (window.MathJax && window.MathJax.typesetPromise) {
            setTimeout(function () {
                window.MathJax.typesetPromise().then(function () {
                    ampliarEcuaciones();
                    setTimeout(ampliarEcuaciones, 120);
                });
            }, 60);
        }
        return window.dash_clientside.no_update;
    }
    """,
    Output("mathjax-trigger", "children"),
    Input("main-content", "children"),
    prevent_initial_call=True,
)


if __name__ == "__main__":
    app.run(debug=False, use_reloader=False, host="127.0.0.1", port=8050,
            jupyter_mode="external")
