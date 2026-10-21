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
    "E-06_Anexo4_Ficha_formalizacion.docx",
}

# Fuentes individuales de OneDrive. Se leen por HTTPS tanto localmente como
# desde el servidor publicado a partir de GitHub.
PND_DOWNLOAD_URL = (
    "https://educacionec-my.sharepoint.com/:x:/g/personal/"
    "ricardo_castellanos_educacion_gob_ec/"
    "IQDzi44ILLX-SK4FCICFqdzTAdE_QO0OSjNVppH-qRznzq0"
    "?e=TkTQu5&download=1"
)
KPI_ESTRATEGICOS_DOWNLOAD_URL = (
    "https://educacionec-my.sharepoint.com/:x:/g/personal/"
    "ricardo_castellanos_educacion_gob_ec/"
    "IQAhMnqcBe4IQJP7ewr3ZObQAalNMmiIpYKi3jcTS3DuJpQ"
    "?e=4Tfw34&download=1"
)
KPI_INSTITUCIONALES_DOWNLOAD_URL = (
    "https://educacionec-my.sharepoint.com/:x:/g/personal/"
    "ricardo_castellanos_educacion_gob_ec/"
    "IQB3pmjVw-PZT5WJKBhUWYHlAcKdvPZcKeloHyma2JtFrwY"
    "?e=xsKWf6&download=1"
)
VERSIONES_INDICADORES_REMOTOS = {}


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
# Estas fuentes siguen aportando los conteos operativos. La vista ejecutiva usa,
# además, el último consolidado institucional de inversión por régimen.
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


def _origen_indicador_onedrive(url, nombre_archivo):
    """Descarga una matriz de indicadores y registra su hash de contenido."""
    try:
        origen, version, _ = _descargar_excel_onedrive(url, nombre_archivo)
        VERSIONES_INDICADORES_REMOTOS[nombre_archivo] = version
        print(f"[OneDrive] {nombre_archivo}: descarga correcta ({version[:12]}).")
        return origen, None
    except Exception as exc:
        mensaje = f"No fue posible descargar '{nombre_archivo}' desde OneDrive: {exc}"
        print(f"[OneDrive] ERROR: {mensaje}")
        return None, mensaje


# ---------------------------------------------------------------------------
# Carga de datos · Plan Nacional de Desarrollo
# ---------------------------------------------------------------------------
def cargar_base_pnd():
    """Lee el Excel del PND aunque Windows añada (1), (2), etc. al nombre."""
    # La matriz vigente indicada por el usuario es "Indicadores_PND version 1".
    # Se conserva el nombre anterior únicamente como respaldo para despliegues
    # donde todavía no se haya reemplazado el archivo.
    archivo, error_origen = _origen_indicador_onedrive(
        PND_DOWNLOAD_URL, "Indicadores_PND version 1.xlsx"
    )
    if error_origen:
        return pd.DataFrame(), error_origen
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

        # En los indicadores cuya unidad es "Número", el Excel puede guardar
        # conteos abreviados con decimales (107,56 representa 107.560). Se
        # recupera aquí el conteo entero para que tarjetas, barras y ejes usen
        # el punto como separador de miles: 107.560, 116.600, 357.730, etc.
        unidad_normalizada = df["Unidad de medida"].astype(str).map(
            lambda x: "".join(
                c for c in unicodedata.normalize("NFKD", x.strip().lower())
                if not unicodedata.combining(c)
            )
        )
        es_conteo = unidad_normalizada.eq("numero")
        for col in ["Línea base", "Meta", "Estimador"]:
            es_abreviado = (es_conteo & df[col].notna()
                            & (df[col].abs() < 1000)
                            & (df[col] % 1 != 0))
            df.loc[es_abreviado, col] = (df.loc[es_abreviado, col] * 1000).round()
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
    return VERSIONES_INDICADORES_REMOTOS.get("Indicadores_PND version 1.xlsx", 0)


# ---------------------------------------------------------------------------
# Carga de datos · Excel con estructura tipo PND (VICEMINISTERIO, NOMBRE DEL
# INDICADOR, Línea base, Meta, Estimador, Alerta, Observación), usada tanto por
# KPI's Estratégicos como por KPI's Institucionales, con "Mes" opcional
# (vacío para indicadores anuales, con valor para los mensuales/semestrales).
# ---------------------------------------------------------------------------
MES_ORDEN = {mes.lower(): i for i, mes in enumerate(MESES, start=1)}


def cargar_base_tipo_pnd(prefijo, nombre_legible, url=None):
    if url:
        archivo, error_origen = _origen_indicador_onedrive(url, nombre_legible)
        if error_origen:
            return pd.DataFrame(), error_origen
    else:
        archivo = buscar_archivo(prefijo)
        if not archivo:
            return pd.DataFrame(), f"No se encontró '{nombre_legible}'."
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

        # Las matrices institucionales también pueden registrar conteos
        # abreviados: por ejemplo, 107,56 representa 107.560 personas. Esta
        # corrección se aplica exclusivamente cuando la unidad es "Número";
        # porcentajes y demás unidades conservan su escala original.
        unidad_normalizada = df["Unidad de medida"].astype(str).map(
            lambda x: "".join(
                c for c in unicodedata.normalize("NFKD", x.strip().lower())
                if not unicodedata.combining(c)
            )
        )
        es_conteo = unidad_normalizada.eq("numero")
        for col in ["Línea base", "Meta", "Estimador"]:
            es_abreviado = (es_conteo & df[col].notna()
                            & (df[col].abs() < 1000)
                            & (df[col] % 1 != 0))
            df.loc[es_abreviado, col] = (df.loc[es_abreviado, col] * 1000).round()

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
    return cargar_base_tipo_pnd(
        "kpi_estratégicos", "kpi_estratégicos.xlsx", KPI_ESTRATEGICOS_DOWNLOAD_URL
    )


def version_kpi():
    return VERSIONES_INDICADORES_REMOTOS.get("kpi_estratégicos.xlsx", 0)


# ---------------------------------------------------------------------------
# Carga de datos · KPI's Institucionales
# ---------------------------------------------------------------------------
def cargar_base_kpi_inst():
    return cargar_base_tipo_pnd(
        "kpi_institucional", "KPI_INSTITUCIONALES.xlsx",
        KPI_INSTITUCIONALES_DOWNLOAD_URL,
    )


def version_kpi_inst():
    return VERSIONES_INDICADORES_REMOTOS.get("KPI_INSTITUCIONALES.xlsx", 0)


# ---------------------------------------------------------------------------
# Carga de datos · Gestión Educativa (Alimentación Escolar, Uniformes, Textos
# Escolares y Mobiliario y Transporte). A diferencia del PND/KPI, cada Excel
# es una tabla de beneficiarios por Zona/Provincia/Cantón (o por institución,
# en Mobiliario), sin la estructura de indicador con Meta/Línea base. Aquí
# desde estas bases se agregan conteos (beneficiarios, instituciones y
# estudiantes). Los montos que aparecen en la vista ejecutiva proceden del
# consolidado institucional por régimen, no de una suma implícita de estas
# matrices operativas.
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


def _columna_por_fragmentos(mapa, fragmentos, excluir=()):
    """Reconoce encabezados con prefijos o sufijos añadidos en la matriz."""
    for clave, original in mapa.items():
        if (any(fragmento in clave for fragmento in fragmentos)
                and not any(fragmento in clave for fragmento in excluir)):
            return original
    return None


def _metricas_gestion(df, mapa):
    dato = {}
    col_benef = (_columna_por_claves(
        mapa, ("BENEFICIARIOS", "TOTAL_BENEFICIARIOS", "NRO_BENEFICIARIOS")
    ) or _columna_por_fragmentos(mapa, ("BENEFICIAR",)))
    col_ie = (_columna_por_claves(
        mapa, ("NRO_IE", "NO_IE", "NO_I_E", "NUMERO_IE", "TOTAL_IE")
    ) or _columna_por_fragmentos(
        mapa, ("INSTITUCIONES_EDUCATIVAS", "INSTITUCION_EDUCATIVA"),
        excluir=("NOMBRE", "CODIGO"),
    ))
    col_amie = (_columna_por_claves(
        mapa, ("AMIE", "CODIGO_AMIE", "COD_AMIE", "CODIGO_INSTITUCION")
    ) or _columna_por_fragmentos(mapa, ("AMIE",)))
    col_nombre_ie = _columna_por_claves(
        mapa,
        ("NOMBRE_IE", "NOMBRE_INSTITUCION", "NOMBRE_INSTITUCION_EDUCATIVA"),
    )
    col_est = (_columna_por_claves(
        mapa,
        ("TOTAL_ESTUDIANTES", "ESTUDIANTES", "NRO_ESTUDIANTES",
         "NO_ESTUDIANTES", "NUMERO_ESTUDIANTES", "ESTUDIANTES_BENEFICIADOS"),
    ) or _columna_por_fragmentos(
        mapa, ("ESTUDIANT",), excluir=("NOMBRE", "CEDULA", "CODIGO"),
    ))
    col_prov = mapa.get("PROVINCIA")
    col_sost = _columna_por_claves(mapa, ("SOSTENIMIENTO", "TIPO_SOSTENIMIENTO"))
    col_area = _columna_por_claves(mapa, ("AREA", "AREA_GEOGRAFICA", "URBANO_RURAL"))

    if col_benef:
        dato["beneficiarios"] = float(pd.to_numeric(df[col_benef], errors="coerce").sum())
    if col_ie:
        dato["instituciones"] = float(pd.to_numeric(df[col_ie], errors="coerce").sum())
    elif col_amie:
        dato["instituciones"] = float(df[col_amie].dropna().astype(str).str.strip().nunique())
    elif col_nombre_ie:
        dato["instituciones"] = float(
            df[col_nombre_ie].dropna().astype(str).str.strip().replace("", pd.NA).nunique()
        )
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
    df = None
    if nombre_legible == "Mobiliario y Transporte Escolar":
        # El libro trae más de una hoja; la que sirve es la que tiene la
        # columna Clasificación (Equipamiento y mobiliario / Transporte
        # escolar). Sin esto se leía otra hoja y las tarjetas quedaban vacías.
        df = _localizar_hoja_con_columna(archivo, "CLASIFICACION")
    if df is None:
        df = _localizar_hoja_con_columna(archivo, "BENEFICIARIOS")
    if df is None:
        df = _localizar_hoja_con_columna(archivo, "TOTAL_ESTUDIANTES")
    if df is None:
        df = _localizar_hoja_con_columna(archivo, "ESTUDIANTES")
    if df is None:
        df = _localizar_hoja_con_columna(archivo, "NRO_IE")
    if df is None:
        raise ValueError(
            f"'{nombre_legible}' no tiene una columna de Beneficiarios ni "
            "de Total Estudiantes reconocible."
        )
    df.columns = [" ".join(str(c).strip().split()) for c in df.columns]
    mapa = {_normalizar_encabezado_presupuesto(c): c for c in df.columns}
    dato = _metricas_gestion(df, mapa)
    col_regimen = next((c for k, c in mapa.items() if "REGIMEN" in k), None)
    col_clasificacion = _columna_por_claves(
        mapa,
        ("CLASIFICACION", "TIPO_RECURSO", "RECURSO", "COMPONENTE",
         "TIPO_INTERVENCION"),
    )
    if col_clasificacion is None:
        col_clasificacion = next(
            (c for k, c in mapa.items() if "CLASIFICACION" in k), None
        )
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
        es_mobiliario = (
            clasificacion_normalizada.eq("EQUIPAMIENTO_Y_MOBILIARIO")
            | clasificacion_normalizada.str.contains("MOBILIARIO", na=False)
        )
        es_transporte = (
            clasificacion_normalizada.eq("TRANSPORTE_ESCOLAR")
            | clasificacion_normalizada.str.contains("TRANSPORTE", na=False)
        )
        dato["por_recurso"] = {
            "mobiliario": _metricas_gestion(
                df.loc[es_mobiliario], mapa
            ),
            "transporte": _metricas_gestion(
                df.loc[es_transporte], mapa
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
    separador_cache = "&" if "?" in url else "?"
    url_sin_cache = f"{url}{separador_cache}_ts={int(datetime.now().timestamp())}"
    respuesta = requests.get(
        url_sin_cache,
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
        <style>
        /* Información complementaria de SAFPI y NAP. Se mantiene la paleta y
           la tipografía de las tarjetas institucionales existentes. */
        .programa-con-ayuda { position: relative; overflow: visible !important; }
        .programa-info-tooltip {
          position: absolute; left: 14px; right: 14px;
          top: auto; bottom: calc(100% + 10px);
          z-index: 1001; padding: 12px 14px; border: 1px solid #dfe3ef;
          border-radius: 10px; background: #ffffff; color: #4d536b;
          box-shadow: 0 10px 26px rgba(24,37,87,.18);
          font-family: Arial, sans-serif; font-size: 12px; line-height: 1.5;
          text-align: justify; opacity: 0; visibility: hidden;
          transform: translateY(5px); transition: opacity .16s ease, transform .16s ease;
          pointer-events: none;
        }
        .programa-info-tooltip::after {
          content: ""; position: absolute; left: 28px; top: 100%;
          border: 7px solid transparent; border-top-color: #ffffff;
          filter: drop-shadow(0 2px 1px rgba(24,37,87,.10));
        }
        .programa-con-ayuda:hover .programa-info-tooltip,
        .programa-con-ayuda:focus-within .programa-info-tooltip {
          opacity: 1; visibility: visible; transform: translateY(0);
        }
        .programa-con-ayuda:hover,
        .programa-con-ayuda:focus-within { z-index: 1000 !important; }

        /* Portada de Visión Ejecutiva sin infografía: los cinco accesos ocupan
           el área central en dos filas equilibradas. */
        .vision-landing {
          justify-content: flex-start;
        }
        .vision-landing-intro {
          width: min(100%, 1120px); margin: 44px auto 0;
          text-align: center;
        }
        .vision-landing-subtitle {
          max-width: 720px; margin: 12px auto 0; color: #66708d;
          font-size: 15px; line-height: 1.5;
        }
        .vision-landing-menu {
          width: 100%; margin-top: 54px !important;
        }
        .vision-landing-menu .vision-vice-menu {
          display: grid !important;
          grid-template-columns: repeat(6, minmax(0, 1fr)) !important;
          gap: 18px !important; width: min(100%, 1120px) !important;
          max-width: 1120px !important; margin: 0 auto !important;
          padding: 0 0 24px !important; overflow: visible !important;
        }
        .vision-landing-menu .vision-vice-row {
          grid-column: span 2; min-width: 0 !important; padding: 0 !important;
          border: 0 !important; border-radius: 18px !important;
          background: transparent !important; box-shadow: none !important;
        }
        .vision-landing-menu .vision-vice-row:nth-child(4) {
          grid-column: 2 / span 2;
        }
        .vision-landing-menu .vision-vice-row:nth-child(5) {
          grid-column: 4 / span 2;
        }
        .vision-landing-menu .vision-vice-pill {
          display: grid !important;
          grid-template-columns: minmax(0, 1fr) !important;
          grid-template-rows: 1fr auto !important;
          align-items: center !important;
          width: 100% !important; min-height: 132px !important;
          padding: 20px !important; border: 1px solid #dfe3ef !important;
          border-top: 5px solid #4f3c98 !important;
          border-radius: 18px !important; background: #ffffff !important;
          color: #232d5a !important; text-align: center !important;
          box-shadow: 0 10px 26px rgba(24,37,87,.09) !important;
          cursor: pointer; transition: transform .16s ease, box-shadow .16s ease;
        }
        .vision-landing-menu .vision-vice-pill:hover,
        .vision-landing-menu .vision-vice-pill:focus-visible {
          transform: translateY(-4px);
          box-shadow: 0 16px 32px rgba(24,37,87,.16) !important;
          outline: 3px solid rgba(241,182,32,.35);
        }
        .vision-landing-vice-nombre {
          align-self: end; color: #232d5a; font-size: 15px;
          font-weight: 900; line-height: 1.25; text-align: center;
        }
        .vision-landing-vice-accion {
          align-self: start; margin-top: 6px; color: #7a829b;
          font-size: 11px; font-weight: 800; text-align: center;
        }
        .programa-anio-control {
          display: flex; align-items: center; gap: 5px;
          width: fit-content; margin: 0 0 5px;
        }
        .programa-anio-control > span {
          display: block; margin: 0; color: #4f3c98;
          font-size: 10px; font-weight: 900; text-transform: uppercase;
        }
        /* Dash 3.x usa el selector propio (.dash-dropdown); se cubren también
           las clases antiguas de react-select por compatibilidad. */
        .programa-anio-dropdown { width: 150px; min-width: 150px; }
        .programa-anio-dropdown,
        .programa-anio-dropdown.dash-dropdown,
        .programa-anio-dropdown .dash-dropdown,
        .programa-anio-dropdown .Select-control {
          height: 24px !important; min-height: 24px !important; max-height: 24px !important;
          padding-top: 0 !important; padding-bottom: 0 !important;
          border-color: #d8dced; border-radius: 6px;
          font-family: Arial, sans-serif; box-sizing: border-box;
        }
        .programa-anio-dropdown *,
        .programa-anio-dropdown .Select-value-label {
          font-size: 11px !important; font-weight: 800 !important;
          line-height: 22px !important; color: #202a59;
        }
        .programa-anio-dropdown .Select-value,
        .programa-anio-dropdown .Select-placeholder { line-height: 22px !important; }
        .programa-anio-dropdown .Select-input { height: 22px !important; }
        .programa-anio-dropdown .Select-arrow-zone { padding-right: 5px; }
        .programa-anio-dropdown svg { width: 12px !important; height: 12px !important; }
        .programa-periodo-fijo {
          display: inline-flex; align-items: center; gap: 7px;
          width: fit-content; margin: 0 0 10px; padding: 6px 9px;
          border: 1px solid #d8dced; border-radius: 7px; background: #fff;
          color: #202a59; font-size: 11px; font-weight: 800;
        }
        .programa-periodo-fijo > span {
          color: #4f3c98; font-size: 9px; font-weight: 900;
          letter-spacing: .04em; text-transform: uppercase;
        }
        .programa-metricas-horizontales {
          display: grid;
          grid-template-columns: minmax(0, 1fr) minmax(0, .78fr) minmax(0, 1.5fr);
          gap: 10px; margin-top: 14px;
        }
        .programa-metrica-periodo {
          min-width: 0; padding: 12px 8px; border-radius: 9px;
          background: #f3f4fa; text-align: center;
        }
        .programa-metrica-periodo span {
          display: block; color: #626a83; font-size: 9px; font-weight: 800;
          text-transform: uppercase;
        }
        .programa-metrica-periodo strong {
          display: block; margin-top: 3px; color: #17245b;
          font-size: clamp(18px, 1.5vw, 25px); font-weight: 900;
          white-space: nowrap;
        }
        .programa-metrica-periodo.inversion strong {
          font-size: clamp(15px, 1vw, 19px); letter-spacing: -.025em;
        }
        .programa-metricas-dos {
          grid-template-columns: repeat(2, minmax(0, 1fr));
        }
        .programa-metricas-dos .programa-metrica-periodo strong {
          font-size: clamp(20px, 1.7vw, 28px);
        }
        .programa-metricas-dos .programa-metrica-periodo.inversion strong {
          font-size: clamp(16px, 1.15vw, 21px);
        }
        .programa-metricas-priorizadas {
          margin-top: 8px;
        }
        .programa-estudiantes-principal {
          display: flex; align-items: baseline; gap: 6px;
          margin: 0 0 8px;
        }
        .programa-estudiantes-principal strong {
          color: #17245b; font-size: clamp(25px, 2vw, 34px);
          font-weight: 900; line-height: 1;
        }
        .programa-estudiantes-principal span {
          color: #515b7a; font-size: 10px; font-weight: 900;
          text-transform: uppercase;
        }
        .programa-estudiantes-regimen-fila {
          display: grid;
          grid-template-columns: minmax(0, 1fr) repeat(2, minmax(72px, .48fr));
          align-items: stretch;
          gap: 8px;
          margin: 0 0 8px;
        }
        .programa-estudiantes-regimen-fila .programa-estudiantes-principal {
          margin: 0;
          align-self: center;
        }
        .programa-regimen-mini {
          min-width: 0;
          padding: 7px 6px;
          border: 1px solid #f8f4e8;
          border-radius: 9px;
          background: #f8f4e8;
          text-align: center;
        }
        .programa-regimen-mini span {
          display: block;
          color: #69738c;
          font-size: 8px;
          font-weight: 800;
          line-height: 1.1;
          text-transform: uppercase;
        }
        .programa-regimen-mini strong {
          display: block;
          margin-top: 4px;
          color: #17245b;
          font-size: clamp(15px, 1.15vw, 19px);
          font-weight: 900;
          line-height: 1;
          white-space: nowrap;
        }
        .programa-desglose-inferior {
          display: grid; grid-template-columns: minmax(95px, .72fr) minmax(0, 1.45fr);
          gap: 10px;
        }
        .programa-desglose-inferior .programa-metrica-periodo {
          padding: 9px 8px;
        }
        .programa-desglose-inferior .programa-metrica-periodo strong {
          font-size: clamp(19px, 1.45vw, 25px);
        }
        .programa-desglose-inferior .programa-metrica-periodo.inversion strong {
          font-size: clamp(17px, 1.2vw, 22px);
        }
        .reinsercion-desglose {
          display: grid;
          grid-template-columns: repeat(2, minmax(0, 1fr));
          gap: 7px;
          margin-top: 8px;
        }
        .reinsercion-desglose .programa-metrica-periodo {
          padding: 8px 7px;
        }
        .reinsercion-desglose .programa-metrica-periodo strong {
          font-size: clamp(16px, 1.2vw, 20px);
        }
        .hospital-desglose {
          display: grid;
          grid-template-columns: minmax(70px, .65fr) minmax(75px, .65fr) minmax(0, 1.45fr);
          gap: 7px;
        }
        .hospital-desglose .programa-metrica-periodo {
          padding: 8px 6px;
        }
        .hospital-desglose .programa-metrica-periodo strong {
          font-size: clamp(16px, 1.15vw, 20px);
        }
        .hospital-desglose .programa-metrica-periodo.inversion strong {
          font-size: clamp(14px, .95vw, 17px);
        }
        .programa-periodo-nota {
          margin: 7px 0 0 !important; color: #59627d !important;
          font-size: 9px !important; line-height: 1.35 !important;
          font-weight: 400 !important;
          text-align: justify;
        }
        .educacion-media-periodo-pie {
          display: flex; flex-direction: column; gap: 2px;
          margin-top: 8px; color: #59627d;
          /* Igual que las etiquetas secundarias de las demás tarjetas. */
          font-size: 12px; line-height: 1.4;
          font-weight: 400;
        }
        .nee-desglose {
          display: grid;
          grid-template-columns: repeat(2, minmax(0, 1fr));
          gap: 7px;
          margin-top: 9px;
        }
        .nee-grupo-titulo {
          grid-column: 1 / -1;
          margin: 3px 0 -1px;
          color: #4f5876;
          font-size: 11px;
          font-weight: 700;
          text-transform: uppercase;
          letter-spacing: .02em;
        }
        .nee-desglose .vision-mini-stat {
          min-width: 0;
          padding: 8px 10px !important;
        }
        .nee-desglose .vision-mini-stat span {
          font-size: 12px !important;
          line-height: 1.2 !important;
          font-weight: 400 !important;
        }
        .nee-desglose .vision-mini-stat strong {
          margin-top: 3px !important;
          font-size: 17px !important;
          line-height: 1.05 !important;
        }
        /* Los recursos educativos (Alimentación, Uniformes, Textos, etc.)
           manejan cifras de estudiantes mucho más grandes que SAFPI/NAP
           (hasta 7 dígitos), así que reducimos un poco más estos tamaños
           para que "Estudiantes" y los números de Costa/Sierra no se
           corten dentro del cuadro. */
        .recurso-educativo-actualizado .programa-estudiantes-principal {
          flex-wrap: wrap;
          row-gap: 0;
        }
        .recurso-educativo-actualizado .programa-estudiantes-principal strong,
        .recurso-educativo-actualizado .programa-estudiantes-principal span {
          min-width: 0;
        }
        .recurso-educativo-actualizado .programa-estudiantes-principal strong {
          font-size: clamp(18px, 1.5vw, 26px);
        }
        .recurso-educativo-actualizado .programa-estudiantes-principal span {
          font-size: 7px;
          white-space: normal;
          overflow-wrap: break-word;
        }
        .recurso-educativo-actualizado .programa-regimen-mini {
          min-width: 0;
          padding: 6px 4px;
        }
        .recurso-educativo-actualizado .programa-regimen-mini strong {
          font-size: clamp(10px, .85vw, 13px);
          white-space: normal;
          overflow-wrap: break-word;
          word-break: break-word;
        }
        .recurso-resumen-nacional {
          display: grid;
          grid-template-columns: minmax(90px, .72fr) minmax(0, 1.35fr);
          gap: 7px;
          margin-top: 8px;
        }
        .recurso-resumen-nacional .programa-metrica-periodo {
          padding: 8px 7px;
        }
        .recurso-resumen-nacional .programa-metrica-periodo strong {
          font-size: clamp(16px, 1.2vw, 20px);
        }
        .recurso-resumen-nacional .programa-metrica-periodo.inversion strong {
          font-size: clamp(14px, .95vw, 17px);
        }
        .recurso-regimenes {
          display: grid;
          grid-template-columns: repeat(2, minmax(0, 1fr));
          gap: 7px;
          margin-top: 7px;
        }
        .recurso-regimen-panel {
          min-width: 0;
          padding: 8px 9px;
          border-radius: 9px;
          background: #f3f4fa;
        }
        .recurso-regimen-panel.secondary { background: #f8f4e9; }
        .recurso-regimen-panel h4 {
          margin: 0 0 5px;
          color: #4f3c98;
          font-size: 11px;
          font-weight: 900;
          text-transform: uppercase;
        }
        .recurso-regimen-fila {
          display: flex;
          justify-content: space-between;
          gap: 5px;
          margin-top: 3px;
          color: #626a83;
          font-size: 10px;
        }
        .recurso-regimen-fila strong {
          color: #17245b;
          white-space: nowrap;
        }
        .recurso-regimen-fila.inversion {
          display: block;
          margin-top: 5px;
        }
        .recurso-regimen-fila.inversion strong {
          display: block;
          margin-top: 2px;
          font-size: 11px;
          overflow-wrap: anywhere;
          white-space: normal;
        }
        .instituciones-anio-control,
        .estudiantes-anio-control,
        .docentes-anio-control {
          margin-bottom: 8px !important;
          display: grid !important;
          grid-template-columns: 76px 150px;
          align-items: center;
          gap: 5px;
        }
        .instituciones-anio-control > span,
        .estudiantes-anio-control > span,
        .docentes-anio-control > span {
          white-space: nowrap;
        }
        .instituciones-anio-control .programa-anio-dropdown,
        .estudiantes-anio-control .programa-anio-dropdown,
        .docentes-anio-control .programa-anio-dropdown {
          width: 150px !important;
          min-width: 150px !important;
          border-color: #8f80ca !important;
          box-shadow: none !important;
          outline: none !important;
        }
        .instituciones-anio-control .programa-anio-dropdown:focus,
        .instituciones-anio-control .programa-anio-dropdown:focus-within,
        .estudiantes-anio-control .programa-anio-dropdown:focus,
        .estudiantes-anio-control .programa-anio-dropdown:focus-within,
        .docentes-anio-control .programa-anio-dropdown:focus,
        .docentes-anio-control .programa-anio-dropdown:focus-within {
          border-color: #8f80ca !important;
          box-shadow: none !important;
          outline: none !important;
        }
        .instituciones-sostenimientos {
          display: grid !important;
          grid-template-columns: repeat(2, minmax(0, 1fr)) !important;
          gap: 7px !important;
          margin-top: 9px;
        }
        .instituciones-sostenimientos .vision-mini-stat {
          min-width: 0;
          padding: 9px 11px !important;
        }
        .instituciones-sostenimientos .vision-mini-stat span {
          overflow-wrap: anywhere;
          /* Igual que las etiquetas "Educación Especial/Ordinaria". */
          font-size: 12px !important;
          line-height: 1.2 !important;
          font-weight: 400 !important;
        }
        .instituciones-sostenimientos .vision-mini-stat strong {
          margin-top: 4px !important;
          font-size: 18px !important;
          line-height: 1.05 !important;
        }
        .desglose-demografico {
          display: grid !important;
          grid-template-columns: repeat(2, minmax(0, 1fr)) !important;
          gap: 7px !important;
          margin-top: 9px;
        }
        .desglose-demografico .vision-mini-stat {
          min-width: 0;
          padding: 8px 10px !important;
        }
        .desglose-demografico .vision-mini-stat span {
          overflow-wrap: anywhere;
          /* Misma jerarquía tipográfica en sexo, área y régimen. */
          font-size: 12px !important;
          line-height: 1.2 !important;
          font-weight: 400 !important;
        }
        .desglose-demografico .vision-mini-stat strong {
          margin-top: 3px !important;
          font-size: 17px !important;
          line-height: 1.05 !important;
          white-space: nowrap;
        }
        /* Ajustes incluidos en app.py para que funcionen aunque el navegador
           conserve una versión anterior de assets/portada.css. */
        .datos-generales-grid.educacion-media-cifras-grid,
        .datos-generales-grid.educacion-media-segunda-fila {
          width: 100% !important; max-width: none !important;
        }
        .datos-generales-grid.educacion-media-cifras-grid {
          grid-template-columns: repeat(3, minmax(0, 1fr)) !important;
        }
        .datos-generales-grid.educacion-media-segunda-fila {
          grid-template-columns: repeat(6, minmax(0, 1fr)) !important;
        }
        .educacion-media-cifras-grid > .vision-exec-card {
          width: auto !important; min-width: 0 !important;
          max-width: none !important; align-self: stretch !important;
        }
        .educacion-media-segunda-fila > .vision-exec-card {
          grid-column: span 2; width: auto !important; min-width: 0 !important;
          max-width: none !important; align-self: stretch !important;
        }
        .educacion-media-segunda-fila > .vision-exec-card:first-child {
          grid-column: 2 / span 2;
        }
        .educacion-media-segunda-fila > .vision-exec-card:nth-child(2) {
          grid-column: 4 / span 2;
        }
        .recursos-gestion-grid,
        .recursos-gestion-grid.vision-exec-grid {
          width: 100% !important; max-width: none !important;
          grid-template-columns: repeat(6, minmax(0, 1fr)) !important;
        }
        .recursos-gestion-grid > .vision-exec-card {
          grid-column: span 2; width: auto !important; min-width: 0 !important;
          max-width: none !important; align-self: stretch !important;
        }
        .recursos-gestion-grid > .vision-exec-card:nth-child(4) {
          grid-column: 2 / span 2;
        }
        .recursos-gestion-grid > .vision-exec-card:nth-child(5) {
          grid-column: 4 / span 2;
        }
        .vision-card-periodo {
          display: block; margin-top: 5px; color: #65708c;
          font-size: 10px; font-weight: 400;
        }
        .programas-complementarios-grid {
          grid-template-columns: repeat(6, minmax(0, 1fr)) !important;
        }
        .programas-complementarios-grid > .vision-exec-card {
          grid-column: span 2; width: auto !important; min-width: 0 !important;
          max-width: none !important; align-self: stretch !important;
        }
        .programas-complementarios-grid > .vision-exec-card:first-child {
          grid-column: 2 / span 2;
        }
        .programas-complementarios-grid > .vision-exec-card:nth-child(2) {
          grid-column: 4 / span 2;
        }
        .programas-complementarios-grid > .vision-exec-card:only-child {
          grid-column: 3 / span 2;
        }
        .deporte-informacion-grid,
        .deporte-informacion-grid.vision-exec-grid {
          width: 100% !important; max-width: none !important;
          grid-template-columns: repeat(5, minmax(0, 1fr)) !important;
          gap: 12px !important;
        }
        .deporte-informacion-grid > .vision-exec-card {
          width: auto !important; min-width: 0 !important;
          max-width: none !important; min-height: 142px;
          align-self: stretch !important;
        }
        .deporte-informacion-grid .vision-card-label {
          font-size: 12px; line-height: 1.18;
        }
        .deporte-informacion-grid .vision-card-total {
          font-size: clamp(25px, 2vw, 34px);
        }
        .deporte-subseccion-titulo {
          margin: 18px 0 10px; color: #232d5a;
          font-size: 15px; font-weight: 900;
        }
        .deporte-programas-generales-grid,
        .deporte-programas-generales-grid.vision-exec-grid {
          width: 100% !important; max-width: none !important;
          grid-template-columns: repeat(6, minmax(0, 1fr)) !important;
          gap: 12px !important;
        }
        .deporte-programas-generales-grid > .vision-exec-card {
          grid-column: span 2; width: auto !important; min-width: 0 !important;
          max-width: none !important; align-self: stretch !important;
        }
        .deporte-programas-generales-grid > .vision-exec-card:first-child {
          grid-column: 2 / span 2;
        }
        .deporte-programas-generales-grid > .vision-exec-card:nth-child(2) {
          grid-column: 4 / span 2;
        }
        .deporte-programa-metricas {
          display: grid; grid-template-columns: repeat(2, minmax(0, 1fr));
          gap: 8px; margin-top: 10px;
        }
        .deporte-programa-metrica {
          min-width: 0; padding: 10px 8px; border-radius: 9px;
          background: #f3f4fa; text-align: center;
        }
        .deporte-programa-metrica span {
          display: block; min-height: 22px; color: #626a83;
          font-size: 9px; font-weight: 800; line-height: 1.18;
          text-transform: uppercase;
        }
        .deporte-programa-metrica strong {
          display: block; margin-top: 4px; color: #17245b;
          font-size: clamp(20px, 1.6vw, 28px); font-weight: 900;
        }
        .deporte-informacion-fuente {
          margin: 10px 0 0 !important;
        }
        .deporte-informacion-nota {
          margin: 4px 0 16px !important;
        }
        .cultura-subseccion-titulo {
          margin: 18px 0 10px; color: #232d5a;
          font-size: 15px; font-weight: 900;
        }
        .cultura-artistas-grid,
        .cultura-artistas-grid.vision-exec-grid {
          width: 100% !important; max-width: none !important;
          grid-template-columns: repeat(3, minmax(0, 1fr)) !important;
          gap: 10px !important;
        }
        .cultura-organizaciones-grid,
        .cultura-organizaciones-grid.vision-exec-grid {
          width: 100% !important; max-width: none !important;
          grid-template-columns: repeat(6, minmax(0, 1fr)) !important;
          gap: 12px !important;
        }
        .cultura-artistas-grid > .vision-exec-card {
          width: auto !important; min-width: 0 !important;
          max-width: none !important; min-height: 142px;
          align-self: stretch !important;
        }
        .cultura-organizaciones-grid > .vision-exec-card {
          grid-column: span 2;
          width: auto !important; min-width: 0 !important;
          max-width: none !important; min-height: 142px;
          align-self: stretch !important;
        }
        .cultura-organizaciones-grid > .vision-exec-card:first-child {
          grid-column: 2 / span 2;
        }
        .cultura-organizaciones-grid > .vision-exec-card:nth-child(2) {
          grid-column: 4 / span 2;
        }
        .cultura-artistas-grid .vision-card-label,
        .cultura-organizaciones-grid .vision-card-label {
          font-size: 11.5px; line-height: 1.18;
        }
        .cultura-artistas-grid .vision-card-total,
        .cultura-organizaciones-grid .vision-card-total {
          font-size: clamp(24px, 1.8vw, 32px);
        }
        .cultura-artistas-grid > .vision-exec-card:last-child,
        .cultura-organizaciones-grid > .vision-exec-card:last-child {
          background: #f7f5fd;
          border-top-color: #503a98;
        }
        .cultura-informacion-fuente {
          margin: 10px 0 0 !important;
        }
        .cultura-informacion-nota {
          margin: 4px 0 16px !important;
        }
        /* Navegación adaptable: en escritorio el menú conserva su posición;
           en tablet y celular funciona como un panel lateral desplegable. */
        .mobile-menu-button,
        .mobile-menu-overlay {
          display: none;
        }
        @media (max-width: 1400px) {
          body.mobile-menu-open {
            overflow: hidden;
          }
          .header-bar {
            position: sticky; top: 0; z-index: 1000;
            min-height: 60px; flex-wrap: nowrap !important;
          }
          .module-layout {
            display: block !important;
            grid-template-columns: minmax(0, 1fr) !important;
            min-height: calc(100dvh - 60px);
          }
          .mobile-menu-button {
            position: fixed; z-index: 1040; top: 8px; left: 10px;
            height: 44px; padding: 0 13px;
            display: inline-flex; align-items: center; justify-content: center;
            gap: 7px; border: 1px solid rgba(255,255,255,.42);
            border-radius: 10px; color: #fff; background: #4F449A;
            box-shadow: 0 5px 14px rgba(24,37,87,.22);
            font-size: 13px; font-weight: 900; cursor: pointer;
          }
          .mobile-menu-icon {
            font-size: 22px; line-height: 1;
          }
          .module-sidebar {
            position: fixed !important; z-index: 1030;
            top: 60px !important; bottom: 0; left: 0;
            width: min(86vw, 330px) !important;
            height: calc(100dvh - 60px) !important;
            max-height: none !important; overflow-y: auto !important;
            transform: translateX(-105%);
            transition: transform .22s ease;
            box-shadow: 8px 0 28px rgba(24,37,87,.28);
          }
          body.mobile-menu-open .module-sidebar {
            transform: translateX(0);
          }
          .module-sidebar .sidebar-logo,
          .module-sidebar .sidebar-kicker {
            display: none;
          }
          .module-sidebar .side-back {
            margin-top: 4px;
          }
          .mobile-menu-overlay {
            position: fixed; z-index: 1020;
            inset: 60px 0 0 0; width: 100%; height: auto;
            display: block; padding: 0; border: 0;
            background: rgba(20,27,72,.46);
            opacity: 0; visibility: hidden; pointer-events: none;
            transition: opacity .2s ease, visibility .2s ease;
          }
          body.mobile-menu-open .mobile-menu-overlay {
            opacity: 1; visibility: visible; pointer-events: auto;
          }
          .module-main {
            width: 100% !important; min-width: 0 !important;
            min-height: calc(100dvh - 60px) !important;
            padding: 18px 14px 34px !important;
          }
          .vision-vice-menu-compacta {
            max-width: 100%; overflow-x: auto;
            overscroll-behavior-inline: contain;
          }
          .budget-general-chart-grid,
          .budget-kpi-grid {
            grid-template-columns: minmax(0, 1fr) !important;
          }
          .vision-landing-intro { margin-top: 30px; }
          .vision-landing-menu { margin-top: 36px !important; }
          .vision-landing-menu .vision-vice-menu {
            grid-template-columns: repeat(2, minmax(0, 1fr)) !important;
          }
          .vision-landing-menu .vision-vice-row,
          .vision-landing-menu .vision-vice-row:nth-child(4) {
            grid-column: auto !important;
          }
          .vision-landing-menu .vision-vice-row:nth-child(5) {
            grid-column: 1 / -1 !important; width: calc(50% - 9px);
            justify-self: center;
          }
        }
        @media (max-width: 600px) {
          .mobile-menu-button {
            width: 44px; padding: 0;
          }
          .mobile-menu-label {
            display: none;
          }
          .header-bar {
            min-height: 60px; padding: 6px 9px 6px 62px !important;
          }
          .header-title {
            min-width: 0; margin-left: 0 !important;
          }
          .header-title strong {
            overflow: hidden; font-size: 12px !important;
            text-overflow: ellipsis; white-space: nowrap;
          }
          .header-title span {
            font-size: 8px !important;
          }
          .cutoff-date {
            display: none !important;
          }
          .header-actions {
            gap: 5px; padding-left: 4px !important;
          }
          .home-button,
          .print-button {
            width: 40px !important; height: 40px !important;
          }
          .module-main {
            padding: 12px 8px 28px !important;
            overflow-x: hidden !important;
          }
          .vision-landing-intro { margin-top: 18px; }
          .vision-landing-title { font-size: 25px !important; }
          .vision-landing-subtitle { font-size: 13px; }
          .vision-landing-menu { margin-top: 26px !important; }
          .vision-landing-menu .vision-vice-menu {
            grid-template-columns: minmax(0, 1fr) !important;
            gap: 12px !important;
          }
          .vision-landing-menu .vision-vice-row,
          .vision-landing-menu .vision-vice-row:nth-child(4),
          .vision-landing-menu .vision-vice-row:nth-child(5) {
            grid-column: auto !important; width: 100% !important;
          }
          .vision-landing-menu .vision-vice-pill {
            min-height: 104px !important; padding: 15px !important;
          }
          .content-heading,
          .indicator-detail-card,
          .chart-card,
          .budget-table-card {
            padding: 12px !important;
          }
          .content-heading h1,
          .vision-exec-header h1 {
            overflow-wrap: anywhere;
          }
          .status-row {
            display: flex !important; overflow-x: auto;
            gap: 7px; padding-bottom: 5px;
          }
          .status-year-item {
            flex: 0 0 92px;
          }

          /* En teléfonos, las reglas de escritorio de estas cuadrículas
             tienen mayor especificidad. Se las sobrescribe de forma
             explícita para que ninguna tarjeta conserve columnas estrechas. */
          .module-main .vision-exec-page .datos-generales-grid.educacion-media-cifras-grid,
          .module-main .vision-exec-page .datos-generales-grid.educacion-media-segunda-fila,
          .module-main .vision-exec-page .recursos-gestion-grid.vision-exec-grid,
          .module-main .vision-exec-page .programas-complementarios-grid,
          .module-main .vision-exec-page .deporte-informacion-grid.vision-exec-grid,
          .module-main .vision-exec-page .deporte-programas-generales-grid.vision-exec-grid,
          .module-main .vision-exec-page .cultura-artistas-grid.vision-exec-grid,
          .module-main .vision-exec-page .cultura-organizaciones-grid.vision-exec-grid,
          .module-main .vision-exec-page .programas-servicios-grid.vision-exec-grid {
            display: grid !important;
            width: 100% !important;
            max-width: 100% !important;
            grid-template-columns: minmax(0, 1fr) !important;
            gap: 12px !important;
            margin-left: 0 !important;
            margin-right: 0 !important;
          }
          .module-main .vision-exec-page .vision-exec-grid > .vision-exec-card,
          .module-main .vision-exec-page .educacion-media-segunda-fila > .vision-exec-card,
          .module-main .vision-exec-page .recursos-gestion-grid > .vision-exec-card,
          .module-main .vision-exec-page .programas-complementarios-grid > .vision-exec-card,
          .module-main .vision-exec-page .deporte-programas-generales-grid > .vision-exec-card {
            grid-column: 1 / -1 !important;
            width: 100% !important;
            min-width: 0 !important;
            max-width: 100% !important;
            min-height: 0 !important;
            box-sizing: border-box !important;
          }
          .module-main .vision-exec-page .vision-exec-card {
            padding: 16px !important;
            overflow: hidden !important;
          }
          .module-main .vision-exec-page .vision-card-label,
          .module-main .vision-exec-page .vision-card-total,
          .module-main .vision-exec-page .vision-card-unit,
          .module-main .vision-exec-page .vision-exec-card > p {
            max-width: 100% !important;
            overflow-wrap: anywhere !important;
            word-break: normal !important;
            white-space: normal !important;
          }
          .module-main .vision-exec-page .vision-card-label {
            font-size: 14px !important;
            line-height: 1.25 !important;
          }
          .module-main .vision-exec-page .vision-card-total {
            font-size: clamp(28px, 9vw, 38px) !important;
            line-height: 1.08 !important;
          }
          .module-main .vision-exec-page .vision-card-details {
            width: 100% !important;
            grid-template-columns: repeat(2, minmax(0, 1fr)) !important;
            gap: 8px !important;
          }
          .module-main .vision-exec-page .vision-mini-stat {
            min-width: 0 !important;
            padding: 10px 8px !important;
          }
          .module-main .vision-exec-page .vision-mini-stat span,
          .module-main .vision-exec-page .vision-mini-stat strong {
            max-width: 100% !important;
            overflow-wrap: anywhere !important;
            white-space: normal !important;
          }
          .module-main .vision-exec-page .vision-mini-stat strong {
            font-size: clamp(20px, 7vw, 29px) !important;
          }
        }
        @media (max-width: 1300px) {
          .deporte-informacion-grid,
          .deporte-informacion-grid.vision-exec-grid {
            grid-template-columns: repeat(3, minmax(0, 1fr)) !important;
          }
          .cultura-artistas-grid,
          .cultura-artistas-grid.vision-exec-grid {
            grid-template-columns: repeat(3, minmax(0, 1fr)) !important;
          }
          .cultura-organizaciones-grid,
          .cultura-organizaciones-grid.vision-exec-grid {
            grid-template-columns: repeat(2, minmax(0, 1fr)) !important;
          }
        }
        @media (min-width: 1025px) and (max-width: 1450px) {
          /* Portátiles de 13 y 14 pulgadas: el menú lateral reduce el ancho
             real del contenido. Las cuadrículas densas pasan a tres columnas
             y se compacta únicamente la tipografía secundaria. */
          .module-main .vision-exec-page .programas-servicios-grid.vision-exec-grid,
          .module-main .vision-exec-page .deporte-informacion-grid.vision-exec-grid,
          .module-main .vision-exec-page .cultura-artistas-grid.vision-exec-grid {
            grid-template-columns: repeat(3, minmax(0, 1fr)) !important;
            gap: 10px !important;
          }
          .module-main .vision-exec-page .cultura-organizaciones-grid.vision-exec-grid {
            grid-template-columns: repeat(2, minmax(0, 1fr)) !important;
          }
          .module-main .vision-exec-page .datos-generales-grid.educacion-media-segunda-fila,
          .module-main .vision-exec-page .programas-complementarios-grid {
            grid-template-columns: repeat(2, minmax(0, 1fr)) !important;
            gap: 12px !important;
          }
          .module-main .vision-exec-page .educacion-media-segunda-fila > .vision-exec-card,
          .module-main .vision-exec-page .programas-complementarios-grid > .vision-exec-card {
            grid-column: auto !important;
            width: 100% !important;
            min-width: 0 !important;
            max-width: 100% !important;
          }
          .module-main .vision-exec-page .vision-exec-card {
            padding-left: 11px !important;
            padding-right: 11px !important;
          }
          .module-main .vision-exec-page .vision-card-label,
          .module-main .vision-exec-page .programa-resumen-card .vision-card-label {
            min-height: 0 !important;
            font-size: 12px !important;
            line-height: 1.22 !important;
            overflow-wrap: anywhere !important;
          }
          .module-main .vision-exec-page .vision-card-total,
          .module-main .vision-exec-page .programa-card-resultado strong {
            font-size: clamp(23px, 2vw, 29px) !important;
          }
          .module-main .vision-exec-page .vision-mini-stat {
            min-width: 0 !important;
            padding: 7px 6px !important;
          }
          .module-main .vision-exec-page .vision-mini-stat span {
            font-size: 10px !important;
            overflow-wrap: anywhere !important;
          }
          .module-main .vision-exec-page .educacion-media-periodo-pie {
            font-size: 10px;
          }
          .module-main .vision-exec-page .vision-mini-stat strong {
            font-size: 15px !important;
            overflow-wrap: anywhere !important;
          }
          .module-main .vision-exec-page .programa-metrica-periodo {
            padding-left: 6px !important;
            padding-right: 6px !important;
          }
          .module-main .vision-exec-page .programa-metrica-periodo strong {
            font-size: clamp(17px, 1.45vw, 22px) !important;
          }
          .budget-kpi-grid {
            grid-template-columns: repeat(3, minmax(0, 1fr)) !important;
          }
          .vision-vice-menu-compacta {
            max-width: 100%; overflow-x: auto;
            overscroll-behavior-inline: contain;
          }
        }
        @media (max-width: 900px) {
          .educacion-media-cifras-grid,
          .educacion-media-segunda-fila,
          .programas-gestion-consolidados .programas-servicios-grid {
            grid-template-columns: minmax(0, 1fr) !important;
          }
          .educacion-media-segunda-fila > .vision-exec-card,
          .educacion-media-segunda-fila > .vision-exec-card:first-child,
          .educacion-media-segunda-fila > .vision-exec-card:nth-child(2) {
            grid-column: 1 !important;
          }
          .recursos-gestion-grid,
          .recursos-gestion-grid.vision-exec-grid {
            grid-template-columns: minmax(0, 1fr) !important;
          }
          .recursos-gestion-grid > .vision-exec-card,
          .recursos-gestion-grid > .vision-exec-card:nth-child(4),
          .recursos-gestion-grid > .vision-exec-card:nth-child(5) {
            grid-column: 1 !important;
          }
          .programas-complementarios-grid,
          .programas-complementarios-grid > .vision-exec-card,
          .programas-complementarios-grid > .vision-exec-card:first-child,
          .programas-complementarios-grid > .vision-exec-card:nth-child(2) {
            grid-template-columns: minmax(0, 1fr) !important;
            grid-column: 1 !important;
          }
          .deporte-informacion-grid,
          .deporte-informacion-grid.vision-exec-grid {
            grid-template-columns: minmax(0, 1fr) !important;
          }
          .deporte-programas-generales-grid,
          .deporte-programas-generales-grid.vision-exec-grid,
          .deporte-programas-generales-grid > .vision-exec-card,
          .deporte-programas-generales-grid > .vision-exec-card:first-child,
          .deporte-programas-generales-grid > .vision-exec-card:nth-child(2) {
            grid-template-columns: minmax(0, 1fr) !important;
            grid-column: 1 !important;
          }
          .cultura-artistas-grid,
          .cultura-artistas-grid.vision-exec-grid,
          .cultura-organizaciones-grid,
          .cultura-organizaciones-grid.vision-exec-grid {
            grid-template-columns: minmax(0, 1fr) !important;
          }
        }
        @media (min-width: 601px) and (max-width: 1400px) {
          /* En tablet se aprovecha el espacio con dos tarjetas por fila, sin
             conservar las seis columnas auxiliares usadas en escritorio. */
          .module-main .vision-exec-page .datos-generales-grid.educacion-media-cifras-grid,
          .module-main .vision-exec-page .datos-generales-grid.educacion-media-segunda-fila,
          .module-main .vision-exec-page .recursos-gestion-grid.vision-exec-grid,
          .module-main .vision-exec-page .programas-complementarios-grid,
          .module-main .vision-exec-page .deporte-informacion-grid.vision-exec-grid,
          .module-main .vision-exec-page .deporte-programas-generales-grid.vision-exec-grid,
          .module-main .vision-exec-page .cultura-artistas-grid.vision-exec-grid,
          .module-main .vision-exec-page .cultura-organizaciones-grid.vision-exec-grid {
            width: 100% !important;
            max-width: 100% !important;
            grid-template-columns: repeat(2, minmax(0, 1fr)) !important;
            gap: 12px !important;
          }
          .module-main .vision-exec-page .vision-exec-grid > .vision-exec-card {
            grid-column: auto !important;
            width: 100% !important;
            min-width: 0 !important;
            max-width: 100% !important;
            box-sizing: border-box !important;
          }
        }

        /* Esta regla debe ir después de todos los ajustes responsive. Algunas
           reglas móviles establecen overflow:hidden con mayor especificidad y
           recortaban la ayuda de las tarjetas de recursos educativos. */
        .module-main .vision-exec-page .vision-exec-grid > .vision-exec-card.programa-con-ayuda {
          position: relative !important;
          overflow: visible !important;
          cursor: help;
        }
        .module-main .vision-exec-page .vision-exec-grid > .vision-exec-card.programa-con-ayuda:hover,
        .module-main .vision-exec-page .vision-exec-grid > .vision-exec-card.programa-con-ayuda:focus,
        .module-main .vision-exec-page .vision-exec-grid > .vision-exec-card.programa-con-ayuda:focus-within {
          z-index: 1000 !important;
        }
        .module-main .vision-exec-page .recursos-gestion-grid.vision-exec-grid {
          overflow: visible !important;
        }
        .module-main .vision-exec-page .programa-info-tooltip {
          z-index: 1001 !important;
        }
        </style>
        <style id="minedec-institutional-theme">
        /* Tema institucional unificado según el manual técnico de Presidencia.
           La navegación usa el morado oficial #4F449A y el amarillo queda
           reservado para selección y acentos. */
        :root {
          --minedec-navy: #232D5A;
          --minedec-blue: #446381;
          --minedec-teal: #80BBBF;
          --minedec-purple: #4F449A;
          --minedec-yellow: #F1B620;
          --minedec-surface: #FFFFFF;
          --minedec-canvas: #F4F7F9;
          --minedec-line: #D9E2E8;
        }
        .header-bar,
        .module-sidebar {
          background: #4F449A !important;
        }
        .indicadores-cabecera {
          background: #232D5A !important;
        }
        .vision-exec-header h1,
        .vision-landing-title {
          color: #232D5A !important;
        }
        .portada-boton,
        .indicadores-opciones .portada-boton {
          background: linear-gradient(125deg, #232D5A 0%, #344A71 100%) !important;
          box-shadow: 0 9px 0 rgba(79, 68, 154, .18),
                      0 15px 26px rgba(35, 45, 90, .16) !important;
        }
        .portada-boton:hover,
        .portada-boton:focus-visible {
          outline-color: rgba(128, 187, 191, .58) !important;
        }
        .side-section.active {
          color: #4F449A !important;
          background: #F1B620 !important;
          border-left: 4px solid #F1B620 !important;
        }
        .vision-title-mark {
          background: linear-gradient(#80BBBF, #446381) !important;
        }
        .vision-eje-banner {
          color: #FFFFFF !important;
          background: #232D5A !important;
          border-color: #232D5A !important;
          box-shadow: 0 7px 16px rgba(35, 45, 90, .13) !important;
        }
        .vision-eje-banner::before {
          background: #F1B620 !important;
        }
        .vision-vice-pill.active {
          color: #FFFFFF !important;
          background: #232D5A !important;
          border-color: #232D5A !important;
        }
        .vision-exec-card {
          --card-accent: #446381 !important;
          border-color: #D9E2E8 !important;
          border-top-color: #446381 !important;
          box-shadow: 0 8px 20px rgba(35, 45, 90, .08) !important;
        }
        .vision-exec-card:nth-child(3n + 2) {
          border-top-color: #80BBBF !important;
        }
        .vision-exec-card:nth-child(3n) {
          border-top-color: #232D5A !important;
        }
        .vision-card-symbol {
          color: #232D5A !important;
          background: #EDF4F6 !important;
          border-color: #C9DDE1 !important;
        }
        .metric-summary-row .metric-card,
        .metric-summary-row .detail-item.highlight {
          color: #232D5A !important;
          background: #FFFFFF !important;
          border: 1px solid #D9E2E8 !important;
          border-top: 4px solid #80BBBF !important;
          box-shadow: 0 6px 16px rgba(35, 45, 90, .07) !important;
        }
        .metric-summary-row > :nth-child(2) {
          border-top-color: #232D5A !important;
        }
        .metric-card .metric-year {
          color: #526079 !important;
          background: #EDF4F6 !important;
        }
        .budget-kpi-card {
          background: #FFFFFF !important;
          border: 1px solid #D9E2E8 !important;
          border-top: 3px solid #446381 !important;
          box-shadow: 0 5px 14px rgba(35, 45, 90, .07) !important;
        }
        .budget-kpi-card:nth-child(even) {
          border-top-color: #80BBBF !important;
        }
        .budget-kpi-icon {
          background: #EDF4F6 !important;
          border-color: #C9DDE1 !important;
        }
        .budget-investment-badge,
        .budget-updated-badge {
          color: #232D5A !important;
          background: #E8F1F3 !important;
        }

        /* Portadas internas: Inventario, Visión Ejecutiva y Documentación
           comparten la misma jerarquía visual de la portada de Indicadores. */
        .institutional-landing-header {
          width: 100% !important;
          max-width: none !important;
          margin: 4px 0 0 !important;
          padding: 26px 30px !important;
          box-sizing: border-box;
          border: 0 !important;
          border-radius: 20px !important;
          color: #FFFFFF !important;
          background: #232D5A !important;
          box-shadow: 0 12px 26px rgba(35, 45, 90, .14) !important;
          text-align: left !important;
        }
        .institutional-landing-header .content-kicker,
        .institutional-landing-header .documentation-kicker {
          margin: 0 0 7px !important;
          color: #F1B620 !important;
          font-size: 12px !important;
          font-weight: 900 !important;
          letter-spacing: .13em !important;
        }
        .institutional-landing-header h1,
        .institutional-landing-header .vision-landing-title {
          margin: 0 0 7px !important;
          color: #FFFFFF !important;
          font-size: clamp(28px, 2.5vw, 42px) !important;
          line-height: 1.1 !important;
          text-align: left !important;
        }
        .institutional-landing-header > p:last-child,
        .institutional-landing-header .vision-landing-subtitle {
          max-width: 900px !important;
          margin: 0 !important;
          color: #EDF2F7 !important;
          font-size: 14px !important;
          line-height: 1.5 !important;
          text-align: left !important;
        }
        .institutional-access-grid {
          display: grid !important;
          grid-template-columns: repeat(3, minmax(0, 1fr));
          gap: 24px !important;
          margin-top: 56px !important;
        }
        .institutional-access-card {
          min-width: 0;
          min-height: 150px;
          padding: 22px 24px;
          display: flex;
          align-items: center;
          justify-content: center;
          box-sizing: border-box;
          border: 1px solid #51678B;
          border-radius: 20px;
          color: #FFFFFF !important;
          background: #344A71 !important;
          box-shadow: 0 9px 0 rgba(79, 68, 154, .18),
                      0 15px 26px rgba(35, 45, 90, .12);
          font-size: 17px;
          font-weight: 800;
          line-height: 1.25;
          text-align: center;
          text-decoration: none !important;
          transition: transform .18s ease, box-shadow .18s ease;
        }
        .institutional-access-card:hover,
        .institutional-access-card:focus-visible {
          transform: translateY(-4px);
          outline: 3px solid rgba(128, 187, 191, .55);
          outline-offset: 2px;
        }
        .vision-landing-menu {
          margin-top: 42px !important;
        }
        .vision-landing-menu .vision-vice-pill {
          grid-template-rows: 1fr !important;
          min-height: 150px !important;
          border: 1px solid #51678B !important;
          border-top: 1px solid #51678B !important;
          color: #FFFFFF !important;
          background: #344A71 !important;
          box-shadow: 0 9px 0 rgba(79, 68, 154, .18),
                      0 15px 26px rgba(35, 45, 90, .12) !important;
        }
        .vision-landing-menu .vision-landing-vice-nombre {
          align-self: center !important;
          color: #FFFFFF !important;
          font-size: 16px !important;
        }
        .documentation-page {
          width: 100% !important;
        }
        .documentation-section {
          border-color: #D9E2E8 !important;
        }
        .document-deliverable-card {
          border-color: #D9E2E8 !important;
          border-top-color: #446381 !important;
          box-shadow: 0 7px 18px rgba(35, 45, 90, .07);
        }
        .document-deliverable-card h3 {
          color: #232D5A !important;
        }
        .document-file-icon {
          background: #EDF4F6 !important;
        }
        .document-action {
          background: #344A71 !important;
        }
        .document-action:hover,
        .document-action:focus-visible {
          background: #232D5A !important;
        }

        /* Portátiles de 13–14 pulgadas: se prioriza ancho útil para que las
           etiquetas no se monten. En pantallas grandes vuelve la composición
           de dos columnas prevista por el manual (mesa óptima 1920 × 1080). */
        @media (min-width: 1181px) and (max-width: 1650px) {
          .indicator-workspace {
            grid-template-columns: minmax(0, 1fr) !important;
            gap: 14px !important;
          }
          .indicator-detail-card .details-grid {
            grid-template-columns: repeat(auto-fit, minmax(190px, 1fr)) !important;
          }
          .indicator-detail-card .full,
          .indicator-detail-card .source {
            grid-column: auto !important;
          }
          .budget-kpi-grid {
            grid-template-columns: repeat(3, minmax(0, 1fr)) !important;
          }
          .datos-generales-grid,
          .recursos-gestion-grid,
          .cultura-artistas-grid {
            grid-template-columns: repeat(3, minmax(0, 1fr)) !important;
          }
          .cultura-organizaciones-grid {
            grid-template-columns: repeat(2, minmax(0, 1fr)) !important;
          }
        }
        /* En portátiles angostos, una serie mensual conserva el espacio que
           necesita para mostrar cifras horizontales. Solo ese gráfico puede
           desplazarse; las series cortas siguen ocupando exactamente el 100%. */
        @media (min-width: 901px) and (max-width: 1400px) {
          .chart-card {
            overflow-x: auto !important;
            overscroll-behavior-inline: contain;
          }
          .indicator-chart-long {
            min-width: 1080px !important;
          }
        }
        @media (min-width: 1181px) and (max-width: 1480px) {
          .budget-general-chart-grid {
            grid-template-columns: minmax(0, 1fr) !important;
          }
          .budget-panel {
            min-height: 300px;
          }
        }
        @media (max-width: 900px) {
          .institutional-access-grid {
            grid-template-columns: minmax(0, 1fr) !important;
            gap: 14px !important;
            margin-top: 26px !important;
          }
          .institutional-access-card {
            min-height: 96px;
          }
          .institutional-landing-header {
            padding: 21px 20px !important;
          }
          .chart-card {
            overflow-x: auto !important;
            overscroll-behavior-inline: contain;
            -webkit-overflow-scrolling: touch;
          }
          .indicator-chart {
            width: 100% !important;
            min-width: 0 !important;
            height: 410px !important;
          }
          .indicator-chart-short {
            min-width: 620px !important;
          }
          .indicator-chart-medium {
            min-width: 760px !important;
          }
          .indicator-chart-long {
            min-width: 1080px !important;
          }
        }
        /* En celular se muestra la serie completa, sin desplazamiento lateral.
           Plotly compacta las categorías y el ajuste JavaScript gira solamente
           las cifras superiores para que sigan siendo legibles. Esta regla va
           después de la de tabletas para reemplazar sus anchos mínimos. */
        @media (max-width: 600px) {
          .chart-card {
            overflow-x: hidden !important;
          }
          .indicator-chart,
          .indicator-chart-short,
          .indicator-chart-medium,
          .indicator-chart-long {
            width: 100% !important;
            min-width: 0 !important;
            max-width: 100% !important;
            height: 410px !important;
          }
          .status-row {
            display: grid !important;
            grid-template-columns: repeat(3, minmax(0, 1fr)) !important;
            overflow: visible !important;
            gap: 10px 6px !important;
          }
          .status-year-item {
            width: auto !important;
            min-width: 0 !important;
            flex: none !important;
          }
        }

        /* Ajuste final de cuadrículas. Se ubica al final para que ninguna
           regla anterior vuelva a estrechar o cortar las tarjetas. */
        .module-main .vision-exec-page .vision-exec-card {
          min-width: 0 !important;
          box-sizing: border-box !important;
        }
        .module-main .vision-exec-page .vision-card-label,
        .module-main .vision-exec-page .vision-card-unit,
        .module-main .vision-exec-page .vision-mini-stat span {
          max-width: 100% !important;
          overflow-wrap: break-word !important;
          word-break: normal !important;
          white-space: normal !important;
        }
        .module-main .vision-exec-page .vision-card-total,
        .module-main .vision-exec-page .vision-mini-stat strong {
          max-width: 100% !important;
          white-space: nowrap;
        }
        @media (min-width: 1401px) {
          .module-main .vision-exec-page .cultura-artistas-grid.vision-exec-grid {
            grid-template-columns: repeat(3, minmax(0, 1fr)) !important;
          }
          .module-main .vision-exec-page .cultura-organizaciones-grid.vision-exec-grid {
            grid-template-columns: repeat(6, minmax(0, 1fr)) !important;
          }
          .module-main .vision-exec-page .cultura-organizaciones-grid > .vision-exec-card:first-child {
            grid-column: 2 / span 2 !important;
          }
          .module-main .vision-exec-page .cultura-organizaciones-grid > .vision-exec-card:nth-child(2) {
            grid-column: 4 / span 2 !important;
          }
          .module-main .vision-exec-page .educacion-superior-cifras-grid.vision-exec-grid {
            grid-template-columns: repeat(3, minmax(0, 1fr)) !important;
          }
        }
        @media (min-width: 601px) and (max-width: 1400px) {
          .module-main .vision-exec-page .datos-generales-grid.educacion-media-cifras-grid,
          .module-main .vision-exec-page .datos-generales-grid.educacion-media-segunda-fila,
          .module-main .vision-exec-page .cultura-artistas-grid.vision-exec-grid,
          .module-main .vision-exec-page .cultura-organizaciones-grid.vision-exec-grid,
          .module-main .vision-exec-page .educacion-superior-cifras-grid.vision-exec-grid,
          .module-main .vision-exec-page .recursos-gestion-grid.vision-exec-grid,
          .module-main .vision-exec-page .programas-servicios-grid.vision-exec-grid,
          .module-main .vision-exec-page .programas-complementarios-grid,
          .module-main .vision-exec-page .deporte-informacion-grid.vision-exec-grid,
          .module-main .vision-exec-page .deporte-programas-generales-grid.vision-exec-grid {
            grid-template-columns: repeat(2, minmax(0, 1fr)) !important;
          }
          .module-main .vision-exec-page .educacion-media-segunda-fila > .vision-exec-card,
          .module-main .vision-exec-page .programas-complementarios-grid > .vision-exec-card,
          .module-main .vision-exec-page .deporte-programas-generales-grid > .vision-exec-card {
            grid-column: auto !important;
            width: 100% !important;
          }
          .module-main .vision-exec-page .cultura-organizaciones-grid > .vision-exec-card {
            grid-column: auto !important;
          }
        }
        @media (max-width: 600px) {
          .module-main .vision-exec-page .datos-generales-grid.educacion-media-cifras-grid,
          .module-main .vision-exec-page .datos-generales-grid.educacion-media-segunda-fila,
          .module-main .vision-exec-page .cultura-artistas-grid.vision-exec-grid,
          .module-main .vision-exec-page .cultura-organizaciones-grid.vision-exec-grid,
          .module-main .vision-exec-page .educacion-superior-cifras-grid.vision-exec-grid,
          .module-main .vision-exec-page .recursos-gestion-grid.vision-exec-grid,
          .module-main .vision-exec-page .programas-servicios-grid.vision-exec-grid,
          .module-main .vision-exec-page .programas-complementarios-grid,
          .module-main .vision-exec-page .deporte-informacion-grid.vision-exec-grid,
          .module-main .vision-exec-page .deporte-programas-generales-grid.vision-exec-grid {
            grid-template-columns: minmax(0, 1fr) !important;
          }
          .module-main .vision-exec-page .educacion-media-segunda-fila > .vision-exec-card,
          .module-main .vision-exec-page .programas-complementarios-grid > .vision-exec-card,
          .module-main .vision-exec-page .deporte-programas-generales-grid > .vision-exec-card,
          .module-main .vision-exec-page .cultura-organizaciones-grid > .vision-exec-card {
            grid-column: 1 !important;
          }
          .module-main .vision-exec-page .vision-card-total,
          .module-main .vision-exec-page .vision-mini-stat strong {
            white-space: normal !important;
            overflow-wrap: break-word !important;
          }
        }

        /* Cultura reutiliza el lenguaje visual de Gestión Educativa:
           tarjetas, selector de período, métricas internas y pie de corte. */
        .cultura-resumen-grid,
        .cultura-programas-grid {
          width: 100% !important;
          max-width: none !important;
          gap: 12px !important;
          margin: 12px 0 18px;
        }
        .cultura-resumen-grid {
          grid-template-columns: repeat(6, minmax(0, 1fr)) !important;
        }
        .cultura-resumen-grid > .cultura-resumen-card {
          grid-column: span 2;
          width: auto !important;
          min-width: 0 !important;
          max-width: none !important;
          min-height: 220px;
          align-self: stretch !important;
        }
        .cultura-resumen-grid > .cultura-resumen-card:nth-child(4) {
          grid-column: 3 / span 2;
        }
        .cultura-programas-grid {
          grid-template-columns: repeat(3, minmax(0, 1fr)) !important;
        }
        .cultura-programas-grid > .cultura-programa-card {
          width: auto !important;
          min-width: 0 !important;
          max-width: none !important;
          min-height: 255px;
          align-self: stretch !important;
        }
        .cultura-resumen-card .vision-card-title-row,
        .cultura-programa-card .vision-card-title-row {
          min-height: 42px;
          align-items: center !important;
        }
        .cultura-resumen-card .vision-card-label,
        .cultura-programa-card .vision-card-label {
          min-height: 0 !important;
          color: #17245b;
          font-size: 12px !important;
          line-height: 1.2 !important;
          font-weight: 900 !important;
          overflow-wrap: anywhere;
        }
        .cultura-resumen-desglose,
        .cultura-programa-desglose {
          display: grid;
          grid-template-columns: repeat(2, minmax(0, 1fr));
          gap: 8px;
          margin: 8px 0 10px;
        }
        .cultura-resumen-desglose.items-3,
        .cultura-programa-desglose {
          grid-template-columns: repeat(3, minmax(0, 1fr));
        }
        .cultura-resumen-item,
        .cultura-programa-metrica {
          display: flex;
          flex-direction: column;
          align-items: center;
          justify-content: center;
          min-width: 0;
          min-height: 76px;
          padding: 9px 6px;
        }
        .cultura-resumen-item:nth-child(even),
        .cultura-programa-metrica:nth-child(even) {
          background: #f8f4e8;
        }
        .cultura-resumen-item span,
        .cultura-programa-metrica span {
          min-height: 22px;
          font-size: 9px !important;
          line-height: 1.15;
          overflow-wrap: anywhere;
        }
        .cultura-resumen-item strong,
        .cultura-programa-metrica strong {
          font-size: clamp(18px, 1.45vw, 24px) !important;
        }
        .cultura-resumen-item.pendiente strong,
        .cultura-programa-metrica.pendiente strong {
          color: #9aa2b3;
        }
        @media (min-width: 701px) and (max-width: 1400px) {
          .cultura-resumen-grid,
          .cultura-programas-grid {
            grid-template-columns: repeat(2, minmax(0, 1fr)) !important;
          }
          .cultura-resumen-grid > .cultura-resumen-card,
          .cultura-resumen-grid > .cultura-resumen-card:nth-child(4) {
            grid-column: auto !important;
          }
          .cultura-programas-grid > .cultura-programa-card:last-child {
            grid-column: 1 / -1;
            width: calc(50% - 7px) !important;
            justify-self: center;
          }
        }
        @media (max-width: 700px) {
          .cultura-resumen-grid,
          .cultura-programas-grid {
            grid-template-columns: minmax(0, 1fr) !important;
            gap: 12px;
          }
          .cultura-resumen-grid > .cultura-resumen-card,
          .cultura-resumen-grid > .cultura-resumen-card:nth-child(4),
          .cultura-programas-grid > .cultura-programa-card,
          .cultura-programas-grid > .cultura-programa-card:last-child {
            grid-column: 1;
            width: 100% !important;
          }
          .cultura-resumen-desglose.items-3 {
            grid-template-columns: repeat(3, minmax(0, 1fr));
          }
          .cultura-programa-desglose {
            grid-template-columns: repeat(3, minmax(0, 1fr));
          }
          .cultura-resumen-item strong,
          .cultura-programa-metrica strong {
            font-size: clamp(18px, 6vw, 23px) !important;
          }
        }
        </style>
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
        <script>
        /* Las cifras permanecen horizontales en escritorio y tableta. Solo en
           celular se colocan verticales para mostrar todos los períodos a la
           vez, sin obligar al usuario a desplazar el gráfico. */
        (function () {
          let temporizador;
          function ajustarEtiquetas() {
            if (!window.Plotly) return;
            document.querySelectorAll('.indicator-chart .js-plotly-plot').forEach(function (grafico) {
              const categorias = (grafico.data || []).reduce(function (maximo, serie) {
                return Math.max(maximo, Array.isArray(serie.x) ? serie.x.length : 0);
              }, 0);
              const vistaMovil = window.innerWidth <= 600;
              const angulo = vistaMovil && categorias >= 5 ? -90 : 0;
              const tamano = vistaMovil ? 8 : (categorias >= 9 ? 9 : 10);
              const estado = angulo + ':' + tamano + ':' + categorias;
              if (grafico.dataset.etiquetasResponsivas === estado) return;
              grafico.dataset.etiquetasResponsivas = estado;
              window.Plotly.restyle(grafico, {
                textangle: angulo,
                'textfont.size': tamano
              });
              if (vistaMovil) {
                window.Plotly.relayout(grafico, {
                  'margin.l': 44,
                  'margin.r': 6,
                  'margin.t': 72,
                  'margin.b': 62,
                  'xaxis.tickfont.size': 8,
                  'xaxis.title.font.size': 11,
                  'yaxis.title.font.size': 11,
                  'legend.font.size': 9
                });
              }
            });
          }
          function programarAjuste() {
            clearTimeout(temporizador);
            temporizador = setTimeout(ajustarEtiquetas, 90);
          }
          window.addEventListener('resize', programarAjuste);
          window.addEventListener('load', function () {
            programarAjuste();
            new MutationObserver(programarAjuste).observe(document.body, {
              childList: true,
              subtree: true
            });
          });
        })();
        </script>
        <script>
        /* Menú móvil por delegación de eventos: funciona también cuando Dash
           reconstruye por completo el módulo después de cada selección. */
        (function () {
          function botonMenu() {
            return document.getElementById('mobile-menu-button');
          }
          function cerrarMenu() {
            document.body.classList.remove('mobile-menu-open');
            const boton = botonMenu();
            if (boton) {
              boton.setAttribute('aria-expanded', 'false');
              boton.setAttribute('title', 'Abrir menú de navegación');
            }
          }
          function alternarMenu() {
            const abierto = document.body.classList.toggle('mobile-menu-open');
            const boton = botonMenu();
            if (boton) {
              boton.setAttribute('aria-expanded', abierto ? 'true' : 'false');
              boton.setAttribute('title', abierto ? 'Cerrar menú de navegación' : 'Abrir menú de navegación');
            }
          }
          document.addEventListener('click', function (evento) {
            if (evento.target.closest('#mobile-menu-button')) {
              alternarMenu();
              return;
            }
            if (evento.target.closest(
              '#mobile-menu-overlay, .side-back, .side-section, .vice-button, .indicator-button'
            )) {
              cerrarMenu();
            }
          });
          document.addEventListener('keydown', function (evento) {
            if (evento.key === 'Escape') cerrarMenu();
          });
          window.addEventListener('resize', function () {
            if (window.innerWidth > 1400) cerrarMenu();
          });
        })();
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
                    if codigo in {"pnd", "presupuesto"}:
                        # Orden institucional de los cinco ejes estratégicos.
                        # Las dependencias que no pertenecen a esos cinco ejes
                        # se conservan después, sin eliminarlas.
                        orden_ejes = {
                            "viceministerio de gestion educativa": 1,
                            "viceministerio de educacion superior": 2,
                            "viceministerio de educacion": 3,
                            "viceministerio del deporte": 4,
                            "viceministerio de deporte": 4,
                            "viceministerio de cultura": 5,
                        }

                        def clave_orden_eje(nombre):
                            texto = unicodedata.normalize("NFKD", str(nombre).strip().lower())
                            texto = "".join(c for c in texto if not unicodedata.combining(c))
                            return (orden_ejes.get(texto, 99), texto)

                        opciones_vice = sorted(opciones_vice, key=clave_orden_eje)
                    if codigo == "presupuesto":
                        opciones_vice = [
                            x for x in opciones_vice
                            if _normalizar_encabezado_presupuesto(x)
                            not in {"SIN_CLASIFICACION", "SIN_CLASIFICAR", "NO_APLICA"}
                        ]
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


def formato_valor_indicador(valor, unidad):
    """Formatea según la unidad: puntos de miles para conteos y % aparte."""
    unidad_normalizada = "".join(
        c for c in unicodedata.normalize("NFKD", str(unidad or "").strip().lower())
        if not unicodedata.combining(c)
    )
    return formato_conteo(valor) if unidad_normalizada == "numero" else formato_valor(valor)


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
        html.Div([html.Strong(formato_valor_indicador(valor, unidad)), html.B(sufijo)], className="metric-value"),
        html.Small(f"Período de referencia: {etiqueta_periodo(año)}", className="metric-year"),
    ], className="detail-item highlight metric-card")


def ficha_metrica_periodo(etiqueta, valor, periodo, unidad):
    """Variante para KPI: el período puede ser un año simple o un año lectivo (2024-2025)."""
    unidad_texto = str(unidad or "").strip()
    sufijo = "%" if "porcentaje" in unidad_texto.lower() else ""
    return html.Div([
        html.Span(etiqueta),
        html.Div([html.Strong(formato_valor_indicador(valor, unidad)), html.B(sufijo)], className="metric-value"),
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
    texto = formato_valor_indicador(valor, unidad)
    if "porcentaje" in str(unidad or "").lower():
        texto = f"{texto}%"
    return f"<b>{texto}</b>"


def datos_hover(periodos, valores, unidad):
    """Período y valor ya formateado para evitar comas anglosajonas en hover."""
    return [[periodo, formato_texto_barra(valor, unidad)]
            for periodo, valor in zip(periodos, valores)]


def clase_grafico_responsivo(cantidad_periodos):
    """Clasifica el ancho útil del gráfico para portátil, tableta y celular."""
    if cantidad_periodos >= 9:
        return "indicator-chart indicator-chart-long"
    if cantidad_periodos >= 5:
        return "indicator-chart indicator-chart-medium"
    return "indicator-chart indicator-chart-short"


def geometria_barras_por_periodo(grupo, columnas):
    """Ubica solo las barras con datos usando un ancho visual uniforme.

    Plotly reserva normalmente un espacio para cada serie, incluso cuando su
    valor es NaN. Esta geometría centra una barra y separa dos o tres, pero
    mantiene exactamente el mismo ancho en Línea base, Meta y Estimador.
    """
    posiciones = list(range(len(grupo)))
    x_por_columna = {columna: [] for columna in columnas}

    ancho_uniforme = .34
    desplazamientos_por_cantidad = {
        1: [0.0],
        2: [-.20, .20],
        3: [-.36, 0.0, .36],
    }
    for posicion, (_, fila) in zip(posiciones, grupo.iterrows()):
        disponibles = [columna for columna in columnas if pd.notna(fila[columna])]
        desplazamientos = desplazamientos_por_cantidad.get(
            len(disponibles), [0.0] * len(disponibles)
        )
        ubicaciones = dict(zip(disponibles, desplazamientos))
        for columna in columnas:
            x_por_columna[columna].append(posicion + ubicaciones.get(columna, 0.0))

    # Se devuelve un único escalar. Plotly puede aplicar ajustes distintos si
    # recibe listas de anchos, aunque todos sus elementos tengan el mismo valor.
    return posiciones, x_por_columna, ancho_uniforme


# ---------------------------------------------------------------------------
# Plan Nacional de Desarrollo · gráfico y ficha
# ---------------------------------------------------------------------------
def crear_grafico_pnd(grupo):
    grupo = grupo.sort_values("_orden_periodo")
    años = list(grupo["Año"].map(etiqueta_periodo))
    unidad = primer_texto(grupo, "Unidad de medida", "")

    # Posiciones reales dentro de cada año. Una barra sola queda centrada y
    # ancha; cuando coinciden varias, se separan sin amontonar sus etiquetas.
    columnas_barras = ["Línea base", "Meta", "Estimador"]
    posiciones, x_barras, ancho_barra = geometria_barras_por_periodo(
        grupo, columnas_barras
    )
    # Horizontal por defecto. Un ajuste JavaScript responsivo cambia a vertical
    # únicamente cuando el contenedor del gráfico queda estrecho.
    angulo_texto = 0

    fig = go.Figure()
    fig.add_bar(x=x_barras["Línea base"], y=grupo["Línea base"], name="Línea base",
                marker=dict(color="#80BBBF", line=dict(color="#446381", width=1.4)),
                text=[formato_texto_barra(v, unidad) for v in grupo["Línea base"]], textposition="outside",
                textfont=dict(size=9, color="#232D5A", family="Arial"), constraintext="none",
                textangle=angulo_texto, width=ancho_barra, cliponaxis=False,
                customdata=datos_hover(años, grupo["Línea base"], unidad),
                hovertemplate="Período %{customdata[0]}<br>Línea base: %{customdata[1]}<extra></extra>")
    fig.add_bar(x=x_barras["Meta"], y=grupo["Meta"], name="Meta",
                marker=dict(color="#446381", line=dict(color="#232D5A", width=1.1)),
                text=[formato_texto_barra(v, unidad) for v in grupo["Meta"]], textposition="outside",
                textfont=dict(size=9, color="#232D5A", family="Arial"), constraintext="none",
                textangle=angulo_texto, width=ancho_barra, cliponaxis=False,
                customdata=datos_hover(años, grupo["Meta"], unidad),
                hovertemplate="Período %{customdata[0]}<br>Meta: %{customdata[1]}<extra></extra>")
    fig.add_bar(x=x_barras["Estimador"], y=grupo["Estimador"], name="Estimador",
                marker=dict(color="#ECED80", line=dict(color="#4F449A", width=1.3)),
                text=[formato_texto_barra(v, unidad) for v in grupo["Estimador"]],
                textposition="outside",
                textfont=dict(size=9, color="#232D5A", family="Arial"), constraintext="none",
                textangle=angulo_texto, width=ancho_barra, cliponaxis=False,
                customdata=datos_hover(años, grupo["Estimador"], unidad),
                hovertemplate="Período %{customdata[0]}<br>Estimador: %{customdata[1]}<extra></extra>")
    candidatos = pd.concat([grupo["Línea base"], grupo["Meta"], grupo["Estimador"]]).dropna()
    techo = candidatos.max() * 1.22 if not candidatos.empty else None
    fig.update_layout(
        template="plotly_white", barmode="overlay",
        height=410,
        margin=dict(l=55, r=25, t=82, b=52), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="#fff",
        font=dict(family="Arial", color="#232D5A", size=13),
        title=dict(text="Evolución del indicador", x=.02, font=dict(size=20)),
        legend=dict(orientation="h", y=1.13, x=1, xanchor="right"),
        xaxis=dict(title="Año / período lectivo", showgrid=False, fixedrange=True,
                   tickmode="array", tickvals=posiciones, ticktext=años,
                   range=[-.60, posiciones[-1] + .60] if posiciones else [-.60, .60]),
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
                        className=clase_grafico_responsivo(len(grupo)),
                        style={"width": "100%", "height": "410px"}),
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
    angulo_texto = 0
    claves_periodo = list(grupo["Periodo_id"])
    periodos = list(grupo["Periodo_grafico"])
    periodos_completos = list(grupo["Periodo"])
    # Un intervalo uniforme por período aprovecha todo el eje X. Las barras son
    # más anchas en series cortas y mantienen aire suficiente en las mensuales.
    columnas_barras = ["Línea base", "Meta", "Estimador"]
    posiciones, x_barras, ancho_barra = geometria_barras_por_periodo(
        grupo, columnas_barras
    )

    fig = go.Figure()
    fig.add_bar(x=x_barras["Línea base"], y=grupo["Línea base"], name="Línea base",
                marker=dict(color="#80BBBF", line=dict(color="#446381", width=1.4)),
                text=[formato_texto_barra(v, unidad) for v in grupo["Línea base"]],
                textposition="outside", textfont=dict(size=etiqueta_size, color="#232D5A", family="Arial"), constraintext="none",
                textangle=angulo_texto, width=ancho_barra, cliponaxis=False,
                customdata=datos_hover(periodos_completos, grupo["Línea base"], unidad),
                hovertemplate="%{customdata[0]}<br>Línea base: %{customdata[1]}<extra></extra>")
    fig.add_bar(x=x_barras["Meta"], y=grupo["Meta"], name="Meta",
                marker=dict(color="#446381", line=dict(color="#232D5A", width=1.1)),
                text=[formato_texto_barra(v, unidad) for v in grupo["Meta"]],
                textposition="outside", textfont=dict(size=etiqueta_size, color="#232D5A", family="Arial"),
                constraintext="none", textangle=angulo_texto,
                width=ancho_barra, cliponaxis=False,
                customdata=datos_hover(periodos_completos, grupo["Meta"], unidad),
                hovertemplate="%{customdata[0]}<br>Meta: %{customdata[1]}<extra></extra>")
    fig.add_bar(x=x_barras["Estimador"], y=grupo["Estimador"], name="Ejecutado",
                marker=dict(color="#ECED80", line=dict(color="#4F449A", width=1.3)),
                text=[formato_texto_barra(v, unidad) for v in grupo["Estimador"]],
                textposition="outside", textfont=dict(size=etiqueta_size, color="#232D5A", family="Arial"),
                constraintext="none", textangle=angulo_texto,
                width=ancho_barra, cliponaxis=False,
                customdata=datos_hover(periodos_completos, grupo["Estimador"], unidad),
                hovertemplate="%{customdata[0]}<br>Ejecutado: %{customdata[1]}<extra></extra>")

    candidatos = pd.concat([grupo["Línea base"], grupo["Meta"], grupo["Estimador"]]).dropna()
    techo = candidatos.max() * (1.30 if len(grupo) >= 9 else 1.24) if not candidatos.empty else None

    fig.update_layout(
        template="plotly_white", barmode="overlay",
        height=410, autosize=True,
        margin=dict(l=55, r=15, t=82, b=70), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="#fff",
        font=dict(family="Arial", color="#232D5A", size=13),
        title=dict(text="Evolución del indicador", x=.02, font=dict(size=20)),
        legend=dict(orientation="h", y=1.13, x=1, xanchor="right"),
        # Categoría simple con orden forzado explícitamente: es la forma
        # confiable de garantizar el orden cronológico. El eje "multicategory"
        # ignoraba tanto categoryorder="trace" como categoryarray.
        xaxis=dict(title="Período", showgrid=False, type="linear",
                   tickmode="array", tickvals=posiciones, ticktext=periodos,
                   fixedrange=True, tickangle=0,
                   range=[-.60, posiciones[-1] + .60] if posiciones else [-.60, .60],
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
                        className=clase_grafico_responsivo(len(grupo)),
                        style={"width": "100%", "height": "410px"}),
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

    def contenido_boton(nombre):
        if compacto:
            return nombre
        return html.Span(nombre, className="vision-landing-vice-nombre")

    return html.Nav(
        [html.Div(
            html.Button(contenido_boton(nombre),
                        id={"type": "vision-tab", "index": codigo}, n_clicks=0,
                        className="vision-vice-pill active" if codigo == activo else "vision-vice-pill",
                        type="button"),
            className="vision-vice-row",
        ) for codigo, nombre in VISION_TABS],
        className="vision-vice-menu vision-vice-menu-compacta" if compacto else "vision-vice-menu",
        **{"aria-label": "Viceministerios de Visión Ejecutiva"},
    )


ICONOS_TARJETAS = {
    "institucion": "<path d='M7 20h34M11 20v20m26-20v20M5 40h38M24 7 6 17h36z'/><path d='M17 25v9m7-9v9m7-9v9'/>",
    "estudiantes": "<path d='m5 17 19-9 19 9-19 9z'/><path d='M12 21v10c7 6 17 6 24 0V21M43 17v13'/>",
    "docentes": "<circle cx='17' cy='13' r='6'/><path d='M6 38c1-9 5-14 11-14s10 5 11 14M29 13h13v18H31M34 19h4m-4 6h4'/>",
    "inclusion": "<circle cx='24' cy='8' r='4'/><path d='M10 16h28M24 13v12m0 0-9 15m9-15 9 15M15 20l-5 11m23-11 5 11'/>",
    "alimentacion": "<circle cx='24' cy='25' r='14'/><path d='M24 11v28M10 25h28M5 9v12m4-12v12M7 21v18M41 9v30m0-30c-5 4-5 12 0 14'/>",
    "uniformes": "<path d='m16 8 8 5 8-5 10 8-7 8-3-4v20H16V20l-3 4-7-8z'/><path d='M20 10c1 4 7 4 8 0'/>",
    "textos": "<path d='M7 9h13c3 0 4 2 4 5v27c0-3-1-5-4-5H7zM41 9H28c-3 0-4 2-4 5v27c0-3 1-5 4-5h13z'/><path d='M12 16h7m-7 6h7m10-6h7m-7 6h7'/>",
    "mobiliario": "<path d='M12 9h24v17H12zM9 26h30v6H9zM14 32v8m20-8v8'/><path d='M18 15h12'/>",
    "transporte": "<path d='M7 14h29l6 9v13H6V20z'/><path d='M12 14v10h26M19 14v10'/><circle cx='14' cy='37' r='4'/><circle cx='35' cy='37' r='4'/>",
    "epja": "<circle cx='15' cy='13' r='5'/><path d='M6 34c1-8 4-13 9-13s8 5 9 13M27 12h14v23H27zM31 18h6m-6 6h6m-6 6h4'/>",
    "becas": "<path d='m5 16 19-9 19 9-19 9z'/><path d='M12 21v9c7 6 17 6 24 0v-9M40 18v11'/><path d='m35 36 3 2 4-3-1 5 3 3-5 1-2 4-2-4-5-1 4-3z'/>",
    "psicologia": "<path d='M18 39h12M24 39v-7'/><path d='M16 31c-6-3-9-9-7-15 2-7 8-10 15-10s13 5 14 12c1 6-2 11-7 14'/><path d='M18 17c0-3 2-5 5-5 4 0 6 4 4 7 4 0 6 5 3 8-2 2-4 2-6 1-3 2-7 0-7-4-3-1-3-6 1-7'/>",
    "seguridad": "<path d='M24 5 39 11v11c0 10-6 17-15 21C15 39 9 32 9 22V11z'/><path d='m17 24 5 5 10-11'/>",
    "reinsercion": "<path d='m6 17 18-9 18 9-18 9z'/><path d='M13 22v9c6 5 15 6 22 1M36 25a12 12 0 1 1-7-11'/><path d='m35 8-6 6 8 2'/>",
    "hospital": "<path d='M8 13h32v28H8zM15 13V7h18v6'/><path d='M24 19v15m-7-7h14'/>",
    "bachillerato": "<path d='M8 9h24l8 8v22H8zM32 9v9h8'/><path d='M15 25h18M15 31h18'/><path d='m15 19 2 2 4-5'/>",
    "safpi": "<path d='M6 22 24 8l18 14'/><path d='M10 20v20h28V20'/><circle cx='18' cy='25' r='4'/><circle cx='30' cy='25' r='4'/><path d='M13 37c1-5 3-8 5-8s4 3 5 8m2 0c1-5 3-8 5-8s4 3 5 8'/>",
    "nap": "<path d='M7 39h9v-8h8v-8h8v-8h9'/><path d='m34 8 7 7-7 7'/><path d='M9 10h13v12H9zM13 14h5m-5 4h5'/>",
    "cupos": "<rect x='8' y='7' width='32' height='34' rx='4'/><path d='M17 7v6m14-6v6M14 19h20'/><path d='m16 29 5 5 11-11'/>",
    "internacional": "<circle cx='24' cy='24' r='17'/><path d='M7 24h34M24 7c5 5 8 11 8 17s-3 12-8 17c-5-5-8-11-8-17s3-12 8-17z'/><path d='m31 29 4 3 5-5'/>",
    "validatec": "<path d='M9 8h25l6 6v26H9zM34 8v7h6'/><path d='M16 20h16M16 26h10'/><circle cx='30' cy='33' r='7'/><path d='m27 33 2 2 4-5'/>",
    "tecnica": "<path d='m6 16 18-9 18 9-18 9zM12 21v9c7 6 17 6 24 0v-9'/><path d='M35 34l7 7m-1-9-9 9'/><circle cx='36.5' cy='36.5' r='5.5'/>",
    "cultura": "<path d='M24 7c-10 0-18 7-18 16 0 8 6 14 13 14h4c2 0 3-2 2-4l-1-2c-1-2 1-4 3-4h5c6 0 10-4 10-9C42 11 34 7 24 7z'/><circle cx='14' cy='20' r='2'/><circle cx='20' cy='14' r='2'/><circle cx='29' cy='14' r='2'/><circle cx='35' cy='20' r='2'/>",
    "deporte": "<circle cx='24' cy='24' r='18'/><path d='m24 14 7 5-3 8h-8l-3-8zM24 6v8M7 19l10 1M12 37l8-10m16 10-8-10m13-8-10 1'/>",
    "digital": "<rect x='6' y='8' width='36' height='25' rx='3'/><path d='M16 41h16M24 33v8'/><path d='m17 22 5 5 10-11'/>",
    "programa": "<circle cx='24' cy='24' r='17'/><circle cx='24' cy='24' r='9'/><path d='m24 7 3 6-3 4-3-4zM41 24l-6 3-4-3 4-3zM24 41l-3-6 3-4 3 4zM7 24l6-3 4 3-4 3z'/>",
    "vitalicios": "<circle cx='24' cy='17' r='10'/><path d='m18 26-3 15 9-5 9 5-3-15'/><path d='m24 11 2 4 5 .7-3.5 3.4.8 4.9-4.3-2.3-4.3 2.3.8-4.9-3.5-3.4 5-.7z'/>",
    "alto-rendimiento": "<path d='M8 37h32M12 37V25h9v12m0 0V18h9v19m0 0V29h7v8'/><path d='m24 7 2.2 4.5 5 .7-3.6 3.5.9 5-4.5-2.4-4.5 2.4.9-5-3.6-3.5 5-.7z'/>",
    "entrenadores": "<path d='M9 9h30v32H9zM17 9V6h14v3'/><path d='M16 19h16M16 26h10M16 33h7'/><circle cx='34' cy='31' r='5'/><path d='m31 31 2 2 4-5'/>",
    "usuarios-deporte": "<circle cx='17' cy='15' r='6'/><circle cx='33' cy='17' r='5'/><path d='M6 39c1-10 5-16 11-16s10 6 11 16M27 27c2-3 4-5 7-5 5 0 8 5 9 13'/>",
    "juegos-juveniles": "<path d='M17 8h14v8c0 7-3 12-7 12s-7-5-7-12z'/><path d='M17 12h-6v4c0 5 3 8 8 8m12-12h6v4c0 5-3 8-8 8M24 28v7m-8 6h16M19 35h10'/>",
    "activate": "<circle cx='28' cy='8' r='4'/><path d='m25 14-7 8 6 5 5 12M20 21l-9 4m14-10 8 8 7 1M24 27l-8 12'/>",
    "vamos-cancha": "<rect x='6' y='10' width='36' height='28' rx='3'/><path d='M6 18h9v12H6m36-12h-9v12h9M24 10v28'/><circle cx='24' cy='24' r='5'/>",
    "patrimonio-cultural": "<path d='M6 18h36L24 7zM9 39h30M12 18v17m8-17v17m8-17v17m8-17v17M7 35h34'/>",
    "memoria-social": "<path d='M9 8h25l5 5v27H9zM34 8v6h6'/><path d='M16 20h16M16 27h16M16 34h10'/><circle cx='13' cy='13' r='3'/>",
    "organizaciones": "<circle cx='24' cy='11' r='5'/><circle cx='12' cy='25' r='5'/><circle cx='36' cy='25' r='5'/><path d='M19 14l-5 7m15-7 5 7M17 27h14M6 41c1-7 3-11 6-11s5 4 6 11m12 0c1-7 3-11 6-11s5 4 6 11M18 41c1-9 3-14 6-14s5 5 6 14'/>",
    "empresa-cultural": "<path d='M8 41V14h20v27M28 22h12v19M13 20h4m5 0h2m-11 7h4m5 0h2m-11 7h4m5 0h2m10-6h3m-3 7h3M5 41h38'/>",
}

PALETA_ICONOS_TARJETAS = {
    "institucion": ("#b18400", "#fff7d6"), "estudiantes": ("#287bb5", "#eaf6ff"),
    "docentes": ("#7b2f78", "#f8ecf7"), "inclusion": ("#3c6eab", "#eef5ff"),
    "alimentacion": ("#16806d", "#e7f8f3"), "uniformes": ("#2f8f4e", "#eaf7ee"),
    "textos": ("#b57d00", "#fff5d9"), "mobiliario": ("#633ba3", "#f1ecfb"),
    "transporte": ("#247cb3", "#e8f5fc"), "epja": ("#6c3c9f", "#f3edfb"),
    "becas": ("#8050b5", "#f3ecfb"), "psicologia": ("#b14d77", "#fbeef3"),
    "seguridad": ("#267d65", "#e9f7f2"), "reinsercion": ("#356bb0", "#eaf2fc"),
    "hospital": ("#b83d55", "#fdeef1"), "bachillerato": ("#4d56a7", "#eef0ff"),
    "safpi": ("#9a4e85", "#faedf6"), "nap": ("#d07b18", "#fff1df"),
    "cupos": ("#287bb5", "#eaf6ff"), "internacional": ("#347c72", "#eaf7f4"),
    "validatec": ("#5a55a7", "#efeffb"), "tecnica": ("#b06a13", "#fff2df"),
    "cultura": ("#a03f78", "#faedf5"), "deporte": ("#238071", "#e9f7f4"),
    "digital": ("#3d67ad", "#ebf2fd"), "programa": ("#503a98", "#f1eefb"),
    "vitalicios": ("#a16f00", "#fff5d8"),
    "alto-rendimiento": ("#503a98", "#f1eefb"),
    "entrenadores": ("#287bb5", "#eaf6ff"),
    "usuarios-deporte": ("#238071", "#e9f7f4"),
    "juegos-juveniles": ("#a03f78", "#faedf5"),
    "activate": ("#c18a00", "#fff6d9"),
    "vamos-cancha": ("#27806e", "#e9f7f3"),
    "patrimonio-cultural": ("#9a6b00", "#fff5d8"),
    "memoria-social": ("#5c55a8", "#efeffb"),
    "organizaciones": ("#287bb5", "#eaf6ff"),
    "empresa-cultural": ("#a03f78", "#faedf5"),
}


def _clave_icono_tarjeta(codigo="", titulo=""):
    """Selecciona un símbolo semántico sin depender de nombres exactos del Excel."""
    texto = _sin_tildes(f"{codigo} {titulo}")
    reglas = [
        (("exdeportistas vitalicios", "ex deportistas vitalicios"), "vitalicios"),
        (("alto rendimiento",), "alto-rendimiento"),
        (("entrenadores", "entrenador"), "entrenadores"),
        (("usuarios de deporte",), "usuarios-deporte"),
        (("juegos deportivos nacionales", "ciclo juvenil"), "juegos-juveniles"),
        (("activate",), "activate"),
        (("vamos a la cancha",), "vamos-cancha"),
        (("patrimonio cultural",), "patrimonio-cultural"),
        (("memoria social",), "memoria-social"),
        (("organizaciones sociales",), "organizaciones"),
        (("empresa cultural",), "empresa-cultural"),
        (("educacion cultural",), "estudiantes"),
        (("instituciones educativas",), "institucion"),
        (("total de artistas", "total artistas"), "usuarios-deporte"),
        (("otras categorias culturales",), "programa"),
        (("safpi", "atencion familiar", "primera infancia"), "safpi"),
        (("nap", "nivelacion y aceleracion"), "nap"),
        (("cupos", "admision"), "cupos"),
        (("because he is nice", "cooperacion", "asuntos internacionales"), "internacional"),
        (("validatec", "validacion de trayectoria"), "validatec"),
        (("becas tecnicas", "tecnologicas financiadas", "formacion tecnica"), "tecnica"),
        (("robotica", "transformacion digital", "brecha digital"), "digital"),
        (("deporte", "activate", "cancha"), "deporte"),
        (("cultura", "arte", "museo", "cine", "festival"), "cultura"),
        (("discapacidad", "necesidades educativas", "inclusion"), "inclusion"),
        (("alimentacion",), "alimentacion"),
        (("uniforme",), "uniformes"),
        (("texto", "libro", "lectura"), "textos"),
        (("mobiliario",), "mobiliario"),
        (("transporte",), "transporte"),
        (("institucion", "universidad", "infraestructura"), "institucion"),
        (("docente", "profesional de la educacion"), "docentes"),
        (("estudiante",), "estudiantes"),
        (("epja", "jovenes y adultos"), "epja"),
        (("beca",), "becas"),
        (("dece", "psicolog"), "psicologia"),
        (("comunidades seguras", "escudo", "seguridad"), "seguridad"),
        (("reinsercion", "todos al aula"), "reinsercion"),
        (("hospitalaria",), "hospital"),
        (("bachillerato",), "bachillerato"),
    ]
    return next((icono for palabras, icono in reglas if any(p in texto for p in palabras)),
                "programa")


def icono_tarjeta(codigo="", titulo=""):
    """Genera un icono SVG lineal, ligero y autónomo para cada tarjeta."""
    clave = _clave_icono_tarjeta(codigo, titulo)
    color, fondo = PALETA_ICONOS_TARJETAS[clave]
    svg = (
        "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 48 48' "
        f"fill='none' stroke='{color}' stroke-width='2.2' "
        f"stroke-linecap='round' stroke-linejoin='round'>{ICONOS_TARJETAS[clave]}</svg>"
    )
    return html.Span(
        html.Img(
            src="data:image/svg+xml;utf8," + quote(svg), alt="",
            style={"display": "block", "width": "24px", "height": "24px"},
            **{"aria-hidden": "true"},
        ),
        className="vision-card-symbol",
        style={
            "display": "inline-flex", "width": "36px", "height": "36px",
            "marginBottom": "0", "alignItems": "center", "justifyContent": "center",
            "border": f"1px solid {color}33", "borderRadius": "11px",
            "background": fondo, "boxShadow": f"0 5px 12px {color}1a",
            "flex": "0 0 auto",
        },
    )


def cabecera_tarjeta(codigo, titulo, ayuda_texto=None):
    """Alinea símbolo y título; la ayuda aparece al pasar por la tarjeta."""
    return html.Div([
        icono_tarjeta(codigo, titulo),
        html.Span(
            titulo, className="vision-card-label",
            style={"minHeight": "0", "margin": "0", "flex": "1 1 auto"},
        ),
    ], className="vision-card-title-row", style={
        "display": "flex", "alignItems": "center", "gap": "10px",
        "minWidth": "0", "marginBottom": "7px",
    })


def _ayuda_tarjeta(nombre, descripcion):
    """Tooltip institucional: nombre en negrita y explicación justificada."""
    return html.Div([
        html.Strong(nombre),
        html.Span(" " + descripcion),
    ], className="programa-info-tooltip", role="tooltip")


def tarjeta_gestion_educativa(codigo, titulo, subtitulo, valor, detalles=None,
                              unidad=None, filtros=None, contenido_id=None, valor_id=None,
                              ayuda_nombre=None, ayuda_descripcion=None, periodo=None):
    tiene_ayuda = bool(ayuda_nombre and ayuda_descripcion)
    ayuda_texto = (f"{ayuda_nombre} {ayuda_descripcion}" if tiene_ayuda else None)
    contenido = [
        _ayuda_tarjeta(ayuda_nombre, ayuda_descripcion) if tiene_ayuda else None,
        cabecera_tarjeta(codigo, titulo, ayuda_texto),
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
        html.Small(periodo, className="vision-card-periodo") if periodo else None,
        html.Div(filtros, className="vision-card-filters") if filtros else None,
    ]
    return html.Article(
        contenido,
        className=(f"vision-exec-card {codigo}"
                   f"{' programa-con-ayuda' if tiene_ayuda else ''}"),
        tabIndex=0 if tiene_ayuda else None,
    )


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


# ---------------------------------------------------------------------------
# Educación Superior · Universidades y Escuelas Politécnicas / Institutos
# Superiores Técnicos Tecnológicos. Cifras de la matriz institucional
# "POLITECNICAS" / "Instituto Superior Técnico Tecnológico"; se muestran
# únicamente los totales ("Nro."), sin los porcentajes de la matriz.
DATOS_UNIVERSIDADES_POLITECNICAS = {
    "instituciones_total": 64,
    "instituciones_financiamiento": {
        "Particular Autofinanciada": 20,
        "Particular Cofinanciada": 8,
        "Pública": 36,
    },
    "estudiantes_total": 909116,
    "estudiantes_financiamiento": {
        "Particular Autofinanciada": 162081,
        "Particular Cofinanciada": 170545,
        "Pública": 576490,
    },
    "estudiantes_sexo": {"Hombre": 391975, "Mujer": 517141},
    "estudiantes_nivel": {"Tercer Nivel": 805739, "Cuarto Nivel": 103377},
    "docentes_total": 38828,
    "docentes_financiamiento": {
        "Particular Autofinanciada": 7820,
        "Particular Cofinanciada": 10443,
        "Pública": 20565,
    },
    "docentes_sexo": {"Hombre": 22124, "Mujer": 16704},
    "corte": "año 2025",
    "actualizacion": "septiembre de cada año",
}

DATOS_INSTITUTOS_TECNICOS_TECNOLOGICOS = {
    "instituciones_total": 214,
    "instituciones_financiamiento": {"Público": 152, "Particular": 62},
    "estudiantes_total": 122044,
    "estudiantes_financiamiento": {"Público": 39318, "Particular": 82726},
    "estudiantes_sexo": {"Hombre": 57011, "Mujer": 65033},
    "docentes_total": 9315,
    "docentes_financiamiento": {"Público": 3586, "Particular": 5729},
    "docentes_sexo": {"Hombre": 5312, "Mujer": 4003},
    "corte": "II semestre 2025",
    "actualizacion": "junio de cada año",
}


def _pie_corte_actualizacion(corte, actualizacion):
    """Mismo pie de dos líneas (período / frecuencia) que usan las tarjetas
    de Educación Media, pero con el texto 'Corte' en vez de 'Año lectivo'."""
    return html.Div([
        html.Span(f"Corte: {corte}"),
        html.Span(f"Actualización: {actualizacion}"),
    ], className="educacion-media-periodo-pie")


def _desglose_con_grupos(grupos):
    """Arma los mini-cuadros de desglose (mismo formato que usan las
    tarjetas de Educación Media). 'grupos' es una lista de tuplas
    (titulo_del_grupo_o_None, {etiqueta: valor}); si hay más de un grupo se
    antepone un rótulo (igual que 'Sostenimiento'/'Régimen' en la tarjeta de
    necesidades educativas especiales)."""
    mostrar_titulos = len(grupos) > 1
    hijos = []
    for titulo, datos in grupos:
        if mostrar_titulos and titulo:
            hijos.append(html.Span(titulo, className="nee-grupo-titulo"))
        hijos.extend(
            html.Div([
                html.Span(nombre),
                html.Strong(formato_valor(valor)),
            ], className=("vision-mini-stat secondary" if i % 2 else "vision-mini-stat"))
            for i, (nombre, valor) in enumerate(datos.items())
        )
    return hijos


def _tarjeta_educacion_superior_nivel(codigo, titulo, total, grupos, pie):
    """Tarjeta con el mismo formato que las de Educación Media: total
    principal + desglose(s) en mini-cuadros + pie de período."""
    return html.Article([
        cabecera_tarjeta(codigo, titulo),
        html.Div([
            html.Strong(formato_valor(total), className="vision-card-total"),
        ], className="vision-card-value-row"),
        html.Div(_desglose_con_grupos(grupos), className="nee-desglose"),
        pie,
    ], className=f"vision-exec-card {codigo} educacion-superior-nivel-card")


def tarjetas_universidades_politecnicas():
    """Instituciones, Estudiantes y Docentes de Universidades y Escuelas
    Politécnicas, con el mismo formato que Educación Media."""
    d = DATOS_UNIVERSIDADES_POLITECNICAS
    pie = _pie_corte_actualizacion(d["corte"], d["actualizacion"])
    return [
        _tarjeta_educacion_superior_nivel(
            "politecnicas-instituciones", "Instituciones",
            d["instituciones_total"],
            [(None, d["instituciones_financiamiento"])],
            pie,
        ),
        _tarjeta_educacion_superior_nivel(
            "politecnicas-estudiantes", "Estudiantes",
            d["estudiantes_total"],
            [
                ("Sexo", d["estudiantes_sexo"]),
                ("Nivel de formación", d["estudiantes_nivel"]),
            ],
            pie,
        ),
        _tarjeta_educacion_superior_nivel(
            "politecnicas-docentes", "Docentes",
            d["docentes_total"],
            [("Sexo", d["docentes_sexo"])],
            pie,
        ),
    ]


def tarjetas_institutos_tecnicos_tecnologicos():
    """Instituciones, Estudiantes y Docentes de Institutos Superiores
    Técnicos Tecnológicos, con el mismo formato que Educación Media."""
    d = DATOS_INSTITUTOS_TECNICOS_TECNOLOGICOS
    pie = _pie_corte_actualizacion(d["corte"], d["actualizacion"])
    return [
        _tarjeta_educacion_superior_nivel(
            "itts-instituciones", "Instituciones",
            d["instituciones_total"],
            [(None, d["instituciones_financiamiento"])],
            pie,
        ),
        _tarjeta_educacion_superior_nivel(
            "itts-estudiantes", "Estudiantes",
            d["estudiantes_total"],
            [
                ("Financiamiento", d["estudiantes_financiamiento"]),
                ("Sexo", d["estudiantes_sexo"]),
            ],
            pie,
        ),
        _tarjeta_educacion_superior_nivel(
            "itts-docentes", "Docentes",
            d["docentes_total"],
            [
                ("Financiamiento", d["docentes_financiamiento"]),
                ("Sexo", d["docentes_sexo"]),
            ],
            pie,
        ),
    ]


# ---------------------------------------------------------------------------
# Educación Superior · Programas y servicios de becas (actualizado a pedido
# de la Viceministra, octubre 2026). Reemplaza la carga anterior desde
# "Mapeo - Rendicion de cuentas.xlsx": estos 6 programas se mantienen fijos
# en código hasta que se defina una fuente propia para becas.
EJE2_BECAS_EDUCACION_SUPERIOR = [
    {
        "titulo": "Becas Fulbright",
        "beneficiarios": 8,
        "unidad_beneficiarios": "Beneficiarios",
        "inversion": "$677.800,00",
        "descripcion": (
            "Las Becas Fulbright financian estudios de posgrado (maestrías) en "
            "Estados Unidos, con la comisión Fulbright Ecuador."
        ),
    },
    {
        "titulo": "Becas Nacionales para Fortalecimiento de la Industria de Hidrocarburos",
        "beneficiarios": 141,
        "unidad_beneficiarios": "Beneficiarios",
        "inversion": "$899.870,10",
        "descripcion": (
            "Fortalece las competencias técnicas y de cuarto nivel en áreas "
            "estratégicas de la industria de hidrocarburos, articulando las "
            "necesidades del sector energético con la oferta académica nacional "
            "de las instituciones de educación superior."
        ),
    },
    {
        "titulo": "Programa de Educación Continua “Because He Is Nice”",
        "beneficiarios": 71570,
        "unidad_beneficiarios": "Beneficiarios",
        "inversion": "$1.431.400,00",
        "descripcion": (
            "El programa de Educación Continua “Because He Is Nice” ofrece "
            "becas para estudiar inglés gratis y en línea. Está dirigido a "
            "mayores de 18 años con título de bachiller registrado en el país."
        ),
    },
    {
        "titulo": "Jóvenes del Nuevo Ecuador",
        "beneficiarios": 3276,
        "unidad_beneficiarios": "Beneficiarios",
        "inversion": "$2.715.705,00",
        "descripcion": (
            "Promueve la permanencia de jóvenes de 18 a 29 años en la educación "
            "superior, mediante la entrega de becas parciales que brindan un "
            "apoyo económico único para cubrir gastos de manutención y "
            "facilitar la continuidad de sus estudios de tercer nivel."
        ),
    },
    {
        "titulo": "Mujeres que Transforman",
        "beneficiarios": 2627,
        "unidad_beneficiarios": "Beneficiarias",
        "inversion": "$1.266.214,00",
        "descripcion": (
            "Apoya los gastos de permanencia (alimentación, vivienda, servicios "
            "básicos y transporte interno) durante un período académico del "
            "segundo semestre de 2026, con una subvención de un SBU."
        ),
    },
    {
        "titulo": "Talento Joven",
        "beneficiarios": 3200,
        "unidad_beneficiarios": "Beneficiarios",
        "inversion": "$1.821.536,00",
        "descripcion": (
            "Evitar la deserción: contribuir con un solo pago a los estudiantes "
            "de grado en instituciones públicas para que continúen y culminen "
            "sus carreras universitarias."
        ),
    },
]


def _tarjeta_beca_educacion_superior(dato):
    """Beneficiarios + inversión de un programa de becas, con la nota de
    actualización/corte fija que pidió la Viceministra (no es un período
    histórico seleccionable, como SAFPI/NAP)."""
    ayuda_texto = f"{dato['titulo']}. {dato['descripcion']}"
    return html.Article([
        _ayuda_tarjeta(dato["titulo"], dato["descripcion"]),
        cabecera_tarjeta("becas", dato["titulo"], ayuda_texto),
        html.Div([
            html.Strong(formato_valor(dato["beneficiarios"]), className="vision-card-total"),
            html.Span(dato["unidad_beneficiarios"], className="vision-card-unit"),
        ], className="vision-card-value-row"),
        html.Div([
            html.Span("Inversión"),
            html.Strong(dato["inversion"]),
        ], className="programa-metrica-periodo inversion"),
        html.Div([
            html.Span("Corte: octubre 2026"),
            html.Span("Actualización: conforme a planificación de cada programa"),
        ], className="educacion-media-periodo-pie"),
    ], className="vision-exec-card becas programa-con-ayuda", tabIndex=0)


def tarjetas_becas_educacion_superior():
    return [_tarjeta_beca_educacion_superior(d) for d in EJE2_BECAS_EDUCACION_SUPERIOR]


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


PROGRAMAS_CON_DETALLE = {
    "safpi": {
        "nombre_descripcion": "Servicio de Atención Familiar para la Primera Infancia (SAFPI).",
        "descripcion": (
            "Promueve el desarrollo integral de niñas y niños de 3 y 4 años que no asisten "
            "a una institución educativa, mediante el fortalecimiento del acompañamiento de "
            "sus familias y la participación de la comunidad. Esta iniciativa forma parte de "
            "la estrategia del Ministerio de Educación, Deporte y Cultura para ampliar la "
            "cobertura y mejorar la calidad educativa."
        ),
        "filas": [
            ("2023-2024", "18.578", "746", "$ 9.537.415,44"),
            ("2024-2025", "19.142", "746", "$ 9.544.875,44"),
            ("2025-2026", "20.564", "747", "$ 9.565.140,18"),
            ("2026-2027*", "20.669", "752", "$ 9.638.187,88"),
        ],
        "regimenes": {
            "2023-2024": {"Costa": 9098, "Sierra": 9480},
            "2024-2025": {"Costa": 10147, "Sierra": 8995},
            "2025-2026": {"Costa": 10274, "Sierra": 10290},
            "2026-2027*": {"Costa": 10434, "Sierra": 10235},
        },
        "actualizacion": "julio y diciembre de cada año",
        "nota": "*Nota: En el régimen Sierra-Amazonía se sigue levantando información de niños y niñas.",
    },
    "nap": {
        "nombre_descripcion": "Nivelación y Aceleración Pedagógica (NAP).",
        "descripcion": (
            "Es un servicio de educación formal "
            "que permite nivelar los conocimientos y habilidades de niñas, niños y adolescentes "
            "con escolaridad inconclusa y rezago escolar, mediante un proceso acelerado que "
            "facilita su reincorporación al grado o curso correspondiente a su edad."
        ),
        "filas": [
            ("2023-2024", "6.660", "322", "$ 4.116.686,02"),
            ("2024-2025", "8.172", "385", "$ 4.925.974,59"),
            ("2025-2026", "7.810", "401", "$ 5.134.700,42"),
            ("2026-2027", "8.232", "398", "$ 5.101.062,20"),
        ],
        "regimenes": {
            "2023-2024": {"Costa": 2831, "Sierra": 3829},
            "2024-2025": {"Costa": 3463, "Sierra": 4709},
            "2025-2026": {"Costa": 3313, "Sierra": 4497},
            "2026-2027": {"Costa": 3492, "Sierra": 4740},
        },
        "actualizacion": "julio y diciembre de cada año",
        "nota": "Fuente: Registros administrativos MINEDEC_DEIB.",
    },
}


REINSERCION_POR_ANIO_LECTIVO = {
    "2023-2024": {"Fiscal": 11456, "Fiscomisional": 153, "Municipal": 26, "Particular": 256},
    "2024-2025": {"Fiscal": 45641, "Fiscomisional": 18, "Municipal": 1, "Particular": 40},
    "2025-2026": {"Fiscal": 62512, "Fiscomisional": 2244, "Municipal": 248, "Particular": 5859},
    "2026-2027": {"Fiscal": 40061, "Fiscomisional": 1046, "Municipal": 110, "Particular": 9564},
}


AULAS_HOSPITALARIAS_POR_ANIO = {
    "2023": {"aulas": 90, "beneficiarios": 46148, "docentes": 151, "inversion": "$1.937.946,98"},
    "2024": {"aulas": 97, "beneficiarios": 46248, "docentes": 157, "inversion": "$2.016.521,50"},
    "2025": {"aulas": 98, "beneficiarios": 42965, "docentes": 158, "inversion": "$2.030.945,58"},
    "2026": {"aulas": 99, "beneficiarios": 27142, "docentes": 156, "inversion": "$1.537.195,06"},
}


PROGRAMAS_GESTION_COMPLEMENTARIOS = [
    {
        "clave": "hospital",
        "titulo": "Aulas Hospitalarias",
        "descripcion": (
            "Son espacios educativos y pedagógicos ubicados dentro de recintos de salud que "
            "brindan atención escolar a niños, niñas y jóvenes en situación de enfermedad "
            "física o mental que no pueden asistir de forma regular a las instituciones educativas."
        ),
    },
]


def _ayuda_programa(nombre, descripcion):
    """Tooltip justificado con el nombre institucional destacado al inicio."""
    return _ayuda_tarjeta(nombre, descripcion)


# Definiciones respaldadas por los portales oficiales de Educación Superior
# del MINEDEC/SIAU. Los valores numéricos continúan procediendo de la matriz
# institucional "Mapeo - Rendicion de cuentas.xlsx".
DEFINICIONES_PROGRAMAS_EDUCACION_SUPERIOR = [
    (("cupos aceptados",),
     "Cupos aceptados en Educación Superior.",
     "Resultado del proceso mediante el cual una persona acepta el cupo asignado "
     "por una institución de educación superior y continúa el procedimiento de "
     "matriculación."),
    (("because he is nice",),
     "Because He Is Nice.",
     "Programa de becas de educación continua que cubre totalmente cursos "
     "virtuales de inglés para fortalecer las competencias y oportunidades de "
     "la ciudadanía."),
    (("validatec",),
     "ValidaTec.",
     "Programa de validación del ejercicio profesional dirigido a docentes "
     "bachilleres, mediante el cual se reconoce su experiencia, formación y "
     "conocimientos para acceder a un título de tercer nivel técnico-tecnológico."),
    (("becas tecnicas", "tecnologicas financiadas por aliados"),
     "Becas técnicas y tecnológicas financiadas por aliados.",
     "Becas gestionadas mediante cooperación con instituciones, gobiernos y "
     "entidades aliadas para facilitar el acceso a estudios de tercer nivel "
     "técnico y tecnológico."),
]


def _definicion_programa_educacion_superior(nombre_programa):
    """Encuentra la definición aunque el Excel cambie mayúsculas o tildes."""
    clave = _sin_tildes(nombre_programa)
    for patrones, nombre, descripcion in DEFINICIONES_PROGRAMAS_EDUCACION_SUPERIOR:
        if all(patron in clave for patron in patrones):
            return nombre, descripcion
    return None


# Definiciones con respaldo en fuentes oficiales del Ministerio de Educación.
# Las cifras de beneficiarios se mantienen vinculadas exclusivamente a la
# matriz institucional "Mapeo - Rendicion de cuentas.xlsx".
DEFINICIONES_PROGRAMAS_EDUCACION = [
    (("mecapacito",),
     "Mecapacito.",
     "Plataforma de formación permanente del Ministerio que ofrece cursos e "
     "itinerarios destinados a fortalecer las competencias profesionales, "
     "pedagógicas y didácticas de los profesionales de la educación."),
    (("profesionales dece",),
     "Departamentos de Consejería Estudiantil (DECE).",
     "Equipos de profesionales que brindan acompañamiento y atención "
     "psicosocial, socioemocional, psicoeducativa y de orientación a "
     "estudiantes, familias y comunidad educativa."),
    (("comunidades seguras",),
     "Comunidades educativas seguras.",
     "Acciones orientadas a fortalecer entornos educativos protectores, "
     "prevenir situaciones de riesgo y promover relaciones respetuosas, "
     "saludables y seguras entre los integrantes de la comunidad educativa."),
]


def _definicion_programa_educacion(nombre_programa):
    """Devuelve solo definiciones educativas que tienen respaldo oficial."""
    clave = _sin_tildes(nombre_programa)
    for patrones, nombre, descripcion in DEFINICIONES_PROGRAMAS_EDUCACION:
        if all(patron in clave for patron in patrones):
            return nombre, descripcion
    return None


# Definiciones de los programas del Viceministerio del Deporte con respaldo
# en los portales oficiales del Ministerio del Deporte. Las cifras provienen
# de la matriz "Mapeo - Rendicion de cuentas.xlsx".
DEFINICIONES_PROGRAMAS_DEPORTE = [
    (("vamos a la cancha",),
     "Vamos a la Cancha.",
     "Proyecto del Ministerio del Deporte que, mediante prácticas socio "
     "deportivas, promueve la iniciación y práctica deportiva de niños, "
     "jóvenes y sus familias, fomentando la cultura de paz, la cohesión "
     "social, la inclusión y la convivencia."),
    (("activate",),
     "Actívate.",
     "Estrategia del Ministerio del Deporte para incrementar la actividad "
     "física de la población, mediante clases de actividad física y "
     "recreativa que impulsan el buen uso del tiempo libre."),
]


def _definicion_programa_deporte(nombre_programa):
    """Devuelve la definición oficial del programa de Deporte, si existe."""
    clave = _sin_tildes(nombre_programa)
    for patrones, nombre, descripcion in DEFINICIONES_PROGRAMAS_DEPORTE:
        if all(patron in clave for patron in patrones):
            return nombre, descripcion
    return None


# Definiciones incorporadas únicamente cuando existe respaldo institucional
# del sector Cultura. Las cifras mostradas en las tarjetas continúan leyendo
# la matriz institucional; estos textos solo explican el alcance del programa.
DEFINICIONES_PROGRAMAS_CULTURA = [
    (("visitas a espacios culturales",),
     "Visitas a espacios culturales.",
     "Indicador de asistencia de la ciudadanía a museos, archivos y "
     "bibliotecas vinculados al sector cultural."),
    (("festival internacional de artes vivas de loja",),
     "Festival Internacional de Artes Vivas de Loja.",
     "Encuentro cultural institucionalizado como espacio de encuentro y "
     "fomento de las expresiones culturales y la identidad del Ecuador, con "
     "participación de artistas y gestores locales, nacionales e "
     "internacionales en teatros, calles y plazas de Loja."),
]


def _definicion_programa_cultura(nombre_programa):
    """Devuelve solo definiciones culturales respaldadas por fuentes oficiales."""
    clave = _sin_tildes(nombre_programa)
    for patrones, nombre, descripcion in DEFINICIONES_PROGRAMAS_CULTURA:
        if all(patron in clave for patron in patrones):
            return nombre, descripcion
    return None


def _metricas_programa_por_anio(clave, anio):
    detalle = PROGRAMAS_CON_DETALLE[clave]
    fila = next((f for f in detalle["filas"] if f[0] == anio), detalle["filas"][-1])
    _, estudiantes, docentes, inversion = fila
    inversion = inversion.replace("$ ", "$")
    estudiantes_principal = html.Div([
        html.Strong(estudiantes),
        html.Span("Estudiantes"),
    ], className="programa-estudiantes-principal")
    if clave in {"safpi", "nap"}:
        regimenes = detalle.get("regimenes", {}).get(anio, {})
        fila_estudiantes = html.Div([
            estudiantes_principal,
            *[
                html.Div([
                    html.Span(regimen),
                    html.Strong(formato_valor(valor)),
                ], className="programa-regimen-mini")
                for regimen, valor in regimenes.items()
            ],
        ], className="programa-estudiantes-regimen-fila")
    else:
        fila_estudiantes = estudiantes_principal

    contenido = [html.Div([
        fila_estudiantes,
        html.Div([
            html.Div([html.Span("Docentes"), html.Strong(docentes)],
                     className="programa-metrica-periodo"),
            html.Div([html.Span("Inversión"), html.Strong(inversion)],
                     className="programa-metrica-periodo inversion"),
        ], className="programa-desglose-inferior"),
    ], className="programa-metricas-priorizadas")]
    nota = detalle.get("nota", "")
    # La advertencia de levantamiento corresponde exclusivamente al último
    # período de SAFPI (2026-2027). Las notas de fuente de otros programas
    # continúan visibles en todos sus períodos.
    es_nota_condicional = clave == "safpi" and nota.startswith("*Nota:")
    ultimo_periodo = detalle["filas"][-1][0]
    if nota and (not es_nota_condicional or anio == ultimo_periodo):
        contenido.append(html.P(nota, className="programa-periodo-nota"))
    if clave in {"safpi", "nap"}:
        contenido.append(html.Div([
            html.Span(f"Año lectivo {str(anio).replace('-', '–')}"),
            html.Span(
                "Actualización: " + detalle.get(
                    "actualizacion", "julio y diciembre de cada año"
                )
            ),
        ], className="educacion-media-periodo-pie"))
    return contenido


def _control_anio_programa(clave):
    detalle = PROGRAMAS_CON_DETALLE[clave]
    periodos = [fila[0] for fila in detalle["filas"]]
    return [
        _ayuda_programa(detalle["nombre_descripcion"], detalle["descripcion"]),
        html.Div([
            html.Span("Año"),
            dcc.Dropdown(
                id={"type": "programa-anio", "index": clave},
                options=[{"label": periodo, "value": periodo} for periodo in periodos],
                value=periodos[-1], clearable=False, searchable=False,
                className="programa-anio-dropdown",
            ),
        ], className="programa-anio-control"),
    ]


def _control_anio_complementario(clave, periodos, etiqueta):
    """Selector compacto para Reinserción y Aulas Hospitalarias."""
    return html.Div([
        html.Span(etiqueta),
        dcc.Dropdown(
            id={"type": "programa-complementario-anio", "index": clave},
            options=[
                {"label": periodo.replace("-", "–"), "value": periodo}
                for periodo in periodos
            ],
            value=periodos[-1],
            clearable=False,
            searchable=False,
            className="programa-anio-dropdown",
        ),
    ], className="programa-anio-control")


def _pie_programa_actualizable(etiqueta_periodo, periodo):
    periodo_visible = str(periodo).replace("-", "–")
    return html.Div([
        html.Span(f"{etiqueta_periodo} {periodo_visible}"),
        html.Span("Actualización: diciembre de cada año"),
    ], className="educacion-media-periodo-pie")


def _metricas_reinsercion(anio_lectivo):
    if anio_lectivo not in REINSERCION_POR_ANIO_LECTIVO:
        anio_lectivo = list(REINSERCION_POR_ANIO_LECTIVO)[-1]
    datos = REINSERCION_POR_ANIO_LECTIVO[anio_lectivo]
    return [
        html.Div([
            html.Strong(formato_valor(sum(datos.values()))),
            html.Span("Estudiantes"),
        ], className="vision-card-value-row programa-card-resultado"),
        html.Div([
            html.Div([
                html.Span(nombre),
                html.Strong(formato_valor(valor)),
            ], className="programa-metrica-periodo")
            for nombre, valor in datos.items()
        ], className="reinsercion-desglose"),
        _pie_programa_actualizable("Año lectivo", anio_lectivo),
    ]


def _metricas_aulas_hospitalarias(anio):
    if anio not in AULAS_HOSPITALARIAS_POR_ANIO:
        anio = list(AULAS_HOSPITALARIAS_POR_ANIO)[-1]
    datos = AULAS_HOSPITALARIAS_POR_ANIO[anio]
    return [
        html.Div([
            html.Strong(formato_valor(datos["beneficiarios"])),
            html.Span("Beneficiarios"),
        ], className="programa-estudiantes-principal"),
        html.Div([
            html.Div([
                html.Span("Aulas hospitalarias"),
                html.Strong(formato_valor(datos["aulas"])),
            ], className="programa-metrica-periodo"),
            html.Div([
                html.Span("Docentes"),
                html.Strong(formato_valor(datos["docentes"])),
            ], className="programa-metrica-periodo"),
            html.Div([
                html.Span("Inversión"),
                html.Strong(datos["inversion"]),
            ], className="programa-metrica-periodo inversion"),
        ], className="hospital-desglose"),
        _pie_programa_actualizable("Año", anio),
    ]


def _tarjetas_programas_eje2(datos):
    """Presenta cada programa con el mismo lenguaje visual de las cifras."""
    tarjetas = []
    tarjetas_complementarias = []
    incluye_reinsercion = False
    colores = ["azul", "verde", "amarillo", "morado"]
    for i, dato in enumerate(datos):
        clave_programa = _sin_tildes(dato["programa"])
        ayuda_superior = _definicion_programa_educacion_superior(dato["programa"])
        ayuda_educacion = _definicion_programa_educacion(dato["programa"])
        ayuda_deporte = _definicion_programa_deporte(dato["programa"])
        ayuda_cultura = _definicion_programa_cultura(dato["programa"])
        ayuda_oficial = (ayuda_superior or ayuda_educacion or
                         ayuda_deporte or ayuda_cultura)
        es_dece = clave_programa.startswith("profesionales dece")
        es_comunidades = clave_programa.startswith("comunidades seguras")
        clave_detalle = ("safpi" if "safpi" in clave_programa
                         else "nap" if re.search(r"(^|\s)nap($|\s|\()", clave_programa)
                         else None)
        es_reinsercion = "reinsercion" in clave_programa or "todos al aula" in clave_programa
        incluye_reinsercion = incluye_reinsercion or es_reinsercion
        titulo = ("Profesionales DECE (Apoyo Psicológico)"
                  if es_dece else dato["programa"])
        # Se conserva la corrección de la versión consolidada anterior: DECE
        # representa profesionales contratados, no el tipo genérico de
        # beneficiario registrado en la matriz.
        unidad = ("Comunidad Educativa" if es_comunidades
                  else "DECEs contratados" if es_dece
                  else str(dato["tipo"]).strip())
        ayuda_reinsercion = (
            "Todos al Aula por el Nuevo Ecuador.",
            "Promueve el acceso y/o reinserción escolar de niños, niñas, adolescentes, "
            "jóvenes y adultos que se encuentran fuera del Sistema Nacional de Educación, "
            "a través de acciones focalizadas en territorio y coordinación interinstitucional.",
        )
        contenido = []
        if clave_detalle:
            contenido.extend(_control_anio_programa(clave_detalle))
        elif es_reinsercion:
            contenido.extend([
                _ayuda_programa(*ayuda_reinsercion),
                _control_anio_complementario(
                    "reinsercion", list(REINSERCION_POR_ANIO_LECTIVO), "Año lectivo"
                ),
            ])
        elif ayuda_oficial:
            contenido.append(_ayuda_programa(ayuda_oficial[0], ayuda_oficial[1]))
        texto_ayuda = (f"{ayuda_reinsercion[0]} {ayuda_reinsercion[1]}"
                       if es_reinsercion else
                       f"{ayuda_oficial[0]} {ayuda_oficial[1]}"
                       if ayuda_oficial else None)
        codigo_icono = "reinsercion" if es_reinsercion else "programa"
        contenido.append(cabecera_tarjeta(codigo_icono, titulo, texto_ayuda))
        if clave_detalle:
            periodo_inicial = PROGRAMAS_CON_DETALLE[clave_detalle]["filas"][-1][0]
            contenido.append(html.Div(
                _metricas_programa_por_anio(clave_detalle, periodo_inicial),
                id={"type": "programa-metricas", "index": clave_detalle},
            ))
        elif es_reinsercion:
            periodo_inicial = list(REINSERCION_POR_ANIO_LECTIVO)[-1]
            contenido.append(html.Div(
                _metricas_reinsercion(periodo_inicial),
                id={"type": "programa-complementario-metricas", "index": "reinsercion"},
            ))
        else:
            contenido.append(html.Div([
                html.Strong(formato_valor(dato["beneficiarios"])),
                html.Span(unidad, className="unidad-sin-mayusculas" if es_dece else None),
            ], className="vision-card-value-row programa-card-resultado"))
        if es_dece:
            contenido.append(html.P("La brecha nacional se ha reducido en 45%",
                                    className="programa-card-destacado"))
        tiene_ayuda = bool(clave_detalle or es_reinsercion or ayuda_oficial)
        tarjetas.append(html.Article(
            contenido,
            className=(f"vision-exec-card programa-resumen-card "
                       f"programa-{colores[i % len(colores)]}"
                       f"{' programa-con-ayuda' if tiene_ayuda else ''}"),
            tabIndex=0 if tiene_ayuda else None,
        ))

    if incluye_reinsercion:
        for j, programa in enumerate(PROGRAMAS_GESTION_COMPLEMENTARIOS):
            clave = programa["clave"]
            periodos = list(AULAS_HOSPITALARIAS_POR_ANIO)
            periodo_inicial = periodos[-1]
            tarjetas_complementarias.append(html.Article([
                _ayuda_programa(programa["titulo"] + ".", programa["descripcion"]),
                _control_anio_complementario(clave, periodos, "Año"),
                cabecera_tarjeta("hospital", programa["titulo"]),
                html.Div(
                    _metricas_aulas_hospitalarias(periodo_inicial),
                    id={"type": "programa-complementario-metricas", "index": clave},
                ),
            ], className=("vision-exec-card programa-resumen-card programa-con-ayuda "
                          f"programa-{colores[(len(tarjetas) + j) % len(colores)]}"),
               tabIndex=0))

    estilo_grid = ({"gridTemplateColumns": "repeat(3, minmax(0, 1fr))"}
                   if len(tarjetas) == 3 else None)
    fila_principal = html.Div(
        tarjetas,
        className="vision-exec-grid gestion-educativa-grid programas-servicios-grid",
        style=estilo_grid,
    )
    if not tarjetas_complementarias:
        return fila_principal
    return html.Div([
        fila_principal,
        html.Div(
            tarjetas_complementarias,
            className=("vision-exec-grid gestion-educativa-grid programas-servicios-grid "
                       "programas-complementarios-grid"),
            style={"marginTop": "12px"},
        ),
    ], className="programas-gestion-consolidados")


@app.callback(
    Output({"type": "programa-metricas", "index": MATCH}, "children"),
    Input({"type": "programa-anio", "index": MATCH}, "value"),
    State({"type": "programa-anio", "index": MATCH}, "id"),
)
def actualizar_metricas_programa(anio, identificador):
    """Actualiza SAFPI o NAP sin reconstruir el resto de la página."""
    clave = identificador["index"]
    return _metricas_programa_por_anio(clave, anio)


@app.callback(
    Output({"type": "programa-complementario-metricas", "index": MATCH}, "children"),
    Input({"type": "programa-complementario-anio", "index": MATCH}, "value"),
    State({"type": "programa-complementario-anio", "index": MATCH}, "id"),
)
def actualizar_programa_complementario(periodo, identificador):
    """Actualiza de forma independiente Reinserción o Aulas Hospitalarias."""
    clave = identificador["index"]
    if clave == "reinsercion":
        return _metricas_reinsercion(periodo)
    if clave == "hospital":
        return _metricas_aulas_hospitalarias(periodo)
    return no_update


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
                      html.P("Programas vigentes para ampliar el acceso, reconocer trayectorias "
                             "académicas y fortalecer la formación técnica, tecnológica y universitaria.")])]),
        html.H2("Información de Universidades y Escuelas Politécnicas",
                className="vision-eje-banner"),
        html.Div(
            className="vision-exec-grid gestion-educativa-grid datos-generales-grid "
                      "educacion-superior-cifras-grid universidades-politecnicas-grid",
            children=tarjetas_universidades_politecnicas(),
            style={"gridTemplateColumns": "repeat(3, minmax(0, 1fr))"},
        ),
        html.H2("Información de Institutos Superiores Técnicos Tecnológicos",
                className="vision-eje-banner"),
        html.Div(
            className="vision-exec-grid gestion-educativa-grid datos-generales-grid "
                      "educacion-superior-cifras-grid institutos-tecnicos-grid",
            children=tarjetas_institutos_tecnicos_tecnologicos(),
            style={"gridTemplateColumns": "repeat(3, minmax(0, 1fr))"},
        ),
        html.H2("Programas y servicios de Educación Superior", className="vision-eje-banner"),
        html.Div(
            className="vision-exec-grid gestion-educativa-grid programas-servicios-grid",
            children=tarjetas_becas_educacion_superior(),
            style={"gridTemplateColumns": "repeat(3, minmax(0, 1fr))"},
        ),
    ])


# Información institucional del Viceministerio del Deporte. Corresponde al
# corte de agosto de 2026 entregado para la Visión Ejecutiva.
DATOS_GENERALES_DEPORTE = [
    ("vitalicios", "Exdeportistas vitalicios", 334, "Exdeportistas"),
    ("alto-rendimiento", "Deportistas de alto rendimiento", 401, "Deportistas"),
    ("entrenadores", "Entrenadores", 324, "Entrenadores"),
    ("usuarios-deporte", "Usuarios de deporte", 455632, "Usuarios"),
    ("juegos-juveniles", "Juegos Deportivos Nacionales · Ciclo Juvenil", 2538,
     "Deportistas"),
]

DATOS_PROGRAMAS_GENERALES_DEPORTE = [
    {
        "codigo": "activate", "titulo": "Actívate", "puntos_activos": 492,
        "instructores": 12, "beneficiarios_permanentes": 22030,
        "beneficiarios_eventos": 27370,
    },
    {
        "codigo": "vamos-cancha", "titulo": "Vamos a la Cancha", "puntos_activos": 415,
        "instructores": 9, "beneficiarios_permanentes": 13232,
        "beneficiarios_eventos": 16683,
    },
]


def _tarjeta_programa_general_deporte(dato):
    metricas = [
        ("Puntos activos", dato["puntos_activos"]),
        ("Instructores", dato["instructores"]),
        ("Beneficiarios permanentes", dato["beneficiarios_permanentes"]),
        ("Beneficiarios de eventos masivos", dato["beneficiarios_eventos"]),
    ]
    return html.Article([
        cabecera_tarjeta(dato["codigo"], dato["titulo"]),
        html.Div([
            html.Div([
                html.Span(etiqueta),
                html.Strong(formato_valor(valor)),
            ], className="deporte-programa-metrica")
            for etiqueta, valor in metricas
        ], className="deporte-programa-metricas"),
        html.Small("Corte: agosto de 2026", className="vision-card-periodo"),
    ], className=(f"vision-exec-card deporte-programa-general "
                  f"{dato['codigo']}"))


def seccion_informacion_general_deporte():
    """Cifras generales de Deporte, separadas de sus programas y servicios."""
    tarjetas = [
        tarjeta_gestion_educativa(
            codigo, titulo, None, formato_valor(valor), unidad=unidad,
            periodo="Corte: agosto de 2026",
        )
        for codigo, titulo, valor, unidad in DATOS_GENERALES_DEPORTE
    ]
    return html.Div([
        html.H2("Información general", className="vision-eje-banner"),
        html.Div(
            tarjetas,
            className=("vision-exec-grid gestion-educativa-grid "
                       "datos-generales-grid deporte-informacion-grid"),
        ),
        html.H3("Actívate y Vamos a la Cancha", className="deporte-subseccion-titulo"),
        html.Div(
            [_tarjeta_programa_general_deporte(dato)
             for dato in DATOS_PROGRAMAS_GENERALES_DEPORTE],
            className=("vision-exec-grid gestion-educativa-grid "
                       "deporte-programas-generales-grid"),
        ),
        html.P([
            html.Strong("Fuente: "),
            "Dirección de Alto Rendimiento, Subsecretaría de Servicios del "
            "Sistema Deportivo y reportes de recaudación de usuarios de centros activos.",
        ], className="eje2-source-note deporte-informacion-fuente"),
        html.P([
            html.Strong("Nota: "),
            "Los registros del Viceministerio del Deporte no todos cuentan con "
            "desagregación a nivel cantonal.",
        ], className="eje2-source-note deporte-informacion-nota"),
    ], className="deporte-informacion-general")


# Información general de Cultura. Cada encabezado de la matriz corresponde a
# una tarjeta y sus filas se presentan como desagregaciones internas.
DATOS_INFORMACION_GENERAL_CULTURA = [
    {
        "codigo": "cultura",
        "titulo": "Artistas y gestores culturales",
        "items": [
            ("Artes e innovación", 23990),
            ("Patrimonio cultural", 916),
            ("Memoria social", 479),
        ],
        "periodo": "Corte: julio de 2026",
    },
    {
        "codigo": "organizaciones",
        "titulo": "Organizaciones sociales y empresas culturales",
        "items": [
            ("Organizaciones sociales", 227),
            ("Empresa cultural", 482),
        ],
        "periodo": "Corte: julio de 2026",
    },
    {
        "codigo": "patrimonio-cultural",
        "titulo": "Espacios e infraestructura",
        "items": [
            ("Museos", None),
            ("Archivos", None),
            ("Bibliotecas", None),
            ("Teatro", None),
        ],
        "periodo": "Información pendiente",
    },
    {
        "codigo": "cultura",
        "titulo": "Colectivos",
        "items": [
            ("Gremios", None),
            ("Asociaciones", None),
        ],
        "periodo": "Información pendiente",
    },
]


# Estructura prevista en la matriz de Cultura. Todavía no contiene cifras,
# pero se deja lista para incorporarlas sin rediseñar la página.
PROGRAMAS_SERVICIOS_CULTURA = [
    {
        "codigo": "memoria-social",
        "titulo": "Visitas de repositorios de la memoria social",
        "anios": ["2023", "2024", "2025"],
        "categorias": ["Museos", "Archivos", "Bibliotecas"],
    },
    {
        "codigo": "cultura",
        "titulo": "Beneficiarios de Líneas de Fomento",
        "anios": ["2023", "2024", "2025"],
        "categorias": ["Artes e innovación", "Memoria social", "Patrimonio cultural"],
    },
    {
        "codigo": "cultura",
        "titulo": "Festival Internacional de Artes Vivas de Loja",
        "anios": ["2024", "2025"],
        "categorias": ["Asistentes", "Artistas", "Eventos"],
    },
]


def _tarjeta_resumen_cultura(dato, color):
    """Tarjeta principal con la desagregación de una tabla de Cultura."""
    items = [
        html.Div([
            html.Span(etiqueta),
            html.Strong(formato_valor(valor) if valor is not None else "—"),
        ], className=("programa-metrica-periodo cultura-resumen-item"
                      + (" pendiente" if valor is None else "")))
        for etiqueta, valor in dato["items"]
    ]
    return html.Article([
        cabecera_tarjeta(dato["codigo"], dato["titulo"]),
        html.Div(items, className=f"cultura-resumen-desglose items-{len(items)}"),
        html.Div([
            html.Span(dato["periodo"]),
            html.Span("Actualización: conforme al corte institucional"),
        ], className="educacion-media-periodo-pie"),
    ], className=(f"vision-exec-card programa-resumen-card "
                  f"cultura-resumen-card programa-{color} {dato['codigo']}"))


def _tarjeta_programa_cultura(dato, color):
    """Reserva visual con el mismo patrón de SAFPI/NAP, aún sin valores."""
    anio_inicial = dato["anios"][-1]
    metricas = [
        html.Div([
            html.Span(categoria),
            html.Strong("—"),
        ], className="programa-metrica-periodo cultura-programa-metrica pendiente")
        for categoria in dato["categorias"]
    ]
    return html.Article([
        html.Div([
            html.Span("Año"),
            dcc.Dropdown(
                options=[{"label": anio, "value": anio} for anio in dato["anios"]],
                value=anio_inicial, clearable=False, searchable=False,
                className="programa-anio-dropdown",
            ),
        ], className="programa-anio-control"),
        cabecera_tarjeta(dato["codigo"], dato["titulo"]),
        html.Div(metricas, className="cultura-programa-desglose"),
        html.Div([
            html.Span(f"Año {anio_inicial}"),
            html.Span("Actualización: pendiente de carga"),
        ], className="educacion-media-periodo-pie"),
    ], className=(f"vision-exec-card programa-resumen-card "
                  f"cultura-programa-card programa-{color} {dato['codigo']}"))


def seccion_informacion_general_cultura():
    """Cuatro tablas generales de Cultura convertidas en tarjetas ejecutivas."""
    return html.Div([
        html.H2("Información general", className="vision-eje-banner"),
        html.Div(
            [_tarjeta_resumen_cultura(dato, color)
             for dato, color in zip(
                 DATOS_INFORMACION_GENERAL_CULTURA,
                 ["azul", "verde", "morado", "amarillo"],
             )],
            className=("vision-exec-grid gestion-educativa-grid "
                       "programas-servicios-grid cultura-resumen-grid"),
        ),
        html.P([
            html.Strong("Fuente: "),
            "Registro Único de Artistas y Gestores Culturales.",
        ], className="eje2-source-note cultura-informacion-fuente"),
    ], className="cultura-informacion-general")


def seccion_programas_servicios_cultura():
    """Tres tablas previstas por Cultura, visibles aun cuando falten cifras."""
    return html.Div([
        html.H2("Programas y servicios de Cultura", className="vision-eje-banner"),
        html.P("Espacios preparados para incorporar las cifras institucionales por año.",
               className="eje1-subtitulo"),
        html.Div(
            [_tarjeta_programa_cultura(dato, color)
             for dato, color in zip(
                 PROGRAMAS_SERVICIOS_CULTURA,
                 ["azul", "verde", "morado"],
             )],
            className=("vision-exec-grid gestion-educativa-grid "
                       "programas-servicios-grid cultura-programas-grid"),
        ),
    ], className="cultura-programas-servicios")


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
    programas, aviso = (([], None) if codigo == "cultura"
                        else cargar_programas_viceministerio(eje["vice_excel"]))
    informacion_general = (
        seccion_informacion_general_deporte() if codigo == "deporte"
        else seccion_informacion_general_cultura() if codigo == "cultura"
        else None
    )
    programas_y_servicios = (
        seccion_programas_servicios_cultura() if codigo == "cultura"
        else html.Div([
            html.H2(f"Programas y servicios de {eje['vice_excel']}",
                    className="vision-eje-banner"),
            html.Div(
                _tarjetas_programas_eje2(programas),
                className="eje-programas-fila-unica",
            ) if programas else None,
            html.P(aviso, className="eje2-source-note") if aviso else None,
        ])
    )
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
        informacion_general,
        programas_y_servicios,
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
    "nee_total": 35582,
    "nee_sostenimiento": {
        "Fiscal": 24658,
        "Fiscomisional": 2665,
        "Municipal": 571,
        "Particular": 7688,
    },
    "nee_regimen": {
        "Costa": 19527,
        "Sierra": 16055,
    },
    "escolaridad_inconclusa_total": 242941,
    "escolaridad_inconclusa_sexo": {
        "Hombres": 116326,
        "Mujeres": 126615,
    },
}


# Serie histórica entregada para Instituciones Educativas. El selector abre
# siempre en el último año lectivo y actualiza el total y sus sostenimientos.
INSTITUCIONES_EDUCATIVAS_POR_ANIO = {
    "2023-2024": {
        "Fiscal": 12389, "Fiscomisional": 617,
        "Municipal": 112, "Particular": 3022,
    },
    "2024-2025": {
        "Fiscal": 12327, "Fiscomisional": 606,
        "Municipal": 110, "Particular": 3109,
    },
    "2025-2026": {
        "Fiscal": 12287, "Fiscomisional": 607,
        "Municipal": 108, "Particular": 3213,
    },
}


# Series históricas tomadas de los totales de la matriz entregada. Las seis
# categorías son independientes: sexo, área y régimen escolar. No se presenta
# un cruce Rural/Urbana por Costa/Sierra.
ESTUDIANTES_POR_ANIO_LECTIVO = {
    "2023-2024": {
        "Total": 4202798, "Hombres": 2116518, "Mujeres": 2086280,
        "Rural": 997456, "Urbana": 3205342,
        "Sierra": 1763290, "Costa": 2439508,
    },
    "2024-2025": {
        "Total": 4106819, "Hombres": 2067589, "Mujeres": 2039230,
        "Rural": 941760, "Urbana": 3165059,
        "Sierra": 1727355, "Costa": 2379464,
    },
    "2025-2026": {
        "Total": 4039650, "Hombres": 2034159, "Mujeres": 2005491,
        "Rural": 929157, "Urbana": 3110493,
        "Sierra": 1703367, "Costa": 2336283,
    },
}

DOCENTES_POR_ANIO_LECTIVO = {
    "2023-2024": {
        "Total": 213468, "Hombres": 59257, "Mujeres": 154211,
        "Rural": 57356, "Urbana": 156112,
        "Sierra": 99592, "Costa": 113876,
    },
    "2024-2025": {
        "Total": 215030, "Hombres": 59439, "Mujeres": 155591,
        "Rural": 55633, "Urbana": 159397,
        "Sierra": 99979, "Costa": 115051,
    },
    "2025-2026": {
        "Total": 217693, "Hombres": 60064, "Mujeres": 157629,
        "Rural": 56516, "Urbana": 161177,
        "Sierra": 101094, "Costa": 116599,
    },
}

DATOS_ABIERTOS_MINEDEC_URL = "https://educacion.gob.ec/datos-abiertos-minedec/"


def _pie_actualizacion_educacion_media(anio_lectivo):
    """Período y frecuencia de actualización de las cifras generales."""
    periodo = str(anio_lectivo).replace("-", "–")
    return html.Div([
        html.Span(f"Año lectivo {periodo}"),
        html.Span("Actualización: marzo de cada año"),
    ], className="educacion-media-periodo-pie")


def _contenido_instituciones_educativas(anio_lectivo):
    """Total y desglose por sostenimiento para el año lectivo seleccionado."""
    if anio_lectivo not in INSTITUCIONES_EDUCATIVAS_POR_ANIO:
        anio_lectivo = list(INSTITUCIONES_EDUCATIVAS_POR_ANIO)[-1]
    sostenimientos = INSTITUCIONES_EDUCATIVAS_POR_ANIO[anio_lectivo]
    total = sum(sostenimientos.values())
    return [
        html.Div([
            html.Strong(formato_valor(total), className="vision-card-total"),
        ], className="vision-card-value-row"),
        html.Div([
            html.Div([
                html.Span(nombre),
                html.Strong(formato_valor(valor)),
            ], className=("vision-mini-stat secondary" if i % 2 else "vision-mini-stat"))
            for i, (nombre, valor) in enumerate(sostenimientos.items())
        ], className="vision-card-details instituciones-sostenimientos"),
        _pie_actualizacion_educacion_media(anio_lectivo),
    ]


def tarjeta_instituciones_educativas():
    """Tarjeta histórica, con el mismo control compacto usado por SAFPI."""
    periodos = list(INSTITUCIONES_EDUCATIVAS_POR_ANIO)
    periodo_reciente = periodos[-1]
    return html.Article([
        html.Div([
            html.Span("Año lectivo"),
            dcc.Dropdown(
                id="instituciones-anio-lectivo",
                options=[{"label": periodo.replace("-", "–"), "value": periodo}
                         for periodo in periodos],
                value=periodo_reciente,
                clearable=False,
                searchable=False,
                className="programa-anio-dropdown",
            ),
        ], className="programa-anio-control instituciones-anio-control"),
        cabecera_tarjeta("institucion", "Instituciones Educativas (IE)"),
        html.Div(
            _contenido_instituciones_educativas(periodo_reciente),
            id="instituciones-contenido",
        ),
    ], className="vision-exec-card generales instituciones-historicas-card")


@app.callback(
    Output("instituciones-contenido", "children"),
    Input("instituciones-anio-lectivo", "value"),
)
def actualizar_instituciones_educativas(anio_lectivo):
    """Actualiza total y sostenimientos sin reconstruir Visión Ejecutiva."""
    return _contenido_instituciones_educativas(anio_lectivo)


def _contenido_estudiantes(anio_lectivo):
    """Total y desgloses independientes por sexo, área y régimen escolar."""
    if anio_lectivo not in ESTUDIANTES_POR_ANIO_LECTIVO:
        anio_lectivo = list(ESTUDIANTES_POR_ANIO_LECTIVO)[-1]
    datos = ESTUDIANTES_POR_ANIO_LECTIVO[anio_lectivo]
    categorias = ("Hombres", "Mujeres", "Rural", "Urbana", "Sierra", "Costa")

    return [
        html.Div([
            html.Strong(formato_valor(datos["Total"]), className="vision-card-total"),
        ], className="vision-card-value-row"),
        html.Div([
            html.Div([
                html.Span(categoria),
                html.Strong(formato_valor(datos[categoria])),
            ], className=("vision-mini-stat secondary" if i % 2 else "vision-mini-stat"))
            for i, categoria in enumerate(categorias)
        ], className="vision-card-details desglose-demografico"),
        _pie_actualizacion_educacion_media(anio_lectivo),
    ]


def tarjeta_estudiantes():
    """Tarjeta histórica basada solo en Suma de Total_Estudiantes."""
    periodos = list(ESTUDIANTES_POR_ANIO_LECTIVO)
    periodo_reciente = periodos[-1]
    return html.Article([
        html.Div([
            html.Span("Año lectivo"),
            dcc.Dropdown(
                id="estudiantes-anio-lectivo",
                options=[{"label": periodo.replace("-", "–"), "value": periodo}
                         for periodo in periodos],
                value=periodo_reciente,
                clearable=False,
                searchable=False,
                className="programa-anio-dropdown",
            ),
        ], className="programa-anio-control estudiantes-anio-control"),
        cabecera_tarjeta("estudiantes", "Estudiantes"),
        html.Div(
            _contenido_estudiantes(periodo_reciente),
            id="estudiantes-contenido",
        ),
    ], className="vision-exec-card generales estudiantes-historicos-card")


@app.callback(
    Output("estudiantes-contenido", "children"),
    Input("estudiantes-anio-lectivo", "value"),
)
def actualizar_estudiantes(anio_lectivo):
    """Actualiza el total y los desgloses independientes de estudiantes."""
    return _contenido_estudiantes(anio_lectivo)


def _contenido_docentes(anio_lectivo):
    """Total y desgloses independientes de docentes para el período elegido."""
    if anio_lectivo not in DOCENTES_POR_ANIO_LECTIVO:
        anio_lectivo = list(DOCENTES_POR_ANIO_LECTIVO)[-1]
    datos = DOCENTES_POR_ANIO_LECTIVO[anio_lectivo]
    categorias = ("Hombres", "Mujeres", "Rural", "Urbana", "Sierra", "Costa")
    return [
        html.Div([
            html.Strong(formato_valor(datos["Total"]), className="vision-card-total"),
        ], className="vision-card-value-row"),
        html.Div([
            html.Div([
                html.Span(categoria),
                html.Strong(formato_valor(datos[categoria])),
            ], className=("vision-mini-stat secondary" if i % 2 else "vision-mini-stat"))
            for i, categoria in enumerate(categorias)
        ], className="vision-card-details desglose-demografico"),
        _pie_actualizacion_educacion_media(anio_lectivo),
    ]


def tarjeta_docentes():
    """Tarjeta histórica de docentes por sexo, área y régimen escolar."""
    periodos = list(DOCENTES_POR_ANIO_LECTIVO)
    periodo_reciente = periodos[-1]
    return html.Article([
        html.Div([
            html.Span("Año lectivo"),
            dcc.Dropdown(
                id="docentes-anio-lectivo",
                options=[{"label": periodo.replace("-", "–"), "value": periodo}
                         for periodo in periodos],
                value=periodo_reciente,
                clearable=False,
                searchable=False,
                className="programa-anio-dropdown",
            ),
        ], className="programa-anio-control docentes-anio-control"),
        cabecera_tarjeta("docentes", "Docentes"),
        html.Div(
            _contenido_docentes(periodo_reciente),
            id="docentes-contenido",
        ),
    ], className="vision-exec-card generales docentes-historicos-card")


@app.callback(
    Output("docentes-contenido", "children"),
    Input("docentes-anio-lectivo", "value"),
)
def actualizar_docentes(anio_lectivo):
    """Actualiza el total y los desgloses independientes de docentes."""
    return _contenido_docentes(anio_lectivo)


def tarjetas_datos_generales():
    """Cifras generales de Educación Media y escolaridad inconclusa."""
    tarjetas = [
        tarjeta_instituciones_educativas(),
        tarjeta_estudiantes(),
        tarjeta_docentes(),
        tarjeta_escolaridad_inconclusa(),
    ]
    return tarjetas


def tarjeta_escolaridad_inconclusa():
    """Total de estudiantes con escolaridad inconclusa desagregado por sexo."""
    d = DATOS_GENERALES_GESTION_EDUCATIVA
    detalles = [
        html.Div([
            html.Span(nombre),
            html.Strong(formato_valor(valor)),
        ], className=("vision-mini-stat secondary" if i % 2 else "vision-mini-stat"))
        for i, (nombre, valor) in enumerate(d["escolaridad_inconclusa_sexo"].items())
    ]
    return html.Article([
        cabecera_tarjeta("epja", "Escolaridad Inconclusa"),
        html.Div([
            html.Strong(
                formato_valor(d["escolaridad_inconclusa_total"]),
                className="vision-card-total",
            ),
            html.Span("Estudiantes", className="vision-card-unit"),
        ], className="vision-card-value-row"),
        html.Div(detalles, className="vision-card-details"),
        _pie_actualizacion_educacion_media(d["anio_lectivo"]),
    ], className="vision-exec-card epja escolaridad-inconclusa-card")


def tarjeta_necesidades_educativas_especiales():
    """Estudiantes con NEE por sostenimiento y régimen, en valores absolutos."""
    d = DATOS_GENERALES_GESTION_EDUCATIVA
    detalles = [html.Span("Sostenimiento", className="nee-grupo-titulo")]
    detalles.extend(
        html.Div([
            html.Span(nombre),
            html.Strong(formato_valor(valor)),
        ], className=("vision-mini-stat secondary" if i % 2 else "vision-mini-stat"))
        for i, (nombre, valor) in enumerate(d["nee_sostenimiento"].items())
    )
    detalles.append(html.Span("Régimen", className="nee-grupo-titulo"))
    detalles.extend(
        html.Div([
            html.Span(nombre),
            html.Strong(formato_valor(valor)),
        ], className=("vision-mini-stat secondary" if i % 2 else "vision-mini-stat"))
        for i, (nombre, valor) in enumerate(d["nee_regimen"].items())
    )
    return html.Article([
        cabecera_tarjeta("generales", "Estudiantes con necesidades educativas especiales"),
        html.Div([
            html.Strong(formato_valor(d["nee_total"]), className="vision-card-total"),
            html.Span("Estudiantes", className="vision-card-unit"),
        ], className="vision-card-value-row"),
        html.Div(detalles, className="nee-desglose"),
        _pie_actualizacion_educacion_media(d["anio_lectivo"]),
    ], className="vision-exec-card generales nee-card")


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


RECURSOS_EDUCATIVOS_ACTUALIZADOS = [
    {
        "codigo": "alimentacion", "titulo": "Alimentación Escolar",
        "nacional": {"instituciones": 12493, "estudiantes": 2493993,
                     "inversion": "$190.476.720,79"},
        "regimen": {
            "Costa": {"instituciones": 7602, "estudiantes": 1436502,
                      "inversion": "$109.711.691,40"},
            "Sierra": {"instituciones": 4891, "estudiantes": 1057491,
                       "inversion": "$80.765.029,39"},
        },
    },
    {
        "codigo": "uniformes", "titulo": "Uniformes Escolares",
        "nacional": {"instituciones": 10136, "estudiantes": 1291398,
                     "inversion": "$42.845.083,30"},
        "regimen": {
            "Costa": {"instituciones": 5608, "estudiantes": 664234,
                      "inversion": "$21.974.683,86"},
            "Sierra": {"instituciones": 4528, "estudiantes": 627164,
                       "inversion": "$20.870.399,44"},
        },
    },
    {
        "codigo": "textos", "titulo": "Textos Escolares",
        "nacional": {"instituciones": 12603, "estudiantes": 2859894,
                     "inversion": "$35.156.870,66"},
        "regimen": {
            "Costa": {"instituciones": 6939, "estudiantes": 1599739,
                      "inversion": "$18.890.301,65"},
            "Sierra": {"instituciones": 5664, "estudiantes": 1260155,
                       "inversion": "$16.266.569,01"},
        },
    },
    {
        "codigo": "mobiliario", "titulo": "Equipamiento y Mobiliario",
        "nacional": {"instituciones": 1413, "estudiantes": 307800,
                     "inversion": "$6.991.867,72"},
        "regimen": {
            "Costa": {"instituciones": 906, "estudiantes": 198640,
                      "inversion": "$4.831.279,59"},
            "Sierra": {"instituciones": 507, "estudiantes": 109160,
                       "inversion": "$2.160.588,13"},
        },
    },
    {
        "codigo": "transporte", "titulo": "Transporte Escolar",
        "nacional": {"instituciones": 167, "estudiantes": 23525,
                     "inversion": "$13.738.150,61"},
        "regimen": {
            "Costa": {"instituciones": 63, "estudiantes": 9761,
                      "inversion": "$5.091.583,64"},
            "Sierra": {"instituciones": 104, "estudiantes": 13764,
                       "inversion": "$8.646.566,97"},
        },
    },
]

DEFINICIONES_RECURSOS_GESTION = {
    "alimentacion": (
        "Alimentación General.",
        "Se entrega a los estudiantes una ración alimenticia diaria, compuesta "
        "por una galleta y leche, durante 200 días.",
    ),
    "uniformes": (
        "Uniformes.",
        "Entrega gratuita de kits de uniformes escolares a estudiantes de "
        "instituciones educativas fiscales de zonas rurales y vulnerables, "
        "para facilitar su acceso a la educación.",
    ),
    "textos": (
        "Textos.",
        "Se entregan libros de Lenguaje, Matemáticas y Ciencias a estudiantes "
        "de instituciones educativas fiscales.",
    ),
    "mobiliario": (
        "Mobiliario Escolar.",
        "Comprende la dotación de bienes destinados a equipar las instituciones "
        "educativas, como mesas, sillas, armarios, pizarras y kits de mobiliario escolar.",
    ),
    "transporte": (
        "Transporte Escolar.",
        "Servicio personalizado que moviliza a estudiantes desde sus hogares o "
        "paradas hasta las instituciones educativas y viceversa, de acuerdo con "
        "sus necesidades.",
    ),
}


def tarjeta_recurso_educativo(dato):
    """Resumen nacional y desglose Costa/Sierra de un recurso educativo, en
    el mismo formato compacto (cuadros Costa/Sierra) que usan SAFPI y NAP:
    el número principal de Estudiantes con los dos cuadros de régimen al
    lado, y Instituciones/Inversión nacional debajo."""
    codigo = dato["codigo"]
    nacional = dato["nacional"]
    regimen = dato["regimen"]
    ayuda = DEFINICIONES_RECURSOS_GESTION.get(codigo, (None, None))

    estudiantes_principal = html.Div([
        html.Strong(formato_valor(nacional["estudiantes"])),
        html.Span("Estudiantes"),
    ], className="programa-estudiantes-principal")

    fila_estudiantes = html.Div([
        estudiantes_principal,
        *[
            html.Div([
                html.Span(nombre_regimen),
                html.Strong(formato_valor(valores["estudiantes"])),
            ], className="programa-regimen-mini")
            for nombre_regimen, valores in regimen.items()
        ],
    ], className="programa-estudiantes-regimen-fila")

    return html.Article([
        _ayuda_tarjeta(ayuda[0], ayuda[1]) if all(ayuda) else None,
        cabecera_tarjeta(
            codigo, dato["titulo"],
            f"{ayuda[0]} {ayuda[1]}" if all(ayuda) else None,
        ),
        html.Div([
            fila_estudiantes,
            html.Div([
                html.Div([
                    html.Span("Instituciones"),
                    html.Strong(formato_valor(nacional["instituciones"])),
                ], className="programa-metrica-periodo"),
                html.Div([
                    html.Span("Inversión"),
                    html.Strong(nacional["inversion"]),
                ], className="programa-metrica-periodo inversion"),
            ], className="programa-desglose-inferior"),
        ], className="programa-metricas-priorizadas"),
        html.Div([
            html.Span("Año lectivo 2025–2026"),
            html.Span("Actualización: diciembre de cada año"),
        ], className="educacion-media-periodo-pie"),
    ], className=(f"vision-exec-card {codigo} programa-con-ayuda "
                  "recurso-educativo-actualizado"), tabIndex=0)


def pagina_gestion_educativa():
    """Visión consolidada de Educación Media y sus programas institucionales."""
    programas_gestion, aviso_programas = cargar_programas_viceministerio("Gestión Educativa")
    programas_gestion = [
        p for p in programas_gestion
        if "epja" not in _sin_tildes(p.get("programa", ""))
        and not (
            "alimentacion escolar" in _sin_tildes(p.get("programa", ""))
            and "bienestar" in _sin_tildes(p.get("programa", ""))
        )
    ]
    tarjetas_media = tarjetas_datos_generales()
    tarjetas_media_principales = tarjetas_media[:3]
    # Orden visual solicitado: primero Necesidades Educativas Especiales y,
    # a continuación, Escolaridad Inconclusa.
    tarjetas_media_segunda_fila = [
        tarjeta_necesidades_educativas_especiales()
    ] + tarjetas_media[3:]

    tarjetas = [tarjeta_recurso_educativo(dato)
                for dato in RECURSOS_EDUCATIVOS_ACTUALIZADOS]

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
                html.P(""),
            ]),
        ]),

        html.H2("Información de Educación Media", className="vision-eje-banner"),
        html.Div(
                 className="vision-exec-grid gestion-educativa-grid datos-generales-grid "
                           "educacion-media-cifras-grid",
                 children=tarjetas_media_principales,
                 style={"gridTemplateColumns": "repeat(3, minmax(0, 1fr))"}),
        html.Div(
            className="vision-exec-grid gestion-educativa-grid datos-generales-grid "
                      "educacion-media-segunda-fila",
            children=tarjetas_media_segunda_fila,
            style={"gridTemplateColumns": "repeat(3, minmax(0, 1fr))", "marginTop": "12px"},
        ),

        html.H2("Programas y servicios de Gestión Educativa", className="vision-eje-banner"),
        html.P("Cifras consolidadas de los programas y servicios institucionales.",
               className="eje1-subtitulo"),
        _tarjetas_programas_eje2(programas_gestion) if programas_gestion else None,
        html.P(aviso_programas, className="eje2-source-note") if aviso_programas else None,

        html.H2("Recursos educativos y complementarios",
                className="vision-eje-banner"),
        html.Div(
            className="vision-exec-grid gestion-educativa-grid recursos-gestion-grid",
            children=tarjetas,
            style={"gridTemplateColumns": "repeat(3, minmax(0, 1fr))"},
        ),
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
            html.Div([
                html.P("VISIÓN EJECUTIVA", className="content-kicker"),
                html.H1(
                    "MINISTERIO DE EDUCACIÓN, DEPORTE Y CULTURA",
                    className="vision-landing-title",
                    style={
                        "margin": "0", "color": "#4e3cab",
                        "fontSize": "clamp(25px, 2.35vw, 40px)",
                        "fontWeight": "900", "lineHeight": "1.08",
                        "letterSpacing": ".01em", "textAlign": "center",
                    },
                ),
                html.P(
                    "Seleccione un viceministerio para consultar sus cifras, programas y servicios.",
                    className="vision-landing-subtitle",
                ),
            ], className="vision-landing-intro institutional-landing-header"),
            html.Div(menu, className="vision-landing-menu"),
        ], className="vision-landing",
           style={"minHeight": "calc(100vh - 92px)", "padding": "18px 28px 34px",
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
        marker={"colors": ["#232D5A", "#E5EDF0"]},
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
    # Una sola familia cromática institucional: azul principal, azul secundario
    # y turquesa de apoyo. El color ya no compite con el significado del monto.
    metricas = [
        ("Codificado", "CODIFICADO", "#232D5A"),
        ("Certificado", "PRECOMPROMISO", "#80BBBF"),
        ("Comprometido", "COMPROMISO", "#446381"),
        ("Devengado", "DEVENGADO", "#232D5A"),
        ("Saldo disponible", "SALDO_DISPONIBLE", "#80BBBF"),
    ]
    etiquetas = [nombre for nombre, _, _ in metricas]
    colores = [color for _, _, color in metricas]
    valores = [
        (float(df[col].sum()) / 1_000_000) if (not df.empty and col in df.columns) else 0.0
        for _, col, _ in metricas
    ]
    montos = [v * 1_000_000 for v in valores]
    figura = go.Figure(go.Bar(
        x=etiquetas, y=valores,
        marker={"color": colores, "line": {"color": "#232D5A", "width": .7}},
        width=.46,
        text=[f"{_millones(v)} M" for v in valores], textposition="outside",
        customdata=[_moneda(m) for m in montos],
        hovertemplate="<b>%{x}</b><br>%{customdata}<extra></extra>",
    ))
    figura.update_layout(
        title={"text": "Montos presupuestarios (millones de USD) · Inversión", "x": .5,
               "font": {"size": 16, "color": "#17245b"}},
        template="plotly_white", height=300, margin=dict(l=55, r=25, t=55, b=45),
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
                    "borderRadius": "7px", "background": "#EDF4F6",
                }),
                html.Div([
                    html.Span(titulo, style={
                        "display": "block", "fontSize": "8.5px", "fontWeight": "800",
                        "textTransform": "uppercase", "letterSpacing": ".02em",
                        "color": "#446381", "whiteSpace": "normal",
                    }),
                    html.Strong(valor, style={
                        "display": "block", "fontSize": "clamp(11px, .95vw, 14px)",
                        "color": "#232d5a", "margin": "2px 0 1px", "overflowWrap": "anywhere",
                    }),
                    html.Small(nota, style={
                        "display": "block", "fontSize": "8px", "color": "#667085",
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
                 className="content-heading institutional-landing-header"),
        html.Div([
            html.A(nombre, href=url, target="_blank", rel="noopener noreferrer",
                   className="institutional-access-card")
            for nombre, url in enlaces
        ], className="institutional-access-grid"),
    ], className="indicator-content")


def contenido_documentacion():
    """Repositorio documental exigido para el entregable E-02 (DP-SI-020)."""
    icono_documento = (
        "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 48 48' "
        "fill='none' stroke='#446381' stroke-width='2.6' "
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
        tarjeta_entregable(
            "E-06", "Ficha de formalización",
            "Anexo 4 con la ficha institucional para la formalización de la solución de información.",
            "DOCX", "E-06_Anexo4_Ficha_formalizacion.docx",
        ),
    ]
    return html.Div([
        html.Section([
            html.Div([
                html.P("DOCUMENTACIÓN", className="documentation-kicker"),
                html.H1("Documentación"),
                html.P("Manual de usuario, diccionario de datos e indicadores, fichas metodológicas y ficha de formalización de la solución de información."),
            ]),
        ], className="documentation-hero institutional-landing-header"),
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
    return html.Div([
                     html.Button([
                         html.Span("☰", className="mobile-menu-icon", **{"aria-hidden": "true"}),
                         html.Span("Menú", className="mobile-menu-label"),
                     ], id="mobile-menu-button", className="mobile-menu-button",
                        n_clicks=0, type="button", title="Abrir menú de navegación",
                        **{"aria-label": "Abrir menú de navegación", "aria-expanded": "false"}),
                     html.Button("", id="mobile-menu-overlay", className="mobile-menu-overlay",
                                 n_clicks=0, type="button", title="Cerrar menú",
                                 **{"aria-label": "Cerrar menú de navegación"}),
                     menu_lateral(seccion, vice, indicador),
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

    # Descarga nuevamente las tres matrices cada cinco minutos y compara sus
    # hashes. Si cambia alguna, el Store provoca la reconstrucción del gráfico.
    if _intervalo % 20 == 0:
        nueva_data, nuevo_error = cargar_base_pnd()
        nueva_version_pnd = version_pnd()
        if (not nuevo_error and not nueva_data.empty
                and nueva_version_pnd != version_actual.get("pnd")):
            DATA_PND, ERROR_PND, VERSION_PND = nueva_data, None, nueva_version_pnd
            version_actual["pnd"] = nueva_version_pnd
            cambio = True

        nueva_data, nuevo_error = cargar_base_kpi()
        nueva_version_kpi = version_kpi()
        if (not nuevo_error and not nueva_data.empty
                and nueva_version_kpi != version_actual.get("kpi-estrategicos")):
            DATA_KPI, ERROR_KPI, VERSION_KPI = nueva_data, None, nueva_version_kpi
            version_actual["kpi-estrategicos"] = nueva_version_kpi
            cambio = True

        nueva_data, nuevo_error = cargar_base_kpi_inst()
        nueva_version_kpi_inst = version_kpi_inst()
        if (not nuevo_error and not nueva_data.empty
                and nueva_version_kpi_inst != version_actual.get("kpi-institucionales")):
            DATA_KPI_INST, ERROR_KPI_INST, VERSION_KPI_INST = nueva_data, None, nueva_version_kpi_inst
            version_actual["kpi-institucionales"] = nueva_version_kpi_inst
            cambio = True

    # Las bases de Gestión Educativa (algunas en OneDrive) se vuelven a leer
    # cada 15 minutos, igual que el presupuesto: no hay forma barata de saber
    # si cambiaron sin descargarlas, así que se relee y se compara después.
    if _intervalo % 20 == 0:
        nuevo_resumen, nuevos_pendientes = cargar_resumen_gestion_educativa()
        nueva_version_gestion_educ = version_gestion_educativa(nuevo_resumen, nuevos_pendientes)
        if nueva_version_gestion_educ != version_actual.get("gestion-educativa"):
            RESUMEN_GESTION_EDUCATIVA, PENDIENTES_GESTION_EDUCATIVA = nuevo_resumen, nuevos_pendientes
            VERSION_GESTION_EDUCATIVA = nueva_version_gestion_educ
            version_actual["gestion-educativa"] = nueva_version_gestion_educ
            cambio = True

    # SharePoint se consulta cada 15 minutos para evitar solicitudes innecesarias.
    if _intervalo % 20 == 0:
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
