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
from dash import ALL, Dash, Input, Output, State, ctx, dash_table, dcc, html, no_update

BASE_DIR = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()
SECCIONES = [
    ("vision", "Visión Ejecutiva"),
    ("pnd", "Plan Nacional de Desarrollo"),
    ("kpi-estrategicos", "KPI´s Estratégicos"),
    ("kpi-institucionales", "KPI´s Institucionales"),
    ("presupuesto", "Ejecución Presupuestaria"),
    ("inventario", "Inventario y Recurso de Información"),
]

# Iconografía lineal, monocromática y minimalista según el manual técnico.
ICONOS = {
    "vision": "◎",
    "pnd": "⌖",
    "kpi-estrategicos": "◇",
    "kpi-institucionales": "◫",
    "presupuesto": "▥",
    "inventario": "▤",
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


def _descargar_esigef_actual():
    """Descarga el archivo presupuestario vigente desde OneDrive.

    Reproduce la estrategia usada en Shiny/httr: sigue redirecciones y evita
    reutilizar una copia almacenada en caché.
    """
    respuesta = requests.get(
        ESIGEF_DOWNLOAD_URL,
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
            "El enlace de ESIGEF_actual.xlsx no devolvió un archivo Excel. "
            "Revise que el vínculo permita descargar el archivo."
        )

    version = hashlib.sha256(contenido).hexdigest()
    return BytesIO(contenido), {
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
DATA_VISION, ERROR_VISION = cargar_base_vision()
DATA_PRESUPUESTO, ERROR_PRESUPUESTO = cargar_base_presupuesto()
VERSION_PND = version_pnd()
VERSION_KPI = version_kpi()
VERSION_KPI_INST = version_kpi_inst()
VERSION_VISION = version_vision()
VERSION_PRESUPUESTO = version_presupuesto()

# Metadatos de las secciones que sí tienen datos tabulares (menú lateral con acordeón).
SECCIONES_CON_DATOS = {
    "vision": {"vice_col": "Sección", "indicador_col": "Tabla"},
    "pnd": {"vice_col": "VICEMINISTERIO", "indicador_col": "NOMBRE DEL INDICADOR"},
    "kpi-estrategicos": {"vice_col": "VICEMINISTERIO", "indicador_col": "NOMBRE DEL INDICADOR"},
    "kpi-institucionales": {"vice_col": "VICEMINISTERIO", "indicador_col": "NOMBRE DEL INDICADOR"},
    "presupuesto": {"vice_col": "Viceministerio", "indicador_col": "Viceministerio"},
}


def obtener_datos(seccion):
    if seccion == "vision":
        return DATA_VISION, ERROR_VISION
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
    def tarjeta(codigo, nombre, numero):
        return html.Button([
            html.Span(f"{numero:02d}", className="card-number"),
            html.Span(nombre, className="card-label"), html.Span("→", className="card-arrow")
        ], id={"type": "home-card", "index": codigo}, className=f"section-card card-{numero}", n_clicks=0)

    return html.Main(html.Section([
        html.Div([html.P("PANEL INSTITUCIONAL", className="eyebrow"),
                  html.H1("PANEL INTEGRADO DE GESTIÓN MINEDEC")], className="title-banner"),
        html.P("Acceso integrado a la información estratégica, institucional y presupuestaria del Ministerio.",
               className="intro-text"),
        html.Div([tarjeta(c, n, i) for i, (c, n) in enumerate(SECCIONES, 1)], className="section-grid"),
        html.Footer([html.Div(className="footer-line"),
                     html.Img(src=app.get_asset_url("logo-minedec.png"), className="footer-logo")],
                    className="page-footer"),
    ], className="dashboard-card"), className="page-content")


def menu_lateral(seccion_activa=None, vice_activo=None, indicador_activo=None):
    items = []
    for codigo, nombre in SECCIONES:
        activo = codigo == seccion_activa
        items.append(html.Button([
            html.Span(ICONOS[codigo], className="side-icon"), html.Span(nombre),
            html.Span("⌄" if activo else "", className="side-chevron")
        ], id={"type": "side-section", "index": codigo}, n_clicks=0,
           className="side-section active" if activo else "side-section"))
        if activo and codigo in SECCIONES_CON_DATOS:
            df, _ = obtener_datos(codigo)
            cfg = SECCIONES_CON_DATOS[codigo]
            if not df.empty:
                if codigo == "vision":
                    opciones_vice = (df[[cfg["vice_col"], "Orden sección"]].drop_duplicates()
                                     .sort_values("Orden sección")[cfg["vice_col"]].tolist())
                else:
                    opciones_vice = sorted(df[cfg["vice_col"]].dropna().unique())
                    if codigo == "presupuesto":
                        opciones_vice = sorted(
                            opciones_vice,
                            key=lambda x: (
                                str(x).strip().lower() == "coordinaciones generales",
                                str(x).lower(),
                            ),
                        )
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
    """Tarjeta fija, estilo documento, para fórmulas LaTeX extensas."""
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

    return html.Div([
        html.Div([
            html.Span("∑", className="formula-accordion-icon"),
            html.Strong(etiqueta),
        ], className="formula-static-header"),
        html.Div(bloques, className="formula-accordion-body formula-content"),
    ], className=f"detail-item {clase} formula-static".strip())


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
        html.Section(formula_documento, className="formula-document") if formula_documento else None,
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
    claves_periodo = list(grupo["Periodo_id"])
    periodos = list(grupo["Periodo_grafico"])
    periodos_completos = list(grupo["Periodo"])

    fig = go.Figure()
    fig.add_bar(x=claves_periodo, y=grupo["Línea base"], name="Línea base",
                marker=dict(color="#80BBBF", line=dict(color="#446381", width=1.4)),
                text=[formato_texto_barra(v, unidad) for v in grupo["Línea base"]],
                textposition="outside", textfont=dict(size=10, color="#232D5A", family="Arial"), constraintext="none",
                width=.34, offset=-.17, cliponaxis=False,
                customdata=periodos_completos,
                hovertemplate="%{customdata}<br>Línea base: %{y:,.2f}<extra></extra>")
    fig.add_bar(x=claves_periodo, y=grupo["Meta"], name="Meta",
                marker=dict(color="#F1B620", line=dict(color="#C88F00", width=1.4)),
                text=[formato_texto_barra(v, unidad) for v in grupo["Meta"]],
                textposition="outside", textfont=dict(size=10, color="#232D5A", family="Arial"),
                constraintext="none", width=.34, offset=-.36, cliponaxis=False,
                customdata=periodos_completos,
                hovertemplate="%{customdata}<br>Meta: %{y:,.2f}<extra></extra>")
    fig.add_bar(x=claves_periodo, y=grupo["Estimador"], name="Ejecutado",
                marker=dict(color="#4F449A", line=dict(color="#232D5A", width=1.4)),
                text=[formato_texto_barra(v, unidad) for v in grupo["Estimador"]],
                textposition="outside", textfont=dict(size=10, color="#232D5A", family="Arial"),
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
                   automargin=True, tickfont=dict(size=10)),
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
        html.Section(formula_documento, className="formula-document") if formula_documento else None,
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
                     html.P("Seleccione un viceministerio y luego el indicador que desea consultar.")],
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
                     html.P("Seleccione un viceministerio y luego el indicador que desea consultar.")],
                    className="module-welcome")


# ---------------------------------------------------------------------------
# Visión Ejecutiva · tablas actualizables desde Excel
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


def contenido_vision(seccion=None):
    if ERROR_VISION:
        return html.Div([html.H2("Base no disponible"), html.P(ERROR_VISION)], className="load-error")
    if DATA_VISION.empty:
        return html.Div([html.H2("Sin información"),
                         html.P("El archivo de Visión Ejecutiva no contiene registros.")],
                        className="load-error")

    tablas = []
    if seccion:
        datos_seccion = DATA_VISION.loc[DATA_VISION["Sección"] == seccion]
        tabla_orden = (datos_seccion[["Tabla", "Orden tabla"]].drop_duplicates()
                       .sort_values("Orden tabla"))
        for _, registro_tabla in tabla_orden.iterrows():
            nombre_tabla = registro_tabla["Tabla"]
            tablas.append(tarjeta_tabla_vision(
                datos_seccion.loc[datos_seccion["Tabla"] == nombre_tabla]
            ))

    introduccion_seccion = (html.P(
        "La información corresponde a las instituciones educativas, estudiantes y docentes "
        "que forman parte de la oferta educativa desde Educación Inicial hasta Tercero de Bachillerato.",
        className="vision-section-intro"
    ) if seccion == "Educación Media" else None)

    cuerpo = (html.Div([
                  introduccion_seccion,
                  html.Div(tablas, className="vision-tables-stack")
              ], className="vision-section-body vision-selected-section")
              if tablas else html.Div([
                  html.H2("Seleccione una sección"),
                  html.P("Use el menú lateral para consultar las tablas de Educación Media, "
                         "Educación Superior, Cultura o Viceministerio de Deporte.")
              ], className="module-welcome vision-welcome"))

    return html.Div([
        html.Div([html.P("VISIÓN EJECUTIVA", className="content-kicker"),
                  html.H1(seccion or "Comunidad Educativa"),
                  html.P("Información consolidada de educación, cultura y deporte.")],
                 className="vision-banner"),
        cuerpo,
    ], className="vision-content")


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

    grupo = DATA_PRESUPUESTO.loc[DATA_PRESUPUESTO["Viceministerio"] == vice].copy()
    totales = grupo[PRESUPUESTO_MONETARIAS].sum()
    codificado = float(totales["CODIFICADO"])
    devengado = float(totales["DEVENGADO"])
    ejecucion = (devengado / codificado * 100) if codificado else 0.0

    kpis = [
        ("Codificado", _moneda(codificado), "Presupuesto vigente"),
        ("Compromiso", _moneda(totales["COMPROMISO"]), "Obligaciones comprometidas"),
        ("Devengado", _moneda(devengado), f"{ejecucion:.2f}% de ejecución".replace(".", ",")),
        ("Pagado", _moneda(totales["PAGADO"]), "Monto pagado"),
        ("Saldo disponible", _moneda(totales["SALDO_DISPONIBLE"]), "Disponibilidad registrada"),
    ]

    pendiente = max(100.0 - ejecucion, 0.0)
    fig_avance = go.Figure(go.Pie(
        labels=["Ejecutado (devengado)", "Pendiente por ejecutar"],
        values=[ejecucion, pendiente], hole=0.68, sort=False,
        marker={"colors": ["#55479d", "#e8eaf3"], "line": {"color": "white", "width": 2}},
        textinfo="none",
        hovertemplate="%{label}: %{value:.2f}%<extra></extra>",
    ))
    fig_avance.add_annotation(
        text=f"<b>{ejecucion:.1f}%</b><br><span style='font-size:12px'>ejecutado</span>",
        x=0.5, y=0.5, showarrow=False, font={"size": 25, "color": "#25335f"}
    )
    fig_avance.update_layout(
        height=310, margin=dict(l=25, r=25, t=58, b=30), paper_bgcolor="white",
        title={"text": "Avance de ejecución presupuestaria", "x": 0.5, "xanchor": "center"},
        legend={"orientation": "h", "y": -0.06, "x": 0.5, "xanchor": "center"},
    )

    dimension = "TIPO" if "TIPO" in grupo.columns else "NOM_GRUPO"
    resumen_tipo = (grupo.groupby(dimension, dropna=False)[["CODIFICADO", "DEVENGADO"]]
                    .sum().reset_index())
    resumen_tipo[dimension] = resumen_tipo[dimension].fillna("Sin clasificación").astype(str)
    resumen_tipo["EJECUCION"] = resumen_tipo.apply(
        lambda r: (r["DEVENGADO"] / r["CODIFICADO"] * 100) if r["CODIFICADO"] else 0.0,
        axis=1
    )
    resumen_tipo = resumen_tipo.sort_values("EJECUCION", ascending=True).tail(10)
    colores_tipo = ["#83bac0" if v >= 70 else "#f7bd24" if v >= 40 else "#55479d"
                    for v in resumen_tipo["EJECUCION"]]
    fig_tipo = go.Figure(go.Bar(
        x=resumen_tipo["EJECUCION"], y=resumen_tipo[dimension], orientation="h",
        marker_color=colores_tipo,
        text=[f"{v:.1f}%" for v in resumen_tipo["EJECUCION"]], textposition="outside",
        customdata=resumen_tipo[["CODIFICADO", "DEVENGADO"]],
        hovertemplate=("<b>%{y}</b><br>Ejecución: %{x:.2f}%"
                       "<br>Codificado: $%{customdata[0]:,.2f}"
                       "<br>Devengado: $%{customdata[1]:,.2f}<extra></extra>"),
    ))
    fig_tipo.update_layout(
        height=330, margin=dict(l=145, r=55, t=55, b=45),
        title={"text": "Porcentaje ejecutado por tipo de gasto", "x": 0.5, "xanchor": "center"},
        paper_bgcolor="white", plot_bgcolor="white", showlegend=False
    )
    fig_tipo.update_xaxes(range=[0, max(105, float(resumen_tipo["EJECUCION"].max()) + 12)],
                          ticksuffix="%", gridcolor="#e5e7f0", title="Devengado / Codificado")
    fig_tipo.update_yaxes(showgrid=False, automargin=True)

    columnas_preferidas = ["SUBSECRETARÍA", "NOM_PROGRAMA", "NOM_PROYECTO", "NOM_ACTIVIDAD",
                           "ZONA", "PROVINCIA", "CANTÓN",
                           "CODIFICADO", "COMPROMISO", "DEVENGADO", "PAGADO", "SALDO_DISPONIBLE"]
    columnas = [c for c in columnas_preferidas if c in grupo.columns]
    detalle = grupo.loc[grupo["CODIFICADO"] > 0, columnas].copy()
    for col in [c for c in PRESUPUESTO_MONETARIAS if c in detalle.columns]:
        detalle[col] = detalle[col].map(_moneda)

    return html.Div([
        html.Div([html.P("EJECUCIÓN PRESUPUESTARIA", className="content-kicker"),
                  html.H1(vice),
                  html.P("Resumen de ejecución del presupuesto codificado."),
                  insignia_actualizacion],
                 className="content-heading"),
        html.Div([html.Div([html.Span(t), html.Strong(v), html.Small(s)],
                           style={"borderTop": f"4px solid {color}"})
                  for (t, v, s), color in zip(kpis, ["#55479d", "#83bac0", "#f7bd24", "#25335f", "#8d83c7"])],
                 className="budget-kpi-grid"),
        html.Div([html.Div(dcc.Graph(figure=fig_avance, config={"displayModeBar": False}), className="budget-panel"),
                  html.Div(dcc.Graph(figure=fig_tipo, config={"displayModeBar": False}), className="budget-panel")],
                 className="budget-chart-grid"),
        html.Section([
            html.Div([html.H2("Detalle presupuestario"),
                      html.Span(f"{len(detalle):,} registros".replace(",", "."), id="budget-table-count")],
                     className="budget-table-title"),
            html.Div([
                html.Div([html.Label("Tipo"), dcc.Dropdown(
                    id="budget-filter-tipo", multi=True, clearable=True, placeholder="Todos",
                    options=[{"label": x, "value": x} for x in sorted(
                        grupo.get("TIPO", pd.Series(dtype=str)).dropna().astype(str).unique())]
                )]),
            ], style={"maxWidth": "420px", "marginBottom": "14px"}),
            dash_table.DataTable(
                id="budget-detail-table",
                data=detalle.to_dict("records"), columns=[{"name": c.replace("_", " ").title(), "id": c} for c in columnas],
                page_size=12, sort_action="native", filter_action="none", fixed_rows={"headers": True},
                style_table={"overflowX": "auto", "maxHeight": "520px"},
                style_cell={"fontFamily": "Arial", "fontSize": "11px", "padding": "8px",
                            "minWidth": "105px", "maxWidth": "260px", "whiteSpace": "normal", "textAlign": "left"},
                style_header={"backgroundColor": "#e6e8ef", "color": "#25335f", "fontWeight": "700"},
                style_data_conditional=[{"if": {"row_index": "odd"}, "backgroundColor": "#f7f7fb"}],
            )
        ], className="budget-table-card"),
        html.P(f"Fuente: ESIGEF · {PRESUPUESTO_META.get('origen', '')}", className="budget-source")
    ], className="indicator-content budget-content")


# ---------------------------------------------------------------------------
# Enrutamiento genérico de módulos
# ---------------------------------------------------------------------------
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
              Input({"type": "side-section", "index": ALL}, "n_clicks"),
              State("active-section", "data"), prevent_initial_call=True)
def cambiar_seccion(_home, _cards, _side, actual):
    disparador = ctx.triggered_id
    if disparador == "home-button":
        return ("home", None, None) if _home else (actual, no_update, no_update)
    if isinstance(disparador, dict):
        tipo = disparador.get("type")
        valores = _cards if tipo == "home-card" else _side
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
    if vice not in set(df[cfg["vice_col"]].dropna().unique()):
        return actual, no_update
    # Segundo clic sobre el mismo viceministerio: recoge el acordeón y limpia la ficha.
    if vice == actual:
        return None, None
    if seccion in {"vision", "presupuesto"}:
        return vice, None
    indicadores = df.loc[df[cfg["vice_col"]] == vice, cfg["indicador_col"]].dropna().unique()
    primer_indicador = indicadores[0] if len(indicadores) else None
    return vice, primer_indicador


@app.callback(Output("selected-indicator", "data"),
              Input({"type": "indicator-button", "index": ALL}, "n_clicks"),
              State("selected-vice", "data"), State("selected-indicator", "data"),
              State("active-section", "data"), prevent_initial_call=True)
def seleccionar_indicador(_clicks, vice, actual, seccion):
    cfg = SECCIONES_CON_DATOS.get(seccion)
    df, _ = obtener_datos(seccion)
    if (not isinstance(ctx.triggered_id, dict) or not cfg or not vice or not _clicks
            or not any((v or 0) > 0 for v in _clicks)):
        return actual
    indicador = ctx.triggered_id["index"]
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
    return modulo_seccion(seccion, vice, indicador)


@app.callback(
    Output("budget-detail-table", "data"),
    Output("budget-table-count", "children"),
    Input("budget-filter-tipo", "value"),
    State("selected-vice", "data"),
    prevent_initial_call=True,
)
def filtrar_detalle_presupuesto(tipos, vice):
    """Filtra por tipo y conserva únicamente partidas con codificado positivo."""
    if not vice or DATA_PRESUPUESTO.empty:
        return [], "0 registros"

    detalle = DATA_PRESUPUESTO.loc[
        DATA_PRESUPUESTO["Viceministerio"] == vice
    ].copy()
    detalle = detalle[detalle["CODIFICADO"] > 0]
    if tipos and "TIPO" in detalle.columns:
        detalle = detalle[detalle["TIPO"].astype(str).isin(tipos)]

    columnas_preferidas = [
        "SUBSECRETARÍA", "NOM_PROGRAMA", "NOM_PROYECTO", "NOM_ACTIVIDAD",
        "ZONA", "PROVINCIA", "CANTÓN",
        "CODIFICADO", "COMPROMISO", "DEVENGADO", "PAGADO", "SALDO_DISPONIBLE",
    ]
    columnas = [c for c in columnas_preferidas if c in detalle.columns]
    detalle = detalle[columnas].copy()
    for col in [c for c in PRESUPUESTO_MONETARIAS if c in detalle.columns]:
        detalle[col] = detalle[col].map(_moneda)
    cantidad = f"{len(detalle):,} registros".replace(",", ".")
    return detalle.to_dict("records"), cantidad


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
    global DATA_VISION, ERROR_VISION, VERSION_VISION
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

    nueva_version_vision = version_vision()
    if nueva_version_vision and nueva_version_vision != version_actual.get("vision"):
        nueva_data, nuevo_error = cargar_base_vision()
        if not nuevo_error and not nueva_data.empty:
            DATA_VISION, ERROR_VISION, VERSION_VISION = nueva_data, None, nueva_version_vision
            version_actual["vision"] = nueva_version_vision
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
