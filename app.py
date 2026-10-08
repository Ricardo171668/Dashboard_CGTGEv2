"""
Panel de Inversión MINEDEC — versión Dash

Arquitectura orientada a rendimiento:
- Dash/Flask: los callbacks actualizan solo el componente afectado.
- SharePoint institucional como fuente (sin OneDrive personal).
- Caché en memoria con TTL: no vuelve a descargar Excel en cada clic.
- Carga diferida: presupuesto al entrar; ejecución al mostrar Gantt;
  Banco de Proyectos/PRETT únicamente al consultar un detalle.
"""

from __future__ import annotations

import base64
import hashlib
import io
import os
import re
import threading
import time
import unicodedata
from datetime import datetime
from difflib import get_close_matches
from urllib.parse import parse_qs, quote, unquote

import pandas as pd
import openpyxl

try:
    import python_calamine  # noqa: F401
    _HAS_CALAMINE = True
except ImportError:
    _HAS_CALAMINE = False
import plotly.graph_objects as go
import requests
from dash import ALL, Dash, Input, Output, State, ctx, dcc, html, no_update
from flask import Flask


# =============================================================================
# 1. CONFIGURACIÓN
# =============================================================================

URL_PRESUPUESTO = (
    "https://educacionec.sharepoint.com/:x:/s/DocumentacinDNSE/"
    "IQAF7raaOY2gRLSqM8J-kLEQAcrTKgB8Ga1VyXIrMcc79qI"
    "?e=dsBNDf&CID=3a90d370-fcb6-8f8c-28fc-1c63a63a09bd"
)

URL_BANCO_PROYECTOS = (
    "https://educacionec.sharepoint.com/:x:/s/DocumentacinDNSE/"
    "IQDVr77shy8qT7nYBIahKQQNAbRhPkTnK71B2x9iMiAz3LA"
    "?e=NpYkAW&CID=0c26ad54-e9f6-147c-f500-88d9b0cd0075"
)

URLS_PRETT = [
    (
        "https://educacionec.sharepoint.com/:x:/s/DocumentacinDNSE/"
        "IQDCY7-LdgDhT4QqcNHV1DyTARigN6HoJSzIaB1jxxMehvk"
        "?e=1XxBAn&CID=e2e3ce30-6b70-71c6-c006-afb03d8f2247"
    )
]

URL_EJECUCION_MENSUAL = (
    "https://educacionec.sharepoint.com/:x:/s/DocumentacinDNSE/"
    "IQAx97NfQnZNTozEs_I5kCcrARhZAXfKhUbv6Rby25YX2vo"
    "?rtime=JTIrh8MC30g"
)

CACHE_TTL_SECONDS = int(os.getenv("CACHE_TTL_SECONDS", "3600"))  # 60 min
REQUEST_CONNECT_TIMEOUT = 4
REQUEST_READ_TIMEOUT = 18

PURPLE_900 = "#25206D"
PURPLE_800 = "#332A8F"
PURPLE_700 = "#4E3CC5"
PURPLE_600 = "#6550D8"
PURPLE_500 = "#7A66E4"
PURPLE_300 = "#B9AEF3"
PURPLE_100 = "#EEEAFE"
PURPLE_50 = "#F8F6FF"
INK = "#18203A"
MUTED = "#717A91"
BORDER = "#E5E7F0"
GREEN = "#2FA66F"
GRAY = "#9AA2B2"


# =============================================================================
# 2. SERVIDOR / DASH
# =============================================================================


def _resolve_base_dir() -> str:
    """
    Dash necesita saber en qué carpeta física vive este archivo para
    encontrar assets/ al lado. Su detección automática usa __file__ por
    debajo, lo cual falla al correr con el botón "Run" de Spyder (runfile).
    Por eso lo resolvemos nosotros mismos, con respaldo al directorio de
    trabajo actual si __file__ no está disponible.
    """
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except NameError:
        return os.getcwd()


_BASE_DIR = _resolve_base_dir()

_CACHE_DIR = os.path.join(_BASE_DIR, "cache")
_EJECUCION_CACHE_FILE = os.path.join(_CACHE_DIR, "ejecucion_mensual.pkl")
os.makedirs(_CACHE_DIR, exist_ok=True)

server = Flask(__name__)
server.secret_key = os.getenv("FLASK_SECRET_KEY") or os.urandom(32)

app = Dash(
    __name__,
    server=server,
    suppress_callback_exceptions=True,
    title="Panel de Inversión MINEDEC",
    update_title=None,
    assets_folder=os.path.join(_BASE_DIR, "assets"),
)

# El CSS vive en assets/style.css Y ADEMÁS embebido aquí directo como
# respaldo — así el estilo no depende de que Dash encuentre bien la carpeta
# assets/, que ya ha fallado antes en este entorno según cómo se corra el
# script. Con el CSS embebido, aplica siempre, sin importar rutas.
app.index_string = """<!DOCTYPE html>
<html>
    <head>
        {%metas%}
        <title>{%title%}</title>
        {%favicon%}
        {%css%}
        <style>
:root{
  --purple-900:#25206D;
  --purple-800:#332A8F;
  --purple-700:#4E3CC5;
  --purple-600:#6550D8;
  --purple-500:#7A66E4;
  --purple-300:#B9AEF3;
  --purple-100:#EEEAFE;
  --purple-50:#F8F6FF;
  --ink:#18203A;
  --muted:#717A91;
  --border:#E5E7F0;
  --green:#2FA66F;
  --gray:#9AA2B2;
  --page:#F5F6FA;
  --white:#FFFFFF;
  --shadow:0 8px 30px rgba(37,32,109,.08);
}

*{box-sizing:border-box;}
html,body,#react-entry-point,#_dash-app-content{
  margin:0;
  min-height:100%;
}
body{
  font-family:Inter,Segoe UI,Roboto,Arial,sans-serif;
  color:var(--ink);
  background:var(--page);
}
a{text-decoration:none;color:inherit;}
button,input,select{font:inherit;}

/* LOGIN */
.login-page{
  min-height:100vh;
  width:100%;
  display:flex;
  align-items:center;
  justify-content:center;
  padding:28px;
  background:
    radial-gradient(circle at 16% 18%,rgba(185,174,243,.55),transparent 34%),
    linear-gradient(135deg,#8F89BF 0%,#667BA8 55%,#526D9A 100%);
}
.login-center{
  width:100%;
  display:flex;
  justify-content:center;
}
.login-card{
  width:min(430px,92vw);
  background:#fff;
  border-radius:28px;
  padding:42px 44px 36px;
  box-shadow:0 30px 80px rgba(37,32,109,.22);
  text-align:center;
}
.login-logo{
  width:340px;
  height:auto;
  max-height:200px;
  object-fit:contain;
  display:block;
  margin:0 auto 30px;
}
.login-title{
  font-size:34px;
  line-height:1.12;
  margin:0 0 10px;
  letter-spacing:-.5px;
  color:#20243A;
}
.login-subtitle{
  margin:0 0 22px;
  color:#8A8F9D;
  font-size:14px;
}
.password-input{
  width:100%;
  height:52px;
  border:1px solid #E4E6EE;
  border-radius:10px;
  padding:0 14px;
  line-height:normal;
  box-sizing:border-box;
  background:#F7F8FB;
  outline:none;
  color:var(--ink);
  font-size:15px;
  transition:.2s ease;
}
.password-input:focus{
  border-color:var(--purple-500);
  box-shadow:0 0 0 3px rgba(122,102,228,.12);
  background:#fff;
}
.login-button{
  width:100%;
  height:56px;
  margin-top:16px;
  padding:0 18px;
  border:0;
  border-radius:10px;
  color:#fff;
  font-weight:700;
  font-size:16px;
  cursor:pointer;
  background:linear-gradient(90deg,#533DD7,#A548D1);
  box-shadow:0 10px 24px rgba(83,61,215,.18);
}
.login-button:hover{filter:brightness(1.03);}
.login-error{
  min-height:22px;
  margin-top:10px;
  color:#C83F55;
  font-size:13px;
}
.login-foot{
  margin:14px 0 0;
  color:#999EAA;
  font-size:12px;
}

/* APP SHELL */
.app-shell{
  min-height:100vh;
  display:grid;
  grid-template-columns:260px minmax(0,1fr);
  background:var(--page);
}
.sidebar{
  position:sticky;
  top:0;
  height:100vh;
  padding:20px 16px 18px;
  display:flex;
  flex-direction:column;
  color:#fff;
  background:linear-gradient(180deg,#2D2778 0%,#243E83 100%);
  box-shadow:10px 0 32px rgba(37,32,109,.06);
  z-index:20;
}
.collapse-symbol{
  text-align:right;
  font-size:22px;
  font-weight:800;
  margin:0 6px 22px;
  opacity:.95;
}
.logo-box{
  background:#fff;
  border-radius:8px;
  min-height:140px;
  display:flex;
  align-items:center;
  justify-content:center;
  padding:14px;
  margin-bottom:36px;
}
.sidebar-logo{
  width:220px;
  max-width:100%;
  height:110px;
  object-fit:contain;
}
.sidebar-title{
  margin:0 0 10px;
  font-size:18px;
  font-weight:800;
}
.sidebar-subtitle{
  margin:0 0 28px;
  color:#BEC7E0;
  font-size:13px;
}
.sidebar nav{
  display:flex;
  flex-direction:column;
  gap:14px;
}
.nav-button{
  width:100%;
  border:1px solid rgba(255,255,255,.08);
  background:rgba(255,255,255,.07);
  color:#fff;
  padding:14px 16px;
  border-radius:11px;
  font-size:14px;
  text-align:center;
  transition:.18s ease;
}
.nav-button:hover{
  background:rgba(255,255,255,.13);
  transform:translateY(-1px);
}
.sidebar-spacer{flex:1;}
.sidebar-date{
  font-size:12px;
  color:#BBC6E2;
  margin:0 0 16px;
}
.logout-button{
  width:100%;
  padding:13px;
  border-radius:10px;
  border:1px solid rgba(255,255,255,.10);
  background:rgba(255,255,255,.08);
  color:#fff;
  cursor:pointer;
  font-weight:600;
}
.logout-button:hover{background:rgba(255,255,255,.14);}

.main-content{
  width:100%;
  max-width:1460px;
  margin:0 auto;
  padding:18px 26px 38px;
}

/* HEADER */
.header-card{
  display:grid;
  grid-template-columns:minmax(0,1fr) 220px;
  gap:14px;
  padding:14px;
  border:1px solid var(--border);
  border-radius:14px;
  margin-bottom:18px;
  background:#fff;
}
.top-gradient{
  border-radius:12px;
  padding:22px 26px;
  min-height:106px;
  display:flex;
  flex-direction:column;
  justify-content:center;
  color:#fff;
  background:linear-gradient(90deg,#573FD3 0%,#30257C 100%);
}
.top-title{
  font-size:27px;
  line-height:1.15;
  font-weight:800;
}
.top-subtitle{
  margin-top:8px;
  color:#DED9FF;
  font-size:13px;
}
.updated-card{
  display:flex;
  flex-direction:column;
  justify-content:center;
  align-items:flex-start;
  border:1px solid var(--border);
  border-radius:12px;
  padding:16px 20px;
  background:#fff;
}
.updated-label{
  font-size:12px;
  margin-bottom:6px;
}
.updated-value{
  font-size:31px;
  line-height:1;
  font-weight:400;
}

/* KPI */
.kpi-grid{
  display:grid;
  grid-template-columns:repeat(6,minmax(0,1fr));
  gap:14px;
  margin-bottom:18px;
}
.kpi-card{
  min-width:0;
  min-height:122px;
  padding:16px 18px;
  border:1px solid var(--border);
  border-radius:14px;
  background:#fff;
  box-shadow:0 2px 12px rgba(37,32,109,.02);
}
.kpi-icon{
  width:36px;
  height:36px;
  border-radius:9px;
  display:flex;
  align-items:center;
  justify-content:center;
  font-size:19px;
  margin-bottom:11px;
}
.kpi-value{
  font-size:21px;
  font-weight:800;
  white-space:nowrap;
  overflow:hidden;
  text-overflow:ellipsis;
}
.kpi-label{
  color:#5E6880;
  margin-top:6px;
  font-size:13px;
}

/* GENERAL CARDS */
.content-card{
  border:1px solid #D9DCE6;
  border-radius:14px;
  background:#fff;
  padding:16px 18px;
  margin-bottom:18px;
}
.section-title{
  font-size:17px;
  font-weight:800;
  margin:0 0 5px;
}
.section-subtitle{
  font-size:12px;
  color:#7E879B;
  margin:0 0 10px;
}
.graph{
  width:100%;
}
.hint{
  color:#8B92A1;
  font-size:13px;
  margin-top:4px;
}
.alert,.info-alert{
  padding:14px 16px;
  border-radius:9px;
  background:#E7F0FF;
  color:#2763A3;
  border:1px solid #D6E5FA;
}

/* RADIO */
.vice-radio{
  display:flex!important;
  flex-wrap:wrap;
  gap:8px;
  margin:6px 0 12px;
}
.radio-label{
  display:inline-flex!important;
  align-items:center;
  gap:7px;
  border:1px solid var(--border);
  background:#fff;
  padding:8px 13px;
  border-radius:999px;
  cursor:pointer;
  color:#26324C;
  font-size:13px;
}
.radio-label:has(input:checked){
  background:var(--purple-700);
  color:#fff;
  border-color:var(--purple-700);
  font-weight:700;
}
.radio-input{accent-color:#FF5964;}

/* PROJECTS */
.project-filters{
  display:grid;
  grid-template-columns:minmax(250px,1fr) minmax(220px,360px);
  gap:14px;
  margin-bottom:16px;
}
.search-input{
  width:100%;
  border:1px solid var(--border);
  background:#fff;
  border-radius:10px;
  padding:12px 14px;
  outline:none;
}
.vice-dropdown .Select-control,
.vice-dropdown .Select__control{
  border-radius:10px!important;
  border-color:var(--border)!important;
  min-height:43px!important;
}
.project-row{
  display:grid;
  grid-template-columns:46px minmax(0,1fr) 150px 120px 120px;
  gap:12px;
  align-items:center;
  padding:17px 0;
  border-bottom:1px solid #E6E8F0;
}
.project-row:last-child{border-bottom:0;}
.project-icon{
  width:38px;height:38px;
  display:flex;align-items:center;justify-content:center;
  border-radius:10px;
  color:var(--purple-700);
  background:var(--purple-100);
  font-size:18px;
}
.project-name{
  font-weight:750;
  line-height:1.3;
  margin-bottom:7px;
}
.soft-tag{
  display:inline-block;
  padding:4px 8px;
  background:var(--purple-100);
  color:var(--purple-700);
  border-radius:6px;
  font-size:10px;
  font-weight:700;
}
.project-metric{
  display:flex;
  flex-direction:column;
  gap:5px;
}
.project-metric strong{font-size:14px;}
.project-metric small{font-size:11px;color:#8D94A3;}
.detail-link,.back-link{
  border:1px solid #DCD8FF;
  color:var(--purple-700);
  border-radius:9px;
  padding:10px 12px;
  text-align:center;
  font-size:12px;
  background:#fff;
}
.detail-link:hover,.back-link:hover{
  background:var(--purple-50);
}

/* DETAIL */
.detail-container{width:100%;}
.detail-main-title{
  margin:0 0 8px;
  font-size:23px;
}
.detail-kpi-grid{
  display:grid;
  grid-template-columns:repeat(4,minmax(0,1fr));
  gap:14px;
  margin:16px 0 18px;
}
.detail-two-cols{
  display:grid;
  grid-template-columns:1fr 1fr;
  gap:14px;
  margin-bottom:18px;
}
.detail-two-cols.compact{align-items:start;}
.summary-box{
  padding:14px 16px;
  border-radius:9px;
  background:#E9F2FF;
  color:#13589D;
  margin:10px 0 12px;
  line-height:1.45;
}
.meta-three{
  display:grid;
  grid-template-columns:1fr 1fr 1fr;
  gap:14px;
  padding-top:8px;
  font-size:13px;
}
.native-details{
  border:1px solid #D9DCE6;
  border-radius:9px;
  background:#fff;
  margin-bottom:12px;
  overflow:hidden;
}
.native-details summary{
  cursor:pointer;
  padding:12px 14px;
  font-size:13px;
  list-style:none;
}
.native-details summary::-webkit-details-marker{display:none;}
.native-details summary:before{
  content:"›";
  display:inline-block;
  margin-right:9px;
  font-weight:800;
  color:var(--purple-700);
}
.native-details[open] summary:before{transform:rotate(90deg);}
.details-body{
  border-top:1px solid #ECEEF4;
  padding:14px 16px;
  font-size:13px;
  line-height:1.5;
}
.prett-block{margin-top:18px;}
.activity-card{
  padding:12px 14px;
  border:1px solid #E1E4EC;
  border-radius:9px;
  background:#fff;
  margin-top:9px;
}
.activity-head{
  display:flex;
  gap:12px;
  justify-content:space-between;
  align-items:flex-start;
}
.activity-title{font-size:13px;}
.activity-pct{font-weight:800;color:var(--purple-700);}
.progress-track{
  width:100%;
  height:7px;
  border-radius:999px;
  overflow:hidden;
  background:#E9ECF3;
  margin:10px 0;
}
.progress-fill{
  height:100%;
  border-radius:999px;
  background:linear-gradient(90deg,#6550D8,#4E3CC5);
}
.activity-meta,.macro-meta{
  display:flex;
  flex-wrap:wrap;
  gap:10px 20px;
  color:#717A91;
  font-size:11px;
}

/* Dash loading overlay – keeps the UI visible instead of washing out whole page */
._dash-loading{
  color:var(--purple-700)!important;
}


/* Los nombres del Gantt son elementos interactivos */
.js-plotly-plot .scatterlayer text{
  cursor:pointer!important;
}


/* ===== Línea de tiempo HTML nativa ===== */
.timeline-matrix{width:100%;overflow-x:auto;padding-bottom:4px;}
.timeline-grid{min-width:1120px;display:grid;align-items:stretch;}
.timeline-head{font-size:13px;font-weight:700;color:#4D5870;padding:10px 8px 12px;text-align:center;border-bottom:1px solid #E9EBF3;}
.timeline-head.project-col{text-align:left;padding-left:4px;}
.timeline-project-cell,.timeline-month-cell,.timeline-progress-cell,.timeline-segments-cell{min-height:62px;display:flex;align-items:center;border-bottom:1px solid #EFF1F6;}
.timeline-project-cell{padding:8px 12px 8px 4px;}
.timeline-project-button{width:100%;border:0;background:transparent;color:#30257C;text-decoration:underline;text-underline-offset:2px;font-weight:700;font-size:14px;line-height:1.3;text-align:left;cursor:pointer;padding:6px 28px 6px 0;position:relative;}
.timeline-project-button::after{content:"›";position:absolute;right:7px;top:50%;transform:translateY(-50%);color:#7A66E4;font-size:20px;}
.timeline-project-button:hover{color:#573FD3;}
.timeline-month-cell{justify-content:center;padding:8px 4px;}
.timeline-segment{width:100%;height:28px;border-radius:5px;transition:transform .12s ease,box-shadow .12s ease;}
.timeline-segment.exec{background:linear-gradient(90deg,#4E3CC5,#573FD3);}
.timeline-segment.noexec{background:#B9AEF3;}
.timeline-segment.empty{background:transparent;}
.timeline-segment:not(.empty):hover{transform:translateY(-1px);box-shadow:0 4px 12px rgba(78,60,197,.18);}
.timeline-progress-cell{padding:8px 10px 8px 16px;flex-direction:column;align-items:flex-start;justify-content:center;gap:6px;}
.timeline-progress-label{font-size:13px;font-weight:700;color:#202A43;}
.timeline-mini-track{width:78px;height:5px;border-radius:999px;background:#E9EAF2;overflow:hidden;}
.timeline-mini-fill{height:100%;border-radius:999px;background:#4E3CC5;}
.timeline-segments-cell{justify-content:center;font-size:13px;font-weight:700;color:#44506A;}
.timeline-legend{display:flex;justify-content:flex-end;gap:22px;align-items:center;margin:3px 0 10px;font-size:13px;color:#44506A;}
.timeline-legend-item{display:flex;align-items:center;gap:7px;}
.timeline-legend-box{width:13px;height:13px;border-radius:2px;}
.timeline-legend-box.exec{background:#573FD3;}
.timeline-legend-box.noexec{background:#B9AEF3;}
.timeline-tip{margin-top:12px;padding:12px 14px;border:1px solid #DCE3F3;background:#F7FAFF;color:#63708A;border-radius:9px;font-size:12px;}
.timeline-selected-wrap{margin-top:14px;}
.timeline-detail-title{margin:0 0 12px;font-size:18px;font-weight:800;color:#30257C;}


/* ===== Detalle ejecutivo sin gráficos ===== */
.project-context-grid{display:grid;grid-template-columns:1.6fr 1fr 1fr;gap:14px;margin:14px 0 16px;}
.context-card{background:#fff;border:1px solid #E2E5EF;border-radius:14px;padding:17px 18px;min-height:105px;box-shadow:0 3px 12px rgba(37,32,109,.035);}
.context-label{font-size:11px;text-transform:uppercase;letter-spacing:.045em;color:#7B8498;font-weight:800;margin-bottom:8px;}
.context-value{font-size:13px;line-height:1.5;color:#18203A;font-weight:600;}
.project-values-title{font-size:16px;font-weight:800;color:#25206D;margin:18px 0 10px;}
.component-details{border:1px solid #D9DDF0;border-radius:14px;background:#fff;overflow:hidden;margin-top:16px;}
.component-details > summary{list-style:none;cursor:pointer;padding:17px 20px;display:flex;align-items:center;justify-content:space-between;color:#fff;background:linear-gradient(90deg,#4E3CC5,#30257C);font-size:15px;font-weight:800;}
.component-details > summary::-webkit-details-marker{display:none;}
.component-details > summary::after{content:'+';width:25px;height:25px;display:flex;align-items:center;justify-content:center;border-radius:50%;background:rgba(255,255,255,.15);font-size:18px;}
.component-details[open] > summary::after{content:'−';}
.component-body{padding:18px;}
.macro-stage{border:1px solid #E1E4EE;border-left:4px solid #573FD3;border-radius:12px;background:#fff;margin-bottom:14px;overflow:hidden;}
.macro-header{padding:15px 17px 13px;background:#F8F7FE;}
.macro-topline{display:flex;justify-content:space-between;align-items:flex-start;gap:14px;}
.macro-title{font-size:15px;font-weight:850;color:#17223C;line-height:1.35;}
.macro-code{display:inline-block;font-size:10px;color:#6657BA;font-weight:800;margin-bottom:4px;}
.status-pill{display:inline-flex;align-items:center;padding:5px 9px;border-radius:999px;font-size:10px;font-weight:800;white-space:nowrap;background:#EEEAFE;color:#4E3CC5;}
.status-pill.done{background:#E8F7F0;color:#247C58;}.status-pill.due{background:#FCEBEC;color:#B4384D;}.status-pill.review{background:#FFF3DF;color:#B56C09;}
.macro-meta-grid{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:10px;margin-top:12px;}
.meta-chip{background:#fff;border:1px solid #E5E7F0;border-radius:9px;padding:9px 10px;}
.meta-chip .lbl{font-size:9px;text-transform:uppercase;letter-spacing:.035em;color:#8A91A1;font-weight:800;margin-bottom:4px;}
.meta-chip .val{font-size:11px;color:#26324C;font-weight:700;line-height:1.3;}
.micro-wrap{padding:0 16px 16px;}.micro-table{width:100%;border-collapse:separate;border-spacing:0;font-size:11px;overflow:hidden;border:1px solid #E8EAF1;border-radius:10px;}
.micro-table th{background:#F3F4F8;color:#556077;text-align:left;padding:9px 8px;font-size:9px;text-transform:uppercase;letter-spacing:.02em;border-bottom:1px solid #E5E7EF;white-space:nowrap;}
.micro-table td{padding:9px 8px;border-bottom:1px solid #EFF1F5;vertical-align:top;color:#26324C;line-height:1.35;}.micro-table tr:last-child td{border-bottom:0;}
.micro-activity{font-weight:700;min-width:220px;}.micro-alert-vencida{color:#B4384D;font-weight:850;}.micro-alert-finalizada{color:#247C58;font-weight:850;}.empty-note{color:#8A91A1;font-style:italic;}
@media(max-width:1100px){.project-context-grid{grid-template-columns:1fr;}.macro-meta-grid{grid-template-columns:repeat(2,1fr);}}


/* =========================================================
   DETALLE JERÁRQUICO PROYECTO -> COMPONENTE -> MACRO -> MICRO
   ========================================================= */
.project-context-block{
  display:flex;
  flex-direction:column;
  gap:12px;
  margin:14px 0 18px;
}

.context-objective{
  width:100%;
  padding:18px 20px;
}

.objective-text{
  line-height:1.55;
  font-size:13.5px;
}

.project-context-secondary{
  display:grid;
  grid-template-columns:repeat(4,minmax(0,1fr));
  gap:12px;
}

.detail-project-name{
  margin:0 0 14px;
  font-size:19px;
  line-height:1.35;
  color:#17223C;
}

.component-details{
  margin-top:18px;
}

.component-details > summary{
  font-size:15px;
  font-weight:800;
  color:#30257C;
  background:#FFFFFF;
  border:1px solid #D8DEF0;
  border-radius:11px;
  padding:15px 17px;
  cursor:pointer;
  list-style:none;
  transition:.15s ease;
}

.component-details > summary:hover{
  border-color:#AFA1F2;
  background:#FBFAFF;
}

.component-details[open] > summary{
  border-color:#7A66E4;
  box-shadow:0 4px 16px rgba(83,63,211,.08);
}

.component-body{
  margin-top:12px;
}

@media(max-width:1000px){
  .project-context-secondary{
    grid-template-columns:repeat(2,minmax(0,1fr));
  }
}

@media(max-width:650px){
  .project-context-secondary{
    grid-template-columns:1fr;
  }
}


/* =========================================================
   TEMA EJECUTIVO - NAVEGACIÓN EN CASCADA
   Inspiración: admin dashboard claro, sidebar oscuro,
   tarjetas limpias y controles compactos.
   ========================================================= */

.route-cascade-shell{
  margin-top:20px;
  padding:20px;
  border:1px solid #DDE3F0;
  border-radius:16px;
  background:
    radial-gradient(circle at 100% 0%, rgba(103,79,219,.08), transparent 260px),
    #FFFFFF;
  box-shadow:0 10px 30px rgba(25,35,72,.045);
}

.cascade-header{
  display:flex;
  align-items:flex-start;
  justify-content:space-between;
  gap:20px;
  margin-bottom:18px;
  padding-bottom:15px;
  border-bottom:1px solid #EDF0F6;
}

.cascade-kicker{
  margin-bottom:4px;
  color:#6D5BD0;
  font-size:10px;
  font-weight:900;
  letter-spacing:.09em;
  text-transform:uppercase;
}

.cascade-section-title{
  margin:0;
  color:#17223C;
  font-size:19px;
  font-weight:850;
}

.cascade-section-subtitle{
  margin:6px 0 0;
  color:#768198;
  font-size:12px;
  line-height:1.45;
}

.cascade-count{
  min-width:96px;
  padding:9px 12px;
  border:1px solid #E1DCF9;
  border-radius:12px;
  background:#F7F4FF;
  text-align:center;
}

.cascade-count-number{
  display:block;
  color:#5038C9;
  font-size:20px;
  font-weight:900;
  line-height:1;
}

.cascade-count-label{
  display:block;
  margin-top:4px;
  color:#766F98;
  font-size:10px;
  font-weight:750;
}

.cascade-step-card{
  margin-top:12px;
  padding:16px;
  border:1px solid #E2E6F0;
  border-radius:13px;
  background:#FCFCFE;
}

.cascade-step-heading{
  display:flex;
  align-items:center;
  gap:11px;
  margin-bottom:12px;
}

.cascade-step-number{
  flex:0 0 auto;
  width:30px;
  height:30px;
  display:grid;
  place-items:center;
  border-radius:9px;
  background:linear-gradient(135deg,#5A43D8,#7865E8);
  box-shadow:0 5px 12px rgba(90,67,216,.18);
  color:#FFF;
  font-size:12px;
  font-weight:900;
}

.cascade-step-title{
  color:#202B46;
  font-size:13px;
  font-weight:850;
}

.cascade-step-help{
  margin-top:2px;
  color:#8A93A5;
  font-size:10.5px;
}

.cascade-dropdown .Select-control,
.cascade-dropdown .Select__control{
  min-height:44px!important;
  border:1px solid #DADFEB!important;
  border-radius:10px!important;
  box-shadow:none!important;
  background:#FFF!important;
}

.cascade-dropdown .Select-control:hover,
.cascade-dropdown .Select__control:hover{
  border-color:#8D7AE8!important;
}

.cascade-radio-list{
  display:flex;
  flex-direction:column;
  gap:7px;
}

.cascade-radio-label{
  position:relative;
  display:flex!important;
  align-items:center;
  width:100%;
  min-height:45px;
  margin:0!important;
  padding:10px 12px!important;
  border:1px solid #E1E5EF;
  border-radius:10px;
  background:#FFF;
  color:#35415B;
  font-size:11.5px;
  font-weight:700;
  line-height:1.35;
  cursor:pointer;
  transition:all .14s ease;
}

.cascade-radio-label:hover{
  border-color:#A99AF0;
  background:#FAF9FF;
  transform:translateX(2px);
}

.cascade-radio-label:has(input:checked){
  border-color:#6A54DF;
  background:linear-gradient(90deg,#F2EFFF,#FBFAFF);
  box-shadow:0 4px 13px rgba(91,69,207,.08);
  color:#35258E;
}

.cascade-radio-input{
  flex:0 0 auto;
  margin:0 10px 0 0!important;
  accent-color:#6048D9;
}

.macro-radio-list{
  max-height:375px;
  overflow-y:auto;
  padding-right:4px;
}

.micro-radio-list{
  max-height:360px;
  overflow-y:auto;
  padding-right:4px;
}

.cascade-selection-card{
  margin-bottom:13px;
  overflow:hidden;
  border:1px solid #DCDFF0;
  border-radius:12px;
  background:#FFF;
}

.cascade-selected-head{
  display:flex;
  justify-content:space-between;
  align-items:flex-start;
  gap:14px;
  padding:14px 15px;
  background:linear-gradient(90deg,#F5F2FF,#FBFAFF);
  border-bottom:1px solid #E8E5F5;
}

.cascade-eyebrow{
  display:flex;
  align-items:center;
  gap:7px;
  margin-bottom:4px;
}

.cascade-tag{
  padding:3px 6px;
  border-radius:5px;
  background:#5B43D2;
  color:#FFF;
  font-size:8px;
  font-weight:900;
  letter-spacing:.08em;
}

.cascade-code{
  color:#7B849A;
  font-size:10px;
  font-weight:800;
}

.cascade-selected-title{
  margin:0;
  color:#17223C;
  font-size:15px;
  font-weight:850;
}

.cascade-detail-grid{
  display:grid;
  grid-template-columns:repeat(4,minmax(0,1fr));
  gap:8px;
  padding:13px 15px 15px;
}

.cascade-metric{
  min-height:57px;
  padding:9px 10px;
  border:1px solid #E8EAF1;
  border-radius:8px;
  background:#FAFBFD;
}

.cascade-metric-wide{
  grid-column:1/-1;
}

.cascade-metric-label{
  margin-bottom:4px;
  color:#8991A3;
  font-size:8.5px;
  font-weight:850;
  letter-spacing:.025em;
  text-transform:uppercase;
}

.cascade-metric-value{
  color:#28344D;
  font-size:11.5px;
  font-weight:700;
  line-height:1.35;
  overflow-wrap:anywhere;
}

.cascade-micro-selector{
  padding-top:3px;
}

.cascade-list-caption{
  margin:0 0 9px;
  color:#647087;
  font-size:10.5px;
  font-weight:750;
}

.cascade-empty-state{
  padding:15px;
  border:1px dashed #D9DDE8;
  border-radius:10px;
  background:#FAFBFD;
  color:#8992A5;
  font-size:11px;
}

.cascade-activity-detail{
  margin-top:12px;
  overflow:hidden;
  border:1px solid #D7DCEF;
  border-radius:13px;
  background:#FFF;
  box-shadow:0 8px 22px rgba(38,42,86,.045);
}

.cascade-activity-head{
  display:flex;
  justify-content:space-between;
  align-items:flex-start;
  gap:16px;
  padding:16px;
  background:linear-gradient(100deg,#F7F5FF,#FFFFFF);
  border-bottom:1px solid #E9EAF3;
}

.cascade-activity-title{
  margin:2px 0 4px;
  color:#18223B;
  font-size:16px;
  font-weight:850;
  line-height:1.35;
}

.cascade-activity-code{
  color:#7868C8;
  font-size:10px;
  font-weight:850;
}

.cascade-reveal{
  animation:cascadeReveal .18s ease both;
}

@keyframes cascadeReveal{
  from{opacity:0;transform:translateY(5px);}
  to{opacity:1;transform:translateY(0);}
}

/* Ajustes visuales generales para una lectura ejecutiva */
.detail-container{
  max-width:100%;
}

.detail-kpi-grid .kpi-card{
  box-shadow:0 5px 16px rgba(28,36,67,.035);
}

.context-card{
  box-shadow:0 5px 16px rgba(28,36,67,.035);
}

@media(max-width:1050px){
  .cascade-detail-grid{
    grid-template-columns:repeat(2,minmax(0,1fr));
  }
}

@media(max-width:650px){
  .route-cascade-shell{
    padding:14px;
  }
  .cascade-header{
    flex-direction:column;
  }
  .cascade-detail-grid{
    grid-template-columns:1fr;
  }
}


/* =========================================================
   RESUMEN COMPACTO DEL MACRO
   ========================================================= */
.macro-executive-summary{
  display:grid;
  grid-template-columns:2fr repeat(3,1fr);
  gap:0;
  border-top:1px solid #E9E7F4;
  background:#FFFFFF;
}

.macro-summary-item{
  padding:12px 14px;
  border-right:1px solid #EDF0F5;
}

.macro-summary-item:last-child{
  border-right:0;
}

.macro-summary-label{
  margin-bottom:4px;
  color:#8A92A5;
  font-size:8.5px;
  font-weight:850;
  letter-spacing:.035em;
  text-transform:uppercase;
}

.macro-summary-value{
  color:#26324A;
  font-size:11.5px;
  font-weight:800;
  line-height:1.3;
}

.macro-summary-period .macro-summary-value{
  color:#4C3BBC;
}

@media(max-width:850px){
  .macro-executive-summary{
    grid-template-columns:1fr 1fr;
  }

  .macro-summary-item{
    border-bottom:1px solid #EDF0F5;
  }
}

@media(max-width:550px){
  .macro-executive-summary{
    grid-template-columns:1fr;
  }

  .macro-summary-item{
    border-right:0;
  }
}

/* Responsive */
@media(max-width:1200px){
  .kpi-grid{grid-template-columns:repeat(3,1fr);}
  .project-row{grid-template-columns:44px minmax(0,1fr) 130px 110px;}
  .detail-link{grid-column:2/-1;justify-self:end;}
}
@media(max-width:900px){
  .app-shell{grid-template-columns:1fr;}
  .sidebar{
    position:relative;
    height:auto;
    min-height:0;
  }
  .sidebar-spacer{display:none;}
  .main-content{padding:14px;}
  .header-card{grid-template-columns:1fr;}
  .updated-card{min-height:86px;}
  .kpi-grid{grid-template-columns:repeat(2,1fr);}
  .detail-kpi-grid{grid-template-columns:repeat(2,1fr);}
  .detail-two-cols{grid-template-columns:1fr;}
  .meta-three{grid-template-columns:1fr;}
}
@media(max-width:640px){
  .login-card{padding:32px 24px;}
  .login-title{font-size:28px;}
  .project-filters{grid-template-columns:1fr;}
  .project-row{
    grid-template-columns:42px 1fr;
  }
  .project-metric,.detail-link{
    grid-column:2;
  }
  .kpi-grid,.detail-kpi-grid{grid-template-columns:1fr;}
}

        </style>
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


@server.after_request
def _no_cache(response):
    """
    Fuerza al navegador a no guardar en caché NINGUNA respuesta de esta app
    (ni el HTML, ni el CSS, ni el JS de Dash, ni las llamadas de los
    callbacks). Esto evita que quede pegada una versión vieja y rota de la
    página después de haber corregido el código — el síntoma exacto que
    hemos visto repetidas veces en esta sesión.
    """
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response


# =============================================================================
# 3. CACHÉ EN MEMORIA
# =============================================================================

_CACHE: dict[str, tuple[float, object]] = {}
_CACHE_LOCK = threading.RLock()


def cached_value(key: str, loader, ttl: int = CACHE_TTL_SECONDS):
    now = time.time()
    with _CACHE_LOCK:
        item = _CACHE.get(key)
        if item and now - item[0] < ttl:
            return item[1]
    value = loader()
    with _CACHE_LOCK:
        _CACHE[key] = (time.time(), value)
    return value


def clear_data_cache():
    with _CACHE_LOCK:
        _CACHE.clear()


# =============================================================================
# 4. DESCARGA SHAREPOINT
# =============================================================================


def variantes_url(url: str) -> list[str]:
    if not url:
        return []
    if "download=1" in url:
        return [url]
    if "?" in url:
        return [f"{url}&download=1", url]
    return [f"{url}?download=1", url]


def descargar_excel(url: str) -> io.BytesIO | None:
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 Chrome/124 Safari/537.36"
        ),
        "Accept": (
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,"
            "application/octet-stream,*/*"
        ),
    }
    for candidate in variantes_url(url):
        try:
            r = requests.get(
                candidate,
                timeout=(REQUEST_CONNECT_TIMEOUT, REQUEST_READ_TIMEOUT),
                allow_redirects=True,
                headers=headers,
            )
            r.raise_for_status()
            content = r.content
            ctype = r.headers.get("Content-Type", "").lower()
            if len(content) > 1000 and content[:2] == b"PK" and "text/html" not in ctype:
                return io.BytesIO(content)
        except Exception:
            continue
    return None


# =============================================================================
# 5. AUXILIARES
# =============================================================================


_LOGO_CACHE = {}


def logo_src():
    """
    Busca logo-minedec.png en varias ubicaciones posibles (assets/ junto al
    script, assets/ en el directorio de trabajo, y la carpeta del proyecto
    directa) y lo incrusta como data URI en base64. Esto no depende de que
    Dash resuelva bien la ruta de assets/ ni de rutas servidas por Flask —
    funciona sin importar cómo se haya corrido el script.
    """
    if "src" in _LOGO_CACHE:
        return _LOGO_CACHE["src"]

    candidatos = [
        os.path.join(_BASE_DIR, "assets", "logo-minedec.png"),
        os.path.join(os.getcwd(), "assets", "logo-minedec.png"),
        os.path.join(_BASE_DIR, "logo-minedec.png"),
        os.path.join(os.getcwd(), "logo-minedec.png"),
    ]
    for ruta in candidatos:
        try:
            with open(ruta, "rb") as f:
                encoded = base64.b64encode(f.read()).decode("ascii")
            src = f"data:image/png;base64,{encoded}"
            _LOGO_CACHE["src"] = src
            return src
        except Exception:
            continue

    _LOGO_CACHE["src"] = ""
    return ""


def normalizar(s: str) -> str:
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", s).strip().lower()


def parse_dated_entries(text, max_entries=6):
    if not isinstance(text, str) or not text.strip():
        return []
    pattern = re.compile(r"(\d{2}/\d{2}/\d{4})\s*[•\-–]\s*")
    parts = pattern.split(text)
    entries = []
    for i in range(1, len(parts), 2):
        fecha = parts[i].strip()
        cuerpo = parts[i + 1].strip() if i + 1 < len(parts) else ""
        cuerpo = re.sub(r"\s{2,}", " ", cuerpo.replace("\n", " ")).strip()
        if cuerpo:
            entries.append({"fecha": fecha, "texto": cuerpo[:800]})
    return entries[-max_entries:]


def vice_short(v):
    return (v or "Sin viceministerio").replace("Viceministerio de ", "")


def fmt_money(v):
    return f"$ {float(v):,.2f}"


def fmt_mm(v):
    v = float(v)
    return f"$ {v/1_000_000:,.2f} MM" if abs(v) >= 1_000_000 else f"$ {v:,.0f}"


def auth_ok(password: str) -> bool:
    env_hash = os.getenv("APP_PASSWORD_HASH")
    if env_hash:
        return hashlib.sha256(password.encode("utf-8")).hexdigest() == env_hash
    expected = os.getenv("APP_PASSWORD", "minedec2026")
    return password == expected


def match_project(nombre: str, nombres_norm: dict[str, str], cutoff=0.88):
    clave = normalizar(nombre)
    real = nombres_norm.get(clave)
    if real:
        return real
    if not nombres_norm:
        return None
    matches = get_close_matches(clave, list(nombres_norm.keys()), n=1, cutoff=cutoff)
    return nombres_norm[matches[0]] if matches else None


# =============================================================================
# 6. PROCESAMIENTO: PRESUPUESTO
# =============================================================================


def procesar_presupuesto(buf: io.BytesIO) -> dict:
    df = pd.read_excel(buf, sheet_name="HOY")
    df.columns = [str(c).strip() for c in df.columns]
    name_col = "NOMBRE DEL PROYECTO"
    vice_col = "VICEMINISTERIO PROYECTOS"
    df = df[df[name_col].astype(str).str.strip().str.lower().ne("corriente")].copy()

    proyectos = {}
    for nombre, g in df.groupby(name_col):
        if not isinstance(nombre, str) or not nombre.strip():
            continue
        vice = g[vice_col].dropna()
        cod = float(g["CODIFICADO"].sum())
        dev = float(g["DEVENGADO"].sum())
        proyectos[nombre.strip()] = {
            "codificado": cod,
            "comprometido": float(g["COMPROMISO"].sum()),
            "devengado": dev,
            "precompromiso": float(g["PRECOMPROMISO"].sum()),
            "asignado": float(g["ASIGNADO"].sum()),
            "saldo_disponible": float(g["SALDO_DISPONIBLE"].sum()),
            "viceministerio": str(vice.iloc[0]) if not vice.empty else "Sin asignar",
            "pct_ejecucion": round(dev / cod * 100, 2) if cod else 0.0,
        }
    return proyectos


def load_presupuesto():
    def _loader():
        buf = descargar_excel(URL_PRESUPUESTO)
        if buf is None:
            return {}
        return procesar_presupuesto(buf)
    return cached_value("presupuesto", _loader)


# =============================================================================
# 7. PROCESAMIENTO: BANCO / PRETT
# =============================================================================


def procesar_fichas(buf: io.BytesIO) -> dict:
    df = pd.read_excel(buf, sheet_name="Banco Proyectos", header=None)
    data = df.iloc[1:]
    COL_PROY, COL_CUP, COL_PERIODO = 1, 2, 6
    COL_OBJ, COL_COMP = 8, 9
    COL_LOGROS, COL_NUDOS, COL_ACCIONES, COL_RESP = 10, 11, 15, 17

    # Detecta Dictamen y Oficio si la base los incorpora como encabezados.
    header_norm = {}
    for c in df.columns:
        for v in df[c].head(5).dropna().astype(str):
            nv = normalizar(v)
            if nv:
                header_norm[c] = nv
                break

    def localizar_columna(*tokens):
        toks = [normalizar(t) for t in tokens]
        for c, h in header_norm.items():
            if all(t in h for t in toks):
                return c
        return None

    col_dictamen = localizar_columna("dictamen")
    col_oficio = localizar_columna("oficio")

    fichas = {}
    for nombre, g in data.groupby(data[COL_PROY]):
        if not isinstance(nombre, str) or not nombre.strip():
            continue
        if nombre.strip().rstrip(":").upper() == "PROYECTO":
            continue
        primera = g.iloc[0]
        componentes = [str(c).strip() for c in g[COL_COMP].dropna().unique() if str(c).strip()]
        dictamen = str(primera[col_dictamen]).strip() if col_dictamen is not None and pd.notna(primera[col_dictamen]) else "—"
        oficio = str(primera[col_oficio]).strip() if col_oficio is not None and pd.notna(primera[col_oficio]) else "—"
        fichas[nombre.strip()] = {
            "cup": str(primera[COL_CUP]) if pd.notna(primera[COL_CUP]) else "—",
            "gerente": str(primera[COL_RESP]) if pd.notna(primera[COL_RESP]) else "—",
            "periodo_prioridad": str(primera[COL_PERIODO]) if pd.notna(primera[COL_PERIODO]) else "—",
            "objetivo": str(primera[COL_OBJ]).strip() if pd.notna(primera[COL_OBJ]) else "",
            "dictamen_prioridad": dictamen,
            "oficio": oficio,
            "componentes": componentes,
            "logros": parse_dated_entries(primera[COL_LOGROS]),
            "nudos_criticos": parse_dated_entries(primera[COL_NUDOS]),
            "acciones_gestion": parse_dated_entries(primera[COL_ACCIONES]),
        }
    return fichas


def procesar_prett(buf: io.BytesIO):
    """
    PRETT = información del PROYECTO.
    HOJA_RUTA = desagregación del PROYECTO en:
        Componente -> Macro -> Micro

    No se calculan ni completan valores inexistentes.
    Una celda vacía se conserva vacía y en la interfaz se representa con "—".
    """

    def clean(v):
        if pd.isna(v):
            return ""
        s = str(v).strip()
        return "" if s.lower() in ("nan", "none", "nat") else s

    def fmt_fecha(v):
        if pd.isna(v) or clean(v) == "":
            return ""
        dt = pd.to_datetime(v, errors="coerce")
        if pd.isna(dt):
            return clean(v)
        return dt.strftime("%d/%m/%Y")

    # ---------------------------------------------------------
    # CONEXION: solo para identificar el proyecto y respaldo
    # ---------------------------------------------------------
    buf.seek(0)
    conexion = pd.read_excel(buf, sheet_name="CONEXION", header=None)
    data = conexion.iloc[1:]

    if data.empty:
        return None, None, None

    fila = data.iloc[0]
    nombre_conexion = clean(fila.iloc[0]) if len(fila) > 0 else ""

    # ---------------------------------------------------------
    # PRETT: datos generales de la ficha
    # ---------------------------------------------------------
    buf.seek(0)
    prett = pd.read_excel(buf, sheet_name="PRETT", header=None)

    def cell_norm(v):
        return normalizar(clean(v))

    def buscar_valor_derecha(*terminos):
        terms = [normalizar(t) for t in terminos]

        for r in range(len(prett)):
            row = prett.iloc[r]

            for c in range(len(row)):
                nv = cell_norm(row.iloc[c])
                if not nv:
                    continue

                if all(t in nv for t in terms):
                    for cc in range(c + 1, len(row)):
                        val = clean(row.iloc[cc])
                        if val:
                            return val

        return ""

    def buscar_texto_debajo(*terminos):
        terms = [normalizar(t) for t in terminos]

        for r in range(len(prett)):
            row = prett.iloc[r]
            found = False

            for c in range(len(row)):
                nv = cell_norm(row.iloc[c])
                if nv and all(t in nv for t in terms):
                    found = True
                    break

            if not found:
                continue

            for rr in range(r + 1, min(r + 8, len(prett))):
                valores = [clean(v) for v in prett.iloc[rr].tolist()]
                valores = [v for v in valores if v]

                if not valores:
                    continue

                texto = " ".join(valores).strip()
                nt = normalizar(texto)

                # No tomar el siguiente encabezado como contenido.
                if (
                    "desagregacion saldo disponible" in nt
                    or "presupuesto del proyecto" in nt
                    or "datos generales del proyecto" in nt
                ):
                    break

                return texto

        return ""

    nombre_prett = buscar_valor_derecha("proyecto")
    nombre_proyecto = nombre_prett or nombre_conexion

    # Este campo es ÚNICO en la ficha:
    # "Oficio último Dictamen de Prioridad o actualización de prioridad"
    oficio_ultimo_dictamen = buscar_valor_derecha(
        "oficio", "dictamen", "prioridad"
    )

    fecha_emision_raw = buscar_valor_derecha("fecha", "emision")
    fecha_emision = fmt_fecha(fecha_emision_raw)

    objetivo_prett = buscar_valor_derecha("objetivo", "general")

    ficha = {
        "cup": buscar_valor_derecha("cup"),
        "no_esigef": buscar_valor_derecha("no esigef"),
        "gerente": buscar_valor_derecha("gerente", "responsable"),
        "periodo_prioridad": buscar_valor_derecha("periodo", "prioridad"),
        "unidad_responsable": buscar_valor_derecha("unidad", "responsable"),
        "objetivo": objetivo_prett,
        "oficio_ultimo_dictamen": oficio_ultimo_dictamen,
        "fecha_emision_dictamen": fecha_emision,
        "devengado_a_fecha": buscar_texto_debajo("devengado", "fecha"),
        "desagregacion_saldo_disponible": buscar_texto_debajo(
            "desagregacion", "saldo", "disponible"
        ),
        "es_prett": True,
    }

    # ---------------------------------------------------------
    # HOJA_RUTA
    # ---------------------------------------------------------
    buf.seek(0)
    hr = pd.read_excel(buf, sheet_name="HOJA_RUTA", header=0)
    hr.columns = [
        str(c).replace("\n", " ").strip()
        for c in hr.columns
    ]

    def find_col(*aliases):
        aliases_n = [normalizar(a.replace("_", " ")) for a in aliases]

        for c in hr.columns:
            nc = normalizar(str(c).replace("_", " ").replace("\n", " "))
            if any(a == nc for a in aliases_n):
                return c

        # Segundo intento por inclusión, para tolerar "(automático)".
        for c in hr.columns:
            nc = normalizar(str(c).replace("_", " ").replace("\n", " "))
            if any(a in nc or nc in a for a in aliases_n):
                return c

        return None

    C_PROY = find_col("ID Proyecto", "ID_Proyecto")
    C_TIPO = find_col("Tipo Registro", "Tipo_Registro")
    C_ID = find_col("ID Registro", "ID_Registro")
    C_IDBI = find_col("ID BI", "ID_BI")
    C_COMP = find_col("Componente")
    C_NOMBRE = find_col(
        "Nombre Actividad / Tarea / Descripción",
        "Nombre_Actividad / Tarea / Descripción",
        "Nombre Actividad Tarea Descripción",
    )
    C_INICIO = find_col("Fecha Inicio", "Fecha_Inicio")
    C_FIN = find_col("Fecha Fin", "Fecha_Fin")
    C_EFECTIVA = find_col(
        "Fecha efectiva finalización",
        "Fecha_efectiva_finalización",
    )
    C_AVANCE = find_col(
        "% Avance (automático)",
        "Avance automático",
    )
    C_ESTADO = find_col("Estado")
    C_RESP = find_col("Responsable")
    C_ATENCION = find_col(
        "Requiere Atención Ministra",
        "Requiere_Atención Ministra",
    )
    C_OBS = find_col("Observaciones")
    C_ALERTA = find_col("Alerta de vencimiento")

    required = [C_TIPO, C_ID, C_COMP, C_NOMBRE]

    # Si faltan columnas estructurales, devolvemos la hoja como no utilizable,
    # pero NO inventamos datos.
    if any(c is None for c in required):
        return nombre_proyecto, ficha, {
            "componentes": [],
            "columnas_detectadas": list(hr.columns),
        }

    def raw(row, col):
        if col is None or col not in row.index:
            return ""
        return clean(row[col])

    def avance_text(row):
        if C_AVANCE is None or pd.isna(row[C_AVANCE]):
            return ""

        v = row[C_AVANCE]

        # Si Excel/pandas ya devuelve texto como 100%, lo respetamos.
        if isinstance(v, str):
            return clean(v)

        try:
            x = float(v)
            if x <= 1:
                return f"{x * 100:.0f}%"
            return f"{x:.0f}%"
        except Exception:
            return clean(v)

    # Separación exacta según Tipo_Registro.
    tipo = hr[C_TIPO].astype(str).map(normalizar)
    macros_df = hr[tipo.eq("macro")].copy()
    micros_df = hr[tipo.eq("micro")].copy()

    componentes = []

    component_names = [
        clean(x)
        for x in hr[C_COMP].dropna().tolist()
        if clean(x)
    ]
    # Orden de aparición, sin inventar ni ordenar alfabéticamente.
    component_names = list(dict.fromkeys(component_names))

    for comp_name in component_names:
        comp_macro = macros_df[
            macros_df[C_COMP].astype(str).str.strip().eq(comp_name)
        ].copy()

        comp_micro = micros_df[
            micros_df[C_COMP].astype(str).str.strip().eq(comp_name)
        ].copy()

        macros = []

        for _, m in comp_macro.iterrows():
            macro_id = raw(m, C_ID)

            # La relación Macro -> Micro se hace con ID_BI, como en la hoja.
            if C_IDBI is not None:
                hijos = comp_micro[
                    comp_micro[C_IDBI]
                    .astype(str)
                    .str.strip()
                    .eq(macro_id)
                ].copy()
            else:
                hijos = pd.DataFrame(columns=comp_micro.columns)

            micros = []

            for _, mi in hijos.iterrows():
                micros.append({
                    "id_proyecto": raw(mi, C_PROY),
                    "tipo_registro": raw(mi, C_TIPO),
                    "id_registro": raw(mi, C_ID),
                    "id_bi": raw(mi, C_IDBI),
                    "componente": raw(mi, C_COMP),
                    "nombre": raw(mi, C_NOMBRE),
                    "fecha_inicio": fmt_fecha(mi[C_INICIO]) if C_INICIO else "",
                    "fecha_fin": fmt_fecha(mi[C_FIN]) if C_FIN else "",
                    "fecha_efectiva_finalizacion": (
                        fmt_fecha(mi[C_EFECTIVA]) if C_EFECTIVA else ""
                    ),
                    "avance": avance_text(mi),
                    "estado": raw(mi, C_ESTADO),
                    "responsable": raw(mi, C_RESP),
                    "requiere_atencion_ministra": raw(mi, C_ATENCION),
                    "observaciones": raw(mi, C_OBS),
                    "alerta_vencimiento": raw(mi, C_ALERTA),
                })

            macros.append({
                "id_proyecto": raw(m, C_PROY),
                "tipo_registro": raw(m, C_TIPO),
                "id_registro": macro_id,
                "id_bi": raw(m, C_IDBI),
                "componente": raw(m, C_COMP),
                "nombre": raw(m, C_NOMBRE),
                "fecha_inicio": fmt_fecha(m[C_INICIO]) if C_INICIO else "",
                "fecha_fin": fmt_fecha(m[C_FIN]) if C_FIN else "",
                "fecha_efectiva_finalizacion": (
                    fmt_fecha(m[C_EFECTIVA]) if C_EFECTIVA else ""
                ),
                "avance": avance_text(m),
                "estado": raw(m, C_ESTADO),
                "responsable": raw(m, C_RESP),
                "requiere_atencion_ministra": raw(m, C_ATENCION),
                "observaciones": raw(m, C_OBS),
                "alerta_vencimiento": raw(m, C_ALERTA),
                "micros": micros,
            })

        componentes.append({
            "nombre": comp_name,
            "macros": macros,
        })

    return nombre_proyecto, ficha, {
        "componentes": componentes,
        "columnas_detectadas": list(hr.columns),
    }


def load_details_for_projects(proyectos: dict):
    def _loader_raw():
        fichas_raw = {}
        prett_raw = []
        b = descargar_excel(URL_BANCO_PROYECTOS)
        if b is not None:
            try:
                fichas_raw = procesar_fichas(b)
            except Exception:
                fichas_raw = {}
        for url in URLS_PRETT:
            b = descargar_excel(url)
            if b is None:
                continue
            try:
                result = procesar_prett(b)
                if result and result[0]:
                    prett_raw.append(result)
            except Exception:
                continue
        return fichas_raw, prett_raw

    fichas_raw, prett_raw = cached_value("details_raw", _loader_raw)
    nombres_norm = {normalizar(p): p for p in proyectos}
    fichas, roadmaps = {}, {}
    for nombre_banco, ficha in fichas_raw.items():
        real = match_project(nombre_banco, nombres_norm)
        if real:
            fichas[real] = ficha
    for nombre_excel, ficha_prett, roadmap in prett_raw:
        real = match_project(nombre_excel, nombres_norm)

        if real:
            # Para el proyecto que tiene ficha PRETT usamos únicamente
            # los campos que efectivamente provienen de esa ficha.
            fichas[real] = ficha_prett
            roadmaps[real] = roadmap

    return fichas, roadmaps


# =============================================================================
# 8. EJECUCIÓN MENSUAL / GANTT
# =============================================================================


def _leer_ejecucion_crudo(buf: io.BytesIO) -> pd.DataFrame:
    """
    Lee solo las 5 columnas necesarias de la hoja EJECUCION (~108,000 filas,
    54 columnas). Usa python-calamine si está instalado (motor en Rust,
    varias veces más rápido para archivos grandes); si no está disponible,
    cae de vuelta a openpyxl en modo read_only, que sigue siendo más
    liviano que pd.read_excel normal para un archivo de este tamaño.
    """
    necesarias = ["DATE", "NOMBRE DEL PROYECTO", "VICEMINISTERIO PROYECTOS", "CODIFICADO", "DEVENGADO"]

    if _HAS_CALAMINE:
        buf.seek(0)
        df = pd.read_excel(
            buf,
            sheet_name="EJECUCION",
            engine="calamine",
            usecols=lambda c: str(c).strip() in necesarias,
        )
        df.columns = [str(c).strip() for c in df.columns]
        df = df[df["NOMBRE DEL PROYECTO"].astype(str).str.strip().str.lower().ne("corriente")].copy()
        return df

    buf.seek(0)
    wb = openpyxl.load_workbook(buf, read_only=True, data_only=True)
    ws = wb["EJECUCION"]

    filas = ws.iter_rows(values_only=True)
    encabezados = [str(c).strip() if c is not None else "" for c in next(filas)]

    indices = {}
    for nombre in necesarias:
        for i, h in enumerate(encabezados):
            if h == nombre:
                indices[nombre] = i
                break

    faltantes = [n for n in necesarias if n not in indices]
    if faltantes:
        wb.close()
        raise ValueError(f"Columnas no encontradas en EJECUCION: {faltantes}")

    registros = []
    for fila in filas:
        nombre_proy = fila[indices["NOMBRE DEL PROYECTO"]]
        if not nombre_proy or str(nombre_proy).strip().lower() == "corriente":
            continue
        registros.append({
            "DATE": fila[indices["DATE"]],
            "NOMBRE DEL PROYECTO": str(nombre_proy).strip(),
            "VICEMINISTERIO PROYECTOS": fila[indices["VICEMINISTERIO PROYECTOS"]],
            "CODIFICADO": fila[indices["CODIFICADO"]] or 0,
            "DEVENGADO": fila[indices["DEVENGADO"]] or 0,
        })
    wb.close()
    return pd.DataFrame(registros)


def procesar_ejecucion_mensual(buf: io.BytesIO) -> pd.DataFrame:
    df = _leer_ejecucion_crudo(buf)
    df["DATE"] = pd.to_datetime(df["DATE"], errors="coerce")
    df = df.dropna(subset=["DATE"])

    mensual = (
        df.groupby(["NOMBRE DEL PROYECTO", "DATE"], as_index=False)
        .agg(
            codificado=("CODIFICADO", "sum"),
            devengado=("DEVENGADO", "sum"),
            vice=("VICEMINISTERIO PROYECTOS", "first"),
        )
        .sort_values(["NOMBRE DEL PROYECTO", "DATE"])
    )
    mensual = mensual[(mensual["codificado"] > 0) | (mensual["devengado"] > 0)].copy()

    mensual["devengado_mes"] = (
        mensual.groupby("NOMBRE DEL PROYECTO")["devengado"]
        .diff()
        .fillna(mensual["devengado"])
        .clip(lower=0)
    )
    mensual["estado_mes"] = mensual["devengado_mes"].gt(0).map(
        {True: "Con ejecución", False: "Sin ejecución"}
    )
    mensual["pct_corte"] = (
        mensual["devengado"]
        .div(mensual["codificado"].replace(0, pd.NA))
        .mul(100)
        .fillna(0)
        .clip(lower=0)
        .round(1)
    )
    return mensual.rename(columns={"NOMBRE DEL PROYECTO": "proyecto", "DATE": "fecha"})


def _guardar_cache_ejecucion(df: pd.DataFrame):
    """Guarda el Gantt ya procesado para que el siguiente arranque sea rápido."""
    if df is None or df.empty:
        return
    try:
        tmp = _EJECUCION_CACHE_FILE + ".tmp"
        df.to_pickle(tmp)
        os.replace(tmp, _EJECUCION_CACHE_FILE)
    except Exception:
        pass


def _leer_cache_ejecucion() -> pd.DataFrame:
    """Lee la última versión procesada disponible en disco."""
    try:
        if os.path.exists(_EJECUCION_CACHE_FILE):
            df = pd.read_pickle(_EJECUCION_CACHE_FILE)
            if isinstance(df, pd.DataFrame) and not df.empty:
                return df
    except Exception:
        pass
    return pd.DataFrame()


def _descargar_y_procesar_ejecucion() -> pd.DataFrame:
    """
    Consulta SharePoint solo cuando corresponde refrescar. Si la descarga
    funciona, actualiza también la copia local procesada.
    """
    b = descargar_excel(URL_EJECUCION_MENSUAL)
    if b is None:
        return pd.DataFrame()

    df = procesar_ejecucion_mensual(b)
    if df is not None and not df.empty:
        _guardar_cache_ejecucion(df)
    return df


def load_ejecucion(proyectos: dict) -> pd.DataFrame:
    """
    Prioridad:
    1. RAM (respuesta instantánea durante la sesión).
    2. Copia procesada en disco (respuesta rápida al reabrir la app).
    3. SharePoint (solo si todavía no existe una copia).
    """
    with _CACHE_LOCK:
        item = _CACHE.get("ejecucion_mensual")
        if item and item[1] is not None and not item[1].empty:
            df = item[1]
        else:
            df = None

    if df is None:
        local = _leer_cache_ejecucion()
        if not local.empty:
            df = local
            with _CACHE_LOCK:
                _CACHE["ejecucion_mensual"] = (time.time(), df)
        else:
            df = _descargar_y_procesar_ejecucion()
            if df is None or df.empty:
                return pd.DataFrame()
            with _CACHE_LOCK:
                _CACHE["ejecucion_mensual"] = (time.time(), df)

    nombres_norm = {normalizar(p): p for p in proyectos}
    out = df.copy()
    out["proyecto"] = out["proyecto"].apply(
        lambda n: match_project(n, nombres_norm) or str(n).strip()
    )
    return out


def refrescar_ejecucion_en_segundo_plano():
    """
    Actualiza SharePoint sin bloquear la interfaz. El usuario puede trabajar
    con la última copia disponible mientras se termina la sincronización.
    """
    try:
        nuevo = _descargar_y_procesar_ejecucion()
        if nuevo is not None and not nuevo.empty:
            with _CACHE_LOCK:
                _CACHE["ejecucion_mensual"] = (time.time(), nuevo)
    except Exception:
        pass


def gantt_figure(df: pd.DataFrame, vice_selected: str) -> go.Figure:
    if df is None or df.empty:
        fig = go.Figure()
        fig.add_annotation(
            text="No se pudo cargar EJECUCION_MENSUAL",
            showarrow=False, x=.5, y=.5,
        )
        fig.update_layout(height=260, xaxis_visible=False, yaxis_visible=False)
        return fig

    data = df.copy()
    data["vice_short"] = data["vice"].apply(vice_short)
    data = data[data["vice_short"] == vice_selected].copy()

    if data.empty:
        fig = go.Figure()
        fig.add_annotation(
            text="Sin registros para este viceministerio",
            showarrow=False, x=.5, y=.5,
        )
        fig.update_layout(height=260, xaxis_visible=False, yaxis_visible=False)
        return fig

    data["mes_inicio"] = data["fecha"].dt.to_period("M").dt.start_time

    order = (
        data.groupby("proyecto")["fecha"]
        .max()
        .sort_values(ascending=False)
        .index
        .tolist()
    )

    summary = data.groupby("proyecto").agg(
        total_meses=("estado_mes", "size"),
        meses_con_ejecucion=(
            "estado_mes",
            lambda s: int((s == "Con ejecución").sum()),
        ),
    )

    fig = go.Figure()
    colors = {
        "Con ejecución": PURPLE_700,
        "Sin ejecución": PURPLE_300,
    }

    xmin = data["mes_inicio"].min()
    xmax = data["mes_inicio"].max() + pd.offsets.MonthBegin(2)

    # -----------------------------------------------------------------
    # ETIQUETAS CLICABLES
    # -----------------------------------------------------------------
    # En lugar de depender de los ticks del eje Y, que Plotly no devuelve
    # como clickData, se dibujan los nombres como un trace de texto.
    # Ese texto sí genera clickData y lleva el nombre del proyecto.
    fig.add_trace(go.Scatter(
        x=[xmin] * len(order),
        y=order,
        mode="text",
        text=order,
        textposition="middle left",
        textfont=dict(
            size=11,
            color=PURPLE_800,
            family="Segoe UI",
        ),
        customdata=[[project, "__PROJECT_LABEL__"] for project in order],
        hovertemplate="<b>%{customdata[0]}</b><br>Haz clic para ver el detalle<extra></extra>",
        showlegend=False,
        cliponaxis=False,
        name="Proyecto",
    ))

    # -----------------------------------------------------------------
    # SEGMENTOS MENSUALES
    # -----------------------------------------------------------------
    for estado in ["Con ejecución", "Sin ejecución"]:
        sub = data[data["estado_mes"] == estado].copy()
        if sub.empty:
            continue

        durations = []
        custom = []

        for _, row in sub.iterrows():
            ini = pd.Timestamp(row["mes_inicio"])
            fin = ini + pd.offsets.MonthBegin(1)
            durations.append(int((fin - ini).total_seconds() * 1000))
            custom.append([
                row["proyecto"],
                pd.Timestamp(row["fecha"]).strftime("%b %Y"),
                float(row["codificado"]),
                float(row["devengado"]),
                float(row["devengado_mes"]),
                float(row["pct_corte"]),
                estado,
                "__MONTH_SEGMENT__",
            ])

        fig.add_trace(go.Bar(
            name=estado,
            x=durations,
            base=sub["mes_inicio"].tolist(),
            y=sub["proyecto"].tolist(),
            orientation="h",
            marker=dict(
                color=colors[estado],
                line=dict(color="white", width=1),
            ),
            width=.62,
            customdata=custom,
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "Mes: %{customdata[1]}<br>"
                "Estado mensual: %{customdata[6]}<br>"
                "Codificado al corte: $ %{customdata[2]:,.2f}<br>"
                "Devengado acumulado: $ %{customdata[3]:,.2f}<br>"
                "Devengado del mes: $ %{customdata[4]:,.2f}<br>"
                "Ejecución al corte: %{customdata[5]:.1f}%"
                "<extra></extra>"
            ),
        ))

    # Porcentaje de meses con ejecución
    for project in order:
        r = summary.loc[project]
        total = int(r["total_meses"])
        executed = int(r["meses_con_ejecucion"])
        pct = 100 * executed / total if total else 0

        fig.add_annotation(
            x=xmax - pd.Timedelta(days=4),
            y=project,
            text=f"{pct:.0f}% ({executed}/{total})",
            showarrow=False,
            xanchor="right",
            font=dict(size=10, color=INK),
        )

    fig.update_layout(
        height=max(330, 62 * len(order) + 110),

        # Margen izquierdo amplio: aquí se muestran los nombres clicables.
        margin=dict(l=330, r=16, t=50, b=35),

        xaxis=dict(
            type="date",
            range=[xmin, xmax],
            tickformat="%b",
            dtick="M1",
            gridcolor="#E9EBF3",
        ),

        # Ocultamos los labels normales porque ahora los nombres clicables
        # los dibuja el trace Scatter.
        yaxis=dict(
            categoryorder="array",
            categoryarray=order,
            autorange="reversed",
            showticklabels=False,
            automargin=False,
        ),

        barmode="overlay",
        legend=dict(
            orientation="h",
            y=1.08,
            x=.28,
        ),
        plot_bgcolor="white",
        paper_bgcolor="white",
        clickmode="event+select",
        hovermode="closest",
    )

    return fig



def _timeline_months(data: pd.DataFrame) -> list[pd.Timestamp]:
    if data is None or data.empty:
        return []
    vals = (
        data["fecha"].dt.to_period("M")
        .dropna().drop_duplicates().sort_values().tolist()
    )
    return [pd.Period(v, freq="M").to_timestamp() for v in vals]


def _timeline_hover(row) -> str:
    return (
        f"{pd.Timestamp(row['fecha']).strftime('%b %Y')} | "
        f"{str(row.get('estado_mes', ''))} | "
        f"Codificado: $ {float(row.get('codificado', 0)):,.2f} | "
        f"Devengado acumulado: $ {float(row.get('devengado', 0)):,.2f} | "
        f"Devengado del mes: $ {float(row.get('devengado_mes', 0)):,.2f} | "
        f"Ejecución al corte: {float(row.get('pct_corte', 0)):.1f}%"
    )


def timeline_matrix_component(df: pd.DataFrame, vice_selected: str):
    if df is None or df.empty:
        return html.Div("La ejecución mensual aún no está disponible. El resto del panel puede seguir utilizándose mientras se actualizan los datos.", className="alert")

    data = df.copy()
    data["vice_short"] = data["vice"].apply(vice_short)
    data = data[data["vice_short"] == vice_selected].copy()

    if data.empty:
        return html.Div("Sin registros para este viceministerio.", className="alert")

    data["month_start"] = data["fecha"].dt.to_period("M").dt.start_time
    months = _timeline_months(data)
    if len(months) > 8:
        months = months[-8:]

    data = data[data["month_start"].isin(set(months))].copy()

    order = (
        data.groupby("proyecto")["fecha"]
        .max().sort_values(ascending=False).index.tolist()
    )

    month_labels = {1:"Ene",2:"Feb",3:"Mar",4:"Abr",5:"May",6:"Jun",7:"Jul",8:"Ago",9:"Sep",10:"Oct",11:"Nov",12:"Dic"}

    legend = html.Div([
        html.Span([html.Span(className="timeline-legend-box exec"), "Con ejecución"], className="timeline-legend-item"),
        html.Span([html.Span(className="timeline-legend-box noexec"), "Sin ejecución"], className="timeline-legend-item"),
    ], className="timeline-legend")

    grid = [html.Div("Proyecto", className="timeline-head project-col")]
    for m in months:
        grid.append(html.Div(month_labels[m.month], className="timeline-head"))
    grid.extend([
        html.Div("Avance", className="timeline-head"),
        html.Div("Segmentos", className="timeline-head"),
    ])

    for row_idx, project in enumerate(order):
        sub = data[data["proyecto"] == project].sort_values("month_start").copy()
        by_month = {pd.Timestamp(r["month_start"]): r for _, r in sub.iterrows()}
        executed = int((sub["estado_mes"] == "Con ejecución").sum())
        total = int(len(sub))
        pct = 100 * executed / total if total else 0

        grid.append(html.Div(
            html.Button(
                project,
                id={"type":"timeline-project","index":row_idx},
                value=project,
                n_clicks=0,
                className="timeline-project-button",
                title="Ver detalle del proyecto",
            ),
            className="timeline-project-cell"
        ))

        for m in months:
            r = by_month.get(pd.Timestamp(m))
            if r is None:
                cls = "timeline-segment empty"
                title = "Sin registro para este mes"
            else:
                cls = "timeline-segment exec" if str(r.get("estado_mes")) == "Con ejecución" else "timeline-segment noexec"
                title = _timeline_hover(r)
            grid.append(html.Div(html.Div(className=cls, title=title), className="timeline-month-cell"))

        grid.append(html.Div([
            html.Div(f"{pct:.0f}%", className="timeline-progress-label"),
            html.Div(
                html.Div(className="timeline-mini-fill", style={"width":f"{max(0,min(pct,100)):.1f}%"}),
                className="timeline-mini-track"
            ),
        ], className="timeline-progress-cell"))

        grid.append(html.Div(f"{executed} / {total}", className="timeline-segments-cell"))

    template = f"minmax(310px,1.9fr) repeat({len(months)}, minmax(72px,.6fr)) 112px 86px"

    return html.Div([
        legend,
        html.Div(
            html.Div(grid, className="timeline-grid", style={"gridTemplateColumns":template}),
            className="timeline-matrix"
        ),
        html.Div(
            "Haga clic sobre el nombre del proyecto para consultar su detalle. "
            "Pase el mouse sobre cada mes para revisar el movimiento presupuestario.",
            className="timeline-tip"
        ),
    ])


# =============================================================================
# 9. FIGURAS / COMPONENTES UI
# =============================================================================


def aggregate_kpis(proyectos: dict):
    n = len(proyectos)
    cod = sum(p["codificado"] for p in proyectos.values())
    comp = sum(p["comprometido"] for p in proyectos.values())
    dev = sum(p["devengado"] for p in proyectos.values())
    disp = sum(p["saldo_disponible"] for p in proyectos.values())
    pct = round(dev / cod * 100, 1) if cod else 0
    return n, cod, comp, dev, disp, pct


def kpi_card(icon, value, label, bg):
    return html.Div([
        html.Div(icon, className="kpi-icon", style={"background": bg}),
        html.Div(value, className="kpi-value"),
        html.Div(label, className="kpi-label"),
    ], className="kpi-card")


def render_kpis(proyectos):
    n, cod, comp, dev, disp, pct = aggregate_kpis(proyectos)
    values = [
        ("📁", str(n), "Proyectos", PURPLE_100),
        ("💵", fmt_mm(cod), "Codificado", "#E9F3FF"),
        ("📝", fmt_mm(comp), "Comprometido", "#F1ECFF"),
        ("✅", fmt_mm(dev), "Devengado", "#E8F7F0"),
        ("👛", fmt_mm(disp), "Disponible", "#FFF3E3"),
        ("📊", f"{pct}%", "% Ejecución", PURPLE_100),
    ]
    return html.Div([kpi_card(*x) for x in values], className="kpi-grid")


def presupuesto_vice_figure(proyectos):
    rows = []
    for p in proyectos.values():
        rows.append({
            "vice": vice_short(p.get("viceministerio")),
            "asignado": p["codificado"],
            "ejecutado": p["devengado"],
        })
    df = pd.DataFrame(rows)
    if df.empty:
        return go.Figure()
    df = df.groupby("vice", as_index=False)[["asignado", "ejecutado"]].sum().sort_values("asignado")
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=df["asignado"] / 1e6, y=df["vice"], orientation="h",
        name="Presupuesto asignado", marker_color=PURPLE_300,
        text=[f"$ {v/1e6:,.1f} MM" for v in df["asignado"]], textposition="outside",
    ))
    fig.add_trace(go.Bar(
        x=df["ejecutado"] / 1e6, y=df["vice"], orientation="h",
        name="Ejecutado", marker_color=PURPLE_700,
        text=[f"$ {v/1e6:,.1f} MM" for v in df["ejecutado"]], textposition="outside",
    ))
    fig.update_layout(
        barmode="group", height=325, margin=dict(l=0, r=80, t=20, b=10),
        xaxis=dict(visible=False), legend=dict(orientation="h", y=1.12, x=.08),
        plot_bgcolor="white", paper_bgcolor="white",
    )
    return fig


def header_block(title="Panel de seguimiento de proyectos de inversión — MINEDEC", subtitle="Presupuesto general y control de actividades"):
    return html.Div([
        html.Div([
            html.Div(title, className="top-title"),
            html.Div(subtitle, className="top-subtitle"),
        ], className="top-gradient"),
        html.Div([
            html.Div("Actualizado", className="updated-label"),
            html.Div(datetime.now().strftime("%d/%m/%Y"), className="updated-value"),
        ], className="updated-card"),
    ], className="header-card")


def sidebar():
    return html.Aside([
        html.Div("‹‹", className="collapse-symbol"),
        html.Div([
            html.Img(src=logo_src(), className="sidebar-logo", id="sidebar-logo"),
        ], className="logo-box"),
        html.H3("Panel de Inversión", className="sidebar-title"),
        html.P("Seguimiento institucional", className="sidebar-subtitle"),
        html.Nav([
            dcc.Link("⌂  Inicio", href="/inicio", className="nav-button"),
            dcc.Link("▦  Proyectos", href="/proyectos", className="nav-button"),
        ]),
        html.Div(className="sidebar-spacer"),
        html.P(f"Actualizado {datetime.now().strftime('%d/%m/%Y')}", className="sidebar-date"),
        html.Button("↪ Cerrar sesión", id="logout-button", className="logout-button", n_clicks=0),
    ], className="sidebar")


def app_shell(content):
    return html.Div([
        sidebar(),
        html.Main(content, className="main-content"),
    ], className="app-shell")


def card(children, className="content-card"):
    return html.Div(children, className=className)


# =============================================================================
# 10. LOGIN
# =============================================================================


def login_layout(error_text=""):
    return html.Div([
        html.Div([
            html.Div([
                html.Img(src=logo_src(), className="login-logo"),
                html.H1("Seguimiento de proyectos de inversión", className="login-title"),
                html.P("Panel de Inversión", className="login-subtitle"),
                dcc.Input(
                    id="password-input", type="password", placeholder="Ingresa la contraseña",
                    className="password-input", debounce=False,
                ),
                html.Button("Ingresar", id="login-button", className="login-button", n_clicks=0),
                html.Div(error_text, id="login-error", className="login-error"),
                html.P("Acceso exclusivo para personal autorizado", className="login-foot"),
            ], className="login-card")
        ], className="login-center")
    ], className="login-page")


# =============================================================================
# 11. INICIO
# =============================================================================


def home_layout():
    proyectos = load_presupuesto()
    if not proyectos:
        return app_shell([
            header_block(),
            card(html.Div("No se pudo leer MINEDEC 2.0 desde SharePoint.", className="alert")),
        ])

    # La línea de tiempo necesita la base de ejecución mensual.
    # Si por cualquier motivo SharePoint/cache falla, el resto del dashboard
    # sigue funcionando y la sección temporal muestra un aviso en vez de
    # provocar un NameError o tumbar la página completa.
    try:
        ejecucion = load_ejecucion(proyectos)
    except Exception:
        ejecucion = pd.DataFrame()

    vices = sorted({vice_short(p.get("viceministerio")) for p in proyectos.values()})
    default_vice = vices[0] if vices else None

    return app_shell([
        header_block(),
        render_kpis(proyectos),
        card([
            html.H3("Presupuesto asignado vs. ejecutado por viceministerio", className="section-title"),
            html.P("Montos en millones USD", className="section-subtitle"),
            dcc.Graph(
                id="budget-vice-graph", figure=presupuesto_vice_figure(proyectos),
                config={"displayModeBar": False}, className="graph",
            ),
        ]),
        card([
            html.H3("Línea de tiempo por proyecto", className="section-title"),
            html.P(
                "Movimiento presupuestario mensual · morado oscuro: con ejecución en el mes · morado claro: sin ejecución en el mes",
                className="section-subtitle",
            ),
            dcc.RadioItems(
                id="vice-radio",
                options=[{"label": v, "value": v} for v in vices],
                value=default_vice,
                inline=True,
                className="vice-radio",
                inputClassName="radio-input",
                labelClassName="radio-label",
            ),
            html.Div(
                timeline_matrix_component(ejecucion, default_vice),
                id="timeline-matrix-container",
            ),
        ]),
        html.Div(
            id="home-project-detail",
            className="timeline-selected-wrap",
        ),
    ])


# =============================================================================
# 12. PROYECTOS
# =============================================================================


def project_row(name, p):
    return html.Div([
        html.Div("▦", className="project-icon"),
        html.Div([
            html.Div(name, className="project-name"),
            html.Span(vice_short(p.get("viceministerio")), className="soft-tag"),
        ], className="project-main"),
        html.Div([html.Strong(fmt_mm(p["codificado"])), html.Small("Codificado")], className="project-metric"),
        html.Div([html.Strong(f"{p['pct_ejecucion']}%"), html.Small("Ejecución")], className="project-metric"),
        dcc.Link("Ver detalle ›", href=f"/detalle?project={quote(name)}", className="detail-link"),
    ], className="project-row")


def projects_layout():
    proyectos = load_presupuesto()
    vices = sorted({vice_short(p.get("viceministerio")) for p in proyectos.values()})
    return app_shell([
        header_block("Proyectos de inversión", "Consulta y seguimiento individual de los proyectos"),
        html.Div([
            dcc.Input(id="project-search", placeholder="Buscar proyecto por nombre", className="search-input", debounce=True),
            dcc.Dropdown(
                id="project-vice-filter",
                options=[{"label": "Todos", "value": "Todos"}] + [{"label": v, "value": v} for v in vices],
                value="Todos", clearable=False, className="vice-dropdown",
            ),
        ], className="project-filters"),
        html.Div(id="project-list"),
    ])


# =============================================================================
# 13. DETALLE
# =============================================================================


def budget_detail_fig(p):
    labels = ["Codificado", "Comprometido", "Devengado", "Disponible"]
    values = [p["codificado"], p["comprometido"], p["devengado"], p["saldo_disponible"]]
    fig = go.Figure(go.Bar(
        x=values, y=labels, orientation="h",
        marker_color=[PURPLE_300, PURPLE_500, PURPLE_700, PURPLE_100],
        text=[fmt_money(v) for v in values], textposition="outside",
    ))
    fig.update_layout(height=280, margin=dict(l=0, r=120, t=10, b=10), xaxis_visible=False, yaxis_autorange="reversed", plot_bgcolor="white", paper_bgcolor="white")
    return fig


def composition_fig(p):
    cod, dev = float(p["codificado"]), float(p["devengado"])
    pending = max(cod - dev, 0)
    fig = go.Figure(go.Pie(
        labels=["Devengado", "Pendiente de devengar"], values=[dev, pending], hole=.72,
        marker_colors=[PURPLE_700, PURPLE_100], textinfo="percent",
    ))
    fig.update_layout(
        height=280, margin=dict(l=0, r=0, t=10, b=10),
        annotations=[dict(text=f"{p['pct_ejecucion']}%<br>ejecutado", x=.5, y=.5, showarrow=False, font_size=16)],
    )
    return fig


def dated_list(title, entries):
    return html.Details([
        html.Summary(f"{title} ({len(entries)})"),
        html.Div([
            html.P([html.Strong(f"{e['fecha']} — "), e["texto"]]) for e in reversed(entries)
        ] if entries else [html.P("No hay registros con fecha.")], className="details-body"),
    ], className="native-details")


def _display(v):
    """Representación visual de una celda vacía, sin inventar contenido."""
    if v is None:
        return "—"
    s = str(v).strip()
    return s if s and s.lower() not in ("nan", "none", "nat") else "—"


def _status_class(value):
    n = normalizar(value)

    if any(k in n for k in ["finalizada", "completada", "completado"]):
        return "status-pill done"

    if "vencida" in n:
        return "status-pill due"

    if any(k in n for k in ["revision", "actualizacion", "espera", "pendiente"]):
        return "status-pill review"

    return "status-pill"


def _route_field(label, value, wide=False):
    return html.Div(
        [
            html.Div(label, className="route-field-label"),
            html.Div(_display(value), className="route-field-value"),
        ],
        className=("route-field route-field-wide" if wide else "route-field"),
    )



def _route_value(value):
    """Muestra exactamente el dato leído; vacío = raya visual."""
    if value is None:
        return "—"
    s = str(value).strip()
    return s if s and s.lower() not in ("nan", "none", "nat") else "—"


def _route_metric(label, value, wide=False):
    return html.Div(
        [
            html.Div(label, className="cascade-metric-label"),
            html.Div(_route_value(value), className="cascade-metric-value"),
        ],
        className=(
            "cascade-metric cascade-metric-wide"
            if wide
            else "cascade-metric"
        ),
    )


def _macro_summary_card(macro):
    """Resumen ejecutivo del Macro seleccionado."""
    alerta = _route_value(macro.get("alerta_vencimiento"))

    return html.Div(
        [
            html.Div(
                [
                    html.Div(
                        [
                            html.Span("MACRO", className="cascade-tag"),
                            html.Span(
                                _route_value(macro.get("id_registro")),
                                className="cascade-code",
                            ),
                        ],
                        className="cascade-eyebrow",
                    ),
                    html.H4(
                        _route_value(macro.get("nombre")),
                        className="cascade-selected-title",
                    ),
                ]
            ),
            html.Span(
                alerta,
                className=_status_class(alerta),
            ),
        ],
        className="cascade-selected-head",
    )


def _route_detail_grid(item):
    """
    Campos reales de HOJA_RUTA.
    No calcula ni completa ninguna variable.
    """
    return html.Div(
        [
            _route_metric("Fecha inicio", item.get("fecha_inicio")),
            _route_metric("Fecha fin", item.get("fecha_fin")),
            _route_metric(
                "Fecha efectiva finalización",
                item.get("fecha_efectiva_finalizacion"),
            ),
            _route_metric("% Avance", item.get("avance")),
            _route_metric("Estado", item.get("estado")),
            _route_metric("Responsable", item.get("responsable")),
            _route_metric(
                "Requiere Atención Ministra",
                item.get("requiere_atencion_ministra"),
            ),
            _route_metric(
                "Alerta de vencimiento",
                item.get("alerta_vencimiento"),
            ),
            _route_metric(
                "Observaciones",
                item.get("observaciones"),
                wide=True,
            ),
        ],
        className="cascade-detail-grid",
    )


def _find_component(rm, component_name):
    for comp in (rm or {}).get("componentes", []):
        if str(comp.get("nombre", "")) == str(component_name):
            return comp
    return None


def _find_macro(rm, component_name, macro_id):
    comp = _find_component(rm, component_name)
    if not comp:
        return None

    for macro in comp.get("macros", []):
        if str(macro.get("id_registro", "")) == str(macro_id):
            return macro
    return None


def _find_micro(rm, component_name, macro_id, micro_id):
    macro = _find_macro(rm, component_name, macro_id)
    if not macro:
        return None

    for micro in macro.get("micros", []):
        if str(micro.get("id_registro", "")) == str(micro_id):
            return micro
    return None


def prett_section(rm):
    """
    Navegación ejecutiva:
        Componente -> Macro -> Micro -> detalle
    Nada se despliega masivamente al inicio.
    """
    componentes = (rm or {}).get("componentes", [])

    if not componentes:
        return html.Div(
            "No se encontraron registros de Componente/Macro/Micro en HOJA_RUTA.",
            className="info-alert",
        )

    component_options = [
        {
            "label": str(comp.get("nombre", "")),
            "value": str(comp.get("nombre", "")),
        }
        for comp in componentes
        if str(comp.get("nombre", "")).strip()
    ]

    return html.Div(
        [
            # Los datos ya procesados se guardan en el navegador para que
            # los cambios Componente/Macro/Micro sean instantáneos.
            dcc.Store(
                id="route-roadmap-store",
                data=rm,
                storage_type="memory",
            ),

            html.Div(
                [
                    html.Div(
                        [
                            html.Div(
                                "Seguimiento operativo",
                                className="cascade-kicker",
                            ),
                            html.H3(
                                "Desagregación por componente",
                                className="cascade-section-title",
                            ),
                            html.P(
                                "Seleccione un componente y avance por las etapas Macro "
                                "hasta consultar el detalle Micro.",
                                className="cascade-section-subtitle",
                            ),
                        ]
                    ),
                    html.Div(
                        [
                            html.Span(
                                str(len(component_options)),
                                className="cascade-count-number",
                            ),
                            html.Span(
                                "componente" if len(component_options) == 1 else "componentes",
                                className="cascade-count-label",
                            ),
                        ],
                        className="cascade-count",
                    ),
                ],
                className="cascade-header",
            ),

            # PASO 1
            html.Div(
                [
                    html.Div(
                        [
                            html.Span("1", className="cascade-step-number"),
                            html.Div(
                                [
                                    html.Div(
                                        "Componente",
                                        className="cascade-step-title",
                                    ),
                                    html.Div(
                                        "Seleccione el componente que desea revisar",
                                        className="cascade-step-help",
                                    ),
                                ]
                            ),
                        ],
                        className="cascade-step-heading",
                    ),
                    dcc.Dropdown(
                        id="route-component-select",
                        options=component_options,
                        value=None,
                        clearable=True,
                        searchable=False,
                        placeholder="Seleccione un componente",
                        className="cascade-dropdown",
                    ),
                ],
                className="cascade-step-card",
            ),

            # PASOS SIGUIENTES
            html.Div(id="route-macro-area"),
            html.Div(id="route-micro-area"),
            html.Div(id="route-activity-detail"),
        ],
        className="route-cascade-shell",
    )


def _strip_synthesis_date(texto):
    if not texto: return texto
    return re.sub(r"^\s*\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\s*[—–-]\s*", "", str(texto)).strip()

def _context_card(label,value):
    return html.Div([html.Div(label,className="context-label"),html.Div(value if value not in [None,"","nan"] else "—",className="context-value")],className="context-card")

def detail_component(name: str, embedded=False):
    proyectos = load_presupuesto()
    p = proyectos.get(name)

    if not p:
        return card(
            html.Div(
                "No se encontró el proyecto seleccionado.",
                className="alert",
            )
        )

    fichas, roadmaps = load_details_for_projects(proyectos)
    ficha = fichas.get(name)
    rm = roadmaps.get(name)

    es_prett = bool(
        ficha
        and ficha.get("es_prett")
    )

    blocks = []

    # ---------------------------------------------------------
    # ENCABEZADO
    # ---------------------------------------------------------
    if embedded:
        blocks.extend([
            html.H2(
                "Detalle del proyecto seleccionado",
                className="detail-main-title",
            ),
            html.H3(
                name,
                className="detail-project-name",
            ),
        ])
    else:
        blocks.extend([
            dcc.Link(
                "← Volver a proyectos",
                href="/proyectos",
                className="back-link",
            ),
            header_block(
                name,
                "Seguimiento consolidado del proyecto",
            ),
        ])

    # ---------------------------------------------------------
    # DATOS DEL PROYECTO
    # ---------------------------------------------------------
    if ficha:
        objetivo = ficha.get("objetivo", "")
        periodo = ficha.get("periodo_prioridad", "")

        # PRETT tiene un único campo:
        # "Oficio último Dictamen de Prioridad o actualización de prioridad".
        if ficha.get("es_prett"):
            blocks.append(
                html.Div(
                    [
                        html.Div(
                            [
                                html.Div(
                                    "Objetivo general",
                                    className="context-label",
                                ),
                                html.Div(
                                    _display(objetivo),
                                    className="context-value objective-text",
                                ),
                            ],
                            className="context-card context-objective",
                        ),

                        html.Div(
                            [
                                _context_card(
                                    "Oficio último Dictamen de Prioridad o actualización de prioridad",
                                    ficha.get("oficio_ultimo_dictamen", ""),
                                ),
                                _context_card(
                                    "Fecha de emisión",
                                    ficha.get("fecha_emision_dictamen", ""),
                                ),
                                _context_card(
                                    "Período de prioridad",
                                    periodo,
                                ),
                                _context_card(
                                    "Unidad responsable",
                                    ficha.get("unidad_responsable", ""),
                                ),
                            ],
                            className="project-context-secondary prett-context-secondary",
                        ),
                    ],
                    className="project-context-block",
                )
            )

        else:
            # Otros proyectos conservan exclusivamente lo que proviene
            # de Banco de Proyectos.
            blocks.append(
                html.Div(
                    [
                        html.Div(
                            [
                                html.Div(
                                    "Objetivo general",
                                    className="context-label",
                                ),
                                html.Div(
                                    _display(objetivo),
                                    className="context-value objective-text",
                                ),
                            ],
                            className="context-card context-objective",
                        ),
                        html.Div(
                            [
                                _context_card(
                                    "Dictamen de prioridad",
                                    ficha.get("dictamen_prioridad", ""),
                                ),
                                _context_card(
                                    "Oficio",
                                    ficha.get("oficio", ""),
                                ),
                                _context_card(
                                    "Período de prioridad",
                                    periodo,
                                ),
                            ],
                            className="project-context-secondary",
                        ),
                    ],
                    className="project-context-block",
                )
            )

    # ---------------------------------------------------------
    # VALORES PRESUPUESTARIOS
    # ---------------------------------------------------------
    blocks.append(
        html.Div(
            "Valores presupuestarios",
            className="project-values-title",
        )
    )

    blocks.append(
        html.Div(
            [
                kpi_card(
                    "💵",
                    fmt_money(p["codificado"]),
                    "Codificado",
                    "#E9F3FF",
                ),
                kpi_card(
                    "📝",
                    fmt_money(p["comprometido"]),
                    "Comprometido",
                    "#F1ECFF",
                ),
                kpi_card(
                    "✅",
                    fmt_money(p["devengado"]),
                    "Devengado",
                    "#E8F7F0",
                ),
                kpi_card(
                    "📊",
                    f"{p['pct_ejecucion']}%",
                    "Ejecución",
                    PURPLE_100,
                ),
            ],
            className="detail-kpi-grid",
        )
    )

    # ---------------------------------------------------------
    # DEVENGADO A LA FECHA
    # Solo PRETT usa esta sección específica de su ficha.
    # ---------------------------------------------------------
    if es_prett and ficha.get("devengado_a_fecha"):
        blocks.append(
            card(
                [
                    html.H3(
                        "Devengado a la fecha",
                        className="section-title",
                    ),
                    html.P(
                        "Síntesis descriptiva del devengado registrado en la ficha del proyecto",
                        className="section-subtitle",
                    ),
                    html.Div(
                        ficha.get("devengado_a_fecha"),
                        className="summary-box",
                    ),
                    html.Div(
                        [
                            html.Span(
                                [
                                    html.Strong("CUP: "),
                                    ficha.get("cup", "—"),
                                ]
                            ),
                            html.Span(
                                [
                                    html.Strong(
                                        "Gerente/responsable: "
                                    ),
                                    ficha.get(
                                        "gerente",
                                        "—",
                                    ),
                                ]
                            ),
                            html.Span(
                                [
                                    html.Strong(
                                        "Período de prioridad: "
                                    ),
                                    ficha.get(
                                        "periodo_prioridad",
                                        "—",
                                    ),
                                ]
                            ),
                        ],
                        className="meta-three",
                    ),
                ]
            )
        )

    else:
        # Para proyectos que no son PRETT se conserva la síntesis
        # de seguimiento disponible en Banco de Proyectos.
        synthesis = ""

        if ficha:
            if ficha.get("logros"):
                synthesis = _strip_synthesis_date(
                    ficha["logros"][-1].get(
                        "texto",
                        "",
                    )
                )
            elif ficha.get("objetivo"):
                synthesis = ficha.get(
                    "objetivo",
                    "",
                )

        state = [
            html.H3(
                "Estado de ejecución del proyecto",
                className="section-title",
            ),
            html.P(
                "Síntesis descriptiva de avances",
                className="section-subtitle",
            ),
            html.Div(
                synthesis
                or "No existe una síntesis descriptiva registrada.",
                className="summary-box",
            ),
        ]

        if ficha:
            state.append(
                html.Div(
                    [
                        html.Span(
                            [
                                html.Strong("CUP: "),
                                ficha.get("cup", "—"),
                            ]
                        ),
                        html.Span(
                            [
                                html.Strong(
                                    "Gerente/responsable: "
                                ),
                                ficha.get(
                                    "gerente",
                                    "—",
                                ),
                            ]
                        ),
                        html.Span(
                            [
                                html.Strong(
                                    "Período de prioridad: "
                                ),
                                ficha.get(
                                    "periodo_prioridad",
                                    "—",
                                ),
                            ]
                        ),
                    ],
                    className="meta-three",
                )
            )

        blocks.append(card(state))

    # ---------------------------------------------------------
    # COMPONENTE
    # El componente solo aparece si ESTE proyecto tiene roadmap.
    # ---------------------------------------------------------
    if rm and rm.get("componentes"):
        blocks.append(
            prett_section(rm)
        )

    elif es_prett:
        blocks.append(
            html.Div(
                "La ficha PRETT está relacionada con este proyecto, "
                "pero HOJA_RUTA no contiene componentes utilizables.",
                className="info-alert",
            )
        )

    return html.Div(
        blocks,
        className="detail-container",
    )


def detail_page(project_name: str):
    return app_shell(detail_component(project_name, embedded=False))


# =============================================================================
# 14. APP LAYOUT / ROUTING
# =============================================================================

app.layout = html.Div([
    dcc.Location(id="url", refresh="callback-nav"),
    dcc.Store(
        id="auth-store",
        storage_type="session",
        data={"authenticated": False},
    ),
    html.Div(id="page-content"),
])


@app.callback(
    Output("page-content", "children"),
    Input("url", "pathname"),
    Input("url", "search"),
    Input("auth-store", "data"),
)
def route(pathname, search, auth_data):
    authenticated = bool((auth_data or {}).get("authenticated", False))
    pathname = pathname or "/"

    if not authenticated:
        return login_layout()

    if pathname in ("/", "/inicio"):
        return home_layout()

    if pathname == "/proyectos":
        return projects_layout()

    if pathname == "/detalle":
        params = parse_qs((search or "").lstrip("?"))
        name = unquote(params.get("project", [""])[0])
        return detail_page(name)

    return home_layout()


@app.callback(
    Output("auth-store", "data", allow_duplicate=True),
    Output("url", "pathname", allow_duplicate=True),
    Output("login-error", "children"),
    Input("login-button", "n_clicks"),
    Input("password-input", "n_submit"),
    State("password-input", "value"),
    prevent_initial_call=True,
)
def do_login(n_clicks, n_submit, password):
    if not n_clicks and not n_submit:
        return no_update, no_update, no_update
    if auth_ok(password or ""):
        return {"authenticated": True}, "/inicio", ""
    return no_update, no_update, "Contraseña incorrecta. Intenta nuevamente."


@app.callback(
    Output("auth-store", "data", allow_duplicate=True),
    Output("url", "pathname", allow_duplicate=True),
    Input("logout-button", "n_clicks"),
    prevent_initial_call=True,
)
def do_logout(n_clicks):
    if n_clicks:
        return {"authenticated": False}, "/"
    return no_update, no_update


@app.callback(
    Output("timeline-matrix-container", "children"),
    Output("home-project-detail", "children", allow_duplicate=True),
    Input("vice-radio", "value"),
    prevent_initial_call=True,
)
def update_timeline_matrix(vice):
    proyectos = load_presupuesto()
    if not proyectos:
        return (
            html.Div(
                "No se pudo leer MINEDEC 2.0 desde SharePoint.",
                className="alert",
            ),
            "",
        )

    try:
        df = load_ejecucion(proyectos)
    except Exception:
        df = pd.DataFrame()

    # Al cambiar de viceministerio se limpia cualquier detalle previo.
    return timeline_matrix_component(df, vice), ""


@app.callback(
    Output("home-project-detail", "children"),
    Input({"type":"timeline-project","index":ALL}, "n_clicks"),
    State({"type":"timeline-project","index":ALL}, "value"),
    prevent_initial_call=True,
)
def timeline_project_detail(clicks, values):
    if not clicks or not values:
        return no_update

    triggered = ctx.triggered_id
    if not isinstance(triggered, dict):
        return no_update

    try:
        button_idx = int(triggered.get("index"))
    except Exception:
        return no_update

    if button_idx < 0 or button_idx >= len(values):
        return no_update

    project = values[button_idx]
    if not project:
        return no_update

    return detail_component(str(project), embedded=True)



@app.callback(
    Output("route-macro-area", "children"),
    Output("route-micro-area", "children"),
    Output("route-activity-detail", "children"),
    Input("route-component-select", "value"),
    State("route-roadmap-store", "data"),
    prevent_initial_call=True,
)
def route_select_component(component_name, rm):
    if not component_name:
        return "", "", ""

    comp = _find_component(rm, component_name)

    if not comp:
        return (
            html.Div(
                "No se encontraron registros Macro para el componente seleccionado.",
                className="info-alert",
            ),
            "",
            "",
        )

    macros = comp.get("macros", [])

    macro_options = []
    for macro in macros:
        macro_id = str(macro.get("id_registro", "")).strip()
        nombre = str(macro.get("nombre", "")).strip()
        avance = str(macro.get("avance", "")).strip()

        label = f"{macro_id} · {nombre}"
        if avance:
            label += f"   ·   {avance}"

        macro_options.append({
            "label": label,
            "value": macro_id,
        })

    return (
        html.Div(
            [
                html.Div(
                    [
                        html.Span("2", className="cascade-step-number"),
                        html.Div(
                            [
                                html.Div(
                                    "Etapa Macro",
                                    className="cascade-step-title",
                                ),
                                html.Div(
                                    f"{len(macro_options)} etapas disponibles en {component_name}",
                                    className="cascade-step-help",
                                ),
                            ]
                        ),
                    ],
                    className="cascade-step-heading",
                ),
                dcc.RadioItems(
                    id="route-macro-select",
                    options=macro_options,
                    value=None,
                    className="cascade-radio-list macro-radio-list",
                    inputClassName="cascade-radio-input",
                    labelClassName="cascade-radio-label",
                ),
            ],
            className="cascade-step-card cascade-reveal",
        ),
        "",
        "",
    )


@app.callback(
    Output("route-micro-area", "children", allow_duplicate=True),
    Output("route-activity-detail", "children", allow_duplicate=True),
    Input("route-macro-select", "value"),
    State("route-component-select", "value"),
    State("route-roadmap-store", "data"),
    prevent_initial_call=True,
)
def route_select_macro(macro_id, component_name, rm):
    if not macro_id or not component_name:
        return "", ""

    macro = _find_macro(rm, component_name, macro_id)

    if not macro:
        return (
            html.Div(
                "No se encontró la etapa Macro seleccionada.",
                className="info-alert",
            ),
            "",
        )

    micros = macro.get("micros", [])

    micro_options = []
    for micro in micros:
        micro_id = str(micro.get("id_registro", "")).strip()
        nombre = str(micro.get("nombre", "")).strip()
        label = f"{micro_id} · {nombre}"

        micro_options.append({
            "label": label,
            "value": micro_id,
        })

    # Resumen compacto del Macro.
    # El detalle completo se reserva para la actividad Micro seleccionada,
    # evitando repetir la misma información dos veces.
    macro_card = html.Div(
        [
            _macro_summary_card(macro),
            html.Div(
                [
                    html.Div(
                        [
                            html.Div("Período", className="macro-summary-label"),
                            html.Div(
                                f"{_route_value(macro.get('fecha_inicio'))}  →  "
                                f"{_route_value(macro.get('fecha_fin'))}",
                                className="macro-summary-value",
                            ),
                        ],
                        className="macro-summary-item macro-summary-period",
                    ),
                    html.Div(
                        [
                            html.Div("Avance", className="macro-summary-label"),
                            html.Div(
                                _route_value(macro.get("avance")),
                                className="macro-summary-value",
                            ),
                        ],
                        className="macro-summary-item",
                    ),
                    html.Div(
                        [
                            html.Div("Estado", className="macro-summary-label"),
                            html.Div(
                                _route_value(macro.get("estado")),
                                className="macro-summary-value",
                            ),
                        ],
                        className="macro-summary-item",
                    ),
                    html.Div(
                        [
                            html.Div("Actividades Micro", className="macro-summary-label"),
                            html.Div(
                                str(len(micros)),
                                className="macro-summary-value",
                            ),
                        ],
                        className="macro-summary-item",
                    ),
                ],
                className="macro-executive-summary",
            ),
        ],
        className="cascade-selection-card",
    )

    micro_selector = html.Div(
        [
            html.Div(
                [
                    html.Span("3", className="cascade-step-number"),
                    html.Div(
                        [
                            html.Div(
                                "Actividad Micro",
                                className="cascade-step-title",
                            ),
                            html.Div(
                                f"{len(micro_options)} actividades asociadas a esta etapa",
                                className="cascade-step-help",
                            ),
                        ]
                    ),
                ],
                className="cascade-step-heading",
            ),

            macro_card,

            html.Div(
                [
                    html.Div(
                        "Seleccione una actividad Micro para consultar su información",
                        className="cascade-list-caption",
                    ),
                    dcc.RadioItems(
                        id="route-micro-select",
                        options=micro_options,
                        value=None,
                        className="cascade-radio-list micro-radio-list",
                        inputClassName="cascade-radio-input",
                        labelClassName="cascade-radio-label",
                    )
                    if micro_options
                    else html.Div(
                        "Esta etapa no tiene registros Micro asociados en HOJA_RUTA.",
                        className="cascade-empty-state",
                    ),
                ],
                className="cascade-micro-selector",
            ),
        ],
        className="cascade-step-card cascade-reveal",
    )

    return micro_selector, ""


@app.callback(
    Output("route-activity-detail", "children", allow_duplicate=True),
    Input("route-micro-select", "value"),
    State("route-macro-select", "value"),
    State("route-component-select", "value"),
    State("route-roadmap-store", "data"),
    prevent_initial_call=True,
)
def route_select_micro(micro_id, macro_id, component_name, rm):
    if not micro_id or not macro_id or not component_name:
        return ""

    micro = _find_micro(
        rm,
        component_name,
        macro_id,
        micro_id,
    )

    if not micro:
        return html.Div(
            "No se encontró la actividad Micro seleccionada.",
            className="info-alert",
        )

    alerta = _route_value(micro.get("alerta_vencimiento"))

    return html.Div(
        [
            html.Div(
                [
                    html.Div(
                        [
                            html.Div(
                                "Detalle de la actividad",
                                className="cascade-kicker",
                            ),
                            html.H4(
                                _route_value(micro.get("nombre")),
                                className="cascade-activity-title",
                            ),
                            html.Div(
                                _route_value(micro.get("id_registro")),
                                className="cascade-activity-code",
                            ),
                        ]
                    ),
                    html.Span(
                        alerta,
                        className=_status_class(alerta),
                    ),
                ],
                className="cascade-activity-head",
            ),
            _route_detail_grid(micro),
        ],
        className="cascade-activity-detail cascade-reveal",
    )


@app.callback(
    Output("project-list", "children"),
    Input("project-search", "value"),
    Input("project-vice-filter", "value"),
)
def filter_projects(query, vice):
    proyectos = load_presupuesto()
    names = list(proyectos)
    if query:
        q = normalizar(query)
        names = [n for n in names if q in normalizar(n)]
    if vice and vice != "Todos":
        names = [n for n in names if vice_short(proyectos[n].get("viceministerio")) == vice]
    names.sort(key=lambda n: proyectos[n]["pct_ejecucion"])
    if not names:
        return card(html.Div("No se encontraron proyectos con los filtros seleccionados.", className="hint"))
    return card([project_row(n, proyectos[n]) for n in names])


# =============================================================================
# 15. EJECUCIÓN
# =============================================================================

def _precalentar_cache_en_segundo_plano():
    """
    Al iniciar:
    - carga presupuesto;
    - pone inmediatamente en RAM la última ejecución procesada disponible;
    - refresca SharePoint después, sin bloquear la interfaz.
    """
    try:
        proyectos = load_presupuesto()

        local = _leer_cache_ejecucion()
        if not local.empty:
            with _CACHE_LOCK:
                _CACHE["ejecucion_mensual"] = (time.time(), local)

        refrescar_ejecucion_en_segundo_plano()

    except Exception:
        pass


if __name__ == "__main__":
    host = os.getenv("DASH_HOST", "127.0.0.1")
    port = int(os.getenv("PORT", "8050"))

    threading.Thread(target=_precalentar_cache_en_segundo_plano, daemon=True).start()

    run_kwargs = {
        "debug": False,
        "host": host,
        "port": port,
    }

    try:
        get_ipython  # noqa: F821
        run_kwargs["jupyter_mode"] = "external"
    except NameError:
        pass

    app.run(**run_kwargs)