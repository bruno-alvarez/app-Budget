import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import re
import json
import csv
import hashlib
import urllib.request
from datetime import datetime

# =============================================================================
# 0. CONFIGURACIÓN
# =============================================================================
st.set_page_config(page_title="Mining Control - Dashboard Oficial", layout="wide",
                   initial_sidebar_state="expanded")

HOJA_FORECAST = "Forecast 5+7"
HOJA_BUDGET = "BUDGET 2027-2031"

# Nombres candidatos del libro maestro (contiene ambas hojas)
CANDIDATOS = ["Budget_2027-2031.xlsx", "Budget 2027-2031.xlsx",
              "Datos_Proyecto_Mejora__2026__3_.xlsx", "Datos Proyecto Mejora  2026 (3).xlsx"]

MESES_REALES = ['Jan-26', 'Feb-26', 'Mar-26', 'Apr-26', 'May-26']
MESES_PROY = ['Jun-26', 'Jul-26', 'Aug-26', 'Sep-26', 'Oct-26', 'Nov-26', 'Dec-26']
MESES_27 = ['Jan-27', 'Feb-27', 'Mar-27', 'Apr-27', 'May-27', 'Jun-27',
            'Jul-27', 'Aug-27', 'Sep-27', 'Oct-27', 'Nov-27', 'Dec-27']
FY_COLS = ['FY27', 'FY28', 'FY29', 'FY30', 'FY31']
N_REAL = 5

# Precios de referencia (ancla para traducir precio en vivo -> % de variación)
PRECIO_REF = {"dolar": 950.0, "uf": 38000.0, "cobre": 4.30, "wti": 70.0}


# =============================================================================
# 1. RESOLUCIÓN ROBUSTA DEL ARCHIVO (arregla el FileNotFoundError)
# =============================================================================
def _buscar_en_disco():
    # (a) nombres conocidos
    for n in CANDIDATOS:
        if os.path.exists(n):
            return n
    # (b) cualquier .xlsx en el directorio que tenga las dos hojas necesarias
    for f in os.listdir('.'):
        if f.lower().endswith('.xlsx') and not f.startswith('~$'):
            try:
                hojas = pd.ExcelFile(f).sheet_names
                if HOJA_FORECAST in hojas and HOJA_BUDGET in hojas:
                    return f
            except Exception:
                continue
    return None


@st.cache_data(show_spinner="Procesando planilla maestra...")
def parsear(_data: bytes, _key: str):
    """Parsea el libro (Forecast 5+7 + BUDGET 2027-2031). _key evita re-hashear los bytes."""
    bio = io.BytesIO(_data)

    # ---- Forecast 5+7 (header en la fila 2) ----
    df = pd.read_excel(bio, sheet_name=HOJA_FORECAST, header=1)
    for col in ['Resp', 'Proc', 'Item', 'CC']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    for col in ['VP', 'Gerencia', 'Classif']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
    for col in MESES_REALES + MESES_PROY + ['YTD', 'Forecast FY', 'Budget FY']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    df['YTD'] = df['YTD'].where(df['YTD'] != 0, df[MESES_REALES].sum(axis=1))
    df['Proyeccion_base'] = df['Forecast FY'] - df['YTD']
    df['Driver'] = _asignar_driver(df)

    # ---- BUDGET 2027-2031 (header en la fila 1) — SE RESPETA TAL CUAL ----
    bio.seek(0)
    bud = pd.read_excel(bio, sheet_name=HOJA_BUDGET, header=0)
    for col in ['Resp', 'Proc', 'Item', 'CC']:
        if col in bud.columns:
            bud[col] = bud[col].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    for col in ['VP', 'Gerencia', 'Classif']:
        if col in bud.columns:
            bud[col] = bud[col].astype(str).str.strip()
    for col in MESES_27 + FY_COLS:
        if col in bud.columns:
            bud[col] = pd.to_numeric(bud[col], errors='coerce').fillna(0)
    bud = bud.dropna(subset=['Item']).reset_index(drop=True)
    bud['Driver'] = _asignar_driver(bud)
    return df, bud


def _asignar_driver(d):
    di = d['Desc Item'].astype(str)
    cl = d['Classif'].astype(str).str.strip()
    return np.select(
        [(cl == 'Fuel') | di.str.contains('Combustible|Diesel|Diésel|Petr', case=False, regex=True),
         cl == 'Labor',
         di.str.contains('Internacional|Extranjer|Importad', case=False, regex=True)
         | cl.isin(['Spare Parts', 'S&C'])],
        ['Combustible', 'Mano de Obra', 'Divisas'], default='Local / IPC')


# =============================================================================
# 2. PRECIOS DE MERCADO EN VIVO (mindicador.cl + Stooq) — con respaldo
# =============================================================================
_UA = ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
       '(KHTML, like Gecko) Chrome/124.0 Safari/537.36')


def _http(url, t=8, intentos=3):
    """GET con reintentos. Devuelve el texto crudo de la respuesta."""
    ult = None
    for _ in range(intentos):
        try:
            req = urllib.request.Request(url, headers={'User-Agent': _UA, 'Accept': '*/*'})
            with urllib.request.urlopen(req, timeout=t) as r:
                return r.read().decode('utf-8', 'replace')
        except Exception as e:
            ult = e
    raise ult


def _mindicador(codigo):
    """Lee un indicador de mindicador.cl tolerando estructuras distintas."""
    d = json.loads(_http(f"https://mindicador.cl/api/{codigo}"))
    serie = d.get("serie")
    if serie and isinstance(serie, list) and serie and "valor" in serie[0]:
        return float(serie[0]["valor"])
    if "valor" in d:                       # algunos endpoints devuelven el valor directo
        return float(d["valor"])
    raise ValueError(f"sin 'serie' para {codigo}")


def _stooq(sym):
    """Último cierre desde Stooq. Prueba varios formatos de respuesta."""
    txt = _http(f"https://stooq.com/q/l/?s={sym}&f=sd2t2ohlcv&h&e=csv")
    rows = list(csv.DictReader(io.StringIO(txt)))
    if rows and rows[0].get('Close') not in (None, '', 'N/D'):
        return float(rows[0]['Close'])
    raise ValueError(f"Stooq sin Close para {sym}")


def _primero(*fuentes):
    """Devuelve el primer fetcher que tenga éxito; si todos fallan, relanza el último error."""
    ult = None
    for fn in fuentes:
        try:
            v = fn()
            if v is not None:
                return float(v)
        except Exception as e:
            ult = e
    raise ult if ult else ValueError("sin fuentes")


@st.cache_data(ttl=1800, show_spinner=False)
def obtener_precios():
    res = {}

    def tryf(k, fn, fb, u):
        try:
            res[k] = {"valor": float(fn()), "estado": "vivo", "unidad": u}
        except Exception:
            res[k] = {"valor": float(fb), "estado": "referencia", "unidad": u}

    # USD/CLP: mindicador -> Stooq (usdclp)
    tryf("dolar", lambda: _primero(lambda: _mindicador("dolar"), lambda: _stooq("usdclp")),
         PRECIO_REF["dolar"], "CLP/USD")
    # UF: mindicador
    tryf("uf", lambda: _mindicador("uf"), PRECIO_REF["uf"], "CLP")
    # Cobre: mindicador -> Stooq cobre futuros (hg.f, USD/lb)
    tryf("cobre", lambda: _primero(lambda: _mindicador("cobre"), lambda: _stooq("hg.f")),
         PRECIO_REF["cobre"], "USD/lb")
    # Crudo: Stooq WTI (wti.us) -> WTI futuros (cl.f) -> Brent (cb.f) como proxy
    tryf("wti", lambda: _primero(lambda: _stooq("wti.us"), lambda: _stooq("cl.f"), lambda: _stooq("cb.f")),
         PRECIO_REF["wti"], "USD/bbl")

    # IPC anualizado (suma de los últimos 12 valores mensuales) para indexar mano de obra
    try:
        y = datetime.now().year
        serie = json.loads(_http(f"https://mindicador.cl/api/ipc/{y}")).get("serie", [])
        try:
            serie += json.loads(_http(f"https://mindicador.cl/api/ipc/{y-1}")).get("serie", [])
        except Exception:
            pass
        ult12 = [s["valor"] for s in serie[:12]]
        if not ult12:
            raise ValueError("ipc vacío")
        res["ipc_anual"] = {"valor": float(sum(ult12)), "estado": "vivo", "unidad": "% 12m"}
    except Exception:
        res["ipc_anual"] = {"valor": 3.5, "estado": "referencia", "unidad": "% 12m"}

    res["timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    return res


def var_pct(actual, ref):
    return round((actual / ref - 1) * 100, 1) if ref else 0.0


# =============================================================================
# 3. MOTOR DE SIMULACIÓN  (Forecast 5+7  ->  efectos en Budget 2027-2031)
# =============================================================================
def simular(df, bud, shf, shx, shl):
    mm = {'Combustible': 1 + shf / 100.0, 'Divisas': 1 + shx / 100.0,
          'Mano de Obra': 1 + shl / 100.0, 'Local / IPC': 1.0}

    # -------- FORECAST 5+7 2026 (solo se ajusta la proyección "+7") --------
    fc = df.copy()
    mult_f = fc['Driver'].map(mm).values
    fc['Forecast FY'] = fc['YTD'] + fc['Proyeccion_base'] * mult_f
    fc['Var'] = fc['Forecast FY'] - fc['Budget FY']
    for mes in MESES_PROY:                       # meses proyectados se mueven con el shock
        fc[mes] = fc[mes] * fc['Driver'].map(mm)
    # (Jan-May reales se respetan tal cual)

    # -------- BUDGET 2027-2031 (se RESPETA y se escala por driver) --------
    bd = bud.copy()
    mult_b = bd['Driver'].map(mm).values
    for col in MESES_27 + FY_COLS:               # mensual 2027 + FY27..FY31
        if col in bd.columns:
            bd[col] = bd[col].values * mult_b
    # FY27 sigue siendo la suma de los 12 meses (se preserva la identidad)
    return fc, bd


# =============================================================================
# 4. INFORME FINANCIERO (Gemini API + respaldo local detallado)
# =============================================================================
def _api_key():
    """Busca la API key de Gemini de forma robusta:
    1) .streamlit/secrets.toml  (varias convenciones de nombre, incl. sección [gemini])
    2) campo manual en el panel   3) variable de entorno.
    Formato recomendado en secrets.toml:   GEMINI_API_KEY = "tu_clave"
    (o como sección)                         [gemini]\n    api_key = "tu_clave"
    """
    try:
        s = st.secrets
        for k in ("GEMINI_API_KEY", "gemini_api_key", "GOOGLE_API_KEY", "google_api_key", "api_key"):
            if k in s and s[k]:
                return s[k]
        if "gemini" in s:                      # sección [gemini]
            sec = s["gemini"]
            for k in ("api_key", "API_KEY", "key", "GEMINI_API_KEY"):
                if k in sec and sec[k]:
                    return sec[k]
    except Exception:
        pass
    return os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")


def generar_informe(datos, modelo="gemini-2.5-flash"):
    """Devuelve (texto, fuente). Intenta Gemini; si no hay clave o falla, informe local."""
    prompt = _prompt_informe(datos)
    key = _api_key()
    if key:
        try:
            url = (f"https://generativelanguage.googleapis.com/v1beta/models/"
                   f"{modelo}:generateContent?key={key}")
            body = json.dumps({"contents": [{"parts": [{"text": prompt}]}],
                               "generationConfig": {"temperature": 0.4, "maxOutputTokens": 2048}}).encode()
            req = urllib.request.Request(url, data=body, headers={'Content-Type': 'application/json'})
            with urllib.request.urlopen(req, timeout=40) as r:
                data = json.loads(r.read().decode('utf-8'))
            txt = data["candidates"][0]["content"]["parts"][0]["text"]
            if not txt or len(txt.strip()) < 80:        # respuesta vacía/cortada -> respaldo
                return _informe_local(datos), "local (respaldo)"
            return txt.strip(), "Gemini (" + modelo + ")"
        except Exception as e:
            return _informe_local(datos) + f"\n\n> Nota: no se pudo usar Gemini ({e}); resumen generado localmente.", "local"
    return _informe_local(datos), "local"


def _prompt_informe(d):
    return f"""Eres un analista financiero senior de una minera chilena. Redacta un RESUMEN EJECUTIVO BREVE
en español (Markdown), de 200 a 320 palabras, como feedback para la gerencia. NO pongas un título principal
(ya va dentro de una sección). NO repitas tablas de cifras (ya están en el informe). Usa exactamente estas 3
subsecciones con encabezados '### ':
### Lectura de la ejecución 2026
(1 párrafo: interpreta la desviación Forecast vs Budget y qué la explica)
### Sensibilidades
(1 párrafo: qué palanca pesa más y por qué, con lectura económica)
### Recomendaciones
(3 o 4 viñetas accionables y concretas)
Sé incisivo y orientado a decisión; menciona 2 o 3 cifras clave, sin saturar.

DATOS (USD): Budget 2026 {d['budget_2026']:,.0f}; Forecast 2026 {d['forecast_adj']:,.0f};
Desviación {d['desv']:,.0f} ({d['desv_pct']:.2f}%); Shocks Combustible {d['shf']:+.1f}%, Divisas {d['shx']:+.1f}%,
Mano de obra {d['shl']:+.1f}%; Impacto en forecast: Combustible {d['imp_fuel']:,.0f}, Divisas {d['imp_fx']:,.0f},
Mano de obra {d['imp_labor']:,.0f}; Pesos del gasto: Combustible {d['w_fuel']:.1f}%, Divisas {d['w_fx']:.1f}%,
Mano de obra {d['w_labor']:.1f}%, Local {d['w_local']:.1f}%; Budget FY27 {d['fy_adj'][0]} -> FY31 {d['fy_adj'][4]},
CAGR {d['cagr']:.2f}%; Mercado: USD/CLP {d['px']['dolar']['valor']:.1f}, WTI {d['px']['wti']['valor']:.1f},
Cobre {d['px']['cobre']['valor']:.2f}, IPC 12m {d['px']['ipc_anual']['valor']:.1f}%."""


def _informe_local(d):
    signo = "sobre-ejecución (presión de costos)" if d['desv'] > 0 else "ahorro frente al presupuesto"
    palanca = max([("combustible", d['imp_fuel']), ("divisas", d['imp_fx']),
                   ("mano de obra", d['imp_labor'])], key=lambda t: abs(t[1]))[0]
    return f"""### Lectura de la ejecución 2026
El Forecast 5+7 proyecta un cierre de **USD {d['forecast_adj']:,.0f}** frente a un budget de
**USD {d['budget_2026']:,.0f}**, una desviación de **USD {d['desv']:,.0f} ({d['desv_pct']:+.2f}%)** — situación de
**{signo}**. La proyección combina 5 meses reales con 7 estimados, por lo que solo la porción proyectada
reacciona a las palancas de mercado bajo los supuestos vigentes (combustible {d['shf']:+.1f}%, divisas
{d['shx']:+.1f}%, mano de obra {d['shl']:+.1f}%).

### Sensibilidades
La palanca de mayor impacto en este escenario es **{palanca}**. Estructuralmente, **divisas** concentra el mayor
apalancamiento ({d['w_fx']:.1f}% del gasto proyectado): una depreciación del peso encarece de inmediato insumos
y repuestos importados. Combustible ({d['w_fuel']:.1f}%) y mano de obra ({d['w_labor']:.1f}%) agregan volatilidad
adicional. El plan quinquenal crece a un CAGR de **{d['cagr']:.2f}%** hacia FY31.

### Recomendaciones
- Coberturas de tipo de cambio para la fracción importada del gasto ({d['w_fx']:.1f}%).
- Cláusulas de indexación y contratos de combustible para acotar su volatilidad.
- Monitoreo mensual del run-rate 5+7 frente al budget para anticipar desviaciones.
- Gestionar las líneas de mayor impacto identificadas en "Líneas más impactadas".
"""


_AZUL = (31, 78, 120)


def _sanit(s):
    rep = {'—': '-', '–': '-', '•': '-', '“': '"', '”': '"', '‘': "'", '’': "'",
           '…': '...', '→': '->', '↔': '<->', '✓': 'OK', '⚠': '!', '🔝': '', '🧠': '', '°': 'o',
           '⛽': '', '💱': '', '👷': '', '🌐': '', '📊': '', '📈': '', '💰': '', '🗓': '', 'Δ': 'Delta '}
    for k, v in rep.items():
        s = s.replace(k, v)
    s = re.sub(r'(?<!\*)\*(?!\*)', '', s)            # quita itálicas de un solo asterisco
    return s.encode('latin-1', 'replace').decode('latin-1')


def _png(fig):
    import io as _io
    buf = _io.BytesIO()
    fig.savefig(buf, format='png', bbox_inches='tight')
    buf.seek(0)
    return buf


def _charts(rep):
    """Genera los gráficos del informe como PNG en memoria. Si no hay matplotlib, devuelve []."""
    try:
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as plt
    except Exception:
        return []
    out = []
    # 1. Budget mensual 2027
    fig, ax = plt.subplots(figsize=(7, 3), dpi=130)
    x = range(12); w = 0.4
    ax.bar([i - w/2 for i in x], [v/1e6 for v in rep['bud_mensual_base']], w, label='Base', color='#9DB4C8')
    ax.bar([i + w/2 for i in x], [v/1e6 for v in rep['bud_mensual_adj']], w, label='Ajustado', color='#1F4E78')
    ax.set_xticks(list(x)); ax.set_xticklabels(rep['meses27_labels'], fontsize=7)
    ax.set_title('Budget mensual 2027 (M USD): base vs ajustado', fontsize=9, weight='bold')
    ax.legend(fontsize=7); ax.spines[['top', 'right']].set_visible(False); ax.tick_params(labelsize=7)
    out.append(_png(fig)); plt.close(fig)
    # 2. Trayectoria FY27-31
    fig, ax = plt.subplots(figsize=(7, 2.8), dpi=130)
    ax.plot(rep['fy_labels'], [v/1e6 for v in rep['fy_b']], 'o-', label='Base', color='#9DB4C8')
    ax.plot(rep['fy_labels'], [v/1e6 for v in rep['fy_adj']], 'o-', label='Ajustado', color='#1F4E78')
    ax.set_title('Trayectoria Budget 2027-2031 (M USD)', fontsize=9, weight='bold')
    ax.legend(fontsize=7); ax.spines[['top', 'right']].set_visible(False); ax.tick_params(labelsize=7)
    out.append(_png(fig)); plt.close(fig)
    # 3. Aporte por palanca
    fig, ax = plt.subplots(figsize=(7, 2.4), dpi=130)
    labs = list(rep['impactos'].keys()); vals = [v/1e6 for v in rep['impactos'].values()]
    ax.barh(labs, vals, color=['#C0392B' if v < 0 else '#2E7D50' for v in vals])
    ax.set_title('Aporte de cada palanca al forecast 2026 (M USD)', fontsize=9, weight='bold')
    ax.axvline(0, color='#888', lw=.8); ax.spines[['top', 'right']].set_visible(False); ax.tick_params(labelsize=7)
    out.append(_png(fig)); plt.close(fig)
    return out


def _tabla(pdf, W, headers, rows, cw, align=None):
    h = 7
    pdf.set_font('Helvetica', 'B', 9); pdf.set_fill_color(*_AZUL); pdf.set_text_color(255)
    for c, t in zip(cw, headers):
        pdf.cell(c, h, _sanit(t), border=1, align='C', fill=True)
    pdf.ln(); pdf.set_text_color(0)
    for i, row in enumerate(rows):
        pdf.set_fill_color(*((242, 246, 250) if i % 2 else (255, 255, 255)))
        for j, val in enumerate(row):
            bold = (j == 0)
            pdf.set_font('Helvetica', 'B' if bold else '', 9)
            a = (align[j] if align else ('L' if j == 0 else 'R'))
            pdf.cell(cw[j], h, _sanit(str(val)), border=1, align=a, fill=True)
        pdf.ln()
    pdf.ln(3)


def _render_md(pdf, W, md):
    LM = dict(new_x="LMARGIN", new_y="NEXT")     # vuelve siempre al margen izquierdo
    lines = md.split('\n'); i = 0
    while i < len(lines):
        ln = lines[i].rstrip()
        if ln.strip().startswith('|') and i + 1 < len(lines) and \
                set(lines[i + 1].replace('|', '').strip()) <= set('-: '):
            i += 1
            while i < len(lines) and lines[i].strip().startswith('|'):
                i += 1
            continue   # las tablas ya se imprimieron arriba como KPIs
        pdf.set_x(pdf.l_margin)
        if ln.startswith('# '):
            pdf.set_font('Helvetica', 'B', 13); pdf.set_text_color(*_AZUL)
            pdf.multi_cell(W, 7, _sanit(ln[2:].replace('**', '')), **LM); pdf.set_text_color(0)
        elif ln.startswith('## '):
            pdf.ln(1); pdf.set_font('Helvetica', 'B', 11.5); pdf.set_text_color(*_AZUL)
            pdf.multi_cell(W, 6.5, _sanit(ln[3:].replace('**', '')), **LM); pdf.set_text_color(0)
        elif ln.startswith('### '):
            pdf.ln(1); pdf.set_font('Helvetica', 'B', 10.5)
            pdf.multi_cell(W, 6, _sanit(ln[4:].replace('**', '')), **LM)
        elif ln.strip().startswith(('- ', '* ')):
            pdf.set_font('Helvetica', '', 9.5)
            pdf.multi_cell(W, 5.5, _sanit('  -  ' + ln.strip()[2:]), markdown=True, **LM)
        elif re.match(r'^\d+\.\s', ln.strip()):
            pdf.set_font('Helvetica', '', 9.5)
            pdf.multi_cell(W, 5.5, _sanit('  ' + ln.strip()), markdown=True, **LM)
        elif ln.startswith('> '):
            pdf.set_font('Helvetica', 'I', 9); pdf.multi_cell(W, 5.5, _sanit(ln[2:].replace('**', '')), **LM)
        elif ln.strip() == '':
            pdf.ln(2)
        elif set(ln.strip()) <= set('-*=') and ln.strip():
            pass
        else:
            pdf.set_font('Helvetica', '', 9.5); pdf.multi_cell(W, 5.5, _sanit(ln), markdown=True, **LM)
        i += 1


def construir_pdf(rep, narrativa, fuente="local"):
    """Informe ejecutivo completo: portada + KPIs + gráficos + sensibilidades + top movers + análisis."""
    from fpdf import FPDF
    from fpdf.enums import XPos, YPos
    pdf = FPDF(format='A4'); pdf.set_auto_page_break(True, margin=15); pdf.add_page(); pdf.set_margins(15, 15, 15)
    W = pdf.w - 30

    def titulo(n, t):
        if pdf.get_y() > 250:
            pdf.add_page()
        pdf.set_font('Helvetica', 'B', 12); pdf.set_text_color(*_AZUL)
        pdf.cell(0, 8, _sanit(f"{n}. {t}"), new_x=XPos.LMARGIN, new_y=YPos.NEXT); pdf.set_text_color(0)

    # ---- Encabezado ----
    pdf.set_fill_color(*_AZUL); pdf.rect(0, 0, pdf.w, 26, 'F')
    pdf.set_xy(15, 7); pdf.set_text_color(255); pdf.set_font('Helvetica', 'B', 17)
    pdf.cell(0, 8, _sanit("Informe Financiero Ejecutivo"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_x(15); pdf.set_font('Helvetica', '', 9)
    pdf.cell(0, 5, _sanit(f"Proyecto Minero  |  {rep['fecha']}  |  Analisis: {fuente}"))
    pdf.set_text_color(0); pdf.set_y(30)
    pdf.set_font('Helvetica', 'I', 8.5); pdf.multi_cell(W, 5, _sanit(rep['scope']), new_x="LMARGIN", new_y="NEXT"); pdf.ln(2)

    # ---- 1. KPIs ----
    titulo(1, "Resumen de KPIs")
    f = lambda v: f"{v:,.0f}"
    _tabla(pdf, W, ["Indicador (USD)", "Base", "Con shocks"],
           [["Budget FY 2026", f(rep['budget_2026']), "-"],
            ["Forecast FY 2026", f(rep['forecast_base']), f(rep['forecast_adj'])],
            ["Desviacion (Var)", f(rep['forecast_base'] - rep['budget_2026']),
             f"{f(rep['desv'])} ({rep['desv_pct']:+.2f}%)"],
            ["Budget FY27", f(rep['fy_b'][0]), f(rep['fy_adj'][0])],
            ["Budget FY31", f(rep['fy_b'][4]), f(rep['fy_adj'][4])],
            ["Acumulado 2027-2031", f(sum(rep['fy_b'])), f(sum(rep['fy_adj']))],
            ["CAGR 27-31", f"{rep['cagr_b']:.2f}%", f"{rep['cagr_adj']:.2f}%"]],
           [W*0.4, W*0.3, W*0.3])

    # ---- 2. Gráficos ----
    titulo(2, "Visualizaciones")
    imgs = _charts(rep)
    if imgs:
        for buf in imgs:
            if pdf.get_y() > 225:
                pdf.add_page()
            pdf.image(buf, x=15, w=W); pdf.ln(2)
    else:
        pdf.set_font('Helvetica', 'I', 9)
        pdf.multi_cell(W, 6, _sanit("(Instala matplotlib para incrustar los graficos: pip install matplotlib)"),
                       new_x="LMARGIN", new_y="NEXT")
        pdf.ln(2)

    # ---- 3. Sensibilidades ----
    titulo(3, "Analisis de Sensibilidades")
    _tabla(pdf, W, ["Palanca", "Shock", "Impacto fcst 2026 (USD)", "Peso gasto"],
           [["Combustible / Diesel", f"{rep['shf']:+.1f}%", f(rep['imp_fuel']), f"{rep['w_fuel']:.1f}%"],
            ["Tipo de cambio USD/CLP", f"{rep['shx']:+.1f}%", f(rep['imp_fx']), f"{rep['w_fx']:.1f}%"],
            ["Mano de obra", f"{rep['shl']:+.1f}%", f(rep['imp_labor']), f"{rep['w_labor']:.1f}%"]],
           [W*0.34, W*0.16, W*0.30, W*0.20],
           align=['L', 'C', 'R', 'C'])

    # ---- 4. Top movers ----
    if rep.get('movers'):
        titulo(4, "Lineas mas impactadas (Delta FY27)")
        _tabla(pdf, W, ["Item", "VP", "Driver", "Delta FY27 (USD)"],
               [[str(d)[:40], str(v)[:12], str(dr)[:14], f"{x:,.0f}"] for d, v, dr, x in rep['movers'][:8]],
               [W*0.40, W*0.16, W*0.18, W*0.26], align=['L', 'L', 'L', 'R'])

    # ---- 5. Análisis narrativo ----
    titulo(5, "Analisis y Recomendaciones")
    _render_md(pdf, W, narrativa)
    return bytes(pdf.output())


# =============================================================================
# 5. INTERFAZ
# =============================================================================
st.title("📊 Dashboard Corporativo Minero — Forecast 5+7  ↔  Budget 2027-2031")
st.markdown("##### Volatilidades de mercado en vivo · El Forecast 5+7 impacta el Budget 2027-2031")

# ---- obtención robusta del archivo ----
archivo = _buscar_en_disco()
data_bytes, fuente_archivo = None, None
if archivo:
    with open(archivo, 'rb') as f:
        data_bytes = f.read()
    fuente_archivo = archivo
else:
    st.warning("No encontré el libro maestro en la carpeta. Súbelo aquí 👇 (debe contener las hojas "
               f"'{HOJA_FORECAST}' y '{HOJA_BUDGET}').")
    up = st.file_uploader("Cargar libro maestro (.xlsx)", type=['xlsx'])
    if up is not None:
        data_bytes = up.read()
        fuente_archivo = up.name

if data_bytes is None:
    st.stop()

try:
    key = hashlib.md5(data_bytes).hexdigest()
    df_base, bud_base = parsear(data_bytes, key)
    st.caption(f"📁 Fuente: **{fuente_archivo}** · {len(df_base):,} líneas forecast · {len(bud_base):,} líneas budget")

    # ---------------- SIDEBAR: precios ----------------
    st.sidebar.header("🌐 Precios de Mercado (en vivo)")
    if st.sidebar.button("🔄 Actualizar precios a la fecha", use_container_width=True):
        obtener_precios.clear()
    px = obtener_precios()
    b = lambda e: "🟢 en vivo" if e == "vivo" else "🟡 referencia"
    c1, c2 = st.sidebar.columns(2)
    c1.metric("USD/CLP", f"${px['dolar']['valor']:,.1f}", b(px['dolar']['estado']))
    c2.metric("WTI", f"${px['wti']['valor']:,.1f}", b(px['wti']['estado']))
    c3, c4 = st.sidebar.columns(2)
    c3.metric("Cobre USD/lb", f"{px['cobre']['valor']:,.2f}", b(px['cobre']['estado']))
    c4.metric("IPC 12m", f"{px['ipc_anual']['valor']:,.1f}%", b(px['ipc_anual']['estado']))
    st.sidebar.caption(f"Actualizado: {px['timestamp']} · mindicador.cl / Stooq")

    var_fuel = var_pct(px['wti']['valor'], PRECIO_REF['wti'])
    var_fx = var_pct(px['dolar']['valor'], PRECIO_REF['dolar'])

    # ---------------- SIDEBAR: sensibilidades ----------------
    st.sidebar.write("---")
    st.sidebar.header("⚙️ Sensibilidades")
    if 'shf' not in st.session_state:
        st.session_state.update(shf=0.0, shx=0.0, shl=0.0)

    bc1, bc2 = st.sidebar.columns(2)
    if bc1.button("📈 Aplicar mercado", use_container_width=True,
                  help="Fija combustible y divisas según el precio en vivo vs. referencia."):
        st.session_state.shf = float(np.clip(var_fuel, -30, 30))
        st.session_state.shx = float(np.clip(var_fx, -30, 30))
    if bc2.button("↺ Reset", use_container_width=True):
        st.session_state.update(shf=0.0, shx=0.0, shl=0.0)

    idx_labor = st.sidebar.checkbox("👷 Indexar mano de obra al IPC 12m",
                                    help="Fija el reajuste de mano de obra al IPC anualizado en vivo.")
    if idx_labor:
        st.session_state.shl = float(np.clip(px['ipc_anual']['valor'], -30, 30))

    shf = st.sidebar.slider("⛽ Combustible / Diésel (%)", -30.0, 30.0, st.session_state.shf, 0.5,
                            key='shf', help=f"Mercado sugiere {var_fuel:+.1f}% (WTI vs ref).")
    shx = st.sidebar.slider("💱 Tipo de cambio USD/CLP (%)", -30.0, 30.0, st.session_state.shx, 0.5,
                            key='shx', help=f"Mercado sugiere {var_fx:+.1f}%. + = peso más débil.")
    shl = st.sidebar.slider("👷 Mano de obra (%)", -30.0, 30.0, st.session_state.shl, 0.5,
                            key='shl', disabled=idx_labor)

    # ---------------- SIDEBAR: filtros ----------------
    st.sidebar.write("---")
    st.sidebar.header("🔍 Filtros")
    vps = sorted([v for v in df_base['VP'].unique() if v.lower() != 'nan' and v])
    vp_sel = st.sidebar.multiselect("Vicepresidencia(s)", vps, default=vps)
    gers = sorted([g for g in df_base[df_base['VP'].isin(vp_sel)]['Gerencia'].unique()
                   if g.lower() != 'nan' and g])
    ger_sel = st.sidebar.multiselect("Gerencia(s)", gers, default=gers)

    fmask = df_base['VP'].isin(vp_sel) & df_base['Gerencia'].isin(ger_sel)
    bmask = bud_base['VP'].isin(vp_sel) & bud_base['Gerencia'].isin(ger_sel)
    df_f, bud_f = df_base[fmask].copy(), bud_base[bmask].copy()

    if df_f.empty:
        st.warning("⚠️ No hay datos para los filtros seleccionados.")
        st.stop()

    # ---------------- SIMULACIÓN ----------------
    fcst, bud = simular(df_f, bud_f, shf, shx, shl)
    fcst0, bud0 = simular(df_f, bud_f, 0, 0, 0)

    budget_2026 = fcst['Budget FY'].sum()
    fc_total = fcst['Forecast FY'].sum()
    fc_base = fcst0['Forecast FY'].sum()
    desv = fc_total - budget_2026
    desv_pct = desv / budget_2026 * 100 if budget_2026 else 0

    # ---------------- KPIs 2026 ----------------
    st.subheader("Ejecución 2026 · Forecast 5+7")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("💰 Budget FY 2026", f"USD {budget_2026:,.0f}")
    k2.metric("📈 Forecast FY 2026", f"USD {fc_total:,.0f}",
              f"{(fc_total/fc_base-1)*100:+.2f}% vs base" if fc_base else None)
    k3.metric("⚠️ Desviación (Var)", f"USD {desv:,.0f}", f"{desv_pct:.2f}%", delta_color="inverse")
    k4.metric("🌐 Impacto volatilidad", f"USD {fc_total-fc_base:,.0f}", "sobre forecast", delta_color="off")

    # ---------------- KPIs Budget 2027-2031 ----------------
    st.subheader("Proyección Quinquenal · Budget 2027-2031  (respeta la hoja base)")
    fy_adj = [bud[y].sum() for y in FY_COLS]
    fy_b = [bud0[y].sum() for y in FY_COLS]
    q1, q2, q3, q4 = st.columns(4)
    q1.metric("FY27", f"USD {fy_adj[0]:,.0f}", f"{(fy_adj[0]/fy_b[0]-1)*100:+.2f}%" if fy_b[0] else None)
    q2.metric("FY31", f"USD {fy_adj[4]:,.0f}")
    cagr = (fy_adj[4]/fy_adj[0])**(1/4)-1 if fy_adj[0] > 0 else 0
    q3.metric("CAGR 27→31", f"{cagr*100:.2f}%")
    q4.metric("Acumulado 5 años", f"USD {sum(fy_adj):,.0f}",
              f"{(sum(fy_adj)/sum(fy_b)-1)*100:+.2f}% vs base" if sum(fy_b) else None)

    # ---------------- KPIs de exposición / sensibilidad ----------------
    st.subheader("Exposición a Volatilidad · % del gasto proyectado por driver")
    _pesos = df_f.groupby('Driver')['Proyeccion_base'].sum()
    _tot = _pesos.sum() or 1
    x1, x2, x3, x4 = st.columns(4)
    x1.metric("⛽ Combustible", f"{_pesos.get('Combustible', 0)/_tot*100:.1f}%",
              f"USD {simular(df_f, bud_f, shf, 0, 0)[0]['Forecast FY'].sum()-fc_base:,.0f}", delta_color="off")
    x2.metric("💱 Divisas (importado)", f"{_pesos.get('Divisas', 0)/_tot*100:.1f}%",
              f"USD {simular(df_f, bud_f, 0, shx, 0)[0]['Forecast FY'].sum()-fc_base:,.0f}", delta_color="off")
    x3.metric("👷 Mano de obra", f"{_pesos.get('Mano de Obra', 0)/_tot*100:.1f}%",
              f"USD {simular(df_f, bud_f, 0, 0, shl)[0]['Forecast FY'].sum()-fc_base:,.0f}", delta_color="off")
    x4.metric("🏭 Local / IPC", f"{_pesos.get('Local / IPC', 0)/_tot*100:.1f}%", "no sensible", delta_color="off")

    st.write("---")

    # ---------------- GRÁFICOS ----------------
    g1, g2 = st.columns(2)
    with g1:
        st.markdown("**🗓️ Budget mensual 2027 (M USD)** — base vs ajustado")
        graf_m = pd.DataFrame({
            'Base': [bud0[m].sum()/1e6 for m in MESES_27],
            'Ajustado': [bud[m].sum()/1e6 for m in MESES_27],
        }, index=[m.split('-')[0] for m in MESES_27])
        st.bar_chart(graf_m, y_label="M USD", stack=False, use_container_width=True)
    with g2:
        st.markdown("**📈 Trayectoria FY27→FY31 (M USD)**")
        st.line_chart(pd.DataFrame({'Base': [v/1e6 for v in fy_b], 'Ajustado': [v/1e6 for v in fy_adj]},
                                   index=FY_COLS), y_label="M USD", use_container_width=True)

    g3, g4 = st.columns(2)
    with g3:
        st.markdown("**📊 Budget vs Forecast — 2° sem. 2026 (M USD)**")
        bmes = budget_2026/12
        graf_h2 = pd.DataFrame({'Budget (mensual)': [bmes/1e6]*7,
                                'Forecast ajustado': [fcst[m].sum()/1e6 for m in MESES_PROY]},
                               index=[m.split('-')[0] for m in MESES_PROY])
        st.bar_chart(graf_h2, y_label="M USD", stack=False, use_container_width=True)
    with g4:
        st.markdown("**🌐 Aporte de cada palanca al forecast 2026 (M USD)**")
        ap = {}
        for nom, a in [('⛽ Combustible', (shf, 0, 0)), ('💱 Divisas', (0, shx, 0)), ('👷 Mano de obra', (0, 0, shl))]:
            fs, _ = simular(df_f, bud_f, *a)
            ap[nom] = (fs['Forecast FY'].sum()-fc_base)/1e6
        st.bar_chart(pd.DataFrame.from_dict(ap, orient='index', columns=['M USD']),
                     horizontal=True, use_container_width=True)

    # ---------------- TOP MOVERS ----------------
    st.markdown("**🔝 Líneas más impactadas por los shocks actuales (FY27, USD)**")
    movers = bud[['Desc Item', 'VP', 'Classif', 'Driver']].copy()
    movers['Δ FY27'] = bud['FY27'].values - bud0['FY27'].values
    movers = movers[movers['Δ FY27'].abs() > 0].reindex(
        movers['Δ FY27'].abs().sort_values(ascending=False).index).head(10)
    st.dataframe(movers.style.format({'Δ FY27': '{:,.0f}'}), use_container_width=True, hide_index=True)

    # ---------------- INFORME FINANCIERO (GEMINI) ----------------
    st.write("---")
    st.subheader("🧠 Informe Financiero (IA)")
    with st.expander("Configuración de Gemini"):
        st.caption("🔒 La API key se lee de forma privada desde `.streamlit/secrets.toml` "
                   "(o variable de entorno). No se ingresa ni se guarda en la app.")
        modelo = st.text_input("Modelo", value="gemini-2.5-flash")
    if st.button("📝 Generar informe financiero detallado", use_container_width=True):
        pesos = df_f.groupby(df_f['Driver'])['Proyeccion_base'].sum()
        tot = pesos.sum() or 1
        imp_fuel = simular(df_f, bud_f, shf, 0, 0)[0]['Forecast FY'].sum() - fc_base
        imp_fx = simular(df_f, bud_f, 0, shx, 0)[0]['Forecast FY'].sum() - fc_base
        imp_labor = simular(df_f, bud_f, 0, 0, shl)[0]['Forecast FY'].sum() - fc_base
        w_fuel = pesos.get('Combustible', 0)/tot*100
        w_fx = pesos.get('Divisas', 0)/tot*100
        w_labor = pesos.get('Mano de Obra', 0)/tot*100
        w_local = pesos.get('Local / IPC', 0)/tot*100

        # datos para la narrativa (Gemini / local)
        datos = dict(
            budget_2026=budget_2026, forecast_base=fc_base, forecast_adj=fc_total,
            desv=desv, desv_pct=desv_pct, shf=shf, shx=shx, shl=shl,
            imp_fuel=imp_fuel, imp_fx=imp_fx, imp_labor=imp_labor,
            fy_base=[f"{v:,.0f}" for v in fy_b], fy_adj=[f"{v:,.0f}" for v in fy_adj],
            cagr=cagr*100, px=px, w_fuel=w_fuel, w_fx=w_fx, w_labor=w_labor, w_local=w_local)

        # datos completos para el PDF (KPIs + gráficos + tablas)
        rep = dict(
            fecha=datetime.now().strftime("%Y-%m-%d %H:%M"),
            scope=(f"Alcance: {len(vp_sel)} VP, {len(bud):,} lineas budget. "
                   f"Shocks aplicados -> Combustible {shf:+.1f}%, Divisas {shx:+.1f}%, Mano de obra {shl:+.1f}%."),
            budget_2026=budget_2026, forecast_base=fc_base, forecast_adj=fc_total,
            desv=desv, desv_pct=desv_pct,
            fy_b=fy_b, fy_adj=fy_adj,
            cagr_b=((fy_b[4]/fy_b[0])**(1/4)-1)*100 if fy_b[0] > 0 else 0, cagr_adj=cagr*100,
            meses27_labels=[m.split('-')[0] for m in MESES_27],
            bud_mensual_base=[bud0[m].sum() for m in MESES_27],
            bud_mensual_adj=[bud[m].sum() for m in MESES_27],
            fy_labels=FY_COLS,
            impactos={'Combustible': imp_fuel, 'Divisas': imp_fx, 'Mano de obra': imp_labor},
            shf=shf, shx=shx, shl=shl, imp_fuel=imp_fuel, imp_fx=imp_fx, imp_labor=imp_labor,
            w_fuel=w_fuel, w_fx=w_fx, w_labor=w_labor, w_local=w_local, px=px,
            movers=[(r['Desc Item'], r['VP'], r['Driver'], r['Δ FY27']) for _, r in movers.iterrows()],
        )
        with st.spinner("Redactando informe y generando PDF..."):
            texto, fuente = generar_informe(datos, modelo)
            try:
                pdf_bytes = construir_pdf(rep, texto, fuente)
            except Exception as ex:
                pdf_bytes = None
                st.session_state['informe_pdf_error'] = str(ex)
        st.session_state['informe_txt'] = texto
        st.session_state['informe_fuente'] = fuente
        st.session_state['informe_pdf'] = pdf_bytes

    if st.session_state.get('informe_txt'):
        st.success(f"Informe generado · análisis: {st.session_state.get('informe_fuente')}")
        st.markdown(st.session_state['informe_txt'])
        if st.session_state.get('informe_pdf'):
            st.download_button("📄 Descargar informe completo (PDF)", st.session_state['informe_pdf'],
                               file_name="Informe_Financiero.pdf", mime="application/pdf",
                               use_container_width=True)
        else:
            st.info(f"Para el PDF con gráficos instala: `pip install fpdf2 matplotlib`. "
                    f"{st.session_state.get('informe_pdf_error', '')}")
            st.download_button("📥 Descargar informe (.md)", st.session_state['informe_txt'].encode(),
                               file_name="Informe_Financiero.md", mime="text/markdown",
                               use_container_width=True)

    # ---------------- EXPORTACIÓN (ambas hojas, estructura respetada) ----------------
    st.write("---")
    cols_fc = ['Resp', 'Desc Resp', 'VP', 'Gerencia', 'Proc', 'Desc Proc', 'Item', 'Desc Item',
               'Classif', 'CC', 'Driver'] + MESES_REALES + MESES_PROY + ['YTD', 'Forecast FY', 'Budget FY', 'Var']
    cols_fc = [c for c in cols_fc if c in fcst.columns]
    cols_bud = ['Resp', 'Desc Resp', 'VP', 'Gerencia', 'Proc', 'Desc Proc', 'Item', 'Desc Item',
                'Classif', 'CC'] + MESES_27 + FY_COLS   # estructura EXACTA de la hoja base

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as w:
        fcst[cols_fc].to_excel(w, index=False, sheet_name='Forecast 5+7 Proyectado')
        bud[cols_bud].to_excel(w, index=False, sheet_name='BUDGET 2027-2031')
    e1, e2 = st.columns([2, 1])
    e1.markdown(f"**💾 Planilla maestra** — Forecast 5+7 + BUDGET 2027-2031 recalculados "
                f"(⛽ {shf:+.1f}% · 💱 {shx:+.1f}% · 👷 {shl:+.1f}%). La hoja budget conserva mes a mes + FY.")
    e2.download_button("📥 Descargar (.xlsx)", buffer.getvalue(),
                       file_name="Forecast_y_Budget_Quinquenal.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True)

    # ---------------- TABLAS ----------------
    t1, t2 = st.tabs([f"🗓️ BUDGET 2027-2031 ({len(bud):,} líneas)", f"🔍 Forecast 5+7 ({len(fcst):,} líneas)"])
    t1.dataframe(bud[cols_bud], use_container_width=True, height=420)
    t2.dataframe(fcst[cols_fc], use_container_width=True, height=420)

except Exception as e:
    st.error(f"❌ Error al procesar el tablero: {e}")
    st.exception(e)