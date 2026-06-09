"""
Generador de Presupuesto de Construcción — Interfaz Web
Stackbone MVP · Junio 2026
"""

import streamlit as st
import io
import re
import json
import tempfile
from pathlib import Path
from datetime import date

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Presupuesto IA · Construcción",
    page_icon="🏗️",
    layout="centered",
)

# ── Styles ────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main { max-width: 760px; }
    .stButton>button {
        background-color: #1F4E79;
        color: white;
        font-weight: 600;
        border-radius: 6px;
        padding: 0.5rem 2rem;
        border: none;
        width: 100%;
        font-size: 1rem;
    }
    .stButton>button:hover { background-color: #2E75B6; }
    .result-box {
        background: #EBF3FB;
        border-left: 4px solid #1F4E79;
        padding: 1rem 1.2rem;
        border-radius: 0 6px 6px 0;
        margin: 1rem 0;
    }
    .step {
        background: #F8F9FA;
        border: 1px solid #DEE2E6;
        border-radius: 6px;
        padding: 0.6rem 1rem;
        margin: 0.3rem 0;
        font-size: 0.9rem;
    }
    .step.active { border-color: #1F4E79; background: #EBF3FB; }
    .step.done { border-color: #28A745; background: #F0FFF4; }
</style>
""", unsafe_allow_html=True)

# ── Imports funcionales ───────────────────────────────────────────────────────
import pdfplumber
import olefile
import anthropic
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.enums import TA_CENTER

# ── Lógica de extracción y generación (importada de presupuesto.py) ───────────

PROMPT_EXTRACCION = """Eres un estimador de costos de construcción experto en proyectos chilenos.

Se te proporciona el texto de una memoria de cálculo estructural. Extrae la información clave para armar un presupuesto.

Sé CONCISO. Responde ÚNICAMENTE con JSON válido y completo, sin texto adicional.

Formato requerido:
{
  "proyecto": {
    "tipo": "descripción breve del tipo de edificación",
    "superficie_total_estimada": número_en_m2_o_null,
    "num_pisos": número_o_null,
    "ubicacion": "si se menciona, sino null",
    "descripcion": "1-2 oraciones sobre el proyecto"
  },
  "dimensiones_clave": {
    "luz_principal_m": número_o_null,
    "largo_edificio_m": número_o_null,
    "ancho_edificio_m": número_o_null,
    "altura_m": número_o_null,
    "superficie_cubierta_m2": número_o_null,
    "perimetro_m": número_o_null
  },
  "sistemas": {
    "estructura": "descripción breve",
    "cubierta": "descripción breve",
    "fundaciones": "descripción breve",
    "muros": "descripción breve",
    "instalaciones": "descripción o null"
  },
  "materiales_clave": [
    {"item": "nombre", "especificacion": "grado/tipo", "cantidad": número_o_null, "unidad": "unidad", "estimado": true}
  ],
  "notas_presupuesto": ["nota breve"]
}

TEXTO DEL DOCUMENTO:
"""

# ── TABLA DE PRECIOS FIJOS ────────────────────────────────────────────────────
# Precios CLP 2025, trabajo completo (material + mano de obra + herramientas).
# Zona Santa Cruz, VI Región — incluye ~12% flete sobre materiales.
# Fuentes: ONDAC Manual de Precios Chile, MINVU DS27 2025, mercado local.
# Claude SOLO aporta cantidades. Los precios vienen exclusivamente de esta tabla.
#
# IMPORTANTE: las claves deben coincidir EXACTAMENTE con lo que Claude devuelve
# en el campo "precio_clave". No hay matching parcial ni fallback.
#
# Estructura: "clave": ("unidad", precio_clp)
PRECIOS_FIJOS = {
    # ── Obras preliminares ────────────────────────────────────────────────────
    "instalacion faenas":           ("global",  1_500_000),  # bodega, baño, cerco
    "trazado nivelacion":           ("m2",      3_500),
    "excavacion":                   ("m3",      18_000),     # excavación manual+máquina
    "retiro escombros":             ("m3",      14_000),     # retiro y transporte
    "relleno compactado":           ("m3",      22_000),

    # ── Fundaciones ──────────────────────────────────────────────────────────
    "solado hormigon":              ("m3",      85_000),     # H-10 hormigón magro
    "cimiento corrido":             ("m3",      240_000),    # HA H-25 inc. moldaje+fierro
    "zapata aislada":               ("m3",      270_000),    # HA H-25 inc. moldaje+fierro
    "radier":                       ("m2",      28_000),     # H-20 e=10cm inc. membrana
    "enfierradura":                 ("kg",      1_000),      # fierro corrugado puesto

    # ── Estructura ───────────────────────────────────────────────────────────
    "pilar hormigon armado":        ("m3",      480_000),    # inc. moldaje metálico+fierro
    "viga hormigon armado":         ("m3",      460_000),    # inc. moldaje+fierro
    "losa nervada":                 ("m2",      120_000),    # alivianada h=20cm inc. EPS
    "losa maciza":                  ("m2",      95_000),     # e=12cm inc. moldaje+fierro
    "escalera hormigon":            ("global",  2_500_000),  # inc. moldaje, fierro, barandas
    "estructura metalica":          ("kg",      2_200),

    # ── Albañilería ──────────────────────────────────────────────────────────
    "muro bloque":                  ("m2",      52_000),     # bloque 15cm inc. mortero+MO
    "tabique":                      ("m2",      35_000),     # Volcanita doble inc. estructura
    "estuco":                       ("m2",      14_000),     # mortero+yeso, ambas caras
    "pintura muro":                 ("m2",      7_500),      # 2 manos látex + sellador

    # ── Cubierta ─────────────────────────────────────────────────────────────
    "impermeabilizacion":           ("m2",      22_000),     # membrana asfáltica bicapa
    "pendiente hormigon liviano":   ("m2",      12_000),
    "cubierta sandwich":            ("m2",      38_000),     # panel 60mm inc. correas
    "cubierta teja":                ("m2",      28_000),
    "canalon":                      ("ml",      7_000),
    "bajante aguas lluvias":        ("ml",      12_000),     # PVC 110mm inc. receptores

    # ── Instalaciones sanitarias ─────────────────────────────────────────────
    "punto agua potable":           ("punto",   130_000),    # inc. tubería, llaves, fittings
    "punto alcantarillado":         ("punto",   100_000),    # inc. tubería PVC, cámara
    "artefacto sanitario":          ("un",      320_000),    # WC, lavamanos o ducha c/u
    "calefon":                      ("un",      550_000),    # calefón + conexión gas

    # ── Instalaciones eléctricas ─────────────────────────────────────────────
    "punto electrico":              ("punto",   65_000),     # inc. conductor, tubería, salida
    "tablero electrico":            ("un",      420_000),    # inc. termomagnéticas+diferencial
    "acometida electrica":          ("global",  850_000),    # empalme NSEG/SEC
    "punto datos tv":               ("punto",   38_000),

    # ── Terminaciones ────────────────────────────────────────────────────────
    "porcelanato":                  ("m2",      50_000),     # 60x60 inc. adhesivo+fragüe
    "piso laminado":                ("m2",      30_000),     # AC4 8mm inc. fieltro
    "ceramica bano":                ("m2",      42_000),     # muro baño inc. adhesivo
    "cielo yeso":                   ("m2",      26_000),     # metal desplegado+estuco
    "pintura exterior":             ("m2",      9_000),      # impermeabilizante pintante

    # ── Carpintería ──────────────────────────────────────────────────────────
    "puerta interior":              ("un",      250_000),    # entamborada inc. marco+cerradura
    "puerta exterior":              ("un",      650_000),    # maciza/metálica inc. marco
    "ventana aluminio":             ("m2",      120_000),    # corredera vidrio 6mm inc. reja
    "mueble cocina":                ("ml",      380_000),    # alto+bajo MDF inc. cubierta

    # ── Obras exteriores ─────────────────────────────────────────────────────
    "pavimento exterior":           ("m2",      32_000),     # hormigón H-20 e=8cm
    "cierre perimetral":            ("ml",      75_000),     # muro bloque + reja h=1.8m
    "jardin":                       ("m2",      15_000),     # pasto + tierra vegetal
    "aseo obra":                    ("global",  900_000),    # limpieza final + retiro
}

PROMPT_PARTIDAS = """Eres un estimador de costos de construcción para proyectos en Chile.

Con base en el análisis técnico, genera un presupuesto por partidas. SOLO aportas cantidades.
Los precios los asigna el sistema desde una tabla fija — NO los inventes.

Incluye máximo 8-10 partidas, con 3-6 ítems cada una (máximo 50 ítems total).

Para cada ítem debes elegir la "precio_clave" EXACTA de esta lista (copia la clave textualmente):

CLAVES DISPONIBLES (cópialas exactas, sin modificar):
instalacion faenas | trazado nivelacion | excavacion | retiro escombros | relleno compactado
solado hormigon | cimiento corrido | zapata aislada | radier | enfierradura
pilar hormigon armado | viga hormigon armado | losa nervada | losa maciza | escalera hormigon | estructura metalica
muro bloque | tabique | estuco | pintura muro
impermeabilizacion | pendiente hormigon liviano | cubierta sandwich | cubierta teja | canalon | bajante aguas lluvias
punto agua potable | punto alcantarillado | artefacto sanitario | calefon
punto electrico | tablero electrico | acometida electrica | punto datos tv
porcelanato | piso laminado | ceramica bano | cielo yeso | pintura exterior
puerta interior | puerta exterior | ventana aluminio | mueble cocina
pavimento exterior | cierre perimetral | jardin | aseo obra

Campos por ítem:
- partida: Obras Preliminares | Fundaciones | Estructura | Albañilería | Cubierta | Instalaciones Sanitarias | Instalaciones Eléctricas | Terminaciones | Carpintería | Obras Exteriores
- item: descripción breve del trabajo
- precio_clave: UNA clave exacta de la lista anterior
- cantidad: número (si estimas, supuesto: true)
- supuesto: true/false

Devuelve ÚNICAMENTE JSON:
{
  "resumen_proyecto": {"nombre": "...", "tipo": "...", "superficie_m2": número, "descripcion": "..."},
  "partidas": [
    {"partida": "...", "item": "...", "precio_clave": "...", "cantidad": número, "supuesto": false}
  ]
}

ANÁLISIS TÉCNICO:
"""


def extraer_texto_pdf(data: bytes) -> str:
    texto_total = []
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for i, pagina in enumerate(pdf.pages):
            texto = pagina.extract_text()
            if texto and texto.strip():
                texto_total.append(f"[PÁGINA {i+1}]\n{texto}")
    return "\n\n".join(texto_total)


def extraer_texto_doc(data: bytes) -> str:
    with tempfile.NamedTemporaryFile(suffix='.doc', delete=False) as f:
        f.write(data)
        tmp_path = f.name
    ole = olefile.OleFileIO(tmp_path)
    streams = []
    for name in ['WordDocument', '1Table', '0Table']:
        try:
            streams.append(ole.openstream(name).read())
        except Exception:
            pass
    combined = b''.join(streams)
    decoded = combined.decode('utf-16-le', errors='ignore')
    lines = decoded.split('\r')
    good_lines = []
    for line in lines:
        clean = re.sub(r'[^\x20-\x7E\xC0-\xFF\n\t]', '', line).strip()
        if len(clean) > 20:
            good_lines.append(clean)
    import os; os.unlink(tmp_path)
    return '\n'.join(good_lines)


def extraer_texto_docx(data: bytes) -> str:
    from docx import Document as DocxDocument
    doc = DocxDocument(io.BytesIO(data))
    return '\n'.join(p.text for p in doc.paragraphs if p.text.strip())


def parse_json_safe(raw: str) -> dict:
    raw = re.sub(r'^```json\s*', '', raw.strip(), flags=re.MULTILINE)
    raw = re.sub(r'\s*```\s*$', '', raw)
    # Strip trailing commas before closing brackets (common truncation artifact)
    raw = re.sub(r',\s*([}\]])', r'\1', raw)
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        # Try standard structural suffixes first
        for suffix in ['"}', '"]}', ']}', ']}]}', ']}]}]}']:
            try:
                return json.loads(raw + suffix)
            except Exception:
                pass
        # Last resort: truncate to last complete key-value pair and close the object
        # Find last comma at top level and try closing from there
        try:
            last_comma = raw.rfind('",')
            if last_comma > 0:
                trimmed = raw[:last_comma + 1].rstrip(',')
                for suffix in ['"}', '"}]}', '}']:
                    try:
                        return json.loads(trimmed + suffix)
                    except Exception:
                        pass
        except Exception:
            pass
        raise


def generar_presupuesto(texto: str, nombre_proyecto: str, api_key: str):
    client = anthropic.Anthropic(api_key=api_key)

    # Paso 1: extraer datos técnicos (temperature=0 → determinístico)
    r1 = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4096,
        temperature=0,
        messages=[{"role": "user", "content": PROMPT_EXTRACCION + texto[:80_000]}]
    )
    datos = parse_json_safe(r1.content[0].text)

    # Paso 2: generar cantidades por partida (temperature=0 → determinístico)
    r2 = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=16000,
        temperature=0,
        messages=[{"role": "user", "content": PROMPT_PARTIDAS + json.dumps(datos, ensure_ascii=False)}]
    )
    presupuesto = parse_json_safe(r2.content[0].text)

    return datos, presupuesto


def calcular_totales(presupuesto: dict) -> pd.DataFrame:
    """Calcula totales con precios fijos — Claude aporta cantidades y precio_clave exacta."""
    filas = []
    for p in presupuesto.get("partidas", []):
        cantidad = float(p.get("cantidad", 0) or 0)
        clave = (p.get("precio_clave") or "").strip()

        # Lookup exacto — sin fallback para evitar asignaciones incorrectas
        if clave in PRECIOS_FIJOS:
            unidad, pu = PRECIOS_FIJOS[clave]
        else:
            # Clave no reconocida: precio 0, marcado para revisión
            unidad = "?"
            pu = 0

        filas.append({
            "Partida": p.get("partida", ""),
            "Ítem": p.get("item", ""),
            "Clave": clave,
            "Unidad": unidad,
            "Cantidad": cantidad,
            "P.U. (CLP)": pu,
            "Total (CLP)": cantidad * pu,
            "Supuesto": "(*)" if p.get("supuesto") or pu == 0 else "",
            "Nota": "" if pu > 0 else "SIN PRECIO — revisar clave",
        })
    df = pd.DataFrame(filas)
    return df


def build_excel(df: pd.DataFrame, resumen: dict, nombre_proyecto: str, uf_valor: int = 38_500) -> bytes:
    wb = openpyxl.Workbook()

    # ── Hoja 1: Presupuesto ───────────────────────────────────────────────────
    ws = wb.active
    ws.title = "Presupuesto"

    az1 = PatternFill("solid", fgColor="1F4E79")
    az2 = PatternFill("solid", fgColor="2E75B6")
    sub = PatternFill("solid", fgColor="D6E4F0")
    ama = PatternFill("solid", fgColor="FFF2CC")
    ver = PatternFill("solid", fgColor="E2EFDA")
    borde = Border(left=Side(style="thin"), right=Side(style="thin"),
                   top=Side(style="thin"), bottom=Side(style="thin"))

    for col, w in {"A":4,"B":46,"C":9,"D":11,"E":17,"F":18,"G":5,"H":32}.items():
        ws.column_dimensions[col].width = w

    # Título
    ws.merge_cells("A1:H1")
    ws["A1"] = f"PRESUPUESTO DE OBRA — {nombre_proyecto.upper()}"
    ws["A1"].font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = az1; ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:H2")
    ws["A2"] = (f"Tipo: {resumen.get('tipo','')}   |   "
                f"Superficie: {resumen.get('superficie_m2','N/D')} m²   |   "
                f"Fecha: {date.today().strftime('%d/%m/%Y')}")
    ws["A2"].font = Font(name="Calibri", size=9, italic=True, color="FFFFFF")
    ws["A2"].fill = az2; ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[3].height = 5

    # Headers — col A=N°, B=Descripción, C=Unid, D=Cantidad, E=P.U.(CLP), F=Total, G=(*), H=Nota
    for col, h in enumerate(["N°","Ítem / Descripción","Unid.","Cantidad","P.U. (CLP)","Total (CLP)","(*)","Notas"], 1):
        c = ws.cell(row=4, column=col, value=h)
        c.font = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
        c.fill = az1; c.alignment = Alignment(horizontal="center", vertical="center"); c.border = borde
    ws.row_dimensions[4].height = 20

    # Build items with formulas
    row = 5
    subtotal_refs = []   # (partida_name, first_row, last_row) for each group
    item_rows = []       # track all item rows for grand total SUM
    partidas_order = list(dict.fromkeys(df["Partida"].tolist()))

    for partida in partidas_order:
        grupo = df[df["Partida"] == partida]

        # Partida header row
        ws.merge_cells(f"A{row}:H{row}")
        ws[f"A{row}"] = partida.upper()
        ws[f"A{row}"].font = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
        ws[f"A{row}"].fill = az2
        ws[f"A{row}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[row].height = 16
        row += 1

        first_item_row = row
        item_num = 1
        for _, item in grupo.iterrows():
            es_sup = item["Supuesto"] == "(*)"
            fill = ama if es_sup else None
            pu_val = float(item["P.U. (CLP)"])

            vals = [item_num, item["Ítem"], item["Unidad"], item["Cantidad"], pu_val, None, item["Supuesto"], item["Nota"]]
            for ci, val in enumerate(vals, 1):
                c = ws.cell(row=row, column=ci, value=val)
                c.font = Font(name="Calibri", size=9)
                c.border = borde
                if fill: c.fill = fill
                if ci == 1:
                    c.alignment = Alignment(horizontal="center")
                if ci == 2:
                    c.alignment = Alignment(wrap_text=True, indent=1)
                if ci == 4:
                    c.number_format = '#,##0.00'; c.alignment = Alignment(horizontal="right")
                if ci == 5:
                    c.number_format = '#,##0'; c.alignment = Alignment(horizontal="right")
                if ci == 6:
                    # FORMULA: Total = Cantidad × P.U.
                    c.value = f"=D{row}*E{row}"
                    c.number_format = '#,##0'; c.alignment = Alignment(horizontal="right")
                    if fill: c.fill = fill
            ws.row_dimensions[row].height = 15
            item_rows.append(row)
            item_num += 1
            row += 1

        last_item_row = row - 1

        # Subtotal row (formula)
        for ci in range(1, 9):
            c = ws.cell(row=row, column=ci); c.fill = sub; c.border = borde
        ws.cell(row=row, column=2, value=f"SUBTOTAL {partida.upper()}").font = Font(name="Calibri", size=9, bold=True)
        ws.cell(row=row, column=2).fill = sub; ws.cell(row=row, column=2).border = borde
        c_sub = ws.cell(row=row, column=6, value=f"=SUM(F{first_item_row}:F{last_item_row})")
        c_sub.font = Font(name="Calibri", size=9, bold=True)
        c_sub.number_format = '#,##0'; c_sub.alignment = Alignment(horizontal="right")
        c_sub.fill = sub; c_sub.border = borde
        subtotal_refs.append((partida, row))
        ws.row_dimensions[row].height = 16
        row += 1

    # ── Gastos generales block ────────────────────────────────────────────────
    # Subtotal neto = SUM of all subtotal cells
    subtotal_cells = "+".join([f"F{r}" for _, r in subtotal_refs])
    row_sub_neto = row
    ws.merge_cells(f"A{row}:E{row}")
    ws[f"A{row}"] = "SUBTOTAL NETO (sin GG)"
    ws[f"A{row}"].font = Font(name="Calibri", size=9, bold=True)
    ws[f"A{row}"].alignment = Alignment(horizontal="right"); ws[f"A{row}"].fill = sub
    for ci in range(1,9): ws.cell(row=row,column=ci).fill=sub; ws.cell(row=row,column=ci).border=borde
    c_neto = ws.cell(row=row, column=6, value=f"={subtotal_cells}")
    c_neto.number_format = '#,##0'; c_neto.alignment = Alignment(horizontal="right")
    c_neto.font = Font(name="Calibri", size=9, bold=True); c_neto.fill = sub; c_neto.border = borde
    row += 2

    # GG rows referencing the neto subtotal
    gg_items = [
        ("Gastos generales", 0.10),
        ("Utilidad empresa", 0.08),
        ("Imprevistos", 0.05),
    ]
    gg_rows = []
    for label, pct in gg_items:
        ws.merge_cells(f"A{row}:D{row}")
        ws[f"A{row}"] = f"{label} ({int(pct*100)}%)"
        ws[f"A{row}"].font = Font(name="Calibri", size=9); ws[f"A{row}"].fill = ver
        ws[f"A{row}"].alignment = Alignment(horizontal="right")
        for ci in range(1,9): ws.cell(row=row,column=ci).fill=ver; ws.cell(row=row,column=ci).border=borde
        c_gg = ws.cell(row=row, column=6, value=f"=F{row_sub_neto}*{pct}")
        c_gg.number_format = '#,##0'; c_gg.alignment = Alignment(horizontal="right")
        c_gg.font = Font(name="Calibri", size=9); c_gg.fill = ver; c_gg.border = borde
        gg_rows.append(row)
        ws.row_dimensions[row].height = 15
        row += 1

    # ── TOTAL GENERAL ────────────────────────────────────────────────────────
    row += 1
    total_formula = f"=F{row_sub_neto}+" + "+".join([f"F{r}" for r in gg_rows])
    row_total = row
    ws.merge_cells(f"A{row}:E{row}")
    ws[f"A{row}"] = "TOTAL GENERAL (incl. GG, utilidad e imprevistos)"
    ws[f"A{row}"].font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    ws[f"A{row}"].fill = az1; ws[f"A{row}"].alignment = Alignment(horizontal="right", vertical="center")
    ws[f"A{row}"].border = borde
    for ci in [6,7,8]: ws.cell(row=row,column=ci).fill=az1; ws.cell(row=row,column=ci).border=borde
    c_tot = ws.cell(row=row, column=6, value=total_formula)
    c_tot.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    c_tot.fill = az1; c_tot.number_format = '#,##0'; c_tot.alignment = Alignment(horizontal="right"); c_tot.border = borde
    ws.row_dimensions[row].height = 24
    row += 1

    # UF equivalent
    ws.merge_cells(f"A{row}:E{row}")
    ws[f"A{row}"] = f"Equivalente en UF  (1 UF = ${uf_valor:,})"
    ws[f"A{row}"].font = Font(name="Calibri", size=9, italic=True)
    ws[f"A{row}"].alignment = Alignment(horizontal="right")
    c_uf = ws.cell(row=row, column=6, value=f"=F{row_total}/{uf_valor}")
    c_uf.number_format = '#,##0.0" UF"'; c_uf.alignment = Alignment(horizontal="right")
    c_uf.font = Font(name="Calibri", size=9, bold=True)
    row += 2

    # Notas finales
    ws.merge_cells(f"A{row}:H{row}")
    ws[f"A{row}"] = "(*) Ítems en amarillo = cantidades estimadas. Ajusta la columna Cantidad según planos y cubicaciones definitivas — los totales se recalculan automáticamente."
    ws[f"A{row}"].font = Font(name="Calibri", size=8, italic=True, color="666666")
    ws[f"A{row}"].fill = PatternFill("solid", fgColor="FFFFD0")
    row += 1
    ws.merge_cells(f"A{row}:H{row}")
    ws[f"A{row}"] = "PRESUPUESTO REFERENCIAL — No incluye IVA. Precios de referencia en tab 'Tabla de Precios'. Sujeto a cubicaciones y cotizaciones definitivas."
    ws[f"A{row}"].font = Font(name="Calibri", size=8, italic=True, color="CC0000")

    ws.freeze_panes = "B5"

    # ── Hoja 2: Tabla de Precios ──────────────────────────────────────────────
    wp = wb.create_sheet("Tabla de Precios")
    for col, w in {"A":30,"B":10,"C":18,"D":45}.items():
        wp.column_dimensions[col].width = w

    wp.merge_cells("A1:D1")
    wp["A1"] = "TABLA DE PRECIOS UNITARIOS — REFERENCIA"
    wp["A1"].font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    wp["A1"].fill = az1; wp["A1"].alignment = Alignment(horizontal="center", vertical="center")
    wp.row_dimensions[1].height = 22

    wp.merge_cells("A2:D2")
    wp["A2"] = f"Precios CLP 2025 · Zona Santa Cruz, VI Región (incl. ~12% flete) · Trabajo completo: material + mano de obra · Fecha: {date.today().strftime('%d/%m/%Y')}"
    wp["A2"].font = Font(name="Calibri", size=8, italic=True, color="FFFFFF")
    wp["A2"].fill = az2; wp["A2"].alignment = Alignment(horizontal="center")

    for ci, h in enumerate(["Ítem / Descripción", "Unidad", "Precio Unit. (CLP)", "Notas / Alcance"], 1):
        c = wp.cell(row=3, column=ci, value=h)
        c.font = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
        c.fill = az1; c.alignment = Alignment(horizontal="center"); c.border = borde
    wp.row_dimensions[3].height = 18

    # Categorías y descripciones para la tabla de precios
    categorias = {
        "OBRAS PRELIMINARES": [
            ("instalacion faenas",      "Instalación de faenas (bodega, baño, cerco)"),
            ("trazado nivelacion",      "Trazado y nivelación de terreno"),
            ("excavacion",              "Excavación en terreno (manual + máquina)"),
            ("retiro escombros",        "Retiro y transporte de material sobrante"),
            ("relleno compactado",      "Relleno compactado con material seleccionado"),
        ],
        "FUNDACIONES": [
            ("solado hormigon",         "Hormigón de limpieza H-10 (solado bajo cimientos)"),
            ("cimiento corrido",        "Cimiento corrido HA H-25 (inc. moldaje + fierro)"),
            ("zapata aislada",          "Zapata aislada HA H-25 (inc. moldaje + fierro)"),
            ("radier",                  "Radier HA H-20 e=10cm (inc. membrana)"),
            ("enfierradura",            "Enfierradura/acero corrugado puesto en obra"),
        ],
        "ESTRUCTURA": [
            ("pilar hormigon armado",   "Pilar HA H-25 (inc. moldaje metálico + fierro)"),
            ("viga hormigon armado",    "Viga/dala HA H-25 (inc. moldaje + fierro)"),
            ("losa nervada",            "Losa nervada aligerada h=20cm (inc. EPS + compresión)"),
            ("losa maciza",             "Losa maciza e=12cm (inc. moldaje + fierro)"),
            ("escalera hormigon",       "Escalera HA (inc. moldaje, fierro, baranda)"),
            ("estructura metalica",     "Estructura metálica (perfiles acero puesto)"),
        ],
        "ALBAÑILERÍA": [
            ("muro bloque",             "Muro bloque hormigón 15cm (inc. mortero + MO)"),
            ("tabique",                 "Tabique Volcanita/yeso-cartón (inc. estructura)"),
            ("estuco",                  "Estuco mortero cemento-arena 1:4 (ambas caras)"),
            ("pintura muro",            "Pintura interior 2 manos látex + sellador"),
        ],
        "CUBIERTA": [
            ("impermeabilizacion",      "Impermeabilización membrana asfáltica bicapa"),
            ("pendiente hormigon liviano", "Formación de pendientes con hormigón liviano"),
            ("cubierta sandwich",       "Cubierta panel sandwich 60mm (inc. correas)"),
            ("cubierta teja",           "Cubierta teja (inc. estructura + membrana)"),
            ("canalon",                 "Canalón zinc/PVC (inc. soportes)"),
            ("bajante aguas lluvias",   "Bajante aguas lluvias PVC 110mm (inc. receptores)"),
        ],
        "INSTALACIONES SANITARIAS": [
            ("punto agua potable",      "Punto agua potable (inc. tubería + llaves + fittings)"),
            ("punto alcantarillado",    "Punto alcantarillado (inc. tubería PVC + cámara)"),
            ("artefacto sanitario",     "Artefacto sanitario c/u (WC, lavamanos o ducha)"),
            ("calefon",                 "Calefón a gas + conexión (inc. tubería + accesorios)"),
        ],
        "INSTALACIONES ELÉCTRICAS": [
            ("punto electrico",         "Punto eléctrico (inc. conductor, tubería, salida)"),
            ("tablero electrico",       "Tablero eléctrico (inc. termomagnéticas + diferencial)"),
            ("acometida electrica",     "Acometida + empalme NSEG/SEC"),
            ("punto datos tv",          "Punto red datos / TV cable"),
        ],
        "TERMINACIONES": [
            ("porcelanato",             "Porcelanato piso 60×60cm (inc. adhesivo + fragüe)"),
            ("piso laminado",           "Piso laminado AC4 8mm (inc. fieltro amortiguador)"),
            ("ceramica bano",           "Cerámica muro baño (inc. adhesivo + fragüe)"),
            ("cielo yeso",              "Cielo falso yeso metal desplegado (inc. estuco)"),
            ("pintura exterior",        "Pintura exterior impermeabilizante 2 manos"),
        ],
        "CARPINTERÍA": [
            ("puerta interior",         "Puerta interior entamborada (inc. marco + cerradura)"),
            ("puerta exterior",         "Puerta exterior maciza/metálica (inc. marco + chapa)"),
            ("ventana aluminio",        "Ventana aluminio vidrio 6mm (inc. reja protectora)"),
            ("mueble cocina",           "Mueble cocina alto+bajo MDF (inc. cubierta)"),
        ],
        "OBRAS EXTERIORES": [
            ("pavimento exterior",      "Pavimento hormigón H-20 e=8cm exterior"),
            ("cierre perimetral",       "Cierre perimetral muro bloque + reja h=1.8m"),
            ("jardin",                  "Jardín pasto natural + tierra vegetal"),
            ("aseo obra",               "Aseo final obra + retiro escombros + entrega"),
        ],
    }

    pr = 4  # precio row
    for cat, items in categorias.items():
        # Category header
        wp.merge_cells(f"A{pr}:D{pr}")
        wp[f"A{pr}"] = cat
        wp[f"A{pr}"].font = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
        wp[f"A{pr}"].fill = az2; wp[f"A{pr}"].alignment = Alignment(indent=1)
        wp.row_dimensions[pr].height = 16
        pr += 1
        for clave, desc in items:
            unidad, precio = PRECIOS_FIJOS.get(clave, ("?", 0))
            row_fill = PatternFill("solid", fgColor="F2F2F2") if pr % 2 == 0 else None
            for ci, val in enumerate([desc, unidad, precio, f"Clave: {clave}"], 1):
                c = wp.cell(row=pr, column=ci, value=val)
                c.font = Font(name="Calibri", size=9)
                c.border = borde
                if row_fill: c.fill = row_fill
                if ci == 3:
                    c.number_format = '#,##0'; c.alignment = Alignment(horizontal="right")
                    c.font = Font(name="Calibri", size=9, bold=True)
                if ci == 4:
                    c.font = Font(name="Calibri", size=8, italic=True, color="666666")
            wp.row_dimensions[pr].height = 14
            pr += 1
        pr += 1  # blank row between categories

    # Footer note on pricing tab
    wp.merge_cells(f"A{pr}:D{pr}")
    wp[f"A{pr}"] = "Para actualizar precios: modifica la columna 'Precio Unit. (CLP)' en esta tabla y actualiza los valores correspondientes en la hoja Presupuesto."
    wp[f"A{pr}"].font = Font(name="Calibri", size=8, italic=True, color="CC0000")

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_pdf(df: pd.DataFrame, resumen: dict, datos: dict, nombre_proyecto: str) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter,
                            rightMargin=1.8*cm, leftMargin=1.8*cm,
                            topMargin=1.5*cm, bottomMargin=1.5*cm)
    estilos = getSampleStyleSheet()
    subtotal_pdf = df["Total (CLP)"].sum()
    total = subtotal_pdf * (1 + 0.10 + 0.08 + 0.05)
    uf_val = 38_500
    total_uf = total / uf_val

    E = lambda name, **kw: ParagraphStyle(name, parent=estilos["Normal"], **kw)
    s_titulo   = E("t", fontSize=17, fontName="Helvetica-Bold", textColor=colors.HexColor("#1F4E79"), spaceAfter=2)
    s_sub      = E("s", fontSize=10, fontName="Helvetica", textColor=colors.HexColor("#2E75B6"), spaceAfter=6)
    s_seccion  = E("sec", fontSize=9, fontName="Helvetica-Bold", textColor=colors.white,
                   backColor=colors.HexColor("#1F4E79"), borderPad=4, spaceAfter=4, spaceBefore=8)
    s_normal   = E("n", fontSize=8.5, fontName="Helvetica", leading=12)
    s_nota     = E("no", fontSize=7.5, fontName="Helvetica-Oblique", textColor=colors.HexColor("#666666"), spaceAfter=3)
    s_disc     = E("d", fontSize=7, fontName="Helvetica-Oblique", textColor=colors.HexColor("#CC0000"), alignment=TA_CENTER)

    contenido = []
    contenido.append(Paragraph("PRESUPUESTO REFERENCIAL DE OBRA", s_titulo))
    contenido.append(Paragraph(nombre_proyecto.upper(),
        E("np", fontSize=13, fontName="Helvetica-Bold", textColor=colors.HexColor("#2E75B6"), spaceAfter=2)))
    contenido.append(Paragraph(
        f"Tipo: {resumen.get('tipo','N/D')}  |  "
        f"Superficie: {resumen.get('superficie_m2','N/D')} m²  |  "
        f"Fecha: {date.today().strftime('%d/%m/%Y')}", s_sub))
    contenido.append(HRFlowable(width="100%", thickness=2, color=colors.HexColor("#1F4E79"), spaceAfter=8))

    # Descripción
    contenido.append(Paragraph("DESCRIPCIÓN DEL PROYECTO", s_seccion))
    desc = resumen.get("descripcion") or datos.get("proyecto", {}).get("descripcion", "")
    if desc:
        contenido.append(Paragraph(desc, s_normal))
    contenido.append(Spacer(1, 0.3*cm))

    # Resumen por partida
    contenido.append(Paragraph("RESUMEN DE COSTOS POR PARTIDA", s_seccion))
    res_partidas = df.groupby("Partida")["Total (CLP)"].sum().reset_index()
    res_partidas = res_partidas.sort_values("Total (CLP)", ascending=False)
    tbl_data = [["Partida", "Total (CLP)", "% del Total"]]
    for _, r in res_partidas.iterrows():
        pct = r["Total (CLP)"] / total * 100 if total else 0
        tbl_data.append([r["Partida"], f"$ {r['Total (CLP)']:,.0f}", f"{pct:.1f}%"])
    tbl = Table(tbl_data, colWidths=[9*cm, 5*cm, 3.5*cm])
    tbl.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#1F4E79")),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1),8),
        ("FONTNAME",(0,1),(-1,-1),"Helvetica"),
        ("ALIGN",(1,0),(-1,-1),"RIGHT"),("ALIGN",(0,0),(0,-1),"LEFT"),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, colors.HexColor("#EBF3FB")]),
        ("GRID",(0,0),(-1,-1),0.5,colors.HexColor("#CCCCCC")),
        ("TOPPADDING",(0,0),(-1,-1),3),("BOTTOMPADDING",(0,0),(-1,-1),3),
        ("LEFTPADDING",(0,0),(-1,-1),6),
    ]))
    contenido.append(tbl)
    contenido.append(Spacer(1, 0.4*cm))

    # Total
    tbl_total = Table([["TOTAL GENERAL (incl. GG, utilidad e imprevistos)",
                        f"$ {total:,.0f}", f"{total_uf:,.0f} UF"]],
                      colWidths=[9*cm, 5*cm, 3.5*cm])
    tbl_total.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),colors.HexColor("#1F4E79")),
        ("TEXTCOLOR",(0,0),(-1,-1),colors.white),
        ("FONTNAME",(0,0),(-1,-1),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1),9),
        ("ALIGN",(1,0),(-1,-1),"RIGHT"),("ALIGN",(0,0),(0,-1),"LEFT"),
        ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5),
        ("LEFTPADDING",(0,0),(-1,-1),6),
    ]))
    contenido.append(tbl_total)
    contenido.append(Spacer(1, 0.5*cm))

    # Top 5
    contenido.append(Paragraph("PRINCIPALES IMPULSORES DE COSTO", s_seccion))
    top5_data = [["#", "Partida", "Monto (CLP)", "% del Total"]]
    for i, (_, r) in enumerate(res_partidas.head(5).iterrows(), 1):
        pct = r["Total (CLP)"] / total * 100 if total else 0
        top5_data.append([str(i), r["Partida"], f"$ {r['Total (CLP)']:,.0f}", f"{pct:.1f}%"])
    tbl5 = Table(top5_data, colWidths=[1*cm, 9*cm, 5*cm, 3.5*cm])
    tbl5.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#2E75B6")),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1),8.5),("FONTNAME",(0,1),(-1,-1),"Helvetica"),
        ("ALIGN",(2,0),(-1,-1),"RIGHT"),("ALIGN",(0,0),(1,-1),"LEFT"),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, colors.HexColor("#EBF3FB")]),
        ("GRID",(0,0),(-1,-1),0.5,colors.HexColor("#CCCCCC")),
        ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),
        ("LEFTPADDING",(0,0),(-1,-1),6),
    ]))
    contenido.append(tbl5)
    contenido.append(Spacer(1, 0.5*cm))

    # Supuestos
    contenido.append(Paragraph("SUPUESTOS Y CONDICIONES", s_seccion))
    for nota in [
        "Precios referenciales de mercado chileno 2025, no incluyen IVA.",
        "Incremento de 12% en materiales por flete a zona rural (Santa Cruz, VI Región).",
        "Ítems marcados como supuestos requieren validación con planos definitivos.",
        "GG (10%), utilidad empresa (8%) e imprevistos (5%) incluidos en el total.",
        f"Equivalencia UF calculada a valor referencial ${uf_val:,}/UF.",
    ] + datos.get("notas_presupuesto", [])[:3]:
        contenido.append(Paragraph(f"• {nota}", s_nota))

    contenido.append(Spacer(1, 0.3*cm))
    contenido.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#CCCCCC"), spaceAfter=4))
    contenido.append(Paragraph(
        "DOCUMENTO REFERENCIAL — Generado con asistencia de IA. "
        "Requiere validación profesional antes de ser utilizado en licitaciones o contratos.", s_disc))

    doc.build(contenido)
    return buf.getvalue()


# ── UI ────────────────────────────────────────────────────────────────────────

# Session state — persiste resultados entre re-runs (evita reset al descargar)
if "resultado" not in st.session_state:
    st.session_state.resultado = None

st.markdown("## 🏗️ Generador de Presupuesto de Construcción")
st.markdown(
    "Sube la memoria de cálculo o especificaciones técnicas del proyecto. "
    "La IA analiza el documento y genera automáticamente el presupuesto por partidas."
)
st.divider()

# API key desde secrets
api_key = st.secrets.get("ANTHROPIC_API_KEY", "") if hasattr(st, "secrets") else ""

col1, col2 = st.columns([2, 1])
with col1:
    nombre_proyecto = st.text_input(
        "Nombre del proyecto",
        placeholder="Ej: Casa Familia Martínez — Santa Cruz",
        help="Este nombre aparecerá en los documentos de salida."
    )
with col2:
    uf_valor = st.number_input("Valor UF (CLP)", value=38_500, step=100,
                                help="Actualiza con el valor del día.")

archivo = st.file_uploader(
    "Memoria de cálculo / Especificaciones técnicas",
    type=["pdf", "doc", "docx"],
    help="Formatos soportados: PDF, DOC, DOCX"
)

if archivo:
    st.caption(f"📄 {archivo.name}  ({archivo.size/1024:.0f} KB)")

st.divider()

if st.button("⚡ Generar Presupuesto", disabled=not (archivo and nombre_proyecto)):

    if not api_key:
        st.error("No se encontró la API key de Anthropic. Configura ANTHROPIC_API_KEY en los secrets de la app.")
        st.stop()

    # Limpiar resultado anterior
    st.session_state.resultado = None

    data = archivo.read()
    ext = Path(archivo.name).suffix.lower()

    progreso = st.container()
    with progreso:
        st.markdown("**Procesando...**")
        p1 = st.empty()
        p2 = st.empty()
        p3 = st.empty()
        p4 = st.empty()

    try:
        p1.markdown("⏳ **[1/4]** Leyendo documento...")
        if ext == ".pdf":
            texto = extraer_texto_pdf(data)
        elif ext == ".doc":
            texto = extraer_texto_doc(data)
        else:
            texto = extraer_texto_docx(data)
        p1.markdown(f"✅ **[1/4]** Documento leído — {len(texto):,} caracteres extraídos")

        p2.markdown("⏳ **[2/4]** Analizando especificaciones técnicas con IA...")
        datos, presupuesto = generar_presupuesto(texto, nombre_proyecto, api_key)
        resumen = presupuesto.get("resumen_proyecto", {})
        p2.markdown(f"✅ **[2/4]** Proyecto identificado: {resumen.get('tipo','N/D')} · {resumen.get('superficie_m2','N/D')} m²")

        p3.markdown("⏳ **[3/4]** Calculando presupuesto...")
        df = calcular_totales(presupuesto)
        if uf_valor != 38_500:
            df["P.U. (CLP)"] = df["P.U. (CLP)"] * uf_valor / 38_500
        df["Total (CLP)"] = df["Cantidad"] * df["P.U. (CLP)"]
        subtotal = df["Total (CLP)"].sum()
        total = subtotal * (1 + 0.10 + 0.08 + 0.05)
        total_uf = total / uf_valor
        n_items = len(presupuesto.get("partidas", []))
        sin_precio = int((df["P.U. (CLP)"] == 0).sum())
        msg = f"✅ **[3/4]** {n_items} ítems · Total: **${total:,.0f} CLP** ({total_uf:,.0f} UF)"
        if sin_precio > 0:
            msg += f" · ⚠️ {sin_precio} ítems sin precio (clave no reconocida)"
        p3.markdown(msg)

        p4.markdown("⏳ **[4/4]** Generando archivos...")
        excel_bytes = build_excel(df, resumen, nombre_proyecto, int(uf_valor))
        pdf_bytes = build_pdf(df, resumen, datos, nombre_proyecto)
        p4.markdown("✅ **[4/4]** Archivos listos para descargar")

        # Guardar en session_state — sobrevive re-runs por clicks de descarga
        nombre_slug = nombre_proyecto.lower().replace(" ", "_").replace("/", "-")[:40]
        st.session_state.resultado = {
            "df": df, "resumen": resumen, "datos": datos,
            "excel_bytes": excel_bytes, "pdf_bytes": pdf_bytes,
            "total": total, "total_uf": total_uf, "n_items": n_items,
            "nombre_slug": nombre_slug, "uf_valor": uf_valor,
        }

    except Exception as e:
        st.error(f"Error al procesar el documento: {e}")
        st.exception(e)

# ── Mostrar resultados desde session_state (persiste sin re-generar) ──────────
if st.session_state.resultado:
    r = st.session_state.resultado
    df, resumen, datos = r["df"], r["resumen"], r["datos"]
    total, total_uf = r["total"], r["total_uf"]
    n_items, nombre_slug = r["n_items"], r["nombre_slug"]
    excel_bytes, pdf_bytes = r["excel_bytes"], r["pdf_bytes"]

    st.divider()
    st.success("**Presupuesto generado correctamente**")

    st.markdown(f"""
    <div class="result-box">
        <b>Proyecto:</b> {resumen.get('tipo','N/D')}<br>
        <b>Superficie:</b> {resumen.get('superficie_m2','N/D')} m²<br>
        <b>Ítems presupuestados:</b> {n_items}<br>
        <b>Total estimado:</b> ${total:,.0f} CLP &nbsp;·&nbsp; {total_uf:,.0f} UF
    </div>
    """, unsafe_allow_html=True)

    dc1, dc2 = st.columns(2)
    with dc1:
        st.download_button(
            label="📥 Descargar Excel (presupuesto detallado)",
            data=excel_bytes,
            file_name=f"presupuesto_{nombre_slug}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with dc2:
        st.download_button(
            label="📥 Descargar PDF (resumen ejecutivo)",
            data=pdf_bytes,
            file_name=f"resumen_{nombre_slug}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

    # Vista previa del resumen por partidas
    with st.expander("Ver resumen por partidas"):
        res = df.groupby("Partida")["Total (CLP)"].sum().reset_index()
        res = res.sort_values("Total (CLP)", ascending=False)
        res["% del Total"] = (res["Total (CLP)"] / total * 100).round(1).astype(str) + "%"
        res["Total (CLP)"] = res["Total (CLP)"].apply(lambda x: f"${x:,.0f}")
        st.dataframe(res, hide_index=True, use_container_width=True)

# Footer
st.divider()
st.caption("Presupuesto referencial generado con IA · Precios mercado Chile 2025 · No incluye IVA · Requiere validación profesional")
