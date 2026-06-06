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

PROMPT_PARTIDAS = """Eres un estimador de costos de construcción para proyectos en Chile.

Con base en el siguiente análisis técnico de un proyecto, genera un presupuesto por partidas.

Incluye máximo 8-10 partidas principales, con 3-6 ítems cada una (máximo 50 ítems en total).

Para cada ítem:
- partida: categoría principal
- item: descripción concisa
- unidad: m2, m3, kg, ml, un, global, punto
- cantidad: número estimado
- precio_unitario_clp: precio referencial CLP 2025
- supuesto: true si la cantidad es estimada, false si es del documento
- nota: máximo 10 palabras si necesario

Usa precios de mercado chilenos 2025. Zona rural (Santa Cruz, VI Región): +12% en materiales.

Devuelve ÚNICAMENTE JSON:
{
  "resumen_proyecto": {
    "nombre": "...",
    "tipo": "...",
    "superficie_m2": número,
    "descripcion": "..."
  },
  "partidas": [
    {"partida": "...", "item": "...", "unidad": "...", "cantidad": número,
     "precio_unitario_clp": número, "supuesto": false, "nota": "..."}
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
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        for suffix in [']}', ']}]}', ']}]}]}']:
            try:
                return json.loads(raw + suffix)
            except Exception:
                pass
        raise


def generar_presupuesto(texto: str, nombre_proyecto: str, api_key: str):
    client = anthropic.Anthropic(api_key=api_key)

    # Paso 1: extraer datos técnicos
    r1 = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=2048,
        messages=[{"role": "user", "content": PROMPT_EXTRACCION + texto[:80_000]}]
    )
    datos = parse_json_safe(r1.content[0].text)

    # Paso 2: generar partidas
    r2 = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=16000,
        messages=[{"role": "user", "content": PROMPT_PARTIDAS + json.dumps(datos, ensure_ascii=False)}]
    )
    presupuesto = parse_json_safe(r2.content[0].text)

    return datos, presupuesto


def calcular_totales(presupuesto: dict) -> pd.DataFrame:
    filas = []
    for p in presupuesto.get("partidas", []):
        cantidad = float(p.get("cantidad", 0) or 0)
        pu = float(p.get("precio_unitario_clp", 0) or 0)
        filas.append({
            "Partida": p.get("partida", ""),
            "Ítem": p.get("item", ""),
            "Unidad": p.get("unidad", ""),
            "Cantidad": cantidad,
            "P.U. (CLP)": pu,
            "Total (CLP)": cantidad * pu,
            "Supuesto": "(*)" if p.get("supuesto") else "",
            "Nota": p.get("nota", ""),
        })
    df = pd.DataFrame(filas)
    subtotal = df["Total (CLP)"].sum()
    extras = pd.DataFrame([
        {"Partida": "Gastos Generales", "Ítem": "Gastos generales (10%)", "Unidad": "global",
         "Cantidad": 1, "P.U. (CLP)": subtotal * 0.10, "Total (CLP)": subtotal * 0.10, "Supuesto": "", "Nota": ""},
        {"Partida": "Gastos Generales", "Ítem": "Utilidad empresa (8%)", "Unidad": "global",
         "Cantidad": 1, "P.U. (CLP)": subtotal * 0.08, "Total (CLP)": subtotal * 0.08, "Supuesto": "", "Nota": ""},
        {"Partida": "Gastos Generales", "Ítem": "Imprevistos (5%)", "Unidad": "global",
         "Cantidad": 1, "P.U. (CLP)": subtotal * 0.05, "Total (CLP)": subtotal * 0.05, "Supuesto": "", "Nota": ""},
    ])
    return pd.concat([df, extras], ignore_index=True)


def build_excel(df: pd.DataFrame, resumen: dict, nombre_proyecto: str) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Presupuesto"
    anchos = {"A": 22, "B": 45, "C": 10, "D": 12, "E": 16, "F": 18, "G": 5, "H": 35}
    for col, w in anchos.items():
        ws.column_dimensions[col].width = w

    fondo_azul    = PatternFill("solid", fgColor="1F4E79")
    fondo_azul2   = PatternFill("solid", fgColor="2E75B6")
    fondo_sub     = PatternFill("solid", fgColor="D6E4F0")
    fondo_amarillo = PatternFill("solid", fgColor="FFF2CC")
    borde = Border(left=Side(style="thin"), right=Side(style="thin"),
                   top=Side(style="thin"), bottom=Side(style="thin"))

    # Título
    ws.merge_cells("A1:H1")
    ws["A1"] = f"PRESUPUESTO DE OBRA — {nombre_proyecto.upper()}"
    ws["A1"].font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = fondo_azul
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:H2")
    ws["A2"] = (f"Tipo: {resumen.get('tipo','')}   |   "
                f"Superficie: {resumen.get('superficie_m2','N/D')} m²   |   "
                f"Fecha: {date.today().strftime('%d/%m/%Y')}")
    ws["A2"].font = Font(name="Calibri", size=9, italic=True, color="FFFFFF")
    ws["A2"].fill = fondo_azul2
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[3].height = 6

    # Headers
    headers = ["Partida", "Ítem / Descripción", "Unidad", "Cantidad", "P.U. (CLP)", "Total (CLP)", "(*)", "Notas"]
    row = 4
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=col, value=h)
        c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        c.fill = fondo_azul
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = borde
    ws.row_dimensions[row].height = 22

    row = 5
    partidas_order = list(dict.fromkeys(df["Partida"].tolist()))
    for partida in partidas_order:
        grupo = df[df["Partida"] == partida]
        ws.merge_cells(f"A{row}:H{row}")
        ws[f"A{row}"] = partida.upper()
        ws[f"A{row}"].font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        ws[f"A{row}"].fill = fondo_azul2
        ws[f"A{row}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[row].height = 18
        row += 1
        subtotal = 0
        for _, item in grupo.iterrows():
            es_sup = item["Supuesto"] == "(*)"
            for col, val in enumerate([
                "", item["Ítem"], item["Unidad"], item["Cantidad"],
                item["P.U. (CLP)"], item["Total (CLP)"], item["Supuesto"], item["Nota"]
            ], 1):
                c = ws.cell(row=row, column=col, value=val)
                c.font = Font(name="Calibri", size=9)
                c.border = borde
                if es_sup:
                    c.fill = fondo_amarillo
                if col == 4:
                    c.number_format = '#,##0.00'; c.alignment = Alignment(horizontal="right")
                if col in (5, 6):
                    c.number_format = '#,##0'; c.alignment = Alignment(horizontal="right")
                if col == 2:
                    c.alignment = Alignment(wrap_text=True, indent=1)
            subtotal += float(item["Total (CLP)"])
            ws.row_dimensions[row].height = 15
            row += 1
        # Subtotal
        for col in range(1, 9):
            c = ws.cell(row=row, column=col)
            c.fill = fondo_sub; c.border = borde
        ws.cell(row=row, column=2, value=f"SUBTOTAL {partida.upper()}").font = Font(name="Calibri", size=9, bold=True)
        ws.cell(row=row, column=2).fill = fondo_sub; ws.cell(row=row, column=2).border = borde
        c_s = ws.cell(row=row, column=6, value=subtotal)
        c_s.font = Font(name="Calibri", size=9, bold=True)
        c_s.number_format = '#,##0'; c_s.alignment = Alignment(horizontal="right")
        c_s.fill = fondo_sub; c_s.border = borde
        ws.row_dimensions[row].height = 16
        row += 1

    # Total general
    total = df["Total (CLP)"].sum()
    uf = total / 38_500
    row += 1
    ws.merge_cells(f"A{row}:E{row}")
    ws[f"A{row}"] = "TOTAL GENERAL (con GG, utilidad e imprevistos)"
    ws[f"A{row}"].font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    ws[f"A{row}"].fill = fondo_azul
    ws[f"A{row}"].alignment = Alignment(horizontal="right", vertical="center")
    ws[f"A{row}"].border = borde
    c_t = ws.cell(row=row, column=6, value=total)
    c_t.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    c_t.fill = fondo_azul; c_t.number_format = '#,##0'; c_t.alignment = Alignment(horizontal="right"); c_t.border = borde
    for col in [7, 8]:
        ws.cell(row=row, column=col).fill = fondo_azul; ws.cell(row=row, column=col).border = borde
    ws.row_dimensions[row].height = 22
    row += 1
    ws.merge_cells(f"A{row}:E{row}")
    ws[f"A{row}"] = "Equivalente en UF (1 UF = $38.500)"
    ws[f"A{row}"].alignment = Alignment(horizontal="right")
    ws[f"A{row}"].font = Font(name="Calibri", size=9, italic=True)
    c_uf = ws.cell(row=row, column=6, value=round(uf, 1))
    c_uf.font = Font(name="Calibri", size=9, bold=True); c_uf.number_format = '#,##0.0" UF"'; c_uf.alignment = Alignment(horizontal="right")
    row += 2
    ws.merge_cells(f"A{row}:H{row}")
    ws[f"A{row}"] = "(*) Ítems en amarillo: cantidades estimadas. Validar con planos definitivos y cubicaciones."
    ws[f"A{row}"].font = Font(name="Calibri", size=8, italic=True, color="666666")
    ws[f"A{row}"].fill = PatternFill("solid", fgColor="FFFFD0")
    row += 1
    ws.merge_cells(f"A{row}:H{row}")
    ws[f"A{row}"] = "PRESUPUESTO REFERENCIAL — No incluye IVA. Sujeto a cubicaciones y cotizaciones definitivas."
    ws[f"A{row}"].font = Font(name="Calibri", size=8, italic=True, color="CC0000")
    ws.freeze_panes = "A5"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_pdf(df: pd.DataFrame, resumen: dict, datos: dict, nombre_proyecto: str) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter,
                            rightMargin=1.8*cm, leftMargin=1.8*cm,
                            topMargin=1.5*cm, bottomMargin=1.5*cm)
    estilos = getSampleStyleSheet()
    total = df["Total (CLP)"].sum()
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

st.markdown("## 🏗️ Generador de Presupuesto de Construcción")
st.markdown(
    "Sube la memoria de cálculo o especificaciones técnicas del proyecto. "
    "La IA analiza el documento y genera automáticamente el presupuesto por partidas."
)
st.divider()

# API key desde secrets o input manual (para desarrollo local)
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
    ext = Path(archivo.name).suffix.lower()
    st.caption(f"📄 {archivo.name}  ({archivo.size/1024:.0f} KB)")

st.divider()

if st.button("⚡ Generar Presupuesto", disabled=not (archivo and nombre_proyecto)):

    if not api_key:
        st.error("No se encontró la API key de Anthropic. Configura ANTHROPIC_API_KEY en los secrets de la app.")
        st.stop()

    data = archivo.read()
    ext = Path(archivo.name).suffix.lower()

    # Contenedor de progreso
    progreso = st.container()
    with progreso:
        st.markdown("**Procesando...**")
        p1 = st.empty()
        p2 = st.empty()
        p3 = st.empty()
        p4 = st.empty()

    try:
        # 1. Extraer texto
        p1.markdown("⏳ **[1/4]** Leyendo documento...")
        if ext == ".pdf":
            texto = extraer_texto_pdf(data)
        elif ext == ".doc":
            texto = extraer_texto_doc(data)
        else:
            texto = extraer_texto_docx(data)
        p1.markdown(f"✅ **[1/4]** Documento leído — {len(texto):,} caracteres extraídos")

        # 2. Análisis IA
        p2.markdown("⏳ **[2/4]** Analizando especificaciones técnicas con IA...")
        datos, presupuesto = generar_presupuesto(texto, nombre_proyecto, api_key)
        resumen = presupuesto.get("resumen_proyecto", {})
        p2.markdown(f"✅ **[2/4]** Proyecto identificado: {resumen.get('tipo','N/D')} · {resumen.get('superficie_m2','N/D')} m²")

        # 3. Calcular totales
        p3.markdown("⏳ **[3/4]** Calculando presupuesto...")
        df = calcular_totales(presupuesto)
        df["Total (CLP)"] = df["Total (CLP)"].apply(lambda x: x * uf_valor / 38_500 if uf_valor != 38_500 else x)
        total = df["Total (CLP)"].sum()
        total_uf = total / uf_valor
        n_items = len(presupuesto.get("partidas", []))
        p3.markdown(f"✅ **[3/4]** {n_items} ítems generados · Total: **${total:,.0f} CLP** ({total_uf:,.0f} UF)")

        # 4. Generar archivos
        p4.markdown("⏳ **[4/4]** Generando archivos...")
        excel_bytes = build_excel(df, resumen, nombre_proyecto)
        pdf_bytes = build_pdf(df, resumen, datos, nombre_proyecto)
        p4.markdown("✅ **[4/4]** Archivos listos para descargar")

        # Resultados
        st.divider()
        st.success(f"**Presupuesto generado correctamente**")

        st.markdown(f"""
        <div class="result-box">
            <b>Proyecto:</b> {resumen.get('tipo','N/D')}<br>
            <b>Superficie:</b> {resumen.get('superficie_m2','N/D')} m²<br>
            <b>Ítems presupuestados:</b> {n_items}<br>
            <b>Total estimado:</b> ${total:,.0f} CLP &nbsp;·&nbsp; {total_uf:,.0f} UF
        </div>
        """, unsafe_allow_html=True)

        nombre_slug = nombre_proyecto.lower().replace(" ", "_").replace("/", "-")[:40]
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

    except Exception as e:
        st.error(f"Error al procesar el documento: {e}")
        st.exception(e)

# Footer
st.divider()
st.caption("Presupuesto referencial generado con IA · Precios mercado Chile 2025 · No incluye IVA · Requiere validación profesional")
