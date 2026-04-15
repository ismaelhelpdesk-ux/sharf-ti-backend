"""
core/pdf.py — Generador de acta PDF con ReportLab
Más rápido y confiable que html2canvas + GAS.
Genera el mismo formato que el acta actual.
"""

import io, base64, logging
from datetime import datetime
from typing   import Optional

from reportlab.lib           import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units     import mm
from reportlab.lib.styles    import ParagraphStyle
from reportlab.platypus      import (SimpleDocTemplate, Table, TableStyle,
                                     Paragraph, Spacer, Image, HRFlowable)
from reportlab.lib.enums     import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.pdfbase       import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from core.models import PayloadDevolucion

log = logging.getLogger("sharf.pdf")

# ── Paleta Scharff ────────────────────────────────────────────────────────
C_ROJO   = colors.HexColor("#9B1035")
C_ROSA   = colors.HexColor("#FF6568")
C_PALIDO = colors.HexColor("#FFC1C2")
C_BLANCO = colors.white
C_NEGRO  = colors.black
C_GRIS   = colors.HexColor("#f5f5f5")
C_GRIS2  = colors.HexColor("#e0e0e0")

def _hoy() -> str:
    return datetime.now().strftime("%d/%m/%Y")

def _esc(v) -> str:
    return str(v or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def generar_acta_pdf(p: PayloadDevolucion, dev_id: str) -> bytes:
    """
    Genera el acta de devolución en PDF usando ReportLab.
    Devuelve bytes del PDF listo para guardar en Drive o adjuntar al correo.
    """
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=14*mm, rightMargin=14*mm,
        topMargin=12*mm, bottomMargin=14*mm,
        title=f"Acta Devolución {dev_id}",
        author="SHARF Mesa de Servicio TI",
    )

    emp  = p.empleado or {}
    tec  = p.tecnico  or {}
    acc  = p.accesorios or {}
    accD = p.accesoriosDevueltos or {}
    accE = p.accesoriosEstado or {}
    accO = p.accesoriosObs or {}
    accC = p.accesoriosCosto or {}
    accCot = p.accesoriosCotizar or {}

    story = []

    # ── CABECERA ─────────────────────────────────────────────────────────
    logo_img = None
    if p.logoBase64:
        try:
            raw = base64.b64decode(p.logoBase64.split(",")[-1])
            logo_img = Image(io.BytesIO(raw), width=38*mm, height=16*mm)
        except Exception as e:
            log.warning(f"Logo: {e}")

    cab_logo = logo_img or Paragraph(
        '<font size="18"><b><i>sharf</i></b></font>', ParagraphStyle("", alignment=TA_LEFT))

    cab_centro = Paragraph(
        '<font size="12"><b>Devolución de Equipos y Materiales</b></font><br/>'
        '<font size="7" color="#555555"><i>'
        '(Documento imprescindible a efectos de la Liquidación de Beneficios Sociales)'
        '</i></font><br/>'
        f'<font size="6" color="#888888">T&amp;C-RG-03 · V.10 · {dev_id} · {_hoy()} · {_esc(tec.get("nombre",""))}</font>',
        ParagraphStyle("cab", alignment=TA_CENTER)
    )
    cab_cod = Paragraph(
        "T&C-RG-03<br/>V.10<br/>Pág.: 1 de 1",
        ParagraphStyle("cod", fontSize=7, alignment=TA_CENTER)
    )

    cab_table = Table([[cab_logo, cab_centro, cab_cod]],
                      colWidths=[42*mm, 105*mm, 25*mm])
    cab_table.setStyle(TableStyle([
        ("VALIGN",    (0,0), (-1,-1), "MIDDLE"),
        ("BOX",       (2,0), (2,0), 0.4, C_NEGRO),
        ("LEFTPADDING",  (0,0), (-1,-1), 3),
        ("RIGHTPADDING", (0,0), (-1,-1), 3),
    ]))
    story.append(cab_table)
    story.append(HRFlowable(width="100%", thickness=2, color=C_ROSA, spaceAfter=3))

    # ── DATOS COLABORADOR ─────────────────────────────────────────────────
    S = ParagraphStyle("s", fontSize=8, leading=11)
    SB = ParagraphStyle("sb", fontSize=8, leading=11, fontName="Helvetica-Bold")

    def cell(txt, bold=False):
        return Paragraph(_esc(txt), SB if bold else S)

    datos_tbl = Table([
        [cell("Nombres y Apellidos:", True), cell(emp.get("nombre","—")),
         cell("DNI / CI:", True),           cell(emp.get("dni","—"))],
        [cell("Área:", True),               cell(emp.get("area","—")),
         cell("Cargo:", True),              cell(emp.get("cargo","—"))],
        [cell("Sede:", True),               cell(emp.get("sede","—")),
         cell("Fecha:", True),              cell(_hoy())],
    ], colWidths=[32*mm, 70*mm, 20*mm, 50*mm])
    datos_tbl.setStyle(TableStyle([
        ("GRID",     (0,0), (-1,-1), 0.4, C_NEGRO),
        ("VALIGN",   (0,0), (-1,-1), "MIDDLE"),
        ("FONTSIZE", (0,0), (-1,-1), 8),
        ("TOPPADDING",    (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ("LEFTPADDING",   (0,0), (-1,-1), 4),
    ]))
    story.append(datos_tbl)

    # Empresa
    emp_actual = (emp.get("empresa","") or "").upper()
    empresas   = [
        ("SICCSA", "SICCSA - Scharff Int. Courier & Cargo"),
        ("SLI",    "SLI - Scharff Logística Integrada"),
        ("SR",     "SR - Scharff Representaciones"),
        ("SB",     "SB - Scharff Bolivia"),
    ]
    emp_html = "<b>Empresa donde laboró:</b><br/>"
    for cod, nombre_e in empresas:
        marca = "☑" if cod in emp_actual else "☐"
        negrita = "<b>" if cod in emp_actual else ""
        fin_neg = "</b>" if cod in emp_actual else ""
        emp_html += f"{marca} {negrita}{_esc(nombre_e)}{fin_neg}<br/>"
    emp_tbl = Table([[cell("Empresa:", True), Paragraph(emp_html, ParagraphStyle("e", fontSize=8, leading=11))]],
                    colWidths=[32*mm, 140*mm])
    emp_tbl.setStyle(TableStyle([
        ("GRID",    (0,0), (-1,-1), 0.4, C_NEGRO),
        ("VALIGN",  (0,0), (-1,-1), "TOP"),
        ("TOPPADDING",    (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ("LEFTPADDING",   (0,0), (-1,-1), 4),
    ]))
    story.append(emp_tbl)

    # ── TÍTULO SOPORTE TI ─────────────────────────────────────────────────
    story.append(Spacer(1, 2*mm))
    tit_tbl = Table([[Paragraph("<font color='white'><b>Soporte TI</b></font>",
                                ParagraphStyle("tit", fontSize=9, alignment=TA_CENTER))]],
                    colWidths=[172*mm])
    tit_tbl.setStyle(TableStyle([("BACKGROUND",(0,0),(0,0),C_ROSA),
                                  ("GRID",(0,0),(0,0),0.4,C_NEGRO),
                                  ("TOPPADDING",(0,0),(0,0),3),
                                  ("BOTTOMPADDING",(0,0),(0,0),3)]))
    story.append(tit_tbl)

    # ── TABLA DE ACCESORIOS ───────────────────────────────────────────────
    def cb(v: bool) -> str:
        return "☑" if v else "☐"

    th = ParagraphStyle("th", fontSize=6.5, fontName="Helvetica-Bold",
                         alignment=TA_CENTER, textColor=C_BLANCO)
    td = ParagraphStyle("td", fontSize=7.5, alignment=TA_CENTER)
    tl = ParagraphStyle("tl", fontSize=7.5, alignment=TA_LEFT)
    ob = ParagraphStyle("ob", fontSize=6.5, alignment=TA_LEFT)

    def p_th(t): return Paragraph(t, th)
    def p_td(t): return Paragraph(str(t), td)
    def p_tl(t): return Paragraph(_esc(str(t)), tl)
    def p_ob(t): return Paragraph(_esc(str(t)), ob)

    # Encabezados (2 filas por rowspan manual)
    hdr1 = [p_th("No."), p_th("Descripción"), p_th("Cant."),
            p_th("Recibido"), p_th(""), p_th(""),
            p_th("Estado"), p_th(""),
            p_th("Costo S/.\n(trabajador)"), p_th("Observaciones")]
    hdr2 = [p_th(""), p_th(""), p_th(""),
            p_th("SI"), p_th("NO"), p_th("NA"),
            p_th("Bueno"), p_th("Malo"),
            p_th(""), p_th("")]

    filas = [hdr1, hdr2]

    # Items
    items = [
        {"key":"activo",   "label": p.tipoActivo or "Laptop",
         "sn": f"S/N: {p.serial}", "devuelto": True,
         "bueno": p.equipoBueno,
         "obs": p.observacionDesc or ("Ver fotos adjuntas" if p.hayObservaciones else "Sin obs.")},
    ]
    acc_defs = [
        ("cargador", "Cargador de Laptop", acc.get("cargadorSerial","")),
        ("mouse",    "Mouse",              acc.get("mouseDesc","")),
        ("mochila",  "Mochila",            acc.get("mochilaDesc","")),
        ("docking",  "Docking Station",    acc.get("dockingDesc","")),
        ("teclado",  "Teclado",            acc.get("tecladoDesc","")),
    ]
    for key, label, desc in acc_defs:
        if desc or accD.get(key) or p.accesoriosNuevos.get(key):
            items.append({
                "key": key, "label": label,
                "sn": desc, "devuelto": bool(accD.get(key)),
                "bueno": (accE.get(key,"bueno") == "bueno"),
                "obs": accO.get(key,""),
            })

    n = 1
    for it in items:
        dev = it["devuelto"]
        bue = it["bueno"]
        cos = accC.get(it["key"], 0)
        cot = accCot.get(it["key"], False)
        obs_txt = (it["sn"] + ". " if it["sn"] else "") + it["obs"]
        if cot: obs_txt += " ⚠️ COTIZACIÓN"
        # Compromisos diferidos
        for cd in (p.compromisosDiferidos or []):
            if it["label"].split()[0].lower() in (cd.get("label","") or "").lower():
                if cd.get("fechaCompromiso"):
                    obs_txt += f" Pendiente: {cd['fechaCompromiso']}"
        costo_str = f"S/. {cos:.2f}" if cos and cos > 0 else "—"
        filas.append([
            p_td(n), p_tl(it["label"]), p_td("1"),
            p_td(cb(dev)), p_td(cb(not dev)), p_td("☐"),
            p_td(cb(dev and bue)), p_td(cb(dev and not bue)),
            p_td(costo_str), p_ob(obs_txt),
        ])
        n += 1

    extras = ["Backup del Equipo","¿Se canceló Cuenta de Correo?",
              "¿Se canceló Cuenta de Legacy?","¿Se canceló Cuenta de Sintad?",
              "¿Se canceló Cuenta de Dominio?"]
    for ex in extras:
        filas.append([p_td(n), p_tl(ex), p_td("—"),
                      p_td("☐"), p_td("☐"), p_td("—"),
                      p_td("—"), p_td("—"), p_td("—"), p_ob("")])
        n += 1

    # Total row
    total_costo = sum(accC.get(k, 0) for k in accC if accC.get(k, 0) > 0)
    filas.append([
        Paragraph(f"<b>(*) Total: S/. {total_costo:.2f}</b>",
                  ParagraphStyle("tot", fontSize=7.5, alignment=TA_RIGHT,
                                 textColor=C_ROSA)),
        p_td(""), p_td(""), p_td(""), p_td(""), p_td(""),
        p_td(""), p_td(""), p_td(""), p_td(""),
    ])

    COL_W = [8*mm, 40*mm, 9*mm, 9*mm, 9*mm, 9*mm, 11*mm, 11*mm, 18*mm, 48*mm]
    acc_tbl = Table(filas, colWidths=COL_W, repeatRows=2)

    HEADER_ROWS = [(0, i) for i in range(10)] + [(1, i) for i in range(10)]
    acc_tbl.setStyle(TableStyle([
        ("GRID",       (0,0), (-1,-1), 0.4, C_NEGRO),
        ("BACKGROUND", (0,0), (-1,1), C_ROSA),
        ("TEXTCOLOR",  (0,0), (-1,1), C_BLANCO),
        ("ALIGN",      (0,0), (-1,-1), "CENTER"),
        ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),
        ("FONTSIZE",   (0,0), (-1,-1), 7),
        ("TOPPADDING", (0,0), (-1,-1), 2),
        ("BOTTOMPADDING",(0,0),(-1,-1), 2),
        ("LEFTPADDING",(0,0),(-1,-1),  2),
        # Fila total
        ("SPAN",      (0,-1),(8,-1)),
        ("BACKGROUND",(9,-1),(9,-1),C_ROSA),
        # Rowspan manual: merge header cells
        ("SPAN",(0,0),(0,1)), ("SPAN",(1,0),(1,1)), ("SPAN",(2,0),(2,1)),
        ("SPAN",(3,0),(5,0)), ("SPAN",(6,0),(7,0)),
        ("SPAN",(8,0),(8,1)), ("SPAN",(9,0),(9,1)),
    ]))
    story.append(acc_tbl)

    # ── COMENTARIOS ───────────────────────────────────────────────────────
    story.append(Spacer(1, 2*mm))
    obs_gen = f"Equipo S/N: {p.serial}."
    if p.observacionDesc: obs_gen += f" {p.observacionDesc}."
    elif p.hayObservaciones: obs_gen += " Hallazgos registrados. Ver fotos adjuntas."
    else: obs_gen += " Sin observaciones adicionales."
    for cd in (p.compromisosDiferidos or []):
        obs_gen += f" {cd.get('label','')} pendiente" + (f" al {cd['fechaCompromiso']}" if cd.get("fechaCompromiso") else "") + "."
    obs_gen += f" Tipo devolución: {p.tipoDev}."

    obs_tbl = Table([[Paragraph(f"<b>Comentarios de Soporte TI:</b><br/>{_esc(obs_gen)}",
                                ParagraphStyle("c", fontSize=7.5, leading=11))]],
                    colWidths=[172*mm])
    obs_tbl.setStyle(TableStyle([
        ("BOX",(0,0),(0,0),0.4,C_NEGRO),
        ("TOPPADDING",(0,0),(0,0),3),
        ("BOTTOMPADDING",(0,0),(0,0),3),
        ("LEFTPADDING",(0,0),(0,0),4),
    ]))
    story.append(obs_tbl)

    # ── FIRMA TÉCNICO ─────────────────────────────────────────────────────
    story.append(Spacer(1, 4*mm))
    firma_content = []
    if p.firmaBase64:
        try:
            raw = base64.b64decode(p.firmaBase64.split(",")[-1])
            firma_img = Image(io.BytesIO(raw), width=35*mm, height=16*mm)
            firma_content.append(firma_img)
        except: pass
    firma_content.append(Paragraph(
        f"{'_'*35}<br/><b>{_esc(tec.get('nombre',''))}</b><br/>"
        f"<font size='6'>VB del Área de Soporte TI</font><br/>"
        f"<font size='6'>Fecha: {_hoy()}</font>",
        ParagraphStyle("f", fontSize=8, alignment=TA_CENTER, leading=11)))

    firma_tbl = Table([[Spacer(1,1), firma_content[0] if len(firma_content)>1 else firma_content[-1]]],
                      colWidths=[86*mm, 86*mm])
    firma_tbl2 = Table([firma_content[-1:]], colWidths=[86*mm])

    vb_tbl = Table(
        [["",
          Paragraph("<font color='white'><b>VB del Área de Soporte TI</b></font>",
                    ParagraphStyle("vb", fontSize=8, alignment=TA_CENTER))]],
        colWidths=[86*mm, 86*mm])
    vb_tbl.setStyle(TableStyle([
        ("BACKGROUND",(1,0),(1,0),C_ROSA),
        ("GRID",(1,0),(1,0),0.4,C_NEGRO),
        ("TOPPADDING",(0,0),(-1,-1),3),
        ("BOTTOMPADDING",(0,0),(-1,-1),3),
    ]))
    story.append(vb_tbl)

    f_cell = []
    if len(firma_content) > 1:
        f_cell.append(firma_content[0])
    f_cell.append(Paragraph(
        f"<b>{_esc(tec.get('nombre',''))}</b><br/>"
        f"<font size='6'>VB del Área de Soporte TI · Fecha: {_hoy()}</font>",
        ParagraphStyle("fn", fontSize=8, alignment=TA_CENTER, leading=11)))
    firma_inner = Table([f_cell], colWidths=[86*mm] * len(f_cell))
    firma_inner.setStyle(TableStyle([
        ("ALIGN",(0,0),(-1,-1),"CENTER"),("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("GRID",(0,0),(-1,-1),0.4,C_NEGRO),
        ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),6),
    ]))
    firma_row = Table([["", firma_inner]], colWidths=[86*mm, 86*mm])
    firma_row.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP")]))
    story.append(firma_row)

    # ── PIE DE PÁGINA ─────────────────────────────────────────────────────
    story.append(Spacer(1, 3*mm))
    story.append(HRFlowable(width="100%", thickness=0.5, color=C_GRIS2))
    story.append(Paragraph(
        "<font size='6' color='#666666'><i>"
        "De existir observaciones con lo indicado en el presente documento, sírvase proporcionar "
        "el formato de descuento correspondiente debidamente firmado para hacerse responsable por "
        "los equipos y/o materiales no entregados, y/o los devueltos en mal estado."
        "</i></font>",
        ParagraphStyle("pie", fontSize=6, alignment=TA_LEFT, leading=9)
    ))

    # ── Fotos (página 2 si hay) ───────────────────────────────────────────
    if p.fotosBase64:
        from reportlab.platypus import PageBreak
        story.append(PageBreak())
        story.append(Paragraph("<b>Evidencia fotográfica</b>",
                               ParagraphStyle("fp", fontSize=11, spaceAfter=6)))
        foto_rows = []
        row_actual = []
        for i, b64 in enumerate(p.fotosBase64[:9]):
            try:
                raw  = base64.b64decode(b64.split(",")[-1])
                fimg = Image(io.BytesIO(raw), width=55*mm, height=42*mm)
                row_actual.append(fimg)
            except:
                row_actual.append(Paragraph("Foto no disponible", S))
            if len(row_actual) == 3:
                foto_rows.append(row_actual)
                row_actual = []
        if row_actual:
            while len(row_actual) < 3:
                row_actual.append("")
            foto_rows.append(row_actual)
        if foto_rows:
            foto_tbl = Table(foto_rows, colWidths=[58*mm]*3)
            foto_tbl.setStyle(TableStyle([
                ("GRID",(0,0),(-1,-1),0.4,C_GRIS2),
                ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                ("ALIGN",(0,0),(-1,-1),"CENTER"),
                ("TOPPADDING",(0,0),(-1,-1),3),
                ("BOTTOMPADDING",(0,0),(-1,-1),3),
            ]))
            story.append(foto_tbl)

    doc.build(story)
    pdf_bytes = buf.getvalue()
    log.info(f"PDF generado: {len(pdf_bytes)//1024} KB")
    return pdf_bytes
