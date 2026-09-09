"""Generación de PDF de una cotización usando reportlab.

Estética alineada a la identidad de Skybridge ya usada en la app (navy +
naranja, tipografía sans-serif limpia, líneas finas en vez de grillas
pesadas): ver los tokens de color en _inyectar_estilos() de app.py.
"""
from io import BytesIO
from datetime import datetime

from reportlab.lib import colors
from reportlab.lib.enums import TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, HRFlowable, PageBreak
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# Misma paleta que --sb-navy / --sb-orange / etc. en app.py.
NAVY = colors.HexColor("#0F172A")
ORANGE = colors.HexColor("#E8652A")
ORANGE_DARK = colors.HexColor("#C04E18")
TEXT_SECONDARY = colors.HexColor("#64748B")
BORDER = colors.HexColor("#E2E8F0")
ZONE = colors.HexColor("#F1F5F9")


def _money(v, symbol="USD"):
    try:
        return f"{symbol} {v:,.2f}"
    except (TypeError, ValueError):
        return f"{symbol} 0.00"


def _fmt_fecha(valor):
    if not valor:
        return "-"
    try:
        return datetime.strptime(str(valor)[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        return str(valor)


def generar_pdf_cotizacion(cab: dict, resultado: dict, completo: bool = True) -> bytes:
    """completo=False arma solo la página 1 (Resumen + Detalle de
    mercadería), pensada para mandarle al cliente sin exponer el desglose
    línea por línea de gastos. completo=True agrega la página 2 con ese
    desglose completo."""
    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=18 * mm, rightMargin=18 * mm, topMargin=16 * mm, bottomMargin=16 * mm,
    )
    styles = getSampleStyleSheet()
    normal = ParagraphStyle("NormalSB", parent=styles["Normal"], fontName="Helvetica", fontSize=8.5, textColor=NAVY, leading=12)
    meta_right = ParagraphStyle("MetaRight", parent=normal, alignment=TA_RIGHT)
    # leftIndent=-6: reportlab arma el Frame de SimpleDocTemplate con 6pt de
    # padding interno por defecto, que se le aplica a los Paragraph sueltos
    # del story pero NO a las tablas (donde ya llevamos ese padding a 0 en
    # la columna 0). Sin este ajuste los títulos y el pie quedan ~6pt más a
    # la derecha que el resto — se compensa acá para que todo arranque en
    # el mismo margen real de la página.
    section_h = ParagraphStyle(
        "SectionH", parent=styles["Normal"], fontName="Helvetica-Bold",
        fontSize=11, textColor=ORANGE_DARK, spaceBefore=14, spaceAfter=4, leftIndent=-6,
    )
    label_style = ParagraphStyle("Label", parent=normal, fontName="Helvetica", fontSize=7.5, textColor=TEXT_SECONDARY, leading=10)
    value_style = ParagraphStyle("Value", parent=normal, fontName="Helvetica-Bold", fontSize=9, textColor=NAVY, leading=12)
    footer_style = ParagraphStyle("Footer", parent=normal, fontSize=7.5, textColor=TEXT_SECONDARY, leftIndent=-6)

    content_width = doc.width
    story = []

    # ---- Encabezado: wordmark + subtítulo a la izquierda, número/fecha a
    # la derecha, con una línea fina debajo (mismo patrón que los títulos
    # de sección en la app: color + border-bottom, sin bloques de color). ----
    wordmark = Paragraph(
        '<font color="#0F172A"><b>SKY</b></font><font color="#E8652A"><b>BRIDGE</b></font>'
        '<br/><font color="#64748B" size="7.5">PRESUPUESTO DE IMPORTACIÓN</font>',
        normal,
    )
    meta = Paragraph(
        f'<font color="#0F172A"><b>Cotización {cab.get("numero") or "-"}</b></font>'
        f'<br/><font color="#64748B" size="8">{datetime.now().strftime("%d/%m/%Y")}</font>',
        meta_right,
    )
    header_tbl = Table([[wordmark, meta]], colWidths=[content_width * 0.6, content_width * 0.4])
    header_tbl.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    story.append(header_tbl)
    story.append(Spacer(1, 6))
    story.append(HRFlowable(width="100%", thickness=1, color=BORDER, spaceAfter=10))

    # ---- Datos del cliente / operación: pares etiqueta-valor sin grilla,
    # mismo lenguaje visual que .sb-field-label / .sb-field-value. ----
    def _campo(etiqueta, valor):
        return Paragraph(f"{etiqueta}<br/><font color=\"#0F172A\"><b>{valor or '-'}</b></font>", label_style)

    info_rows = [
        [_campo("CLIENTE", cab.get("cliente_nombre")), _campo("CUIT", cab.get("cliente_cuit"))],
        [_campo("DETALLE PEDIDO", cab.get("detalle_pedido")), _campo("CONDICIÓN DE VENTA", cab.get("contenedor"))],
        [_campo("ORIGEN", cab.get("origen_cond_venta")), _campo("ETD", _fmt_fecha(cab.get("carrier")))],
        [_campo("TIPO DE ENVÍO", cab.get("etd_eta")), _campo("ETA", _fmt_fecha(cab.get("freetime")))],
    ]
    info_tbl = Table(info_rows, colWidths=[content_width * 0.5, content_width * 0.5])
    info_tbl.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 12),
        ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(info_tbl)

    # ---- Resumen de la operación ----
    story.append(Paragraph("Resumen de la operación", section_h))
    resumen_data = [
        ["Concepto", "USD", "ARS"],
        ["① Mercadería + costo financiero", _money(resultado["fob_total_sum"] + resultado["costo_financiero"]),
         _money((resultado["fob_total_sum"] + resultado["costo_financiero"]) * resultado["cab"]["tc_oper"], "ARS")],
        ["② Tributos VEP (AFIP)", _money(resultado["total_tributos_vep_usd"]), _money(resultado["total_tributos_vep_ars"], "ARS")],
        ["③ Gastos operativos + IVA", _money(resultado["subtotal_oper_usd"] + resultado["subtotal_ivater_usd"]),
         _money(resultado["subtotal_oper_ars"] + resultado["subtotal_ivater_ars"], "ARS")],
        ["TOTAL desembolsado (c/IVA)", _money(resultado["total_c_iva_usd"]), _money(resultado["total_c_iva_ars"], "ARS")],
        ["Incidencia s/FOB", f"{resultado['incidencia_fob']*100:.1f}%", ""],
    ]
    t2 = Table(resumen_data, colWidths=[content_width * 0.5, content_width * 0.25, content_width * 0.25])
    t2.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME", (0, -2), (-1, -2), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("TEXTCOLOR", (0, 0), (-1, 0), TEXT_SECONDARY),
        ("TEXTCOLOR", (0, 1), (-1, -1), NAVY),
        ("LINEBELOW", (0, 0), (-1, 0), 0.75, BORDER),
        ("LINEBELOW", (0, 1), (-1, -2), 0.4, BORDER),
        ("BACKGROUND", (0, -2), (-1, -2), ZONE),
        ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
        ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 6), ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        # La col. 0 y la última alinean con el margen real de la página (0),
        # igual que el wordmark y los datos del cliente arriba — si no, el
        # contenido de la tabla queda corrido respecto del título.
        ("LEFTPADDING", (0, 0), (0, -1), 0), ("RIGHTPADDING", (-1, 0), (-1, -1), 0),
    ]))
    story.append(t2)

    # ---- Detalle de mercadería ----
    story.append(Paragraph("Detalle de mercadería", section_h))
    # Encabezados largos ("Costo c/IVA U. USD"/"ARS") como Paragraph en vez
    # de texto plano: así ajustan a 2 líneas dentro de su columna en lugar
    # de invadir la columna vecina (con texto plano reportlab no corta la
    # celda del header, solo la del contenido).
    head_style = ParagraphStyle(
        "ColHead", parent=styles["Normal"], fontName="Helvetica-Bold",
        fontSize=7.5, textColor=TEXT_SECONDARY, leading=9,
    )
    head_style_r = ParagraphStyle("ColHeadR", parent=head_style, alignment=TA_RIGHT)

    def _head(texto, alinear_derecha=True):
        return Paragraph(texto, head_style_r if alinear_derecha else head_style)

    # Descripción/NCM como Paragraph (no texto plano): con texto plano
    # reportlab no corta la celda cuando el contenido no entra en el ancho
    # de columna, así que un nombre de producto largo quedaba superpuesto
    # con las columnas vecinas (bug reportado por Tom en la cotización
    # 0013). Con Paragraph el texto ajusta a 2+ líneas dentro de su propia
    # columna y la fila simplemente crece de alto — se ve el nombre
    # completo, sin invadir FOB Unit./Cant./etc.
    cell_style = ParagraphStyle(
        "Cell", parent=styles["Normal"], fontName="Helvetica",
        fontSize=8, textColor=NAVY, leading=10,
    )

    def _celda(texto):
        return Paragraph(str(texto), cell_style)

    prod_rows = [[
        _head("Descripción", False), _head("NCM", False), _head("FOB Unit."), _head("Cant."),
        _head("FOB Total"), _head("Costo c/IVA U. USD"), _head("Costo c/IVA U. ARS"),
    ]]
    for p in resultado["productos"]:
        if not p.get("descripcion") and p["cantidad"] == 0:
            continue
        prod_rows.append([
            # Solo Descripción va en Paragraph: es la que puede ser larga y
            # se superponía. El NCM es un código corto de formato fijo
            # (nunca fue el problema reportado) — se deja como texto plano
            # para que no se corte en un punto intermedio del código.
            _celda(p.get("descripcion") or "-"), p.get("ncm") or "-",
            f"{p['fob_unit']:,.2f}", f"{p['cantidad']:,.2f}", f"{p['fob_total']:,.2f}",
            f"{p['costo_civa_unit']:,.2f}", f"{p['costo_civa_unit_ars']:,.2f}",
        ])
    t3 = Table(prod_rows, colWidths=[
        content_width * 0.20, content_width * 0.09, content_width * 0.11, content_width * 0.08,
        content_width * 0.14, content_width * 0.19, content_width * 0.19,
    ])
    t3.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 1), (-1, -1), 8),
        ("VALIGN", (0, 0), (-1, 0), "BOTTOM"),
        # Top, no Middle: con Descripción/NCM ahora en Paragraph (ajustan a
        # varias líneas si el nombre es largo), filas de distinto alto se
        # ven mejor con todo arrancando arriba en vez de centrado.
        ("VALIGN", (0, 1), (-1, -1), "TOP"),
        ("TEXTCOLOR", (0, 0), (-1, 0), TEXT_SECONDARY),
        ("TEXTCOLOR", (0, 1), (-1, -1), NAVY),
        ("LINEBELOW", (0, 0), (-1, 0), 0.75, BORDER),
        ("LINEBELOW", (0, 1), (-1, -1), 0.4, BORDER),
        ("ALIGN", (2, 0), (-1, -1), "RIGHT"),
        ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 6), ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (0, -1), 0), ("RIGHTPADDING", (-1, 0), (-1, -1), 0),
    ]))
    story.append(t3)

    # ---- Desglose de TODOS los componentes del costo (mercadería, costo
    # financiero, tributos aduaneros y gastos operativos) — cada uno con su
    # monto neto (s/IVA), el IVA que se le suma y el total final, todo en
    # USD. Los tributos (Derechos, Tasa, Antidumping, IVA, IVA Adicional,
    # Ganancias, IIBB, Arancel SIM) no llevan un IVA propio encima — son en
    # sí mismos cargos de tipo impositivo — así que ahí la col. IVA va en 0
    # y Final = s/IVA. El total de esta tabla reconcilia exacto con el
    # "TOTAL desembolsado (c/IVA)" del Resumen de arriba.
    #
    # Solo en la versión completa: expone la estructura interna de costos
    # línea por línea, así que la versión simplificada (pensada para el
    # cliente) la omite por completo y termina en "Detalle de mercadería".
    #
    # Salto de página deliberado: con ~14+ filas, el desglose completo no
    # entra debajo del resto sin partirse a la mitad ni dejar el pie de
    # página huérfano solo en la 2da hoja — arranca siempre limpio en su
    # propia página, con el mini-header de _draw_pagina_2 arriba. ----
    if completo:
        story.append(PageBreak())
        story.append(Paragraph("Desglose de gastos", section_h))
        gastos_rows = [["Concepto", "s/IVA (USD)", "IVA (USD)", "Final (USD)"]]
        tot_sviva = tot_iva = tot_final = 0.0

        def _fila(concepto, monto, iva=0.0):
            # Los totales suman siempre, pero un concepto en USD 0 (costo no
            # cargado o no aplicable a esta operación) no se lista — así el
            # desglose impreso solo muestra los ítems que efectivamente
            # tienen costo, sin ensuciar la hoja con filas en cero.
            nonlocal tot_sviva, tot_iva, tot_final
            final = monto + iva
            tot_sviva += monto
            tot_iva += iva
            tot_final += final
            if round(final, 2) == 0:
                return
            gastos_rows.append([concepto, f"{monto:,.2f}", f"{iva:,.2f}", f"{final:,.2f}"])

        _fila("Mercadería (FOB)", resultado["fob_total_sum"])
        _fila("Costo financiero", resultado["costo_financiero"])
        _fila("Derechos de importación", resultado["subtotal_der"])
        _fila("Tasa estadística", resultado["subtotal_tasa"])
        _fila("Antidumping", resultado["subtotal_antid"])
        _fila("IVA (percepción)", resultado["subtotal_iva"])
        _fila("IVA Adicional", resultado["subtotal_iva_ad"])
        _fila("Ganancias (percepción)", resultado["subtotal_gcias"])
        _fila("IIBB (percepción)", resultado["subtotal_iibb"])
        _fila("Arancel SIM", resultado["cab"]["arancel_sim"])
        for g in resultado["gastos"]:
            concepto = g.get("concepto") or "-"
            if concepto == "Honorarios consultora":
                concepto = "Consultoría integral"
            _fila(concepto, g["usd"], g["iva_terceros_usd"])

        gastos_rows.append(["TOTAL", f"{tot_sviva:,.2f}", f"{tot_iva:,.2f}", f"{tot_final:,.2f}"])
        t4 = Table(gastos_rows, colWidths=[content_width * 0.4, content_width * 0.2, content_width * 0.2, content_width * 0.2])
        t4.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("TEXTCOLOR", (0, 0), (-1, 0), TEXT_SECONDARY),
            ("TEXTCOLOR", (0, 1), (-1, -1), NAVY),
            ("LINEBELOW", (0, 0), (-1, 0), 0.75, BORDER),
            ("LINEBELOW", (0, 1), (-1, -1), 0.4, BORDER),
            ("BACKGROUND", (0, -1), (-1, -1), ZONE),
            ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
            ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING", (0, 0), (-1, -1), 6), ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ("LEFTPADDING", (0, 0), (0, -1), 0), ("RIGHTPADDING", (-1, 0), (-1, -1), 0),
        ]))
        story.append(t4)

    story.append(Spacer(1, 14))
    story.append(HRFlowable(width="100%", thickness=0.5, color=BORDER, spaceAfter=6))
    story.append(Paragraph("Valores provisorios sujetos a confirmación.", footer_style))

    # ---- Numeración de página, en ambas hojas: con 2 páginas ya no es obvio
    # a simple vista que el documento sigue en la siguiente, así que un
    # "Página N" abajo a la derecha cierra esa duda sin depender del total
    # de páginas (que puede variar según cuántos productos tenga la
    # cotización). ----
    def _draw_num_pagina(canvas_obj):
        canvas_obj.setFont("Helvetica", 7.5)
        canvas_obj.setFillColor(TEXT_SECONDARY)
        canvas_obj.drawRightString(A4[0] - 18 * mm, 10 * mm, f"Página {canvas_obj.getPageNumber()}")

    def _draw_primera_pagina(canvas_obj, _doc):
        canvas_obj.saveState()
        _draw_num_pagina(canvas_obj)
        canvas_obj.restoreState()

    if completo:
        # ---- Mini-header de continuidad para la página 2: como el desglose
        # arranca en su propia hoja (PageBreak arriba), sin esto se vería
        # como una página "suelta" sin identidad — repite wordmark + nº de
        # cotización en una franja liviana dentro del margen superior
        # existente. Solo aplica a la versión completa, que es la única con
        # más de una página. ----
        def _draw_pagina_2(canvas_obj, _doc):
            canvas_obj.saveState()
            canvas_obj.setFont("Helvetica-Bold", 9)
            canvas_obj.setFillColor(NAVY)
            canvas_obj.drawString(18 * mm, A4[1] - 11 * mm, "SKYBRIDGE")
            canvas_obj.setFont("Helvetica", 8)
            canvas_obj.setFillColor(TEXT_SECONDARY)
            canvas_obj.drawRightString(
                A4[0] - 18 * mm, A4[1] - 11 * mm, f"Cotización {cab.get('numero') or '-'}"
            )
            canvas_obj.setStrokeColor(BORDER)
            canvas_obj.setLineWidth(0.75)
            canvas_obj.line(18 * mm, A4[1] - 13.5 * mm, A4[0] - 18 * mm, A4[1] - 13.5 * mm)
            _draw_num_pagina(canvas_obj)
            canvas_obj.restoreState()

        doc.build(story, onFirstPage=_draw_primera_pagina, onLaterPages=_draw_pagina_2)
    else:
        doc.build(story, onFirstPage=_draw_primera_pagina)
    return buf.getvalue()
