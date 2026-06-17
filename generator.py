from io import BytesIO
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


TEMPLATE_PATH = "template.docx"

# Tab stop positions in cm from left margin
TAB_DESC   = 3.5
TAB_CANT   = 13.2
TAB_PRECIO = 15.5
TAB_TOTAL  = 18.5


def _add_tab_stops(para, stops):
    pPr = para._p.get_or_add_pPr()
    tabs_el = OxmlElement("w:tabs")
    for pos_cm, align in stops:
        tab = OxmlElement("w:tab")
        tab.set(qn("w:val"), align)
        tab.set(qn("w:pos"), str(int(Cm(pos_cm).pt * 20)))
        tabs_el.append(tab)
    pPr.append(tabs_el)


def _remove_table_borders(table):
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    tblBorders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "none")
        tblBorders.append(el)
    tblPr.append(tblBorders)


def _header_table(doc, invoice_no, invoice_date, due_date, terms):
    """Two-column borderless table: client info left, invoice details right."""
    table = doc.add_table(rows=1, cols=2)
    _remove_table_borders(table)

    # Set column widths
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"), str(int(Cm(19).pt * 20)))
    tblW.set(qn("w:type"), "dxa")
    tblPr.append(tblW)

    left_cell  = table.cell(0, 0)
    right_cell = table.cell(0, 1)

    # Fix widths
    for cell, w in ((left_cell, 9000), (right_cell, 5400)):
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        tcW = OxmlElement("w:tcW")
        tcW.set(qn("w:w"), str(w))
        tcW.set(qn("w:type"), "dxa")
        tcPr.append(tcW)

    # Left: client info
    left_cell.text = ""
    _cell_line(left_cell, "Cliente: SUPRICOM CCS 21, C.A.", bold=True, size=10)
    _cell_line(left_cell, "CALLE LOS LABORATORIOS EDIF. OFINCA PISO PB LOCAL 2-A, LOS RUISES, CARACAS, MIRANDA", size=9)
    _cell_line(left_cell, "Distrito Capital DTC Distrito Capital", size=9)
    _cell_line(left_cell, "Venezuela — J501193738", size=9)

    # Right: invoice details (right-aligned)
    right_cell.text = ""
    _cell_line(right_cell, f"Número de Factura: {invoice_no}", bold=True, size=10, align=WD_ALIGN_PARAGRAPH.RIGHT)
    _cell_line(right_cell, f"Fecha de Emisión: {invoice_date}", size=9, align=WD_ALIGN_PARAGRAPH.RIGHT)
    _cell_line(right_cell, f"Termino de Pago: {terms}", size=9, align=WD_ALIGN_PARAGRAPH.RIGHT)

    return table


def _cell_line(cell, text, bold=False, size=9, align=WD_ALIGN_PARAGRAPH.LEFT):
    p = cell.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(1)
    p.alignment = align
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.bold = bold


def _row_para(doc, cols, bold=False, font_size=9):
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after  = Pt(1)

    stops = [
        (TAB_DESC,   "left"),
        (TAB_CANT,   "right"),
        (TAB_PRECIO, "right"),
        (TAB_TOTAL,  "right"),
    ]
    _add_tab_stops(para, stops)

    text = f"{cols[0]}\t{cols[1]}\t{cols[2]}\t{cols[3]}\t{cols[4]}"
    run = para.add_run(text)
    run.font.size = Pt(font_size)
    run.bold = bold
    return para


def _divider(doc):
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after  = Pt(0)
    run = para.add_run("─" * 110)
    run.font.size = Pt(7)
    run.font.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)


def generate_docx(invoice_data: dict, iva_rate: float = 0.16) -> bytes:
    hdr   = invoice_data["header"]
    items = invoice_data["items"]

    doc = Document()

    for section in doc.sections:
        section.top_margin    = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin   = Cm(2)
        section.right_margin  = Cm(1.5)

    # --- Header: two-column layout ---
    _header_table(
        doc,
        invoice_no   = hdr.get("invoice_no", ""),
        invoice_date = hdr.get("invoice_date", ""),
        due_date     = hdr.get("due_date", ""),
        terms        = hdr.get("terms", ""),
    )
    _space(doc)
    _divider(doc)
    _space(doc)

    # --- Column headers (no fill, bold only) ---
    _row_para(doc,
              ["CÓDIGO", "DESCRIPCIÓN", "CANT.", "PRECIO", "TOTAL"],
              bold=True, font_size=9)
    _divider(doc)

    # --- Item rows ---
    for item in items:
        _row_para(doc, [
            item["codigo"],
            item["descripcion"],
            _fmt_qty(item["cantidad"]),
            _fmt_money(item["precio"]),
            _fmt_money(item["total"]),
        ], font_size=8.5)

    _divider(doc)
    _space(doc)

    # --- Totals ---
    subtotal      = sum(i["total"] for i in items)
    iva_amount    = round(subtotal * iva_rate, 2)
    total_general = round(subtotal + iva_amount, 2)

    _total_row(doc, "Subtotal:",                    subtotal)
    _total_row(doc, f"IVA ({int(iva_rate * 100)})%:", iva_amount)
    _total_row(doc, "TOTAL GENERAL:",                total_general, bold=True, size=11)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ── Helpers ──────────────────────────────────────────────────────────────────

def _space(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(4)


def _total_row(doc, label, amount, bold=False, size=10):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(2)
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run(f"{label}    {_fmt_money(amount)}")
    r.font.size = Pt(size)
    r.bold = bold


def _fmt_money(val) -> str:
    return f"{val:,.2f}"


def _fmt_qty(val) -> str:
    if float(val) == int(float(val)):
        return str(int(float(val)))
    return str(val)
