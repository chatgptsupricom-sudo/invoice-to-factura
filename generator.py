from io import BytesIO
from datetime import date
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


TEMPLATE_PATH = "template.docx"

TAB_DESC   = 3.5
TAB_CANT   = 13.2
TAB_PRECIO = 15.5
TAB_TOTAL  = 18.5

FOOTER_LEGAL = (
    "FACTURA SEGUN CONVENIO CAMBIARIO N°1 GOE N°6.405 DE FECHA 07/09/2018. "
    "FACTURA CALCULADA A LA TASA BCV ACTUAL Bs. POR US$ DOLAR "
    "{tasa}PARA EFECTOS DEL IVA DE CONFORMIDAD CON LO ESTABLECIDO EN EL ARTICULO 116 DEL BCV. "
    "PAGARA EN BOLIVARES A LA TASA BCV DEL DIA."
)

IGTF_LEGAL = (
    "En caso de pagar la totalidad de la factura o porción con moneda extranjera, "
    "se cobrará el 3% de IGTF según Gaceta Oficial 42.339 de fecha 2/02/2022"
)


def _add_tab_stops(para, stops):
    pPr = para._p.get_or_add_pPr()
    tabs_el = OxmlElement("w:tabs")
    for pos_cm, align in stops:
        tab = OxmlElement("w:tab")
        tab.set(qn("w:val"), align)
        tab.set(qn("w:pos"), str(int(Cm(pos_cm).pt * 20)))
        tabs_el.append(tab)
    pPr.append(tabs_el)


def _remove_borders(table):
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


def _set_borders(table):
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    tblBorders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        tblBorders.append(el)
    tblPr.append(tblBorders)


def _set_cell_width(cell, twips):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcW = OxmlElement("w:tcW")
    tcW.set(qn("w:w"), str(twips))
    tcW.set(qn("w:type"), "dxa")
    tcPr.append(tcW)


def _cell_para(cell, text, bold=False, size=8, align=WD_ALIGN_PARAGRAPH.LEFT):
    p = cell.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    p.alignment = align
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.bold = bold
    return p


def _header_table(doc, invoice_no, invoice_date, due_date, terms):
    table = doc.add_table(rows=1, cols=2)
    _remove_borders(table)

    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"), str(int(Cm(19).pt * 20)))
    tblW.set(qn("w:type"), "dxa")
    tblPr.append(tblW)

    left  = table.cell(0, 0)
    right = table.cell(0, 1)

    _set_cell_width(left,  9000)
    _set_cell_width(right, 5400)

    left.text = ""
    _cell_para(left, "Cliente: SUPRICOM CCS 21, C.A.", bold=True, size=9)
    _cell_para(left, "CALLE LOS LABORATORIOS EDIF. OFINCA PISO PB LOCAL 2-A, LOS RUISES, CARACAS, MIRANDA", size=8)
    _cell_para(left, "Distrito Capital DTC Distrito Capital", size=8)
    _cell_para(left, "Venezuela — J501193738", size=8)

    right.text = ""
    _cell_para(right, f"Número de Factura: {invoice_no}", bold=True, size=9, align=WD_ALIGN_PARAGRAPH.RIGHT)
    _cell_para(right, f"Fecha de Emisión: {invoice_date}", size=8, align=WD_ALIGN_PARAGRAPH.RIGHT)
    _cell_para(right, f"Termino de Pago: {terms}", size=8, align=WD_ALIGN_PARAGRAPH.RIGHT)


def _row_para(doc, cols, bold=False, font_size=5.5):
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after  = Pt(0.5)

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


def _divider(doc, size=7):
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after  = Pt(0)
    run = para.add_run("─" * 120)
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor(0xAA, 0xAA, 0xAA)


def _totals_box(doc, subtotal, iva_rate, tasa_bcv):
    """Bordered two-column table with USD and Bs. totals."""
    iva_pct    = int(iva_rate * 100)
    iva_amount = round(subtotal * iva_rate, 2)
    total_usd  = round(subtotal + iva_amount, 2)

    sub_bs     = round(subtotal   * tasa_bcv, 2)
    iva_bs     = round(iva_amount * tasa_bcv, 2)
    total_bs   = round(total_usd  * tasa_bcv, 2)

    table = doc.add_table(rows=5, cols=2)
    _set_borders(table)

    # Column widths
    for row in table.rows:
        _set_cell_width(row.cells[0], 7100)
        _set_cell_width(row.cells[1], 7200)

    # Header row (row 0): column titles
    h_left  = table.cell(0, 0)
    h_right = table.cell(0, 1)
    h_left.text  = ""
    h_right.text = ""
    _cell_para(h_left,  "Monto en Dólares  (U$S.)", bold=True, size=6)
    _cell_para(h_right, "Monto en Bolívares  (Bs.)", bold=True, size=6)

    rows_data = [
        # label_usd,                 val_usd,     label_bs,                          val_bs
        ("Sub-Total:",               subtotal,    "Sub-Total Ref:",                  sub_bs),
        ("Base Imponible:",          subtotal,    "Base Imponible Ref:",              sub_bs),
        (f"IVA {iva_pct}% Sobre {_fmt(subtotal - iva_amount / iva_rate * (1 - iva_rate)):.3f}:",
                                     iva_amount,
         f"IVA {iva_pct}% Ref Sobre {_fmt(sub_bs / (1 + iva_rate)):.3f}  :",
                                     iva_bs),
        ("IGTF 3% Sugerido:",        0.00,        "IGTF 3% Sugerido Ref:",           0.00),
        ("Total General U$S:",       total_usd,   "Total General Ref Bs.:",          total_bs),
    ]

    for i, (lu, vu, lb, vb) in enumerate(rows_data):
        r = table.rows[i]
        r.cells[0].text = ""
        r.cells[1].text = ""
        is_total = (i == 4)
        # Left cell: label + value side by side
        _biline(r.cells[0], lu, _fmt(vu), bold=is_total, size=6)
        _biline(r.cells[1], lb, _fmt(vb), bold=is_total, size=6)

    return table


def _biline(cell, label, value, bold=False, size=6):
    """One paragraph: label left-aligned, value right-aligned via tab."""
    p = cell.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)

    # Tab stop at far right of cell (~8.5 cm)
    pPr = p._p.get_or_add_pPr()
    tabs_el = OxmlElement("w:tabs")
    tab = OxmlElement("w:tab")
    tab.set(qn("w:val"), "right")
    tab.set(qn("w:pos"), str(int(Cm(8.8).pt * 20)))
    tabs_el.append(tab)
    pPr.append(tabs_el)

    r = p.add_run(f"{label}\t{value}")
    r.font.size = Pt(size)
    r.bold = bold


def _small_para(doc, text, size=5.5, bold=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(1)
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.bold = bold


def generate_docx(invoice_data: dict, iva_rate: float = 0.16, tasa_bcv: float = 1.0) -> bytes:
    hdr   = invoice_data["header"]
    items = invoice_data["items"]

    doc = Document()

    for section in doc.sections:
        section.top_margin    = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin   = Cm(2)
        section.right_margin  = Cm(1.5)

    # Header
    _header_table(
        doc,
        invoice_no   = hdr.get("invoice_no", ""),
        invoice_date = hdr.get("invoice_date", ""),
        due_date     = hdr.get("due_date", ""),
        terms        = hdr.get("terms", ""),
    )
    _space(doc, 4)
    _divider(doc)
    _space(doc, 1)

    # Column headers
    _row_para(doc, ["CÓDIGO", "DESCRIPCIÓN", "CANT.", "PRECIO", "TOTAL"],
              bold=True, font_size=6)
    _divider(doc)

    # Items (same order as PDF)
    for item in items:
        _row_para(doc, [
            item["codigo"],
            item["descripcion"],
            _fmt_qty(item["cantidad"]),
            _fmt_money(item["precio"]),
            _fmt_money(item["total"]),
        ], font_size=5.5)

    _divider(doc)
    _space(doc, 4)

    # Totals box
    subtotal = sum(i["total"] for i in items)
    _totals_box(doc, subtotal, iva_rate, tasa_bcv)

    _space(doc, 2)

    # Legal texts
    _small_para(doc, IGTF_LEGAL, size=5.5)
    today = date.today().strftime("%-d/%m/%Y")
    tasa_fmt = f"{tasa_bcv:,.4f}".replace(",", "X").replace(".", ",").replace("X", ".")
    _small_para(doc, f"Tasa Vigente al {today} según el BCV: {tasa_fmt}", size=5.5)
    _space(doc, 2)
    _divider(doc)
    _space(doc, 1)
    _small_para(doc, FOOTER_LEGAL.format(tasa=tasa_fmt), size=5.5)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ── Helpers ──────────────────────────────────────────────────────────────────

def _space(doc, after_pt=4):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(after_pt)


def _fmt(val) -> float:
    return val


def _fmt_money(val) -> str:
    return f"{val:,.2f}"


def _fmt_qty(val) -> str:
    if float(val) == int(float(val)):
        return str(int(float(val)))
    return str(val)
