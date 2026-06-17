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
    "{tasa}PARA EFECTOS DEL IVA DE CONFORMIDAD CON LO ESTABLECIDO "
    "EN EL ARTICULO 116 DEL BCV. PAGARA EN BOLIVARES A LA TASA BCV DEL DIA."
)

IGTF_LEGAL = (
    "En caso de pagar la totalidad de la factura o porción con moneda extranjera, "
    "se cobrará el 3% de IGTF según Gaceta Oficial 42.339 de fecha 2/02/2022"
)

_BORDER_SINGLE = {"val": "single", "sz": "4", "space": "0", "color": "000000"}
_BORDER_NONE   = {"val": "none"}


# ── XML helpers ───────────────────────────────────────────────────────────────

def _add_tab_stops(para, stops):
    pPr = para._p.get_or_add_pPr()
    tabs_el = OxmlElement("w:tabs")
    for pos_cm, align in stops:
        tab = OxmlElement("w:tab")
        tab.set(qn("w:val"), align)
        tab.set(qn("w:pos"), str(int(Cm(pos_cm).pt * 20)))
        tabs_el.append(tab)
    pPr.append(tabs_el)


def _apply_tbl_borders(table, sides):
    """sides: dict of side_name -> border_attrs_dict"""
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    # Remove existing tblBorders
    old = tblPr.find(qn("w:tblBorders"))
    if old is not None:
        tblPr.remove(old)
    tblBorders = OxmlElement("w:tblBorders")
    for side, attrs in sides.items():
        el = OxmlElement(f"w:{side}")
        for k, v in attrs.items():
            el.set(qn(f"w:{k}"), v)
        tblBorders.append(el)
    tblPr.append(tblBorders)


def _apply_cell_borders(cell, sides):
    """Apply borders to a specific cell. sides: dict side -> attrs_dict"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    old = tcPr.find(qn("w:tcBorders"))
    if old is not None:
        tcPr.remove(old)
    tcBorders = OxmlElement("w:tcBorders")
    for side, attrs in sides.items():
        el = OxmlElement(f"w:{side}")
        for k, v in attrs.items():
            el.set(qn(f"w:{k}"), v)
        tcBorders.append(el)
    tcPr.append(tcBorders)


def _set_cell_width(cell, twips):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    old = tcPr.find(qn("w:tcW"))
    if old is not None:
        tcPr.remove(old)
    tcW = OxmlElement("w:tcW")
    tcW.set(qn("w:w"), str(twips))
    tcW.set(qn("w:type"), "dxa")
    tcPr.append(tcW)


def _tight_cell(cell, twips=None):
    """Zero padding on a cell."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    old = tcPr.find(qn("w:tcMar"))
    if old is not None:
        tcPr.remove(old)
    mar = OxmlElement("w:tcMar")
    for side in ("top", "bottom", "left", "right"):
        s = OxmlElement(f"w:{side}")
        val = "30" if side in ("left", "right") else "0"
        s.set(qn("w:w"), val)
        s.set(qn("w:type"), "dxa")
        mar.append(s)
    tcPr.append(mar)
    if twips:
        _set_cell_width(cell, twips)


def _remove_borders(table):
    _apply_tbl_borders(table, {
        s: _BORDER_NONE for s in ("top","left","bottom","right","insideH","insideV")
    })


def _p(cell, text, bold=False, size=6, align=WD_ALIGN_PARAGRAPH.LEFT):
    para = cell.add_paragraph()
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after  = Pt(0)
    para.alignment = align
    run = para.add_run(text)
    run.font.size = Pt(size)
    run.bold = bold
    return para


def _p_tab(cell, label, value, bold=False, size=6):
    """Label left + value right-aligned via tab."""
    para = cell.add_paragraph()
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after  = Pt(0)
    pPr = para._p.get_or_add_pPr()
    tabs_el = OxmlElement("w:tabs")
    tab = OxmlElement("w:tab")
    tab.set(qn("w:val"), "right")
    tab.set(qn("w:pos"), str(int(Cm(8.6).pt * 20)))
    tabs_el.append(tab)
    pPr.append(tabs_el)
    run = para.add_run(f"{label}\t{value}")
    run.font.size = Pt(size)
    run.bold = bold


# ── Header ────────────────────────────────────────────────────────────────────

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
    _p(left, "Cliente: SUPRICOM CCS 21, C.A.", bold=True, size=9)
    _p(left, "CALLE LOS LABORATORIOS EDIF. OFINCA PISO PB LOCAL 2-A, LOS RUISES, CARACAS, MIRANDA", size=8)
    _p(left, "Distrito Capital DTC Distrito Capital", size=8)
    _p(left, "Venezuela — J501193738", size=8)

    right.text = ""
    _p(right, f"Número de Factura: {invoice_no}", bold=True, size=9, align=WD_ALIGN_PARAGRAPH.RIGHT)
    _p(right, f"Fecha de Emisión: {invoice_date}", size=8, align=WD_ALIGN_PARAGRAPH.RIGHT)
    _p(right, f"Termino de Pago: {terms}", size=8, align=WD_ALIGN_PARAGRAPH.RIGHT)


# ── Item rows ─────────────────────────────────────────────────────────────────

def _row_para(doc, cols, bold=False, font_size=5.5):
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after  = Pt(0.5)
    _add_tab_stops(para, [
        (TAB_DESC,   "left"),
        (TAB_CANT,   "right"),
        (TAB_PRECIO, "right"),
        (TAB_TOTAL,  "right"),
    ])
    run = para.add_run(f"{cols[0]}\t{cols[1]}\t{cols[2]}\t{cols[3]}\t{cols[4]}")
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


# ── Totals box ────────────────────────────────────────────────────────────────

def _totals_box(doc, subtotal, iva_rate, tasa_bcv):
    """
    Single outer-bordered box. Structure (9 rows × 2 cols):
      Row 0  : column titles  (2 cols, vertical divider between them)
      Rows 1-5: data rows     (2 cols, vertical divider between them)
      Row 6  : IGTF text      (merged full-width, no borders inside)
      Row 7  : Tasa vigente   (merged full-width, no borders inside)
      Row 8  : footer legal   (merged full-width, TOP border = the only horiz line)

    Table-level borders: outer only (top/left/bottom/right), insideH=none, insideV=none
    Cell-level: left cell of data rows gets right border = vertical divider
    Footer cell gets top border = horizontal divider
    """
    subtotal   = round(subtotal, 2)
    iva_pct    = int(iva_rate * 100)
    iva_amount = round(subtotal * iva_rate, 2)
    total_usd  = round(subtotal + iva_amount, 2)

    sub_bs   = round(subtotal   * tasa_bcv, 2)
    iva_bs   = round(iva_amount * tasa_bcv, 2)
    total_bs = round(total_usd  * tasa_bcv, 2)

    today    = date.today().strftime("%-d/%m/%Y")
    tasa_fmt = f"{tasa_bcv:,.4f}"

    N_DATA = 6   # title row + 5 data rows
    table = doc.add_table(rows=N_DATA + 3, cols=2)

    # Table-level: outer border only, no inside lines
    _apply_tbl_borders(table, {
        "top":     _BORDER_SINGLE,
        "left":    _BORDER_SINGLE,
        "bottom":  _BORDER_SINGLE,
        "right":   _BORDER_SINGLE,
        "insideH": _BORDER_NONE,
        "insideV": _BORDER_NONE,
    })

    COL_W = 7150
    for row in table.rows:
        for cell in row.cells:
            _tight_cell(cell, twips=COL_W)

    # ── Row 0: titles ────────────────────────────────────────────────────────
    r0l = table.cell(0, 0)
    r0r = table.cell(0, 1)
    r0l.text = r0r.text = ""
    _p(r0l, "Monto en Dólares  (U$S.)", bold=True, size=6, align=WD_ALIGN_PARAGRAPH.CENTER)
    _p(r0r, "Monto en Bolívares  (Bs.)", bold=True, size=6, align=WD_ALIGN_PARAGRAPH.CENTER)
    # vertical divider on title row
    _apply_cell_borders(r0l, {"right": _BORDER_SINGLE})
    _apply_cell_borders(r0r, {"left":  _BORDER_NONE})

    # ── Rows 1-5: data ───────────────────────────────────────────────────────
    data_rows = [
        ("Sub-Total:",          _m(subtotal),    "Sub-Total Ref:",           _m(sub_bs)),
        ("Base Imponible:",     _m(subtotal),    "Base Imponible Ref:",      _m(sub_bs)),
        (f"IVA {iva_pct}% Sobre {_m(subtotal)} :", _m3(iva_amount),
         f"IVA {iva_pct}% Ref Sobre {_m(sub_bs)}  :", _m3(iva_bs)),
        ("IGTF 3% Sugerido:",   _m(0),           "IGTF 3% Sugerido Ref:",   _m(0)),
        ("Total General U$S:",  _m(total_usd),   "Total General Ref Bs.:",  _m(total_bs)),
    ]

    for i, (lu, vu, lb, vb) in enumerate(data_rows):
        row = table.rows[i + 1]
        cl, cr = row.cells[0], row.cells[1]
        cl.text = cr.text = ""
        is_total = (i == 4)
        _p_tab(cl, lu, vu, bold=is_total, size=6)
        _p_tab(cr, lb, vb, bold=is_total, size=6)
        # vertical divider only
        _apply_cell_borders(cl, {"right": _BORDER_SINGLE})
        _apply_cell_borders(cr, {"left":  _BORDER_NONE})

    # ── Row 6: IGTF legal (full width, no extra borders) ─────────────────────
    igtf_cell = table.cell(6, 0).merge(table.cell(6, 1))
    igtf_cell.text = ""
    _tight_cell(igtf_cell)
    _apply_cell_borders(igtf_cell, {})   # no extra borders
    _p(igtf_cell, IGTF_LEGAL, size=5.5)

    # ── Row 7: Tasa vigente (full width, no extra borders) ───────────────────
    tasa_cell = table.cell(7, 0).merge(table.cell(7, 1))
    tasa_cell.text = ""
    _tight_cell(tasa_cell)
    _apply_cell_borders(tasa_cell, {})
    _p(tasa_cell, f"Tasa Vigente al {today} según el BCV: {tasa_fmt}", size=5.5)

    # ── Row 8: footer legal (full width, TOP border = the only horiz line) ───
    footer_cell = table.cell(8, 0).merge(table.cell(8, 1))
    footer_cell.text = ""
    _tight_cell(footer_cell)
    _apply_cell_borders(footer_cell, {"top": _BORDER_SINGLE})
    _p(footer_cell, FOOTER_LEGAL.format(tasa=tasa_fmt), size=5.5)


# ── Main ──────────────────────────────────────────────────────────────────────

def generate_docx(invoice_data: dict, iva_rate: float = 0.16, tasa_bcv: float = 1.0) -> bytes:
    hdr   = invoice_data["header"]
    items = invoice_data["items"]

    doc = Document()

    for section in doc.sections:
        section.top_margin    = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin   = Cm(2)
        section.right_margin  = Cm(1.5)

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

    _row_para(doc, ["CÓDIGO", "DESCRIPCIÓN", "CANT.", "PRECIO", "TOTAL"],
              bold=True, font_size=6)
    _divider(doc)

    for item in items:
        _row_para(doc, [
            item["codigo"],
            item["descripcion"],
            _fmt_qty(item["cantidad"]),
            _m(item["precio"]),
            _m(item["total"]),
        ], font_size=5.5)

    _divider(doc)
    _space(doc, 4)

    subtotal = round(sum(i["total"] for i in items), 2)
    _totals_box(doc, subtotal, iva_rate, tasa_bcv)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ── Formatting helpers ────────────────────────────────────────────────────────

def _space(doc, after_pt=4):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(after_pt)


def _m(val) -> str:
    return f"{round(float(val), 2):,.2f}"


def _m3(val) -> str:
    return f"{round(float(val), 3):,.3f}"


def _fmt_qty(val) -> str:
    if float(val) == int(float(val)):
        return str(int(float(val)))
    return str(val)
