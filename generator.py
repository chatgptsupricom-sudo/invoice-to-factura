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

BORDER = {"val": "single", "sz": "4", "space": "0", "color": "000000"}
NONE_B = {"val": "none"}

COL_W = 7150  # twips per column


# ── XML helpers ──────────────────────────────────────────────────────────────

def _border_el(side, props):
    el = OxmlElement(f"w:{side}")
    for k, v in props.items():
        el.set(qn(f"w:{k}"), v)
    return el


def _set_table_borders(table, top=BORDER, left=BORDER, bottom=BORDER,
                        right=BORDER, insideH=NONE_B, insideV=NONE_B):
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    # Remove existing tblBorders
    for old in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(old)
    tblBorders = OxmlElement("w:tblBorders")
    for side, props in (("top", top), ("left", left), ("bottom", bottom),
                         ("right", right), ("insideH", insideH), ("insideV", insideV)):
        tblBorders.append(_border_el(side, props))
    tblPr.append(tblBorders)


def _set_cell_borders(cell, top=None, left=None, bottom=None, right=None):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:tcBorders")):
        tcPr.remove(old)
    tcBorders = OxmlElement("w:tcBorders")
    for side, props in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        tcBorders.append(_border_el(side, props if props else NONE_B))
    tcPr.append(tcBorders)


def _set_cell_width(cell, twips):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:tcW")):
        tcPr.remove(old)
    tcW = OxmlElement("w:tcW")
    tcW.set(qn("w:w"), str(twips))
    tcW.set(qn("w:type"), "dxa")
    tcPr.append(tcW)


def _tight_cell(cell):
    """Minimal top/bottom cell margins."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:tcMar")):
        tcPr.remove(old)
    mar = OxmlElement("w:tcMar")
    for side, twips in (("top", "0"), ("bottom", "0"), ("left", "60"), ("right", "60")):
        s = OxmlElement(f"w:{side}")
        s.set(qn("w:w"), twips)
        s.set(qn("w:type"), "dxa")
        mar.append(s)
    tcPr.append(mar)


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
    _set_table_borders(table, top=NONE_B, left=NONE_B, bottom=NONE_B,
                       right=NONE_B, insideH=NONE_B, insideV=NONE_B)


# ── Paragraph helpers ─────────────────────────────────────────────────────────

def _zero_para(para, line_spacing_pt=None):
    """Set spacing to zero on an existing paragraph."""
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after  = Pt(0)
    if line_spacing_pt:
        para.paragraph_format.line_spacing = Pt(line_spacing_pt)


def _cp(cell, text, bold=False, size=6, align=WD_ALIGN_PARAGRAPH.LEFT):
    para = cell.add_paragraph()
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after  = Pt(0)
    para.alignment = align
    run = para.add_run(text)
    run.font.size = Pt(size)
    run.bold = bold
    return para


def _cp_tab(cell, label, value, bold=False, size=6):
    para = cell.add_paragraph()
    _zero_para(para, line_spacing_pt=7)
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
    return para


# ── Header ───────────────────────────────────────────────────────────────────

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
    _cp(left, "Cliente: SUPRICOM CCS 21, C.A.", bold=True, size=9)
    _cp(left, "CALLE LOS LABORATORIOS EDIF. OFINCA PISO PB LOCAL 2-A, LOS RUISES, CARACAS, MIRANDA", size=8)
    _cp(left, "Distrito Capital DTC Distrito Capital", size=8)
    _cp(left, "Venezuela — J501193738", size=8)

    right.text = ""
    _cp(right, f"Número de Factura: {invoice_no}", bold=True, size=9, align=WD_ALIGN_PARAGRAPH.RIGHT)
    _cp(right, f"Fecha de Emisión: {date.today().strftime('%-d/%m/%Y')}", size=8, align=WD_ALIGN_PARAGRAPH.RIGHT)
    _cp(right, f"Termino de Pago: {terms}", size=8, align=WD_ALIGN_PARAGRAPH.RIGHT)


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
    Single bordered box. Structure (9 rows):
      0  : header titles  [USD col | Bs col]  — vertical divider only
      1-5: data rows      [USD col | Bs col]  — vertical divider only
      6  : IGTF legal     [merged full width] — no internal lines
      7  : Tasa vigente   [merged full width] — no internal lines
      8  : footer legal   [merged full width] — top border (horizontal line)

    Box: outer border on all 4 sides.
    Internal lines: only the vertical divider between col 0 and col 1 (rows 0-7),
                    and the horizontal line above row 8.
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

    table = doc.add_table(rows=9, cols=2)

    # Outer border only, no internal lines
    _set_table_borders(table,
                       top=BORDER, left=BORDER, bottom=BORDER, right=BORDER,
                       insideH=NONE_B, insideV=NONE_B)

    # Set widths, tight margins, and zero-space default paragraph for all cells
    for row in table.rows:
        for cell in row.cells:
            _set_cell_width(cell, COL_W)
            _tight_cell(cell)
            _zero_para(cell.paragraphs[0], line_spacing_pt=7)

    # Apply vertical divider (right border of left cell) for rows 0-7
    # and no other internal borders on any cell
    for i in range(8):
        row = table.rows[i]
        _set_cell_borders(row.cells[0], right=BORDER)   # vertical divider
        _set_cell_borders(row.cells[1], left=NONE_B)    # no double border

    # Row 8 (footer): top border = horizontal divider line
    _set_cell_borders(table.rows[8].cells[0], top=BORDER, right=NONE_B)
    _set_cell_borders(table.rows[8].cells[1], top=BORDER)

    # ── Row 0: column titles ──
    _cp(table.cell(0, 0), "Monto en Dólares  (U$S.)", bold=True, size=6,
        align=WD_ALIGN_PARAGRAPH.CENTER)
    _cp(table.cell(0, 1), "Monto en Bolívares  (Bs.)", bold=True, size=6,
        align=WD_ALIGN_PARAGRAPH.CENTER)

    # ── Rows 1-5: data ──
    data = [
        ("Sub-Total:",                          _m(subtotal),      "Sub-Total Ref:",                       _m(sub_bs)),
        ("Base Imponible:",                     _m(subtotal),      "Base Imponible Ref:",                  _m(sub_bs)),
        (f"IVA {iva_pct}% Sobre {_m(subtotal)} :", _m3(iva_amount), f"IVA {iva_pct}% Ref Sobre {_m(sub_bs)}  :", _m3(iva_bs)),
        ("IGTF 3% Sugerido:",                   _m(0),             "IGTF 3% Sugerido Ref:",                _m(0)),
        ("Total General U$S:",                  _m(total_usd),     "Total General Ref Bs.:",               _m(total_bs)),
    ]

    for i, (lu, vu, lb, vb) in enumerate(data):
        r = table.rows[i + 1]
        is_total = (i == 4)
        _cp_tab(r.cells[0], lu, vu, bold=is_total, size=6)
        _cp_tab(r.cells[1], lb, vb, bold=is_total, size=6)

    # ── Rows 6-8: full-width merged ──
    def _merge(row_idx):
        merged = table.cell(row_idx, 0).merge(table.cell(row_idx, 1))
        return merged

    igtf_cell   = _merge(6)
    tasa_cell   = _merge(7)
    footer_cell = _merge(8)

    _tight_cell(igtf_cell)
    _tight_cell(tasa_cell)
    _tight_cell(footer_cell)

    _cp(igtf_cell,   IGTF_LEGAL, size=5.5)
    _cp(tasa_cell,   f"Tasa Vigente al {today} según el BCV: {tasa_fmt}", size=5.5)
    _cp(footer_cell, FOOTER_LEGAL.format(tasa=tasa_fmt), size=5.5)


# ── Main ─────────────────────────────────────────────────────────────────────

def generate_docx(invoice_data: dict, iva_rate: float = 0.16, tasa_bcv: float = 1.0) -> bytes:
    hdr   = invoice_data["header"]
    items = invoice_data["items"]

    doc = Document()

    for section in doc.sections:
        section.top_margin    = Cm(3.5)  # Espacio superior (~5 líneas vacías para membrete)
        section.bottom_margin = Cm(2.5)  # NUEVO: Espacio inferior controlado para evitar desbordes
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


# ── Format helpers ────────────────────────────────────────────────────────────

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
