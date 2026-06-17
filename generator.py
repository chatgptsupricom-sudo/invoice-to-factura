from io import BytesIO
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


TEMPLATE_PATH = "template.docx"

# Tab stop positions in cm from left margin
# Código(0) | Descripción(3.5) | Cantidad(13) | Precio(15) | Total(18)
TAB_CODIGO = 0
TAB_DESC   = 3.5
TAB_CANT   = 13.2
TAB_PRECIO = 15.5
TAB_TOTAL  = 18.5


def _cm(val):
    return Cm(val)


def _add_tab_stops(para, stops):
    """Add tab stops to a paragraph. stops = list of (position_cm, alignment)."""
    pPr = para._p.get_or_add_pPr()
    tabs_el = OxmlElement("w:tabs")
    for pos_cm, align in stops:
        tab = OxmlElement("w:tab")
        tab.set(qn("w:val"), align)
        tab.set(qn("w:pos"), str(int(Cm(pos_cm).pt * 20)))
        tabs_el.append(tab)
    pPr.append(tabs_el)


def _row_para(doc, cols, bold=False, font_size=9, header=False):
    """Add a paragraph with tab-separated columns."""
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(1)

    stops = [
        (TAB_DESC,   "left"),
        (TAB_CANT,   "right"),
        (TAB_PRECIO, "right"),
        (TAB_TOTAL,  "right"),
    ]
    _add_tab_stops(para, stops)

    # Build: col0 \t col1 \t col2 \t col3 \t col4
    text = f"{cols[0]}\t{cols[1]}\t{cols[2]}\t{cols[3]}\t{cols[4]}"
    run = para.add_run(text)
    run.font.size = Pt(font_size)
    run.bold = bold
    if header:
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    if header:
        # Gray background via paragraph shading
        pPr = para._p.get_or_add_pPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), "2D3748")
        pPr.append(shd)

    return para


def _divider(doc):
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(0)
    run = para.add_run("─" * 110)
    run.font.size = Pt(7)
    run.font.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)


def generate_docx(invoice_data: dict, iva_rate: float = 0.16) -> bytes:
    header = invoice_data["header"]
    items = invoice_data["items"]

    doc = Document()

    # --- Page margins ---
    for section in doc.sections:
        section.top_margin    = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin   = Cm(2)
        section.right_margin  = Cm(1.5)

    # --- Company header ---
    _heading(doc, "SUPRICOM CCS 21, C.A.", size=14, bold=True)
    _heading(doc, "CALLE LOS LABORATORIOS EDIF. OFINCA PISO PB LOCAL 2-A, LOS RUISES, CARACAS, MIRANDA", size=9)
    _heading(doc, "Distrito Capital — Venezuela — J501193738", size=9)
    _space(doc)

    # Invoice details block
    _detail(doc, "N° Factura",     header.get("invoice_no", ""))
    _detail(doc, "Fecha",          header.get("invoice_date", ""))
    _detail(doc, "Vencimiento",    header.get("due_date", ""))
    _detail(doc, "Términos",       header.get("terms", ""))
    _space(doc)

    # --- Column header row ---
    _row_para(doc,
              ["CÓDIGO", "DESCRIPCIÓN", "CANT", "PRECIO", "TOTAL"],
              bold=True, font_size=9, header=True)

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
    subtotal     = sum(i["total"] for i in items)
    iva_amount   = round(subtotal * iva_rate, 2)
    total_general = round(subtotal + iva_amount, 2)

    _total_row(doc, "Subtotal:",                   subtotal)
    _total_row(doc, f"IVA ({int(iva_rate*100)})%:", iva_amount)
    _total_row(doc, "TOTAL GENERAL:",               total_general, bold=True, size=11)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ── Helpers ──────────────────────────────────────────────────────────────────

def _heading(doc, text, size=10, bold=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(2)
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.bold = bold


def _detail(doc, label, value):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(1)
    r1 = p.add_run(f"{label}: ")
    r1.font.size = Pt(9)
    r1.bold = True
    r2 = p.add_run(value)
    r2.font.size = Pt(9)


def _space(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(4)


def _total_row(doc, label, amount, bold=False, size=10):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(2)
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    stops = [(TAB_TOTAL, "right")]
    _add_tab_stops(p, stops)

    r = p.add_run(f"{label}    {_fmt_money(amount)}")
    r.font.size = Pt(size)
    r.bold = bold


def _fmt_money(val) -> str:
    return f"{val:,.2f}"


def _fmt_qty(val) -> str:
    if float(val) == int(float(val)):
        return str(int(float(val)))
    return str(val)
