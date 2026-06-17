from io import BytesIO
from copy import deepcopy
from docx import Document
from docx.shared import Pt, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


TEMPLATE_PATH = "template.docx"

# Column widths in twips (from template measurement, total ~11287)
COL_WIDTHS = [1800, 5800, 900, 1300, 1487]  # Código, Descripción, Cantidad, Precio, Total


def _set_col_width(cell, width_twips):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcW = tcPr.find(qn("w:tcW"))
    if tcW is None:
        tcW = OxmlElement("w:tcW")
        tcPr.append(tcW)
    tcW.set(qn("w:w"), str(width_twips))
    tcW.set(qn("w:type"), "dxa")


def _set_cell(cell, text, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT, font_size=9):
    cell.text = ""
    para = cell.paragraphs[0]
    para.alignment = align
    run = para.add_run(str(text))
    run.font.size = Pt(font_size)
    if bold:
        run.bold = True


def _add_row_with_widths(table, values, bold=False, aligns=None, font_size=9):
    """Add a row with fixed column widths."""
    row = table.add_row()
    if aligns is None:
        aligns = [WD_ALIGN_PARAGRAPH.LEFT] * len(values)
    for i, (cell, val, width) in enumerate(zip(row.cells, values, COL_WIDTHS)):
        _set_col_width(cell, width)
        _set_cell(cell, val, bold=bold, align=aligns[i], font_size=font_size)
    return row


def generate_docx(invoice_data: dict, iva_rate: float = 0.16) -> bytes:
    header = invoice_data["header"]
    items = invoice_data["items"]

    doc = Document(TEMPLATE_PATH)

    # --- Update header paragraphs ---
    for para in doc.paragraphs:
        t = para.text
        if "Número de Factura:" in t:
            _replace_para(para, f"Número de Factura: {header.get('invoice_no', '')}")
        elif "Fecha de Emisión:" in t:
            _replace_para(para, f"Fecha de Emisión: {header.get('invoice_date', '')}")
        elif "Termino de Pago:" in t:
            _replace_para(para, f"Termino de Pago: {header.get('terms', '')}")

    # --- Work with items table (first table) ---
    table = doc.tables[0]

    # Fix header row widths
    header_row = table.rows[0]
    header_texts = ["Código", "Descripción", "Cantidad", "Precio", "Total"]
    for i, (cell, width) in enumerate(zip(header_row.cells, COL_WIDTHS)):
        _set_col_width(cell, width)

    # Remove all rows after header
    rows_to_remove = list(table.rows)[1:]
    for row in rows_to_remove:
        table._tbl.remove(row._tr)

    # Right-align for numeric columns
    R = WD_ALIGN_PARAGRAPH.RIGHT
    L = WD_ALIGN_PARAGRAPH.LEFT

    # Add item rows
    for item in items:
        _add_row_with_widths(
            table,
            [
                item["codigo"],
                item["descripcion"],
                _fmt_qty(item["cantidad"]),
                _fmt_money(item["precio"]),
                _fmt_money(item["total"]),
            ],
            aligns=[L, L, R, R, R],
        )

    # --- Totals ---
    subtotal = sum(i["total"] for i in items)
    iva_amount = round(subtotal * iva_rate, 2)
    total_general = round(subtotal + iva_amount, 2)

    # Spacer
    _add_row_with_widths(table, ["", "", "", "", ""])

    # Subtotal
    _add_row_with_widths(
        table, ["", "", "", "SUBTOTAL:", _fmt_money(subtotal)],
        aligns=[L, L, L, R, R], font_size=9,
    )
    # IVA
    _add_row_with_widths(
        table, ["", "", "", f"IVA ({int(iva_rate*100)}%):", _fmt_money(iva_amount)],
        aligns=[L, L, L, R, R], font_size=9,
    )
    # Total General
    _add_row_with_widths(
        table, ["", "", "", "TOTAL GENERAL:", _fmt_money(total_general)],
        bold=True, aligns=[L, L, L, R, R], font_size=10,
    )

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


def _replace_para(para, new_text):
    for run in para.runs:
        run.text = ""
    if para.runs:
        para.runs[0].text = new_text
    else:
        para.add_run(new_text)


def _fmt_money(val) -> str:
    return f"{val:,.2f}"


def _fmt_qty(val) -> str:
    if val == int(val):
        return str(int(val))
    return str(val)
