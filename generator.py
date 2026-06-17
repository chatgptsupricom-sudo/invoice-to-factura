from io import BytesIO
from copy import deepcopy
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import copy


TEMPLATE_PATH = "template.docx"


def _set_cell_text(cell, text, bold=False, font_size=None, align=None):
    cell.text = ""
    para = cell.paragraphs[0]
    run = para.add_run(str(text))
    if bold:
        run.bold = True
    if font_size:
        run.font.size = Pt(font_size)
    if align:
        para.alignment = align


def _copy_row_format(source_row, target_row):
    """Copy XML element properties from source row to target row."""
    for src_cell, tgt_cell in zip(source_row.cells, target_row.cells):
        tgt_cell._tc.get_or_add_tcPr()
        src_tcPr = src_cell._tc.find(qn("w:tcPr"))
        if src_tcPr is not None:
            tgt_tcPr = tgt_cell._tc.find(qn("w:tcPr"))
            if tgt_tcPr is not None:
                tgt_cell._tc.remove(tgt_tcPr)
            tgt_cell._tc.insert(0, deepcopy(src_tcPr))


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

    # Keep only header row and one blank row as template, remove rest
    # Find first data row index (skip header row 0 and blank row 1)
    template_data_row = table.rows[1]  # blank row used as format template

    # Remove all rows after row 0 (keep header)
    rows_to_remove = list(table.rows)[1:]
    for row in rows_to_remove:
        tbl = table._tbl
        tbl.remove(row._tr)

    # Add item rows
    for item in items:
        row = table.add_row()
        _copy_row_format(template_data_row, row)
        row.cells[0].text = str(item["codigo"])
        row.cells[1].text = str(item["descripcion"])
        row.cells[2].text = _fmt_qty(item["cantidad"])
        row.cells[3].text = _fmt_money(item["precio"])
        row.cells[4].text = _fmt_money(item["total"])

    # --- Totals section ---
    subtotal = sum(i["total"] for i in items)
    iva_amount = round(subtotal * iva_rate, 2)
    total_general = round(subtotal + iva_amount, 2)

    # Add spacer row
    spacer = table.add_row()
    for cell in spacer.cells:
        cell.text = ""

    # Subtotal row
    _add_total_row(table, "SUBTOTAL:", subtotal)
    # IVA row
    _add_total_row(table, f"IVA ({int(iva_rate*100)}%):", iva_amount)
    # Total general row
    _add_total_row(table, "TOTAL GENERAL:", total_general, bold=True)

    # Save to buffer
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


def _add_total_row(table, label: str, amount: float, bold: bool = False):
    row = table.add_row()
    # Merge first 4 cells for label
    row.cells[0].text = ""
    row.cells[1].text = ""
    row.cells[2].text = ""
    row.cells[3].text = label
    para = row.cells[3].paragraphs[0]
    para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    if bold:
        for run in para.runs:
            run.bold = True
    row.cells[4].text = _fmt_money(amount)
    para4 = row.cells[4].paragraphs[0]
    para4.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    if bold:
        for run in para4.runs:
            run.bold = True
