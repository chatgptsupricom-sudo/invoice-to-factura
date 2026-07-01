import pdfplumber
import re
from io import BytesIO


# ── Format A: original Invoice (e.g. Invoice_1042) ───────────────────────────
# Matches: optional "NN. " prefix, SKU, description, qty, $rate, $amount
_ITEM_RE = re.compile(
    r"(?:^\d+\.\s+)?"
    r"([A-Z0-9][\w/\-\.+]+)"
    r"\s+(.+?)\s+"
    r"(\d[\d,]*)\s+"
    r"\$([0-9,]+\.[0-9]{2})\s+"
    r"\$([0-9,]+\.[0-9]{2})"
)

# ── Format B: Factura with [CODE] in brackets ─────────────────────────────────
# Lines like: [90NR0N06-M00950] ASUS TUF F16 ... 4,00 1.100,00 $ 4.400,00
_BRACKET_ITEM_RE = re.compile(
    r"^\[([^\]]+)\]\s+"          # [CODE]
    r"(.+?)\s+"                  # description (non-greedy)
    r"(\d[\d.]*,\d{2})\s+"       # qty  e.g. 4,00
    r"(\d[\d.]*,\d{2})\s+"       # price e.g. 1.100,00
    r"\$\s*([\d.,]+)"            # $ amount
)


def _parse_number_en(s: str) -> float:
    """Parse English-format number: 1,234.56"""
    return float(s.replace(",", "").replace("$", "").strip())


def _parse_number_es(s: str) -> float:
    """Parse Spanish-format number: 1.234,56"""
    return float(s.replace(".", "").replace(",", ".").replace("$", "").strip())


def _detect_format(full_text: str) -> str:
    """Return 'bracket' if Format B, else 'classic'."""
    if re.search(r"\[[A-Z0-9][^\]]{2,}\]", full_text):
        return "bracket"
    return "classic"


def extract_invoice(pdf_bytes: bytes) -> dict:
    header = {"invoice_no": "", "invoice_date": "", "due_date": "", "terms": ""}
    items = []

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        full_text = ""
        for page in pdf.pages:
            full_text += (page.extract_text() or "") + "\n"

    fmt = _detect_format(full_text)

    if fmt == "bracket":
        _extract_bracket(full_text, header, items)
    else:
        _extract_classic(full_text, header, items)

    return {"header": header, "items": items}


# ── Format A parser ───────────────────────────────────────────────────────────

def _extract_classic(full_text, header, items):
    m = re.search(r"Invoice no\.:\s*(\S+)", full_text)
    if m: header["invoice_no"] = m.group(1)
    m = re.search(r"Invoice date:\s*([\d/]+)", full_text)
    if m: header["invoice_date"] = m.group(1)
    m = re.search(r"Due date:\s*([\d/]+)", full_text)
    if m: header["due_date"] = m.group(1)
    m = re.search(r"Terms:\s*(.+)", full_text)
    if m: header["terms"] = m.group(1).strip()

    lines = full_text.splitlines()
    pending = None

    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            continue
        stripped = re.sub(r"^\d+\.\s*", "", line)
        m = _ITEM_RE.match(stripped) or _ITEM_RE.match(line)
        if m:
            if pending:
                result = _try_parse_pending(pending)
                if result: items.append(result)
                pending = None
            qty  = _parse_number_en(m.group(3))
            rate = _parse_number_en(m.group(4))
            items.append({
                "codigo": m.group(1),
                "descripcion": m.group(2),
                "cantidad": qty,
                "precio": rate,
                "total": round(qty * rate, 2),
            })
        else:
            if pending:
                pending += " " + stripped
                result = _try_parse_pending(pending)
                if result:
                    items.append(result)
                    pending = None
            elif re.match(r"^[A-Z0-9][\w/\-\.+]+\s+", stripped):
                pending = stripped

    if pending:
        result = _try_parse_pending(pending)
        if result: items.append(result)


def _try_parse_pending(text: str):
    m = _ITEM_RE.match(text) or _ITEM_RE.match(re.sub(r"^\d+\.\s+", "", text))
    if m:
        qty  = _parse_number_en(m.group(3))
        rate = _parse_number_en(m.group(4))
        return {"codigo": m.group(1), "descripcion": m.group(2),
                "cantidad": qty, "precio": rate, "total": round(qty * rate, 2)}
    return None


# ── Format B parser ───────────────────────────────────────────────────────────

def _extract_bracket(full_text, header, items):
    # Header fields
    m = re.search(r"Factura\s+(\d+)", full_text)
    if m: header["invoice_no"] = m.group(1)
    # Dates may be on same line or next line after the label
    m = re.search(r"Fecha de factura:\s*([\d/]+)", full_text)
    if m:
        header["invoice_date"] = m.group(1)
    else:
        # Labels on one line, values on next: "Fecha de factura: ...\n30/06/2026 ..."
        m = re.search(r"Fecha de factura:.*?\n([\d/]+)", full_text)
        if m: header["invoice_date"] = m.group(1)
    m = re.search(r"Fecha l[ií]mite:\s*([\d/]+)", full_text)
    if m:
        header["due_date"] = m.group(1)
    else:
        m = re.search(r"Fecha l[ií]mite:.*?\n.*?([\d/]+)", full_text)
        if m: header["due_date"] = m.group(1)
    m = re.search(r"T[eé]rminos de pago:\s*(.+)", full_text)
    if m: header["terms"] = m.group(1).strip()

    lines = full_text.splitlines()
    pending_code = None
    pending_desc = None
    pending_rest = ""   # accumulated continuation lines

    SKIP_PATTERNS = re.compile(
        r"^(OFFICE SOLUTIONS|AV |GFIVETECH|MANZANA|DELICIAS|Maracay|Valencia|Venezuela|"
        r"RFC:|P[aá]gina:|Descripci[oó]n|Cantidad|Precio|Impuesto|Importe|unitario|"
        r"ZONA|INDUSTRIAL|GALPON|Total|Fuente:|Fecha)"
    , re.IGNORECASE)

    def try_flush():
        nonlocal pending_code, pending_desc, pending_rest
        if pending_code is None:
            return
        text = (pending_desc or "") + " " + pending_rest
        # Try to extract qty, price, amount from combined text
        m = re.search(
            r"(\d[\d.]*,\d{2})\s+(\d[\d.]*,\d{2})\s+\$\s*([\d.,]+)",
            text
        )
        if m:
            desc_part = text[:m.start()].strip().rstrip(",").strip()
            qty   = _parse_number_es(m.group(1))
            price = _parse_number_es(m.group(2))
            items.append({
                "codigo": pending_code,
                "descripcion": desc_part,
                "cantidad": qty,
                "precio": price,
                "total": round(qty * price, 2),
            })
        pending_code = pending_desc = None
        pending_rest = ""

    for raw_line in lines:
        line = raw_line.strip()
        if not line or line == "Unidades":
            continue
        if SKIP_PATTERNS.match(line):
            continue

        # New item starting with [CODE]
        bracket_m = _BRACKET_ITEM_RE.match(line)
        if bracket_m:
            try_flush()
            pending_code = bracket_m.group(1).strip()
            pending_desc = bracket_m.group(2).strip()
            pending_rest = ""
            # Already has numbers on same line — flush immediately
            qty   = _parse_number_es(bracket_m.group(3))
            price = _parse_number_es(bracket_m.group(4))
            items.append({
                "codigo": pending_code,
                "descripcion": pending_desc,
                "cantidad": qty,
                "precio": price,
                "total": round(qty * price, 2),
            })
            pending_code = pending_desc = None
            pending_rest = ""
            continue

        # Line starts with [CODE] but numbers are on continuation lines
        bracket_start = re.match(r"^\[([^\]]+)\]\s+(.*)", line)
        if bracket_start:
            try_flush()
            pending_code = bracket_start.group(1).strip()
            pending_desc = bracket_start.group(2).strip()
            pending_rest = ""
            continue

        # Continuation line for a pending item
        if pending_code is not None:
            pending_rest += " " + line
            # Try to resolve now if numbers are present
            text = (pending_desc or "") + " " + pending_rest
            m = re.search(
                r"(\d[\d.]*,\d{2})\s+(\d[\d.]*,\d{2})\s+\$\s*([\d.,]+)",
                text
            )
            if m:
                try_flush()

    try_flush()
