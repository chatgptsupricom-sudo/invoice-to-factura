import pdfplumber
import re
from io import BytesIO


# Matches: optional "NN. " prefix, SKU, description, qty, $rate, $amount
_ITEM_RE = re.compile(
    r"(?:^\d+\.\s+)?"                 # optional line number "12. "
    r"([A-Z0-9][\w/\-\.+]+)"          # SKU: starts uppercase/digit, no spaces
    r"\s+(.+?)\s+"                    # description (non-greedy)
    r"(\d[\d,]*)\s+"                  # quantity (integer or decimal)
    r"\$([0-9,]+\.[0-9]{2})\s+"       # rate  $X.XX
    r"\$([0-9,]+\.[0-9]{2})"          # amount $X.XX
)

# Standalone item number at start of text block "75.\n" means next block continues
_NUM_PREFIX_RE = re.compile(r"^\d+\.\s*$")


def _parse_number(s: str) -> float:
    return float(s.replace(",", "").replace("$", "").strip())


def extract_invoice(pdf_bytes: bytes) -> dict:
    header = {
        "invoice_no": "",
        "invoice_date": "",
        "due_date": "",
        "terms": "",
    }
    items = []

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        full_text = ""
        page_texts = []
        for page in pdf.pages:
            t = page.extract_text() or ""
            full_text += t + "\n"
            page_texts.append(t)

    # Extract header from full text
    m = re.search(r"Invoice no\.:\s*(\S+)", full_text)
    if m:
        header["invoice_no"] = m.group(1)
    m = re.search(r"Invoice date:\s*([\d/]+)", full_text)
    if m:
        header["invoice_date"] = m.group(1)
    m = re.search(r"Due date:\s*([\d/]+)", full_text)
    if m:
        header["due_date"] = m.group(1)
    m = re.search(r"Terms:\s*(.+)", full_text)
    if m:
        header["terms"] = m.group(1).strip()

    # Parse items line by line across all pages
    # Build a continuous list of lines, stripping the "NN." prefix per line
    lines = full_text.splitlines()
    pending = None  # pending incomplete item text

    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            continue

        # Try to match a complete item on this line (possibly with leading number)
        # Strip leading "NN. " if present
        stripped = re.sub(r"^\d+\.\s*", "", line)

        m = _ITEM_RE.match(stripped) or _ITEM_RE.match(line)
        if m:
            # Save any pending partial item first
            if pending:
                result = _try_parse_pending(pending)
                if result:
                    items.append(result)
                pending = None
            # Full match on this line
            sku = m.group(1)
            desc = m.group(2)
            qty = _parse_number(m.group(3))
            rate = _parse_number(m.group(4))
            total = round(qty * rate, 2)
            items.append({
                "codigo": sku,
                "descripcion": desc,
                "cantidad": qty,
                "precio": rate,
                "total": total,
            })
        else:
            # Could be a continuation line or header garbage — append to pending
            if pending:
                pending += " " + stripped
                # Try to parse now that we have more text
                result = _try_parse_pending(pending)
                if result:
                    items.append(result)
                    pending = None
            else:
                # Start new pending if line looks like beginning of an item
                if re.match(r"^[A-Z0-9][\w/\-\.+]+\s+", stripped):
                    pending = stripped

    if pending:
        result = _try_parse_pending(pending)
        if result:
            items.append(result)

    return {"header": header, "items": items}


def _try_parse_pending(text: str):
    m = _ITEM_RE.match(text)
    if not m:
        # Try without leading number
        stripped = re.sub(r"^\d+\.\s+", "", text)
        m = _ITEM_RE.match(stripped)
    if m:
        qty = _parse_number(m.group(3))
        rate = _parse_number(m.group(4))
        return {
            "codigo": m.group(1),
            "descripcion": m.group(2),
            "cantidad": qty,
            "precio": rate,
            "total": round(qty * rate, 2),
        }
    return None
