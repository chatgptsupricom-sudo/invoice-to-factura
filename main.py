import json
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import Response, FileResponse
from fastapi.staticfiles import StaticFiles

from extractor import extract_invoice
from generator import generate_docx

app = FastAPI(title="Invoice to Factura")

app.mount("/static", StaticFiles(directory="static"), name="static")


@app.get("/")
async def index():
    return FileResponse("static/index.html")


@app.post("/convert")
async def convert(
    pdf: UploadFile = File(...),
    iva_rate: float = Form(0.16),
    tasa_bcv: float = Form(1.0),
    items_json: str = Form(""),
    header_json: str = Form(""),
):
    if not pdf.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="El archivo debe ser un PDF.")

    # Use items/header sent from frontend (may reflect user deletions)
    if items_json and header_json:
        try:
            invoice_data = {"header": json.loads(header_json), "items": json.loads(items_json)}
        except Exception as e:
            raise HTTPException(status_code=400, detail=f"items_json inválido: {e}")
    else:
        pdf_bytes = await pdf.read()
        try:
            invoice_data = extract_invoice(pdf_bytes)
        except Exception as e:
            raise HTTPException(status_code=422, detail=f"No se pudo leer el PDF: {e}")

    if not invoice_data["items"]:
        raise HTTPException(status_code=422, detail="No se encontraron ítems en el PDF.")

    try:
        docx_bytes = generate_docx(invoice_data, iva_rate=iva_rate, tasa_bcv=tasa_bcv)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error al generar el Word: {e}")

    invoice_no = invoice_data["header"].get("invoice_no", "factura")
    filename = f"Factura_Fiscal_{invoice_no}.docx"

    return Response(
        content=docx_bytes,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.post("/preview")
async def preview(pdf: UploadFile = File(...)):
    """Return extracted items as JSON for preview."""
    if not pdf.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="El archivo debe ser un PDF.")

    pdf_bytes = await pdf.read()

    try:
        invoice_data = extract_invoice(pdf_bytes)
    except Exception as e:
        raise HTTPException(status_code=422, detail=f"No se pudo leer el PDF: {e}")

    return invoice_data
