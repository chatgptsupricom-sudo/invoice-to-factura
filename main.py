import json
import os
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Request, Query
from fastapi.responses import Response, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from extractor import extract_invoice
from generator import generate_docx
from audit import init_db, log_conversion_full, get_all

ADMIN_KEY = os.environ.get("ADMIN_KEY", "supricom2024")

app = FastAPI(title="Invoice to Factura")
init_db()

app.mount("/static", StaticFiles(directory="static"), name="static")


@app.get("/")
async def index():
    return FileResponse("static/index.html")


@app.post("/convert")
async def convert(
    request: Request,
    pdf: UploadFile = File(...),
    iva_rate: float = Form(0.16),
    tasa_bcv: float = Form(1.0),
    items_json: str = Form(""),
    header_json: str = Form(""),
    original_qty: int = Form(0),
    deleted_items_json: str = Form(""),
):
    if not pdf.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="El archivo debe ser un PDF.")

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

    # Audit log
    ip = request.headers.get("x-forwarded-for", request.client.host if request.client else "?")
    ip = ip.split(",")[0].strip()
    log_conversion_full(
        ip=ip,
        invoice_no=invoice_data["header"].get("invoice_no", "?"),
        original_qty=original_qty or len(invoice_data["items"]),
        final_items=invoice_data["items"],
        deleted_items_json=deleted_items_json or None,
        iva_rate=iva_rate,
        tasa_bcv=tasa_bcv,
    )

    invoice_no = invoice_data["header"].get("invoice_no", "factura")
    filename = f"Factura_Fiscal_{invoice_no}.docx"

    return Response(
        content=docx_bytes,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.post("/preview")
async def preview(pdf: UploadFile = File(...)):
    if not pdf.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="El archivo debe ser un PDF.")
    pdf_bytes = await pdf.read()
    try:
        invoice_data = extract_invoice(pdf_bytes)
    except Exception as e:
        raise HTTPException(status_code=422, detail=f"No se pudo leer el PDF: {e}")
    return invoice_data


@app.get("/admin")
async def admin_page(key: str = Query("")):
    if key != ADMIN_KEY:
        raise HTTPException(status_code=401, detail="No autorizado.")
    return FileResponse("static/admin.html")


@app.get("/admin/logs")
async def admin_logs(key: str = Query("")):
    if key != ADMIN_KEY:
        raise HTTPException(status_code=401, detail="No autorizado.")
    return JSONResponse(get_all())
