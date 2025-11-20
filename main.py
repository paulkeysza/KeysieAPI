from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
import json, base64
from io import BytesIO
from openpyxl import Workbook
from datetime import datetime

app = FastAPI(
    title="KeysieAPI JSON→Excel Converter",
    version="1.0.0",
    description="Convert raw JSON to Excel (Base64 output)."
)

class ConvertRequest(BaseModel):
    jsonInput: str

@app.get("/health")
def health():
    return {"status": "ok"}

@app.post("/convert")
def convert(req: ConvertRequest):
    try:
        data = json.loads(req.jsonInput)
    except:
        raise HTTPException(status_code=400, detail="ERR002: Invalid JSON format.")

    if isinstance(data, dict):
        data=[data]
    if not isinstance(data, list):
        raise HTTPException(status_code=400, detail="ERR003: JSON must be array or object.")

    try:
        wb=Workbook()
        ws=wb.active
        headers=set()
        for item in data:
            if isinstance(item, dict):
                headers.update(item.keys())
            else:
                raise HTTPException(status_code=400, detail="ERR003: Items must be objects.")
        headers=list(headers)
        ws.append(headers)
        for item in data:
            ws.append([item.get(h) for h in headers])

        bio=BytesIO()
        wb.save(bio)
        b64=base64.b64encode(bio.getvalue()).decode()
    except:
        raise HTTPException(status_code=500, detail="ERR004: Excel generation failed.")

    ts=datetime.utcnow().strftime("%Y%m%d%H%M%S")
    return {"fileName": f"KeysieAPI-{ts}.xlsx", "excelFile": b64}