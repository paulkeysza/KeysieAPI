from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
import json
import base64
import csv
from io import BytesIO, StringIO
from openpyxl import Workbook
from datetime import datetime

app = FastAPI(
    title="Qiesi API Toolkit",
    version="1.1.0",
    description="""
A lightweight API toolkit providing:

• JSON → Excel conversion
• Text → CSV conversion

"""
)

class ConvertRequest(BaseModel):
    jsonInput: str


class MessageCSVRequest(BaseModel):
    message: str


# -----------------------------
# System Endpoints
# -----------------------------

@app.get("/health", tags=["System"])
def health():
    return {"status": "ok"}


@app.get("/info", tags=["System"])
def info():
    return {
        "name": "Qiesi API Toolkit",
        "version": "1.1.0",
        "author": "Paul Keys",
        "description": "My API toolkit",
        "endpoints": {
            "health": "/health",
            "json_to_xlsx": "/JSON_to_XLSX",
            "txt_to_csv": "/TXT_to_CSV",
            "docs": "/docs",
            "openapi": "/openapi.json"
        }
    }


@app.get("/ping", tags=["System"])
def ping():
    return {"message": "pong", "api": "Qiesi JSON → Excel"}


# -----------------------------
# JSON → Excel Conversion
# -----------------------------

@app.post(
        "/JSON-to-XLSX", 
        tags=["Conversion"],
        summary="JSON-to-MS Excel",
        description="Accepts JSON and returns it as a Base64 encoded Excel (.xlsx) file.")
def convert(req: ConvertRequest):

    try:

        data = json.loads(req.jsonInput)

        if isinstance(data, dict):
            data = [data]

        wb = Workbook()
        ws = wb.active
        ws.title = "Data"

        headers = list(data[0].keys())
        ws.append(headers)

        for row in data:
            ws.append([row.get(h) for h in headers])

        buffer = BytesIO()
        wb.save(buffer)
        buffer.seek(0)

        encoded = base64.b64encode(buffer.read()).decode()

        filename = f"QiesiAPI-{datetime.utcnow().strftime('%Y%m%d%H%M%S')}.xlsx"

        return {
            "fileName": filename,
            "excelFile": encoded
        }

    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


# -----------------------------
# Message → CSV Conversion
# Supports multiline input
# -----------------------------

@app.post(
    "/TXT-to-CSV",
    tags=["Conversion"],
    summary="Text-To-CSV",
    description="Accepts text and returns it as a Base64 encoded CSV file."
)
def message_to_csv(req: MessageCSVRequest):

    try:

        output = StringIO()
        writer = csv.writer(output)

        lines = req.message.splitlines()

        for line in lines:
            writer.writerow([line])

        csv_bytes = output.getvalue().encode("utf-8")

        encoded = base64.b64encode(csv_bytes).decode()

        filename = f"QiesiAPI-{datetime.utcnow().strftime('%Y%m%d%H%M%S')}.csv"

        return {
            "fileName": filename,
            "csvFile": encoded
        }

    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))