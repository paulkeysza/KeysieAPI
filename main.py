from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
import json, base64
from io import BytesIO
from openpyxl import Workbook
from datetime import datetime

app = FastAPI(
    title="Qiesi JSON → Excel API",
    version="1.1.0",
    description="""
A lightweight, blazing-fast JSON → Excel converter API for demos, proof-of-concepts, and workflow automation.

**Features**
- Accepts raw JSON (object or array)
- Converts to Excel (.xlsx)
- Returns Base64 file inline
- Perfect for RPA / NAC / integration demos
"""
)

class ConvertRequest(BaseModel):
    jsonInput: str

@app.get("/health", tags=["System"])
def health():
    return {"status": "ok"}

@app.get("/info", tags=["System"])
def info():
    return {
        "name": "Qiesi JSON → Excel API",
        "version": "1.1.0",
        "author": "Paul Keys",
        "description": "High-speed JSON → Excel conversion API for demo and automation scenarios.",
        "endpoints": {
            "health": "/health",
            "convert": "/convert",
            "docs": "/docs",
            "openapi": "/openapi.json"
        }
    }

@app.get("/ping", tags=["System"])
def ping():
    return {"message": "pong", "api": "Qiesi JSON → Excel"}

@app.post("/convert", tags=["Conversion"])
def convert(req: ConvertRequest):
    ...
