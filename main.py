import os
import re
import json
import base64
import binascii
import csv
import mimetypes
from io import BytesIO, StringIO
from datetime import datetime
from typing import Optional, BinaryIO

from fastapi import FastAPI, HTTPException, UploadFile, File, Body
from pydantic import BaseModel
from openpyxl import Workbook
from markitdown import MarkItDown, StreamInfo, PRIORITY_SPECIFIC_FILE_FORMAT
from markitdown import UnsupportedFormatException, MissingDependencyException, FileConversionException

try:
    from markitdown.converters._doc_intel_converter import DocumentIntelligenceConverter
except Exception:
    DocumentIntelligenceConverter = None

app = FastAPI(
    title="Qiesi API Toolkit",
    version="1.1.2",
    description="""
A lightweight API toolkit providing:

• JSON → Excel conversion  
• Text → CSV conversion  
• Document/Markdown conversion for AI, workflows, and automation

Designed for Nintex NAC workflow integrations.
"""
)


# -----------------------------
# Request Models
# -----------------------------

class ConvertRequest(BaseModel):
    jsonInput: str


class MessageCSVRequest(BaseModel):
    message: str


class DocumentBase64Request(BaseModel):
    fileName: str
    fileContentBase64: str

SUPPORTED_DOCUMENT_EXTENSIONS = {
    ".pdf",
    ".docx",
    ".xlsx",
    ".pptx",
    ".txt",
    ".csv",
    ".html",
    ".htm",
    ".json",
    ".xml",
    ".md",
    ".jpeg",
    ".jpg",
    ".png",
    ".bmp",
    ".tiff",
}

OCR_DOCUMENT_EXTENSIONS = {".pdf", ".jpeg", ".jpg", ".png", ".bmp", ".tiff"}

DEFAULT_CONTENT_TYPE_MAPPING = {
    ".pdf": "application/pdf",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    ".txt": "text/plain",
    ".csv": "text/csv",
    ".html": "text/html",
    ".htm": "text/html",
    ".json": "application/json",
    ".xml": "application/xml",
    ".md": "text/markdown",
    ".jpeg": "image/jpeg",
    ".jpg": "image/jpeg",
    ".png": "image/png",
    ".bmp": "image/bmp",
    ".tiff": "image/tiff",
}

DOCUMENT_INTELLIGENCE_ENDPOINT_ENV = "AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT"


def _get_file_extension(file_name: str) -> str:
    return os.path.splitext(file_name)[1].lower()


def _get_content_type(file_name: str, extension: str) -> str:
    content_type, _ = mimetypes.guess_type(file_name)
    if content_type:
        return content_type
    return DEFAULT_CONTENT_TYPE_MAPPING.get(extension, "application/octet-stream")


def _ensure_supported_extension(file_name: str, content_type: Optional[str]) -> str:
    extension = _get_file_extension(file_name)
    if not extension and content_type:
        extension = mimetypes.guess_extension(content_type) or ""
    if extension in SUPPORTED_DOCUMENT_EXTENSIONS:
        return extension
    raise HTTPException(
        status_code=415,
        detail=(
            f"Unsupported file type for '{file_name}'. "
            f"Supported extensions: {', '.join(sorted(SUPPORTED_DOCUMENT_EXTENSIONS))}."
        ),
    )


def _decode_base64_file_content(file_content_base64: str) -> bytes:
    try:
        decoded = base64.b64decode(file_content_base64, validate=True)
    except binascii.Error as exc:
        raise HTTPException(status_code=400, detail="Invalid base64 fileContentBase64.") from exc
    if not decoded:
        raise HTTPException(status_code=400, detail="Decoded file content is empty.")
    return decoded


def _convert_document_to_markdown(
    file_bytes: bytes,
    file_name: str,
    content_type: Optional[str],
    use_ocr: bool = False,
) -> dict:
    extension = _ensure_supported_extension(file_name, content_type)
    content_type = _get_content_type(file_name, extension)
    stream_info = StreamInfo(mimetype=content_type, extension=extension, filename=file_name)

    markdowner = MarkItDown(enable_builtins=True)
    ocr_applied = False

    if use_ocr:
        if DocumentIntelligenceConverter is None:
            raise HTTPException(
                status_code=400,
                detail=(
                    "OCR endpoint requires optional dependencies from `markitdown[az-doc-intel]`. "
                    "Install the optional package and configure the Azure Document Intelligence endpoint."
                ),
            )

        endpoint = os.getenv(DOCUMENT_INTELLIGENCE_ENDPOINT_ENV)
        if not endpoint:
            raise HTTPException(
                status_code=400,
                detail=(
                    f"OCR endpoint requires the {DOCUMENT_INTELLIGENCE_ENDPOINT_ENV} environment variable "
                    "to be configured with a valid Azure Document Intelligence endpoint."
                ),
            )

        try:
            doc_intel_converter = DocumentIntelligenceConverter(endpoint=endpoint)
            markdowner.register_converter(doc_intel_converter, priority=PRIORITY_SPECIFIC_FILE_FORMAT)
        except MissingDependencyException as exc:
            raise HTTPException(
                status_code=400,
                detail=(
                    "OCR support requires optional Azure Document Intelligence dependencies. "
                    "Install with `pip install markitdown[az-doc-intel]`."
                ),
            ) from exc

        ocr_applied = extension in OCR_DOCUMENT_EXTENSIONS or (content_type and content_type.startswith("image/"))

    try:
        result = markdowner.convert_stream(BytesIO(file_bytes), stream_info=stream_info)
    except UnsupportedFormatException as exc:
        raise HTTPException(status_code=415, detail=str(exc)) from exc
    except MissingDependencyException as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except FileConversionException as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Document conversion failed: {str(exc)}") from exc

    markdown = result.markdown or ""
    response = {
        "fileName": file_name,
        "contentType": content_type,
        "markdown": markdown,
        "textLength": len(markdown),
        "success": True,
    }

    if use_ocr:
        response["ocrApplied"] = ocr_applied

    return response


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
        "version": "1.1.2",
        "author": "Paul Keys",
        "description": "API toolkit for Nintex workflow integrations",
        "endpoints": {
            "health": "/health",
            "ping": "/ping",
            "json_to_xlsx": "/JSON-to-XLSX",
            "text_to_csv": "/TXT-to-CSV",
            "document_to_markdown": "/documents/markdown",
            "document_to_markdown_ocr": "/documents/markdown/ocr",
            "docs": "/docs",
            "openapi": "/openapi.json"
        }
    }


@app.get("/ping", tags=["System"])
def ping():
    return {
        "message": "pong",
        "api": "Qiesi Toolkit API"
    }


# -----------------------------
# JSON → Excel Conversion
# -----------------------------

@app.post(
    "/JSON-to-XLSX",
    tags=["Conversion"],
    summary="JSON-to-MS Excel",
    description="Accepts JSON and returns it as a Base64 encoded Excel (.xlsx) file."
)
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

        filename = f"QiesiAPI-JSONtoXLSX-{datetime.utcnow().strftime('%Y%m%d%H%M%S')}.xlsx"

        return {
            "fileName": filename,
            "excelFile": encoded
        }

    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post(
    "/documents/markdown",
    tags=["Document Conversion"],
    summary="Document-to-Markdown",
    description="Convert uploaded documents or base64 document payloads into Markdown for AI, automation, and workflow scenarios.",
)
async def document_to_markdown(
    file: Optional[UploadFile] = File(None),
    payload: Optional[DocumentBase64Request] = Body(None),
):
    if file is not None and payload is not None:
        raise HTTPException(
            status_code=400,
            detail="Provide either a multipart file upload or a JSON base64 payload, not both.",
        )

    if file is None and payload is None:
        raise HTTPException(
            status_code=400,
            detail="No file upload or JSON base64 payload provided.",
        )

    if file is not None:
        file_bytes = await file.read()
        if not file_bytes:
            raise HTTPException(status_code=400, detail="Uploaded file is empty.")

        return _convert_document_to_markdown(
            file_bytes=file_bytes,
            file_name=file.filename or "uploaded-document",
            content_type=file.content_type,
        )

    return _convert_document_to_markdown(
        file_bytes=_decode_base64_file_content(payload.fileContentBase64),
        file_name=payload.fileName,
        content_type=None,
    )


@app.post(
    "/documents/markdown/ocr",
    tags=["Document Conversion"],
    summary="Document OCR-to-Markdown",
    description="Convert scanned or image-heavy documents into Markdown using optional Azure Document Intelligence OCR support.",
)
async def document_to_markdown_ocr(
    file: Optional[UploadFile] = File(None),
    payload: Optional[DocumentBase64Request] = Body(None),
):
    if file is not None and payload is not None:
        raise HTTPException(
            status_code=400,
            detail="Provide either a multipart file upload or a JSON base64 payload, not both.",
        )

    if file is None and payload is None:
        raise HTTPException(
            status_code=400,
            detail="No file upload or JSON base64 payload provided.",
        )

    if file is not None:
        file_bytes = await file.read()
        if not file_bytes:
            raise HTTPException(status_code=400, detail="Uploaded file is empty.")

        return _convert_document_to_markdown(
            file_bytes=file_bytes,
            file_name=file.filename or "uploaded-document",
            content_type=file.content_type,
            use_ocr=True,
        )

    return _convert_document_to_markdown(
        file_bytes=_decode_base64_file_content(payload.fileContentBase64),
        file_name=payload.fileName,
        content_type=None,
        use_ocr=True,
    )


# -----------------------------
# Text → CSV Conversion
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

        filename = f"QiesiAPI-TextToCSV-{datetime.utcnow().strftime('%Y%m%d%H%M%S')}.csv"

        return {
            "fileName": filename,
            "csvFile": encoded
        }

    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))