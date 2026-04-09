# QIESi API Toolkit

QIESi API Toolkit is a lightweight FastAPI-based service for workflow and automation scenarios. It provides simple file conversions for Nintex, AI, RAG, search indexing, compliance ingestion, and document-processing workflows.

## What it does

- Converts JSON into Excel (`.xlsx`)
- Converts plain text into CSV
- Converts uploaded or base64-supplied documents into Markdown/text using Microsoft MarkItDown
- Optionally supports OCR-style markdown conversion when Azure Document Intelligence is configured

## Installation

### Python version

- Python 3.10+ is recommended

### Setup

1. Create a virtual environment:

```bash
python -m venv .venv
```

2. Activate the virtual environment:

- Windows:

```powershell
.\.venv\Scripts\Activate.ps1
```

- macOS / Linux:

```bash
source .venv/bin/activate
```

3. Install requirements:

```bash
pip install -r requirements.txt
```

### Optional OCR support

To enable Azure Document Intelligence OCR support for `/documents/markdown/ocr`, install MarkItDown with the optional Azure extension:

```bash
pip install markitdown[az-doc-intel]
```

Then configure the environment variable:

```bash
export AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT="https://<your-endpoint>"
export AZURE_API_KEY="<your-key>"
```

On Windows PowerShell:

```powershell
$env:AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT = "https://<your-endpoint>"
$env:AZURE_API_KEY = "<your-key>"
```

## Run locally

```bash
uvicorn main:app --reload --host 0.0.0.0 --port 8000
```

Then open `http://127.0.0.1:8000/docs` for the interactive API docs.

## Endpoints

### GET /health

Purpose: Verify the API is running.

Use cases: health checks from monitoring, load balancers, and workflow sensors.

Sample response:

```json
{
  "status": "ok"
}
```

---

### GET /info

Purpose: Return basic service metadata and endpoint references.

Use cases: discovery, integration validation, lightweight service info.

Sample response:

```json
{
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
```

---

### GET /ping

Purpose: Simple connectivity check.

Use cases: lightweight ping/pong validation for workflow services and tooling.

Sample response:

```json
{
  "message": "pong",
  "api": "Qiesi Toolkit API"
}
```

---

### POST /JSON-to-XLSX

Purpose: Convert a JSON payload into a Base64 encoded Excel `.xlsx` file.

Use cases: export workflow JSON into spreadsheets, integrate Nintex JSON outputs into Excel, archive structured data.

Request format:

```json
{
  "jsonInput": "[{\"name\": \"Alice\", \"email\": \"alice@example.com\"}]"
}
```

Sample curl:

```bash
curl -X POST http://127.0.0.1:8000/JSON-to-XLSX \
  -H "Content-Type: application/json" \
  -d '{"jsonInput":"[{\"name\":\"Alice\",\"email\":\"alice@example.com\"}]"}'
```

Sample response:

```json
{
  "fileName": "QiesiAPI-JSONtoXLSX-20240101120000.xlsx",
  "excelFile": "<base64-encoded-xlsx>"
}
```

Error notes:

- Returns `400` when JSON is invalid or conversion fails.

---

### POST /TXT-to-CSV

Purpose: Convert plain text into a Base64 encoded CSV file.

Use cases: normalize multiline text into CSV rows, process notes or lists for workflow ingestion.

Request format:

```json
{
  "message": "line1\nline2\nline3"
}
```

Sample curl:

```bash
curl -X POST http://127.0.0.1:8000/TXT-to-CSV \
  -H "Content-Type: application/json" \
  -d '{"message":"line1\nline2\nline3"}'
```

Sample response:

```json
{
  "fileName": "QiesiAPI-TextToCSV-20240101120000.csv",
  "csvFile": "<base64-encoded-csv>"
}
```

Error notes:

- Returns `400` when conversion fails.

---

### POST /documents/markdown

Purpose: Convert uploaded or base64-supplied documents into Markdown/text using Microsoft MarkItDown.

Supported file types:

- PDF
- DOCX
- XLSX
- PPTX
- TXT
- CSV
- HTML
- JSON
- XML
- MD
- JPEG/PNG/BMP/TIFF

Use cases:

- Normalize supplier invoices for AI extraction
- Convert contracts to machine-readable text
- Ingest policies and procedures into search or knowledge bases
- Prepare reports, manuals, and meeting packs for summaries
- Convert operational reports for RAG and workflow automation

Request support:

Option 1 — multipart file upload:

```bash
curl -X POST http://127.0.0.1:8000/documents/markdown \
  -F "file=@invoice.pdf"
```

Option 2 — JSON body with base64 payload:

```bash
curl -X POST http://127.0.0.1:8000/documents/markdown \
  -H "Content-Type: application/json" \
  -d '{
    "fileName": "invoice.pdf",
    "fileContentBase64": "<base64-file-content>"
  }'
```

Sample response:

```json
{
  "fileName": "invoice.pdf",
  "contentType": "application/pdf",
  "markdown": "# Converted content...",
  "textLength": 12345,
  "success": true
}
```

Error notes:

- Returns `400` for invalid or empty uploads.
- Returns `400` for missing JSON fields or invalid base64.
- Returns `415` for unsupported file types.

---

### POST /documents/markdown/ocr

Purpose: Convert scanned or image-heavy documents into Markdown using optional Azure Document Intelligence support.

Use cases:

- Scanned invoices and receipts
- Photo-based reports
- PDFs with screenshots or scanned pages
- Low-quality scanned forms and bank reports

Request support:

Option 1 — multipart file upload:

```bash
curl -X POST http://127.0.0.1:8000/documents/markdown/ocr \
  -F "file=@scanned-document.pdf"
```

Option 2 — JSON body with base64 payload:

```bash
curl -X POST http://127.0.0.1:8000/documents/markdown/ocr \
  -H "Content-Type: application/json" \
  -d '{
    "fileName": "scanned-document.pdf",
    "fileContentBase64": "<base64-file-content>"
  }'
```

Sample response:

```json
{
  "fileName": "scanned-invoice.pdf",
  "contentType": "application/pdf",
  "markdown": "# OCR converted content...",
  "textLength": 9876,
  "ocrApplied": true,
  "success": true
}
```

Error notes:

- Returns `400` when OCR dependencies or Azure configuration are missing.
- Returns `415` for unsupported file types.
- If optional OCR support is not installed, the endpoint returns a clear message explaining `markitdown[az-doc-intel]` and environment variables.

## Real business use cases

- Accounts payable invoice normalization for extraction engines
- Contract review preparation for clause analysis and summarization
- Knowledge base ingestion for policy, procedure, and training content
- Compliance and risk document ingestion for searchable archives
- Bank statement, report, and operational data pre-processing
- Nintex / workflow automation file ingestion pipelines
- AI, RAG, and search indexing of business documents

## Testing examples

### JSON → XLSX

```bash
curl -X POST http://127.0.0.1:8000/JSON-to-XLSX \
  -H "Content-Type: application/json" \
  -d '{"jsonInput":"[{\"id\":1,\"name\":\"Widget\"}]"}'
```

### TXT → CSV

```bash
curl -X POST http://127.0.0.1:8000/TXT-to-CSV \
  -H "Content-Type: application/json" \
  -d '{"message":"line1\nline2\nline3"}'
```

### Documents → Markdown

```bash
curl -X POST http://127.0.0.1:8000/documents/markdown \
  -F "file=@example.docx"
```

### Documents → Markdown (base64)

```bash
curl -X POST http://127.0.0.1:8000/documents/markdown \
  -H "Content-Type: application/json" \
  -d '{"fileName":"example.pdf","fileContentBase64":"<base64>"}'
```

### Documents → Markdown OCR

```bash
curl -X POST http://127.0.0.1:8000/documents/markdown/ocr \
  -F "file=@scanned.pdf"
```

## Backward compatibility

All original routes remain supported and unchanged:

- `GET /health`
- `GET /info`
- `GET /ping`
- `POST /JSON-to-XLSX`
- `POST /TXT-to-CSV`

New endpoints were added as extensions only, preserving existing API behavior.
