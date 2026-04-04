# MIDNIGHT
### Policy Migration Engine — by Takeoff

> Convert legacy policy documents into audit-ready, template-faithful Word output. No manual reconstruction.

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue?style=flat-square)](https://python.org)
[![FastAPI](https://img.shields.io/badge/FastAPI-Backend-green?style=flat-square)](https://fastapi.tiangolo.com/)
[![Groq](https://img.shields.io/badge/Groq-LLaMA%203.3%2070B-orange?style=flat-square)](https://groq.com)
[![License](https://img.shields.io/badge/License-Private-lightgrey?style=flat-square)]()

---

## What it does

Midnight is a document migration and policy creation engine built for compliance and policy operations teams.

### Migrate a Policy
Upload a legacy `.docx`, `.txt`, or `.md` file.  
AI extracts every field, section, bullet, table, and revision entry, then rebuilds it into the target template.

### Create a Policy
Structured intake → mapped into a fixed schema → rendered into a template-faithful Word document.

---

## Architecture

Midnight uses three strictly separated layers:

```
Frontend (HTML UI)
        ↓
FastAPI Backend
        ↓
Document Engine (Builder)
```

No UI logic inside the backend.  
No extraction logic inside the builder.  
Each layer does one job.

---

## Processing Pipeline

```
Layer 1 — Extraction
  Read source document → extract all content → output POLICY_DATA

Layer 2 — Mapping
  Normalize into fixed schema → validate required fields
  Preserve original wording exactly (no paraphrasing)

Layer 3 — Rendering
  Build final .docx → strict layout engine
  Deterministic output regardless of content length
```

---

## Project Structure

```
Midnight-2.0/
│
├── backend/
│   ├── api.py
│   ├── services.py
│   ├── hps_policy_migration_builder.py
│   ├── requirements.txt
│
├── frontend/
│   ├── index.html
│
├── download/ (optional)
└── README.md
```

---

## Stack

| Layer | Technology |
|------|-----------|
| Frontend | HTML / JavaScript |
| Backend | FastAPI |
| AI Extraction | Groq — LLaMA 3.3 70B |
| Document Output | python-docx |

---

## How it works

```
Upload Policy → Extract → Preview → Generate → Download
```

Everything runs through the browser → API → document engine.

---

## Backend Setup (Local)

```bash
cd backend
pip install -r requirements.txt
set GROQ_API_KEY=your_key
uvicorn api:app --reload
```

Then open:

```
frontend/index.html
```

---

## Key Design Decisions

- Layout engine is deterministic (fixed widths, controlled rows)
- Banner height and structure are locked
- Section content flows across pages without breaking layout
- Revision history is always extracted and rendered
- Extraction preserves original wording — no summarization
- Logo handling is optional and never breaks generation
- Temporary file handling is isolated and safe

---

## Status

Midnight 2.0 — Active Development  
Core migration pipeline operational

---

## Takeoff Platform

Midnight is part of the **Takeoff** compliance system:

```
Pre-Flight → Boarding → Takeoff → In-Flight → Landing
```

Midnight sits at **Boarding** — preparing documentation for execution and audit readiness.

---

*Takeoff — wordups*
