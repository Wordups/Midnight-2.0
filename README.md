[README (3).md](https://github.com/user-attachments/files/26475625/README.3.md)
# MIDNIGHT
### Policy Migration Engine 

> Convert legacy policy documents into audit-ready, template-faithful Word output. No manual reconstruction.

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue?style=flat-square)](https://python.org)
[![Streamlit](https://img.shields.io/badge/Streamlit-1.32%2B-red?style=flat-square)](https://streamlit.io)
[![Groq](https://img.shields.io/badge/Groq-LLaMA%203.3%2070B-orange?style=flat-square)](https://groq.com)
[![License](https://img.shields.io/badge/License-Private-lightgrey?style=flat-square)]()

---

## What it does

Midnight is a document migration and policy creation engine built for compliance and policy operations teams.

**Migrate a Policy** — Upload a legacy `.docx`, `.txt`, or `.md` file. AI extracts every field, section, bullet, table, and revision entry, then rebuilds it into the target template.

**Create a Policy** — Structured intake form with smart date defaults and live preview. Fill in what you know — Midnight handles layout and template fidelity.

---

## How it works

Midnight uses three strictly separated layers. No bleed-through between stages.

```
Layer 1 — Extraction
  Read source document → pull all fields, sections, bullets,
  tables, revision history → output structured POLICY_DATA

Layer 2 — Mapping
  Normalize into fixed schema → validate required fields
  Preserve original wording exactly — no paraphrasing

Layer 3 — Rendering
  Rebuild into target template → handle layout only
  Deterministic output regardless of content length
```

---

## Stack

| Component | Technology |
|-----------|-----------|
| UI | Streamlit |
| AI Extraction | Groq — LLaMA 3.3 70B Versatile |
| Document Output | python-docx |
| Hosting | Streamlit Cloud |

---

## Files

```
midnight/
├── app.py                          # Streamlit UI — three pages
├── hps_policy_migration_builder.py # Strict layout engine — DOCX generation
├── requirements.txt                # Dependencies
└── README.md
```

---

## Requirements

```
streamlit>=1.32.0
groq>=0.4.0
python-docx>=1.1.0
```

---

## Local setup

```bash
git clone https://github.com/wordups/midnight
cd midnight
pip install -r requirements.txt
streamlit run app.py
```

Set your Groq API key in `.streamlit/secrets.toml`:

```toml
GROQ_API_KEY = "gsk_..."
```

---

## Architecture notes

**Builder design decisions:**

- Banner row locked at exact 720-twip height — never expands
- All column widths fixed in `dxa` units — content never shifts layout
- Section content rows flow freely across pages — no orphaning
- Revision history handles `tuple`, `list`, and `dict` entries
- Logo renders from path if valid, falls back to text — never crashes
- All file I/O uses `/tmp` — safe on any host

**Extraction rules enforced by prompt:**

- No summarization or paraphrasing
- Original wording preserved
- All procedure types classified: `para`, `heading`, `bullet`, `sub-bullet`, `bold_intro`, `bold_intro_semi`, `empty`
- Revision history treated as hard requirement — must extract, map, and render

---

## Methodology

Midnight is part of the **Takeoff** governance and compliance platform.

```
Pre-Flight → Boarding → Takeoff → In-Flight → Landing
```

Midnight sits at **Boarding** — getting documentation structured and compliant before the rest of the workflow runs.

---

## Status

Production. 13+ policies migrated in first batch run.

---

*Takeoff — wordups*
