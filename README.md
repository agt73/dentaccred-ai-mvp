# DentAccred AI — Render-ready MVP

DentAccred AI is an independent dental accreditation-readiness and evidence-review demonstration for predoctoral DDS/DMD programs. It maps uploaded evidence to atomic requirements, applies conservative status rules, records consultant disposition, and produces a printable report.

It is not affiliated with or endorsed by CODA or the American Dental Association and does not make official compliance or accreditation decisions.

## Included functionality

- 12 selected predoctoral demonstration criteria with source metadata and atomic requirements
- PDF and DOCX validation and text extraction with page/section locations
- SCOPE-E-inspired evidence mapping and deterministic readiness categories
- Supporting evidence, gaps, corrective actions, reviewer questions, confidence and limitations
- Prompt-injection detection and expert-review escalation
- Consultant accept/modify/reject/defer controls with required rationale and audit trail
- Four fictional review paths and printable report
- `/health` endpoint and Render Blueprint configuration

## Run locally

```bash
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Open `http://localhost:10000`. No demo login is required.

## Deploy on Render

1. Create a GitHub repository and upload the contents of this folder—not the ZIP itself.
2. In Render, choose **New → Blueprint**, connect the repository, and apply `render.yaml`.
3. Alternatively, create a free Python Web Service with:
   - Build: `pip install -r requirements.txt`
   - Start: `gunicorn app:app`
   - Health check: `/health`
   - Environment variables: generated `SECRET_KEY`; `MAX_UPLOAD_MB=10`
4. Verify `/health`, then run all four fictional review paths.

If the repository files are inside a subfolder, set Render's **Root Directory** to that folder.

## Demonstration limitations

- Rules-assisted fallback, not a full semantic AI accreditation analysis
- No OCR, persistent database, authentication, SSO, production audit storage, or cross-document analysis
- Reviews disappear when the free Render process restarts or sleeps
- DOCX locations are approximate heading/paragraph locations
- Criteria are selected summaries that require qualified source verification before institutional use
- A single document cannot prove program-wide implementation or official compliance

## Data and cost statement

Use fictional or fully de-identified material only. Uploaded bytes are processed in memory and are not intentionally written to disk. No paid service, database, domain, AI subscription, software license, or Render upgrade is required.

Project owner: Mohammad Ali Saghiri, D.Eng., MS, PhD. Specification v2.0, August 2026.
