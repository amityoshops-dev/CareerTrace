# CareerTrace v1

Local-first job application intelligence agent using FastAPI + LangGraph + Gemini + SQLite + LangSmith.

## Truthfulness / ATS rule
CareerTrace does not fabricate qualifications. It separates supported JD keywords from missing/unsupported keywords and can suggest JD wording only where the resume already supports the underlying experience.

## Codespaces
```bash
bash bootstrap.sh
source .venv/bin/activate
cp .env.example .env
uvicorn app.main:app --host 0.0.0.0 --port 8000
```

Prefer Codespaces Secrets / environment variables for `GOOGLE_API_KEY` and `LANGSMITH_API_KEY`; never commit `.env`.

Put the resume at `data/resume.docx`.

## Windows
Codespaces cannot directly access the Windows Downloads folder. In a local Windows clone, run `powershell -ExecutionPolicy Bypass -File .\copy_resume.ps1`, or manually copy the resume into `data\resume.docx`.
