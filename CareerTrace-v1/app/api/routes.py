from fastapi import APIRouter, HTTPException
from pydantic import BaseModel
from ..database import get_conn
from ..agents.graph import build_graph
from ..services.resume_parser import extract_resume
from ..config import RESUME_PATH

router = APIRouter()

class ApplicationIn(BaseModel):
    company: str
    role: str
    jd: str
    source: str = ""
    recruiter: str = ""
    recruiter_email: str = ""

@router.get("/health")
def health():
    return {"status": "ok", "service": "CareerTrace", "version": "v1"}

@router.get("/resume")
def resume_status():
    return {"configured": RESUME_PATH.exists(), "path": str(RESUME_PATH)}

@router.post("/applications/analyze")
def analyze_application(item: ApplicationIn):
    resume = extract_resume(RESUME_PATH)
    if not resume:
        raise HTTPException(400, "Resume not found. Put it at data/resume.docx or set CAREERTRACE_RESUME_PATH.")
    result = build_graph().invoke({
        "company": item.company, "role": item.role,
        "jd": item.jd, "resume": resume
    })
    if result.get("error"):
        raise HTTPException(500, result["error"])
    a = result["analysis"]
    conn = get_conn()
    cur = conn.execute(
        """INSERT INTO applications
        (company, role, jd, source, recruiter, recruiter_email, status,
         match_summary, ats_keywords, missing_keywords, tailored_summary)
        VALUES (?, ?, ?, ?, ?, ?, 'APPLICATION_READY', ?, ?, ?, ?)""",
        (item.company, item.role, item.jd, item.source, item.recruiter,
         item.recruiter_email, a.get("match_summary",""),
         ", ".join(a.get("supported_keywords", [])),
         ", ".join(a.get("missing_keywords", [])),
         a.get("tailored_summary",""))
    )
    conn.commit()
    app_id = cur.lastrowid
    conn.close()
    return {"application_id": app_id, "analysis": a}

@router.get("/applications")
def applications():
    conn = get_conn()
    rows = [dict(r) for r in conn.execute(
        "SELECT id, company, role, status, match_summary, ats_keywords, missing_keywords, created_at FROM applications ORDER BY id DESC"
    )]
    conn.close()
    return rows
