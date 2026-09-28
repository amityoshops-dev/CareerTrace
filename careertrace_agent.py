#!/usr/bin/env python3
"""CareerTrace: resume -> Gemini plan -> Crustdata jobs -> truthful match -> pitch -> tracker."""
from __future__ import annotations
import argparse, csv, json, os, re, sys
from datetime import date
from pathlib import Path
from typing import Any, TypedDict
import requests
from dotenv import load_dotenv
from langchain_google_genai import ChatGoogleGenerativeAI
from langgraph.graph import END, START, StateGraph

load_dotenv()
CRUSTDATA_URL = "https://api.crustdata.com/job/search"
CRUSTDATA_VERSION = "2025-11-01"
MODEL = os.getenv("GEMINI_MODEL", "gemini-2.5-flash")


class State(TypedDict, total=False):
    resume_path: str; resume_text: str; role_override: str; location: str
    top_n: int; dry_run: bool; profile: dict; jobs: list; scored: list
    selected: list; tracker_path: str; report_path: str; log: list


def log(state, msg):
    print(f"[CareerTrace] {msg}", flush=True)
    return state.get("log", []) + [msg]


def _content(x) -> str:
    if isinstance(x, list):
        return "".join(p.get("text", "") if isinstance(p, dict) else str(p) for p in x)
    return str(x)


def llm_json(prompt: str) -> Any:
    out = _content(ChatGoogleGenerativeAI(model=MODEL, temperature=0).invoke(prompt).content)
    out = re.sub(r"^```(?:json)?|```$", "", out.strip(), flags=re.M).strip()
    return json.loads(out)


def llm_text(prompt: str) -> str:
    return _content(ChatGoogleGenerativeAI(model=MODEL, temperature=0.3).invoke(prompt).content).strip()


def load_resume(state):
    p = Path(state["resume_path"])
    if not p.exists():
        sys.exit(f"Resume not found: {p}")
    if p.suffix.lower() == ".pdf":
        from pypdf import PdfReader
        text = "\n".join(pg.extract_text() or "" for pg in PdfReader(str(p)).pages)
    else:
        text = p.read_text(errors="ignore")
    if len(text.strip()) < 50:
        sys.exit("Resume text looks empty (scanned PDF?). Use a text PDF or .txt")
    return {"resume_text": text, "log": log(state, f"Loaded resume ({len(text)} chars)")}


def plan_search(state):
    prof = llm_json(f"""Read this resume. Return ONLY JSON:
{{"titles": [up to 4 target job titles, most specific first],
 "skills": [up to 25 technical skills that literally appear in the resume],
 "seniority": "intern|junior|mid|senior"}}
RESUME:
{state['resume_text'][:8000]}""")
    if state.get("role_override"):
        prof["titles"] = [state["role_override"]] + prof.get("titles", [])
    return {"profile": prof, "log": log(state, f"Titles: {prof['titles']} | {len(prof['skills'])} skills")}


def _flatten(i):
    jd = i.get("job_details", {}) or {}
    co = (i.get("company", {}) or {}).get("basic_info", {}) or {}
    loc = i.get("location", {}) or {}
    return {"title": jd.get("title") or i.get("title", ""),
            "company": co.get("name") or i.get("company_name", ""),
            "location": loc.get("raw") or i.get("location", "") or "",
            "url": jd.get("url") or i.get("url", ""),
            "workplace_type": jd.get("workplace_type", "") or "",
            "description": (i.get("content", {}) or {}).get("description") or i.get("description", "") or ""}


def crustdata_search(state):
    if state.get("dry_run"):
        jobs = [{"title": "Data Engineer", "company": "Acme Analytics", "location": "Pune", "url": "https://example.com/1",
                 "workplace_type": "Hybrid", "description": "Python, SQL, Airflow, Spark, AWS required."},
                {"title": "ML Engineer", "company": "Nimbus AI", "location": "Remote", "url": "https://example.com/2",
                 "workplace_type": "Remote", "description": "Python, PyTorch, Docker, Kubernetes, LangChain."}]
        return {"jobs": jobs, "log": log(state, "DRY RUN: sample jobs")}
    key = os.getenv("CRUSTDATA_API_KEY") or sys.exit("CRUSTDATA_API_KEY missing")
    conds = [{"op": "or", "conditions": [{"field": "job_details.title", "type": "(.)", "value": t}
                                        for t in state["profile"]["titles"][:4]]}]
    if state.get("location"):
        conds.append({"field": "location.raw", "type": "(.)", "value": state["location"]})
    body = {"filters": {"op": "and", "conditions": conds},
            "sorts": [{"field": "metadata.date_added", "order": "desc"}], "limit": 50}
    r = requests.post(CRUSTDATA_URL, json=body, timeout=60,
                      headers={"Authorization": f"Bearer {key}", "x-api-version": CRUSTDATA_VERSION})
    if r.status_code != 200:
        sys.exit(f"Crustdata error {r.status_code}: {r.text[:500]}")
    data = r.json()
    raw = data.get("job_listings") or data.get("results") or data.get("jobs") or []
    if not raw:
        print("Crustdata raw response (0 parsed):", json.dumps(data)[:600])
    jobs = [_flatten(x) for x in raw]
    return {"jobs": jobs, "log": log(state, f"Crustdata returned {len(jobs)} jobs")}


def match_jobs(state):
    resume = state["resume_text"].lower()
    scored = []
    for j in state["jobs"]:
        try:
            req = llm_json("List up to 15 technical skills/tools in this job text. "
                           "Return ONLY a JSON array of short strings.\n\n" + (j["description"][:4000] or j["title"]))
        except Exception:
            req = []
        sup, gaps = [], []
        for s in req:
            pat = r"(?<![a-z0-9])" + re.escape(str(s).lower()) + r"(?![a-z0-9])"
            (sup if re.search(pat, resume) else gaps).append(s)
        scored.append({**j, "required": req, "supported": sup, "gaps": gaps,
                       "score": round(100 * len(sup) / len(req)) if req else 0})
    scored.sort(key=lambda x: x["score"], reverse=True)
    return {"scored": scored, "log": log(state, f"Scored {len(scored)} jobs")}


def select_and_tailor(state):
    top = state["scored"][: state.get("top_n", 5)]
    for j in top:
        j["pitch"] = llm_text(f"""Write a 3-sentence application pitch. Mention ONLY these verified skills: {j['supported']}.
Never claim these gaps: {j['gaps']}. No invented employers, numbers or degrees.
Role: {j['title']} at {j['company']}""")
    return {"selected": top, "log": log(state, f"Tailored top {len(top)}")}


def track_and_report(state):
    out = Path("output"); out.mkdir(exist_ok=True)
    t = out / "applications.csv"; new = not t.exists()
    with t.open("a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["date", "company", "title", "location", "score", "url", "status"])
        for j in state["selected"]:
            w.writerow([date.today(), j["company"], j["title"], j["location"], j["score"], j["url"], "To apply"])
    lines = [f"# CareerTrace Report - {date.today()}\n"]
    for i, j in enumerate(state["selected"], 1):
        lines += [f"## {i}. {j['title']} @ {j['company']} ({j['score']}% match)",
                  f"- Location: {j['location']} {j['workplace_type']}", f"- Apply: {j['url']}",
                  f"- Verified skills: {', '.join(j['supported']) or 'none'}",
                  f"- Gaps (do NOT claim): {', '.join(j['gaps']) or 'none'}", f"\n{j['pitch']}\n"]
    rp = out / "report.md"; rp.write_text("\n".join(lines), encoding="utf-8")
    return {"tracker_path": str(t), "report_path": str(rp), "log": log(state, f"Saved {t} and {rp}")}


def build_graph():
    g = StateGraph(State)
    for n, fn in [("load_resume", load_resume), ("plan_search", plan_search), ("crustdata_search", crustdata_search),
                  ("match_jobs", match_jobs), ("select_and_tailor", select_and_tailor),
                  ("track_and_report", track_and_report)]:
        g.add_node(n, fn)
    g.add_edge(START, "load_resume"); g.add_edge("load_resume", "plan_search")
    g.add_edge("plan_search", "crustdata_search")
    g.add_conditional_edges("crustdata_search", lambda s: "match_jobs" if s.get("jobs") else END)
    g.add_edge("match_jobs", "select_and_tailor"); g.add_edge("select_and_tailor", "track_and_report")
    g.add_edge("track_and_report", END)
    return g.compile()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--resume", required=True); ap.add_argument("--role", default="")
    ap.add_argument("--location", default=""); ap.add_argument("--top", type=int, default=5)
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()
    if not os.getenv("GOOGLE_API_KEY"):
        sys.exit("GOOGLE_API_KEY missing (check .env)")
    final = build_graph().invoke({"resume_path": a.resume, "role_override": a.role, "location": a.location,
                                  "top_n": a.top, "dry_run": a.dry_run, "log": []})
    if not final.get("selected"):
        print("No jobs found. Broaden --role or drop --location."); return
    print("\nDONE ->", final["report_path"], "|", final["tracker_path"])


if __name__ == "__main__":
    main()
