from typing import TypedDict
import json
import re
from langgraph.graph import StateGraph, END
from langchain_google_genai import ChatGoogleGenerativeAI
from ..config import GOOGLE_API_KEY

class CareerState(TypedDict, total=False):
    company: str
    role: str
    jd: str
    resume: str
    analysis: dict
    error: str

RULE = """You are CareerTrace, a factual job-application assistant.
Never invent employment, projects, achievements, tools, certifications, metrics,
employers, responsibilities, or experience.
You may identify JD keywords already supported by the resume.
You may suggest expressing an existing supported concept using the JD wording.
If a JD keyword is not supported by the resume, put it in missing_keywords.
Do not fabricate ATS keywords.
Return valid JSON only."""

def analyze(state: CareerState):
    if not GOOGLE_API_KEY:
        state["error"] = "GOOGLE_API_KEY is not configured."
        return state
    model = ChatGoogleGenerativeAI(
        model="gemini-2.0-flash",
        google_api_key=GOOGLE_API_KEY,
        temperature=0
    )
    prompt = f"""{RULE}

Resume:
{state.get("resume", "")}

Job description:
{state["jd"]}

Return JSON with exactly these fields:
{{
  "role_type": "string",
  "domain": ["string"],
  "required_keywords": ["string"],
  "supported_keywords": ["string"],
  "missing_keywords": ["string"],
  "match_summary": "factual 2-4 sentence summary",
  "ats_safe_keyword_edits": [
    {{"keyword":"string", "resume_section":"string", "reason":"string"}}
  ],
  "tailored_summary": "truthful summary using only supported experience"
}}
"""
    try:
        result = model.invoke(prompt)
        raw = result.content
        raw = re.sub(r"^```json\s*|^```\s*|\s*```$", "", raw.strip())
        state["analysis"] = json.loads(raw)
    except Exception as e:
        state["error"] = str(e)
    return state

def build_graph():
    graph = StateGraph(CareerState)
    graph.add_node("analyze", analyze)
    graph.set_entry_point("analyze")
    graph.add_edge("analyze", END)
    return graph.compile()
