#!/usr/bin/env bash
# CareerTrace Pro - NON-DESTRUCTIVE add-on. Run from the repo root:  bash upgrade_careertrace.sh
set -e
[ -f careertrace_ui.py ] || { echo "Run this from the CareerTrace repo root (where careertrace_ui.py is)."; exit 1; }
PY=$(command -v python3 || command -v python)
cp -n careertrace_ui.py careertrace_ui.py.bak_pro
mkdir -p data exports pro_static output
$PY -m pip install -q openpyxl python-docx 2>/dev/null || $PY -m pip install -q --break-system-packages openpyxl python-docx 2>&1 | tail -1

# ================= backend module =================
cat > careertrace_pro.py <<'PYEOF'
"""CareerTrace Pro: multi-portal search, resume builder, Excel tracker, signed emails, LinkedIn drafts.
Loaded by careertrace_ui.py via careertrace_pro.install(app, globals()). A failure here never stops the app."""
import os, re, io, csv, json, base64, uuid, ssl, smtplib, urllib.request, urllib.parse, urllib.error, datetime as dt
from concurrent.futures import ThreadPoolExecutor
from email.message import EmailMessage
from pathlib import Path
from typing import List
from urllib.parse import urlparse
from dotenv import load_dotenv
from fastapi import APIRouter, HTTPException
from fastapi.responses import FileResponse, HTMLResponse, Response
from pydantic import BaseModel

load_dotenv()
load_dotenv(".env.pro")
HERE = Path(__file__).parent
STATIC = HERE / "pro_static"
G = {}
STATE = {"conns": [], "msgs": {}}
AG = None
RESUME_OUT = "output/Amit_Mankar_Resume_Updated.docx"
TRACKER = "exports/CareerTrace_Tracker.xlsx"

DEFAULT_PROFILE = {
    "name": "Amit Mankar", "email": "avmankar001@gmail.com", "phone": "9922080307",
    "linkedin": "https://www.linkedin.com/in/amit-mankar5",
    "blog": "https://substack.com/@intentcuriositysphere",
    "alumni_terms": "Kotak, Aditya Birla, University of Mumbai, Amravati",
    "projects": [
        {"name": "TBG CORE // Institutional Transaction Banking Platform", "url": "https://tbg-engine.onrender.com/", "summary": ""},
        {"name": "NPCI Multi-Rail Gateway Cockpit & Architecture Blueprint", "url": "https://npci-b2b-engine.onrender.com/", "summary": ""},
        {"name": "Institutional Fintech Suite - Business Banking Suite", "url": "https://institutional-fintech-suite.onrender.com/", "summary": ""},
    ],
}


def env(k):
    return os.environ.get(k, "")


def profile():
    os.makedirs("data", exist_ok=True)
    if not os.path.exists("data/profile.json"):
        json.dump(DEFAULT_PROFILE, open("data/profile.json", "w"), indent=1)
    P = json.load(open("data/profile.json"))
    P.setdefault("alumni_terms", DEFAULT_PROFILE["alumni_terms"])
    return P


def norm(s):
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9+#]+", " ", (s or "").lower())).strip()


def has(text, kw):
    k = norm(kw)
    return bool(k) and re.search(r"(?<![a-z0-9])" + re.escape(k) + r"(?![a-z0-9])", text) is not None


def split(s):
    return [x.strip() for x in re.split(r"[,\n;]", s or "") if x.strip()]


# ---------------- signature on every email ----------------
def footer():
    P = profile()
    pr = "\n".join(f"  - {p['name']}: {p['url']}" for p in P.get("projects", []) if p.get("url"))
    return (f"{P['name']}\n{P['email']} | {P['phone']}\nLinkedIn: {P['linkedin']}\n"
            f"Blog (Intent Curiosity Sphere): {P['blog']}\n\nSelected projects:\n{pr}")


def sign(text):
    P = profile()
    u = next((p["url"] for p in P.get("projects", []) if p.get("url")), "")
    if u and u in text:
        return text.replace("[Your Name]", P["name"])
    t = text.replace("[Your Name]", "").rstrip()
    return t + ("\n" if t.endswith(",") else "\n\n") + footer()


# ---------------- job sources ----------------
PORTALS = {"linkedin.com": "LinkedIn", "naukri.com": "Naukri", "wellfound.com": "Wellfound", "angel.co": "Wellfound",
           "apna.co": "Apna", "indeed.com": "Indeed", "foundit.in": "Foundit", "monsterindia": "Foundit",
           "glassdoor": "Glassdoor", "instahyre.com": "Instahyre", "cutshort.io": "Cutshort", "shine.com": "Shine",
           "timesjobs.com": "TimesJobs", "iimjobs": "iimjobs", "hirist": "Hirist"}
ALIAS = {"mumbai": ["mumbai", "navi mumbai", "thane", "bombay", "mahape", "andheri", "powai", "bkc", "vashi"],
         "pune": ["pune", "pimpri", "chinchwad", "hinjewadi", "kharadi", "baner"],
         "bengaluru": ["bengaluru", "bangalore"], "gurgaon": ["gurgaon", "gurugram"],
         "delhi": ["delhi", "new delhi", "noida", "ncr"], "hyderabad": ["hyderabad", "secunderabad"]}


class NoKey(Exception):
    pass


def portal_of(url, hint=""):
    h = urlparse(url or "").netloc.lower()
    for k, v in PORTALS.items():
        if k in h:
            return v
    hn = norm(hint)
    for k, v in PORTALS.items():
        if hn and norm(k.split(".")[0]) in hn:
            return v
    return (hint or h.replace("www.", "") or "Other")[:24]


def _http(url, headers=None, data=None):
    req = urllib.request.Request(url, data=data, headers=headers or {"User-Agent": "CareerTrace/2"})
    with urllib.request.urlopen(req, timeout=25) as r:
        return json.loads(r.read().decode("utf-8", "ignore"))


def _job(title, company, location, url, desc, posted, hint, src):
    return {"title": title or "", "company": company or "", "location": location or "", "url": url or "",
            "description": desc or "", "posted": (posted or "")[:10], "portal": portal_of(url, hint),
            "workplace_type": "", "source": src}


def p_jsearch(q, loc, cfg):
    k = env("RAPIDAPI_KEY")
    if not k:
        raise NoKey("not configured - add RAPIDAPI_KEY to .env.pro")
    d = cfg.get("days") or 0
    dp = "all" if not d else "today" if d <= 1 else "3days" if d <= 3 else "week" if d <= 7 else "month"
    qs = urllib.parse.urlencode({"query": f"{q} in {loc}", "page": 1, "num_pages": 2, "country": "in", "date_posted": dp})
    data = _http("https://jsearch.p.rapidapi.com/search?" + qs,
                 {"X-RapidAPI-Key": k, "X-RapidAPI-Host": "jsearch.p.rapidapi.com"})
    return [_job(x.get("job_title"), x.get("employer_name"),
                 ", ".join(filter(None, [x.get("job_city"), x.get("job_state"), x.get("job_country")])),
                 x.get("job_apply_link"), x.get("job_description"), x.get("job_posted_at_datetime_utc"),
                 x.get("job_publisher"), "JSearch") for x in data.get("data", [])]


def p_adzuna(q, loc, cfg):
    i, k = env("ADZUNA_APP_ID"), env("ADZUNA_APP_KEY")
    if not (i and k):
        raise NoKey("not configured - add ADZUNA_APP_ID and ADZUNA_APP_KEY to .env.pro")
    p = {"app_id": i, "app_key": k, "results_per_page": 30, "what": q, "content-type": "application/json"}
    if norm(loc) not in ("india", "all india", ""):
        p["where"] = loc
    if cfg.get("days"):
        p["max_days_old"] = cfg["days"]
    data = _http("https://api.adzuna.com/v1/api/jobs/in/search/1?" + urllib.parse.urlencode(p))
    return [_job(x.get("title"), (x.get("company") or {}).get("display_name"),
                 (x.get("location") or {}).get("display_name"), x.get("redirect_url"),
                 x.get("description"), x.get("created"), "", "Adzuna") for x in data.get("results", [])]


def p_jooble(q, loc, cfg):
    k = env("JOOBLE_KEY")
    if not k:
        raise NoKey("not configured - add JOOBLE_KEY to .env.pro")
    body = json.dumps({"keywords": q, "location": "" if norm(loc) in ("india", "all india") else loc, "page": 1}).encode()
    data = _http("https://jooble.org/api/" + k, {"Content-Type": "application/json"}, body)
    return [_job(x.get("title"), x.get("company"), x.get("location"), x.get("link"),
                 re.sub("<[^>]+>", " ", x.get("snippet") or ""), x.get("updated"), x.get("source"), "Jooble")
            for x in data.get("jobs", [])]


def crust_all(titles, locs, kws, limit):
    """Your existing Crustdata search, run once per location so 'Pune, Mumbai' works."""
    out, errs = [], []
    for l in locs:
        try:
            jobs, _ = G["crust"](titles, l, kws, limit)
            out += jobs
        except HTTPException as e:
            if e.status_code == 400 and "CRUSTDATA_API_KEY" in str(e.detail):
                raise NoKey("not configured - CRUSTDATA_API_KEY missing in .env")
            errs.append(f"{l}: {str(e.detail)[:100]}")
    if errs and not out:
        raise RuntimeError("; ".join(errs))
    for j in out:
        j["portal"] = j.get("portal") or portal_of(j.get("url", ""), "LinkedIn")
    return out


SAMPLES = [
    _job("Cash Management Sales - Transaction Banking", "Sample Bank", "Mumbai, Maharashtra", "https://www.linkedin.com/jobs/view/1",
         "Own corporate cash management and collections portfolio, payments, virtual accounts, H2H. " * 4, dt.date.today().isoformat(), "", "sample"),
    _job("Relationship Manager - Cash Management", "Sample Bank 2", "Bengaluru", "https://wellfound.com/jobs/3",
         "Manage client relationships for cash management services and payments. " * 5, "", "", "sample"),
    _job("Software Engineer - Java", "Sample Tech", "Pune", "https://apna.co/job/4", "Build microservices in Java. " * 15, "", "", "sample"),
    _job("Transaction Banking Analyst", "Sample Bank 3", "Pune", "https://in.indeed.com/viewjob?jk=5",
         "Analyse collections and payments data for corporate clients; cash management reporting. " * 5, "", "", "sample"),
]


DEAD = ("no longer accepting applications", "no longer available", "job has expired", "this job has expired", "job expired",
        "position has been filled", "position is no longer", "job is closed", "job has been closed", "posting has expired",
        "no longer open", "job not found", "page not found", "vacancy has been closed", "this job is no longer",
        "listing has expired", "job is no longer", "vacancy is closed", "expired job")


def check_live(url):
    """live = page loads with no closed notice; closed = 404/410 or 'no longer available' text; unknown = blocked/unreachable."""
    if not url or "example.com" in url:
        return "unknown"
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36",
                                                   "Accept-Language": "en"})
        with urllib.request.urlopen(req, timeout=8) as r:
            html = r.read(200000).decode("utf-8", "ignore").lower()
        return "closed" if any(x in html for x in DEAD) else "live"
    except urllib.error.HTTPError as e:
        return "closed" if e.code in (404, 410) else "unknown"
    except Exception:
        return "unknown"


def loc_ok(jloc, locs):
    ln = [norm(l) for l in locs]
    if not ln or any(l in ("india", "all india", "anywhere", "remote") for l in ln):
        return True
    j = norm(jloc)
    if not j or "remote" in j:
        return True
    return any(any(a in j for a in ALIAS.get(l, [l])) for l in ln)


def score_job(j, cfg, dropped):
    titles, anyk = split(cfg["titles"]), split(cfg["keywords"])
    must, excl = split(cfg["must"]), split(cfg["exclude"])
    locs = split(cfg["location"]) or ["India"]
    desc = j.get("description") or ""
    ttl = norm(j.get("title"))
    text = norm((j.get("title") or "") + " " + desc)
    full = len(desc) >= 300
    if any(has(text, e) for e in excl):
        dropped["excluded"] += 1
        return None
    if cfg["strict_location"] and not loc_ok(j.get("location", ""), locs):
        dropped["location"] += 1
        return None
    days = cfg["days"] or (60 if cfg.get("only_live", True) else 0)
    if days and j.get("posted"):
        try:
            if (dt.date.today() - dt.date.fromisoformat(j["posted"])).days > days:
                dropped["too_old"] += 1
                return None
        except ValueError:
            pass
    mh = [k for k in must if has(text, k)]
    ah = [k for k in anyk if has(text, k)]
    if must and len(mh) < len(must) and full:
        dropped["missing_must_have"] += 1
        return None
    if anyk and not ah and full:
        dropped["no_keyword_match"] += 1
        return None
    ts = 1.0 if not titles else 0.0
    for t in titles:
        w = [x for x in norm(t).split() if len(x) > 2] or norm(t).split()
        if w:
            ts = max(ts, 1.0 if has(ttl, t) else 0.8 * sum(has(ttl, x) for x in w) / len(w))
    kws = must + anyk
    matched = mh + [k for k in ah if k not in mh]
    cov = len(matched) / len(kws) if kws else 1.0
    sc = round(50 * ts + 40 * cov + (10 if j.get("location") else 5))
    return dict(j, portal=j.get("portal") or portal_of(j.get("url", "")), kw_score=sc, score=sc,
                matched=matched, partial=not full, supported=[], gaps=[], resume_pct=None)


def resume_text(path=""):
    try:
        return G["read_resume"](path or (RESUME_OUT if os.path.exists(RESUME_OUT) else G["find_resume"]()))
    except Exception:
        return ""


def resume_match(j, rtext):
    try:
        req = AG.llm_json("List up to 12 key skills, domain terms, products or qualifications required in this job text. "
                          "Return ONLY a JSON array of short strings.\n\n" + ((j.get("description") or "")[:4000] or j["title"]))
    except Exception:
        req = []
    if req and rtext:
        sup, gaps, pct = G["match"](rtext, req)
        j.update(supported=sup, gaps=gaps, resume_pct=pct, score=round(0.6 * j["kw_score"] + 0.4 * pct))


def portal_links(titles, locs):
    out = []
    for t in titles[:5]:
        for l in locs[:3]:
            q, lq = urllib.parse.quote_plus(t), urllib.parse.quote_plus(l)
            slug = re.sub(r"[^a-z0-9]+", "-", t.lower()).strip("-")
            ls = re.sub(r"[^a-z0-9]+", "-", l.lower()).strip("-")
            india = norm(l) in ("india", "all india")
            g = lambda d: "https://www.google.com/search?q=" + urllib.parse.quote_plus(f'site:{d} "{t}" {l}')
            out.append({"title": t, "location": l, "links": {
                "LinkedIn": f"https://www.linkedin.com/jobs/search/?keywords={q}&location={lq}",
                "Naukri": f"https://www.naukri.com/{slug}-jobs" + ("" if india else f"-in-{ls}"),
                "Indeed": f"https://in.indeed.com/jobs?q={q}&l={lq}",
                "Foundit": f"https://www.foundit.in/srp/results?query={q}&locations={lq}",
                "Wellfound": g("wellfound.com/jobs"), "Apna": g("apna.co"),
                "Instahyre": g("instahyre.com"), "Cutshort": g("cutshort.io")}})
    return out


def run_search(cfg):
    titles = split(cfg["titles"])
    locs = split(cfg["location"]) or ["India"]
    kws = split(cfg["keywords"]) + split(cfg["must"])
    status, raw = {}, []
    if cfg["dry_run"]:
        raw = list(G.get("DEMO", [])) + SAMPLES
        status["Sample data"] = {"ok": True, "count": len(raw), "error": ""}
    else:
        pairs = [(t, l) for t in titles for l in locs][:cfg["max_queries"]]
        calls = [("Crustdata (your existing source)", lambda: crust_all(titles, locs, kws, min(cfg["max_jobs"] * 2, 50)))]
        for nm, fn in (("JSearch (LinkedIn/Naukri/Indeed/Glassdoor/Apna via Google Jobs)", p_jsearch),
                       ("Adzuna", p_adzuna), ("Jooble", p_jooble)):
            calls.append((nm, (lambda fn=fn: [x for q, l in pairs for x in fn(q, l, cfg)])))

        def run(c):
            try:
                return c[0], c[1](), ""
            except NoKey as e:
                return c[0], [], str(e)
            except Exception as e:
                return c[0], [], f"{type(e).__name__}: {str(e)[:140]}"

        with ThreadPoolExecutor(4) as ex:
            for nm, r, err in ex.map(run, calls):
                status[nm] = {"ok": not err, "count": len(r), "error": err}
                raw += r
    dropped = {k: 0 for k in ("excluded", "location", "too_old", "closed", "missing_must_have", "no_keyword_match", "low_score")}
    best = {}
    for j in raw:
        sj = score_job(j, cfg, dropped)
        if not sj:
            continue
        key = norm(sj["title"]) + "|" + norm(sj["company"])
        if key not in best or len(sj.get("description", "")) > len(best[key].get("description", "")):
            best[key] = sj
    only_live = cfg.get("only_live", True)
    jobs = sorted(best.values(), key=lambda x: -x["score"])[: cfg["max_jobs"] * (2 if only_live else 1)]
    for j in jobs:
        j["live"] = "unknown"
    if only_live and not cfg["dry_run"]:
        with ThreadPoolExecutor(8) as ex:
            for j, st in zip(jobs, ex.map(lambda j: check_live(j.get("url", "")), jobs)):
                j["live"] = st
        n0 = len(jobs)
        jobs = [j for j in jobs if j["live"] != "closed"]
        dropped["closed"] += n0 - len(jobs)
    jobs = jobs[: cfg["max_jobs"]]
    rpath = cfg.get("resume_path") or (RESUME_OUT if os.path.exists(RESUME_OUT) else G["find_resume"]())
    rt = resume_text(rpath)
    with ThreadPoolExecutor(4) as ex:
        list(ex.map(lambda j: resume_match(j, rt), jobs[:12]))
    keep = []
    for j in jobs:
        if j["score"] < cfg["min_score"]:
            dropped["low_score"] += 1
        else:
            keep.append(j)
    keep.sort(key=lambda x: -x["score"])
    terms = split(profile().get("alumni_terms", ""))
    for j in keep:
        j["description"] = (j.get("description") or "")[:500]
        ms = find_matches(STATE["conns"], j.get("company", ""), j.get("title", ""), terms)
        j["referrals"] = [{"person": f"{p['first']} {p['last']}".strip(), "position": p["position"], "company": p["company"],
                           "profile": p["url"], "why": p["why"]} for p in ms]
        cq = urllib.parse.quote_plus(j.get("company", ""))
        rq = urllib.parse.quote_plus(j.get("company", "") + " recruiter")
        net = "&network=%5B%22F%22%2C%22S%22%5D&origin=FACETED_SEARCH"
        j["li_links"] = {"My connections + 2nd-degree at company": f"https://www.linkedin.com/search/results/people/?keywords={cq}{net}",
                         "Recruiters at company": f"https://www.linkedin.com/search/results/people/?keywords={rq}{net}"}
    note = "" if any(s["ok"] for s in status.values()) or cfg["dry_run"] else \
        "No job source responded. Check keys, or use the portal links below."
    return {"jobs": keep, "providers": status, "dropped": dropped, "raw_count": len(raw), "note": note,
            "portal_links": portal_links(titles, locs), "resume": rpath, "resume_read": bool(rt)}


# ---------------- Excel tracker ----------------
STATUSES = ["Saved", "Applied", "Referral Requested", "Recruiter Contacted", "Interview Scheduled",
            "Interview Done", "Offer", "Rejected", "On Hold", "Withdrawn"]
APP = ["Application ID", "Date Added", "Job Title", "Company", "Location", "Portal", "Job URL", "Status", "Date Applied",
       "Match Score", "Matched Keywords", "Keyword Gaps (not in resume)", "Contact Name", "Contact Email",
       "Follow-up Date", "Notes", "Last Updated"]
LOG = ["Timestamp", "Application ID", "Company", "Job Title", "Old Status", "New Status", "Note"]
OUT = ["Timestamp", "Application ID", "Company", "Person", "Relationship", "Channel", "Message", "Status"]


def _now():
    return dt.datetime.now().strftime("%Y-%m-%d %H:%M")


def _wb():
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.worksheet.datavalidation import DataValidation
    if os.path.exists(TRACKER):
        return load_workbook(TRACKER)
    os.makedirs("exports", exist_ok=True)
    wb = Workbook()
    wb.remove(wb.active)

    def sh(name, cols, w=None):
        ws = wb.create_sheet(name)
        ws.append(cols)
        ws.freeze_panes = "A2"
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F3864")
            ws.column_dimensions[c.column_letter].width = (w or {}).get(c.value, 18)
        return ws

    a = sh("Applications", APP, {"Job Title": 32, "Company": 24, "Job URL": 40, "Notes": 40,
                                 "Matched Keywords": 30, "Keyword Gaps (not in resume)": 30})
    dv = DataValidation(type="list", formula1='"%s"' % ",".join(STATUSES))
    a.add_data_validation(dv)
    dv.add("H2:H2000")
    sh("Status Log", LOG, {"Job Title": 32, "Company": 24, "Note": 40})
    sh("Outreach", OUT, {"Message": 80, "Person": 24, "Company": 24})
    d = wb.create_sheet("Dashboard", 0)
    d["A1"] = "CareerTrace Dashboard"
    d["A1"].font = Font(bold=True, size=16, color="1F3864")
    for i, s in enumerate(STATUSES, 3):
        d.cell(i, 1, s)
        d.cell(i, 2, f"=COUNTIF(Applications!H:H,A{i})")
    d.cell(len(STATUSES) + 3, 1, "Total")
    d.cell(len(STATUSES) + 3, 2, "=COUNTA(Applications!A:A)-1")
    d.column_dimensions["A"].width = 26
    wb.save(TRACKER)
    return wb


def list_apps():
    ws = _wb()["Applications"]
    return [dict(zip(APP, ["" if v is None else str(v) for v in r])) for r in ws.iter_rows(min_row=2, values_only=True) if r[0]]


def add_app(job, status="Saved"):
    for r in list_apps():
        if (job.get("url") and r["Job URL"] == job["url"]) or \
                norm(r["Job Title"]) + norm(r["Company"]) == norm(job.get("title", "")) + norm(job.get("company", "")):
            return r
    wb = _wb()
    aid = "CT-" + dt.date.today().strftime("%y%m%d") + "-" + uuid.uuid4().hex[:4].upper()
    row = {"Application ID": aid, "Date Added": _now(), "Job Title": job.get("title", ""), "Company": job.get("company", ""),
           "Location": job.get("location", ""), "Portal": job.get("portal", ""), "Job URL": job.get("url", ""),
           "Status": status, "Date Applied": _now() if status == "Applied" else "", "Match Score": job.get("score", ""),
           "Matched Keywords": ", ".join(job.get("matched", [])),
           "Keyword Gaps (not in resume)": ", ".join(job.get("gaps", [])), "Last Updated": _now()}
    wb["Applications"].append([row.get(c, "") for c in APP])
    wb["Status Log"].append([_now(), aid, row["Company"], row["Job Title"], "", status, "Created"])
    wb.save(TRACKER)
    return row


def set_status(aid, status, note="", **fields):
    wb = _wb()
    ws = wb["Applications"]
    h = {c.value: i for i, c in enumerate(ws[1])}
    for r in ws.iter_rows(min_row=2):
        if r[0].value == aid:
            old = r[h["Status"]].value
            r[h["Status"]].value = status
            r[h["Last Updated"]].value = _now()
            if status == "Applied" and not r[h["Date Applied"]].value:
                r[h["Date Applied"]].value = _now()
            if note:
                r[h["Notes"]].value = ((r[h["Notes"]].value or "") + " | " + note).strip(" |")
            for k, v in fields.items():
                if k in h and v:
                    r[h[k]].value = v
            wb["Status Log"].append([_now(), aid, r[h["Company"]].value, r[h["Job Title"]].value, old, status, note])
            wb.save(TRACKER)
            return True
    return False


def log_outreach(aid, company, person, msg):
    wb = _wb()
    wb["Outreach"].append([_now(), aid, company, person, "1st-degree", "LinkedIn", msg, "Drafted"])
    wb.save(TRACKER)


# ---------------- LinkedIn network ----------------
def parse_connections(t):
    lines = t.lstrip("\ufeff").splitlines()
    i = next((k for k, l in enumerate(lines) if l.lower().startswith("first name")), 0)
    return [{"first": r.get("First Name", "").strip(), "last": r.get("Last Name", "").strip(), "url": r.get("URL", "").strip(),
             "company": r.get("Company", "").strip(), "position": r.get("Position", "").strip(),
             "connected_on": r.get("Connected On", "").strip()} for r in csv.DictReader(io.StringIO("\n".join(lines[i:])))]


def parse_messages(t, me="amit-mankar5"):
    out = {}
    for r in csv.DictReader(io.StringIO(t.lstrip("\ufeff"))):
        r = {(k or "").upper(): v for k, v in r.items()}
        d = (r.get("DATE") or "")[:10]
        blob = (r.get("SENDER PROFILE URL") or "") + " " + (r.get("RECIPIENT PROFILE URLS") or "")
        for u in re.findall(r"https?://[^\s,;\"]*linkedin\.com/in/[^\s,;\"/]+", blob):
            u = u.lower().rstrip("/")
            if me not in u and d > out.get(u, ""):
                out[u] = d
    return out


def _fmt(d):
    try:
        return dt.datetime.strptime(d[:10], "%Y-%m-%d").strftime("%b %Y")
    except Exception:
        return ""


HR = ("recruit", "talent", "hiring", "human resources", " hr", "hr ", "acquisition", "people")


def find_matches(conns, company, title="", terms=None):
    c = norm(company)
    kw = [w for w in norm(title).split() if len(w) > 3]
    terms = [norm(t) for t in (terms or []) if norm(t)]
    a, al, b, d = [], [], [], []
    for p in conns:
        pc, pos = norm(p["company"]), norm(p["position"])
        if c and pc and (c in pc or pc in c):
            a.append(dict(p, why="Works at " + company))
        elif any(t in pc or t in pos for t in terms):
            al.append(dict(p, why="Alumni / ex-colleague network"))
        elif any(t in " " + pos + " " for t in HR):
            b.append(dict(p, why="Recruiter / HR"))
        elif kw and sum(w in pos for w in kw) >= 2:
            d.append(dict(p, why="Similar role"))
    return a[:5] + al[:3] + b[:3] + d[:2]


def li_draft(p, job, hist, prior):
    P = profile()
    first = p["first"] or "there"
    co = job.get("company", "the company")
    title = job.get("title", "this role")
    y = re.search(r"\d{4}", p.get("connected_on", ""))
    sup = ", ".join((job.get("supported") or job.get("matched") or [])[:4]) or "cash management and transaction banking"
    at_co = bool(norm(p["company"]) and norm(co) and (norm(co) in norm(p["company"]) or norm(p["company"]) in norm(co)))
    proj = next((x["url"] for x in P.get("projects", []) if x.get("url")), "")
    facts = (f"They work as {p['position'] or 'a professional'} at {p['company'] or 'their company'}. "
             + (f"Connected since {y.group()}. " if y else "") + (f"Last conversation with them: {prior}. " if prior else ""))
    prompt = (f"Write a LinkedIn direct message, max 75 words, from {P['name']} to {first}. Use ONLY these facts: {facts}"
              f"Target role: {title} at {co}. Candidate's verified skills for this role: {sup}. "
              f"Candidate's note: {hist or 'none'}. Goal: politely ask whether they can refer him or help him get this role in any way "
              f"(a referral, an intro to the hiring manager, or advice). "
              f"{'Mention that they work at the same company.' if at_co else 'They may not work at the company; ask for an intro or advice.'} "
              f"Warm, specific, not pushy. Plain text, no subject line, sign with first name only. Invent nothing.")
    try:
        t = AG.llm_text(prompt).strip()
    except Exception:
        t = (f"Hi {first}, hope you're doing well."
             + (f" It's been a while since we last spoke ({prior})." if prior else (f" We've been connected since {y.group()}." if y else ""))
             + (" " + hist.strip() if hist.strip() else "")
             + (f" I saw that {co} has an opening for {title}, which lines up with my background in {sup}. "
                f"Since you're there, would you be open to referring me or sharing how the team hires?" if at_co else
                f" I'm exploring {title} at {co} and thought you might know the right person to speak with. "
                f"Could you introduce me or share any advice? My background: {sup}.")
             + " Thank you!\n\n" + P["name"].split()[0])
    return t + (f"\n\nMy work: {proj}" if proj else "")


# ---------------- resume builder ----------------
def build_resume():
    from docx import Document
    from docx.shared import Pt
    from docx.oxml import OxmlElement
    from docx.text.paragraph import Paragraph
    base = G["find_resume"]()
    if not base or not base.lower().endswith(".docx"):
        raise HTTPException(400, "No .docx resume found. Upload your resume (.docx) in the Resume tab first.")
    P = profile()
    d = Document(base)
    pars = d.paragraphs
    full = "\n".join(p.text for p in pars).lower()
    changes = []
    blog = P["blog"].replace("https://", "").replace("http://", "")
    for p in pars:
        t = p.text.lower()
        if "@" in t and "linkedin" in t and p.runs:
            last = p.runs[-1]
            add = []
            li = P["linkedin"].replace("https://", "").replace("http://", "").replace("www.", "")
            if P["email"].lower() not in t:
                add.append(P["email"])
            if li.lower() not in t.replace("www.", ""):
                add.append(li)

            def clone(text, brk=False, p=p, last=last):
                r = p.add_run()
                if brk:
                    r.add_break()
                r.add_text(text)
                r.font.size = last.font.size
                r.font.name = last.font.name
                if last.font.color is not None and last.font.color.rgb is not None:
                    r.font.color.rgb = last.font.color.rgb

            if add:
                clone(" | " + " | ".join(add))
                changes.append("Added missing email/LinkedIn to header")
            if blog.lower() not in t:
                clone("Blog: " + blog, brk=True)
                changes.append("Added blog link to header")
            break
    host = lambda u: re.sub(r"^https?://|/.*$", "", u.lower())
    missing = [x for x in P.get("projects", []) if x.get("url") and host(x["url"]) not in full]

    def after(par, text, bold=False, italic=False, size=10.5, before=4):
        e = OxmlElement("w:p")
        par._p.addnext(e)
        n = Paragraph(e, par._parent)
        r = n.add_run(text)
        r.bold = bold
        r.italic = italic
        r.font.size = Pt(size)
        r.font.name = "Calibri"
        n.paragraph_format.space_before = Pt(before)
        n.paragraph_format.space_after = Pt(1)
        return n

    if missing:
        head = next((p for p in pars if p.text.strip().upper().startswith(
            ("TRANSACTION BANKING PROTOTYPES", "PROJECTS", "TRANSACTION BANKING PROJECTS"))), None)
        if head is None:
            head = after(pars[-1], "PROJECTS", bold=True, size=11, before=10)
            changes.append("Created PROJECTS section")
        blocks = []
        for x in missing:
            blocks += [(x["name"] + " | Live", dict(bold=True)), ("Link: " + x["url"], dict(italic=True, before=1))]
            if x.get("summary"):
                blocks.append((x["summary"], dict(size=9.5, before=1)))
        for text, kw in reversed(blocks):
            after(head, text, **kw)
        changes.append("Added %d project(s) from your profile" % len(missing))
    try:
        for rel in d.part.rels.values():
            if rel.is_external and "google.com/search?q=" in rel.target_ref:
                rel._target = rel.target_ref.split("q=", 1)[1]
                changes.append("Fixed redirect link")
    except Exception:
        pass
    os.makedirs("output", exist_ok=True)
    d.save(RESUME_OUT)
    return {"path": RESUME_OUT, "base": base, "changes": changes or ["Resume already contains all profile details"]}


# ---------------- API ----------------
class SearchReq(BaseModel):
    titles: str = ""
    location: str = "India"
    keywords: str = ""
    must: str = ""
    exclude: str = ""
    days: int = 0
    strict_location: bool = True
    min_score: int = 20
    max_jobs: int = 25
    max_queries: int = 10
    dry_run: bool = False
    only_live: bool = True
    resume_path: str = ""


class Job(BaseModel):
    title: str
    company: str = ""
    location: str = ""
    url: str = ""
    portal: str = ""
    score: int = 0
    description: str = ""
    matched: List[str] = []
    gaps: List[str] = []
    supported: List[str] = []


class DraftReq(BaseModel):
    job: Job
    recruiter: str = "Hiring Team"
    history: str = ""
    extra: str = ""


class StatusReq(BaseModel):
    id: str
    status: str
    note: str = ""


class SendReq(BaseModel):
    to: str
    subject: str
    body: str
    app_id: str = ""


class LogReq(BaseModel):
    id: str = ""
    company: str = ""
    person: str
    message: str


class Csv(BaseModel):
    csv: str


class B64(BaseModel):
    b64: str


def _load(f, d):
    try:
        return json.load(open(f))
    except Exception:
        return d


def install(app, g):
    global AG
    import careertrace_agent as ag
    AG = ag
    if not getattr(g["make_email"], "_pro", False):
        orig = g["make_email"]

        def signed(*a, **k):
            return sign(orig(*a, **k))
        signed._pro = True
        g["make_email"] = signed
    G.update(g)
    profile()
    r = APIRouter()
    state = STATE
    state["conns"] = _load("data/connections.json", [])
    state["msgs"] = _load("data/messages.json", {})

    @r.get("/pro", response_class=HTMLResponse)
    def page():
        return (STATIC / "pro.html").read_text(encoding="utf8")

    @r.get("/pro/app.js")
    def js():
        return Response((STATIC / "app.js").read_text(encoding="utf8"), media_type="application/javascript")

    @r.get("/api/pro/health")
    def health():
        return {"resume": G["find_resume"](), "resume_built": os.path.exists(RESUME_OUT),
                "connections": len(state["conns"]), "messages": len(state["msgs"]),
                "keys": {"Crustdata": bool(env("CRUSTDATA_API_KEY")), "JSearch": bool(env("RAPIDAPI_KEY")),
                         "Adzuna": bool(env("ADZUNA_APP_ID") and env("ADZUNA_APP_KEY")),
                         "Jooble": bool(env("JOOBLE_KEY")), "Gmail send": bool(env("GMAIL_APP_PASSWORD"))}}

    @r.post("/api/pro/search")
    def search(q: SearchReq):
        return run_search(q.model_dump())

    @r.get("/api/pro/statuses")
    def statuses():
        return STATUSES

    @r.get("/api/pro/apps")
    def apps():
        return list_apps()

    @r.post("/api/pro/apps")
    def add(j: Job, status: str = "Saved"):
        row = add_app(j.model_dump(), status)
        if status != "Saved" and row["Status"] != status:
            set_status(row["Application ID"], status)
        return row

    @r.post("/api/pro/status")
    def set_st(s: StatusReq):
        return {"ok": set_status(s.id, s.status, s.note)}

    @r.post("/api/pro/email")
    def email(d: DraftReq):
        row = add_app(d.job.model_dump(), "Saved")
        P = profile()
        j = d.job
        try:
            txt = G["make_email"](j.title, j.company, j.location, j.supported or j.matched, j.gaps, d.extra)
        except Exception:
            txt = sign(f"Subject: Application: {j.title} - {P['name']}\n\nDear Hiring Team,\n\nI am writing to apply for the "
                       f"{j.title} role at {j.company}. My resume is attached and I would welcome a conversation.\n\nRegards,\n[Your Name]")
        m = re.match(r"\s*Subject:\s*(.*?)\n+(.*)", txt, re.S)
        subj, body = (m.group(1).strip(), m.group(2).strip()) if m else (f"Application: {j.title} - {P['name']}", txt)
        if d.recruiter and d.recruiter != "Hiring Team":
            body = body.replace("Hiring Team", d.recruiter, 1)
        return {"subject": subj, "body": body, "app_id": row["Application ID"]}

    @r.post("/api/pro/email/send")
    def send(s: SendReq):
        pw = env("GMAIL_APP_PASSWORD")
        P = profile()
        if not pw:
            raise HTTPException(400, "Add GMAIL_APP_PASSWORD (a Google App Password) to .env.pro to send from the app, or use 'Open in Gmail'.")
        m = EmailMessage()
        m["From"], m["To"], m["Subject"] = P["email"], s.to, s.subject
        m.set_content(s.body)
        att = RESUME_OUT if os.path.exists(RESUME_OUT) else G["find_resume"]()
        if att and att.lower().endswith(".docx"):
            m.add_attachment(open(att, "rb").read(), maintype="application",
                             subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
                             filename="Amit_Mankar_Resume.docx")
        try:
            with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=ssl.create_default_context()) as sm:
                sm.login(P["email"], pw)
                sm.send_message(m)
        except Exception as e:
            raise HTTPException(400, f"Send failed: {e}")
        if s.app_id:
            set_status(s.app_id, "Recruiter Contacted", "Email sent to " + s.to, **{"Contact Email": s.to})
        return {"ok": True}

    @r.post("/api/pro/connections")
    def conns(c: Csv):
        state["conns"] = parse_connections(c.csv)
        json.dump(state["conns"], open("data/connections.json", "w"))
        return {"count": len(state["conns"])}

    @r.post("/api/pro/messages")
    def msgs(c: Csv):
        state["msgs"] = parse_messages(c.csv)
        json.dump(state["msgs"], open("data/messages.json", "w"))
        return {"count": len(state["msgs"])}

    @r.post("/api/pro/outreach")
    def outreach(d: DraftReq):
        ms = find_matches(state["conns"], d.job.company, d.job.title, split(profile().get("alumni_terms", "")))
        jd = d.job.model_dump()

        def one(p):
            prior = _fmt(state["msgs"].get((p.get("url") or "").lower().rstrip("/"), ""))
            return {"person": f"{p['first']} {p['last']}".strip(), "position": p["position"], "company": p["company"],
                    "profile": p["url"], "prior": prior, "why": p.get("why", ""), "message": li_draft(p, jd, d.history, prior)}

        with ThreadPoolExecutor(4) as ex:
            res = list(ex.map(one, ms))
        return {"connections_loaded": len(state["conns"]), "matches": res}

    @r.post("/api/pro/outreach/log")
    def outlog(l: LogReq):
        log_outreach(l.id, l.company, l.person, l.message)
        if l.id:
            set_status(l.id, "Referral Requested", "LinkedIn draft for " + l.person)
        return {"ok": True}

    @r.get("/api/pro/profile")
    def gp():
        return profile()

    @r.post("/api/pro/profile")
    def sp(p: dict):
        json.dump(p, open("data/profile.json", "w"), indent=1)
        return {"ok": True}

    @r.post("/api/pro/resume/upload")
    def rup(b: B64):
        os.makedirs("data", exist_ok=True)
        open("data/resume.docx", "wb").write(base64.b64decode(b.b64))
        return {"ok": True}

    @r.post("/api/pro/resume/build")
    def rb():
        return build_resume()

    @r.get("/api/pro/resume/download")
    def rd():
        p = RESUME_OUT if os.path.exists(RESUME_OUT) else G["find_resume"]()
        if not p:
            raise HTTPException(404, "No resume")
        return FileResponse(p, filename="Amit_Mankar_Resume.docx")

    @r.get("/api/pro/export")
    def export():
        _wb()
        return FileResponse(TRACKER, filename="CareerTrace_Tracker.xlsx")

    app.include_router(r)

    @app.middleware("http")
    async def inject(request, call_next):
        resp = await call_next(request)
        p = request.url.path
        if request.method == "GET" and not p.startswith(("/pro", "/api", "/docs", "/redoc", "/openapi")) \
                and "text/html" in resp.headers.get("content-type", ""):
            body = b"".join([c async for c in resp.body_iterator]).decode("utf-8", "ignore")
            tag = '<script src="/pro/app.js"></script>'
            body = body.replace("</body>", tag + "</body>", 1) if "</body>" in body else body + tag
            h = {k: v for k, v in resp.headers.items() if k.lower() not in ("content-length", "content-type")}
            return Response(body, status_code=resp.status_code, headers=h, media_type="text/html")
        return resp

    print("CareerTrace Pro loaded: open the app and click the blue Pro button (or /pro).")
PYEOF

# ================= UI files =================
cat > pro_static/app.js <<'JSEOF'
(function(){
if(window.__proLoaded)return;window.__proLoaded=1;
var b=document.createElement("button");
b.textContent="\u26A1 Pro: Search \u00B7 Tracker \u00B7 Resume \u00B7 LinkedIn";
b.style.cssText="position:fixed;right:18px;bottom:18px;z-index:99999;background:#3b82f6;color:#fff;border:0;border-radius:24px;padding:12px 18px;font:600 14px system-ui;cursor:pointer;box-shadow:0 4px 18px rgba(0,0,0,.4)";
var o=document.createElement("div");
o.style.cssText="position:fixed;inset:0;z-index:100000;background:#0b1220;display:none";
var f=document.createElement("iframe");
f.style.cssText="width:100%;height:100%;border:0";
o.appendChild(f);
b.onclick=function(){if(!f.src)f.src="/pro";o.style.display="block"};
addEventListener("message",function(e){if(e.data==="pro-close")o.style.display="none"});
document.body.appendChild(b);document.body.appendChild(o);
})();
JSEOF

cat > pro_static/pro.html <<'HTMLEOF'
<!doctype html>
<html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>CareerTrace Pro</title>
<style>
:root{--bg:#0b1220;--card:#121a2b;--line:#22304a;--tx:#e6edf7;--mu:#8fa0bd;--ac:#3b82f6;--ok:#22c55e;--warn:#f59e0b}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--tx);font:14px system-ui,Segoe UI,sans-serif}
a{color:var(--ac)}
header{display:flex;gap:20px;align-items:center;padding:12px 24px;border-bottom:1px solid var(--line);flex-wrap:wrap}
header b{font-size:20px}
nav button{background:none;border:0;color:var(--mu);padding:8px 12px;cursor:pointer;font-size:14px;border-bottom:2px solid transparent}
nav button.on{color:var(--tx);border-color:var(--ac)}
main{padding:24px;max-width:1400px;margin:auto}
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:16px;margin-bottom:16px}
input,textarea,select{width:100%;background:#0b1220;color:var(--tx);border:1px solid var(--line);border-radius:6px;padding:8px;font:inherit}
input[type=checkbox]{width:auto}
label{color:var(--mu);font-size:12px;display:block;margin:8px 0 4px}
.g{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:12px}
button.b{background:var(--ac);color:#fff;border:0;border-radius:6px;padding:9px 16px;cursor:pointer;font-size:14px}
button.b:disabled{opacity:.5}
button.s{background:#1b2740;color:var(--tx);border:1px solid var(--line);border-radius:6px;padding:5px 9px;cursor:pointer;margin:2px}
table{width:100%;border-collapse:collapse}
th,td{text-align:left;padding:8px;border-bottom:1px solid var(--line);vertical-align:top}
th{color:var(--mu);font-weight:500}
.kpi{display:flex;gap:12px;flex-wrap:wrap}
.kpi div{flex:1;min-width:120px;background:var(--card);border:1px solid var(--line);border-radius:10px;padding:12px}
.kpi h2{margin:4px 0 0}
.tag{display:inline-block;padding:2px 8px;border-radius:99px;background:#1b2740;font-size:12px;margin:1px}
.tag.ok{background:#12351f;color:#7ee2a0}
.tag.w{background:#3a2a0a;color:#fbbf24}
.sc{font-weight:700;font-size:16px}
.hide{display:none}
.mu{color:var(--mu)}
#toast{position:fixed;right:20px;bottom:20px;background:#1b2740;border:1px solid var(--ac);padding:10px 16px;border-radius:8px;display:none;max-width:420px;z-index:9}
</style></head>
<body>
<header>
<b>CareerTrace <span style="color:var(--ac)">Pro</span></b>
<nav id="nav"></nav>
<span style="margin-left:auto"><a href="/api/pro/export">Download Excel tracker</a> &nbsp;
<button class="s" id="cl" onclick="parent.postMessage('pro-close','*')">Back to CareerTrace</button></span>
</header>
<main>
<section id="dash">
<div class="kpi" id="kpi"></div>
<div class="card" style="margin-top:16px"><h3>Setup status</h3><div id="health"></div>
<p class="mu">Add missing keys to <code>.env.pro</code> (not .env, which run.sh rewrites), then restart ./run.sh.</p></div>
</section>

<section id="find" class="hide">
<div class="card"><div class="g">
<div><label>Job titles (comma separated)</label><textarea id="t" rows="4">Transaction Banking, Cash Management, Cash Management Services, Corporate Cash Management, CMS Product</textarea></div>
<div><label>Location(s) ("India" = all India)</label><input id="l" value="Pune, Mumbai">
<label style="color:var(--tx)"><input type="checkbox" id="sl" checked> Only jobs in these locations</label></div>
<div><label>Keywords - any match</label><input id="k" value="cash management, collections, payments">
<label>Must-have keywords - ALL required</label><input id="mu" placeholder="e.g. RTGS, H2H"></div>
<div><label>Exclude words</label><input id="ex" placeholder="e.g. intern, telecaller">
<label>Posted within</label><select id="dy"><option value="0">Any time</option><option value="1">24 hours</option><option value="3">3 days</option><option value="7">7 days</option><option value="14">14 days</option><option value="30">30 days</option></select></div>
<div><label>Minimum match score (0-100)</label><input id="ms" type="number" value="20">
<label>Max jobs</label><input id="m" type="number" value="25"></div>
<div><label>Mode</label><label style="color:var(--tx)"><input type="checkbox" id="ol" checked> Only open jobs (checks each posting; hides closed/expired)</label><label style="color:var(--tx)"><input type="checkbox" id="d"> Dry run (sample jobs)</label></div>
</div>
<p><button class="b" id="sb" onclick="search()">Search all portals</button> <span id="note" class="mu"></span></p>
<div id="prov"></div></div>
<div class="card"><table><thead><tr><th>Job</th><th>Portal</th><th>Match</th><th>Keywords / skills</th><th>Actions</th></tr></thead>
<tbody id="jobs"><tr><td colspan="5" class="mu">Set your search and press Search.</td></tr></tbody></table>
<p id="drop" class="mu"></p></div>
<div class="card"><h3>Open every portal with your search pre-filled</h3><div id="plinks" class="mu">Run a search to generate links.</div></div>
</section>

<section id="trk" class="hide"><div class="card"><table><thead><tr><th>ID</th><th>Job</th><th>Company</th><th>Portal</th><th>Status</th><th>Applied</th><th>Updated</th></tr></thead><tbody id="apps"></tbody></table></div></section>

<section id="net" class="hide"><div class="card"><h3>LinkedIn network</h3>
<p class="mu">LinkedIn blocks automated access to your connections, so import your own export: Settings &gt; Data privacy &gt; Get a copy of your data &gt; Connections (and Messages for history). Upload the CSVs. Drafts are for you to review and send yourself.</p>
<label>Connections.csv</label><input type="file" id="cf" accept=".csv">
<p><button class="b" onclick="upl('cf','connections','cn')">Upload connections</button> <span id="cn" class="mu"></span></p>
<label>messages.csv (optional - drafts will mention past conversations)</label><input type="file" id="mf" accept=".csv">
<p><button class="b" onclick="upl('mf','messages','mn')">Upload messages</button> <span id="mn" class="mu"></span></p></div></section>

<section id="res" class="hide">
<div class="card"><h3>Resume</h3><p id="rinfo" class="mu"></p><input type="file" id="rf" accept=".docx">
<p><button class="b" onclick="rup()">Upload resume (.docx)</button> <button class="b" onclick="rbuild()">Build updated resume</button> <a id="rdl" href="/api/pro/resume/download" class="hide">Download updated resume</a></p>
<pre id="rout" class="mu"></pre></div>
<div class="card"><h3>Profile used in the resume and every email</h3>
<div class="g"><div><label>Name</label><input id="pn"></div><div><label>Email</label><input id="pe"></div><div><label>Phone</label><input id="pp"></div><div><label>LinkedIn</label><input id="pl"></div><div><label>Blog</label><input id="pb"></div><div><label>Alumni / ex-colleague tags (employers, colleges; comma separated)</label><input id="pa"></div></div>
<h4>Projects (added to the resume and to every email)</h4><div id="prj"></div>
<p><button class="s" onclick="addp()">+ Add project</button> <button class="b" onclick="savep()">Save profile</button></p></div></section>

<section id="draft" class="hide">
<div class="card"><h3 id="dt"></h3><div class="g">
<div><label>Recruiter / hiring manager</label><input id="rn" value="Hiring Team"></div>
<div><label>Recipient email</label><input id="to" placeholder="recruiter@company.com"></div>
<div><label>Extra instruction for the email</label><input id="ext" placeholder="optional"></div>
<div><label>Your history with this role (LinkedIn drafts)</label><input id="hx"></div></div>
<p><button class="b" id="mkb" onclick="mk()">Generate email + LinkedIn drafts</button></p></div>
<div id="dout"></div></section>
</main>
<div id="toast"></div>
<script>
const $=i=>document.getElementById(i),esc=s=>String(s??"").replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const S=[["dash","Dashboard"],["find","Find Jobs"],["trk","Applications"],["net","Network"],["res","Resume & Profile"],["draft","Drafts"]];
let JOBS=[],cur=null,PR=[],OM=[],ST=[],CONN=0,RM=[];
if(window.self===window.top)$("cl").classList.add("hide");
function toast(m){const t=$("toast");t.textContent=m;t.style.display="block";clearTimeout(window._t);window._t=setTimeout(()=>t.style.display="none",4500)}
async function api(u,b,m){
  const o=b===undefined?{}:{method:m||"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(b)};
  const r=await fetch("/api/pro/"+u,o);
  if(!r.ok){let t=await r.text();try{t=JSON.parse(t).detail||t}catch(e){}throw new Error(String(t).slice(0,300))}
  return r.json()}
function b64(buf){let s="",b=new Uint8Array(buf);for(let i=0;i<b.length;i+=8192)s+=String.fromCharCode.apply(null,b.subarray(i,i+8192));return btoa(s)}
$("nav").innerHTML=S.map(s=>`<button id="n${s[0]}" onclick="go('${s[0]}')">${s[1]}</button>`).join("");
function go(s){S.forEach(x=>{$(x[0]).classList.toggle("hide",x[0]!=s);$("n"+x[0]).classList.toggle("on",x[0]==s)});if(s=="dash"||s=="trk")load();if(s=="res")loadRes()}
async function load(){
  ST=await api("statuses");const a=await api("apps"),c={};a.forEach(x=>c[x.Status]=(c[x.Status]||0)+1);
  $("kpi").innerHTML=`<div>Total tracked<h2>${a.length}</h2></div>`+ST.map(s=>`<div>${s}<h2>${c[s]||0}</h2></div>`).join("");
  $("apps").innerHTML=a.map(x=>`<tr><td>${esc(x["Application ID"])}</td><td>${esc(x["Job Title"])}</td><td>${esc(x.Company)}</td><td>${esc(x.Portal)}</td><td><select onchange="st('${esc(x["Application ID"])}',this.value)">${ST.map(s=>`<option ${s==x.Status?"selected":""}>${s}</option>`).join("")}</select></td><td>${esc(x["Date Applied"])}</td><td>${esc(x["Last Updated"])}</td></tr>`).join("");
  const h=await api("health");
  $("health").innerHTML=Object.entries(h.keys).map(([k,v])=>`<span class="tag ${v?"ok":"w"}">${k}: ${v?"ready":"not configured"}</span>`).join("")+`<p class="mu">Resume: ${esc(h.resume||"not found")} | Connections loaded: ${h.connections} | Conversations indexed: ${h.messages}</p>`}
async function st(id,s){await api("status",{id,status:s});load()}
async function search(){
  const b=$("sb");b.disabled=true;$("note").textContent="Searching all configured sources... (can take up to a minute)";
  try{
    const r=await api("search",{titles:$("t").value,location:$("l").value,keywords:$("k").value,must:$("mu").value,exclude:$("ex").value,days:+$("dy").value,min_score:+$("ms").value,max_jobs:+$("m").value,strict_location:$("sl").checked,dry_run:$("d").checked,only_live:$("ol").checked});
    JOBS=r.jobs;try{CONN=(await api("health")).connections}catch(e){}
    $("note").textContent=r.note||`${JOBS.length} matching jobs from ${r.raw_count} fetched`+(r.resume_read?"":" (resume not read - skill gaps unavailable)")+(CONN?` | ${CONN} LinkedIn connections checked for referrals`:" | no LinkedIn connections loaded: upload Connections.csv in the Network tab");
    $("prov").innerHTML=Object.entries(r.providers).map(([k,v])=>`<span class="tag ${v.ok?"ok":"w"}" title="${esc(v.error)}">${esc(k)}: ${v.ok?v.count+" jobs":esc(v.error)}</span>`).join("");
    $("jobs").innerHTML=JOBS.map((x,i)=>`<tr><td><b>${esc(x.title)}</b><br><span class="mu">${esc(x.company)} | ${esc(x.location)} ${x.posted?"| "+esc(x.posted):""}</span>${x.partial?'<br><span class="tag w">snippet only - verify description</span>':""}${liveTag(x)}${refHtml(x,i)}</td><td><span class="tag">${esc(x.portal)}</span></td><td><span class="sc" style="color:${x.score>=70?"var(--ok)":x.score>=50?"var(--warn)":"var(--mu)"}">${x.score}</span>${x.resume_pct!=null?`<br><span class="mu">resume fit ${x.resume_pct}%</span>`:""}</td><td>${x.matched.map(k=>`<span class="tag ok">${esc(k)}</span>`).join("")}${(x.supported||[]).map(k=>`<span class="tag ok">${esc(k)}</span>`).join("")}${(x.gaps||[]).map(k=>`<span class="tag w" title="Required by the job but not in your resume">gap: ${esc(k)}</span>`).join("")}</td><td>${x.url?`<a href="${esc(x.url)}" target="_blank"><button class="s">Open / Apply</button></a>`:""}<button class="s" onclick="sv(${i},'Saved')">Save</button><button class="s" onclick="sv(${i},'Applied')">Mark applied</button><button class="s" onclick="ref(${i})">Referral drafts</button><button class="s" onclick="dr(${i})">Email draft</button></td></tr>`).join("")||'<tr><td colspan="5" class="mu">No jobs matched. Loosen filters or configure more sources.</td></tr>';
    $("drop").textContent="Filtered out: "+(Object.entries(r.dropped).filter(x=>x[1]).map(x=>x[0].replace(/_/g," ")+" "+x[1]).join(", ")||"nothing");
    $("plinks").innerHTML=r.portal_links.map(p=>`<div style="margin:6px 0"><b>${esc(p.title)}</b> <span class="mu">in ${esc(p.location)}</span> &nbsp; ${Object.entries(p.links).map(([k,u])=>`<a href="${esc(u)}" target="_blank"><span class="tag">${k}</span></a>`).join("")}</div>`).join("")
  }catch(e){$("note").textContent="Error: "+e.message}
  b.disabled=false}
async function sv(i,s){const r=await api("apps?status="+s,JOBS[i]);toast(`${s}: ${r["Application ID"]}`)}
function liveTag(x){return x.live==="live"?'<br><span class="tag ok">looks open</span>':(x.live==="unknown"&&!$("d").checked?'<br><span class="tag w" title="Portal blocked the check - open the link to confirm">availability unverified</span>':"")}
function refHtml(x,i){
  const r=x.referrals||[];
  let h=r.length?`<div style="margin-top:6px"><span class="tag ok">${r.length} referral contact${r.length>1?"s":""}</span> ${r.map(m=>`<a href="${esc(m.profile)}" target="_blank" title="${esc(m.position)} @ ${esc(m.company)}"><span class="tag">${esc(m.person)} - ${esc(m.why)}</span></a>`).join("")}</div>`:`<div class="mu" style="margin-top:6px">No matches in loaded connections${CONN?"":" (upload Connections.csv in the Network tab)"}</div>`;
  h+=`<div style="margin-top:4px">${Object.entries(x.li_links||{}).map(([k,u])=>`<a href="${esc(u)}" target="_blank"><span class="tag">${esc(k)}</span></a>`).join("")}</div><div id="rf${i}"></div>`;
  return h}
async function ref(i){
  const x=JOBS[i],box=$("rf"+i);box.innerHTML='<p class="mu">Drafting referral messages...</p>';
  try{
    const a=await api("apps?status=Saved",x);
    const o=await api("outreach",{job:x,recruiter:"Hiring Team",history:$("hx").value,extra:""});
    x.app_id=a["Application ID"];RM[i]=o.matches;
    box.innerHTML=o.matches.length?o.matches.map((m,k)=>`<div class="card" style="margin:8px 0"><b>${esc(m.person)}</b> <span class="tag">${esc(m.why)}</span> <span class="tag">${esc(m.position)} @ ${esc(m.company)}</span>${m.prior?`<span class="tag ok">last chat ${esc(m.prior)}</span>`:""}<textarea id="rm${i}_${k}" rows="7">${esc(m.message)}</textarea><button class="s" onclick="rcp(${i},${k})">Copy + log</button> <a href="${esc(m.profile)}" target="_blank"><button class="s">Open profile</button></a></div>`).join(""):'<p class="mu">No matching connections. Upload Connections.csv in the Network tab, or use the LinkedIn search links above.</p>'
  }catch(e){box.innerHTML='<p class="mu">'+esc(e.message)+'</p>'}}
async function rcp(i,k){const m=RM[i][k],t=$("rm"+i+"_"+k).value;try{await navigator.clipboard.writeText(t)}catch(e){}await api("outreach/log",{id:JOBS[i].app_id||"",company:m.company,person:m.person,message:t});toast("Copied and logged in the Outreach tab")}
function dr(i){cur=JOBS[i];$("dt").textContent=cur.title+" @ "+cur.company;$("dout").innerHTML="";go("draft")}
async function mk(){
  const b0=$("mkb");b0.disabled=true;
  try{
    const b={job:cur,recruiter:$("rn").value,history:$("hx").value,extra:$("ext").value};
    const [e,o]=await Promise.all([api("email",b),api("outreach",b)]);
    cur.app_id=e.app_id;OM=o.matches;
    let h=`<div class="card"><h3>Email <span class="tag">${esc(e.app_id)}</span></h3><label>Subject</label><input id="es" value="${esc(e.subject)}"><label>Body</label><textarea id="eb" rows="18">${esc(e.body)}</textarea><p><button class="b" onclick="gm()">Open in Gmail</button> <button class="b" onclick="sendm()">Send with resume attached</button> <a href="/api/pro/resume/download">Download resume</a></p></div><div class="card"><h3>LinkedIn drafts (${OM.length} matches from ${o.connections_loaded} connections)</h3>${OM.length?"":'<p class="mu">No matching connections. Upload Connections.csv in the Network tab.</p>'}`;
    OM.forEach((m,k)=>{h+=`<div style="margin-bottom:14px"><b>${esc(m.person)}</b> <span class="tag">${esc(m.position)} @ ${esc(m.company)}</span>${m.prior?`<span class="tag ok">last chat ${esc(m.prior)}</span>`:""} <a href="${esc(m.profile)}" target="_blank">profile</a><textarea id="lm${k}" rows="8">${esc(m.message)}</textarea><button class="s" onclick="cp(${k})">Copy + log</button></div>`});
    $("dout").innerHTML=h+"</div>"
  }catch(e){toast(e.message)}
  b0.disabled=false}
function gm(){window.open("https://mail.google.com/mail/?view=cm&fs=1&to="+encodeURIComponent($("to").value)+"&su="+encodeURIComponent($("es").value)+"&body="+encodeURIComponent($("eb").value),"_blank")}
async function sendm(){const to=$("to").value.trim();if(!to)return toast("Enter the recipient email first");if(!confirm("Send this email to "+to+" with your resume attached?"))return;
  try{await api("email/send",{to,subject:$("es").value,body:$("eb").value,app_id:cur.app_id});toast("Sent. Status set to Recruiter Contacted")}catch(e){toast(e.message)}}
async function cp(k){const m=OM[k],t=$("lm"+k).value;try{await navigator.clipboard.writeText(t)}catch(e){}
  await api("outreach/log",{id:cur.app_id||"",company:m.company,person:m.person,message:t});toast("Copied and logged in the Outreach tab")}
async function upl(f,ep,o){const x=$(f).files[0];if(!x)return toast("Choose a file first");const r=await api(ep,{csv:await x.text()});$(o).textContent=r.count+" records loaded"}
async function loadRes(){
  const h=await api("health");$("rinfo").textContent="Detected resume: "+(h.resume||"none - upload one");$("rdl").classList.toggle("hide",!h.resume_built);
  const p=await api("profile");$("pn").value=p.name;$("pe").value=p.email;$("pp").value=p.phone;$("pl").value=p.linkedin;$("pb").value=p.blog;$("pa").value=p.alumni_terms||"";PR=p.projects||[];rp()}
function rp(){$("prj").innerHTML=PR.map((x,i)=>`<div class="g" style="margin-bottom:8px"><input value="${esc(x.name)}" onchange="PR[${i}].name=this.value" placeholder="Project name"><input value="${esc(x.url)}" onchange="PR[${i}].url=this.value" placeholder="URL"><input value="${esc(x.summary||"")}" onchange="PR[${i}].summary=this.value" placeholder="One line: what it does (optional)"><button class="s" onclick="PR.splice(${i},1);rp()">Remove</button></div>`).join("")}
function addp(){PR.push({name:"",url:"",summary:""});rp()}
async function savep(){await api("profile",{name:$("pn").value,email:$("pe").value,phone:$("pp").value,linkedin:$("pl").value,blog:$("pb").value,alumni_terms:$("pa").value,projects:PR});toast("Profile saved")}
async function rup(){const x=$("rf").files[0];if(!x)return toast("Choose a .docx first");await api("resume/upload",{b64:b64(await x.arrayBuffer())});toast("Resume uploaded");loadRes()}
async function rbuild(){await savep();try{const r=await api("resume/build",{});$("rout").textContent="Built from "+r.base+"\n- "+r.changes.join("\n- ");$("rdl").classList.remove("hide")}catch(e){$("rout").textContent=e.message}}
go("dash");
</script></body></html>
HTMLEOF

# ================= config, safe loader, rollback check =================
[ -f .env.pro ] || printf '# CareerTrace Pro keys (run.sh rewrites .env, so extra keys live here)\nRAPIDAPI_KEY=\nADZUNA_APP_ID=\nADZUNA_APP_KEY=\nJOOBLE_KEY=\nGMAIL_APP_PASSWORD=\n' > .env.pro
for f in .env.pro data/connections.json data/messages.json data/profile.json exports/ output/ '*.bak_pro'; do
  grep -qxF "$f" .gitignore 2>/dev/null || echo "$f" >> .gitignore
done
if ! grep -q "careertrace_pro" careertrace_ui.py; then
cat >> careertrace_ui.py <<'LOADEOF'


# --- CareerTrace Pro (safe loader: a failure here never stops the app) ---
try:
    import careertrace_pro
    careertrace_pro.install(app, globals())
except Exception as _e:
    print("CareerTrace Pro not loaded:", _e)
LOADEOF
fi
$PY -m py_compile careertrace_pro.py
if $PY -c "import careertrace_ui as u; assert '/api/pro/health' in u.app.openapi()['paths']" >/tmp/pro_check.log 2>&1; then
  echo "OK: CareerTrace Pro installed. Restart your app (Ctrl+C, then ./run.sh), open port 8000, and click the blue Pro button (or go to /pro)."
else
  cp careertrace_ui.py.bak_pro careertrace_ui.py
  echo "Install check failed, your app was restored unchanged. Details:"; tail -15 /tmp/pro_check.log
fi
