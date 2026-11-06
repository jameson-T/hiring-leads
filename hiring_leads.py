#!/usr/bin/env python3
"""
Daily hiring-companies generator.

Pulls live job postings from legit job APIs (Jooble, Adzuna, Reed), keeps one row per
company, verifies each one, skips companies you already got in the last 30 days, and
saves 40-60 fresh companies to CSV + XLSX every run.

Setup: see README.md. Needs: pip install requests openpyxl
"""
import csv
import json
import os
import random
import re
import sys
import time
from datetime import datetime, timedelta, timezone
from concurrent.futures import ThreadPoolExecutor
from urllib.parse import quote_plus, urlparse

import requests

# ----------------------------- CONFIG ---------------------------------------
TARGET = int(os.getenv("TARGET", "50"))              # companies per run (40-60)
MAX_AGE_DAYS = int(os.getenv("MAX_AGE_DAYS", "14"))  # only jobs posted within this window
SKIP_SEEN_DAYS = int(os.getenv("SKIP_SEEN_DAYS", "3650"))  # ~10 years = never repeat a company
PAGES = int(os.getenv("PAGES", "3"))
KEYWORDS = [k.strip() for k in os.getenv(
    "KEYWORDS", "marketing,designer,developer,sales,manager").split(",") if k.strip()]
EXCLUDE_RECRUITERS = os.getenv("EXCLUDE_RECRUITERS", "1") == "1"
REMOTE_FOCUS = os.getenv("REMOTE_FOCUS", "1") == "1"  # 1 = search remote roles first
REMOTE_ONLY = os.getenv("REMOTE_ONLY", "0") == "1"    # 1 = drop anything not clearly remote

# API keys (set as environment variables, see README)
JOOBLE_KEY = os.getenv("JOOBLE_KEY", "")
ADZUNA_ID = os.getenv("ADZUNA_APP_ID", "")
ADZUNA_KEY = os.getenv("ADZUNA_APP_KEY", "")
REED_KEY = os.getenv("REED_KEY", "")
COMPANIES_HOUSE_KEY = os.getenv("COMPANIES_HOUSE_KEY", "")  # free, UK verification
SERPAPI_KEY = os.getenv("SERPAPI_KEY", "")                  # optional, finds real LinkedIn page

# Add or remove countries here. adzuna = Adzuna country code (if supported).
COUNTRIES = {
    "United Kingdom": {"adzuna": "gb", "reed": True, "jooble": "United Kingdom"},
    "UAE":            {"jooble": "United Arab Emirates"},
    "Lebanon":        {"jooble": "Lebanon"},
    "France":         {"adzuna": "fr", "jooble": "France"},
    "Germany":        {"adzuna": "de", "jooble": "Germany"},
    "Canada":         {"adzuna": "ca", "jooble": "Canada"},
    # Remote-only job board, companies hiring worldwide (no API key needed)
    "Remote (Worldwide)": {"remotive": True},
}

# ---------------- UAE B2B MODE (UAE_MODE=1) ----------------
# UAE only, split by city so the 60 companies come from a wide pool.
UAE_MODE = os.getenv("UAE_MODE", "0") == "1"
PREFIX = "uae_b2b" if UAE_MODE else "hiring_companies"
if UAE_MODE:
    if "MAX_AGE_DAYS" not in os.environ:
        MAX_AGE_DAYS = 7   # UAE mode: only jobs from the last 7 days (set MAX_AGE_DAYS to change)
    if "KEYWORDS" not in os.environ:
        KEYWORDS = ["freelance", "contract", "project manager", "digital marketing",
                    "web developer", "graphic designer", "social media", "seo",
                    "video editor", "branding", "crm", "erp", "business development"]
    COUNTRIES = {
        "UAE - Dubai":     {"jooble": "Dubai"},
        "UAE - Abu Dhabi": {"jooble": "Abu Dhabi"},
        "UAE - Sharjah":   {"jooble": "Sharjah"},
        "UAE - Other":     {"jooble": "United Arab Emirates"},
    }

# Obvious consumer-facing / non-B2B jobs get dropped in UAE mode
B2C_WORDS = ("restaurant", "cafe", "café", "hotel", "salon", "barber", "driver", "cashier",
             "waiter", "waitress", "barista", "chef", "cook", "nanny", "maid", "housekeeping",
             "security guard", "supermarket", "hypermarket", "beautician", "nurse", "teacher",
             "cleaner", "delivery rider", "storekeeper")
PROJECT_WORDS = ("freelance", "contract", "project", "consultant", "consulting", "temporary",
                 "short-term", "short term", "part-time", "part time", "per project")

RECRUITER_WORDS = ("recruit", "staffing", "talent", "resourc", "headhunt", "personnel",
                   "manpower", "randstad", "adecco", "hays", "michael page", "robert half",
                   "reed specialist", "search & selection", "consultancy ltd")

HERE = os.path.dirname(os.path.abspath(__file__))
HISTORY = os.path.join(HERE, "seen.json")
UA = {"User-Agent": "Mozilla/5.0 (compatible; HiringLeads/1.0)"}


def norm(name):
    n = re.sub(r"[^a-z0-9 ]", "", name.lower())
    n = re.sub(r"\b(ltd|limited|llc|inc|plc|sarl|sa|gmbh|fze|fzco|co|the|group|uk|corp|corporation)\b",
               "", n)
    return re.sub(r"\s+", " ", n).strip()


def parse_date(s):
    if not s:
        return None
    try:
        d = datetime.fromisoformat(str(s).replace("Z", "+00:00")[:25])
        return d if d.tzinfo else d.replace(tzinfo=timezone.utc)
    except Exception:
        return None


def fresh(d):
    # unknown date = allow; known date must be inside the window
    return d is None or d >= datetime.now(timezone.utc) - timedelta(days=MAX_AGE_DAYS)


def is_recruiter(name):
    low = name.lower()
    return EXCLUDE_RECRUITERS and any(w in low for w in RECRUITER_WORDS)


def is_b2c(company, role):
    blob = f"{company} {role}".lower()
    return any(w in blob for w in B2C_WORDS)


def is_project(*texts):
    blob = " ".join(t for t in texts if t).lower()
    return any(w in blob for w in PROJECT_WORDS)


def is_remote(*texts):
    blob = " ".join(t for t in texts if t).lower()
    return any(w in blob for w in ("remote", "work from home", "work-from-home",
                                   "télétravail", "teletravail", "home office",
                                   "anywhere", "hybrid"))


# ----------------------------- SOURCES --------------------------------------
def from_jooble(country, loc, kw, page):
    if not JOOBLE_KEY:
        return []
    try:
        r = requests.post(f"https://jooble.org/api/{JOOBLE_KEY}",
                          json={"keywords": kw, "location": loc, "page": str(page)}, timeout=20)
        jobs = r.json().get("jobs", [])
    except Exception as e:
        print(f"  jooble {country}/{kw}: {e}")
        return []
    return [dict(company=j.get("company", "").strip(), role=j.get("title", ""),
                 url=j.get("link", ""), posted=parse_date(j.get("updated")),
                 country=country, source="Jooble", ptext=j.get("snippet"),
                 remote=is_remote(j.get("title"), j.get("location"), j.get("snippet"))
                 ) for j in jobs]


def from_adzuna(country, code, kw, page):
    if not (ADZUNA_ID and ADZUNA_KEY):
        return []
    try:
        r = requests.get(f"https://api.adzuna.com/v1/api/jobs/{code}/search/{page}",
                         params={"app_id": ADZUNA_ID, "app_key": ADZUNA_KEY,
                                 "results_per_page": 50, "what": kw,
                                 "max_days_old": MAX_AGE_DAYS,
                                 "content-type": "application/json"}, timeout=20)
        res = r.json().get("results", [])
    except Exception as e:
        print(f"  adzuna {country}/{kw}: {e}")
        return []
    return [dict(company=(j.get("company") or {}).get("display_name", "").strip(),
                 role=j.get("title", ""), url=j.get("redirect_url", ""),
                 posted=parse_date(j.get("created")), country=country, source="Adzuna",
                 remote=is_remote(j.get("title"), (j.get("location") or {}).get("display_name"),
                                  j.get("description")))
            for j in res]


def from_reed(country, kw, page):
    if not REED_KEY:
        return []
    try:
        r = requests.get("https://www.reed.co.uk/api/1.0/search",
                         params={"keywords": kw, "resultsToTake": 100,
                                 "resultsToSkip": (page - 1) * 100},
                         auth=(REED_KEY, ""), timeout=20)
        res = r.json().get("results", [])
    except Exception as e:
        print(f"  reed {kw}: {e}")
        return []
    out = []
    for j in res:
        d = None
        try:
            d = datetime.strptime(j.get("date", ""), "%d/%m/%Y").replace(tzinfo=timezone.utc)
        except Exception:
            pass
        out.append(dict(company=j.get("employerName", "").strip(), role=j.get("jobTitle", ""),
                        url=j.get("jobUrl", ""), posted=d, country=country, source="Reed",
                        remote=is_remote(j.get("jobTitle"), j.get("locationName"),
                                         j.get("jobDescription"))))
    return out


def from_remotive(country):
    """Remote-only jobs from Remotive (free, no key). One call per run on purpose."""
    try:
        r = requests.get("https://remotive.com/api/remote-jobs", params={"limit": 300},
                         headers=UA, timeout=30)
        jobs = r.json().get("jobs", [])
    except Exception as e:
        print(f"  remotive: {e}")
        return []
    out = []
    for j in jobs:
        out.append(dict(company=j.get("company_name", "").strip(), role=j.get("title", ""),
                        url=j.get("url", ""), posted=parse_date(j.get("publication_date")),
                        country=country, source="Remotive", remote=True,
                        tags=" ".join(j.get("tags") or []) + " " + (j.get("category") or "")))
    # prefer roles matching your keywords, but never return an empty list
    kws = [k.lower() for k in KEYWORDS]
    matched = [o for o in out if any(k in (o["role"] + " " + o["tags"]).lower() for k in kws)]
    return matched if len(matched) >= 15 else out


# --------------------------- VERIFICATION -----------------------------------
def job_is_live(url):
    """dead / live / unknown. Many job boards block bots, so 403 = unknown, not dead."""
    if not url:
        return "unknown"
    try:
        r = requests.get(url, headers=UA, timeout=10, allow_redirects=True, stream=True)
        r.close()
        if r.status_code in (404, 410):
            return "dead"
        return "live" if r.status_code < 400 else "unknown"
    except Exception:
        return "unknown"


# STRICT = only keep a job if its page was actually opened and looks open. Anything the site
# blocked or that we could not check is DROPPED (set STRICT_VERIFY=0 to allow those).
STRICT = os.getenv("STRICT_VERIFY", "1") == "1"
MAX_CHECK = int(os.getenv("MAX_CHECK", "600"))
_status_cache = {}
CLOSED_PHRASES = (
    "no longer available", "no longer accepting", "no longer open", "has expired",
    "job expired", "expired job", "position has been filled", "position is filled",
    "vacancy has been filled", "this job is closed", "job is closed", "vacancy is closed",
    "has been closed", "applications are closed", "applications closed", "job not found",
    "this job has been removed", "job has been removed", "posting has been removed",
    "this position is no longer", "this vacancy is no longer", "no longer active",
    "page not found", "404 not found", "job no longer exists", "this job does not exist",
    "offre n'est plus disponible", "offre expirée", "cette offre a expiré",
    "n'est plus disponible", "stellenanzeige ist nicht mehr", "nicht mehr verfügbar",
    "هذه الوظيفة لم تعد", "انتهت صلاحية",
)


def _check_job(url):
    if not url:
        return "unknown"
    try:
        r = requests.get(url, headers=UA, timeout=12, allow_redirects=True)
    except Exception:
        return "unknown"
    if r.status_code in (404, 410):
        return "dead"
    if r.status_code >= 400:
        return "unknown"            # blocked / error: cannot confirm
    final, orig = urlparse(r.url), urlparse(url)
    if final.path.strip("/") == "" and orig.path.strip("/") != "":
        return "dead"               # bounced to a homepage = the job was removed
    text = r.text[:400000].lower()
    if any(p in text for p in CLOSED_PHRASES):
        return "dead"
    return "live"


def job_status(url):
    """live = page opens and shows no closed/expired notice. dead = gone/closed. unknown = could not check."""
    if url not in _status_cache:
        _status_cache[url] = _check_job(url)
    return _status_cache[url]


def companies_house(name):
    """Returns (found_active, text) for UK companies."""
    if not COMPANIES_HOUSE_KEY:
        return None, ""
    try:
        r = requests.get("https://api.company-information.service.gov.uk/search/companies",
                         params={"q": name, "items_per_page": 5},
                         auth=(COMPANIES_HOUSE_KEY, ""), timeout=15)
        for it in r.json().get("items", []):
            if norm(it.get("title", "")) == norm(name):
                ok = it.get("company_status") == "active"
                return ok, f"Companies House {it.get('company_status')} #{it.get('company_number')}"
    except Exception:
        pass
    return False, "Not found on Companies House"


def linkedin_link(name):
    """Real company page if SerpAPI is set, else a LinkedIn company search link."""
    if SERPAPI_KEY:
        try:
            r = requests.get("https://serpapi.com/search.json",
                             params={"q": f'site:linkedin.com/company "{name}"',
                                     "api_key": SERPAPI_KEY, "num": 3}, timeout=20)
            for it in r.json().get("organic_results", []):
                m = re.match(r"https?://[a-z.]*linkedin\.com/company/[^/?#]+", it.get("link", ""))
                if m:
                    return m.group(0), "page found via search"
        except Exception:
            pass
    return ("https://www.linkedin.com/search/results/companies/?keywords=" + quote_plus(name),
            "search link (open and pick the right page)")


# ------------------------------ MAIN ----------------------------------------
def load_history():
    """Every company ever given to you: seen.json PLUS every old hiring_companies_*.csv file
    in the folder. Even if seen.json gets lost, old files still block repeats."""
    import glob
    h = {}
    try:
        with open(HISTORY) as f:
            h = json.load(f)
    except Exception:
        pass
    now = datetime.now(timezone.utc).isoformat()
    paths = (glob.glob(os.path.join(HERE, "hiring_companies_*.csv")) +
             glob.glob(os.path.join(HERE, "uae_b2b_*.csv")))
    for path in paths:
        try:
            with open(path, newline="", encoding="utf-8-sig") as f:
                for row in csv.DictReader(f):
                    n = norm(row.get("Company", ""))
                    if n:
                        h.setdefault(n, now)
        except Exception:
            pass
    return h


def main():
    if not (JOOBLE_KEY or (ADZUNA_ID and ADZUNA_KEY) or REED_KEY):
        sys.exit("No API keys set. Open README.md and add at least JOOBLE_KEY.")

    history = load_history()
    cutoff = datetime.now(timezone.utc) - timedelta(days=SKIP_SEEN_DAYS)
    seen_recent = {k for k, v in history.items()
                   if (parse_date(v) or datetime.min.replace(tzinfo=timezone.utc)) > cutoff}

    pool = {}  # (country, normname) -> record

    def ingest(country, jobs):
        for j in jobs:
            n = norm(j["company"])
            if not n or len(n) < 3 or n in seen_recent or is_recruiter(j["company"]):
                continue
            if not fresh(j["posted"]):
                continue
            if REMOTE_ONLY and not j.get("remote"):
                continue
            if UAE_MODE and is_b2c(j["company"], j["role"]):
                continue
            j["project"] = is_project(j["role"], j.get("ptext"))
            key = (country, n)
            if key in pool:
                pool[key]["roles"].add(j["role"])
                pool[key]["sources"].add(j["source"])
                pool[key]["remote"] = pool[key]["remote"] or j.get("remote", False)
                pool[key]["project"] = pool[key].get("project") or j["project"]
            else:
                pool[key] = dict(j, roles={j["role"]}, sources={j["source"]},
                                 remote=j.get("remote", False))

    # remote-flavoured searches first, then the normal ones
    terms = ([f"remote {k}" for k in KEYWORDS] + KEYWORDS) if REMOTE_FOCUS else KEYWORDS
    if REMOTE_ONLY:
        terms = [f"remote {k}" for k in KEYWORDS]

    if UAE_MODE:
        terms = list(KEYWORDS)  # no "remote" prefix, that pulls worldwide jobs into UAE
    for country, cfg in COUNTRIES.items():
        print(f"Collecting {country}...")
        if cfg.get("remotive"):
            ingest(country, from_remotive(country))
            continue
        for kw in terms:
            for page in range(1, PAGES + 1):
                jobs = []
                if "jooble" in cfg:
                    jobs += from_jooble(country, cfg["jooble"], kw, page)
                if "adzuna" in cfg:
                    jobs += from_adzuna(country, cfg["adzuna"], kw, page)
                if cfg.get("reed"):
                    jobs += from_reed(country, kw, page)
                ingest(country, jobs)
                time.sleep(0.3)

    # round-robin across countries so one country doesn't eat the whole list
    by_country = {}
    floor = datetime.min.replace(tzinfo=timezone.utc)
    for rec in pool.values():
        by_country.setdefault(rec["country"], []).append(rec)
    for lst in by_country.values():
        # NEWEST DAY FIRST. Within the same day: project/contract, then remote, then more boards.
        # Jobs with no date go last.
        lst.sort(key=lambda r: ((r["posted"] or floor).date(), r.get("project", False),
                                r["remote"], len(r["sources"])),
                 reverse=True)

    results, queues = [], {c: list(v) for c, v in by_country.items()}
    used = set()  # a company can only appear once per file, even if it hires in several countries
    used_urls = set()
    stats = {"checked": 0, "live": 0, "closed": 0, "unverifiable": 0}

    # open the job pages in parallel (fast), in the same order we will use them
    order, qs = [], {c: list(v) for c, v in queues.items()}
    while any(qs.values()) and len(order) < MAX_CHECK:
        for c in list(qs):
            if qs[c] and len(order) < MAX_CHECK:
                order.append(qs[c].pop(0)["url"])
    print(f"Opening {len(order)} job pages to confirm they are still open...")
    with ThreadPoolExecutor(max_workers=10) as ex:
        list(ex.map(job_status, order))
    print(f"Pool: {sum(len(v) for v in queues.values())} candidate companies. Verifying...")
    while len(results) < TARGET and any(queues.values()):
        for c in list(queues):
            if len(results) >= TARGET:
                break
            if not queues[c]:
                continue
            rec = queues[c].pop(0)
            if norm(rec["company"]) in used or rec["url"] in used_urls:
                continue
            live = job_status(rec["url"])
            stats["checked"] += 1
            if live == "dead":
                stats["closed"] += 1
                continue
            if live != "live" and STRICT:
                stats["unverifiable"] += 1
                continue
            stats["live"] += 1
            verify = [f"job page opened on {datetime.now().strftime('%Y-%m-%d')}, no closed notice"
                      if live == "live" else "job page could not be checked"]
            if c == "United Kingdom" and COMPANIES_HOUSE_KEY:
                ok, txt = companies_house(rec["company"])
                if not ok:
                    continue  # UK company not active/found -> drop it
                verify.append(txt)
            if len(rec["sources"]) > 1:
                verify.append("on " + " + ".join(sorted(rec["sources"])))
            li, li_note = linkedin_link(rec["company"])
            used.add(norm(rec["company"]))
            used_urls.add(rec["url"])
            results.append({
                "Company": rec["company"], "Country": c,
                "Remote?": "Yes" if rec["remote"] else "Not stated",
                "Signal": "Contract / freelance / project" if rec.get("project") else "Direct hire",
                "Roles hiring": " | ".join(sorted(rec["roles"]))[:200],
                "Job link": rec["url"],
                "Posted": rec["posted"].strftime("%Y-%m-%d") if rec["posted"] else "unknown",
                "LinkedIn": li, "LinkedIn note": li_note,
                "Verification": "; ".join(verify),
            })

    if not results:
        sys.exit("No NEW companies found today (everything returned was already given to you). "
                 "Add more KEYWORDS or more API keys and run again.")

    # final safety net: no duplicate company can ever be written
    assert len({norm(r["Company"]) for r in results}) == len(results), "duplicate company!"
    if len(results) < 40:
        print(f"Note: only {len(results)} brand-new companies today. Add more KEYWORDS/API keys "
              f"to get 40-60.")

    stamp = datetime.now().strftime("%Y-%m-%d")
    base = os.path.join(HERE, f"{PREFIX}_{stamp}")
    fields = list(results[0].keys())
    with open(base + ".csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(results)
    try:
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Hiring today"
        ws.append(fields)
        for r in results:
            ws.append([r[k] for k in fields])
        for col, width in zip("ABCDEFGHIJ", (30, 18, 12, 28, 45, 50, 12, 55, 30, 45)):
            ws.column_dimensions[col].width = width
        wb.save(base + ".xlsx")
    except ImportError:
        print("openpyxl not installed, CSV only (pip install openpyxl for Excel).")

    now = datetime.now(timezone.utc).isoformat()
    for r in results:
        history[norm(r["Company"])] = now
    with open(HISTORY, "w") as f:
        json.dump(history, f)

    counts = {}
    for r in results:
        counts[r["Country"]] = counts.get(r["Country"], 0) + 1
    print(f"\nDone. {len(results)} companies saved to {base}.xlsx")
    print("By country:", counts)
    vals = list(_status_cache.values())
    print(f"Job pages opened: {len(vals)} | open: {vals.count('live')} | "
          f"closed/expired/removed: {vals.count('dead')} | could not verify: {vals.count('unknown')} "
          f"| {len(results)} confirmed-open companies sent to you")


if __name__ == "__main__":
    main()
