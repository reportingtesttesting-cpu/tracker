import os
import socket

_orig_getaddrinfo = socket.getaddrinfo
def _ipv4_only(host, port, family=0, type=0, proto=0, flags=0):
    return _orig_getaddrinfo(host, port, socket.AF_INET, type, proto, flags)
socket.getaddrinfo = _ipv4_only

import json
import time
import tempfile
from datetime import datetime, timezone, timedelta

import gspread
from google.oauth2.credentials import Credentials as UserCredentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import pandas as pd
from dotenv import load_dotenv
from jinja2 import Environment, FileSystemLoader
import requests


load_dotenv()

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive",
]

GOOGLE_CLIENT_ID = os.environ.get("GOOGLE_CLIENT_ID", "").strip()
GOOGLE_CLIENT_SECRET = os.environ.get("GOOGLE_CLIENT_SECRET", "").strip()
GOOGLE_REFRESH_TOKEN = os.environ.get("GOOGLE_REFRESH_TOKEN", "").strip()
DEEPSEEK_API_KEY = os.environ.get("DEEPSEEK_API_KEY", "").strip()

SHEET_ID = os.environ.get("SHEET_ID", "").strip()

TABS = ["Config", "People", "Tasks", "Defects", "Dependencies", "Updates"]

TEMPLATES_DIR = os.path.join(os.path.dirname(__file__), "..", "templates")
DASHBOARDS_DIR = os.path.join(os.path.dirname(__file__), "..", "dashboards")

CLOSED_DEFECT_STATES = ["Closed", "Verified", "Rejected", "Duplicate"]

_credentials_cache = None


def get_credentials():
    global _credentials_cache
    if _credentials_cache is None:
        _credentials_cache = UserCredentials(
            token=None,
            refresh_token=GOOGLE_REFRESH_TOKEN,
            token_uri="https://oauth2.googleapis.com/token",
            client_id=GOOGLE_CLIENT_ID,
            client_secret=GOOGLE_CLIENT_SECRET,
            scopes=SCOPES,
        )
        _credentials_cache.refresh(Request())
        print("  OAuth token refreshed")
    return _credentials_cache


def connect():
    creds = get_credentials()
    client = gspread.authorize(creds)
    return client.open_by_key(SHEET_ID)


def read_tab(ss, name):
    try:
        ws = ss.worksheet(name)
    except gspread.WorksheetNotFound:
        print(f"  Tab '{name}' not found")
        return pd.DataFrame()
    data = ws.get_all_values()
    if not data:
        return pd.DataFrame()
    header = data[0]
    rows = data[1:]
    if not rows:
        return pd.DataFrame(columns=header)
    return pd.DataFrame(rows, columns=header)


def read_config(ss):
    ws = ss.worksheet("Config")
    data = ws.get_all_values()
    cfg = {}
    for row in data:
        if len(row) >= 2 and row[0]:
            cfg[row[0]] = row[1]
    return cfg


def parse_date(v):
    try:
        return datetime.strptime(str(v)[:10], "%Y-%m-%d").date()
    except Exception:
        return None


def build_maps(people_df):
    company_map = {}
    name_map = {}
    if len(people_df) > 0:
        for _, p in people_df.iterrows():
            pid = str(p.get("person_id", "")).strip()
            if pid:
                company_map[pid] = str(p.get("company", "")).strip()
                name_map[pid] = str(p.get("name", "")).strip()
    return company_map, name_map


def build_tasks(tasks_df, company_map, name_map):
    today = datetime.now(timezone.utc).date()
    out = []
    for _, r in tasks_df.iterrows():
        if not r.get("task_id"):
            continue
        owner_id = str(r.get("owner_id", ""))
        status = str(r.get("status", "Not Started"))
        
        be = parse_date(r.get("baseline_end"))
        fe = parse_date(r.get("forecast_end"))
        af = parse_date(r.get("actual_finish"))
        
        pct = int(float(r.get("percent_complete", 0) or 0))
        pct = min(100, max(0, pct)) # Clamp 0-100

        # Slip Logic: If complete, use actual_finish or today. Else use forecast.
        slip = 0
        if be:
            if status == "Complete":
                finish_date = af or today
                slip = (finish_date - be).days
            elif fe:
                slip = (fe - be).days

        out.append({
            "id": str(r.get("task_id")),
            "parent_id": str(r.get("parent_task_id", "")).strip(),
            "name": str(r.get("name", "")),
            "description": str(r.get("description", "")),
            "owner_id": owner_id,
            "owner_name": name_map.get(owner_id, owner_id),
            "owner_company": company_map.get(owner_id, ""),
            "status": status,
            "percent": pct,
            "priority": str(r.get("priority", "P2")),
            "critical": str(r.get("critical_flag", "")).upper() == "TRUE",
            "type": str(r.get("type", "contractual")),
            "baseline_start": str(r.get("baseline_start", ""))[:10],
            "baseline_end": str(r.get("baseline_end", ""))[:10],
            "forecast_start": str(r.get("forecast_start", ""))[:10],
            "forecast_end": str(r.get("forecast_end", ""))[:10],
            "actual_start": str(r.get("actual_start", ""))[:10],
            "actual_finish": str(r.get("actual_finish", ""))[:10],
            "slip_days": slip,
            "cascade_start": "",
            "cascade_end": "",
            "cascade_slip_days": 0,
        })
    return out


def build_defects(defects_df, company_map, name_map):
    today = datetime.now(timezone.utc).date()
    out = []
    for _, r in defects_df.iterrows():
        if not r.get("defect_id"):
            continue
        
        assigned_raw = str(r.get("assigned_to", ""))
        assigned_ids = [x.strip() for x in assigned_raw.split(",") if x.strip()]
        
        companies = set()
        names = set()
        for aid in assigned_ids:
            if aid in company_map: companies.add(company_map[aid])
            if aid in name_map: names.add(name_map[aid])
            
        assigned_name = ", ".join(names) if names else assigned_raw
        assigned_company = ", ".join(companies) if companies else ""

        det = parse_date(r.get("detected_at"))
        clo = parse_date(r.get("closed_at"))
        age = (today - det).days if det else None
        mttr = (clo - det).days if (det and clo) else None
        status = str(r.get("status", "New"))
        
        out.append({
            "id": str(r.get("defect_id")),
            "title": str(r.get("title", "")),
            "description": str(r.get("description", "")),
            "severity": str(r.get("severity", "Minor")),
            "priority": str(r.get("priority", "P2")),
            "status": status,
            "assigned_to": assigned_raw,
            "assigned_name": assigned_name,
            "assigned_company": assigned_company,
            "detected_at": str(r.get("detected_at", ""))[:10],
            "closed_at": str(r.get("closed_at", ""))[:10],
            "root_cause": str(r.get("root_cause", "")),
            "corrective_action": str(r.get("corrective_action", "")),
            "linked_task_id": str(r.get("linked_task_id", "")),
            "environment": str(r.get("environment", "")),
            "reproducibility": str(r.get("reproducibility", "")),
            "age_days": age,
            "mttr_days": mttr,
            "is_open": status not in CLOSED_DEFECT_STATES,
        })
    return out


def build_dependencies(deps_df):
    out = []
    if len(deps_df) == 0:
        return out
    for _, r in deps_df.iterrows():
        p = str(r.iloc[0] if len(r) > 0 else "").strip()
        s = str(r.iloc[1] if len(r) > 1 else "").strip()
        if not p or not s:
            continue
        out.append({
            "predecessor_id": p,
            "successor_id": s,
            "type": str(r.iloc[2] if len(r) > 2 else "FS"),
            "lag": int(float(r.iloc[3])) if len(r) > 3 and str(r.iloc[3]).strip() else 0,
        })
    return out


def build_updates(updates_df):
    out = []
    if len(updates_df) == 0:
        return out
    for _, r in updates_df.iterrows():
        ts = str(r.get("timestamp", ""))[:10]
        out.append({
            "timestamp": ts,
            "entity_type": str(r.get("entity_type", "")),
            "entity_id": str(r.get("entity_id", "")),
            "person_id": str(r.get("person_id", "")),
            "field": str(r.get("field", "")),
            "old_value": str(r.get("old_value", "")),
            "new_value": str(r.get("new_value", "")),
            "notes": str(r.get("notes", "")),
            "source": str(r.get("source", "")),
        })
    out.sort(key=lambda x: x["timestamp"], reverse=True)
    return out


def build_history(updates):
    history = {}
    for u in updates:
        eid = u["entity_id"]
        if not eid: continue
        history.setdefault(eid, []).append(u)
    return history


def apply_cascade(tasks, dependencies):
    task_map = {t["id"]: t for t in tasks}
    for t in tasks:
        cs = parse_date(t.get("forecast_start")) or parse_date(t.get("baseline_start"))
        if t["status"] == "Complete":
            ce = parse_date(t.get("actual_finish")) or parse_date(t.get("forecast_end")) or parse_date(t.get("baseline_end"))
        else:
            ce = parse_date(t.get("forecast_end")) or parse_date(t.get("baseline_end"))
        t["_cs"] = cs
        t["_ce"] = ce

    preds_of = {}
    for d in dependencies:
        if d["type"] != "FS": continue
        preds_of.setdefault(d["successor_id"], []).append(d)

    for _ in range(50):
        changed = False
        for t in tasks:
            preds = preds_of.get(t["id"], [])
            if not preds or not t["_cs"] or not t["_ce"]: continue
            latest = None
            for d in preds:
                p = task_map.get(d["predecessor_id"])
                if not p or not p["_ce"]: continue
                candidate = p["_ce"] + timedelta(days=d["lag"])
                if latest is None or candidate > latest:
                    latest = candidate
            if latest and latest > t["_cs"]:
                delta = (latest - t["_cs"]).days
                t["_cs"] = latest
                t["_ce"] = t["_ce"] + timedelta(days=delta)
                changed = True
        if not changed: break

    for t in tasks:
        if t["_cs"] and t["_ce"]:
            t["cascade_start"] = t["_cs"].isoformat()
            t["cascade_end"] = t["_ce"].isoformat()
            be = parse_date(t.get("baseline_end"))
            if be: t["cascade_slip_days"] = (t["_ce"] - be).days
        t.pop("_cs", None)
        t.pop("_ce", None)
    return tasks


def compute_completion(tasks):
    if not tasks: return {"best": "", "likely": "", "worst": "", "avg_slip": 0}
    best_dates = [parse_date(t["baseline_end"]) for t in tasks if parse_date(t["baseline_end"])]
    likely_dates = [parse_date(t["cascade_end"]) for t in tasks if parse_date(t["cascade_end"])]
    slips = [t["slip_days"] for t in tasks if t["slip_days"] > 0]
    avg_slip = sum(slips) / len(slips) if slips else 0
    worst_candidates = []
    for t in tasks:
        be = parse_date(t["baseline_end"])
        if be: worst_candidates.append(be + timedelta(days=int(avg_slip)))
    return {
        "best": max(best_dates).isoformat() if best_dates else "",
        "likely": max(likely_dates).isoformat() if likely_dates else "",
        "worst": max(worst_candidates).isoformat() if worst_candidates else "",
        "avg_slip": round(avg_slip, 1),
    }


def generate_summaries(tasks, defects, history):
    if not DEEPSEEK_API_KEY: return {}

    def batch(items, kind):
        payload = []
        for it in items:
            h = history.get(it["id"], [])[:10]
            payload.append({
                "id": it["id"],
                "name": it.get("name") or it.get("title"),
                "status": it.get("status"),
                "history": [f"{u['timestamp']} {u['field']}: {u['old_value']} -> {u['new_value']} {u['notes']}" for u in h],
            })
            
        prompt = (
            f"You are a project controls analyst. For each {kind} below, write a one-sentence summary of the current situation based on its history. "
            "Focus on the ACTUAL CAUSES, BLOCKERS, and NOTES. Do not mention technical field names (like 'percent_complete' or 'forecast_end'). "
            "Summarize the narrative. If the item is complete, state if it was finished late or on time. "
            "Output strict JSON: {\"ID\": \"summary\", ...}. "
            f"DATA: {json.dumps(payload)}"
        )
        
        try:
            r = requests.post(
                "https://api.deepseek.com/chat/completions",
                headers={"Authorization": f"Bearer {DEEPSEEK_API_KEY}", "Content-Type": "application/json"},
                json={
                    "model": "deepseek-chat",
                    "messages": [{"role": "system", "content": "Output only valid JSON."}, {"role": "user", "content": prompt}],
                    "temperature": 0, "response_format": {"type": "json_object"},
                }, timeout=120,
            )
            if r.status_code != 200: return {}
            return json.loads(r.json()["choices"][0]["message"]["content"])
        except Exception: return {}

    out = {}
    out.update(batch(tasks, "task"))
    out.update(batch(defects, "defect"))
    print(f"  Generated {len(out)} summaries")
    return out


def compute_kpis(tasks, defects):
    total = len(tasks)
    complete = sum(1 for t in tasks if t["status"] == "Complete")
    in_progress = sum(1 for t in tasks if t["status"] == "In Progress")
    blocked = sum(1 for t in tasks if t["status"] == "Blocked")
    today = datetime.now(timezone.utc).date()
    overdue = 0
    for t in tasks:
        fe = parse_date(t.get("forecast_end"))
        if fe and fe < today and t["status"] != "Complete": overdue += 1
    positive_slips = [t["slip_days"] for t in tasks if t["slip_days"] > 0]
    avg_slip = round(sum(positive_slips) / len(positive_slips), 1) if positive_slips else 0

    open_defects = [d for d in defects if d["is_open"]]
    critical_defects = sum(1 for d in open_defects if d["severity"] == "Critical")

    health = "green"
    if critical_defects > 0 or overdue > 5: health = "red"
    elif overdue > 0 or avg_slip > 7: health = "amber"

    return {
        "total_tasks": total, "complete": complete, "in_progress": in_progress, "blocked": blocked,
        "overdue": overdue, "avg_slip": avg_slip, "open_defects": len(open_defects),
        "critical_defects": critical_defects, "health": health,
    }


def render(audience, role, config, kpis, completion, tasks, defects, dependencies, updates, history, summaries):
    os.makedirs(DASHBOARDS_DIR, exist_ok=True)
    env = Environment(loader=FileSystemLoader(TEMPLATES_DIR), autoescape=False)
    template = env.get_template("dashboard.html")
    data = {
        "config": {k: str(v) for k, v in config.items()},
        "kpis": kpis, "completion": completion, "tasks": tasks, "defects": defects,
        "dependencies": dependencies, "updates": updates[:50], "history": history, "summaries": summaries,
    }
    data_json = json.dumps(data, default=str).replace("<", "\\u003c").replace(">", "\\u003e")
    html = template.render(
        project_name=config.get("name", "Project"), client=config.get("client", ""),
        contract_id=config.get("contract_id", ""), generated_at=datetime.now().strftime("%Y-%m-%d %H:%M"),
        audience=role, data_json=data_json,
    )
    out_dir = os.path.join(DASHBOARDS_DIR, audience)
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "index.html")
    with open(out_path, "w", encoding="utf-8") as f: f.write(html)
    print(f"  Wrote {out_path} ({len(html)} bytes)")
    return out_path


def upload_to_drive(local_path, folder_id, filename):
    creds = get_credentials()
    service = build("drive", "v3", credentials=creds)
    last_error = None
    for attempt in range(3):
        try:
            query = f"name='{filename}' and '{folder_id}' in parents and trashed=false"
            existing = service.files().list(q=query, fields="files(id)").execute().get("files", [])
            media = MediaFileUpload(local_path, mimetype="text/html", resumable=False)
            if existing:
                service.files().update(fileId=existing[0]["id"], media_body=media).execute()
                print(f"  Drive updated: {filename}")
            else:
                meta = {"name": filename, "parents": [folder_id]}
                service.files().create(body=meta, media_body=media, fields="id").execute()
                print(f"  Drive created: {filename}")
            return True
        except Exception as e:
            last_error = e
            if attempt < 2: time.sleep(5)
    print(f"  Drive upload FAILED: {last_error}")
    return False


def main():
    print("=== generate.py ===")
    ss = connect()
    config = read_config(ss)

    raw = {}
    for tab in TABS: raw[tab] = read_tab(ss, tab)

    company_map, name_map = build_maps(raw["People"])
    tasks = build_tasks(raw["Tasks"], company_map, name_map)
    defects = build_defects(raw["Defects"], company_map, name_map)
    dependencies = build_dependencies(raw["Dependencies"])
    updates = build_updates(raw["Updates"])

    tasks = apply_cascade(tasks, dependencies)
    history = build_history(updates)
    summaries = generate_summaries(tasks, defects, history)
    folder_id = config.get("drive_dashboard_folder_id")

    def do_render(audience, role, t_list, d_list):
        kpis = compute_kpis(t_list, d_list)
        completion = compute_completion(t_list)
        p = render(audience, role, config, kpis, completion, t_list, d_list, dependencies, updates, history, summaries)
        if folder_id: upload_to_drive(p, folder_id, f"{audience}_dashboard.html")

    do_render("executive", "executive", tasks, defects)

    for co in ["TomatoFarms", "MTG", "Subcon"]:
        co_tasks = [t for t in tasks if t["owner_company"] == co]
        co_defects = [d for d in defects if co in d["assigned_company"]]
        if co_tasks or co_defects:
            do_render(co.lower().replace(" ", "_"), co.lower().replace(" ", "_"), co_tasks, co_defects)

    print("Done.")

if __name__ == "__main__":
    main()