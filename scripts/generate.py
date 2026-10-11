import os
import json
import tempfile
from datetime import datetime, timezone

import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from dotenv import load_dotenv
from jinja2 import Environment, FileSystemLoader


load_dotenv()

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

_raw_creds = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()

if _raw_creds.startswith("{"):
    _tmp = tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False)
    _tmp.write(_raw_creds)
    _tmp.close()
    SERVICE_ACCOUNT_FILE = _tmp.name
else:
    SERVICE_ACCOUNT_FILE = _raw_creds or os.path.join(
        os.path.dirname(__file__), "..", "secrets", "service-account.json"
    )

SHEET_ID = os.environ.get("SHEET_ID", "")

TABS = ["Config", "People", "Tasks", "Defects", "Dependencies", "Updates"]

TEMPLATES_DIR = os.path.join(os.path.dirname(__file__), "..", "templates")
DASHBOARDS_DIR = os.path.join(os.path.dirname(__file__), "..", "dashboards")

CLOSED_DEFECT_STATES = ["Closed", "Verified", "Rejected", "Duplicate"]

REDIRECTS_CONTENT = (
    "/executive/*     200!  Role=executive\n"
    "/tomato_farms/*  200!  Role=tomato_farms\n"
    "/*               /  200\n"
)


def connect():
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
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


def build_tasks(tasks_df, people_df):
    company_map = {}
    if len(people_df) > 0:
        for _, p in people_df.iterrows():
            company_map[p.get("person_id", "")] = p.get("company", "")

    out = []
    for _, r in tasks_df.iterrows():
        if not r.get("task_id"):
            continue
        owner_id = str(r.get("owner_id", ""))
        be = parse_date(r.get("baseline_end"))
        fe = parse_date(r.get("forecast_end"))
        slip = (fe - be).days if (be and fe) else 0
        out.append({
            "id": str(r.get("task_id")),
            "name": str(r.get("name", "")),
            "description": str(r.get("description", "")),
            "owner_id": owner_id,
            "owner_company": company_map.get(owner_id, ""),
            "status": str(r.get("status", "Not Started")),
            "percent": int(float(r.get("percent_complete", 0) or 0)),
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
        })
    return out


def build_defects(defects_df, people_df):
    company_map = {}
    if len(people_df) > 0:
        for _, p in people_df.iterrows():
            company_map[p.get("person_id", "")] = p.get("company", "")

    today = datetime.now(timezone.utc).date()
    out = []
    for _, r in defects_df.iterrows():
        if not r.get("defect_id"):
            continue
        assigned = str(r.get("assigned_to", ""))
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
            "assigned_to": assigned,
            "assigned_company": company_map.get(assigned, ""),
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
    return out[:200]


def compute_kpis(tasks, defects):
    total = len(tasks)
    complete = sum(1 for t in tasks if t["status"] == "Complete")
    in_progress = sum(1 for t in tasks if t["status"] == "In Progress")
    blocked = sum(1 for t in tasks if t["status"] == "Blocked")
    today = datetime.now(timezone.utc).date()
    overdue = 0
    for t in tasks:
        fe = parse_date(t.get("forecast_end"))
        if fe and fe < today and t["status"] != "Complete":
            overdue += 1
    positive_slips = [t["slip_days"] for t in tasks if t["slip_days"] > 0]
    avg_slip = round(sum(positive_slips) / len(positive_slips), 1) if positive_slips else 0

    open_defects = [d for d in defects if d["is_open"]]
    critical_defects = sum(1 for d in open_defects if d["severity"] == "Critical")

    health = "green"
    if critical_defects > 0 or overdue > 5:
        health = "red"
    elif overdue > 0 or avg_slip > 7:
        health = "amber"

    return {
        "total_tasks": total,
        "complete": complete,
        "in_progress": in_progress,
        "blocked": blocked,
        "overdue": overdue,
        "avg_slip": avg_slip,
        "open_defects": len(open_defects),
        "critical_defects": critical_defects,
        "health": health,
    }


def render(audience, role, config, kpis, tasks, defects, dependencies, updates):
    env = Environment(loader=FileSystemLoader(TEMPLATES_DIR), autoescape=False)
    template = env.get_template("dashboard.html")

    data = {
        "config": {k: str(v) for k, v in config.items()},
        "kpis": kpis,
        "tasks": tasks,
        "defects": defects,
        "dependencies": dependencies,
        "updates": updates,
    }

    data_json = json.dumps(data, default=str).replace("<", "\\u003c").replace(">", "\\u003e")

    html = template.render(
        project_name=config.get("name", "Project"),
        client=config.get("client", ""),
        contract_id=config.get("contract_id", ""),
        generated_at=datetime.now().strftime("%Y-%m-%d %H:%M"),
        audience=role,
        data_json=data_json,
    )

    out_dir = os.path.join(DASHBOARDS_DIR, audience)
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "index.html")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"  Wrote {out_path} ({len(html)} bytes)")


def render_root():
    env = Environment(loader=FileSystemLoader(TEMPLATES_DIR), autoescape=False)
    template = env.get_template("root.html")
    html = template.render()
    out_path = os.path.join(DASHBOARDS_DIR, "index.html")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"  Wrote {out_path} ({len(html)} bytes)")


def write_redirects():
    out_path = os.path.join(DASHBOARDS_DIR, "_redirects")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(REDIRECTS_CONTENT)
    print(f"  Wrote {out_path}")


def main():
    print("=== generate.py ===")
    print("Sheet ID:", SHEET_ID)

    ss = connect()
    print("Connected:", ss.title)

    config = read_config(ss)

    raw = {}
    for tab in TABS:
        raw[tab] = read_tab(ss, tab)
        print(f"  {tab}: {len(raw[tab])} rows")

    people = raw["People"]
    tasks = build_tasks(raw["Tasks"], people)
    defects = build_defects(raw["Defects"], people)
    dependencies = build_dependencies(raw["Dependencies"])
    updates = build_updates(raw["Updates"])
    kpis = compute_kpis(tasks, defects)

    print("KPIs:", json.dumps(kpis))

    print("Rendering root...")
    render_root()

    print("Writing redirects...")
    write_redirects()

    print("Rendering audiences...")
    render("executive", "executive", config, kpis, tasks, defects, dependencies, updates)
    render("tomato_farms", "tomato_farms", config, kpis, tasks, defects, dependencies, updates)

    print("Done.")


if __name__ == "__main__":
    main()