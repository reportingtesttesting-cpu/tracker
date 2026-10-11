import os
import json
from datetime import datetime

import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from dotenv import load_dotenv


load_dotenv()

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

SERVICE_ACCOUNT_FILE = os.environ.get(
    "GOOGLE_SERVICE_ACCOUNT_JSON",
    os.path.join(os.path.dirname(__file__), "..", "secrets", "service-account.json"),
)

SHEET_ID = os.environ.get("SHEET_ID", "")

TABS = ["Config", "People", "Tasks", "Defects", "Dependencies", "Updates"]


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


def compute_kpis(tasks, defects):
    today = datetime.utcnow().date()
    total = len(tasks)

    def count_status(status):
        if total == 0:
            return 0
        return int((tasks["status"] == status).sum())

    complete = count_status("Complete")
    in_progress = count_status("In Progress")
    blocked = count_status("Blocked")

    overdue = 0
    slips = []
    for _, r in tasks.iterrows():
        be = parse_date(r.get("baseline_end"))
        fe = parse_date(r.get("forecast_end"))
        st = r.get("status")
        if be and fe:
            slips.append((fe - be).days)
        if fe and fe < today and st != "Complete":
            overdue += 1

    positive_slips = [s for s in slips if s > 0]
    avg_slip = round(sum(positive_slips) / len(positive_slips), 1) if positive_slips else 0

    if len(defects) > 0:
        open_mask = ~defects["status"].isin(["Closed", "Verified", "Rejected", "Duplicate"])
        open_defects = defects[open_mask]
        critical_defects = int((open_defects["severity"] == "Critical").sum()) if len(open_defects) else 0
    else:
        open_defects = pd.DataFrame()
        critical_defects = 0

    return {
        "total_tasks": total,
        "complete": complete,
        "in_progress": in_progress,
        "blocked": blocked,
        "overdue": overdue,
        "avg_slip": avg_slip,
        "open_defects": len(open_defects),
        "critical_defects": critical_defects,
    }


def main():
    print("=== generate.py ===")
    print("Sheet ID:", SHEET_ID)

    ss = connect()
    print("Connected:", ss.title)

    config = read_config(ss)
    print("Config keys:", list(config.keys()))

    data = {}
    for tab in TABS:
        df = read_tab(ss, tab)
        data[tab] = df
        print(f"  {tab}: {len(df)} rows")

    kpis = compute_kpis(data["Tasks"], data["Defects"])
    print("KPIs:")
    print(json.dumps(kpis, indent=2))


if __name__ == "__main__":
    main()