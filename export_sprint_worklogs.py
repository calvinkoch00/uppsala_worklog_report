import os
import sys
import requests
import pandas as pd
from requests.auth import HTTPBasicAuth
from datetime import datetime, timedelta
from dotenv import load_dotenv

load_dotenv()

JIRA_DOMAIN = os.getenv("JIRA_DOMAIN", "")
JIRA_DOMAIN = JIRA_DOMAIN.replace("https://", "").replace("http://", "").strip("/")
JIRA_EMAIL = os.getenv("JIRA_EMAIL")
JIRA_API_TOKEN = os.getenv("JIRA_API_TOKEN")

AUTH = HTTPBasicAuth(JIRA_EMAIL, JIRA_API_TOKEN)
HEADERS = {
    "Accept": "application/json",
    "Content-Type": "application/json"
}
CSV_DIR = "csv"
XLSX_FILE = "sprint_worklog_report.xlsx"

# Fixed team mapping for first and last names
KNOWN_AUTHORS = {
    "712020:1793fffe-a3b4-4b8b-b162-77b2e79e19da": {"firstname": "Calvin", "lastname": "Koch"},
    "712020:991ea2e9-c055-4d6e-a6b7-a44bccafe0a6": {"firstname": "Anton", "lastname": "Hedlund"},
    "712020:ba6fc434-9839-4353-a732-a8f4b60b9695": {"firstname": "Klas", "lastname": "Widebeck"},
    "712020:7e8593ee-4747-44a1-89c2-28e81af19f99": {"firstname": "Seleman", "lastname": "Hassan"},
    "712020:873506cb-9f09-47b4-8689-67e320ded9be": {"firstname": "Einar", "lastname": "Eriksson Wahlin"},
    "712020:7a155d42-06e0-4222-9d4e-f39dede68758": {"firstname": "Suchith", "lastname": "Sai Kurra"}
}


def parse_names(account_id, display_name):
    """Return explicit (firstname, lastname) from lookup or parse safely from display_name."""
    if account_id in KNOWN_AUTHORS:
        return KNOWN_AUTHORS[account_id]["firstname"], KNOWN_AUTHORS[account_id]["lastname"]
    parts = display_name.strip().split()
    if len(parts) >= 2:
        return parts[0], " ".join(parts[1:])
    elif len(parts) == 1:
        return parts[0], ""
    return "Unknown", "Unknown"


def extract_plain_text_from_adf(node):
    """Recursively extract plain text from an Atlassian Document Format (ADF) node."""
    if not node:
        return ""
    if isinstance(node, str):
        return node.strip()
    if isinstance(node, dict):
        if node.get("type") == "text":
            return node.get("text", "")
        content = node.get("content", [])
        extracted = [extract_plain_text_from_adf(child) for child in content]
        return " ".join([t for t in extracted if t]).strip()
    if isinstance(node, list):
        extracted = [extract_plain_text_from_adf(child) for child in node]
        return " ".join([t for t in extracted if t]).strip()
    return ""


def export_xlsx_to_csvs(xlsx_path=XLSX_FILE, output_dir=CSV_DIR):
    """Reads all sheets from the Excel file and saves each as an individual CSV."""
    if not os.path.exists(xlsx_path):
        print(f"Error: '{xlsx_path}' does not exist.")
        sys.exit(1)

    os.makedirs(output_dir, exist_ok=True)
    excel_file = pd.ExcelFile(xlsx_path)
    for sheet_name in excel_file.sheet_names:
        df = pd.read_excel(xlsx_path, sheet_name=sheet_name)
        csv_path = os.path.join(output_dir, f"{sheet_name}.csv")
        df.to_csv(csv_path, index=False)
        print(f"Exported: {csv_path}")


# Project timezone (CEST = UTC+2 in summer, CET = UTC+1 in winter)
TIMEZONE = "Europe/Stockholm"

def get_all_sprints():
    """Fetch all sprints and compute 2-week calendar boundaries in local project time."""
    sprints_raw = {}
    boards_url = f"https://{JIRA_DOMAIN}/rest/agile/1.0/board"
    try:
        boards_res = requests.get(boards_url, headers=HEADERS, auth=AUTH)
        if boards_res.status_code != 200:
            return []

        boards = boards_res.json().get("values", [])
        for board in boards:
            board_id = board.get("id")
            sprint_url = f"https://{JIRA_DOMAIN}/rest/agile/1.0/board/{board_id}/sprint"
            sprint_res = requests.get(sprint_url, headers=HEADERS, auth=AUTH)
            if sprint_res.status_code != 200:
                continue

            for sp in sprint_res.json().get("values", []):
                sprint_id = sp.get("id")
                if sprint_id not in sprints_raw:
                    sprints_raw[sprint_id] = {
                        "sprint_id": sprint_id,
                        "sprint_name": sp.get("name"),
                        "state": sp.get("state"),
                        "raw_start_date": sp.get("startDate")
                    }
    except Exception as e:
        print(f"Warning: Could not fetch agile sprints ({e}).")

    sprint_list = list(sprints_raw.values())
    if not sprint_list:
        return []

    # Sort sprints chronologically by raw start date
    sprint_list.sort(key=lambda s: pd.to_datetime(s.get("raw_start_date") or "2099-01-01", utc=True))

    for sprint in sprint_list:
        raw_start = sprint.get("raw_start_date")
        if raw_start:
            # Convert UTC timestamp to local project timezone first
            dt_local = pd.to_datetime(raw_start, utc=True).tz_convert(TIMEZONE)
            # Find Monday 00:00:00 of that week in local time
            monday_local = (dt_local - timedelta(days=dt_local.weekday())).replace(
                hour=0, minute=0, second=0, microsecond=0
            )
            # Exactly 2 weeks later: Sunday 23:59:59
            sunday_local = monday_local + timedelta(days=13, hours=23, minutes=59, seconds=59)

            # Convert back to UTC for consistent comparisons with Jira worklogs
            eff_start_utc = monday_local.tz_convert("UTC")
            eff_end_utc = sunday_local.tz_convert("UTC")
        else:
            eff_start_utc = None
            eff_end_utc = None
            monday_local = None
            sunday_local = None

        sprint["effective_start_dt"] = eff_start_utc
        sprint["effective_end_dt"] = eff_end_utc
        sprint["theoretical_start_dt"] = eff_start_utc
        sprint["theoretical_end_dt"] = eff_end_utc

        sprint["theoretical_start_date"] = monday_local.strftime("%Y-%m-%d %H:%M:%S") if monday_local else ""
        sprint["theoretical_end_date"] = sunday_local.strftime("%Y-%m-%d %H:%M:%S") if sunday_local else ""
        sprint["effective_start_date"] = monday_local.strftime("%Y-%m-%d %H:%M:%S") if monday_local else ""
        sprint["effective_end_date"] = sunday_local.strftime("%Y-%m-%d %H:%M:%S") if sunday_local else ""
        sprint.pop("raw_start_date", None)

    return sprint_list


def get_all_jira_issues():
    """Fetch ALL issues in Jira site using the /search/jql API."""
    issues = []
    next_page_token = None
    url = f"https://{JIRA_DOMAIN}/rest/api/3/search/jql"

    while True:
        payload = {
            "jql": "created is not EMPTY ORDER BY created DESC",
            "fields": ["key", "summary", "status", "issuetype", "timespent", "created", "updated"],
            "maxResults": 50
        }
        if next_page_token:
            payload["nextPageToken"] = next_page_token

        res = requests.post(url, headers=HEADERS, auth=AUTH, json=payload)
        if res.status_code != 200:
            print(f"Jira API Error ({res.status_code}): {res.text}")
            res.raise_for_status()

        data = res.json()
        issues.extend(data.get("issues", []))

        next_page_token = data.get("nextPageToken")
        if not next_page_token:
            break

    return issues


def get_worklogs_for_issue(issue_key):
    """Fetch raw worklog entries for a specific issue."""
    url = f"https://{JIRA_DOMAIN}/rest/api/3/issue/{issue_key}/worklog"
    res = requests.get(url, headers=HEADERS, auth=AUTH)
    if res.status_code != 200:
        return []
    return res.json().get("worklogs", [])


def match_sprint_by_date(started_dt, sprint_list):
    """Assign worklog to a sprint using clean 2-week boundary windows."""
    if started_dt is None or not sprint_list:
        return "Unassigned / Backlog"

    for sprint in sprint_list:
        s_start = sprint.get("effective_start_dt")
        s_end = sprint.get("effective_end_dt")
        if s_start and s_end and s_start <= started_dt <= s_end:
            return sprint["sprint_name"]

    # If work was logged prior to official sprint 1 start, attribute to Sprint 1
    if sprint_list and sprint_list[0].get("effective_start_dt"):
        if started_dt < sprint_list[0]["effective_start_dt"]:
            return sprint_list[0]["sprint_name"]

    return "Outside Sprints"


def build_star_report():
    print("Fetching sprint definitions and calculating gapless windows...")
    sprint_list = get_all_sprints()

    print("Fetching ALL issues for dim_issues...")
    all_issues = get_all_jira_issues()

    raw_timelogs = []
    dim_issues = {}
    dim_authors = {}

    # Pre-populate dim_authors with known team members
    for acc_id, name_data in KNOWN_AUTHORS.items():
        dim_authors[acc_id] = {
            "account_id": acc_id,
            "display_name": f"{name_data['firstname']} {name_data['lastname']}".strip(),
            "firstname": name_data["firstname"],
            "lastname": name_data["lastname"]
        }

    print(f"Processing {len(all_issues)} total issues...")
    for issue in all_issues:
        key = issue["key"]
        fields = issue.get("fields", {})
        total_time_spent = fields.get("timespent") or 0

        # Build full dim_issues table
        dim_issues[key] = {
            "issue_key": key,
            "summary": fields.get("summary", ""),
            "issue_type": fields.get("issuetype", {}).get("name", "Unknown"),
            "status": fields.get("status", {}).get("name", "Unknown"),
            "timespent_hours": round(total_time_spent / 3600.0, 2),
            "created": fields.get("created", "")[:19].replace("T", " "),
            "updated": fields.get("updated", "")[:19].replace("T", " "),
            "issue_url": f"https://{JIRA_DOMAIN}/browse/{key}"
        }

        # Query worklogs if work was logged
        if total_time_spent > 0:
            worklogs = get_worklogs_for_issue(key)
            for wl in worklogs:
                worklog_id = str(wl.get("id"))
                author_data = wl.get("author", {})
                author_id = author_data.get("accountId", "unknown")
                author_name = author_data.get("displayName", "Unknown")

                if author_id not in dim_authors:
                    fn, ln = parse_names(author_id, author_name)
                    dim_authors[author_id] = {
                        "account_id": author_id,
                        "display_name": author_name,
                        "firstname": fn,
                        "lastname": ln
                    }

                started_raw = wl.get("started")
                started_dt = pd.to_datetime(started_raw, utc=True) if started_raw else None

                if started_dt is not None:
                    iso_year, iso_week, _ = started_dt.isocalendar()
                    calendar_week = f"{iso_year}-W{iso_week:02d}"
                    clean_timestamp = started_dt.strftime("%Y-%m-%d %H:%M:%S")
                    clean_date = started_dt.strftime("%Y-%m-%d")
                else:
                    calendar_week = "Unknown"
                    clean_timestamp = ""
                    clean_date = ""

                seconds = wl.get("timeSpentSeconds", 0)
                hours = round(seconds / 3600.0, 2)
                sprint_assigned = match_sprint_by_date(started_dt, sprint_list)

                raw_timelogs.append({
                    "worklog_id": worklog_id,
                    "timestamp": clean_timestamp,
                    "date": clean_date,
                    "calendar_week": calendar_week,
                    "sprint_by_date": sprint_assigned,
                    "worklog_author": author_name,
                    "author_id": author_id,
                    "issue_key": key,
                    "hours_logged": hours,
                    "comment": extract_plain_text_from_adf(wl.get("comment"))
                })

    df_facts = pd.DataFrame(raw_timelogs)
    df_issues = pd.DataFrame(list(dim_issues.values()))
    df_authors = pd.DataFrame(list(dim_authors.values()))

    # Clean dim_sprints: omit internal datetime objects
    sprints_clean = [
        {k: v for k, v in sp.items() if not k.endswith("_dt")}
        for sp in sprint_list
    ]
    df_sprints = pd.DataFrame(sprints_clean)

    # Build Overviews
    if not df_facts.empty:
        df_overview_week = df_facts.pivot_table(
            index=["calendar_week", "worklog_author"],
            values="hours_logged",
            aggfunc="sum"
        ).reset_index().rename(columns={"hours_logged": "total_hours"})

        df_overview_sprint = df_facts.pivot_table(
            index=["sprint_by_date", "worklog_author"],
            values="hours_logged",
            aggfunc="sum"
        ).reset_index().rename(columns={"hours_logged": "total_hours"})
    else:
        df_overview_week = pd.DataFrame(columns=["calendar_week", "worklog_author", "total_hours"])
        df_overview_sprint = pd.DataFrame(columns=["sprint_by_date", "worklog_author", "total_hours"])

    # Write multi-sheet Excel
    with pd.ExcelWriter(XLSX_FILE, engine="openpyxl") as writer:
        df_overview_week.to_excel(writer, sheet_name="Overview_Weekly", index=False)
        df_overview_sprint.to_excel(writer, sheet_name="Overview_Sprint", index=False)
        df_facts.to_excel(writer, sheet_name="Fact_Timelogs", index=False)
        df_sprints.to_excel(writer, sheet_name="Dim_Sprints", index=False)
        df_issues.to_excel(writer, sheet_name="Dim_Issues", index=False)
        df_authors.to_excel(writer, sheet_name="Dim_Authors", index=False)

    print(f"Generated {XLSX_FILE}")
    export_xlsx_to_csvs(XLSX_FILE, CSV_DIR)


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1].lower() in ["update-csv", "update_csv", "csv"]:
        print(f"Reading {XLSX_FILE} and synchronizing CSVs...")
        export_xlsx_to_csvs(XLSX_FILE, CSV_DIR)
    else:
        build_star_report()