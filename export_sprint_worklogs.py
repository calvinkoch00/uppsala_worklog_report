import os
import requests
import pandas as pd
from requests.auth import HTTPBasicAuth

# 1. Configuration (Loaded from Environment Variables)
JIRA_DOMAIN = os.getenv("JIRA_DOMAIN")        # e.g., "your-company.atlassian.net"
JIRA_EMAIL = os.getenv("JIRA_EMAIL")          # e.g., "your-email@example.com"
JIRA_API_TOKEN = os.getenv("JIRA_API_TOKEN")  # API token from id.atlassian.com

AUTH = HTTPBasicAuth(JIRA_EMAIL, JIRA_API_TOKEN)
HEADERS = {"Accept": "application/json"}

def get_issues_with_worklogs(jql_query):
    """Fetch issues matching JQL using the new Jira Cloud /search/jql API."""
    issues = []
    next_page_token = None
    url = f"https://{JIRA_DOMAIN}/rest/api/3/search/jql"

    while True:
        payload = {
            "jql": jql_query,
            "fields": ["key", "summary", "sprint", "customfield_10020"],
            "maxResults": 50
        }
        if next_page_token:
            payload["nextPageToken"] = next_page_token

        res = requests.post(url, headers={"Accept": "application/json", "Content-Type": "application/json"}, auth=AUTH, json=payload)
        res.raise_for_status()
        data = res.json()

        issues.extend(data.get("issues", []))

        # Check if another page exists
        next_page_token = data.get("nextPageToken")
        if not next_page_token:
            break

    return issues

def extract_sprint_name(issue_fields):
    """Extract sprint name whether it's under 'sprint' or custom fields."""
    sprint_data = issue_fields.get("sprint") or issue_fields.get("customfield_10020")
    if isinstance(sprint_data, list) and sprint_data:
        # Most recent sprint on the issue
        return sprint_data[-1].get("name", "No Sprint")
    elif isinstance(sprint_data, dict):
        return sprint_data.get("name", "No Sprint")
    return "No Sprint"

def get_worklogs_for_issue(issue_key):
    """Retrieve individual worklogs for a given issue."""
    url = f"https://{JIRA_DOMAIN}/rest/api/3/issue/{issue_key}/worklog"
    res = requests.get(url, headers=HEADERS, auth=AUTH)
    if res.status_code != 200:
        return []
    return res.json().get("worklogs", [])

def build_report():
    # Example JQL: Adjust to target specific projects or recent dates
    jql = "timespent > 0 AND worklogDate >= -30d ORDER BY updated DESC"
    issues = get_issues_with_worklogs(jql)
    
    rows = []
    for issue in issues:
        key = issue["key"]
        sprint_name = extract_sprint_name(issue["fields"])
        worklogs = get_worklogs_for_issue(key)

        for wl in worklogs:
            author = wl.get("author", {}).get("displayName", "Unknown")
            seconds = wl.get("timeSpentSeconds", 0)
            hours = round(seconds / 3600.0, 2)
            started_date = wl.get("started", "")[:10]

            rows.append({
                "Sprint": sprint_name,
                "Worklog Author": author,
                "Hours Logged": hours,
                "Date": started_date,
                "Issue": key
            })

    if not rows:
        print("No worklogs found for the specified period.")
        return

    df = pd.DataFrame(rows)

    # Summary table: Total hours grouped by Sprint and Author
    summary = df.pivot_table(
        index=["Sprint", "Worklog Author"],
        values="Hours Logged",
        aggfunc="sum"
    ).reset_index()

    # Write both raw logs and aggregated summary to Excel
    output_file = "sprint_worklog_report.xlsx"
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Sprint Summary", index=False)
        df.to_excel(writer, sheet_name="Raw Worklogs", index=False)

    print(f"Report successfully generated: {output_file}")

if __name__ == "__main__":
    build_report()
