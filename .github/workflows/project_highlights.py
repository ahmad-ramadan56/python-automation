import requests
import pandas as pd
import os
from datetime import datetime
from msal import ConfidentialClientApplication


# ============ CONFIGURATION ============

# Fetch credentials injected by GitHub Actions at runtime

TENANT_ID = os.environ.get('TENANT_ID')
CLIENT_ID = os.environ.get('CLIENT_ID')
CLIENT_SECRET = os.environ.get('CLIENT_SECRET')

# Fallback check to make troubleshooting easier

if not all([TENANT_ID, CLIENT_ID, CLIENT_SECRET]):
    raise ValueError("Missing critical Azure credentials in environment variables!")


SITE_NAME = "pwrapps"

OLD_LIST = "Archived data : Pulse Dashboard Weekly Update"
NEW_LIST = "op_weekly_project_status_update"
HIGHLIGHTS_LIST = "op_ProjectHighlights"


# ============ METRICS ============

METRICS = [
    {"name": "Customer", "comment": "CustomerComment"},  # Match API, not UI
    {"name": "Safety", "comment": "SafetyComment"},
    {"name": "Quality", "comment": "QualityComment"},
    {"name": "Delivery", "comment": "DeliveryComment"},
    {"name": "Capacity", "comment": "CapacityComment"},
    {"name": "Material", "comment": "MaterialComment"},
    {"name": "Components", "comment": "ComponentsComment"},
    {"name": "Finance", "comment": "FinanceComment"},
    {"name": "Headcount", "comment": "HeadcountComment"},
    {"name": "SickLeave", "comment": "SickLeaveComment"},
    {"name": "Overtime", "comment": "OvertimeComment"},
    {"name": "Turnover", "comment": "TurnoverComment"}
]


# ============ AUTH ============

def get_access_token():
    """Get access token using MSAL"""
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = ConfidentialClientApplication(
        CLIENT_ID,
        authority=authority,
        client_credential=CLIENT_SECRET
    )
    scopes = ["https://graph.microsoft.com/.default"]
    result = app.acquire_token_for_client(scopes=scopes)
    
    if "access_token" in result:
        return result["access_token"]
    raise Exception(f"Failed to get token: {result}")


# ============ GRAPH HELPERS ============

def graph_headers(token, json_type=True):
    """Create headers for Graph API requests"""
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "ConsistencyLevel": "eventual"
    }
    if json_type:
        headers["Content-Type"] = "application/json"
    return headers


def get_site_id(token):
    """Get SharePoint site ID"""
    headers = graph_headers(token, False)
    url = (
        f"https://graph.microsoft.com/v1.0/"
        f"sites/euromaintab.sharepoint.com:/sites/{SITE_NAME}:/"
        f"?$select=id,webUrl"
    )
    r = requests.get(url, headers=headers)
    
    if r.status_code == 200:
        data = r.json()
        print("Resolved site:", data.get("webUrl"))
        return data["id"]
    
    raise Exception(f"Failed to get site ID: {r.status_code} - {r.text}")


def get_list_id(token, site_id, list_name):
    """Get list ID by name"""
    headers = graph_headers(token, False)
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists"
    r = requests.get(url, headers=headers)
    
    if r.status_code != 200:
        raise Exception(f"List fetch failed: {r.status_code} - {r.text}")
    
    lists = r.json().get("value", [])
    
    for lst in lists:
        if (lst.get("displayName") == list_name or lst.get("name") == list_name):
            return lst["id"]
    
    raise Exception(f"List not found: {list_name}")


def get_list_items(token, site_id, list_name):
    """Get all items from a SharePoint list"""
    try:
        list_id = get_list_id(token, site_id, list_name)
        headers = graph_headers(token, False)
        items = []
        url = (
            f"https://graph.microsoft.com/v1.0/"
            f"sites/{site_id}/lists/{list_id}/items"
            f"?expand=fields"
        )
        
        while url:
            r = requests.get(url, headers=headers)
            
            if r.status_code != 200:
                raise Exception(f"{r.status_code} - {r.text}")
            
            data = r.json()
            
            for item in data.get("value", []):
                items.append(item["fields"])
            
            url = data.get("@odata.nextLink")
        
        return pd.DataFrame(items)
    
    except Exception as e:
        print(f"Error reading list '{list_name}': {e}")
        return pd.DataFrame()


# ============ ANALYSIS ============

def are_comments_similar(comment1, comment2):
    """Check if two comments are similar"""
    if pd.isna(comment1) or pd.isna(comment2):
        return False
    
    c1 = str(comment1).lower().strip()
    c2 = str(comment2).lower().strip()
    
    if not c1 or not c2:
        return False
    
    # Exact match
    if c1 == c2:
        return True
    
    # Check if one contains the other
    if c1 in c2 or c2 in c1:
        return True
    
    # Check word overlap
    words1 = set(c1.split())
    words2 = set(c2.split())
    
    if len(words1) == 0 or len(words2) == 0:
        return False
    
    overlap = len(words1.intersection(words2))
    similarity = overlap / max(len(words1), len(words2))
    
    return similarity > 0.5  # 50% word overlap


def build_highlight(row, metric, comment_field, status, value):
    """Build a single highlight dictionary with correct data types"""
    
    # Safely get string values
    def safe_str(field, default=''):
        val = row.get(field, default)
        if pd.notna(val):
            return str(val).strip()
        return default
    
    # Safely get integer values
    def safe_int(field, default=0):
        val = row.get(field, default)
        if pd.notna(val):
            try:
                return int(float(val))
            except (ValueError, TypeError):
                return default
        return default
    
    project = safe_str('Project', 'Unknown')
    week = safe_int('Week', 0)
    year = safe_int('Year', 0)
    group = safe_str('Group', '')
    comment = safe_str(comment_field, '')
    
    icon = "🔴" if status == "Red" else "🟡"
    
    return {
        "fields": {
            "Title": f"{project}-W{week}-{metric}",
            "Project": project if project else "Unknown",
            "Week": str(week),  # CHANGED: Convert to string
            "Year": year if year else 0,  # Number type - keep as int
            "Group": group if group else "",
            "MetricType": metric,
            "Status": status,
            "StatusValue": value,
            "Count": 1,  # Number type - keep as int
            "Comment": comment if comment else "",
            "HighlightText": f"{icon} {project} (Week {week}): {metric} is {status.upper()}" + (f" - {comment}" if comment else ""),
            "Priority": "High" if status == "Red" else "Medium",
            "Category": f"{status} Alert",
            "DateGenerated": datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ"),  # CHANGED: Proper format
            "IsActive": True
            
        }
    }

def build_runs(rows):
    """
    rows: list of dicts
    Expected keys: Year, Week, Status, Comment
    """

    # Sort by year, then week
    rows = sorted(rows, key=lambda x: (int(x["Year"]), int(x["Week"])))

    runs = []
    current_run = None

    for row in rows:
        year = int(row["Year"])
        week = int(row["Week"])
        status = row["Status"]
        comment = (row.get("Comment") or "").strip()

        if current_run is None:
            current_run = {
                "status": status,
                "weeks": [(year, week)],
                "comments": [comment]
            }
            continue

        prev_year, prev_week = current_run["weeks"][-1]

        is_consecutive = (
            (year == prev_year and week == prev_week + 1)
            or
            (year == prev_year + 1 and prev_week >= 52 and week == 1)
        )

        if status == current_run["status"] and is_consecutive:
            current_run["weeks"].append((year, week))
            current_run["comments"].append(comment)
        else:
            runs.append(current_run)
            current_run = {
                "status": status,
                "weeks": [(year, week)],
                "comments": [comment]
            }

    if current_run:
        runs.append(current_run)

    return runs


def analyze_metrics(df):
    highlights = []

    # Normalize data
    df = df.copy()
    df["Week"] = df["Week"].astype(int)
    df["Year"] = df["Year"].astype(int)

    for metric in METRICS:
        metric_name = metric["name"]
        comment_field = metric["comment"]

        if metric_name not in df.columns or comment_field not in df.columns:
            continue

        for project in df["Project"].dropna().unique():
            pdata = df[df["Project"] == project]

            rows = []
            for _, r in pdata.iterrows():
                status_val = r.get(metric_name)

                if status_val == 1:
                    status = "RED"
                elif status_val == 2:
                    status = "YELLOW"
                else:
                    continue  # ignore GREEN

                rows.append({
                    "Year": int(r["Year"]),
                    "Week": int(r["Week"]),
                    "Status": status,
                    "Comment": str(r.get(comment_field, "")).strip(),
                    "Group": str(r.get("Group", ""))
                })

            if len(rows) < 2:
                continue

            rows = sorted(rows, key=lambda x: (x["Year"], x["Week"]))
            runs = build_runs(rows)

            for run in runs:
                # FIXED: Access the dict correctly
                weeks_count = len(run["weeks"])  # Number of weeks in this run
                status = run["status"]  # Status for this run
                comments = run["comments"]  # List of comments
                unique_comments = list(dict.fromkeys([c for c in comments if c]))  # Remove duplicates, keep order
                
                # Get the last week's data
                last_year, last_week = run["weeks"][-1]
                last_comment = comments[-1] if comments else ""
                
                # Get group from the last matching row
                last_row = [r for r in rows if r["Year"] == last_year and r["Week"] == last_week]
                group = last_row[0]["Group"] if last_row else ""

                rule = None
                text = None

                # 🔴 Rule A: RED for 2+ weeks with SAME comment
                if status == "RED" and weeks_count >= 2 and len(unique_comments) == 1:
                    rule = "A"
                    text = (
                        f"{project} – {metric_name} reported RED 🔴 "
                        f"for {weeks_count} consecutive weeks: {unique_comments[0]}"
                    )

                # 🟡 Rule B: YELLOW for 3+ weeks with SAME comment
                elif status == "YELLOW" and weeks_count >= 3 and len(unique_comments) == 1:
                    rule = "B"
                    text = (
                        f"{project} – {metric_name} reported YELLOW 🟡 "
                        f"for {weeks_count} consecutive weeks: {unique_comments[0]}"
                    )

                # 🔴 Rule C: RED for 3+ weeks with DIFFERENT comments
                elif status == "RED" and weeks_count >= 3 and len(unique_comments) > 1:
                    rule = "C"
                    joined = "; ".join(unique_comments)
                    text = (
                        f"{project} – {metric_name} reported RED 🔴 "
                        f"for {weeks_count} consecutive weeks with evolving issues: {joined}"
                    )

                if rule:
                    highlights.append({
                        "fields": {
                            "Title": f"{project}-W{last_week}-{metric_name}-{rule}",
                            "Project": project,
                            "Week": str(last_week),
                            "Year": last_year,
                            "Group": group,
                            "MetricType": metric_name,
                            "Status": status.title(),
                            "StatusValue": 1 if status == "RED" else 2,
                            "Count": weeks_count,
                            "Comment": last_comment,
                            "HighlightText": text,
                            "Priority": "High" if status == "RED" else "Medium",
                            "Category": "Weekly Highlight",
                            "DateGenerated": datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ"),
                            "IsActive": True
                        }
                    })

    # ---------------- LIMIT TO TOP 10 ----------------
    # Sort by: Red first, then by count (descending), then by project name
    highlights = sorted(
        highlights,
        key=lambda h: (
            0 if h["fields"]["Status"] == "Red" else 1,  # Red first
            -h["fields"]["Count"],  # More weeks = higher priority
            h["fields"]["Project"]  # Alphabetical
        )
    )

    return highlights[:10]
# ============ WRITE BACK ============

def delete_old_highlights(token, site_id):
    """Delete ALL items from ProjectHighlights list"""
    try:
        list_id = get_list_id(token, site_id, HIGHLIGHTS_LIST)
        headers = graph_headers(token)

        # Get all items
        url = (
            f"https://graph.microsoft.com/v1.0/"
            f"sites/{site_id}/lists/{list_id}/items"
        )
        r = requests.get(url, headers=headers)

        if r.status_code != 200:
            raise Exception(r.text)

        all_items = r.json().get("value", [])

        if len(all_items) == 0:
            print("   No old items to delete")
            return

        # Delete each item
        count = 0
        for item in all_items:
            item_id = item["id"]
            delete_url = (
                f"https://graph.microsoft.com/v1.0/"
                f"sites/{site_id}/lists/{list_id}/items/{item_id}"
            )
            delete_response = requests.delete(delete_url, headers=headers)

            if delete_response.status_code in [200, 204]:
                count += 1

        print(f"   ✓ Deleted {count} old highlights")

    except Exception as e:
        print(f"   Delete error: {e}")


def write_highlights(token, site_id, highlights):
    """Write new highlights to SharePoint"""
    try:
        list_id = get_list_id(token, site_id, HIGHLIGHTS_LIST)
        headers = graph_headers(token)
        url = (
            f"https://graph.microsoft.com/v1.0/"
            f"sites/{site_id}/lists/{list_id}/items"
        )

        count = 0
        errors = 0
        
        for i, h in enumerate(highlights, 1):
            r = requests.post(url, headers=headers, json=h)
            if r.status_code == 201:
                count += 1
                if count % 100 == 0:  # Progress indicator
                    print(f"  Created {count}/{len(highlights)}...")
            else:
                errors += 1
                if errors <= 3:  # Only show first 3 errors
                    print(f"Create failed: {r.text}")
        
        print(f"\n✓ Created {count} highlights")
        if errors > 0:
            print(f"⚠ {errors} items failed to create")
    
    except Exception as e:
        print(f"Write error: {e}")


# ============ MAIN ============

def main():
    """Main execution"""
    print("\n" + "="*50)
    print("   PROJECT HIGHLIGHTS ANALYZER")
    print("="*50 + "\n")
    
    try:
        print("1. Getting authentication token...")
        token = get_access_token()
        print("   ✓ Token obtained\n")
        
        print("2. Resolving SharePoint site...")
        site_id = get_site_id(token)
        print(f"   ✓ Site ID: {site_id}\n")
        
        print("3. Reading data from SharePoint lists...")
        print(f"   - Old list: {OLD_LIST}")
        old_df = get_list_items(token, site_id, OLD_LIST)
        print(f"     ✓ {len(old_df)} rows")
        
        print(f"   - New list: {NEW_LIST}")
        new_df = get_list_items(token, site_id, NEW_LIST)
        print(f"     ✓ {len(new_df)} rows\n")
        
        # Combine data
        data = pd.concat([old_df, new_df], ignore_index=True)
        print(f"4. Combined data: {len(data)} total rows\n")
        
        # Show column names for debugging
        print("   Columns in data:")
        for col in sorted(data.columns)[:15]:  # Show first 15
            print(f"     - {col}")
        if len(data.columns) > 15:
            print(f"     ... and {len(data.columns) - 15} more\n")
        else:
            print()
        
        # Clean data - handle NaN values
        if 'Year' in data.columns:
            data['Year'] = data['Year'].fillna(0)
        if 'Week' in data.columns:
            data['Week'] = data['Week'].fillna(0)
        if 'Project' in data.columns:
            data['Project'] = data['Project'].fillna('Unknown')
        if 'Group' in data.columns:
            data['Group'] = data['Group'].fillna('')
        
        print("5. Analyzing metrics for highlights...")
        highlights = analyze_metrics(data)
        print(f"\n   ✓ Found {len(highlights)} highlights\n")
        
        if highlights:
            print("6. Deleting old highlights...")
        delete_old_highlights(token, site_id)
        print()

        if highlights:
            print("7. Writing new highlights to SharePoint...")
            write_highlights(token, site_id, highlights)
        else:
            print("   No new highlights to write\n")
        
        print("\n" + "="*50)
        print("   ✅ ANALYSIS COMPLETE!")
        print("="*50 + "\n")
    
    except Exception as e:
        print("\n" + "="*50)
        print("   ❌ ERROR OCCURRED")
        print("="*50)
        print(f"\n{e}\n")
        import traceback
        traceback.print_exc()
        print()


if __name__ == "__main__":
    main()
