import requests
import openpyxl
from msal import PublicClientApplication
from urllib.parse import quote

# ==== CONFIGURATION ====
CANVAS_API_BASE_URL = "https://[your instance].instructure.com/api/v1" # Replace with your Canvas instance URL
CANVAS_API_TOKEN = "" # Your Canvas API token here
CANVAS_ACCOUNT_ID = 1  # Usually 1 for root account
CANVAS_PAGE_SIZE = 100
CANVAS_TERM_ID = 328 # Replace with your term ID
CANVAS_MATCH_STRING = "2025/Summer" # Replace with your desired match string for better filtering

GRAPH_CLIENT_ID = "" # Your Azure AD application client ID here
GRAPH_TENANT_ID = "" # Your Azure AD tenant ID here
GRAPH_SCOPES = ["Group.Read.All", "User.Read"]
# =======================

def get_all_canvas_courses():
    headers = {"Authorization": f"Bearer {CANVAS_API_TOKEN}"}
    url = f"{CANVAS_API_BASE_URL}/accounts/{CANVAS_ACCOUNT_ID}/courses?enrollment_term_id={CANVAS_TERM_ID}&per_page={CANVAS_PAGE_SIZE}"
    courses = []

    while url:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        page_courses = response.json()
        courses.extend(page_courses)

        print(f"📄 Fetched {len(page_courses)} courses. Total: {len(courses)}")

        links = response.headers.get("Link", "")
        next_url = None
        for link in links.split(","):
            if 'rel="next"' in link:
                next_url = link[link.find("<") + 1:link.find(">")]
        url = next_url

    print(f"✅ Total Canvas courses fetched: {len(courses)}")
    return courses

def filter_courses(courses, match_string={CANVAS_MATCH_STRING}):
    filtered = [
        course for course in courses
        if isinstance(course.get("sis_course_id"), str) and match_string in course["sis_course_id"]
    ]
    print(f"🎯 Courses matching '{match_string}': {len(filtered)}")
    return filtered

def authenticate_graph():
    app = PublicClientApplication(GRAPH_CLIENT_ID, authority=f"https://login.microsoftonline.com/{GRAPH_TENANT_ID}")
    flow = app.initiate_device_flow(scopes=GRAPH_SCOPES)
    print("🔑 Please authenticate:")
    print(flow["message"])

    result = app.acquire_token_by_device_flow(flow)
    if "access_token" in result:
        print("✅ Graph authentication successful.")
        return result["access_token"]
    else:
        raise Exception("Authentication failed: " + str(result))

def find_team_by_display_name(display_name, access_token):
    headers = {
        "Authorization": f"Bearer {access_token}",
        "ConsistencyLevel": "eventual"
    }

    safe_display_name = quote(f"startswith(displayName,'{display_name[:60]}')")

    query = f"https://graph.microsoft.com/v1.0/groups?$filter={safe_display_name}"
    print(f"📡 Graph API request: {query}")

    try:
        response = requests.get(query, headers=headers)
        response.raise_for_status()
        results = response.json().get("value", [])

        for team in results:
            if display_name in team.get("displayName", ""):
                return {
                    "team_id": team["id"],
                    "team_name": team["displayName"],
                    "matched_by": "sis_course_id"
                }

        print(f"❌ No match by displayName: {display_name}")
        return None

    except Exception as e:
        print(f"⚠️ Error during Graph API search: {e}")
        return None

def get_team_web_url(team_id, access_token):
    url = f"https://graph.microsoft.com/v1.0/teams/{team_id}"
    headers = {
        "Authorization": f"Bearer {access_token}"
    }

    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        team = response.json()
        return team.get("webUrl")
    except Exception as e:
        print(f"⚠️ Failed to get team webUrl for {team_id}: {e}")
        return None

def find_team_fallback(course, access_token):
    sis_id = course.get("sis_course_id", "")
    course_name = course.get("name", "")

    team = find_team_by_display_name(sis_id, access_token)

    if not team:
        team = find_team_by_display_name(course_name, access_token)
        if team:
            team["matched_by"] = "name"

    return team

def add_teams_redirect_tool(course_id, team_link):
    url = f"{CANVAS_API_BASE_URL}/courses/{course_id}/external_tools"
    headers = {
        "Authorization": f"Bearer {CANVAS_API_TOKEN}",
        "Content-Type": "application/json"
    }

    payload = {
        "name": "Microsoft Teams Class",
        "consumer_key": "key",
        "shared_secret": "key",
        "privacy_level": "anonymous",
        "url": "https://www.edu-apps.org/redirect",
        "course_navigation": {
            "enabled": True,
            "visibility": "public",
            "label": "Microsoft Teams Class",
            "selection_width": 800,
            "selection_height": 400,
            "icon_url": "https://www.edu-apps.org/assets/lti_redirect_engine/redirect_icon.png"
        },
        "custom_fields": {
            "new_tab": 1,
            "url": team_link
        }
    }

    response = requests.post(url, headers=headers, json=payload)
    if response.status_code in (200, 201):
        print(f"🧩 Redirect Tool successfully added to course {course_id}")
    else:
        print(f"⚠️ Failed to add tool to course {course_id}: {response.status_code} - {response.text}")

def write_to_excel(course_team_data, filename="canvas_teams_links.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Team Links"
    ws.append(["SIS Course ID", "Course Name", "Canvas Course ID", "Team Name", "Team Deep Link", "Matched By"])

    for entry in course_team_data:
        ws.append([
            entry["sis_course_id"],
            entry["course_name"],
            entry["canvas_course_id"],
            entry["team_name"],
            entry["team_link"],
            entry["matched_by"]
        ])

    wb.save(filename)
    print(f"📁 Excel file written: {filename}")

def main():
    print("🚀 Fetching Canvas courses...")
    all_courses = get_all_canvas_courses()

    print("🔍 Filtering for Summer 2025 courses...")
    summer_courses = filter_courses(all_courses)

    print("🔐 Authenticating with Microsoft Graph...")
    access_token = authenticate_graph()

    course_team_data = []

    for course in summer_courses:
        sis_id = course.get("sis_course_id", "")
        canvas_id = course.get("id")
        course_name = course.get("name", "")

        team_info = find_team_fallback(course, access_token)
        if team_info:
            team_id = team_info["team_id"]
            team_name = team_info["team_name"]
            match_method = team_info["matched_by"]

            team_link = get_team_web_url(team_id, access_token)
            if team_link:
                print(f"✅ Found team: {team_name} (matched by {match_method})")
                add_teams_redirect_tool(canvas_id, team_link)
            else:
                print(f"⚠️ Skipping course {canvas_id} due to missing team link.")

        else:
            team_link = "Not Found"
            team_name = "Not Found"
            match_method = "none"

        course_team_data.append({
            "sis_course_id": sis_id,
            "course_name": course_name,
            "canvas_course_id": canvas_id,
            "team_name": team_name,
            "team_link": team_link,
            "matched_by": match_method
        })

    write_to_excel(course_team_data)
    print("🏁 Done!")

if __name__ == "__main__":
    main()
