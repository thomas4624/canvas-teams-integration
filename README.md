
# 🎓 Insert Teams into Canvas LMS Nav

This Python script automates the integration of **Canvas LMS** courses with corresponding **Microsoft Teams**. It fetches Canvas courses for a specific term, attempts to match them with Microsoft Teams based on SIS Course ID or course name, and creates a redirect link in each Canvas course pointing to the matched Microsoft Team. Results are exported to an Excel file.

---

## ✨ Features

- Fetch all Canvas courses from a specific enrollment term.
- Filter courses based on a custom string match (e.g., `2025/Summer`).
- Authenticate with Microsoft Graph via device flow.
- Match Canvas courses to Microsoft Teams by `displayName`.
- Automatically create an external tool link in Canvas to the matching Team.
- Export course-team match results to an Excel file.

---

## 🔧 Prerequisites

### Python Packages

Install the required packages:

```bash
pip install requests openpyxl msal
```

### Configuration

Update the following configuration constants at the top of the script:

```python
CANVAS_API_BASE_URL = "https://[your instance].instructure.com/api/v1"
CANVAS_API_TOKEN = "your_canvas_token"
CANVAS_TERM_ID = 328  # Example term ID
CANVAS_MATCH_STRING = "2025/Summer" # Example match string

GRAPH_CLIENT_ID = "your_graph_client_id"
GRAPH_TENANT_ID = "your_tenant_id"
GRAPH_SCOPES = ["Group.Read.All", "User.Read"]
```

---

## 🚀 Usage

Run the script with:

```bash
python your_script_name.py
```

### What it does:
1. Authenticates via Azure AD (you’ll need to visit a URL and enter a code).
2. Fetches all courses from Canvas for the specified term.
3. Filters for courses matching your `CANVAS_MATCH_STRING`.
4. Attempts to find Microsoft Teams with matching names.
5. Adds an LTI redirect tool to the Canvas course navigation if a match is found.
6. Outputs the final mapping to an Excel file (`canvas_teams_links.xlsx`).

---

## 📁 Output Example

The script creates an Excel file with the following columns:

| SIS Course ID | Course Name | Canvas Course ID | Team Name | Team Deep Link | Matched By |
|---------------|-------------|------------------|-----------|----------------|-------------|

---

## ⚠️ Notes

- Canvas API and Microsoft Graph API tokens/permissions must be properly configured.
- Ensure your Azure AD app has the necessary permissions (`Group.Read.All`, `User.Read`).
- Matching logic is basic (`startswith(displayName, ...)`), so results may vary—review manually if needed.
- Tool added to Canvas is a generic redirect (not a full LTI integration).

---


## 📬 Contact

For questions, suggestions, or issues, feel free to reach out or open an issue in your repository.
