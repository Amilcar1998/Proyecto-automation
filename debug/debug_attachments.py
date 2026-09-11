import sys
import os
import requests

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from jira.jira_utilidades import JiraClient

client = JiraClient()
url = f"{client.base_url}/rest/api/3/search/jql"
jql = 'project = "MAS" AND issuetype IN (Story, EPIC) ORDER BY created DESC'
payload = {"jql": jql, "maxResults": 5, "fields": ["summary", "attachment"]}
response = requests.post(url, json=payload, headers=client.headers_json, auth=client.auth, timeout=30)

if response.status_code == 200:
    issues = response.json().get("issues", [])
    for issue in issues:
        print(f"\nIssue: {issue['key']}")
        attachments = issue['fields'].get('attachment', [])
        print(f"Attachments: {len(attachments)}")
        for att in attachments:
            print(f" - {att.get('filename')}")
else:
    print(response.status_code)
