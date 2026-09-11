import sys
import os
import requests

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from jira.jira_utilidades import JiraClient

client = JiraClient()
url = f"{client.base_url}/rest/api/3/search/jql"
jql = 'project = "MAS" AND issuetype IN (Story, EPIC) ORDER BY created DESC'
payload = {"jql": jql, "maxResults": 2, "fields": ["summary", "status", "issuetype", "subtasks"]}
response = requests.post(url, json=payload, headers=client.headers_json, auth=client.auth, timeout=30)

if response.status_code == 200:
    issues = response.json().get("issues", [])
    for issue in issues:
        print(f"\nIssue: {issue['key']}")
        print(f"Subtasks field: {issue['fields'].get('subtasks')}")
        
        # We can also check links just in case it's a linked issue and not a subtask
        print("Issue Links:")
        # Wait, issue links are in 'issuelinks' field.
else:
    print(response.status_code)

# Let's also fetch with issuelinks
payload = {"jql": jql, "maxResults": 2, "fields": ["summary", "status", "issuetype", "subtasks", "issuelinks"]}
response = requests.post(url, json=payload, headers=client.headers_json, auth=client.auth, timeout=30)
if response.status_code == 200:
    issues = response.json().get("issues", [])
    for issue in issues:
        print(f"\nIssue: {issue['key']} Links:")
        print(issue['fields'].get('issuelinks'))

