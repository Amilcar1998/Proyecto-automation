import os
import requests
from requests.auth import HTTPBasicAuth
import json

URL = "https://grupounicomer.atlassian.net"
USER = "eliseo_lopezp@unicomer.com"
TOKEN = "ATATT3xFfGF0AVjOUbXoEUUHF5WGZqYlOaPDMCFUXeoExn2772LdgxwCFFWzHrB3e8ONkqOmZ26-Kd6n9UniwxpVRecOwZzVSSvLGOEx_hba73_NT25iBmPfkou3uXoSncd8MAercwOlPEOoqrKqWUsgJZueewd-K9fLIspWbVZFNvil16jstiA=05557851"

auth = HTTPBasicAuth(USER, TOKEN)
headers = {"Accept": "application/json"}

# Fetch createmeta for ML project
response = requests.get(
    f"{URL}/rest/api/3/issue/createmeta?projectKeys=ML&expand=projects.issuetypes.fields",
    headers=headers,
    auth=auth
)

with open(r"C:\Users\eliseo_lopezp\proyecto1\test\jira_meta.json", "w", encoding="utf-8") as f:
    json.dump(response.json(), f, indent=4)
    
print("Metadata guardada en jira_meta.json")
