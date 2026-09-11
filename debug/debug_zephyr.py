import requests

zephyr_token = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJjb250ZXh0Ijp7ImJhc2VVcmwiOiJodHRwczovL2dydXBvdW5pY29tZXIuYXRsYXNzaWFuLm5ldCIsInVzZXIiOnsiYWNjb3VudElkIjoiNzEyMDIwOjMyN2NjNWU1LWNkYjQtNGVhYS1iY2JjLTcxMGNkZTI3ZDAyYSIsInRva2VuSWQiOiI4MjEzZTdmNS0yZjJlLTRhYTQtODI5Ni00MjkwNjY0MGRhZWEifX0sImlzcyI6ImNvbS5rYW5vYWgudGVzdC1tYW5hZ2VyIiwic3ViIjoiOWNiODYxZGYtOTE4ZC0zZjFhLTgxODYtNjNkMjkxOGMxNTkwIiwiZXhwIjoxNzk1MjkzMjc1LCJpYXQiOjE3NjM3NTcyNzV9.xmr44h6gz_gdVioh9CFE_gwm69QlY9BXQaaV7i0QJmA"
headers = {
    "Authorization": f"Bearer {zephyr_token}",
    "Accept": "application/json"
}
issue_key = "MAS-42"
# Jira Issue MAS-42 ID is 757658 (from previous logs)
issue_id = "757658"

# 1. Try links API
url_links = f"https://api.zephyrscale.smartbear.com/v2/links/issues/{issue_key}/testcases"
print("Links using issueKey:")
resp_links = requests.get(url_links, headers=headers)
print(resp_links.json())

url_links2 = f"https://api.zephyrscale.smartbear.com/v2/testcases?query=jiraIssueKeys=\"{issue_key}\""
print("Testcases using query jiraIssueKeys:")
resp_tc = requests.get(url_links2, headers=headers)
print(resp_tc.json())

url_tc = f"https://api.zephyrscale.smartbear.com/v2/testcases?projectKey=MAS&maxResults=2"
print("General test cases:")
print(requests.get(url_tc, headers=headers).json())
