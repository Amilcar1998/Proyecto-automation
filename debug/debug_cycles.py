import requests

zephyr_token = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJjb250ZXh0Ijp7ImJhc2VVcmwiOiJodHRwczovL2dydXBvdW5pY29tZXIuYXRsYXNzaWFuLm5ldCIsInVzZXIiOnsiYWNjb3VudElkIjoiNzEyMDIwOjMyN2NjNWU1LWNkYjQtNGVhYS1iY2JjLTcxMGNkZTI3ZDAyYSIsInRva2VuSWQiOiI4MjEzZTdmNS0yZjJlLTRhYTQtODI5Ni00MjkwNjY0MGRhZWEifX0sImlzcyI6ImNvbS5rYW5vYWgudGVzdC1tYW5hZ2VyIiwic3ViIjoiOWNiODYxZGYtOTE4ZC0zZjFhLTgxODYtNjNkMjkxOGMxNTkwIiwiZXhwIjoxNzk1MjkzMjc1LCJpYXQiOjE3NjM3NTcyNzV9.xmr44h6gz_gdVioh9CFE_gwm69QlY9BXQaaV7i0QJmA"
headers = {
    "Authorization": f"Bearer {zephyr_token}",
    "Accept": "application/json"
}

# Fetch first 3 test cycles to see their structure
url_cycles = "https://api.zephyrscale.smartbear.com/v2/testcycles?projectKey=MAS&maxResults=3"
resp = requests.get(url_cycles, headers=headers)
print("Test Cycles:")
print(resp.json())

# Fetch test cases linked to a Jira issue correctly (maybe /v2/testcases requires JQL?)
# Let's try searching test executions using a query
url_execs = 'https://api.zephyrscale.smartbear.com/v2/testexecutions'
params = {
    "projectKey": "MAS",
    "maxResults": 10
}
print("\nTest Executions (general):")
print(requests.get(url_execs, headers=headers, params=params).json())
