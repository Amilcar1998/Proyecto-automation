import requests

zephyr_token = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJjb250ZXh0Ijp7ImJhc2VVcmwiOiJodHRwczovL2dydXBvdW5pY29tZXIuYXRsYXNzaWFuLm5ldCIsInVzZXIiOnsiYWNjb3VudElkIjoiNzEyMDIwOjMyN2NjNWU1LWNkYjQtNGVhYS1iY2JjLTcxMGNkZTI3ZDAyYSIsInRva2VuSWQiOiI4MjEzZTdmNS0yZjJlLTRhYTQtODI5Ni00MjkwNjY0MGRhZWEifX0sImlzcyI6ImNvbS5rYW5vYWgudGVzdC1tYW5hZ2VyIiwic3ViIjoiOWNiODYxZGYtOTE4ZC0zZjFhLTgxODYtNjNkMjkxOGMxNTkwIiwiZXhwIjoxNzk1MjkzMjc1LCJpYXQiOjE3NjM3NTcyNzV9.xmr44h6gz_gdVioh9CFE_gwm69QlY9BXQaaV7i0QJmA"
headers = {
    "Authorization": f"Bearer {zephyr_token}",
    "Accept": "application/json"
}

url_execs = "https://api.zephyrscale.smartbear.com/v2/testexecutions?projectKey=MAS&maxResults=1000"
resp = requests.get(url_execs, headers=headers)
data = resp.json()
print(f"Total executions returned: {len(data.get('values', []))}")
if data.get('values'):
    c = data['values'][0]
    print(c.keys())
    print("Execution example testCycle:", c.get("testCycle"))
    print("Execution example executionStatus:", c.get("executionStatus"))
