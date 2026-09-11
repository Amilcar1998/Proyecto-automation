import requests

zephyr_token = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJjb250ZXh0Ijp7ImJhc2VVcmwiOiJodHRwczovL2dydXBvdW5pY29tZXIuYXRsYXNzaWFuLm5ldCIsInVzZXIiOnsiYWNjb3VudElkIjoiNzEyMDIwOjMyN2NjNWU1LWNkYjQtNGVhYS1iY2JjLTcxMGNkZTI3ZDAyYSIsInRva2VuSWQiOiI4MjEzZTdmNS0yZjJlLTRhYTQtODI5Ni00MjkwNjY0MGRhZWEifX0sImlzcyI6ImNvbS5rYW5vYWgudGVzdC1tYW5hZ2VyIiwic3ViIjoiOWNiODYxZGYtOTE4ZC0zZjFhLTgxODYtNjNkMjkxOGMxNTkwIiwiZXhwIjoxNzk1MjkzMjc1LCJpYXQiOjE3NjM3NTcyNzV9.xmr44h6gz_gdVioh9CFE_gwm69QlY9BXQaaV7i0QJmA"
headers = {
    "Authorization": f"Bearer {zephyr_token}",
    "Accept": "application/json"
}

url_cycles = "https://api.zephyrscale.smartbear.com/v2/testcycles?projectKey=MAS&maxResults=100"
resp = requests.get(url_cycles, headers=headers)
data = resp.json()
print(f"Total cycles returned: {len(data.get('values', []))}")
# Look at the first cycle to see if links are included
if data.get('values'):
    c = data['values'][0]
    print("First cycle links:")
    print(c.get("links", {}).get("issues", []))
