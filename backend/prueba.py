import socket
import requests

print(socket.gethostbyname("five.epicollect.net"))

API_URL = "https://five.epicollect.net/api/export/entries/csa-ups-instalacion"

headers = {
    "User-Agent": "Mozilla/5.0"
}

response = requests.get(
    API_URL,
    headers=headers,
    verify=False,
    timeout=10
)

print(response.status_code)
print(response.text[:500])