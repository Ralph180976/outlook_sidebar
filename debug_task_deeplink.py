import sys
import requests

url = "https://outlook.live.com/mail/0/deeplink/tasks"

# Use a valid task id
task_id = "AQMkADAwATM3ZmYBLWFjNTEtOTdjYS0wMAItMDAKAEYAAANITyg22N35R4N4xB_2p1IHBwAA2RxgJY1sEEGfWKWXQD-YKAAAAgESAAAA2RxgJY1sEEGfWKWXQD-YKAAAAqTiAAAA"

test_url = f"{url}?itemId={task_id}"
print("Testing:", test_url)

try:
    resp = requests.get(test_url, allow_redirects=False)
    print("Status:", resp.status_code)
    print("Location:", resp.headers.get("Location"))
except Exception as e:
    print(e)
