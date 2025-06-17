import requests


def check_access():
    url = "https://raw.githubusercontent.com/ikigu/excel_auto/main/access.json"
    try:
        resp = requests.get(url, timeout=3)
        if resp.status_code == 200 and resp.json().get("access_granted"):
            return True
        else:
            print("Access denied.")
            return False
    except:
        print("Could not verify access.")
        return False
