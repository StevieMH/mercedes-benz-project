"""
scrape_used.py
──────────────
Scrapes all USED car listings from Mercedes-Benz UK.
Auto-refreshes the token — no manual intervention needed.

FIRST TIME SETUP:
  1. Paste a valid TOKEN below (get from DevTools)
  2. Run: python scrape_used.py
  3. The script handles all future token refreshes automatically
"""

import requests
import json
import time
import os
import glob

# ─── CREDENTIALS ──────────────────────────────────────────────────────────────
# ⚠️ Only paste TOKEN once to get started. Everything else is already set.
# After the first run, credentials.json is created and TOKEN is no longer needed.
TOKEN = "Bearer eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9.eyJwcm9maWxlSUQiOiJkYWUyZDZkMWUwMGY0OWE2YTA5MTIyNGUzZDA4ZGVhYyIsInN1YiI6InNaTzF3ejNaWm5KbmZhMllCeS9XN21QQjJLelF1SUp2WGVab1plV3Ywc0U9IiwiaWF0IjoxNzc0NTYxMDQ5Njk2LCJleHAiOjE3NzQ1NjE5NDk2OTZ9.Q3BBJBkijVfCyUVeoioiSNOCoIaLjH0JKt6Wo5iMJKJGqaDjzcsvfPCr9gObzazgYQlAgrb_S7frzRuN4XHHwd5De2IssB45SLAHelAS888mBk38MQebKUgzjzgXz88ZfD3TLKcXsvGOt6eXZo2EFidbuZKRB6RWoNi_zBMpGQEJyGvKQL9S5kXpoUC8BOruA4lTkuwFIgXSoA5kY2gd7P-RiH7H_0hI0rsm8GmqPnwRsWPkOc2AiDpFqGRO78LGPotH9ubDAXOQQfdrEfXNUiek5gBzuwtqoYWD5zxSi2hzUl0uN94Cp0kKLD43_60QCcLf6lA3MEdm9ZTFeeg_Rg"
PROFILE_ID = "dae2d6d1e00f49a6a091224e3d08deac"
REFRESH_TOKEN = "tZGC1LMJKs2o5jmpe14+A28AAUUI0LuxW8UXRpFBrik="
COOKIE = "optimizely_user=30eb928c-7a2d-4058-afb5-fd354a93d844; _gcl_au=1.1.991989149.1774551131; _scid=mHL7wRI6vYG7syztQdA5kp8Ht0XUkY06DMhG8g; _pin_unauth=dWlkPVpUZG1OemhtTXpVdE16bGpNQzAwTjJNeExUa3pNalF0WlRZNU5EaGhNV0ZpWldGaw; _tt_enable_cookie=1; _ttp=01KMNQTS41CY3G0747M15V6DA1_.tt.2; _ScCbts=%5B%2268%3Bchrome.2%3A2%3A5%22%5D; uslk_umm_2127_s=ewAiAHYAZQByAHMAaQBvAG4AIgA6ACIAMQAiACwAIgBkAGEAdABhACIAOgB7AH0AfQA=; _gid=GA1.3.219232878.1774551134; _fbp=fb.2.1774551156755.783415520639627941; AWSALB=8xMEFOD87w042NcQGP1GQ7V8BOUv4I3wu5mAu2mq1xrbuLIQOi3iyrWXkNf7nTmSA8n/kn39iVmvDGDrdPk4WV1nkiAz1akGPb7AXa7DUjXzOK1OnGanYloFE/cg; AWSALBCORS=8xMEFOD87w042NcQGP1GQ7V8BOUv4I3wu5mAu2mq1xrbuLIQOi3iyrWXkNf7nTmSA8n/kn39iVmvDGDrdPk4WV1nkiAz1akGPb7AXa7DUjXzOK1OnGanYloFE/cg; profile=%7B%22id%22%3A%22dae2d6d1e00f49a6a091224e3d08deac%22%2C%22refreshToken%22%3A%22tZGC1LMJKs2o5jmpe14%2BA28AAUUI0LuxW8UXRpFBrik%3D%22%2C%22authToken%22%3A%22eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9%22%7D; _ga=GA1.3.1776843931.1774551134; _uetsid=a50ace90293411f19aaaef9c9ec7e001; _uetvid=baa45e20173611f1bb8313bfbf987a3c; _ga_N1E9P5N35X=GS2.1.s1774551133$o1$g1$t1774554196$j44$l0$h1523598671; settings=%7B%22isAgentUser%22%3Afalse%2C%22viewportDesktop%22%3Afalse%2C%22viewportPortrait%22%3Atrue%2C%22mode%22%3A%7B%22finance%22%3Afalse%7D%2C%22footer%22%3Anull%7D"
# ──────────────────────────────────────────────────────────────────────────────

REFRESH_URL = "https://shop.mercedes-benz.co.uk/api/v3/token/refresh"
SEARCH_URL = "https://shop.mercedes-benz.co.uk/api/v4/vehicles/search/used"
PROGRESS_FILE = "data/used/progress.json"
CHUNKS_DIR = "data/used/chunks"
FINAL_FILE = "data/used/mercedes_used_cars_FULL.json"
CREDS_FILE = "data/used/credentials.json"
TOTAL_PAGES = 403


def load_credentials():
    if os.path.exists(CREDS_FILE):
        with open(CREDS_FILE) as f:
            return json.load(f)
    return {"token": TOKEN, "profile_id": PROFILE_ID, "refresh_token": REFRESH_TOKEN}


def save_credentials(token, profile_id, refresh_token):
    os.makedirs(os.path.dirname(CREDS_FILE), exist_ok=True)
    with open(CREDS_FILE, "w") as f:
        json.dump({"token": token, "profile_id": profile_id,
                  "refresh_token": refresh_token}, f)


def do_refresh(current_token, profile_id, refresh_tok):
    print("🔄 Refreshing token automatically...")
    headers = {
        "accept": "application/json",
        "authorization": current_token,
        "content-type": "application/json",
        "origin": "https://shop.mercedes-benz.co.uk",
        "cookie": COOKIE
    }
    payload = {"profileID": profile_id, "refreshToken": refresh_tok}
    try:
        r = requests.post(REFRESH_URL, headers=headers,
                          json=payload, timeout=15)
        if r.status_code == 200:
            data = r.json()
            new_token = "Bearer " + data.get("authToken", "")
            new_refresh = data.get("refreshToken", refresh_tok)
            print("✅ Token refreshed successfully")
            return new_token, new_refresh
        else:
            print(f"❌ Refresh failed ({r.status_code}): {r.text[:200]}")
            return None, None
    except Exception as e:
        print(f"❌ Refresh error: {e}")
        return None, None


def make_headers(token):
    return {
        "accept": "application/json",
        "accept-language": "es-ES,es;q=0.9,en;q=0.8",
        "authorization": token,
        "content-type": "application/json",
        "origin": "https://shop.mercedes-benz.co.uk",
        "sec-fetch-dest": "empty",
        "sec-fetch-mode": "cors",
        "sec-fetch-site": "same-origin",
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/146.0.0.0 Safari/537.36",
        "cookie": COOKIE
    }


def make_payload(page):
    return {
        "Criteria": {
            "VehicleType": 2,
            "LimitToMotability": False,
            "LatestModel": False,
            "PreviousModel": False,
            "PromotionalOfferVehiclesOnly": False,
            "RetailerGroupId": []
        },
        "Sort": {"Id": 1},
        "Finance": {"Criteria": {
            "Key": "PCP", "Name": "Agility (Personal Contract Plan)",
            "Type": "PCP", "IsDefault": True,
            "Term": {"Options": [{"IsDefault": True, "Value": 48}]},
            "Deposit": {"Default": "17.5%"},
            "Mileage": {"Options": [{"IsDefault": True, "Value": 10000}]},
            "MonthlyPrice": {"Min": 50, "Max": 4000},
            "IsPersonalised": False, "CustomerType": "Personal",
            "VehicleType": "UNASSIGNED", "AdvanceRentals": None, "RegularPayment": None
        }},
        "BestMatchCriteriaId": "6A9BCA02-BD6C-4854-ABAB-3560FC2CDB70",
        "DisableBestMatch": False,
        "IncludeOffers": True, "IncludeReservations": False, "IncludeQuotes": False,
        "Paging": {"ResultsPerPage": 20, "PageIndex": page}
    }


def get_start_page():
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE) as f:
            data = json.load(f)
        start = data["last_page"] + 1
        print(f"🔄 Resuming from page {start}/{TOTAL_PAGES}")
        return start
    print("🆕 Starting fresh — Used Cars")
    return 0


def save_progress(page):
    os.makedirs(os.path.dirname(PROGRESS_FILE), exist_ok=True)
    with open(PROGRESS_FILE, "w") as f:
        json.dump({"last_page": page}, f)


def save_chunk(page, vehicles):
    os.makedirs(CHUNKS_DIR, exist_ok=True)
    with open(os.path.join(CHUNKS_DIR, f"page_{page:04d}.json"), "w") as f:
        json.dump(vehicles, f)


def merge_chunks():
    files = sorted(glob.glob(os.path.join(CHUNKS_DIR, "page_*.json")))
    seen, all_vehicles = set(), []
    for fp in files:
        with open(fp) as f:
            for v in json.load(f):
                if v.get("Id") not in seen:
                    seen.add(v["Id"])
                    all_vehicles.append(v)
    return all_vehicles


def scrape():
    os.makedirs(CHUNKS_DIR, exist_ok=True)
    start_page = get_start_page()

    creds = load_credentials()
    current_token = creds["token"]
    profile_id = creds["profile_id"]
    current_refresh = creds["refresh_token"]

    for page in range(start_page, TOTAL_PAGES):
        try:
            r = requests.post(SEARCH_URL, headers=make_headers(
                current_token), json=make_payload(page))

            if r.status_code == 403:
                print(
                    f"\n⚠️  Token expired at page {page} — attempting auto-refresh...")
                new_token, new_refresh = do_refresh(
                    current_token, profile_id, current_refresh)

                if new_token:
                    current_token = new_token
                    current_refresh = new_refresh
                    save_credentials(
                        current_token, profile_id, current_refresh)
                    r = requests.post(SEARCH_URL, headers=make_headers(
                        current_token), json=make_payload(page))
                    if r.status_code != 200:
                        print(
                            f"❌ Still failing after refresh ({r.status_code}). Stopping.")
                        save_progress(page - 1)
                        break
                else:
                    print(
                        "❌ Auto-refresh failed. Paste a fresh TOKEN into the script and re-run.")
                    save_progress(page - 1)
                    break

            if r.status_code != 200:
                print(f"❌ Error {r.status_code} on page {page}, skipping...")
                time.sleep(2)
                continue

            vehicles = r.json()["SearchResults"]["Vehicles"]
            save_chunk(page, vehicles)
            save_progress(page)
            print(
                f"✅ Page {page + 1}/{TOTAL_PAGES} — {len(vehicles)} vehicles")
            time.sleep(0.5)

        except Exception as e:
            print(f"💥 Exception on page {page}: {e}")
            save_progress(page - 1)
            time.sleep(3)

    print("\n📦 Merging chunks...")
    all_vehicles = merge_chunks()
    with open(FINAL_FILE, "w") as f:
        json.dump(all_vehicles, f, indent=2)
    print(f"🏁 Done! {len(all_vehicles):,} vehicles → {FINAL_FILE}")

    if os.path.exists(PROGRESS_FILE):
        os.remove(PROGRESS_FILE)


scrape()
