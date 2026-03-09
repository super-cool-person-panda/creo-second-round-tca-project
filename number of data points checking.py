import requests
import pandas as pd

base_url = "https://dqcqexieqlfqylqwogsj.supabase.co/rest/v1/inbound_leads?select=*"
headers = {
    "apikey": "sb_publishable_iFgXlEVP5UqmrZx6l1nVgw_6WD5CPpy",
}

all_data = []
limit = 250  # ← changed from 500 to 250
offset = 0

while True:
    url = f"{base_url}&limit={limit}&offset={offset}"
    response = requests.get(url, headers=headers)
    batch = response.json()

    print(f"Offset {offset}: got {len(batch)} rows")

    if not isinstance(batch, list) or len(batch) == 0:
        break

    all_data.extend(batch)
    offset += limit

df = pd.DataFrame(all_data)
print(f"\nTotal rows fetched: {len(df)}")

