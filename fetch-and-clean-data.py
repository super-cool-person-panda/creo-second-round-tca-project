import requests
import pandas as pd

# ── STEP 1: FETCH ALL DATA FROM API ──────────────────────────────────

base_url = "https://dqcqexieqlfqylqwogsj.supabase.co/rest/v1/inbound_leads?select=*"
headers = {
    "apikey": "sb_publishable_iFgXlEVP5UqmrZx6l1nVgw_6WD5CPpy",
}

all_data = []
limit = 250
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

# ── STEP 2: CLEAN ALL STRING COLUMNS FIRST ───────────────────────────

df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)

# ── STEP 3: CLEAN DATE COLUMN ─────────────────────────────────────────

df['lead_date'] = pd.to_datetime(df['lead_date'], format='mixed', errors='coerce').dt.date

# ── STEP 4: CLEAN ALL OTHER COLUMNS ──────────────────────────────────

neighbourhood_mapping = {
    'West End': 'West End', 'Westend': 'West End', 'Westend ': 'West End',
    ' West End': 'West End', 'west end': 'West End',
    'Downtown': 'Downtown', 'Down Town': 'Downtown', 'Down town': 'Downtown',
    'downtown': 'Downtown', ' Downtown': 'Downtown',
    'Sydenham Ward': 'Sydenham Ward', 'Sydenhamm Ward': 'Sydenham Ward',
    'sydenham ward': 'Sydenham Ward',
    'Strathcona Park': 'Strathcona Park', 'strathcona park': 'Strathcona Park',
    'Portsmouth Village': 'Portsmouth Village', 'Portsmoth Village': 'Portsmouth Village',
    'portsmouth village': 'Portsmouth Village',
    'Calvin Park': 'Calvin Park', ' Calvin Park ': 'Calvin Park', 'calvin park': 'Calvin Park',
}

property_mapping = {
    'Detached': 'Detached', 'detached': 'Detached', 'DETACHED': 'Detached',
    'Semi-Detached': 'Semi-Detached', 'Semi Detached': 'Semi-Detached', 'semi-detached': 'Semi-Detached',
    'Townhouse': 'Townhouse', 'townhouse': 'Townhouse', 'Town House': 'Townhouse',
    'Apartment': 'Apartment', 'apartment': 'Apartment',
    'Heritage Home': 'Heritage Home', 'heritage home': 'Heritage Home',
}

referral_mapping = {
    'Facebook Ads': 'Facebook Ads', 'FaceBook': 'Facebook Ads', 'Facebook': 'Facebook Ads',
    'facebook ads': 'Facebook Ads', 'FB Ads': 'Facebook Ads',
    'Lawn Signs': 'Lawn Signs', 'LawnSign': 'Lawn Signs', 'Lawn Sign': 'Lawn Signs',
    'lawn signs': 'Lawn Signs',
    'Door-to-Door': 'Door-to-Door', 'Door 2 Door': 'Door-to-Door', 'door-to-door': 'Door-to-Door',
    'Door to Door': 'Door-to-Door',
    'Word of Mouth/Referral': 'Word of Mouth/Referral', 'Word-of-mouth': 'Word of Mouth/Referral',
    'Word of Mouth': 'Word of Mouth/Referral', 'word of mouth': 'Word of Mouth/Referral',
    'Referral': 'Word of Mouth/Referral', 'WOM': 'Word of Mouth/Referral',
}

timeline_mapping = {
    'ASAP': 'ASAP', 'asap': 'ASAP', 'A.S.A.P': 'ASAP',
    '1-2 weeks': '1-2 weeks', '1-2 Weeks': '1-2 weeks', '1 - 2 weeks': '1-2 weeks',
    '1 month': '1 month', '1 Month': '1 month', 'One month': '1 month',
    'Flexible': 'Flexible', 'flexible': 'Flexible', 'Anytime': 'Flexible', 'No Rush': 'Flexible',
}

homeowner_mapping = {
    'Own': 'Own', 'own': 'Own', 'Owner': 'Own',
    'Rent': 'Rent', 'rent': 'Rent', 'Renting': 'Rent',
    'Recently Purchased': 'Recently Purchased', 'recently purchased': 'Recently Purchased',
    'Recent Purchase': 'Recently Purchased', 'New Purchase': 'Recently Purchased',
}

contact_mapping = {
    'Email': 'Email', 'email': 'Email',
    'SMS': 'SMS', 'sms': 'SMS', 'Text': 'SMS',
    'Phone Call': 'Phone Call', 'phone call': 'Phone Call', 'Phone': 'Phone Call', 'Call': 'Phone Call',
}

weather_mapping = {
    'Sunny': 'Sunny', 'sunny': 'Sunny',
    'Cloudy': 'Cloudy', 'cloudy': 'Cloudy',
    'Rain': 'Rain', 'rain': 'Rain', 'Rainy': 'Rain',
    'Snow': 'Snow', 'snow': 'Snow',
    'Windy': 'Windy', 'windy': 'Windy',
}

age_mapping = {
    '18-24': '18-24', '25-34': '25-34', '35-44': '35-44',
    '45-54': '45-54', '55-64': '55-64', '65+': '65+', '65 +': '65+',
}

weekday_mapping = {
    'Monday': 'Monday', 'Tuesday': 'Tuesday', 'Wednesday': 'Wednesday',
    'Thursday': 'Thursday', 'Friday': 'Friday', 'Saturday': 'Saturday', 'Sunday': 'Sunday',
}

profit_mapping = {
    'High': 'High', 'high': 'High',
    'Medium': 'Medium', 'medium': 'Medium', 'Med': 'Medium',
    'Low': 'Low', 'low': 'Low',
}

df['neighbourhood']        = df['neighbourhood'].map(neighbourhood_mapping)
df['property_type']        = df['property_type'].map(property_mapping)
df['referral_source']      = df['referral_source'].map(referral_mapping)
df['requested_timeline']   = df['requested_timeline'].map(timeline_mapping)
df['homeowner_status']     = df['homeowner_status'].map(homeowner_mapping)
df['preferred_contact']    = df['preferred_contact'].map(contact_mapping)
df['lead_capture_weather'] = df['lead_capture_weather'].map(weather_mapping)
df['customer_age_bracket'] = df['customer_age_bracket'].map(age_mapping)
df['lead_weekday']         = df['lead_weekday'].map(weekday_mapping)
df['expected_profit_band'] = df['expected_profit_band'].map(profit_mapping)

# ── STEP 5: SPLIT DATA ────────────────────────────────────────────────

df_labelled   = df[df['expected_profit_band'].notna()]
df_unlabelled = df[df['expected_profit_band'].isna()]

df_labelled_complete   = df_labelled[df_labelled.notna().all(axis=1)]
df_labelled_incomplete = df_labelled[df_labelled.isna().any(axis=1)]
df_unlabelled_complete = df_unlabelled[df_unlabelled.drop(columns=['expected_profit_band']).notna().all(axis=1)]

print(f"\nLabelled rows (have profit band):      {len(df_labelled)}")
print(f"  - Complete (no missing values):      {len(df_labelled_complete)}")
print(f"  - Incomplete (has missing values):   {len(df_labelled_incomplete)}")
print(f"Unlabelled rows (missing profit band): {len(df_unlabelled)}")
print(f"  - Complete (no missing values):      {len(df_unlabelled_complete)}")

# ── STEP 6: SAVE TO EXCEL ─────────────────────────────────────────────

with pd.ExcelWriter("leads_data_organised.xlsx", engine='openpyxl') as writer:
    df.to_excel(writer,                     sheet_name='All Data',                     index=False)
    df_labelled.to_excel(writer,            sheet_name='With Profit Band',             index=False)
    df_labelled_complete.to_excel(writer,   sheet_name='Complete Labelled',            index=False)
    df_labelled_incomplete.to_excel(writer, sheet_name='Incomplete Labelled',          index=False)
    df_unlabelled.to_excel(writer,          sheet_name='Without Profit Band',          index=False)
    df_unlabelled_complete.to_excel(writer, sheet_name='Complete Without Profit Band', index=False)

print("\nSaved to leads_data_organised.xlsx!")
print("  Sheet 1: All Data")
print("  Sheet 2: With Profit Band")
print("  Sheet 3: Complete Labelled")
print("  Sheet 4: Incomplete Labelled")
print("  Sheet 5: Without Profit Band")
print("  Sheet 6: Complete Without Profit Band")
print("\nDone!")
