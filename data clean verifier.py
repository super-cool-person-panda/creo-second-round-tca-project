import pandas as pd

sheets = [
    'All Data',
    'With Profit Band',
    'Complete Labelled',
    'Incomplete Labelled',
    'Without Profit Band',
    'Complete Without Profit Band'
]

cols_to_check = [
    'property_type',
    'neighbourhood',
    'requested_timeline',
    'referral_source',
    'homeowner_status',
    'preferred_contact',
    'lead_capture_weather',
    'customer_age_bracket',
    'has_pets',
    'lead_weekday',
    'expected_profit_band',
]

for sheet in sheets:
    df = pd.read_excel("leads_data_organised.xlsx", sheet_name=sheet)

    print("\n" + "=" * 60)
    print(f"SHEET: {sheet} ({len(df)} rows)")
    print("=" * 60)

    print("\n── UNIQUE VALUES PER COLUMN ──")
    for col in cols_to_check:
        if col in df.columns:
            print(f"\n{col}:")
            print(df[col].value_counts(dropna=False).to_string())

    print("\n── DUPLICATE ROWS ──")
    dupes = df[df.duplicated()]
    print(f"Number of duplicate rows: {len(dupes)}")
    if len(dupes) > 0:
        print(dupes)

    print("\n── DUPLICATE LEAD IDs ──")
    dupe_ids = df[df['lead_id'].duplicated()]
    print(f"Number of duplicate lead IDs: {len(dupe_ids)}")
    if len(dupe_ids) > 0:
        print(dupe_ids[['lead_id', 'lead_date', 'property_type', 'neighbourhood']])
