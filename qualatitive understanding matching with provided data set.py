import pandas as pd

# ── LOAD DATA ─────────────────────────────────────────────────────────

df = pd.read_excel("leads_data_organised.xlsx", sheet_name='Complete Labelled')

profit_bands = ['High', 'Medium', 'Low']

# ── FUNCTION TO BUILD FREQUENCY + PROBABILITY TABLE ──────────────────

def build_table(col):
    df_temp = df[df[col].notna()].copy()

    counts = df_temp.groupby([col, 'expected_profit_band']).size().unstack(fill_value=0)

    for band in profit_bands:
        if band not in counts.columns:
            counts[band] = 0
    counts = counts[profit_bands]

    counts['Total']    = counts[profit_bands].sum(axis=1)
    counts['Total %']  = (counts['Total'] / counts['Total'].sum() * 100).round(1)
    counts['High %']   = (counts['High']   / counts['Total'] * 100).round(1)
    counts['Medium %'] = (counts['Medium'] / counts['Total'] * 100).round(1)
    counts['Low %']    = (counts['Low']    / counts['Total'] * 100).round(1)

    return counts[['High', 'Medium', 'Low', 'Total', 'Total %', 'High %', 'Medium %', 'Low %']]

# ── FUNCTION TO BUILD CONDITIONAL PROBABILITY TABLE ──────────────────

def build_conditional(cols):
    df_temp = df[cols + ['expected_profit_band']].dropna().copy()

    counts = df_temp.groupby(cols + ['expected_profit_band']).size().unstack(fill_value=0)

    for band in profit_bands:
        if band not in counts.columns:
            counts[band] = 0
    counts = counts[profit_bands]

    counts['Total']    = counts.sum(axis=1)
    counts['High %']   = (counts['High']   / counts['Total'] * 100).round(1)
    counts['Medium %'] = (counts['Medium'] / counts['Total'] * 100).round(1)
    counts['Low %']    = (counts['Low']    / counts['Total'] * 100).round(1)

    return counts[['High', 'Medium', 'Low', 'Total', 'High %', 'Medium %', 'Low %']]

# ── SAVE TO EXCEL ─────────────────────────────────────────────────────

with pd.ExcelWriter("frequency_probability_tables_from_qualitative_understanding.xlsx", engine='openpyxl') as writer:

    # Individual factor tables
    build_table('neighbourhood').to_excel(writer,        sheet_name='Neighbourhood')
    print("✓ Done: Neighbourhood")

    build_table('customer_age_bracket').to_excel(writer, sheet_name='Age Bracket')
    print("✓ Done: Age Bracket")

    build_table('homeowner_status').to_excel(writer,     sheet_name='Homeowner Status')
    print("✓ Done: Homeowner Status")

    build_table('property_type').to_excel(writer,        sheet_name='Property Type')
    print("✓ Done: Property Type")

    # Conditional probability tables
    build_conditional([
        'neighbourhood',
        'customer_age_bracket'
    ]).to_excel(writer, sheet_name='Neigh → Age')
    print("✓ Done: Neighbourhood → Age")

    build_conditional([
        'neighbourhood',
        'customer_age_bracket',
        'property_type'
    ]).to_excel(writer, sheet_name='Neigh → Age → Property')
    print("✓ Done: Neighbourhood → Age → Property")

    build_conditional([
        'neighbourhood',
        'customer_age_bracket',
        'property_type',
        'homeowner_status'
    ]).to_excel(writer, sheet_name='All 4 Factors')
    print("✓ Done: All 4 Factors")

print("\nSaved to frequency_probability_tables_from_qualitative_understanding.xlsx!")
