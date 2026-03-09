import pandas as pd
import numpy as np

# ── LOAD DATA ─────────────────────────────────────────────────────────

df = pd.read_excel("leads_data_organised.xlsx", sheet_name='Complete Labelled')

profit_bands = ['High', 'Medium', 'Low']

# ── REMOVE OUTLIERS AND BIN PROPERTY SIZE ────────────────────────────

job_sizes = df['estimated_job_size_sqft'].dropna()

print("Job Size Stats (before removing outliers):")
print(f"  Min:    {job_sizes.min():.0f} sqft")
print(f"  Max:    {job_sizes.max():.0f} sqft")
print(f"  Mean:   {job_sizes.mean():.0f} sqft")
print(f"  Median: {job_sizes.median():.0f} sqft")
print(f"  Std:    {job_sizes.std():.0f} sqft")

# Remove outliers using IQR method
q1  = job_sizes.quantile(0.25)
q3  = job_sizes.quantile(0.75)
iqr = q3 - q1
upper_bound = q3 + 1.5 * iqr
lower_bound = max(0, q1 - 1.5 * iqr)

print(f"\nOutlier Bounds:")
print(f"  Lower: {lower_bound:.0f} sqft")
print(f"  Upper: {upper_bound:.0f} sqft")

outliers = df[df['estimated_job_size_sqft'] > upper_bound]
print(f"  Outliers removed: {len(outliers)} rows")

# Cap outliers instead of removing them
df['estimated_job_size_sqft_clean'] = df['estimated_job_size_sqft'].clip(upper=upper_bound)

job_sizes_clean = df['estimated_job_size_sqft_clean'].dropna()

print(f"\nJob Size Stats (after removing outliers):")
print(f"  Min:    {job_sizes_clean.min():.0f} sqft")
print(f"  Max:    {job_sizes_clean.max():.0f} sqft")
print(f"  Mean:   {job_sizes_clean.mean():.0f} sqft")
print(f"  Median: {job_sizes_clean.median():.0f} sqft")

# Use Sturges Rule only (most stable for this data)
n = len(job_sizes_clean)
n_bins = int(np.ceil(np.log2(n) + 1))
print(f"\nOptimal number of bins (Sturges): {n_bins}")

# Generate equal-width bins based on clean data
bin_min = 0
bin_max = int(np.ceil(job_sizes_clean.max() / 100) * 100)
bin_width = int(np.ceil((bin_max - bin_min) / n_bins / 100) * 100)

bin_edges = list(range(bin_min, bin_max + bin_width, bin_width))
bin_edges[-1] = bin_max + 1
labels = [f"{bin_edges[i]}-{bin_edges[i+1]-1}" for i in range(len(bin_edges)-1)]

print(f"Bin width: {bin_width} sqft")
print(f"Bins: {labels}")

# Apply bins
df['job_size_bin'] = pd.cut(
    df['estimated_job_size_sqft_clean'],
    bins=bin_edges,
    labels=labels,
    right=False,
    include_lowest=True
)

print("\nJob Size Bin Distribution:")
print(df['job_size_bin'].value_counts().sort_index())

# ── FUNCTION TO BUILD FREQUENCY + PROBABILITY TABLE ──────────────────

def build_table(col):
    df_temp = df[df[col].notna()].copy()

    counts = df_temp.groupby([col, 'expected_profit_band'], observed=True).size().unstack(fill_value=0)

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

    counts = df_temp.groupby(cols + ['expected_profit_band'], observed=True).size().unstack(fill_value=0)

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

with pd.ExcelWriter("quantitative_analysis.xlsx", engine='openpyxl') as writer:

    build_table('neighbourhood').to_excel(writer,       sheet_name='Neighbourhood')
    print("\n✓ Done: Neighbourhood")

    build_table('customer_age_bracket').to_excel(writer, sheet_name='Age Bracket')
    print("✓ Done: Age Bracket")

    build_table('property_type').to_excel(writer,       sheet_name='Property Type')
    print("✓ Done: Property Type")

    build_table('homeowner_status').to_excel(writer,    sheet_name='Homeowner Status')
    print("✓ Done: Homeowner Status")

    build_table('job_size_bin').to_excel(writer,        sheet_name='Job Size')
    print("✓ Done: Job Size")

    build_table('requested_timeline').to_excel(writer,  sheet_name='Timeline')
    print("✓ Done: Timeline")

    build_conditional([
        'neighbourhood',
        'customer_age_bracket',
        'property_type',
        'homeowner_status',
        'job_size_bin',
        'requested_timeline',
    ]).to_excel(writer, sheet_name='All 6 Factors')
    print("✓ Done: All 6 Factors")

print("\nSaved to quantitative_analysis.xlsx!")
