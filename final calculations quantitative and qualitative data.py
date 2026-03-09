import pandas as pd
import numpy as np
from scipy.stats import chi2_contingency

# ── STEP 1: LOAD DATA ─────────────────────────────────────────────────

df = pd.read_excel("leads_data_organised.xlsx", sheet_name='Complete Labelled')

profit_bands = ['High', 'Medium', 'Low']

# ── BIN PROPERTY SIZE ─────────────────────────────────────────────────

job_sizes   = df['estimated_job_size_sqft'].dropna()
q1          = job_sizes.quantile(0.25)
q3          = job_sizes.quantile(0.75)
iqr         = q3 - q1
upper_bound = q3 + 1.5 * iqr

df['estimated_job_size_sqft_clean'] = df['estimated_job_size_sqft'].clip(upper=upper_bound)

bin_edges = [0, 750, 1500, 2250, float('inf')]
labels    = ['Small (0-750)', 'Medium (751-1500)', 'Large (1501-2250)', 'Extra Large (2250+)']

df['job_size_bin'] = pd.cut(
    df['estimated_job_size_sqft_clean'],
    bins=bin_edges,
    labels=labels,
    right=False,
    include_lowest=True
)

# ── STEP 2: CRAMER'S V FOR ALL FACTORS ───────────────────────────────

all_factors = {
    'neighbourhood':         'Neighbourhood',
    'customer_age_bracket':  'Age Bracket',
    'property_type':         'Property Type',
    'homeowner_status':      'Homeowner Status',
    'job_size_bin':          'Job Size',
    'requested_timeline':    'Timeline',
    'referral_source':       'Referral Source',
    'preferred_contact':     'Preferred Contact',
    'lead_capture_weather':  'Weather',
    'has_pets':              'Has Pets',
    'lead_weekday':          'Lead Weekday',
    'distance_to_queens_km': 'Distance from Queens',
}

def cramers_v_full(col):
    df_temp       = df[[col, 'expected_profit_band']].dropna()
    contingency   = pd.crosstab(df_temp[col], df_temp['expected_profit_band'])
    chi2, p, _, _ = chi2_contingency(contingency)
    n             = contingency.sum().sum()
    r, k          = contingency.shape
    v             = np.sqrt(chi2 / (n * (min(r, k) - 1)))
    return round(v, 4), round(p, 4)

def strength(v):
    if v >= 0.5:   return 'Strong'
    elif v >= 0.3: return 'Moderate'
    elif v >= 0.1: return 'Weak'
    else:          return 'Negligible'

def significance(p):
    return 'Significant' if p < 0.05 else 'Not Significant'

def use_in_model(v, p):
    return 'Yes' if v >= 0.1 and p < 0.05 else 'No'

print("=" * 60)
print("STEP 2: CRAMER'S V ANALYSIS")
print("=" * 60)

cramers_rows = []

for col, name in all_factors.items():
    if col not in df.columns:
        print(f"  Skipping {name} - column not found")
        continue
    try:
        v, p = cramers_v_full(col)
        cramers_rows.append({
            'Factor':       name,
            'Column':       col,
            "Cramer's V":   v,
            'Strength':     strength(v),
            'P-Value':      p,
            'Significance': significance(p),
            'Use in Model': use_in_model(v, p),
        })
        print(f"  ✓ {name}: V={v}, p={p}, {strength(v)}, {significance(p)}")
    except Exception as e:
        print(f"  ✗ {name}: Error - {e}")

cramers_df = pd.DataFrame(cramers_rows).sort_values("Cramer's V", ascending=False).reset_index(drop=True)

# ── STEP 3: SELECT SIGNIFICANT FACTORS FOR MODEL ─────────────────────

model_features = [
    row['Column'] for _, row in cramers_df.iterrows()
    if row['Use in Model'] == 'Yes'
]

print(f"\nSignificant features selected: {model_features}")

# ── STEP 4: CALCULATE DATA-DRIVEN WEIGHTS FROM CRAMER'S V ────────────

model_cramers = {
    row['Column']: row["Cramer's V"]
    for _, row in cramers_df.iterrows()
    if row['Use in Model'] == 'Yes'
}

total_v  = sum(model_cramers.values())
weights  = {col: round(v / total_v, 4) for col, v in model_cramers.items()}

print("\n" + "=" * 60)
print("STEP 4: DATA-DRIVEN WEIGHTS")
print("=" * 60)
for col, w in weights.items():
    print(f"  {col}: {w*100:.1f}%")

# ── STEP 5: FREQUENCY TABLES ──────────────────────────────────────────

def build_frequency_table(col):
    df_temp = df[df[col].notna()].copy()
    counts  = df_temp.groupby([col, 'expected_profit_band'], observed=True).size().unstack(fill_value=0)

    for band in profit_bands:
        if band not in counts.columns:
            counts[band] = 0
    counts = counts[profit_bands]

    counts['Total']    = counts.sum(axis=1)
    counts['Total %']  = (counts['Total'] / counts['Total'].sum() * 100).round(1)
    counts['High %']   = (counts['High']   / counts['Total'] * 100).round(1)
    counts['Medium %'] = (counts['Medium'] / counts['Total'] * 100).round(1)
    counts['Low %']    = (counts['Low']    / counts['Total'] * 100).round(1)

    return counts[['High', 'Medium', 'Low', 'Total', 'Total %', 'High %', 'Medium %', 'Low %']]

# ── STEP 6: DATA SCORE LOOKUP (0-10) ─────────────────────────────────

def build_score_lookup(col):
    df_temp = df[df[col].notna()].copy()
    counts  = df_temp.groupby([col, 'expected_profit_band'], observed=True).size().unstack(fill_value=0)

    for band in profit_bands:
        if band not in counts.columns:
            counts[band] = 0

    counts['Total']     = counts[profit_bands].sum(axis=1)
    counts['High %']    = counts['High']   / counts['Total']
    counts['Low %']     = counts['Low']    / counts['Total']
    counts['raw_score'] = counts['High %'] - counts['Low %']

    min_s = counts['raw_score'].min()
    max_s = counts['raw_score'].max()

    counts['score'] = 5.0 if max_s == min_s else \
        ((counts['raw_score'] - min_s) / (max_s - min_s) * 10).round(2)

    return counts['score'].to_dict()

score_lookups = {}
for col in model_features:
    score_lookups[col] = build_score_lookup(col)

# ── STEP 7: QUALITATIVE SCORES FROM LEAD SCORER DOC ──────────────────

qualitative_scores = {
    'neighbourhood': {
        'West End':           9,
        'Strathcona Park':    8,
        'Calvin Park':        7,
        'Sydenham Ward':      6,
        'Portsmouth Village': 4,
        'Downtown':           2,
    },
    'customer_age_bracket': {
        '65+':   10,
        '55-64':  9,
        '45-54':  8,
        '35-44':  5,
        '25-34':  3,
        '18-24':  2,
    },
    'homeowner_status': {
        'Recently Purchased': 10,
        'Own':                 9,
        'Rent':                2,
    },
    'property_type': {
        'Detached':      8,
        'Townhouse':     7,
        'Semi-Detached': 6,
        'Heritage Home': 6,
        'Apartment':     2,
    },
}

qualitative_weights = {
    'neighbourhood':        0.35,
    'customer_age_bracket': 0.25,
    'homeowner_status':     0.25,
    'property_type':        0.15,
}

# ── STEP 8: SANITY CHECK ──────────────────────────────────────────────

print("\n" + "=" * 60)
print("STEP 8: QUALITATIVE SANITY CHECK")
print("=" * 60)

sanity_rows = []

for col, qual_scores in qualitative_scores.items():
    df_temp = df[df[col].notna()].copy()
    counts  = df_temp.groupby([col, 'expected_profit_band'], observed=True).size().unstack(fill_value=0)

    for band in profit_bands:
        if band not in counts.columns:
            counts[band] = 0

    counts['Total']  = counts[profit_bands].sum(axis=1)
    counts['High %'] = (counts['High'] / counts['Total'] * 100).round(1)
    counts['Low %']  = (counts['Low']  / counts['Total'] * 100).round(1)

    for val, qual_score in qual_scores.items():
        if val not in counts.index:
            continue

        actual_high  = counts.loc[val, 'High %']
        actual_low   = counts.loc[val, 'Low %']
        data_score   = round(actual_high - actual_low, 1)
        match        = '✅' if (qual_score >= 6 and actual_high > actual_low) or \
                               (qual_score <= 4 and actual_low > actual_high) else '⚠️'

        # Qualitative score normalized to 0-10
        qual_normalized = round(qual_score / 10 * 10, 2)

        sanity_rows.append({
            'Factor':                   col,
            'Value':                    val,
            'Qualitative Score (1-10)': qual_score,
            'Qualitative Weight':       qualitative_weights.get(col, 'N/A'),
            'Cramer V Weight %':        round(weights.get(col, 0) * 100, 1),
            'Actual High %':            actual_high,
            'Actual Low %':             actual_low,
            'Data Score (High-Low)':    data_score,
            'Match':                    match,
        })

        print(f"  {match} {col} - {val}: Qual={qual_score}, High%={actual_high}, Low%={actual_low}")

sanity_df    = pd.DataFrame(sanity_rows)
matches      = (sanity_df['Match'] == '✅').sum()
total_checks = len(sanity_df)
sanity_pct   = round(matches / total_checks * 100, 1)

print(f"\nSanity Check: {matches}/{total_checks} ({sanity_pct}%) match qualitative research")

# ── STEP 9: CALCULATE FINAL SCORE PER LEAD ───────────────────────────
# Data Score = weighted score from significant Cramer factors
# Qualitative Score = weighted score from lead scorer doc
# Final Score = blend of both (60% data, 40% qualitative)

DATA_WEIGHT = 0.60
QUAL_WEIGHT = 0.40

def get_qual_score(row):
    weighted_sum  = 0
    total_weight  = 0

    for col, qual_scores in qualitative_scores.items():
        value = row.get(col)
        if pd.isna(value):
            continue
        score        = qual_scores.get(value, 5)
        weight       = qualitative_weights.get(col, 0)
        weighted_sum += weight * score
        total_weight += weight

    if total_weight == 0:
        return 5.0

    # Normalize to 0-10
    return round(weighted_sum / total_weight, 2)

def get_data_score(row):
    weighted_sum = 0
    total_weight = 0

    for col in model_features:
        value = row.get(col)
        if pd.isna(value):
            continue
        score        = score_lookups[col].get(value, 5.0)
        weight       = weights[col]
        weighted_sum += weight * score
        total_weight += weight

    if total_weight == 0:
        return 5.0

    return round(weighted_sum / total_weight, 2)

df['data_score']       = df.apply(get_data_score, axis=1)
df['qual_score']       = df.apply(get_qual_score, axis=1)
df['final_score']      = (DATA_WEIGHT * df['data_score'] + QUAL_WEIGHT * df['qual_score']).round(2)

# ── STEP 10: PROFITABILITY PROJECTION + TIER ─────────────────────────

def project_profitability(score):
    if score >= 7.0:   return 'High'
    elif score >= 4.0: return 'Medium'
    else:              return 'Low'

def assign_tier(score):
    if score >= 8.0:   return 'Priority Lead'
    elif score >= 6.5: return 'Strong Lead'
    elif score >= 5.0: return 'Moderate Lead'
    elif score >= 3.0: return 'Weak Lead'
    else:              return 'Pass'

df['projected_profit_band'] = df['final_score'].apply(project_profitability)
df['lead_tier']             = df['final_score'].apply(assign_tier)

# ── STEP 11: ACCURACY + ERROR ANALYSIS ───────────────────────────────

correct  = (df['projected_profit_band'] == df['expected_profit_band']).sum()
total    = len(df)
accuracy = round(correct / total * 100, 1)

print("\n" + "=" * 60)
print("STEP 11: ACCURACY + ERROR ANALYSIS")
print("=" * 60)
print(f"  Overall Accuracy: {correct}/{total} = {accuracy}%")

band_stats = []
for band in profit_bands:
    band_df      = df[df['expected_profit_band'] == band]
    band_correct = (band_df['projected_profit_band'] == band).sum()
    band_total   = len(band_df)
    band_acc     = round(band_correct / band_total * 100, 1) if band_total > 0 else 0
    band_error   = round(100 - band_acc, 1)
    print(f"  {band}: {band_correct}/{band_total} correct = {band_acc}% accuracy, {band_error}% error")
    band_stats.append({
        'Profit Band':  band,
        'Total Leads':  band_total,
        'Correct':      band_correct,
        'Incorrect':    band_total - band_correct,
        'Accuracy %':   band_acc,
        'Error %':      band_error,
    })

# Confusion matrix
confusion = pd.crosstab(
    df['expected_profit_band'],
    df['projected_profit_band'],
    rownames=['Actual'],
    colnames=['Predicted']
)
print("\nConfusion Matrix:")
print(confusion)

# ── SAVE TO EXCEL ─────────────────────────────────────────────────────

with pd.ExcelWriter("lead_scoring_model3.xlsx", engine='openpyxl') as writer:

    # Cramer's V
    cramers_df.drop(columns=['Column']).to_excel(writer, sheet_name="Cramer's V", index=False)
    print("\n✓ Done: Cramer's V")

    # Frequency tables
    feature_names = {
        'neighbourhood':        'Neighbourhood',
        'customer_age_bracket': 'Age Bracket',
        'property_type':        'Property Type',
        'homeowner_status':     'Homeowner Status',
        'job_size_bin':         'Job Size',
        'requested_timeline':   'Timeline',
        'referral_source':      'Referral Source',
        'preferred_contact':    'Preferred Contact',
        'lead_capture_weather': 'Weather',
        'has_pets':             'Has Pets',
        'lead_weekday':         'Lead Weekday',
    }

    for col, name in feature_names.items():
        if col in df.columns:
            build_frequency_table(col).to_excel(writer, sheet_name=name[:31])
            print(f"✓ Done: {name}")

    # Score lookups
    score_rows = []
    for col, lookup in score_lookups.items():
        for val, score in lookup.items():
            score_rows.append({
                'Factor':         col,
                'Value':          val,
                'Score (0-10)':   score,
                'Weight':         weights[col],
                'Weighted Score': round(score * weights[col], 4),
            })
    pd.DataFrame(score_rows).to_excel(writer, sheet_name='Score Lookups', index=False)
    print("✓ Done: Score Lookups")

    # Sanity check
    sanity_df.to_excel(writer, sheet_name='Sanity Check', index=False)
    print("✓ Done: Sanity Check")

    # All leads scored
    output_cols   = [
        'lead_id', 'neighbourhood', 'customer_age_bracket',
        'property_type', 'homeowner_status', 'job_size_bin',
        'requested_timeline', 'expected_profit_band',
        'data_score', 'qual_score', 'final_score',
        'projected_profit_band', 'lead_tier',
    ]
    existing_cols = [c for c in output_cols if c in df.columns]
    df[existing_cols].to_excel(writer, sheet_name='All Leads Scored', index=False)
    print("✓ Done: All Leads Scored")

    # Accuracy summary
    accuracy_df = pd.DataFrame({
        'Metric': [
            'Total Leads',
            'Correct Predictions',
            'Incorrect Predictions',
            'Overall Accuracy %',
            'Overall Error %',
            'Data Score Weight',
            'Qualitative Score Weight',
            'Sanity Check Pass Rate %',
        ],
        'Value': [
            total,
            correct,
            total - correct,
            accuracy,
            round(100 - accuracy, 1),
            f"{int(DATA_WEIGHT*100)}%",
            f"{int(QUAL_WEIGHT*100)}%",
            sanity_pct,
        ]
    })
    accuracy_df.to_excel(writer, sheet_name='Accuracy Summary', index=False)
    print("✓ Done: Accuracy Summary")

    # Per band error analysis
    pd.DataFrame(band_stats).to_excel(writer, sheet_name='Band Error Analysis', index=False)
    print("✓ Done: Band Error Analysis")

    # Confusion matrix
    confusion.to_excel(writer, sheet_name='Confusion Matrix')
    print("✓ Done: Confusion Matrix")

    # Tier summary
    tier_summary = df.groupby(['lead_tier', 'expected_profit_band']).size().unstack(fill_value=0)
    tier_summary['Total'] = tier_summary.sum(axis=1)
    tier_summary.to_excel(writer, sheet_name='Tier Summary')
    print("✓ Done: Tier Summary")

print(f"\nSaved to lead_scoring_model3.xlsx!")
print(f"Overall Accuracy: {accuracy}%")
print(f"Sanity Check: {matches}/{total_checks} ({sanity_pct}%) passed")
