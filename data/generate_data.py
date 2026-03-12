import pandas as pd
import numpy as np
from datetime import date, timedelta

np.random.seed(42)

START_DATE = date(2026, 1, 1)
END_DATE = date(2026, 3, 31)
dates = pd.date_range(START_DATE, END_DATE, freq="D")

BRANDS = {
    "Apex Athletics": "sportswear",
    "Luminara Beauty": "cosmetics",
    "TechNova": "consumer electronics",
}

CHANNELS = ["Google Search", "Meta (Facebook/Instagram)", "YouTube Video", "Programmatic Display", "TikTok"]

# Channel-specific baseline profiles:
#   impressions_range, ctr_base, cpc_base, conv_rate_base, avg_order_value
CHANNEL_PROFILES = {
    "Google Search":              ((5_000, 50_000),    0.045, 2.50, 0.065, 120),
    "Meta (Facebook/Instagram)":  ((20_000, 300_000),  0.018, 1.20, 0.030, 95),
    "YouTube Video":              ((10_000, 200_000),  0.008, 0.60, 0.015, 80),
    "Programmatic Display":       ((50_000, 500_000),  0.004, 0.35, 0.010, 70),
    "TikTok":                     ((20_000, 300_000),  0.022, 0.80, 0.018, 75),
}

# Brand-level spend weight (share of total weekly budget ~$1M)
BRAND_WEIGHTS = {"Apex Athletics": 0.35, "Luminara Beauty": 0.30, "TechNova": 0.35}

# Channel-level spend weight within each brand (heavier on Meta and Google)
CHANNEL_SPEND_WEIGHTS = {
    "Google Search": 0.30,
    "Meta (Facebook/Instagram)": 0.30,
    "YouTube Video": 0.15,
    "Programmatic Display": 0.10,
    "TikTok": 0.15,
}

rows = []

for single_date in dates:
    day_index = (single_date.date() - START_DATE).days
    day_of_week = single_date.dayofweek  # 0=Mon, 6=Sun
    is_weekend = day_of_week >= 5

    # Weekend dampening factor
    weekend_factor = np.random.uniform(0.75, 0.85) if is_weekend else 1.0

    # Gradual optimization trend: +0 to ~18% improvement over 90 days
    trend_factor = 1.0 + 0.20 * (day_index / 89)

    for brand in BRANDS:
        brand_weight = BRAND_WEIGHTS[brand]

        # TechNova mid-February dip (Feb 8 – Feb 22)
        tech_dip = 1.0
        if brand == "TechNova" and date(2026, 2, 8) <= single_date.date() <= date(2026, 2, 22):
            dip_center = date(2026, 2, 15)
            dist = abs((single_date.date() - dip_center).days)
            tech_dip = 0.55 + 0.05 * dist  # deepest ~0.55 at center, tapering out

        for channel in CHANNELS:
            imp_range, ctr_base, cpc_base, conv_base, aov = CHANNEL_PROFILES[channel]
            ch_spend_weight = CHANNEL_SPEND_WEIGHTS[channel]

            # Daily base spend: ~$1M/week ÷ 7 days ≈ $142K/day total
            daily_budget_share = (1_000_000 / 7) * brand_weight * ch_spend_weight
            spend_noise = np.random.uniform(0.85, 1.15)
            spend = daily_budget_share * weekend_factor * spend_noise * tech_dip
            spend = round(spend, 2)

            # Impressions
            base_imp = np.random.uniform(*imp_range)
            impressions = int(base_imp * weekend_factor * trend_factor * tech_dip)

            # CTR with noise and trend
            ctr = ctr_base * trend_factor * np.random.uniform(0.85, 1.15) * tech_dip
            ctr = min(ctr, 0.15)
            clicks = max(1, int(impressions * ctr))

            # CPC derives from spend/clicks but we add noise
            cpc = spend / max(clicks, 1)

            # Conversion rate with trend
            conv_rate = conv_base * trend_factor * np.random.uniform(0.85, 1.15) * tech_dip
            conv_rate = min(conv_rate, 0.25)
            conversions = max(0, int(clicks * conv_rate))

            # Revenue
            revenue_per_conv = aov * np.random.uniform(0.80, 1.20) * trend_factor
            revenue = round(conversions * revenue_per_conv, 2)

            # Derived metrics
            roas = round(revenue / spend, 4) if spend > 0 else 0
            cpa = round(spend / conversions, 2) if conversions > 0 else 0
            cpm = round((spend / impressions) * 1000, 2) if impressions > 0 else 0

            campaign_name = f"{brand} — {channel} Q1"

            rows.append({
                "date": single_date.strftime("%Y-%m-%d"),
                "brand": brand,
                "campaign_name": campaign_name,
                "channel": channel,
                "impressions": impressions,
                "clicks": clicks,
                "ctr": round(ctr, 6),
                "spend_usd": spend,
                "conversions": conversions,
                "conversion_rate": round(conv_rate, 6),
                "revenue_usd": revenue,
                "roas": roas,
                "cpa_usd": cpa,
                "cpc_usd": round(cpc, 2),
                "cpm_usd": cpm,
            })

df = pd.DataFrame(rows)
output_path = "campaign_data.csv"
df.to_csv(output_path, index=False)

print(f"Generated {len(df)} rows → {output_path}")
print(f"Date range: {df['date'].min()} to {df['date'].max()}")
print(f"Brands: {df['brand'].unique().tolist()}")
print(f"Campaigns: {df['campaign_name'].nunique()}")
print(f"\nSample stats:")
print(df.describe().to_string())
