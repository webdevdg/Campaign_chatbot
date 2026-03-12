import os
from typing import Optional

import pandas as pd

_df: Optional[pd.DataFrame] = None

DATA_PATH = os.path.join(os.path.dirname(__file__), "..", "..", "data", "campaign_data.csv")


def load_data() -> pd.DataFrame:
    global _df
    if _df is None:
        _df = pd.read_csv(DATA_PATH, parse_dates=["date"])
    return _df


def get_data_summary() -> str:
    df = load_data()
    lines = [
        "=== Dataset Summary ===",
        f"Rows: {len(df)}, Columns: {list(df.columns)}",
        f"Date range: {df['date'].min().date()} to {df['date'].max().date()}",
        f"Brands: {df['brand'].unique().tolist()}",
        f"Channels: {df['channel'].unique().tolist()}",
        f"Campaigns: {df['campaign_name'].unique().tolist()}",
        "",
        "=== Descriptive Statistics ===",
        df.describe().to_string(),
    ]
    return "\n".join(lines)


def get_csv_text() -> str:
    df = load_data()
    return df.to_csv(index=False)
