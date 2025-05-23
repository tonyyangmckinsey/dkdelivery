from collections import defaultdict
import pandas as pd

def calculate_avg_mrc_per_day(df, date_columns, steps):
    """
    Calculates the average MRC (Monthly Recurring Charge) per day in each delivery step.
    Returns a DataFrame with columns: ['Step', 'Avg MRC per Day']
    """
    df.columns = df.columns.str.strip()  # Strip columns once
    step_daily_mrc = defaultdict(float)

    for col in date_columns:
        for _, row in df.iterrows():
            step = row[col]
            if step in steps:
                value = row.get("Delta MRC", 0)
                if pd.isna(value) or str(value).strip() == "":
                    mrc = 0
                else:
                    try:
                        mrc = float(str(value).strip())
                    except:
                        mrc = 0
                step_daily_mrc[step] += mrc

    num_days = len(date_columns)
    data = []
    for step in steps:
        total_mrc = step_daily_mrc[step]
        avg_mrc = total_mrc / num_days if num_days > 0 else 0
        data.append({"Step": step, "Avg MRC per Day": round(avg_mrc, 2)})

    return pd.DataFrame(data)
