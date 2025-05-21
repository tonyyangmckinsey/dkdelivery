import pandas as pd
from collections import defaultdict

def calculate_avg_orders_per_day(df, date_columns, steps):
    """
    Calculates the average number of orders in each step per day.
    Returns a DataFrame with columns: ['Step', 'Avg Orders per Day']
    """
    # Initialize count per step per day
    step_daily_counts = defaultdict(int)

    for col in date_columns:
        for _, row in df.iterrows():
            step = row[col]
            if step in steps:
                step_daily_counts[step] += 1

    # Average per step across number of days
    num_days = len(date_columns)
    data = []
    for step in steps:
        total_occurrences = step_daily_counts[step]
        avg_per_day = total_occurrences / num_days if num_days > 0 else 0
        data.append({"Step": step, "Avg Orders per Day": round(avg_per_day, 2)})

    return pd.DataFrame(data)
