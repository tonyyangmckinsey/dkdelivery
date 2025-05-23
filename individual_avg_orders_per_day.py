from collections import defaultdict
import pandas as pd

def calculate_avg_orders_per_day_per_dpm(df, date_columns, steps):
    """
    Calculates the average number of orders in each step per day, grouped by Delivery Project Manager.
    Returns a DataFrame with Delivery Project Managers as rows and Steps as columns.
    """
    step_daily_counts = defaultdict(lambda: defaultdict(int))  # {dpm: {step: count}}

    for col in date_columns:
        for _, row in df.iterrows():
            step = row[col]
            dpm = row.get("Delivery Project Manager", "Unknown")
            if step in steps:
                step_daily_counts[dpm][step] += 1

    num_days = len(date_columns)
    dpm_list = sorted(step_daily_counts.keys())
    data = []

    for dpm in dpm_list:
        row_data = {"Delivery Project Manager": dpm}
        for step in steps:
            total = step_daily_counts[dpm].get(step, 0)
            row_data[step] = round(total / num_days, 2)
        data.append(row_data)

    return pd.DataFrame(data)
