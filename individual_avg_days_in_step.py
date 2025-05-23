from collections import defaultdict
import pandas as pd

def calculate_avg_days_in_step_per_dpm(df, date_columns, steps):
    """
    Calculates average number of days each order spends in each delivery step, grouped by Delivery Project Manager.
    Returns a DataFrame with Delivery Project Managers as rows and Steps as columns.
    """
    step_days = defaultdict(lambda: defaultdict(int))        # {dpm: {step: total_days}}
    step_order_counts = defaultdict(lambda: defaultdict(int))  # {dpm: {step: num_orders}}

    for _, row in df.iterrows():
        dpm = row.get("Delivery Project Manager", "Unknown")
        previous_step = None
        current_step = None
        days_in_step = 0

        for col in date_columns + [None]:  # Add sentinel
            current_step = row[col] if col else None

            if current_step == previous_step:
                days_in_step += 1
            else:
                if previous_step in steps:
                    step_days[dpm][previous_step] += days_in_step
                    step_order_counts[dpm][previous_step] += 1
                days_in_step = 1
                previous_step = current_step

    # Build DataFrame
    data = []
    for dpm in sorted(step_days.keys()):
        row_data = {"Delivery Project Manager": dpm}
        for step in steps:
            total = step_days[dpm].get(step, 0)
            count = step_order_counts[dpm].get(step, 0)
            row_data[step] = round(total / count, 2) if count > 0 else 0
        data.append(row_data)

    return pd.DataFrame(data)
