import pandas as pd
from collections import defaultdict

def calculate_avg_days_in_step(df, date_columns, steps):
    """
    Calculates the average number of days each order spends in each delivery step.
    Returns a DataFrame with columns: ['Step', 'Avg Days in Step']
    """
    # Dictionary to accumulate total days per step
    step_days = defaultdict(int)
    # Dictionary to count how many orders visited each step
    step_order_counts = defaultdict(int)

    for _, row in df.iterrows():
        previous_step = None
        current_step = None
        days_in_step = 0

        for col in date_columns + [None]:  # Add None as sentinel to trigger final flush
            current_step = row[col] if col else None

            if current_step == previous_step:
                days_in_step += 1
            else:
                if previous_step in steps:
                    step_days[previous_step] += days_in_step
                    step_order_counts[previous_step] += 1
                days_in_step = 1  # Reset for new step
                previous_step = current_step

    # Compute average days per step
    data = []
    for step in steps:
        total_days = step_days[step]
        count = step_order_counts[step]
        avg_days = total_days / count if count > 0 else 0
        data.append({"Step": step, "Avg Days in Step": round(avg_days, 2)})

    return pd.DataFrame(data)
