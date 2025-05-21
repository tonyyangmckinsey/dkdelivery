import pandas as pd
from collections import defaultdict

def calculate_unique_orders(df, date_columns, steps):
    """
    Calculates how many unique orders ever entered each delivery step.
    Returns a DataFrame with columns: ['Step', '# Orders (Unique)']
    """
    # Dictionary to track whether an order has been in a step
    step_orders = defaultdict(set)

    for _, row in df.iterrows():
        order_id = (row["ServiceID_Crid"], row["SalesProjectID_ContractNumber"])
        seen_steps = set()

        for col in date_columns:
            step = row[col]
            if step in steps and step not in seen_steps:
                step_orders[step].add(order_id)
                seen_steps.add(step)

    # Build the result table
    data = []
    for step in steps:
        count = len(step_orders[step])
        data.append({"Step": step, "# Orders (Unique)": count})

    return pd.DataFrame(data)
