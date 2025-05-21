import pandas as pd
from collections import defaultdict

def calculate_total_mrc(df, date_columns, steps):
    """
    Calculates the total MRC (Monthly Recurring Charge) across all orders that passed through each step.
    Returns a DataFrame with columns: ['Step', 'Total MRC (Unique Orders)']
    """
    step_mrc = defaultdict(float)
    order_seen_in_step = defaultdict(set)

    for _, row in df.iterrows():
        order_id = (row["ServiceID_Crid"], row["SalesProjectID_ContractNumber"])
        mrc_raw = row.get("Delta MRC", 0)

        try:
            mrc = float(str(mrc_raw).strip()) if pd.notna(mrc_raw) else 0
        except:
            mrc = 0

        seen_steps = set()

        for col in date_columns:
            step = row[col]
            if step in steps and step not in seen_steps:
                if order_id not in order_seen_in_step[step]:
                    step_mrc[step] += mrc
                    order_seen_in_step[step].add(order_id)
                seen_steps.add(step)

    # Build result table
    data = []
    for step in steps:
        data.append({"Step": step, "Total MRC (Unique Orders)": round(step_mrc[step], 2)})

    return pd.DataFrame(data)
