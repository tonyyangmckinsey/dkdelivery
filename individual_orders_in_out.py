import pandas as pd
from collections import defaultdict

def calculate_orders_in_out_per_dpm(df, date_columns, step, T1, T2):
    T1_index = date_columns.index(T1)
    T2_index = date_columns.index(T2) + 1
    selected_dates = date_columns[T1_index:T2_index]

    stats = defaultdict(lambda: {"Entered": 0, "Stayed": 0, "Exited": 0})

    for _, row in df.iterrows():
        dpm = row.get("Delivery Project Manager", "Unknown")
        order_id = f"{row['ServiceID_Crid']}__{row['SalesProjectID_ContractNumber']}"
        step_sequence = [row[date] for date in selected_dates if pd.notna(row[date])]

        if not step_sequence:
            continue

        at_T1 = step_sequence[0]
        at_T2 = step_sequence[-1]

        if step in step_sequence and at_T1 != step:
            stats[dpm]["Entered"] += 1

        if all(s == step for s in step_sequence):
            stats[dpm]["Stayed"] += 1

        if step in step_sequence and at_T2 != step:
            stats[dpm]["Exited"] += 1

    data = []
    for dpm in sorted(stats.keys()):
        row = {"Delivery Project Manager": dpm}
        row.update(stats[dpm])
        data.append(row)

    return pd.DataFrame(data)
