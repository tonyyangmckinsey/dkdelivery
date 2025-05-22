import pandas as pd
from collections import defaultdict

def calculate_orders_in_out(df, date_columns, steps, T1, T2):
    """
    Calculates for each step between T1 and T2:
    - Number of unique orders that entered
    - Number of unique orders that stayed
    - Number of unique orders that exited
    
    Returns a DataFrame with columns: ['Step', 'Entered', 'Stayed', 'Exited']
    """
    # Validate that T1 and T2 are within the date columns
    print(date_columns)
    if T1 not in date_columns or T2 not in date_columns:
        raise ValueError(f"T1 ({T1}) and/or T2 ({T2}) not found in date columns.")

    # Slice the snapshot period
    T1_index = date_columns.index(T1)
    T2_index = date_columns.index(T2) + 1  # inclusive
    selected_dates = date_columns[T1_index:T2_index]

    result = []

    for step in steps:
        entered_orders = set()
        stayed_orders = set()
        exited_orders = set()

        for _, row in df.iterrows():
            order_id = f"{row['ServiceID_Crid']}__{row['SalesProjectID_ContractNumber']}"
            step_sequence = [row[date] for date in selected_dates if pd.notna(row[date])]

            if not step_sequence:
                continue

            at_T1 = step_sequence[0]
            at_T2 = step_sequence[-1]
            was_in_step = [s == step for s in step_sequence]

            # Check for entry
            if step in step_sequence and at_T1 != step:
                entered_orders.add(order_id)

            # Check for stayed
            if all(s == step for s in step_sequence):
                stayed_orders.add(order_id)

            # Check for exit
            if step in step_sequence:
                if at_T2 != step:
                    exited_orders.add(order_id)

        result.append({
            "Step": step,
            "Entered": len(entered_orders),
            "Stayed": len(stayed_orders),
            "Exited": len(exited_orders)
        })

    return pd.DataFrame(result)
