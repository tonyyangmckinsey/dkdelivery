from collections import defaultdict
import pandas as pd

def calculate_avg_mrc_per_day_per_dpm(df, date_columns, steps):
    """
    Calculates the average MRC per day in each delivery step per Delivery Project Manager.
    Returns a DataFrame with Delivery Project Managers as rows and Steps as columns.
    """
    step_daily_mrc = defaultdict(lambda: defaultdict(float))  # {dpm: {step: mrc_sum}}

    for col in date_columns:
        for _, row in df.iterrows():
            step = row[col]
            dpm = row.get("Delivery Project Manager", "Unknown")

            if step in steps:
                value = row.get("Delta MRC", 0)
                if pd.isna(value) or str(value).strip() == "":
                    mrc = 0
                else:
                    try:
                        mrc = float(str(value).strip())
                    except:
                        mrc = 0

                step_daily_mrc[dpm][step] += mrc

    num_days = len(date_columns)
    data = []
    for dpm in sorted(step_daily_mrc.keys()):
        row_data = {"Delivery Project Manager": dpm}
        for step in steps:
            total_mrc = step_daily_mrc[dpm].get(step, 0)
            row_data[step] = round(total_mrc / num_days, 2)
        data.append(row_data)

    return pd.DataFrame(data)
