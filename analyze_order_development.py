import pandas as pd
import os
import glob
import re
import time
from metric_avg_days_in_step import calculate_avg_days_in_step
from metric_orders_unique import calculate_unique_orders
from metric_avg_orders_per_day import calculate_avg_orders_per_day
from metric_total_mrc import calculate_total_mrc
from metric_avg_mrc_per_day import calculate_avg_mrc_per_day
from metric_orders_in_out import calculate_orders_in_out
# from metric_still_in_step import calculate_percent_still_in_step

start_time = time.time()

def normalize_keys(df):
    df["ServiceID_Crid"] = df["ServiceID_Crid"].astype(str).str.strip()
    df["SalesProjectID_ContractNumber"] = df["SalesProjectID_ContractNumber"].astype(str).str.strip()
    return df

def extract_date_columns(df):
    return [col for col in df.columns if re.match(r"^\\d{8}$", str(col)) and df[col].notna().any()]

def analyze_order_development(file_path):
    df = pd.read_excel(file_path, engine="openpyxl")
    df = normalize_keys(df)
    date_columns = extract_date_columns(df)
    print("Her er datecolumns", date_columns)

    steps = [
        "1. Order entry and validation",
        "1.9 PMO delivery planning",
        "2. Delivery planning",
        "3. Installation preparation and digging order",
        "4. Final design",
        "5. Configuration",
        "6. Cabling splicing and installation",
        "7. Service activation",
        "8. Ready for service",
        "9. Vendor invoice",
        "Delivered",
        "Cancelled/terminated/On hold"
    ]

    avg_days_df = calculate_avg_days_in_step(df, date_columns, steps)
    avg_orders_per_day_df = calculate_avg_orders_per_day(df, date_columns, steps)
    avg_mrc_per_day_df = calculate_avg_mrc_per_day(df, date_columns, steps)

    summary_df = avg_days_df.merge(avg_orders_per_day_df, on="Step").merge(avg_mrc_per_day_df, on="Step")

    # === New: Orders In/Out Analysis ===
    T1 = "20250521"
    T2 = "20250523"
    in_out_df = calculate_orders_in_out(df, date_columns, steps, T1, T2)

    # === Write both tables to same Excel file ===
    output_path = os.path.join(
        r"C:\Users\Tony Yang\OneDrive - McKinsey & Company\Documents\Python\FinalOutputFolder",
        os.path.basename(file_path).replace(".xlsx", "_step_metrics_summary.xlsx")
    )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, index=False, sheet_name="Step Metrics Summary")
        in_out_df.to_excel(writer, index=False, sheet_name="Step Entry-Exit Summary")

    print(f"✅ Summary exported to: {output_path}")

def get_latest_order_output(folder):
    pattern = os.path.join(folder, "*_Order_Development_Output.xlsx")
    files = glob.glob(pattern)
    if not files:
        raise FileNotFoundError("❌ No Order_Development_Output.xlsx files found.")
    latest_file = max(files, key=os.path.getmtime)
    return latest_file

# Run
output_folder = r"C:\Users\Tony Yang\OneDrive - McKinsey & Company\Documents\Python\AnalysisOutputFolder"
file_path = get_latest_order_output(output_folder)
analyze_order_development(file_path)

end_time = time.time()
print(f"⏱️ Execution time: {end_time - start_time:.2f} seconds")
