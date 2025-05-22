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
from consolidate_heatmap_performance_sheets import consolidate_performance_sheets
#from metric_orders_entered_exited import calculate_orders_in_out
#from metric_still_in_step import calculate_percent_still_in_step

start_time = time.time()


def normalize_keys(df):
    """
    Ensures key identifier columns are stripped and treated as strings.
    """
    df["ServiceID_Crid"] = df["ServiceID_Crid"].astype(str).str.strip()
    df["SalesProjectID_ContractNumber"] = df["SalesProjectID_ContractNumber"].astype(str).str.strip()
    return df

def extract_date_columns(df):
    date_cols = []
    for col in df.columns:
        if re.match(r"^\d{8}$", str(col)):
            if df[col].notna().any():  # At least one non-null value
                date_cols.append(col)
    return date_cols

def analyze_order_development(file_path):
    # === Load the full Order_Development_Output Excel ===
    df = pd.read_excel(file_path, engine="openpyxl")

    # === Normalize order keys and prepare date columns ===
    df = normalize_keys(df)                # e.g., strip spaces, convert types
    date_columns = extract_date_columns(df)  # identifies all date columns (YYYYMMDD)
    print("Her er datecolumns", date_columns)

    # === Prepare delivery step list ===
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

    # === Compute each metric (each returns a DataFrame with 'Step' as index) ===
    avg_days_df = calculate_avg_days_in_step(df, date_columns, steps)
    #unique_orders_df = calculate_unique_orders(df, date_columns, steps)
    avg_orders_per_day_df = calculate_avg_orders_per_day(df, date_columns, steps)
    #total_mrc = calculate_total_mrc(df, date_columns, steps)
    avg_mrc_per_day_df = calculate_avg_mrc_per_day(df, date_columns, steps)
    #in_out_df = calculate_orders_in_out(df, date_columns, steps)
    #still_in_step_df = calculate_percent_still_in_step(df, date_columns, steps)

    # === Merge all metric tables into one consolidated summary ===
    summary_df = avg_days_df.merge(avg_orders_per_day_df, on="Step").merge(avg_mrc_per_day_df, on="Step")
    #    .merge(in_out_df, on="Step")
    #    .merge(still_in_step_df, on="Step")
    #)


 # === Calculate in/out analysis between T1 and T2 ===
    T1 = "20250519"  # <-- can change these dynamically later
    T2 = "20250521"
    in_out_df = calculate_orders_in_out(df, date_columns, steps, T1, T2)

    # === Export both summary and in/out analysis to same Excel file ===
    output_path = os.path.join(
        r"C:\Users\Tony Yang\OneDrive - McKinsey & Company\Documents\Python\FinalOutputFolder",
        os.path.basename(file_path).replace(".xlsx", "_step_metrics_summary.xlsx")
    )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, index=False, sheet_name="Step Metrics Summary")
        in_out_df.to_excel(writer, index=False, sheet_name="Step Entry-Exit Summary")


        # === Export both summary and in/out analysis to same Excel file ===
    output_path = os.path.join(
        r"C:\Users\Tony Yang\OneDrive - McKinsey & Company\Documents\Python\FinalOutputFolder",
        os.path.basename(file_path).replace(".xlsx", "_step_metrics_summary.xlsx")
    )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, index=False, sheet_name="Step Metrics Summary")
        in_out_df.to_excel(writer, index=False, sheet_name="Step Entry-Exit Summary")

    # === Now write custom notes to cell A15 in both sheets ===
    from openpyxl import load_workbook

    wb = load_workbook(output_path)
    sheet1 = wb["Step Metrics Summary"]
    sheet2 = wb["Step Entry-Exit Summary"]

    # Use latest snapshot date in the file
    start_date = "20250519"
    today_date = date_columns[-1]  # latest available snapshot

    # Sheet 1 note
    sheet1["A15"] = f"Data as of {start_date} - {today_date}"

    # Sheet 2 note
    sheet2["A15"] = f"Interval between T1 and T2: {T1} - {T2}"

    wb.save(output_path)

    print(f"✅ Summary exported to: {output_path}")

def get_latest_order_output(folder):
    pattern = os.path.join(folder, "*_Order_Development_Output.xlsx")
    files = glob.glob(pattern)
    if not files:
        raise FileNotFoundError("❌ No Order_Development_Output.xlsx files found.")
    latest_file = max(files, key=os.path.getmtime)
    return latest_file

# Define folder where the outputs are stored
output_folder = r"C:\Users\Tony Yang\OneDrive - McKinsey & Company\Documents\Python\AnalysisOutputFolder"
file_path = get_latest_order_output(output_folder)
analyze_order_development(file_path)

end_time = time.time()
elapsed_time = end_time - start_time
print(f"⏱️ Execution time: {elapsed_time:.2f} seconds")






