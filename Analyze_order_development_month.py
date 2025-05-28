import pandas as pd
import os
import glob
import re
import time
from openpyxl import load_workbook
from metric_avg_days_in_step import calculate_avg_days_in_step
from metric_avg_orders_per_day import calculate_avg_orders_per_day
from metric_avg_mrc_per_day import calculate_avg_mrc_per_day
from metric_orders_in_out import calculate_orders_in_out
from individual_avg_orders_per_day import calculate_avg_orders_per_day_per_dpm
from individual_avg_mrc_per_day import calculate_avg_mrc_per_day_per_dpm
from individual_avg_days_in_step import calculate_avg_days_in_step_per_dpm
from individual_orders_in_out import calculate_orders_in_out_per_dpm

def normalize_keys(df):
    df["ServiceID_Crid"] = df["ServiceID_Crid"].astype(str).str.strip()
    df["SalesProjectID_ContractNumber"] = df["SalesProjectID_ContractNumber"].astype(str).str.strip()
    return df

def extract_date_columns(df):
    return [col for col in df.columns if re.match(r"^\d{8}$", str(col)) and df[col].notna().any()]

def run_filtered_analysis():
    start_time = time.time()

    output_folder = r"C:\Users\Tony Yang\OneDrive - McKinsey & Company\Documents\Python\AnalysisOutputFolder"
    pattern = os.path.join(output_folder, "*_Order_Development_Output.xlsx")
    files = glob.glob(pattern)
    if not files:
        raise FileNotFoundError("❌ No Order_Development_Output.xlsx files found.")
    file_path = max(files, key=os.path.getmtime)

    df = pd.read_excel(file_path, engine="openpyxl")
    df = normalize_keys(df)

    # Ask user for Committed RFS month
    target_month = input("Enter Committed RFS month to filter (format: YYYY-MM): ").strip()
    df["Committed RFS"] = pd.to_datetime(df["Committed RFS"], errors="coerce")
    df = df[df["Committed RFS"].dt.strftime("%Y-%m") == target_month]

    if df.empty:
        print("⚠️ No orders found for the selected month.")
        return

    date_columns = extract_date_columns(df)
    print(f"✅ Filtered {len(df)} rows for Committed RFS month: {target_month}")
    print("📅 Date columns:", date_columns)

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

    tracked_steps = [
        "1.9 PMO delivery planning",
        "2. Delivery planning",
        "3. Installation preparation and digging order",
        "4. Final design",
        "5. Configuration",
        "6. Cabling splicing and installation",
        "7. Service activation",
        "8. Ready for service"
    ]

    avg_days_df = calculate_avg_days_in_step(df, date_columns, steps)
    avg_orders_per_day_df = calculate_avg_orders_per_day(df, date_columns, steps)
    avg_mrc_per_day_df = calculate_avg_mrc_per_day(df, date_columns, steps)
    summary_df = avg_days_df.merge(avg_orders_per_day_df, on="Step").merge(avg_mrc_per_day_df, on="Step")

    T1 = input("Enter T1 date for order in & out analysis (format: YYYYMMDD): ")
    T2 = input("Enter T2 date for order in & out analysis (format: YYYYMMDD): ")
    in_out_df = calculate_orders_in_out(df, date_columns, steps, T1, T2)

    avg_orders_per_dpm_df = calculate_avg_orders_per_day_per_dpm(df, date_columns, steps)
    avg_mrc_per_dpm_df = calculate_avg_mrc_per_day_per_dpm(df, date_columns, steps)
    avg_days_per_dpm_df = calculate_avg_days_in_step_per_dpm(df, date_columns, steps)
    
    output_path = os.path.join(
        r"C:\Users\Tony Yang\OneDrive - McKinsey & Company\Documents\Python\FinalOutputFolderMonth",
        os.path.basename(file_path).replace(".xlsx", f"_metrics_summary_in_{target_month}.xlsx")
    )


    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, index=False, sheet_name="Step Metrics Summary")
        in_out_df.to_excel(writer, index=False, sheet_name="Step Entry-Exit Summary")
        avg_orders_per_dpm_df.to_excel(writer, index=False, sheet_name="DPM Avg Orders per Day")
        avg_mrc_per_dpm_df.to_excel(writer, index=False, sheet_name="DPM Avg MRC per Day")
        avg_days_per_dpm_df.to_excel(writer, index=False, sheet_name="DPM Avg Days in Step")

        for step in tracked_steps:
            per_step_df = calculate_orders_in_out_per_dpm(df, date_columns, step, T1, T2)
            writer.book.create_sheet(step[:31])
            per_step_df.to_excel(writer, index=False, sheet_name=step[:31])

    print(f"📤 Filtered summary saved to: {output_path}")
    print(f"⏱️ Execution time: {time.time() - start_time:.2f} seconds")
