import pandas as pd
import os
from datetime import datetime, timedelta

def consolidate_performance_sheets():
    # === Setup paths ===
    base_path = r"C:\Users\Tony Yang\OneDrive - McKinsey & Company\Documents\Python"
    snapshot_folder = os.path.join(base_path, "3. DKTrackerSnapshots")
    output_folder = os.path.join(base_path, "ConsolidationOutputFolder")
    overall_output_path = os.path.join(output_folder, "Consolidated_overall_performance_heatmap.xlsx")
    individual_output_path = os.path.join(output_folder, "Consolidated_individual_performance_heatmap.xlsx")

    # === Create ExcelWriter objects ===
    overall_writer = pd.ExcelWriter(overall_output_path, engine="openpyxl", mode='w')
    individual_writer = pd.ExcelWriter(individual_output_path, engine="openpyxl", mode='w')

    # === Loop through dates from 20250521 to 20251231 ===
    start_date = datetime.strptime("20250521", "%Y%m%d")
    end_date = datetime.strptime("20251231", "%Y%m%d")
    current_date = start_date

    while current_date <= end_date:
        date_str = current_date.strftime("%Y%m%d")
        snapshot_filename = f"{date_str}_DK_backlog_snapshot.xlsx"
        snapshot_path = os.path.join(snapshot_folder, snapshot_filename)

        print(f"🔍 Processing snapshot: {snapshot_filename}")
        if not os.path.exists(snapshot_path):
            print(f"⚠️ Snapshot not found: {snapshot_filename}")
            current_date += timedelta(days=1)
            continue

        try:
            # Load sheet 2 and 3 by index
            overall_df = pd.read_excel(snapshot_path, sheet_name=1, engine="openpyxl")  # Sheet index 1 = "2.1"
            individual_df = pd.read_excel(snapshot_path, sheet_name=2, engine="openpyxl")  # Sheet index 2 = "2.2"

            # Write to respective output files using date as sheet name
            overall_df.to_excel(overall_writer, sheet_name=date_str, index=False)
            individual_df.to_excel(individual_writer, sheet_name=date_str, index=False)
            print(f"✅ Stored sheets for {date_str}")

        except Exception as e:
            print(f"❌ Failed to process {snapshot_filename}: {e}")

        current_date += timedelta(days=1)

    # === Save the consolidated files ===
    overall_writer.close()
    individual_writer.close()
    print("📤 Consolidated files saved successfully:")
    print(f"   - {overall_output_path}")
    print(f"   - {individual_output_path}")


