import pandas as pd
import os
from datetime import datetime

today_str = datetime.today().strftime("%Y%m%d")  # Format: YYYYMMDD
snapshot_date_str = "20250521"  # <-- snapshot date to process

def add_initial_new_orders():
    # === Setup paths ===
    base_path = r"C:\Users\Tony Yang\OneDrive - McKinsey & Company\Documents\Python"
    snapshot_path = os.path.join(base_path, "3. DKTrackerSnapshots", f"{snapshot_date_str}_DK_backlog_snapshot.xlsx")
    master_file_path = os.path.join(base_path, "AnalysisOutputFolder", "Populated_UniqueCrids.xlsx")
    output_file_path = os.path.join(base_path, "AnalysisOutputFolder", f"{today_str}_Order_Development_Output.xlsx")

    # === Load master file and snapshot ===
    print("📥 Loading master file and new snapshot...")
    master_df = pd.read_excel(master_file_path, engine="openpyxl")
    snapshot_df = pd.read_excel(snapshot_path, sheet_name=0, engine="openpyxl")

    # === Normalize key columns ===
    master_df["ServiceID_Crid"] = master_df["ServiceID_Crid"].astype(str).str.strip()
    master_df["SalesProjectID_ContractNumber"] = master_df["SalesProjectID_ContractNumber"].astype(str).str.strip()
    snapshot_df["ServiceID_Crid"] = snapshot_df["ServiceID_Crid"].astype(str).str.strip()
    snapshot_df["SalesProjectID_ContractNumber"] = snapshot_df["SalesProjectID_ContractNumber"].astype(str).str.strip()

    # === Identify new orders ===
    master_keys = set(zip(master_df["ServiceID_Crid"], master_df["SalesProjectID_ContractNumber"]))
    snapshot_keys = set(zip(snapshot_df["ServiceID_Crid"], snapshot_df["SalesProjectID_ContractNumber"]))
    new_keys = snapshot_keys - master_keys
    print(f"🆕 Found {len(new_keys)} new orders to add.")

    # === Columns to bring from snapshot
    cols_to_keep = [
        "ServiceID_Crid",
        "SalesProjectID_ContractNumber",
        "Customer name",
        "Delta MRC",
        "Delivery Project Manager",
        "Segment Red.",
        "ProjectTypeValue",
        "Input: Which delivery step is the order in? (Drop down list)"
    ]

    new_rows = snapshot_df[
        snapshot_df.set_index(["ServiceID_Crid", "SalesProjectID_ContractNumber"]).index.isin(new_keys)
    ][cols_to_keep].copy()

    # === Add 20250521 column with delivery step info
    new_rows[snapshot_date_str] = new_rows["Input: Which delivery step is the order in? (Drop down list)"]
    new_rows.drop(columns=["Input: Which delivery step is the order in? (Drop down list)"], inplace=True)

    # === Add empty columns for other dates ===
    date_columns = [col for col in master_df.columns if col.isdigit() and col != snapshot_date_str]
    for col in date_columns:
        new_rows[col] = ""

    # === Reorder new rows to match master columns ===
    ordered_columns = list(master_df.columns)
    new_rows = new_rows.reindex(columns=ordered_columns)

    # === Append new rows ===
    updated_df = pd.concat([master_df, new_rows], ignore_index=True)

    # === Update 20250521 column for existing orders ===
    print("🔄 Updating existing rows with delivery step for 20250521...")
    delivery_col = "Input: Which delivery step is the order in? (Drop down list)"
    snapshot_lookup = snapshot_df.set_index(["ServiceID_Crid", "SalesProjectID_ContractNumber"])[delivery_col].to_dict()

    filled_count = 0
    for idx, row in updated_df.iterrows():
        key = (row["ServiceID_Crid"], row["SalesProjectID_ContractNumber"])
        if key in snapshot_lookup:
            updated_df.at[idx, snapshot_date_str] = snapshot_lookup[key]
            filled_count += 1

    print(f"✅ {filled_count} total rows updated for {snapshot_date_str}.")

    # === Export final updated file ===
    updated_df.to_excel(output_file_path, index=False)
    print(f"📤 Final file saved to: {output_file_path}")
