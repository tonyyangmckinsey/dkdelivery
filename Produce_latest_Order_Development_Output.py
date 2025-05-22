import pandas as pd
import os
import re
from datetime import datetime

def update_solution_from_snapshots():
    # === Setup paths ===
    base_path = r"C:\Users\Tony Yang\OneDrive - McKinsey & Company\Documents\Python"
    snapshot_folder = os.path.join(base_path, "3. DKTrackerSnapshots")
    output_folder = os.path.join(base_path, "AnalysisOutputFolder")

    # === Identify the latest Order_Development_Output file ===
    files = os.listdir(output_folder)
    solution_files = [f for f in files if re.match(r"\d{8}_Order_Development_Output\.xlsx", f)]
    latest_file = sorted(solution_files)[-1]
    latest_date_str = latest_file.split("_")[0]
    print(f"🗂️ Latest solution file found: {latest_file}")

    master_file_path = os.path.join(output_folder, latest_file)
    master_df = pd.read_excel(master_file_path, engine="openpyxl")

    # === Normalize master keys ===
    master_df["ServiceID_Crid"] = master_df["ServiceID_Crid"].astype(str).str.strip()
    master_df["SalesProjectID_ContractNumber"] = master_df["SalesProjectID_ContractNumber"].astype(str).str.strip()

    # === Identify all snapshot files newer than the latest solution ===
    snapshot_files = sorted([
        f for f in os.listdir(snapshot_folder)
        if re.match(r"\d{8}_DK_backlog_snapshot\.xlsx", f)
        and f[:8] > latest_date_str
    ])
    

    if not snapshot_files:
        print("✅ No new snapshot files to process.")
        return

    print(f"📦 Found {len(snapshot_files)} snapshot(s) to process: {[f[:8] for f in snapshot_files]}")

    for snapshot_file in snapshot_files:
        snapshot_date = snapshot_file[:8]
        snapshot_path = os.path.join(snapshot_folder, snapshot_file)
        print(f"🔁 Processing snapshot: {snapshot_file}")

        try:
            snapshot_df = pd.read_excel(snapshot_path, sheet_name=0, engine="openpyxl")
        except Exception as e:
            print(f"⚠️ Failed to read snapshot {snapshot_file}: {e}")
            continue

        # === Normalize snapshot keys ===
        snapshot_df["ServiceID_Crid"] = snapshot_df["ServiceID_Crid"].astype(str).str.strip()
        snapshot_df["SalesProjectID_ContractNumber"] = snapshot_df["SalesProjectID_ContractNumber"].astype(str).str.strip()

        # === Identify new orders ===
        master_keys = set(zip(master_df["ServiceID_Crid"], master_df["SalesProjectID_ContractNumber"]))
        snapshot_keys = set(zip(snapshot_df["ServiceID_Crid"], snapshot_df["SalesProjectID_ContractNumber"]))
        new_keys = snapshot_keys - master_keys
        print(f"🆕 {len(new_keys)} new orders to add from {snapshot_date}")

        # === Add new orders ===
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

        # Fill in current date column for new rows
        new_rows[snapshot_date] = new_rows["Input: Which delivery step is the order in? (Drop down list)"]
        new_rows.drop(columns=["Input: Which delivery step is the order in? (Drop down list)"], inplace=True)

        # Fill blanks for all other date columns in master
        date_columns = [col for col in master_df.columns if col.isdigit() and col != snapshot_date]
        for col in date_columns:
            new_rows[col] = ""

        # Ensure column alignment
        for col in master_df.columns:
            if col not in new_rows.columns:
                new_rows[col] = ""

        new_rows = new_rows[master_df.columns]
        master_df = pd.concat([master_df, new_rows], ignore_index=True)

        # === Update delivery step value for all orders ===
        print(f"✏️ Updating delivery steps in column {snapshot_date}...")
        delivery_col = "Input: Which delivery step is the order in? (Drop down list)"
        snapshot_lookup = snapshot_df.set_index(["ServiceID_Crid", "SalesProjectID_ContractNumber"])[delivery_col].to_dict()

        filled = 0
        for idx, row in master_df.iterrows():
            key = (row["ServiceID_Crid"], row["SalesProjectID_ContractNumber"])
            if key in snapshot_lookup:
                master_df.at[idx, snapshot_date] = snapshot_lookup[key]
                filled += 1
        print(f"✅ Updated {filled} rows for {snapshot_date}")

    # === Export final updated solution ===
    latest_snapshot_date = snapshot_files[-1][:8]
    output_file = os.path.join(output_folder, f"{latest_snapshot_date}_Order_Development_Output.xlsx")
    master_df.to_excel(output_file, index=False)
    print(f"📤 Final solution saved to: {output_file}")
