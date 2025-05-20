import pandas as pd
import os
from datetime import datetime

def compute_initial_solution():
    # === Setup paths ===
    base_path = r"C:\Users\Tony Yang\OneDrive - McKinsey & Company\Documents\Python"
    snapshot_folder = os.path.join(base_path, "3. DKTrackerSnapshots")
    master_file_path = os.path.join(base_path, "AnalysisOutputFolder", "UniqueCrids.xlsx")
    output_file_path = os.path.join(base_path, "AnalysisOutputFolder", "Populated_UniqueCrids.xlsx")

    # === Load master UniqueCrids file ===
    print("📥 Loading master UniqueCrids file...")
    master_df = pd.read_excel(master_file_path, engine="openpyxl")
    
    # === Identify all date columns ===
    date_columns = [col for col in master_df.columns if col.isdigit() and len(col) == 8]

    # === Normalize matching keys ===
    master_df["ServiceID_Crid"] = master_df["ServiceID_Crid"].astype(str).str.strip()
    master_df["SalesProjectID_ContractNumber"] = master_df["SalesProjectID_ContractNumber"].astype(str).str.strip()

    for date_str in date_columns:
        snapshot_filename = f"{date_str}_DK_backlog_snapshot.xlsx"
        snapshot_path = os.path.join(snapshot_folder, snapshot_filename)

        print(f"🔍 Processing snapshot: {snapshot_filename}")

        if not os.path.exists(snapshot_path):
            print(f"⚠️ Snapshot file not found: {snapshot_filename} — skipping")
            continue

        try:
            snapshot_df = pd.read_excel(snapshot_path, sheet_name=0, engine="openpyxl")
        except Exception as e:
            print(f"⚠️ Failed to read {snapshot_filename}: {e}")
            continue

        # Normalize keys in snapshot
        snapshot_df["ServiceID_Crid"] = snapshot_df["ServiceID_Crid"].astype(str).str.strip()
        snapshot_df["SalesProjectID_ContractNumber"] = snapshot_df["SalesProjectID_ContractNumber"].astype(str).str.strip()

        # Create a lookup dictionary {(ServiceID_Crid, SalesProjectID_ContractNumber): DeliveryStep}
        delivery_step_col = "Input: Which delivery step is the order in? (Drop down list)"
        snapshot_lookup = snapshot_df.set_index(["ServiceID_Crid", "SalesProjectID_ContractNumber"])[delivery_step_col].to_dict()

        # Fill in values in master_df
        filled_count = 0
        for idx, row in master_df.iterrows():
            key = (str(row["ServiceID_Crid"]).strip(), str(row["SalesProjectID_ContractNumber"]).strip())
            if key in snapshot_lookup:
                master_df.at[idx, date_str] = snapshot_lookup[key]
                filled_count += 1

        print(f"✅ {filled_count} rows updated for {date_str}")

    # === Save the updated file ===
    master_df.to_excel(output_file_path, index=False)
    print(f"📤 Final populated file saved to: {output_file_path}")
