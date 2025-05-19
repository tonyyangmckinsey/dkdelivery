import pandas as pd
from datetime import datetime
today_str = datetime.today().strftime("%Y%m%d")  # Format: YYYYMMDD

# === Load master and new data files ===
#Change name of master path
master_path = r"C:\Users\Tony Yang\OneDrive - McKinsey & Company\Documents\DK backlog python\2. Update existing backlog (add latest backlog here)\DK backlog tracker.xlsx"
new_path = rf"C:\Users\Tony Yang\OneDrive - McKinsey & Company\Documents\DK backlog python\{today_str}_1. Final__latest_backlog_output (remember to update the first columns).xlsx"

master_df = pd.read_excel(master_path)
new_df = pd.read_excel(new_path)

# === Clean column names ===
master_df.columns = master_df.columns.str.strip()
new_df.columns = new_df.columns.str.strip()

# === Normalize keys ===
for df in [master_df, new_df]:
    df["ServiceID_Crid"] = df["ServiceID_Crid"].astype(str).str.strip()
    df["SalesProjectID_ContractNumber"] = df["SalesProjectID_ContractNumber"].astype(str).str.strip()

# === Align columns ===
new_df = new_df[master_df.columns]  # ensure consistent column structure

# === Identify new rows ===
merged_df = pd.merge(
    new_df,
    master_df,
    on=["ServiceID_Crid", "SalesProjectID_ContractNumber"],
    how="left",
    indicator=True
)
new_orders = new_df.loc[merged_df["_merge"] == "left_only"]
new_orders = new_orders[master_df.columns]

# === Append new rows ===
updated_master = pd.concat([master_df, new_orders], ignore_index=True)

# === Fields to update if new values exist ===
fields_to_update = [
    "Committed RFS", "Customer AgreedRFS/ConfirmedDeliveryDate", "Actual RFS",
    "Planning complete date", "Design complete date", "Correct RFS",
    "Expected RFS", "Deliveryconfirmationdate_Configsignoffdate", 
    "Customer Agreed RFS", "CUSTOMERCONFIRMEDDELIVERYDATE", 
    "CUSTOMERDEFERREDDELIVERYDATE"
]

# === Merge update fields ===
update_fields = ["ServiceID_Crid", "SalesProjectID_ContractNumber"] + fields_to_update
update_df = new_df[update_fields].copy()
update_df = update_df.rename(columns={col: f"{col}_new" for col in fields_to_update})

updated_master = pd.merge(
    updated_master,
    update_df,
    on=["ServiceID_Crid", "SalesProjectID_ContractNumber"],
    how="left"
)

# === Apply updates and track how many were changed ===
update_counts = {}
for col in fields_to_update:
    new_col = f"{col}_new"
    original = updated_master[col]
    updated_master[col] = updated_master[new_col].combine_first(original)
    updated = (original != updated_master[col]) & updated_master[new_col].notna()
    update_counts[col] = updated.sum()

# === Drop helper columns
updated_master.drop(columns=[f"{col}_new" for col in fields_to_update], inplace=True)

# === Export final file
output_path = rf"C:\Users\Tony Yang\OneDrive - McKinsey & Company\Documents\DK backlog python\{today_str}_2. Merged_and_appended_backlog_DKBACKLOGTRACKER.xlsx"
updated_master.to_excel(output_path, index=False)

# === Print summary
print("✅ Master file updated and exported.")
for field, count in update_counts.items():
    print(f"🔁 {field}: {count} updates applied")
