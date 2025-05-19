import pandas as pd

# === Load master and new data files ===
master_path = r"C:\Users\Tony Yang\OneDrive - McKinsey & Company\Documents\Python input\DK backlog list\20250430 DK order backlog as of 20250430 TEST 1.xlsx"
new_path = r"C:\Users\Tony Yang\OneDrive - McKinsey & Company\Documents\Python input\DK backlog list\20250430 DK order backlog as of 20250430 TEST 2.xlsx"

master_df = pd.read_excel(master_path)
new_df = pd.read_excel(new_path)

# === Clean column names: remove trailing/leading spaces ===
master_df.columns = master_df.columns.str.strip()
new_df.columns = new_df.columns.str.strip()

# === Normalize ID columns ===
for df in [master_df, new_df]:
    df["ServiceID_Crid"] = df["ServiceID_Crid"].astype(str).str.strip()
    df["SalesProjectID_ContractNumber"] = df["SalesProjectID_ContractNumber"].astype(str).str.strip()

# === Align columns ===
new_df = new_df[master_df.columns]  # enforce same column order and content

# === Identify NEW rows based on composite key ===
merged_df = pd.merge(
    new_df,
    master_df,
    on=["ServiceID_Crid", "SalesProjectID_ContractNumber"],
    how="left",
    indicator=True
)

# 🛠️ Instead of taking rows from merged_df, take clean ones from new_df using the same index
new_orders = new_df.loc[merged_df["_merge"] == "left_only"]
new_orders = new_orders[master_df.columns]  # drop the _merge column and match column structure

# === Append and export ===
updated_master = pd.concat([master_df, new_orders], ignore_index=True)

#output_path = r"C:\Users\Tony Yang\OneDrive - McKinsey & Company\Documents\Python input\DK backlog list\Merged and updated file.xlsx"
output_path = r"C:\Users\Tony Yang\OneDrive - McKinsey & Company\Documents\Python input\DK backlog list\output_test.xlsx"
updated_master.to_excel(output_path, index=False)

print("✅ Master file exported successfully.")
print("🆕 New orders added:", len(new_orders))
