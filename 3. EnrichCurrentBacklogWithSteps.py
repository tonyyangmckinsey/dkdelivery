import pandas as pd
from datetime import datetime
today_str = datetime.today().strftime("%Y%m%d")  # Format: YYYYMMDD

# === Load the final output (master) and the DK backlog tracker (new data) ===
final_output_path = rf"C:\Users\Tony Yang\OneDrive - McKinsey & Company\Documents\Python\2. Output\{today_str}_1. Final__latest_backlog_output (remember to update the first columns).xlsx"
tracker_path = r"C:\Users\Tony Yang\OneDrive - McKinsey & Company\Documents\Python\1. Input\DK backlog tracker.xlsx"

final_df = pd.read_excel(final_output_path)
tracker_df = pd.read_excel(tracker_path)

# === Clean and normalize ID columns ===
final_df.columns = final_df.columns.str.strip()
tracker_df.columns = tracker_df.columns.str.strip()

final_df["ServiceID_Crid"] = final_df["ServiceID_Crid"].astype(str).str.strip()
final_df["SalesProjectID_ContractNumber"] = final_df["SalesProjectID_ContractNumber"].astype(str).str.strip()
tracker_df["ServiceID_Crid"] = tracker_df["ServiceID_Crid"].astype(str).str.strip()
tracker_df["SalesProjectID_ContractNumber"] = tracker_df["SalesProjectID_ContractNumber"].astype(str).str.strip()

# ✅ Drop duplicates in tracker to prevent row expansion during merge
tracker_df = tracker_df.drop_duplicates(subset=["ServiceID_Crid", "SalesProjectID_ContractNumber"])

# === Define columns to enrich ===
enrich_cols = [
    "Input: Which delivery step is the order in? (Drop down list)",
    "Input: What is the status of the order? (Drop down list)",
    "(Optional) Input: Comment"
]

# === Merge tracker data into final output based on ID columns ===
final_enriched = final_df.merge(
    tracker_df[["ServiceID_Crid", "SalesProjectID_ContractNumber"] + enrich_cols],
    on=["ServiceID_Crid", "SalesProjectID_ContractNumber"],
    how="left",
    suffixes=("", "_tracker")
)

# === Overwrite/enrich columns if values from tracker exist ===
for col in enrich_cols:
    enriched_col = col + "_tracker"
    final_enriched[col] = final_enriched[enriched_col].combine_first(final_enriched[col])
    final_enriched.drop(columns=[enriched_col], inplace=True)

# === Save the enriched final output ===
output_path = rf"C:\Users\Tony Yang\OneDrive - McKinsey & Company\Documents\Python\2. Output\{today_str}_3. Final_enriched_latest_backlog.xlsx"
final_enriched.to_excel(output_path, index=False)

print("✅ Enrichment complete.")
print(f"📄 Output saved to: {output_path}")
