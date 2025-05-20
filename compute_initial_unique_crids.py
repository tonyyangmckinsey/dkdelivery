def compute_unique_crids():

    import pandas as pd
    import os
    from datetime import datetime

    today_str = datetime.today().strftime("%Y%m%d")  # Format: YYYYMMDD
  
    # === Setup paths ===
    base_path = r"C:\Users\Tony Yang\OneDrive - McKinsey & Company\Documents\Python"
    snapshot_folder = os.path.join(base_path, "3. DKTrackerSnapshots")
    output_file = os.path.join(base_path, "AnalysisOutputFolder", f"UniqueCrids.xlsx")

    # === Columns to extract ===
    cols_to_keep = [
        "ServiceID_Crid",
        "SalesProjectID_ContractNumber",
        "Customer name",
        "Delta MRC",
        "Delivery Project Manager",
        "Segment Red.",
        "ProjectTypeValue"
    ]

    # === Load and collect data from all snapshot files ===
    combined_df = pd.DataFrame()
    for filename in os.listdir(snapshot_folder):
        if filename.endswith(".xlsx"):
            file_path = os.path.join(snapshot_folder, filename)
            print(f"🔍 Reading: {filename}")
            try:
                df = pd.read_excel(file_path, sheet_name=0, engine="openpyxl")
                df_subset = df[cols_to_keep].copy()
                combined_df = pd.concat([combined_df, df_subset], ignore_index=True)
            except Exception as e:
                print(f"⚠️ Could not process {filename}: {e}")

    # === Drop duplicates ===
    print(f"📊 Combined rows before deduplication: {len(combined_df)}")
    unique_df = combined_df.drop_duplicates(subset=["ServiceID_Crid", "SalesProjectID_ContractNumber"])
    print(f"✅ Unique rows after deduplication: {len(unique_df)}")

    # === Export result ===
    unique_df.to_excel(output_file, index=False)
    print(f"📤 Unique crids exported to: {output_file}")

    # === Add date columns from today until end of year ===
    from datetime import timedelta
    start_date = datetime.strptime("20250519", "%Y%m%d")

    end_date = datetime(start_date.year, 12, 31)
    date_range = pd.date_range(start=start_date, end=end_date)

    for date in date_range:
        col_name = date.strftime("%Y%m%d")
        unique_df[col_name] = ""

    # === Export result ===
    unique_df.to_excel(output_file, index=False)
    print(f"📤 Unique crids exported to: {output_file}")

  