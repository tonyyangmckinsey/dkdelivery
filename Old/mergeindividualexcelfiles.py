import os
import pandas as pd

# Define paths
input_folder = r"C:\Users\Tony Yang\OneDrive - McKinsey & Company\Documents\Python input\DK backlog files"
output_path = r"C:\Users\Tony Yang\OneDrive - McKinsey & Company\Documents\Python\DK backlog output\merged_output.xlsx"

# Get list of Excel files
excel_files = [os.path.join(input_folder, f) for f in os.listdir(input_folder) if f.endswith(".xlsx")]

# Read the first file fully
df_base = pd.read_excel(excel_files[0], engine='openpyxl')

# Identify rows and columns to merge (row 3 onward, columns C to F → index 2 to 5)
start_row = 2
cols_to_merge = df_base.columns[2:6]  # Columns C to F

# Loop through the rest of the files and merge values into df_base
for file in excel_files[1:]:
    df_temp = pd.read_excel(file, engine='openpyxl')
    for col in cols_to_merge:
        df_base.loc[start_row:, col] = df_temp.loc[start_row:, col].combine_first(df_base.loc[start_row:, col])

# Export merged result
df_base.to_excel(output_path, index=False)
print("✅ Merged file exported successfully to:", output_path)
