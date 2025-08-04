import pandas as pd

# Load Excel file and sheet
file_path = "2024_06_20_Software_Analysis_All.xlsx"  # Update path as needed
df = pd.read_excel(file_path, sheet_name='Software_All')

# Clean 'undefined' and empty strings
df_cleaned = df.replace("undefined", pd.NA)
df_cleaned['DisplayName'] = df_cleaned['DisplayName'].replace(r'^\s*$', pd.NA, regex=True)

# Drop rows where DisplayName is NaN
df_valid = df_cleaned.dropna(subset=['DisplayName'])

# Identify duplicated DisplayNames (keep all duplicates)
duplicates_mask = df_valid.duplicated(subset=['DisplayName'], keep=False)
df_duplicates_all = df_valid[duplicates_mask]

# Keep only the first entry for each duplicate group
df_duplicates_first = df_duplicates_all.drop_duplicates(subset=['DisplayName'], keep='first')

# Export to Excel
df_duplicates_first.to_excel("Software_Duplicates_Report.xlsx", index=False)
