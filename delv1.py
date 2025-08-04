import pandas as pd

# Load Excel file and target sheet
file_path = "2024_06_20_Software_Analysis_All.xlsx"  # Replace with your actual file path
df = pd.read_excel(file_path, sheet_name='Software_All')

# Clean 'undefined' and blank strings
df['DisplayName'] = df['DisplayName'].replace("undefined", pd.NA)
df['DisplayName'] = df['DisplayName'].replace(r'^\s*$', pd.NA, regex=True)

# Drop rows with missing DisplayName
df_valid = df.dropna(subset=['DisplayName'])

# ✅ Keep first occurrence of each DisplayName
df_unique_first = df_valid.drop_duplicates(subset=['DisplayName'], keep='first')

# Export to Excel
df_unique_first.to_excel("Software_Unique_Updated_Report.xlsx", index=False)
