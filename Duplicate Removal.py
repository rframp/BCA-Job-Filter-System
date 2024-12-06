import pandas as pd

# Define the file path and date columns to format
file_path = r"C:\Users\RFrampton\OneDrive - Cowan Motor Group\Desktop\BCA Work.xlsx"
date_columns = ['StartDate', 'EndDate', 'AgreedDate']  # Adjust these names if needed

# Load data from Sheet1
data = pd.read_excel(file_path, sheet_name="Sheet1")

# Remove duplicate rows based on 'JobNumber' column, keeping the first occurrence
data = data.drop_duplicates(subset="JobNumber", keep="first")

# Format date columns to 'dd/mm/yyyy'
for column in date_columns:
    if column in data.columns:
        data[column] = pd.to_datetime(data[column], errors='coerce').dt.strftime('%d/%m/%Y')

# Write the processed data to Sheet2
with pd.ExcelWriter(file_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    data.to_excel(writer, sheet_name="Sheet2", index=False)

print("Duplicate removal and date formatting completed. Check 'Sheet2' in the Excel file.")
