import pandas as pd

# 1. Load the raw Indian sales Excel file safely using a raw string path
df = pd.read_excel(r"C:\Users\divay\Downloads\raw_indian_sales.xlsx")

# 2. Fix missing values in the Amount column using the median
df['Amount (INR)'] = df['Amount (INR)'].fillna(df['Amount (INR)'].median())

# 3. Clean and standardise text formatting
df['Customer Name'] = df['Customer Name'].fillna("Unknown").str.strip().str.title()
df['State'] = df['State'].fillna("Unknown").str.strip().str.title()
df['Category'] = df['Category'].fillna("Unknown").str.strip().str.capitalize()

# 4. FIX: Standardise mixed formats safely WITHOUT losing rows
# format='mixed' allows Python to parse both DD-MM-YYYY and YYYY/MM/DD automatically
df['Date'] = pd.to_datetime(df['Date'], format='mixed', errors='coerce')

# If any dates are still unreadable, fill them with a default date instead of deleting the row
df['Date'] = df['Date'].fillna(pd.to_datetime('2026-01-01'))

# 5. FIX: Ensure missing Order IDs are given a temporary ID instead of dropping the row
df['Order ID'] = df['Order ID'].fillna("IN-2026-UNKNOWN")

# 6. Save the final output as a clean Excel file for Power BI
df.to_excel(r"C:\Users\divay\Downloads\clean_indian_sales.xlsx", index=False)
print("Data successfully cleaned! All rows preserved.")

