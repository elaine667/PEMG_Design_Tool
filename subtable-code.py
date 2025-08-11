
# import statements
import pandas as pd

# obtains the file name
file_path = 'Cores-Info-Database.xlsx'  # your Excel file path

# exception handler
try:
    df = pd.read_excel(file_path) # reads the excel spreadsheet
except FileNotFoundError:
    print(f"Error: File '{file_path}' not found.") # prints an erro statement
    exit()

df.columns = [col.strip() for col in df.columns]  # Clean column names - rids of any whitespace, case issues, etc

# Ask user for core type
core_type = input("Enter the magnetic core type (e.g., E 5, RM 6): ").strip()

# Filter DataFrame for matching core type (case insensitive)
filtered_rows = df[df['Core Type'].str.lower() == core_type.lower()] # traverses through the sheet to find the matching core type

# Print result
if filtered_rows.empty:
    print(f"No cores found for type '{core_type}'.") # prints if no such core was found
else:
    print(f"\nCores of type '{core_type}':\n")
    for index, row in filtered_rows.iterrows():
        for col in df.columns:
            print(f"{col}: {row[col]}")
        print()  # blank line between cores