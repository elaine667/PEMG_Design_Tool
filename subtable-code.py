
# import statements
import pandas as pd

# obtains the file name
file_path = 'PandasSubtablePractice.xlsx'  # your Excel file path

# exception handler
try:
    df = pd.read_excel(file_path) # reads the excel spreadsheet
except FileNotFoundError:
    print(f"Error: File '{file_path}' not found.") # prints an erro statement
    exit()

df.columns = [col.strip() for col in df.columns]  # Clean column names - rids of any whitespace, case issues, etc

# Ask user for core type
core_type = input("Enter the magnetic core type (e.g., Toroid, E-Core): ").strip()

# Filter DataFrame for matching core type (case insensitive)
filtered_rows = df[df['Core Type'].str.lower() == core_type.lower()] # traverses through the sheet to find the matching core type

# Print result
if filtered_rows.empty:
    print(f"No cores found for type '{core_type}'.") # prints if no such core was found
else:
    print(f"\nAll cores of type '{core_type}':\n")
    print(filtered_rows.to_string(index=False)) # prints the category name as well as the actual value