
# Imports the pandas library
import pandas as pd

# Load Excel into a dataframe(df)
df = pd.read_excel('Cores-Info-Database.xlsx')

# Normalize column names - removes any whitespace and makes names case-insensitive
df.columns = [str(col).strip().lower() for col in df.columns]

# Get user input for the core name as well as the required dimension
core_type = input("Enter the core type (e.g., E 5): ").strip()
dimension = input("Enter the dimension to retrieve (e.g., Effective magnetic path length: ").strip().lower()

# Filter by core name - traverses through the spreadsheet to find the matching row
# inner df[] obtains a list of core types and matches it to the one desired
# outer df[] returns the row corresponding to the specific row type with all dimension values included
row = df[df['core type'].str.lower() == core_type.lower()]

if row.empty: # checks if the row is empty, meaning that the desired values weren't present in the spreadsheet
    print(f"No core found with name '{core_type}'.") # print statement if the specific core type wasn't found
elif dimension not in df.columns: # checks if the specified dimension is the 
    print(f"Dimension '{dimension}' not found in spreadsheet.") # print statement if the specific dimension isn't located in the spreadsheet
else:
    value = row.iloc[0][dimension] # obtains the dimension value 
    print(f"{dimension.title()} for {core_type.title()}: {value}") # prints the value