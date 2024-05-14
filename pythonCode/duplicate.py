import openpyxl
import pandas as pd

# Load Excel files using openpyxl
task1_wb = openpyxl.load_workbook("task1.xlsx")

# Access sheets
comp_sheet = task1_wb.active


# Read data into lists
task_data = []
for row in comp_sheet.iter_rows(values_only=True):
    task_data.append(row)


# Convert lists to pandas DataFrames
task_df = pd.DataFrame(task_data[1:], columns=task_data[0])

# Filter Unique and Duplicate Drivers
unique_drivers = task_df.drop_duplicates(subset=["DriverName"], keep=False)
duplicate_drivers = task_df[task_df.duplicated(subset=["DriverName"], keep=False)]

# Save Results

unique_drivers.to_excel('unique_drivers.xlsx', index=False)
duplicate_drivers.to_excel('duplicate_drivers.xlsx', index=False)

keep_one = duplicate_drivers.drop_duplicates(subset=["DriverName"], keep="first")


# put it in excell
keep_one.to_excel('keep_one.xlsx', index=False)
