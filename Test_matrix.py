import pandas as pd
import os
import win32com.client as win32

# Define the updated test matrix data
data = {
    "Test Case ID": [
        f"TC{str(i).zfill(3)}" for i in range(1, 13)  # Generate TC001 to TC012
    ],
    "Applicant Type": [
        "Single", "BBB", "Couple", "Double", "Single", "BBB", "Couple", "Double", "Single", "BBB", "Couple", "Double"
    ],
    "Vehicle Condition": [
        "NNN", "NNN", "NNN", "NNN", "UUU", "UUU", "UUU", "UUU", "CCC", "CCC", "CCC", "CCC"
    ],
    "Product Type": [
        "MMM", "MMM", "MMM", "MMM", "MMM", "MMM", "MMM", "MMM", "RRR", "RRR", "RRR", "RRR"
    ],
    "State": [
        "ORRR", "TXTX", "ORRR", "TXTX", "ORRR", "TXTX", "ORRR", "TXTX", "ORRR", "TXTX", "ORRR", "TXTX"
    ],
    "Space Amount (Condition)": [
        "$200", "$200", "$200", "$200", "5% of Money Calculated", "5% of Money Calculated", "5% of Money Calculated", "5% of Money Calculated",
        "$200", "$200", "5% of Money Calculated", "5% of Money Calculated"
    ],
    "Rule Name": [
        "Space Amount Over-Advance" if condition == "$200" else "Space Amount Over-Advance (TXTX & ORRR)"
        for condition in [
            "$200", "$200", "$200", "$200", "5% of Money Calculated", "5% of Money Calculated", "5% of Money Calculated", "5% of Money Calculated",
            "$200", "$200", "5% of Money Calculated", "5% of Money Calculated"
        ]
    ],
    "Expected Result": [
        "Rule should not be triggered" if rule == "Space Amount Over-Advance" else "Rule should be triggered"
        for rule in [
            "Space Amount Over-Advance" if condition == "$200" else "Space Amount Over-Advance (TXTX & ORRR)"
            for condition in [
                "$200", "$200", "$200", "$200", "5% of Money Calculated", "5% of Money Calculated", "5% of Money Calculated", "5% of Money Calculated",
                "$200", "$200", "5% of Money Calculated", "5% of Money Calculated"
            ]
        ]
    ]
}

# Create a DataFrame
df = pd.DataFrame(data)

# Define the output file path
output_file = os.path.join(os.getcwd(), "Updated_TestMatrix_SpaceAmountOverAdvance.xlsm")

# Save the DataFrame to an Excel file
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Add()
ws = wb.Worksheets(1)

# Write the DataFrame to the Excel sheet
for i, col in enumerate(df.columns):
    ws.Cells(1, i + 1).Value = col  # Write headers
    for j, value in enumerate(df[col]):
        ws.Cells(j + 2, i + 1).Value = value  # Write data

# Save the file as a macros-enabled Excel file
wb.SaveAs(output_file, FileFormat=52)  # Save as .xlsm
wb.Close()
excel.Application.Quit()

print(f"Updated test matrix has been created and saved as {output_file}.")