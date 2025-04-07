import pandas as pd
import os
import win32com.client as win32

# Step 1: Load the test matrix data from Updated_TestMatrix_SpaceAmountOverAdvance.xlsx
matrix_file = os.path.join(os.getcwd(), "Updated_TestMatrix_SpaceAmountOverAdvance.xlsx")  # Replace with the actual file name
df = pd.read_excel(matrix_file)

# Step 2: Define the design template file paths
template_200 = os.path.join(os.getcwd(), "CR 59513_Design Template.xlsx")  # Template for Space Amount = 200
template_5_percent = os.path.join(os.getcwd(), "CR 59513_Design Template_Resubmit.xlsx")  # Template for Space Amount = 5% of Money Calculated

# Step 3: Generate test case files for each row in the test matrix
excel = win32.gencache.EnsureDispatch('Excel.Application')

for _, row in df.iterrows():
    test_case_id = row["Test Case ID"]
    output_file = os.path.join(os.getcwd(), f"{test_case_id}_TestDesign.xlsm")  # Ensure unique file name for each test case ID

    # Select the appropriate design template based on Space Amount (Condition)
    if row['Space Amount (Condition)'] == "$200":
        template_file = template_200
    elif row['Space Amount (Condition)'] == "5% of Money Calculated":
        template_file = template_5_percent
    else:
        # Default to the first template if the condition is unknown
        print(f"Unknown Space Amount (Condition) for Test Case ID: {test_case_id}. Using default template.")
        template_file = template_200

    # Open the selected design template file
    wb = excel.Workbooks.Open(template_file)
    ws = wb.Worksheets(1)

    # Update the placeholders in the design template dynamically
    # Row 7 (B7): Update Product, Vehicle Condition, and Applicant Type
    ws.Cells(7, 2).Value = f"Product: {row['Product Type']}, Vehicle condition: {row['Vehicle Condition']}, Applicant type: {row['Applicant Type']}"

    # Row 8 (D8): Update Applicant Type
    ws.Cells(8, 4).Value = f"Applicant type: {row['Applicant Type']}"

    # Row 9 (D9): Update Applicant Type
    ws.Cells(9, 4).Value = f"Applicant type: {row['Applicant Type']}"

    # Row 10 (D10): Update Product and Vehicle Condition
    ws.Cells(10, 4).Value = f"Product: {row['Product Type']}, Vehicle condition: {row['Vehicle Condition']}"

    # Row 13 (D13): Update Space Amount (Condition)
    ws.Cells(13, 4).Value = f"Space Amount: {row['Space Amount (Condition)']}"

    # Row 14 (D14): Update Rule Name
    ws.Cells(14, 4).Value = f"Rule Name: {row['Rule Name']}"

    # Row 14 (E14): Update Expected Result
    ws.Cells(14, 5).Value = f"Expected Result: {row['Expected Result']}"

    # Save the populated file as .xlsm
    wb.SaveAs(output_file, FileFormat=52)  # Save as .xlsm
    wb.Close()

excel.Application.Quit()

print("Test case files for all rows in the test matrix have been created successfully based on the respective templates.")