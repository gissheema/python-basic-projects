# import win32com.client 

# excel = win32com.client.Dispatch("Excel.Application")
# excel.Visible = True

# workbook = excel.Workbooks.Open(r"C:\Users\sheem\python-project-web\sample.xlsx")
# sheet = workbook.Worksheets("Sheet1")

# # Write data
# sheet.Cells(1, 1).Value = "Name"
# sheet.Cells(1, 2).Value = "Score"

# # Add sample data
# names = ["Arun", "Priya", "Kumar"]

# for i, name in enumerate(names, start=2):
#     sheet.Cells(i, 1).Value = name
#     sheet.Cells(i, 2).Value = i * 10

# # Format header
# sheet.Range("A1:B1").Font.Bold = True

# # Save and close
# workbook.Save()
# workbook.Close()
# excel.Quit()




import win32com.client as win32 #pip install pywin32

# Start Excel
excel = win32.Dispatch("Excel.Application")
excel.Visible = True   # Set False to run in background

try:
    # Open an existing workbook
    file_path = r"C:\\Users\sheem\python-project-web\sample1.xlsx"
    workbook = excel.Workbooks.Open(file_path,ReadOnly=False)

    # Select worksheet
    sheet = workbook.Worksheets("Sheet1")

    # -------------------------------
    # 1. Write Header
    # -------------------------------
    sheet.Cells(1, 1).Value = "Name"
    sheet.Cells(1, 2).Value = "Department"
    sheet.Cells(1, 3).Value = "Salary"

    # -------------------------------
    # 2. Add Data
    # -------------------------------
    employees = [
        ("Arun", "IT", 30000),
        ("Priya", "HR", 28000),
        ("Kumar", "Finance", 35000),
        ("Divya", "IT", 40000)
    ]

    for i, emp in enumerate(employees, start=2):
        sheet.Cells(i, 1).Value = emp[0]
        sheet.Cells(i, 2).Value = emp[1]
        sheet.Cells(i, 3).Value = emp[2]

    # -------------------------------
    # 3. Formatting
    # -------------------------------
    header_range = sheet.Range("A1:C1")
    header_range.Font.Bold = True
    header_range.Interior.ColorIndex = 36  # Light yellow

    # Auto-fit columns
    sheet.Columns("A:C").AutoFit()

    # -------------------------------
    # 4. Insert a New Row
    # -------------------------------
    sheet.Rows(2).Insert()
    sheet.Cells(2, 1).Value = "New Employee"
    sheet.Cells(2, 2).Value = "Admin"
    sheet.Cells(2, 3).Value = 25000

    # -------------------------------
    # 5. Calculate Total Salary
    # -------------------------------
    last_row = sheet.UsedRange.Rows.Count
    sheet.Cells(last_row + 1, 2).Value = "Total"
    sheet.Cells(last_row + 1, 3).Formula = f"=SUM(C2:C{last_row})"

    # -------------------------------
    # 6. Save File
    # -------------------------------
   # workbook.Save()

    # Optional: Save as new file
    workbook.SaveAs(r"C:\\Users\sheem\python-project-web\updated.xlsx")

    print("✅ Excel file updated successfully!")

except Exception as e:
    print("❌ Error:", e)

finally:
    # -------------------------------
    # 7. Close Excel
    # -------------------------------
    workbook.Close()
    excel.Quit()