import openpyxl

try:
    # Create a new workbook and save it
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = 'Attendance'  # Optional: Set sheet title
    sheet.append(["Name", "Date Time"])  # Adding headers
    wb.save('Attendance.xlsx')
    print("Attendance.xlsx file created successfully.")
except Exception as e:
    print(f"An error occurred: {e}")
