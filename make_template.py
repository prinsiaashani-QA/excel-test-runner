"""Creates input/test_cases.xlsx with 3 sample test cases (2 should pass, 1 is a demo fail)."""
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from core.rules import TC_COLUMNS, PRIORITIES

wb = Workbook()
ws = wb.active
ws.title = "TestCases"

ws["A1"] = "Login Module - Test Cases"
ws["A1"].font = Font(bold=True, size=14)

ws.append(TC_COLUMNS)
for c in ws[2]:
    c.font = Font(bold=True, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor="305496")

LOGIN = ("1. open\n"
         "2. type | placeholder=Username | {{USERNAME}}\n"
         "3. type | placeholder=Password | {{PASSWORD}}\n"
         "4. click | role=button:Login")

ws.append(["TC-001", "Login", "Valid login", "User is on login page",
           LOGIN + "\n5. verify_url | inventory.html\n6. verify_text | css=.title | Products",
           "standard_user / secret_sauce", "Products page opens", "", "High", ""])

ws.append(["TC-002", "Login", "Login with empty username", "User is on login page",
           "1. open\n2. type | placeholder=Password | {{PASSWORD}}\n3. click | role=button:Login\n"
           "4. verify_contains | css=[data-test='error'] | Username is required",
           "secret_sauce", "Error: Username is required", "", "Medium", ""])

ws.append(["TC-003", "Login", "DEMO: this must fail (wrong expected text)", "User is on login page",
           LOGIN + "\n5. verify_text | css=.title | Product Lists",
           "standard_user / secret_sauce", "Title is 'Product Lists'", "", "Low", ""])

for i, w in enumerate([10, 12, 36, 24, 52, 28, 32, 36, 10, 10], 1):
    ws.column_dimensions[chr(64 + i)].width = w
for row in ws.iter_rows(min_row=3):
    for cell in row:
        cell.alignment = Alignment(wrap_text=True, vertical="top")
ws.freeze_panes = "A3"

dv = DataValidation(type="list", formula1='"' + ",".join(PRIORITIES) + '"')
ws.add_data_validation(dv)
dv.add("I3:I500")

Path("input").mkdir(exist_ok=True)
wb.save("input/test_cases.xlsx")
print("Template ready: input/test_cases.xlsx")
