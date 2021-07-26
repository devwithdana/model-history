from openpyxl import load_workbook

dell_tickets_file = "C:\\Users\\hagand\\Downloads\\Dell_TechDirect_Tickets.xlsx"
asset_tiger_file = "C:\\Users\\hagand\\Downloads\\AssetTiger_7460AIO.xlsx"

dell_tickets_sheet = load_workbook(dell_tickets_file)
dell_tickets = dell_tickets_sheet["Sheet1"]
test_dell_cell = dell_tickets['A1'].value

print(test_dell_cell)