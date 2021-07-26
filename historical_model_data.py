from openpyxl import load_workbook

dell_tickets_file = "C:\\Users\\hagand\\Downloads\\Dell_TechDirect_Tickets.xlsx"
asset_tiger_file = "C:\\Users\\hagand\\Downloads\\AssetTiger_7460AIO.xlsx"

dell_tickets_workbook = load_workbook(dell_tickets_file)
dell_tickets_sheet = dell_tickets_workbook["Sheet1"]

asset_tiger_workbook = load_workbook(asset_tiger_file)
asset_tiger_sheet = asset_tiger_workbook["Asset"]

test_dell_cell = dell_tickets_sheet['A1'].value
print(test_dell_cell)

at_device_id = []
for cell in asset_tiger_sheet['B']:
  at_device_id.append(cell.value)

# Remove header from list of service tags
at_device_id.pop(0)

print(at_device_id[0])
print(at_device_id[1])