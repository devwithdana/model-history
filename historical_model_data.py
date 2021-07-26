from openpyxl import load_workbook

dell_tickets_file = "C:\\Users\\hagand\\Downloads\\Dell_TechDirect_Tickets.xlsx"
asset_tiger_file = "C:\\Users\\hagand\\Downloads\\AssetTiger_7460AIO.xlsx"

dell_tickets_workbook = load_workbook(dell_tickets_file)
dell_tickets_sheet = dell_tickets_workbook["Sheet1"]

asset_tiger_workbook = load_workbook(asset_tiger_file)
asset_tiger_sheet = asset_tiger_workbook["Asset"]

test_dell_cell = dell_tickets_sheet['A1'].value
print(test_dell_cell)

at_device_ids = []
for cell in asset_tiger_sheet['B']:
  at_device_ids.append(cell.value)

# Remove header from list of service tags
at_device_ids.pop(0)

print(at_device_ids[0])
print(at_device_ids[1])

dt_all_device_ids = []
for cell in dell_tickets_sheet['C']:
  dt_all_device_ids.append(cell.value)

# Remove header from list of service tags
dt_all_device_ids.pop(0)

dt_devices = []
for idx, st_id in enumerate(dt_all_device_ids):
  device = st_id
  if device in at_device_ids:
    dt_devices.append([device, idx])

# print("This is a list of devices of the given model that had Dell Tickets were submitted:")
# for device in dt_devices:
#   print(device[0])

# WARNING: Some devices from dt_all_device_ids are the correct model we need but were not available in Asset Tiger
# passed = "29D4XQ2" in at_device_ids
# print(passed)