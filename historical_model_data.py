from openpyxl import load_workbook

dell_tickets_file = "C:\\Users\\hagand\\Downloads\\Dell_TechDirect_Tickets.xlsx"
asset_tiger_file = "C:\\Users\\hagand\\Downloads\\AssetTiger_AllDevices_Warranty.xlsx"

dell_tickets_workbook = load_workbook(dell_tickets_file)
dell_tickets_sheet = dell_tickets_workbook["Sheet1"]

# Using a special exported sheet - see Ryder for details
asset_tiger_workbook = load_workbook(asset_tiger_file)
asset_tiger_sheet = asset_tiger_workbook["Export"]

searched_model = "OPTIPLEX 7460 AIO"

at_device_ids = []
at_device_warranties = []
for idx, cell in enumerate(asset_tiger_sheet['A']):
  device_st = cell.value
  device_model = asset_tiger_sheet['D' + str(idx + 1)].value
  device_warranty = asset_tiger_sheet['C' + str(idx + 1)].value
  if device_model in searched_model:
    at_device_ids.append(device_st)
    at_device_warranties.append([device_st, device_warranty])

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

#To get the number of machines - ignores multiple tickets.
dt_devices_st_only = []
for device in dt_devices:
  device_st = device[0]
  if device_st not in dt_devices_st_only:
    dt_devices_st_only.append(device_st)




# What we wanted to know
print("Device count: " + str(len(at_device_ids)))
print("Devices that have had Dell tickets: " + str(len(dt_devices_st_only)))
print("Devices that have multiple tickets: " + str(len(dt_devices)-len(dt_devices_st_only)))