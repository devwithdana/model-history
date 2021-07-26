from openpyxl import load_workbook

# Using an excel sheet with special exporting - see Ryder.
# Not all devices of the given model may be found in Asset Tiger

dell_tickets_file = "C:\\Users\\hagand\\Downloads\\Dell_TechDirect_Tickets.xlsx"
asset_tiger_file = "C:\\Users\\hagand\\Downloads\\AssetTiger_AllDevices_Warranty.xlsx"

dell_tickets_workbook = load_workbook(dell_tickets_file)
dell_tickets_sheet = dell_tickets_workbook["Sheet1"]

asset_tiger_workbook = load_workbook(asset_tiger_file)
asset_tiger_sheet = asset_tiger_workbook["Export"]

searched_model = "OPTIPLEX 7460 AIO"

at_devices = dict()
for idx, cell in enumerate(asset_tiger_sheet['A']):
  device_st = cell.value
  device_model = asset_tiger_sheet['D' + str(idx + 1)].value
  device_warranty = asset_tiger_sheet['C' + str(idx + 1)].value
  if device_model in searched_model:
    at_devices[device_st] = device_warranty

dt_devices = dict()
ticket_count = 0
for idx, cell, in enumerate(dell_tickets_sheet['C']):
  if idx != 0:
    device_st = cell.value
    if device_st in list(at_devices):
      ticket_count = ticket_count + 1
      if device_st in list(dt_devices):
        ticket_list = dt_devices[device_st]
        ticket_list.append(idx + 1)
        dt_devices[device_st] = ticket_list
      else:
        dt_devices[device_st] = [idx + 1]


# What we wanted to know
print("Device count: " + str(len(at_devices)))
print("Devices that have had Dell tickets: " + str(len(dt_devices)))
print("Devices that have multiple tickets: " + str(ticket_count -len(dt_devices)))