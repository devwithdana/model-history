from openpyxl import load_workbook
from openpyxl import Workbook

from objects import Device
from objects import Ticket

# Using an excel sheet with special exporting - see Ryder.
# Not all devices of the given model may be found in Asset Tiger

dell_tickets_file = "C:\\Users\\hagand\\Downloads\\Dell_TechDirect_Tickets.xlsx"
asset_tiger_file = "C:\\Users\\hagand\\Downloads\\AssetTiger_AllDevices_Warranty.xlsx"

dell_tickets_workbook = load_workbook(dell_tickets_file)
dell_tickets_sheet = dell_tickets_workbook["Sheet1"]

asset_tiger_workbook = load_workbook(asset_tiger_file)
asset_tiger_sheet = asset_tiger_workbook["Export"]

search_tag = ""
search_model = "OPTIPLEX 7460 AIO"
search_purchase_date = ""
search_warranty = ""


# Read in Asset Tiger
devices_dict = dict()
for idx, cell in enumerate(asset_tiger_sheet['A']):
  if idx != 0:
    st = cell.value
    model = asset_tiger_sheet['D' + str(idx + 1)].value
    warranty = asset_tiger_sheet['C' + str(idx + 1)].value
    devices_dict[st] = Device(st, model, None, warranty)


# Read in Dell Tickets
for idx, cell in enumerate(dell_tickets_sheet['A']):
  if idx != 0:
    work_order = cell.value
    device_tag = dell_tickets_sheet['C' + str(idx)].value
    status = dell_tickets_sheet['B' + str(idx)].value
    problem = dell_tickets_sheet['E' + str(idx)].value
    date_created = dell_tickets_sheet['F' + str(idx)].value
    ticket = Ticket(work_order, device_tag, status, problem, date_created)

    device = devices_dict.get(device_tag)
    if device != None: 
      device.tickets.append(ticket)

filtered_devices = devices_dict
filtered_devices = {k:v for (k, v) in filtered_devices.items() if search_tag in v.model}
filtered_devices = {k:v for (k, v) in filtered_devices.items() if search_model in v.model}
filtered_devices = {k:v for (k, v) in filtered_devices.items() if search_purchase_date in v.model}
filtered_devices = {k:v for (k, v) in filtered_devices.items() if search_warranty in v.model}

# What we wanted to know
print("Model: " + search_model)
print("Device count: " + str(len(filtered_devices)))

in_warranty = {k:v for (k, v) in filtered_devices.items() if v.hasExpired()}
print("Devices out of warranty: " + str(len(in_warranty)))

has_tickets = {k:v for (k, v) in filtered_devices.items() if len(v.tickets) > 0}
print("Devices that have had Dell tickets: " + str(len(has_tickets)))

have_multiple_tickets = {k:v for (k, v) in has_tickets.items() if len(v.tickets) > 1}
print("Devices that have multiple tickets: " + str(len(have_multiple_tickets)))

mobo_issues = {k:v for (k, v) in has_tickets.items() if v.tickets[0].category == "MOBO"}
print("Devices that had motherboard issues: " + str(len(mobo_issues)))
hdd_issues = {k:v for (k, v) in has_tickets.items() if v.tickets[0].category == "HDD"}
print("Devices that had hard drive issues: " + str(len(hdd_issues)))
lcd_issues = {k:v for (k, v) in has_tickets.items() if v.tickets[0].category == "LCD"}
print("Devices that had LCD issues: " + str(len(lcd_issues)))
psu_issues = {k:v for (k, v) in has_tickets.items() if v.tickets[0].category == "PSU"}
print("Devices that had power supply issues: " + str(len(psu_issues)))
ram_issues = {k:v for (k, v) in has_tickets.items() if v.tickets[0].category == "RAM"}
print("Devices that had RAM issues: " + str(len(ram_issues)))
other_issues = {k:v for (k, v) in has_tickets.items() if v.tickets[0].category == "OTHER"}
print("Devices that had other issues: " + str(len(other_issues)))
for device in other_issues:
  tickets = other_issues[device].tickets
  for ticket in tickets:
    print(ticket.problem)