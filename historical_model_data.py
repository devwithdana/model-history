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
problem_categories = {
  "MOBO": ["motherboard"], 
  "HDD": ["hard drive", "hdd", "solid state drive", "ssd"], 
  "LCD": ["lcd", "display", "screen"], 
  "PSU": ["power"], 
  "RAM": ["memory"],
  "OTHER": []}
tickets_by_category = {
  "MOBO": [], 
  "HDD": [], 
  "LCD": [], 
  "PSU": [], 
  "RAM": [], 
  "OTHER": []
}

# Read in Asset Tiger
at_devices = dict()
in_warranty_count = 0
for idx, cell in enumerate(asset_tiger_sheet['A']):
  device_st = cell.value
  device_model = asset_tiger_sheet['D' + str(idx + 1)].value
  device_warranty = asset_tiger_sheet['C' + str(idx + 1)].value
  device_in_warranty = asset_tiger_sheet['E' + str(idx + 1)].value
  if device_model in searched_model:
    at_devices[device_st] = device_warranty
    if device_in_warranty == "Yes":
      in_warranty_count = in_warranty_count + 1

# Read in Dell Tickets
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

# Review tickets
for device in dt_devices:
  row = dt_devices[device]
  for ticket in row:
    ticket_description = dell_tickets_sheet['E' + str(ticket)].value
    ticket_problem = ""
    for category in problem_categories:
      key_words = problem_categories[category]
      for word in key_words:
        if word.lower() in ticket_description.lower():
          if ticket_problem == "":
            ticket_problem = category
          else:
            print("WARNING: " + device + " has triggered multiple description categories.")
            print(ticket_problem + " and " + category)
            print(ticket_description)
            ticket_problem = "OTHER"
    if ticket_problem == "":
      ticket_problem = "OTHER"
    tickets_by_category[ticket_problem].append(device)

# What we wanted to know

print("\n")

print("Model: " + searched_model)
print("Device count: " + str(len(at_devices)))
print("Devices in warranty: " + str(in_warranty_count))
print("Devices that have had Dell tickets: " + str(len(dt_devices)))
print("Devices that have multiple tickets: " + str(ticket_count -len(dt_devices)))

print("\n")

print("Devices that had motherboard issues: " + str(len(tickets_by_category["MOBO"])))
print("Devices that had hard drive issues: " + str(len(tickets_by_category["HDD"])))
print("Devices that had LCD issues: " + str(len(tickets_by_category["LCD"])))
print("Devices that had power supply issues: " + str(len(tickets_by_category["PSU"])))
print("Devices that had RAM issues: " + str(len(tickets_by_category["RAM"])))

print("\n")


print("Devices that had other issues: " + str(len(tickets_by_category["OTHER"])))
for device in tickets_by_category["OTHER"]:
  ticket_rows = dt_devices[device]
  print(dell_tickets_sheet['E' + str(ticket_rows[0])].value)