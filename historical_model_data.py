from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.styles import Font

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

def format_date(raw_date):
  if raw_date is None:
    return raw_date
  proper_date = str(raw_date).split()[0].split("-")
  proper_date = proper_date[1] + "-" + proper_date[2] + "-" + proper_date[0]
  return proper_date

def format_date_yrmo(raw_date):
  if raw_date is None:
    return raw_date
  proper_date = str(raw_date).split()[0].split("-")
  proper_date = proper_date[0] + "-" + proper_date[1]
  return proper_date

wb = Workbook()
dest_filename = 'asset_list.xlsx'

def create_overview_assessment():
  ws0 = wb.active
  ws0.title = "Overview"

  # ======= Search Criteria
  ws0['A1'] = "Search Criteria"
  ws0['A2'] = "Service Tag"
  if search_tag == "":
    ws0['B2'] = "N/A"
  else:
    ws0['B2'] = search_tag

  ws0['A3'] = "Model"
  if search_model == "":
    ws0['B3'] = "N/A"
  else:
    ws0['B3'] = search_model

  ws0['A4'] = "Purchase Date"
  if search_purchase_date == "":
    ws0['B4'] = "N/A"
  else:
    ws0['B4'] = search_purchase_date

  ws0['A5'] = "Warranty Date"
  if search_warranty == "":
    ws0['B5'] = "N/A"
  else:
    ws0['B5'] = search_warranty

  # ======= Result Info
  ws0['A7'] = "Device Count"
  ws0['B7'] = len(filtered_devices)

  expired = {k:v for (k, v) in filtered_devices.items() if v.hasExpired()}
  ws0['A8'] = "Out of Warranty"
  ws0['B8'] = len(expired)

  has_tickets = {k:v for (k, v) in filtered_devices.items() if len(v.tickets) > 0}
  ws0['A9'] = "Have Had Tickets"
  ws0['B9'] = len(has_tickets)

  have_multiple_tickets = {k:v for (k, v) in has_tickets.items() if len(v.tickets) > 1}
  ws0['A10'] = "Have Had Multiple Tickets"
  ws0['B10'] = len(have_multiple_tickets)

  ws0['A12'] = "Ticket Categories"

  mobo_issues = {k:v for (k, v) in has_tickets.items() if v.tickets[0].category == "MOBO"}
  ws0['A13'] = "Motherboard"
  ws0['B13'] = len(mobo_issues)

  hdd_issues = {k:v for (k, v) in has_tickets.items() if v.tickets[0].category == "HDD"}
  ws0['A14'] = "Hard Drive"
  ws0['B14'] = len(hdd_issues)
  
  lcd_issues = {k:v for (k, v) in has_tickets.items() if v.tickets[0].category == "LCD"}
  ws0['A15'] = "LCD/Screen"
  ws0['B15'] = len(lcd_issues)

  psu_issues = {k:v for (k, v) in has_tickets.items() if v.tickets[0].category == "PSU"}
  ws0['A16'] = "Power Supply"
  ws0['B16'] = len(psu_issues)

  ram_issues = {k:v for (k, v) in has_tickets.items() if v.tickets[0].category == "RAM"}
  ws0['A17'] = "RAM/Memory"
  ws0['B17'] = len(ram_issues)

  other_issues = {k:v for (k, v) in has_tickets.items() if v.tickets[0].category == "OTHER"}
  ws0['A18'] = "Other"
  ws0['B18'] = len(other_issues)
  ws0['A19'] = "Other Issues Include:"
  idx = 19
  for device in other_issues:
    tickets = other_issues[device].tickets
    unique_tickets = []
    for ticket in tickets:
      if ticket.problem not in unique_tickets:
        unique_tickets.append(ticket.problem)
        ws0['B' + str(idx)] = ticket.problem
        idx = idx + 1

def create_gen_asset_info_ws():
  ws1 = wb.create_sheet()
  ws1.title = "General Asset Info"

  # Set Headers
  ws1['A1'] = "Purchased"
  ws1['B1'] = "Model"
  ws1['C1'] = "Service Tag"
  ws1['D1'] = "Warranty Expiration"
  ws1['E1'] = "Tickets"
  ws1['F1'] = "Ticket Categories"

  idx = 2
  for asset in filtered_devices:
    device = filtered_devices[asset]
    p_date = ""
    we_date_config = str(device.warranty_expiration).split()[0].split("-")
    we_date = we_date_config[1] + "-" + we_date_config[2] + "-" + we_date_config[0]
    ws1['A' + str(idx)] = format_date(device.purchase_date)
    ws1['B' + str(idx)] = device.model
    ws1['C' + str(idx)] = device.service_tag
    ws1['D' + str(idx)] = format_date(device.warranty_expiration)
    ws1['E' + str(idx)] = len(device.tickets)
    problem_types = ""
    for ticket in device.tickets:
      cat = ticket.category
      if cat not in problem_types:
        problem_types = problem_types + " " + cat
    ws1['F' + str(idx)] = problem_types
    idx = idx + 1

def create_ticket_view_ws():
  ws2 = wb.create_sheet()
  ws2.title = "Ticket Statistics"

  has_tickets = {k:v for (k, v) in filtered_devices.items() if len(v.tickets) > 0}
  # Set Headers
  ws2['A1'] = "Service Tag"
  ws2['B1'] = "Warranty Expiration Date"
  ws2['C1'] = "Ticket Created"
  ws2['D1'] = "Category"
  ws2['E1'] = "Problem"
  ws2['F1'] = "Status"

  idx = 2
  for asset in has_tickets:
    device = has_tickets[asset]
    tickets = device.tickets
    for ticket in tickets:
      ws2['A' + str(idx)] = device.service_tag
      ws2['B' + str(idx)] = format_date(device.warranty_expiration)
      ws2['C' + str(idx)] = format_date(ticket.date_created)
      ws2['D' + str(idx)] = ticket.category
      ws2['E' + str(idx)] = ticket.problem
      ws2['F' + str(idx)] = ticket.status
      idx = idx + 1

def create_warranty_ticket_ratios():
  ws3 = wb.create_sheet()
  ws3.title = "Warranty to Ticket Ratio"

  # Key = Month-Year, Value = [Expiring, Tickets]
  ws3['B1'] = "Devices Expiring"
  ws3['C1'] = "Tickets Created"

  ratio = dict()
  for asset in filtered_devices:
    device = filtered_devices[asset]
    # Device Updates
    warranty_date = format_date_yrmo(device.warranty_expiration)
    if warranty_date in ratio:
      ratio_date = ratio[warranty_date]
      ratio_date = [ratio_date[0] + 1, ratio_date[1]]
      ratio[warranty_date] = ratio_date
    else:
      ratio[warranty_date] = [1, 0]

    # Ticket Updates
    tickets = device.tickets
    for ticket in tickets:
      ticket_date = format_date_yrmo(ticket.date_created)
      if ticket_date in ratio:
        ratio_date = ratio[ticket_date]
        ratio_date = [ratio_date[0], ratio_date[1] + 1]
        ratio[ticket_date] = ratio_date
      else:
        ratio[ticket_date] = [0, 1]

  idx = 2
  for key in sorted (ratio.keys()):
    date = ratio[key]
    display_date = key.split("-")
    ws3['A' + str(idx)] = display_date[1] + "-" + display_date[0]
    ws3['B' + str(idx)] = date[0]
    ws3['C' + str(idx)] = date[1]
    idx = idx + 1

create_overview_assessment()
create_gen_asset_info_ws()
create_ticket_view_ws()
create_warranty_ticket_ratios()
wb.save(filename = dest_filename)