import datetime

class Device: 
  def __init__(self, service_tag, model, purchase_date, warranty_expiration):
    self.service_tag = service_tag
    self.model = model
    self.purchase_date = purchase_date
    parsing_warranty = warranty_expiration.split('/')
    month = int(parsing_warranty[0])
    day = int(parsing_warranty[1])
    year = int(parsing_warranty[2])
    warranty_date = datetime.datetime(year, month, day)
    self.warranty_expiration = warranty_date

  def hasExpired(self):
    current_date = datetime.datetime.today()
    return current_date > warranty_date

class Ticket:
  def __init__(self, work_order, status, problem, date_created):
    self.work_order = work_order
    self.status = status
    self.problem = problem
    self.date_created = convertDate(date_created)
    self.category = determineCategory()

  def convertDate(self, date_created):
    parsing_date = date_created.split()[0].split('/')
    month = int(parsing_date[0])
    day = int(parsing_date[1])
    year = int(parsing_date[2])
    ticket_date = datetime.datetime(year, month, day)

  def determineCategory(self):
    problem_categories = {
      "MOBO": ["motherboard"], 
      "HDD": ["hard drive", "hdd", "solid state drive", "ssd"], 
      "LCD": ["lcd", "display", "screen"], 
      "PSU": ["power"], 
      "RAM": ["memory"],
      "OTHER": []
    }

    ticket_problem = ""
    multi_category = ""
    for category in problem_categories:
      key_words = problem_categories[category]
      for word in key_words:
        if word.lower() in self.problem.lower():
          if ticket_problem == "":
            ticket_problem = category
          else:
            multi_category = ticket_problem + ", " + category
            ticket_problem = "OTHER"

    if ticket_problem == "OTHER":
      print("Warning: This device might fall into multiple categories.")
      print("Device may fall into any of these categories: " + multi_category)
      
    if ticket_problem == "":
      ticket_problem = "OTHER"

    return ticket_problem

  def occurredBeforeExpiration(self, warranty_expiration):
    return self.date_created < warranty_expiration