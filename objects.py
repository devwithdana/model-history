import datetime

class Device: 
  def __init__(self, service_tag, model, purchase_date, warranty_expiration):
    self.service_tag = service_tag
    self.model = model
    self.purchase_date = purchase_date
    self.warranty_expiration = warranty_expiration
    self.tickets = []

  def hasExpired(self):
    current_date = datetime.datetime.today()
    return current_date > self.warranty_expiration

class Ticket:
  def __init__(self, work_order, device_tag, status, problem, date_created):
    self.device_tag = device_tag
    self.work_order = work_order
    self.status = status
    self.problem = problem
    self.date_created = date_created
    self.category = self.determineCategory()

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
          elif ticket_problem != category:
            multi_category = ticket_problem + ", " + category
            ticket_problem = "OTHER"

    if ticket_problem == "OTHER":
      print("Warning: This device might fall into multiple categories.")
      print("Device may fall into any of these categories: " + multi_category + "\n")

    if ticket_problem == "":
      ticket_problem = "OTHER"

    return ticket_problem

  def occurredBeforeExpiration(self, warranty_expiration):
    return self.date_created < warranty_expiration