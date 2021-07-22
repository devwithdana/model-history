# Device History By Model

## Intent
Create a software to analyze the number of devices of a given model that had tickets created in Dell Tech Direct and the type of ticket to determine whether or not the warranty should be extended.

## Data Sources
Asset Tiger - Filtered by model name
Dell Tech Direct - All tickets submitted 

**Asset Tiger's headers:** 
* Asset Photo
* Asset Tag ID
* Brand	
* Description	
* Site	
* Location	
* Status	
* Category	
* Assigned to

**Dell's Tech Direct headers:**
* Work Order Code
* Status
* Service Tag
* Customer Full Name
* Description
* Create Timestamp Local	
* Dispatch Number


## Game Plan
* Filter Dell Tech Direct tickets by service tag to tickets that share a service tag with Asset Tiger's list of devices.
* Sort Dell Tech Direct tickets by ticket Description - this is the problem headline submitted on the ticket.
* Create categories based on Description column - Simplify by what the bad part was
* Compile pertinent information and export as JSON

## What We Want to Know
* Number of devices we have of that model
* How many devices had tickets submitted
* Of the devices that had tickets, what parts went bad?
* Number of devices in/out warranty
