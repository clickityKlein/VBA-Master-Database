# VBA-Master-Database
This VBA project allows for a command center to parse through different excel files, gathering information to send to a "master" database.

## Background
The task was to create a database that housed information about each client, updated from their client Excel workbooks. The main obstacle was that each client's workbook (and sometimes each year in the workbook) were unique. Essentially, it wasn't feasible to create a scraper that simply went through the workbooks and gathered this information. However, placing a uniform data input page in each of these workbooks would allow for easy scraping of information into a single place. The datapoints in the final project were extensive and exhaustive for the information contained across all clients. What's presented here is the template design for this scraper, illustrated with just a few datapoints.

## Initial Setup
Each client will have an input tab that has general information (constant through the years of data) and variable information (likely to change through the years of data).

Here is an example (Client1):
![Cient Input Example](Pictures/ClientExample.png)
