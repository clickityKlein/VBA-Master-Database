# VBA-Master-Database
This VBA project allows for a command center to parse through different excel files, gathering information to send to a "master" database.

## Background
The task was to create a database that housed information about each client, updated from their client Excel workbooks. The main obstacle was that each client's workbook (and sometimes each year in the workbook) were unique. Essentially, it wasn't feasible to create a scraper that simply went through the workbooks and gathered this information. However, placing a uniform data input page in each of these workbooks would allow for easy scraping of information into a single place. The datapoints in the final project were extensive and exhaustive for the information contained across all clients. What's presented here is the template design for this scraper, illustrated with just a few datapoints.

## Initial Setup
Each client will have an input tab that has general information (constant through the years of data) and variable information (likely to change through the years of data).

Here is an example (Client1):
![Client Input Example](Pictures/ClientExample.png)
Please note that the name of the input tabs DO NOT need to be the same across all clients. But, do take note of the tab name.

After the input page is setup, open the "Directory" tab in the "MDB Controls - Link" file (this is essentially the command center). Entered on this page is the client name, location to the excel file, and location to the input tab. Once again, the name of the input tabs DO NOT need to be the same across all clients. They just are in this example.

Six mock clients are entered in the example, with the "ClientError" client being there for an error handling example.

Example:
![Directory Example](Pictures/Directory/Directory_Before.png)
Don't worry about entering in the number of plan years, this is filled automatically from the client input tab when the program is run.
![Directory Example](Pictures/Directory/Directory_After.png)

The hard part is over! Just a few selections on the Controls page, and the automation will commence.

## Updating the Database
Starting on the "Controls" tab of the "MDB Controls - Link" file, there are a few buttons present:
- Full Pull
- Select Client(s)
- Subset Pull

![Controls](Pictures/Controls/Controls_Basic.png)

Full Pull: this will update the database across every client in the directory. After the database has been updated, a pop-up will appear:
![Controls](Pictures/Controls/Controls_Full_Pull.png)

Select Client(s): If only a subset of clients need to be updated, use the Select Client(s), which will open a form:
![Controls](Pictures/Controls/Controls_Select.png)

By holding ctrl, it is possible to select multiple clients:
![Controls](Pictures/Controls/Controls_Select_Indv.png)

After the client selection is made, the client(s) selected will appear under "Update the Following Subset" text:
![Controls](Pictures/Controls/Controls_Select_Indv_After.png)

Subset Pull: only the client(s) which appear under the "Update the Following Subset" text will be updated when this button is pressed. After the database has been updated, a pop-up will appear:
![Controls](Pictures/Controls/Controls_Subset_Pull.png)

Additionally, notice how the under the text "Directory Issues Found for the Following Client(s)" is our error handle testing client "ClientError". A false address was given to this client to illustrate that in the event a directory error occurs, the code will keep running, but the issue will be noted.

Note: in the back-end of things, directoy errors are caught when the program fetches number of years of data. If there is an error, the number of years will simply be 0 (look back at the directory example above to see the years of data for ClientError).


















