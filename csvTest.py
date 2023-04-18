import gspread
import csv
from google.oauth2.service_account import Credentials

# define the scope of the credentials
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]

# create the credentials object
creds = Credentials.from_service_account_file(
    r"C:\Users\Andreo\Desktop\Backlog\airtable-backlog-c8c87395ab42.json", scopes=scope
)

gc = gspread.service_account(r"C:\Users\Andreo\Desktop\Backlog\airtable-backlog-c8c87395ab42.json")

# authenticate with gspread using the credentials
gc = gspread.authorize(creds)

# open the spreadsheet by URL
url = "https://docs.google.com/spreadsheets/d/1QX9rUCmFRTY7IjwaZsLiFVql_iOb6OPfMfoyk5FtBDA/edit#gid=395646271"
sheet = gc.open_by_url(url).sheet1

# get all values from the first column
item_numbers = sheet.col_values(1)

# open the airtable using the airtable-python-wrapper
from airtable import Airtable

# fill in your airtable details
base_key = "appOo3NIs0nT86pYB"
table_name = "Formulas Test"
api_key = "keyeDQ87k9Looanna"
item_no_col = "ItemNo+Col"

# create the airtable object
airtable = Airtable(base_key, table_name, api_key)

# create two lists to store the found and not found item numbers
found_item_numbers = []
not_found_item_numbers = []

# loop through each item number in the spreadsheet
for item_number in item_numbers[1:]: # skip the header row
    # search for the item number in the airtable column
    result = airtable.search(item_no_col, item_number)
    if result:
        # add the item number to the found list
        found_item_numbers.append(item_number)
        print(f"Found {item_number}")
    else:
        # add the item number to the not found list
        not_found_item_numbers.append(item_number)
        print(f"Not found {item_number}")

# create two CSV files to store the results
with open("found_items.csv", "w", newline="") as found_file, open("not_found_items.csv", "w", newline="") as not_found_file:
    found_writer = csv.writer(found_file)
    not_found_writer = csv.writer(not_found_file)
    # write the header row to each file
    found_writer.writerow(["Item Number"])
    not_found_writer.writerow(["Item Number"])
    # write the found item numbers to the first file
    for item_number in found_item_numbers:
        found_writer.writerow([item_number])
        print(f"Added {item_number} to found list")
    # write the not found item numbers to the second file
    for item_number in not_found_item_numbers:
        not_found_writer.writerow([item_number])
        print(f"Added {item_number} to not found list")

# print a message to indicate that the script has finished running
print("Done!")
