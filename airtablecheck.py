import gspread
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

# create a new spreadsheet to store the results
folder_id = "1Z5d6pbiBt4Rn9mpKHS3t-G2WASbzIxv2"
results_sheet = gc.create("Item Number Results", folder_id)

# create two worksheets in the new spreadsheet
found_sheet = results_sheet.add_worksheet(title="Found", rows=1, cols=1)
not_found_sheet = results_sheet.add_worksheet(title="Not Found", rows=1, cols=1)

# write the found item numbers to the first sheet
found_sheet.insert_row(["Item Number"], 1)
for item_number in found_item_numbers:
    found_sheet.insert_row([item_number], found_sheet.row_count + 1)
    print(f"Added {item_number} to found list")

# write the not found item numbers to the second sheet
not_found_sheet.insert_row(["Item Number"], 1)
for item_number in not_found_item_numbers:
    not_found_sheet.insert_row([item_number], not_found_sheet.row_count + 1)
    print(f"Added {item_number} to not found list")

# print a message to indicate that the script has finished running
print("Done!")
