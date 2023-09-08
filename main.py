import pandas as pd
from pprint import pprint
from Google import create_service
from sheets.sheets import update_sheet_cells, get_data_range
from drive.drive import create_folder

# CLIENT_SECRET_FILE = 'client_secret_566868905417-g3difppfkga03a99hdvofmnguokj4hgn.apps.googleusercontent.com.json'
# API_NAME = 'sheets'
# API_VERSION = 'v4'
# SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# service = create_service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)
spreadsheet_id = '1AH8H_MRaQLyNL0h7WYR6gw64Pqad2E-dDKmySJggbe4'

# UPDATE CELLS
# values = [[115, 'Mouseee', 1000, 1, 10000], [116, 'Keyboard', 5000, 5, 50000], [
#     115, 'Mouseee', 1000, 1, 10000], [116, 'Keyboard', 5000, 5, 50000]]

# update_sheet_cells(spreadsheet_id=spreadsheet_id,
#                    range_='Sales North!B18:F21', values=values)


# CREATING FOLDER
row = get_data_range('Sales North!B2:F3')
print(row)
row_product = []
row_product.append(row)
parent_folder_name = ['1Rd5MVPuQUQ97G8JFfH8DvPCh96AWsy8W']
create_folder(folder_names=row_product, parent_folder_name=parent_folder_name)
