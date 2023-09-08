import os
import pandas as pd
from pprint import pprint
from Google import create_service

CLIENT_SECRET_FILE = 'client_secret_566868905417-g3difppfkga03a99hdvofmnguokj4hgn.apps.googleusercontent.com.json'
API_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

service = create_service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)
spreadsheet_id = '1AH8H_MRaQLyNL0h7WYR6gw64Pqad2E-dDKmySJggbe4'


def get_data_range(range: str):
    """Getting data knowing the end of file"""
    response = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        majorDimension='ROWS',
        range=range
    ).execute()

    columns = response['values'][0]
    data = response['values'][1:]
    df = pd.DataFrame(data, columns=columns)
    df2 = df.set_index('Invoice')
    print(df2)
    return df2['Product']


# def get_data(sheet: str, cells_range=list[str]):
#     """Getting data not knowing the end of file"""
#     response = service.spreadsheets().values().get(
#         spreadsheetId=spreadsheet_id,
#         majorDimension='ROWS',
#         # 'Sales North!B2:F3'
#         range=f'{sheet}!{cells_range[0]}:{cells_range[1]}'
#     ).execute()

#     # print(response['values'])
#     # columns = response['values'][1][1:]
#     columns = ['Invoice', 'Product', 'Quantity', 'Unit Price', 'Sales Total']
#     data = [item[0:] for item in response['values'][1:]]
#     df = pd.DataFrame(data, columns=columns)
#     print(df)
#     return df


def get_data_batch():
    """Getting data using batchGet method"""
    value_ranges_body = [
        'Sales North!B2:F17',
        'Sales South!B2:F11',
        'Sales West!B2:F11',
        'Sales East!B2:F11'
    ]

    response = service.spreadsheets().values().batchGet(
        spreadsheetId=spreadsheet_id,
        majorDimension='ROWS',
        ranges=value_ranges_body
    ).execute()
    print(response.keys())
    print(response['valueRanges'])

    dataset = {}
    for item in response['valueRanges']:
        dataset[item['range']] = item['values']
    print(dataset.keys())

    df = {}
    for index, k in enumerate(dataset):
        columns = dataset[k][0]
        data = dataset[k][1:]
        df[index] = pd.DataFrame(data, columns=columns)

    print(df.keys())
    print(df[0])


# update cells values
def update_sheet_cells(spreadsheet_id: str, range_: str, values: list):

    value_input_option = 'RAW'
    value_range_body = {
        'values':
        values
    }
    request = service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_,
        body=value_range_body,
        valueInputOption=value_input_option,
    )
    response = request.execute()
    pprint(response)
