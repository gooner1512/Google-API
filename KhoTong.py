from googleapiclient.http import MediaFileUpload
from Google import create_service
from sqlserver import Database
import pandas as pd
from decimal import *
import datetime

def construct_request_body(value_array, dimension: str='ROWS') -> dict:
    try:
        request_body = {
            'majorDimension': dimension,
            'values': value_array
        }
        return request_body
    except Exception as e:
        print(e)
        return {}


CLIENT_SECRET_FILE = 'credentials.json'
API_SERVICE_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

service = create_service(CLIENT_SECRET_FILE, API_SERVICE_NAME, API_VERSION, SCOPES)


spreadsheet_id = '1CwJ2N2PkGdFtf_reVZC-LePvhEWPuBUdYSoSmvxg0EU'
mySpreadsheets = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()

worksheetname_list = ['nk_xk_tk_1','nk_xk_tk_2','nk_xk_tk_3']

db = Database()
query="""
EXEC USP_KIDS_LayDuLieuKhoTong '20230313','20230313'
"""

dataframes = db.run_query_multi_tables(query)

for x in range(len(dataframes)):    
    df = dataframes[x]
    recordset = df.values.tolist()
    for i in range(len(recordset)):
        for j in range(len(recordset[i])):
            if isinstance(recordset[i][j], Decimal):
                recordset[i][j] = float(recordset[i][j])
            elif isinstance(recordset[i][j], datetime.date):
                recordset[i][j] = recordset[i][j].strftime('%Y-%m-%d')

    columns = df.columns.tolist()   

    # Determine the number of rows already in the sheet
    response = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"{worksheetname_list[x]}!A1:A",
        majorDimension="COLUMNS",
    ).execute()
    num_rows = len(response["values"][0])
    range_start = num_rows + 1
    #Insert Data rows
    request_body_values = construct_request_body(recordset)
    service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        valueInputOption='USER_ENTERED',
        insertDataOption='INSERT_ROWS',
        range=f'{worksheetname_list[x]}!A{range_start}',
        body=request_body_values
    ).execute()
    #Clear data
    # service.spreadsheets().values().clear(
    #     spreadsheetId=spreadsheet_id,
    #     range=f'{worksheetname_list[x]}',
    # ).execute()
    #Insert Header row
    # request_body_columns = construct_request_body([columns])
    # service.spreadsheets().values().update(
    #     spreadsheetId=spreadsheet_id,
    #     valueInputOption='USER_ENTERED',
    #     range=f'{worksheetname_list[x]}!A1',
    #     body=request_body_columns
    # ).execute()
    #Insert Data rows
    # request_body_values = construct_request_body(recordset)
    # service.spreadsheets().values().update(
    #     spreadsheetId=spreadsheet_id,
    #     valueInputOption='USER_ENTERED',
    #     range=f'{worksheetname_list[x]}!A{range_start}',
    #     body=request_body_values
    # ).execute()

print('Task is complete')


