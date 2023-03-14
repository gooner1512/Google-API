import pickle
import os.path
import os
import logging
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCOPES = ['https://www.googleapis.com/auth/drive']
SHEET_TOKEN = "credentials.json"
SHEET_PICKLE = "token.json"

class PySheet:
    def __init__(self, scopes=SCOPES, token_file=SHEET_TOKEN):
        creds = None
        if os.path.exists(SHEET_PICKLE):
            with open(SHEET_PICKLE, 'rb') as token:
                creds = pickle.load(token)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(token_file, scopes)
                # creds = flow.run_console()
                creds = flow.run_local_server()
            with open(SHEET_PICKLE, 'wb') as token:
                pickle.dump(creds, token)
        self.gg_api = build(
            'sheets',
            'v4',
            credentials=creds,
            cache_discovery=False)

    def gs_clear(self, spreadsheet_id, range):
        '''

        :param spreadsheet_id:
        :param range:
        :return:
        '''
        clear = self.gg_api\
            .spreadsheets()\
            .values()\
            .clear(
                spreadsheetId=spreadsheet_id,
                range=range).execute()
        print(clear)


    def gs_write(self, spreadsheet_id, range, data):
        '''

        :param spreadsheet_id:
        :param range:
        :param data:
        :return:
        '''
        append_to_sheets = self.gg_api\
            .spreadsheets()\
            .values()\
            .append(
                spreadsheetId=spreadsheet_id,
                range=range,
                valueInputOption="USER_ENTERED",
                body={'values': data}).execute()
        print(append_to_sheets)

    def gs_read(self, spreadsheet_id, range):
        '''

        :param spreadsheet_id:
        :param range:
        :return:
        '''
        get_data = self.gg_api\
            .spreadsheets()\
            .values()\
            .get(
                spreadsheetId=spreadsheet_id,
                range=range)
        result = get_data.execute()
        df = []
        for i in result.get('values'):
            df.append(i)
        return df

    def gs_batch_update(self, spreadsheet_id, data):
        '''

        :param spreadsheet_id:
        :param data:
        :return:
        '''
        batch_update = self.gg_api\
            .spreadsheets()\
            .values()\
            .batchUpdate(
                spreadsheetId=spreadsheet_id,
                body=data).execute()
        print(batch_update)

