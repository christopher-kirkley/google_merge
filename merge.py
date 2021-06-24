from __future__ import print_function
import os.path
import time
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

from config import TEMPLATE_DOCUMENT_ID, GOOGLE_SHEETS_ID, OUTPUT_FOLDER_ID

SCOPES = ['https://www.googleapis.com/auth/documents',
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive']


def auth():
    """Shows basic usage of the Docs API.
    Prints the title of a sample document.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def main():
    creds = auth()

    # Create instances
    service_DOCS = build('docs', 'v1', credentials=creds)

    time.sleep(2)

    # Build drive service
    service_DRIVE = build('drive', 'v3', credentials=creds)

    time.sleep(2)

    # Build drive service
    service_SHEETS = build('sheets', 'v4', credentials=creds)

    #Retrieve the documents contents from the Docs service.
    document = service_DOCS.documents().get(documentId=TEMPLATE_DOCUMENT_ID).execute()

    responses = {}

    responses['sheets'] = sheet = service_SHEETS.spreadsheets().values().get(
            spreadsheetId=GOOGLE_SHEETS_ID,
            range="Sheet1",
            majorDimension="ROWS",
            ).execute()

    columns = responses['sheets']['values'][0]
    records = responses['sheets']['values'][1:]

    # Iterate and perform mail merge

    def mapping(merge_field, value=''):
        json_representation = {
                'replaceAllText': {
                    'replaceText': value,
                    'containsText': {
                        'matchCase': 'true',
                        'text': '{{{{{0}}}}}'.format(merge_field)
                        }
                    }
                }
        return json_representation

    for i, record in enumerate(records, start=1):
        print('Processing record {0}...'.format(i))

        # copy template doc file as new file
        document_title = f'{record[0]}_{record[1]}_contract'

        responses['docs'] = service_DRIVE.files().copy(
                fileId=TEMPLATE_DOCUMENT_ID,
                body={
                    'parents': [OUTPUT_FOLDER_ID],
                    'name': document_title
                    }
                ).execute()
        document_id = responses['docs']['id']
        print(document_id)

        # update Google Docs document
        merge_fields_information = [mapping(columns[indx], value) for indx, value in enumerate(record)]
        service_DOCS.documents().batchUpdate(
                documentId=document_id,
                body={
                    'requests': merge_fields_information
                    }
                ).execute()



if __name__ == '__main__':
    main()
