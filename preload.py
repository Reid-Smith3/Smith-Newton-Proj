import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import tkinter as tk

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# setting the id and ranges for the movie database
MOVIES_SPREADSHEET_ID = '1XrT1fRhhiS_Ga7yX9AEikkK5uFoGOfIIX6UpOXKLPLI'
test_range = 'A2:I158548'

""" This file is intended to be used if the user wants
to store the database as a dictionary for repeated access
on their computer, essentially making it a desktop application.
Otherwise, continue loading from the API. It would be good
practice for the user to run this file every month or so
to update their local database with changes from the main database."""

if __name__ == '__main__':
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result_test = sheet.values().get(spreadsheetId=MOVIES_SPREADSHEET_ID,
                                range=test_range).execute()
    
    # save to pickle file
    test_values = result_test.get('values', [])
    test_dict = {info[0]:info[1:] for info in test_values}

    movie_file = open('moviedata.pickle', 'wb')
    pickle.dump(test_dict, movie_file)
    movie_file.close()