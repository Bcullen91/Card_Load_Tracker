from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import base64
import email
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
from bs4 import BeautifulSoup

# Must pip install these packages -- use pip3 if on a mac.
# pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
# pip install gspread
# pip install beautifulsoup4
# pip install oauth2client

# CHANGE THIS TO YOUR PATH. Must use double \\'s
path = 'C:\\Users\\Brian\\Desktop\\Projects\\Card_Loads\\'

sheet_name = 'Serve/BB Loads'


# MUST DO THE STEPS HERE BEFORE IT WILL WORK. Store the files in the same working directory as this file.
# https://developers.google.com/gmail/api/quickstart/python
# If modifying these scopes, delete the file token.pickle.
scope = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']

drive_creds = ServiceAccountCredentials.from_json_keyfile_name(path + 'client_secret.json', scope)
client = gspread.authorize(drive_creds)
sheet = client.open(sheet_name).sheet1
num_row = sheet.col_values(1)
row = len(num_row) + 1
user_id = 'me'
subject = 'You added cash to your Account'
exceptions = 0


def get_service():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(path + 'token.pickle'):
        with open(path + 'token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(path +
                'credentials.json', scope)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(path + 'token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    service = build('gmail', 'v1', credentials=creds)
    return service


# This function searches gmail folder for the search string and returns the selected email id's.
def search_messages(service, user_id, search_string):
    try:
        search_id = service.users().messages().list(userId=user_id, q=search_string, labelIds=['UNREAD']).execute()
        number_results = search_id['resultSizeEstimate']
        final_list = []
        if number_results > 0:
            message_ids = search_id['messages']
            for ids in message_ids:
                final_list.append(ids['id'])
            return final_list
        else:
            print('There were no results for your search.')
            return ""
    except Exception as e:
        print('An error occurred in the search message function.')
        print(e)


# This function gets the found email information and changes the data type to a readable datatype.
def get_message(service, user_id, msg_id):
    try:
        # Makes the connection and GETS the emails in RAW format.
        message = service.users().messages().get(userId=user_id, id=msg_id, format='raw').execute()
        # Changes format from RAW to ASCII
        msg_raw = base64.urlsafe_b64decode(message['raw'].encode('ASCII'))
        # Changes format type again
        msg_str = email.message_from_bytes(msg_raw)
        # This line checks what the content is, if multipart (plaintext and html) or single part
        if 'Serve®' in message['snippet']:
            soup = BeautifulSoup(msg_str.get_payload(), 'lxml')
            text = base64.b64decode(soup.get_text())
            soup2 = BeautifulSoup(text, 'lxml')
            text2 = soup2.get_text()
            return text2
        if 'Bluebird' in message['snippet']:
            soup = BeautifulSoup(msg_str.get_payload(), 'lxml')
            return soup.get_text()
    except Exception as e:
        print('An error has occured during the get_message function.')
        print(e)


def get_data(raw_email):
    amount_line = 'no'
    date_line = 'no'
    date = ''
    amount = ''
    account = ''
    for line in raw_email:
        if amount_line == 'yes':
            amount = line.strip()
            amount_line = 'no'
        if date_line == 'yes':
            date = line.strip()
            date_line = 'no'
        if 'Account ending' in line:
            account = line.strip()[-4:]
        if "Amount" in line:
            amount_line = 'yes'
        if 'Submitted on' in line:
            date_line = 'yes'
    return date, amount, account


def mark_read(service, user_id, msg_id):
    msg = service.users().messages().modify(userId=user_id, id=msg_id, body={'removeLabelIds': ['UNREAD']}).execute()
    return msg


if __name__ == '__main__':
    service = get_service()
    last_checked = sheet.col_values(10)
    latest_email = [id for id in last_checked if id != ''][-1]
    times_ran = 0
    while True:
        times_ran += 1
        account = ''
        amount = ''
        date = ''
        try:
            service = get_service()
            msg_ids = search_messages(service, user_id, subject)
            if len(msg_ids) > 0:
                latest_email = msg_ids[0]
            print('This script has been ran ' + str(times_ran) + ' times.')
            if latest_email in last_checked:
                pass
            else:
                raw_text = get_message(service, user_id, latest_email).split('\n')
                date, amount, account = get_data(raw_text)
                drive_creds = ServiceAccountCredentials.from_json_keyfile_name(path + 'client_secret.json', scope)
                client = gspread.authorize(drive_creds)
                sheet = client.open('Serve/BB Loads').sheet1
                if amount != '' and account != '' and date != '':
                    print('-------------------------------UPDATING!!!!-----------------------------------')
                    print('date: ' + date + ' Amount: ' + amount + ' Account: ' + account)
                    sheet.update_cell(row, 1, date)
                    sheet.update_cell(row, 2, amount)
                    sheet.update_cell(row, 3, account)
                    sheet.update_cell(row, 4, 'This was inserted by code.')
                    sheet.update_cell(row, 10, latest_email)
                    row += 1
                    mark_read(service, user_id, latest_email)
                    print('A Transaction was successfully updated')
                    sheet.update_cell(48, 1, str(latest_email))
                    last_checked = str(latest_email)
                else:
                    print("The script was unable to pull the information from the email.")
            time.sleep(10)
        except Exception as e:
            exceptions += 1
            print('something went wrong and the entire script errorred out')
            print(e)
            time.sleep(100)

