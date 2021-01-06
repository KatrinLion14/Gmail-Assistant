import pickle
import os
import os.path
import json
import base64
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from idna import unicode

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']


def auth():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('gmail', 'v1', credentials=creds)

    print_emails(service)
    logout()


def print_emails(service):
    result = service.users().messages().list(userId='me').execute()
    messages = result.get('messages')
    emails = []

    for msg in messages:
        txt = service.users().messages().get(userId='me', id=msg['id']).execute()
        try:
            payload = txt['payload']
            headers = payload['headers']
            subject = ''
            sender = ''
            for d in headers:
                if d['name'] == 'Subject':
                    subject = d['value']
                if d['name'] == 'From':
                    sender = d['value']
            parts = payload.get('parts')[0]
            data = parts['body']['data']
            data = data.replace("-", "+").replace("_", "/")
            body = unicode(base64.b64decode(data), 'utf-8')
            mail = dict()
            mail['subject'] = str(subject)
            mail['from'] = str(sender)
            mail['message'] = str(body)
            emails.append(mail)
        except:
            pass

    f = open('test.json', 'w', encoding="utf-8")
    f.write(json.dumps(emails, indent=4, ensure_ascii=False))
    f.close()


def logout():
    print('Do you want to logout? y/n')
    if input() == 'y':
        if os.path.exists('token.pickle'):
            os.remove(os.path.abspath('token.pickle'))


if __name__ == '__main__':
    auth()
