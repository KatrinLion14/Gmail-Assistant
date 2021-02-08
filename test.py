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

urgent_keywords = ['срочно', 'СРОЧНО', 'Срочно', 'важно', 'ВАЖНО', 'Важно']


def log_in():
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
    return service


def decoding_message(mail):
    data = mail.replace("-", "+").replace("_", "/")
    body = unicode(base64.b64decode(data), 'utf-8')
    return body


def print_into_json(name, obj):
    f = open(name, 'w', encoding="utf-8")
    f.write(json.dumps(obj, indent=4, ensure_ascii=False))
    f.close()


def get_emails(service, query):
    result = service.users().messages().list(userId='me', q=query).execute()
    messages = result.get('messages')
    emails = []
    # txts = []
    for msg in messages:
        txt = service.users().messages().get(userId='me', id=msg['id']).execute()
        # txts.append(txt)
        # print(txt)
        # if 'body' in txt['payload']:
        #     attachment = service.users().messages().attachments().get(userId='me', messageId=txt['id'], id=txt['payload']['body']['attachmentId']).execute()
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
            if 'parts' in parts:
                data = parts['parts'][0]['body']['data']
            else:
                data = parts['body']['data']
            body = decoding_message(data)
            mail = dict()
            mail['subject'] = str(subject)
            mail['from'] = str(sender)
            mail['message'] = str(body)
            emails.append(mail)
        except:
            pass
    # print_into_json('fixing.json', txts)
    return emails


def find_keywords_emails(service, query, keywords, keyword):
    emails = get_emails(service, query)
    urgent_emails = []
    for mail in emails:
        for word in keywords:
            if (word in mail['message']) or (word in mail['subject']):
                urgent_emails.append(mail)
    name = str(keyword) + '_messages.json'
    print_into_json(name, urgent_emails)


def find_urgent_emails(service):
    find_keywords_emails(service, '', urgent_keywords, 'urgent')


def print_emails(service, query):
    emails = get_emails(service, query)
    print_into_json('all_emails.json', emails)


def logout():
    print('Do you want to logout? y/n')
    if input() == 'y':
        if os.path.exists('token.pickle'):
            os.remove(os.path.abspath('token.pickle'))


def main():
    service = log_in()
    print_emails(service, '')
    find_urgent_emails(service)
    keyword = []
    keyword.append('Testing')
    find_keywords_emails(service, '', keyword, 'testing')
    logout()


if __name__ == '__main__':
    main()
