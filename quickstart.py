from __future__ import print_function
import httplib2
import os
import base64
import email
import re
import csv

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/gmail-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/gmail.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Gmail API Python Quickstart'
csv_file_dest = 'orders.csv'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'gmail-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def GetMessageContents(mime_msg):
    messageMainType = mime_msg.get_content_maintype()
    if messageMainType == 'multipart':
        for part in mime_msg.get_payload():
            return GetMessageContents(part)
    elif messageMainType == 'text':
        return mime_msg.get_payload()


def GetMessageBody(service, user_id, msg_id):
    try:
            message = service.users().messages().get(userId=user_id, id=msg_id, format='raw').execute()
            msg_str = base64.urlsafe_b64decode(message['raw'].encode('ASCII'))
            mime_msg = email.message_from_string(msg_str)
            return GetMessageContents(mime_msg)
            # messageMainType = mime_msg.get_content_maintype()

            # if messageMainType == 'multipart':
            #         for part in mime_msg.get_payload():
            #                 if part.get_content_maintype() == 'text':
            #                         return part.get_payload()
            #         return ""
            # elif messageMainType == 'text':
            #         return mime_msg.get_payload()
    except Exception as error:
            print('An error occurred: %s' % error)

def parseOrderEmail(msg):
    msg = msg.replace('\r','')
    msg = msg.replace('=\n','')
    # msg = msg.replace('\n\n', '')
    order_number = re.findall(r"Order #: ([0-9]+)", msg)[0]
    order_date = re.findall(r"Order placed: ([0-9a-zA-Z, ]+)", msg)[0]
    total_cost = re.findall(r"Total cost: \$([0-9.]+)", msg)[0]
    order_tax = re.findall(r"Shipping:[ \t]+\$[0-9.]+[ \t]+\$([0-9.]+) -+", msg)[0]
    header = re.search(r"\*+\nORDER DETAILS\n\*+", msg)
    footer = re.search(r"----------------------------------", msg)
    order_table = msg[header.end()+1:footer.start()-1]
    order_lines_raw = order_table.split('\n')
    order_lines = []
    i = 0
    order_line = {"item_notes":""}
    for line in order_lines_raw:
        if ("Item:" in line):
            item_description = re.findall(r"Item: ([^\n]+)", line)[0]
            order_line["item_description"] = item_description
        elif ("Quantity:" in line):
            a = line.split('Quantity:')
            quantity = re.findall(r"Quantity: ([0-9.]+)", line)[0]
            order_line["item_notes"] += a[0].replace('\t', '')
            order_line["quantity"] = a[1]
        elif ("Item Total:" in line):
            order_line["line_total"] = re.findall(r"Item Total: \$([0-9.]+)", line)[0]
            order_lines.append(order_line)
            order_line = {"item_notes":""}
    order = {
        "order_number":order_number,
        "order_date":order_date,
        "order_tax":order_tax,
        "order_total":total_cost,
        "order_lines":order_lines
    }
    return(order)



def main():
    """Shows basic usage of the Gmail API.

    Creates a Gmail API service object and outputs a list of label names
    of the user's Gmail account.
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('gmail', 'v1', http=http)

    results = service.users().messages().list(userId='me', q='from:no-reply-commerce@wix.com OR from:no-reply@mystore.wix.com').execute()
    messages = results.get('messages', [])

    if not messages:
        print('No messages found.')
    else:
        print('messages:')
        row = '%(order_number)s\t%(order_date)s\t%(order_tax)s\t%(order_total)s\t%(item_description)s\t%(quantity)s\t%(line_total)s\t%(item_notes)s\n'
        processed_lines = []
        with open(csv_file_dest, 'w') as outputFile:
            outputFile.write(row % {
                'order_number': "order_number",
                'order_date': "order_date",
                'order_tax': "order_tax",
                'order_total': "order_total",
                'item_description': "item_description",
                'quantity': "quantity",
                'line_total': "line_total",
                'item_notes': "item_notes"
            })
            for msg in messages:
                msgBody = GetMessageBody(service, 'me', msg['id'])
                order = parseOrderEmail(msgBody)
                print(order)
                for line in order["order_lines"]:
                    line_key = order["order_number"]+"_"+line["item_description"]+"_"+line["item_notes"]
                    if line_key in processed_lines:
                        print(line_key +" already processed")
                    else:
                        outputFile.write(row % {
                            'order_number': order["order_number"],
                            'order_date': order["order_date"],
                            'order_tax': order["order_tax"],
                            'order_total': order["order_total"],
                            'item_description': line["item_description"],
                            'quantity': line["quantity"],
                            'line_total': line["line_total"],
                            'item_notes': line["item_notes"]
                        })
                        processed_lines.append(line_key)



if __name__ == '__main__':
    main()