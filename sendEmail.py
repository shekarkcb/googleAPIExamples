
"""
Script : sendEmail.py
Desc   : Sending email using gmail API's
        1. Create a project in google api console.
        2. Allow gmail API permissions to that project
        3. Create oauth tokens for that project
        4. download the token informaton as token.json and keep in the same directory as this file
        5. sample.csv is the file used as attachment to be sent.
        6. fromAddress - address from which gmail should send email. 
        7. On the first run, it will try to authorize using the token.json plus gmail authentication page.
        8. Enter the same creds as fromAddress to authenticate/authorize and click allow.
        9. Flow is complete, then it downloads token.pickel, both toekn.json & token.pickel is required to send emails
        without hitting authorization page again.
        10. message_text - this has sample html page body to be sent. Modify accordingly
Modified : @shekarkcb        
"""
from __future__ import print_function
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import base64, httplib2
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import mimetypes
from email.mime.image import MIMEImage
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
import pickle

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.send']

fromAddress='your_from_address@your_email.com'
subject="Testing sending email using gmail API"
file2send="/path/to/your/sample/file"
message_text="""
    <!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sending email</title>
</head>
<body>
    <h6>How are you doing today</h6>
    <ul>
        <li>1.Hi</li>
        <li>2.Hello</li>
        <li>3.How are you</li>
    </ul>
</body>
</html>
"""


def create_message_with_attachment(
    sender, to, subject, message_text, file):
  """Create a message for an email.

  Args:
    sender: Email address of the sender.
    to: Email address of the receiver.
    subject: The subject of the email message.
    message_text: The text of the email message.
    file: The path to the file to be attached.

  Returns:
    An object containing a base64url encoded email object.
  """
  message = MIMEMultipart()
  message['to'] = to
  message['from'] = sender
  message['subject'] = subject

  msg = MIMEText(message_text)
  message.attach(msg)

  content_type, encoding = mimetypes.guess_type(file)
  # print(f"Got the content type {content_type},{encoding}")
  if content_type is None or encoding is not None:
    content_type = 'application/octet-stream'
  main_type, sub_type = content_type.split('/', 1)
  # print(f"Got the content type {content_type},{encoding},{main_type},{sub_type}")
  if main_type == 'text':
    fp = open(file, 'r')
    msg = MIMEText(fp.read(), _subtype=sub_type)
    # msg = MIMEText(fp.read().as_bytes(),_subtype=sub_type)
    fp.close()
  elif main_type == 'image':
    fp = open(file, 'r')
    msg = MIMEImage(fp.read(), _subtype=sub_type)
    fp.close()
  elif main_type == 'audio':
    fp = open(file, 'r')
    msg = MIMEAudio(fp.read(), _subtype=sub_type)
    fp.close()
  else:
    fp = open(file, 'r')
    msg = MIMEBase(main_type, sub_type)
    msg.set_payload(fp.read())
    fp.close()
  filename = os.path.basename(file)
  msg.add_header('Content-Disposition', 'attachment', filename=filename)
  message.attach(msg)

  # return {'raw': base64.urlsafe_b64encode(message.as_string())}
  # return {'raw': base64.urlsafe_b64encode(message.as_bytes())}
  # return {'raw': base64.urlsafe_b64encode(pickle.dumps(message))}
  # raw = base64.urlsafe_b64encode(pickle.dumps(message))
  # For python3.x this is required.
  raw = base64.urlsafe_b64encode(message.as_bytes())
  raw = raw.decode()
  return {'raw': raw}


def send_message(service, user_id, message):
  """Send an email message.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    message: Message to be sent.

  Returns:
    Sent Message.
  """
  try:
    message = (service.users().messages().send(userId=user_id, body=message)
               .execute())
    print('Message Id: %s' % message['id'])
    return message
  except httplib2.HttpLib2Error as  error:
    print ('An error occurred: %s' % error)


def main():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
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

    service = build('gmail', 'v1', credentials=creds)

    # # Call the Gmail API to list lables.
    # results = service.users().labels().list(userId='me').execute()
    # labels = results.get('labels', [])

    # if not labels:
    #     print('No labels found.')
    # else:
    #     print('Labels:')
    #     for label in labels:
    #         print(label['name'])

    message2Send = create_message_with_attachment(
        fromAddress,
        subject,
        message_text,
        file2send)
    print(f'sending {message2Send}')
    send_message(service, "me", message2Send)

if __name__ == '__main__':
    main()
