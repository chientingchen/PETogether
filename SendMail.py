#-*- coding: utf-8 -*-
import base64
import httplib2
import mimetypes
import os,csv,codecs,sys,time
from os import listdir
from os.path import isfile, join
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from apiclient import errors
from apiclient.discovery import build
from oauth2client.client import flow_from_clientsecrets
from oauth2client.file import Storage
from oauth2client.tools import run

def SendMessage(service, user_id, message):
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
    print 'Message Id: %s' % message['id']
    return message
  except errors.HttpError, error:
    print 'An error occurred: %s' % error


def CreateMessage(sender, to, subject, message_text):
  """Create a message for an email.

  Args:
    sender: Email address of the sender.
    to: Email address of the receiver.
    subject: The subject of the email message.
    message_text: The text of the email message.

  Returns:
    An object containing a base64url encoded email object.
  """
  message = MIMEText(message_text)
  message['to'] = to
  message['from'] = sender
  message['subject'] = subject
  return {'raw': base64.urlsafe_b64encode(message.as_string())}


def CreateMessageWithAttachment(
    sender, to, subject, message_text, file_dir, filename):
  """Create a message for an email.

  Args:
    sender: Email address of the sender.
    to: Email address of the receiver.
    subject: The subject of the email message.
    message_text: The text of the email message.
    file_dir: The directory containing the file to be attached.
    filename: The name of the file to be attached.

  Returns:
    An object containing a base64url encoded email object.
  """
  message = MIMEMultipart()
  message['to'] = to
  message['from'] = sender
  message['subject'] = subject

  msg = MIMEText(message_text)
  message.attach(msg)

  path = os.path.join(file_dir, filename)
  content_type, encoding = mimetypes.guess_type(path)

  if content_type is None or encoding is not None:
    content_type = 'application/octet-stream'
  main_type, sub_type = content_type.split('/', 1)
  if main_type == 'text':
    fp = open(path, 'rb')
    msg = MIMEText(fp.read(), _subtype=sub_type)
    fp.close()
  elif main_type == 'image':
    fp = open(path, 'rb')
    msg = MIMEImage(fp.read(), _subtype=sub_type)
    fp.close()
  elif main_type == 'audio':
    fp = open(path, 'rb')
    msg = MIMEAudio(fp.read(), _subtype=sub_type)
    fp.close()
  else:
    fp = open(path, 'rb')
    msg = MIMEBase(main_type, sub_type)
    msg.set_payload(fp.read())
    fp.close()

  msg.add_header('Content-Disposition', 'attachment', filename=filename)
  message.attach(msg)

  return {'raw': base64.urlsafe_b64encode(message.as_string())}

def CreateMessageWithMultipleAttachmentsAndCC(
    sender, to, subject, message_text, file_dir, cc_addresses):
  """Create a message for an email.

  Args:
    sender: Email address of the sender.
    to: Email address of the receiver.
    subject: The subject of the email message.
    message_text: The text of the email message.
    file_dir: The directory containing the file to be attached.
    filename: The name of the file to be attached.

  Returns:
    An object containing a base64url encoded email object.
  """
  message = MIMEMultipart()
  message['to'] = to.strip()
  message['from'] = sender.strip()
  message['subject'] = subject.strip()
  message['Cc'] = cc_addresses.strip()

  msg = MIMEText(message_text)
  message.attach(msg)

  #iterate all filename under the folder and go through following section.
  filenames = [ f for f in listdir(file_dir) if isfile(join(file_dir,f)) ]

  for filename in filenames:
    print 'Attach file: ' + filename
    path = os.path.join(file_dir, filename)
    content_type, encoding = mimetypes.guess_type(path)

    if content_type is None or encoding is not None:
      content_type = 'application/octet-stream'
    main_type, sub_type = content_type.split('/', 1)
    if main_type == 'text':
      fp = open(path, 'rb')
      msg = MIMEText(fp.read(), _subtype=sub_type)
      fp.close()
    elif main_type == 'image':
      fp = open(path, 'rb')
      msg = MIMEImage(fp.read(), _subtype=sub_type)
      fp.close()
    elif main_type == 'audio':
      fp = open(path, 'rb')
      msg = MIMEAudio(fp.read(), _subtype=sub_type)
      fp.close()
    else:
      fp = open(path, 'rb')
      msg = MIMEBase(main_type, sub_type)
      msg.set_payload(fp.read())
      fp.close()

    msg.add_header('Content-Disposition', 'attachment', filename=filename)
    message.attach(msg)

  return {'raw': base64.urlsafe_b64encode(message.as_string())}


def unicode_csv_reader(utf8_data, dialect=csv.excel, **kwargs):
    csv_reader = csv.reader(utf8_data, dialect=dialect, **kwargs)
    for row in csv_reader:
        yield [unicode(cell, 'utf-8') for cell in row]

def GetMasterName(Contact_line):
  line = Contact_line.split(',')
  return line[0]

def GetPetName(Contact_line):
  line = Contact_line.split(',')
  return line[1]

def GetReceipent(Contact_line):
  line = Contact_line.split(',')
  return line[2]

if __name__=="__main__" :
  # Path to the client_secret.json file downloaded from the Developer Console
  CLIENT_SECRET_FILE = r'C:\Users\jaxon_chen\Desktop\PETogether_send_mail\client_secret.json'

  # Check https://developers.google.com/gmail/api/auth/scopes for all available scopes
  OAUTH_SCOPE = 'https://www.googleapis.com/auth/gmail.compose'

  # Location of the credentials storage file
  STORAGE = Storage('gmail.storage')

  # Start the OAuth flow to retrieve credentials
  flow = flow_from_clientsecrets(CLIENT_SECRET_FILE, scope=OAUTH_SCOPE)
  http = httplib2.Http()

  # Try to retrieve credentials from storage or run the flow to generate them
  credentials = STORAGE.get()
  if credentials is None or credentials.invalid:
    credentials = run(flow, STORAGE, http=http)

  # Authorize the httplib2.Http object with our credentials
  http = credentials.authorize(http)

  # Build the Gmail service from discovery
  gmail_service = build('gmail', 'v1', http=http)

  #Recipients processing.
  MasterName = None
  PETName = None
  MailAddress = None
  MailTitle = '您好，我們是PETogether!'.decode('utf8')  
  cc_addresses = 'petogether.tw@gmail.com'

  reload(sys)

  sys.setdefaultencoding('utf8')
  Contact_Book = None

  with codecs.open(r'C:\Users\jaxon_chen\Desktop\PETogether_send_mail\DataTable.csv','r','utf8') as f:
    Contact_Book = f.readlines()
    for line in Contact_Book:
      print line.decode('utf8').encode('big5')
  
  MailContent = None
  with codecs.open(r'C:\Users\jaxon_chen\Desktop\PETogether_send_mail\Reply_template_ASCII.txt','r') as f:
    text = f.read()
    MailContent = text
  
  #Replace keyword in mail content and send the mail to receipent.
  for line in Contact_Book:
    MasterName = GetMasterName(line.decode('utf8').encode('big5')).strip()
    PETName = GetPetName(line.decode('utf8').encode('big5')).strip()
    receipent = GetReceipent(line.decode('utf8').encode('big5')).strip()
    MailContent_for_sending = MailContent.replace('XXX',MasterName).strip()
    MailContent_for_sending = MailContent_for_sending.replace('OOO',PETName).strip()

    AttachmentPath = r'C:\Users\jaxon_chen\Desktop\PETogether_send_mail\Attachment'

    print os.path.join(AttachmentPath,PETName)
    #print MailTitle.decode('utf8').encode('big5')
    #msg = CreateMessage('petogether.tw@gmail.com',receipent, MailTitle, MailContent_for_sending)
    ##msg = CreateMessageWithMultipleAttachmentsAndCC('petogether.tw@gmail.com', receipent, MailTitle, MailContent_for_sending, r'C:\Users\jaxon_chen\Desktop\PETogether_send_mail\Attachment',cc_addresses)
    msg = CreateMessageWithMultipleAttachmentsAndCC('petogether.tw@gmail.com', receipent, MailTitle, MailContent_for_sending, os.path.join(AttachmentPath,PETName), cc_addresses)
    msg = SendMessage(gmail_service,'me',msg)        
  #print MasterName
  #print PETName
  #print MailAddress
  #print MailContent
  #print MailTitle  


