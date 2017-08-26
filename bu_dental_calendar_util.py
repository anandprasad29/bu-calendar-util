from __future__ import print_function
import httplib2
import os
import re

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

import base64
import datetime
import logging
import time
from apiclient import errors

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

logging.basicConfig(filename='/Users/anandprasad/Dropbox (Personal)/logs/bu-calendar-util.log',
                    format='%(asctime)s %(filename)s %(funcName)s %(lineno)s %(message)s',
                    datefmt='%m/%d/%Y %I:%M:%S %p',
                    level=logging.INFO)

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/bu-calendar-util.json
SCOPES = 'https://www.googleapis.com/auth/gmail.modify https://www.googleapis.com/auth/calendar'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'BU Dental Calendar Util'

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
                                   'bu-calendar-util.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        logging.info('Storing credentials to ' + credential_path)
    return credentials

def get_message(service, user_id, msg_id):
  """Get a Message with given ID.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    msg_id: The ID of the Message required.

  Returns:
    A Message.
  """
  try:
    message = service.users().messages().get(userId=user_id, id=msg_id).execute()
    return message
  except errors.HttpError, error:
    logging.error('An error occurred: %s' % error)

def addProcessedLabel(mail_service, msg_id):
  """Add Processed Appt. label to the given Message.

  Args:
    service: Authorized Gmail API service instance.
    msg_id: The id of the message required.
  """
  msg_labels = {'removeLabelIds': [], 'addLabelIds': ['Label_1']}

  try:
    mail_service.users().messages().modify(userId='me', id=msg_id,
                                           body=msg_labels).execute()
  except errors.HttpError, error:
    logging.error('An error occurred: %s' % error)

def get_messages(mail_service, message_ids):
  messages = []
  for message_data in message_ids:
    msg = get_message(mail_service, 'me', message_data['id'])
    messages.append(msg)
  return messages


def list_messages_matching_query(service, user_id, query=''):
  """List all Messages of the user's mailbox matching the query.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    query: String used to filter messages returned.
    Eg.- 'from:user@some_domain.com' for Messages from a particular sender.

  Returns:
    List of Messages that match the criteria of the query. Note that the
    returned list contains Message IDs, you must use get with the
    appropriate ID to get the details of a Message.
  """
  try:
    response = service.users().messages().list(userId=user_id,
                                               q=query).execute()
    messages = []
    if 'messages' in response:
      messages.extend(response['messages'])

    while 'nextPageToken' in response:
      page_token = response['nextPageToken']
      response = service.users().messages().list(userId=user_id, q=query,
                                         pageToken=page_token).execute()
      messages.extend(response['messages'])

    return messages
  except errors.HttpError, error:
    logging.error('An error occurred: %s' % error)

def generate_start_end_time(text):
    """Generate start, end time objects after parsing text.

    Args:
      text: Text to parse.

    Returns:
      Tuple of 2 datetime objects representing the startTime and endTime.
    """
    regex = '(\d{2})/(\d{2})/(\d{4}).*(\d{2}):(\d{2})\s-\s(\d{2}):(\d{2})'
    match = re.search(regex, text, re.DOTALL)
    if match is not None:
      mm =  int(match.group(1))
      dd =  int(match.group(2))
      yyyy =  int(match.group(3))
      start_hh =  int(match.group(4))
      start_mm =  int(match.group(5))
      end_hh =  int(match.group(6))
      end_mm =  int(match.group(7))
      startTime = datetime.datetime(yyyy, mm, dd, start_hh, start_mm)
      endTime = datetime.datetime(yyyy, mm, dd, end_hh, end_mm)
      return (startTime, endTime)

def generate_per_line_start_end_time(body_text):
  """Loop through each line of the text and extract datetime info.

  Args:
    text: Text to parse.

  Returns:
    List of tuples, each consisting of 2 datetime objects representing the
    startTime and endTime.
  """
  lines = body_text.split("\n")
  list_times = []
  for line in lines:
    date_time_tuple = generate_start_end_time(line)
    if date_time_tuple is not None:
      list_times.append(date_time_tuple)
  return list_times

def generate_per_body_start_end_time(body_text):
  list_times = []
  list_times.append(generate_start_end_time(body_text))
  return list_times

def create_event(body_text, summary, startTime, endTime):
  event = {}
  event['summary'] = summary
  event['description'] = body_text
  event['start'] = {
    'dateTime': startTime.isoformat("T"),
    'timeZone': 'America/New_York',
  }
  event['end'] = {
    'dateTime': endTime.isoformat("T"),
    'timeZone': 'America/New_York',
  }
  return event

def insert_event_into_calendar(cal_service, body_text, summary,
                               startTime, endTime):
  event = create_event(body_text, summary, startTime, endTime)
  gen_event = cal_service.events().insert(calendarId='primary', body=event).execute()
  logging.info('Event created: %s' % (gen_event.get('htmlLink')))
  logging.info('Event description: %s' % (body_text))

def insert_unique_event_into_calendar(cal_service, body_text, summary,
                                      list_times):
  should_add = False
  for (startTime, endTime) in list_times:
    matched_events = list_matching_cal_events(cal_service, startTime, endTime)
    if matched_events:
      for matched_event in matched_events:
        if "Appointment" in matched_event['summary']:
          logging.info("Appointment matching this date and time already exists"
                       ", ignoring this e-mail:%s", body_text)
        else:
          # The matched events don't have the same summary, so might be
          # of a different type.
          should_add = True
    else:
      # No matched events, add an appointment
      should_add = True
    if should_add:
      insert_event_into_calendar(cal_service, body_text, "Appointment",
                                 startTime, endTime)

def add_event(cal_service, body_text, subject):
  list_times = []
  if "New Salud Alert" in subject:
    list_times = generate_per_line_start_end_time(body_text)
    if not list_times:
      logging.error("Failed to determine Regular Appointment info from " + body_text)
      return
    summary = 'Appointment'
  else:
    list_times = generate_per_body_start_end_time(body_text)
    if not list_times:
      logging.error("Failed to determine custom Appointment info from " + body_text)
      return
    summary = subject[:-7]+" Appointment"

  for (startTime, endTime) in list_times:
    insert_unique_event_into_calendar(cal_service, body_text, subject,
                                      list_times)

def list_matching_cal_events(cal_service, startTime, endTime):
  local_time = time.localtime()
  if local_time.tm_isdst == 1:
    EST_offset = "-04:00"
  else:
    EST_offset = "-05:00"
  startTimeStr = startTime.isoformat("T")+EST_offset
  endTimeStr = endTime.isoformat("T")+EST_offset

  matched_events = []
  page_token = None
  while True:
    events = cal_service.events().list(calendarId='primary', pageToken=page_token,
                                       timeMin=startTimeStr,
                                       timeMax=endTimeStr).execute()
    for event in events['items']:
      if (event['start']['dateTime'] == startTimeStr and
          event['end']['dateTime'] == endTimeStr):
        matched_events.append(event)

    page_token = events.get('nextPageToken')
    if not page_token:
      break
  return matched_events

def update_event(cal_service, body_text, subject):
    list_times = generate_per_line_start_end_time(body_text)
    if list_times is None:
      logging.error("Failed to determine Appointment info")
      return
    insert_unique_event_into_calendar(cal_service, body_text, subject, list_times)


def cancel_event(cal_service, body_text):
  list_times = generate_per_line_start_end_time(body_text)
  if list_times is None:
    logging.error("Failed to determine Appointment info")
    return

  for (startTime, endTime) in list_times:
    matched_events = list_matching_cal_events(cal_service, startTime, endTime)
    for matched_event in matched_events:
      logging.info("Cancelling event with description: %s",
                   matched_event['description'])
      cal_service.events().delete(calendarId='primary', eventId=matched_event['id']).execute()


def create_calendar_event(cal_service, subject, body_text):
  """ Extract Appointment information from message and update the calendar.
  """
  if "added" in body_text or "successful booking" in body_text:
      add_event(cal_service, body_text, subject)
  if "updated" in body_text:
      update_event(cal_service, body_text, subject)
  elif "cancelled" in body_text:
      cancel_event(cal_service, body_text)

def main():
  credentials = get_credentials()
  http = credentials.authorize(httplib2.Http())
  mail_service = discovery.build('gmail', 'v1', http=http)
  cal_service = discovery.build('calendar', 'v3', http=http)

  query_str = 'from:jgt@bu.edu NOT label:Processed-Appt.'
  message_ids = list_messages_matching_query(mail_service, 'me', query_str)
  messages = get_messages(mail_service, message_ids)
  messages.sort(key=lambda x: int(x['internalDate']))

  for msg in messages:
    headers = msg['payload']['headers']
    for header in headers:
      if ("Subject" in header["name"]):
        subject = header["value"]
    body = base64.urlsafe_b64decode(msg['payload']['body']['data'].encode('utf-8'))
    create_calendar_event(cal_service, subject, body)
    addProcessedLabel(mail_service, msg['id'])


if __name__ == '__main__':
    main()
