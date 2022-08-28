#!/usr/bin/env python3

import datetime
import sys
import os.path
import argparse
import re
import logging
import smtplib
import logging.config
from datetime import datetime
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from dateutil.relativedelta import relativedelta
from email.message import EmailMessage

LOG_FILE = f'/var/log/{os.path.basename(__file__).split(".")[0]}.log'
# LOG_FILE = f'{os.path.abspath(os.path.dirname(__file__))}/logs/{os.path.basename(__file__).split(".")[0]}.log'
SENDER_MAIL = '<MAIL_SENDER>'
SENDER_PASSWORD = '<SMTP_PASS>
SMTP_SERVER = '<SMTP_SERVER>'
SMTP_PORT = <SMTP_PORT>
COMPANY_DOMAIN = '<DOMAIN_NAME>'
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']  # IF YOU MODIFY THE SCOPES DELETE token.json.

logging.basicConfig(filename=LOG_FILE, format='%(asctime)s | %(name)s | %(levelname)s | %(message)s', datefmt='%Y-%m-%d %H:%M:%S', level=logging.DEBUG)


def monitor_upcoming_events(service, calendar_id, time_min, time_max, max_results, mail_recipient):
    events_result = service.events().list(  # call calendar api
        calendarId=calendar_id,
        timeMin=time_min.isoformat() + 'Z',  # 'Z' indicates UTC time
        timeMax=time_max.isoformat() + 'Z',
        maxResults=max_results,
        singleEvents=True,
        orderBy='startTime').execute()
    calendar = service.calendars().get(calendarId=calendar_id).execute()
    events = events_result.get('items', [])
    if not events:
        logging.info(f'not found upcoming events in "{calendar["summary"]}" for {time_min.strftime("%Y-%m-%d %H:%M")} - {time_max.strftime("%Y-%m-%d %H:%M")}')
        return
    meetings_list = []
    for event in events:
        start_date = datetime.strptime(event['start'].get('dateTime'), '%Y-%m-%dT%H:%M:%S%z')  # start date of meeting
        end_date = datetime.strptime(event["end"].get('dateTime'), '%Y-%m-%dT%H:%M:%S%z')  # end date of meeting
        summary = event['summary']  # title of meeting
        meeting_link = event['htmlLink']  # link to meeting
        attendees = [event['attendees']][0]  # all participants at the meeting
        if not employee_on_meeting(attendees):  # add to list if only organizer at the meeting or no redge employee
            logging.debug(f'"{summary}" meeting append to list')
            meetings_list.append(f'{summary}\n\t\t{start_date.strftime("%m-%d %H:%M")} - {end_date.strftime("%m-%d %H:%M")}\n\t\t{meeting_link}')  # saves meetings without attendees or when there is only inviting bot
    if any(meetings_list):
        send_mail(meetings_list, mail_recipient, calendar['summary'], datetime.now().strftime('%Y-%m-%d'))
        logging.debug(f'reminder sent to {mail_recipient}, found {len(meetings_list)} meeting/s without attendees in "{calendar["summary"]}" for {time_min.strftime("%Y-%m-%d %H:%M")} - {time_max.strftime("%Y-%m-%d %H:%M")}')
    else:
        logging.info(f'found {len(events)} upcoming event/s in "{calendar["summary"]}" for {time_min.strftime("%Y-%m-%d %H:%M")} - {time_max.strftime("%Y-%m-%d %H:%M")}, attendees list is correct, reminder not sent')


def send_mail(meetings, mail_recipient, calendar, now):
    body_meetings = ''
    for meeting in meetings:
        body_meetings += f'\t{meeting} \n'
    message = EmailMessage()
    message['Subject'] = f'{calendar} meetings information {now}'
    message['From'] = SENDER_MAIL
    message['To'] = mail_recipient
    message.set_content(f'List of meetings that no one from Redge is signed up: \n{body_meetings}')
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(SENDER_MAIL, SENDER_PASSWORD)
        server.send_message(message)


def employee_on_meeting(attendees):
    for attendee in attendees:
        if f'@{COMPANY_DOMAIN}' in attendee['email']:
            return True
    return False


def get_api_service():
    credentials = None
    if os.path.exists(f'{os.path.abspath(os.path.dirname(__file__))}/config/token.json'):
        credentials = Credentials.from_authorized_user_file(f'{os.path.abspath(os.path.dirname(__file__))}/config/token.json', SCOPES)
    if not credentials or not credentials.valid:  # if there are no (valid) credentials available, user will log in
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(f'{os.path.abspath(os.path.dirname(__file__))}/config/credentials.json', SCOPES)
            credentials = flow.run_local_server(port=0)
        with open(f'{os.path.abspath(os.path.dirname(__file__))}/config/token.json', 'w') as token:  # save credentials for the next run
            token.write(credentials.to_json())
    try:
        return build('calendar', 'v3', credentials=credentials)
    except HttpError as error:
        logging.error(f'An error occurred: {error}')
        print(f'An error occurred: {error}')
        sys.exit(0)


def get_start_of_day():
    return datetime.now() + relativedelta(hour=0, minute=0)


def get_end_of_day():
    return datetime.now() + relativedelta(hour=23, minute=59)


def valid_date(date):
    regex_date_pattern = r'^(\d{4}-\d{2}-\d{2})'
    if isinstance(date, str) and bool(re.match(regex_date_pattern, date)):
        new_date = date.split('-')
        return datetime(int(new_date[0]), int(new_date[1]), int(new_date[2]))
    else:
        raise argparse.ArgumentTypeError("not a valid date: {0!r}".format(date))


def parse_args():
    arg_parser = argparse.ArgumentParser(description='check if only organizer is signed up at the meeting and if there is no company employee, if so - send email reminder')
    arg_parser.add_argument('-c', '--calendarId', help='', type=str, required=True)
    arg_parser.add_argument('-b', '--beginDate', help='', type=valid_date, default=get_start_of_day())
    arg_parser.add_argument('-e', '--endDate', help='', type=valid_date, default=get_end_of_day())
    arg_parser.add_argument('-m', '--maxResults', help='max item amount in meeting list', type=int, default=10)
    arg_parser.add_argument('-M', '--mailRecipient', help='Reminder recipient', type=str, required=True)
    return arg_parser.parse_args()


if __name__ == '__main__':
    args = parse_args()
    api_service = get_api_service()
    monitor_upcoming_events(api_service, args.calendarId, args.beginDate, args.endDate, args.maxResults, args.mailRecipient)
