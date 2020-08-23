from __future__ import print_function

import argparse
from collections import defaultdict
import itertools
import os
import pickle

import arrow
from docxtpl import DocxTemplate, RichText
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pandas as pd


DEFAULT_TEMPLATE_PATH = 'template.docx'
DEFAULT_OUTPUT_DIR = 'newsletters'
DEFAULT_START_NUM = 1

GOOGLE_CAL_SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']
SPORTS_CALENDAR_ID = '342i46upg047nehs8t089f7b1c@group.calendar.google.com'
EVENTS_CALENDAR_ID = 'quen.mcr@gmail.com'

parser = argparse.ArgumentParser()
parser.add_argument('bulletin_csv')
parser.add_argument('--as-of', type=lambda d: arrow.Arrow.strptime(d, '%Y%m%d'), help="Generate the newsletter as of this date")
parser.add_argument('--disable-events', action='store_true', help="Use this flag to disable pulling the events section")
parser.add_argument('--disable-sports', action='store_true', help="Use this flag to disable pulling the sports section")
parser.add_argument('--start_num', default=DEFAULT_START_NUM, type=int, help="Override the start number of the newsletter headings")
parser.add_argument('--template-path', default=DEFAULT_TEMPLATE_PATH, help="Override the path to the Microsoft Word template")
parser.add_argument('--output-dir', default=DEFAULT_OUTPUT_DIR, help="Dir to write newsletters")
parser.add_argument('--google-credentials', default='./credentials.json', help="Path to Google credentials file")
args = parser.parse_args()


college = []
cambridge = []
jobs = []


def normalised_begin(event):
    """Return the start time of an event."""
    return arrow.get(event['start']['dateTime'])


def authenticate_google_cal(creds_path):
    """Ask the user to log into Google Cal and save access tokens on disk."""
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
            flow = InstalledAppFlow.from_client_secrets_file(creds_path, SCOPES)
            creds = flow.run_local_server(port=0)

        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    return creds


def events(calendar_id, creds):
    """Fetch events for a given calendar id."""
    if args.as_of:
        now = args.as_of
    else:
        now = arrow.utcnow()

    monday_start = now.shift(weekday=0).floor('day')
    sunday_end = monday_start.shift(weeks=+1).shift(minutes=-1)

    service = build('calendar', 'v3', credentials=creds)
    events_this_week = service.events().list(
        calendarId=calendar_id,
        timeMin=monday_start.datetime.isoformat(),
        timeMax=sunday_end.datetime.isoformat(),
        singleEvents=True,
        orderBy='startTime').execute().get('items', [])

    grouped = defaultdict(list)
    for e in events_this_week:
        day_start = normalised_begin(e).floor('day')
        grouped[day_start].append(e)

    events_context = []
    for k, v in grouped.items():
        events_in_day = []
        for e in v:
            events_in_day.append({
                'name': e['summary'],
                'location': e.get('location', ''),
                'when': normalised_begin(e).format('HH:mm')
            })

        events_context.append({
            'day': k.format('dddd Do MMMM'),
            'events': events_in_day
        })

    return events_context



def bulletins_dataframe(path):
    """Read bulletin entries from CSV and return a dataframe."""
    bulletins = pd.read_csv(args.bulletin_csv, converters={'Approved': lambda x: x == 'TRUE', 'Sent': lambda x: x == 'TRUE'})
    bulletins['"Email contact" address'] = bulletins['"Email contact" address'].fillna('')
    bulletins['"Further information" link'] = bulletins['"Further information" link'].fillna('')
    bulletins['"Facebook" link'] = bulletins['"Facebook" link'].fillna('')
    bulletins['"Apply now" link'] = bulletins['"Apply now" link'].fillna('')

    to_include = bulletins[bulletins['Approved'] & ~ bulletins['Sent']]

    return to_include


def bulletin_to_template_entry(bulletin):
    """Transform a bulletin item into a renderable template entry."""
    entry = {}

    for k, v in bulletin.items():
        entry[k.replace(' ', '_').replace('"', '')] = v

    # Title-case the title
    entry['Title'] = entry['Title'].title()

    # Add contact email
    email_link = 'mailto:{}'.format(entry['Email_contact_address'])

    # Add links
    rt = RichText(entry['Email_contact_address'], url_id=doc.build_url_id(email_link))
    entry['Email_link_rt'] = rt

    rt = RichText('Apply now', url_id=doc.build_url_id(entry['Apply_now_link']))
    entry['Apply_now_link_rt'] = rt

    rt = RichText('Further information', url_id=doc.build_url_id(entry['Further_information_link']))
    entry['Further_information_link_rt'] = rt

    rt = RichText('Facebook', url_id=doc.build_url_id(entry['Facebook_link']))
    entry['Facebook_link_rt'] = rt

    return entry


def filter_entries(entries, section):
    """Filter a list of entries and only return those in the matching section."""
    return [e for e in entries if e['Section'].strip() == section]


def create_newsletter(template_path, college, cambridge, jobs):
    """Open the template and populate with content."""
    return DocxTemplate(template_path).render({
        'college_entries': college,
        'cambridge_entries': cambridge,
        'job_entries': jobs
    })


def save_newsletter(doc, output_dir):
    """Write the newsletter to disk."""
    current_date = arrow.utcnow().datetime
    output_filename = 'newsletter-{}.docx'.format(current_date.isoformat()) 
    output_path = os.path.join(output_dir, output_filename)

    doc.save(output_path)

    return output_path


def open_word(path):
    os.system('open {}'.format(output_path))


if __name__ == '__main__':
    # Read the bulletins CSV and add the newsletter entries
    bulletins = bulletins_dataframe(args.bulletin_csv)
    entries = [bulletin_to_template_entry(b) for b in bulletins.to_dict(orient='records')]

    # Separate the entries into the three newsletter sections
    college = filter_entries(entries, 'College')
    cambridge = filter_entries(entries, 'Cambridge')
    jobs = filter_entries(entries, 'Jobs')

    # Prepend empty bulletin entries to the start of the newsletter
    # if the start number has been overriden
    college_prepend = []
    for i in range(args.start_num - 1):
        college_prepend.append({})

    # Prepend the calendar sections: one for events, one for sports
    creds = authenticate_google_cal(args.google_credentials)

    if not args.disable_events:
        college_prepend.append({
            'Title': 'MCR Events Bulletin',
            'events': events(EVENTS_CALENDAR_ID, creds),
        })

    if not args.disable_sports:
        college_prepend.append({
            'Title': 'MCR Sports Bulletin',
            'events': events(SPORTS_CALENDAR_ID, creds),
        })

    college = list(itertools.chain(college_prepend, college))

    # Chain the sections together and update the numbering accordingly
    for i, entry in enumerate(itertools.chain(college, cambridge, jobs)):
        entry['number'] = i+1

    # Create, save and open the newsletter
    newsletter = create_newsletter(template_path, college, cambridge, jobs)
    path = save(newsletter, output_dir)
    open_word(path)
