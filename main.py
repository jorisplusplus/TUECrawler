"""A quickstart example showing usage of the Google Calendar API."""
import datetime
import os

from apiclient.discovery import build
from httplib2 import Http
import oauth2client
from oauth2client import client
from oauth2client import tools
from pushbullet import PushBullet
from BeautifulSoup import BeautifulSoup
import mechanize
import cookielib
import time
import csv
import dateutil.tz

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

SCOPES = 'https://www.googleapis.com/auth/calendar'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Calendar API Quickstart'

Username = 'TUEuser'
Password = 'TUEpass'
ApiKey = 'PushAPI'

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
                                   'calendar-api-quickstart.json')

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatability with Python 2.6
            credentials = tools.run(flow, store)
        print 'Storing credentials to ' + credential_path
    return credentials


def main():
    """Shows basic usage of the Google Calendar API.

    Creates a Google Calendar API service object and outputs a list of the next
    10 events on the user's calendar.
    """
    start_time = time.time()
    br = mechanize.Browser()
    cj = cookielib.LWPCookieJar()
    br.set_cookiejar(cj)

    p = PushBullet(ApiKey)
    devices = p.getDevices()
    p.pushNote(devices[0]["iden"], 'Rooster', 'Starting rooster import')

    localtz = dateutil.tz.tzlocal()
    localoffset = localtz.utcoffset(datetime.datetime.now(localtz))
    hours = localoffset.total_seconds()/3600
    print 'utcoffset ',hours

    credentials = get_credentials()
    service = build('calendar', 'v3', http=credentials.authorize(Http()))
    page_token = None
    roosterID = None
    while True:
        calendar_list = service.calendarList().list(pageToken=page_token).execute()
        for calendar_list_entry in calendar_list['items']:
            if calendar_list_entry['summary'] == 'Rooster':
		        roosterID = calendar_list_entry['id']
        page_token = calendar_list.get('nextPageToken')
        if not page_token:
            break
    if roosterID == None:
        calendar = {
            'summary': 'Rooster',
            'timeZone': 'Europe/Amsterdam'
        }
        created_calendar = service.calendars().insert(body=calendar).execute()
        roosterID = created_calendar.get('id')
    else:
        print 'Clearing old rooster: ',roosterID
        page_token = None
        while True:
            events = service.events().list(calendarId=roosterID, pageToken=page_token).execute()
            eventList = events.get('items',[])
            event = None
            for event in eventList:
                eventID = event.get('id')
                if eventID:
                    print 'Delete event: ',eventID
                    service.events().delete(calendarId=roosterID, eventId=eventID).execute()
                    time.sleep(1)
            page_token = events.get('nextPageToken')
            if not page_token:
                break

    br.set_handle_robots(False)
    br.addheaders = [('User-agent', 'Mozilla/5.0 (X11; U; Linux i686; en-US; rv:1.9.0.1) Gecko/2008071615 Fedora/3.0.1-1.fc9 Firefox/3.0.1')]
    br.open('https://onderwijssso.tue.nl/CookieAuth.dll?GetLogon?curl=Z2F&reason=0&formdir=3')
    br.select_form(nr = 0)
    br.form['username'] = Username
    br.form['password'] = Password
    time.sleep(1)
    response = br.submit()
    print response
    br.open('https://onderwijssso.tue.nl/Activiteiten/Pages/TimeTable.aspx?mode=kwartiel')
    time.sleep(1)
    #print br.response().read()
    for form in br.forms():
        if form.attrs['id'] == 'aspnetForm':
            br.form = form
    br.select_form('aspnetForm')
    button = None
    for control in br.form.controls:
        if br[control.name] == ['ICS']:
            br[control.name] = ['CSV']
        #if not control.type == 'hidden':
            #print "type=%s, name=%s value=%s" % (control.type, control.name, br[control.name])
        if 'Export' in control.name:
            button = control.name
    response = br.submit(name=button)
    print 'export response: ',response
    csvURL = None
    for line in response.read().splitlines():
        if 'ExportToOutlook' in line:
            print "Found URL line: ",line
            soup = BeautifulSoup(line)
            csvURL = soup.findAll('iframe')[0]['src']
            print csvURL
    if csvURL:
        br.open('https://onderwijssso.tue.nl/'+csvURL)
        result = [row for row in csv.reader(br.response().read().splitlines(), delimiter=',')]
        rownum = 0
        for row in result:
            if not rownum == 0:
                print row
                event = {
                    'summary': row[5],
                    'location': row[4],
                    'start': {
                        'dateTime': row[0].strip()+'T'+row[2].strip()+':00.000+0'+str("{0:.0f}".format(hours))+':00'
                    },
                    'end': {
                        'dateTime': row[1].strip()+'T'+row[3].strip()+':00.000+0'+str("{0:.0f}".format(hours))+':00'
                    },
                    }
                created_event = service.events().insert(calendarId=roosterID, body=event).execute()
            rownum += 1
            time.sleep(1)
        print 'Finished saved ',rownum,' events'
        dur = "{:.1f}".format(time.time() - start_time)
        print("--- %s seconds ---" % (time.time() - start_time))
        p.pushNote(devices[0]["iden"], 'Rooster', 'Job Succesfull. Time '+dur)
    else:
        p.pushNote(devices[0]["iden"], 'Rooster', 'Job Failed')
        print 'Job failed'

if __name__ == '__main__':
    main()
