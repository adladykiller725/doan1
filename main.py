from __future__ import print_function

import datetime

import httplib2
import os
import csv
import re
from googleapiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# constant
CALENDAR_NAME = 'csv-to-calendar'


# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/calendar-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/calendar'
CLIENT_SECRET_FILE = '.client_secret.json'
APPLICATION_NAME = 'CSV to Calendar'


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
                                   'calendar-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else:  # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


def main():
    """Shows basic usage of the Google Calendar API.

    Creates a Google Calendar API service object and outputs a list of the next
    10 events on the user's calendar.
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('calendar', 'v3', http=http)

    calendar_id = initCalendar(service, CALENDAR_NAME)

    mat = getMatrixFromCSV('timetable2.csv')
    events = parseMatrixIntoEvents(mat)
    for x in events:
        uploadEvent(service, x, calendar_id)


def deleteAllCalendars_NO(service, summary):
    page_token = None
    while True:
        calendar_list = service.calendarList().list(pageToken=page_token).execute()
        for calendar_list_entry in calendar_list['items']:
            if calendar_list_entry['summary'] == summary:
                service.calendars().delete(calendarId=calendar_list_entry['id']).execute()
                # service.calendars().delete('secondaryCalendarId').execute()
                print('Calendar ' + calendar_list_entry['summary'] + ' has been deleted')

        page_token = calendar_list.get('nextPageToken')
        if not page_token:
            break

    return None


def getCalendarId(service, calendarSummary):
    page_token = None
    while True:
        calendar_list = service.calendarList().list(pageToken=page_token).execute()
        for calendar_list_entry in calendar_list['items']:
            if calendar_list_entry['summary'] == calendarSummary:
                return calendar_list_entry['id']
        page_token = calendar_list.get('nextPageToken')
        if not page_token:
            break

    return None


def uploadEvent(service, event, calenderID):
    #print(event)
    print('Creating event... ' + event['summary'] + ' @ ' + event['start']['dateTime'])

    event = service.events().insert(calendarId=calenderID, body=event).execute()
    print('Event created!')


def initCalendar(service, summary):
    calenderID = getCalendarId(service, summary)
    newID = calenderID
    if calenderID != None:
        service.calendars().delete(calendarId=calenderID).execute()
        print('Calendar ' + summary +' has been deleted')

    calendar = {
    'summary': summary,
    'timeZone': 'Asia/Ho_Chi_Minh'
    }

    created_calendar = service.calendars().insert(body=calendar).execute()

    print('Calendar ' + summary +' has been created')
    newID = created_calendar['id']

    return newID


def getColours(service):
    colors = service.colors().get().execute()

    print(colors)
    # Print available calendarListEntry colors.
    for id, color in colors['calendar']:
        print('colorId: %s'.format(id))
        #print'  Background: %s' % color['background']
        #print '  Foreground: %s' % color['foreground']

    # Print available event colors.
    for id, color in colors['event']:
        print('colorId: %s'.format(id))
        #print '  Background: %s' % color['background']
        #print '  Foreground: %s' % color['foreground']


def createEvent(year, weekofyear, date, start, end, subject, location,colour):

    year = int(year)
    date = int(date)
    weekofyear = int(weekofyear)
    firstdayofweek = datetime.datetime.strptime(f'{year}-W{int(weekofyear)}-1', "%Y-W%W-%w").date()
    timestudy = firstdayofweek + datetime.timedelta(days = date - 2)

    startTime = start
    endTime = end

    googleEvent = {
      'summary': subject,
      'location': location,
      'start': {
        'dateTime': str(timestudy) + 'T' + str(startTime) + ':00+07:00',
        'timeZone': 'Asia/Ho_Chi_Minh'
      },
      'end': {
        'dateTime': str(timestudy) + 'T' + str(endTime) + ':00+07:00',
        'timeZone': 'Asia/Ho_Chi_Minh'
      },
      'recurrence': [
        'RRULE:FREQ=WEEKLY;COUNT=1'
      ],
      'reminders': {
        'useDefault': False,
        'overrides': [
          {'method': 'popup', 'minutes': 15}
        ],
      },
      'colorId': str(colour)
    }

    return googleEvent


def parseMatrixIntoEvents(mat):



    DATA_START_ROW = 3
    DATA_START_COL = 1
    HEADER_ROW = 2

    eventList = []

    rowCount = 0

    year1 = ''.join(mat[0][0])
    year2 = re.findall(r'\d+', year1)
    semester = year2[0]
    year3 = year2[1]
    emptymatrix = []
    colour1 = 3

    for row in mat:
        colCount = 0

        for data in row:

            if rowCount >= DATA_START_ROW:
                if colCount >= DATA_START_COL:
                 if data != '':
                     if mat[HEADER_ROW][colCount] == 'TEN MON HOC':
                        # Merge entries below

                        content1 = mat[] + mat[rowCount][colCount-1] + ' - ' +  mat[rowCount][colCount] + ' - ' + mat[rowCount][colCount + 3]
                        content2 = ''.join(content1)

                        location1 = mat[rowCount][colCount + 7] + ' - ' + mat[rowCount][colCount + 8]
                        location2 = ''.join(location1)

                        timeweek1 = ''.join(mat[rowCount][colCount + 9])
                        timeweek2 = re.findall(r'\d+',timeweek1)

                        timestart = ''.join(mat[rowCount][colCount + 6]).split(' - ')[0]
                        timeend = ''.join(mat[rowCount][colCount + 6]).split(' - ')[1]
                        if mat[rowCount][colCount + 4] != '--':
                            if timeweek2 != emptymatrix:
                                for data2 in timeweek2:
                                    if semester != '1':
                                        year3 = year2[2]
                                    else:
                                        if int(data2) <= 10:
                                            year3 = year2[2]
                                        else:
                                            year3 = year2[1]
                                    tmpEntryDict = createEvent(year3, data2,
                                    mat[rowCount][colCount + 4], timestart,
                                    timeend, content2,
                                    location2,colour1)

                                    eventList.append(tmpEntryDict)
            colCount += 1

        rowCount += 1



    return eventList


def getMatrixFromCSV(csvFile):

    rowCount = 0

    matrix = []

    with open('timetable4.csv', newline='') as f:
        reader = csv.reader(f)
        for row in reader:
            colCount = 0

            if rowCount >= len(matrix):
                #print('Adding row to matrix')
                matrix.append([])

            for col in row:
                matrix[rowCount].append(col)
                #print('[' + str(rowCount) + '][' + str(colCount) + '] ' + col)
                colCount += 1
            rowCount += 1

    return matrix


if __name__ == '__main__':
    main()

