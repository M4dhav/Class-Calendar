from __future__ import print_function
import asyncio
import datetime
import os.path
import io
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import openpyxl
import streamlit as st
from httpx_oauth.clients.google import GoogleOAuth2
from dotenv import load_dotenv

load_dotenv()

client_id = os.getenv("CLIENT_ID")
client_secret = os.getenv("CLIENT_SECRET")
redirect_uri = os.getenv("REDIRECT_URI")
client = GoogleOAuth2(client_id, client_secret)

def create_event(summary, location, description,dstart,tstart, tend,calid, creds ):
    service = build('calendar', 'v3', credentials=creds)

        # Call the Calendar API
    event = {
            'summary': summary,
            'location': location,
            'description': description,
            'start': {
                'dateTime': '2023-08-'+dstart+'T'+tstart+':30:00+05:30',
                'timeZone': 'Asia/Kolkata',
            },
            'end': {
                'dateTime': '2023-08-'+dstart+'T'+tend+':30:00+05:30',
                'timeZone': 'Asia/Kolkata',
            },
            'recurrence': [
                'RRULE:FREQ=WEEKLY;COUNT=22'
            ],
            'attendees': [
            ],
            'reminders': {
                'useDefault': True,
            },
        }
    event = service.events().insert(calendarId=calid, body=event).execute()
    print('Event created: %s' % (event.get('htmlLink')))

def parse(data,specialisation):
    match specialisation:
        case "AI":
            splcourse = "CSET211"
        case "Blockchain":
            splcourse = "CSET212"
        case "Cyber Security":
            splcourse = "CSET213"
        case "Data Science":
            splcourse = "CSET214"
        case "Gaming":
            splcourse = "CSET215"
        case "Core":
            splcourse = "CSET216"
        case "DevOps":
            splcourse = "CSET217"
        case "Full Stack":
            splcourse = "CSET218"
        case "Quantum Computing":
            splcourse = "CSET219"
        case "Drones":
            splcourse = "CSET220"
        case "Robotics":
            splcourse = "CSET221"
        case "IoT":
            splcourse = "CSET222"
        case "AR/VR":
            splcourse = "CSET223"
        case "Product Design":
            splcourse = "CSET238"
        case "Cloud Computing":
            splcourse = "CSET224"
    ttwb = openpyxl.load_workbook(io.BytesIO(data))
    tt = ttwb.active
    coursenames = [[],[],[],[],[]]
    rooms = [[],[],[],[],[]]
    c = 0
    for i in range(2,7):
        for j in range(5,14):
            value = tt.cell(row = j, column = i).value
            if value == None:
                coursenames[c].append("Free")
                rooms[c].append("Free")
                continue
            i1 = value.index("{")
            i2 = value.index("}")
            room = value[i1+1:i2]
            if "CSET201" in value:
                coursenames[c].append(value)
                rooms[c].append(room)
            elif "CSET202" in value:
                coursenames[c].append(value)
                rooms[c].append(room)
            elif "CSET203" in value:
                coursenames[c].append(value)
                rooms[c].append(room)
            elif "CSET240" in value:
                coursenames[c].append(value)
                rooms[c].append(room)
            elif "CSET205" in value:
                coursenames[c].append(value)
                rooms[c].append(room)
            elif splcourse in value:
                i1 = value.index(splcourse)
                value = value[i1:]
                i2 = value.index("}")
                value = value[:i2+1]
                i3 = value.index("{")
                room = value[i3+1:i2]
                rooms[c].append(room)
                coursenames[c].append(value)
            else:
                coursenames[c].append("Free")
                rooms[c].append("Free")
        c += 1
    return coursenames, rooms

async def write_authorization_url(client,
                                  redirect_uri):
    authorization_url = await client.get_authorization_url(
        redirect_uri,
        scope=["https://www.googleapis.com/auth/calendar"],
        extras_params={"access_type": "offline"},
    )
    return authorization_url
authorization_url = asyncio.run(
    write_authorization_url(client=client,
                            redirect_uri=redirect_uri))

st.title('Timetable Excel Sheet to Google Calendar')
print(st.experimental_get_query_params())
st.session_state.token = None
try:
    st.session_state.token = st.experimental_get_query_params()['code']
    st.write(st.session_state.token)
except:
    pass
st.write(f'''<h1><a target="_self"href="{authorization_url}">LOGIN</a></h1>''',unsafe_allow_html=True)
specialisation = st.selectbox('What is your specialisation?',("AI","Blockchain","Cyber Security","Data Science","Gaming","Core","DevOps","Full Stack","Quantum Computing","Drones","Robotics","IoT","AR/VR","Product Design","Cloud Computing"))
if st.session_state.token:
    uploaded_file = st.file_uploader("Choose a file")


    if uploaded_file is not None:
        bytes_data = uploaded_file.read()
        coursenames, rooms=parse(bytes_data,specialisation)
    

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/calendar']


def main_c():
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

    try:
        calendar = {
        'summary': 'Bennett Sem 3 Timetable',
        'timeZone': 'Asia/Kolkata'
        }
        service = build('calendar', 'v3', credentials=creds)
        created_calendar = service.calendars().insert(body=calendar).execute()
        calid = created_calendar['id']
        for i in range(5):
            for j in range(9):
                value = coursenames[i][j]
                match i:
                    case 0:
                        dstart = '07'
                    case 1:
                        dstart = '01'
                    case 2:
                        dstart = '02'
                    case 3:
                        dstart = '03'
                    case 4:
                        dstart = '04'
                match j:
                    case 0:
                        tstart = '08'
                        tend = '09'
                    case 1:
                        tstart = '09'
                        tend = '10'
                    case 2:
                        tstart = '10'
                        tend = '11'
                    case 3:
                        tstart = '11'
                        tend = '12'
                    case 4:
                        tstart = '12'
                        tend = '13'
                    case 5:
                        tstart = '13'
                        tend = '14'
                    case 6:
                        tstart = '14'
                        tend = '15'
                    case 7:
                        tstart = '15'
                        tend = '16'
                    case 8:
                        tstart = '16'
                        tend = '17'
                location = rooms[i][j]
                description = coursenames[i][j]
                if coursenames[i][j] == "Free":
                        continue
                elif "CSET201" in value:
                    summary = "Information Management Systems "
                elif "CSET202" in value:
                    summary = "Data Structures using C++ "
                elif "CSET203" in value:
                    summary = "Microprocessors and Computer Architecture "
                elif "CSET240" in value:
                    summary = "Probability and Statistics "
                elif "CSET205" in value:
                    summary = "Software Engineering "
                elif "CSET211" in value:
                    summary = "Statistical Machine Learning "
                if "(L)" in value:
                    summary += "Lecture"
                elif "(T)" in value:
                    summary += "Tutorial"
                elif "(P)" in value:
                    summary += "Lab"
                create_event(summary, location, description,dstart,tstart, tend,calid, creds)
                print("Successfull")

    except HttpError as error:
        print('An error occurred: %s' % error)

