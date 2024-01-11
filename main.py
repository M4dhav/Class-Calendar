from __future__ import print_function
import threading
import tkinter as tk
import os
from tkinter import HORIZONTAL, filedialog
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import openpyxl
from tkinter import Label,messagebox,ttk
from xls2xlsx import XLS2XLSX
import time
import env

client_id = os.environ.get('CLIENT_ID')
client_secret = os.environ.get('CLIENT_SECRET')
redirect_uri = ['urn:ietf:wg:oauth:2.0:oob',"http://localhost"]
project_id = "timetable-to-cal"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"

global t1
global splcourses
global maincourses
global eleccourses
splcourses = {"AI":["CSET301", "Artificial Intelligence and Machine Learning"], "Cloud Computing":["CSET232", "Design of Cloud Architectural Solutions"] }
maincourses = {"CSET206":"Design and Analysis of Algorithms", "CSET207":"Computer Networks", "CSET208":"Ethics for Engineers, Patents, Copyrights and IPR",\
    "CSET209":"Operating Systems","CSET210":"Design Thinking & Innovation"}
def create_event(summary, location, description,dstart,tstart, tend,calid, creds ):
    service = build('calendar', 'v3', credentials=creds)

        # Call the Calendar API
    event = {
            'summary': summary,
            'location': location,
            'description': description,
            'start': {
                'dateTime': '2024-01-'+dstart+'T'+tstart+':30:00+05:30',
                'timeZone': 'Asia/Kolkata',
            },
            'end': {
                'dateTime': '2024-01-'+dstart+'T'+tend+':30:00+05:30',
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

def parse(specialisation,wb):
    try:
        splcourse = splcourses[specialisation]
    except KeyError:
        messagebox.showerror('Missing Resource', 'Error: The data for this specialisation is not updated yet, please wait for a future version')
    ttwb = wb
    tt = ttwb.active
    coursenames = [[],[],[],[],[]]
    rooms = [[],[],[],[],[]]
    c = 0
    for i in range(2,7):
        for j in range(5,14):
            value = tt.cell(row = j, column = i).value
            if value == None or value == "":
                coursenames[c].append("Free")
                rooms[c].append("Free")
                continue
            i1 = value.index("{")
            i2 = value.index("}")
            room = value[i1+1:i2]
            if any(k in value for k in list(maincourses.keys())):
                coursenames[c].append(value)
                rooms[c].append(room)
            elif splcourse[0] in value:
                i1 = value.index(splcourse[0])
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
    try:
        t1 = threading.Thread(target=Connection_API, args=(coursenames,rooms,splcourse))
        t1.start()
    except TimeoutError:
        messagebox.showerror('Authentication Timed Out', 'Error: Please try using a different internet connection')

def convert(file,specialisation):
    x2x = XLS2XLSX(file)
    parse(specialisation,x2x.to_xlsx())
    

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/calendar']


def Connection_API(coursenames, rooms, splcourse):
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            client_config = {
        "installed": {
            "client_id": client_id,
            "project_id": project_id,
            "auth_uri": auth_uri,
            "token_uri": token_uri,
            "auth_provider_x509_cert_url": auth_provider_x509_cert_url,
            "client_secret": client_secret,
            "redirect_uris": redirect_uri
            
        }
    }
            flow = InstalledAppFlow.from_client_config(client_config,SCOPES)
            creds = flow.run_local_server(port=0)

    try:
        calendar = {
        'summary': 'Bennett Sem 4 Timetable',
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
                        dstart = '15'
                    case 1:
                        dstart = '16'
                    case 2:
                        dstart = '17'
                    case 3:
                        dstart = '11'
                    case 4:
                        dstart = '12'
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
                progBar['value'] += 2.22222222222
                if coursenames[i][j] == "Free":
                        continue
                else:
                    for k in list(maincourses.keys()):
                        if k in value:
                            summary = maincourses[k]
                            continue
                    if splcourse[0] in value:
                        summary = splcourse[1]
                if "(L)" in value:
                    summary += " Lecture"
                elif "(T)" in value:
                    summary += " Tutorial"
                elif "(P)" in value:
                    summary += " Lab"
                create_event(summary, location, description,dstart,tstart, tend,calid, creds)
        messagebox.showinfo("Completed", "Your Google Calendar has been updated with all the classes")
        progBar["value"] = 0
        time.sleep(2)
        root.quit()
        root.destroy()
        
    except HttpError as error:
        print('An error occurred: %s' % error)
        
def add_file(text = "Please upload the timetable file downloaded from the mail in xls format", types = (("xls files","*.xls*"),("Allfiles","*.*"))):
    specialisation = value_inside.get()
    if specialisation not in options_list:
        messagebox.showerror('Specialisation not chosen', 'Error: Please choose a specialisation')
    else:
        d = filedialog.askopenfilename(initialdir="%userprofile%\downloads",title=text,filetypes = types)
        try:
            convert(d,specialisation)
        except ValueError:
            messagebox.showerror('File Format Error', 'Please choose a valid .XLS file')
            
            
root = tk.Tk()
root.geometry("350x350")
root.maxsize(350,350)
root.minsize(350,350)
root.title("Class Calendar")
l = Label(root, text = '''Made by Madhav Gupta''')
toppad = Label(root, text = '''
Class Calendar
''')
medpad = Label(root, text='''
Please be patient if it appears to be stuck''')
botpad = Label(root, text = '''
               ''')
l.config(font =("Courier", 12))
toppad.config(font =("Monospace", 20))
options_list = ["AI","Blockchain","Cyber Security","Data Science","Gaming","Core","DevOps","Full Stack","Quantum Computing","Drones","Robotics","IoT","AR/VR","Product Design","Cloud Computing"]
value_inside = tk.StringVar(root)
value_inside.set("Choose your specialization")
question_menu = tk.OptionMenu(root, value_inside, *options_list)
b = tk.Button(root, text = "Start", command=add_file)
progBar = ttk.Progressbar(root,orient=HORIZONTAL, length=300,mode="determinate")
toppad.pack()
question_menu.pack()
b.pack()
medpad.pack()
progBar.pack()
botpad.pack()
l.pack()
root.mainloop()