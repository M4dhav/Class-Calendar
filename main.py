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

client_id = os.environ.get('CLIENT_ID')
client_secret = os.environ.get('CLIENT_SECRET')
redirect_uri = ['urn:ietf:wg:oauth:2.0:oob',"http://localhost"]
project_id = "timetable-to-cal"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"

global t1

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

def parse(specialisation,wb):
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
    ttwb = wb
    tt = ttwb.active
    coursenames = [[],[],[],[],[]]
    rooms = [[],[],[],[],[]]
    c = 0
    for i in range(2,7):
        for j in range(5,14):
            value = tt.cell(row = j, column = i).value
            print(value)
            if value == None or value == "":
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
    try:
        t1 = threading.Thread(target=Connection_API, args=(coursenames,rooms,))
        t1.start()
    except TimeoutError:
        messagebox.showerror('Authentication Timed Out', 'Error: Please try using a different internet connection')

def convert(file,specialisation):
    x2x = XLS2XLSX(file)
    parse(specialisation,x2x.to_xlsx())
    

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/calendar']


def Connection_API(coursenames, rooms):
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
                progBar['value'] += 2.22222222222
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
                elif "CSET212" in value:
                    summary = "Blockchain Foundations "
                elif "CSET213" in value:
                    summary = "Linux and Shell Programming "
                elif "CSET214" in value:
                    summary = "Data Analysis using Python "
                elif "CSET215" in value:
                    summary = "Graphics and Visual Computing "
                elif "CSET216" in value:
                    summary = "UI/UX Design for Human Computer Interface "
                elif "CSET217" in value:
                    summary = "Software Development with DevOps "
                elif "CSET218" in value:
                    summary = "Full Stack Development "
                elif "CSET219" in value:
                    summary = "Quantum Computing Foundations "
                elif "CSET220" in value:
                    summary = "Unmanned Aerial Vehicles "
                elif "CSET221" in value:
                    summary = "Robotic Process Automation Essentials "
                elif "CSET222" in value:
                    summary = "Microcontrollers, Robotics & Embedded Systems "
                elif "CSET223" in value:
                    summary = "Augmented Reality Foundations "
                elif "CSET224" in value:
                    summary = "Cloud Computing "
                elif "CSET238" in value:
                    summary = "Product Design Principles and Practice "
                if "(L)" in value:
                    summary += "Lecture"
                elif "(T)" in value:
                    summary += "Tutorial"
                elif "(P)" in value:
                    summary += "Lab"
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