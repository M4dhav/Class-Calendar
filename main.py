from __future__ import print_function
import datetime
import os.path
import openpyxl
import streamlit as st
from dotenv import load_dotenv
from xls2xlsx import XLS2XLSX
from ics import Calendar, Event

load_dotenv()

def connectionAPI(coursenames, rooms, cal):
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
            create_event(summary, location, description,dstart,tstart, tend,cal)


def create_event(summary, location, description,dstart,tstart, tend, cal_ref ):
        # Call the Calendar API  
    dstart = int(dstart)
    tstart = int(tstart)
    tend = int(tend)
    tstart = datetime.datetime(2023, 8, dstart, tstart, 30)
    tend = datetime.datetime(2023, 8, dstart, tend, 30)

    for i in range(22):
        e = Event()
        e.name=summary
        e.begin = tstart
        e.end = tend
        e.location = location
        e.description = description
        cal_ref.events.add(e)
        tstart = tstart + datetime.timedelta(days=7)
        tend = tend + datetime.timedelta(days=7)
def parse(wb,specialisation):
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
    ttwb = openpyxl.load_workbook(wb)
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


st.title('Timetable Excel Sheet to Google Calendar')
st.session_state.token = None

specialisation = st.selectbox('What is your specialisation?',("AI","Blockchain","Cyber Security","Data Science","Gaming","Core","DevOps","Full Stack","Quantum Computing","Drones","Robotics","IoT","AR/VR","Product Design","Cloud Computing"))

uploaded_file = st.file_uploader("Choose a file", ["xls","xlsx"])


if uploaded_file is not None:
    try:
        x2x = XLS2XLSX(uploaded_file)
        current_timestamp = datetime.datetime.now().timestamp()
        x2x.to_xlsx(f"{current_timestamp}.xlsx")
        coursenames, rooms=parse(f"{current_timestamp}.xlsx",specialisation)
        os.remove(f"{current_timestamp}.xlsx")
    except ValueError:
        coursenames, rooms = parse(uploaded_file, specialisation)
    cal = Calendar()
    connectionAPI(coursenames, rooms, cal)
    with open('my.ics', 'w') as f:
        f.writelines(cal.serialize_iter())
    st.write("Download the generated ics file")
    st.download_button(label="Download",data=cal.serialize(),file_name="timetable.ics",mime="text/calendar")
    st.write("We recommend creating a new calendar in your google calendar and importing the ics file into it.")
    st.write("You can then delete the calendar after the semester is over.")

    st.write("[Create new calendar in google calendar ↗](https://calendar.google.com/calendar/u/0/r/settings/createcalendar)")


    st.write("Import the ics file into your new calendar")

    st.write("[Import ics file into new calendar ↗](https://calendar.google.com/calendar/u/0/r/settings/export)")

