
import sys
import traceback
import openpyxl
from openpyxl.utils import get_column_letter
import datetime
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE  # Class in which the shape type is defined
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.dml.color import RGBColor    # Color management class
from pptx.util import Cm,Pt            # Class to specify units (centimeters, points)
from pptx.util import Inches, Pt
import subprocess
import time
import numpy as np
from pptx.enum.text import PP_ALIGN
import easygui
from tkinter import filedialog
from tkinter import Text
from tkinter import ttk
from tkinter import *
from tkinter.filedialog import asksaveasfile
import tkinter as tk
from openpyxl.utils.exceptions import *
import easygui
import json
import os
import calendar
from enum import Enum


excelEventsDictionary = {}

MAX_FIRST_MONTH_EVENT_COUNT = 31
MAX_SECOND_MONTH_EVENT_COUNT = 31
EXCEL_COULUMN_NUMBER = 9
FORBIDDEN_SENTENCE_HEBREW = "עלות פלוס: חינם | יאללה: חינם"
FORBIDDEN_SENTENCE_ARABIC = "السعر: شيكل"


class MetaData:
    def __init__ (self, firstMonthInteger = None, secondMonthInteger = None, firstMonthName = None, secondMonthName = None, year = None, contactName = None, contactPhone = None, splitYear = False,
        hasCacheFile = False):
        self.month1 = month1
        self.month2 = month2
        self.year=year
        self.contactName = contactName
        self.contactPhone = contactPhone
        self.splitYear = splitYear
        self.hasCacheFile = hasCacheFile


# An object holding the fields of the Excel input event: date (datetime), day (string), hour (datetime.time), title (string), 
# text1 (string), location (string), price (string), isFree (bool), isZoom (bool) and month (int) of a planned event 
class ExcelEvent:
    def __init__(self, date, hour, title, location, price, community, link, eventType):
        self.date = date
        self.hour = hour
        self.title = title
        self.location = location
        self.price = price
        self.community = community
        self.link = link
        self.eventType = eventType
        self.month = getMonth(date)
        self.day = getDayFromDate(date)

    #Comparison function, comparing Event objects
    #by date and hour (if dates are equal)
    def __lt__(self, other):
        return self.date.month <= other.date.month

    def __str__(self):
        return "date = %s, hour = %s, title = %s, month = %s"%(self.date.date(),self.hour, self.title, str(self.month))



# An object holding the different shapes of the input powerpoint event shape: dateShape (shape), dayShape (shape), hourShape (shape), titleShape (shape), 
# text1Shape (shape), locationShape (shape) priceShape, freeShape (shape), zoomShape (shape)
# and shape (the root element holding them all together)
class SingleEventShape:
    def __init__(self, zoomShape, freeShape, dateShape, titleShape, text1Shape, locShape, priceShape, shape, isTreated = False):
        self.dateShape = dateShape
        self.titleShape = titleShape
        self.text1Shape = text1Shape
        self.locShape = locShape
        self.priceShape = priceShape
        self.freeShape = freeShape
        self.zoomShape = zoomShape
        self.isTreated = isTreated
        self.shape = shape

class DoubleEventShape:
    def __init__(self, dateShape, title1Shape, loc1Shape, zoom1Shape, free1Shape, price1Shape, title2Shape, loc2Shape, zoom2Shape, free2Shape, price2Shape, shape, isTreated = False):
        self.dateShape = dateShape
        self.title1Shape = title1Shape
        self.loc1Shape = loc1Shape
        self.zoom1Shape = zoom1Shape
        self.free1Shape = free1Shape
        self.price1Shape = price1Shape
        self.title2Shape = title2Shape
        self.loc2Shape = loc2Shape
        self.zoom2Shape = zoom2Shape
        self.free2Shape = free2Shape
        self.price2Shape = price2Shape
        self.isTreated = isTreated
        self.shape = shape



# readExcel(excel_name)
#     Reading and analyzing the given excel file and returns a lists of events it contains
#     Arguments:
#         excel_name: a string path to the xlsx file holding
#             the data of all the planned events to be analyzed
#     Returns:
#         first_months_events: Event list containing all the events from excel that happen in the first month, 
#                              ignoring every event after MAX_FIRST_MONTH_EVENTS_COUNT
#         second_months_events: Event list containing all the events from excel that happen in the second month
#                              ignoring every event after MAX_SECOND_MONTH_EVENTs_COUNT
#         months: list of int containing the two months the excel deals with
#     Exceptions:
#         * CellCoordinatesException
#         * IllegalCharacterError
#         * InvalidFileException
#         * SheetTitleException
#         * MonthCountException: Less than two different months found in excel file
#         * MonthCountException: More than two different months found in excel file
#         * ValueError: time data does not match format

def readExcel(excel_name):

    # Define variable to load the dataframe
    try:
        ws = openpyxl.load_workbook(excel_name)

        # Define variable to read sheet
        sheet = ws.active
    
        events = getEventsFromExcel(sheet)

        if len(months)<2:
            raise Exception("קובץ האקסל מכיל פחות משני חודשים")


        return splitExcelEventsByMonths(events)
    except CellCoordinatesException:
        raise Exception("שגיאת המרה בין ערך נומרי ל-A1-style")
    except IllegalCharacterError:
        raise Exeption("קובץ האקסל מכיל תוים לא חוקייים")
    except InvalidFileException:
        raise Exception("שיגאה בעת נסיון פתיחת קובץ שאינו קובץ אקסל")


#Split the Event list into first month and second month events
def splitExcelEventsByMonths(events):
    first_months_events = []
    second_months_events = []

    for event in events:
        if event.month == metaData.firstMonthInteger:
            first_months_events.append(event)
        elif event.month == metaData.secondMonthInteger:
            second_months_events.append(event)
        else:
            raise Exception("קובץ האקסל מכיל מעל לשני חודשיים")
    return first_months_events, second_months_events


def trimLeadingZero(inputString):
    if inputString.startswith('0'):
        return str(inputString[1:])
    else:
        return inputString

def checkDateStringValidity(inputString):
    if inputString != None:
        if len(inputString)==5:
            inputString = str(trimLeadingZero(inputString))
            if '.' in inputString:
                arrayString = inputString.split('.')
                if len(arrayString)==2:
                    if len(arrayString[1])==2:
                        return True
    return False

def getMonth(eventDate):
    eventDategArr = eventDate.split('.')
    if len(eventDategArr) < 2:
       eventDateMiddleSplit = eventDate.split(' ')
       eventDategArr = eventDateMiddleSplit[0].split('-')
    return int(eventStringArr[1])

def getDay(eventDate):
    eventDategArr = eventDate.split('.')
    if len(eventDategArr) < 2:
       eventDateMiddleSplit = eventDate.split(' ')
       eventDategArr = eventDateMiddleSplit[0].split('-')
    return int(eventStringArr[0])



# getEventsAndMonthsFromExcel(sheet)
#     Itereates through excel rows and cells to create Event and months lists
#     Arguments:
#         sheet: a Worksheet object to itereate through
#     Returns:
#         events: a sorted Event list containing all the events from the Worksheet
#         months: a sorted int list containing the months the excel Worksheet deals with
#     Exceptions:
#         ValueError

def getEventsFromExcel(sheet):
    excelEvents = []
    excelMonths = []
    community, link, eventType = None
    # Iterate the loop to read the cell values
    # Ignore first row with column titles
    for rowIndex in range(2, sheet.max_row+1):
        date = sheet.cell(row=rowIndex,column=1).value
        if not checkDateStringValidity(str(date)):
            if str(date)=="":
                if rowIsEmpty(sheet, rowIndex, EXCEL_COULUMN_NUMBER):
                    continue;
                else:
                    raise Exception("שורה \n"+str(rowIndex)+"\n עמודה: \n"+str(1)+"\nהערך ריק או לא תואם את התבנית:\n "+"DD.MM")
        
        eventDayInMonth = getDay(str(date))

        day = sheet.cell(row=rowIndex,column=2).value
        if day is None: 
            raise Exception("שורה \n"+str(rowIndex)+"\n עמודה: \n"+str(2)+"\nהערך ריק ")

        hourValue = sheet.cell(row=rowIndex,column=3).value
        #May throw: ValueError: time data does not match format
        # try: 
        #     if isinstance(hourValue, datetime.time):
        #         hour = hourValue.strftime("%H:%M")
        #     elif isinstance(hourValue, datetime.datetime):
        #         hour = hourValue.time().strftime("%H:%M")
        #     else:
        #         hour = str(hourValue)
        # except ValueError:
        #     raise Exception("שורה \n"+str(rowIndex)+"\n עמודה: \n"+str(3)+"\nהערך ריק או לא תואם את התבנית:\n "+"HH:MM")

        title = sheet.cell(row=rowIndex,column=4).value
        if title is None:
            title = ""


        text1 = sheet.cell(row=rowIndex,column=5).value
        if text1 is None:
            text1 = ""

        location = sheet.cell(row=rowIndex,column=6).value
        if location is None:
            location = ""

        price = sheet.cell(row=rowIndex, column=7).value
        if price is None:
            price = ""


        isFree = sheet.cell(row = rowIndex, column = 8).value
        if isFree is None:
            isFree = False

        isZoom = sheet.cell(row=rowIndex, column = 9).value
        if isZoom is None:
            isZoom = False

        excelEvent = ExcelEvent(str(date), str(hourValue), title, location, price, community, link, eventType)
        excelEvents.append(excelEvent)
        if excelEvent.month not in excelMonths:
             months.append(excelEvent.month)
    if not (12 in excelMonths and 1 in excelMonths):
        excelMonths.sort()
    metaData.firstMonthInteger = excelMonths[0]
    metaData.secondMonthInteger = excelMonths[1]
    excelEventsDictionary[eventDayInMonth] = event
    return events



def rowIsEmpty(sheet, rowIndex, maxColIndex):
    for i in range(1,maxColIndex+1):
        if sheet.cell(row=rowIndex,column=i).value !=None:
            return False
        else:
            if sheet.cell(row = rowIndex, column = 7) == None and sheet.cell(row=rowIndex, column =8) == False:
                return False
            return True

 
# get_text_boxes(slide, months)
#     Itereates through all powerpoint shapes (using iter_textframed_shapes method) that contain textboxes, sort them,
#     create month textboxes and returns an event_shapes sorted list of all textable leaf (not group) EventShape objects
#     Arguments:
#         slide: a Presentation Slide object to analize for shapes
#         months: array of ints, representing the numeric (int) representation of months presenting in the excel events sheet
#     Returns:
#         event_shapes: a sorted list of all textable leaf (not group) EventShape objects in the given slide
#     Exceptions:
#         * Not all cells of Excel are filled

def get_text_boxes(slide, months, area, contact):
    singles, doubles, fridays = find_groups(slide.shapes)

    singles = reversed(singles)
    doubles = reversed(doubles)
    fridays.sort(key= lambda x: x.name)

    single_zoom = []
    single_free = []
    single_dates = [] 
    single_titles = []
    single_texts1 = []
    single_locs = []
    single_prices = []
    single_shapes = []

    double_dates = []
    double_titles1 = []
    double_locs1 = [] 
    double_zoom1 = []
    double_free1 = []
    double_prices1 = []
    double_titles2 = []
    double_locs2 = []
    double_zoom2 = []
    double_free2 = []
    double_prices2 = []
    double_shapes = []

    friday_shapes = []
    friday_dates = []
    friday_titles = []
    friday_locs = []
    friday_prices = []

    for g in singles:
        single_shapes.append(g)
        for shape in iter_textable_shapes(g.shapes):
            if shape.name == 'DATE':
                single_dates.append(shape)
            if shape.name == 'TITLE':
                single_titles.append(shape)
            if shape.name == 'TEXT1':
                single_texts1.append(shape)
            if shape.name == 'LOC':
                single_locs.append(shape)
            if shape.name == 'PRICE':
                single_prices.append(shape)
            if shape.name == 'FREE':
                single_free.append(shape)
            if shape.name == 'ZOOM':
                single_zoom.append(shape)
    for g in doubles:
        double_shapes.append(g)
        for shape in iter_textable_shapes(g.shapes):
            if shape.name == 'DATE':
                double_dates.append(shape)
            if shape.name == 'TITLE1':
                double_titles1.append(shape)
            if shape.name == 'LOC1':
                double_locs1.append(shape)
            if shape.name == 'PRICE1':
                double_prices1.append(shape)
            if shape.name == 'FREE1':
                double_free1.append(shape)
            if shape.name == 'ZOOM1':
                double_zoom1.append(shape)
            if shape.name == 'TITLE2':
                double_titles2.append(shape)
            if shape.name == 'LOC2':
                double_locs2.append(shape)
            if shape.name == 'ZOOM2':
                double_zoom2.append(shape)
            if shape.name == 'FREE2':
                double_free2.append(shape)
            if shape.name == 'PRICE2':
                double_prices2.append(shape)
    for g in fridays:
        friday_shapes.append(g)
        for shape in iter_textable_shapes(g.shapes):
            if shape.name == "DATE":
                friday_dates.append(shape)
            if shape.name == "TITLE":
                friday_titles.append(shape)
            if shape.name == "LOC":
                friday_locs.append(shape)
            if shape.name == "PRICE":
                friday_prices.append(shape)


    textable_shapes = list(iter_textframed_shapes(slide.shapes))
    ordered_textable_shapes = sorted(
        textable_shapes, key=lambda shape: (shape.top, shape.left)
    )

    for shape in ordered_textable_shapes:
        if shape.name.startswith('MONTH1'):
            createMonthObjects(shape, months, 0)
        if shape.name.startswith('MONTH2'):
            createMonthObjects(shape, months, 1)
        if shape.name == "AREA":
            createAreaObjects(shape, area)
        if shape.name == "CONTACT":
            createContactObject(shape, contact)

    # days = createRightOrder(days)
    #dates = createRightOrder(dates)
    # hours = createRightOrder(hours)
    # titles = createRightOrder(titles)
    # texts1 = createRightOrder(texts1)
    # locs = createRightOrder(locs)
    # prices = createRightOrder(prices)
    # free = createRightOrder(free)
    # zoom = createRightOrder(zoom)
    #shapes = sorted(shapes, key=lambda x: get_int_from_shape_name(x.name))

    single_event_shapes = []
    double_event_shapes = []
    friday_event_shapes = []

    for i in range(25):
        event_shape = SingleEventShape(single_zoom[i],single_free[i], single_dates[i], single_titles[i], single_texts1[i], single_locs[i], single_prices[i], single_shapes[i])
        single_event_shapes.append(event_shape)
    for i in range(25):
        event_shape = DoubleEventShape(double_dates[i],double_titles1[i], double_locs1[i], double_zoom1[i], double_free1[i], double_prices1[i], double_titles2[i], 
                                        double_locs2[i], double_zoom2[i], double_free2[i], double_prices2[i], double_shapes[i])
        double_event_shapes.append(event_shape)
    for i in range(len(friday_shapes)):
        event_shape = SingleEventShape(None,None, friday_dates[i], friday_titles[i], None, friday_locs[i], friday_prices[i], friday_shapes[i])
        friday_event_shapes.append(event_shape)
    return single_event_shapes, double_event_shapes, friday_event_shapes


# get_int_from_shape_name(shape_name)
#     Get the number of the given shape name
#     Arguments:
#         shape_name: the shape object name, expecting name of format "XXX %d"
#     Returns:
#         the number in the shape name
#     Exceptions:
#         * Wrong element name format

def find_groups(shapes):
    doubles = []
    singles = []
    fridays = []
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            if shape.name.startswith("DOUBLE"):
                doubles.append(shape)
            elif shape.name.startswith("ELEMENT"):
                singles.append(shape)
            elif shape.name.startswith("FRIDAY"):
                fridays.append(shape)
    return singles, doubles, fridays


def get_int_from_shape_name(shape_name):
    try: 
        shape_name_list = shape_name.split()
        return int(shape_name_list[1])
    except:
        raise Exception("שמות האלמנטים בקובץ הפאוורפוינט לדוגמא חייבים להיות מהצורה ELEMENT num, למשל ELEMENT 12")


# createMonthObjects(shape, months, monthIndex):
#     Fill the given shape with the string representation of the months list in the given monthIndex
#     Arguments:
#         shape: shape Object to be filled with the month string representation
#         months: a list of numerical months representation (for example: [1,4])
#         monthIndex: the index of the month in the months array to be written in the shape in it's string representation
#     Exceptions: 
#         * Month should be 1-12. no month representing for "+str(int_month)

def createMonthObjects(shape, months, monthIndex):
    month = months[monthIndex]
    month_text_frame = shape.text_frame
    p_month = month_text_frame.paragraphs[0]
    p_month.alignment = PP_ALIGN.RIGHT
    clearTextboxText(p_month, month_text_frame)
    run_month = p_month.runs[0]
    run_month.text = month

def createAreaObjects(shape, area):
    area_text_frame = shape.text_frame
    p_area = area_text_frame.paragraphs[0]
    p_area.alignment = PP_ALIGN.RIGHT
    clearTextboxText(p_area, area_text_frame)
    run_area = p_area.runs[0]
    run_area.text = area

def createContactObject(shape, contact):
    contact_text_frame = shape.text_frame
    p_contact = contact_text_frame.paragraphs[0]
    p_contact.alignment = PP_ALIGN.CENTER
    clearTextboxText(p_contact, contact_text_frame)
    run_contact = p_contact.runs[0]
    run_contact.text = contact



# iter_textframed_shapes(shapes):
#     Itereates through all powerpoint shapes (using iter_textable_shapes method) that contain textboxes and yields
#     the leaf shapes with text boxes and their parent group shapes
#     Arguments:
#         shapes: a list of shapes (Shape objects) to iterate throuth
#     Returns:
#         a leaf shape with textbox or a group shape hilding shapes with textboxes

def iter_textframed_shapes(shapes):
    """Generate shape objects in shapes that can contain text.

    Shape objects are generated in document order (z-order), bottom to top.
    """
    for shape in shapes:
        # ---recurse on group shapes---
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            group_shape = shape
            yield shape
            iter_textframed_shapes(group_shape.shapes)
            for shape in iter_textable_shapes(group_shape.shapes):
                yield shape
            continue

        # ---otherwise, treat shape as a "leaf" shape---
        if shape.has_text_frame:
            yield shape


# iter_textable_shapes(shapes):
#     Itereates through shapes in group shape to find and yield the leaf shape with textboxes
#     Arguments:
#         shapes: a list of shapes (Shape objects) to iterate throuth
#     Returns:
#         a leaf shape with textbox

def iter_textable_shapes(shapes):
    for shape in shapes:
        if shape.has_text_frame:
            yield shape


# iter_textable_shapes(shapes):
#     *some magic*
#     Arguments:
#         array: the list to be right-ordered
#     Returns:
#         right-ordered list

def createRightOrder(array):
    return np.append(array[MAX_SECOND_MONTH_EVENT_COUNT:], array[:MAX_SECOND_MONTH_EVENT_COUNT])



# create_program_GUI():
#     creates the program GUI window

def create_program_GUI():

    useUserSavedInput = False
    userSavedInput = None

    try:
        cache_file = open("cache.txt", "r")
        content = cache_file.read()
        if os.stat("cache.txt").st_size != 0:
            userSavedInput = json.loads(content)
            useUserSavedInput = True
        cache_file.close()
    except FileNotFoundError:
        pass

    def assignExcelPath(value):
        global excelFilePath
        validity, extension = checkValidity(value, "xlsx")
        excelFilePath = value
        textbox1.configure(state="normal")
        if len(textbox1.get("1.0", tk.END))>1:
            textbox1.delete(1.0,tk.END)
        textbox1.insert(tk.END, excelFilePath)
        textbox1.configure(state="disabled")

    def assignPowerpointPath(value):
        global pptxFilePath
        pptxFilePath = value
        validity, extension = checkValidity(value, "pptx")
        textbox2.configure(state="normal")
        if len(textbox2.get("1.0", tk.END))>1:
            textbox2.delete(1.0,tk.END)
        textbox2.insert(tk.END, pptxFilePath)
        textbox2.configure(state="disabled")


    def assignFileLocation(value):
        global fileLocation 
        fileLocation = value;
        place_textbox.configure(state = "normal")
        place_textbox.insert(tk.END, fileLocation)
        place_textbox.configure(state="disabled")

    try:

        app = tk.Tk()
        app.wm_title("2MonthPlanCreator")
        app.geometry("650x250")
        app.padx = (100,0)
        app['background']='#015293'

        headLabel = tk.Label(text='פרטים לתוכנית ארועים',  font='Arial 14 bold')
        headLabel.grid(sticky="e", row = 0, column = 5, columnspan = 2)
        headLabel['background']='#015293'
        headLabel.config(fg= "#fdc53e")

        assignFirstMonthLabel = tk.Label(text="חודש ראשון")
        assignFirstMonthLabel.grid(sticky="e", row = 1, column =6)
        assignFirstMonthLabel['background']='#015293'
        assignFirstMonthLabel.config(fg="white")

        assignFirstMonthTextbox = ttk.Entry(justify=tk.RIGHT)
        assignFirstMonthTextbox.grid(stick="e", row=1, column=5)
        if useUserSavedInput == True and userSavedInput != None:
            assignFirstMonthTextbox.insert(0, userSavedInput.get("month1"))


        assignSecondMonthLabel = tk.Label(text="חודש שני")
        assignSecondMonthLabel.grid(sticky="e", row=1, column = 4)
        assignSecondMonthLabel['background']='#015293'
        assignSecondMonthLabel.config(fg="white")

        assignSecondMonthTextbox = ttk.Entry(justify=tk.RIGHT)
        assignSecondMonthTextbox.grid(sticky="e", row=1, column = 3, padx=(10,10))
        if useUserSavedInput == True and userSavedInput != None:
            assignSecondMonthTextbox.insert(0, userSavedInput.get("month2"))

        assignYearLabel = tk.Label(text = "שנה")
        assignYearLabel.grid(sticky="e", row = 1, column = 2)
        assignYearLabel['background']='#015293'
        assignYearLabel.config(fg="white")

        assignYearTextbox = ttk.Entry(justify=tk.RIGHT, width="10")
        assignYearTextbox.grid(sticky="e", row = 1, column = 1, padx=(10,10))
        if useUserSavedInput == True and userSavedInput != None:
            assignYearTextbox.insert(0, userSavedInput.get("year"))

        language = tk.StringVar()

        hebrew = Radiobutton(app, text="עברית", variable = language, value = 1)
        hebrew['background']='#015293'
        hebrew.config(fg="white", selectcolor="#4cb263")
        hebrew.grid(row = 2, column = 6, sticky="e")
        arabic = Radiobutton(app, text="ערבית", variable = language, value = 2)
        arabic['background']='#015293'
        arabic.config(fg="white", selectcolor="#4cb263")
        arabic.grid(row = 2, column = 5, sticky="e")

        if useUserSavedInput == True and userSavedInput != None:
            language.set(userSavedInput.get("language"))
        else:
            language.set(1)

        assignLocationLabel = tk.Label(text="מיקום")
        assignLocationLabel.grid(sticky="e", row = 3, column = 6)
        assignLocationLabel['background']='#015293'
        assignLocationLabel.config(fg="white")

        assignLocationTextbox = ttk.Entry(justify=tk.RIGHT)
        assignLocationTextbox.grid(sticky="e", row = 3, column = 5)
        if useUserSavedInput == True and userSavedInput != None:
            assignLocationTextbox.insert(0, userSavedInput.get("area"))

        assignContactLabel = tk.Label(text = "לקבלת פרטים לפנות אל")
        assignContactLabel.grid(sticky="e", row = 4, column = 6)
        assignContactLabel['background']='#015293'
        assignContactLabel.config(fg="white")

        assignContactTextbox = ttk.Entry(justify=tk.RIGHT)
        assignContactTextbox.grid(sticky="e", row = 4, column = 5)
        if useUserSavedInput == True and userSavedInput != None:
            assignContactTextbox.insert(tk.END, userSavedInput.get("contact"))

        directions1 = tk.Label(text = "אנא בחרו קובץ אקסל עם המידע הרלוונטי וקובץ פאוורפוינט לדוגמא", justify="right")
        directions1['background']='#015293'
        directions1.config(fg="white")
        directions1.grid(sticky="e", row=5, column=4, columnspan=4)

        btn1 = tk.Button(text="בחרו קובץ אקסל", command=lambda: assignExcelPath(filedialog.askopenfilename(filetypes=[("קובץ אקסל" , ".xlsx")])))
        btn1.grid(sticky="e", row=6, column=6)
        btn1['background']='#008fd1'
        btn1.config(fg="white")

        # Create text widget and specify size.
        textbox1 = Text(app, height = 1, width = 55)
        textbox1.grid(sticky="e", row=6, column=1, columnspan=5)
        textbox1.configure(state="disabled")

        btn2 = tk.Button(text="בחרו קובץ פאוורפוינט", command=lambda: assignPowerpointPath(filedialog.askopenfilename(filetypes=[("קובץ פאוורפוינט" , ".pptx")])))
        btn2.grid(sticky="e", row=7, column = 6)
        btn2['background']='#008fd1'
        btn2.config(fg="white")

        # Create text widget and specify size.
        textbox2 = Text(app, height = 1, width = 55)
        textbox2.grid(sticky="e", row=7, column=1, columnspan=5)
        textbox2.configure(state="disabled")

        btn4 = tk.Button(text="צרו תוכנית דו-חודשית!", command= lambda: createPptxPlan(assignFirstMonthTextbox.get(), assignSecondMonthTextbox.get(), assignLocationTextbox.get(), assignContactTextbox.get(),assignYearTextbox.get(), language.get()))
        btn4.grid(sticky="e", row=9, column = 5)
        btn4['background']='#4cb263'
        btn4.config(fg="white")


        btn5 = tk.Button(text="סגירה", command=app.destroy)
        btn5.grid(sticky="e", row=9, column = 4)
        btn5['background']='#008fd1'
        btn5.config(fg="white")

        app.mainloop()  

    except Exception as e:
        easygui.msgbox(e)
        raise Exception("קובץ ה-cache אינו תקין")




# checkValidity(value, extension):
#     Checks if the given filepath's extension matches the desired extesion and returns 
#     Arguments:
#         value: the given filepath to be checked
#         extension: the desired extension
#     Returns:
#         a tuple of: 
#             True/False depending on whether the given value file extension matches the desired extension
#             a string containing the real extension of the file given 
#             if the value does not contain an extension ('.xxx') returns (False, "")

def checkValidity(value, extension):
    if extension == "":
    #check for filename validity
        try:
            nameExtension = value.split('.')
            if len(nameExtension)>2:
                return False, ""
            elif len(nameExtension) == 2 and nameExtension[1]!="pptx":
                return True, ""
            elif len(nameExtension) ==2 and nameExtension[1] == "[pptx]":
                return True, ""
            else:
                return True, "pptx"
        except:
            return False, ""
    else:
        try:
            valueExtension = value.split('.')[1]
            valueExtension = ''.join(c for c in valueExtension if c.isprintable())
            return valueExtension == extension, valueExtension
        except:
            return False, ""


# createPptxPlan(month1, month2, area, contact, year, language):
#     Creates and saves the two-month plan pptx file 
#     Arguments:
#         month1:
#         month2:
#         area:
#         contact:
#         year:
#         language: 

def createPptxPlan(month1, month2, area, contact, year, language):
    try: 
        cache_file = open("cache.txt", "w")
        userInput = {
            "month1": month1,
            "month2": month2,
            "area": area,
            "contact": contact,
            "year": year,
            "language": language
        }
        cache_file.write(json.dumps(userInput))
        cache_file.close
        months_names = [month1,month2]

        excel_name=excelFilePath
        pptx_name = pptxFilePath

        first_months_events, second_months_events, monthsFromExcel = readExcel(excel_name)

        presentation = Presentation(pptx_name)

        slide = presentation.slides[0]
        single_event_shapes, double_event_shapes, friday_event_shapes = get_text_boxes(slide, months_names, area, contact)
        increaseYear = False

        calendar.setfirstweekday(6)
        if len(year)!=4:
            raise Exception("שדה השנה צריך להכיל 4 תוים בדיוק. למשל: 2023")
        createCalendarDates(single_event_shapes, year, monthsFromExcel[0])
        createCalendarDates(double_event_shapes, year, monthsFromExcel[0])

        writeTextToTextboxes(slide, first_months_events, single_event_shapes, double_event_shapes, friday_event_shapes, language, year, monthsFromExcel[0])

        try:
            slide = presentation.slides[1]
        except IndexError:
            raise Exception("תבנית ה-powerpoint מכילה פחות מ-2 שקפים")

        single_event_shapes, double_event_shapes, friday_event_shapes = get_text_boxes(slide, months_names, area, contact)

        calendar.setfirstweekday(6)
        if monthsFromExcel[1]==1:
            increaseYear = True
        createCalendarDates(single_event_shapes, year, monthsFromExcel[1], increaseYear)
        createCalendarDates(double_event_shapes, year, monthsFromExcel[1], increaseYear)

        writeTextToTextboxes(slide, second_months_events, single_event_shapes, double_event_shapes, friday_event_shapes, language, year, monthsFromExcel[1], increaseYear)


        files = [('Powerpoint files', '*.pptx')]
        file = asksaveasfile(filetypes = files, defaultextension = files)
        try:
            presentation.save(file.name)
            easygui.msgbox("הקובץ נוצר בהצלחה")
        except PermissionError: 
            easygui.msgbox("הקובץ \n"+ file.name+"\n פתוח. יש לסגור אותו בעת הרצת התוכנית")

    except Exception as e:
        easygui.msgbox("שגיאה :"+ str(e))

def clearTextboxText(first_paragrpah, text_frame):
    if len(first_paragrpah.runs)>1:
        for i in range(1,len(first_paragrpah.runs)):
            first_paragrpah.runs[i].text=""
    if len(text_frame.paragraphs)>1:
        for i in range(1, len(text_frame.paragraphs)):
            para = text_frame.paragraphs[i]
            for j in range (0, len(para.runs)):
                para.runs[j].text=""

def get_number_of_shape(year, month, day):
    numOfDaysInMonth = calendar.monthrange(year, month)[1]
    x = np.array(calendar.monthcalendar(year, month))
    week_of_month = np.where(x==day)[0][0] # 0 is first week
    day_of_week = np.where(x == day)[1][0]+1 # 1 is Sunday
    if day_of_week > 5: 
        if day_of_week == 6:
            return 100
        return -1
    first_day_of_month = np.where(x == 1)[1][0] + 1 # for removing first week that starts on Friday or Saturday
    if first_day_of_month <= 5: 
        return week_of_month*5 + day_of_week
    else:
        return (week_of_month - 1)*5 + day_of_week

def createCalendarDates(eventsShape, year, month, increaseYear = False):

    year = int(year)
    if increaseYear:
        year += 1

    numOfDaysInMonth = calendar.monthrange(year, month)[1]
    for i in range(1, numOfDaysInMonth+1):
        num_of_shape = get_number_of_shape(year, month, i) - 1
        if num_of_shape < 0 or num_of_shape >= 99:
            continue
        else:
            event_shape = eventsShape[num_of_shape]
            date_text_frame = event_shape.dateShape.text_frame
            p_date = date_text_frame.paragraphs[0]
            p_date.alignment = PP_ALIGN.RIGHT
            clearTextboxText(p_date, date_text_frame)
            run_date = p_date.runs[0]
            run_date.text =str(i) 

def findEventDay(event):
    date = event.date
    dateArray = date.split('.')
    if len(dateArray) < 2:
       dateMiddleSplit = date.split(' ')
       return dateMiddleSplit[0].split('-')[2]
    return int(dateArray[0])



def writeTextToTextboxes(slide, monthsEvents, singleEventsShape, doubleEventsShape, fridayEventsShape, language, year, month, increaseYear = False):

    year = int(year)
    if increaseYear:
        year += 1
    numOfDaysInMonth = calendar.monthrange(year, month)[1]
    first_day_of_week1 = get_number_of_shape(year, month, 1) - 1

    i = 0
    fridayCount = 0

    while i< len(monthsEvents):

        event = monthsEvents[i]
        day = findEventDay(event)

        next_event = None
        next_next_event = None
        isDouble = False
        isFriday = False

        event_shape_num = get_number_of_shape(year, month, day)-1
        if event_shape_num >= 99:
            isFriday = True
            fridayCount += 1
            if fridayCount >= 6:
                raise Exception("לא ניתן להזין מכל ל-5 ימי שישי בחודש")

        if (i+1) < len(monthsEvents):
            next_event = monthsEvents[i+1]
            next_event_day = findEventDay(next_event)
            if next_event_day == day:
                if (i+2) < len(monthsEvents):
                    next_next_event = monthsEvents[i+2]
                    next_next_event_day = findEventDay(next_next_event)
                    if next_event_day == next_next_event_day:
                        raise Exception("אין לכלול בקובץ האקסל מעל ל-2 ארועים באותו יום")
                isDouble = True
                if isFriday:
                    raise Exception("ביום שישי \n"+str(day)+"."+str(month)+"\n"+"קיימים שניים או יותר ארועים. לא ניתן להזין מעל לארוע אחד בכל יום שישי")
                i+=1

        if event_shape_num >= 99 and fridayEventsShape:
            friday_event_shape = fridayEventsShape[fridayCount -1]
        elif event_shape_num < 0:
            single_event_shape = double_event_shape = None
        elif event_shape_num >= 0 and event_shape_num < 99:
            single_event_shape = singleEventsShape[event_shape_num]
            single_event_shape.isTreated = True
            double_event_shape = doubleEventsShape[event_shape_num]
            double_event_shape.isTreated = True

        if event_shape_num < 0:
            i+=1
            continue
        else:
            if isDouble == False and isFriday == False: # It's a standart single event
                event_shape = single_event_shape
                #deleteDoubleShape
                slide.shapes.element.remove(double_event_shape.shape.element)

                title_text_frame = event_shape.titleShape.text_frame
                p_title = title_text_frame.paragraphs[0]
                p_title.runs[0].text = ""
                p_title.alignment = PP_ALIGN.RIGHT
                clearTextboxText(p_title, title_text_frame)
                run_title = p_title.runs[0]
                run_title.text = event.hour + ": " + event.title

                text1_text_frame = event_shape.text1Shape.text_frame
                p_text1 = text1_text_frame.paragraphs[0]
                p_text1.runs[0].text = ""
                p_text1.alignment = PP_ALIGN.RIGHT
                clearTextboxText(p_text1, text1_text_frame)
                run_text1 = p_text1.runs[0]
                run_text1.text = event.text1

                loc_text_frame = event_shape.locShape.text_frame
                p_loc = loc_text_frame.paragraphs[0]
                p_loc.alignment = PP_ALIGN.RIGHT
                clearTextboxText(p_loc, loc_text_frame)
                run_loc = p_loc.runs[0]
                run_loc.text = event.location

                price_text_frame = event_shape.priceShape.text_frame
                p_price = price_text_frame.paragraphs[0]
                p_price.runs[0].text = ""
                p_price.alignment = PP_ALIGN.RIGHT
                clearTextboxText(p_price, price_text_frame)
                run_price = p_price.runs[0]
                if language == "1": 
                    if event.price != "" and event.price != FORBIDDEN_SENTENCE_HEBREW:
                        run_price.text = event.price
                else:
                    if event.price != "" and event.price != FORBIDDEN_SENTENCE_ARABIC:
                        run_price.text = event.price

                if event.isFree == False:
                    freeElement = event_shape.freeShape
                    event_shape.shape.shapes.element.remove(freeElement.element)

                if event.isZoom == False:
                    zoomElement = event_shape.zoomShape
                    event_shape.shape.shapes.element.remove(zoomElement.element)

            elif isDouble == False and isFriday == True and fridayEventsShape: # It's a Friday event
                    event_shape = friday_event_shape

                    date_text_frame = event_shape.dateShape.text_frame
                    p_date = date_text_frame.paragraphs[0]
                    p_date.runs[0].text = ""
                    p_date.alignment = PP_ALIGN.RIGHT
                    clearTextboxText(p_date, date_text_frame)
                    run_date = p_date.runs[0]
                    run_date.text = event.date #TODO: add the word Friday (in hebrew or arabic, according to the language parameter)

                    title_text_frame = event_shape.titleShape.text_frame
                    p_title = title_text_frame.paragraphs[0]
                    p_title.runs[0].text = ""
                    p_title.alignment = PP_ALIGN.RIGHT
                    clearTextboxText(p_title, title_text_frame)
                    run_title = p_title.runs[0]
                    run_title.text = event.hour + ": " + event.title

                    loc_text_frame = event_shape.locShape.text_frame
                    p_loc = loc_text_frame.paragraphs[0]
                    p_loc.alignment = PP_ALIGN.RIGHT
                    clearTextboxText(p_loc, loc_text_frame)
                    run_loc = p_loc.runs[0]
                    run_loc.text = event.location

                    price_text_frame = event_shape.priceShape.text_frame
                    p_price = price_text_frame.paragraphs[0]
                    p_price.runs[0].text = ""
                    p_price.alignment = PP_ALIGN.RIGHT
                    clearTextboxText(p_price, price_text_frame)
                    run_price = p_price.runs[0]
                    if language == "1": 
                        if event.price != "" and event.price != FORBIDDEN_SENTENCE_HEBREW:
                            run_price.text = event.price
                    else:
                        if event.price != "" and event.price != FORBIDDEN_SENTENCE_ARABIC:
                            run_price.text = event.price

            elif isFriday == False: #isDouble == True, isFriday should be False, it's a double event
                event_shape = double_event_shape
                slide.shapes.element.remove(single_event_shape.shape.element)


                title1_text_frame = event_shape.title1Shape.text_frame
                p_title1 = title1_text_frame.paragraphs[0]
                p_title1.runs[0].text = ""
                p_title1.alignment = PP_ALIGN.RIGHT
                clearTextboxText(p_title1, title1_text_frame)
                run_title1 = p_title1.runs[0]
                run_title1.text = event.hour + ": " + event.title

                title2_text_frame = event_shape.title2Shape.text_frame
                p_title2 = title2_text_frame.paragraphs[0]
                p_title2.runs[0].text = ""
                p_title2.alignment = PP_ALIGN.RIGHT
                clearTextboxText(p_title2, title2_text_frame)
                run_title2 = p_title2.runs[0]
                run_title2.text = next_event.hour + ": " + next_event.title

                loc1_text_frame = event_shape.loc1Shape.text_frame
                p_loc1 = loc1_text_frame.paragraphs[0]
                p_loc1.alignment = PP_ALIGN.RIGHT
                clearTextboxText(p_loc1, loc1_text_frame)
                run_loc1 = p_loc1.runs[0]
                run_loc1.text = event.location

                loc2_text_frame = event_shape.loc2Shape.text_frame
                p_loc2 = loc2_text_frame.paragraphs[0]
                p_loc2.alignment = PP_ALIGN.RIGHT
                clearTextboxText(p_loc2, loc2_text_frame)
                run_loc2 = p_loc2.runs[0]
                run_loc2.text = next_event.location

                price1_text_frame = event_shape.price1Shape.text_frame
                p_price1 = price1_text_frame.paragraphs[0]
                p_price1.runs[0].text = ""
                p_price1.alignment = PP_ALIGN.RIGHT
                clearTextboxText(p_price1, price1_text_frame)
                run_price1 = p_price1.runs[0]
                if language == "1": 
                    if event.price != "" and event.price != FORBIDDEN_SENTENCE_HEBREW:
                        run_price1.text = event.price
                else:
                    if event.price != "" and event.price != FORBIDDEN_SENTENCE_ARABIC:
                        run_price1.text = event.price

                if event.isFree == False:
                    free1Element = event_shape.free1Shape
                    event_shape.shape.shapes.element.remove(free1Element.element)

                if event.isZoom == False:
                    zoom1Element = event_shape.zoom1Shape
                    event_shape.shape.shapes.element.remove(zoom1Element.element)

                price2_text_frame = event_shape.price2Shape.text_frame
                p_price2 = price2_text_frame.paragraphs[0]
                p_price2.runs[0].text = ""
                p_price2.alignment = PP_ALIGN.RIGHT
                clearTextboxText(p_price2, price2_text_frame)
                run_price2 = p_price2.runs[0]
                if language == "1": 
                    if next_event.price != "" and next_event.price != FORBIDDEN_SENTENCE_HEBREW:
                        run_price2.text = next_event.price
                else:
                    if next_event.price != "" and next_event.price != FORBIDDEN_SENTENCE_ARABIC:
                        run_price2.text = next_event.price

                if next_event.isFree == False:
                    run_price2.text = next_event.price
                    free2Element = event_shape.free2Shape
                    event_shape.shape.shapes.element.remove(free2Element.element)

                if next_event.isZoom == False:
                    zoom2Element = event_shape.zoom2Shape
                    event_shape.shape.shapes.element.remove(zoom2Element.element)

            i+=1

    firstDayReached = False 
    lastDayReached = False
    lastDayOfMonth = calendar.monthrange(year, month)[1]
    numOfShapeOfLastDay = get_number_of_shape(year, month, lastDayOfMonth)
    lastDayIsFriday = (numOfShapeOfLastDay == 100)
    lastDayIsSaurday = (numOfShapeOfLastDay == -1)

    for i in range(25):
        if singleEventsShape[i].dateShape.text == "6" and ((not firstDayReached) and (not lastDayReached)):
            slide.shapes.element.remove(doubleEventsShape[i].shape.element)
            singleEventsShape[i].shape.shapes.element.remove(singleEventsShape[i].dateShape.element)
            singleEventsShape[i].shape.shapes.element.remove(singleEventsShape[i].titleShape.element)
            singleEventsShape[i].shape.shapes.element.remove(singleEventsShape[i].text1Shape.element)
            singleEventsShape[i].shape.shapes.element.remove(singleEventsShape[i].locShape.element)
            singleEventsShape[i].shape.shapes.element.remove(singleEventsShape[i].priceShape.element)
            singleEventsShape[i].shape.shapes.element.remove(singleEventsShape[i].freeShape.element)
            singleEventsShape[i].shape.shapes.element.remove(singleEventsShape[i].zoomShape.element)
            continue
        else:
            firstDayReached = True
        if singleEventsShape[i].isTreated == False and doubleEventsShape[i].isTreated == False:
            slide.shapes.element.remove(doubleEventsShape[i].shape.element)

            #if lastDayReached and singleEventsShape[i].dateShape.text == "6":
            if lastDayReached:
                singleEventsShape[i].shape.shapes.element.remove(singleEventsShape[i].dateShape.element)
                singleEventsShape[i].shape.shapes.element.remove(singleEventsShape[i].titleShape.element)
                singleEventsShape[i].shape.shapes.element.remove(singleEventsShape[i].text1Shape.element)
                singleEventsShape[i].shape.shapes.element.remove(singleEventsShape[i].locShape.element)
                singleEventsShape[i].shape.shapes.element.remove(singleEventsShape[i].priceShape.element)
                singleEventsShape[i].shape.shapes.element.remove(singleEventsShape[i].freeShape.element)
                singleEventsShape[i].shape.shapes.element.remove(singleEventsShape[i].zoomShape.element)
                if (lastDayIsFriday or lastDayIsSaurday):
                    slide.shapes.element.remove(singleEventsShape[i].shape.element)
            else:
                singleEventsShape[i].shape.shapes.element.remove(singleEventsShape[i].titleShape.element)
                singleEventsShape[i].shape.shapes.element.remove(singleEventsShape[i].text1Shape.element)
                singleEventsShape[i].shape.shapes.element.remove(singleEventsShape[i].locShape.element)
                singleEventsShape[i].shape.shapes.element.remove(singleEventsShape[i].priceShape.element)
                singleEventsShape[i].shape.shapes.element.remove(singleEventsShape[i].freeShape.element)
                singleEventsShape[i].shape.shapes.element.remove(singleEventsShape[i].zoomShape.element)
        if not lastDayReached:
            if int(singleEventsShape[i].dateShape.text) == lastDayOfMonth:
                lastDayReached = True
            elif lastDayIsFriday and int(singleEventsShape[i].dateShape.text) == lastDayOfMonth - 1:
                lastDayReached = True
            elif lastDayIsSaurday and int(singleEventsShape[i].dateShape.text) == lastDayOfMonth - 2:
                lastDayReached = True


    if fridayEventsShape:
        for i in range(fridayCount, 5):
            slide.shapes.element.remove(fridayEventsShape[i].shape.element)

def main():
    try:
        metaData = MetaData()
        create_program_GUI()
    except Exception as e:
        if str(e)=="קובץ ה-cache אינו תקין":
            easygui.msgbox("שגיאה: \nאאי אפשר לקרוא נתונים מוזנים אחרונים. הקובץ \n\"cache.txt\"\nאינו תקין")
            easygui.msgbox(e)

if __name__ == "__main__":
    main()