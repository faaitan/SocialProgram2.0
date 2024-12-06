
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


firstMonthExcelEventsDictionary = {}
secondMonthExcelEventsDictionary = {}
firstMonthPeakExcelEventsDictionary = {}
secondMonthPeakExcelEventsDictionary = {}

excelEventsDictionary = None

def addToExcelEventDictionary(excelEvent, Dictionary):
    eventDayInMonth = excelEvent.day
    if eventDayInMonth in Dictionary:
        if len(Dictionary[eventDayInMonth]) >= 2:
            raise Exception("בתאריך " + excelEvent.date + " קיימים מעל לשני אירועים") #TODO: check this case
        Dictionary[eventDayInMonth].append(excelEvent)
    else:
        Dictionary[eventDayInMonth] = [excelEvent]


MAX_FIRST_MONTH_EVENT_COUNT = 31
MAX_SECOND_MONTH_EVENT_COUNT = 31
EXCEL_COULUMN_NUMBER = 9
FORBIDDEN_SENTENCE_HEBREW = "עלות פלוס: חינם | יאללה: חינם"
FORBIDDEN_SENTENCE_ARABIC = "السعر: شيكل"


class MetaData:
    def __init__ (self, firstMonthInteger = None, secondMonthInteger = None, firstMonthName = None, secondMonthName = None, zone = None, year = None, contactName = None, contactPhone = None, language = None,
    splitYear = False,
        hasCacheFile = False):
        self.firstMonthInteger = firstMonthInteger
        self.secondMonthInteger = secondMonthInteger
        self.firstMonthName = firstMonthName
        self.secondMonthName = secondMonthName
        self.zone = zone
        self.year=year
        self.contactName = contactName
        self.contactPhone = contactPhone
        self.language = language
        self.splitYear = splitYear
        self.hasCacheFile = hasCacheFile
        self.isFirstPresentation = True
        self.isSecondPresentation = False



# An object holding the fields of the Excel input event: date (datetime), day (string), hour (datetime.time), title (string), 
# text1 (string), location (string), price (string), isFree (bool), isZoom (bool) and month (int) of a planned event 
class ExcelEvent:
    def __init__(self, date, hour, title, location, price, community, link, eventType):
        self.date = date
        self.day = getDay(date)
        self.hour = hour
        self.title = title
        self.location = location
        self.price = price
        self.community = community
        self.link = link
        self.eventType = eventType
        self.month = getMonth(date)
        self.dayInMonth = getDay(date)

    #Comparison function, comparing Event objects
    #by date and hour (if dates are equal)
    def __lt__(self, other):
        return self.date.month <= other.date.month

    def __str__(self):
        return "date = %s, hour = %s, title = %s, month = %s"%(self.date.date(),self.hour, self.title, str(self.month))


class Community(Enum):
    SINGLES = 1
    TIULA = 2
    YOLO = 3
    WOMEN = 4
    GOLDERS = 5
    KULTURA = 6

class Communities:
    SinglesString = "סינגלס"
    Singles = Community.SINGLES
    TiulaString = "טיולא"
    Tiula = Community.TIULA
    YoloString = "YOLO"
    Yolo = Community.YOLO
    WomenString = "נשים"
    Women = Community.WOMEN
    GoldersString = "גולדרס"
    Golders = Community.GOLDERS
    KulturaString = "קולטורה"
    Kultura = Community.KULTURA

    communitiesArray = None
    communitiesStringArray = None

    def __init__(self):
        communitiesStringArray = [attr for atrr in dir(Communities) if (not attr.startswith('__') and "String" in attr)]
        communitiesArray = [attr for atrr in dir(Communities) if (not attr.startswith('__') and not "String" in attr)]

class EventType(Enum):
    REGULAR = 1
    WIDE = 2
    PEAK = 3

class EventTypes:
    def __init__(self):
        self.wideString = "רוחבי"
        self.wide = EventType.WIDE
        self.peakString = "ארוע שיא"
        self.peak = EventType.PEAK

class ShapeType(Enum):
    SINGLE = 1
    DOUBLE = 2
    PIC = 3
    OFF = 4


    

# An object holding the different shapes of the input powerpoint event shape: dateShape (shape), dayShape (shape), hourShape (shape), titleShape (shape), 
# text1Shape (shape), locationShape (shape) priceShape, freeShape (shape), zoomShape (shape)
# and shape (the root element holding them all together)
class SingleEventShape:
    def __init__(self, titleShape, locationShape, priceShape, tagSingles, tagTiula, tagYolo, tagWomen, tagGolders, tagKultura, countShape, countOffShape, dayShape, dayOffShape,
    spineBgShape, spineBgOffShape, bgHighlightShape, bgPicShape, bgOffShape, bgShape, shape, isTreated = False):
        self.titleShape = titleShape
        self.locationShape = locationShape
        self.priceShape = priceShape
        self.tagSingles = tagSingles
        self.tagTiula = tagTiula
        self.tagYolo = tagYolo
        self.tagWomen = tagWomen
        self.tagGolders = tagGolders
        self.tagKultura = tagKultura
        self.countShape = countShape
        self.countOffShape = countOffShape
        self.dayShape = dayShape
        self.dayOffShape = dayOffShape
        self.spineBgShape = spineBgShape
        self.spineBgOffShape = spineBgOffShape
        self.bgHighlightShape = bgHighlightShape
        self.bgPicShape = bgPicShape
        self.bgOffShape = bgOffShape
        self.bgShape = bgShape
        self.shape = shape

class DoubleEventShape:
    def __init__(self, titleShape1, locationShape1, priceShape1, tagSingles1, tagTiula1, tagYolo1, tagWomen1, tagGolders1, tagKultura1,
     titleShape2, locationShape2, priceShape2, tagSingles2, tagTiula2, tagYolo2, tagWomen2, tagGolders2, tagKultura2, countShape, countShape, countOffShape, dayShape, dayOffShape, 
     spineBgShape, spineBgOffShape, bgHighlightShape, bgPicShape, bgOffShape, bgShape, shape, isTreated = False):
        self.titleShape1 = titleShape1
        self.locationShape1 = locationShape1
        self.priceShape1 = priceShape1
        self.tagSingles1 = tagSingles1
        self.tagTiula1 = tagTiula1
        self.tagYolo1 = tagYolo1
        self.tagWomen1 = tagWomen1
        self.tagGolders1 = tagGolders1
        self.tagKultura1 = tagKultura1
        self.titleShape2 = titleShape2
        self.locationShape2 = locationShape2
        self.priceShape2 = priceShape2
        self.tagSingles2 = tagSingles2
        self.tagTiula2 = tagTiula2
        self.tagYolo2 = tagYolo2
        self.tagWomen2 = tagWomen2
        self.tagGolders2 = tagGolders2
        self.tagKultura2 = tagKultura2
        self.countShape = countShape
        self.countOffShape = countOffShape
        self.dayShape = dayShape
        self.dayOffShape = dayOffShape
        self.spineBgShape = spineBgShape
        self.spineBgOffShape = spineBgOffShape
        self.bgHighlightShape = bgHighlightShape
        self.bgPicShape = bgPicShape
        self.bgOffShape = bgOffShape
        self.bgShape = bgShape
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
    
        getEventsFromExcel(sheet)
          
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
    excelMonths = []
    dictionary = firstMonthExcelEventsDictionary
    peakEventsDictionary = firstMonthPeakExcelEventsDictionary
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

        day = sheet.cell(row=rowIndex,column=2).value
        if day is None: 
            raise Exception("שורה \n"+str(rowIndex)+"\n עמודה: \n"+str(2)+"\nהערך ריק ")

        hourValue = sheet.cell(row=rowIndex,column=3).value

        title = sheet.cell(row=rowIndex,column=4).value
        if title is None:
            title = ""

        location = sheet.cell(row=rowIndex,column=6).value
        if location is None:
            location = ""

        price = sheet.cell(row=rowIndex, column=7).value
        if price is None:
            price = ""

        communityString = sheet.cell(row = rowIndex, column = 8).value
        community = getCommunityFromString(communityString) #TODO: write this one

        link = sheet.cell(row = rowIndex, column = 9).value
        #TODO: check link validity

        eventTypeString = sheet.cell(row = rowIndex, column = 10).value
        eventType = getEventTypeFromString(eventTypeString) #TODO: write this one too


        excelEvent = ExcelEvent(str(date), str(hourValue), title, location, price, community, link, eventType)

        if excelEvent.month not in excelMonths:
            excelMonths.append(excelEvent.month)
            if len(excelMonths) == 2:
                dictionary = secondMonthExcelEventsDictionary
                peakEventsDictionary = secondMonthPeakExcelEventsDictionary
            elif len(excelMonths) > 2:
                raise Exception("קובץ האקסל מכיל מעל לשני חודשים")

        if excelEvent.eventType is EventType.PEAK:
            peakEventsDictionary[excelEvent.Community] = excelEvent
        else:
            addToExcelEventDictionary(excelEvent, dictionary)
             

    if len(excelMonths) < 2:
        raise Exception("קובץ האקסל מכיל פחות משני חודשים")

    if not (12 in excelMonths and 1 in excelMonths):
        excelMonths.sort()
    else:
        metaData.increaseYear = True

    metaData.firstMonthInteger = excelMonths[0]
    metaData.secondMonthInteger = excelMonths[1]

    excelEventsDictionary = { metaData.firstMonthInteger : firstMonthExcelEventsDictionary, metaData.secondMonthInteger : secondMonthExcelEventsDictionary }



def rowIsEmpty(sheet, rowIndex, maxColIndex):
    for i in range(1,maxColIndex+1):
        if sheet.cell(row=rowIndex,column=i).value !=None:
            return False
        else:
            if sheet.cell(row = rowIndex, column = 7) == None and sheet.cell(row=rowIndex, column =8) == False:
                return False
            return True


def getCommunityFromString(communityString):
    communities = Communities()
    if communities.SinglesString in communityString:
        return communities.Singles
    elif communities.TiulaString in communityString:
        return communities.Tiula
    elif communities.YoloString in communityString:
        return communities.Yolo
    elif communities.WomenString in communityString:
        return communities.Women
    elif communities.GoldersString in communityString:
        return communities.Golders
    elif communities.KulturaString in communityString:
        return communities.Kultura
    else:
        return None


def getEventTypeFromString(eventTypeString):
    eventTypes = EventTypes()
    if eventTypes.wideString in eventTypeString:
        return eventType.wide
    elif eventTypes.peakString in eventTypeString:
        return eventType.peak
    else:
        return None



 
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

def get_slide_shapes(slide):
    singles, doubles, header_shape = find_groups(slide.shapes)
    header_shape_dict = {}

    singles = reversed(singles) #Maybe unnecessary
    doubles = reversed(doubles)

    single_titles = []
    single_locations = []
    single_prices = [] 
    single_Singles_tags = []
    single_Tiula_tags = []
    single_Yolo_tags = []
    single_Women_tags = []
    single_Golders_tags = []
    single_Kultura_tags = []
    single_counts = []
    single_count_offs = []
    single_days = []
    single_day_offs = []
    single_spine_bgs = []
    single_spine_bg_offs = []
    single_bg_highlights = []
    single_bg_pics = []
    single_bg_offs = []
    single_bgs = []
    single_shapes = []

    double_titles1 = []
    double_locations1 = []
    double_prices1 = [] 
    double_Singles_tags1 = []
    double_Tiula_tags1 = []
    double_Yolo_tags1 = []
    double_Women_tags1 = []
    double_Golders_tags1 = []
    double_Kultura_tags1 = []
    double_titles2 = []
    double_locations2 = []
    double_prices2 = [] 
    double_Singles_tags2 = []
    double_Tiula_tags2 = []
    double_Yolo_tags2 = []
    double_Women_tags2 = []
    double_Golders_tags2 = []
    double_Kultura_tags2 = []
    double_counts = []
    double_count_offs = []
    double_days = []
    double_day_offs = []
    double_spine_bgs = []
    double_spine_bg_offs = []
    double_bg_highlights = []
    double_bg_pics = []
    double_bg_offs = []
    double_bgs = []
    double_shapes = []


    for g in singles:
        single_shapes.append(g)
        for shape in iter_textable_shapes(g.shapes):
            if shape.name == 'TITLE':
                single_titles.append(shape)
            elif shape.name == 'LOCATION':
                single_locations.append(shape)
            elif shape.name == 'PRICE':
                single_prices.append(shape)
            elif shape.name == 'TAG SINGLES 1':
                single_Singles_tags.append(shape)
            elif shape.name == 'TAG TIULA 1':
                single_Tiula_tags.append(shape)
            elif shape.name == 'TAG YOLO 1':
                single_Yolo_tags.append(shape)
            elif shape.name == 'TAG WOMEN 1':
                single_Women_tags.append(shape)
            elif shape.name == 'TAG GOLDERS 1':
                single_Golders_tags.append(shape)
            elif shape.name == 'TAG KULTURA 1':
                single_Kultura_tags.append(shape)
            elif shape.name == 'COUNT':
                single_counts.append(shape)
            elif shape.name == 'COUNT OFF':
                single_count_offs.append(shape)
            elif shape.name == 'DAY':
                single_days.append(shape)
            elif shape.name == 'DAY OFF':
                single_day_offs.append(shape)
            elif shape.name == 'SPINE BG':
                single_sping_bgs.append(shape)
            elif shape.name == 'SPINE BG OFF':
                single_spine_bg_offs.append(shape)
            elif shape.name == 'BG HIGHLIGHT':
                single_bg_highlights.append(shape)
            elif shape.name == 'BG PIC':
                single_bg_pics.append(shape)
            elif shape.name == 'BG OFF':
                single_bg_offs.append(shape)
            elif shape.name == 'BG':
                single_bgs.append(shape)

    for g in doubles:
        double_shapes.append(g)
        for shape in iter_textable_shapes(g.shapes):
            if shape.name == 'TITLE 1':
                double_titles1.append(shape)
            elif shape.name == 'LOCATION 1':
                double_locations1.append(shape)
            elif shape.name == 'PRICE 1':
                double_prices1.append(shape)
            elif shape.name == 'TAG SINGLES 1':
                double_Singles_tags1.append(shape)
            elif shape.name == 'TAG TIULA 1':
                double_Tiula_tags1.append(shape)
            elif shape.name == 'TAG YOLO 1':
                double_Yolo_tags1.append(shape)
            elif shape.name == 'TAG WOMEN 1':
                double_Women_tags1.append(shape)
            elif shape.name == 'TAG GOLDERS 1':
                double_Golders_tags1.append(shape)
            elif shape.name == 'TAG KULTURA 1':
                double_Kultura_tags1.append(shape)
            elif shape.name == 'TITLE 2':
                double_titles2.append(shape)
            elif shape.name == 'LOCATION 2':
                double_locations2.append(shape)
            elif shape.name == 'PRICE 2':
                double_prices2.append(shape)
            elif shape.name == 'TAG SINGLES 2':
                double_Singles_tags2.append(shape)
            elif shape.name == 'TAG TIULA 2':
                double_Tiula_tags2.append(shape)
            elif shape.name == 'TAG YOLO 2':
                double_Yolo_tags2.append(shape)
            elif shape.name == 'TAG WOMEN 2':
                double_Women_tags2.append(shape)
            elif shape.name == 'TAG GOLDERS 2':
                double_Golders_tags2.append(shape)
            elif shape.name == 'TAG KULTURA 2':
                double_Kultura_tags2.append(shape)
            elif shape.name == 'COUNT':
                double_counts.append(shape)
            elif shape.name == 'COUNT OFF':
                double_count_offs.append(shape)
            elif shape.name == 'DAY':
                double_days.append(shape)
            elif shape.name == 'DAY OFF':
                double_day_offs.append(shape)
            elif shape.name == 'SPINE BG':
                double_sping_bgs.append(shape)
            elif shape.name == 'SPINE BG OFF':
                double_spine_bg_offs.append(shape)
            elif shape.name == 'BG HIGHLIGHT':
                double_bg_highlights.append(shape)
            elif shape.name == 'BG PIC':
                double_bg_pics.append(shape)
            elif shape.name == 'BG OFF':
                double_bg_offs.append(shape)
            elif shape.name == 'BG':
                double_bgs.append(shape)

    for shape in iter_textable_shapes(header_shape):
        if shape.name == 'ZONE':
            header_shape_dict["zoneShape"] = shape
        elif shape.name == 'MONTH':
            header_shape_dict["monthShape"] = shape
        elif shape.name == 'YEAR':
            header_shape_dict["yearShape"] = shape


    # textable_shapes = list(iter_textframed_shapes(slide.shapes))
    # ordered_textable_shapes = sorted(
    #     textable_shapes, key=lambda shape: (shape.top, shape.left)
    # )

    # for shape in ordered_textable_shapes:
    #     if shape.name.startswith('MONTH1'):
    #         createMonthObjects(shape, months, 0)
    #     if shape.name.startswith('MONTH2'):
    #         createMonthObjects(shape, months, 1)
    #     if shape.name == "AREA":
    #         createAreaObjects(shape, area)
    #     if shape.name == "CONTACT":
    #         createContactObject(shape, contact)

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

    # def __init__(self, titleShape, locationShape, priceShape, tagSingles, tagTiula, tagYolo, tagWomen, tagGolders, tagKultura, countShape, countOffShape, dayShape, dayOffShape,
    # spineBgShape, spineBgOffShape, bgHighlightShape, bgPicShape, bgOffShape, bgShape, shape, isTreated = False):

    for i in range(36):
        single_event_shape = SingleEventShape(single_titles[i], single_locations[i], single_prices[i], single_Singles_tags[i], single_Tiula_tags[i], single_Yolo_tags[i], single_Women_tags[i],
            single_Golders_tags[i], single_Kultura_tags[i], single_counts[i], single_count_offs[i], single_days[i], single_day_offs[i], single_spine_bgs[i], single_spine_bg_offs[i],
            single_bg_highlights[i], single_bg_pics[i], single_bg_offs[i], single_bgs[i], single_shapes[i])
        single_event_shapes.append(single_event_shape)

        double_event_shape = DoubleEventShape(double_titles1[i], double_locations1[i], double_prices1[i], double_Singles_tags1[i], double_Tiula_tags1[i], double_Yolo_tags1[i], double_Women_tags1[i],
            double_Golders_tags1[i], double_Kultura_tags1[i], double_titles2[i], double_locations2[i], double_prices2[i], double_Singles_tags2[i], double_Tiula_tags2[i], double_Yolo_tags2[i],
            double_Women_tags2[i], double_Golders_tags2[i], double_Kultura_tags2[i], double_counts[i], double_count_offs[i], double_days[i], double_day_offs[i], double_spine_bgs[i],
            double_spine_bg_offs[i],double_bg_highlights[i], double_bg_pics[i], double_bg_offs[i], double_bgs[i], double_shapes[i])
        double_event_shapes.append(double_event_shape)

    return single_event_shapes, double_event_shapes, header_shape_dict


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
    header = None

    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            if shape.name.startswith("DOUBLE"):
                doubles.append(shape)
            elif shape.name.startswith("ELEMENT"):
                singles.append(shape)
            elif shape.name.startswith("HEADER"):
                header = shape
    return singles, doubles, header


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

        btn4 = tk.Button(text="צרו תוכנית דו-חודשית!", command= lambda: createPptxPlans(assignFirstMonthTextbox.get(), assignSecondMonthTextbox.get(), assignLocationTextbox.get(), assignContactTextbox.get(),assignYearTextbox.get(), language.get()))
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

def createPptxPlans(month1, month2, area, contact, year, language):
    if len(year)!=4:
            raise Exception("שדה השנה צריך להכיל 4 תוים בדיוק. למשל: 2023")
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

        metaData.firstMonthName = month1
        metaData.secondMonthName = month2
        metaData.year = year
        metaData.zone = area
        metaData.language = language


        readExcel(excelFilePath)

        #createPptxPlan(pptxFilePath)
        try:
            presentation = Presentation(pptxFilePath)

            first_slide = presentation.slides[0]
            second_slide = presentation.slides[1] #TODO: wrap in try catch and throw excecption?

            processFirstSlide(first_slide, isFirstPresentation)
            processSecondSlide(second_slide, isFirstPresentation)

        except IndexError:
            raise Exception("תבנית ה-powerpoint מכילה פחות מ-2 שקפים")
        

        

        single_event_shapes, double_event_shapes, friday_event_shapes = get_text_boxes(slide, months_names, area, contact)

        


        files = [('Powerpoint files', '*.pptx')]
        file = asksaveasfile(filetypes = files, defaultextension = files)
        try:
            presentation.save(file.name)
            easygui.msgbox("הקובץ נוצר בהצלחה")
        except PermissionError: 
            easygui.msgbox("הקובץ \n"+ file.name+"\n פתוח. יש לסגור אותו בעת הרצת התוכנית")

    except Exception as e:
        easygui.msgbox("שגיאה :"+ str(e))

    


def processFirstSlide(slide, isFirstPresentation):
    pass


def processSecondSlide(slide, isFirstPresentation):
    single_event_shapes, double_event_shapes, header_shape = get_slide_shapes(slide)

    calendar.setfirstweekday(6)

    month = None 
    if isFirstPresentation:
        month = metaData.firstMonthInteger
    else:
        month = metaData.secondMonthInteger
    
    if isFirstPresentation:

        calendar.setfirstweekday(6)

        createCalendarDates(single_event_shapes, double_event_shapes, month)

        writeTextToTextboxes(slide, single_event_shapes, double_event_shapes, header_shape, month)



def clearTextboxText(text_frame):
    for i in range(len(text_frame.paragraphs)):
        para = text_frame.paragraphs[i]
        for j in range (len(para.runs)):
            para.runs[j].text=""

def get_number_of_shape(year, month, day):
    numOfDaysInMonth = calendar.monthrange(year, month)[1]
    x = np.array(calendar.monthcalendar(year, month))
    week_of_month = np.where(x==day)[0][0] # 0 is first week
    day_of_week = np.where(x == day)[1][0]+1 # 1 is Sunday
    if day_of_week > 6: 
        return -1
    first_day_of_month = np.where(x == 1)[1][0] + 1 # for removing first week that starts on Friday or Saturday
    if first_day_of_month <= 6: 
        return week_of_month*6 + day_of_week
    else:
        return (week_of_month - 1)*6 + day_of_week


def createCalendarDates(slide, singleEventShapes, doubleEventShapes, month, increaseYear = False):
    year = int(metaData.year)
    if increaseYear:
        year += 1

    processFirstDayOffs(slide, singleEventShapes, doubleEventShapes, year, month)

    numOfDaysInMonth = calendar.monthrange(year, month)[1]
    

    for i in range(1, numOfDaysInMonth+1):
        num_of_shape = get_number_of_shape(year, month, i) - 1
        if num_of_shape < 0:
            continue
        else:
            single_event_shape = singleEventShapes[num_of_shape]
            double_event_shape = doubleEventShapes[num_of_shape]

            removeShapeOffElements(single_event_shape)
            writeTextToTextbox(single_event_shape.countShape.text_frame, str(i))
            writeTextToTextbox(single_event_shape.dayShape.text_frame, hebrew_letter_of_day(get_week_day(i, month, year))) #TODO: deal with arabic

            removeShapeOffElements(double_event_shape)
            writeTextToTextbox(double_event_shape.countShape.text_frame, str(i))
            writeTextToTextbox(double_event_shape.dayShape.text_frame, hebrew_letter_of_day(get_week_day(i, month, year))) #TODO: deal with arabic

    processLastDayOffs(slide, singleEventShapes, doubleEventShapes, year, month)


def removeShapeOffElements(shape):
    slide.shapes.element.remove(shape.countOffShape.shape.element)
    slide.shapes.element.remove(shape.dayOffShape.shape.element)
    slide.shapes.element.remove(shape.spineBgOffShape.shape.element)
    slide.shapes.element.remove(shape.bgOffShape.shape.element)


def processFirstDayOffs(slide, singleEventShapes, doubleEventShapes, year, month):

    num_of_days_in_previous_month = calendar.monthrange(year, month - 1)[1] if month != 1 else calendar.monthrange(year - 1, 12)[1]

    if month != 1:
        num_of_days_in_previous_month = calendar.monthrange(year, month - 1)[1]
    else:
        num_of_days_in_previous_month = calendar.monthrange(year - 1, 12)[1]


    first_day_of_week1 = get_number_of_shape(year, month, 1) - 1


    i = first_day_of_week - 1
    j = num_of_days_in_previous_month
    while i >= 0:
        single_event_shape = singleEventShapes[i]
        double_event_shape = doubleEventShapes[i]

        writeTextToTextbox(single_event_shape.countOffShape.text_frame, str(j))
        writeTextToTextbox(single_event_shape.dayOffShape.text_frame, hebrew_letter_of_day(get_week_day(j, month, year))) #TODO: deal with arabic
 
        treat_off_shape(slide, single_event_shape, double_event_shape)

        i -= 1
        j -= 1


def processLastDayOffs(slide, singleEventShapes, doubleEventShapes, year, month):

    num_of_days_in_month = calendar.monthrange(year, month)[1]
    last_shape_in_month_index = get_number_of_shape(year, month, num_of_days_in_month)

    i = last_shape_in_month_index
    j = 1
    while i < 36:
        single_event_shape = singleEventShapes[i]
        double_event_shape = doubleEventShapes[i]

        writeTextToTextbox(single_event_shape.countOffShape.text_frame, str(j))
        writeTextToTextbox(single_event_shape.countOffShape.text_frame, hebrew_letter_of_day(get_week_day(j, month, year))) #TODO: deal with arabic

        i += 1
        j += 1



def treat_off_shape(slide, single_event_shape, double_event_shape):

    slide.shapes.element.remove(double_event_shape.shape.element)

    slide.shapes.element.remove(single_event_shape.titleShape.shape.element)
    slide.shapes.element.remove(single_event_shape.locationShape.shape.element)
    slide.shapes.element.remove(single_event_shapes.priceShape.shape.element)
    slide.shapes.element.remove(single_event_shape.tagSinglesShape.shape.element)
    slide.shapes.element.remove(single_event_shape.tagTiulaShape.shape.element)
    slide.shapes.element.remove(single_event_shape.tagYoloShape.shape.element)
    slide.shapes.element.remove(single_event_shape.tagWomenShape.shape.element)
    slide.shapes.element.remove(single_event_shape.tagGoldersShape.shape.element)
    slide.shapes.element.remove(single_event_shape.tagKulturaShape.shape.element)
    slide.shapes.element.remove(single_event_shape.countShape.shape.element)
    slide.shapes.element.remove(single_event_shape.dayShape.shape.element)
    slide.shapes.element.remove(single_event_shape.spineBgShape.shape.element)
    slide.shapes.element.remove(single_event_shape.bgHighlightShape.shape.element)
    slide.shapes.element.remove(single_event_shape.bgPicShape.shape.element)
    slide.shapes.element.remove(single_event_shape.bgShape.shape.element)


def get_week_day(day, month, year):
    date = datetime.datetime(year, month, day)
    return date.weekday()


def hebrew_letter_of_day(weekday_int):
    match weekday_int:
        case 0:
            return  "ב"
        case 1:
            return "ג"
        case 2:
            return "ד"
        case 3:
            return "ה"
        case 4:
            return "ו"
        case 5:
            raise Exception("Shouldn't get Saturday weekday. You're doing something wrong!")
        case 6:
            return "א"


def writeTextToTextboxes(slide, singleEventShapes, doubleEventShapes, headerShape, month, increaseYear = False):
    year = int(metaData.year)
    if increaseYear:
        year += 1

    numOfDaysInMonth = calendar.monthrange(year, numOfMonth)[1]


    single_event_shape, double_event_shape = None

    for i in range(1, numOfDaysInMonth + 1):

        shape_index = get_number_of_shape(year, month, i) - 1

        if shape_index >= 0:

            single_event_shape = singleEventShapes[shape_index]
            double_event_shepe = doubleEventShapes[shape_index]

            shape_day_in_month = int(single_event_shape.countShape.text_frame.paragraphs[0].runs[0].text)

            type = get_shape_type(shape_day_in_month, month)

            match type:
                case ShapeType.SINGLE:
                    treatSingleShape(slide, single_event_shape, double_event_shape, shape_day_in_month, month)
                case ShapeType.DOUBLE:
                    treatDoubleShape(slide, single_event_shape, double_event_shape, shape_day_in_month, month)
                case ShapeType.PIC:
                    treatPicShape(slide, single_event_shape, double_event_shape)

    writeTextToTextbox(header_shape[zone].text_frame, metaData.area)
    writeTextToTextbox(header_shape[month].text_frame, metaData.month)
    writeTextToTextbox(header_shape[year].text_frame, metaData.year)




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

def get_shape_type(day_in_month, month):
    excelEvents = excelEventsDictionary[month]
    if day_in_month in excelEvents:
        excelEventArray = excelEvent[dayInMonth]
        if len(excelEventArray) == 1:
            return ShapeType.SINGLE
        elif len(excelEventArray) == 2:
            return ShapeType.DOUBLE
    else:
        return ShapeType.PIC


def treatSingleShape(slide, single_event_shape, double_event_shape, shape_day_in_month, month):
    excelEvents = excelEventsDictionary[month]
    excelEvent = excelEvents[shape_day_in_month]

    slide.shapes.element.remove(double_event_shape.shape.element)
    treatTags(excelEvent, single_event_shape) #TODO: implement me

    if excelEvent.shapeType is shapeType.REGULAR:
        slide.shapes.element.remove(single_event_shape.bgPicShape.shape.element)
        slide.shapes.element.remove(single_event_shape.bgHighlightShape.shape.element)
    elif excelEvent.shapeType is shapeType.WIDE:
        slide.shapes.element.remove(single_event_shape.bgPicShape.shape.element)
        slide.shapes.element.remove(single_event_shape.bgShape.shape.element)

    writeTextToTextbox(single_event_shape.titleShape.text_frame, excelEvent.hour + " - " + excelEvent.title) #TODO: make the title an hyperlink
    writeTextToTextbox(single_event_shape.locationShape.text_frame, excelEvent.location)
    writeTextToTextbox(single_event_shape.priceShape.text_frame, excelEvent.price)


def treatDoubleShape(slide, single_event_shape, double_event_shape, shape_day_in_month, month):
    excelEvents = excelEventsDictionary[month]
    firstExcelEvent = excelEvents[shape_day_in_month][0]
    secondExcelEvent = excelEvents[shape_day_in_month][1]

    slide.shapes.element.remove(single_event_shape.shape.element)
    treatTags(firstExcelEvent, secondExcelEvent, double_event_shape) #TODO: implement me

    slide.shapes.element.remove(double_event_shape.bgPicShape.shape.element)

    writeTextToTextbox(double_event_shape.titleShape1.text_frame, firstExcelEvent.hour + " - " + firstExcelEvent.title) #TODO: make the title an hyperlink
    writeTextToTextbox(double_event_shape.locationShape1.text_frame, firstExcelEvent.location)
    writeTextToTextbox(double_event_shape.priceShape1.text_frame, firstExcelEvent.price)
    writeTextToTextbox(double_event_shape.titleShape2.text_frame, secondExcelEvent.hour + " - " + secondExcelEvent.title) #TODO: make the title an hyperlink
    writeTextToTextbox(double_event_shape.locationShape2.text_frame, secondExcelEvent.location)
    writeTextToTextbox(double_event_shape.priceShape2.text_frame, secondExcelEvent.price)


def treatPicShape(slide, single_event_shape, double_event_shape):
    slide.shapes.element.remove(double_event_shape.shape.element)
    slide.shapes.element.remove(single_event_shape.bgShape.shape.element)
    slide.shapes.element.remove(single_event_shape.bgHighlightShapeShape.shape.element)
    slide.shapes.element.remove(single_event_shape.titleShape.shape.element)
    slide.shapes.element.remove(single_event_shape.locationShape.shape.element)
    slide.shapes.element.remove(single_event_shape.priceShape.shape.element)
    #TODO: implement random picture for bgPicShape


def writeTextToTextbox(shape_text_frame, text):
    text_frame_paragraphs = shape_text_frame.paragraphs
    clearTextboxText(text_frame_paragraph)
    text_frame_paragraph = shape_text_frame.paragraphs[0]
    text_frame_paragraph.alignment = PP_ALIGN.RIGHT
    run = text_frame_paragraph.runs[0]
    run.text = text


def treatTags(slide, excelEvent, secondExcelEvent = None, single_event_shape, double_event_shape = None):
    excelEventCommunity = excelEvent.community
    secondExcelEventCommunity = None
    if secondExcelEvent != None:
        secondExcelEventCommunity = secondExcelEvent.community

    allCommunities = Communities.communitiesArray
    allCommunitiesStrings = Communities.communitiesStringArray

    for i in range(len(allCommunities)):
        if double_event_shape == None:
            if allCommunities[i] is not excelEventCommunity:
                commString = allCommunitiesStrings[i]
                suffix = "1" if double_event_shape != None else ""
                attrToRemove = getattr(singleEventsShape, "tag" + commString.removesuffix('String') + suffix)
                slide.shapes.element.remove(attrToRemove.shape.element)
        if double_event_shape != None:
            if allCommunities[i] is not secondExcelEventCommunity:
                commString = allCommunitiesStrings[i]
                suffix = 2
                attrToRemove = getattr(singleEventsShape, "tag" + commString.removesuffix('String') + suffix)
                slide.shapes.element.remove(attrToRemove.shape.element)


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