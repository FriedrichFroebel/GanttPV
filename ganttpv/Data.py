#!/usr/bin/env python
# Data Tables - includes update logic, date routines, and gantt calculations

# Copyright 2004, 2005 by Brian C. Christensen

#    This file is part of GanttPV.
#
#    GanttPV is free software; you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation; either version 2 of the License, or
#    (at your option) any later version.
#
#    GanttPV is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#
#    You should have received a copy of the GNU General Public License
#    along with GanttPV; if not, write to the Free Software
#    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA

# 040407 - first version w/ FillTable and Update to create some sample data
# 040408 - added date conversion tables
# 040409 - added GanttCalculation
# 040410 - Update will now update
# 040412 - added CheckChange and SetUndo; added OpenReports
# 040413 - added SetEmptyData
# 040414 - renamed file to Data; added load and store routines; corrections to Update;
#          reserved ReportID's 1 & 2 for 'Project/Report List' and 'Resource List'
# 040415 - moved menu adjust logic here from Main.py; made various changes to correct or improve logic
# 040416 - changed gantt calculation to process all projects; wrote Undo and Redo routines
# 040417 - added optional parameter to Update to specify whether change should be added to UndoStack;
#          in CheckChange now sets ChangedReport flag if zzStatus appears; moved flags and variables to
#          front of file this; added MakeReady to do all the necessary computations after a new database
#          is loaded; added AdjustReportRows to add report rows when records are added
# 040419 - reworked AdjustReportRows logic
# 040420 - added routines to open and close reports; revised MakeReady
# 040422 - fix so rows added by AdjustReportRows will always have a link to their parent; added row type colors
#          for Project/Report frame
# 040423 - added ReorderReportRows; moved AutoGantt to Option; added ConfirmScripts to Option
# 040424 - added ScriptDirectory to Option; added LoadOption and SaveOption; added GetRowList;
#          fixed reset error in date tables;
# 040426 - fixed Recalculate to check ChangedReport; added check for None to GetRowList and ReorderReportRows;
#          added table Assignment; in GanttCalculation ignore deleted Dependency records
# 040427 - in Resource table changed 'LongName' to 'Name'; added Width to ReportColumn; revised ReportColumn fields
# 040503 - added ReportType and ColumnType tables; moved UpdateDataPointers here from Main and GanttReport; added
#          "Prerequisite" synonym for "Task" in Database
# 040504 - added default Width to ColumnType; in gantt calculation, will ignore tasks w/o projects;
#          added GetColumnList and ReorderReportColumns; added GetToday
# 040505 - fixed "Computed"/"Calculated" inconsistency; changed some column type labels
# 040506 - add info line at beginning of saved files; added BaseBar to chart options
# 040508 - start w/ empty data instead of sample data
# 040510 - added header line to option file
# 040512 - added 'T' (source table) to ColumnType; several bug fixes; fixed handling of holidays in date tables
# 040520 - fixed bug where Assignment pointer was not set for Loaded files
# 040605 - fixed bug in ReorderReportColumns to properly handle zzStatus == None
# 040715 - Pierre_Rouleau@impathnetworks.com: removed all tabs, now use 4-space indentation level to comply with Official Python Guideline.
# 040815 - FileName not set to None on New (because not on globals list)
# 040906 - add "project name / report name" to report titles
# 040928 - Alexander - when making DateConv, ignore incorrectly formatted dates; in other places, ignore dates not in DateConv
# 041001 - added FindID and FindIDs
# 041009 - tightened edits on FirstDate and LastDate; added ValidDate() routine
# 041012 - moved AddTable, AddRow, and AddReportType here from "Add Earned Value Tracking.py"
# 041031 - don't automatically add report rows for "deleted" records
# 041203 - GetColumnHeader and GetCellValue logic moved here from GanttReport
# 041231 - added xref to find first day of next month; changed GetColumnHead to work for months and quarters; added routine to add months to a date index
# 050101 - added hours units
# 050104 - added backwards pass and float calculations
# 050328 - derive parent dates from children
# 050329 - support multi-value list type columns (for predecessors, successors, children, and resource names)
# 050402 - handle task w/ self for parent
# 050407 - Alexander - close deleted reports and prevent them from opening
# 050409 - Alexander - added GetModuleNames
# 050503 - Alexander - added App, for program quiting; and ActiveReport, for script-running and window-switching
# 050423 - moved GetPeriodInfo logic to calculate period start and hours to Data from GanttReport.py; added GetColumnDate; save "SubtaskCount" in Task table 
# 050504 - Alexander - moved script-running logic here; added logic to prevent no-value entries in the database; added logic to update the Window menu.
# 050513 - Alexander - revised some dictionary fetches to ignore invalid keys (instead of raising an exception); tightened _Do logic;
# 050519 - Brian - use TaskID instead of ParentTaskID to designate Task parent.

# import calendar
import datetime
import cPickle
import wx, os
import Menu 
import GanttReport
# import Main

debug = 1
if debug: print "load Data.py"
# On making changes in GanttPV
# 1- Use Update for all changes
# 2- Make change will decide the impact of the changes
# 3- Use SetUndo to tell GanttPV to fix anything in response to the changes

def GetModuleNames():
    """ Return the GanttPV modules as a namespace dictionary.

    This dictionary should be passed to any scripts run with execfile.
    """
    import Data, GanttPV, GanttReport, ID, Menu, UI, wx
    return locals()

App = 0           # the application itself
ActiveReport = 1  # ReportID of most recently active frame

# Clients should treat these tables as read only data
Database = {}           # will contain pointers to all of these tables
ChangedData = False     # True if database needs to be saved
FileName = None         # Filename of current database
Path = ""               # Directory containing executed program, look for option file here

def CreateEmptyTables():
    global Database, Project, Task, Dependency, Report, ReportColumn, ReportRow
    global Resource, Assignment, Holiday, ReportType, ColumnType, OtherData, Other, NextID
    Database = {}
    Project = {}
    Task = {}
    Dependency = {}
    Report = {}
    ReportColumn = {}
    ReportRow = {}
    Resource = {}
    Assignment = {}
    Holiday = {}
    ReportType = {}
    ColumnType = {}
    OtherData = {}  # used in updates
    Other = {}
    NextID = {}

# Successor = { }  # dependency xref (used?)
# Predecessor = { }

# Date conversion tables
DateConv = {}   # usage: index = DateConv['2004-01-01']
DateIndex = []  # usage: date = DateIndex[1]
DateInfo = []   # dayhours, cumhours, dow = DateInfo[1]
DateNextMonth = {} # usage: index = DateNextMonth['2004-01']  # index is first day of next month

# Time units conversion
WeekToHour = 40
DayToHour = 8
AllowHourToDay = True

# save impact of change until it can be addressed
ChangedCalendar = False  # will affect calendar hours tables (gantt will have to be redone, too)
ChangedSchedule = False  # gantt calculation must be redone
ChangedReport = False    # report layout needs to be redone
ChangedRow = False       # a record was added or deleted, the rows on a report may be affected

UndoStack = []
RedoStack = []
OpenReports = {}  # keys are report id's of open reports. Main is always "1"; Report Options is always "2"

# these options are user preferences
Option = {  'AutoGantt' : True,  # Delay calculation of calendar and schedule changes
            'ConfirmScripts': True,  # scripts should display a confirm dialog
            'ScriptDirectory': None,  # where script directory is located
#           'HelpfulWarnings': True,  # display warning popups  -- do I want to do this?
# ProjecReport Colors
            'ParentColor': ( 153, 153, 255 ),
            'ChildColor': ( 238, 238, 238 ),
            'HiddenColor': ( 204, 204, 204 ),
            'DeletedColor': ( 255, 102, 204 ),
# Gantt Chart Colors
# background
            'WorkDay': (230, 230, 230),
#           'WorkDay': (221, 221, 221),
            'WorkDaySelected': (153, 204, 255),     # make this lighter    maybe use white only for the selected column??
            'NotWorkDay': (204, 204, 204),                                  # more light focused on the selection
            'NotWorkDaySelected': (136, 187, 238),
# bars
            'PlanBar': (0, 153, 102),   # green             # maybe use same color for selected & not
            'PlanBarSelected': (0, 153, 102),
            'ActualBar': (0, 0, 255),    # yellow??
            'ActualBarSelected': (0, 0, 255),
            'BaseBar': (0, 255, 0),   # green??
            'BaseBarSelected': (0, 255, 0),
            'CompletionBar': None,          # based on BaseBar or PlanBar
            'CompletionBarSelected': None,

#           'ResourceProblem': None,
#           'ResourceProblemSelected': None,
        }

# -------------- Setup Empty and Sample databases

# reserved table column names: 'Table', 'ID', 'zzStatus', anything ending with 'ID", anything starting with 'xx'

def SetOther():
    """ These are the values that only occur once in the database """
    global Other
    Other = {   'ID' : 1,  # every record needs an ID
                'BaseDate' : None,  # date numbers are relative to this date
                'WeekHours' : (8, 8, 8, 8, 8, 0, 0),  # hours per day MTWHFSS
            }
    Database['OtherData'] = OtherData  # for consistent access
    OtherData[1] = Other  # for use w/ Update
    Database['Other'] = Other  # for simpler access
    Database['NextID'] = NextID
    Database['Prerequisite'] = Database['Task']  # create a synonym for Task
# don't add these until Purge has been updated to recognize them
#    Database['SpecialTables'] = ['Other', 'NextID', 'Prerequisite', 'SpecialTables', 'Indices', 'PriorMeasurement']
#    Database['Indices'] = { 
#        'Dependency': ('TaskID', 'PrerequisiteID'), 
#        'Assignment': ('TaskID', 'ResourceID'),
#        'Holiday': ('Date', ),
#        'ProjectMeasurement': ('ProjectID', 'MeasurementID'),
#        'MeasurementDependency': ('MeasurementID', 'PriorMeasurementID'),
#        'ProjectWeek': ('ProjectID', 'Week'),
#        'ProjectDay': ('ProjectID', 'Day'),
#    }

def FillTable(name, t, Columns, Data):
    """ Utility routine to load sample (or initial) data into database """
    for i, d in enumerate(Data):
        id = i+1
        t[id] = {}
        t[id]['ID'] = id

        for c, r in zip( Columns, d ):
            if r or r == 0:
                t[id][c] = r
    NextID[name] = len(Data) + 1
    Database[name] = t

def SetTypes():
    """ Create default report types """
    # ReportType
    # Also ( allow user to select from this report type's columns also )
    Columns = ( 'Also', 'Name', 'TableA', 'TableB', 'Label' )
    Data = ( ( None, 'Project', 'Project' ),
             ( 1,    'Project/Report', 'Project', 'Report' ),
             ( None, 'Task', 'Task' ),
             ( 3,    'Task/Assignment', 'Task', 'Assignment' ),
             ( 3,    'Task/Dependency', 'Task', 'Dependency' ),
             ( None, 'Report', 'Report' ),
             ( 6,    'Report/ReportColumn', 'Report', 'ReportColumn' ),
             ( None, 'Resource', 'Resource' ),
             ( 8,    'Resource/Assignment', 'Resource', 'Assignment' ),
             ( None, 'Holiday', 'Holiday'), 
             ( None, 'ReportType', 'ReportType' ), 
             ( 11, 'ReportType/ColumnType', 'ReportType', 'ColumnType' ), 
            )
    FillTable('ReportType', ReportType, Columns, Data)

    # ColumnType
    # DataType ( i = integer, t = text, d = date, g = graphic )
    # AccessType ( da = direct table a, db = direct table b, bi = indirect, s = special )
    # Edit ( True = allow user to change value, False = display only column )
    Columns = ('ReportTypeID', 'DataType', 'AccessType', 'T', 'Edit', 'Width', 'Name', 'Label')
    Data = ( ( 1, 'i', 'd', 'A', False, None, 'ID' ),                # Project
             ( 1, 't', 'd', 'A', True,  140, 'Name' ),
             ( 1, 'd', 'd', 'A', True,  80, 'StartDate', 'Start Date' ),
             ( 1, 'd', 'd', 'A', True,  80, 'TargetEndDate', 'Target\nEnd Date' ),
             ( 2, 'i', 'd', 'B', False, None, 'ID' ),                # Project/Report
             ( 2, 't', 'd', 'B', True,  140, 'Name'), 
             ( 3, 'i', 'd', 'A', False, None, 'ID' ),                # Task
             ( 3, 'i', 'd', 'A', False, 80, 'ProjectID' ),
             ( 3, 't', 'd', 'A', True,  140, 'Name' ),
             ( 3, 'd', 'd', 'A', True,  80, 'StartDate' ),
             ( 3, 'i', 'd', 'A', True,  80, 'DurationHours', 'Duration\n(Hours)' ),
             ( 3, 'd', 'd', 'A', False, 80, 'CalculatedStartDate', 'Start Date\n(Calculated)' ),
             ( 3, 'd', 'd', 'A', False, 80, 'CalculatedEndDate', 'End Date\n(Calculated)' ),
             ( 3, 'g', 's', 'X', False, None, 'Day/Gantt' ),
             ( 3, 'i', 's', 'X', False, None, 'Day/Hours' ),
             ( 4, 'i', 'd', 'B', False, None, 'ID' ),                # Task/Assignment
             ( 4, 't', 'i', 'B', True,  140, 'Resource/Name' ),
             # ( 4, 'i', 's', 'x', False, None, 'Day/Hours' ),
             ( 5, 'i', 'd', 'B', False, None, 'ID' ),                # Task/Dependency
             ( 5, 't', 'i', 'B', True,  140, 'Prerequisite/Name' ),
             # ( 5, 'i', 's', 'X', False, None, 'Day/Hours' ),
             ( 6, 'i', 'd', 'A', False, None, 'ID' ),                # Report
             ( 6, 't', 'd', 'A', True,  140, 'Name' ),
             # ( 6, 't', 'd', 'a', False, 80, 'TableA' ),
             # ( 'TableB', 
             ( 6, 't', 'd', 'A', False, 80, 'SelectColumn' ),
             ( 6, 'i', 'd', 'A', False, 80, 'SelectValue' ),
             ( 6, 'i', 'd', 'A', False, None, 'ProjectID' ),
             ( 7, 'i', 'd', 'B', False, None, 'ID' ),                # Report/ReportColumn
             ( 7, 'i', 'd', 'B', True,  40, 'Width' ), 
             ( 7, 't', 'd', 'B', True,  140, 'Label' ),
             # ( 7, ' 'A', 'TypeA', 'B', 'TypeB', 'Time', 
             ( 7, 'i', 'd', 'B', True,  40, 'Periods' ),
             ( 7, 'd', 'd', 'B', True,  80, 'FirstDate' ),
             ( 8, 'i', 'd', 'A', False, None, 'ID' ),                # Resource
             ( 8, 't', 'd', 'A', True,  80, 'ShortName' ),
             ( 8, 't', 'd', 'A', True,  140, 'Name' ),
             ( 9, 'i', 'd', 'B', False, None, 'ID' ),                # Resource/Assignment
             ( 9, 't', 'i', 'B', True,  140, 'Task/Name' ),
             ( 9, 'i', 's', 'X', False, None, 'Day/Hours' ),
             ( 10, 'i', 'd', 'A', False, None, 'ID' ),               # Holiday
             ( 10, 't', 'd', 'A', True,  140, 'Name' ),
             ( 10, 'd', 'd', 'A', True,  80, 'Date' ), 
             ( 10, 'i', 'd', 'A', True,  40, 'Hours' ),
             ( 11, 'i', 'd', 'A', False, None, 'ID' ),               # ReportType
             ( 11, 'i', 'd', 'A', False, 40, 'Also' ),
             ( 11, 't', 'd', 'A', False, 140, 'Name' ),  # 42
             ( 11, 't', 'd', 'A', False, 100, 'TableA' ),
             ( 11, 't', 'd', 'A', False, 100, 'TableB' ),
             ( 11, 't', 'd', 'A', True,  120, 'Label' ),
             ( 12, 'i', 'd', 'B', False, None, 'ID' ),               # ColumnType
             # ( 12, 'i', 'd', 'B', False, None, 'ReportTypeID' ),
             ( 12, 't', 'd', 'B', False, None, 'DataType' ),
             ( 12, 't', 'd', 'B', False, None, 'AccessType' ),
             ( 12, 'b', 'd', 'B', False, None, 'Edit' ),
             ( 12, 't', 'd', 'B', False, 140, 'Name' ),
             ( 12, 't', 'd', 'B', True,  140, 'Label' ),
            )
    FillTable('ColumnType', ColumnType, Columns, Data)

def SetEmptyData():
    """ Create empty data tables that are ready to use """
    global FileName, Database, UndoStack, RedoStack, ChangedData
    if debug: print "Start SetEmptyData"

    FileName = None  # Filename of current database
    CreateEmptyTables()
    UndoStack = []
    RedoStack = []
    ChangedData = False

    # fill in tables w/ no default values
    Database['Dependency'] = Dependency
    Database['Resource'] = Resource
    Database['Assignment'] = Assignment
    Database['Holiday'] = Holiday
    NextID['Dependency'] = 1
    NextID['Resource'] = 1
    NextID['Assignment'] = 1
    NextID['Holiday'] = 1

    # Project
    Columns = ('Name', 'StartDate', 'TargetEndDate')
    Data = ( ( 'All Projects',),
             ( 'New Project',),
           )
    FillTable('Project', Project, Columns, Data)

    # Task
    Columns = ('ProjectID', 'Name' )
    Data = ( ( 2, 'First Task' ),
           )
    FillTable('Task', Task, Columns, Data)

    # Report
    Columns = ('ReportTypeID', 'Name', 'FirstColumn', 'FirstRow', 'SelectColumn', 'SelectValue', 'ProjectID')
    
    # TableB must contain a column that is named TableA + "ID" that can be used to link the tables
    # SelectColumn must be in TableA
    Data = ( ( 2, 'Project/Report List', 1, 1, None, None, 1 ),
                ( 12, 'Report Options', 4, 8, None, None, 1 ),
                ( 8, 'Resource List', 5, 0,  None, None, 1 ),
                ( 3, 'Gantt Chart', 6, 4, 'ProjectID', 2, 2 ),
            )
    FillTable('Report', Report, Columns, Data)
    #
    Columns = ('ColumnTypeID', 'ReportID', 'NextColumn', 'Width', 'Label', 'A', 'TypeA', 'B', 'TypeB', 'Time', 'Periods', 'FirstDate')
    Data = ( ( 2, 1, 2, 140,         'Name',                 'Name',         'CHAR',         'Name','CHAR',  '', 0, '' ),
                ( 3, 1, 3, 80,          'Start Date',           'StartDate',    'DATE',         '','',          '', 0, '' ),
                ( 4, 1, 0, 80,          'Target Date',          'TargetEndDate', 'DATE',        '','',          '', 0, '' ),
                ( 42, 2, 12, 140,       'Name',                 'Name',         'CHAR',         '','',          '', 0, '' ),
                ( 32, 3, 0, 140,        'Name',                 'Name',         'CHAR',         '','',          '', 0, '' ),
                ( 9, 4, 7, 140,         'Name',                 'Name',         'CHAR',         '','',          '', 0, '' ),
                ( 10, 4, 8, 80,         'Start Date',           'StartDate',    'DATE',         '','',          '', 0, '' ),
                ( 11, 4, 9, 80,         'Duration',             'DurationHours', 'INT',         '','',          '', 0, '' ),
                ( 12, 4, 10, 80,        'Start Date\n(Calculated)', 'CalculatedStartDate', 'DATE',      '','',          '', 0, '' ),
                ( 13, 4, 11, 80,        'End Date\n(Calculated)', 'CalculatedEndDate', 'DATE',  '','',          '', 0, '' ),
                ( 14, 4, 0, None,       '',                     '',             'CHART',        '','',          'Day', 21, GetToday() ),
                (  50, 2, None, 140, None, None, None, None, None, None, None, None )
            )
    FillTable('ReportColumn', ReportColumn, Columns, Data)

    # dump of file (didn't use this, just edited the above
    # Columns = ( 'ColumnTypeID', 'ReportID', 'NextColumn', 'Width', 'Label', 'A', 'TypeA', 'B', 'TypeB', 'Time', 'Periods', 'FirstDate' )
    # Data = (
    #         (  2, 1, 2, 140, 'Name', 'Name', 'CHAR', 'Name', 'CHAR', '', 0, '' )
    #         (  3, 1, 3, 80, 'Start Date', 'StartDate', 'DATE', '', '', '', 0, '' )
    #         (  4, 1, 0, 80, 'Target Date', 'TargetEndDate', 'DATE', '', '', '', 0, '' )
    #         (  42, 2, 12, 140, 'Name', 'Name', 'CHAR', '', '', '', 0, '' )
    #         (  32, 3, 0, 140, 'Name', 'Name', 'CHAR', '', '', '', 0, '' )
    #         (  9, 4, 7, 140, 'Name', 'Name', 'CHAR', '', '', '', 0, '' )
    #         (  10, 4, 8, 80, 'Start Date', 'StartDate', 'DATE', '', '', '', 0, '' )
    #         (  11, 4, 9, 80, 'Duration', 'DurationHours', 'INT', '', '', '', 0, '' )
    #         (  12, 4, 10, 80, 'Start Date\n(Calculated)', 'CalculatedStartDate', 'DATE', '', '', '', 0, '' )
    #         (  13, 4, 11, 80, 'End Date\n(Calculated)', 'CalculatedEndDate', 'DATE', '', '', '', 0, '' )
    #         (  14, 4, 0, None, '', '', 'CHART', '', '', 'Day', 21, '2004-04-13' )
    #         (  50, 2, None, 140, None, None, None, None, None, None, None, None )
    #         )
    #         FillTable(' ReportColumn ',  ReportColumn , Columns, Data)

    # old version
    # Columns = ('ReportID', 'TableName', 'TableID', 'NextRow')
    # Data = (
    #         (1, 'Project', 1,  2 ),
    #         (1, 'Project', 2,  3 ),
    #         (1, 'Report', 4, 0 ),  # lists only the gantt chart report
    #         (4, 'Task', 1, 0 ),
    #         )
    # FillTable('ReportRow', ReportRow, Columns, Data)

    Columns = ( 'ReportID', 'TableName', 'TableID', 'NextRow', 'Hidden' )
    Data = ( (  1, 'Project', 1, 5, None ),
             (  1, 'Project', 2, 3, None ),
             (  1, 'Report', 4, 0, None ),
             (  4, 'Task', 1, 0, None ),
             (  1, 'Report', 1, 6, None ),
             (  1, 'Report', 2, 7, True ),
             (  1, 'Report', 3, 2, None ),
             (  2, 'ReportType', 1, 20, None ),
             (  2, 'ReportType', 2, 24, None ),
             (  2, 'ReportType', 3, 26, None ),
             (  2, 'ReportType', 4, 35, None ),
             (  2, 'ReportType', 5, 37, None ),
             (  2, 'ReportType', 6, 39, True ),
             (  2, 'ReportType', 7, 44, True ),
             (  2, 'ReportType', 8, 49, None ),
             (  2, 'ReportType', 9, 52, None ),
             (  2, 'ReportType', 10, 55, None ),
             (  2, 'ReportType', 11, 59, True ),
             (  2, 'ReportType', 12, 65, True ),
             (  2, 'ColumnType', 1, 21, None ),
             (  2, 'ColumnType', 2, 22, None ),
             (  2, 'ColumnType', 3, 23, None ),
             (  2, 'ColumnType', 4, 9, None ),
             (  2, 'ColumnType', 5, 25, None ),
             (  2, 'ColumnType', 6, 10, None ),
             (  2, 'ColumnType', 7, 27, None ),
             (  2, 'ColumnType', 8, 28, None ),
             (  2, 'ColumnType', 9, 29, None ),
             (  2, 'ColumnType', 10, 30, None ),
             (  2, 'ColumnType', 11, 31, None ),
             (  2, 'ColumnType', 12, 32, None ),
             (  2, 'ColumnType', 13, 33, None ),
             (  2, 'ColumnType', 14, 34, None ),
             (  2, 'ColumnType', 15, 11, True ),
             (  2, 'ColumnType', 16, 36, None ),
             (  2, 'ColumnType', 17, 12, None ),
             (  2, 'ColumnType', 18, 38, None ),
             (  2, 'ColumnType', 19, 13, None ),
             (  2, 'ColumnType', 20, 40, True ),
             (  2, 'ColumnType', 21, 41, True ),
             (  2, 'ColumnType', 22, 42, True ),
             (  2, 'ColumnType', 23, 43, True ),
             (  2, 'ColumnType', 24, 14, True ),
             (  2, 'ColumnType', 25, 45, True ),
             (  2, 'ColumnType', 26, 46, True ),
             (  2, 'ColumnType', 27, 47, True ),
             (  2, 'ColumnType', 28, 48, True ),
             (  2, 'ColumnType', 29, 15, True ),
             (  2, 'ColumnType', 30, 50, None ),
             (  2, 'ColumnType', 31, 51, None ),
             (  2, 'ColumnType', 32, 16, None ),
             (  2, 'ColumnType', 33, 53, None ),
             (  2, 'ColumnType', 34, 54, None ),
             (  2, 'ColumnType', 35, 17, True ),
             (  2, 'ColumnType', 36, 56, None ),
             (  2, 'ColumnType', 37, 57, None ),
             (  2, 'ColumnType', 38, 58, None ),
             (  2, 'ColumnType', 39, 18, None ),
             (  2, 'ColumnType', 40, 60, True ),
             (  2, 'ColumnType', 41, 61, True ),
             (  2, 'ColumnType', 42, 62, True ),
             (  2, 'ColumnType', 43, 63, True ),
             (  2, 'ColumnType', 44, 64, True ),
             (  2, 'ColumnType', 45, 19, True ),
             (  2, 'ColumnType', 46, 66, True ),
             (  2, 'ColumnType', 47, 67, True ),
             (  2, 'ColumnType', 48, 68, True ),
             (  2, 'ColumnType', 49, 69, True ),
             (  2, 'ColumnType', 50, 70, True ),
             (  2, 'ColumnType', 51, 0, True ),
           )
    FillTable('ReportRow',  ReportRow , Columns, Data)
    #
    SetTypes()
    SetOther()
    if debug: print "End SetEmptyData"
        

# set up sample database
def SetSampleData():
    """ Fill data tables with sample/demonstration data """
    global FileName, Database, UndoStack, RedoStack, ChangedData
    if debug: print "Start SetSampleData"

    FileName = None  # Filename of current database
    CreateEmptyTables()
    UndoStack = []
    RedoStack = []
    ChangedData = False
    #
    Columns = ('Name', 'StartDate', 'TargetEndDate')
    Data = ( ( 'All Projects', None, None ),
             ( 'Big Project', '2004-04-04', '2004-12-04'),
             ( 'Important Project', '2004-04-04', '2004-12-04')
           )
    FillTable('Project', Project, Columns, Data)
    #
    Columns = ('ProjectID', 'Name', 'StartDate', 'DurationHours', 
            'CalculatedStartDate', 'CalculatedEndDate', 'CalculatedDurationHours' )
    Data = ( ( 2, 'First Task', '2003-12-31', 12, '2004-04-04', '2004-04-14', 10 ),
             ( 2, 'Second Task', None, None, '2004-04-14', '2004-04-19', 10 ),
             ( 2, 'Third Task', None, None, '2004-04-19', '2004-04-20', 10 ),
             ( 3, 'Only Task', '2004-04-04', None, '2004-04-04', '2004-04-14', 10 )
           )
    FillTable('Task', Task, Columns, Data)
    #
    Columns = ('PrerequisiteID', 'TaskID', 'Name')
    Data = ( ( 1, 2 ),
             ( 2, 3 )
           )
    FillTable('Dependency', Dependency, Columns, Data)
    #
    Columns = ('ReportTypeID', 'Name', 'FirstColumn', 'FirstRow', 'TableA', 'TableB', 'SelectColumn', 'SelectValue', 'ProjectID')
    # TableB must contain a column that is named TableA + "ID" that can be used to link the tables
    # SelectColumn must be in TableA
    Data = ( ( 2, 'Project/Report List', 1, 1, 'Project', 'Report', None, None, None ),
             ( 12, 'Report Column Options', 4, 0 ),
             ( 8, 'Resource List', 5, 3, 'Resource', None, None, None, None ),
             ( 3, 'Gantt Chart', 6, 4, 'Task', None, 'ProjectID', 2, 2 ),
           )
    FillTable('Report', Report, Columns, Data)
    #
    Columns = ('ColumnTypeID', 'ReportID', 'NextColumn', 'Width', 'Label', 'A', 'TypeA', 'B', 'TypeB', 'Time', 'Periods', 'FirstDate')
    Data = ( ( 2, 1, 2, 140,         'Name',                 'Name',         'CHAR',         'Name','CHAR',  '', 0, '' ),
             ( 3, 1, 3, 80,          'Start Date',           'StartDate',    'DATE',         '','',          '', 0, '' ),
             ( 4, 1, 0, 80,          'Target Date',          'TargetEndDate', 'DATE',        '','',          '', 0, '' ),
             ( 42, 2, 0, 140,        'Name',                 'Name',         'CHAR',         '','',          '', 0, '' ),
             ( 32, 3, 0, 140,        'Name',                 'Name',         'CHAR',         '','',          '', 0, '' ),
             ( 9, 4, 7, 140,         'Name',                 'Name',         'CHAR',         '','',          '', 0, '' ),
             ( 10, 4, 8, 80,         'Start Date',           'StartDate',    'DATE',         '','',          '', 0, '' ),
             ( 11, 4, 9, 80,         'Duration\n (Hours)',   'DurationHours', 'INT',         '','',          '', 0, '' ),
             ( 14, 4, 0, None,       '',                     '',             'CHART',        '','',          'Day', 10, '2003-12-30' ),
           )
    FillTable('ReportColumn', ReportColumn, Columns, Data)
    #
    Columns = ('ReportID', 'TableName', 'TableID', 'NextRow')
    Data = ( ( 1,  'Project', 1,  2 ),
             ( 1,  'Report', 1, 7 ),
             ( 3,  'Resource', 1, 0 ),
             ( 4,  'Task', 1, 5 ),
             ( 4,  'Task', 2, 6 ),
             ( 4,  'Task', 3, 0 ),
             ( 1,  'Report', 2, 8 ),
             ( 1,  'Report', 3, 9 ),
             ( 1,  'Project', 2, 10 ),
             ( 1,  'Report', 4, 11 ),
             ( 1,  'Project', 3, 0 ),
           )
    FillTable('ReportRow', ReportRow, Columns, Data)
    #
    Columns = ('ShortName', 'Name')
    Data = ( ( 'Brian', 'Brian Christensen' ),
             ( 'Alex', 'Alexander Christensen' ),
            )
    FillTable('Resource', Resource, Columns, Data)
    #
    Columns = ('ResourceID', 'TaskID')
    Data = ( ( 1, 1 ),
           )
    FillTable('Assignment', Assignment, Columns, Data)
    #
    Columns = ('Date', 'Hours')
    Data = ( ( '2004-03-12', 0 ),
           )
    FillTable('Holiday', Holiday, Columns, Data)

    SetTypes()
    SetOther()
    if debug: print "End SetSampleData"

# ---------
# Needed changes  <----
# set "Also" --> replace name with ID   (done)
# don't change labels if they already exist  (done)

def AddTable(name):
    # add the table if it doesn't already exist
    # (Note: this part of the change isn't reversed by Undo.)
    if name and not Database.has_key(name):
        Database[name] = {}
        Database['NextID'][name] = 1

def AddRow(change):  # add or update row
    changeTable = Database.get(change.get("Table"))
    changeName = change.get("Name")
    if not changeTable or not changeName: return
    cid = 0
    for k, v in changeTable.items():
        if v.get('Name') == changeName:
            cid = k
            break
    if cid:  # already exists
        change['ID'] = cid
    Update(change)

def AddReportType(reportType, columnTypes):
    # should add code to ensure all values are valid

    # convert "Also" name to record id
    also = reportType.get("Also")
    if also:
        alsoid = 0
        for k, v in Database['ReportType'].items():
            if v.get('Name') == also:
                alsoid = k
                break
        if alsoid:
            reportType['Also'] = alsoid
        else:
            if debug: print "couldn't convert Also value", also
            del reportType['Also']

    # Make sure tables exist that are referenced by ReportType
    tableNames =  map( reportType.get, ["TableA", "TableB"] )
    for name in tableNames:
        AddTable(name)

    # Either add or update ReportType
    reportTypeName = reportType["Name"]
    rtid = 0
    for k, v in Database['ReportType'].items():
        if v.get('Name') == reportTypeName:
            rtid = k
            break
    change = reportType
    change["Table"] = 'ReportType'
    if rtid:  # already exists
        change['ID'] = rtid
        Update(change)
        oldRT = True
    else:  # new
        undo = Update(change) 
        rtid = undo['ID']
        oldRT = False

    # expects a list of ColumnTypes
    for change in columnTypes:
        typeName = change["Name"]

        # check whether column type already exists
        ctid = 0
        if oldRT:
            for k, v in Database['ColumnType'].items():
                if v.get('ReportTypeID') == rtid and v.get('Name') == typeName:
                    ctid = k
                    break
        if ctid: 
            change['ID'] = ctid  # if column type already exists, change this to an update instead of an add
            if change.has_key('Label'): del change['Label']  # don't change label on existing column type, is test necessary?
        change["Table"] = "ColumnType"
        change["ReportTypeID"] = rtid

        # check for required fields ??
        Update(change) 

# ---------
def FindID(table, field1, value1, field2, value2):
    # if debug: print "start FindID", table, field1, value1, field2, value2
    if not Database.has_key(table): return 0
    if not field2:
        for k, v in Database[table].iteritems():
            if (v.get(field1) == value1): return k
    else:
        for k, v in Database[table].iteritems():
            if (v.get(field1) == value1) and (v.get(field2) == value2): return k

def FindIDs(table, field1, value1, field2, value2):
    if not Database.has_key(table): return []
    result = []
    if not field2:
        for k, v in Database[table].iteritems():
            if (v.get(field1) == value1): result.append(k)
    else:
        for k, v in Database[table].iteritems():
            if (v.get(field1) == value1) and (v.get(field2) == value2): result.append(k)
    return result

# --------------------- ( is this used for anything? )
# def SetDependencyXref():
#     for i, d in Dependency.iteritems():
#         p = d['PrerequisiteID']
#         s = d['TaskID']
#         if not Predecessor.has_key(s):
#             Predecessor[s] = []
#         if not Successor.has_key(p):
#             Successor[p] = []
#         Predecessor[s].append(p)
#         Successor[p].append(s)

# ----------------------

# check for important changes
def CheckChange(change):  # change contains the undo info for the changes (only changed columns included)
    """ Check for important changes. """
    global ChangedCalendar, ChangedSchedule, ChangedReport, ChangedRow
    if debug: print "Start CheckChange"
    if debug: print 'change', change
    if not change.has_key('Table'):
        if debug: print "change does not specify table"
        return  

    if change.has_key('zzStatus') and not change['Table'] in ('ReportRow', 'ReportColumn'): ChangedRow = True  # something has been added or deleted
        # 'zzStatus' is not set for new record; ChangedRow flag is set in Update when adding a record
        # zzStatus on all new records just takes up space.

    if change['Table'] == 'OtherData':  # check for change in hours per day
        # for k in ('WeekHours'):  # which of these to use?
        for k in ('WeekHours',):
            if change.has_key(k): ChangedCalendar = True; break
    elif change['Table'] == 'ReportRow':
        for k in ('NextRow', 'Hidden', 'zzStatus'):
            if change.has_key(k): ChangedReport = True; break
    elif change['Table'] == 'ReportColumn':
        for k in ('Type', 'NextColumn', 'Time', 'Periods', 'FirstDate', 'zzStatus'):
            if change.has_key(k): ChangedReport = True; break
    elif change['Table'] == 'Report':
        for k in ('Name', 'FirstColumn', 'FirstRow', 'ShowHidden', 'zzStatus'):
            if change.has_key(k): ChangedReport = True; break
            # should I include 'Name' here? it doesn't really change the report structure to change the name

    elif ChangedCalendar: pass  # everything will be refreshed anyway, don't bother looking further
    elif change['Table'] == 'Holiday':
        for k in ('Date', 'Hours', 'zzStatus'):
            if change.has_key(k): ChangedCalendar = True; break

    elif ChangedSchedule: pass  # 
    elif change['Table'] == 'Task':
        for k in ('StartDate', 'DurationHours', 'zzStatus', 'ProjectID', 'TaskID'):
            if change.has_key(k): ChangedSchedule = True; break
    elif change['Table'] == 'Dependency':  # allows dependencies that refer to different projects
        for k in ('PrerequisiteID', 'TaskID', 'zzStatus'):
            if change.has_key(k): ChangedSchedule = True; break
    elif change['Table'] == 'Project':
        for k in ('zzStatus', 'StartDate'):
            if change.has_key(k) :  ChangedSchedule = True; break
    if debug: print "End Check Change"
# --------------------------

def RefreshReports():
    if debug: print "Start RefreshReports"
    for k, v in OpenReports.items():  # invalid or deleted reports might be closed during this loop
        if debug: print 'reportid', k
        if v == None: continue
        if Report[k].get('zzStatus', 'active') == 'deleted':
            CloseReport(k)
        if k != 1: v.SetReportTitle()  # only grid reports
        Menu.AdjustMenus(v)
        v.Refresh()
        # v.Report.RefreshReport()  # update displayed data
        v.Report.Refresh()  # update displayed data (needed for Main on Windows, not needed on Mac)
    if debug: print "End RefreshReports"

# ----- undo and redo

def Recalculate(autogantt=True):
    global ChangedCalendar, ChangedSchedule, ChangedRow, ChangedReport

    UpdateCalendar = ChangedCalendar and autogantt
    UpdateGantt = (ChangedCalendar or ChangedSchedule) and autogantt
    UpdateReports = UpdateCalendar or UpdateGantt or ChangedReport or ChangedRow

    # these routines shouldn't add anything to the undo stack
    if UpdateCalendar:
        SetupDateConv(); ChangedCalendar = False
    if UpdateGantt:
        GanttCalculation(); ChangedSchedule = False
    if ChangedRow:
        AdjustReportRows(); ChangedRow = False
    if UpdateReports:
        ChangedReport = False
        for k, v in OpenReports.items():  # an invalid report may be closed in this loop
            if v == None: continue
            v.UpdatePointers()
    RefreshReports()  # adjust visibility/appearance of user interface objects

def SetUndo(message):
    """ This is the last step in submitting a group of changes to the database. 
            It tells the system to adjust any system calculations and to update the displays. """
    global ChangedData, UndoStack, RedoStack
    if debug: print "Start SetUndo"
    if debug: print "message", message

    UndoStack.append(message)
    ChangedData = True  # file needs to be saved
    RedoStack = []  # clear out the redo stack
    autogantt = Option.get('AutoGantt')
    Recalculate(autogantt)

    if debug: print "End SetUndo"

def _Do(fromstack, tostack):
    global ChangedData
    if fromstack and isinstance(fromstack[-1], string):
        savemessage = fromstack.pop()
        while fromstack and isinstance(fromstack[-1], dict):
            change = fromstack.pop()
            redo = Update(change, 0)  # '0' means don't put change into Undo Stack
            tostack.append(redo)
        tostack.append(savemessage)
        ChangedData = True  # file needs to be saved
        Recalculate()

def DoUndo():
    _Do(UndoStack, RedoStack)

def DoRedo():
    _Do(RedoStack, UndoStack)

# --------------------------

# update routine
def Update(change, push=1):
    global ChangedRow
    if debug: print 'Start Update'
    if debug: print 'change', change
    tname = change.get('Table')
    if not tname:
        if debug: print 'change does not specify table'
        raise KeyError

    table = Database.get(tname)
    if table == None:
        if debug: print 'change specifies invalid table:', tname
        raise KeyError

    undo = {'Table': tname}
    if change.has_key('ID'):
        if debug: print "Change existing record"
        id = undo['ID'] = change['ID']
        record = table[id]
        for c, newval in change.iteritems():  # process each new field
            if c == 'Table' or c == 'ID': continue
            oldval = record.get(c)
            if newval != oldval:
                undo[c] = oldval
                if newval or newval == 0:
                    record[c] = newval
                else:
                    del record[c]
        CheckChange(undo)
    else:
        if debug: print "Add new record"
        record = {}
        for c, newval in change.iteritems():  # process each new field
            if newval or newval == 0:
                record[c] = newval
        undo['zzStatus'] = 'deleted'
        # record['zzStatus'] = 'active'  # not needed; a record without a zzStatus is assumed to be active
        id = NextID[tname]
        if debug: print "new id:", id
        NextID[tname] = id + 1
        record['ID'] = undo['ID'] = id
        CheckChange(record)
        ChangedRow = True  # CheckChange doesn't recognize this for new records
        del record['Table']  # save space in the database
        table[id] = record
    if push: UndoStack.append(undo)
    if debug: print "End Update"
    return undo

# ------------------

### date conversion routines
def HoursToDate(hours):
    """" return date that includes an hour and the offset into that date """
    dmax = len(DateInfo)
    hmax = DateInfo[-1][1]  # cum hours
    d = int((dmax * hours) / hmax)
    if DateInfo[d][1] > hours: i = -1
    else: i = 1
    while not (DateInfo[d][1] <= hours < (DateInfo[d][0] + DateInfo[d][1])): d += i
    return DateIndex[d], hours - DateInfo[d][1]

def ValidDate(value):
    if value < '1901-01-01':
        return ""
    elif len(value) == 10 and value[4] == '-' and value[7] == '-':
        try:
            datetime.date(int(value[0:4]),int(value[5:7]), int(value[8:10]))
        except ValueError:
            return ""
    return value

def SetupDateConv():
    """" Setup tables that will be used for all schedule date calculations """
    global DateConv, DateIndex, DateInfo, DateNextMonth, WeekToHour, DayToHour
    # global ChangedCalendar
    # --->> don't use 'Update' in this processing <<----
    # dow = calendar.weekday(yyyy,mm,dd)  # monday = 0
    # dow, days = calendar. monthrange(yyyy,mm)  #  days in month

    #   # d = date object
    # dow = d.weekday()
    # firstdayofweek = )

    def StringToDate(stringdate):
        return datetime.date(int(stringdate[0:4]),int(stringdate[5:7]), int(stringdate[8:10]))

    def makeFOW(d):
        return d - (d.weekday() * datetime.timedelta(1))   # backup to first day of week

    Today = GetToday()

    DateConv = {}  # usage: index = DateConv['2004-01-01']
    DateIndex = []  # usage: date = DateIndex[1]
    DateInfo = [] # dayhours, cumhours, dow = DateInfo[1]
    DateNextMonth = {} # usage: index = DateNextMonth['2004-01']  # index is first day of next month

    FirstDate = Today
    LastDate = Today
    oneday = datetime.timedelta(1)
    oneweek = oneday * 7

    # find the earliest and last dates that have been specified anywhere
    for p in Project.values():
        d = p.get('StartDate')
        if d:
            if d < FirstDate and ValidDate(d): FirstDate = d
            elif d > LastDate and ValidDate(d): LastDate = d
        d = p.get('TargetEndDate')
        if d:
            if d < FirstDate and ValidDate(d): FirstDate = d
            elif d > LastDate and ValidDate(d): LastDate = d

    for p in Task.values():
        d = p.get('StartDate')
        if d:
            if d < FirstDate and ValidDate(d): FirstDate = d
            elif d > LastDate and ValidDate(d): LastDate = d
        d = p.get('EndDate')
        if d:
            if d < FirstDate and ValidDate(d): FirstDate = d
            elif d > LastDate and ValidDate(d): LastDate = d

    d1 = StringToDate(FirstDate)
    d1 = makeFOW(d1) - (oneweek * 50)  # allow 50 weeks of room before the first date in the file

    d2 = StringToDate(LastDate)
    d2 = makeFOW(d2) - oneday + (oneweek * 105)

    if debug: print "first & last", d1.strftime("%Y-%m-%d"), d2.strftime("%Y-%m-%d")

    wh = Other.get('WeekHours') or (8,8,8,8,8,0,0)

    WeekToHour = 0
    WeekToDay = 0
    DayToHour = 0
    AllowHourToDay = True

    for dayLength in wh:
        if dayLength > 0:
            WeekToHour += dayLength
            WeekToDay += 1
            if AllowHourToDay:
                if DayToHour == 0:
                    DayToHour = dayLength
                elif DayToHour != dayLength:
                    AllowHourToDay = False
    DayToHour = WeekToHour.__truediv__(WeekToDay)

    hxref = {}
    for k, v in Holiday.iteritems():
        if v.get('zzStatus', 'active') == 'deleted': continue
        date = v.get('Date')
        hours = v.get('Hours')
        if date and hours != None:
            hxref[date] = hours
    cumhours = 0
    i = 0
    lastmo = FirstDate[0:7] # keep track of prior month
    while d1 <= d2:
        date = d1.strftime("%Y-%m-%d")
        dow = i % 7
        dh = hxref.get(date, wh[dow])  # use holiday hours if provided, else use week hours
        # if debug: print "date", date, dh
        DateConv[date] = i
        DateIndex.append(date)
        DateInfo.append( (dh, cumhours, dow), )  # work hours on date, work hours prior to date, day of week
        if date[0:7] != lastmo:  # save first index of this month under prior month
            DateNextMonth[lastmo] = i
            lastmo = date[0:7]
        cumhours += dh; i += 1
        d1 += oneday
    # if debug: print "DateNextMonth", DateNextMonth

    Other['BaseDate'] = DateIndex[0]  # The base date is a reflexion of the date conversion tables that are
                                      # in place. It must always reflect the current tables. It is not 'undo-able'

def AddMonths(ix, months):
    """
    adds specified number of months to index
    returns first day of month
    """
    if months > 0:
        cnt = 0
        while cnt < months and ix < len(DateIndex):
            yymm = DateIndex[ix][0:7]
            if DateNextMonth.has_key(yymm):
                ix = DateNextMonth[yymm]
            cnt += 1
    else:
        months *= -1
        yy, mm = int(DateIndex[ix][0:4]), int(DateIndex[ix][5:7])
        cnt = 0
        while cnt < months:
            mm += -1
            if mm < 1: mm = 12; yy += -1
            cnt += 1
        newdate = "%04d-%02d-01" % (yy, mm)
        if DateConv.has_key(newdate):
            ix = DateConv[newdate]
    return ix

def GetPeriodStart(period, ix, of):
    return GetPeriodInfo(period, ix, of, 1)

def GetPeriodInfo(period, ix, of, index=0):  # needed by server
    """
    Returns: hours in period, hours index of start of period, day of week of start of period 
    """
    period = period[0:3]
    if period == "Wee":
        ix -= DateInfo[ ix ][2]  # convert to beginning of week
        dh, cumh, dow = DateInfo[ix  + of * 7]
        dh2, cumh2, dow2 = DateInfo[ix  + (of + 1) * 7]
        dh = cumh2 - cumh
    elif period == "Mon":
        ix -= int(DateIndex[ ix ][8:10]) - 1  # convert to beginning of month
        ix = AddMonths(ix, of)
        dh, cumh, dow = DateInfo[ix]
        dh2, cumh2, dow2 = DateInfo[AddMonths(ix, 1)]
        dh = cumh2 - cumh
    elif period == "Qua":
        year = DateIndex[ ix ][0:4]
        mo = DateIndex[ ix ][5:7]
        if   mo <= "03": mo = "01"
        elif mo <= "06": mo = "04"
        elif mo <= "09": mo = "07"
        else:            mo = "10"
        ix = DateConv[ year + '-' + mo + '-01' ]  # convert to beginning of quarter
        ix = AddMonths(ix, of * 3)
        dh, cumh, dow = DateInfo[ix]
        dh2, cumh2, dow2 = DateInfo[AddMonths(ix, 3)]
        dh = cumh2 - cumh
    else: # period == "Day":
        ix += of
        dh, cumh, dow = DateInfo[ix]
    if index:
        return ix
    else:
        return (dh, cumh, dow)

# -----------------
def GetToday():
    dToday = datetime.date.today()      # get date object for today
    return dToday.strftime("%Y-%m-%d")  # convert to standard database format

def GanttCalculation(): # Gantt chart calculations - all dates are in hours
    # change = { 'Table': 'Task' }  # will be used to apply updates  --->> Don't Use 'Update' here <<--
    # Set the project start date (use specified date or today)  later may be adjusted by the tasks' start dates
    Today = GetToday()
    if debug: print "today", Today
    ps = {}  # project start dates indexed by project id
    for k, v in Project.iteritems():
        if v.get('zzStatus', 'active') == 'deleted': continue
        sd = v.get('StartDate')
        # if debug: print "project, startdate", k, sd
        if sd == "": sd = Today
        elif sd == None: sd = Today
        elif not DateConv.has_key(sd): sd = Today  # default project start dates to today
        ps[k] = sd
        #ProjectStart = Project[projectid].get('StartDate', Today.strftime("%Y-%m-%d"))
        # if debug: print "project, startdate (adjusted)", k, sd

    # dependencies
    pre = {}  # task prerequisites indexed by task id number
    suc = {}  # task successors
    precnt = {}  # count of all unprocessed prerequisites
    succnt = {}  # count of all unprocessed successors
    tpid = {}  # task's projectid
    parent = {}  # children indexed by parent id
    for k, v in Task.iteritems():  # init dependency counts, xrefs, and start dates
        if v.get('zzStatus', 'active') == 'deleted': continue
        pid = v.get('ProjectID')
        if not pid or not Project.has_key(pid): continue  # silently ignore tasks w/o projects

        precnt[k] = 0
        succnt[k] = 0
        pre[k] = []
        suc[k] = []
        tpid[k] = pid
        # if debug: "task's project", k, pid
        tsd = v.get('StartDate')
        if tsd == "": tsd = None
        if tsd and tsd < ps[pid]: ps[pid] = tsd  # adjust project start date if task starts are earlier

        # if debug: print "task data", k, v
        p = v.get('TaskID')  # parent task id
        if p and p != k and Task.has_key(p):  # ignore parent pointer if it points to self
            if parent.has_key(p):
                parent[p] += 1  # count the number of children
            else:
                parent[p] = 1  # add to list of parents

    if debug: print "parent counts", parent
    for k in parent.keys():  # clear parent dates, will calculate dates from children
        del precnt[k] # remove parents from pass calculations
        del succnt[k]
        # del pre[k]
        # del suc[k]

        # if debug: print "clearing parent", k, Task[k]
        # update database -- doesn't use Update, but may in the future
        Task[k]['hES'] = None
        Task[k]['hEF'] = None
        Task[k]['CalculatedStartDate'], Task[k]['CalculatedStartHour'] = None, None
        Task[k]['CalculatedEndDate'], Task[k]['CalculatedEndHour'] = None, None

        Task[k]['hLS'] = None
        Task[k]['hLF'] = None
        Task[k]['FreeFloatHours'] = None
        Task[k]['TotalFloatHours'] = None
        # if debug: print "cleared parent", k, Task[k]

    for k, v in Dependency.iteritems():
        # if debug: print "dependency record", v
        if v.get('zzStatus', 'active') == 'deleted': continue
        p = v['PrerequisiteID']
        t = v['TaskID']
        if parent.has_key(p) or parent.has_key(t): continue  # skip parent dependencies
        if suc.has_key(p) and pre.has_key(t):  # silently ignore dependencies for tasks not in list
            precnt[t] += 1
            succnt[p] += 1
            pre[t].append(p)
            suc[p].append(t)

    # convert project start dates to hours format
    ProjectStartHour = {}; ProjectEndHour = {}
    for k, v in ps.iteritems():  # k = project id, v = start date
        if not DateConv.has_key(v): continue
        si = DateConv[v]  # get starting date index
        sh = DateInfo[si][1]  # get cum hours for start date
        ProjectStartHour[k] = sh
        if debug: "project, start hour", k, sh
        ProjectEndHour[k] = 0  # prepare to save project end hour

    # forward pass
    moretodo = True
    while moretodo:
        moretodo = False  # if it doesn't process at least one task per loop it will quit
                                # dependency loops are silently ignored
        for k, v in precnt.iteritems():  # k is the task id
            if v == 0:  # all prereqs have been processed; set to -1 when done
                moretodo = True  # will make an extra final pass through all tasks

                # calculate early start for task
                es = ProjectStartHour[tpid[k]]
                for t in pre[k]:
                    ef = Task[t]["hEF"]
                    if ef > es: es = ef
                # if start date was specified, use it if possible
                # (note: the end date is not currently used to compute a start date)
                tsd = Task[k].get('StartDate')
                if tsd == "": tsd = None
                ted = Task[k].get('EndDate')
                if ted == "": ted = None
                td  = Task[k].get('DurationHours')
                if td == "": td = None
                if tsd and DateConv.has_key(tsd):
                    tsi = DateConv[tsd]  # date index
                    tsh = DateInfo[tsi][1]  # date hour
                    if tsh > es: es = tsh  # use the date if dependencies allow

                # calculate early finish
                if tsd and ted and not td and DateConv.has_key(ted):  # use difference in dates to compute duration
                    tei = DateConv[ted]  # date index
                    teh = DateInfo[tei+1][1]  # first hour of next day
                    ef = es + (tsh - teh)
                elif td:
                    ef = es + td
                else:
                    ef = es + 8

                if ef > ProjectEndHour[tpid[k]]: ProjectEndHour[tpid[k]] = ef  # keep track of project end date 

                # update database -- doesn't use Update, but may in the future
                Task[k]['hES'] = es
                Task[k]['hEF'] = ef
                Task[k]['CalculatedStartDate'], Task[k]['CalculatedStartHour'] = HoursToDate(es)
                Task[k]['CalculatedEndDate'], Task[k]['CalculatedEndHour'] = HoursToDate(ef)

                # change['ID'] = k
                # change['hES'] = es
                # change['hEF'] = ef
                # change['CalculatedStartDate'], change['CalculatedStartHour'] = HoursToDate(es)
                # change['CalculatedEndDate'], change['CalculatedEndHour'] = HoursToDate(ef)
                # Update(change, 0)

                # tell successor that I'm ready
                for t in suc[k]: precnt[t] -= 1
                precnt[k] = -1
    vs = precnt.values()
    if vs.count(-1) != len(vs):
        if debug: print "forward pass did ", vs.count(-1), " of ", len(vs), ". Probably a dependency loop."
        return  # didn't finish forward pass, skip backward pass

    # backward pass
    moretodo = True
    while moretodo:
        moretodo = False  # if it doesn't process at least one task per loop it will quit
                                # dependency loops are silently ignored
        for k, v in succnt.iteritems():  # k is the task id
            if v == 0:  # all successors have been processed; set to -1 when done
                moretodo = True  # will make an extra final pass through all tasks

                # calculate late finish and free float for task
                lf = ProjectEndHour[tpid[k]]
                ses = lf  # successor early start
                for t in suc[k]:
                    ls = Task[t]["hLS"]
                    if ls < lf: lf = ls
                    es = Task[t]["hES"]
                    if es < ses: ses = es
                ff = ses - Task[k]["hEF"]

                # calculate late start
                td  = Task[k]["hEF"] - Task[k]["hES"]
                ls = lf - td

                # update database -- doesn't use Update, but may in the future
                Task[k]['hLS'] = ls
                Task[k]['hLF'] = lf
                Task[k]['FreeFloatHours'] = ff
                Task[k]['TotalFloatHours'] = ls - Task[k]["hES"]

                # change['ID'] = k
                # change['hLS'] = ls
                # change['hLF'] = lf
                # ---------- remember to add float lines
                # Update(change, 0)

                # tell predecessor that I'm ready
                for t in pre[k]: succnt[t] -= 1
                succnt[k] = -1

    for k, v in Task.iteritems():  # derive parent dates from children
            # reminder: make sure this handles deleted & purged records properly
        if v.get('zzStatus', 'active') == 'deleted': continue
        if parent.has_key(k):
            Task[k]['SubtaskCount'] = parent[k]  # save count of child tasks
            continue  # skip parent tasks
        if Task[k].has_key('SubtaskCount'): del Task[k]['SubtaskCount']  # do I need to test first?
        p = v.get('TaskID')  # parent task id
        if not p or k == p: continue  # ignore all tasks w/o parents (or w/ self for parent)
        # if debug: print "adjust parents of ", k, v
        # update database -- doesn't use Update, but may in the future
        hes, hef, hls, hlf = [ v.get(x) for x in ['hES', 'hEF', 'hLS', 'hLF']]
        loopcnt = 0
        while p and Task.has_key(p) and loopcnt < 10:
            # if debug: print "adjusting parent ", p, Task[p]
            loopcnt += 1
            phes, phef, phls, phlf = [ Task[p].get(x) for x in ['hES', 'hEF', 'hLS', 'hLF']]
            if not phes or hes < phes:
                Task[p]['hES'] = hes
                Task[p]['CalculatedStartDate'], Task[p]['CalculatedStartHour'] = HoursToDate(hes)
            if not phef or hef > phef:
                Task[p]['hEF'] = hef
                Task[p]['CalculatedEndDate'], Task[p]['CalculatedEndHour'] = HoursToDate(hef)
            if not phls or hls < phls: Task[p]['hLS'] = hls
            if not phlf or hlf > phlf: Task[p]['hLF'] = hlf

            # if debug: print "adjusted parent ", p, Task[p]

            p = Task[p].get('TaskID')  # parent task id
                                       
# end of GanttCalculation
# -----------------

def GetColumnDate(colid, of):  # required by server
        """
    colid == ReportColumn ID; of == the offset; returns the index of first day of period
        """
        if of == -1:
		return None
        else:
            ctid = ReportColumn[colid].get('ColumnTypeID')  # get column type record that corresponds to this column
            if not ctid or not ColumnType.has_key(ctid): return ""  # shouldn't happen
            ct = ColumnType[ctid]
            ctperiod, ctfield = ct.get('Name').split("/")

            firstdate = ReportColumn[colid].get('FirstDate')

            if not DateConv.has_key(firstdate): firstdate = GetToday()
            index = DateConv[ firstdate ]

            if ctperiod == "Day":
                result = index + of
            elif ctperiod == "Week":
                index -= DateInfo[ index ][2]  # convert to beginning of week
                result = index + (of * 7)
            elif ctperiod == "Month":
                index -= int(DateIndex[ index ][8:10]) - 1  # convert to beginning of month
                result = AddMonths(index, of)
            elif ctperiod == "Quarter":
                year = firstdate[0:4]
                mo = firstdate[5:7]
                if   mo <= "03": mo = "01"
                elif mo <= "06": mo = "04"
                elif mo <= "09": mo = "07"
                else:            mo = "10"
                index = DateConv[ year + '-' + mo + '-01' ]  # convert to beginning of quarter
                result = AddMonths(index, of * 3)
            else:
                return None  # unknown time scale
        if result > len(DateConv): return None
        return result  # should test to make sure it is valid

def GetColumnHeader(colid, of):
        """
    colid == ReportColumn ID; of == the offset; returns the first and second header line
        """
        if of == -1:
            label = ReportColumn[colid].get('Label')
            if not label or label == "":
                ctid = ReportColumn[colid].get('ColumnTypeID')  # get column type record that corresponds to this column
                if not ctid or not ColumnType.has_key(ctid): return ""  # shouldn't happen
                ct = ColumnType[ctid]
                label = ct.get('Label') or ct.get('Name')
        else:
            ctid = ReportColumn[colid].get('ColumnTypeID')  # get column type record that corresponds to this column
            if not ctid or not ColumnType.has_key(ctid): return ""  # shouldn't happen
            ct = ColumnType[ctid]
            ctperiod, ctfield = ct.get('Name').split("/")

            firstdate = ReportColumn[colid].get('FirstDate')

            if not DateConv.has_key(firstdate): firstdate = GetToday()
            index = DateConv[ firstdate ]

            if ctfield == 'Gantt':
                if ctperiod == "Day":
                    date = DateIndex[ index + of ]
                    if of == 0 or date[8:10] == '01': label = date[5:7]
                    else: 
                        dow = DateInfo[ index + of ][2]
                        label = 'MTWHFSS'[dow]
                    label += '\n' +  date[8:10]
                elif ctperiod == "Week":
                    index -= DateInfo[ index ][2]  # convert to beginning of week
                    date = DateIndex[ index + (of * 7) ]
                    if of == 0 or date[8:10] <= '07': label = date[5:7]
                    else: label = ''
                    label += '\n' +  date[8:10]
                elif ctperiod == "Month":
                    index -= int(DateIndex[ index ][8:10]) - 1  # convert to beginning of month
                    date = DateIndex[ AddMonths(index, of) ]
                    if of == 0 or date[5:7] == '01': label = date[2:4]
                    else: label = ''
                    label += '\n' +  date[5:7]
                elif ctperiod == "Quarter":
                    year = firstdate[0:4]
                    mo = firstdate[5:7]
                    if   mo <= "03": mo = "01"
                    elif mo <= "06": mo = "04"
                    elif mo <= "09": mo = "07"
                    else:            mo = "10"
                    index = DateConv[ year + '-' + mo + '-01' ]  # convert to beginning of quarter
                    date = DateIndex[ AddMonths(index, of * 3) ]
                    if of == 0 or date[5:7] == '01': label = date[2:4]
                    else: label = ''
                    label += '\nQ' +  str((int(date[5:7]) + 2)//3) 
                else:
                    label = "-"  # unknown time scale
            else: 
                if ctperiod == "Day":
                    date = DateIndex[ index + of ]
                elif ctperiod == "Week":
                    index -= DateInfo[ index ][2]  # convert to beginning of week
                    date = DateIndex[ index + (of * 7) ]
                else:
                    return "-"  # unknown time scale

                if of == 0:
                    label = ctfield[:5]  # ??try column width mod 8??
                else:
                    label = ''
                label += '\n' + date[5:7] + "/" + date[8:10]
        return label

def GetCellValue(rowid, colid, of):
        # of = self.coloffset[col]
        ctid = ReportColumn[colid].get('ColumnTypeID')  # get column type record that corresponds to this column
        if not ctid or not ColumnType.has_key(ctid): return ""  # shouldn't happen
        ct = ColumnType[ctid]
        # ct = self.columntype[self.ctypes[col]]
        if of == -1:
            rr = ReportRow[rowid]
            rtable = rr.get('TableName')
            tid = rr['TableID']  # was 'TaskID' -> changed to generic ID

            # rc = self.reportcolumn[self.columns[col]]

            t = ct.get('T', 'X')
            reportid = rr.get('ReportID')
            if not reportid or not Report.has_key(reportid): return ""  # shouldn't happen
            rtid = Report[reportid].get('ReportTypeID')
            ctable = ReportType[rtid].get('Table' + t)

            dt = ct.get('DataType')
            at = ct.get('AccessType')
            if rtable != ctable:
                value = ''
            elif at == 'd':
                column = ct.get('Name')
                # print column  # it prints each column twice - why???
                value = Database[rtable][tid].get(column, "")
            elif at == 'i':
                try:
                    it, ic = ct.get('Name').split('/')  # indirect table & column
                except ValueError:
                    if debug: print "Indirect column w/o 'Table/Column', Name is: ", ct.get('Name')
                    value = ""
                else:
                    iid = Database[rtable][tid].get(it+'ID')
                    # if debug: print "rtable, tid, it, ic, iid", rtable, tid, it, ic, iid
                    if iid:
                        value = Database[it][iid].get(ic, "")
                    else:
                        value = ""
            elif at == 'list':  # create a comma separated list of data
            # use this Column's value --> go to this Table --> select rows where this Column = value --> return this Column's values
            #    Table2 -> Column2
            # 'Prerequisites': ID/Dependency/TaskID/PrerequisiteID
            # 'Successors': ID/Dependency/PrerequisiteID/TaskID
            # 'ChildTasks': ID/Task/TaskID/ID  (TaskID == parent's id)
            # 'ResourceNames': ID/Assignment/TaskID/ResourceID/Resource/Name
                try:
                    listcol, listtable, listselect, listtarget, listtable2, listcol2 = ct.get('Path').split('/')  # path to values
                except ValueError:
                    if debug: print "List column needs valid path, has: ", ct.get('Path')
                    value = ""
                else:
                    listvalue = Database[rtable][tid].get(listcol)
                    rows = FindIDs(listtable, listselect, listvalue, None, None)
                    if rows:
                        vals = [ Database[listtable][x].get(listtarget) for x in rows if not Database[listtable][x].get('zzStatus') == 'deleted']
                        if listtable2 and listcol2:
                            vals = [ Database[listtable2][x].get(listcol2) or "-" for x in vals if not Database[listtable2][x].get('zzStatus') == 'deleted']
                        value = ", ".join( [ str(x) for x in vals ] )  # need to check this with unicode
                    else:
                        value = ""
            if dt == 'u' and isinstance(value, int) and value > 0:
                w, h = divmod(value, WeekToHour)
                if h > int(WeekToHour): w += 1; h = 0

                if AllowHourToDay:
                    d, h = divmod(h, DayToHour)
                    if h > int(DayToHour): d += 1; h = 0
                else:
                    d = 0

                value = []
                if w: value.append(str(int(w)) + 'w')
                if d: value.append(str(int(d)) + 'd')
                if h: value.append(str(int(h)) + 'h')
                return ' '.join(value)
        else:
            # ---- Here are some examples that this should handle ----
            # -- Report Type => Column Type --
            # TaskDay => Day/Gantt, Day/Hours
            # ResourceDay => Day/Hours
            # ProjectDay or ProjectWeek => Day/Measurement, Week/Measurement
            # TaskWeek => Week/PercentComplete, Week/Effort
            # ResourceWeek => Week/Effort

            ctperiod, ctfield = ct.get('Name').split("/")

            if ctfield == "Gantt":  # don't display a value
                value = ''
            else:  # table name, field name, time period, and record id
                rr = ReportRow[rowid]
                tablename = rr.get('TableName')
                tid = rr['TableID']
                if ctfield == 'Measurement':  # find field name
                    mid = Database[tablename][tid].get('MeasurementID')  # point at measurement record
                    if mid: 
                        fieldname = Database['Measurement'][mid].get('Name')  # measurement name == field name
                    else:
                        fieldname = None
                    tid = Database[tablename][tid].get('ProjectID')  # point at measurement record
                    tablename = 'Project'  # only supports project measurements
                else:
                    fieldname = ctfield

                # find the period date
                firstdate = ReportColumn[colid].get('FirstDate')
                if not DateConv.has_key(firstdate): firstdate = GetToday()
                index = DateConv[ firstdate ]
                if ctperiod == "Day":
                    date = DateIndex[ index + of ]
                elif ctperiod == "Week":
                    index -= DateInfo[ index ][2]  # convert to beginning of week
                    date = DateIndex[ index + (of * 7) ]
                else:
                    date = None
                timename = tablename + ctperiod

                timeid = FindID(timename, tablename + "ID", tid, 'Period', date)
                # if debug: print "timeid", timeid
                if timeid:
                    value = Database[timename][timeid].get(fieldname)
                    # if debug: print "timename, timeid, fieldname, value: ", timename, timeid, fieldname, value
                    # if debug: print "record: ", Data.Database[timename][timeid]
                else:
                    value = None
                    # if debug: print "didn't find timeid", timeid, value

        if value == None: value = ''
        return value

# -----------------
def UpdateDataPointers(self, reportid):  # self is either a list or a table behind a grid
    if debug: print "Start UpdateDataPointers"
    # create local pointers to database
    self.data = Database
    # pointers to one record
    self.report = self.data["Report"][reportid]
    reporttypeid = self.report['ReportTypeID']
    #if debug: print 'reporttypeid', reporttypeid
    self.reporttype = self.data["ReportType"][reporttypeid]
    # pointers to tables
    self.reportcolumn = self.data["ReportColumn"]
    self.columntype = self.data["ColumnType"]
    self.reportrow = self.data["ReportRow"]
    if debug: print "End UpdateDataPointers"

def UpdateRowPointers(self):  # self is object that stores 'rows', used by all reports to create local pointers to report rows
    rr = self.reportrow  # pointer to report row table
    show = self.report.get('ShowHidden', False)
    r = self.report.get('FirstRow', 0)
    self.rows = []
    while r != 0 and r != None:
        if show:
            self.rows.append( r )
        else:
            hidden = rr[r].get('Hidden', False)
            table = rr[r].get('TableName')
            id = rr[r].get('TableID')
            active = table and id and (Database[table][id].get('zzStatus', 'active') == 'active')
            if (not hidden) and active:
                self.rows.append( r )
        r = rr[r].get('NextRow', 0)
        assert self.rows.count( r ) == 0, 'Loop in report row pointers' 

# ----

def GetColumnList(reportid):  # returns list of column id's in the current order
    ids = []
    loopcheck = 0
    k = Report[reportid].get('FirstColumn', 0)
    while k != 0 and k != None:
        ids.append(k)
        k = ReportColumn[k].get('NextColumn', 0)
        loopcheck += 1
        if loopcheck > 100000:  break  # prevent endless loop
    return ids

def ReorderReportColumns(reportid, ids):
    """ Use the ids to resequence report columns for report 
        Any ids not included in the new list are appended in the same sequence as before
        Deleted columns are omitted.  (Remove deleted columns by calling w/ empty 'ids' list)
    """
    if debug: print "Start ReorderReportColumns"
    newseq = []
    keys = {  }# collect id's of all valid columns
    for k, v in ReportColumn.iteritems():  
        status = v.get('zzStatus')
        if v.get('ReportID') == reportid and (status == 'active' or status == None):
            keys[k] = None
    # make sure all id's in new list are valid, remove duplicates
    for k in ids:
        if keys.has_key(k):
            newseq.append(k)
            del keys[k]
    # append any remaining report column records at end
    loopcheck = 0
    k = Report[reportid].get('FirstColumn', 0)
    while k != 0 and k != None:
        if keys.has_key(k):  # valid and not already in list
            newseq.append(k)
            del keys[k]
        k = ReportColumn[k].get('NextColumn', 0)
        loopcheck += 1
        if loopcheck > 100000:  break  # prevent endless loop
    # newseq.extend(keys.keys())  # if not in new or prior list, just ignore it silently
    # perhaps in the future these should be 'deleted'???
    newseq.append(0)  # end of list marker

    # apply the new sequence to the report Columns
    newid = newseq.pop(0)
    if Report[reportid].get('FirstColumn', 0) != newid:
        change = {'Table': 'Report', 'ID': reportid, 'FirstColumn': newid }
        Update(change)
    change = {'Table': 'ReportColumn', 'ID': None, 'NextColumn': None }
    while len(newseq) > 0:
        prior = newid
        newid = newseq.pop(0)
        if ReportColumn[prior].get('NextColumn', 0) != newid:
            change['ID'] = prior
            change['NextColumn'] = newid
            Update(change)
    # calling routine must call 'SetUndo'
    if debug: print "End ReorderReportColumns"

# ----

def GetRowList(reportid):  # returns an order list or row id's in the current order
    rowids = []
    loopcheck = 0
    k = Report[reportid].get('FirstRow', 0)
    while k != 0 and k != None:
        rowids.append(k)
        k = ReportRow[k].get('NextRow', 0)
        loopcheck += 1
        if loopcheck > 100000:  break  # prevent endless loop
    return rowids

def ReorderReportRows(reportid, rowids):
    """ Use the rowids to resequence report rows for report 
        Any rowids not included in the list are appended in the same sequence as before
    """
    if debug: print "Start ReorderReportRows"
    newseq = []
    keys = {}
    for k, v in ReportRow.iteritems():  
        if v.get('ReportID') == reportid:
            keys[k] = None
    # make sure all id's are valid, remove duplicates
    for k in rowids:
        if keys.has_key(k):
            newseq.append(k)
            del keys[k]
    # append any remaining reportrow records
    loopcheck = 0
    k = Report[reportid].get('FirstRow', 0)
    while k != 0 and k != None:
        if keys.has_key(k):
            newseq.append(k)
            del keys[k]
        k = ReportRow[k].get('NextRow', 0)
        loopcheck += 1
        if loopcheck > 100000:  break  # prevent endless loop
    newseq.extend(keys.keys())  # anything not in prior list is appended to the new one
    newseq.append(0)  # end of list marker

    # apply the new sequence to the report rows
    newid = newseq.pop(0)
    if Report[reportid].get('FirstRow', 0) != newid:
        change = {'Table': 'Report', 'ID': reportid, 'FirstRow': newid }
        Update(change)
    change = {'Table': 'ReportRow', 'ID': None, 'NextRow': None }
    while len(newseq) > 0:
        prior = newid
        newid = newseq.pop(0)
        if ReportRow[prior].get('NextRow', 0) != newid:
            change['ID'] = prior
            change['NextRow'] = newid
            Update(change)
    # calling routine must call 'SetUndo'
    if debug: print "End ReorderReportRows"

# ----

def AdjustReportRows(): 
    """
# Reports rows may be affected by table adds or deletions
#   this routine makes sure that every report has the right number of rows with 
#   the right links to rows in other tables
# 1- build a list of all of the data that should appear in the report
# 2- remove from that list everything that currently appears in the report
# 3- add additional rows for everything that doesn't
# question: should I have a flag in report rows for deleted records?
# answer: no, let the individual reports check. they have to scan through all the rows anyway for hidden rows
#
# Don't add report rows for deleted records. This is to fix a bug of Insert, Undo, Redo which was adding
#   a second report row for the same record.
# Also not adding rows for children of deleted records. Don't know if this is necessary.
    """
    if debug: print "Start AdjustReportRows"
    newrow = {'Table': 'ReportRow', 'TableName': None, 'TableID': None, 'NextRow': None}
    oldrow = {'Table': 'ReportRow', 'ID': None, 'NextRow': None}
    oldreport = {'Table': 'Report', 'ID': None, 'FirstRow': None}

    ta, tb, selcol, selval = None, None, None, None  # variables used in selection filters
        # I'm not sure whether I need to mention them here.
        # This is a precaution because I want to make sure the variable scope rules let the functions use them.
    def sela(key):  # if the selection column has the right value, keep it
        return Database[ta][key].get(selcol) == selval

    def selb(key):  # if selection column in b's parent has the right value, keep it
        if debug: print "ta, tb, key, selcol", ta, tb, key, selcol
        parentid = Database[tb][key].get(ta + 'ID')
        if debug: print 'parentid', parentid
        return Database[ta][parentid].get(selcol) == selval

    for rk, r in Report.iteritems():  # process all active reports
        if r.get('zzStatus', 'active') == 'deleted': continue
        newrow["ReportID"] = rk  # make sure any new rows know their parent report is
        rtid = r.get('ReportTypeID')
        if not rtid or not ReportType.has_key(rtid):
            if debug: print "invalid ReportType: report", rk, ", type", rtid
            continue
        rt = ReportType[rtid]
        ta, tb = rt.get('TableA'), rt.get('TableB')  # get the names of the tables used
        if not ta: continue  # silently skip report records w/ no primary table
        if tb:  # if the report uses two tables
            # make lists of everything that should be in report
            selcol, selval = r.get('SelectColumn'), r.get('SelectValue')
            if selcol:
                shoulda = filter(sela, Database[ta].keys())
                shouldb = filter(selb, Database[tb].keys())
            else:
                shoulda = Database[ta].keys()
                shouldb = Database[tb].keys()
            if debug: print "tablea", ta, "shoulda", shoulda
            if debug: print "tableb", tb, "shouldb", shouldb
            # compare that with everything that is already there
            iref = {}  # insert row references - where to insert table 'b' rows for each table 'a' record
            loopguard = 0
            rowk = r.get('FirstRow', 0)
            saverowk = 0  # this will stay 0 if no report rows
#            while rowk != 0:    # testing change 040717
            while rowk and ReportRow.has_key(rowk):
                saverowk = rowk
                rr = ReportRow[rowk]
                t = rr.get('TableName'); id = rr.get('TableID')
                if t == ta:
                    try:
                        shoulda.remove(id)  # remove from list all that have rows
                    except ValueError:  # occurs if something already in the list shouldn't be there
                        if debug: print "id that doesn't belong was found in row (a of ab)", id
                    iref[id] = rowk  # insertion at beginning (may be overridden in elif)
                    tak = id # most recent table 'a' id, used for insertion at end
                elif t == tb:
                    try:
                        shouldb.remove(id)  # remove from list all that have rows
                    except ValueError:  # occurs if something already in the list shouldn't be there
                        if debug: print "id that doesn't belong was found in row (b of ab)", id
                    iref[tak] = rowk  # current guess at list end
                # else: pass  # silently ignore any invalid table references
                            # I am also ignoring any incorrect report rows that don't appear in the linked list
                rowk = rr.get('NextRow', 0)
                # saverowk points to the last row in the report
                loopguard += 1
                if loopguard > 10000: break

            # add report rows for everything that is missing
            newrow['TableName'] = ta; newrow['NextRow'] = 0
            for ka in shoulda:
                if Database[ta][ka].get("zzStatus") == "deleted": continue # -- don't add rows for deleted
                newrow['TableID'] = ka
                undo = Update(newrow, 0)  # don't allow undo on these
                if saverowk == 0:  # this is the first report row
                    oldreport['ID'] = rk; oldreport['FirstRow'] = undo['ID']
                    Update(oldreport, 0)
                else:
                    oldrow['ID'] = saverowk; oldrow['NextRow'] = undo['ID']
                    Update(oldrow, 0)
                saverowk = undo['ID']
                iref[ka] = saverowk  # insert 'b's after this 'a'

            newrow['TableName'] = tb;
            tableb = Database[tb]
            for kb in shouldb:
                if tableb[kb].get("zzStatus") == "deleted": continue # -- don't add rows for deleted
                tak = tableb[kb].get(ta + "ID")  # table 'a' id record of 'b's parent
                if not tak: continue  # if the key is NULL or 0
                if Database[ta][tak].get("zzStatus") == "deleted": continue # -- don't add rows for deleted
                insertafter = iref[tak]  # insert after this row
                nextr = ReportRow[insertafter].get('NextRow', 0)
                newrow['TableID'] = kb; newrow['NextRow'] = nextr
                undo = Update(newrow, 0)
                oldrow['ID'] = insertafter; oldrow['NextRow'] = undo['ID']
                Update(oldrow, 0)
                iref[tak] = undo['ID']

        else:  # if the report uses one table
            # make lists of everything that should be in report
            selcol, selval = r.get('SelectColumn'), r.get('SelectValue')
            if selcol:
                shoulda = filter(sela, Database[ta].keys())
            else:
                shoulda = Database[ta].keys()

            if debug: print "tablea", ta, "shoulda", shoulda
            # compare that with everything that is already there
            rowk = r.get('FirstRow', 0)
            loopguard = 0
            saverowk = 0  # this will stay 0 if no report rows
#            while rowk != 0:  # ditto debug change 040717
            while rowk and ReportRow.has_key(rowk):
                saverowk = rowk
                rr = ReportRow[rowk]
                id = rr.get('TableID')
                try:
                    shoulda.remove(id)  # remove from list all that have rows
                except ValueError:  # occurs if something already in the list shouldn't be there
                    if debug: print "id that doesn't belong was found in row", id
                rowk = rr.get('NextRow', 0)
                loopguard += 1
                if loopguard > 10000: break
            # saverowk points to the last row in the report

            # add report rows for everything that is missing
            newrow['TableName'] = ta; newrow['NextRow'] = 0
            for ka in shoulda:
                if Database[ta][ka].get("zzStatus") == "deleted": continue # -- don't add rows for deleted
                newrow['TableID'] = ka
                undo = Update(newrow, 0)  # don't allow undo on these
                if saverowk == 0:  # this is the first report row
                    oldreport['ID'] = rk; oldreport['FirstRow'] = undo['ID']
                    Update(oldreport, 0)
                else:
                    oldrow['ID'] = saverowk; oldrow['NextRow'] = undo['ID']
                    Update(oldrow, 0)
                saverowk = undo['ID']
    if debug: print "End AdjustReportRows"
# end of AdjustReportRows

# --------- Routines to load and save data ---------
OptionFile = None

def LoadOption(directory=None):
    """ Load the option file.
    """
    global OptionFile, Option
    if directory: opDir = directory
    else: opDir = os.getcwd()

    OptionFile = os.path.join(opDir, "Options.ganttpvo")
    if debug: print "loading option file", OptionFile
    try: 
        f = open(OptionFile, "rb")
    except IOError:
        if debug: print "LoadOption io error"
    else:
        header = f.readline()  # header will identify need for conversion of earlier versions or use of different file formats
        Option = cPickle.load(f)
        f.close()
        Menu.GetScriptNames()

def SaveOption():
    """ Save the Option file.
    """
    if debug: print "saving option file", OptionFile
    try: 
        f = open(OptionFile, "wb")
    except IOError:
        if debug: print "SaveOption io error"
    else:
        f.write("GanttPV\t0.1\tO\n")
        cPickle.dump(Option, f)
        f.close()

# --------- Project files

def OpenReport(id):
    """ Open a report and bring it to the front.  """
    r = Report.get(id)
    if not r or r.get('zzStatus') == 'deleted':
        return
    r['Open'] = True
    if not OpenReports.get(id):
        pos = (r.get('FramePositionX') or -1, r.get('FramePositionY') or -1)
        size = (r.get('FrameSizeW') or 768, r.get('FrameSizeH') or 311)
        OpenReports[id] = GanttReport.GanttReportFrame(id, None, -1, "", pos, size)

#         frame = GanttReport.GanttReportFrame(id, None, -1, "")
#         if r.get('FramePositionX') and r.get('FramePositionY'):
#             frame.SetPosition(wx.Point(r['FramePositionX'], r['FramePositionY']))
#         if r.get('FrameSizeW') and r.get('FrameSizeH'):
#             frame.SetSize(wx.Size(r['FrameSizeW'], r['FrameSizeH']))
#         OpenReports[id] = frame

        Menu.UpdateWindowMenuItem(id)
        OpenReports[id].Show(True)
    OpenReports[id].Raise()

def CloseReport(id):
    """ Close a report. """
    if OpenReports.has_key(id):
        del Report[id]['Open']
        OpenReports[id].Destroy()
        del OpenReports[id]
        Menu.UpdateWindowMenuItem(id)

def CloseReports():
    """ Close all reports except #1. """
    for id, frame in OpenReports.items():
        if id == 1: continue
        del Report[id]['Open']
        frame.Destroy()
        del OpenReports[id]
    Menu.ResetWindowMenus()

def MakeReady():
    """ After an database has been loaded or created, set all other values to match """
    global ChangedData, ChangedCalendar, ChangedSchedule, ChangedReport, UndoStack, RedoStack
    global Project, Task, Dependency, Report, ReportColumn, ReportRow
    global Resource, Assignment, Holiday, ReportType, ColumnType, OtherData, Other, NextID

    Project =       Database['Project']
    Task =          Database['Task']
    Dependency =    Database['Dependency']
    Report =        Database['Report']
    ReportColumn =  Database['ReportColumn']
    ReportRow =     Database['ReportRow']
    Resource =      Database['Resource']
    Assignment =    Database['Assignment']
    Holiday =       Database['Holiday']
    ReportType =    Database['ReportType']
    ColumnType =    Database['ColumnType']
    OtherData =     Database['OtherData']  # used in updates
    Other =         OtherData[1]
    NextID =        Database['NextID']

    SetupDateConv()
    GanttCalculation()
    AdjustReportRows()  # probably not needed, but the report frames need to set themselves up when opened
    # open all 'open' reports
    for k, v in Report.iteritems():
        if k == 1:
            if OpenReports.has_key(1):
                frame = OpenReports[1]
                frame.UpdatePointers(1) # new database
                Menu.AdjustMenus(frame)
                frame.Refresh()
                frame.Report.Refresh()  # update displayed data (needed for Main on Windows, not needed on Mac)
        if Database['Report'][k].get('Open'):  OpenReport(k)
    UndoStack  = []
    RedoStack = []
    ChangedData = False  # true if database needs to be saved
    ChangedCalendar, ChangedSchedule, ChangedReport, ChangedRow = False, False, False, False

#    RefreshReports()  # needed?? -- not here, this routine is back end

def LoadContents():
    """ Load the contents of our document into memory. """
    global Database
    CloseReports()
    try:
        f = open(FileName, "rb")
        header = f.readline()  # add read line of text - will allow conversion of earlier versions or use of different file formats
        Database = cPickle.load(f)
        f.close()
    except IOError:
        if debug: print "LoadContents io error"
    else:
        MakeReady()
        return True

def SaveContents():
    """ Save the contents of our document to disk. """
    global ChangedData

    f = open(FileName, "wb")
    # add write line of text 
    f.write("GanttPV\t0.1\ta\n")
    cPickle.dump(Database, f)
    f.close()

    ChangedData = False

def GetScriptDirectory():
    """ Return the directory to search for script files """
    return Option.get('ScriptDirectory') or os.path.join(Path, "Scripts")

def OpenFile(path):
    """ Open a file (any type) """
    ext = os.path.splitext(path)[1]
    if ext == '.ganttpv':
        OpenDatabase(path)
    elif ext == '.py':
        RunScript(path)
    else:
        if debug: print "unknown file extension:", ext

def OpenDatabase(path):
    """ Open a database file (.ganttpv) """
    global FileName
    FileName = path
    if not LoadContents(): return
    title = os.path.basename(path)
    OpenReports[1].SetTitle(title)
    OpenReports[1].Show(True)

def RunScript(path):
    """ Run a script file (.py) """
    if debug: print "begin script:", path

    name_dict = GetModuleNames()
    name_dict['self'] = OpenReports.get(ActiveReport)
    name_dict['thisfile'] = path
    name_dict['debug'] = debug

    try:
        execfile(path, name_dict)
    except:
        import sys
        error_info = sys.exc_info()
        sys.excepthook(*error_info)

    if debug: print "end of script:", path

def AskIfUserWantsToSave(action):
    """ Give the user the opportunity to save the current document.

    'action' is a string describing the action about to be taken.  If
    the user wants to save the document, it is saved immediately.  If
    the user cancels, we return False.
    """
    global FileName
    if not ChangedData: return True # Nothing to do.

    response = wx.MessageBox("Save changes before " + action + "?",
                                "Confirm", wx.YES_NO | wx.CANCEL)

    if response == wx.YES:
        if FileName == None:
            tempFileName = wx.FileSelector("Save File As", "Saving",
                                            default_filename='Untitled.ganttpv',
                                            default_extension="ganttpv",
                                            wildcard="*.ganttpv",
                                            flags = wx.SAVE | wx.OVERWRITE_PROMPT)
            if tempFileName == "": return False # User cancelled.
            FileName = tempFileName

        SaveContents()
        return True
    elif response == wx.NO:
        return True # User doesn't want changes saved.
    elif response == wx.CANCEL:
        return False # User cancelled.

# ------- setup data for testing

SetEmptyData()
# SetSampleData()
MakeReady()

if debug: print "end Data.py"
