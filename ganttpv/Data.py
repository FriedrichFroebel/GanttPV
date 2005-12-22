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
# 040414 - renamed file to Data; added load and store routines; corrections to Update; reserved ReportID's 1 & 2 for 'Project/Report List' and 'Resource List'
# 040415 - moved menu adjust logic here from Main.py; made various changes to correct or improve logic
# 040416 - changed gantt calculation to process all projects; wrote Undo and Redo routines
# 040417 - added optional parameter to Update to specify whether change should be added to UndoStack; in CheckChange now sets ChangedReport flag if zzStatus appears; moved flags and variables to front of file this; added MakeReady to do all the necessary computations after a new database is loaded; added AdjustReportRows to add report rows when records are added
# 040419 - reworked AdjustReportRows logic
# 040420 - added routines to open and close reports; revised MakeReady
# 040422 - fix so rows added by AdjustReportRows will always have a link to their parent; added row type colors  for Project/Report frame
# 040423 - added ReorderReportRows; moved AutoGantt to Option; added ConfirmScripts to Option
# 040424 - added ScriptDirectory to Option; added LoadOption and SaveOption; added GetRowList; fixed reset error in date tables;
# 040426 - fixed Recalculate to check ChangedReport; added check for None to GetRowList and ReorderReportRows; added table Assignment; in GanttCalculation ignore deleted Dependency records
# 040427 - in Resource table changed 'LongName' to 'Name'; added Width to ReportColumn; revised ReportColumn fields
# 040503 - added ReportType and ColumnType tables; moved UpdateDataPointers here from Main and GanttReport; added "Prerequisite" synonym for "Task" in Database
# 040504 - added default Width to ColumnType; in gantt calculation, will ignore tasks w/o projects; added GetColumnList and ReorderReportColumns; added GetToday
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
# 050513 - Alexander - revised some dictionary fetches to ignore invalid keys (instead of raising an exception); tightened _Do logic
# 050519 - Brian - use TaskID instead of ParentTaskID to designate Task parent.
# 050521 - Alexander - fixed work week bug in SetupDateConv
# 050527 - Alexander - added 'Generation' column to designate levels in the task-parenting heirarchy; updated in GanttCalculation
# 050531 - Brian - in AdjustReportRows test to make sure parent exists before testing parent's select column value
# 050716 - Brian - in GetCellValue use more general routine to find period
# 050806 - Alexander - rewrote date conversion!
# 050814 - Alexander - added SearchByColumn
# 050826 - Alexander - rewrote row/column ordering!
# 050903 - Alexander - fixed bug in AddRow (was aborting when table was empty)
# 051121 - Brian - add FileSignature number

import datetime, calendar
import cPickle
import wx, os, sys
import Menu, GanttReport
import random

debug = 1
if debug: print "load Data.py"

# On making changes in GanttPV
# 1- Use Update for all changes
# 2- CheckChange will decide the impact of the changes
# 3- Use SetUndo to tell GanttPV to fix anything in response to the changes

def GetModuleNames():
    """ Return the GanttPV modules as a namespace dictionary.

    This dictionary should be passed to any scripts run with execfile.
    """
    import Data, GanttPV, GanttReport, ID, Menu, UI, wx
    return locals()

App = None  # the application itself
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

# save impact of change until it can be addressed
ChangedCalendar = False  # will affect calendar hours tables (gantt will have to be redone, too)
ChangedSchedule = False  # gantt calculation must be redone
ChangedReport = False    # report layout needs to be redone
ChangedRow = False       # a record was created, the rows on a report may be affected

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
            'WorkDaySelected': (153, 204, 255),  # make this lighter  # maybe use white only for the selected column?
            'NotWorkDay': (204, 204, 204),  # more light focused on the selection
            'NotWorkDaySelected': (136, 187, 238),
# bars
            'PlanBar': (0, 153, 102),   # green  # maybe use same color for selected & not
            'PlanBarSelected': (0, 153, 102),
            'ActualBar': (0, 0, 255),    # yellow?
            'ActualBarSelected': (0, 0, 255),
            'BaseBar': (0, 255, 0),   # green?
            'BaseBarSelected': (0, 255, 0),
            'CompletionBar': None,  # based on BaseBar or PlanBar
            'CompletionBarSelected': None,

#           'ResourceProblem': None,
#           'ResourceProblemSelected': None,
        }

# -------------- Setup Empty and Sample databases

# reserved table column names: 'Table', 'ID', 'zzStatus', anything ending with 'ID", anything starting with 'xx'

def SetOther():
    """ These are the values that only occur once in the database """
    OtherData[1] = { 'ID' : 1,
        'WeekHours' : (8, 8, 8, 8, 8, 0, 0),  # hours per day, MTWHFSS
        }
    Database['OtherData'] = OtherData
    Database['NextID'] = NextID
    Database['Prerequisite'] = Database['Task']  # synonym for Task

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
             ( 3, 'i', 'd', 'A', False, 80, 'ProjectID', 'Project ID' ),
             ( 3, 't', 'd', 'A', True,  140, 'Name' ),
             ( 3, 'd', 'd', 'A', True,  80, 'StartDate', 'Start Date' ),
             ( 3, 'i', 'd', 'A', True,  80, 'DurationHours', 'Duration' ),
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
             # ( 6, 't', 'd', 'a', False, 80, 'TableA', 'Table A' ),
             ( 6, 't', 'd', 'A', False, 80, 'SelectColumn', 'Select Column' ),
             ( 6, 'i', 'd', 'A', False, 80, 'SelectValue', 'Select Value' ),
             ( 6, 'i', 'd', 'A', False, None, 'ProjectID', 'Project ID' ),
             ( 7, 'i', 'd', 'B', False, None, 'ID' ),                # Report/ReportColumn
             ( 7, 'i', 'd', 'B', True,  40, 'Width' ), 
             ( 7, 't', 'd', 'B', True,  140, 'Label' ),
             # ( 7, ' 'A', 'TypeA', 'B', 'TypeB', 'Time', 
             ( 7, 'i', 'd', 'B', True,  40, 'Periods' ),
             ( 7, 'd', 'd', 'B', True,  80, 'FirstDate', 'First Date' ),
             ( 8, 'i', 'd', 'A', False, None, 'ID' ),                # Resource
             ( 8, 't', 'd', 'A', True,  80, 'ShortName', 'Short Name' ),
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
             ( 11, 't', 'd', 'A', False, 100, 'TableA', 'Table A' ),
             ( 11, 't', 'd', 'A', False, 100, 'TableB', 'Table B' ),
             ( 11, 't', 'd', 'A', True,  120, 'Label' ),
             ( 12, 'i', 'd', 'B', False, None, 'ID' ),               # ColumnType
             # ( 12, 'i', 'd', 'B', False, None, 'ReportTypeID', 'Report Type ID' ),
             ( 12, 't', 'd', 'B', False, None, 'DataType', 'Data Type' ),
             ( 12, 't', 'd', 'B', False, None, 'AccessType', 'Access Type' ),
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
             ( 11, 4, 9, 80,         'Duration',             'DurationHours', 'INT',         '','',          '', 0, '' ),
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

def AddTable(name):
    """ Add the table if it doesn't already exist """
    # (Note: this isn't reversed by Undo.)
    if name and name not in Database:
        Database[name] = {}
        Database['NextID'][name] = 1

def AddRow(change):
    """ Add or update row """
    changeTable = Database.get(change.get("Table")) or {}
    changeName = change.get("Name")
    for k, v in changeTable.iteritems():
        if v.get('Name') == changeName:
            change['ID'] = k
            break
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
    if not Database.has_key(table): return 0
    if not field2:
        for k, v in Database[table].iteritems():
            if (v.get(field1) == value1): return k
    else:
        for k, v in Database[table].iteritems():
            if (v.get(field1) == value1) and (v.get(field2) == value2): return k

def FindIDs(table, field1, value1, field2, value2):  # deprecated
    if not Database.has_key(table): return []
    result = []
    if not field2:
        for k, v in Database[table].iteritems():
            if (v.get(field1) == value1): result.append(k)
    else:
        for k, v in Database[table].iteritems():
            if (v.get(field1) == value1) and (v.get(field2) == value2): result.append(k)
    return result

def SearchByColumn(table, search, max=None):
    """ Search the table for specific column values

    table -- a dictionary of records
    search -- a dictionary that maps column names to values
    max -- if given, return at most that many records

    Return a subset of the table: a dictionary that contains the matching
    records by ID.  To match, a record must match every item in the search.
    Note that a value of None matches the absence of a value.

    """
    result = {}
    for id, record in table.iteritems():
        for key, val in search.iteritems():
            if record.get(key) != val: break
        else:
            result[id] = record
            if max and len(result) >= max: break
    return result

def CheckChange(change):
    """ Check for important changes

    change -- the undo info for the changes (only changed columns included)
    """
    global ChangedCalendar, ChangedSchedule, ChangedReport, ChangedRow
    if debug: print 'CheckChange:', change
    if not change.has_key('Table'):
        if debug: print "change does not specify table"
        return

    if ChangedRow: pass
    elif 'zzStatus' in change:  # something has been added or deleted
        ChangedRow = True
    else:
        for k in change:
            if k[-2:] == 'ID' and len(k) > 2:  # foreign key was changed
                ChangedRow = True; break

    if ChangedCalendar: pass 
    elif change['Table'] == 'OtherData':
        for k in ('WeekHours',):
            if change.has_key(k): ChangedCalendar = True; break
    elif change['Table'] == 'Holiday':
        for k in ('Date', 'Hours', 'zzStatus'):
            if change.has_key(k): ChangedCalendar = True; break

    elif ChangedSchedule: pass
    elif change['Table'] == 'Task':
        for k in ('StartDate', 'DurationHours', 'zzStatus', 'ProjectID', 'TaskID'):
            if change.has_key(k): ChangedSchedule = True; break
    elif change['Table'] == 'Dependency':
        for k in ('PrerequisiteID', 'TaskID', 'zzStatus'):
            if change.has_key(k): ChangedSchedule = True; break
    elif change['Table'] == 'Project':
        for k in ('zzStatus', 'StartDate'):
            if change.has_key(k) :  ChangedSchedule = True; break

    if ChangedRow or ChangedReport: pass
    elif change['Table'] == 'ReportRow':
        for k in ('NextRow', 'Hidden'):
            if change.has_key(k): ChangedReport = True; break
    elif change['Table'] == 'ReportColumn':
        for k in ('Type', 'NextColumn', 'Time', 'Periods', 'FirstDate'):
            if change.has_key(k): ChangedReport = True; break
    elif change['Table'] == 'Report':
        for k in ('Name', 'FirstColumn', 'FirstRow', 'ShowHidden', 'zzStatus'):
            if change.has_key(k): ChangedReport = True; break

def RefreshReports():
    if debug: print "Start RefreshReports"
    for k, v in OpenReports.items():  # deleted reports are closed during this loop
        if debug: print 'reportid', k
        if not v: continue
        if Report[k].get('zzStatus') == 'deleted':
            CloseReport(k)
        if k != 1: v.SetReportTitle()  # only grid reports
        Menu.AdjustMenus(v)
        v.Refresh()
        v.Report.Refresh()  # update displayed data (needed for Main on Windows, not needed on Mac)
    if debug: print "End RefreshReports"

# ----- undo and redo

def Recalculate(autogantt=True):
    global ChangedCalendar, ChangedSchedule, ChangedReport

    UpdateCalendar = ChangedCalendar and autogantt
    UpdateGantt = UpdateCalendar or (ChangedSchedule and autogantt)
    UpdateReports = ChangedReport or ChangedRow

    # these routines shouldn't add anything to the undo stack
    if UpdateCalendar:
        SetupDateConv(); ChangedCalendar = False
    if UpdateGantt:
        GanttCalculation(); ChangedSchedule = False
    if UpdateReports:
        for v in OpenReports.values():
            if v: v.UpdatePointers()
        ChangedReport = False

    RefreshReports()

def SetUndo(message):
    """ This is the last step in submitting a group of changes to the database. 
    Adjust calculated values and update the displays.
    """
    global ChangedReport, ChangedRow, ChangedData, UndoStack, RedoStack
    if debug: print "Start SetUndo"
    if debug: print "message", message

    if ChangedRow:
        AdjustReportRows(); ChangedReport = True; ChangedRow = False

    UndoStack.append(message)
    RedoStack = []  # clear out the redo stack
    ChangedData = True  # file needs to be saved

    autogantt = Option.get('AutoGantt')
    Recalculate(autogantt)

    if debug: print "End SetUndo"

def _Do(fromstack, tostack):
    global ChangedData, ChangedRow
    if fromstack and isinstance(fromstack[-1], str):
        savemessage = fromstack.pop()
        while fromstack and isinstance(fromstack[-1], dict):
            change = fromstack.pop()
            redo = Update(change, 0)  # '0' means don't put change into Undo Stack
            tostack.append(redo)
        tostack.append(savemessage)
        Recalculate()
        ChangedRow = False  # already handled by last SetUndo
        ChangedData = True  # file needs to be saved

def DoUndo():
    _Do(UndoStack, RedoStack)

def DoRedo():
    _Do(RedoStack, UndoStack)

# --------------------------

# update routine
def Update(change, push=1):
    global ChangedRow
    if debug: print 'Update:', change
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
        id = undo['ID'] = change['ID']
        if id not in table:
            if debug: print 'change specifies invalid record:', id
            raise KeyError

        record = table[id]
        for c, newval in change.iteritems():  # process each field
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
        record = {}
        for c, newval in change.iteritems():  # process each field
            if newval or newval == 0:
                record[c] = newval
        undo['zzStatus'] = 'deleted'
        id = NextID[tname]
        if debug: print "Added new record:", id
        NextID[tname] = id + 1
        record['ID'] = undo['ID'] = id
        CheckChange(record)
        ChangedRow = True  # CheckChange doesn't recognize this for new records
        del record['Table']  # saves space
        table[id] = record
    if push: UndoStack.append(undo)
    return undo

########## @@@@@@@@@@ Start Alex Date Conversion @@@@@@@@@@ ########## 

# calendar setup

def SetupDateConv(): 
    UpdateWorkWeek()
    ReadHolidays()
    UpdateHolidayHours()

def UpdateWorkWeek():
    global WorkWeek, WeekSize, CumWeek
    global HoursPerWeek, DaysPerWeek, HoursPerDay, AllowDaysUnit
    global ChangedCalendar
    if debug: print 'start UpdateWorkWeek'
    Other = Database['OtherData'][1]

    WorkWeek = list(Other.get('WeekHours') or (8, 8, 8, 8, 8, 0, 0)) 
    WeekSize = len(WorkWeek)

    CumWeek = []
    HoursPerWeek = DaysPerWeek = 0
    for day in WorkWeek:
        CumWeek.append(HoursPerWeek)
        if day:
            HoursPerWeek += day
            DaysPerWeek += 1

    HoursPerDay = Other.get('HoursPerDay')
    if HoursPerDay:
        AllowDaysUnit = True
    else:
        HoursPerDay = HoursPerWeek.__truediv__(DaysPerWeek)
        AllowDaysUnit = (DaysPerWeek == WorkWeek.count(HoursPerDay))

    if debug: print 'end UpdateWorkWeek'

HolidayMap = {}
HolidayDate = []
HolidayHour = []
HolidayAdjust = []

def ReadHolidays():
    """ Read the holiday dates and lengths from the database """
    global HolidayMap, HolidayDate
    if debug: print 'start ReadHolidays'
    HolidayMap = {}
    for r in Database['Holiday'].itervalues():
        if r.get('zzStatus') == 'deleted': continue
        datestr = r.get('Date')
        try:
            date = StringToDate(datestr)
        except ValueError:
            continue
        hours = r.get('Hours')
        if hours:
            HolidayMap[date] = hours
        else:
            HolidayMap[date] = 0
    HolidayDate = HolidayMap.keys()
    HolidayDate.sort()
    if debug: print 'end ReadHolidays'

def UpdateHolidayHours():
    """ Prepare the holiday conversions """
    global HolidayHour, HolidayAdjust
    if debug: print 'start UpdateHolidayHours'
    HolidayHour = []
    HolidayAdjust = [0]
    adjust = 0
    for date in HolidayDate:
        w, dow = divmod(date, WeekSize)
        newLength = HolidayMap[date]
        oldLength = WorkWeek[dow]
        hour = w * HoursPerWeek + CumWeek[dow] + adjust + min(newLength, oldLength, 0)
        adjust += newLength - oldLength
        HolidayHour.append(hour)
        HolidayAdjust.append(adjust)
    if debug: print 'end UpdateHolidayHours'


# date / hour conversion

def DateToHours(date):
    """ Convert a date from days to hours """
    key = CutByValue(HolidayDate, date)
    week, dow = divmod(date, WeekSize)
    hour = week * HoursPerWeek + CumWeek[dow] + HolidayAdjust[key]
    if key > 0:
        hour = max(hour, HolidayHour[key-1])
    return hour

def HoursToDate(hours):
    """ Convert a date from hours to days """
    key = CutByValue(HolidayHour, hours)
    h = hours - HolidayAdjust[key]
    if (key > 0) and (h <= HolidayHour[key-1]):
        date = HolidayDate[key-1]
        hours -= HolidayHour[key-1]
    else:
        week, hours = divmod(h, HoursPerWeek)
        dow = CutByValue(CumWeek, hours) - 1
        date = week * WeekSize + dow
        hours -= CumWeek[dow]
    return date, hours

def HoursToDateString(hours):
    date, hours = HoursToDate(hours)
    return DateToString(date), hours

def CutByValue(list, value):
    """ Return the highest index such that list[:index][k] <= value

    list -- an ascending sequence
    value -- the search value

    """
    start = 0
    end = len(list)
    while start < end:
        center = (start + end) >> 1
        if list[center] <= value:
            start = center + 1
        else:
            end = center
    return start


base_date_object = datetime.date(2001, 1, 1)
BaseDate = base_date_object.toordinal()
BaseYear = base_date_object.year

Years_per_Calendar_Cycle = 400  # Gregorian calendar
Days_per_Calendar_Cycle = 146097

DateFormat = "%04d-%02d-%02d"
NegDateFormat = "-" + DateFormat


# the current date

def TodayDate():
    y, m, d = _get_today()
    return _ymd_to_date(y, m, d)

def TodayString():
    y, m, d = _get_today()
    return _ymd_to_str(y, m, d)

GetToday = TodayString

def _get_today():
    dateobj = datetime.date.today()
    return _date_tuple(dateobj)

def _date_tuple(dateobj):
    return dateobj.year, dateobj.month, dateobj.day


# parsing entered dates

def CheckDateString(s):
    try:
        y, m, d = _user_str_to_ymd(s)
        _ymd_to_date(y, m, d)
    except ValueError:
        return ""
    return _ymd_to_str(y, m, d)

def _user_str_to_ymd(s):
    if not s:
        raise ValueError

    if s == '=':
        return _get_today()
    if s[0] == '*':
        today = TodayDate()
        dow = int(s[1:]) - 1
        date = today + (dow - today) % WeekSize
        return _date_to_ymd(date)

    parts = s.split('-')
    if s[0] == '-':
        parts[:2] = ['-' + parts[1]]

    ymd = [int(p) for p in parts]

    if s[0] in ('+', '-'):
        while len(ymd) < 3:
            ymd.append(1)
    elif len(ymd) < 3:
        ymd[:0] = _get_today()[:-len(ymd)]
    else:
        magnitude = 10 ** len(parts[0])
        currentyear = _get_today()[0]

        if currentyear < 0:
            diff = (-currentyear % magnitude) - ymd[0]
        else:
            diff = ymd[0] - (currentyear % magnitude)
        year = currentyear + diff

        if diff < 0:
            if (-diff > magnitude / 2) and not (0 < -year < magnitude):
                year += magnitude
        elif (diff > magnitude / 2) and not (0 < year < magnitude):
            year -= magnitude

        ymd[0] = year

    return ymd


# string to date conversion

def StringToDate(s):
    y, m, d = _str_to_ymd(s)
    date = _ymd_to_date(y, m, d)
    return date

def _str_to_ymd(s):
    if not s:
        raise ValueError
    parts = s.split('-')
    if s[0] == '-':
        parts[:2] = ['-' + parts[1]]
    ymd = [int(p) for p in parts]
    return ymd

def _ymd_to_date(year, month, day):
    cycles, year = divmod(year - BaseYear, Years_per_Calendar_Cycle)
    dateobj = datetime.date(year + BaseYear, month, day)
    date = dateobj.toordinal() - BaseDate + cycles * Days_per_Calendar_Cycle
    return date

def _coerce_ymd_to_date(year, month, day):
    # wrap month around year (e.g. 2005-15 -> 2006-03)
    # if day is too high, use start of following month (e.g. Feb 31 -> Mar 1)

    year += (month - 1) // 12
    month = (month - 1) % 12 + 1

    equiv_year = (year - BaseYear) % Years_per_Calendar_Cycle + BaseYear
    if day > calendar.monthrange(equiv_year, month)[1]:
        year += month // 12
        month = month % 12 + 1

    return _ymd_to_date(year, month, day)


# date to string conversion

def DateToString(date):
    y, m, d = _date_to_ymd(date)
    s = _ymd_to_str(y, m, d)
    return s

def ValidDate(s):
    try:
        StringToDate(s)
        return s
    except ValueError:
        return ""

def _date_to_ymd(date):
    cycles, date = divmod(date, Days_per_Calendar_Cycle) 
    dateobj = datetime.date.fromordinal(int(date) + BaseDate)
    y, m, d = _date_tuple(dateobj)
    y += cycles * Years_per_Calendar_Cycle
    return y, m, d

def _ymd_to_str(year, month, day):
    if year < 0:
        return NegDateFormat % (-year, month, day)
    else:
        return DateFormat % (year, month, day)


# month intervals

def AddMonths(date, months):
    y, m, d = _date_to_ymd(date)
    return _coerce_ymd_to_date(y, m + months, d)


# transition objects (for backwards compatibility)

class _str_to_date:
    def __init__(self):
        pass
    def __contains__(self, s):
        return ValidDate(s)
    def __getitem__(self, s):
        return StringToDate(s)
    def has_key(self, key):
        return key in self

class _date_to_str:
    def __init__(self):
        pass
    def __getitem__(self, d):
        return DateToString(d)

class _date_info:
    def __init__(self):
        pass
    def __getitem__(self, d):
        dow = d % WeekSize
        dayhours = WorkWeek[dow]
        cumhours = DateToHours(d)
        return dayhours, cumhours, dow

class _next_month:
    def __init__(self):
        pass
    def __contains__(self, s):
        return ValidDate(s)
    def __getitem__(self, s):
        date = StringToDate(s)
        return AddMonths(date, 1)
    def has_key(self, key):
        return key in self

# DateConv = {}   # usage: index = DateConv['2004-01-01']
# DateIndex = []  # usage: date = DateIndex[1]
# DateInfo = []   # dayhours, cumhours, dow = DateInfo[1]
# DateNextMonth = {} # usage: index = DateNextMonth['2004-01']  # index is first day of next month

DateConv = _str_to_date()
DateIndex = _date_to_str()
DateInfo = _date_info()
DateNextMonth = _next_month()

########## @@@@@@@@@@ End Alex Date Conversion @@@@@@@@@@ ########## 

def GetPeriodStart(period, ix, of):
    return GetPeriodInfo(period, ix, of, 1)

def GetPeriodInfo(period, ix, of, index=0):  # needed by server
    """
    Returns: hours in period, hours index of start of period, day of week of start of period 
    """
    period = period[0:3]
    if period == "Wee":
        ix -= DateInfo[ ix ][2]  # convert to beginning of week
        ix += of * 7
        dh, cumh, dow = DateInfo[ix]
        dh2, cumh2, dow2 = DateInfo[ix + 7]
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

def GanttCalculation(): # Gantt chart calculations - all dates are in hours
    # change = { 'Table': 'Task' }  # will be used to apply updates  --->> Don't Use 'Update' here <<--
    # Set the project start date (use specified date or today)  later may be adjusted by the tasks' start dates
    Today = GetToday()
    if debug: print "today", Today
    ps = {}  # project start dates indexed by project id
    for k, v in Project.iteritems():
        if v.get('zzStatus') == 'deleted': continue
        sd = v.get('StartDate')
        # if debug: print "project", k, ", startdate", sd
        if not (sd and sd in DateConv): sd = Today  # default project start dates to today
        ps[k] = sd

    # dependencies
    pre = {}  # task prerequisites indexed by task id number
    suc = {}  # task successors
    precnt = {}  # count of all unprocessed prerequisites
    succnt = {}  # count of all unprocessed successors
    tpid = {}  # task's projectid
    parent = {}  # children indexed by parent id
    for k, v in Task.iteritems():  # init dependency counts, xrefs, and start dates
        if v.get('zzStatus') == 'deleted': continue
        pid = v.get('ProjectID')
        if pid not in Project or Project[pid].get('zzStatus') == 'deleted': continue
            # silently ignore task w/ invalid or deleted project

        precnt[k] = 0
        succnt[k] = 0
        pre[k] = []
        suc[k] = []
        tpid[k] = pid
        # if debug: "task's project", k, pid
        tsd = v.get('StartDate')
        if tsd and (tsd < ps[pid]) and (tsd in DateConv):
            # adjust project start date if task starts earlier
            ps[pid] = tsd

        # if debug: print "task data", k, v
        p = v.get('TaskID')  # parent task id
        if p != k and p in Task and Task[p].get('zzStatus') != 'deleted': 
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
        if v.get('zzStatus') == 'deleted': continue
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
                # note: the end date is not currently used to compute a start date
                tsd = Task[k].get('StartDate')
                ted = Task[k].get('EndDate')
                td  = Task[k].get('DurationHours')
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
                    ef = es + int(HoursPerDay)

                if ef > ProjectEndHour[tpid[k]]: ProjectEndHour[tpid[k]] = ef  # keep track of project end date 

                # update database -- doesn't use Update, but may in the future
                Task[k]['hES'] = es
                Task[k]['hEF'] = ef
                Task[k]['CalculatedStartDate'], Task[k]['CalculatedStartHour'] = HoursToDateString(es)
                Task[k]['CalculatedEndDate'], Task[k]['CalculatedEndHour'] = HoursToDateString(ef)

                # change['ID'] = k
                # change['hES'] = es
                # change['hEF'] = ef
                # change['CalculatedStartDate'], change['CalculatedStartHour'] = HoursToDateString(es)
                # change['CalculatedEndDate'], change['CalculatedEndHour'] = HoursToDateString(ef)
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
        if v.get('zzStatus') == 'deleted': continue
        if parent.has_key(k):
            Task[k]['SubtaskCount'] = parent[k]  # save count of child tasks
            continue  # skip parent tasks
        if Task[k].has_key('SubtaskCount'): del Task[k]['SubtaskCount']

        p = v.get('TaskID')  # parent task id
        pid = v.get('ProjectID')
        hes, hef, hls, hlf = [ v.get(x) for x in ['hES', 'hEF', 'hLS', 'hLF']]

        lineage = {}
        while p in Task:
            if p == k or p in lineage or Task[p].get('zzStatus') == 'deleted':
                break
            lineage[p] = None
            p = Task[p].get('TaskID')  # parent task id

        for p in lineage:
            # if debug: print "adjusting parent ", p, Task[p]
            phes, phef, phls, phlf = [ Task[p].get(x) for x in ['hES', 'hEF', 'hLS', 'hLF']]
            if phes == None or hes < phes:
                Task[p]['hES'] = hes
                Task[p]['CalculatedStartDate'], Task[p]['CalculatedStartHour'] = HoursToDateString(hes)
            if phef == None or hef > phef:
                Task[p]['hEF'] = hef
                Task[p]['CalculatedEndDate'], Task[p]['CalculatedEndHour'] = HoursToDateString(hef)
            if phls == None or hls < phls: Task[p]['hLS'] = hls
            if phlf == None or hlf > phlf: Task[p]['hLF'] = hlf
            # if debug: print "adjusted parent ", p, Task[p]

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

            # use this instead?  050716
            # result = GetPeriodStart(ctperiod, index, of)  # convert to beginning of desired period

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
        return result  # should test to make sure it is valid

def GetColumnHeader(colid, of):
        """
    colid == ReportColumn ID; of == the offset; returns the first and second header line
        """
        if of == -1:
            label = ReportColumn[colid].get('Label')
            if not label:
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
        if colid not in ReportColumn: return ""
        rc = ReportColumn[colid]
        ctid = rc.get('ColumnTypeID')
        if ctid not in ColumnType: return ""
        ct = ColumnType[ctid]
        if of == -1:
            rr = ReportRow[rowid]
            rtable = rr.get('TableName')
            tid = rr['TableID']  # was 'TaskID'

            t = ct.get('T') or 'X'
            reportid = rr.get('ReportID')
            if reportid not in Report: return ""
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
                    table = Database.get(listtable) or {}
                    records = SearchByColumn(table, {listselect: listvalue})
                    vals = [(r.get(listtarget) or "-") for r in records.itervalues() if r.get('zzStatus') != 'deleted']
                    if listtable2 and listcol2:
                        table = Database.get(listtable2) or {}
                        records = [table.get(v) for v in vals]
                        vals = [r.get(listcol2) or "-" for r in records if r and r.get('zzStatus') != 'deleted']
                    value = ", ".join( [ str(x) for x in vals ] )
            if dt == 'u' and isinstance(value, int) and value > 0:
                w, h = divmod(value, HoursPerWeek)
                if h > int(HoursPerWeek): w += 1; h = 0

                if AllowDaysUnit:
                    d, h = divmod(h, HoursPerDay)
                    if h > int(HoursPerDay): d += 1; h = 0
                else:
                    d = 0

                if h % 1: h += 1

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
                index = DateConv[firstdate]
                index = GetPeriodStart(ctperiod, index, of)
                date = DateIndex[index]

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
def UpdateDataPointers(self, reportid):
    # self is either a list or a table behind a grid
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

########## @@@@@@@@@@ Start Alex Row/Column Order @@@@@@@@@@ ##########

def UpdateRowPointers(self):
    # self is either a list or a table behind a grid

    reportid = self.report.get('ID')
    rows, rowlevels = GetRowLevels(reportid)

    if self.report.get('ShowHidden'):
        self.rows = rows
        self.rowlevels = rowlevels
        return

    self.rows = []
    self.rowlevels = []
    level = 0  # maximum level of next row

    for rowid, lev in zip(rows, rowlevels):
        if lev > level:
            # parent is not shown
            continue
        level = lev

        r = ReportRow.get(rowid) or {}
        hidden = r.get('Hidden')
        try:
            t = r['TableName']
            id = r['TableID']
            deleted = (Database[t][id].get('zzStatus') == 'deleted')
        except KeyError:
            deleted = True

        if not (hidden or deleted):
            self.rows.append(rowid)
            self.rowlevels.append(lev)
            level += 1

def GetColumnList(reportid):
    """ Return list of column id's in the current order """
    ids = []
    done = {}
    report = Report.get(reportid) or {}
    k = report.get('FirstColumn')
    while k in ReportColumn and k not in done:
        ids.append(k)
        done[k] = None
        k = ReportColumn[k].get('NextColumn')
    return ids

def ReorderReportColumns(reportid, columnids):
    """ Use the list of column ids as the columns sequence of the report

    Ignores duplicates.  Calling routine must call 'SetUndo'.

    """
    if debug: print "Start ReorderReportColumns"

    # remove duplicates
    stack = []
    done = {}
    for id in columnids:
        if (id not in done) and (id in ReportColumn):
            stack.append(id)
            done[id] = None
    stack.reverse()

    # apply new order
    next = None
    for id in stack:
        if ReportColumn[id].get('NextColumn') != next:
            change = {'Table': 'ReportColumn', 'ID': id, 'NextColumn': next}
            Update(change)
        next = id
    if Report[reportid].get('FirstColumn') != next:
        change = {'Table': 'Report', 'ID': reportid, 'FirstColumn': next}
        Update(change)

    if debug: print "End ReorderReportColumns"

def GetRowList(reportid):
    """ Returns a list of row ids in the current order """
    ids = []
    done = {}
    report = Report.get(reportid) or {}
    k = report.get('FirstRow')
    while k in ReportRow and k not in done:
        ids.append(k)
        done[k] = None
        k = ReportRow[k].get('NextRow')
    return ids

def GetRowLevels(reportid):
    """ Return two lists: row ids and row levels (ie - ancestor counts) """
    rlist = GetRowList(reportid)
    rlevels = []

    stack = []
    for rid in rlist:
        rr = ReportRow.get(rid) or {}
        parent = rr.get('ParentRow')
        while stack:
            if stack[-1] == parent:
                break
            stack.pop()
        rlevels.append(len(stack))
        stack.append(rid)

    return rlist, rlevels

def ReorderReportRows(reportid, rowids):
    """ Use the list of rowids as the row sequence of the report

    Ignore duplicates.  Keep children with parents.
    Calling routine must call 'SetUndo'.

    """
    if debug: print "Start ReorderReportRows"

    # remove duplicates
    queue = []
    done = {}
    for id in rowids:
        if (id not in done) and (id in ReportRow):
            queue.append(id)
            done[id] = None

    # examine hierarchy
    godfamily = {}
    family = {}  # {parent: [child, ...]}
        # godfamily is cross-table; family is same-table

    stack = []
    for id in queue:
        record = ReportRow[id]
        parent = record.get('ParentRow')
        if parent in ReportRow:
            if parent not in family:
                godfamily[parent] = []
                family[parent] = []
            prec = ReportRow[parent]
            if prec.get('TableName') == record.get('TableName'):
                family[parent].append(id)
            else:
                godfamily[parent].append(id)
        else:
            stack.append(id)

    # apply new order
    next = None
    while stack:
        id = stack[-1]
        if id in family:
            stack += godfamily.pop(id) + family.pop(id)
        else:
            stack.pop()
            if ReportRow[id].get('NextRow') != next:
                change = {'Table': 'ReportRow', 'ID': id, 'NextRow': next}
                Update(change)
            next = id
    if Report[reportid].get('FirstRow') != next:
        change = {'Table': 'Report', 'ID': reportid, 'FirstRow': next}
        Update(change)

    if debug: print "End ReorderReportRows"

def AdjustReportRows():
    """ Ensure that every report has the correct rows and hierarchy
    
    The primary steps:
    - build a list of all of the records that should appear in the report
    - scan the current rows
        - remove every row whose record isn't in the Should list
        - remember the ids for every row whose record is in the Should list
    - create a row for every record in the Should list that lacks a row
    - remove parenting loops
    - give each a row a reference to its parent row
    - call ReorderReportRows to link the rows

    """
    if debug: print "Start AdjustReportRows"
    newrow = {'Table': 'ReportRow'}
    oldrow = {'Table': 'ReportRow'}

    # process all non-deleted reports
    for rk, r in Report.iteritems():
        if r.get('zzStatus') == 'deleted' or r.get('AdjustRowOption'): continue
        newrow['ReportID'] = rk

        rtid = r.get('ReportTypeID')
        if rtid not in ReportType:
            if debug: print "invalid ReportType id: report", rk, ", type", rtid
            continue
        rt = ReportType[rtid]

        ta, tb = rt.get('TableA'), rt.get('TableB')
        tableA = Database.get(ta)
        if not tableA: continue

        if tb == ta:
            tableB == None
        else:
            tableB = Database.get(tb)

        # search for records that should be present
        should = {}  # {table name: {record id: row id}}

        map = {}
        selcol = r.get('SelectColumn')
        if selcol:
            selval = r.get('SelectValue')
            for id, record in tableA.iteritems():
                if record.get(selcol) == selval:
                    map[id] = None
        else:
            for id, record in tableA.iteritems():
                map[id] = None
        should[ta] = map

        if tableB:
            map = {}
            for id, record in tableB.iteritems():
                parent = record.get(ta + 'ID')
                if parent in should[ta]:
                    map[id] = None
            should[tb] = map

        # correlate records with existing rows
        rlist = []
        for rowid in GetRowList(rk):
            r = ReportRow[rowid]
            t = r.get('TableName')
            tid = r.get('TableID')

            if t in should and tid in should[t] and not should[t][tid]:
                should[t][tid] = rowid
                rlist.append(rowid)

        # create a row for every record that needs one;
        #   map parent records to child rows
        parents = {}  # {row id: (parent table, parent id)}
        godparents = {}
            # here, parents are same-table; godparents are cross-table

        for t, map in should.iteritems():
            newrow['TableName'] = t
            for tid, rowid in map.items():
                if not rowid:
                    newrow['TableID'] = tid
                    rowid = Update(newrow)['ID']
                    map[tid] = rowid
                    rlist.append(rowid)

                record = Database[t][tid]
                parent = record.get(t + 'ID')
                prec = Database[t].get(parent)

                if prec and prec.get('zzStatus') != 'deleted':
                    parents[rowid] = t, parent
                if t != ta:
                    godparents[rowid] = ta, record.get(ta + 'ID')

        # map parent rows to child rows
        parentrows = {}  # {child row: parent row}

        for rowid in rlist:
            lineage = {}
            while rowid in parents:
                t, tid = parents.pop(rowid)
                parent = should[t].get(tid)
                lineage[rowid] = parent
                if parent in lineage:
                    del lineage[parent]
                rowid = parent
            parentrows.update(lineage)

        # give each a row a reference to its parent row
        for rowid in rlist:
            parent = parentrows.get(rowid)
            if (not parent) and (rowid in godparents):
                t, tid = godparents[rowid]
                parent = should[t].get(tid)
            if ReportRow[rowid].get('ParentRow') != parent:
                oldrow['ID'] = rowid
                oldrow['ParentRow'] = parent
                Update(oldrow)

        # link the new rowlist
        ReorderReportRows(rk, rlist)

    if debug: print "End AdjustReportRows"

########## @@@@@@@@@@ End Alex Row/Column Order @@@@@@@@@@ ##########

# --------- Routines to load and save data ---------
OptionFile = None

def LoadOption(directory=None):
    """ Load the option file.
    """
    global OptionFile, Option
    if not directory:
        directory = Path

    OptionFile = os.path.join(directory, "Options.ganttpvo")
    if debug: print "load option file:", OptionFile
    try: 
        f = open(OptionFile, "rb")
    except IOError:
        if debug: print "option file not found"
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
        Menu.UpdateWindowMenuItem(id)
        OpenReports[id].Show(True)
    OpenReports[id].Raise()

def CloseReport(id):
    """ Close a report. """
    if OpenReports.has_key(id):
        if Report.get(id) and Report[id].has_key('Open'):
            del Report[id]['Open']
        OpenReports[id].Destroy()
        del OpenReports[id]
        Menu.UpdateWindowMenuItem(id)

def CloseReports():
    """ Close all reports except #1. """
    for id, frame in OpenReports.items():
        if id == 1: continue
        if Report.get(id) and Report[id].has_key('Open'):
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
    OtherData =     Database['OtherData']
    Other =         OtherData[1]
    NextID =        Database['NextID']

    Database['Other'] = Database['OtherData']  # 'OtherData' is deprecated

    if not Other.get('FileSignature'):
        Other['FileSignature'] = random.randint(1, 1000000000)

    SetupDateConv()
    GanttCalculation()
    AdjustReportRows()
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
    UndoStack = []
    RedoStack = []
    ChangedData = False  # true if database needs to be saved
    ChangedCalendar = ChangedSchedule = ChangedReport = ChangedRow = False

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
    name_dict['debug'] = 1

    try:
        execfile(path, name_dict)
    except:
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

# SetEmptyData()
# MakeReady()

if debug: print "end Data.py"
