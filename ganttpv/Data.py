#!/usr/bin/env python
# Data Tables - includes update logic, date routines, and gantt calculations

# Copyright 2004 by Brian C. Christensen

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

# import calendar
import datetime
import cPickle
import wx, os
import Menu 
import GanttReport
# import Main

debug = 0
if debug: print "load Data.py"
# On making changes in GanttPV
# 1- Use Update for all changes
# 2- Make change will decide the impact of the changes
# 3- Use SetUndo to tell GanttPV to fix anything in response to the changes

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

Successor = { }  # dependency xref (used?)
Predecessor = { }

# Date conversion tables
DateConv = {}   # usage: index = DateConv['2004-01-01']
DateIndex = []  # usage: date = DateIndex[1]
DateInfo = []   # dayhours, cumhours, dow = DateInfo[1]

# save impact of change until it can be addressed
ChangedCalendar = False  # will affect calendar hours tables (gantt will have to be redone, too)
ChangedSchedule = False  # gantt calculation must be redone
ChangedReport = False    # report layout needs to be redone
ChangedRow = False       # a record was added or deleted, the rows on a report may be affected

UndoStack = []
RedoStack = []
OpenReports = {}  # keys are report id's of open reports. Main is always "1"; Resource is always "2"

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

def FillTable(name, t, Columns, Data):
    """ Utility routine to load sample (or initial) data into database """
    for i, d in enumerate(Data):
        id = i+1
        t[id] = {}
        t[id]['ID'] = id

        for c, r in zip( Columns, d ):
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
                ( 14, 4, 0, None,       '',                     '',             'CHART',        '','',          'Day', 21, '2004-04-13' ),
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

# --------------------- ( is this used for anything? )
def SetDependencyXref():
    for i, d in Dependency.iteritems():
        p = d['PrerequisiteID']
        s = d['TaskID']
        if not Predecessor.has_key(s):
            Predecessor[s] = []
        if not Successor.has_key(p):
            Successor[p] = []
        Predecessor[s].append(p)
        Successor[p].append(s)

# ----------------------

# check for important changes
def CheckChange(change):  # change contains the undo info for the changes (only changed columns included)
    """ Check for important changes. """
    global ChangedCalendar, ChangedSchedule, ChangedReport, ChangedRow
    if debug: print "Start CheckChange"
    if debug: print 'change', change
    if change.has_key('zzStatus') and not change['Table'] in ('ReportRow', 'ReportColumn'): ChangedRow = True  # something has been added or deleted
        # Don't really need 'zzStatus' for new record adds if I set the new record add flag in Update???
        # zzStatus on all new records just take up space.
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
        for k in ('StartDate', 'DurationHours', 'zzStatus', 'ProjectID'):
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
    for k, v in OpenReports.items():  # invalid reports might be closed during this loop
        if debug: print 'reportid', k
        if v == None: continue
        if k != 1: v.SetReportTitle()  # only grid reports
        Menu.AdjustMenus(v)
        v.Refresh()
        #v.Report.RefreshReport()  # update displayed data
        v.Report.Refresh()  # update displayed data (needed for Main on Windows, not needed on Mac)
    if debug: print "End RefreshReports"

# set up undo
def SetUndo(message):
    """ This is the last step in submitting a group of changes to the database. 
            It tells the system to adjust any system calculations and to update the displays. """
    global ChangedCalendar, ChangedSchedule, ChangedReport, ChangedRow, ChangedData
    global UndoStack, RedoStack  # is this needed for the UndoStack??

    if debug: print "Start SetUndo"
    if debug: print "message", message
    # these routines may add more information to the undo stack
    UpdateCalendar = ChangedCalendar and Option.get('AutoGantt')
    UpdateGantt = (ChangedCalendar or ChangedSchedule) and Option.get('AutoGantt')
    UpdateReports = UpdateCalendar or UpdateGantt or ChangedReport or ChangedRow

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
    UndoStack.append(message)
    ChangedData = True  # file needs to be saved
    RedoStack = []  # clear out the redo stack
    # adjust visibility/appearance of user interface objects
    # refresh all reports??
    RefreshReports()
    if debug: print "End SetUndo"

# ----- undo and redo

def Recalculate():
    global ChangedCalendar, ChangedSchedule, ChangedRow, ChangedReport

    # these routines shouldn't add anything to the undo stack
    UpdateReports = ChangedCalendar or ChangedSchedule or ChangedReport or ChangedRow
    if ChangedCalendar: SetupDateConv()
    if ChangedSchedule or ChangedCalendar: GanttCalculation(); ChangedSchedule = False; ChangedCalendar = False
    if ChangedRow: AdjustReportRows(); ChangedRow = False  # needed? I think so
    if UpdateReports:
        ChangedReport = False
        for k, v in OpenReports.items():  # an invalid report may be closed in this process
            if v == None: continue
            v.UpdatePointers()
    RefreshReports()

def _Do(fromstack, tostack):
    global ChangedData
    if len(fromstack) > 1:
        savemessage = fromstack.pop()
        while len(fromstack) > 0 and isinstance(fromstack[-1], dict):
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
    if debug: print 'Start Update'
    if debug: print 'change', change
    tname = change['Table']  # exception if not found
    undo = { 'Table' : tname }
    table = Database[tname]
    if change.get('ID') != None:  # change existing record
        if debug: print "Changed existing record"
        id = change['ID']
        record = table[id]
        for c, newval in change.iteritems():  # process each new field
            if c == 'Table' or c == 'ID': continue
            if record.has_key(c):  # change existing value
                if newval != record[c]: # if value has changed
                    undo[c] = record[c]  # old value
                    record[c] = newval
            else:  # new value
                undo[c] = None
                record[c] = newval
    else:  # add a new record
        if debug: print "Added new record"
        record = {}
        for c, newval in change.iteritems():  # process each new field
            if c == 'Table': continue
            undo[c] = None
            record[c] = newval
        id = NextID[tname]
        if debug: print "new id:", id
        NextID[tname] = id + 1
        record['ID'] = id
        table[id] = record
        undo['zzStatus'] = 'deleted'
        record['zzStatus'] = 'active'  # not really needed if I set new record flag here??????
    undo['ID'] = id
    # redo['ID'] = id
    # check 'redo' to see what needs to be refreshed in the displays
    CheckChange(undo)
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

def SetupDateConv():
    """" Setup tables that will be used for all schedule date calculations """
    global DateConv, DateIndex, DateInfo
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

    DateConv = {}  # usage: index = DateConv['2004-01-01']
    DateIndex = []  # usage: date = DateIndex[1]
    DateInfo = [] # dayhours, cumhours, dow = DateInfo[1]

    FirstDate = '9999-99-99'
    LastDate = '0000-00-00'
    oneday = datetime.timedelta(1)
    oneweek = oneday * 7

    # find the earliest and last dates that have been specified anywhere
    for p in Project.values():
        d = p.get('StartDate', '9999-99-99')
        if d and d != '' and d < FirstDate: FirstDate = d
        d = p.get('TargetEndDate', '0000-00-00')
        if d and d != '' and d > LastDate: LastDate = d

    for p in Task.values():
        d = p.get('StartDate', '9999-99-99')
        if d and d != '' and d < FirstDate: FirstDate = d
        d = p.get('EndDate', '0000-00-00')
        if d and d != '' and d > LastDate: LastDate = d

    if FirstDate == "9999-99-99": d1 = datetime.date.today()
    else: d1 = StringToDate(FirstDate)
    d1 = makeFOW(d1) - (oneweek * 50)  # allow 50 weeks of room before the first date in the file

    if LastDate == "0000-00-00": d2 = d1 + (oneweek * 105)
    else: d2 = makeFOW(StringToDate(LastDate)) + (oneweek * 105)

    if debug: print "first & last", d1.strftime("%Y-%m-%d"), d2.strftime("%Y-%m-%d")

    wh = Other.get('WeekHours', (8,8,8,8,8,0,0))
    hxref = {}
    for k, v in Holiday.iteritems():
        if v.get('zzStatus', 'active') == 'deleted': continue
        date = v.get('Date')
        hours = v.get('Hours')
        if date and hours != None:
            hxref[date] = hours
    cumhours = 0
    i = 0
    while d1 <= d2:
        date = d1.strftime("%Y-%m-%d")
        dow = i % 7
        dh = hxref.get(date, wh[dow])  # use holiday hours if provided, else use week hours
        # if debug: print "date", date, dh
        DateConv[date] = i
        DateIndex.append(date)
        DateInfo.append( (dh, cumhours, dow), )  # work hours on date, work hours prior to date, day of week
        cumhours += dh; i += 1
        d1 += oneday

    Other['BaseDate'] = DateIndex[0]  # The base date is a reflexion of the date conversion tables that are
                                      # in place. It must always reflect the current tables. It is not 'undo-able'

# -----------------
def GetToday():
    dToday = datetime.date.today()      # get date object for today
    return dToday.strftime("%Y-%m-%d")  # convert to standard database format

def GanttCalculation(): # Gantt chart calculations - all dates are in hours
    # change = { 'Table': 'Task' }  # will be used to apply updates  --->> Don't Use 'Update' here <<--
    # Set the project start date (use specified date or today)  later may be adjusted by the tasks' start dates
    Today = GetToday()
    if debug: print "today", Today
    ps = {}  # project start dates
    for k, v in Project.iteritems():
        if v.get('zzStatus', 'active') == 'deleted': continue
        sd = v.get('StartDate')
        # if debug: print "project, startdate", k, sd
        if sd == "": sd = Today
        elif sd == None: sd = Today  # default project start dates to today
        ps[k] = sd
        #ProjectStart = Project[projectid].get('StartDate', Today.strftime("%Y-%m-%d"))
        # if debug: print "project, startdate (adjusted)", k, sd

    # dependencies
    pre = {}  # task prerequisites indexed by task id number
    suc = {}  # task successors
    precnt = {}  # count of all unprocessed prerequisites
    tpid = {}  # task's projectid
    for k, v in Task.iteritems():  # init dependency counts, xrefs, and start dates
        if v.get('zzStatus', 'active') == 'deleted': continue
        pid = v['ProjectID']
        if not pid: continue  # silently ignore tasks w/o projects
        precnt[k] = 0
        pre[k] = []
        suc[k] = []
        tpid[k] = pid
        if debug: "task's project", k, pid
        tsd = v.get('StartDate')
        if tsd == "": tsd = None
        if tsd and tsd < ps[pid]: ps[pid] = tsd  # adjust project start date if task starts are earlier
    for k, v in Dependency.iteritems():
        if debug: print "dependency record", v
        if v.get('zzStatus', 'active') == 'deleted': continue
        p = v['PrerequisiteID']
        t = v['TaskID']
        if suc.has_key(p) and pre.has_key(t):  # silently ignore dependencies for tasks not in list
            precnt[t] += 1
            pre[t].append(p)
            suc[p].append(t)

    # convert project start dates to hours format
    ProjectStartHour = {}
    for k, v in ps.iteritems():
        si = DateConv[v]  # get starting date index
        sh = DateInfo[si][1]  # get cum hours for start date
        ProjectStartHour[k] = sh
        if debug: "project, start hour", k, sh

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
                if tsd:
                    tsi = DateConv[tsd]  # date index
                    tsh = DateInfo[tsi][1]  # date hour
                    if tsh > es: es = tsh  # use the date if dependencies allow

                # calculate early finish
                if tsd and ted and not td:  # use difference in dates to compute duration
                    tei = DateConv[ted]  # date index
                    teh = DateInfo[tei+1][1]  # first hour of next day
                    ef = es + (tsh - teh)
                elif td:
                    ef = es + td
                else:
                    ef = es + 8

                # update database
                Task[k]['hES'] = es
                Task[k]['hEF'] = ef
                Task[k]['CalculatedStartDate'], Task[k]['CalculatedStartDateHour'] = HoursToDate(es)
                Task[k]['CalculatedEndDate'], Task[k]['CalculatedEndDateHour'] = HoursToDate(ef)

                # change['ID'] = k
                # change['hES'] = es
                # change['hEF'] = ef
                # change['CalculatedStartDate'], change['CalculatedStartDateHour'] = HoursToDate(es)
                # change['CalculatedEndDate'], change['CalculatedEndDateHour'] = HoursToDate(ef)
                # Update(change, 0)

                # tell successor that I'm ready
                for t in suc[k]: precnt[t] -= 1
                precnt[k] = -1
                                        
# end of GanttCalculation

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

def AdjustReportRows(): # Reports rows may be affected by table adds or deletions
#                               this routine makes sure that every report has the right number of rows with 
#                               the right links to rows in other tables
# 1- build a list of all of the data that should appear in the report
# 2- remove from that list everything that currently appears in the report
# 3- add additional rows for everything that doesn't
# question: should I have a flag in report rows for deleted records?
# answer: no, let the individual reports check. they have to scan through all the rows anyway for hidden rows
#
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
        if r.get('zzStatus', 'active') == 'delete': continue
        newrow["ReportID"] = rk  # make sure any new rows know their parent report is
        rtid = r.get('ReportTypeID')
        if not rtid:  # could happen if report is "undone"
            if debug: print "AdjustReportRows: invalid ReportTypeID (rk, rtid)", rk, rtid
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
            while rowk:
                saverowk = rowk
                rr = ReportRow[rowk]
                t = rr['TableName']; id = rr['TableID']
                if t == ta:
                    shoulda.remove(id)  # remove from list all that have rows
                    iref[id] = rowk  # insertion at beginning (may be overridden in elif)
                    tak = id # most recent table 'a' id, used for insertion at end
                elif t == tb:
                    shouldb.remove(id)  # remove from list all that have rows
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
                tak = tableb[kb].get(ta + "ID")  # table 'a' id record of 'b's parent
                if not tak: continue  # if the key is NULL or 0
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

            if debug: print "shoulda", shoulda
            # compare that with everything that is already there
            rowk = r.get('FirstRow', 0)
            loopguard = 0
            saverowk = 0  # this will stay 0 if no report rows
#            while rowk != 0:  # ditto debug change 040717
            while rowk:
                saverowk = rowk
                rr = ReportRow[rowk]
                id = rr['TableID']
                shoulda.remove(id)  # remove from list all that have rows
                rowk = rr.get('NextRow', 0)
                loopguard += 1
                if loopguard > 10000: break
            # saverowk points to the last row in the report

            # add report rows for everything that is missing
            newrow['TableName'] = ta; newrow['NextRow'] = 0
            for ka in shoulda:
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
        Menu.GetScriptNames()  # must be done after option file is loaded

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

def OpenReport(id):  # this should not be called for report #1
    r = Database['Report'][id]
    if not r.get('ReportTypeID'): return  # don't open if not report type id (can happen if "undone")
    if r.get('Open') and OpenReports.get(id):
        OpenReports[id].Raise()  # allow attempt to open a report that is already open
        return
    type = r.get('Type', 'GanttReport')  # at present Type isn't defined in any report record
    # if id == 1:  frame = OpenReports[1]
    if type == 'GanttReport': frame = GanttReport.GanttReportFrame(id, None, -1, "")
    else: return
    OpenReports[id] = frame
    r['Open'] = True
    if r.get('FramePositionX'):
        frame.SetPosition(wx.Point(r['FramePositionX'], r['FramePositionY']))
        frame.SetSize(wx.Size(r['FrameSizeW'], r['FrameSizeH']))
    frame.SetReportTitle()
    frame.Show(True)

def CloseReport(id):
    if Database['Report'][id].get('Open'):
        v = OpenReports[id]
        del OpenReports[id]
        v.Destroy()
        del Database['Report'][id]['Open']  # delete is the same as setting it to false, but saves file space

def CloseReports():  # close all reports except #1
    klist = OpenReports.keys()
    for k in klist:
        if k != 1: CloseReport(k)

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
        elif Database['Report'][k].get('Open', False):  OpenReport(k)
    UndoStack  = []
    RedoStack = []
    ChangedData = False  # true if database needs to be saved
    ChangedCalendar, ChangedSchedule, ChangedReport, ChangedRow = False, False, False, False

#    RefreshReports()  # needed?? -- not here, this routine is back end

def LoadContents(self):
    """ Load the contents of our document into memory.
    """
    global Database

    f = open(FileName, "rb")
    header = f.readline()  # add read line of text - will allow conversion of earlier versions or use of different file formats
    Database = cPickle.load(f)
    f.close()

    MakeReady()

def SaveContents(self):
    """ Save the contents of our document to disk.
    """
    global ChangedData

    f = open(FileName, "wb")
    # add write line of text 
    f.write("GanttPV\t0.1\ta\n")
    cPickle.dump(Database, f)
    f.close()

    ChangedData = False

def AskIfUserWantsToSave(self, action):
    """ Give the user the opportunity to save the current document.
        'action' is a string describing the action about to be taken.  If
        the user wants to save the document, it is saved immediately.  If
        the user cancels, we return False.
    """
    global FileName
    if not ChangedData: return True # Nothing to do.

    response = wx.MessageBox("Save changes before " + action + "?",
                                "Confirm", wx.YES_NO | wx.CANCEL, self)

    if response == wx.YES:
        if FileName == None:
            tempFileName = wx.FileSelector("Save File As", "Saving",
                                            default_filename='Untitled.ganttpv',
                                            default_extension="ganttpv",
                                            wildcard="*.ganttpv",
                                            flags = wx.SAVE | wx.OVERWRITE_PROMPT)
            if tempFileName == "": return False # User cancelled.
            FileName = tempFileName

        SaveContents(self)
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
