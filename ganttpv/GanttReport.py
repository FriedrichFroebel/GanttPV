#!/usr/bin/env python
# gantt report definition

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

# 040408 - first version of this program
# 040409 - will display task values in report
# 040410 - SetValue will now set Task fields
# 040419 - changes for new ReportRow column names
# 040420 - changes to allow Main.py to open gantt reports; doClose will close this report; saves Report size and position
# 040424 - added doInsertTask, doDuplicateTask, doDeleteTask
# 040426 - added doMoveRow, added doPrerequisite
# 040427 - added doAssignResource, onHide, onShowHidden; fixed doDelete; save the result when the user changes columns widths; prevent changes to row height; changed some column names in ReportColumn
# 040503 - changes to use new ReportType and ColumnType tables; added OnInsertColumn; added row highlighting for two table reports
# 040504 - finished OnInsertColumn; added OnDeleteColumn, and OnMoveColumn
# 040505 - support UI changes (many references to 'Task' replaced with 'Row', format changes, and minor adjustments); revised OnInsertRow, OnDuplicateRow, and OnDeleteRow to work with all tables; added OnScroll and OnScrollToTask; use colors from Data.Option for chart
# 040508 - prevent deletion of project 1 or reports 1 or 2
# 040510 - set minimum row size for non-gantt column in "_updateColAttrs"; set Hidden and Deleted row color; copied Scripts menu processing from Main.py; changed doClose to use Data.doClose
# 040512 - added OnEditorShown; added support for indirect column types; several bug fixes
# 040518 - added doHome
# 040528 - in doDuplicate only dup rows of primary table's type; in onPrerequisite only include Tasks in the list of potential prerequisites
# 040531 - in onDraw make sure all rectangles use new syntax for version 2.5
# 040715 - Pierre_Rouleau@impathnetworks.com: removed all tabs, now use 4-space indentation level to comply with Official Python Guideline.
# 040906 - changed OnInsertColumn to ignore Labels w/ value of None; display "project name / report name" in report title.
# 040914 - handle indirect display if no ID is found; assign project id when inserting rows into reports that can be each.
# 040915 - allow entry of floating point numbers (type "f")
# 040918 - add week time scale; default FirstDate in column type to today or this week
# 040928 - Alexander - ignores dates not present in Data.DateConv; prevents entry of incorrectly formatted dates
# 041001 - display measurements in weekly time scale
# 041009 - Alexander & Brian - default dates to current year and month; Brian - change SetValue to work on measurement time scale data
# 041009 - change scroll to work w/ any time scale column
# 041010 - change "column insert" to set # periods for all timescale columns; changes to allow edit of non-measurement   time scale columns
# 041126 - draw bars for week timescale
# 041203 - moved get column header logic to Data
# 041204 - default time scale first date to today
# 050101 - previously added month  & quarter timescale
# 050101 - added entry of duration units datatype
# 050105 - fix part of bug: avoid report reset unless # rows or # columns changes (report un-scrolls when reset this was a problem when entering start dates and durations); problem still exists when adding a column or a row, but at least it will be less annoying.
# 050106 - fixed bug where deleted records threw off location of inserted rows
# 050202 - remove line feed from insert column menu text
# 050423 - move logic to calculate period start and hours to Data.GetPeriodInfo
# 050424 - fix ScrollToTask for M & Q
# 050426 - Alexander - fixed bug where moving columns threw off displayed widths
# 050503 - make wxPython 2.6.0.0 use same fonts as prior default
# 050504 - in OnPrerequisite, chain tasks if more than one non-parent task is selected
# 050504 - Alexander - implemented Window menu; moved some menu event-handling logic to Menu.py
# 050505 - Alexander - updated time units feature to use the work week
# 050513 - Alexander - replaced UpdateColumnWidths with new UpdateColumns; this fixes the bug where moving columns threw off cell renderers and read-only status
# 050517 - Alexander - fixed bug where row size was not properly adjusted for the presence of a gantt column
# 050520 - Alexander - renamed UpdateColumns into UpdateAttrs, and added a call to _updateRowAttrs; this fixes the bug where moving rows threw off row colors.
# 050527 - Alexander - added IndentedRenderer
# 050630 - Brian - use PlanBarColor to override default bar color
# 050826 - Alexander - rewrote row/column movement!
# 050903 - Brian - prevent negative window positions (fix problem caused when Windows iconizes reports)
# 051207 - Brian - changed GetSelectedRows to work around error in wx that treats shift-click rows as cells not rows
# 060131 - Alexander - store gantt-chart backgrounds in the column attributes (allows print script to override those colors)

import wx, wx.grid
import datetime
from wxPython.lib.dialogs import wxMultipleChoiceDialog
import Data, UI, ID, Menu
# import images
import re

debug = 1
is24 = 1

if debug: print "load GanttReport.py"

# ------------ Table behind grid ---------

class GanttChartTable(wx.grid.PyGridTableBase):
    """ A custom wxGrid Table using user supplied data """
    def __init__(self, reportid):
        """ data is taken from SampleData """
        # The base class must be initialized *first*
        wx.grid.PyGridTableBase.__init__(self)

        # create local pointers to SampleData?
        self.UpdateDataPointers(reportid)
        self.UpdateColumnPointers()
        self.UpdateRowPointers()

        # self.colnames = _getColNames()

        # store the row length and col length to test if the table has changed size
        self._rows = self.GetNumberRows()
        self._cols = self.GetNumberCols()

    def GetNumberCols(self):
        return len(self.columns)

    def GetNumberRows(self):
        return len(self.rows)

    def GetColLabelValue(self, col):
        return Data.GetColumnHeader(self.columns[col], self.coloffset[col])

    def GetValue(self, row, col):
        return Data.GetCellValue(self.rows[row], self.columns[col], self.coloffset[col])

    def GetRawValue(self, row, col):  # same as GetValue  ( I don't know the difference. The example I'm following made them the same. )

        # Alexander - In the wxPython demo, GetValue converted to a string,
        #   while GetRawValue did not. But GetRawValue is not a standard
        #   function of the base class; we don't need to add it.

        value = GetValue(self, row, col)
        return value

    def SetValue(self, row, col, value):
        rr = self.reportrow[self.rows[row]]
        rtable = rr.get('TableName')
        tid = rr['TableID']

        # rc = self.reportcolumn[self.columns[col]]
        ct = self.columntype[self.ctypes[col]]

        # t = ct.get('T', 'X')
        # rtid = self.report.get('ReportTypeID')
        # ctable = Data.ReportType[rtid].get('Table' + t)
        # if rtable != ctable: return  # shouldn't happen

        type = ct.get('DataType', 't')

        of = self.coloffset[col]
        if of != -1:
            ctperiod, ctfield = ct.get('Name').split("/")
            if ctfield == "Gantt":  # shouldn't happen
                return
            else:  # table name, field name, time period, and record id
                if debug: print 'ctfield', ctfield
                if ctfield == 'Measurement':  # find field name
                    mid = Data.Database[rtable][tid].get('MeasurementID')  # point at measurement record
                    if mid: # measurement id
                        fieldname = Data.Database['Measurement'][mid].get('Name')  # measurement name == field name
                        type = Data.Database['Measurement'][mid].get('DataType')  # override column type w/ measurement type
                    else:
                        fieldname = None
                    tid = Data.Database[rtable][tid].get('ProjectID')  # point at measurement record
                    rtable = 'Project'  # only supports project measurements
                else:
                    fieldname = ctfield
            if debug: print 'offset, rtable, fieldname', of, rtable, fieldname
            if fieldname == None: return

            # find the period date
            firstdate = self.reportcolumn[self.columns[col]].get('FirstDate')
            if not Data.DateConv.has_key(firstdate): firstdate = Data.GetToday()  # use standard default
            index = Data.DateConv[ firstdate ]
            if ctperiod == "Day":
                date = Data.DateIndex[ index + of ]
            elif ctperiod == "Week":
                index -= Data.DateInfo[ index ][2]  # convert to beginning of week
                date = Data.DateIndex[ index + (of * 7) ]
            else:
                return  # don't update anything if time period not recognized
            timename = rtable + ctperiod

            timeid = Data.FindID(timename, rtable + "ID", tid, 'Period', date)

        # add processing of other date formats (per user preferences)
        # also add processing of dd, mm-dd, and yy-mm-dd for standard format
        if value == "": v = None
        elif type == 'i': 
            try:
                v = int(value)
            except ValueError:  # should I display an error message?
                return
        elif type == 'f': 
            try:
                v = float(value)
            except ValueError:  # should I display an error message?
                return
        elif type == 'd':
            v = Data.CheckDateString(value)
            if not v: return
        elif type == 'u': 
            result = re.match(r"^\s*(\d+w)?\s*(\d+d)?\s*(\d+h)?\s*$|^\s*(\d+)\s*$", value, re.I)
            if result:
                groups = result.groups() # [weeks, days, hours, None] or [None, None, None, hours]
                if groups[3]:
                    v = int(groups[3])
                else:                    
                    v = 0
                    if groups[2]: v += int(groups[2][:-1])
                    if groups[1]: v += int(groups[1][:-1]) * Data.HoursPerDay
                    if groups[0]: v += int(groups[0][:-1]) * Data.HoursPerWeek
                    if v % 1: v += 1
                    v = int(v)
            else: 
                 return
        else: v = value

        change = {}

        at = ct.get('AccessType')
        if at == 'd':
            column = ct.get('Name')
            table = rtable
            id = tid
            if not id: return  # don't make update if ID is invalid
            change['ID'] = id
        elif at == 'i':
            table, column = ct.get('Name').split('/')  # indirect table & column
            id = self.data[rtable][tid].get(table+'ID')
            if not id: return  # don't make update if ID is invalid
            change['ID'] = id
        elif at == 's':
            table = timename
            column = fieldname
            if timeid: # don't add record if already exists
                change['ID'] = timeid
            else:
                change[rtable + 'ID'] = tid
                change['Period'] = date
        else:
            return  # we don't recognize how to find the record to update, so ignore it

        change['Table'] = table
        change[column] = v
        Data.Update(change)
        Data.SetUndo(column)

    def ResetView(self, grid):  # deprecated
        grid.Reset()

    def UpdateValues(self, grid):  # deprecated
        # This sends an event to the grid table to update all of the values
        msg = wx.grid.GridTableMessage(self, wx.grid.GRIDTABLE_REQUEST_VIEW_GET_VALUES)
        grid.ProcessTableMessage(msg)

    # -- These are not part of the sample interface - my routines to make display easier

    def UpdateDataPointers(self, reportid):
        Data.UpdateDataPointers(self, reportid)

        # self.data = Data.Database
        # self.report = self.data["Report"][reportid]
        # self.reportcolumn = self.data["ReportColumn"]
        # self.reportrow = self.data["ReportRow"]

    def UpdateColumnPointers(self):
        rc = self.reportcolumn  # pointer to table
        rct = self.columntype  # pointer to table

        self.columns = []
        self.ctypes = []
        self.coloffset = []

        reportid = self.report.get('ID')
        for colid in Data.GetColumnList(reportid):
            r = rc.get(colid) or {}
            ctid = r.get('ColumnTypeID')
            ct = rct.get(ctid) or {}
            type = ct.get('AccessType')
            if type == 's':
                # column is a time scale
                for of in xrange(r.get('Periods') or 1):
                    self.columns.append( colid )
                    self.ctypes.append( ctid )
                    self.coloffset.append( of )
            else:
                self.columns.append( colid )
                self.ctypes.append( ctid )
                self.coloffset.append( -1 )

    def UpdateRowPointers(self):
        Data.UpdateRowPointers(self)
        
    def _updateColAttrs(self, grid):
        """
        wxGrid -> update the column attributes to add the
        appropriate renderer given the column type.

        Otherwise default to the default renderer.
        """
        rsize = 22
        for col in xrange(len(self.columns)):
            attr = wx.grid.GridCellAttr()
            cid = self.columns[col]
            rc = Data.ReportColumn.get(cid) or {}
            ctid = self.ctypes[col]
            ct = self.columntype.get(ctid) or {}
            ctname = ct.get('Name')
            of = self.coloffset[col]
            if of > -1 and ctname and ctname[-6:] == "/Gantt":
                renderer = GanttCellRenderer(self)
                grid.SetColSize(col, renderer.colSize)
                if rsize < renderer.rowSize:
                    rsize = renderer.rowSize
                attr.SetReadOnly(True)
                attr.SetRenderer(renderer)
                cid = self.columns[col]
                rc = Data.ReportColumn.get(cid) or {}
                fdate = rc.get('FirstDate')
                if fdate and Data.ValidDate(fdate):
                    ix = Data.StringToDate(fdate)
                else:
                    ix = Data.TodayDate()
                if Data.GetPeriodInfo(ctname, ix, of)[0]:
                    bgcolor = Data.Option.get('WorkDay', wx.BLUE)
                else:
                    bgcolor = Data.Option.get('NotWorkDay', wx.BLUE)
                attr.SetBackgroundColour(bgcolor)
            else:
                if ctname == 'Name':
                    sub_renderer = grid.GetDefaultRenderer()
                    renderer = IndentedRenderer(self, sub_renderer)
                    attr.SetRenderer(renderer)
                csize = rc.get('Width') or grid.GetDefaultColSize()
                grid.SetColSize(col, csize)
                readonly = not ct.get('Edit')
                attr.SetReadOnly(readonly)
            grid.SetColAttr(col, attr)
        if rsize != grid.GetDefaultRowSize():
            grid.SetDefaultRowSize(rsize, True)

    def _updateRowAttrs(self, grid):
        """ Highlight parent rows in split-table reports """
        rtid = self.report['ReportTypeID']
        rt = Data.Database['ReportType'].get(rtid)
        ta, tb = rt.get('TableA'), rt.get('TableB')

        for row, rowid in enumerate(self.rows):
            rr = self.reportrow.get(rowid) or {}
            t = rr.get('TableName')
            tid = rr.get('TableID')
            try:
                deleted = (Data.Database[t][tid]['zzStatus'] == 'deleted')
            except KeyError:
                deleted = False
            hidden = rr.get('Hidden')

            attr = wx.grid.GridCellAttr()
            attr.SetFont(wx.Font(11, wx.SWISS, wx.NORMAL, wx.NORMAL))

            if deleted:
                bgcolor = Data.Option.get('DeletedColor')
            elif hidden:
                bgcolor = Data.Option.get('HiddenColor')
            elif t == ta and tb:
                # parent table in two level report
                bgcolor = Data.Option.get('ParentColor')
            else:
                bgcolor = Data.Option.get('ChildColor')
                # attr.SetTextColour(wxBLACK)
                # attr.SetFont(wxFont(10, wxSWISS, wxNORMAL, wxBOLD))
                # attr.SetReadOnly(True)
                # self.SetRowAttr(row, attr)
            attr.SetBackgroundColour(bgcolor or "pale green")
            grid.SetRowAttr(row, attr)

# ---------------------- Draw Cells of Grid -----------------------------

class IndentedRenderer(wx.grid.PyGridCellRenderer):
    def __init__(self, table, renderer):
        wx.grid.PyGridCellRenderer.__init__(self)
        self.table = table
        self.renderer = renderer
        self.min_width = 20
        self.increment = 10

    def Draw(self, grid, attr, dc, rect, row, col, isSelected):
        x, y, width, height = rect
        max_indent = width - self.min_width
        if max_indent > 0:
             level = self.table.rowlevels[row]

             # subtract level of ancestor in primary table
             #   (because columns are not shared between tables)
             rt = self.table.reporttype
             ta = rt.get('TableA')
             rr = Data.ReportRow[self.table.rows[row]]
             t = rr.get('TableName')
             if t != ta:
                 parent = row
                 while t != ta and parent > 0:
                     parent -= 1
                     rr = Data.ReportRow[self.table.rows[parent]]
                     t = rr.get('TableName')
                 if parent >= 0:
                     level -= self.table.rowlevels[parent] + 1

             indent = level * self.increment
             if indent > 0:
                 if indent > max_indent:
                     indent = max_indent

                 # fill in the background
                 if isSelected:
                     color = wx.SystemSettings_GetColour(wx.SYS_COLOUR_HIGHLIGHT)
                 else:
                     color = attr.GetBackgroundColour()
                 dc.SetBrush(wx.Brush(color))
                 dc.SetPen(wx.Pen(color))
                 dc.DrawRectangle(x, y, indent, height)

                 # indent the drawing area
                 x += indent
                 width -= indent
                 rect = (x, y, width, height)

        self.renderer.Draw(grid, attr, dc, rect, row, col, isSelected)

class GanttCellRenderer(wx.grid.PyGridCellRenderer):
    def __init__(self, table):
        """
        Image Renderer Test.  This just places an image in a cell
        based on the row index.  There are N choices and the
        choice is made by  choice[row%N]
        """
        wx.grid.PyGridCellRenderer.__init__(self)
        self.table = table
        # self._choices = [images.getSmilesBitmap,
        #                  images.getMondrianBitmap,
        #                  images.get_10s_Bitmap,
        #                  images.get_01c_Bitmap]

        self.colSize = 24
        self.rowSize = 28

    def Draw(self, grid, attr, dc, rect, row, col, isSelected):
        # choice = self.table.GetRawValue(row, col)
        # bmp = self._choices[ choice % len(self._choices)]()
        # image = wxMemoryDC()
        # image.SelectObject(bmp)
        o = Data.Option  # colors

        # Get data to draw gantt chart
        # if debug: print 'col & colid', col, self.table.columns[col]
        rc = self.table.reportcolumn[self.table.columns[col]]
        # if debug: print 'rc', rc

        # draw bar chart or not

        # there are two tests. the name test is new, the FirstDate test is used in verion 0.1
        ct = self.table.columntype[self.table.ctypes[col]]  # if the column is in a report it will always have a type
        ctname = ct.get('Name')
        # if debug: print "draw ctname", ctname
        if ctname and ctname[-6:] != "/Gantt": return

        fdate = rc.get('FirstDate') or Data.GetToday()
        if not fdate or not Data.DateConv.has_key(fdate): return  # this routine shouldn't be called for not gantt columns, but it is anyway - just ignore them
                                # the program seems to be refreshing three times when one is needed (040505)

        # if debug: print "-- didn't return --"

        ix = Data.DateConv[fdate]
        of = self.table.coloffset[col]

        dh, cumh, dow = Data.GetPeriodInfo(ctname, ix, of)  # if period not recognized, defaults to day
        # if debug: print 'ix, dh, cumh', ix, dh, cumh

        # clear the background
        dc.SetBackgroundMode(wx.SOLID)
        if isSelected and dh:
            backcolor = o.get('WorkDaySelected', wx.BLUE)
        elif isSelected:
            backcolor = o.get('NotWorkDaySelected', wx.BLUE)
        else:
            backcolor = attr.GetBackgroundColour()

        dc.SetBrush(wx.Brush(backcolor, wx.SOLID))
        dc.SetPen(wx.Pen(backcolor, 1, wx.SOLID))
        if is24:
            dc.DrawRectangle(rect.x, rect.y, rect.width, rect.height)
        else:
            dc.DrawRectangle((rect.x, rect.y), (rect.width, rect.height))
        # draw gantt bar
        if dh > 0:  # only display bars on days that have working hours 

            # get info needed to draw bar
            rr = self.table.reportrow[self.table.rows[row]]
            tname = rr['TableName']
            if tname != 'Task': return  # only draw for task records
            tid = rr['TableID']
            task = self.table.data['Task'][tid]

            # pick color
            if isSelected:
                plancolor = o.get('PlanBarSelected', wx.GREEN)
                # actualcolor = o.get('ActualBarSelected', wx.GREEN)
                # basecolor = o.get('BaseBarSelected', wx.GREEN)
            else:
                plancolor = o.get('PlanBar', wx.GREEN)
                # actualcolor = o.get('ActualBar', wx.GREEN)
                # basecolor = o.get('BaseBar', wx.GREEN)

            override = rr.get('PlanBarColor')
            if override and re.match('^[0-9A-Fa-f]{6}$', override):
                plancolor = [ int(x,16) for x in (override[0:2], override[2:4], override[4:6]) ]

            # calculate bar location
            # (dh should be integer, but just to make sure I don't divide by a fraction below)
            es = task.get('hES', 0)  # if not found don't display gantt chart
            ef = task.get('hEF', 0)

            def drawbar(es, ef, barcolor, yof, yh):
                if es < (dh + cumh) and ef >= cumh:
                    if es <= cumh: xof = 0
                    else: xof = int( rect.width * (es - cumh)/dh)
                    if ef >= (cumh + dh): wof = 0
                    else: wof = int( rect.width * (cumh + dh - ef)/dh)
                    dc.SetBrush(wx.Brush(barcolor, wx.SOLID))
                    dc.SetPen(wx.Pen(barcolor, 1, wx.SOLID))
                    if is24:
                        dc.DrawRectangle(rect.x+xof, rect.y+yof, rect.width-wof-xof, yh)
                    else:
                        dc.DrawRectangle((rect.x+xof, rect.y+yof), (rect.width-wof-xof, yh))

            if task.get('SubtaskCount'):
                drawbar(es, ef, plancolor, 6, (rect.height-12)//2)  # half-height bar
            else:
                drawbar(es, ef, plancolor, 6, rect.height-12)

            # asd = task.get('ActualStartDate')
            # aed = task.get('ActualEndDate') or Data.Today()
            # if asd: drawbar(DateInfo[DateConv[asd]][1], DateInfo[DateConv[ase]][1], actualcolor, 4, 4)

            # bsd = task.get('BaseStartDate')
            # bed = task.get('BaseEndDate')
            # if asd: drawbar(DateInfo[DateConv[asd]][1], DateInfo[DateConv[ase]][1], actualcolor, 20, 4)

        # firstdate = table.reportcolumn[table.columns[col]]['FirstDate']
        # dateindex = Data.DateConv[firstdate] + of
        
        # date = Data.DateIndex[ index + of ]
        # rr = table.reportrow[table.rows[row]]
        # tid = rr['ID']
        # task = table.data['Task'][tid]
        # startdate = task.get('StartDate')

        # startdate = table.rows[row].get('CalculatedStartDate', None)
        # enddate = table.rows[row].get('CalculatedEndDate', None)
        # c, i = columns[col]
        # firstperiod = 
        # thisperiod = GetPeriod(
        # If startdate != None and enddate != None: 
        #         if startdate <= next period and end date > this period then skip
        #   if start date = this period and hours > 0 then no adjustment on left
        #   if end date > this period then no adjustment on right

# ------------------ Grid -------------------------

class GanttChartGrid(wx.grid.Grid):
    def __init__(self, parent, reportid):
        wx.grid.Grid.__init__(self, parent, -1) # initialize base class first

        self.table = GanttChartTable(reportid)
        self.SetTable(self.table)

        self.DisableDragRowSize()
        self.DisableDragGridSize()
        self.SetRowLabelSize(40)
        self.SetDefaultRowSize(22)  # (22, True) would resize existing rows
        self.SetDefaultColSize(40)  # (40, True) would resize existing columns

        self.SetLabelFont(wx.Font(11, wx.SWISS, wx.NORMAL, wx.BOLD))

        wx.grid.EVT_GRID_RANGE_SELECT(self, self.OnSelect)
        wx.grid.EVT_GRID_SELECT_CELL(self, self.OnSelect)

        self.Reset()

    def GetSelectedRows(self):
        """
The original implementation returns shift-selected rows as if they were shift-selected cells.
This routine adds them back as selected rows.
"""
        selrow = wx.grid.Grid.GetSelectedRows(self)  # current selection
        # if debug: print "old selrow", selrow
        seltl = self.GetSelectionBlockTopLeft()  # current selection
        selbr = self.GetSelectionBlockBottomRight()  # current selection
        # if debug: print "seltl", seltl
        # if debug: print "selbr", selbr
        cols = self.table.GetNumberCols()
        # if debug: print "cols", cols        
        if len(seltl):  # if at least one selection of cells
            for i in range(len(seltl)):
                ra, ca = seltl[i]
                rz, cz = selbr[i]
                if (ca == 0) and (cz == cols - 1):  # these look like shift-click selected rows
                    for newrow in range(ra, rz+1):
                        if not newrow in selrow: selrow.append(newrow)  # prevent duplicates from being added
            # selrow.sort()
        # if debug: print "new selrow", selrow
        return selrow

    def OnSelect(self, event):
        # if debug: print "OnSelect"
        reportid = self.table.report['ID']
        f = Data.OpenReports.get(reportid)
        if f:
            Menu.AdjustMenus(f)
        event.Skip()

    def UpdateAttrs(self):
         self.table._updateRowAttrs(self)
         self.table._updateColAttrs(self)

# UpdateColumnWidths is no good -- if the column pointers have changed, then the
# column attributes must be updated, to preserve read-only status and cell renderers.
#    -- Alex

#     def UpdateColumnWidths(self):
#         rc = Data.ReportColumn
#         for i, c in enumerate(self.table.columns):
#             of = self.table.coloffset[i]
#             ct = Data.ColumnType[self.table.ctypes[i]]
#             ctname = ct.get('Name')
#             if of > -1 and ctname and ctname[-6:] == "/Gantt":
#                 self.SetColSize(i, 24)
#             else:
#                 cw = rc[c].get('Width') or 40
#                 self.SetColSize(i, cw)
        
    def Reset(self):
        """ Reset the view based on the data in the table.

        Call this when rows are added or destroyed
        """
        self.BeginBatch()

        for current, new, delmsg, addmsg in [            (self.table._rows, self.table.GetNumberRows(), wx.grid.GRIDTABLE_NOTIFY_ROWS_DELETED, wx.grid.GRIDTABLE_NOTIFY_ROWS_APPENDED),            (self.table._cols, self.table.GetNumberCols(), wx.grid.GRIDTABLE_NOTIFY_COLS_DELETED, wx.grid.GRIDTABLE_NOTIFY_COLS_APPENDED),        ]:            if new < current:                msg = wx.grid.GridTableMessage(self.table,delmsg,new,current-new)                self.ProcessTableMessage(msg)            elif new > current:                msg = wx.grid.GridTableMessage(self.table,addmsg,new-current)                self.ProcessTableMessage(msg)                msg = wx.grid.GridTableMessage(self.table, wx.grid.GRIDTABLE_REQUEST_VIEW_GET_VALUES)
                self.ProcessTableMessage(msg)

        self.table._rows = self.table.GetNumberRows()
        self.table._cols = self.table.GetNumberCols()

        # update the column rendering plugins
        self.UpdateAttrs()

        # update the scrollbars and the displayed part of the grid
        self.AdjustScrollbars()

        self.EndBatch()


#------------------ MultiSelect Frame -----------------------------------

class MultiSelection(UI.MultipleSelection):
    def __init__(self, *args, **kwds):
        # begin wxGlade: ReportFrame.__init__
        UI.MultipleSelection.__init__(self, *args, **kwds)

        wx.EVT_BUTTON(self, self.OK.GetId(), self.onOK)
        wx.EVT_BUTTON(self, self.Cancel.GetId(), self.onCancel)
        
    # TableName  == Dependency table
    # TableIDs  == candidate prerequisites
    # ID  == selected Task ID
    # Status == status of current records

    def onOK(self, event):
        if debug: print "Start onOK"
        vals = self.SelectionListBox.GetSelections()
        if debug: print "selected values", vals
        if self.TableName == 'Dependency':
            for v in vals:  # off sets of user's selection
                k = self.Status[v]  # convert to current Dependency record keys
                if debug: print "dep rec key, task key", k, self.TableIDs[v]
                if k < 0:  # deleted, must be activated
                    change = { 'Table': 'Dependency', 'ID': - k, 'zzStatus': None }
                    Data.Update(change)
                elif k == 0: # no record, must add
                    change = { 'Table': 'Dependency', 'PrerequisiteID': self.TableIDs[v], 'TaskID': self.ID }
                    Data.Update(change)
                else: # k > 0  --  there was an active record, it is still needed
                    self.Status[v] = 0  # prevent change in next step
            for k in self.Status:  # turn off any prior dependencies records that were not selected
                if k > 0:  # active, must delete
                    change = { 'Table': 'Dependency', 'ID': k, 'zzStatus': 'deleted' }
                    Data.Update(change)
            Data.SetUndo('Set Dependencies')
        elif self.TableName == 'Assignment':
            for v in vals:  # off sets of user's selection
                k = self.Status[v]  # convert to current Assignment record keys
                if debug: print "dep rec key, task key", k, self.TableIDs[v]
                if k < 0:  # deleted, must be activated
                    change = { 'Table': 'Assignment', 'ID': - k, 'zzStatus': None }
                    Data.Update(change)
                elif k == 0: # no record, must add
                    change = { 'Table': 'Assignment', 'ResourceID': self.TableIDs[v], 'TaskID': self.ID }
                    Data.Update(change)
                else: # k > 0  --  there was an active record, it is still needed
                    self.Status[v] = 0  # prevent change in next step
            for k in self.Status:  # turn off any prior dependencies records that were not selected
                if k > 0:  # active, must delete
                    change = { 'Table': 'Assignment', 'ID': k, 'zzStatus': 'deleted' }
                    Data.Update(change)
            Data.SetUndo('Set Assignments')
        else: pass  # other uses of this method
        self.Destroy()

    def onCancel(self, event):
        self.Destroy()

#------------------ Gantt Report Frame -----------------------------------

class GanttReportFrame(UI.ReportFrame):
    def __init__(self, reportid, *args, **kwds):
        if debug: "Start GanttReport init"
        if debug: print 'reportid', reportid
        UI.ReportFrame.__init__(self, *args, **kwds)

        # these three commands were moved out of UI.ReportFrame's init
        self.report_window = GanttChartGrid(self.Report_Panel, reportid)
        self.Report = self.report_window
        self.ReportID = reportid

        # self.Report, aka report_window, is a grid, not a report or a frame
        # report_window is referenced in one place in UI.ReportFrame.do_layout
        # Report is referenced in several scripts
        #     -- Alexander

        # Data.OpenReports[reportid] = self

        self.set_properties()  # these are in the parent class
        self.do_layout()

        Menu.doAddScripts(self)
        Menu.FillWindowMenu(self)
        Menu.AdjustMenus(self)
        self.SetReportTitle()

        # file menu events
        wx.EVT_MENU(self, wx.ID_NEW,    Menu.doNew)
        wx.EVT_MENU(self, wx.ID_OPEN,   Menu.doOpen)
        wx.EVT_MENU(self, wx.ID_CLOSE,  self.doClose)
        wx.EVT_MENU(self, wx.ID_CLOSE_ALL, Menu.doCloseReports)
        wx.EVT_MENU(self, wx.ID_SAVE,   Menu.doSave)
        wx.EVT_MENU(self, wx.ID_SAVEAS, Menu.doSaveAs)
        wx.EVT_MENU(self, wx.ID_REVERT, Menu.doRevert)
        wx.EVT_MENU(self, wx.ID_EXIT, Menu.doExit)

        # edit menu events
        wx.EVT_MENU(self, wx.ID_UNDO,          Menu.doUndo)
        wx.EVT_MENU(self, wx.ID_REDO,          Menu.doRedo)

        # script menu events
        wx.EVT_MENU(self, ID.FIND_SCRIPTS, Menu.doFindScripts)
        wx.EVT_MENU_RANGE(self, ID.FIRST_SCRIPT, ID.LAST_SCRIPT, Menu.doScript)

        # window menu events
        wx.EVT_MENU_RANGE(self, ID.FIRST_WINDOW, ID.LAST_WINDOW, self.doBringWindow)

        # help menu events
        wx.EVT_MENU(self, wx.ID_ABOUT, Menu.doShowAbout)
        wx.EVT_MENU(self, ID.HOME_PAGE, Menu.doHome)
        wx.EVT_MENU(self, ID.HELP_PAGE, Menu.doHelp)
        wx.EVT_MENU(self, ID.FORUM, Menu.doForum)

        # frame events
        wx.EVT_ACTIVATE(self, self.OnActivate)
        wx.EVT_CLOSE(self, self.doClose)
        wx.EVT_SIZE(self, self.OnSize)
        wx.EVT_MOVE(self, self.OnMove)

        # grid events
        wx.grid.EVT_GRID_COL_SIZE(self, self.OnColSize)
        wx.grid.EVT_GRID_ROW_SIZE(self, self.OnRowSize)
        wx.grid.EVT_GRID_EDITOR_SHOWN(self, self.OnEditorShown)

        # tool bar events
        wx.EVT_TOOL(self, ID.INSERT_ROW, self.OnInsertRow)
        wx.EVT_TOOL(self, ID.DUPLICATE_ROW, self.OnDuplicateRow)
        wx.EVT_TOOL(self, ID.DELETE_ROW, self.OnDeleteRow)
        wx.EVT_TOOL(self, ID.MOVE_UP, self.OnMoveRow)
        wx.EVT_TOOL(self, ID.MOVE_DOWN, self. OnMoveRow)
        wx.EVT_TOOL(self, ID.PREREQUISITE, self.OnPrerequisite)
        wx.EVT_TOOL(self, ID.ASSIGN_RESOURCE, self.AssignResource)

        wx.EVT_TOOL(self, ID.HIDE_ROW, self.OnHide)
        wx.EVT_TOOL(self, ID.SHOW_HIDDEN, self.OnShowHidden)

        wx.EVT_TOOL(self, ID.INSERT_COLUMN, self.OnInsertColumn)
        wx.EVT_TOOL(self, ID.DELETE_COLUMN, self.OnDeleteColumn)
        wx.EVT_TOOL(self, ID.MOVE_LEFT, self.OnMoveColumn)
        wx.EVT_TOOL(self, ID.MOVE_RIGHT, self.OnMoveColumn)

        # wx.EVT_TOOL(self, ID.GANTT_OPTIONS, self.OnGanttOptions)
        wx.EVT_TOOL(self, ID.SCROLL_LEFT_FAR, self.OnScroll)
        wx.EVT_TOOL(self, ID.SCROLL_LEFT, self.OnScroll)
        wx.EVT_TOOL(self, ID.SCROLL_RIGHT, self.OnScroll)
        wx.EVT_TOOL(self, ID.SCROLL_RIGHT_FAR, self.OnScroll)
        wx.EVT_TOOL(self, ID.SCROLL_TO_TASK, self.OnScrollToTask)
        if debug: "End GanttReport init"

    # ------ Tool Bar Commands ---------

    def OnInsertRow(self, event):
        if debug: print "Start OnInsertRow"

        r = Data.Report[self.ReportID]
        rt = Data.ReportType[r['ReportTypeID']]
        ta = rt['TableA']
        tb = rt.get('TableB')  # if two table report all inserts go at the end (less confusing to user)
        each = rt.get('AllOrEach') in ['both', 'each']

        change = { 'Table': ta, 'Name': '--' }  # new record because no ID specified
        if ta == 'Task' or each:
            pid = r.get('ProjectID')
            change['ProjectID'] = pid 
        elif ta in ('Report', 'ReportColumn', 'ReportRow', 'ReportType', 'ColumnType'): return  # need special handling
        undo = Data.Update(change)
        newid = undo['ID']

        change = { 'Table': 'ReportRow', 'ReportID': self.ReportID, 'TableName': ta, 'TableID': newid }  
        undo = Data.Update(change)  # created here to control where inserted
        newrowid = undo['ID']

        rlist = Data.GetRowList(self.ReportID)  # list of row id's in display order

        s = self.Report.GetSelectedRows()  # current selection

        if not s: rlist.append(newrowid)  # if no selection add to end
        else:
            rid = self.Report.table.rows[min(s)] # convert to rowid
            pos = rlist.index(rid)  # find position of selected row (should always work)
            rlist.insert(pos, newrowid)  # insert before first selected row
        Data.ReorderReportRows(self.ReportID, rlist)

        Data.SetUndo('Insert ' + ta)
        if debug: print "End OnInsertRow"

    def OnDuplicateRow(self, event):
        sel = self.Report.GetSelectedRows()  # current selection
        if len(sel) == 0: 
            if debug: print "can't duplicate, empty selection"
            return
        rtid = Data.Report[self.ReportID].get('ReportTypeID')
        tablea = Data.ReportType[rtid].get('TableA')
        new = []
        for s in sel:
            rid = self.Report.table.rows[s]
            ta = Data.ReportRow[rid]['TableName']
            if ta != tablea: continue  # only duplicate rows of primary table type
            rcopy = Data.ReportRow[rid].copy()  # report row
            tid = rcopy['TableID']
            tcopy = Data.Database[ta][tid].copy()  # table row

            tcopy['Table'] = ta; del tcopy['ID']
            undo = Data.Update(tcopy)
            rcopy['Table'] = 'ReportRow'; del rcopy['ID']
            rcopy['TableID'] = undo['ID']
            undo = Data.Update(rcopy)
            new.append(undo['ID'])

        rlist = Data.GetRowList(self.ReportID)  # list of row id's in display order
        where = max(sel) + 1
        # print "rlist", rlist
        # print "new", new
        rlist[where:where] = new
        # print "new rlist", rlist
        Data.ReorderReportRows(self.ReportID, rlist)

        Data.SetUndo('Duplicate Row')

    def OnDeleteRow(self, event):  # this is a shallow delete
        if debug: print "Start OnDeleteRow"
        sel = self.Report.GetSelectedRows()  # current selection
        if len(sel) < 1: 
            if debug: print "can't delete, no rows selected"
            return  # only move if rows selected
        change = { 'Table': None, 'ID': None, 'zzStatus': 'deleted' }
        cnt = 0
        for s in sel:
            rid = self.Report.table.rows[s]
            ta = Data.ReportRow[rid].get('TableName')
            id = Data.ReportRow[rid].get('TableID')
            if not id: continue  # silently skip invalid table id's
            if ta == 'Project' and id == 1: continue  # certain projects and reports can't be deleted
            elif ta == 'Report' and (id == 1 or id == 2): continue
            if Data.Database[ta][id].get('zzStatus') == 'deleted': 
                change['zzStatus'] = None
            else:
                change['zzStatus'] = 'deleted'
            change['Table'] = ta
            change['ID'] = id
            undo = Data.Update(change)
            cnt += 1
        if cnt > 0: Data.SetUndo('Delete/Reactivate Row')
        if debug: print "End OnDeleteRow"

    def OnMoveRow(self, event):
        """ Move selected rows up or down by one position """
        sel = self.Report.GetSelectedRows()
        if not sel:
            return

        if event.GetId() == ID.MOVE_UP:
            step = -1
        else: # ID.MOVE_DOWN
            step = 1

        rows = self.Report.table.rows
        rowlevels = self.Report.table.rowlevels

        # move only the best-ranking rows of the selection
        level = max(rowlevels)

        move = []
        for x in sel:
            lev = rowlevels[x]
            if lev == level:
                move.append(x)
            elif lev < level:
                level = lev
                move = [x]
        move.sort()

        row_set = dict.fromkeys(move)  # set of row positions
        move_row_ids = [rows[x] for x in move]  # list of report row ids

        # determine whether the selection is contiguous
        #   (ignoring children and invisible rows)
        start = move[0]
        end = move[-1] + 1

        for x in xrange(start, end):
            if x in row_set:
                continue
            lev = rowlevels[x]
            if lev <= level:
                contiguous = False
                break
        else:
            contiguous = True

            # find the next visible, same-parent row
            next_row_id = None

            if step < 0:
                x = start - 1
            else:
                x = end

            while 0 <= x < len(rows):
                lev = rowlevels[x]
                if lev == level:
                    next_row_id = rows[x]
                    break
                elif lev < level:
                    break
                x += step

        # check for invisible rows
        if not self.Report.table.report.get('ShowHidden'):
            # report may include rows that are not displayed

            # get report rows from database; convert row positions
            rows, rowlevels = Data.GetRowLevels(self.ReportID)

            positions = {}
            for x, rid in enumerate(rows):
                positions[rid] = x

            row_set = {}
            for rowid in move_row_ids:
                if rowid in positions:
                    pos = positions[rowid]
                    row_set[pos] = None
            if not row_set:
                return

            move = row_set.keys()
            move.sort()

            start = move[0]
            end = move[-1] + 1
        else:
            # don't modify the row list directly
            rows = rows[:]

        if not contiguous:
            # consolidate rows
            if step < 0:
                end = start + len(move)
            else:
                start = end - len(move)

            move.reverse()
            for x in move:
                del rows[x]
            rows[start:start] = move_row_ids

        elif next_row_id:
            # move next row to the other side of selection
            if step < 0:
                dest = end - 1
            else:
                dest = start
            rows.remove(next_row_id)
            rows.insert(dest, next_row_id)

        else:
            # nothing changed
            return

        # apply the new row order
        Data.ReorderReportRows(self.ReportID, rows)
        Data.SetUndo('Move Row')

    def OnPrerequisite(self, event):
        # list tasks in the same order they appear now  -- use self.Report.table.rows
        # highlight the ones that are currently prerequisites
        sel = self.Report.GetSelectedRows()  # current selection
        if len(sel) < 1: 
            if debug: print "must select at least one row"
            return
        elif len(sel) > 1:
            sel.sort()  # chain tasks in the order they appear on report
            rows = self.Report.table.rows
            alltids = [ Data.ReportRow[rows[x]].get('TableID') for x in sel if Data.ReportRow[rows[x]].get('TableName') == 'Task' ]
            tids = [ x for x in alltids if not Data.Task[x].get('SubtaskCount') ]
            if len(tids) > 1:
                for i in range(len(tids) - 1):  # try to match and link each pair
                    # look for an existing dependency record
                    did = Data.FindID('Depedency', 'PrerequisiteID', tids[i], 'TaskID', tids[i+1])
                    if did:
                        if Data.Database['Dependency'][did].get('zzStatus') == 'deleted':
                            change = { 'Table': 'Dependency', 'ID': did, 'zzStatus': None }
                            Data.Update(change)
                    else:
                        change = { 'Table': 'Dependency', 'PrerequisiteID': tids[i], 'TaskID': tids[i+1] }
                        Data.Update(change)
                Data.SetUndo("Set Dependencies")
            return
        rows = self.Report.table.rows
        sel = sel[0]
        rowid = rows[sel]  # get selection's task id
        ta = Data.ReportRow[rowid].get('TableName')
        if ta != 'Task': return  # only on task rows
        sid = Data.ReportRow[rowid].get('TableID')
        assert sid, "tried to assign prereq's to an invalid task id"

        alltid = [ Data.ReportRow[x].get('TableID') for x in rows if Data.ReportRow[x].get('TableName') == 'Task']
        tid = [ x for x in alltid if Data.Task[x].get('zzStatus') != 'deleted' and x != sid]  # active displayed tasks
        tname = [ Data.Task[x].get('Name', "") or "--" for x in tid ]

        status = [0] * len(tid)  # array of 0's
        for k, v in Data.Dependency.iteritems():
            if sid != v.get('TaskID'): continue
            p = v.get('PrerequisiteID')
            try:
                i = tid.index(p)
            except: pass
            else:
                if v.get('zzStatus') != 'deleted':
                    status[i] = k
                else:
                    status[i] = -k

        dialog = MultiSelection(self, -1, "", size=(240,320))
        dialog.Instructions.SetLabel("Select prerequisite tasks:")
        # dialog.SelectionListBox.Clear()
        dialog.SelectionListBox.Set(tname)
        for i, v in enumerate(status):
            if v > 0:
                dialog.SelectionListBox.SetSelection(i)
        dialog.TableName = 'Dependency'
        dialog.TableIDs = tid
        if debug: print "task id's", tid
        dialog.ID = sid  # selected task id
        if debug: print "selected task id", sid
        dialog.Status = status  # status of dependency candidates
        if debug: print "task status", status

        dialog.Centre()  # centers dialog on screen
        dialog.ShowModal()

    def AssignResource(self, event):
        # list resources in alphabetical order
        # highlight the ones that are currently assigned
        sel = self.Report.GetSelectedRows()  # user's selection
        if len(sel) != 1: 
            if debug: print "only assignments for one task"
            return  # only move if rows selected
        rows = self.Report.table.rows
        sel = sel[0]  # make it not a list
        rowid = rows[sel]  # get selected task's id
        sid = Data.ReportRow[rowid].get('TableID')

        ta = Data.ReportRow[rowid].get('TableName')
        if ta != 'Task': return  # only on task rows

        assert sid and sid > 0, "tried to assign prereq's to an invalid task id"

        res = {}
        for k, v in Data.Resource.iteritems():
            if v.get('zzStatus') == 'deleted': continue
            name = v.get('Name')
            if name: res[name] = k

        names = res.keys()
        names.sort()
        # print 'names', names
        ids = [ res[x] for x in names ]

        status = [ 0 for x in range(len(ids)) ]  # array of 0's
        for k, v in Data.Assignment.iteritems():
            if sid != v.get('TaskID'): continue
            p = v.get('ResourceID')
            try:
                i = ids.index(p)
            except: pass
            else:
                if v.get('zzStatus') != 'deleted':
                    status[i] = k
                else:
                    status[i] = -k

        dialog = MultiSelection(self, -1, "", size=(240,320))
        dialog.Instructions.SetLabel("Select assigned resources:")
        # dialog.SelectionListBox.Clear()
        dialog.SelectionListBox.Set(names)
        for i, v in enumerate(status):
            if v > 0:
                dialog.SelectionListBox.SetSelection(i)
        dialog.TableName = 'Assignment'
        dialog.TableIDs = ids
        if debug: print "resource id's", ids
        dialog.ID = sid  # selected task id
        if debug: print "selected task id", sid
        dialog.Status = status  # status of dependency candidates
        if debug: print "assignment status", status

        dialog.Centre()  # centers dialog on screen
        dialog.ShowModal()

    def OnHide(self, event):
        sel = self.Report.GetSelectedRows()  # user's selection
        Menu.onHide(self.Report.table, event, sel)

    def OnShowHidden(self, event):
        Menu.onShowHidden(self, event)

    def OnInsertColumn(self, event):
        if debug: print "Start OnInsertColumn"

        r = Data.Report[self.ReportID]
        rtid = r.get('ReportTypeID')  # these determine what kind of columns can be inserted
        also = Data.ReportType[rtid].get('Also')

        menuid = []  # list of types to be displayed for selection
        rlist = Data.GetRowList(2)  # report 2 defines the sequence of the report type selection list
        for k in rlist:
            rr = Data.ReportRow[k]
            hidden = rr.get('Hidden', False)
            table = rr.get('TableName')  # should always be 'ReportType' or 'ColumnType'
            id = rr.get('TableID')
            if (not hidden) and table == 'ColumnType' and id:
                xrtid = Data.ColumnType[id].get('ReportTypeID')
                active = Data.ColumnType[id].get('zzStatus') != 'deleted'
                if active and xrtid and ( (rtid == xrtid) or (also and (also == xrtid)) ):
                    menuid.append( id )

        menutext = [ (Data.ColumnType[x].get('Label') or Data.ColumnType[x].get('Name')) for x in menuid ]
        for i, v in enumerate(menutext):  # remove line feeds before menu display
            if '\n' in v:
                menutext[i] = v.replace('\n', ' ')
        if debug: print menuid, menutext
        dlg = wxMultipleChoiceDialog(self,
                         "Select columns to add:",
                            "New Columns", menutext, style=wx.DEFAULT_FRAME_STYLE, size=(240, 320))
        dlg.Centre()
        if (dlg.ShowModal() != wx.ID_OK): return
        newlist = dlg.GetValue()
        addlist = []

        change = { 'Table': 'ReportColumn', 'ReportID': self.ReportID }  # new record because no ID specified
        for n in newlist:
            change['ColumnTypeID'] = menuid[n]
            ct = Data.ColumnType[menuid[n]]
            if ct['Name'] == 'Day/Gantt':  # leaving this for compatibility with version 0.1
                change['Periods'] = 21
                change['FirstDate'] = Data.GetToday()
                undo = Data.Update(change)
                del change['Periods']; del change['FirstDate']
            elif ct.get('AccessType') == 's':
                change['Periods'] = 14
                # change['FirstDate'] = Data.GetToday()
                undo = Data.Update(change)
                del change['Periods']
                # del change['FirstDate']
            else:
                change['Width'] = ct.get('Width')
                undo = Data.Update(change)
                del change['Width']
            addlist.append(undo['ID'])

        clist = Data.GetColumnList(self.ReportID)  # list of column id's in display order

        s = self.Report.GetSelectedCols()  # current selection
        if debug: print 'selection', s
        if len(s) == 0:
            clist.extend(addlist)  # if no selection add to end
        else: 
            addlist.reverse()  # insert the list in reverse order so the same insertion point can be used for all
            for a in addlist:
                clist.insert(min(s), a)  # insert before first selected row

        Data.ReorderReportColumns(self.ReportID, clist)

        Data.SetUndo('Insert Column')
        if debug: print "End OnInsertColumn"

    def OnDeleteColumn(self, event):
        if debug: print "Start OnDeleteColumn"
        sel = self.Report.GetSelectedCols()
        if not sel: return

        clist = Data.GetColumnList(self.ReportID)
        cids = self.Report.table.columns

        sel_cids = [cids[x] for x in sel]
        sel_map = dict.fromkeys(sel_cids)
        clist = [x for x in clist if x not in sel_map]

        # apply the new column order
        Data.ReorderReportColumns(self.ReportID, clist)
        Data.SetUndo('Delete Column')

        # clear the selection
        self.Report.ClearSelection()

        if debug: print "End OnDeleteColumn"

    def OnMoveColumn(self, event):
        """ Move selected columns left or right by one position """
        sel = self.Report.GetSelectedCols()
        if not sel:
            return

        if event.GetId() == ID.MOVE_LEFT:
            step = -1
        else: # ID.MOVE_RIGHT
            step = 1

        cids = self.Report.table.columns

        # get report columns from database
        #   (grid uses multiple grid columns for certain report columns)
        clist = Data.GetColumnList(self.ReportID)
        cids = self.Report.table.columns

        map = {}
        for sel_x in sel:
            try:
                x = clist.index(cids[sel_x])
            except:
                continue
            map[x] = None
        if not map:
            return

        move = map.keys()
        move.sort()

        # perform the movement
        start = move[0]
        end = move[-1] + 1

        if end - start > len(move):
            # consolidate columns
            if step < 0:
                end = start + len(move)
            else:
                start = end - len(move)

            move_cids = [clist[x] for x in move]
            move.reverse()
            for x in move:
                del clist[x]

            clist[start:start] = move_cids
        else:
            # move past the next column
            if step < 0:
                x = start - 1
                dest = end - 1
            else:
                x = end
                dest = start

            if 0 <= x < len(clist):
                cid = clist.pop(x)
                clist.insert(dest, cid)
                start += step
                end += step
            else:
                return

        # apply the new column order
        Data.ReorderReportColumns(self.ReportID, clist)
        Data.SetUndo('Move Column')

    def OnScroll(self, event):  # need test to make sure this is a gantt report??
        """ scroll the selected  """
        if debug: print "Start OnScroll"
        # figure out which gantt chart to scroll -- either the ones w/ selected columns or the first one 
        scrollcols = []
        sel = self.Report.GetSelectedCols()  # current selection
        if len(sel) > 0: 
            for s in sel:
                cid = self.Report.table.columns[s]
                ctid = Data.ReportColumn[cid]['ColumnTypeID']
                if Data.ColumnType[ctid].get('AccessType') == 's':
                    if scrollcols.count(cid) == 0: 
                        scrollcols.append(cid)
        if len(scrollcols) == 0:
            clist = Data.GetColumnList(self.ReportID)  # complete list of column id's in display order
            for cid in clist:
                ctid = Data.ReportColumn[cid]['ColumnTypeID']
                if Data.ColumnType[ctid].get('AccessType') == 's':
                    scrollcols.append(cid)
                    break
        offset = 0
        id = event.GetId()  # move left or right?
        if id == ID.SCROLL_LEFT_FAR: offset = 7
        elif id == ID.SCROLL_LEFT: offset = 1
        elif id == ID.SCROLL_RIGHT: offset = -1
        elif id == ID.SCROLL_RIGHT_FAR: offset = -7
        change = { 'Table': 'ReportColumn', 'FirstDate': None, 'ID': None }
        somethingChanged = False
        for s in scrollcols:
            rc = Data.ReportColumn[s]
            dateConvIndex = rc.get('FirstDate')
            if Data.DateConv.has_key(dateConvIndex):
                 datex = Data.DateConv[dateConvIndex]
            else:
                 datex = Data.DateConv[Data.GetToday()]

            ctid = rc['ColumnTypeID']
            timep = Data.ColumnType[ctid].get('Name')
            if debug: print 'column type name:', timep
            if not timep:
                continue
            elif timep[0]  == 'D':
                datex += offset
            elif timep[0]  == 'W':
                datex -= Data.DateInfo[ datex ][2]  # convert to beginning of week
                datex += offset * 7
            elif timep[0]  == 'M':
                datex = Data.GetPeriodStart(timep, datex, offset)
            elif timep[0]  == 'Q':
                datex = Data.GetPeriodStart(timep, datex, offset)
            else:  # if we don't recognize the time period
                continue
            newdate = Data.DateIndex[datex]
            change['ID'] = s
            change['FirstDate'] = newdate
            Data.Update(change)
            somethingChanged = True
        if somethingChanged: Data.SetUndo('Scroll Timescale Columns')
        if debug: print "End Scroll"

    def OnScrollToTask(self, event):  # need test to make sure this is a gantt report??
        if debug: print "Start ScrollToTask"
        sel = self.Report.GetSelectedRows()  # current selection
        if len(sel) < 1: 
            if debug: print "can't scroll, no tasks selected"
            return  # only move if columns selected
        s = self.Report.table.rows[min(sel)]  # just use first row
        rs = Data.ReportRow[s]
        if rs.get('TableName') != 'Task': 
            if debug: print "tablename", rs.get('TableName')
            return  # can only scroll to tasks
        tid = rs.get('TableID')
        newdate = Data.ValidDate(Data.Task[tid].get('StartDate')) or Data.Task[tid].get('CalculatedStartDate')
        if debug: print "new date", newdate
        if not newdate: return  # not date to scroll to
        scrollcols = []
        clist = Data.GetColumnList(self.ReportID)  # complete list of row id's in display order
        for c in clist:
            ctid = Data.ReportColumn[c]['ColumnTypeID']
            if Data.ColumnType[ctid].get('AccessType') == 's':
                scrollcols.append(c)
        change = { 'Table': 'ReportColumn', 'FirstDate': None, 'ID': None }
        somethingChanged = False
        for s in scrollcols:
            datex = Data.DateConv[newdate]
            ctid = Data.ReportColumn[s]['ColumnTypeID']
            timep = Data.ColumnType[ctid].get('Name')
            if not timep: continue
            ix = Data.GetPeriodStart(timep, datex, 0)
            if not ix: continue
            newdate = Data.DateIndex[datex]
            #elif timep[0]  == 'D':
            #    newdate = Data.DateIndex[datex]
            #elif timep[0]  == 'W':
            #    datex -= Data.DateInfo[ datex ][2]  # convert to beginning of week
            #    newdate = Data.DateIndex[datex]
            #else:  # if we don't recognize the time period
            #    continue
            change['ID'] = s
            change['FirstDate'] = newdate
            Data.Update(change)
            somethingChanged = True
        if somethingChanged: Data.SetUndo('Scroll to Task')
        if debug: print "End ScrollToTask"

        # wx.EVT_TOOL(self, ID.COLUMN_OPTIONS, self.OnGanttOptions)

    # ---- menu Commands -----

    def doClose(self, event):
        Data.CloseReport(self.ReportID)

    def doBringWindow(self, event):
        Menu.doBringWindow(self, event)

    # ----------------- window size/position

    def OnSize(self, event):
        size = event.GetSize()
        r = Data.Database['Report'][self.ReportID]
        r['FrameSizeW'] = size.width
        r['FrameSizeH'] = size.height

        event.Skip()  # to call default handler; needed?

    def OnMove(self, event):
        pos = event.GetPosition()
        if pos.x > 0 and pos.y > 0:
            r = Data.Database['Report'][self.ReportID]
            r['FramePositionX'] = pos.x
            r['FramePositionY'] = pos.y

        event.Skip()  # needed?

    # ---------------- activate

    def OnActivate(self, event):
        if event.GetActive():
            Data.ActiveReport = self.ReportID
        else:
            self.Report.SaveEditControlValue()
            self.Report.HideCellEditControl()
        event.Skip()

    # ---------------- Column/Row resizing
    def OnRowSize(self, evt):
        pass

    def OnColSize(self, evt):
        if debug: print "OnColSize", (evt.GetRowOrCol(), evt.GetPosition())
        if debug: print "get col width", self.Report.GetColSize(evt.GetRowOrCol())
        newsize = self.Report.GetColSize(evt.GetRowOrCol())
        col = evt.GetRowOrCol()
        if self.Report.table.coloffset[col] == -1:  # not a gantt column
            colid = self.Report.table.columns[col]
            change = { 'Table': 'ReportColumn', 'ID': colid, 'Width': newsize }
            Data.Update(change, 0)  #  --------------------- don't allow Undo (until I can figure out how)
            # Data.SetUndo('Change Column Width')

        evt.Skip()

    # -------------- Make sure only the right tables are edited per column
    def OnEditorShown(self, evt):
        rtid = Data.Report[self.ReportID].get('ReportTypeID')

        rid = self.Report.table.rows[evt.GetRow()]
        rtable = Data.ReportRow[rid].get('TableName')

        cid = self.Report.table.columns[evt.GetCol()]
        ctid = Data.ReportColumn[cid].get('ColumnTypeID')
        which = Data.ColumnType[ctid].get('T') or 'Z'  # can be 'A', 'B', or 'X'

        ctable = Data.ReportType[rtid].get('Table' + which)  # 'X' should always yield 'None'

        if which != 'X' and ((not ctable) or rtable != ctable):
            evt.Veto()
            return
        evt.Skip()

    # -----------
    def SetReportTitle(self):
        rname = Data.Report[self.ReportID].get('Name') or "-"
        pid = Data.Report[self.ReportID].get('ProjectID')
        if Data.Project.has_key(pid):
            pname = Data.Project[pid].get('Name') or "-"
        else:
            pname = "-"
        title = pname + " / " + rname
        if self.GetTitle() != title:
            self.SetTitle(title)
            Menu.UpdateWindowMenuItem(self.ReportID)

    def UpdatePointers(self, all=0):  # 1 = new database; 0 = changed report rows or columns
        if debug: print "Start GanttReport UpdatePointers"

        # don't refresh a report if the underlying report record is invalid
        if all or not Data.Report[self.ReportID].get('ReportTypeID'):
            Data.CloseReport(self.ReportID)
            return

        sr = self.Report.table
        rlen, clen = len(sr.rows), len(sr.columns)

        select_rows = {}
        for x in self.Report.GetSelectedRows():
             rowid = sr.rows[x]
             select_rows[rowid] = None
        select_columns = {}
        for x in self.Report.GetSelectedCols():
             colid = sr.columns[x]
             select_columns[colid] = None

        sr.UpdateColumnPointers()
        sr.UpdateRowPointers()

        self.Report.BeginBatch()

        if rlen != len(sr.rows) or clen != len(sr.columns):
            self.Report.Reset()  # tell grid that the number of rows or columns has changed
        else:
            self.Report.UpdateAttrs()

        self.Report.ClearSelection()
        for x, rowid in enumerate(sr.rows):
             if rowid in select_rows:
                 self.Report.SelectRow(x, True)
        for x, colid in enumerate(sr.columns):
             if colid in select_columns:
                 self.Report.SelectCol(x, True)

        self.Report.EndBatch()

        if debug: print "End GanttReport UpdatePointers"

#---------------------------------------------------------------------------

if __name__ == '__main__':
    app = wx.PySimpleApp()
    frame = GanttReportFrame(3, None, -1, "")  # reportid = 3
    frame.Show(True)
    app.MainLoop()

if debug: print "end GanttReport.py"


