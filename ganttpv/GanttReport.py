#!/usr/bin/env python
# gantt report definition

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

# 040408 - first version of this program
# 040409 - will display task values in report
# 040410 - SetValue will now set Task fields
# 040419 - changes for new ReportRow column names
# 040420 - changes to allow Main.py to open gantt reports; doClose will close this report; 
#               saves Report size and position
# 040424 - added doInsertTask, doDuplicateTask, doDeleteTask
# 040426 - added doMoveRow, added doPrerequisite
# 040427 - added doAssignResource, onHide, onShowHidden; fixed doDelete; 
#               save the result when the user changes columns widths; prevent changes
#               to row height; changed some column names in ReportColumn
# 040503 - changes to use new ReportType and ColumnType tables; added OnInsertColumn; 
#               added row highlighting for two table reports
# 040504 - finished OnInsertColumn; added OnDeleteColumn, and OnMoveColumn
# 040505 - support UI changes (many references to 'Task' replaced with 'Row', format changes, and minor adjustments);
#               revised OnInsertRow, OnDuplicateRow, and OnDeleteRow to work with all tables; added OnScroll and
#               OnScrollToTask; use colors from Data.Option for chart
# 040508 - prevent deletion of project 1 or reports 1 or 2
# 040510 - set minimum row size for non-gantt column in "_updateColAttrs"; set Hidden and Deleted row color;
#               copied Scripts menu processing from Main.py; changed doClose to use Data.doClose
# 040512 - added OnEditorShown; added support for indirect column types; several bug fixes
# 040518 - added doHome
# 040528 - in doDuplicate only dup rows of primary table's type; in onPrerequisite only include Tasks in the 
#               list of potential prerequisites
# 040531 - in onDraw make sure all rectangles use new syntax for version 2.5
# 040715 - Pierre_Rouleau@impathnetworks.com: removed all tabs, now use 4-space indentation level to comply with Official Python Guideline.
# 040906 - changed OnInsertColumn to ignore Labels w/ value of None; display "project name / report name" in 
#               report title.
# 040914 - handle indirect display if no ID is found; assign project id when inserting rows into reports that can be each.
# 040915 - allow entry of floating point numbers (type "f")
# 040918 - add week time scale; default FirstDate in column type to today or this week
# 040928 - Alexander - ignores dates not present in Data.DateConv; prevents entry of incorrectly formatted dates
# 041001 - display measurements in weekly time scale
# 041009 - Alexander & Brian - default dates to current year and month; Brian - change SetValue to work on measurement time scale data
# 041009 - change scroll to work w/ any time scale column
# 041010 - change "column insert" to set # periods for all timescale columns; changes to allow edit of non-measurement 
#		time scale columns
# 041126 - draw bars for week timescale

import wx, wx.grid
import datetime
from wxPython.lib.dialogs import wxMultipleChoiceDialog
import Data, UI, ID, Menu
# import images

debug = 1
mac = 1
is24 = 0

if debug: print "load GanttReport.py"

# ------------ Table behind grid ---------

class GanttChartTable(wx.grid.PyGridTableBase):
    """
    A custom wxGrid Table using user supplied data
    """
    def __init__(self, reportidx):
        """ data is taken from SampleData
        """
        # The base class must be initialized *first*
        wx.grid.PyGridTableBase.__init__(self)

        reportid = reportidx

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
        of = self.coloffset[col]
        if of == -1:
            label = self.reportcolumn[self.columns[col]].get('Label')
            if not label or label == "":
                ct = self.columntype[self.ctypes[col]]  # get column type record that corresponds to this column
                label = ct.get('Label') or ct.get('Name')
        else:
            ct = self.columntype[self.ctypes[col]]  # get column type record that corresponds to this column
            ctperiod, ctfield = ct.get('Name').split("/")

            firstdate = self.reportcolumn[self.columns[col]].get('FirstDate')
            if not Data.DateConv.has_key(firstdate): firstdate = Data.GetToday()
            index = Data.DateConv[ firstdate ]

            if ctfield == 'Gantt':
                if ctperiod == "Day":
                    date = Data.DateIndex[ index + of ]
                    if of == 0 or date[8:10] == '01': label = date[5:7]
                    else: 
                        dow = Data.DateInfo[ index + of ][2]
                        label = 'MTWHFSS'[dow]
                    label += '\n' +  date[8:10]
                elif ctperiod == "Week":
                    index -= Data.DateInfo[ index ][2]  # convert to beginning of week
                    date = Data.DateIndex[ index + (of * 7) ]
                    if of == 0 or date[8:10] <= '07': label = date[5:7]
                    else: label = ''
                    label += '\n' +  date[8:10]
                else:
                    label = "-"  # unknown time scale
            else: 
                if ctperiod == "Day":
                    date = Data.DateIndex[ index + of ]
                elif ctperiod == "Week":
                    index -= Data.DateInfo[ index ][2]  # convert to beginning of week
                    date = Data.DateIndex[ index + (of * 7) ]
                else:
                    return "-"  # unknown time scale

                if of == 0:
                    label = ctfield[:5]  # ??try column width mod 8??
                else:
                    label = ''
                label += '\n' + date[5:7] + "/" + date[8:10]
        return label

    # default behavior is to number the rows  ---- option to include the task name ?????
    # def GetRowLabelValue(self, row):
    #    print 'grlv', row  # not currently used???
    #     return row + 1  #  self.reportrow[self.rows[row]]['TaskID']

    def GetValue(self, row, col):
        of = self.coloffset[col]
        ct = self.columntype[self.ctypes[col]]
        if of == -1:
            rr = self.reportrow[self.rows[row]]
            rtable = rr.get('TableName')
            tid = rr['TableID']  # was 'TaskID' -> changed to generic ID

            # rc = self.reportcolumn[self.columns[col]]

            t = ct.get('T', 'X')
            rtid = self.report.get('ReportTypeID')
            ctable = Data.ReportType[rtid].get('Table' + t)

            at = ct.get('AccessType')
            if rtable != ctable:
                value = ''
            elif at == 'd':
                column = ct.get('Name')
                # print column  # it prints each column twice - why???
                value = self.data[rtable][tid].get(column, "")
            elif at == 'i':
                try:
                    it, ic = ct.get('Name').split('/')  # indirect table & column
                except ValueError:
                    if debug: print "Indirect column w/o 'Table/Column', Name is: ", ct.get('Name')
                    value = ""
                else:
                    iid = self.data[rtable][tid].get(it+'ID')
                    # if debug: print "rtable, tid, it, ic, iid", rtable, tid, it, ic, iid
                    if iid:
                        value = self.data[it][iid].get(ic, "")
                    else:
                        value = ""
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
                value = 'gantt'
            else:  # table name, field name, time period, and record id
                rr = self.reportrow[self.rows[row]]
                tablename = rr.get('TableName')
                tid = rr['TableID']
                if ctfield == 'Measurement':  # find field name
                    mid = Data.Database[tablename][tid].get('MeasurementID')  # point at measurement record
                    if mid: 
                        fieldname = Data.Database['Measurement'][mid].get('Name')  # measurement name == field name
                    else:
                        fieldname = None
                    tid = Data.Database[tablename][tid].get('ProjectID')  # point at measurement record
                    tablename = 'Project'  # only supports project measurements
                else:
                    fieldname = ctfield

                # find the period date
                firstdate = self.reportcolumn[self.columns[col]].get('FirstDate')
                if not Data.DateConv.has_key(firstdate): firstdate = Data.GetToday()
                index = Data.DateConv[ firstdate ]
                if ctperiod == "Day":
                    date = Data.DateIndex[ index + of ]
                elif ctperiod == "Week":
                    index -= Data.DateInfo[ index ][2]  # convert to beginning of week
                    date = Data.DateIndex[ index + (of * 7) ]
                else:
                    date = None
                timename = tablename + ctperiod

                timeid = Data.FindID(timename, tablename + "ID", tid, 'Period', date)
                # if debug: print "timeid", timeid
                if timeid:
                    value = Data.Database[timename][timeid].get(fieldname)
                    # if debug: print "timename, timeid, fieldname, value: ", timename, timeid, fieldname, value
                    # if debug: print "record: ", Data.Database[timename][timeid]
                else:
                    value = None
                    # if debug: print "didn't find timeid", timeid, value

        if value == None: value = ''
        return value

    def GetRawValue(self, row, col):  # same as GetValue  ( I don't know the difference. The example I'm following made them the same. )
        value = GetValue(self, row, col)
#        of = self.coloffset[col]
#        if of == -1:
#            rr = self.reportrow[self.rows[row]]
#            rtable = rr.get('TableName')
#            tid = rr['TableID']  # was 'TaskID' -> changed to generic ID

#            # rc = self.reportcolumn[self.columns[col]]
#            ct = self.columntype[self.ctypes[col]]

#            t = ct.get('T', 'X')
#            rtid = self.report.get('ReportTypeID')
#            ctable = Data.ReportType[rtid].get('Table' + t)

#            at = ct.get('AccessType')
#            if rtable != ctable:
#                value = ''
#            elif at == 'd':
#                column = ct.get('Name')
#                # print column  # it prints each column twice - why???
#                value = self.data[rtable][tid].get(column, "")
#            elif at == 'i':
#                it, ic = ct.get('Name').split('/')  # indirect table & column
#                iid = self.data[rtable][tid].get(it+'ID')
#                # if debug: print "rtable, tid, it, ic, iid", rtable, tid, it, ic, iid
#                if iid:
#                    value = self.data[it][iid].get(ic, "")
#                else:
#                    value = ""
#        else:
#            value = 'gantt'
#        if value == None: value = ''
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
            if len(value) in (2, 5, 8):
                td = Data.GetToday()
                value = td[0:-len(value)] + value
            if value < '1901-01-01':
                return
            elif Data.DateConv.has_key(value):
                v = value
            elif len(value) == 10 and value[4] == '-' and value[7] == '-':
                try:
                    datetime.date(int(value[0:4]),int(value[5:7]), int(value[8:10]))
                    v = value
                    Data.ChangedCalendar = True
                except ValueError:
                    return
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
        elif at == 's':  # time scale ??
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
        if debug: print '---- New Value', change
        Data.Update(change)
        Data.SetUndo(column)

    def ResetView(self, grid):
        """
        (wxGrid) -> Reset the grid view.   Call this to
        update the grid if rows and columns have been added or deleted
        """
        grid.BeginBatch()
        for current, new, delmsg, addmsg in [
            (self._rows, self.GetNumberRows(), wx.grid.GRIDTABLE_NOTIFY_ROWS_DELETED, wx.grid.GRIDTABLE_NOTIFY_ROWS_APPENDED),
            (self._cols, self.GetNumberCols(), wx.grid.GRIDTABLE_NOTIFY_COLS_DELETED, wx.grid.GRIDTABLE_NOTIFY_COLS_APPENDED),
        ]:
            if new < current:
                msg = wx.grid.GridTableMessage(self,delmsg,new,current-new)
                grid.ProcessTableMessage(msg)
            elif new > current:
                msg = wx.grid.GridTableMessage(self,addmsg,new-current)
                grid.ProcessTableMessage(msg)
                self.UpdateValues(grid)
        grid.EndBatch()

        self._rows = self.GetNumberRows()
        self._cols = self.GetNumberCols()

        # update the column rendering plugins
        self._updateColAttrs(grid)
        self._updateRowAttrs(grid)  # maybe highlight some rows

        # update the scrollbars and the displayed part of the grid
        grid.AdjustScrollbars()
        grid.ForceRefresh()


    def UpdateValues(self, grid):
        """Update all displayed values"""
        # This sends an event to the grid table to update all of the values
        msg = wx.grid.GridTableMessage(self, wx.grid.GRIDTABLE_REQUEST_VIEW_GET_VALUES)
        grid.ProcessTableMessage(msg)

    # -- These are not part of the standard interface - my routines to make display easier

    def UpdateDataPointers(self, reportid):
        Data.UpdateDataPointers(self, reportid)

        # self.data = Data.Database
        # self.report = self.data["Report"][reportid]
        # self.reportcolumn = self.data["ReportColumn"]
        # self.reportrow = self.data["ReportRow"]

    def UpdateColumnPointers(self):
        if debug: print 'Start UpdateColumnPointers'

        rc = self.reportcolumn  # pointer to table
        rct = self.columntype  # pointer to table
        c = self.report.get('FirstColumn', 0)
        self.columns = []
        self.ctypes = []
        self.coloffset = []
        while c != 0 and c != None:
            # get info to decide whether this is a "timescale" column
            ct = rc[c]['ColumnTypeID']
            type = rct[ct].get('AccessType')
            if (type == 's'):
                for i in range(rc[c].get('Periods', 1)):
                    self.columns.append( c )
                    self.ctypes.append( ct )
                    self.coloffset.append( i )
            else:
                self.columns.append( c )
                self.ctypes.append( ct )
                self.coloffset.append( -1 )
            c = rc[c].get('NextColumn', 0)
            assert self.columns.count( c ) == 0, 'Loop in report column pointers' 
        if debug: print 'End UpdateColumnPointers'

    def UpdateRowPointers(self):
        Data.UpdateRowPointers(self)
        # rr = self.reportrow
        # show = self.report.get('ShowHidden', False)
        # r = self.report.get('FirstRow', 0)
        # self.rows = []
        # while r != 0 and r != None:
        #         if show:
        #                 self.rows.append( r )
        #         else:
        #                 hidden = rr[r].get('Hidden', False)
        #                 table = rr[r].get('TableName')
        #                 id = rr[r].get('TableID')
        #                 active = table and id and (Data.Database[table][id].get('zzStatus', 'active') == 'active')
        #                 if (not hidden) and active:
        #                         self.rows.append( r )
        #         # self.rows.append( r )
        #         r = rr[r].get('NextRow', 0)
        #         assert self.rows.count( r ) == 0, 'Loop in report column pointers' 
        
    def _updateColAttrs(self, grid):
        """
        wxGrid -> update the column attributes to add the
        appropriate renderer given the column type.

        Otherwise default to the default renderer.
        """

        for col in range(len(self.columns)):
            attr = wx.grid.GridCellAttr()
            of = self.coloffset[col]
            ct = self.columntype[self.ctypes[col]]  # if the column is in a report it will always have a type
            ctname = ct.get('Name')
            # if ctfield != "Gantt": return
            if of > -1 and ctname and ctname[-6:] == "/Gantt":
                renderer = GanttCellRenderer(self)
                if renderer.colSize:
                    grid.SetColSize(col, renderer.colSize)
                if renderer.rowSize:
                    grid.SetDefaultRowSize(renderer.rowSize)
                attr.SetReadOnly(True)
                attr.SetRenderer(renderer)
            else:
                # if debug: print "updateColAttrs ctname", ctname
                # add logic here to protect columns
                cid = self.columns[col]
                rc = Data.ReportColumn[cid]
                ctid = rc['ColumnTypeID']
                grid.SetColSize(col, rc.get('Width') or 40)
                grid.SetDefaultRowSize(22)      # make this big enough so text is fully displayed while editted
                attr.SetReadOnly(not Data.ColumnType[ctid].get('Edit'))
                # attr.SetRenderer(None)
            grid.SetColAttr(col, attr)
            col += 1

    def _updateRowAttrs(self, grid):
        """ Highlight parent rows if two type of rows are show """

        rtid = self.report['ReportTypeID']
        rt = Data.ReportType[rtid]
        ta, tb = rt.get('TableA'), rt.get('TableB')
        # if tb == None or tb == '': return

        for row in range(len(self.rows)):
            rr = self.reportrow[self.rows[row]]
            rowtable = rr['TableName']
            rowid = rr['TableID']
            deleted = Data.Database[rowtable][rowid].get('zzStatus', 'active') == 'deleted'
            hidden = rr.get('Hidden', False)

            attr = wx.grid.GridCellAttr()

            if deleted:
                attr.SetBackgroundColour(Data.Option.get('DeletedColor', "pale green"))
            elif hidden:
                attr.SetBackgroundColour(Data.Option.get('HiddenColor', "pale green"))
            elif rr.get('TableName') == ta and tb:  # parent table in two level report
                attr.SetBackgroundColour(Data.Option.get('ParentColor', "pale green"))
            else:
                attr.SetBackgroundColour(Data.Option.get('ChildColor', "pale green"))
                # attr.SetTextColour(wxBLACK)
                # attr.SetFont(wxFont(10, wxSWISS, wxNORMAL, wxBOLD))
                # attr.SetReadOnly(True)
                # self.SetRowAttr(row, attr)
            grid.SetRowAttr(row, attr)  # reset all other rows back to default values
            # row += 1

# ---------------------- Draw Cells of Grid -----------------------------
# Sample wxGrid renderers

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
        # self.colSize = None
        self.colSize = 24
        # self.rowSize = None
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

        fdate = rc.get('FirstDate')
        if not fdate or not Data.DateConv.has_key(fdate): return  # this routine shouldn't be called for not gantt columns, but it is anyway - just ignore them
                                # the program seems to be refreshing three times when one is needed (040505)

        # if debug: print "-- didn't return --"

        ix = Data.DateConv[fdate]
        of = self.table.coloffset[col]

        if ctname[0:3] == "Day":
            dh, cumh, dow = Data.DateInfo[ix  + of]
        elif ctname[0:4] == "Week":
            dh, cumh, dow = Data.DateInfo[ix  + of * 7]
            dh2, cumh2, dow2 = Data.DateInfo[ix  + (of + 1) * 7]
            dh = cumh2 - cumh
        else: return

        # clear the background
        dc.SetBackgroundMode(wx.SOLID)
        if isSelected and dh:
            backcolor = o.get('WorkDaySelected', wx.BLUE)
        elif isSelected:
            backcolor = o.get('NotWorkDaySelected', wx.BLUE)
        elif dh:
            backcolor = o.get('WorkDay', wx.BLUE)
        else:
            backcolor = o.get('NotWorkDay', wx.BLUE)

        dc.SetBrush(wx.Brush(backcolor, wx.SOLID))
        dc.SetPen(wx.Pen(backcolor, 1, wx.SOLID))
        if is24:
            dc.DrawRectangle(rect.x, rect.y, rect.width, rect.height)
        else:
            dc.DrawRectangle((rect.x, rect.y), (rect.width, rect.height))
        # draw gantt bar
        if dh >= 1:  # only display bars on days that have working hours 

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
                    # dc.DrawRectangle(rect.x+xof, rect.y+6, rect.width-wof-xof, rect.height-12)
                    if is24:
                        dc.DrawRectangle(rect.x+xof, rect.y+yof, rect.width-wof-xof, yh)
                    else:
                        dc.DrawRectangle((rect.x+xof, rect.y+yof), (rect.width-wof-xof, yh))
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
        """
        """
        wx.grid.Grid.__init__(self, parent, -1) # initialize base class first
        self._table = GanttChartTable(reportid)
        self.table = self._table  # treat table as an attribute that should be accessable
                                  # (eventually convert all "_table" references to "table")
        self.SetTable(self._table)
        self.DisableDragRowSize()
        self.UpdateColumnWidths()
        self.SetRowLabelSize(40)

        # this has no effect
        self.SetDefaultRowSize(20)  # less than the gantt renderer; (28, True) would mean resize existing rows

        wx.grid.EVT_GRID_RANGE_SELECT(self, self.OnSelect)
        wx.grid.EVT_GRID_SELECT_CELL(self, self.OnSelect)

    def OnSelect(self, event):
        # if debug: print "OnSelect"
        reportid = self._table.report['ID']
        f = Data.OpenReports[reportid]
        Menu.AdjustMenus(f)
        event.Skip()

     # def __set_properties(self):
     #    wx.grid.Grid.__set_properties(self)
     #     self.grid_1.CreateGrid(10, 3)
     #     self.grid_1.EnableDragColSize(0)
     #     self.grid_1.EnableDragRowSize(0)

     # def UpdateColumnWidths(self):
     #    self._table.UpdateColumnWidths():

    def UpdateColumnWidths(self):
        rc = self._table.reportcolumn  # pointer to table
        for i, c in enumerate(self._table.columns):
            of = self._table.coloffset[i]
            if of == -1:
                cw = rc[c].get('Width', 40) or 40
                # print "column number:", i, "width:", cw
                self.SetColSize(i, cw)
        
    def Reset(self):
        """reset the view based on the data in the table.  Call
        this when rows are added or destroyed"""
        self._table.ResetView(self)


#------------------ MultiSelect Frame -----------------------------------

class MultiSelection(UI. MultipleSelection):
    def __init__(self, *args, **kwds):
        # begin wxGlade: ReportFrame.__init__
        kwds["style"] = wx.DEFAULT_FRAME_STYLE
        UI. MultipleSelection.__init__(self, *args, **kwds)

        # these three commands were moved out of UI.ReportFrame's init
        # self.report_window = GanttChartGrid(self, reportid)
        # self.report_window.Reset()
        # self.Report = self.report_window
        # self.ReportID = reportid

        # self.set_properties()  # these are in the parent class
        # self.do_layout()
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
                    change = { 'Table': 'Dependency', 'ID': - k, 'zzStatus': 'active' }
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
                    change = { 'Table': 'Assignment', 'ID': - k, 'zzStatus': 'active' }
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
        if debug: print 'args', args
        if debug: print 'kwds', kwds
        # begin wxGlade: ReportFrame.__init__
        kwds["style"] = wx.DEFAULT_FRAME_STYLE
        UI.ReportFrame.__init__(self, *args, **kwds)

        # these three commands were moved out of UI.ReportFrame's init
        self.report_window = GanttChartGrid(self.Report_Panel, reportid)
        self.title = None # force update of title
        self.report_window.Reset()
        self.Report = self.report_window
        self.ReportID = reportid

        self.set_properties()  # these are in the parent class
        self.do_layout()

        Menu.doAddScripts(self)
        # Data.OpenReports[reportid] = self
        Menu.AdjustMenus(self)

    # def __init__(self, parent, reportid):
    #     wx.Frame.__init__(self, parent, -1,
    #                      "Test Frame", size=(640,480))
 
    #    grid = GanttChartGrid(self, reportid)
    #    grid.Reset()

    # def Reset(self): 
    #    """ Call this routine whenever the database is replaced """
    #    UpdateDataPointers(self)
    #    UpdateColumnPointers(self)
    #    UpdateRowPointers(self)
    #    ResetView(self, self.grid)

    # def Refresh(self):
    #    """ Call this routine whenever data in the database changes """
    #    UpdateValues(self, self.grid)

    # ----- Menu and Toolbars

        # Associate each menu/toolbar item with the method that handles that
        # item.
        if 1:  # mac only [TODO: check OS type instead]
            wx.App_SetMacAboutMenuItemId(ID.ABOUT)
            # wx.App_SetMacPreferencesMenuItemId(),
            wx.App_SetMacExitMenuItemId(wx.ID_EXIT)
            # wx.App_SetMacHelpMenuTitleName("&Help")
            # wx.EVT_MENU(self, wx.ID_ABOUT,    self.doShowAbout)
            # wx.EVT_MENU(self, wx.ID_PREFERENCES,    self.doNew)

        wx.EVT_MENU(self, wx.ID_NEW,    self.doNew)
        wx.EVT_MENU(self, wx.ID_OPEN,   self.doOpen)
        wx.EVT_MENU(self, wx.ID_CLOSE,  self.doClose)
        wx.EVT_MENU(self, wx.ID_SAVE,   self.doSave)
        wx.EVT_MENU(self, wx.ID_SAVEAS, self.doSaveAs)
        # wx.EVT_MENU(self, wx.ID_REVERT, self.doRevert)
        wx.EVT_MENU(self, wx.ID_EXIT, self.doExit)
        # if 0:  # not mac [TODO: check OS type instead]
        #        wx.EVT_MENU(self, wx.ID_EXIT,   self.doExit)

        wx.EVT_MENU(self, ID.UNDO,          self.doUndo)
        wx.EVT_MENU(self, ID.REDO,          self.doRedo)

        wx.EVT_MENU(self, ID.FIND_SCRIPTS, self.doFindScripts)
        # Menu.doAddScripts(self)
        wx.EVT_MENU_RANGE(self, Menu.FirstScriptID, Menu.FirstScriptID + 1000, self.doScript)

        wx.EVT_MENU(self, ID.ABOUT, self.doShowAbout)
        wx.EVT_MENU(self, ID.HOME_PAGE, self.doHome)
        wx.EVT_MENU(self, ID.HELP_PAGE, self.doHelp)
        wx.EVT_MENU(self, ID.FORUM, self.doForum)

        # Install our own method to handle closing the window.  This allows us
        # to ask the user if he/she wants to save before closing the window, as
        # well as keeping track of which windows are currently open.

        wx.EVT_CLOSE(self, self.doClose)
        wx.EVT_SIZE(self, self.OnSize)  # report window size
        wx.EVT_MOVE(self, self.OnMove)
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

        if debug: print " Start OnInsertRow"
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

        if len(s) == 0 or tb: rlist.append(newrowid)  # if no selection add to end
        else: rlist.insert(min(s), newrowid)  # insert before first selected row
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
            rid = self.Report._table.rows[s]
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

    def OnDeleteRow(self, event):  # this is a shallow delete -- should delete deep like I do in Main?????
        if debug: print "Start OnDeleteRow"
        sel = self.Report.GetSelectedRows()  # current selection
        if len(sel) < 1: 
            if debug: print "can't delete, no rows selected"
            return  # only move if rows selected
        change = { 'Table': None, 'ID': None, 'zzStatus': 'deleted' }
        cnt = 0
        for s in sel:
            rid = self.Report._table.rows[s]
            ta = Data.ReportRow[rid].get('TableName')
            id = Data.ReportRow[rid].get('TableID')
            if not id: continue  # silently skip invalid table id's
            if ta == 'Project' and id == 1: continue  # certain projects and reports can't be deleted
            elif ta == 'Report' and (id == 1 or id == 2): continue
            if Data.Database[ta][id].get('zzStatus', 'active') == 'deleted': 
                change['zzStatus'] = 'active'
            else:
                change['zzStatus'] = 'deleted'
            change['Table'] = ta
            change['ID'] = id
            if debug: print "change", change
            undo = Data.Update(change)
            cnt += 1
        if cnt > 0: Data.SetUndo('Delete/Reactivate Row')
        if debug: print "End OnDeleteRow"

    # def OnEdit(self, event):
    #    if self.report_list.currentItem:
    #            # if project
    #            dlg = wx.TextEntryDialog(frame, 'New project name', 'Edit Project', 'Python')
    #            dlg.SetValue("Python is the best!")
    #            if dlg.ShowModal() == wxID_OK:
    #                    log.WriteText('You entered: %s\n' % dlg.GetValue())
    #            dlg.Destroy()

    def OnMoveRow(self, event):
        """ move selected rows up or down one row """
        sel = self.Report.GetSelectedRows()  # current selection
        if len(sel) < 1: 
            if debug: print "can't move, no rows selected"
            return  # only move if rows selected
        id = event.GetId()  # move up or move down?

        # find first and last rows (these use screen row offsets)
        first = min(sel)
        if first == 0 and id == ID.MOVE_UP: 
            if debug: print "can't move up"
            return  # can't move up, we're already there
        last = max(sel)
        rows = self.Report._table.rows  # list of displayed rows
        if last == len(rows) - 1 and id == ID.MOVE_DOWN: 
            if debug: print "can't move down"
            return  # can't move down, we're already there

        rlist = Data.GetRowList(self.ReportID)  # complete list of row id's in display order
        sel.sort()  # probably sorted already, but just in case
        sel.reverse()  # simpler to process them in descending order (all inserts at same offset)

        # screen list of rows may not match complete list (example: ShowHidden == False)
        # adjust screen offsets to offsets in complete list
        # screen row --> report row id --> position of row id in complete list
        seladj = [ rlist.index(rows[x]) for x in sel ]

        if id == ID.MOVE_UP:
            before = rlist.index(rows[first-1])  # destination
            rids = [ rlist.pop(x) for x in seladj  ]  # remove the row ids that are being moved
            for x in rids: rlist.insert(before, x)  # insert the row ids before the "before" row
        elif id == ID.MOVE_DOWN:
            after = rlist.index(rows[last+1])  # destination
            rids = [ rlist.pop(x) for x in seladj ]  # remove the rows that are being moved
            a = after - len(rids) + 1
            for x in rids: rlist.insert(a, x)  # insert the rows after the "after" row on screen
        else: return  # shouldn't happen

        Data.ReorderReportRows(self.ReportID, rlist)
        Data.SetUndo('Move Row')

        # make sure the same rows are selected so the same rows can be moved again
        # selection will be thrown off by undo -- is that a problem? (selection is not an undo-able event)
        # sel = [ self.Report._table.rows.index(x) for x in rids ]  #~~ find the new positions
        self.Report.ClearSelection()
        if id == ID.MOVE_UP: sel = [ x - 1 for x in sel ]  # assume everything moved one row
        else: sel = [ x + 1 for x in sel ]
        # self.Report.SelectRow(sel.pop(), False)
        for x in sel: self.Report.SelectRow(x, True)

    def OnPrerequisite(self, event):
        # list tasks in the same order they appear now  -- use self.Report._table.rows
        # highlight the ones that are currently prerequisites
        sel = self.Report.GetSelectedRows()  # current selection
        if len(sel) != 1: 
            if debug: print "only prereq's for one task"
            return  # only move if rows selected
        rows = self.Report._table.rows
        sel = sel[0]
        rowid = rows[sel]  # get selection's task id
        ta = Data.ReportRow[rowid].get('TableName')
        if ta != 'Task': return  # only on task rows
        sid = Data.ReportRow[rowid].get('TableID')
        assert sid and sid > 0, "tried to assign prereq's to an invalid task id"

        alltid = [ Data.ReportRow[x].get('TableID') for x in rows if Data.ReportRow[x].get('TableName') == 'Task']
        tid = [ x for x in alltid if Data.Task[x].get('zzStatus', 'active') == 'active' and x != sid]  # active displayed tasks
        tname = [ Data.Task[x].get('Name', "") or "--" for x in tid ]

        status = [ 0 for x in range(len(tid)) ]  # array of 0's
        for k, v in Data.Dependency.iteritems():
            if sid != v.get('TaskID'): continue
            p = v.get('PrerequisiteID')
            try:
                i = tid.index(p)
            except: pass
            else:
                if v.get('zzStatus', 'active') == 'active':
                    status[i] = k
                else:
                    status[i] = -k

        dialog = MultiSelection(self, -1, "")
        # dialog = MultiSelection(None, -1, "")
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
        dialog.Show(True)

    def AssignResource(self, event):
        # list resources in alphabetical order
        # highlight the ones that are currently assigned
        sel = self.Report.GetSelectedRows()  # user's selection
        if len(sel) != 1: 
            if debug: print "only assignments for one task"
            return  # only move if rows selected
        rows = self.Report._table.rows
        sel = sel[0]  # make it not a list
        rowid = rows[sel]  # get selected task's id
        sid = Data.ReportRow[rowid].get('TableID')

        ta = Data.ReportRow[rowid].get('TableName')
        if ta != 'Task': return  # only on task rows

        assert sid and sid > 0, "tried to assign prereq's to an invalid task id"

        res = {}
        for k, v in Data.Resource.iteritems():
            if v.get('zzStatus', 'active') != 'active': continue
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
                if v.get('zzStatus', 'active') == 'active':
                    status[i] = k
                else:
                    status[i] = -k

        dialog = MultiSelection(self, -1, "")
        # dialog = MultiSelection(None, -1, "")
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
        dialog.Show(True)

    def OnHide(self, event):
        sel = self.Report.GetSelectedRows()  # user's selection
        Menu.onHide(self.Report._table, event, sel)

    def OnShowHidden(self, event):
        Menu.onShowHidden(self, event)

    def OnInsertColumn(self, event):
        if debug: print "Start OnInsertColumn"

        r = Data.Report[self.ReportID]
        rtid = r.get('ReportTypeID')  # these determine what kind of columns can be inserted
        also = Data.ReportType[rtid].get('Also')

        r2 = Data.Report[2]  # report 2 defines the sequence of the column type selection list
        menuid = []  # list of column types to be displayed for selection
        k = r2.get('FirstRow', 0)
        loopcheck = 0
        while k != 0 and k != None:
            rr = Data.ReportRow[k]
            hidden = rr.get('Hidden', False)
            table = rr.get('TableName')  # should always be 'ReportType' or 'ColumnType'
            id = rr.get('TableID')
            if (not hidden) and table == 'ColumnType' and id:
                xrtid = Data.ColumnType[id].get('ReportTypeID')
                active = Data.ColumnType[id].get('zzStatus', 'active') == 'active'
                if active and xrtid and ( (rtid == xrtid) or (also and (also == xrtid)) ):
                    menuid.append( id )
            k = rr.get('NextRow', 0)
            loopcheck += 1
            if loopcheck > 100000:  break  # prevent endless loop if data is corrupted
        menutext = [ (Data.ColumnType[x].get('Label') or Data.ColumnType[x].get('Name')) for x in menuid ]
        if debug: print menuid, menutext
        dlg = wxMultipleChoiceDialog(self,
                         "Select columns to add:",
                            "New Columns", menutext)
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
        # r = Data.Report[self.ReportID]

        s = self.Report.GetSelectedCols()  # current selection
        if len(s) == 0: return  # nothing to delete

        change = { 'Table': 'ReportColumn', 'zzStatus': 'deleted' }
        for i in s:
            change['ID'] = self.Report._table.columns[i]
            Data.Update(change)

        Data.ReorderReportColumns(self.ReportID, [])  # keeps columns in same order, omits deleted ones

        Data.SetUndo('Delete Column')
        if debug: print "End OnDeleteColumn"
        if debug: print "GetColumnList", Data.GetColumnList(self.ReportID)

    def OnMoveColumn(self, event):
        """ move selected columns left or right one position. I assume selection is contiguous. """
        if debug: print "Start MoveColumn"
        sel = self.Report.GetSelectedCols()  # current selection
        if debug: print "selection", sel
        if len(sel) < 1: 
            if debug: print "can't move, no columns selected"
            return  # only move if columns selected
        id = event.GetId()  # move left or right?
        if debug: print "event id", id
        clist = Data.GetColumnList(self.ReportID)  # complete list of row id's in display order
        first = self.Report._table.columns[min(sel)]  # first and last selected column id's
        firstoff = self.Report._table.coloffset[min(sel)]  # save original offset
        last = self.Report._table.columns[max(sel)]
        if debug: print "first, firstoff, last", first, firstoff, last

        # find position of first and last in clist
        f = clist.index(first)
        l = clist.index(last)
        cnt = l - f + 1  # number of selected column. can't use sel because same column may appear multiple times
        if debug: print "f, l, cnt", f, l, cnt

        # find first and last rows (these use screen row offsets)
        if f == 0 and id == ID.MOVE_LEFT: 
            if debug: print "can't move left"
            return  # can't move left, we're already there
        if l == len(clist) - 1 and id == ID.MOVE_RIGHT: 
            if debug: print "can't move right"
            return  # can't move right, we're already there

        if debug: print "clist", clist
        if id == ID.MOVE_LEFT:
            before = f-1  # destination
            temp = clist.pop(before)  # remove the column in front of selection
            clist.insert(before + cnt, temp)  # re-insert behind
        elif id == ID.MOVE_RIGHT:
            after = l + 1  # destination
            temp = clist.pop(after) # remove the column after selection
            clist.insert(after - cnt, temp)  # re-insert before
        else: return  # shouldn't happen
        if debug: print "clist", clist

        Data.ReorderReportColumns(self.ReportID, clist)
        Data.SetUndo('Move Column')

        # make sure the same columns are selected so the same columns can be moved again
        # selection will be thrown off by undo -- is that a problem? (selection is not an undo-able event)
        for i in range(len(self.Report._table.columns)):
            if self.Report._table.columns[i] == first and self.Report._table.coloffset[i] == firstoff:
                break
        sel = [ i + x for x in range(len(sel)) ]
        if debug: print 'new sel', sel

        self.Report.ClearSelection()
        for x in sel: self.Report.SelectCol(x, True)
        if debug: print "End MoveColumn"
        if debug: print "GetColumnList", Data.GetColumnList(self.ReportID)

    def OnScroll(self, event):  # need test to make sure this is a gantt report??
        """ scroll the selected  """
        if debug: print "Start OnScroll"
        # figure out which gantt chart to scroll -- either the ones w/ selected columns or the first one 
        scrollcols = []
        sel = self.Report.GetSelectedCols()  # current selection
        if len(sel) > 0: 
            for s in sel:
                cid = self.Report._table.columns[s]
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
        if debug: print 'scrollcols, change:', scrollcols, change
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
                newdate = Data.DateIndex[datex + offset]
            elif timep[0]  == 'W':
                datex -= Data.DateInfo[ datex ][2]  # convert to beginning of week
                newdate = Data.DateIndex[datex + offset * 7]
            else:  # if we don't recognize the time period
                continue
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
        s = self.Report._table.rows[min(sel)]  # just use first row
        rs = Data.ReportRow[s]
        if rs.get('TableName') != 'Task': 
            if debug: print "tablename", rs.get('TableName')
            return  # can only scroll to tasks
        tid = rs.get('TableID')
        newdate = Data.ValidDate(Data.Task[tid].get('StartDate')) or Data.Task[tid].get('CalculatedStartDate')
        if debug: print "new date", newdate
        if newdate == None or "": return  # not date to scroll to
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
            if not timep:
                continue
            elif timep[0]  == 'D':
                newdate = Data.DateIndex[datex]
            elif timep[0]  == 'W':
                datex -= Data.DateInfo[ datex ][2]  # convert to beginning of week
                newdate = Data.DateIndex[datex]
            else:  # if we don't recognize the time period
                continue
            change['ID'] = s
            change['FirstDate'] = newdate
            Data.Update(change)
            somethingChanged = True
        if somethingChanged: Data.SetUndo('Scroll to Task')
        if debug: print "End ScrollToTask"

        # wx.EVT_TOOL(self, ID.COLUMN_OPTIONS, self.OnGanttOptions)

    # ---- Menu Command -----

    # File Menu
    def doNew(self, event):
        Menu.doNew(self, event)

    def doOpen(self, event):
        Menu.doOpen(self, event)

    def doClose(self, event):
        # if Data.ChangedData:
        #    if not Data.AskIfUserWantsToSave(self, "closing"): return
        Data.CloseReport(self.ReportID)
        # Data.Report[self.ReportID]["Open"] = False
        # del Data.OpenReports[self.ReportID]
        # self.Destroy()

    def doSave(self, event):
        Menu.doSave(self, event)

    def doSaveAs(self, event):
        Menu.doSaveAs(self, event)

    def doExit(self, event):
        Menu.doExit(self, event)

    # Edit Menu
    def doUndo(self, event):
        # if debug: print Data.ReportColumn[11]
        # if debug: print Data.ReportColumn[13]
        Menu.doUndo(self, event)
        # if debug: print "GetColumnList", Data.GetColumnList(self.ReportID)
        # if debug: print Data.ReportColumn[11]
        # if debug: print Data.ReportColumn[13]

    def doRedo(self, event):
        Menu.doRedo(self, event)
        # if debug: print "GetColumnList", Data.GetColumnList(self.ReportID)

    # Script Menu
    def doFindScripts(self, event):
        Menu.doFindScripts(self, event)

    def doScript(self, event):
        Menu.doScript(self, event)

    # Help Menu
    def doShowAbout(self, event):
        Menu.doShowAbout(self, event)

    def doHome(self, event):
        Menu.doHome(self, event)

    def doHelp(self, event):
        Menu.doHelp(self, event)

    def doForum(self, event):
        Menu.doForum(self, event)

    # ----------------- window size/position 
    def OnSize(self, event):
        size = event.GetSize()
        # print size
        r = Data.Database['Report'][self.ReportID]
        r['FrameSizeW'] = size.width
        r['FrameSizeH'] = size.height

        event.Skip()  # to call default handler; needed?

    def OnMove(self, event):
        pos = event.GetPosition()
        # print pos
        r = Data.Database['Report'][self.ReportID]
        r['FramePositionX'] = pos.x
        r['FramePositionY'] = pos.y

        event.Skip()  # needed?

    # ---------------- Column/Row resizing
    def OnRowSize(self, evt): pass
        # self.log.write("OnRowSize: row %d, %s\n" %
        #                (evt.GetRowOrCol(), evt.GetPosition()))
        # evt.Skip()

    def OnColSize(self, evt):
        if debug: print "OnColSize", (evt.GetRowOrCol(), evt.GetPosition())
        if debug: print "get col width", self.Report.GetColSize(evt.GetRowOrCol())
        newsize = self.Report.GetColSize(evt.GetRowOrCol())
        col = evt.GetRowOrCol()
        if self.Report._table.coloffset[col] == -1:  # not a gantt column
            colid = self.Report._table.columns[col]
            change = { 'Table': 'ReportColumn', 'ID': colid, 'Width': newsize }
            if debug: print 'change', change
            Data.Update(change, 0)  #  --------------------- don't allow Undo (until I can figure out how)
            # Data.SetUndo('Change Column Width')

        evt.Skip()

    # -------------- Make sure only the right tables are edited per column
    def OnEditorShown(self, evt):
        rtid = Data.Report[self.ReportID].get('ReportTypeID')

        rid = self.Report._table.rows[evt.GetRow()]
        rtable = Data.ReportRow[rid].get('TableName')

        cid = self.Report._table.columns[evt.GetCol()]
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
        if pid:
            pname = Data.Project[pid].get('Name') or "-"
        else:
            pname = "-"
        title = pname + " / " + rname
        if self.title != title:
            self.SetTitle(title)

    def UpdatePointers(self, all=0):  # 1 = new database; 0 = changed report rows or columns
        if debug: print "Start Update Gantt Report Pointers"
        # don't refresh a report if the underlying report record is invalid
        if not Data.Report[self.ReportID].get('ReportTypeID'):  # happens if a report is "undone"
            Data.CloseReport(self.ReportID)
            return
        if all:  # shouldn't happen. All reports should be closed before opening a new database
            self.Report._table.UpdateDataPointers()
        self.Report._table.UpdateColumnPointers()
        self.Report._table.UpdateRowPointers()
        self.Report.Reset()  # tell grid that the number of rows or columns has changed
        if debug: print "End Update Gantt Report Pointers"

#---------------------------------------------------------------------------

if __name__ == '__main__':
    app = wx.PySimpleApp()
    frame = GanttReportFrame(3, None, -1, "")  # reportid = 3
    frame.Show(True)
    app.MainLoop()

if debug: print "end GanttReport.py"


