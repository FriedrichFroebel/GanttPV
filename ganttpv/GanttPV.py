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

# 040412 - first version of this program
# 040414 - added menu/toolbar processing; added links to web pages
# 040417 - moved menu processing to Menu.py so it can be more easily shared
# 040419 - changes for new ReportRow column names; added RefreshReport to display updated data
# 040420 - change to show 'About' under the application menu instead of the 'Help' menu;
#               double click on report name to open or show report; close will close this report;
#               will open reports where they were before
# 040421 - fixed New Report button; added Duplicate button logic; added Delete button logic
# 040422 - added reactivate logic to Delete button; added ShowHidden logic; add logic to change 
#               colors for different row types
# 040424 - added menu event handling for scripts
# 040427 - some changes to ReportColumn fields names
# 040503 - changes to use new ReportType and ColumnType tables
# 040505 - support new UI changes; temporarily make all new reports of type "Task"
# 040506 - updated NewReport button to allow selection of report
# 040508 - prevent deletion of project 1 or reports 1 or 2
# 040510 - fixed doClose and doExit
# 040512 - use Alex's report icon instead of smiley face; new reports for specific projects can be any task report
# 040518 - added doHome
# 040520 - renamed this script from Main.py to GanttPV.py
# 040528 - don't allow duplication of project id #1
# 040616 - change Target End Date column width to 120 (doesn't show heading as two lines for some reason)
# 040715 - Pierre_Rouleau@impathnetworks.com: removed all tabs, now use 4-space indentation level to comply with Official Python Guideline.
# 040716 - minor fix to tab changes
# 040914 - ReportTypes can be flagged as all/each/both for all projects or individual project use

import wx
import wx.grid
# from wxPython.lib.dialogs import wxMultipleChoiceDialog  # use single selection for adding new reports
# import webbrowser
import os, os.path, sys
import Data, UI, ID, Menu, GanttReport

# import images

debug = 1
mac = 1
if debug: print "load Main.py"

#----------------------------------------------------------------------

class ProjectReportList(wx.ListCtrl):
    def __init__(self, parent):
        # The base class must be initialized *first*
        wx.ListCtrl.__init__(self, parent, -1,
                            style=wx.LC_REPORT|wx.LC_VIRTUAL|wx.LC_HRULES|wx.LC_VRULES|wx.LC_SINGLE_SEL)
#        self.log = log

        self.currentItem = None
        self.il = wx.ImageList(16, 16)
#        self.idx1 = self.il.Add(images.getSmilesBitmap())
        if debug: print "bitmap path", os.path.join(Data.Path, "icons", "Report.bmp")
        self.idx1 = self.il.Add(wx.Bitmap(os.path.join(Data.Path, "icons", "Report.bmp"), wx.BITMAP_TYPE_BMP))
        self.SetImageList(self.il, wx.IMAGE_LIST_SMALL)

        # create local pointers to data and report definitions
        self.UpdateDataPointers()
        self.UpdateColumnPointers()
        self.UpdateRowPointers()

        self.InsertColumn(0, "Name")
        self.InsertColumn(1, "Start Date")
        self.InsertColumn(2, "Target\nEnd Date")
        self.SetColumnWidth(0, 200)
        self.SetColumnWidth(1, 80)
        self.SetColumnWidth(2, 120)

        other = Data.Option
        self.attrhidden = wx.ListItemAttr()
        self.attrhidden.SetBackgroundColour(other.get('HiddenColor', 'light grey'))

        self.attrdeleted = wx.ListItemAttr()
        self.attrdeleted.SetBackgroundColour(other.get('DeletedColor', "pink"))

        self.attrproject = wx.ListItemAttr()
        self.attrproject.SetBackgroundColour(other.get('ParentColor', "pale green"))

        self.attrreport = wx.ListItemAttr()
        self.attrreport.SetBackgroundColour(other.get('ChildColor', "khaki"))

        wx.EVT_LIST_ITEM_SELECTED(self, self.GetId(), self.OnItemSelected)
        wx.EVT_LIST_ITEM_ACTIVATED(self, self.GetId(), self.OnItemActivated)
        wx.EVT_LIST_ITEM_DESELECTED(self, self.GetId(), self.OnItemDeselected)


    def OnItemSelected(self, event):
        self.currentItem = event.m_itemIndex
        Menu.AdjustMenus(Data.OpenReports[1])
#        self.log.WriteText('OnItemSelected: "%s", "%s", "%s", "%s"\n' %
#                           (self.currentItem,
#                            self.GetItemText(self.currentItem),
#                            self.getColumnText(self.currentItem, 1),
#                            self.getColumnText(self.currentItem, 2)))

    def OnItemActivated(self, event):  # double clicked or "ENTER" while selected
        item = event.m_itemIndex
        self.currentItem = item

        rr = self.reportrow[self.rows[item]]
        rowtable = rr['TableName']
        rowid = rr['TableID']

#       rc = self.reportcolumn[self.columns[col]]
#       column = rc['A']

#       value = self.data[rowtable][rowid].get(column, '---')

        if debug: print 'rowid:', rowid
        if rowtable == 'Project': pass
                # if project open project edit window

        elif rowtable == 'Report':
            frame = Data.OpenReports.get(rowid)  # row id == report id
            if frame:
                #frame.Show(True)  # no, what command will bring the frame up front?
                frame.Raise()  # Window method
            else:
                Data.OpenReport(rowid)

        Menu.AdjustMenus(Data.OpenReports[1])

#        self.log.WriteText("OnItemActivated: %s\nTopItem: %s\n" %
#                           (self.GetItemText(self.currentItem), self.GetTopItem()))

    def getColumnText(self, index, col):
        item = self.GetItem(index, col)
        return item.GetText()

    def OnItemDeselected(self, evt):
        self.currentItem = None
        Menu.AdjustMenus(Data.OpenReports[1])
        if debug: print "OnItemDeselected"

#----------- provide data to list display

    def OnGetItemText(self, item, col):
        rr = self.reportrow[self.rows[item]]
        rowtable = rr['TableName']
        rowid = rr['TableID']
        # rc = self.reportcolumn[self.columns[col]]
        rc = self.columntype[self.ctypes[col]]
        # print 'rc', rc
        column = rc['Name']
        # print table, column  # it prints each column twice???
        value = self.data[rowtable][rowid].get(column, '---')
        if value == None: value = ''
        return value

    def OnGetItemImage(self, item):
        rr = self.reportrow[self.rows[item]]
        rowtable = rr['TableName']
        # rowid = rr['TableID']
        if rowtable == 'Report':
            return self.idx1
        else:
            return -1

    def OnGetItemAttr(self, item):
        rr = self.reportrow[self.rows[item]]
        rowtable = rr['TableName']
        rowid = rr['TableID']
        deleted = Data.Database[rowtable][rowid].get('zzStatus', 'active') == 'deleted'
        hidden = rr.get('Hidden', False)
        if deleted:
            return self.attrdeleted
        elif hidden:
            return self.attrhidden
        elif rowtable == 'Project':
            return self.attrproject
        elif rowtable  == 'Report':
            return self.attrreport
        else:
            return None

# These are not part of the standard interface - my routines to make display easier
    def UpdateDataPointers(self):
        Data.UpdateDataPointers(self, 1)

#    def UpdateDataPointers(self, reportid):
#       if debug: print "Start UpdateDataPointers"
#       # create local pointers to database
#       self.data = Data.Database
#       # pointers to one record
#       self.report = self.data["Report"][1]
#       reporttypeid = self.report['ReportTypeID']
#       #if debug: print 'reporttypeid', reporttypeid
#       self.reporttype = self.data["ReportType"][reporttypeid]
#       # pointers to tables
#       self.reportcolumn = self.data["ReportColumn"]
#       self.columntype = self.data["ColumnType"]
#       self.reportrow = self.data["ReportRow"]
#       if debug: print "End UpdateDataPointers"

    def UpdateColumnPointers(self):
        if debug: print "Start UpdateColumnPointers"
        # if debug: print "self", self
        rc = self.reportcolumn  # pointer to table
        c = self.report.get('FirstColumn', 0)
        self.columns = []
        self.ctypes = []
        self.coloffset = []
        while c != 0 and c != None:
            # tn = rc[c].get('TableName', 'table name not found')
            # cn = rc[c].get('ColumnName', 'column name not found')
            ct = rc[c]['ColumnTypeID']
            self.columns.append( c )
            self.ctypes.append( ct )
            self.coloffset.append( -1 )
            c = rc[c].get('NextColumn', 0)
            assert self.columns.count( c ) == 0, 'Loop in report column pointers' 
        if debug: print "End UpdateColumnPointers"

    def UpdateRowPointers(self):
        if debug: print "Start UpdateRowPointers"
        Data.UpdateRowPointers(self)
        self.SetItemCount(len(self.rows))  # number of items in virtual list
        if debug: print "End UpdateRowPointers"

    def RefreshReport(self):
        if debug: print "Start RefreshReport"
        if debug: print 'self', self
        self.RefreshItems(1, len(self.rows))
        if debug: print "End RefreshReport"
#----------------------------------------------------------------------

# def runTest(frame, nb, log):
#     win = ProjectReportList(nb, log)
#     return win

#----------------------------------------------------------------------
class ProjectReportFrame(UI.MainFrame):
    def __init__(self, reportid, *args, **kwds):
        if debug: print 'ProjectReportFrame __init__'
        # print 'self', self
        if debug: print 'reportid', reportid
        if debug: print 'args', args
        if debug: print 'kwds', kwds
        # begin wxGlade: ReportFrame.__init__
        kwds["style"] = wx.DEFAULT_FRAME_STYLE
        UI.MainFrame.__init__(self, *args, **kwds)

# these three commands were moved out of UI.MainFrame's init
        self.main_list = ProjectReportList(self.Main_Panel)
        self.Report = self.main_list
#        self.main_list = ProjectReportList(self)
#       self.Report = self.main_list
        self.ReportID = reportid
        Data.OpenReports[reportid] = self  # used to refresh report
        Data.Report[reportid]['Open'] = True  # needed by close routine
        Menu.AdjustMenus(self)
        # self. main_list.Reset()  # reset doesn't exist -- I don't remember where I got this line from
        self.set_properties()
        self.do_layout()

# ----- Menu and Toolbars

        # Associate each menu/toolbar item with the method that handles that
        # item.
        if mac:  # mac only
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
#        wx.EVT_MENU(self, wx.ID_REVERT, self.doRevert)
        wx.EVT_MENU(self, wx.ID_EXIT, self.doExit)
        if not mac:  # not mac
            wx.EVT_MENU(self, wx.ID_EXIT,   self.doExit)

        wx.EVT_MENU(self, ID.UNDO,          self.doUndo)
        wx.EVT_MENU(self, ID.REDO,          self.doRedo)

        wx.EVT_MENU(self, ID.FIND_SCRIPTS, self.doFindScripts)
        Menu.doAddScripts(self)
        wx.EVT_MENU_RANGE(self, Menu.FirstScriptID, Menu.FirstScriptID + 1000, self.doScript)

        wx.EVT_MENU(self, ID.ABOUT, self.doShowAbout)
        wx.EVT_MENU(self, ID.HOME_PAGE, self.doHome)
        wx.EVT_MENU(self, ID.HELP_PAGE, self.doHelp)
        wx.EVT_MENU(self, ID.FORUM, self.doForum)

        # Install our own method to handle closing the window.  This allows us
        # to ask the user if he/she wants to save before closing the window, as
        # well as keeping track of which windows are currently open.

        wx.EVT_CLOSE(self, self.doClose)
        wx.EVT_SIZE(self, self.OnSize)
        wx.EVT_MOVE(self, self.OnMove)

        # tool bar events
        wx.EVT_TOOL(self, ID.NEW_PROJECT, self.OnNewProject)
        wx.EVT_TOOL(self, ID.NEW_REPORT, self.OnNewReport)
        wx.EVT_TOOL(self, ID.EDIT, self.OnEdit)
        wx.EVT_TOOL(self, ID.DUPLICATE, self.OnDuplicate)
        wx.EVT_TOOL(self, ID.DELETE, self.OnDelete)
        wx.EVT_TOOL(self, ID.SHOW_HIDDEN_REPORT, self.OnShowHidden)

# ------ Tool Bar Commands ---------

    def OnNewProject(self, event):
        change = { 'Table': 'Project', 'Name': 'New Project' }  # new project because not ID specified
        Data.Update(change)
        Data.SetUndo('New Project')

    def OnNewReport(self, event):
        if debug: print "Start OnNewReport"
        # add report to which project?
        if self.Report.currentItem == None: 
            if debug: print "self.Report.currentItem == None"
            return
        rr = self.Report.reportrow[self.Report.rows[self.Report.currentItem]]  # item -> report row id -> report row record
        rowtable = rr['TableName']
        rowid = rr['TableID']
        if debug: print "rowtable", rowtable, "rowid", rowid
        if rowtable == 'Project':
            pid = rowid
        elif rowtable == 'Report':
            pid = self.Report.data['Report'][rowid].get('ProjectID')
        if not pid: 
            if debug: print "couldn't find program id"
            return  # shouldn't happen
        change = { 'Table': 'Report', 'ProjectID': pid  }
        if pid == 1:  # reports for all projects
            r2 = Data.Report[2]  # report 2 defines the sequence of the report type selection list
            menuid = []  # list of types to be displayed for selection
            k = r2.get('FirstRow', 0)
            loopcheck = 0
            while k != 0 and k != None:
                rr = Data.ReportRow[k]
                hidden = rr.get('Hidden', False)
                table = rr.get('TableName')  # should always be 'ReportType' or 'ColumnType'
                id = rr.get('TableID')
                if (not hidden) and table == 'ReportType' and id:
                    active = Data.ReportType[id].get('zzStatus', 'active') == 'active'
                    allproj = (Data.ReportType[id].get('AllOrEach') or 'all') in ['all', 'both']
                    if active and allproj :
                        menuid.append( id )
                k = rr.get('NextRow', 0)
                loopcheck += 1
                if loopcheck > 100000:  
                    if debug: print "New Report: report rows loop"
                    break  # prevent endless loop if data is corrupted
            menutext = [ (Data.ReportType[x].get('Label') or Data.ReportType[x].get('Name')) for x in menuid ]
            if debug: print menuid, menutext
            dlg = wx.SingleChoiceDialog(self, "Select report to add:", "New Report", menutext)
            if (dlg.ShowModal() != wx.ID_OK): return
            selection = dlg.GetSelection()
            rtid = menuid[selection]
        else:
            r2 = Data.Report[2]  # report 2 defines the sequence of the report type selection list
            menuid = []  # list of types to be displayed for selection
            k = r2.get('FirstRow', 0)
            loopcheck = 0
            while k != 0 and k != None:
                rr = Data.ReportRow[k]
                hidden = rr.get('Hidden', False)
                table = rr.get('TableName')  # should always be 'ReportType' or 'ColumnType'
                id = rr.get('TableID')
                if (not hidden) and table == 'ReportType' and id:
                    active = Data.ReportType[id].get('zzStatus', 'active') == 'active'
                    eachproj = (Data.ReportType[id].get('AllOrEach') or 'all') in ['each', 'both']
                    if active and (eachproj or Data.ReportType[id].get('TableA') == "Task"):
                        menuid.append( id )
                k = rr.get('NextRow', 0)
                loopcheck += 1
                if loopcheck > 100000:  
                    if debug: print "New Report: report rows loop"
                    break  # prevent endless loop if data is corrupted
            menutext = [ (Data.ReportType[x].get('Label') or Data.ReportType[x].get('Name')) for x in menuid ]
            if debug: print menuid, menutext
            dlg = wx.SingleChoiceDialog(self, "Select report to add:", "New Report", menutext)
            if (dlg.ShowModal() != wx.ID_OK): return
            selection = dlg.GetSelection()
            rtid = menuid[selection]
            change['SelectColumn'] = 'ProjectID'
            change['SelectValue'] = pid
        change['ReportTypeID'] = rtid
        change['Name'] = Data.ReportType[rtid].get('Label') or Data.ReportType[rtid].get('Name')  
        Data.Update(change)
        Data.SetUndo('New Report')
        if debug: print "End OnNewReport"

# use this logic when I open a new report???
#        global _docList
#        newFrame = DrawingFrame(None, -1, "Untitled")
#        newFrame.Show(True)
#        _docList.append(newFrame)

# not implemented yet
    def OnEdit(self, event):
        if debug: print 'Start OnEdit'
        rr = self.reportrow[self.rows[self.Report.currentItem]]
        rowtable = rr['TableName']
        rowid = rr['TableID']
        if debug: print 'rowtable:', rowtable
        if debug: print 'rowid:', rowid
        if rowtable == 'Project':
            # if project open project edit window
            dlg = wx.TextEntryDialog(frame, 'New project name', 'Edit Project Name', 'asdf')
            dlg.SetValue("xxxxxxx")
            if dlg.ShowModal() == wxID_OK:
                log.WriteText('You entered: %s\n' % dlg.GetValue())
            dlg.Destroy()

        # change = { 'Table': 'Report', 'Name': 'New Report', 'FirstColumn': id }
        # Data.Update(change)
        # Data.SetUndo('Edit')

        elif rowtable == 'Report':
            # if project open project edit window
            dlg = wx.TextEntryDialog(frame, 'New report name', 'Edit Report Name', 'asdf')
            dlg.SetValue("xxxxx")
            if dlg.ShowModal() == wxID_OK:
                log.WriteText('You entered: %s\n' % dlg.GetValue())
            dlg.Destroy()

        if debug: print 'End OnEdit'

    def OnDuplicate(self, event):
        if debug: print "Start OnDuplicate"
        if self.Report.currentItem == None: 
            if debug: print "self.Report.currentItem == None"
            return
        rr = self.Report.reportrow[self.Report.rows[self.Report.currentItem]]  # item -> report row id -> report row record
        rowtable = rr['TableName']
        rowid = rr['TableID']
        if debug: print "rowtable", rowtable, "rowid", rowid
        if rowtable == 'Project' and rowid:  # duplicate project
            if rowid == 1: return
            # copy project
            new = Data.Project[rowid].copy()
            del new['ID']
            new['Table'] = 'Project'
            undo = Data.Update(new)
            pid = undo['ID']
            # copy tasks
            for k, v in Data.Task.items():
                v = Data.Task[k]
                if v['ProjectID'] != rowid: continue
                new = v.copy()
                del new['ID']
                new['Table'] = 'Task'
                new['ProjectID'] = pid
                Data.Update(new)
            Data.SetUndo('Duplicate Project')

        elif rowtable == 'Report' and rowid:  # duplicate report
            # copy report
            rep = Data.Report[rowid].copy()
            del rep['ID']
            rep['Table'] = 'Report'
            undo = Data.Update(rep)
            rid = undo['ID']

            # copy report columns
            c = rep.get('FirstColumn')
            rc = Data.ReportColumn
            cols = []
            while c:
                new = rc[c].copy()
                del new['ID']
                new['Table'] = 'ReportColumn'
                new['ReportID'] = rid
                cols.append(new)
                c = new.get('NextColumn')
            while len(cols):
                new = cols.pop()
                new['NextColumn'] = c
                undo = Data.Update(new)
                c = undo['ID']
            rep['FirstColumn'] = c

            # copy report rows  -- do I need to do this? won't it automatically create these for me?
            r = rep.get('FirstRow')
            rr = Data.ReportRow
            rows = []
            while r:
                new = rr[r].copy()
                del new['ID']
                new['Table'] = 'ReportRow'
                new['ReportID'] = rid
                rows.append(new)
                r = new.get('NextRow')
            while len(rows):
                new = rows.pop()
                new['NextRow'] = r
                undo = Data.Update(new)
                r = undo['ID']
            rep['FirstRow'] = r

# turn these into a routine?
# def dupchain(parent, parentname, child, childname, first, next)  
#       parent is a record, child is a table, names are table names, first & next are columns names

            rep['ID'] = rid
            Data.Update(rep)
            Data.SetUndo('Duplicate Report')
        if debug: print "End OnDuplicate"
        
    def OnDelete(self, event):
        if debug: print "Start OnDelete"
        if self.Report.currentItem == None: 
            if debug: print "self.Report.currentItem == None"
            return
        rr = self.Report.reportrow[self.Report.rows[self.Report.currentItem]]  # item -> report row id -> report row record
        rowtable = rr['TableName']
        rowid = rr['TableID']
        if debug: print "rowtable", rowtable, "rowid", rowid
        if rowtable == 'Project' and rowid:  # delete project
            if rowid == 1: return  # don't allow deletion of project 1
            # check to see if alrady deleted
            if Data.Database['Project'][rowid].get('zzStatus', 'active') == 'deleted':  # only reactivate project, not tasks
                change = {'Table':'Project', 'ID': rowid, 'zzStatus': 'active'}
                Data.Update(change)
                Data.SetUndo('Reactivate Project')
                return
            # delete project
            change = {'Table':'Project', 'ID': rowid, 'zzStatus': 'deleted'}
            Data.Update(change)
            # delete tasks
            change = {'Table':'Task', 'ID': None, 'zzStatus': 'deleted'}
            for k, v in Data.Task.iteritems():
                if v['ProjectID'] != rowid: continue
                change['ID'] = k
                Data.Update(change)
            Data.SetUndo('Delete Project')

        elif rowtable == 'Report' and rowid:  # duplicate report
            if rowid == 1 or rowid == 2: return  # don't allow deletion of report 1 or 2
            # check to see if already deleted
            if Data.Database['Report'][rowid].get('zzStatus', 'active') == 'deleted':
                newstatus = 'active'
                message = 'Reactivate Report'
            else:
                newstatus = 'deleted'
                message = 'Delete Report'
            # delete report
            change = {'Table':'Report', 'ID': rowid, 'zzStatus': newstatus}
            Data.Update(change)
            # delete report columns
            change = {'Table':'ReportColumn', 'ID': None, 'zzStatus': newstatus}
            for k, v in Data.ReportColumn.iteritems():
                if debug: print "v", v
                if v['ReportID'] != rowid: continue
                change['ID'] = k
                Data.Update(change)
            # delete report rows
            change = {'Table':'ReportRow', 'ID': None, 'zzStatus': newstatus}
            for k, v in Data.ReportRow.iteritems():
                if debug: print "v", v
                if v['ReportID'] != rowid: continue
                change['ID'] = k
                Data.Update(change)
            Data.SetUndo(message)
        if debug: print "End OnDelete"

    def OnHide(self, event):
        if self.Report.currentItem == None: 
            if debug: print "self.Report.currentItem == None"
            return
        Menu.onHide(self.Report, event, [self.Report.currentItem])

    def OnShowHidden(self, event):
        Menu.onShowHidden(self, event)

    # ---- Menu Command -----

    # File Menu
    def doNew(self, event):
        Menu.doNew(self, event)

    def doOpen(self, event):
        Menu.doOpen(self, event)

    def doClose(self, event):
        Menu.doExit(self, event)
        # if Data.ChangedData:
        #     if not Data.AskIfUserWantsToSave(self, "closing"): return
        # 
        # del Data.OpenReports[self.ReportID]
        # self.Destroy()

    def doSave(self, event):
        Menu.doSave(self, event)

    def doSaveAs(self, event):
        Menu.doSaveAs(self, event)

    def doExit(self, event):
        Menu.doExit(self, event)
        if debug: print "doExit didn't exit"
        GanttPV.ExitMainLoop()  # just in case the doExit fails


    # Edit Menu
    def doUndo(self, event):
        Menu.doUndo(self, event)

    def doRedo(self, event):
        Menu.doRedo(self, event)

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
        if debug: print size
        r = Data.Database['Report'][self.ReportID]
        r['FrameSizeW'] = size.width
        r['FrameSizeH'] = size.height

        event.Skip()  # to call default handler; needed?

    def OnMove(self, event):
        pos = event.GetPosition()
        # if debug: print pos
        r = Data.Database['Report'][self.ReportID]
        r['FramePositionX'] = pos.x
        r['FramePositionY'] = pos.y

        event.Skip()  # needed?

    def UpdatePointers(self, all=0):  # 1 = new database; 0 = changed report rows or columns
        if all:
            self.Report.UpdateDataPointers()
        self.Report.UpdateColumnPointers()
        self.Report.UpdateRowPointers()

# -------------

class GanttPVApp(wx.App):
    def OnInit(self):
        wx.InitAllImageHandlers()
        # first parameter is the report id. evidently my new parameters have to be added before the others

        # these are the positional parameters for wx.Frame:
        # None = parent
        # -1 = id
        # "" = title
        # wx.Point(100,100) = where to position window
        # wx.Size(160,100) = how large to make window

        Main = ProjectReportFrame(1, None, -1, "")  # reportid = 1
        # Main = ProjectReportFrame(-1, "")
        self.SetTopWindow(Main)
        Main.Show(1)
        return 1

# end of class GanttPVApp

if debug: print "Test to start main loop"

if __name__ == "__main__":
    if debug: print "Start main loop ------------------------- "
    if debug: print "sys.argv", sys.argv
    # ex_name = 'myprogram.exe' 
    # if mac:
    #    ex_name = 'GanttPV.app' 
    # else:
    #    ex_name = 'GanttPV.exe' 
    # if debug: print "ex_name", ex_name
    # 
    # if sys.path[0][-len(ex_name):] == ex_name: 
    #    path = os.path.split(sys.path[0]) 
    # else: 
    if mac:
        path = os.path.split(__file__) 
        Data.Path = path[0]  # directory from which program was executed
    else:
        path = os.path.abspath(sys.argv[0])
        dirname, basename = os.path.split(path)
        Data.Path = os.path.dirname(dirname)  # back up one level
        # path = (os.getcwd(), 'GanttPV.exe')
    if debug: print "path", path
    # Data.Path = path[0]  #~~ directory from which program was executed
    Data.LoadOption(Data.Path)

    GanttPV = GanttPVApp(0)
    # if len(sys.argv) == 3:
    #     Data.OpenFile(sys.argv[2])

    GanttPV.MainLoop()
    if debug: print "End main loop ---------------------- "

# if __name__ == '__main__':
#    app = wx.PySimpleApp()
#    frame = ProjectReportList(app)
#    frame.Show(True)
#    app.MainLoop()
if debug: print "end Main.py"

