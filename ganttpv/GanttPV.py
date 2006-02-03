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
# 040420 - change to show 'About' under the application menu instead of the 'Help' menu; double click on report name to open or show report; close will close this report; will open reports where they were before
# 040421 - fixed New Report button; added Duplicate button logic; added Delete button logic
# 040422 - added reactivate logic to Delete button; added ShowHidden logic; add logic to change  colors for different row types
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
# 050202 - remove line feed from insert report menu text
# 050409 - Alexander - added command-line support for opening .ganttpv files and .py scripts.
# 050503 - Alexander - moved script-running logic to Data.py
# 050504 - Alexander - implemented Window menu; moved some menu event-handling logic to Menu.py
# 050515 - Alexander - added MacOpenFile
# 060131 - Alexander - when the selection is invalid (index >= number of rows), treat it as empty

import wx, wx.grid
# from wxPython.lib.dialogs import wxMultipleChoiceDialog  # use single selection for adding new reports
# import webbrowser
import os, os.path, sys
import re
# import images
import Data, UI, ID, Menu, GanttReport

debug = 1
mac = 1
platform = sys.platform

if debug: print "load GanttPV.py"

#----------------------------------------------------------------------

class ProjectReportList(wx.ListCtrl):
    def __init__(self, parent):
        wx.ListCtrl.__init__(self, parent, -1, style=wx.LC_REPORT|wx.LC_VIRTUAL|wx.LC_HRULES|wx.LC_VRULES|wx.LC_SINGLE_SEL)

        self.currentItem = None

        self.images = wx.ImageList(16, 16)
        path = os.path.join(Data.Path, "icons", "Report.bmp")
        self.idx1 = self.images.Add(wx.Bitmap(path, wx.BITMAP_TYPE_BMP))
        self.SetImageList(self.images, wx.IMAGE_LIST_SMALL)

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

        wx.EVT_LIST_ITEM_ACTIVATED(self, self.GetId(), self.OnItemActivated)
        wx.EVT_LIST_ITEM_DESELECTED(self, self.GetId(), self.OnItemDeselected)
        wx.EVT_LIST_ITEM_SELECTED(self, self.GetId(), self.OnItemSelected)

    def OnItemSelected(self, event):
        self.currentItem = event.GetIndex()
        Menu.AdjustMenus(Data.OpenReports[1])

    def OnItemActivated(self, event):  # double clicked or "ENTER" while selected
        item = event.GetIndex()
        self.currentItem = item

        rr = self.reportrow[self.rows[item]]
        rowtable = rr['TableName']
        rowid = rr['TableID']

        if debug: print 'rowid:', rowid
        if rowtable == 'Project': pass
                # if project open project edit window

        elif rowtable == 'Report':
            Data.OpenReport(rowid)  # row id == report id

    def getColumnText(self, index, col):
        item = self.GetItem(index, col)
        return item.GetText()

    def OnItemDeselected(self, evt):
        self.currentItem = None
        Menu.AdjustMenus(Data.OpenReports[1])

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
        deleted = Data.Database[rowtable][rowid].get('zzStatus') == 'deleted'
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

    def UpdateColumnPointers(self):
        rc = self.reportcolumn  # pointer to table

        self.columns = []
        self.ctypes = []
        self.coloffset = []

        for colid in Data.GetColumnList(1):
            r = rc.get(colid) or {}
            ctid = r.get('ColumnTypeID')

            self.columns.append( colid )
            self.ctypes.append( ctid )
            self.coloffset.append( -1 )

    def UpdateRowPointers(self):
        Data.UpdateRowPointers(self)
        self.SetItemCount(len(self.rows))  # number of items in virtual list

#----------------------------------------------------------------------

class ProjectReportFrame(UI.MainFrame):
    def __init__(self, reportid, *args, **kwds):
        if debug: print 'ProjectReportFrame __init__'
        if debug: print 'reportid', reportid
        UI.MainFrame.__init__(self, *args, **kwds)

        # these three commands were moved out of UI.MainFrame's init
        self.main_list = ProjectReportList(self.Main_Panel)
        self.Report = self.main_list
        self.ReportID = reportid

        Data.OpenReports[reportid] = self  # used to refresh report
        Data.Report[reportid]['Open'] = True  # needed by close routine

        Menu.GetScriptNames()
        Menu.doAddScripts(self)
        Menu.FillWindowMenu(self)
        Menu.AdjustMenus(self)

        self.set_properties()
        self.do_layout()

        # file menu events
        wx.EVT_MENU(self, wx.ID_NEW,    Menu.doNew)
        wx.EVT_MENU(self, wx.ID_OPEN,   Menu.doOpen)
        wx.EVT_MENU(self, wx.ID_CLOSE,  Menu.doExit)
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
        wx.EVT_CLOSE(self, Menu.doExit)
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
        sel = self.Report.currentItem
        if (sel == None) or (sel >= len(self.Report.rows)):
            if debug: print "no item selected"
            return
        rr = self.Report.reportrow[self.Report.rows[self.Report.currentItem]]  # item -> report row id -> report row record
        rowtable = rr['TableName']
        rowid = rr['TableID']
        if debug: print "rowtable", rowtable, "rowid", rowid
        if rowtable == 'Project':
            pid = rowid
        elif rowtable == 'Report':
            pid = Data.Database['Report'][rowid].get('ProjectID')
        if not pid: 
            if debug: print "couldn't find project id"
            return  # shouldn't happen
        change = { 'Table': 'Report', 'ProjectID': pid  }

        menuid = []  # list of types to be displayed for selection
        rlist = Data.GetRowList(2)  # report 2 defines the sequence of the report type selection list
        for k in rlist:
            rr = Data.ReportRow[k]
            hidden = rr.get('Hidden', False)
            table = rr.get('TableName')  # should always be 'ReportType' or 'ColumnType'
            id = rr.get('TableID')
            if (not hidden) and table == 'ReportType' and id:
                active = Data.ReportType[id].get('zzStatus') != 'deleted'
                if pid == 1:
                    allproj = (Data.ReportType[id].get('AllOrEach') or 'all') in ['all', 'both']
                    if active and allproj :
                        menuid.append( id )
                else:
                    eachproj = (Data.ReportType[id].get('AllOrEach') or 'all') in ['each', 'both']
                    if active and (eachproj or Data.ReportType[id].get('TableA') == "Task"):
                        menuid.append( id )

        menutext = [ (Data.ReportType[x].get('Label') or Data.ReportType[x].get('Name')) for x in menuid ]
        for i, v in enumerate(menutext):  # remove line feeds before menu display
            if '\n' in v:
                menutext[i] = v.replace('\n', ' ')
        if debug: print menuid, menutext
        dlg = wx.SingleChoiceDialog(self, "Select report to add:", "New Report", menutext)
        dlg.SetSize((240,320))
        dlg.Centre()
        if (dlg.ShowModal() != wx.ID_OK): return
        selection = dlg.GetSelection()
        rtid = menuid[selection]
        if pid != 1:
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
                pass
            dlg.Destroy()

        # change = { 'Table': 'Report', 'Name': 'New Report', 'FirstColumn': id }
        # Data.Update(change)
        # Data.SetUndo('Edit')

        elif rowtable == 'Report':
            # if project open project edit window
            dlg = wx.TextEntryDialog(frame, 'New report name', 'Edit Report Name', 'asdf')
            dlg.SetValue("xxxxx")
            if dlg.ShowModal() == wxID_OK:
                pass
            dlg.Destroy()

        if debug: print 'End OnEdit'

    def OnDuplicate(self, event):
        if debug: print "Start OnDuplicate"
        sel = self.Report.currentItem
        if (sel == None) or (sel >= len(self.Report.rows)):
            if debug: print "no item selected"
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
            clist = Data.GetColumnList(rowid)
            cols = []
            for c in clist:
                new = Data.ReportColumn[c].copy()
                del new['ID']
                new['Table'] = 'ReportColumn'
                new['ReportID'] = rid
                cols.append(new)
            cols.reverse()

            c = None
            for new in cols:
                new['NextColumn'] = c
                undo = Data.Update(new)
                c = undo['ID']
            rep['FirstColumn'] = c

            # copy order of report rows
            rlist = Data.GetRowList(rowid)
            rows = []
            for r in rlist:
                new = Data.ReportRow[r].copy()
                del new['ID']
                new['Table'] = 'ReportRow'
                new['ReportID'] = rid
                rows.append(new)
            rows.reverse()

            r = None
            for new in rows:
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
        sel = self.Report.currentItem
        if (sel == None) or (sel >= len(self.Report.rows)):
            if debug: print "no item selected"
            return
        rr = self.Report.reportrow[self.Report.rows[self.Report.currentItem]]  # item -> report row id -> report row record
        rowtable = rr['TableName']
        rowid = rr['TableID']
        if debug: print "rowtable", rowtable, "rowid", rowid
        if rowtable == 'Project' and rowid:  # delete project
            if rowid == 1: return  # don't allow deletion of project 1
            # check to see if already deleted
            if Data.Database['Project'][rowid].get('zzStatus') == 'deleted':  # only reactivate project, not tasks
                change = {'Table':'Project', 'ID': rowid, 'zzStatus': None}
                Data.Update(change)
                Data.SetUndo('Reactivate Project')
            else:
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
            if Data.Database['Report'][rowid].get('zzStatus') == 'deleted':
                newstatus = None
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
                if v['ReportID'] != rowid: continue
                change['ID'] = k
                Data.Update(change)
            # delete report rows
            change = {'Table':'ReportRow', 'ID': None, 'zzStatus': newstatus}
            for k, v in Data.ReportRow.iteritems():
                if v['ReportID'] != rowid: continue
                change['ID'] = k
                Data.Update(change)
            Data.SetUndo(message)
        if debug: print "End OnDelete"

    def OnHide(self, event):
        sel = self.Report.currentItem
        if (sel == None) or (sel >= len(self.Report.rows)):
            if debug: print "no item selected"
            return
        Menu.onHide(self.Report, event, [self.Report.currentItem])

    def OnShowHidden(self, event):
        Menu.onShowHidden(self, event)

    # ---- menu commands -----

    def doBringWindow(self, event):
        Menu.doBringWindow(self, event)

# ---------------- activate

    def OnActivate(self, event):
        if event.GetActive():
            Data.ActiveReport = self.ReportID
        event.Skip()

# ----------------- window size/position 

    def OnSize(self, event):
        size = event.GetSize()
        # if debug: print size
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

        Main = ProjectReportFrame(1, None)
          # first parameter is the report id
          # second parameter is parent for wx.Frame
        self.SetTopWindow(Main)

        self.ParseOptions()
        Main.Show(True)

        return 1

    def ParseOptions(self):
        """ Process any command line options """
        options = sys.argv
        usageStr = 'usage: ' + options[0] + ' [file...]'
        i = 1
        while i < len(options):
            Data.OpenFile(options[i])
            i += 1

    def MacOpenFile(self, path):
        """ Open a file for Mac OS. """
        Data.OpenFile(path)

# end of class GanttPVApp

if __name__ == "__main__":
    path = os.path.abspath(sys.argv[0])
    Data.Path = os.path.dirname(path)
    if debug: print "data path:", Data.Path

    Data.LoadOption()
    Data.SetEmptyData()
    Data.MakeReady()

    if debug: print "Start main loop ------------------------ "
    Data.App = GanttPVApp(0)
    Data.App.MainLoop()
    if debug: print "End main loop -------------------------- "

if debug: print "end GanttPV.py"

