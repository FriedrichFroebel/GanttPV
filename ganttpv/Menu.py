#!/usr/bin/env python
# Implement shared menu commands

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

# 040417 - first version of this program; moved AdjustMenus here from Data
# 040419 - Setup Undo and Redo to use the commands from Data
# 040420 - moved doClose back to individual reports; drafted toolbar adjustment for Main
# 040424 - added scripts menu processing
# 040427 - onHide and onShowHidden moved from Main
# 040503 - added Undo message text to Undo and Redo menus
# 040510 - fixed doOpen & doExit; changed AdjustMenus to adjust GanttReport toolbar
# 040517 - added Bitmap
# 040518 - changed doScript so that it will look in the default location if ScriptDirectory is not defined; added doHome
# 040519 - show Assign Resources and Assign Prerequisite only if single task selected; revise logic to enable Main's buttons
# 040520 - Added final slash to Help and Forum URL's
# 040605 - changed URL directory names to lower case
# 040715 - Pierre_Rouleau@impathnetworks.com: removed all tabs, now use 4-space indentation level to comply with Official Python Guideline.

import wx, os, webbrowser
import Data, ID, UI
# import Main, GanttReport
# import inspect

debug = 1
if debug: print "load Menu.py"

# --------- File Menu ------------

def doNew(self, event):
    """ Respond to the "New" menu command.
    """
    if debug: print "Start doNew"
    if Data.AskIfUserWantsToSave(self, "creating new project file"):
        if debug: print "Setup empty database"
        Data.CloseReports()# close all open reports except #1
        Data.SetEmptyData()
        Data.MakeReady()
    if debug: print "End doNew"

def doOpen(self, event):
    """ Respond to the "Open" menu command.
    """
    if not Data.AskIfUserWantsToSave(self, "open project file"): return

    curDir = os.getcwd()
    fileName = wx.FileSelector("Open File", default_extension=".ganttpv",
                               flags = wx.OPEN | wx.FILE_MUST_EXIST)
    if fileName == "": return
    fileName = os.path.join(os.getcwd(), fileName)
    os.chdir(curDir)

    Data.CloseReports()         # close all open reports except #1
    Data.OpenFile(fileName)

# I'm going to leave the close logic in each frame for the present
# def doClose(self, event):  # what should I close? the report? all reports? 
#     """ Respond to the "Close" menu command.
#    """
#    if Data.ChangedData:
#        if not Data.AskIfUserWantsToSave(self, "closing"): return
#
        # del OpenReports(self.reportid)
        # self.Destroy()

def doSave(self, event):
    """ Respond to the "Save" menu command.
    """
    if Data.FileName != None:
        Data.SaveContents(self)

def doSaveAs(self, event):
    """ Respond to the "Save As" menu command.
    """
    # global Data.FileName
    if Data.FileName == None:
        default = ""
    else:
        default = Data.FileName

    curDir = os.getcwd()
    fileName = wx.FileSelector("Save File As", "Saving",
                              default_filename='Untitled.ganttpv',
                              default_extension="ganttpv",
                              wildcard="*.ganttpv",
                              flags = wx.SAVE | wx.OVERWRITE_PROMPT)
    if fileName == "": return # User cancelled.
    fileName = os.path.join(os.getcwd(), fileName)
    os.chdir(curDir)

    title = os.path.basename(fileName)
    self.SetTitle(title)

    Data.FileName = fileName
    Data.SaveContents(self)

def doExit(self, event):
    """ Respond to the "Quit" menu command.
    """
    if not Data.AskIfUserWantsToSave(self, "exit GanttPV"): return

    if debug: print "Continue doExit"
    Data.CloseReports()
    Data.CloseReport(1)
    # Main.GanttPV.ExitMainLoop()  # needed for windows??
#    _app.ExitMainLoop()
    if debug: print "End doExit"

# not used yet
def doRevert(self, event):
    """ Respond to the "Revert" menu command.
    """
    if not Data.ChangedData: return

    if wxMessageBox("Discard changes made to this file?", "Confirm",
                    style = wx.OK | wx.CANCEL | wx.ICON_QUESTION,
                    parent=self) == wx.CANCEL: return
    Data.LoadContents()

# ---------- Edit Menu -------

def doUndo(self, event):
    """ Respond to the "Undo" menu command.
    """
    Data.DoUndo()

def doRedo(self, event):
    """ Respond to the "Undo" menu command.
    """
    Data.DoRedo()

# ------------ Script Menu -----------------

# look for the scripts directory at Data.Option['ScriptDirectory']

Script = []  # array of script file names
FirstScriptID = 10000
# menuOffset = 2  # required if I insert into the menu

def doAddScripts(self):  # report should use this it initialize scripts menu
    if len(Script) > 0:
        mb = self.GetMenuBar()
        menuItem = mb.FindItemById(ID.FIND_SCRIPTS)
        menu = menuItem.GetMenu()
        i = 0
        while i < len(Script):
            menu.Append(FirstScriptID + i, Script[i][:-3])  # omit ".py" from script names
            i += 1

        #menu.Insert(9, ID, "NewItem " + str(ID))
        #item = wxMenuItem(menu)
        #item.SetId(ID)
        #item.SetText("NewItem " + str(ID))
        #menu.InsertItem(9, item)

def GetScriptNames():
    global Script
    sd = Data.Option.get('ScriptDirectory') or os.path.join(Data.Path, "Scripts")
    try:
        if debug: print "looking for script names in " + sd
        s = os.listdir(sd)
        Script = [ x for x in s if x[-3:] == ".py" ]
        if debug: print 'these are the script names I found: ',  Script
    except OSError:
        if debug: print "doFindScripts os.listdir OSError"

def doFindScripts(self, event):
    global Script     # Data.ScriptDirector
    if debug: print "Start doFindScripts"

    for v in Data.OpenReports.values():    # update menus for all open reports
        if len(Script) > 0:  # Remove old scripts from menu
            mb = v.GetMenuBar()
            menuItem = mb.FindItemById(ID.FIND_SCRIPTS)
            menu = menuItem.GetMenu()
            i = 0
            while i < len(Script):
                menu.Remove(FirstScriptID + i)
                i += 1

    GetScriptNames()

    if len(Script) == 0: 

        curDir = os.getcwd()  # save current directory
        dlg = wx.DirDialog(self, "Choose Script directory",
                                style = wx.DD_DEFAULT_STYLE|wx.DD_NEW_DIR_BUTTON)
        if dlg.ShowModal() == wx.ID_OK:
            sd = dlg.GetPath()
            if Data.Option['ScriptDirectory'] != sd:
                Data.Option['ScriptDirectory'] = sd
                Data.SaveOption()
            GetScriptNames()  # even if directory is same, contents may have changed

        os.chdir(curDir)

    for v in Data.OpenReports.values():    # update menus for all open reports
        doAddScripts(v)

    if debug: print "End doFindScripts"

def doScript(self, event):
    if debug: print "Start doScript"
    id = event.GetId()
    sd = Data.Option.get('ScriptDirectory') or os.path.join(Data.Path, "Scripts")
    script = os.path.join(sd, Script[id - FirstScriptID])
    if debug: print "script file name", script
    try:
        execfile(script)
    except IOError:
        pass
    if debug: print "End doScript"

# ------------ Help Menu -----------------

def doHome(self, event):
    if debug: print "start doHome"
    webbrowser.open_new('http://www.PureViolet.net/ganttpv/')
    if debug: print "end doHome"

def doHelp(self, event):
    webbrowser.open_new('http://www.PureViolet.net/ganttpv/help/')

def doForum(self, event):
    webbrowser.open_new('http://www.SimpleProjectManagement.com/forum/')

def doShowAbout(self, event):
    """ Respond to the "About GanttPV" menu command.
    """

    # dialog = wx.Dialog(self, -1, "About GanttPV") # ,
    #                   #style=wx.DIALOG_MODAL | wx.STAY_ON_TOP)
    # dialog.SetBackgroundColour(wx.WHITE)
    # 
    # panel = wx.Panel(dialog, -1)
    # panel.SetBackgroundColour(wx.WHITE)
    # 
    # panelSizer = wx.BoxSizer(wx.VERTICAL)
    # 
    # boldFont = wx.Font(panel.GetFont().GetPointSize(),
    #                   panel.GetFont().GetFamily(),
    #                   wx.NORMAL, wx.BOLD)
    # 
    # logo = wx.StaticBitmap(panel, -1, wx.Bitmap("images/logo.bmp",
    #                                               wx.BITMAP_TYPE_BMP))
    # 
    # lab1 = wx.StaticText(panel, -1, "GanttPV")
    # lab1.SetFont(wx.Font(36, boldFont.GetFamily(), wx.ITALIC, wx.BOLD))
    # lab1.SetSize(lab1.GetBestSize())
    # 
    # imageSizer = wx.BoxSizer(wx.HORIZONTAL)
    # imageSizer.Add(logo, 0, wx.ALL | wx.ALIGN_CENTRE_VERTICAL, 5)
    # imageSizer.Add(lab1, 0, wx.ALL | wx.ALIGN_CENTRE_VERTICAL, 5)
    # 
    # lab2 = wx.StaticText(panel, -1, "A cross-platform, open source project schedulingented " + \
    #                                "program.")
    # lab2.SetFont(boldFont)
    # lab2.SetSize(lab2.GetBestSize())
    # 
    # lab3 = wx.StaticText(panel, -1, "GanttPV is release under the GPL" + \
    #                                "; please")
    # lab3.SetFont(boldFont)
    # lab3.SetSize(lab3.GetBestSize())
    # 
    # lab4 = wx.StaticText(panel, -1, "enjoy the user of this program. " + \
    #                                "Support it in any way you like.")
    # lab4.SetFont(boldFont)
    # lab4.SetSize(lab4.GetBestSize())
    # 
    # lab5 = wx.StaticText(panel, -1, "Author: Brian Christensen " + \
    #                                    "(brian@SimpleProjectManagement.com)")
    # lab5.SetFont(boldFont)
    # lab5.SetSize(lab5.GetBestSize())
    # 
    # tnOK = wx.Button(panel, wx.ID_OK, "OK")
    # 
    # panelSizer.Add(imageSizer, 0, wx.ALIGN_CENTRE)
    # panelSizer.Add(10, 10) # Spacer.
    # panelSizer.Add(lab2, 0, wx.ALIGN_CENTRE)
    # panelSizer.Add(10, 10) # Spacer.
    # panelSizer.Add(lab3, 0, wx.ALIGN_CENTRE)
    # panelSizer.Add(lab4, 0, wx.ALIGN_CENTRE)
    # panelSizer.Add(10, 10) # Spacer.
    # panelSizer.Add(lab5, 0, wx.ALIGN_CENTRE)
    # panelSizer.Add(10, 10) # Spacer.
    # panelSizer.Add(btnOK, 0, wx.ALL | wx.ALIGN_CENTRE, 5)
    # 
    # panel.SetAutoLayout(True)
    # panel.SetSizer(panelSizer)
    # panelSizer.Fit(panel)
    # 
    # topSizer = wx.BoxSizer(wx.HORIZONTAL)
    # topSizer.Add(panel, 0, wx.ALL, 10)
    # 
    # dialog.SetAutoLayout(True)
    # dialog.SetSizer(topSizer)
    # topSizer.Fit(dialog)

    dialog = UI.AboutDialog(self, -1, "About GanttPV") # ,
                      #style=wx.DIALOG_MODAL | wx.STAY_ON_TOP)
    dialog.Centre()

    btn = dialog.ShowModal()
    dialog.Destroy()

# -- adjust menus

def AdjustMenus(self):
    """ Adjust menus and toolbar to reflect the current state. """
    # if debug: print "Adjusting Menus"
    # if debug: print 'self', self
    canSave   = (Data.FileName != None) and Data.ChangedData
    canRevert = (Data.FileName != None) and Data.ChangedData
    canUndo   = len(Data.UndoStack) > 0
    canRedo   = len(Data.RedoStack) > 0

    # selection = len(self.selection) > 0
    # onlyOne   = len(self.selection) == 1
    # isText    = onlyOne and (self.selection[0].getType() == obj_TEXT)
    # front     = onlyOne and (self.selection[0] == self.contents[0])
    # back      = onlyOne and (self.selection[0] == self.contents[-1])

    # Enable/disable our menu items.

    self.FileMenu.Enable(wx.ID_SAVE,   canSave)
    # self.FileMenu.Enable(wx.ID_REVERT, canRevert)

    self.Edit.Enable(ID.UNDO,      canUndo)
    self.Edit.Enable(ID.REDO,      canRedo)
    mbar = self.GetMenuBar()
    undoItem = mbar.FindItemById(ID.UNDO)
    redoItem = mbar.FindItemById(ID.REDO)
    # if debug and canUndo: print 'Data.UndoStack[-1]', Data.UndoStack[-1]
    # if canUndo: undoItem.SetText('Undo ' + Data.UndoStack[-1])
        # problem - shouldn't have to test for string, but somehow GanttReport onSelect is called while the top of the 
        #       stack is not a string
    if canUndo and isinstance(Data.UndoStack[-1], str): undoItem.SetText('Undo ' + Data.UndoStack[-1])
    else: undoItem.SetText('Undo')
    if canRedo: redoItem.SetText('Redo ' + Data.RedoStack[-1])
    else: redoItem.SetText('Redo')
        
    # self.editMenu.Enable(menu_DUPLICATE, selection)
    # self.editMenu.Enable(menu_EDIT_TEXT, isText)
    # self.editMenu.Enable(menu_DELETE,    selection)

    # self.toolsMenu.Check(menu_SELECT,  self.x == self.y)

    # Enable/disable our toolbar icons.
    if isinstance(self, UI.MainFrame):
        item = self.Report.currentItem
        if debug: "item", item
        isSel = item != None
        if debug: "isSel", isSel
        selReport = False; selProject = False
        if isSel:
            rr = self.Report.reportrow[self.Report.rows[item]]
            selReport = rr['TableName'] == 'Report'
            selProject = rr['TableName'] == 'Project'
        
        self.main_toolbar.EnableTool(ID.NEW_PROJECT, True)
        self.main_toolbar.EnableTool(ID.NEW_REPORT, isSel)
        self.main_toolbar.EnableTool(ID.EDIT, isSel)
        self.main_toolbar.EnableTool(ID.DUPLICATE, isSel)
        self.main_toolbar.EnableTool(ID.DELETE, isSel)
        self.main_toolbar.EnableTool(ID.SHOW_HIDDEN_REPORT, True)  # need to adjust toggle based on report's show hidden flag
        #-------------- the next line needs to be uncommented
        # self.main_toolbar.Check(ID.SHOW_HIDDEN_REPORT, self.report.get('ShowHidden', False))

    elif isinstance(self, UI.ReportFrame):
        r = Data.Report[self.ReportID]
        rtid = r['ReportTypeID']
        rt = Data.ReportType[rtid]
        ta = rt.get('TableA')
        toolbar = self.report_toolbar
        isTask = ta == "Task"
        csel = self.Report.GetSelectedCols()  # current selection?
        rsel = self.Report.GetSelectedRows()  # current selection?
        isSelCol = len(csel) > 0
        isSelRow = len(rsel) > 0
        # self.report_toolbar.EnableTool(ID.INSERT_ROW, True)
        self.report_toolbar.EnableTool(ID.DUPLICATE_ROW, isSelRow and isTask)
        self.report_toolbar.EnableTool(ID.DELETE_ROW, isSelRow)
        self.report_toolbar.EnableTool(ID.MOVE_UP, isSelRow)
        self.report_toolbar.EnableTool(ID.MOVE_DOWN, isSelRow)

        self.report_toolbar.EnableTool(ID.PREREQUISITE, isTask and len(rsel) == 1)
        self.report_toolbar.EnableTool(ID.ASSIGN_RESOURCE, isTask and len(rsel) == 1)

        self.report_toolbar.EnableTool(ID.HIDE_ROW, isSelRow)
        self.report_toolbar.ToggleTool(ID.SHOW_HIDDEN, r.get('ShowHidden', False))

        # self.report_toolbar.EnableTool(ID.INSERT_COLUMN, True)
        self.report_toolbar.EnableTool(ID.DELETE_COLUMN, isSelCol)
        self.report_toolbar.EnableTool(ID.MOVE_LEFT, isSelCol)
        self.report_toolbar.EnableTool(ID.MOVE_RIGHT, isSelCol)
        self.report_toolbar.EnableTool(ID.COLUMN_OPTIONS, False)
        # self.report_toolbar.EnableTool(ID.SCROLL_LEFT_FAR, True)
        # self.report_toolbar.EnableTool(ID.SCROLL_LEFT, True)
        # self.report_toolbar.EnableTool(ID.SCROLL_RIGHT, True)
        # self.report_toolbar.EnableTool(ID.SCROLL_RIGHT_FAR, True)
        self.report_toolbar.EnableTool(ID.SCROLL_TO_TASK, isSelRow)

        #control = tb.FindControl(controlid)
        #flag = tb.GetToolEnabled(toolid)
        # flag = tb.GetToolState(toolid)  # whether toggled on or off
        # tb.ToggleTool(toolid, toggleflag)  # flag sets as on or off


    # if debug: print 'Finished adjusting menus'

# ----------- common tool buttons -------

def onHide(self, event, sel):  # self is the object where 'rows' is stored
    """ Toggle Hide flag on report row """
    if debug: print "Start OnHide"
    for x in sel:   
        rrid = self.rows[x]  # item -> report row id
        rr = self.reportrow[rrid]  # report row id -> report row record
        oldval = rr.get('Hidden', False)
        if debug: print "old hide value", oldval
        change = { 'Table': 'ReportRow', 'ID': rrid, 'Hidden': not oldval }
        Data.Update(change)
        if debug: print "change", change
    Data.SetUndo('Hide Report Row')
    if debug: print "End OnHide"

def onShowHidden(self, event):  # self is the frame
    """ Toggle ShowHidden flag on report """
    if debug: print "Start OnShowHidden"
    rep = Data.Report[self.ReportID]
    change = {'Table':'Report', 'ID': self.ReportID, 'ShowHidden': not (rep.get('ShowHidden', False))}
    # if debug: print "change", change
    Data.Update(change)
    Data.SetUndo('Show Hidden Rows')
    if debug: print "End OnShowHidden"

def Bitmap(file, flags):  # used to replace code that reads bitmaps in UI.py
    parts = file.split('/') # separate out the file name
    path = os.path.join(Data.Path, "icons", parts[-1])
    # if debug: print 'path', path
    bm = wx.Bitmap(path, flags)
    # if not isinstance(bm, wx.Bitmap): bm = wx.NullBitmap
    return bm
    # return wx.Bitmap(os.path.join(Data.Path, "icons", "Report.bmp"), wx.BITMAP_TYPE_BMP)

if debug: print "end Menu.py"
