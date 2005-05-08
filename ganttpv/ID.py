#!/usr/bin/env python
# ID's for menus

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

# 040414 -first version of this file
# 040416 - changes to match revised menus/tool bars in Main & Gantt Report
# 040423 - added Scripts menu
# 040505 - renamed some ID's from 'Task' to 'Row'
# 040715 - Pierre_Rouleau@impathnetworks.com: removed all tabs, now use 4-space indentation level to comply with Official Python Guideline.
# 050503 - Alexander - added Window menu support; re-numbered IDs; used standard wx constants for some menu items

# -------- wx menu item IDs --------

# wx.ID_NEW  # File menu
# wx.ID_OPEN
# wx.ID_CLOSE
# wx.ID_CLOSE_ALL
# wx.ID_SAVE
# wx.ID_SAVEAS
# wx.ID_EXIT

# wx.ID_UNDO  # Edit menu
# wx.ID_REDO

# wx.ID_MAXIMIZE_FRAME  # Window menu
# wx.ID_ICONIZE_FRAME

# wx.ID_ABOUT  # Help menu
# wx.ID_HELP


# -------- our menu item IDs --------

FIND_SCRIPTS   = 301  # Scripts menu

HOME_PAGE   = 501  # Help menu
HELP_PAGE   = 502
FORUM       = 503

FIRST_SCRIPT      = 11000  # ranges for generated items
LAST_SCRIPT       = 11999
FIRST_WINDOW      = 12000
LAST_WINDOW       = 12999


# -------- our tool IDs --------

NEW_PROJECT  = 2101  # Main
NEW_REPORT   = 2102
EDIT         = 2103
DUPLICATE    = 2104
DELETE       = 2105
SHOW_HIDDEN_REPORT = 2106

INSERT_ROW      = 2201  # Report
DUPLICATE_ROW   = 2202
DELETE_ROW      = 2203
MOVE_UP          = 2204
MOVE_DOWN        = 2205
PREREQUISITE     = 2206
ASSIGN_RESOURCE  = 2207
HIDE_ROW        = 2208
SHOW_HIDDEN      = 2209

INSERT_COLUMN   = 2301
# SELECT_CONTENTS = 2302
DELETE_COLUMN   = 2303
MOVE_LEFT       = 2304
MOVE_RIGHT      = 2305

COLUMN_OPTIONS    = 2401

SCROLL_LEFT_FAR  = 2502
SCROLL_LEFT      = 2503
SCROLL_RIGHT     = 2504
SCROLL_RIGHT_FAR = 2505
SCROLL_TO_TASK   = 2506
# NEW_RESOURCE     = 2601  # Resource
# EDIT_RESOURCE    = 2602
# ASSIGN_TASK      = 2603
# DELETE_RESOURCE  = 2604
# HIDE_RESOURCE    = 2605
# SHOW_HIDDEN_RESOURCE = 2606
