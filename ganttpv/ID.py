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

# Our menu item IDs:
# File

UNDO          = 1001 # Edit menu items.
REDO          = 1002
# edit_DUPLICATE     = 1003
# edit_DELETE        = 1005

FIND_SCRIPTS    = 1201 # Scripts menu

ABOUT          = 101 # Help menu items.
HOME_PAGE      = 102
HELP           = 103
HELP_PAGE      = 104
FORUM          = 105

# Our tool IDs:
# Main
NEW_PROJECT  = 2001
NEW_REPORT   = 2002
EDIT         = 2003
DUPLICATE    = 2004
DELETE       = 2005
SHOW_HIDDEN_REPORT = 2006

# Report
INSERT_ROW      = 2101
DUPLICATE_ROW   = 2102
DELETE_ROW      = 2103
MOVE_UP          = 2104
MOVE_DOWN        = 2105
PREREQUISITE     = 2106
ASSIGN_RESOURCE  = 2107
HIDE_ROW        = 2108
SHOW_HIDDEN      = 2109

INSERT_COLUMN   = 2201
SELECT_CONTENTS = 2202
DELETE_COLUMN   = 2303
MOVE_LEFT       = 2304
MOVE_RIGHT      = 2305

COLUMN_OPTIONS    = 2401
SCROLL_LEFT_FAR  = 2402
SCROLL_LEFT      = 2403
SCROLL_RIGHT     = 2404
SCROLL_RIGHT_FAR = 2405
SCROLL_TO_TASK   = 2406

# Resource
NEW_RESOURCE     = 2401
EDIT_RESOURCE    = 2402
ASSIGN_TASK      = 2403
DELETE_RESOURCE  = 2404
HIDE_RESOURCE    = 2405
SHOW_HIDDEN_RESOURCE = 2406
