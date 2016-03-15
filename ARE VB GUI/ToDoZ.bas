Attribute VB_Name = "ToDoZ"
Option Explicit

'ApplicationOptions
'**DONE  Add option for Number of Valid Queries to track.
'**DONE  Add option to load a Default query type
'**DONE  Add option to load a default saved query
'**DONE  Add option for Default Server Name
'**REMOVED  Add default forms to be checked
'**DONE  Add Locked search dialog to open

'SEARCH DIALOG
'**FIXED Hide "OR" button and change "AND" button to "Add" unless Execute_On Selected
'**      then change "Add" to "And" and show "OR",  Update all execute_on's to same
'**FIXED When user clicks on an item, show that item's info in the controls (along with tag
'**      values).
'**FIXED Can not hand enter a field name when searching. Filters
'**FIXED Add Case sensative checkbox
'**FIXED Fix "More"

'CACHE
'**FIXED  Only Refresh cache on startup
'**DONE:  Add abilitiy to search cache for Field ID's
'**Move to own DB (Seperate from saved queries and Results DB)

'TOOLBAR, DELAYED To Ver2: Change toolbar to CoolBar with embedded toolbar and other stuffs...
'**FIXED Fix Search Button on Toolbar to only be active when connected to server.
'**FIXED -Make Functional
'**FIXED Add buttons for user to  assign to saved search queries
'**FIXED -Make Functional
'**FIXED Add Reset Search Param's button, bitmap done
'**FIXED -Make Functional
'**FIXED Add PrintResults Button, bitmap done
'**FIXED      -Make Functional
'**FIXED Add SaveResults Button, bitmap done
'**FIXED      -Make Functional
'**FIXED Add SaveQuery Button
'**FIXED -Make Functional
'**FIXED Add OpenQuery Button
'**FIXED -Make Functional
'**FIXED Add DeleteQuery Button, bitmap done
'**FIXED -Make Functional
'**FIXED Add Disconnect/Connect (Must prompt for Username/Password/Server), bitmap done
'**FIXED -Make Functional
'**FIXED Add Filter Search button
'**FIXED -Make Functional
'**FIXED Add Previous Query button, bitmap done
'**FIXED -Make Functional
'**FIXED Add Next Query button, bitmap done
'**FIXED -Make Functional

'MENUBAR
'**FIXED Remove un-needed menu items
'**FIXED Add needed menu items (Save Search, Load Search, etc..)
'**FIXED Add functionality for menu items

'TREEVIEW (tvTreeView)  For these items, copy functionality in tvIDPicker
'**MOVED TO VER 2  Allow multiple server functionality
'**Add object count to each Parent (ex: 'Forms (12)')

'LISTVIEW (lvListView)
'**FIXED  Add AL Types to the lvListView Columns
'**DELAYED:  Hide unwanted lvListView Columns
'**FIXED  Save order of lvListView Columns
'**FIXED  Save width of lvListView Columns
'**DELAYED:  Create Options that allow user to choose which fields to display in lvListView
'**FIXED  Fix Modified Date Displayed, ARMisc problem, waiting for updated DLL
'**FIXED  Fix Execution Order Sort bug

'SEARCHING
'**FIXED   Make sure only one of each type (with the exception of Execute On) are allowed
'          to be searched in one query.
'**FIXED  Create ability to save predefined search queries
'**FIXED Fix Date/Time conversion bug
'**FIXED Add Filter Search capability-  Needs testing
'**FIXED  -Need to update search dialog
'**FIXED  -Adjust "Searching: " message
'**FIXED Add ability to reverse up to x amount of previous queries (x set by user in prefrences)

'GENERAL
'**Add Status Messages for EVERYTHING
'**Better Error Handling
'**FIXED Add shortcut menus for everything
'**Help File

'ID TREEVIEW (frmIDPicker.tvIDDisplay)
'**Add Count of ID's to parents (ex:  'integers (2)')

'WOULD LIKE TO DO'S
'**Move all search code from frmMain to query collection
