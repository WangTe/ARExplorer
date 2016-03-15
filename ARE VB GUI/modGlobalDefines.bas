Attribute VB_Name = "modGlobalDefines"
Option Explicit
'For the sake of ease, a commented list of constants will be placed before the actual definition
'(make select case's easier)

Global Const PropertyOptional = 1
Global Const PropertyRequired = 2

Global Const CreateModeOPEN = 1
Global Const CreateModePROTECTED = 2

Global Const ChangeHistory = 1
Global Const HelpText = 2

Global Const OriginalInstall = "StartupPosition"
Global Const LastUsedDate = "LastPosition"
Global Const CurrentDate = "CurrentPosition"

Global Const sEmptyString = "<EMPTY>"
Global Const iMaxAssignedCount = 5

Global Const sServerTableName = "Server"
Global Const sFormTableName = "FormName"
Global Const sFieldTableName = "FieldProperties"
Global Const sQueryTableName = "tblQueries"
Global Const sQueryItemsTableName = "tblQueryItems"

'Public field names for the above table.
Global Const fldName = "Name"
Global Const fldID = "ID"
Global Const fldModTime = "ModTime"
Global Const fldARID = "ARID"
Global Const fldType = "Type"
Global Const fldParentServerID = "ParentServerID"
Global Const fldParentFormID = "ParentFormID"
Global Const fldQueryName = "QueryName"
Global Const fldExecuteOnCount = "ExecuteOnCount"
Global Const fldExecuteOnValue = "ExecuteOnValue"
Global Const fldCaseSensitive = "CaseSensitive"
Global Const fldQueryID = "QueryID"
Global Const fldSearchParam = "SearchParamater"
Global Const fldConditionString = "SearchConditionString"
Global Const fldConditionValue = "SearchConditionValue"
Global Const fldSearchType = "SearchType"
Global Const fldSearchValueNumber = "SearchValueNumber"
Global Const fldSearchValueString = "SearchValueString"
Global Const fldSearchANDOR = "SearchANDOR"


'TYPE_AL
'TYPE_FILTER
'TYPE_FORM
Global Const TYPE_AL = "AL"
Global Const TYPE_FILTER = "FILTER"
Global Const TYPE_FORM = "FORM"
Global Const TYPE_FIELD = "FIELD"

'QUERY_OPEN
'QUERY_SAVE
'QUERY_DELETE
'QUERY_ASSIGN
Global Const QUERY_OPEN = "Open"
Global Const QUERY_SAVE = "Save"
Global Const QUERY_DELETE = "Delete"
Global Const QUERY_ASSIGN = "Assign"

'ICON ORDER:
'icoShowSearch
'icoHideSearch
'icoPerformSearch
'icoSaveQuery
'icoOpenQuery
'icoSearch1
'icoSearch2
'icoSearch3
'icoSearch4
'icoSearch5
'icoActiveLink
'icoFilter
'icoSaveResults
'icoConnect
'icoDisconnect
'icoDeleteQuery
'icoResetQuery
'icoPrintResults
'icoPreviousQuery
'icoNextQuery
Global Const icoShowSearch = 1
Global Const icoHideSearch = 2
Global Const icoPerformSearch = 3
Global Const icoSaveQuery = 4
Global Const icoOpenQuery = 5
Global Const icoSearch1 = 6
Global Const icoSearch2 = 7
Global Const icoSearch3 = 8
Global Const icoSearch4 = 9
Global Const icoSearch5 = 10
Global Const icoActiveLink = 11
Global Const icoFilter = 12
Global Const icoSaveResults = 13
Global Const icoConnect = 14
Global Const icoDisconnect = 15
Global Const icoDeleteQuery = 16
Global Const icoResetQuery = 17
Global Const icoPrintResults = 18
Global Const icoPreviousQuery = 19
Global Const icoNextQuery = 20
Global Const icoFields = 23
Global Const icoModify = 21


'BUTTON ORDER:
'Connection
'ResetQuery
'SaveQuery
'SaveResults
'OpenQuery
'DeleteQuery
'PrintResults
'ShowSearch
'SearchAL
'SearchFilter
'SavedQuery1
'SavedQuery2
'SavedQuery3
'SavedQuery4
'SavedQuery5
'PerformSearch
'PreviousQuery
'NextQuery
Global Const ConnectionNumber = 1
Global Const ResetQueryNumber = 3
Global Const SaveQueryNumber = 4
Global Const SaveResultsNumber = 5
Global Const OpenQueryNumber = 6
Global Const DeleteQueryNumber = 7
Global Const PrintResultsNumber = 8
Global Const ShowSearchNumber = 10
Global Const SearchALNumber = 11
Global Const SearchFilterNumber = 12
Global Const SavedQuery1Number = 13
Global Const SavedQuery2Number = 14
Global Const SavedQuery3Number = 15
Global Const SavedQuery4Number = 16
Global Const SavedQuery5Number = 17
Global Const PerformSearchNumber = 18
Global Const PreviousQueryNumber = 20
Global Const NextQueryNumber = 21




'AR_INTEGER
'AR_REAL
'AR_CHAR
'AR_DIARY
'AR_SELECTION
'AR_DATE
'AR_FIXEDDECIMAL
'AR_ATTACHMENT
'AR_TRIM
'AR_CONTROL
'AR_TABLE
'AR_COLUMN
'AR_PAGE
'AR_PAGEHOLDER
Global Const AR_INTEGER = "Integer"
Global Const AR_REAL = "Real"
Global Const AR_CHAR = "Character"
Global Const AR_DIARY = "Diary"
Global Const AR_SELECTION = "Selection"
Global Const AR_DATE = "Date/time"
Global Const AR_FIXEDDECIMAL = "Fixed-point decimal"
Global Const AR_ATTACHMENT = "Attachment"
Global Const AR_TRIM = "Trim"
Global Const AR_CONTROL = "Control"
Global Const AR_TABLE = "Table"
Global Const AR_COLUMN = "Column"
Global Const AR_PAGE = "Page"
Global Const AR_PAGEHOLDER = "Page holder"


'AR_ALNAME_NUMBER
'AR_MODTIME_NUMBER
'AR_ENABLEDDISABLED_NUMBER
'AR_EXECUTIONORDER_NUMBER
'AR_EXECUTEON_NUMBER
'AR_FOCUSFIELDNAME_NUMBER
'AR_BUTTONNAME_NUMBER
'AR_RUNIFTEXT_NUMBER
'AR_FILTERNAME_NUMBER
Global Const AR_ALNAME_NUMBER = 1
Global Const AR_MODTIME_NUMBER = 2
Global Const AR_ENABLEDDISABLED_NUMBER = 3
Global Const AR_EXECUTIONORDER_NUMBER = 4
Global Const AR_EXECUTEON_NUMBER = "8"
Global Const AR_FOCUSFIELDNAME_NUMBER = 5
Global Const AR_BUTTONNAME_NUMBER = 6
Global Const AR_RUNIFTEXT_NUMBER = 7
Global Const AR_FILTERNAME_NUMBER = 1
Global Const AR_FIELDNAME_NUMBER = 1
Global Const AR_FIELDID_NUMBER = 2
Global Const AR_FIELDTYPE_NUMBER = 3
'AR_ENABLED_NUMBER
'AR_DISABLED_NUMBER
'AR_ONBUTTON_NUMBER
'AR_ONRETURN_NUMBER
'AR_ONSUBMIT_NUMBER
'AR_ONMODIFY_NUMBER
'AR_ONDISPLAY_NUMBER
'AR_MODIFYALL_NUMBER
'AR_MENUOPEN_NUMBER
'AR_MENUCHOICE_NUMBER
'AR_LOSEFOCUS_NUMBER
'AR_SETDEFAULT_NUMBER
'AR_ONQUERY_NUMBER
'AR_AFTERMODIFY_NUMBER
'AR_AFTERSUBMIT_NUMBER
'AR_GAINFOCUS_NUMBER
'AR_WINDOWOPEN_NUMBER
'AR_WINDOWCLOSE_NUMBER
Global Const AR_ENABLED_NUMBER = 9
Global Const AR_DISABLED_NUMBER = 10
Global Const AR_ONBUTTON_NUMBER = 11
Global Const AR_ONRETURN_NUMBER = 12
Global Const AR_ONSUBMIT_NUMBER = 13
Global Const AR_ONMODIFY_NUMBER = 14
Global Const AR_ONDISPLAY_NUMBER = 15
Global Const AR_MODIFYALL_NUMBER = 16
Global Const AR_MENUOPEN_NUMBER = 17
Global Const AR_MENUCHOICE_NUMBER = 18
Global Const AR_LOSEFOCUS_NUMBER = 19
Global Const AR_SETDEFAULT_NUMBER = 20
Global Const AR_ONQUERY_NUMBER = 21
Global Const AR_AFTERMODIFY_NUMBER = 22
Global Const AR_AFTERSUBMIT_NUMBER = 23
Global Const AR_GAINFOCUS_NUMBER = 24
Global Const AR_WINDOWOPEN_NUMBER = 25
Global Const AR_WINDOWCLOSE_NUMBER = 26
Global Const AR_GET_NUMBER = 27
Global Const AR_DELETE_NUMBER = 28
Global Const AR_MERGE_NUMBER = 29
Global Const AR_NONE_NUMBER = 30


'AR_FIELDNAME
'AR_FIELDID
'AR_FIELDTYPE
Global Const AR_FIELDNAME = "Database Name"
Global Const AR_FIELDID = "Field ID"
Global Const AR_FIELDTYPE = "Type"


'AR_ALNAME
'AR_MODTIME
'AR_ENABLEDDISABLED
'AR_EXECUTIONORDER
'AR_EXECUTEON
'AR_FOCUSFIELDNAME
'AR_BUTTONNAME
'AR_RUNIFTEXT
'AR_FILTERNAME
Global Const AR_ALNAME = "Active Link Name"
Global Const AR_MODTIME = "Modification Time"
Global Const AR_ENABLEDDISABLED = "Enabled / Disabled"
Global Const AR_EXECUTIONORDER = "Execution Order"
Global Const AR_EXECUTEON = "Execute On"
Global Const AR_FOCUSFIELDNAME = "Focus Field Name"
Global Const AR_BUTTONNAME = "Button Name"
Global Const AR_RUNIFTEXT = "Run If Text"
Global Const AR_FILTERNAME = "Filter Name"


'AR_BEGINSWITH
'AR_CONTAINS
'AR_ENDSWITH
'AR_GREATERTHAN
'AR_DATERANGE
'AR_EXACTDATE
'AR_EQUAL
'AR_GREATERTHANOREQUAL
'AR_LESSTHAN
'AR_LESSTHANOREQUAL
Global Const AR_BEGINSWITH = "Begins with"
Global Const AR_CONTAINS = "Contains"
Global Const AR_ENDSWITH = "Ends with"
Global Const AR_GREATERTHAN = "Greater than"
Global Const AR_DATERANGE = "Date range"
Global Const AR_EXACTDATE = "Exact date"
Global Const AR_EQUAL = "Equal"
Global Const AR_GREATERTHANOREQUAL = "Greater than or equal"
Global Const AR_LESSTHAN = "Less than"
Global Const AR_LESSTHANOREQUAL = "Less than or equal"

'AR_ENABLED
'AR_DISABLED
'AR_ONBUTTON
'AR_ONRETURN
'AR_ONSUBMIT
'AR_ONMODIFY
'AR_ONDISPLAY
'AR_MODIFYALL
'AR_MENUOPEN
'AR_MENUCHOICE
'AR_LOSEFOCUS
'AR_SETDEFAULT
'AR_ONQUERY
'AR_AFTERMODIFY
'AR_AFTERSUBMIT
'AR_GAINFOCUS
'AR_WINDOWOPEN
'AR_WINDOWCLOSE
Global Const AR_ENABLED = "Enabled"
Global Const AR_DISABLED = "Disabled"
Global Const AR_ONBUTTON = "Button/Menu Item"
Global Const AR_ONRETURN = "Return"
Global Const AR_ONSUBMIT = "Submit"
Global Const AR_ONMODIFY = "Modify"
Global Const AR_ONDISPLAY = "Display"
'Global Const AR_MODIFYALL = "MODIFY_ALL"
'Global Const AR_MENUOPEN = "MENU_OPEN"
Global Const AR_MENUCHOICE = "Menu/Row Choice"
Global Const AR_LOSEFOCUS = "Lose Focus"
Global Const AR_SETDEFAULT = "Set Default"
Global Const AR_ONQUERY = "Search"
Global Const AR_AFTERMODIFY = "After Modify"
Global Const AR_AFTERSUBMIT = "After Submit"
Global Const AR_GAINFOCUS = "Gain Focus"
Global Const AR_WINDOWOPEN = "Window Open"
Global Const AR_WINDOWCLOSE = "Window Close"
Global Const AR_NONE = "None"
Global Const AR_GET = "Get"
Global Const AR_DELETE = "Delete"
Global Const AR_MERGE = "Merge"

'KEY_ALNAME
'KEY_FORMNAME
'KEY_MODTIME
'KEY_EXECUTIONMASK
'KEY_EXECUTIONORDER
'KEY_ENABLED
Global Const KEY_ALNAME = "keyALName"
Global Const KEY_FORMNAME = "keyFormName"
Global Const KEY_MODTIME = "keyModTime"
Global Const KEY_EXECUTIONMASK = "keyExecutionMask"
Global Const KEY_EXECUTIONORDER = "keyExecutionOrder"
Global Const KEY_ENABLED = "keyEnabled"

