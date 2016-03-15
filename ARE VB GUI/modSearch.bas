Attribute VB_Name = "modSearch"
Option Explicit

'Just helper text
'ALName
'ALModTime
'EnabledDisabled
'ExecutionOrder
'ExecuteOn
'FocusFieldID
'ButtonID
'RunIfText
Private Const ALName = 1
Private Const ALModTime = 2
Private Const EnabledDisabled = 3
Private Const ExecutionOrder = 4
Private Const ExecuteOn = 5
Private Const FocusFieldID = 6
Private Const ButtonID = 7
Private Const RunIfText = 8

'NameBegins
'NameContains
'NameEnds
Private Const NameBegins = 1
Private Const NameContains = 2
Private Const NameEnds = 3

'ModGreater
'ModLess
'ModRange
'ModExactDate
Private Const ModGreater = 1
Private Const ModLess = 2
Private Const ModRange = 3
Private Const ModExactDate = 4

'ON_BUTTON
'ON_RETURN
'ON_SUBMIT
'ON_MODIFY
'ON_DISPLAY
'MODIFY_ALL
'MENU_OPEN
'MENU_CHOICE
'LOSE_FOCUS
'SET_DEFAULT
'ON_QUERY
'AFTER_MODIFY
'AFTER_SUBMIT
'GAIN_FOCUS
'WINDOW_OPEN
'WINDOW_CLOSE
Private Const ON_BUTTON = 1
Private Const ON_RETURN = 2
Private Const ON_SUBMIT = 4
Private Const ON_MODIFY = 8
Private Const ON_DISPLAY = 16
Private Const MODIFY_ALL = 32
Private Const MENU_OPEN = 64
Private Const MENU_CHOICE = 128
Private Const LOSE_FOCUS = 256
Private Const SET_DEFAULT = 512
Private Const ON_QUERY = 1024
Private Const AFTER_MODIFY = 2048
Private Const AFTER_SUBMIT = 4096
Private Const GAIN_FOCUS = 8192
Private Const WINDOW_OPEN = 16384
Private Const WINDOW_CLOSE = 32768


'ExecEqual
'ExecGreater
'ExecGreaterOrEqual
'ExecLess
'ExecLessOrEqual
Private Const ExecEqual = 1
Private Const ExecGreater = 2
Private Const ExecGreaterOrEqual = 3
Private Const ExecLess = 4
Private Const ExecLessOrEqual = 5

'FocusEqual
Private Const FocusEqual = 1

'ButtonEqual
Private Const ButtonEqual = 1

'RunBegins
'RunContains
'RunEnds
Private Const RunBegins = 1
Private Const RunContains = 2
Private Const RunEnds = 3

'To Do:  Need to limit for Focus Field ID (see iMISC.doc)
'AR_INTEGER_VALUE
'AR_REAL_VALUE
'AR_CHAR_VALUE
'AR_DIARY_VALUE
'AR_SELECTION_VALUE
'AR_DATE_VALUE
'AR_FIXEDDECIMAL_VALUE
'AR_ATTACHMENT_VALUE
'AR_TRIM_VALUE
'AR_CONTROL_VALUE
'AR_TABLE_VALUE
'AR_COLUMN_VALUE
'AR_PAGE_VALUE
'AR_PAGEHOLDER_VALUE
Private Const AR_INTEGER_VALUE = 2
Private Const AR_REAL_VALUE = 3
Private Const AR_CHAR_VALUE = 4
Private Const AR_DIARY_VALUE = 5
Private Const AR_SELECTION_VALUE = 6
Private Const AR_DATE_VALUE = 7
Private Const AR_FIXEDDECIMAL_VALUE = 10
Private Const AR_ATTACHMENT_VALUE = 11
Private Const AR_TRIM_VALUE = 31
Private Const AR_CONTROL_VALUE = 32
Private Const AR_TABLE_VALUE = 33
Private Const AR_COLUMN_VALUE = 34
Private Const AR_PAGE_VALUE = 35
Private Const AR_PAGEHOLDER_VALUE = 36


Private iPropertyValue As Long
Private iConditionValue As Long
Private iSearchValue As Long
Private sSearchString As String



Public Function AddSearchItem(sProperty As String, sCondition As String, sValue As String, Optional lNumValue As Long) As Long
  
  

End Function


Public Function ValidateSearchParamaters(sProperty As String, sCondition As String, sValue As String, Optional lNumValue As Long) As Boolean

  

End Function
