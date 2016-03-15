VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsARIDLL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'***************************************************
'clsARIDLL:  Interfaces with AR DLL for information
'            display.
'***************************************************
'AR.DLL Entry Point Function Prototypes:
'
'***************************************************
'use:  private ARIDLL as new clsARIDLL
'***************************************************
'Default settings:
'ErrorLevel = Developer (Change for final release)
'***************************************************



'****************************************************************************************
'To Do's:
' Check for a valid connection before critical function calls
' Better Error handling is a MUST
' Speed things up
'****************************************************************************************

Const NoneValue = 1073741824



'***************
'Constants
'***************
Const arOK = 0

Private Const AR_ERROR = "AR"
Private Const FILTER_ERROR = "FILTER"
Private Const MISC_ERROR = "MISC"

Const AR_INTEGER = "Integer"
Const AR_REAL = "Real"
Const AR_CHAR = "Character"
Const AR_DIARY = "Diary"
Const AR_SELECTION = "Selection"
Const AR_DATE = "Date/time"
Const AR_FIXEDDECIMAL = "Fixed-point decimal"
Const AR_ATTACHMENT = "Attachment"
Const AR_TRIM = "Trim"
Const AR_CONTROL = "Control"
Const AR_TABLE = "Table"
Const AR_COLUMN = "Column"
Const AR_PAGE = "Page"
Const AR_PAGEHOLDER = "Page holder"


Const AR_INTEGER_VALUE = 2
Const AR_REAL_VALUE = 3
Const AR_CHAR_VALUE = 4
Const AR_DIARY_VALUE = 5
Const AR_SELECTION_VALUE = 6
Const AR_DATE_VALUE = 7
Const AR_FIXEDDECIMAL_VALUE = 10
Const AR_ATTACHMENT_VALUE = 11
Const AR_TRIM_VALUE = 31
Const AR_CONTROL_VALUE = 32
Const AR_TABLE_VALUE = 33
Const AR_COLUMN_VALUE = 34
Const AR_PAGE_VALUE = 35
Const AR_PAGEHOLDER_VALUE = 36


'*******************
'Public Properties
'*******************

'*******************
'Private Properties
'*******************
Private sARServerName As String
Private sUserID As String
Private sUserPassword As String

Private bConnectedToServer As Boolean

'This needs a little explaination.
'Since we're interfacing two COM objects in this class, we need to know which COM
'object the error occured for error retreival.
Private sErrorSource As String

Private AR As New AREConnect            'Active Link Searches
Private ARMisc As New Misc              'Get Field data and time conversion
Private ARFilter As New CSearchFilter   'Filter searches
Private ARLogin As New Login            'Server settings for connect, error text,etc...
Private ARFieldProperties As New ModifyFieldProps  'Modifying field properties
Private ARActiveLink As New ActiveLink  'Get Active Link properties

Private sErrorText As String
Private bErrorFlag As Boolean

Public sFormName As String
Public sEnabled As String
Public sExecutionOrder As String
Public sExecuteMask As String
Public sModifiedTime As String


'*******************
'Public Accessors
'*******************

Public Property Get IsConnected() As Boolean
  IsConnected = bConnectedToServer
End Property

'*******************
'Public Methods
'*******************

Public Function GetALProperty() As Long
Dim sExecOrder As String

  sFormName = ""
  sEnabled = ""
  sExecutionOrder = ""
  sExecuteMask = ""
  sModifiedTime = ""

  GetALProperty = AR.GetALProperties(sFormName, sEnabled, sExecOrder, sExecuteMask, sModifiedTime)
  
  sExecutionOrder = String(4 - Len(sExecOrder), " ") & sExecOrder
End Function

Public Function GetFilterProperty() As Long
Dim sTempModified As String
Dim sExecOrder As String

  sFormName = ""
  sEnabled = ""
  sExecutionOrder = ""
  sExecuteMask = ""
  sModifiedTime = ""

  GetFilterProperty = ARFilter.GetFilterProperties(sFormName, sEnabled, sExecOrder, sExecuteMask, sModifiedTime)
  sExecutionOrder = String(4 - Len(sExecOrder), " ") & sExecOrder
    
End Function


Public Sub SetServerName(sName As String)

  sARServerName = sName

End Sub
Public Function GetServerName() As String

  GetServerName = sARServerName
  
End Function


Public Sub SetUserName(sName As String)
  sUserID = sName
End Sub
Public Function GetUserName() As String
  GetUserName = sUserID
End Function


Public Function SetUserPassword(sPassword As String)
  sUserPassword = sPassword
End Function


Public Function GetErrorText() As String
  GetErrorText = sErrorText
End Function


Public Sub ErrorResolved(bValue As Boolean)
  ResetError
End Sub


Public Function Connected() As Boolean
  Connected = bConnectedToServer
End Function

Public Sub ResetFLSearchList()
  ARFilter.ResetSearch
End Sub

Public Sub ResetALSearchList()
  AR.ResetSearch
End Sub


'Returns 0 if no fields found, non zero is the field count
Public Function SetFieldIDPairs(sFormName As String) As Long
Dim lresult As Long

  ARMisc.FieldsReset
  
  lresult = ARMisc.PopulateFieldNames(sFormName)
  
  If lresult = arOK Then
    SetFieldIDPairs = ARMisc.GetNumberFields()
  Else
    SetError
    SetFieldIDPairs = 0
  End If

End Function

Public Function GetFormModTime(sFormName As String) As Long
Dim lresult As Long
Dim lModTime As Long
  
  lresult = ARMisc.GetFormModifiedTime(sFormName, lModTime)
  
  If lresult = arOK Then
    GetFormModTime = lModTime
  Else
    SetError
    GetFormModTime = 0
  End If

End Function


Public Function GetFieldID(Optional iIndex As Long) As Long
Dim lFieldID As Long
Dim i As Long

  If iIndex > 0 Then
    ARMisc.FieldsJumpToBegin
    
    If iIndex > 1 Then
      For i = 1 To iIndex - 1
        ARMisc.GotoNextFieldName
      Next i
    End If
    
  End If
  
  lFieldID = ARMisc.GetCurrentFieldID
  
  GetFieldID = lFieldID
  
End Function


Public Function GetFieldName(Optional iIndex As Long) As String
Dim sFieldName As String
Dim i As Long

  If iIndex > 0 Then
    ARMisc.FieldsJumpToBegin
    
    If iIndex > 1 Then
      For i = 1 To iIndex - 1
        ARMisc.GotoNextFieldName
      Next i
    End If
    
  End If
  
  sFieldName = ARMisc.GetCurrentFieldName()
  
  GetFieldName = sFieldName

End Function


Public Function GetFieldDataType(Optional lIndex As Long) As Long
Dim i As Long

  If lIndex > 0 Then
    ARMisc.FieldsJumpToBegin
    
    If lIndex > 1 Then
      For i = 1 To lIndex - 1
        ARMisc.GotoNextFieldName
      Next i
    End If
    
  End If

  GetFieldDataType = ARMisc.GetCurrentFieldType
  'GetFieldDateType = AR_INTEGER_VALUE
  
End Function

Public Function GetFieldDataTypeString(Optional lIndex As Long) As String
Dim i As Long

  If lIndex > 0 Then
    ARMisc.FieldsJumpToBegin
    
    If lIndex > 1 Then
      For i = 1 To lIndex - 1
        ARMisc.GotoNextFieldName
      Next i
    End If
    
  End If
  
  i = ARMisc.GetCurrentFieldType
  
  Select Case i
    Case AR_INTEGER_VALUE
      GetFieldDataTypeString = AR_INTEGER
    Case AR_REAL_VALUE
      GetFieldDataTypeString = AR_REAL
    Case AR_CHAR_VALUE
      GetFieldDataTypeString = AR_CHAR
    Case AR_DIARY_VALUE
      GetFieldDataTypeString = AR_DIARY
    Case AR_SELECTION_VALUE
      GetFieldDataTypeString = AR_SELECTION
    Case AR_DATE_VALUE
      GetFieldDataTypeString = AR_DATE
    Case AR_FIXEDDECIMAL_VALUE
      GetFieldDataTypeString = AR_FIXEDDECIMAL
    Case AR_ATTACHMENT_VALUE
      GetFieldDataTypeString = AR_ATTACHMENT
    Case AR_TRIM_VALUE
      GetFieldDataTypeString = AR_TRIM
    Case AR_CONTROL_VALUE
      GetFieldDataTypeString = AR_CONTROL
    Case AR_TABLE_VALUE
      GetFieldDataTypeString = AR_TABLE
    Case AR_COLUMN_VALUE
      GetFieldDataTypeString = AR_COLUMN
    Case AR_PAGE_VALUE
      GetFieldDataTypeString = AR_PAGE
    Case AR_PAGEHOLDER_VALUE
      GetFieldDataTypeString = AR_PAGEHOLDER
  End Select

End Function

Public Function SetFLEnabledDisabledParam(lValue As Long) As Long
Dim lresult As Long
  
  lresult = ARFilter.ParamEnabled(lValue)
  
  If Not lresult = arOK Then
    SetError
  End If
  
  SetFLEnabledDisabledParam = lresult
  
End Function

Public Function SetALEnabledDisabledParam(lValue As Long) As Long
Dim lresult As Long
  
  lresult = AR.ParamEnabled(lValue)
  
  If Not lresult = arOK Then
    SetError
  End If
  
  SetALEnabledDisabledParam = lresult
  
End Function

Public Function SetFLExecutionOrderParam(lValue As Long, lCondition As Long) As Long
Dim lresult As Long

  lresult = ARFilter.ParamExecutionOrder(lValue, lCondition)
  
  If Not lresult = arOK Then
    SetError
  End If
  
  SetFLExecutionOrderParam = lresult
  
End Function

Public Function SetALExecutionOrderParam(lValue As Long, lCondition As Long) As Long
Dim lresult As Long

  lresult = AR.ParamExecutionOrder(lValue, lCondition)
  
  If Not lresult = arOK Then
    SetError
  End If
  
  SetALExecutionOrderParam = lresult
  
End Function

Public Function SetALExecuteOnParam(lValue As Long, lCondition As Long) As Long
Dim lNoneValue As Long
Dim lresult As Long

  If lValue < NoneValue Then
    lNoneValue = 1
  Else
    lValue = lValue - NoneValue
    lNoneValue = 0
  End If
  
  lresult = AR.ParamExecuteOn(lCondition, lValue, lNoneValue)
  
  If Not lresult = arOK Then
    SetError
  End If
  
  SetALExecuteOnParam = lresult
  
End Function

Public Function SetFLExecuteOnParam(lValue As Long, lCondition As Long) As Long
Dim lresult As Long

  lresult = ARFilter.ParamExecuteOn(lCondition, lValue)
  
  If Not lresult = arOK Then
    SetError
  End If
  
  SetFLExecuteOnParam = lresult
  
End Function

Public Function SetALFocusFieldIDParam(lValue As Long) As Long
Dim lresult As Long

  lresult = AR.ParamFocusId(lValue)
  
  If Not lresult = arOK Then
    SetError
  End If

  SetALFocusFieldIDParam = lresult

End Function

Public Function SetALButtonIDParam(lValue As Long) As Long
Dim lresult As Long
  
  lresult = AR.ParamButtonId(lValue)
  
  If Not lresult = arOK Then
    SetError
  End If

  SetALButtonIDParam = lresult
  
End Function


Public Function ConvertDate(sDate As String) As Long
Dim iMonth As Integer
Dim iDay As Integer
Dim iYear As Integer
Dim iHour As Integer
Dim iMinute As Integer
Dim iSeconds As Integer

  iMonth = Month(sDate)
  iDay = Day(sDate)
  iYear = Year(sDate) - 1900
  iHour = Hour(sDate) '- 1
  iMinute = Minute(sDate)
  iSeconds = Second(sDate)
  
  ConvertDate = ARMisc.SetTime(iSeconds, iMinute, iHour, iDay, iMonth, iYear)

End Function


'Note:  sDate must be in "MM/DD/YY HH:MM:SS" format!!!!!
Public Function ConvertDateToJulian(sDate As Variant) As Long
Dim iMonth As Integer
Dim iDay As Integer
Dim iYear As Integer
Dim iHour As Integer
Dim iMinute As Integer
Dim iSeconds As Integer

  iMonth = Month(sDate) - 1
  iDay = Day(sDate)
  iYear = Year(sDate) - 1900
  iHour = Hour(sDate) - 1
  iMinute = Minute(sDate)
  iSeconds = Second(sDate)
  
  ConvertDateToJulian = ARMisc.SetTime(iSeconds, iMinute, iHour, iDay, iMonth, iYear)

End Function

Public Sub AddFLFormSearchItem(sFormName As String)
  ARFilter.AddFormToSearch (sFormName)
End Sub

Public Sub AddALFormSearchItem(sFormName As String)
   AR.AddFormToSearch (sFormName)
End Sub

'Public Function GetFLNumFormsToSearch() As Long
'  'OOPS
'  GetFLNumFormsToSearch = ARFilter.GetNumForms
'End Function

Public Function GetALNumFormsToSearch() As Long
  GetALNumFormsToSearch = AR.GetNumFormsToSearch
End Function


Public Function SetALNameSearchParam(sALName As String, lOperator As Long) As Long
Dim lresult As Long
  
  lresult = AR.ParamALName(sALName, lOperator)
  
  If Not lresult = arOK Then
    SetError
  End If
  
  SetALNameSearchParam = lresult
  
End Function

Public Function SetALCaseSensitive(bTFValue As Boolean) As Long
  If bTFValue = True Then
    AR.SetCaseSensitive
  Else
    AR.SetCaseInsensitive
  End If
End Function

Public Function SetFLCaseSensitive(bTFValue As Boolean) As Long
  If bTFValue = True Then
    ARFilter.SetCaseSensitive
  Else
    ARFilter.SetCaseInsensitive
  End If
End Function

Public Function SetFLNameSearchParam(sFLName As String, lOperator As Long) As Long
Dim lresult As Long

  lresult = ARFilter.ParamFilterName(sFLName, lOperator)
  
  If Not lresult = arOK Then
    SetError
  End If
  
  SetFLNameSearchParam = lresult
  
End Function

Public Function SetFLRunIfSearchParam(sALText As String, sOperator As String) As Long
Dim lresult As Long

  lresult = ARFilter.ParamRunIfText(sALText)
  
  If Not lresult = arOK Then
    SetError
  End If
  
  SetFLRunIfSearchParam = lresult
  
End Function

Public Function SetALRunIfSearchParam(sALText As String, sOperator As String) As Long
Dim lresult As Long

  lresult = AR.ParamRunIfText(sALText)
  
  If Not lresult = arOK Then
    SetError
  End If

  SetALRunIfSearchParam = lresult
  
End Function

Public Function SetALModifiedSearchParam(lTime As Long, lOperator As Long) As Long
Dim lresult As Long

  lresult = AR.ParamModTime(lTime, lOperator)

  If Not lresult = arOK Then
    SetError
  End If

  SetALModifiedSearchParam = lresult
  
End Function

Public Function SetFLModifiedSearchParam(lTime As Long, lOperator As Long) As Long
Dim lresult As Long

  lresult = ARFilter.ParamModTime(lTime, lOperator)

  If Not lresult = arOK Then
    SetError
  End If

  SetFLModifiedSearchParam = lresult

End Function

Public Function SearchActiveLink() As Long
Dim lresult As Long

  lresult = AR.SearchAL()
  
  If lresult <> 0 Then
    SetError
  End If
  
  sErrorSource = AR_ERROR
  SearchActiveLink = lresult
  
End Function

Public Function SearchFilter() As Long
Dim lresult As Long

  lresult = ARFilter.SearchFilters()
  
  If lresult <> 0 Then
    sErrorSource = FILTER_ERROR
    SetError
  End If
  
  SearchFilter = lresult
End Function

Public Function FoundNumberOfActiveLinks() As Long
  FoundNumberOfActiveLinks = AR.GetNumFoundAL
End Function

Public Function FoundNumberOfFilters() As Long
  FoundNumberOfFilters = ARFilter.GetNumFoundFilters()
End Function

Public Function GetFoundALName() As String
Dim sALName As String
Dim lresult As Long

  lresult = AR.GetFoundAL(sALName)
  
  If lresult = arOK Then
    GetFoundALName = sALName
  Else
    sErrorSource = AR_ERROR
    SetError
  End If
  
End Function

Public Function GetFoundFilterName() As String
Dim sFLName As String
Dim lresult As Long

  lresult = ARFilter.GetFoundFilter(sFLName)
  
  If lresult = arOK Then
    GetFoundFilterName = sFLName
  Else
    sErrorSource = FILTER_ERROR
    SetError
  End If
  
End Function

Public Sub GotoNextFoundALName()
  AR.GotoNextALName
End Sub

Public Sub GotoNextFoundFLName()
  ARFilter.GotoNextFilterName
End Sub



Public Function ConnectToARServer(sServer As String, sUser As String, sPass As String) As Boolean
Dim lresult As Long
  
  sARServerName = ""
  sUserID = ""
  sUserPassword = ""
  
  If Len(sServer) > 0 Then
    sARServerName = sServer
  Else
  End If
  If Len(sUser) > 0 Then
    sUserID = sUser
  End If
  If Len(sPass) > 0 Then
    sUserPassword = sPass
  End If
  
  bConnectedToServer = False
  
  AR.SetLogin sUserID, sUserPassword, sARServerName
  
  lresult = AR.DoLogin
  
  If lresult = arOK Then
    bConnectedToServer = True
  Else
    sErrorSource = AR_ERROR
    SetError
  End If
  
  ConnectToARServer = bConnectedToServer

End Function

Public Sub Logout()
  
  AR.ResetSearch
  ARFilter.ResetSearch
  ARMisc.FieldsReset
  AR.DoLogout

End Sub

Public Function DisconnectFromServer() As Boolean

  AR.DoLogout
  
  bConnectedToServer = False
  
  DisconnectFromServer = bConnectedToServer

End Function


Public Function CheckConnection() As Boolean
Dim iResult As Integer

  iResult = AR.GetConnectStatus
  
  If iResult = 0 Then
    CheckConnection = True
  Else
    CheckConnection = False
  End If
  
End Function


Public Function GetNumberOfForms() As Long
Dim iReturnValue As Integer

  iReturnValue = AR.GetForms
  
  If iReturnValue = arOK Then
    AR.GotoBeginForms
    GetNumberOfForms = AR.GetNumForms
  Else
    GetNumberOfForms = 0
    sErrorSource = AR_ERROR
    SetError True, "No forms found on server."
  End If

End Function


Public Function GetFormName(lFormNumber As Long) As String
Dim i As Long
Dim sFormName As String

  AR.GotoBeginForms
  
  If lFormNumber > 1 Then
    For i = 1 To lFormNumber - 1
      AR.GotoNextFormName
    Next i
  End If
  
  If AR.GetCurrentFormName(sFormName) = arOK Then
    GetFormName = sFormName
  Else
    GetFormName = "<Not Found>"
    SetError
  End If
  
End Function 'End of GetFormName()


Public Function LoadPermissions() As Long

  ARMisc.GroupListReset

  If ARMisc.StoreAllGroupInfo = arOK Then
    LoadPermissions = ARMisc.GetNumGroups
  Else
    LoadPermissions = 0
    SetError
  End If

End Function


Public Function GroupOK() As Boolean

  GroupOK = ARMisc.IsCurrentGroupAL

End Function

Public Sub GotoNextGroup()

  ARMisc.GotoNextGroupName

End Sub

Public Function GetCurrentGroup() As String
Dim sName As String

  If ARMisc.GetCurrentGroupName(sName) = arOK Then
    GetCurrentGroup = sName
  Else
    GetCurrentGroup = ""
    SetError
  End If

End Function



'***********************************
'Active Link Modification Functions:
'***********************************
Public Sub ResetActiveLinkModify()

  AR.ResetModify

End Sub

Public Function AddActiveLinkToModify(sActiveLinkName As String) As Long

  If AR.AddALToModify(sActiveLinkName) = arOK Then
    AddActiveLinkToModify = arOK
  Else
    AddActiveLinkToModify = 1
    SetError
  End If

End Function

Public Function SetActiveLinkName(sName As String) As Long

  If AR.SetALName(sName) = arOK Then
    SetActiveLinkName = arOK
  Else
    SetActiveLinkName = 1
    SetError
  End If

End Function

Public Function SetActiveLinkExecutionOrder(lNumber As Long) As Long

  If AR.SetALExecutionOrder(lNumber) = arOK Then
    SetActiveLinkExecutionOrder = arOK
  Else
    SetActiveLinkExecutionOrder = 1
    SetError
  End If

End Function

Public Function SetActiveLinkEnabled(bYesNo As Boolean) As Long
Dim lresult As Long

  If bYesNo = True Then
    lresult = AR.SetALEnabled(1)
  Else
    lresult = AR.SetALEnabled(0)
  End If
  
  If lresult = arOK Then
    SetActiveLinkEnabled = arOK
  Else
    SetActiveLinkEnabled = 1
    SetError
  End If

End Function

Public Function SetActiveLinkChangeHistory(sText As String) As Long

  If AR.SetALChangeHistory(sText) = arOK Then
    SetActiveLinkChangeHistory = arOK
  Else
    SetActiveLinkChangeHistory = 1
    SetError
  End If

End Function

Public Sub SetActiveLinkHelpTextAction(sAction As String)

  Select Case sAction
  Case "Append"
    AR.SetHelpTextActionAppend
  Case "Overwrite"
    AR.SetHelpTextActionOverwrite
  End Select

End Sub


Public Function SetActiveLinkHelpText(sText As String) As Long

  If AR.SetALHelpText(sText) = arOK Then
    SetActiveLinkHelpText = arOK
  Else
    SetActiveLinkHelpText = 1
    SetError
  End If

End Function


Public Function SetActiveLinkPermissionType(sType As String) As Long
Dim lresult As Long

  Select Case sType
    Case "Append"
      AR.SetPermActionAppend
    Case "Overwrite"
      AR.SetPermActionOverwrite
    Case "Remove All"
      AR.SetPermActionRemoveAll
  End Select
  
  SetActiveLinkPermissionType = arOK

End Function


Public Function SetActiveLinkPermissionGroup(sGroup As String) As Long

  If AR.SetALPermissionGroup(sGroup) = arOK Then
    SetActiveLinkPermissionGroup = arOK
  Else
    SetActiveLinkPermissionGroup = 1
    SetError
  End If

End Function


Public Function ExecuteActiveLinkModification() As Long

  If AR.ModifyAL = arOK Then
    ExecuteActiveLinkModification = arOK
  Else
    ExecuteActiveLinkModification = 1
    SetError
  End If

End Function


'******************************
'Filter Modification Functions:
'******************************
Public Function AddFilterToModify(sFilterName As String) As Long

  If ARFilter.AddFilterToModify(sFilterName) = arOK Then
    AddFilterToModify = arOK
  Else
    AddFilterToModify = 1
    SetError
  End If
  

End Function


Public Sub ResetFilterModify()
  
  ARFilter.ResetModify
  
End Sub

Public Function SetFilterNameModify(sName As String) As Long

  If ARFilter.SetFLName(sName) = arOK Then
    SetFilterNameModify = arOK
  Else
    SetFilterNameModify = 1
    SetError
  End If

End Function

Public Function SetFilterExecutionOrder(lNumber As Long) As Long

  If ARFilter.SetFLExecutionOrder(lNumber) = arOK Then
    SetFilterExecutionOrder = arOK
  Else
    SetFilterExecutionOrder = 1
    SetError
  End If

End Function

Public Function SetFilterEnabled(bYesNo As Boolean) As Long
Dim lresult As Long

  If bYesNo = True Then
    lresult = ARFilter.SetFLEnabled(1)
  Else
    lresult = ARFilter.SetFLEnabled(0)
  End If
  
  If lresult = arOK Then
    SetFilterEnabled = arOK
  Else
    SetFilterEnabled = 1
    SetError
  End If
  
End Function

Public Function SetFilterChangeHistory(sText As String) As Long

  If ARFilter.SetFLChangeHistory(sText) = arOK Then
    SetFilterChangeHistory = arOK
  Else
    SetFilterChangeHistory = 1
    SetError
  End If

End Function

Public Sub SetFilterHelpTextAction(sAction As String)

  Select Case sAction
  Case "Append"
    ARFilter.SetHelpTextActionAppend
  Case "Overwrite"
    ARFilter.SetHelpTextActionOverwrite
  End Select

End Sub

Public Function SetFilterHelpText(sText As String) As Long
 
  If ARFilter.SetFLHelpText(sText) = arOK Then
    SetFilterHelpText = arOK
  Else
    SetFilterHelpText = 1
    SetError
  End If

End Function

Public Function ExecuteFilterModification() As Long

  If ARFilter.ModifyFilter = arOK Then
    ExecuteFilterModification = arOK
  Else
    ExecuteFilterModification = 1
    SetError
  End If

End Function


'******************************
'Field Modification Functions:
'ARE Object: ARFieldProperties
'******************************
'Call this first to reset Object
Public Sub ResetFieldObject()

  ARFieldProperties.Reset

End Sub


'Feed the list of FormNames and FieldIDs one at a time here BEFORE any changes are made
Public Function SetFormFieldPair(sFormName As String, lFieldID As Long) As Long
Dim lresult As Long

  lresult = ARFieldProperties.ModifyField(lFieldID, sFormName)
  
  If lresult = arOK Then
    SetFormFieldPair = arOK
  Else
    SetFormFieldPair = lresult
    SetError
  End If
  
End Function


Public Sub SetFieldCreateMode(iMode As Integer)

  Select Case iMode
  Case CreateModeOPEN
    ARFieldProperties.SetCreateModeOpen
  Case CreateModePROTECTED
    ARFieldProperties.SetCreateModeProtected
  End Select

End Sub


Public Function SetFieldDBName(sDBName As String) As Long

  If ARFieldProperties.SetFieldDBName(sDBName) = arOK Then
    SetFieldDBName = arOK
  Else
    SetFieldDBName = 1
    SetError
  End If

End Function


Public Sub SetFieldHelpTextAction(sAction As String)

  Select Case sAction
  Case "Append"
    ARFieldProperties.SetHelpTextActionAppend
  Case "Overwrite"
    ARFieldProperties.SetHelpTextActionOverwrite
  End Select

End Sub


Public Function SetFieldHelpText(sHelpText As String) As Long

  If ARFieldProperties.SetHelpText(sHelpText) = arOK Then
    SetFieldHelpText = arOK
  Else
    SetFieldHelpText = 1
    SetError
  End If

End Function


Public Sub SetFieldOptionalProp(iOption As Integer)

  Select Case iOption
  Case PropertyOptional
    ARFieldProperties.SetOptionalPropOptional
  Case PropertyRequired
    ARFieldProperties.SetOptionalPropRequired
  End Select

End Sub


Public Function SetFieldChangeHistory(sChangeHistory As String) As Long

  If ARFieldProperties.SetChangeHistory(sChangeHistory) = arOK Then
    SetFieldChangeHistory = arOK
  Else
    SetFieldChangeHistory = 1
    SetError
  End If

End Function


Public Function SetFieldLabel(sNewLabel As String) As Long

  If ARFieldProperties.SetGlobalFieldLabel(sNewLabel) = arOK Then
    SetFieldLabel = arOK
  Else
    SetFieldLabel = 1
    SetError
  End If

End Function

'****************************************
'Functions to get Active Link Properties
'****************************************

'Sets the Active Link Name for which we're going to retrieve properties
Public Function ActiveLinkName(sLocalALName As String) As Long
    
    'Store the Active Link Name in AREServer.dll
    ARActiveLink.SetName (sLocalALName)
        
End Function

Public Function GetIfActionList() As Long
    'Get Active Link Actions List and store in AREServer.dll
    If ARActiveLink.GetIfActionList() = arOK Then
        GetIfActionList = arOK
    Else
        GetIfActionList = 1
        SetError
    End If

End Function

Public Function GetElseActionList() As Long
    'Get Active Link Actions List and store in AREServer.dll
    If ARActiveLink.GetElseActionList() = arOK Then
        GetElseActionList = arOK
    Else
        GetElseActionList = 1
        SetError
    End If

End Function

'Gets Active Link if actions from server and stores in AREServer.dll for retrieval
Public Function GetALActionCount() As Long
    
    'Return the number of Actions for the specified active link
    GetALActionCount = ARActiveLink.GetIfActionsCount
    
End Function
'Gets Active Link if actions from server and stores in AREServer.dll for retrieval
Public Function GetALElseActionCount() As Long
    
    'Return the number of Actions for the specified active link
    GetALElseActionCount = ARActiveLink.GetElseActionsCount
    
End Function


Public Function GetALActionType() As String
Dim ActionType As String
Dim returnCode As Long
    
    'Get the Action Type and return it
    returnCode = ARActiveLink.GetIfActionType(ActionType)
    If returnCode = arOK Then
        GetALActionType = ActionType
    Else
        GetALActionType = "<None>"
        SetError
    End If

End Function

Public Function GetALElseActionType() As String
Dim ActionType As String
Dim returnCode As Long
    
    'Get the Action Type and return it
    returnCode = ARActiveLink.GetElseActionType(ActionType)
    If returnCode = arOK Then
        GetALElseActionType = ActionType
    Else
        GetALElseActionType = "<None>"
        SetError
    End If

End Function


Public Function GotoNextALAction() As Long
    
    'Goto the next AL Action and return
    If ARActiveLink.NextIfAction() = arOK Then
        GotoNextALAction = arOK
    Else
        GotoNextALAction = 1
        SetError
    End If
    
End Function

Public Function GotoNextALElseAction() As Long
    
    'Goto the next AL Action and return
    If ARActiveLink.NextElseAction() = arOK Then
        GotoNextALElseAction = arOK
    Else
        GotoNextALElseAction = 1
        SetError
    End If
    
End Function


'*******************
'Private Methods
'*******************

Private Sub SetError(Optional bOveride As Boolean, Optional sOverideErrorText As String)
Dim sText As String

  sErrorText = "                                                                      "


  If Len(sOverideErrorText) > 0 Then
    If bOveride = True Then
      sText = sOverideErrorText
    End If
  End If

  sText = sText & AR.GetErrorText
  
  sErrorText = vbCrLf & sText
  bErrorFlag = True

End Sub


Private Sub ResetError()
  sErrorText = "No error has occured."
  bErrorFlag = False
End Sub



Private Sub Class_Initialize()

  ARLogin.GetExtendedUserInfo (7548652)

End Sub
