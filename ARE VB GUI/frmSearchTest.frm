VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIDPicker 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form Objects"
   ClientHeight    =   1875
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   3750
   Icon            =   "frmSearchTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imlIDImages 
      Left            =   2940
      Top             =   780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearchTest.frx":0442
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvIDDisplay 
      Height          =   1815
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   3201
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   0
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      HotTracking     =   -1  'True
      BorderStyle     =   1
      Appearance      =   1
   End
End
Attribute VB_Name = "frmIDPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const ROOT = "Forms"
Const ID = "ID"

Const ICON_FORMS = 1
Const ICON_BUTTONID = 2
Const ICON_FOCUSID = 3

'Const AR_INTEGER = "Integer"
'Const AR_REAL = "Real"
'Const AR_CHAR = "Character"
'Const AR_DIARY = "Diary"
'Const AR_SELECTION = "Selection"
'Const AR_DATE = "Date/time"
'Const AR_FIXEDDECIMAL = "Fixed-point decimal"
'Const AR_ATTACHMENT = "Attachment"
'Const AR_TRIM = "Trim"
'Const AR_CONTROL = "Control"
'Const AR_TABLE = "Table"
'Const AR_COLUMN = "Column"
'Const AR_PAGE = "Page"
'Const AR_PAGEHOLDER = "Page holder"


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

Const KEY_ROOT_PREFIX = "root"
Const KEY_CHILD_PREFIX = "node"

Private lFieldCount As Long

Private lFormNumber As Long
Private sFormNumberString As String
Private lDataTypeNumber As Long
Private sDataTypeString As String

Private sIDListMode As String


Public Sub ClearForm()

  tvIDDisplay.Nodes.Clear
  lFormNumber = 0
  sFormNumberString = Trim(Str(lFormNumber))
  lDataTypeNumber = 0
  sDataTypeString = Trim(Str(lDataTypeNumber))

End Sub



Private Sub Form_Load()

  ClearForm
  SizeTreeView
  
  Me.Width = GetSetting(App.Title, "Settings", "IDWidth", 3870)
  Me.Height = GetSetting(App.Title, "Settings", "IDHeight", 2235)

End Sub


Private Sub Form_Resize()

  SizeTreeView

End Sub


Private Sub SizeTreeView()

  tvIDDisplay.Top = Me.ScaleTop
  tvIDDisplay.Left = Me.ScaleLeft
  tvIDDisplay.Width = Me.ScaleWidth
  tvIDDisplay.Height = Me.ScaleHeight

End Sub



'The cache will be updated BEFORE this is called
'Changed to pull from Cache to be much faster
'All we need to do is add the form names to the TreeView and query the cache to see if
'the form has any valid fields.
Private Sub AddAllFieldNames(sFormName As String)
Dim i As Long
Dim sType As String
Dim lID As Long
Dim lParentID As Long
Dim sName As String
Dim sNodeName As String
Dim bDisplay As Boolean
Dim rsFieldList As Recordset
Dim sSQL As String

  lParentID = modDatabase.GetFormCacheID(sFormName, frmMain.lCurrentServerID)

  sSQL = "SELECT * "
  sSQL = sSQL & "FROM [" & sFieldTableName & "] "
  sSQL = sSQL & "WHERE [" & fldParentFormID & "] = " & Trim(Str(lParentID)) & " "
  
  Select Case sIDListMode
    Case AR_RUNIFTEXT
      sSQL = sSQL & ";"
    Case AR_FOCUSFIELDNAME
      sSQL = sSQL & "AND (([" & fldType & "] NOT LIKE '" & AR_CONTROL & "') "
      sSQL = sSQL & "AND ([" & fldType & "] NOT LIKE '" & AR_COLUMN & "') "
      sSQL = sSQL & "AND ([" & fldType & "] NOT LIKE '" & AR_PAGEHOLDER & "'));"
    Case AR_BUTTONNAME
      sSQL = sSQL & "AND (([" & fldType & "] LIKE '" & AR_CONTROL & "'));"
  End Select
  
  Set rsFieldList = modDatabase.ExecuteCacheSQL(sSQL)

  On Error Resume Next
  rsFieldList.MoveLast
  rsFieldList.MoveFirst
  On Error GoTo 0
  
  lFieldCount = rsFieldList.RecordCount
  
  If lFieldCount > 0 Then
    For i = 1 To lFieldCount

      lDataTypeNumber = lDataTypeNumber + 1
      sDataTypeString = Trim(Str(lDataTypeNumber))
      
      sName = rsFieldList(fldName) 'frmMain.ARCom.GetFieldName(i)
      sType = rsFieldList(fldType) 'frmMain.ARCom.GetFieldDataType(i)
      lID = rsFieldList(fldARID) 'frmMain.ARCom.GetFieldID(i)
      sNodeName = sType
      
      Call AddTVItem(sFormNumberString & sNodeName, KEY_CHILD_PREFIX & sDataTypeString, sName, lID, ICON_FORMS, False)  ' , bSelected)
      rsFieldList.MoveNext
    Next i
    
  End If
  
  Set rsFieldList = Nothing
  
  Me.tvIDDisplay.Sorted = True
  tvIDDisplay.Refresh

End Sub


Private Sub LimitDataTypes(sParentName As String, sPrefix As String)

  Select Case sIDListMode
  Case AR_RUNIFTEXT
    Call AddTVItem(sParentName, sPrefix, AR_INTEGER, 0, ICON_FORMS)
    Call AddTVItem(sParentName, sPrefix, AR_REAL, 0, ICON_FORMS)
    Call AddTVItem(sParentName, sPrefix, AR_CHAR, 0, ICON_FORMS)
    Call AddTVItem(sParentName, sPrefix, AR_DIARY, 0, ICON_FORMS)
    Call AddTVItem(sParentName, sPrefix, AR_SELECTION, 0, ICON_FORMS)
    Call AddTVItem(sParentName, sPrefix, AR_DATE, 0, ICON_FORMS)
    Call AddTVItem(sParentName, sPrefix, AR_FIXEDDECIMAL, 0, ICON_FORMS)
    Call AddTVItem(sParentName, sPrefix, AR_ATTACHMENT, 0, ICON_FORMS)
    Call AddTVItem(sParentName, sPrefix, AR_TRIM, 0, ICON_FORMS)
    Call AddTVItem(sParentName, sPrefix, AR_CONTROL, 0, ICON_FORMS)
    Call AddTVItem(sParentName, sPrefix, AR_TABLE, 0, ICON_FORMS)
    Call AddTVItem(sParentName, sPrefix, AR_COLUMN, 0, ICON_FORMS)
    Call AddTVItem(sParentName, sPrefix, AR_PAGE, 0, ICON_FORMS)
    Call AddTVItem(sParentName, sPrefix, AR_PAGEHOLDER, 0, ICON_FORMS)
  Case AR_FOCUSFIELDNAME
    Call AddTVItem(sParentName, sPrefix, AR_INTEGER, 0, ICON_FORMS)
    Call AddTVItem(sParentName, sPrefix, AR_REAL, 0, ICON_FORMS)
    Call AddTVItem(sParentName, sPrefix, AR_CHAR, 0, ICON_FORMS)
    Call AddTVItem(sParentName, sPrefix, AR_DIARY, 0, ICON_FORMS)
    Call AddTVItem(sParentName, sPrefix, AR_SELECTION, 0, ICON_FORMS)
    Call AddTVItem(sParentName, sPrefix, AR_DATE, 0, ICON_FORMS)
    Call AddTVItem(sParentName, sPrefix, AR_FIXEDDECIMAL, 0, ICON_FORMS)
    Call AddTVItem(sParentName, sPrefix, AR_ATTACHMENT, 0, ICON_FORMS)
    Call AddTVItem(sParentName, sPrefix, AR_TRIM, 0, ICON_FORMS)
    'Call AddTVItem(sParentName, sPrefix, AR_CONTROL, 0, ICON_FORMS)
    Call AddTVItem(sParentName, sPrefix, AR_TABLE, 0, ICON_FORMS)
    'Call AddTVItem(sParentName, sPrefix, AR_COLUMN, 0, ICON_FORMS)
    Call AddTVItem(sParentName, sPrefix, AR_PAGE, 0, ICON_FORMS)
    'Call AddTVItem(sParentName, sPrefix, AR_PAGEHOLDER, 0, ICON_FORMS)
  Case AR_BUTTONNAME
    Call AddTVItem(sParentName, sPrefix, AR_CONTROL, 0, ICON_FORMS)
  End Select

End Sub


Public Sub AddForm(sFormName As String, bSelected As Boolean, sMode)
Dim sKeyPrefix As String

  sIDListMode = sMode

  lFormNumber = lFormNumber + 1
  sFormNumberString = Trim(Str(lFormNumber))
  
  sKeyPrefix = sFormNumberString
  
  Call AddTVItem("", "", sFormName, 0, ICON_FORMS)
  Call LimitDataTypes(sFormName, sKeyPrefix)
'  Call AddTVItem(sFormName, sKeyPrefix, AR_INTEGER, 0, ICON_FORMS)
'  Call AddTVItem(sFormName, sKeyPrefix, AR_REAL, 0, ICON_FORMS)
'  Call AddTVItem(sFormName, sKeyPrefix, AR_CHAR, 0, ICON_FORMS)
'  Call AddTVItem(sFormName, sKeyPrefix, AR_DIARY, 0, ICON_FORMS)
'  Call AddTVItem(sFormName, sKeyPrefix, AR_SELECTION, 0, ICON_FORMS)
'  Call AddTVItem(sFormName, sKeyPrefix, AR_DATE, 0, ICON_FORMS)
'  Call AddTVItem(sFormName, sKeyPrefix, AR_FIXEDDECIMAL, 0, ICON_FORMS)
'  Call AddTVItem(sFormName, sKeyPrefix, AR_ATTACHMENT, 0, ICON_FORMS)
'  Call AddTVItem(sFormName, sKeyPrefix, AR_TRIM, 0, ICON_FORMS)
'  Call AddTVItem(sFormName, sKeyPrefix, AR_CONTROL, 0, ICON_FORMS)
'  Call AddTVItem(sFormName, sKeyPrefix, AR_TABLE, 0, ICON_FORMS)
'  Call AddTVItem(sFormName, sKeyPrefix, AR_COLUMN, 0, ICON_FORMS)
'  Call AddTVItem(sFormName, sKeyPrefix, AR_PAGE, 0, ICON_FORMS)
'  Call AddTVItem(sFormName, sKeyPrefix, AR_PAGEHOLDER, 0, ICON_FORMS)
  
  If bSelected = True Then
    Me.tvIDDisplay.Nodes(sFormName).EnsureVisible
    Me.tvIDDisplay.Nodes(sFormName).Expanded = True
  End If
  
  AddAllFieldNames (sFormName)
  
End Sub


Private Sub AddTVItem(sParentText As String, sKeyPrefix As String, sDisplayText As String, lValue As Long, iIconIndex As Integer, Optional IsVisible As Boolean)
Dim nNode As Node
Dim sKeyName As String
Dim sParentKey As String

  sKeyName = sKeyPrefix & sDisplayText
  
  tvIDDisplay.Sorted = True
  If (Len(sParentText) > 0) Then
    sParentKey = sParentText
    Set nNode = tvIDDisplay.Nodes.Add(sParentKey, tvwChild, sKeyName, sDisplayText) ', iIconIndex)
  Else
    Set nNode = tvIDDisplay.Nodes.Add(, , sKeyName, sDisplayText) ', iIconIndex)
  End If

  nNode.tag = lValue
  nNode.Sorted = True
  
  If IsVisible = True Then
    nNode.EnsureVisible
  End If
  
End Sub


Private Sub RemoveTVItem(sKey As String)
Dim sKeyName As String

  tvIDDisplay.Nodes.Remove (sKey)

End Sub

Public Sub AddKeywords()

  Call AddTVItem("", "", "Keywords", 0, ICON_FORMS)
  Call AddTVItem("Keywords", "", "APPLICATION", 1, ICON_FORMS)
  Call AddTVItem("Keywords", "", "DATABASE", 1, ICON_FORMS)
  Call AddTVItem("Keywords", "", "DATE", 1, ICON_FORMS)
  Call AddTVItem("Keywords", "", "DEFAULT", 1, ICON_FORMS)
  Call AddTVItem("Keywords", "", "FIELDHELP", 1, ICON_FORMS)
  Call AddTVItem("Keywords", "", "GROUPS", 1, ICON_FORMS)
  Call AddTVItem("Keywords", "", "GUIDE", 1, ICON_FORMS)
  Call AddTVItem("Keywords", "", "GUIDETEXT", 1, ICON_FORMS)
  Call AddTVItem("Keywords", "", "HARDWARE", 1, ICON_FORMS)
  Call AddTVItem("Keywords", "", "LASTCOUNT", 1, ICON_FORMS)
  Call AddTVItem("Keywords", "", "LASTID", 1, ICON_FORMS)
  Call AddTVItem("Keywords", "", "NULL", 1, ICON_FORMS)
  Call AddTVItem("Keywords", "", "OPERATION", 1, ICON_FORMS)
  Call AddTVItem("Keywords", "", "OS", 1, ICON_FORMS)
  Call AddTVItem("Keywords", "", "SCHEMA", 1, ICON_FORMS)
  Call AddTVItem("Keywords", "", "SERVER", 1, ICON_FORMS)
  Call AddTVItem("Keywords", "", "TIME", 1, ICON_FORMS)
  Call AddTVItem("Keywords", "", "TIMESTAMP", 1, ICON_FORMS)
  Call AddTVItem("Keywords", "", "USER", 1, ICON_FORMS)
  Call AddTVItem("Keywords", "", "VERSION", 1, ICON_FORMS)
  Call AddTVItem("Keywords", "", "VUI", 1, ICON_FORMS)
  Call AddTVItem("Keywords", "", "WEEKDAY", 1, ICON_FORMS)

End Sub


Private Sub Form_Unload(Cancel As Integer)

  SaveSetting App.Title, "Settings", "IDWidth", Me.Width
  SaveSetting App.Title, "Settings", "IDHeight", Me.Height

End Sub


Private Sub tvIDDisplay_DblClick()
Dim sReturnValue As String
On Error Resume Next

  If tvIDDisplay.SelectedItem.tag <> 0 Then
  
    If tvIDDisplay.SelectedItem.Parent = "Keywords" Then
      sReturnValue = "$" & tvIDDisplay.SelectedItem.Text & "$"
    Else
      sReturnValue = tvIDDisplay.SelectedItem.Text
    End If
    
    frmMain.cboxValue.Clear
    frmMain.cboxValue.Text = sReturnValue
    frmMain.cboxValue.tag = tvIDDisplay.SelectedItem.tag
    frmMain.cboxValue.AddItem (sReturnValue)
    
    Unload Me
  End If

End Sub
