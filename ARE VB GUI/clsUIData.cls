VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUIData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'
''***************************************************
''clsUIData:  Interfaces with AR Explorer's UI.
''            This class will be responsible for
''            tracking and updating the data used
''            in the UI's TreeView and ListView.
''***************************************************
''use:  private UIData as new clsUIData
''***************************************************
''Default settings:
''ErrorLevel = Developer (Change for final release)
''***************************************************
''***************************************************
'
'
''***************
''Constants
''***************
'
'Const FORMCAPTION = "AR Explorer - "
'
'Const ROOT = "Servers"
'Const SERVER = "Server"    'This will be moved from being a constant to var
'Const ARFORMS = "Forms"
'Const ACTIVELINKS = "Active Links"
'Const FILTERS = "Filters"
'Const ESCALATIONS = "Escalations"
'Const GUIDES = "Guides"
'Const APPLICATIONS = "Applications"
'Const MENUS = "Menus"
'Const GROUPS = "Groups"
'
'Const KEY_PREFIX = "node"
'
'Const ICON_ROOT = 1
'Const ICON_SERVER = 8
'Const ICON_FORMS = 9
'Const ICON_ACTIVELINKS = 10
'Const ICON_FILTERS = 5
'Const ICON_ESCALATIONS = 6
'Const ICON_GUIDES = 7
'Const ICON_APPLICATIONS = 3
'Const ICON_MENUS = 4
'Const ICON_GROUPS = 2
'
'
''*******************
''Private Properties
''*******************
'Private frmUI As Form
'Private tvTree As TreeView
'Private lvList As ListView
'
'Private AR As New clsARIDLL
'
'Private sCheckedItems() As String    'List of Checked items
'
'Private iNumberOfServers As Integer   'Uhh... the number of servers =)
'Private sServerNames() As String      'Server name array, will be redim'd when we know
'                                      'how many there will be.
'
''*******************
''Public Accessors
''*******************
'
'Public Property Set SetForm(frmForm As Form)
'  Set frmUI = frmForm
'  Set tvTree = frmForm.tvTreeView
'  Set lvList = frmForm.lvListView
'End Property
'
''*******************
''Public Methods
''*******************
'
'Public Sub DefaultTreeView()
'
'  SetupTreeView
'
'End Sub
'
'
'Public Sub Connect()
'Dim sServer As String
'Dim sUser As String
'Dim sPassword As String
'Dim i As Integer
'Dim sMSG As String
'
'  If Not (AR.ConnectToARServer(sServer, sUser, sPassword)) Then
'    sMSG = "Connot connect to server '" & sServer & "' "
'    sMSG = sMSG & "with username '" & sUser & "'.  "
'    sMSG = sMSG & "Please verify username and password."
'    i = MsgBox(sMSG, vbOKOnly + vbInformation, "Error connecting to server.")
'  End If
'
'End Sub
'
'
'Public Sub GetCheckedItems()
'Dim i As Integer
'Dim iCount As Integer
'
'  iCount = 0
'
'  For i = 1 To tvTree.Nodes.Count
'
'    If tvTree.Nodes(i).Checked = True Then
'      iCount = iCount + 1
'      ReDim Preserve sCheckedItems(iCount)
'      sCheckedItems(iCount) = tvTree.Nodes(i).Text
'      frmTestOutput.tbOutput.Text = frmTestOutput.tbOutput.Text & vbCrLf & sCheckedItems(iCount)
'    End If
'
'  Next i
'
'  frmTestOutput.Show
'
'End Sub
'
'
''Public Sub AddServer(sServerName As String)
''
''  Call AddTVItem(ROOT, sServerName, ICON_SERVER)
''  Call AddTVItem(sServerName, ARFORMS, ICON_FORMS)
''  Call AddTVItem(sServerName, ACTIVELINKS, ICON_ACTIVELINKS)
''  Call AddTVItem(sServerName, FILTERS, ICON_FILTERS)
''  Call AddTVItem(sServerName, ESCALATIONS, ICON_ESCALATIONS)
''  Call AddTVItem(sServerName, GUIDES, ICON_GUIDES)
''  Call AddTVItem(sServerName, APPLICATIONS, ICON_APPLICATIONS)
''  Call AddTVItem(sServerName, MENUS, ICON_MENUS)
''  Call AddTVItem(sServerName, GROUPS, ICON_GROUPS)
''
''  tvTree.Nodes(KEY_PREFIX & sServerName).Selected = True
''
''
''End Sub
'
'
''*******************
''Private Methods
''*******************
'
''Private Sub AddTVItem(sParentText As String, sDisplayText As String, iIconIndex As Integer)
''Dim nNode As Node
''Dim sKeyName As String
''Dim sParentKey As String
''
''  sKeyName = KEY_PREFIX & sDisplayText
''
''  If (Len(sParentText) > 0) Then
''    sParentKey = KEY_PREFIX & sParentText
''    Set nNode = tvTree.Nodes.Add(sParentKey, tvwChild, sKeyName, sDisplayText, iIconIndex)
''  Else
''    Set nNode = tvTree.Nodes.Add(, , sKeyName, sDisplayText, iIconIndex)
''  End If
''
''  'Set nNode = tvTree.Nodes.Add(sParentKey, tvwChild, sKeyName, sDisplayText, iIconIndex)
''  nNode.EnsureVisible
''
''End Sub
'
''Private Sub RemoveTVItem(sDisplayText As String)
''Dim sKeyName As String
''
''  sKeyName = KEY_PREFIX & sDisplayText
''  tvTree.Nodes.Remove (sKeyName)
''
''End Sub
'
'
'
'Private Sub AddLVItem(sText As String)
'Dim liItem As ListItem
'
'  Set liItem = lvList.ListItems.Add()
'
'  liItem.Text = sText
'  'liItem.SubItems(1) = ""  'Not used, but who knows if we're gonna need it
'
'End Sub
'
'Private Sub RemoveLVItem()
'
'End Sub
'
'Private Sub RemoveAllLVItems()
'
'  lvList.ListItems.Clear
'
'End Sub
'
'
'Private Sub GetHighlightedItems()
'
'
'End Sub
'
'
''***************************************************
''SetupTreeView()
''***************************************************
''Private Sub SetupTreeView()
''Dim i As Integer
''
''  Call AddTVItem("", ROOT, ICON_ROOT)
''
''  Call AddServer("Server1")
''  'Call AddServer("Server2")
'''  Call AddTVItem(ROOT, SERVER, ICON_SERVER)
'''  Call AddTVItem(SERVER, ARFORMS, ICON_FORMS)
'''  Call AddTVItem(SERVER, ACTIVELINKS, ICON_ACTIVELINKS)
'''  Call AddTVItem(SERVER, FILTERS, ICON_FILTERS)
'''  Call AddTVItem(SERVER, ESCALATIONS, ICON_ESCALATIONS)
'''  Call AddTVItem(SERVER, GUIDES, ICON_GUIDES)
'''  Call AddTVItem(SERVER, APPLICATIONS, ICON_APPLICATIONS)
'''  Call AddTVItem(SERVER, MENUS, ICON_MENUS)
'''  Call AddTVItem(SERVER, GROUPS, ICON_GROUPS)
''
''  tvTree.Nodes(KEY_PREFIX & ROOT).Selected = True
''
''End Sub
'
