VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmQueryIO 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5670
   Icon            =   "frmQueryIO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox tbQueryName 
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   3120
      Width           =   5535
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   3540
      Width           =   5535
   End
   Begin MSComctlLib.ListView lvSavedQueries 
      Height          =   2775
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Query Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Search Type"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   2880
      Width           =   5535
   End
End
Attribute VB_Name = "frmQueryIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sFormMode As String
Private sName As String
Private lID As Long


Public Sub SetupForm(sMode As String)

  sFormMode = sMode
  
  Select Case sFormMode
    Case QUERY_OPEN
      Me.Caption = "Open Query"
      cmdGo.Caption = "Open"
    Case QUERY_SAVE
      Me.Caption = "Save Query"
      cmdGo.Caption = "Save"
    Case QUERY_DELETE
      Me.Caption = "Delete Query"
      cmdGo.Caption = "Delete"
  End Select

  ShowCurrentQueries
  
End Sub


Private Sub ShowCurrentQueries()
Dim i As Long
Dim sQueryName As String
Dim sQueryType As String
Dim liItem As ListItem
Dim lisiSubItem As ListSubItem

  For i = 1 To modDatabase.GetSavedQueryCount()
  
    Set liItem = lvSavedQueries.ListItems.Add()
    liItem.tag = modDatabase.GetSavedQueryID(i)
    liItem.Text = modDatabase.GetSavedQueryName(i)
    liItem.SubItems(1) = modDatabase.GetSavedQueryType(i)
    
  Next i

End Sub


Private Sub cmdGo_Click()
Dim i As Long
Dim sMSG As String
Dim bDuplicate As Boolean
Dim bCloseWindow As Boolean

  bCloseWindow = False
  
  If Len(tbQueryName.Text) = 0 Then
    If lvSavedQueries.ListItems.Count > 0 Then
      tbQueryName.Text = lvSavedQueries.SelectedItem.Text
    End If
  End If

  If Len(tbQueryName.Text) > 0 Then
    Select Case sFormMode
      Case QUERY_OPEN
        bCloseWindow = frmMain.OpenQuery(tbQueryName.Text)
        If bCloseWindow = False Then
          sMSG = "The query '" & tbQueryName.Text & "' could not be found."
          i = MsgBox(sMSG, vbOKOnly + vbCritical, "Error..")
        End If
      Case QUERY_SAVE
        For i = 1 To lvSavedQueries.ListItems.Count
          If lvSavedQueries.ListItems(i).Text = tbQueryName.Text Then
            lID = lvSavedQueries.ListItems(i).tag
            bDuplicate = True
          End If
        Next i
        
        i = vbYes
        If bDuplicate = True Then
          sMSG = "Are you sure you wish to overwrite query '" & tbQueryName.Text & "'?"
          i = MsgBox(sMSG, vbYesNo + vbQuestion, "Confirm Overwrite..")
        End If
        
        If i = vbYes Then
          bCloseWindow = frmMain.SaveQuery(tbQueryName.Text, bDuplicate, lID)
        End If
        
      Case QUERY_DELETE
        'To Do:  Delete query from RecentQuery list and/or AssignedQuery list
        sMSG = "Are you sure you wish to delete '" & tbQueryName.Text & "'?"
        i = MsgBox(sMSG, vbYesNo + vbQuestion, "Confirm Delete..")
        
        If i = vbYes Then
        
          If modDatabase.DeleteQuery(tbQueryName.Text) = False Then
            sMSG = "The query '" & tbQueryName.Text & "' was not found.  "
            sMSG = sMSG & "Please check the name and retry."
            i = MsgBox(sMSG, vbOKOnly + vbCritical, "Error..")
            bCloseWindow = False
          Else
            bCloseWindow = True
          End If
          
          'Make sure to remove from RecenQueriesList, AssignedQueryList and if current query is = to the
          'deleted one, need to reset that one as well
          frmMain.RemoveAssignedQuery (tbQueryName.Text)
          frmMain.RemoveQueryFromList (tbQueryName.Text)
          
        End If
        
    End Select
    
  End If
  
  If bCloseWindow = True Then
    Unload Me
  End If

End Sub


Private Sub Form_Load()
Dim lTempWidth As Long

  lTempWidth = Me.lvSavedQueries.Width / 2
  
  Me.lvSavedQueries.ColumnHeaders(1).Width = GetSetting(App.Title, "Settings", "IONameWidth", lTempWidth)
  Me.lvSavedQueries.ColumnHeaders(2).Width = GetSetting(App.Title, "Settings", "IOTypeWidth", lTempWidth)

End Sub

Private Sub Form_Unload(Cancel As Integer)

  SaveSetting App.Title, "Settings", "IONameWidth", lvSavedQueries.ColumnHeaders(1).Width
  SaveSetting App.Title, "Settings", "IOTypeWidth", lvSavedQueries.ColumnHeaders(2).Width

End Sub

Private Sub lvSavedQueries_Click()
  
  If Me.lvSavedQueries.ListItems.Count > 0 Then
    tbQueryName.Text = Me.lvSavedQueries.SelectedItem.Text
  End If

End Sub
