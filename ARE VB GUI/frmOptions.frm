VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   6105
   StartUpPosition =   1  'CenterOwner
   Tag             =   "1068"
   Begin VB.Frame Frame3 
      Caption         =   "Server..."
      Height          =   1035
      Left            =   120
      TabIndex        =   18
      Top             =   3900
      Width           =   5835
      Begin VB.TextBox tbServerName 
         Height          =   285
         Left            =   3420
         TabIndex        =   21
         Top             =   600
         Width           =   1935
      End
      Begin VB.CheckBox ckboxRememberServerName 
         Alignment       =   1  'Right Justify
         Caption         =   "Remember server name between sessions"
         Height          =   255
         Left            =   180
         TabIndex        =   19
         Top             =   300
         Value           =   1  'Checked
         Width           =   3435
      End
      Begin VB.Label lblOr 
         Caption         =   "Or always use:"
         Height          =   195
         Left            =   2220
         TabIndex        =   20
         Top             =   660
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Searching..."
      Height          =   2115
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   5835
      Begin MSComctlLib.Slider sldrNumberOfQueries 
         Height          =   255
         Left            =   1500
         TabIndex        =   22
         Top             =   1740
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   3
         Min             =   1
         Max             =   15
         SelStart        =   1
         TickStyle       =   1
         Value           =   1
      End
      Begin VB.CheckBox ckboxLockSearchDialog 
         Alignment       =   1  'Right Justify
         Caption         =   "Lock search dialog open:"
         Height          =   315
         Left            =   180
         TabIndex        =   17
         Top             =   1020
         Width           =   2595
      End
      Begin VB.ComboBox cboxDefaultQuery 
         Height          =   315
         ItemData        =   "frmOptions.frx":0442
         Left            =   2580
         List            =   "frmOptions.frx":0444
         TabIndex        =   16
         Text            =   "cboxDefaultQuery"
         Top             =   600
         Width           =   1875
      End
      Begin VB.ComboBox cboxDefaultQueryType 
         Height          =   315
         Left            =   2580
         TabIndex        =   14
         Text            =   "cboxDefaultQueryType"
         Top             =   240
         Width           =   1875
      End
      Begin VB.Label Label3 
         Caption         =   "Number of recent qeries to track [1-15]   (the higher the number the more memory needed)"
         Height          =   435
         Left            =   240
         TabIndex        =   23
         Top             =   1440
         Width           =   5415
      End
      Begin VB.Label Label2 
         Caption         =   "Default query to load on startup:"
         Height          =   255
         Left            =   180
         TabIndex        =   15
         Top             =   720
         Width           =   2355
      End
      Begin VB.Label Label1 
         Caption         =   "Default query type:"
         Height          =   255
         Left            =   180
         TabIndex        =   13
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cache..."
      Height          =   1035
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   5835
      Begin VB.CheckBox ckboxSaveCacheOnExit 
         Caption         =   "Save cache between AR Explorer sessions."
         Height          =   255
         Left            =   180
         TabIndex        =   11
         Top             =   300
         Value           =   1  'Checked
         Width           =   3615
      End
      Begin VB.CheckBox ckboxUpdateOnStartup 
         Caption         =   "Update cache on startup, (longer load time)."
         Height          =   255
         Left            =   180
         TabIndex        =   10
         Top             =   600
         Value           =   1  'Checked
         Width           =   3435
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2430
      TabIndex        =   0
      Tag             =   "1075"
      Top             =   5055
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3660
      TabIndex        =   1
      Tag             =   "1074"
      Top             =   5055
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   4860
      TabIndex        =   2
      Tag             =   "1073"
      Top             =   5055
      Width           =   1095
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   2022
         Left            =   505
         TabIndex        =   8
         Tag             =   "1072"
         Top             =   502
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   2022
         Left            =   406
         TabIndex        =   7
         Tag             =   "1071"
         Top             =   403
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   2022
         Left            =   307
         TabIndex        =   5
         Tag             =   "1070"
         Top             =   305
         Width           =   2033
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private bSaveCacheOnExit As Boolean
Private bUpdateCacheOnLoad As Boolean
Private sDefaultQueryType As String
Private sDefaultQueryName As String
Private bLockSearchDialogOpen As Boolean
Private bRememberServerName As Boolean
Private sDefaultServerName As String



Private Sub cboxDefaultQuery_Validate(Cancel As Boolean)
Dim i As Long

  cboxDefaultQueryType.Enabled = True
  
  For i = 0 To cboxDefaultQuery.ListCount - 1
  
    If cboxDefaultQuery.List(i) = cboxDefaultQuery.Text Then
    
      If i = 0 Then
        cboxDefaultQueryType.Text = sDefaultQueryType
      Else
        cboxDefaultQueryType.Text = modDatabase.GetSavedQueryType(i)
        cboxDefaultQueryType.Enabled = False
      End If
      
    End If
    
  Next i
  
  cmdApply.Enabled = True
  
End Sub


Private Sub cboxDefaultQueryType_Change()
  cmdApply.Enabled = True
End Sub

Private Sub ckboxLockSearchDialog_Click()

  cmdApply.Enabled = True
  
End Sub

Private Sub ckboxRememberServerName_Click()

  If ckboxRememberServerName.Value = Checked Then
    Me.tbServerName.Enabled = False
    lblOr.Enabled = False
  Else
    lblOr.Enabled = True
    Me.tbServerName.Enabled = True
  End If
  
  cmdApply.Enabled = True

End Sub

Private Sub ckboxSaveCacheOnExit_Click()

  If ckboxSaveCacheOnExit.Value = Unchecked Then
    ckboxUpdateOnStartup.Value = Checked
    bUpdateCacheOnLoad = True
    ckboxUpdateOnStartup.Enabled = False
    bSaveCacheOnExit = False
  Else
    ckboxUpdateOnStartup.Enabled = True
    bSaveCacheOnExit = True
  End If
  
  cmdApply.Enabled = True

End Sub

Private Sub ckboxUpdateOnStartup_Click()
  
  If ckboxUpdateOnStartup.Value = Checked Then
    bUpdateCacheOnLoad = True
  Else
    bUpdateCacheOnLoad = False
  End If
  
  cmdApply.Enabled = True
  
End Sub

Private Sub Form_Load()
Dim i As Long
Dim iMatchingQuery As Integer
Dim sQueryName As String

  LoadResStrings Me
  
  bSaveCacheOnExit = GetSetting(App.Title, "Options", "SaveCacheOnExit", True)
  bUpdateCacheOnLoad = GetSetting(App.Title, "Options", "UpdateCacheOnLoad", True)
  sDefaultQueryType = GetSetting(App.Title, "Options", "DefaultQueryType", TYPE_AL)
  sDefaultQueryName = GetSetting(App.Title, "Options", "DefaultQueryName", sEmptyString)
  bLockSearchDialogOpen = GetSetting(App.Title, "Options", "LockedSearchDialog", False)
  bRememberServerName = GetSetting(App.Title, "Options", "RememberServerName", True)
  sDefaultServerName = GetSetting(App.Title, "Options", "DefaultServerName", "")
  
  If bSaveCacheOnExit = True Then
    ckboxSaveCacheOnExit.Value = Checked
    ckboxUpdateOnStartup.Enabled = True
  Else
    ckboxSaveCacheOnExit.Value = Unchecked
    ckboxUpdateOnStartup.Value = Checked
    ckboxUpdateOnStartup.Enabled = False
  End If
  
  If bUpdateCacheOnLoad = True Then
    ckboxUpdateOnStartup.Value = Checked
  Else
    ckboxUpdateOnStartup.Value = Unchecked
  End If
    
  cboxDefaultQueryType.AddItem (TYPE_AL)
  cboxDefaultQueryType.AddItem (TYPE_FILTER)
  cboxDefaultQueryType.AddItem (TYPE_FIELD)
  cboxDefaultQueryType.Text = sDefaultQueryType
  
  cboxDefaultQuery.AddItem (sEmptyString)
  For i = 1 To modDatabase.GetSavedQueryCount()
    sQueryName = modDatabase.GetSavedQueryName(i)
    If sQueryName = sDefaultQueryName Then
      iMatchingQuery = i
      cboxDefaultQueryType.Text = modDatabase.GetSavedQueryType(i)
    End If
    cboxDefaultQuery.AddItem sQueryName
  Next i
  cboxDefaultQuery.Text = sDefaultQueryName
  
  Me.sldrNumberOfQueries.Value = GetSetting(App.Title, "Settings", "RecentQueries", 5)
      
  If bLockSearchDialogOpen = True Then
    ckboxLockSearchDialog.Value = Checked
  Else
    ckboxLockSearchDialog.Value = Unchecked
  End If
  
  If bRememberServerName = True Then
    ckboxRememberServerName.Value = Checked
    tbServerName.Enabled = False
    lblOr.Enabled = False
  Else
    ckboxRememberServerName.Value = Unchecked
    lblOr.Enabled = True
    tbServerName.Enabled = True
    tbServerName.Text = sDefaultServerName
  End If
  
  Me.cmdApply.Enabled = False
  
End Sub


Private Sub SaveChanges()

  If ckboxSaveCacheOnExit.Value = Checked Then
    bSaveCacheOnExit = True
  Else
    bSaveCacheOnExit = False
  End If
  
  If ckboxUpdateOnStartup.Value = Checked Then
    bUpdateCacheOnLoad = True
  Else
    bUpdateCacheOnLoad = False
  End If
  
  sDefaultQueryType = cboxDefaultQueryType.Text
  sDefaultQueryName = cboxDefaultQuery.Text
  
  If ckboxLockSearchDialog.Value = Checked Then
    bLockSearchDialogOpen = True
  Else
    bLockSearchDialogOpen = False
  End If
  
  If ckboxRememberServerName.Value = Checked Then
    bRememberServerName = True
  Else
    bRememberServerName = False
  End If

  sDefaultServerName = tbServerName.Text
  
  SaveSetting App.Title, "Options", "SaveCacheOnExit", bSaveCacheOnExit
  SaveSetting App.Title, "Options", "UpdateCacheOnLoad", bUpdateCacheOnLoad
  SaveSetting App.Title, "Options", "DefaultQueryType", sDefaultQueryType
  SaveSetting App.Title, "Options", "DefaultQueryName", sDefaultQueryName
  SaveSetting App.Title, "Options", "LockedSearchDialog", bLockSearchDialogOpen
  SaveSetting App.Title, "Options", "RememberServerName", bRememberServerName
  SaveSetting App.Title, "Options", "DefaultServerName", sDefaultServerName
  SaveSetting App.Title, "Settings", "RecentQueries", sldrNumberOfQueries.Value
  cmdApply.Enabled = False
  
End Sub


Private Sub cmdApply_Click()

  SaveChanges
    
End Sub


Private Sub cmdCancel_Click()

  Unload Me
  
End Sub


Private Sub cmdOk_Click()

  SaveChanges
  Unload Me
  
End Sub


Private Sub sldrNumberOfQueries_Click()
  cmdApply.Enabled = True
End Sub

'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Dim i As Integer
'
'  i = tbsOptions.SelectedItem.Index
'  'handle ctrl+tab to move to the next tab
'  If (Shift And 3) = 2 And KeyCode = vbKeyTab Then
'    If i = tbsOptions.Tabs.Count Then
'      'last tab so we need to wrap to tab 1
'      Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
'    Else
'      'increment the tab
'      Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
'    End If
'  ElseIf (Shift And 3) = 3 And KeyCode = vbKeyTab Then
'    If i = 1 Then
'      'last tab so we need to wrap to tab 1
'      Set tbsOptions.SelectedItem = tbsOptions.Tabs(tbsOptions.Tabs.Count)
'    Else
'      'increment the tab
'      Set tbsOptions.SelectedItem = tbsOptions.Tabs(i - 1)
'    End If
'  End If
'
'End Sub

Private Sub tbServerName_Change()
  cmdApply.Enabled = True
End Sub
