VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSearch 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Search..."
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7080
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5760
      TabIndex        =   12
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox tbSearchText 
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   6855
   End
   Begin VB.Frame Frame2 
      Caption         =   "&Active Link Name Search"
      Height          =   1575
      Left            =   3600
      TabIndex        =   4
      Top             =   120
      Width           =   3375
      Begin VB.OptionButton optEnds 
         Caption         =   "Ends with."
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton optContains 
         Caption         =   "Contains."
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton optBegins 
         Caption         =   "Begins with."
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox tbALName 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Form Last Modified Time/Date"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin MSComCtl2.DTPicker timeSelectTime 
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   24510466
         CurrentDate     =   36850
      End
      Begin VB.CommandButton cmdPickDate 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   2880
         Picture         =   "frmSearch.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Open a Calendar View."
         Top             =   360
         Width           =   315
      End
      Begin VB.OptionButton optLess 
         Caption         =   "Before this date."
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   2295
      End
      Begin VB.OptionButton optGreater 
         Caption         =   "This date and after."
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox tbModTime 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "m/d/yy h:nn AM/PM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2715
      End
   End
   Begin VB.Label Label1 
      Caption         =   "S&earch for string in Active Link Run-If line:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   3375
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Const DATEGREATER = 2
'Const DATELESS = 4
'
'Const ALNAMESTARTS = 2
'Const ALNAMECONTAINS = 4
'Const ALNAMEENDS = 3
'
'Const DEFAULTTIME = "12:00:00 AM"

'Private sALName As String
'Private iALNameOperator As Integer
'
'Private sModTime As String
'Private iModTimeOperator As Integer
'
'Private sALRunIfText As String
'
'Public sDate As String
'Public sTime As String



'Private Sub calDateSelect_DateClick(ByVal DateClicked As Date)
'
'  tbModTime.Text = calDateSelect.Value
'  calDateSelect.Enabled = False
'  calDateSelect.Visible = False
'
'End Sub

Private Sub cmdCancel_Click()

  Unload Me

End Sub


'Private Sub cmdPickDate_Click()
'
'  'Why do I have to add the height twice?
'  frmDatePicker.Top = frmSearch.Top + tbModTime.Top + Frame1.Top + tbModTime.Height * 2 + 20
'  frmDatePicker.Left = frmSearch.Left + tbModTime.Left + Frame1.Left + 20
'
'  frmDatePicker.calToDate.Value = Format(Now())
'
'  frmDatePicker.Show vbModal
'
'  UpdateDateTime
'
'  optGreater.SetFocus
'
'End Sub


'Private Sub cmdSearch_Click()
'Dim i As Integer
'Dim sMSG As String
'
'  If Len(tbSearchText.Text) > 0 Then
'
'    sALRunIfText = tbSearchText.Text
'
'    If Len(tbModTime.Text) > 0 Then
'      sModTime = tbModTime.Text
'
'      If optGreater.Value = True Then
'        iModTimeOperator = DATEGREATER
'      ElseIf optLess.Value = True Then
'        iModTimeOperator = DATELESS
'      Else
'        iModTimeOperator = DATEGREATER
'      End If
'
'    End If
'
'    If Len(tbALName.Text) > 0 Then
'      sALName = tbALName.Text
'
'      If optBegins.Value = True Then
'        iALNameOperator = ALNAMESTARTS
'      ElseIf optContains.Value = True Then
'        iALNameOperator = ALNAMECONTAINS
'      ElseIf optEnds.Value = True Then
'        iALNameOperator = ALNAMEENDS
'      Else
'        iALNameOperator = ALNAMESTARTS
'      End If
'
'    End If
'
'    Unload Me
'
'    Call frmMain.SearchServer(sALName, iALNameOperator, sModTime, iModTimeOperator, sALRunIfText)
'
'  Else
'    sMSG = "You must enter text to search."
'    i = MsgBox(sMSG, vbOKOnly + vbInformation)
'  End If
'
'End Sub


'Private Sub Form_Load()
'
'  sDate = ""
'  sTime = ""
'  timeSelectTime.Value = Now()
'
'End Sub


'Private Sub UpdateDateTime()
'
'  'if they haven't specified a date, assume they want today
'  If Not Len(sDate) > 0 Then
'    sDate = Left(Now(), 8)
'  End If
'
'  'if they haven't specifed a time, assume they want 12am (The default time).
'  If Not Len(sTime) > 0 Then
'    sTime = DEFAULTTIME
'  End If
'
'  tbModTime.Text = sDate & " " & sTime
'
'End Sub


'Private Sub timeSelectTime_Validate(Cancel As Boolean)
'
'  sTime = Format(timeSelectTime.Hour & ":" & timeSelectTime.Minute & ":" & timeSelectTime.Second)
'  UpdateDateTime
'
'End Sub
Private Sub timeSelectTime_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub
