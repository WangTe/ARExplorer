VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDatePicker 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2745
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   2730
   ControlBox      =   0   'False
   Icon            =   "frmDatePicker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   2730
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "&Ok"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   315
      Left            =   2340
      TabIndex        =   2
      Top             =   2400
      Width           =   375
   End
   Begin MSComCtl2.DTPicker dtTime 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   2400
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   556
      _Version        =   393216
      Format          =   24641538
      CurrentDate     =   36651
   End
   Begin MSComCtl2.MonthView calToDate 
      Height          =   2370
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   1
      MultiSelect     =   -1  'True
      StartOfWeek     =   24510465
      CurrentDate     =   36848
   End
End
Attribute VB_Name = "frmDatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub calToDate_Click()

  Me.cmdOk.SetFocus

End Sub

Private Sub calToDate_LostFocus()

  Me.cmdOk.SetFocus

End Sub


Private Sub cmdOk_Click()
Dim iSubValue As Integer
Dim sMonth As String
Dim sDay As String
Dim sYear As String
Dim sAmPm As String
Dim sMinute As String
Dim sSecond As String
Dim sText As String

  sMonth = calToDate.Month
  
  If calToDate.Day < 10 Then
    sDay = "0" & calToDate.Day
  Else
    sDay = calToDate.Day
  End If
  
  sYear = calToDate.Year
  
  If dtTime.Hour > 12 Then
    iSubValue = 12
    sAmPm = "PM"
  Else
    iSubValue = 0
    sAmPm = "AM"
  End If
  
  If dtTime.Minute < 10 Then
    sMinute = "0" & dtTime.Minute
  Else
    sMinute = dtTime.Minute
  End If
  
  If dtTime.Second < 10 Then
    sSecond = "0" & dtTime.Second
  Else
    sSecond = dtTime.Second
  End If
  
  sText = sMonth & "/" & sDay & "/" & sYear & " " & _
    Str(dtTime.Hour - iSubValue) & ":" & sMinute & ":" & _
    sSecond & " " & sAmPm
  
  frmMain.cboxValue.Clear
  frmMain.cboxValue.Text = sText
  frmMain.cboxValue.AddItem (sText)
  'frmMain.cboxValue.Text = Format(calToDate.Value) & " " &
  'frmMain.cboxValue.Text = sMonth & "/" & sDay & "/" & sYear & " " & _
    Str(dtTime.Hour - iSubValue) & ":" & sMinute & ":" & _
    sSecond & " " & sAmPm
  Unload Me

End Sub


Private Sub dtTime_LostFocus()

  Me.cmdOk.SetFocus
  
End Sub


Private Sub Form_Activate()

  Me.cmdOk.SetFocus

End Sub

Private Sub Form_Load()
Dim dDate As Date
Dim sMonthYear As String
Dim sMinHourSecond As String
Dim iAddValue As Integer

  If Len(frmMain.cboxValue.Text) > 0 Then
    dDate = frmMain.cboxValue.Text
'    If Right(dDate, 2) = "PM" Then
'      iAddValue = 12
'    Else
'      iAddValue = 0
'    End If
    sMonthYear = Trim(Str(Month(dDate))) & "/" & Trim(Str(Day(dDate))) & "/" & Trim(Str(Year(dDate)))
    calToDate.Value = sMonthYear
    'calToDate.Month = Month(dDate)
    'calToDate.Day = Day(dDate)
    'calToDate.Year = Year(dDate)
    sMinHourSecond = Trim(Str(Hour(dDate))) & ":" & Trim(Str(Minute(dDate))) & ":" & Trim(Str(Second(dDate)))
    'dtTime.Hour = Hour(dDate)
    'dtTime.Minute = Minute(dDate)
    'dtTime.Second = Second(dDate)
    dtTime.Value = sMinHourSecond
  Else
    dtTime.Value = Now()
    Me.calToDate.Value = Now()
  End If
  
End Sub
