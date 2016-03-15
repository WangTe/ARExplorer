VERSION 5.00
Begin VB.Form frmAssignQuery 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Assign Query"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   Icon            =   "frmAssgnQuery.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   315
      Left            =   3180
      TabIndex        =   10
      Top             =   1980
      Width           =   1215
   End
   Begin VB.ComboBox cboxSaved5 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1560
      Width           =   2955
   End
   Begin VB.ComboBox cboxSaved4 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1200
      Width           =   2955
   End
   Begin VB.ComboBox cboxSaved3 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   840
      Width           =   2955
   End
   Begin VB.ComboBox cboxSaved2 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   480
      Width           =   2955
   End
   Begin VB.ComboBox cboxSaved1 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   120
      Width           =   2955
   End
   Begin VB.Label Label5 
      Caption         =   "Saved Query #4:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1260
      Width           =   1275
   End
   Begin VB.Label Label4 
      Caption         =   "Saved Query #2:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   540
      Width           =   1275
   End
   Begin VB.Label Label3 
      Caption         =   "Saved Query #5:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1620
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   "Saved Query #1:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Saved Query #3:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   900
      Width           =   1275
   End
End
Attribute VB_Name = "frmAssignQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private iFieldFocus As Integer

Public Sub LoadQueries(iFocusNumber As Integer)
Dim i As Long

  iFieldFocus = iFocusNumber

  cboxSaved1.AddItem (sEmptyString)
  cboxSaved2.AddItem (sEmptyString)
  cboxSaved3.AddItem (sEmptyString)
  cboxSaved4.AddItem (sEmptyString)
  cboxSaved5.AddItem (sEmptyString)
  
  For i = 1 To modDatabase.GetSavedQueryCount()
  
    cboxSaved1.AddItem (modDatabase.GetSavedQueryName(i))
    cboxSaved2.AddItem (modDatabase.GetSavedQueryName(i))
    cboxSaved3.AddItem (modDatabase.GetSavedQueryName(i))
    cboxSaved4.AddItem (modDatabase.GetSavedQueryName(i))
    cboxSaved5.AddItem (modDatabase.GetSavedQueryName(i))
    
  Next i

  cboxSaved1.Text = frmMain.AssignedQueries.Item(1).SaveName
  cboxSaved2.Text = frmMain.AssignedQueries.Item(2).SaveName
  cboxSaved3.Text = frmMain.AssignedQueries.Item(3).SaveName
  cboxSaved4.Text = frmMain.AssignedQueries.Item(4).SaveName
  cboxSaved5.Text = frmMain.AssignedQueries.Item(5).SaveName
  
End Sub

Private Sub cmdOK_Click()
Dim i As Integer
  i = frmMain.AssignedQueries.Count
  
  frmMain.AssignedQueries(1).SaveName = cboxSaved1.Text
  frmMain.AssignedQueries(2).SaveName = cboxSaved2.Text
  frmMain.AssignedQueries(3).SaveName = cboxSaved3.Text
  frmMain.AssignedQueries(4).SaveName = cboxSaved4.Text
  frmMain.AssignedQueries(5).SaveName = cboxSaved5.Text
  
  frmMain.LoadAssignedQueries
  
  Unload Me

End Sub

Public Sub SetControlFocus()

  Select Case iFieldFocus
  Case 0
    cboxSaved1.SetFocus
  Case 1
    cboxSaved1.SetFocus
  Case 2
    cboxSaved2.SetFocus
  Case 3
    cboxSaved3.SetFocus
  Case 4
    cboxSaved4.SetFocus
  Case 5
    cboxSaved5.SetFocus
  End Select

End Sub


Private Sub Form_Activate()
  SetControlFocus
End Sub

