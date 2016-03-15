VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3630
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1063"
   Begin VB.TextBox txtServerName 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Text            =   "sever1"
      Top             =   900
      Width           =   2295
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   2100
      TabIndex        =   5
      Tag             =   "1067"
      Top             =   1320
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   495
      TabIndex        =   4
      Tag             =   "1066"
      Top             =   1320
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   60
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "&Server:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   255
      Index           =   1
      Left            =   105
      TabIndex        =   0
      Tag             =   "1065"
      Top             =   540
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   255
      Index           =   0
      Left            =   105
      TabIndex        =   6
      Tag             =   "1064"
      Top             =   120
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long


Public OK As Boolean  'if = true then logged in properly
Public Cancel As Boolean 'if = true then cancel login and quit app


Private Sub Form_Activate()
  If txtUserName.Text = vbNullString Then
    txtUserName.SetFocus
  Else
    txtPassword.SetFocus
  End If

End Sub

Private Sub Form_Load()
  Dim sBuffer As String
  Dim lSize As Long
  Dim sDefaultName As String

  LoadResStrings Me

  Cancel = False

  sBuffer = Space$(255)
  lSize = Len(sBuffer)
  Call GetUserName(sBuffer, lSize)
  If lSize > 0 Then
    sDefaultName = Left$(sBuffer, lSize)
  Else
    sDefaultName = vbNullString
  End If
  
  txtServerName.Text = GetSetting(App.Title, "Settings", "LastServerUsed", "<Type Server Name Here>")
  txtUserName.Text = GetSetting(App.Title, "Settings", "LastUserNameUsed", sDefaultName)
    
End Sub


Private Sub cmdCancel_Click()
    Cancel = True
    Me.Hide
End Sub


Private Sub cmdOk_Click()

  OK = False
  
  If frmMain.Connect(txtServerName.Text, txtUserName.Text, txtPassword.Text) Then
    OK = True
    SaveSetting App.Title, "Settings", "LastServerUsed", txtServerName.Text
    SaveSetting App.Title, "Settings", "LastUserNameUsed", txtUserName.Text
  End If
  
  Me.Hide
  
End Sub


Private Sub txtPassword_GotFocus()
  txtPassword.SelStart = 0
  txtPassword.SelLength = Len(txtPassword.Text)
End Sub

Private Sub txtServerName_GotFocus()
  txtServerName.SelStart = 0
  txtServerName.SelLength = Len(txtServerName.Text)
End Sub

Private Sub txtUserName_GotFocus()
  txtUserName.SelStart = 0
  txtUserName.SelLength = Len(txtUserName)
End Sub
