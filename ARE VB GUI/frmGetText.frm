VERSION 5.00
Begin VB.Form frmGetText 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "&Ok"
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      Top             =   2880
      Width           =   1275
   End
   Begin VB.TextBox tbText 
      Height          =   1395
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   4635
   End
End
Attribute VB_Name = "frmGetText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public iCalledBy As Integer


Private Sub cmdOk_Click()

  Select Case iCalledBy
    Case ChangeHistory
      frmMain.tbChangeHistory.Text = tbText.Text
    Case HelpText
      frmMain.tbHelpText.Text = tbText.Text
  End Select
  
  Unload Me

End Sub


Private Sub Form_Load()
'Load Size properties
  Me.Width = GetSetting(App.Title, "Settings", "TextWidth", 4770)
  Me.Height = GetSetting(App.Title, "Settings", "TextHeight", 3495)

End Sub


Private Sub Form_Resize()
  
  SizeTextBox
  
End Sub


Private Sub Form_Unload(Cancel As Integer)
'Save Size properties
  SaveSetting App.Title, "Settings", "TextWidth", Me.Width
  SaveSetting App.Title, "Settings", "TextHeight", Me.Height

End Sub


Private Sub SizeTextBox()

  If Me.Height <= ((cmdOk.Height + cmdOk.Height + cmdOk.Height) + 160) Then
    Me.Height = (cmdOk.Height + cmdOk.Height + cmdOk.Height) + 160
  End If
  If Me.Width <= (cmdOk.Width + 160) Then
    Me.Width = cmdOk.Width + 160
  End If
  cmdOk.Top = Me.ScaleHeight - (cmdOk.Height + 20)
  cmdOk.Left = Me.Width - (cmdOk.Width + 120)
  tbText.Top = Me.ScaleTop
  tbText.Left = Me.ScaleLeft
  tbText.Width = Me.ScaleWidth
  tbText.Height = cmdOk.Top - 40

End Sub
