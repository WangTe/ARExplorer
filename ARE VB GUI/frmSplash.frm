VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4740
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraMainFrame 
      Height          =   4050
      Left            =   45
      TabIndex        =   0
      Top             =   -15
      Width           =   4680
      Begin VB.PictureBox picLogo 
         Height          =   3075
         Left            =   60
         Picture         =   "frmSplash.frx":0442
         ScaleHeight     =   3015
         ScaleWidth      =   4515
         TabIndex        =   1
         Top             =   120
         Width           =   4575
      End
      Begin VB.Label lblPlatform 
         AutoSize        =   -1  'True
         Caption         =   "Windows 9x, NT, 2000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   5
         Tag             =   "1059"
         Top             =   3300
         Width           =   2775
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         Caption         =   "Version 1.3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   180
         TabIndex        =   4
         Tag             =   "1058"
         Top             =   3660
         Width           =   1380
      End
      Begin VB.Label lblCompany 
         Caption         =   "AR Accelerators, Inc."
         Height          =   255
         Left            =   3120
         TabIndex        =   3
         Tag             =   "1056"
         Top             =   3660
         Width           =   1515
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright:"
         Height          =   255
         Left            =   3120
         TabIndex        =   2
         Tag             =   "1055"
         Top             =   3360
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    LoadResStrings Me
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    'lblProductName.Caption = App.Title
End Sub

Private Sub fraMainFrame_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Unload Me

End Sub

