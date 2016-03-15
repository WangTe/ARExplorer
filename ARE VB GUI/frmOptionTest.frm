VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOptionTest 
   Caption         =   "Option Test"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   7950
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picActiveLink 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      ScaleHeight     =   1575
      ScaleWidth      =   7695
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   7695
      Begin VB.TextBox tbSearchText 
         Height          =   285
         Left            =   0
         TabIndex        =   15
         Top             =   1200
         Width           =   6855
      End
      Begin VB.Frame Frame2 
         Caption         =   "&Active Link Name Search"
         Height          =   855
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   6855
         Begin VB.OptionButton optEnds 
            Caption         =   "Ends with."
            Height          =   255
            Left            =   5040
            TabIndex        =   14
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optContains 
            Caption         =   "Contains."
            Height          =   255
            Left            =   3480
            TabIndex        =   13
            Top             =   480
            Width           =   1335
         End
         Begin VB.OptionButton optBegins 
            Caption         =   "Begins with."
            Height          =   255
            Left            =   3480
            TabIndex        =   12
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox tbALName 
            Height          =   285
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   3135
         End
      End
      Begin VB.Label Label1 
         Caption         =   "S&earch for string in Active Link Run-If line:"
         Height          =   255
         Left            =   0
         TabIndex        =   16
         Top             =   960
         Width           =   3375
      End
   End
   Begin VB.PictureBox picForms 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      ScaleHeight     =   1575
      ScaleWidth      =   7695
      TabIndex        =   2
      Top             =   600
      Width           =   7695
      Begin VB.TextBox Text1 
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
         Left            =   0
         TabIndex        =   6
         Top             =   240
         Width           =   2355
      End
      Begin VB.OptionButton Option2 
         Caption         =   "This date and after."
         Height          =   195
         Left            =   3000
         TabIndex        =   5
         Top             =   120
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Before this date."
         Height          =   195
         Left            =   3000
         TabIndex        =   4
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   2400
         Picture         =   "frmOptionTest.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Open a Calendar View."
         Top             =   240
         Width           =   315
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   0
         TabIndex        =   7
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   24510466
         CurrentDate     =   36850
      End
      Begin VB.Label Label2 
         Caption         =   "Last Modified Date/Time:"
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   2415
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2415
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   4260
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Forms.."
            Object.ToolTipText     =   "Search for Forms"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Active Links.."
            Object.ToolTipText     =   "Search for Active Links"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   6600
      TabIndex        =   0
      Top             =   2760
      Width           =   1335
   End
End
Attribute VB_Name = "frmOptionTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iSelectedTab As Integer

Private Sub cmdQuit_Click()

  Unload Me

End Sub


Private Sub Form_Load()

  AdjustTabs

End Sub

Private Sub AdjustTabs()

  If TabStrip1.SelectedItem.Index = iSelectedTab Then
    Exit Sub
  End If
    
  iSelectedTab = TabStrip1.SelectedItem.Index
  
  Select Case iSelectedTab
    Case 1
      Me.picActiveLink.Enabled = False
      Me.picActiveLink.Visible = False
      Me.picForms.Enabled = True
      Me.picForms.Visible = True
    Case 2
      Me.picForms.Enabled = False
      Me.picForms.Visible = False
      Me.picActiveLink.Enabled = True
      Me.picActiveLink.Visible = True
  End Select

End Sub

Private Sub Tabstrip1_Click()

  AdjustTabs
  iSelectedTab = TabStrip1.SelectedItem.Index
    
End Sub
