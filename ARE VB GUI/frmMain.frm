VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "AR Explorer "
   ClientHeight    =   8280
   ClientLeft      =   1215
   ClientTop       =   960
   ClientWidth     =   10065
   ClipControls    =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   10065
   Begin MSComctlLib.ListView lvListViewActions 
      Height          =   1815
      Left            =   8160
      TabIndex        =   53
      Top             =   2640
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   3201
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.PictureBox picModify 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2475
      Left            =   0
      ScaleHeight     =   2475
      ScaleWidth      =   7815
      TabIndex        =   23
      Top             =   5160
      Visible         =   0   'False
      Width           =   7815
      Begin TabDlg.SSTab tabModify 
         Height          =   2175
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   7725
         _ExtentX        =   13626
         _ExtentY        =   3836
         _Version        =   393216
         Style           =   1
         Tabs            =   5
         Tab             =   4
         TabsPerRow      =   6
         TabHeight       =   520
         TabCaption(0)   =   "General"
         TabPicture(0)   =   "frmMain.frx":0442
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame4"
         Tab(0).Control(1)=   "tbFieldLabel"
         Tab(0).Control(2)=   "tbDatabaseName"
         Tab(0).Control(3)=   "Frame3"
         Tab(0).Control(4)=   "Label17"
         Tab(0).Control(5)=   "Label16"
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Basic"
         TabPicture(1)   =   "frmMain.frx":045E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "tbExecutionOrder"
         Tab(1).Control(1)=   "cboxEnabled"
         Tab(1).Control(2)=   "tbName"
         Tab(1).Control(3)=   "UpDown2"
         Tab(1).Control(4)=   "Label12"
         Tab(1).Control(5)=   "Label11"
         Tab(1).Control(6)=   "Label10"
         Tab(1).ControlCount=   7
         TabCaption(2)   =   "Permissions"
         TabPicture(2)   =   "frmMain.frx":047A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "lboxNoAccess"
         Tab(2).Control(1)=   "lboxAccess"
         Tab(2).Control(2)=   "cmdGiveAccess"
         Tab(2).Control(3)=   "cmdRemoveAccess"
         Tab(2).Control(4)=   "cboxPermissionType"
         Tab(2).Control(5)=   "Label13"
         Tab(2).Control(6)=   "Label14"
         Tab(2).Control(7)=   "Label15"
         Tab(2).ControlCount=   8
         TabCaption(3)   =   "Change History"
         TabPicture(3)   =   "frmMain.frx":0496
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "tbChangeHistory"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Help Text"
         TabPicture(4)   =   "frmMain.frx":04B2
         Tab(4).ControlEnabled=   -1  'True
         Tab(4).Control(0)=   "tbHelpText"
         Tab(4).Control(0).Enabled=   0   'False
         Tab(4).Control(1)=   "cboxHelpText"
         Tab(4).Control(1).Enabled=   0   'False
         Tab(4).ControlCount=   2
         Begin VB.ComboBox cboxHelpText 
            Height          =   315
            ItemData        =   "frmMain.frx":04CE
            Left            =   5400
            List            =   "frmMain.frx":04DB
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   1740
            Width           =   2175
         End
         Begin VB.Frame Frame4 
            Caption         =   "Database Tab: Entry Mode"
            Height          =   735
            Left            =   -74520
            TabIndex        =   48
            Top             =   540
            Width           =   3195
            Begin VB.ComboBox cboxEntryMode 
               Height          =   315
               ItemData        =   "frmMain.frx":04FF
               Left            =   480
               List            =   "frmMain.frx":050C
               Style           =   2  'Dropdown List
               TabIndex        =   49
               Top             =   240
               Width           =   2175
            End
         End
         Begin VB.TextBox tbFieldLabel 
            Height          =   285
            Left            =   -74940
            TabIndex        =   47
            Top             =   1620
            Width           =   3675
         End
         Begin VB.TextBox tbDatabaseName 
            Height          =   285
            Left            =   -71100
            TabIndex        =   46
            Top             =   1620
            Width           =   3675
         End
         Begin VB.Frame Frame3 
            Caption         =   "Permissions Tab: Allow Any User to Submit"
            Height          =   735
            Left            =   -70920
            TabIndex        =   44
            Top             =   540
            Width           =   3495
            Begin VB.ComboBox cboxSubmit 
               Height          =   315
               ItemData        =   "frmMain.frx":0531
               Left            =   300
               List            =   "frmMain.frx":053E
               Style           =   2  'Dropdown List
               TabIndex        =   45
               Top             =   240
               Width           =   2835
            End
         End
         Begin VB.TextBox tbExecutionOrder 
            Height          =   285
            Left            =   -74880
            TabIndex        =   39
            Top             =   1440
            Width           =   1095
         End
         Begin VB.ComboBox cboxEnabled 
            Height          =   315
            ItemData        =   "frmMain.frx":0558
            Left            =   -72360
            List            =   "frmMain.frx":0565
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   1440
            Width           =   1635
         End
         Begin VB.TextBox tbName 
            Height          =   285
            Left            =   -74880
            TabIndex        =   37
            Top             =   720
            Width           =   4095
         End
         Begin VB.ListBox lboxNoAccess 
            Height          =   1425
            Left            =   -74880
            MultiSelect     =   2  'Extended
            Sorted          =   -1  'True
            TabIndex        =   33
            Top             =   600
            Width           =   2295
         End
         Begin VB.ListBox lboxAccess 
            Height          =   1425
            Left            =   -72120
            MultiSelect     =   2  'Extended
            Sorted          =   -1  'True
            TabIndex        =   32
            Top             =   600
            Width           =   2295
         End
         Begin VB.CommandButton cmdGiveAccess 
            Caption         =   ">"
            Height          =   255
            Left            =   -72480
            TabIndex        =   31
            Top             =   840
            Width           =   255
         End
         Begin VB.CommandButton cmdRemoveAccess 
            Caption         =   "<"
            Height          =   255
            Left            =   -72480
            TabIndex        =   30
            Top             =   1200
            Width           =   255
         End
         Begin VB.ComboBox cboxPermissionType 
            Height          =   315
            ItemData        =   "frmMain.frx":057F
            Left            =   -69600
            List            =   "frmMain.frx":058F
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   660
            Width           =   1935
         End
         Begin VB.TextBox tbChangeHistory 
            Height          =   1635
            Left            =   -74940
            MultiLine       =   -1  'True
            TabIndex        =   28
            Top             =   420
            Width           =   7515
         End
         Begin VB.TextBox tbHelpText 
            Height          =   1275
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   27
            Top             =   420
            Width           =   7515
         End
         Begin MSComCtl2.UpDown UpDown2 
            Height          =   285
            Left            =   -73784
            TabIndex        =   40
            Top             =   1440
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "tbExecutionOrder"
            BuddyDispid     =   196617
            OrigLeft        =   1440
            OrigTop         =   1440
            OrigRight       =   1680
            OrigBottom      =   1740
            Max             =   1000
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Label Label17 
            Caption         =   "Display Tab: Label"
            Height          =   195
            Left            =   -74940
            TabIndex        =   51
            Top             =   1380
            Width           =   1395
         End
         Begin VB.Label Label16 
            Caption         =   "Database Tab: Name"
            Height          =   255
            Left            =   -71100
            TabIndex        =   50
            Top             =   1380
            Width           =   1995
         End
         Begin VB.Label Label12 
            Caption         =   "Execution Order:"
            Height          =   255
            Left            =   -74880
            TabIndex        =   43
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label11 
            Caption         =   "Enabled:"
            Height          =   255
            Left            =   -72360
            TabIndex        =   42
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label10 
            Caption         =   "Name:"
            Height          =   255
            Left            =   -74880
            TabIndex        =   41
            Top             =   480
            Width           =   3135
         End
         Begin VB.Label Label13 
            Caption         =   "Permission Type:"
            Height          =   255
            Left            =   -69600
            TabIndex        =   36
            Top             =   420
            Width           =   1695
         End
         Begin VB.Label Label14 
            Caption         =   "No Access:"
            Height          =   195
            Left            =   -74820
            TabIndex        =   35
            Top             =   360
            Width           =   1275
         End
         Begin VB.Label Label15 
            Caption         =   "Permission:"
            Height          =   255
            Left            =   -72120
            TabIndex        =   34
            Top             =   360
            Width           =   915
         End
      End
      Begin VB.Label Label29 
         Caption         =   "Modifying:  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   25
         Top             =   2220
         Width           =   975
      End
      Begin VB.Label lblObjectsModifying 
         Caption         =   "Active Links"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1080
         TabIndex        =   24
         Top             =   2220
         Width           =   1215
      End
   End
   Begin MSComctlLib.StatusBar sbMainStatusBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   22
      Top             =   7965
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12877
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            TextSave        =   "3/5/2003"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "11:12 PM"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlTBDisabled 
      Left            =   7560
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   23
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":05BF
            Key             =   "icoShowSearch"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0B13
            Key             =   "icoHideSearch"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1067
            Key             =   "icoPerformSearch"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15BB
            Key             =   "icoSaveQuery"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B0F
            Key             =   "icoOpenQuery"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2063
            Key             =   "icoSearch1"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":25B7
            Key             =   "icoSearch2"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B0B
            Key             =   "icoSearch3"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":305F
            Key             =   "icoSearch4"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":35B3
            Key             =   "icoSearch5"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3B07
            Key             =   "icoActiveLink"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":405B
            Key             =   "icoFilter"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":45AF
            Key             =   "icoSaveResults"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4B03
            Key             =   "icoConnect"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5057
            Key             =   "icoDisconnect"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":55AB
            Key             =   "icoDeleteQuery"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5AFF
            Key             =   "icoResetQuery"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6053
            Key             =   "icoPrintResults"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":65A7
            Key             =   "icoPreviousQuery"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6AFB
            Key             =   "icoNextQuery"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":704F
            Key             =   "icoModify"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":75A3
            Key             =   "icoExecute"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7AF7
            Key             =   "icoFields"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlTBHighlighted 
      Left            =   6960
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   23
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":804B
            Key             =   "icoShowSearch"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":859F
            Key             =   "icoHideSearch"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8AF3
            Key             =   "icoPerformSearch"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9047
            Key             =   "icoSaveQuery"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":959B
            Key             =   "icoOpenQuery"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9AEF
            Key             =   "icoSearch1"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A043
            Key             =   "icoSearch2"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A597
            Key             =   "icoSearch3"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AAEB
            Key             =   "icoSearch4"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B03F
            Key             =   "icoSearch5"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B593
            Key             =   "icoActiveLink"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BAE7
            Key             =   "icoFilter"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C03B
            Key             =   "icoSaveResults"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C58F
            Key             =   "icoConnect"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CAE3
            Key             =   "icoDisconnect"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D037
            Key             =   "icoDeleteQuery"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D58B
            Key             =   "icoResetQuery"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DADF
            Key             =   "icoPrintResults"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E033
            Key             =   "icoPreviousQuery"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E587
            Key             =   "icoNextQuery"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EADB
            Key             =   "icoModify"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F02F
            Key             =   "icoExecute"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F583
            Key             =   "icoFields"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbMainToolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   15
      Top             =   0
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlTB"
      DisabledImageList=   "imlTBDisabled"
      HotImageList    =   "imlTBHighlighted"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Connection"
            Description     =   "icoConnect"
            Object.ToolTipText     =   "Connect / Disconnect from server."
            ImageKey        =   "icoConnect"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   120
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ResetQuery"
            Description     =   "icoResetQuery"
            Object.ToolTipText     =   "New query."
            ImageKey        =   "icoResetQuery"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SaveQuery"
            Description     =   "icoSaveQuery"
            Object.ToolTipText     =   "Save current query."
            ImageKey        =   "icoSaveQuery"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SaveResults"
            Description     =   "icoSaveResults"
            Object.ToolTipText     =   "Save the current search results."
            ImageKey        =   "icoSaveResults"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OpenQuery"
            Description     =   "icoOpenQuery"
            Object.ToolTipText     =   "Open a saved query."
            ImageKey        =   "icoOpenQuery"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DeleteQuery"
            Description     =   "icoDeleteQuery"
            Object.ToolTipText     =   "Delete a saved query."
            ImageKey        =   "icoDeleteQuery"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PrintResults"
            Description     =   "icoPrintResults"
            Object.ToolTipText     =   "Print the search results."
            ImageKey        =   "icoPrintResults"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ShowDialog"
            Description     =   "icoShowSearch"
            Object.ToolTipText     =   "Show/Hide the dialog."
            ImageKey        =   "icoShowSearch"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ObjectType"
            Description     =   "icoActiveLink"
            Object.ToolTipText     =   "Active Link"
            Object.Tag             =   "ActiveLink"
            ImageKey        =   "icoActiveLink"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ActiveLink"
                  Object.Tag             =   "icoActiveLink"
                  Text            =   "Active Link"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Filter"
                  Object.Tag             =   "icoFilter"
                  Text            =   "Filter"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Field"
                  Object.Tag             =   "icoFields"
                  Text            =   "Field"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ActionType"
            Description     =   "icoPerformSearch"
            Object.ToolTipText     =   "Search"
            Object.Tag             =   "Search"
            ImageKey        =   "icoPerformSearch"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Search"
                  Object.Tag             =   "Search"
                  Text            =   "Search Objects"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Modify"
                  Object.Tag             =   "Modify"
                  Text            =   "Modify Objects"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   1200
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Execute"
            Description     =   "icoExecute"
            Object.ToolTipText     =   "Execute."
            ImageKey        =   "icoExecute"
            Object.Width           =   1400
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SavedQueries"
            Description     =   "Saved Queries"
            Object.ToolTipText     =   "Open a previously saved query."
            Style           =   4
            Object.Width           =   2495
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "PreviousQuery"
            Description     =   "icoPreviousQuery"
            Object.ToolTipText     =   "Show previous successfull search."
            ImageKey        =   "icoPreviousQuery"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "NextQuery"
            Description     =   "icoNextQuery"
            Object.ToolTipText     =   "Show next successfull search."
            ImageKey        =   "icoNextQuery"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.ComboBox cboxSavedQueries 
         Height          =   315
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   0
         Width           =   2475
      End
   End
   Begin MSComctlLib.ImageList imlTB 
      Left            =   6360
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   23
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FAD7
            Key             =   "icoShowSearch"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1002B
            Key             =   "icoHideSearch"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1057F
            Key             =   "icoPerformSearch"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10AD3
            Key             =   "icoSaveQuery"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11027
            Key             =   "icoOpenQuery"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1157B
            Key             =   "icoSearch1"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11ACF
            Key             =   "icoSearch2"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12023
            Key             =   "icoSearch3"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12577
            Key             =   "icoSearch4"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12ACB
            Key             =   "icoSearch5"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1301F
            Key             =   "icoActiveLink"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13573
            Key             =   "icoFilter"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13AC7
            Key             =   "icoSaveResults"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1401F
            Key             =   "icoConnect"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14573
            Key             =   "icoDisconnect"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14AC7
            Key             =   "icoDeleteQuery"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1501B
            Key             =   "icoResetQuery"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1556F
            Key             =   "icoPrintResults"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15AC3
            Key             =   "icoPreviousQuery"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16017
            Key             =   "icoNextQuery"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1656B
            Key             =   "icoModify"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16ABF
            Key             =   "icoExecute"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17013
            Key             =   "icoFields"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSearchOptions 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   0
      ScaleHeight     =   2415
      ScaleWidth      =   7815
      TabIndex        =   14
      Top             =   2700
      Width           =   7815
      Begin VB.CommandButton cmdMore 
         Appearance      =   0  'Flat
         Caption         =   "More -->"
         CausesValidation=   0   'False
         Height          =   255
         Left            =   6600
         MaskColor       =   &H8000000F&
         TabIndex        =   9
         Top             =   2145
         Width           =   915
      End
      Begin VB.ComboBox cboxProperties 
         Height          =   315
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1740
         Width           =   1935
      End
      Begin VB.ComboBox cboxConditions 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1740
         Width           =   1815
      End
      Begin VB.CommandButton cmdAnd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   6000
         TabIndex        =   3
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton cmdOr 
         Caption         =   "OR"
         Height          =   375
         Left            =   6960
         TabIndex        =   4
         Top             =   1680
         Width           =   855
      End
      Begin VB.ComboBox cboxValue 
         Height          =   315
         Left            =   3960
         TabIndex        =   2
         Top             =   1740
         Width           =   1935
      End
      Begin MSComctlLib.ListView lvSearchArguments 
         Height          =   1455
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   2566
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "Property"
            Text            =   "Property"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "Condition"
            Text            =   "Condition"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "Value"
            Text            =   "Value"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "ANDOR"
            Text            =   "And/Or"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CheckBox ckboxCaseSensitive 
         Alignment       =   1  'Right Justify
         Caption         =   "Case Sensitive:"
         CausesValidation=   0   'False
         Height          =   195
         Left            =   4380
         TabIndex        =   5
         Top             =   2160
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
      Begin VB.Label lblObjectsSearching 
         Caption         =   "Active Links"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1140
         TabIndex        =   20
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Searching:  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   19
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Properties:"
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   1500
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Condition:"
         Height          =   255
         Left            =   2040
         TabIndex        =   17
         Top             =   1500
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Search Value:"
         Height          =   255
         Left            =   3960
         TabIndex        =   16
         Top             =   1500
         Width           =   1455
      End
   End
   Begin MSComctlLib.ImageList imlTreeViewIcons 
      Left            =   5760
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17567
            Key             =   "Servers"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17AB9
            Key             =   "Groups"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1800B
            Key             =   "Applications"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1855D
            Key             =   "Menus"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18AAF
            Key             =   "Filters"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19001
            Key             =   "Escalations"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19553
            Key             =   "Guides"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19AA5
            Key             =   "Server"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19FF7
            Key             =   "Forms"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A549
            Key             =   "ActiveLinks"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplitter 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   1200
      Left            =   5520
      ScaleHeight     =   522.532
      ScaleMode       =   0  'User
      ScaleWidth      =   1404
      TabIndex        =   13
      Top             =   1320
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox picTitles 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   9240
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   360
      Width           =   9240
      Begin VB.Label lblListViewTitle 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Details:"
         Height          =   270
         Index           =   1
         Left            =   2078
         TabIndex        =   12
         Tag             =   "1054"
         Top             =   12
         Width           =   3216
      End
      Begin VB.Label lblTreeViewTitle 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Server Source Objects:"
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   11
         Tag             =   "1053"
         Top             =   12
         Width           =   2016
      End
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   5760
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvListView 
      Height          =   1320
      Left            =   2040
      TabIndex        =   8
      Top             =   1320
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   2328
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "keyALName"
         Text            =   "Name"
         Object.Width           =   7408
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "keyFormName"
         Text            =   "Form Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "keyModTime"
         Text            =   "Modification Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "keyExecutionMask"
         Text            =   "Execution Mask"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "keyExecutionOrder"
         Text            =   "Execution Order"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "keyEnabled"
         Text            =   "Enabled"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "keyType"
         Text            =   "Data Type"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.TreeView tvTreeView 
      Height          =   1320
      Left            =   0
      TabIndex        =   7
      Top             =   1320
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   2328
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      HotTracking     =   -1  'True
      ImageList       =   "imlTreeViewIcons"
      Appearance      =   1
   End
   Begin VB.Image imgVSplitter 
      Appearance      =   0  'Flat
      Height          =   1215
      Left            =   5280
      MouseIcon       =   "frmMain.frx":1AA9B
      MousePointer    =   9  'Size W E
      Top             =   1320
      Width           =   165
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileConnection 
         Caption         =   "Connect To Server"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuFileSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New Query"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open Query"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSaveQuery 
         Caption         =   "&Save Query"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveResults 
         Caption         =   "Sa&ve Results"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuFileFind 
         Caption         =   "1002"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSendTo 
         Caption         =   "1003"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileDelete 
         Caption         =   "&Delete Query"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuFileRename 
         Caption         =   "1006"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "1007"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilePrintResults 
         Caption         =   "&Print Results"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileSeperator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "1010"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditSpacer 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "&Remove Query Item"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "1012"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "1013"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditPasteSpecial 
         Caption         =   "1014"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "1015"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditInvertSelection 
         Caption         =   "1016"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuEditSearchDialog 
         Caption         =   "Show Search Dialog"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "1018"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "1019"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "1020"
         Enabled         =   0   'False
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "1021"
         Enabled         =   0   'False
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "1022"
         Enabled         =   0   'False
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "1023"
         Enabled         =   0   'False
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewArrangeIcons 
         Caption         =   "1024"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "&Refresh"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewWebBrowser 
         Caption         =   "1027"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuSearchObjects 
         Caption         =   "Search Active Links"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuSearchFilters 
         Caption         =   "Search Filters"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuSearchFields 
         Caption         =   "Search Fields"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuSearchPredefined 
         Caption         =   "Predefined Queries"
         Shortcut        =   {F12}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSearchSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearchAssigned1 
         Caption         =   "1"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuSearchAssigned2 
         Caption         =   "2"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuSearchAssigned3 
         Caption         =   "3"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuSearchAssigned4 
         Caption         =   "4"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnuSearchAssigned5 
         Caption         =   "5"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnuSearchSeperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearchAssign 
         Caption         =   "Assign Queries"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuSearchSeperator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearchPerformSearch 
         Caption         =   "Execute Search"
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu mnuModify 
      Caption         =   "&Modify"
      Begin VB.Menu mnuModifyActiveLinks 
         Caption         =   "Active Links"
         Shortcut        =   +{F3}
      End
      Begin VB.Menu mnuModifyFilters 
         Caption         =   "Filters"
         Shortcut        =   +{F4}
      End
      Begin VB.Menu mnuModifyFields 
         Caption         =   "Fields"
         Shortcut        =   +{F5}
      End
      Begin VB.Menu mnuModifySeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuModifyExecute 
         Caption         =   "Execute Modification"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsCache 
         Caption         =   "Cache"
         Begin VB.Menu mnuCacheUpdate 
            Caption         =   "Update Cache"
            Shortcut        =   ^U
         End
         Begin VB.Menu mnuCacheRebuild 
            Caption         =   "Rebuild Cache"
            Shortcut        =   ^B
         End
      End
      Begin VB.Menu mnuToolsSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "&Options"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "&Search Help"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuSearchDialog 
      Caption         =   "SearchDialog"
      Visible         =   0   'False
      Begin VB.Menu mnuSearchDialogDelete 
         Caption         =   "Delete Item"
      End
      Begin VB.Menu mnuSearchDialogSpacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearchDialogHide 
         Caption         =   "Hide Search Dialog"
      End
   End
   Begin VB.Menu mnuListView 
      Caption         =   "ListView"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuListViewSortBy 
         Caption         =   "Sort by"
         Begin VB.Menu mnuListViewSortName 
            Caption         =   "Object Name"
         End
         Begin VB.Menu mnuListViewSortForm 
            Caption         =   "Form Name"
         End
         Begin VB.Menu mnuListViewSortMod 
            Caption         =   "Modification Time"
         End
         Begin VB.Menu mnuListViewSortMask 
            Caption         =   "Execution Mask"
         End
         Begin VB.Menu mnuListViewSortExecute 
            Caption         =   "Execute On"
         End
         Begin VB.Menu mnuListViewSortEnabled 
            Caption         =   "Enabled"
         End
         Begin VB.Menu mnuListViewSortID 
            Caption         =   "ID"
         End
         Begin VB.Menu mnuListViewSortDataType 
            Caption         =   "Data Type"
         End
      End
   End
   Begin VB.Menu mnuTreeView 
      Caption         =   "TreeView"
      Visible         =   0   'False
      Begin VB.Menu mnuTreeViewCheckAll 
         Caption         =   "Check All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuTreeViewCheckNone 
         Caption         =   "Check None"
      End
      Begin VB.Menu mnuTreeViewInverseSelection 
         Caption         =   "Inverse Checked Forms"
         Shortcut        =   ^I
      End
   End
   Begin VB.Menu mnuToolBar 
      Caption         =   "ToolBar"
      Visible         =   0   'False
      Begin VB.Menu mnuToolBarRefresh 
         Caption         =   "Refresh Icons"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'///////////////////
'Constants
'///////////////////

Const BasicTab = 1
Const PermissionsTab = 2
Const MiscTab = 3

Const SearchView = 1
Const FieldModifyView = 2
Const ALModifyView = 3
Const FLModifyView = 4

'Print column Max sizes
Const SpacerSize = 2
Const FieldNameSize = 30
Const FormNameSize = 30
Const ExecutionOrderSize = 4
Const ModificationTimeSize = 24
Const EnabledSize = 3
Const ExecuteMaskSize = 19


Const arOK = 0

Const DATEGREATER = 2
Const DATELESS = 4

Const ALNAMESTARTS = 2
Const ALNAMECONTAINS = 4
Const ALNAMEENDS = 3

Const DEFAULTTIME = "12:00:00 AM"

Const NAME_COLUMN = 0
Const TYPE_COLUMN = 1
Const SIZE_COLUMN = 2
Const DATE_COLUMN = 3

Const sglSplitLimit = 500
Const sglHSplitLimit = 3500


Const FORMCAPTION = "AR Explorer - "

Const ROOT = "Servers"
Const SERVER = "Server"    'This will be moved from being a constant to var
Const ARFORMS = "Forms"
Const ACTIVELINKS = "Active Links"
Const FILTERS = "Filters"
Const ESCALATIONS = "Escalations"
Const GUIDES = "Guides"
Const APPLICATIONS = "Applications"
Const MENUS = "Menus"
Const GROUPS = "Groups"

Const KEY_PREFIX = "node"

Const ICON_ROOT = 1
Const ICON_SERVER = 8
Const ICON_FORMS = 9
Const ICON_ACTIVELINKS = 10
Const ICON_FILTERS = 5
Const ICON_ESCALATIONS = 6
Const ICON_GUIDES = 7
Const ICON_APPLICATIONS = 3
Const ICON_MENUS = 4
Const ICON_GROUPS = 2

Const MINVIEWWIDTH = 1500
Const MAXVIEWWIDTH = 1500

Const MINVIEWHEIGTH = 40
Const MAXVIEWHEIGTH = 2415

Const SPLITTERWIDTH = 50

'///////////////////
'Private Properties
'///////////////////

Private lCheckedNodeCount As Long

Private mbMoving As Boolean

Private bSearchViewOpen As Boolean

Private iSelectedTab As Integer

Private ServerName As String
Private UserName As String
Private Password As String

Private sALName As String
Private iALNameOperator As Integer

Private sModTime As String
Private iModTimeOperator As Integer

Private sALRunIfText As String

Public sDate As String
Public sTime As String

Public bEditOk As Boolean

Private sCurrentSearchType As String

Private bRebuildServerCache As Boolean

Private lCurrentTop As Long
Private lCurrentLeft As Long
Private lCurrentWidth As Long
Private lCurrentHeight As Long
Private lOldTop As Long
Private lOldLeft As Long
Private lOldWidth As Long
Private lOldHeight As Long

'Private RecentQueries() As colQueryList
Private RecentQueries As New colQueryGroup

Private NumQueriesInPrevList As Integer

'The current special dialog visible (SearchView, FieldModifyView, ALModifyView, FLModifyView)
Private CurrentView As Integer
'To make things simple as hell, we'll assign this to the current dialog, this will
'reduce the need for 50 billion select case statements, hell, we might not even
'need the CurrentView above
Private picCurrentView As PictureBox

Private bPreviousEnabled As Boolean
Private bNextEnabled As Boolean

'///////////////////
'Public Properties
'///////////////////
'Our AWESOME interface
Public ARCom As New clsARIDLL
Public ARQuery As New colQueryList

Public lSearchValue As Long

Public AssignedQueries As New colQueryGroup

Public iRecentQueriesCount As Integer
Public iCurrentRecentQuery As Integer
Public iNumberInRecentQueryList As Integer

Public sDefaultSearchType As String

Public sCurrentServerName As String
Public lCurrentServerID As Long

'///////////////////
'Function Declorations
'///////////////////

Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
  


Public Sub SetFormConnected()

  Me.Caption = "AR Explorer - " & sCurrentServerName

  'disable menus
  mnuFileConnection.Caption = "Disconnect From Server"
  mnuToolsCache.Enabled = True
  mnuSearchPerformSearch.Enabled = True
  mnuModifyExecute.Enabled = True
  
  'disable toolbar buttons (ToolTipText, enabled & icon)
  tbMainToolbar.Buttons("Connection").ToolTipText = "Disconnect From Server"
  tbMainToolbar.Buttons("Connection").Image = icoDisconnect
  tbMainToolbar.Buttons("Execute").Enabled = True

End Sub

Public Sub SetFormDisconnected()

  Me.Caption = "AR Explorer - " & sCurrentServerName
  
  'disable menus
  mnuFileConnection.Caption = "Connect To Server"
  mnuToolsCache.Enabled = False
  mnuSearchPerformSearch.Enabled = False
  mnuModifyExecute.Enabled = False
  
  'disable toolbar buttons (ToolTipText, enabled & icon)
  tbMainToolbar.Buttons("Connection").ToolTipText = "Connect To Server"
  tbMainToolbar.Buttons("Connection").Image = icoConnect
  tbMainToolbar.Buttons("Execute").Enabled = False

End Sub



Private Sub cboxPermissionType_Click()
Dim i As Long
Dim iRemovedCount As Long

  If cboxPermissionType.Text = "Remove All" Then
  
    iRemovedCount = 0

    For i = 0 To (lboxAccess.ListCount - 1)
    
      lboxNoAccess.AddItem (lboxAccess.List(i - iRemovedCount))
      lboxAccess.RemoveItem (i - iRemovedCount)
      iRemovedCount = iRemovedCount + 1
      
    Next i

    lboxNoAccess.Refresh
    lboxAccess.Refresh
    
    lboxNoAccess.Enabled = False
    lboxAccess.Enabled = False
  Else
    lboxNoAccess.Enabled = True
    lboxAccess.Enabled = True
  End If

End Sub

Private Sub cboxSavedQueries_Click()
Dim i As Integer
Dim qryTemp As colQueryList

  i = cboxSavedQueries.ListIndex + 1

  If i > 0 Then
    If cboxSavedQueries.Text = sEmptyString Then
      frmAssignQuery.LoadQueries i
      frmAssignQuery.Show vbModal
    End If
    
    If Not (AssignedQueries.Item(i).SaveName = sEmptyString) Then
      'cboxSavedQueries.Text = AssignedQueries.Item(i).SaveName
      SetCurrentQuery AssignedQueries.Item(i)
    Else
      Set qryTemp = New colQueryList
      qryTemp.ResetCollection
      SetCurrentQuery qryTemp
      ShowCurrentQuery
    End If
    
  End If

End Sub

Private Sub ckboxCaseSensitive_Click()
  
  ARQuery.CaseSensitive = ckboxCaseSensitive.Value
  ARQuery.Dirty = True

End Sub


Private Sub ckboxCaseSensitive_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

  If Button = 2 Then
    PopupMenu mnuSearchDialog
  End If

End Sub

Private Sub cmdModifyMore_Click()
Dim lExpandAmount As Long

'  lExpandAmount = 7815 - picFieldModification.Width
'
'  If Not Me.WindowState = vbMaximized Then
'    Me.Width = Me.Width + lExpandAmount
'  End If

End Sub

Private Sub cmdGiveAccess_Click()
Dim i As Long
Dim iRemovedCount As Long
  
  iRemovedCount = 0
  
  For i = 0 To (lboxNoAccess.ListCount - 1)
    If lboxNoAccess.Selected(i - iRemovedCount) Then
      lboxAccess.AddItem (lboxNoAccess.List(i - iRemovedCount))
      lboxNoAccess.RemoveItem (i - iRemovedCount)
      iRemovedCount = iRemovedCount + 1
    End If
  Next i

  lboxNoAccess.Refresh
  lboxAccess.Refresh
  
End Sub

Private Sub cmdMore_Click()
Dim lExpandAmount As Long

  lExpandAmount = 7815 - picSearchOptions.Width
  
  If Not Me.WindowState = vbMaximized Then
    Me.Width = Me.Width + lExpandAmount
  End If
  
End Sub

Private Sub cmdShowChangeHistoryText_Click()

  frmGetText.Top = Me.Top + picCurrentView.Top + tbChangeHistory.Top + (tbChangeHistory.Height * 3)
  frmGetText.Left = Me.Left + picCurrentView.Left + tbChangeHistory.Left + 40
  frmGetText.Caption = "Change History"
  frmGetText.iCalledBy = ChangeHistory
  frmGetText.tbText.Text = tbChangeHistory.Text
  frmGetText.Show vbModal
  'tbChangeHistory.Text = frmGetText.tbText.Text

End Sub

Private Sub cmdShowHelpText_Click()

  frmGetText.Top = Me.Top + picCurrentView.Top + tbHelpText.Top + (tbHelpText.Height * 3)
  frmGetText.Left = Me.Left + picCurrentView.Left + tbHelpText.Left + 40
  frmGetText.Caption = "Help Text"
  frmGetText.tbText.Text = tbHelpText.Text
  frmGetText.iCalledBy = HelpText
  frmGetText.Show vbModal
  'tbHelpText.Text = frmGetText.tbText.Text

End Sub

Private Sub cmdRemoveAccess_Click()
Dim i As Long
Dim iRemovedCount As Long
  
  iRemovedCount = 0
  
  For i = 0 To (lboxAccess.ListCount - 1)
    If lboxAccess.Selected(i - iRemovedCount) Then
      lboxNoAccess.AddItem (lboxAccess.List(i - iRemovedCount))
      lboxAccess.RemoveItem (i - iRemovedCount)
      iRemovedCount = iRemovedCount + 1
    End If
  Next i

  lboxNoAccess.Refresh
  lboxAccess.Refresh

End Sub


Private Sub Form_GotFocus()
  
  lCurrentTop = Me.Top
  lCurrentLeft = Me.Left
  lCurrentWidth = Me.Width
  lCurrentHeight = Me.Height
  
  Me.Refresh
  tbMainToolbar.Refresh
  picSearchOptions.Refresh

End Sub


Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    PopupMenu mnuSearchDialog
  End If

End Sub

Private Sub lblMore_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

  If Button = 2 Then
    PopupMenu mnuSearchDialog
  End If

End Sub

Private Sub lblObjectsSearching_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

  If Button = 2 Then
    PopupMenu mnuSearchDialog
  End If

End Sub

Private Sub lvListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

  If (lvListView.SortKey = ColumnHeader.Index - 1) Then
    Select Case lvListView.SortOrder
      Case lvwDescending
        lvListView.SortOrder = lvwAscending
      Case lvwAscending
        lvListView.SortOrder = lvwDescending
    End Select
  Else
    lvListView.SortKey = ColumnHeader.Index - 1
  End If
  
  lvListView.Refresh

End Sub


'Executes when the items are selected in the results list
Private Sub lvListView_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim lSelectedCount As Long
Dim bFirstFound As Boolean
Dim sFirstName As String

  lSelectedCount = 0
  
  If Button = 2 Then
    PopupMenu mnuListView
  End If
  
  bFirstFound = False
  For i = 1 To Me.lvListView.ListItems.Count
    If lvListView.ListItems(i).Selected = True Then
      If bFirstFound = False Then
        bFirstFound = True
        sFirstName = lvListView.ListItems(i).Text
      End If
      lSelectedCount = lSelectedCount + 1
    End If
  Next i
  
  On Error Resume Next
  If lSelectedCount = 1 Then
    tbDatabaseName.Text = sFirstName
    tbName.Text = sFirstName 'Get Active Link name of selected item
    lvListView.tag = 1
    lvListViewActions.ListItems.Clear 'Clear the Action List
    
    'Add items to the Actions List View
    Call FillActionList(tbDatabaseName.Text, TYPE_AL)

  ElseIf lvListView.tag = 1 Then
    lvListView.tag = 0
    tbDatabaseName.Text = ""
    tbName.Text = ""
  End If
  On Error GoTo 0
  
  If lSelectedCount > 0 Then
    SetStatusMessage (Trim(Str(lSelectedCount)) & " Object(s) selected.")
  End If

End Sub

Private Sub lvSearchArguments_Click()
Dim qiQueryItem As clsQueryItem

  If lvSearchArguments.ListItems.Count > 0 Then
    If lvSearchArguments.SelectedItem.Index <= ARQuery.Count Then
      ARQuery.Dirty = True
      Set qiQueryItem = ARQuery.Item(lvSearchArguments.SelectedItem.Text & lvSearchArguments.SelectedItem.SubItems(2))
      cboxProperties.Text = qiQueryItem.SearchType
      If cboxConditions.Enabled = True Then
        cboxConditions.Text = qiQueryItem.SearchConditionString
      Else
        cboxConditions.Enabled = True
        cboxConditions.Locked = True
        cboxConditions.Text = qiQueryItem.SearchConditionString
        cboxConditions.Enabled = False
        cboxConditions.Locked = False
      End If
      cboxConditions.tag = qiQueryItem.SearchParam
      cboxValue.Text = qiQueryItem.SearchValueString
      cboxValue.tag = qiQueryItem.SearchValueNum
    End If
    
  End If
  
End Sub


Private Sub DeleteSearchItem()

  If lvSearchArguments.ListItems.Count > 0 Then
    ARQuery.Remove (lvSearchArguments.SelectedItem.Text & _
      lvSearchArguments.SelectedItem.SubItems(2))
    Me.lvSearchArguments.ListItems.Remove (Me.lvSearchArguments.SelectedItem.Index)
  End If
  
  ShowQueryInStatBar

End Sub

Private Sub lvSearchArguments_DblClick()

  DeleteSearchItem

End Sub


Private Sub lvSearchArguments_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

  If Button = 2 Then
    PopupMenu mnuSearchDialog
  End If

End Sub

Private Sub mnuCacheRebuild_Click()
Dim i As Integer
Dim sMSG As String


  sMSG = "Rebuilding the cache may take several minutes.  "
  sMSG = sMSG & "Are you sure you wish to continue?"
  
  Beep
  i = MsgBox(sMSG, vbYesNo + vbQuestion, "Caution..")
  
  If i = vbYes Then
    frmProgress.lblStatus.Caption = "Rebuilding cache for Form: "
    frmProgress.pbProgress2.Visible = True
    frmProgress.Caption = "Rebuilding Cache"
    frmProgress.Show
    frmProgress.Refresh
    
    DeleteCache (modDatabase.GetServerCacheID(sCurrentServerName))
    CreateCache (sCurrentServerName)
    
    frmProgress.Hide
  End If

End Sub

'possible bug, if admin adds a new form while running ARE then new form won't be logged
Private Sub mnuCacheUpdate_Click()
Dim i As Long
Dim nNode As Node
  
  frmProgress.Caption = "Updating Cache"
  frmProgress.pbProgress2.Visible = True
  frmProgress.lblStatus.Caption = "Checking Form: "
  frmProgress.Show
  frmProgress.Refresh
  frmProgress.pbProgress.Max = tvTreeView.Nodes.Count
  
  For i = 2 To tvTreeView.Nodes.Count
    Set nNode = tvTreeView.Nodes(i)
    If nNode.Parent = ARFORMS Then
      frmProgress.lblStatus.Caption = "Checking Form: " & tvTreeView.Nodes(i).Text
      frmProgress.pbProgress.Value = i
      frmProgress.Refresh
      nNode.tag = UpdateFormCacheByModTime(nNode.Text, nNode.tag)
    End If
  Next i
  
  frmProgress.Hide

End Sub

Private Sub mnuEditSearchDialog_Click()
  SwitchSearchView
End Sub

Private Sub mnuFileConnection_Click()
  ReverseConnection
End Sub

Private Sub mnuFilePrintResults_Click()
  PrintResults
End Sub

Private Sub mnuFileSaveQuery_Click()

  frmQueryIO.SetupForm (QUERY_SAVE)
  frmQueryIO.Show vbModal

End Sub


Private Sub mnuListViewSortDataType_Click()

  If (lvListView.SortKey = 3) Then
    Select Case lvListView.SortOrder
      Case lvwDescending
        lvListView.SortOrder = lvwAscending
      Case lvwAscending
        lvListView.SortOrder = lvwDescending
    End Select
  Else
    lvListView.SortKey = 3
  End If
  
  lvListView.Refresh

End Sub

Private Sub mnuListViewSortEnabled_Click()

  If (lvListView.SortKey = 5) Then
    Select Case lvListView.SortOrder
      Case lvwDescending
        lvListView.SortOrder = lvwAscending
      Case lvwAscending
        lvListView.SortOrder = lvwDescending
    End Select
  Else
    lvListView.SortKey = 5
  End If
  
  lvListView.Refresh

End Sub

Private Sub mnuListViewSortExecute_Click()

  If (lvListView.SortKey = 4) Then
    Select Case lvListView.SortOrder
      Case lvwDescending
        lvListView.SortOrder = lvwAscending
      Case lvwAscending
        lvListView.SortOrder = lvwDescending
    End Select
  Else
    lvListView.SortKey = 4
  End If
  
  lvListView.Refresh

End Sub

Private Sub mnuListViewSortForm_Click()

  If (lvListView.SortKey = 1) Then
    Select Case lvListView.SortOrder
      Case lvwDescending
        lvListView.SortOrder = lvwAscending
      Case lvwAscending
        lvListView.SortOrder = lvwDescending
    End Select
  Else
    lvListView.SortKey = 1
  End If
  
  lvListView.Refresh

End Sub

Private Sub mnuListViewSortID_Click()

  If (lvListView.SortKey = 2) Then
    Select Case lvListView.SortOrder
      Case lvwDescending
        lvListView.SortOrder = lvwAscending
      Case lvwAscending
        lvListView.SortOrder = lvwDescending
    End Select
  Else
    lvListView.SortKey = 2
  End If
  
  lvListView.Refresh

End Sub

Private Sub mnuListViewSortMask_Click()

  If (lvListView.SortKey = 3) Then
    Select Case lvListView.SortOrder
      Case lvwDescending
        lvListView.SortOrder = lvwAscending
      Case lvwAscending
        lvListView.SortOrder = lvwDescending
    End Select
  Else
    lvListView.SortKey = 3
  End If
  
  lvListView.Refresh

End Sub

Private Sub mnuListViewSortMod_Click()

  If (lvListView.SortKey = 2) Then
    Select Case lvListView.SortOrder
      Case lvwDescending
        lvListView.SortOrder = lvwAscending
      Case lvwAscending
        lvListView.SortOrder = lvwDescending
    End Select
  Else
    lvListView.SortKey = 2
  End If
  
  lvListView.Refresh

End Sub

Private Sub mnuListViewSortName_Click()

  If (lvListView.SortKey = 0) Then
    Select Case lvListView.SortOrder
      Case lvwDescending
        lvListView.SortOrder = lvwAscending
      Case lvwAscending
        lvListView.SortOrder = lvwDescending
    End Select
  Else
    lvListView.SortKey = 0
  End If
  
  lvListView.Refresh

End Sub

Private Sub mnuModifyActiveLinks_Click()

  tbMainToolbar.Buttons("ObjectType").tag = TYPE_AL
  mnuModifyExecute.Enabled = True
  tbMainToolbar.Buttons("ActionType").ButtonMenus("Modify").Enabled = True
  SetupSearchParams TYPE_AL, True
  SetupDialog ("Modify")

End Sub

Private Sub mnuModifyExecute_Click()
  Modify
End Sub

Private Sub mnuModifyFields_Click()

  tbMainToolbar.Buttons("ObjectType").tag = TYPE_FIELD
  mnuModifyExecute.Enabled = True
  tbMainToolbar.Buttons("ActionType").ButtonMenus("Modify").Enabled = True
  SetupSearchParams TYPE_FIELD, True
  SetupDialog ("Modify")

End Sub

Private Sub mnuModifyFilters_Click()

  tbMainToolbar.Buttons("ObjectType").tag = TYPE_FILTER
  mnuModifyExecute.Enabled = True
  tbMainToolbar.Buttons("ActionType").ButtonMenus("Modify").Enabled = True
  SetupSearchParams TYPE_FILTER, True
  SetupDialog ("Modify")

End Sub

Private Sub mnuSaveResults_Click()
  ExportResults
End Sub

Private Sub mnuSearchAssign_Click()
  frmAssignQuery.LoadQueries 0
  frmAssignQuery.Show vbModal
End Sub

Private Sub mnuSearchAssigned1_Click()
  If AssignedQueries.Item(1).SaveName = sEmptyString Then
    frmAssignQuery.LoadQueries 1
    frmAssignQuery.Show vbModal
  Else
  End If
  If Not (AssignedQueries.Item(1).SaveName = sEmptyString) Then
    SetCurrentQuery AssignedQueries.Item(1)
  End If
End Sub

Private Sub mnuSearchAssigned2_Click()
  If AssignedQueries.Item(2).SaveName = sEmptyString Then
    frmAssignQuery.LoadQueries 2
    frmAssignQuery.Show vbModal
  Else
  End If
  If Not (AssignedQueries.Item(2).SaveName = sEmptyString) Then
    SetCurrentQuery AssignedQueries.Item(2)
  End If
End Sub

Private Sub mnuSearchAssigned3_Click()
  If AssignedQueries.Item(3).SaveName = sEmptyString Then
    frmAssignQuery.LoadQueries 3
    frmAssignQuery.Show vbModal
  Else
  End If
  If Not (AssignedQueries.Item(3).SaveName = sEmptyString) Then
    SetCurrentQuery AssignedQueries.Item(3)
  End If
End Sub

Private Sub mnuSearchAssigned4_Click()
  If AssignedQueries.Item(4).SaveName = sEmptyString Then
    frmAssignQuery.LoadQueries 4
    frmAssignQuery.Show vbModal
  Else
  End If
  If Not (AssignedQueries.Item(4).SaveName = sEmptyString) Then
    SetCurrentQuery AssignedQueries.Item(4)
  End If
End Sub

Private Sub mnuSearchAssigned5_Click()
  If AssignedQueries.Item(5).SaveName = sEmptyString Then
    frmAssignQuery.LoadQueries 5
    frmAssignQuery.Show vbModal
  Else
  End If
  If Not (AssignedQueries.Item(5).SaveName = sEmptyString) Then
    SetCurrentQuery AssignedQueries.Item(5)
  End If
End Sub


Private Sub mnuSearchDialogDelete_Click()

  DeleteSearchItem

End Sub

Private Sub mnuSearchDialogHide_Click()
  SwitchSearchView
End Sub

Private Sub mnuSearchFields_Click()

  mnuModifyExecute.Enabled = True
  tbMainToolbar.Buttons("ObjectType").tag = TYPE_FIELD
'  tbMainToolbar.Buttons("ActionType").ButtonMenus("Modify").Enabled = True
  SetupSearchParams TYPE_FIELD, True
  SetupDialog ("Search")

End Sub

Private Sub mnuSearchFilters_Click()

  mnuModifyExecute.Enabled = False
  tbMainToolbar.Buttons("ObjectType").tag = TYPE_FILTER
'  tbMainToolbar.Buttons("ActionType").ButtonMenus("Modify").Enabled = False
  SetupSearchParams TYPE_FILTER, True
  SetupDialog ("Search")
  
End Sub


Private Sub mnuSearchPerformSearch_Click()

  AddToRecentQueryList (True)
  Search

End Sub


Private Sub mnuToolBarRefresh_Click()
  tbMainToolbar.Refresh
End Sub

Private Sub mnuTreeViewCheckAll_Click()
Dim i As Long
Dim nTempNode As Node

  If tvTreeView.Nodes.Count > 0 Then
    tvTreeView.Nodes(KEY_PREFIX & ARFORMS).Checked = True
    Set nTempNode = tvTreeView.Nodes(KEY_PREFIX & ARFORMS).Child
    For i = 1 To tvTreeView.Nodes(KEY_PREFIX & ARFORMS).Children
      nTempNode.Checked = True
      Set nTempNode = nTempNode.Next
    Next i
  End If
  
End Sub

Private Sub mnuTreeViewCheckNone_Click()
Dim i As Long
Dim nTempNode As Node

  If tvTreeView.Nodes.Count > 0 Then
    tvTreeView.Nodes(KEY_PREFIX & ARFORMS).Checked = False
    Set nTempNode = tvTreeView.Nodes(KEY_PREFIX & ARFORMS).Child
    For i = 1 To tvTreeView.Nodes(KEY_PREFIX & ARFORMS).Children
      nTempNode.Checked = False
      Set nTempNode = nTempNode.Next
    Next i
  End If
End Sub

Private Sub mnuTreeViewInverseSelection_Click()
Dim i As Long

  For i = 2 To tvTreeView.Nodes.Count
    'If tvTreeView.Nodes(i).Parent.Text = ARFORMS Then
      tvTreeView.Nodes(i).Checked = Not tvTreeView.Nodes(i).Checked
    'End If
  Next i

End Sub


Private Sub picActiveTabView_Click()
Dim i As Integer

  i = 1
  
End Sub


'Private Sub picFieldModification_Resize()
'
'  If picFieldModification.Width < 7815 Then
'    cmdModifyMore.Left = picFieldModification.Width - (cmdModifyMore.Width + 40)
'    ckboxCaseSensitive.ZOrder 1
'    cmdModifyMore.ZOrder 0
'    cmdModifyMore.Visible = True
'  Else
'    cmdModifyMore.Visible = False
'  End If
'
'End Sub

Private Sub picSearchOptions_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

  If Button = 2 Then
    PopupMenu mnuSearchDialog
  End If

End Sub

Private Sub picSearchOptions_Paint()

  lvSearchArguments.Refresh
  Label1.Refresh
  Label2.Refresh
  Label3.Refresh
  Label4.Refresh
  ckboxCaseSensitive.Refresh
  lblObjectsSearching.Refresh
  cboxConditions.Refresh
  cboxProperties.Refresh
  cboxValue.Refresh
  cmdAnd.Refresh
  cmdOr.Refresh

End Sub


'rather then dealing with the internal controls outside of this container object,
'I'll do it here.
Private Sub picSearchOptions_Resize()

  lvSearchArguments.Width = picSearchOptions.Width - 40
  lvSearchArguments.Refresh
  
  'If some controls are being clipped, make sure user knows they're there
  If picSearchOptions.Width < 7815 Then
    cmdMore.Left = picSearchOptions.Width - (cmdMore.Width + 40)
    ckboxCaseSensitive.ZOrder 1
    cmdMore.ZOrder 0
    cmdMore.Visible = True
  Else
    cmdMore.Visible = False
  End If

End Sub


'///////////////////
'Private Methods
'///////////////////
Private Sub Form_Load()
Dim i As Integer
Dim nNode As Node
Dim sTFValue As String
Dim InsertQuery As colQueryList
Dim sKey As String
Dim iWinState As Integer

  lCurrentWidth = 0
  lCurrentHeight = 0
  lOldWidth = 0
  lOldHeight = 0

  LoadResStrings Me
  Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
  Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
  Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
  Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
  
  iWinState = GetSetting(App.Title, "Settings", "WinState", vbNormal)
  
  If Not iWinState = vbMinimized Then
    Me.WindowState = iWinState
  End If
  
  imgVSplitter.Left = GetSetting(App.Title, "Settings", "VSplitter", 1500)
  imgVSplitter.Width = SPLITTERWIDTH
  
  sDate = ""
  sTime = ""
  
  modDatabase.InitializeDB
  
  'Setup assigned queries
  For i = 1 To iMaxAssignedCount
    Set InsertQuery = New colQueryList
    sKey = "Saved" & Trim(Str(i))
    'InsertQuery.SaveName = sEmptyString
    InsertQuery.SaveName = GetSetting(App.Title, "Settings", sKey, sEmptyString)
    AssignedQueries.Add InsertQuery, InsertQuery.SaveName & Trim(Str(i))
    Set InsertQuery = Nothing
  Next i
  LoadAssignedQueries
  
  'Setup RecentQueries list
  iRecentQueriesCount = GetSetting(App.Title, "Settings", "RecentQueries", 5)
  For i = 1 To iRecentQueriesCount
    Set InsertQuery = New colQueryList
    InsertQuery.ResetCollection
    RecentQueries.Add InsertQuery
    Set InsertQuery = Nothing
    iCurrentRecentQuery = i
  Next i
  iNumberInRecentQueryList = 0
  tbMainToolbar.Buttons("PreviousQuery").Enabled = False
  tbMainToolbar.Buttons("NextQuery").Enabled = False
  
  sDefaultSearchType = GetSetting(App.Title, "Options", "DefaultQueryType", TYPE_AL)
  
  SetupSearchParams sDefaultSearchType, True
  
  LoadDefaultQuery (GetSetting(App.Title, "Options", "DefaultQueryName", sEmptyString))
  
  sTFValue = GetSetting(App.Title, "Settings", "SearchOpen", False)
  Select Case sTFValue
    Case "True"
      bSearchViewOpen = False
    Case "False"
      bSearchViewOpen = True
  End Select
  SwitchSearchView

  
  For i = 1 To lvSearchArguments.ColumnHeaders.Count
    lvSearchArguments.ColumnHeaders(i).Width = GetSetting(App.Title, "Settings", "QColumn" & Trim(Str(i)), 1440)
  Next i
  
  If ARCom.IsConnected = False Then
    Me.tbMainToolbar.Buttons("Execute").Enabled = False
  Else
    Me.tbMainToolbar.Buttons("Execute").Enabled = True
  End If
  
  Me.Show
  Me.Refresh
        
End Sub


Private Sub Form_Paint()

    lvListView.View = Val(GetSetting(App.Title, "Settings", "ViewMode", "0"))
    tbMainToolbar.Refresh
    picSearchOptions.Refresh
    Me.Refresh

End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer
Dim QueryIndex As colQueryList
Dim bTFValue As Boolean

  If GetSetting(App.Title, "Options", "SaveCacheOnExit", True) = False Then
    modDatabase.DeleteCache (lCurrentServerID)
  End If

  'close all sub forms
  For i = Forms.Count - 1 To 1 Step -1
      Unload Forms(i)
  Next
  If Me.WindowState = vbNormal Then
      SaveSetting App.Title, "Settings", "MainLeft", Me.Left
      SaveSetting App.Title, "Settings", "MainTop", Me.Top
      SaveSetting App.Title, "Settings", "MainWidth", Me.Width
      SaveSetting App.Title, "Settings", "MainHeight", Me.Height
  Else
      SaveSetting App.Title, "Settings", "MainLeft", lOldLeft
      SaveSetting App.Title, "Settings", "MainTop", lOldTop
      SaveSetting App.Title, "Settings", "MainWidth", lOldWidth
      SaveSetting App.Title, "Settings", "MainHeight", lOldHeight
  End If
  
  SaveSetting App.Title, "Settings", "WinState", Me.WindowState
  
  SaveSetting App.Title, "Settings", "ViewMode", lvListView.View
  SaveSetting App.Title, "Settings", "VSplitter", imgVSplitter.Left
  
  If GetSetting(App.Title, "Options", "RememberServerName", True) = True Then
    If sCurrentServerName = "AR Explorer - <NOT CONNECTED>" Then
      sCurrentServerName = sEmptyString
    Else
      SaveSetting App.Title, "Settings", "LastServerUsed", sCurrentServerName
    End If
  Else
    SaveSetting App.Title, "Settings", "LastServerUsed", GetSetting(App.Title, "Options", "DefaultServerName", sCurrentServerName)
  End If
  
  SaveColumnHeaders (sCurrentSearchType)
  
  'What the HELL is this?  QColumn?  What WAS I thinking?
'  For i = 1 To lvSearchArguments.ColumnHeaders.Count
'    SaveSetting App.Title, "Settings", "QColumn" & Trim(Str(i)), lvSearchArguments.ColumnHeaders(i).Width
'  Next i
  
  Select Case tbMainToolbar.Buttons("ObjectType").Image
  Case icoActiveLink
    SaveSetting App.Title, "Settings", "LastSearched", TYPE_AL
  Case icoFilter
    SaveSetting App.Title, "Settings", "LastSearched", TYPE_FILTER
  Case icoFields
    SaveSetting App.Title, "Settings", "LastSearched", TYPE_FIELD
  End Select
  
  SaveSetting App.Title, "Settings", "SearchOpen", bSearchViewOpen
  
  For i = 1 To iMaxAssignedCount
    SaveSetting App.Title, "Settings", "Saved" & Trim(Str(i)), AssignedQueries.Item(i).SaveName
  Next i
  
  SaveSetting "ARE", "Settings", LastUsedDate, ARCom.ConvertDate(Format(Now, "mm/dd/yy") & " " & Format(Now, "hh:mm:ss AM/PM"))
  
  modDatabase.CloseDB
  
End Sub


Private Sub Form_Resize()

  On Error Resume Next

  'added to stop the reset position of the vsplitter bar.
  If Not Me.WindowState = vbMinimized Then
  
    lOldTop = lCurrentTop
    lOldLeft = lCurrentLeft
    lOldWidth = lCurrentWidth
    lOldHeight = lCurrentHeight
    
    lCurrentTop = Me.Top
    lCurrentLeft = Me.Left
    lCurrentWidth = Me.Width
    lCurrentHeight = Me.Height
    
    If Me.Width < 3000 Then
      Me.Width = 3000
    End If
      
    SizeControls imgVSplitter.Left
    
    If bSearchViewOpen = True Then
      SizeListView (MAXVIEWHEIGTH)
    Else
      SizeListView 0
    End If
    
  End If
  
  On Error GoTo 0
    
End Sub


Private Sub SwitchSearchView(Optional bOverideSwitch As Boolean)
Dim bLocked As Boolean
  
  bLocked = GetSetting(App.Title, "Options", "LockedSearchDialog", False)

  'probably not the best way to do it, but it's late and i'm tired =)
  'bLocked is the "Lock Search Dialog Open" option.
  If bLocked = True Then
    bSearchViewOpen = False
  End If
  
  If bOverideSwitch = True Then
    If bSearchViewOpen = True Then
      SizeListView (MAXVIEWHEIGTH)
    Else
      SizeListView (0)
    End If
  Else
    ' If the Search Dialog is closed then re-size the list view to the top of the window.
    If bSearchViewOpen = True Then
      tbMainToolbar.Buttons("ShowDialog").ToolTipText = "Show Interface Dialog"
      tbMainToolbar.Buttons("ShowDialog").Image = icoShowSearch
      mnuEditSearchDialog.Caption = "Show Interface Dialog"
      SizeListView (0)
      bSearchViewOpen = False
    Else
      'If the Search Dialog is open, then re-size the list view to the bottom of the Search Dialog
      tbMainToolbar.Buttons("ShowDialog").ToolTipText = "Hide Interface Dialog"
      tbMainToolbar.Buttons("ShowDialog").Image = icoHideSearch
      mnuEditSearchDialog.Caption = "Hide Interface Dialog"
      SizeListView (MAXVIEWHEIGTH)
      bSearchViewOpen = True
    End If
  End If
  
  If bLocked = True Then
    tbMainToolbar.Buttons("ShowDialog").Enabled = False
    mnuEditSearchDialog.Enabled = False
    mnuSearchDialogHide.Enabled = False
  Else
    tbMainToolbar.Buttons("ShowDialog").Enabled = True
    mnuEditSearchDialog.Enabled = True
    mnuSearchDialogHide.Enabled = True
  End If
  
End Sub


Private Sub imgVSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With imgVSplitter
      picSplitter.Left = .Left
      picSplitter.Top = .Top
      picSplitter.Width = .Width - 20
      picSplitter.Height = .Height

      picSplitter.Move .Left, .Top, .Width, .Height - 20
    End With
    picSplitter.Visible = True
    mbMoving = True
End Sub


Private Sub imgVSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim sglPos As Single

  If mbMoving Then
    picSplitter.ZOrder (0)
    sglPos = X + imgVSplitter.Left
    If sglPos < sglSplitLimit Then
      picSplitter.Left = sglSplitLimit
    ElseIf sglPos > Me.Width - sglSplitLimit Then
      picSplitter.Left = Me.Width - sglSplitLimit
    Else
      picSplitter.Left = sglPos
    End If
  End If

End Sub


Private Sub imgVSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SizeControls picSplitter.Left
    picSplitter.Visible = False
    mbMoving = False
End Sub


Private Sub EnableModifyOnly()

  tbMainToolbar.Buttons("ResetQuery").Enabled = True
  tbMainToolbar.Buttons("SaveQuery").Enabled = True
  tbMainToolbar.Buttons("OpenQuery").Enabled = True
  tbMainToolbar.Buttons("DeleteQuery").Enabled = True
  tbMainToolbar.Buttons("PreviousQuery").Enabled = True
  tbMainToolbar.Buttons("NextQuery").Enabled = True
  cboxSavedQueries.Enabled = True

End Sub








Private Sub tbDatabaseName_Change()
Dim iCurrentPosition As Integer

  If Len(tbDatabaseName.Text) > 30 Then
    iCurrentPosition = tbDatabaseName.SelStart
    If iCurrentPosition > 30 Then
      tbDatabaseName.Text = Left(tbDatabaseName.Text, 30)
      iCurrentPosition = 30
    Else
      tbDatabaseName.Text = Left(tbDatabaseName.Text, iCurrentPosition - 1) & Right(tbDatabaseName.Text, 30 - iCurrentPosition + 1)
      iCurrentPosition = iCurrentPosition - 1
    End If
    'tbName.Text = Left(tbName.Text, 30)
    tbDatabaseName.SelStart = iCurrentPosition
    tbDatabaseName.SelLength = 0
    Beep
  End If

'  If Len(tbDatabaseName.Text) > 30 Then
'    tbDatabaseName.Text = Left(tbDatabaseName.Text, 30)
'    tbDatabaseName.SelStart = 30
'    tbDatabaseName.SelLength = 0
'    Beep
'  End If

End Sub

Private Sub tbFieldLabel_Change()
Dim iCurrentPosition As Integer

  If Len(tbFieldLabel.Text) > 30 Then
    iCurrentPosition = tbFieldLabel.SelStart
    If iCurrentPosition > 30 Then
      tbFieldLabel.Text = Left(tbFieldLabel.Text, 30)
      iCurrentPosition = 30
    Else
      tbFieldLabel.Text = Left(tbFieldLabel.Text, iCurrentPosition - 1) & Right(tbFieldLabel.Text, 30 - iCurrentPosition + 1)
      iCurrentPosition = iCurrentPosition - 1
    End If
    'tbName.Text = Left(tbName.Text, 30)
    tbFieldLabel.SelStart = iCurrentPosition
    tbFieldLabel.SelLength = 0
    Beep
  End If

'  If Len(tbFieldLabel.Text) > 30 Then
'    tbFieldLabel.Text = Left(tbFieldLabel.Text, 30)
'    tbFieldLabel.SelStart = 30
'    tbFieldLabel.SelLength = 0
'    Beep
'  End If

End Sub

Private Sub tbHelpText_Change()
Dim i As Integer

  i = cboxHelpText.ListIndex
  
  If i = 0 Then
    i = 1
  End If
  
  If Len(tbHelpText.Text) > 0 Then
    cboxHelpText.ListIndex = i
  Else
    cboxHelpText.ListIndex = 0
  End If

End Sub

Private Sub tbExecutionOrder_Change()

  On Error GoTo ErrorHandler
  If Len(tbExecutionOrder) > 0 Then
    tbExecutionOrder.tag = CLng(tbExecutionOrder.Text)
    
    If tbExecutionOrder.tag > 1000 Then
      tbExecutionOrder.tag = 1000
      tbExecutionOrder.Text = 1000
    End If
    
    If tbExecutionOrder.tag < 0 Then
      tbExecutionOrder.tag = 0
      tbExecutionOrder.Text = 0
    End If
    
  End If
    
  On Error GoTo 0
  
  Exit Sub

ErrorHandler:
  tbExecutionOrder.Text = tbExecutionOrder.tag
  Beep
  Resume Next

End Sub

'Private Sub tabActiveLink_Click()
'Dim picActiveTabView As PictureBox
'
''  Set picActiveTabView = Nothing
'  Select Case tabActiveLink.SelectedItem.Index
'  Case BasicTab
'    Set picActiveTabView = picALBasic
'  Case PermissionsTab
'    Set picActiveTabView = picALPermissions
'  Case MiscTab
'    Set picActiveTabView = picALMisc
'  End Select
'
'  picActiveTabView.Visible = True
'  picActiveTabView.Enabled = True
'
'  picActiveTabView.ZOrder 0
'  picActiveTabView.Refresh
'
'End Sub


Private Sub tbMainToolbar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim i As Integer

  Select Case ButtonMenu.Key
  Case "ActiveLink"
    tbMainToolbar.Buttons("ObjectType").tag = TYPE_AL
    'tbMainToolbar.Buttons("ActionType").ButtonMenus("Modify").Enabled = True
    SetupSearchParams TYPE_AL, True
    SetupDialog ("Search")
  Case "Filter"
    tbMainToolbar.Buttons("ObjectType").tag = TYPE_FILTER
    'tbMainToolbar.Buttons("ActionType").ButtonMenus("Modify").Enabled = True
    SetupSearchParams TYPE_FILTER, True
    SetupDialog ("Search")
  Case "Field"
    tbMainToolbar.Buttons("ObjectType").tag = TYPE_FIELD
    'tbMainToolbar.Buttons("ActionType").ButtonMenus("Modify").Enabled = True
    SetupSearchParams TYPE_FIELD, True
    SetupDialog ("Search")
  Case "Search"
    SetupDialog ("Search")
  Case "Modify"
    SetupDialog ("Modify")
    
  End Select

End Sub


Private Sub tbMainToolbar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

  If Button = 2 Then
    PopupMenu mnuToolBar
  End If

End Sub

Private Sub tbName_Change()
Dim iCurrentPosition As Integer

  If Len(tbName.Text) > 30 Then
    iCurrentPosition = tbName.SelStart
    If iCurrentPosition > 30 Then
      tbName.Text = Left(tbName.Text, 30)
      iCurrentPosition = 30
    Else
      tbName.Text = Left(tbName.Text, iCurrentPosition - 1) & Right(tbName.Text, 30 - iCurrentPosition + 1)
      iCurrentPosition = iCurrentPosition - 1
    End If
    'tbName.Text = Left(tbName.Text, 30)
    tbName.SelStart = iCurrentPosition
    tbName.SelLength = 0
    Beep
  End If

End Sub

Private Sub tvTreeView_DragDrop(Source As Control, X As Single, Y As Single)

  If Source = imgVSplitter Then
      SizeControls X
  End If
    
End Sub


'There's still some bugs here.. gotta figure 'em out quick.. but not tonight.. tired
'Update:  I think I found most bugs and moved some control resizing to a more appropriate
'spot in the code.
'Final Update:  This all works great now, code has been optimized and put in more "Logical" areas.
'Mike: You pass in MAXVIEWHEIGHT to size lvListView to the bottom of the Search Dialog.
'Mike: You pass in 0 to size lvListView to the top of the window (i.e. no search dialog)
Sub SizeListView(Y As Integer)

  'Note:  Constants used are defined above
  On Error Resume Next
  
  picCurrentView.Top = tvTreeView.Top 'Position the top of the Search Dialog
  picCurrentView.Height = MAXVIEWHEIGTH 'Size the Search Dialog Height
    
  lvListView.Top = tvTreeView.Top + Y 'Position the top of Resulst List after bottom of Search Dialog
  lvListView.Height = (tvTreeView.Height - Y) / 2 'Size the Results List using leftover space from the Search Dialog
  'The next two lines place and size the Height of the Actions List window
  lvListViewActions.Top = lvListView.Top + lvListView.Height
  lvListViewActions.Height = lvListView.Height
  
'  Mike: I changed this next line so here is the backup copy:
'  lvListView.Height = tvTreeView.Height + (tvTreeView.Top - lvListView.Top)
' Mike: and here is the changed line:
  
  'Mike: This next line is the original code from the line above.
  'picCurrentView.Height = tvTreeView.Height - (lvListView.Height)
  
  If Y = 0 Then
    picCurrentView.Visible = False
  Else
    picCurrentView.Visible = True 'Make the Search Dialog Visible
  End If
  
      
  On Error GoTo 0
  
End Sub


Sub SizeControls(X As Single)

  'Note:  Constants used are defined above
  On Error Resume Next

  'set the width
  If X < MINVIEWWIDTH Then X = MINVIEWWIDTH
  If X > (Me.Width - MAXVIEWWIDTH) Then X = Me.Width - MAXVIEWWIDTH
  tvTreeView.Width = X 'Set left position of the Tree View
  imgVSplitter.Left = X 'Set left position of the window splitter
  lvListView.Left = X + SPLITTERWIDTH  'Set the left position of the Results List
  lvListViewActions.Left = lvListView.Left 'Set the left position of the Actions List
  
  picCurrentView.Left = lvListView.Left + 20
  
  lvListView.Width = Me.Width - (tvTreeView.Width + SPLITTERWIDTH + 100) 'Set width of Results List window
  lvListViewActions.Width = lvListView.Width 'Set width of Action List window
  
  picTitles.Width = Me.Width
  
  picCurrentView.Width = lvListView.Width
  
  lblTreeViewTitle(0).Width = tvTreeView.Width
  lblListViewTitle(1).Left = lvListView.Left + 20
  lblListViewTitle(1).Width = lvListView.Width - 40
  
  'set the top of tvTreeView
  If tbMainToolbar.Visible Then
      tvTreeView.Top = tbMainToolbar.Height + picTitles.Height
  Else
      tvTreeView.Top = picTitles.Height
  End If

  'set the height of tvTreeView
  If sbMainStatusBar.Visible Then
      tvTreeView.Height = Me.ScaleHeight - (picTitles.Top + picTitles.Height + sbMainStatusBar.Height)
  Else
      tvTreeView.Height = Me.ScaleHeight - (picTitles.Top + picTitles.Height)
  End If
  
  imgVSplitter.Top = tvTreeView.Top
  imgVSplitter.Height = tvTreeView.Height
  
  picCurrentView.Refresh
    
  On Error GoTo 0
    
End Sub



'***************************
'ToolBar stuff goes here
'***************************
Private Sub tbMainToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Integer
Dim sMSG As String
  
  ResetStatusMessage
  
  Select Case Button.Key
    Case "Connection"
      ReverseConnection
    Case "ResetQuery"
      Select Case tbMainToolbar.Buttons("ActionType").tag
      Case "Search"
        ResetQuery
      Case "Modify"
        ResetModifyDialog
      End Select
    Case "SaveQuery"
'      If ARQuery.Saved = True Then
'        'simply updating currently saved query.
'        Call SaveQuery(ARQuery.SaveName, True, ARQuery.SavedID)
'      Else
        frmQueryIO.SetupForm (QUERY_SAVE)
        frmQueryIO.Show vbModal
'      End If
    Case "SaveResults"
      ExportResults
    Case "OpenQuery"
      frmQueryIO.SetupForm (QUERY_OPEN)
      frmQueryIO.Show vbModal
    Case "DeleteQuery"
      frmQueryIO.SetupForm (QUERY_DELETE)
      frmQueryIO.Show vbModal
    Case "PrintResults"
      PrintResults
    Case "ShowDialog"
      SwitchSearchView
    Case "ObjectType"
      'Umm.. we do nothing? (since it's covered in the menudropdown click event)
      'tbMainToolbar.Buttons("ObjectType").ButtonMenus.Item (1)
    Case "ActionType"
    Case "Execute"
      Select Case tbMainToolbar.Buttons("ActionType").tag
      Case "Search"
        AddToRecentQueryList (True)
        Search
      Case "Modify"
        Modify
      End Select
'    Case "SearchAL"
'      SetupSearchParams TYPE_AL, True
'    Case "SearchFilter"
'      SetupSearchParams TYPE_FILTER, True
'    Case "PerformSearch"
'      AddToRecentQueryList (True)
'      Search
    Case "PreviousQuery"
      ShowPreviousQuery
    Case "NextQuery"
      ShowNextQuery
  End Select
    
End Sub


Private Sub ResetQuery()

  'AddToRecentQueryList
  ARQuery.ResetCollection
  
  ShowCurrentQuery
  
End Sub


Private Sub CopyARQueryList(ByRef Source As colQueryList, ByRef Dest As colQueryList)
Dim qiQueryItem As clsQueryItem
Dim qiTempQueryItem As clsQueryItem
Dim i As Integer

  Dest.CaseSensitive = Source.CaseSensitive
  Dest.DialogIndex = Source.DialogIndex
  Dest.Dirty = False
  Dest.ExecuteOnANDORValue = Source.ExecuteOnANDORValue
  Dest.Saved = Source.Saved
  Dest.SavedID = Source.SavedID
  Dest.SaveName = Source.SaveName
  Dest.SavePath = Source.SavePath
  Dest.SearchServerName = Source.SearchServerName
  Dest.SearchType = Source.SearchType
  For i = 1 To Source.Count
    Set qiQueryItem = Source.Item(i)
    Set qiTempQueryItem = Dest.Add(qiQueryItem.SearchType, qiQueryItem.SearchParam, qiQueryItem.SearchValueString, _
      qiQueryItem.SearchValueNum, qiQueryItem.SearchCondition, qiQueryItem.SearchConditionString, qiQueryItem.tag)
  Next i

End Sub


Private Sub AddToRecentQueryList(Optional bKeepCurrentQuery As Boolean)
Dim qlQueryList As New colQueryList

'  iNumberInRecentQueryList = iNumberInRecentQueryList + 1
'  If iNumberInRecentQueryList > iRecentQueriesCount Then
'    iNumberInRecentQueryList = iRecentQueriesCount
'  End If
  iNumberInRecentQueryList = iNumberInRecentQueryList - 1
  If iNumberInRecentQueryList < 0 Then
    iNumberInRecentQueryList = 0
  End If
  
  
  'since our list will be reversed, 1 will be the oldest and iRecentQueriesCount will be newest
  RecentQueries.Remove (1)
  CopyARQueryList ARQuery, qlQueryList
  
  'Set qlQueryList = ARQuery
  RecentQueries.Add qlQueryList
  iCurrentRecentQuery = RecentQueries.Count
  tbMainToolbar.Buttons("PreviousQuery").Enabled = True
  tbMainToolbar.Buttons("NextQuery").Enabled = False
  
  If bKeepCurrentQuery = False Then
    Set ARQuery = New colQueryList
  End If
  
  Set qlQueryList = Nothing

End Sub


Private Sub ShowNextQuery()
  
  'If iCurrentRecentQuery > iNumberInRecentQueryList Then
  If iCurrentRecentQuery < iRecentQueriesCount Then
    iCurrentRecentQuery = iCurrentRecentQuery + 1
    Set ARQuery = RecentQueries(iCurrentRecentQuery)
  End If
  
  tbMainToolbar.Buttons("PreviousQuery").Enabled = True
  
  'If iCurrentRecentQuery = iNumberInRecentQueryList Then
  If iCurrentRecentQuery = iRecentQueriesCount Then
    tbMainToolbar.Buttons("NextQuery").Enabled = False
    'iCurrentRecentQuery = iNumberInRecentQueryList + 1
    iCurrentRecentQuery = iRecentQueriesCount + 1
  End If
  
  ShowCurrentQuery

End Sub


Private Sub ShowPreviousQuery()

  If iCurrentRecentQuery = iRecentQueriesCount Then
'    AddToRecentQueryList
  ElseIf iCurrentRecentQuery > iRecentQueriesCount Then
    iCurrentRecentQuery = iRecentQueriesCount
  End If
  
  If iCurrentRecentQuery > 1 Then
    iCurrentRecentQuery = iCurrentRecentQuery - 1
    Set ARQuery = RecentQueries(iCurrentRecentQuery)
  End If
  
  tbMainToolbar.Buttons("NextQuery").Enabled = True
  
  If iCurrentRecentQuery = 1 Then 'Or iCurrentRecentQuery = (iRecentQueriesCount - iNumberInRecentQueryList) Then
    tbMainToolbar.Buttons(PreviousQueryNumber).Enabled = False
  End If

  ShowCurrentQuery

End Sub


Public Function CheckIfAssigned(iIndex As Integer) As Boolean

  If AssignedQueries(iIndex).SaveName = sEmptyString Then
    CheckIfAssigned = False
  Else
    CheckIfAssigned = True
  End If

End Function


Public Sub RemoveAssignedQuery(sQueryName As String)
Dim i As Integer

  For i = 1 To iMaxAssignedCount
    If AssignedQueries(i).SaveName = sQueryName Then
      AssignedQueries(i).ResetCollection
    End If
  Next i

End Sub


Public Sub RemoveQueryFromList(sQueryName As String)
Dim i As Integer
Dim j As Integer

'  For i = 1 To iRecentQueriesCount
'    If RecentQueries(i).SaveName = sQueryName Then
'      RecentQueries.Remove (i)
'      If i < iRecentQueriesCount Then
''        For j = i To iRecentQueriesCount - 1
''          Set RecentQueries(j) = RecentQueries(j + 1)
''        Next j
'      End If
'      RecentQueries(iRecentQueriesCount).ResetCollection
'      'if the query they deleted is also a RecentQuery need to reset CurrentRecentQuery to 0
'      'Basically reset recent list
'      If iCurrentRecentQuery = i Then
'        iCurrentRecentQuery = 0
'      End If
'    End If
'  Next i
'
  If ARQuery.SaveName = sQueryName Then
    ARQuery.SaveName = sEmptyString
    ARQuery.Saved = False
    ARQuery.Dirty = True
    'ARQuery.ResetCollection
  End If

End Sub


Private Sub ReverseConnection()
Dim i As Long
Dim lCount As Long

  If ARCom.IsConnected Then
    ARCom.DisconnectFromServer
    tvTreeView.Nodes.Clear
    
    sCurrentServerName = "AR Explorer - <NOT CONNECTED>"
    SetStatusMessage ("<NOT CONNECTED>")
    SetFormDisconnected
    ARCom.Logout
  Else
    Do
      frmLogin.Show vbModal
    Loop Until frmLogin.OK = True Or frmLogin.Cancel = True
    
            
    If frmLogin.Cancel = True Then
      sCurrentServerName = "AR Explorer - <NOT CONNECTED>"
      SetStatusMessage ("<NOT CONNECTED")
      SetFormDisconnected
    Else
      sCurrentServerName = frmLogin.txtServerName
      SetFormConnected
      frmMain.PopulateTree
      ResetStatusMessage
    End If
    
    Unload frmLogin

  End If

End Sub


Public Sub LoadAssignedQueries()
Dim i As Integer
Dim QueryIndex As colQueryList
  
  cboxSavedQueries.Clear
  
  For i = 1 To iMaxAssignedCount
    Set QueryIndex = AssignedQueries.Item(i)
    If Not (QueryIndex.SaveName = sEmptyString) Then
      QueryIndex.RemoveAll
      If modDatabase.OpenQuery(QueryIndex) = False Then
        QueryIndex.SaveName = sEmptyString
      End If
    Else
      QueryIndex.ResetCollection
    End If
    cboxSavedQueries.AddItem QueryIndex.SaveName
  Next i
  
  cboxSavedQueries.ToolTipText = "Display query: " & AssignedQueries(1).SaveName
  
  mnuSearchAssigned1.Caption = AssignedQueries(1).SaveName
  mnuSearchAssigned2.Caption = AssignedQueries(2).SaveName
  mnuSearchAssigned3.Caption = AssignedQueries(3).SaveName
  mnuSearchAssigned4.Caption = AssignedQueries(4).SaveName
  mnuSearchAssigned5.Caption = AssignedQueries(5).SaveName

End Sub


Public Sub SetCurrentQuery(ByRef tempQuery As colQueryList)
Dim i As Integer
Dim sMSG As String
Dim bOkToSet As Boolean

  'AddToRecentQueryList

  If (ARQuery.Saved = True) And (ARQuery.Dirty = True) Then
    sMSG = "Do you wish to save changes to the current query?"
    i = MsgBox(sMSG, vbYesNoCancel + vbQuestion, "Saved Query Changed")
    
    bOkToSet = True
    
    If i = vbCancel Then
      bOkToSet = False
    ElseIf i = vbYes Then
      're-save query
      SaveQuery ARQuery.SaveName, True ' modDatabase.GetFormCacheID(ARQuery.SaveName)
    End If
    
  Else
    bOkToSet = True
  End If
  
  'tempQuery.SaveName
  If bOkToSet = True Then
    Set ARQuery = tempQuery
    ARQuery.Dirty = False
    ShowCurrentQuery
  End If

End Sub


Private Sub LoadDefaultQuery(sName As String)

  ARQuery.ResetCollection
  ARQuery.SaveName = sName
  ARQuery.Saved = True
  
  If modDatabase.OpenQuery(ARQuery) = True Then
    ShowCurrentQuery
  Else
    ARQuery.ResetCollection
  End If

End Sub


Public Function OpenQuery(sName As String) As Boolean
Dim qryTemp As colQueryList

  If Len(sName) > 0 Then
  
    'AddToRecentQueryList
  
    Set qryTemp = ARQuery
    ARQuery.ResetCollection
    ARQuery.SaveName = sName
    ARQuery.Saved = True
    If modDatabase.OpenQuery(ARQuery) = True Then
      ShowCurrentQuery
      OpenQuery = True
    Else
      Set ARQuery = qryTemp
      OpenQuery = False
    End If
  Else
    OpenQuery = False
  End If

End Function


Public Function SaveQuery(sName As String, Optional bDuplicate As Boolean, Optional lDupQueryKey As Long) As Boolean
Dim i As Integer
Dim sMSG As String

  If bDuplicate = True Then
    modDatabase.DeleteQuery (sName) 'lDupQueryKey)
  End If
  
  ARQuery.SaveName = sName
  
  If modDatabase.SaveQuery(ARQuery) = True Then
    SaveQuery = True
  Else
    Beep
    sMSG = "The query '" & ARQuery.SaveName & "' could not be saved successfully"
    i = MsgBox(sMSG, vbOKOnly + vbCritical, "Error..")
    SaveQuery = False
  End If

End Function


Private Sub ShowCurrentQuery()
Dim i As Integer
Dim liItem As ListItem
Dim qryiItem As clsQueryItem

  lvSearchArguments.ListItems.Clear
  
  Call SetupSearchParams(ARQuery.SearchType, False)
  
  If ARQuery.CaseSensitive = True Then
    ckboxCaseSensitive.Value = 1
  Else
    ckboxCaseSensitive.Value = 0
  End If
  
  If ARQuery.Count > 0 Then
    For i = 1 To ARQuery.Count
      Set qryiItem = ARQuery.Item(i)
      
      Set liItem = lvSearchArguments.ListItems.Add()
  
      'Set the query as the users see it
      liItem.Text = qryiItem.SearchType 'cboxProperties.Text
      liItem.SubItems(1) = qryiItem.SearchConditionString 'cboxConditions.Text
      liItem.SubItems(2) = qryiItem.SearchValueString 'cboxValue.Text
      liItem.SubItems(3) = qryiItem.SearchCondition '"AND"
      liItem.tag = GetIdealPosition(liItem.Text, liItem.SubItems(2))
    Next i
  End If
  
  lblListViewTitle(1).Caption = "Query Name:  " & ARQuery.SaveName
  
  ShowQueryInStatBar

End Sub



'***************************
'Menu stuff goes here
'***************************
Private Sub mnuHelpAbout_Click()

  frmAbout.Show vbModal, Me
    
End Sub



Private Sub mnuHelpSearchForHelpOn_Click()
Dim nRet As Integer

  'if there is no helpfile for this project display a message to the user
  'you can set the HelpFile for your application in the
  'Project Properties dialog
  If Len(App.HelpFile) = 0 Then
    MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
  Else
    On Error Resume Next
    nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
    If Err Then
      MsgBox Err.Description
    End If
    On Error GoTo 0
  End If

End Sub


Private Sub mnuHelpContents_Click()
    Dim nRet As Integer

    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.Path & "\AREXPLORER.HLP", 3, 0)
        'nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
        On Error GoTo 0
    End If

End Sub


Private Sub mnuSearchObjects_Click()

  mnuModifyExecute.Enabled = False
  tbMainToolbar.Buttons("ObjectType").tag = TYPE_AL
'  tbMainToolbar.Buttons("ActionType").ButtonMenus("Modify").Enabled = False
  SetupSearchParams TYPE_AL, True
  SetupDialog ("Search")
  
End Sub


Private Sub mnuToolsOptions_Click()

  frmOptions.Show vbModal, Me
    
End Sub


Private Sub mnuViewOptions_Click()

  frmOptions.Show vbModal
  
  If bSearchViewOpen = True Then
    bSearchViewOpen = False
  Else
    bSearchViewOpen = True
  End If
  
  SwitchSearchView
  
End Sub


Private Sub mnuViewRefresh_Click()

  Me.Refresh
  tbMainToolbar.Refresh
    
End Sub


Private Sub mnuViewStatusBar_Click()

  mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
  sbMainStatusBar.Visible = mnuViewStatusBar.Checked
  SizeControls imgVSplitter.Left
  
End Sub


Private Sub mnuEditCut_Click()

  DeleteSearchItem

End Sub


Private Sub mnuFileClose_Click()

  'unload the form
  Unload Me

End Sub


Private Sub mnuFileDelete_Click()

  frmQueryIO.SetupForm (QUERY_DELETE)
  frmQueryIO.Show vbModal
      
End Sub


Private Sub mnuFileNew_Click()

  ResetQuery

End Sub


Private Sub mnuFileOpen_Click()

  'AddToRecentQueryList
  
  frmQueryIO.SetupForm (QUERY_OPEN)
  frmQueryIO.Show vbModal

'Dim sFile As String
'
'  With dlgCommonDialog
'    .DialogTitle = "Open"
'    .CancelError = False
'    'ToDo: set the flags and attributes of the common dialog control
'    .Filter = "All Files (*.*)|*.*"
'    .ShowOpen
'
'    If Len(.FileName) = 0 Then
'      Exit Sub
'    End If
'
'    sFile = .FileName
'  End With
'
'  'ToDo: add code to process the opened file

End Sub


'***************************
'Misc Form Stuff goes here
'***************************
Public Sub PopulateTree()
Dim i As Long
  
  bRebuildServerCache = GetSetting(App.Title, "Options", "UpdateCacheOnLoad", True)
  
  frmProgress.Caption = "Getting Server Information:"
  
  If bRebuildServerCache = True Then
    frmProgress.lblStatus = "Updating Cache"
    frmProgress.pbProgress2.Visible = True
  Else
    frmProgress.lblStatus = "Getting Form List"
    frmProgress.pbProgress2.Visible = False
  End If
  frmProgress.Show
  frmProgress.Refresh
  
  Call AddServer(ServerName)
  
  'Now that we've added the base, lets add the actual data
  GetAllForms

  frmProgress.Hide
  
  ResetStatusMessage

End Sub


Private Sub AddLVItemField(sName As String, sForm As String, sARID As String, sType As String)
Dim liItem As ListItem
Dim lisiSubItem As ListSubItem

  Set liItem = lvListView.ListItems.Add()

  liItem.Text = sName
  
  liItem.SubItems(1) = sForm
  liItem.SubItems(2) = sARID
  liItem.SubItems(3) = sType
  
End Sub


Private Sub AddLVItemFL(sName As String, sForm As String, sExecution As String, sEnabled As String, _
                      sMask As String, sModTime As String)
Dim liItem As ListItem
Dim lisiSubItem As ListSubItem

  Set liItem = lvListView.ListItems.Add()

  liItem.Text = sName
  
  liItem.SubItems(1) = sForm
  liItem.SubItems(2) = sModTime
  liItem.SubItems(3) = sMask
  liItem.SubItems(4) = sExecution
  liItem.SubItems(5) = sEnabled
  
End Sub


Private Sub AddLVItemAL(sName As String, sForm As String, sExecution As String, sEnabled As String, _
                      sMask As String, sModTime As String)
Dim liItem As ListItem
Dim lisiSubItem As ListSubItem
 
 'Add a List View item and store the reference to it locally
  Set liItem = lvListView.ListItems.Add()

  'Store the AL/FL name in the list item, created in the step above
  liItem.Text = sName
  
  liItem.SubItems(1) = sForm
  liItem.SubItems(2) = sModTime
  liItem.SubItems(3) = sMask
  liItem.SubItems(4) = sExecution
  liItem.SubItems(5) = sEnabled
  
End Sub


Private Sub AddTVItem(sParentText As String, sDisplayText As String, iIconIndex As Integer, lCacheID As Long, Optional IsVisible As Boolean)
Dim nNode As Node
Dim sKeyName As String
Dim sParentKey As String

  sKeyName = KEY_PREFIX & sDisplayText
  
  If (Len(sParentText) > 0) Then
    sParentKey = KEY_PREFIX & sParentText
    Set nNode = tvTreeView.Nodes.Add(sParentKey, tvwChild, sKeyName, sDisplayText, iIconIndex)
  Else
    Set nNode = tvTreeView.Nodes.Add(, , sKeyName, sDisplayText, iIconIndex)
  End If
      
  nNode.Sorted = True
  nNode.tag = lCacheID
      
  If (IsVisible) Then
    nNode.EnsureVisible
  End If
  
End Sub


Public Sub AddServer(sServerName As String)
Dim lServerID As Long

  lServerID = modDatabase.GetServerCacheID(sCurrentServerName)
  
  'if lServerID = 0 then server is NOT in cache.
  'if not cached the user MUST create the cache.
  If lServerID = 0 Then
    lServerID = modDatabase.AddServerToCache(sCurrentServerName)
    bRebuildServerCache = True
  End If
    
  lCurrentServerID = lServerID

  Call AddTVItem("", sServerName, ICON_SERVER, lCurrentServerID, True)
  Call AddTVItem(sServerName, ARFORMS, ICON_FORMS, 0, True)
'  Call AddTVItem(sServerName, ACTIVELINKS, ICON_ACTIVELINKS, True)
'  Call AddTVItem(sServerName, FILTERS, ICON_FILTERS, True)
'  Call AddTVItem(sServerName, ESCALATIONS, ICON_ESCALATIONS, True)
'  Call AddTVItem(sServerName, GUIDES, ICON_GUIDES, True)
'  Call AddTVItem(sServerName, APPLICATIONS, ICON_APPLICATIONS, True)
'  Call AddTVItem(sServerName, MENUS, ICON_MENUS, True)
'  Call AddTVItem(sServerName, GROUPS, ICON_GROUPS, True)
  
  tvTreeView.Nodes(KEY_PREFIX & sServerName).Selected = True
  
End Sub


Private Sub RemoveTVItem(sDisplayText As String)
Dim sKeyName As String

  sKeyName = KEY_PREFIX & sDisplayText
  tvTreeView.Nodes.Remove (sKeyName)

End Sub


'***************************
'Misc AR stuff goes here
'***************************
Public Function Connect(sServer As String, sUser As String, sPass As String) As Boolean
Dim i As Integer
Dim sMSG As String

  ServerName = sServer
  UserName = sUser
  Password = sPass
  
  If ARCom.ConnectToARServer(ServerName, UserName, Password) = True Then
    Connect = True
  Else
    sMSG = "Connot connect to server '" & sServer & "' "
    sMSG = sMSG & "with username '" & sUser & "'.  "
    sMSG = sMSG & "Please verify username and password."
    i = MsgBox(sMSG, vbOKOnly + vbInformation, "Error connecting to server.")
    Connect = False
  End If

End Function


Private Sub GetAllForms()
Dim i As Long
Dim sMSG As String
Dim lFormID As Long
Dim lNewFormID As Long
Dim lFormCount As Long
Dim sFormName As String
      
  'Get forms, if error occurs, return
  lFormCount = ARCom.GetNumberOfForms
  
  If lFormCount = 0 Then
    sMSG = ARCom.GetErrorText
    i = MsgBox(sMSG, vbCritical + vbOKOnly, "Error")
  Else
  
    frmProgress.lblStatus = "Found " & Trim(Str(lFormCount)) & " forms, populating tree."
    frmProgress.pbProgress.Max = lFormCount
    frmProgress.Refresh
    
    If lFormCount > 0 Then  'Excellent, we found forms!
    
      'tvTreeView.Nodes(ARFORMS).Sorted = True
      
      'Populate the Tree View with all forms found
      For i = 1 To lFormCount
        frmProgress.pbProgress.Value = i
        sFormName = ARCom.GetFormName(i)
        lFormID = modDatabase.GetFormCacheID(sFormName, lCurrentServerID)
        
        If bRebuildServerCache = True Then
          lNewFormID = UpdateFormCacheByModTime(sFormName, lFormID)
        Else
          lNewFormID = lFormID
        End If
        Call AddTVItem(ARFORMS, sFormName, ICON_FORMS, lNewFormID)
      Next i
      
    Else    'Oops, no forms were found?
      'To Do:  Raise Exception
      ARCom.ErrorResolved (True)
    End If
  
  End If
        
End Sub 'End of GetAllForms()




'***************************
'SearchView Stuff goes here
'***************************
Private Sub cboxProperties_Click()

  PrefillCondition (cboxProperties.Text)

End Sub


'Here we can make sure the user didn't change a prefilled value
Private Sub cboxValue_Validate(Cancel As Boolean)
Dim i As Integer
Dim sValueText As String
Dim bOk As Boolean

  If bEditOk = False Then
    bOk = False
    sValueText = cboxValue.Text
    For i = 0 To cboxValue.ListCount - 1
      If sValueText = cboxValue.List(i) Then
        bOk = True
      End If
    Next i
    If bOk = False Then
      cboxValue.Text = cboxValue.List(0)
      'Give an audible beep, if no sound card installed it will beep with PC speaker
      Beep
    End If
  End If

End Sub


Private Sub cmdOr_Click()
  
  AddSearchItem ("OR")
  ShowQueryInStatBar
  
End Sub


Private Function GetIdealPosition(sParamater As String, sValue As String) As Integer

  Select Case sParamater
  Case AR_ALNAME
    GetIdealPosition = AR_ALNAME_NUMBER
  Case AR_MODTIME
    GetIdealPosition = AR_MODTIME_NUMBER
  Case AR_ENABLEDDISABLED
    GetIdealPosition = AR_ENABLEDDISABLED_NUMBER
  Case AR_EXECUTIONORDER
    GetIdealPosition = AR_EXECUTIONORDER_NUMBER
  Case AR_FOCUSFIELDNAME
    GetIdealPosition = AR_FOCUSFIELDNAME_NUMBER
  Case AR_BUTTONNAME
    GetIdealPosition = AR_BUTTONNAME_NUMBER
  Case AR_RUNIFTEXT
    GetIdealPosition = AR_RUNIFTEXT_NUMBER
  Case AR_FILTERNAME
    GetIdealPosition = AR_FILTERNAME_NUMBER
  Case AR_NONE
    GetIdealPosition = AR_NONE_NUMBER
  Case AR_FIELDNAME
    GetIdealPosition = AR_FIELDNAME_NUMBER
  Case AR_FIELDID
    GetIdealPosition = AR_FIELDID_NUMBER
  Case AR_FIELDTYPE
    GetIdealPosition = AR_FIELDTYPE_NUMBER
  Case AR_EXECUTEON
    Select Case sValue
    Case AR_ENABLED
      GetIdealPosition = AR_ENABLED_NUMBER
    Case AR_DISABLED
      GetIdealPosition = AR_DISABLED_NUMBER
    Case AR_ONBUTTON
      GetIdealPosition = AR_ONBUTTON_NUMBER
    Case AR_ONRETURN
      GetIdealPosition = AR_ONRETURN_NUMBER
    Case AR_ONSUBMIT
      GetIdealPosition = AR_ONSUBMIT_NUMBER
    Case AR_ONMODIFY
      GetIdealPosition = AR_ONMODIFY_NUMBER
    Case AR_ONDISPLAY
      GetIdealPosition = AR_ONDISPLAY_NUMBER
'    Case AR_MODIFYALL
'      GetIdealPosition = AR_MODIFYALL_NUMBER
'    Case AR_MENUOPEN
'      GetIdealPosition = AR_MENUOPEN_NUMBER
    Case AR_MENUCHOICE
      GetIdealPosition = AR_MENUCHOICE_NUMBER
    Case AR_LOSEFOCUS
      GetIdealPosition = AR_LOSEFOCUS_NUMBER
    Case AR_SETDEFAULT
      GetIdealPosition = AR_SETDEFAULT_NUMBER
    Case AR_ONQUERY
      GetIdealPosition = AR_ONQUERY_NUMBER
    Case AR_AFTERMODIFY
      GetIdealPosition = AR_AFTERMODIFY_NUMBER
    Case AR_AFTERSUBMIT
      GetIdealPosition = AR_AFTERSUBMIT_NUMBER
    Case AR_GAINFOCUS
      GetIdealPosition = AR_GAINFOCUS_NUMBER
    Case AR_WINDOWOPEN
      GetIdealPosition = AR_WINDOWOPEN_NUMBER
    Case AR_WINDOWCLOSE
      GetIdealPosition = AR_WINDOWCLOSE_NUMBER
    Case AR_GET
      GetIdealPosition = AR_GET_NUMBER
    Case AR_DELETE
      GetIdealPosition = AR_DELETE_NUMBER
    Case AR_MERGE
      GetIdealPosition = AR_MERGE_NUMBER
    Case AR_NONE
      GetIdealPosition = AR_NONE_NUMBER
    End Select
  
  End Select

End Function


Private Function FindAddPosition(sParamater As String, sValue As String) As Integer
Dim i As Integer
Dim iFoundPosition As Integer
Dim iIdealPosition As Integer

  iIdealPosition = GetIdealPosition(sParamater, sValue)
  iFoundPosition = 1
  
  For i = 1 To lvSearchArguments.ListItems.Count
    If lvSearchArguments.ListItems(i).tag < iIdealPosition Then
      iFoundPosition = i + 1
    End If
  Next i
  
  FindAddPosition = iFoundPosition
  
End Function

Private Sub SetExecuteOnCondition(sCondition As String)
Dim i As Integer

  For i = 1 To lvSearchArguments.ListItems.Count
    If lvSearchArguments.ListItems(i).Text = AR_EXECUTEON Then
      ARQuery.Item(lvSearchArguments.SelectedItem.Text & _
      lvSearchArguments.SelectedItem.SubItems(2)).SearchCondition = sCondition
      lvSearchArguments.ListItems(i).SubItems(3) = sCondition
    End If
  Next i

End Sub


Private Sub AddSearchItem(sAndOr As String)
Dim liItem As ListItem
Dim sValueString As String
Dim sValueNum As Long
Dim bOkToAdd As Boolean
Dim iDuplicateItemIndex As Integer
Dim j As Integer
Dim i As Integer
Dim sMSG As String
Dim iPosition As Integer
Dim sKey As String
Dim iLikeProperties As Integer
Dim bLikePropertyFound As Boolean

  bOkToAdd = True
  
  iDuplicateItemIndex = -1
  iLikeProperties = 0
  bLikePropertyFound = False
  
  'Search for duplicates
  For i = 1 To lvSearchArguments.ListItems.Count
  
    'Is a duplicate present?
    If lvSearchArguments.ListItems(i).Text = cboxProperties.Text Then
    
      'If we are searching for AR_EXECUTEON or ANY Field type further qualify the duplicate for Value match
      If cboxProperties.Text = AR_EXECUTEON Or tbMainToolbar.Buttons("ObjectType").tag = TYPE_FIELD Then
        
        'Research for the Value
        For j = i To lvSearchArguments.ListItems.Count
        
          'Do the values match?
          'Yes
          If LCase(cboxValue.Text) = LCase(lvSearchArguments.ListItems(j).ListSubItems(2).Text) Then
            bOkToAdd = False
            iDuplicateItemIndex = j
            'Reset this cuz we want to add exactly where the dup occured
            iLikeProperties = 0
            'Ensure we don't continue to add to the count
            bLikePropertyFound = True
          'No
          Else
            'If we've already counted how many like properties there are, don't count again
            'If bLikePropertyFound = False Then
              'iLikeProperties = iLikeProperties + 1
            'End If
            
          End If
          
        Next j 'Value duplication
        
      Else
        bOkToAdd = False
        iDuplicateItemIndex = i
        
      End If
      
      If bLikePropertyFound = False And (tbMainToolbar.Buttons("ObjectType").tag = TYPE_FIELD) Then
        iLikeProperties = iLikeProperties + 1
      End If
    End If
    
  Next i
  
'  If sCurrentSearchType = TYPE_FIELD Then
'    bOkToAdd = True
'    iDuplicateItemIndex = -1
'  Else
    If bOkToAdd = False Then
      
      sMSG = "Adding this item will overide a previous query paramater.  "
      sMSG = sMSG & "Do you wish to overide?"
      i = MsgBox(sMSG, vbYesNo + vbQuestion, "Warning:")
      If i = vbYes Then
        bOkToAdd = True
      End If
    End If
'  End If
  
  If Len(Trim(cboxValue.Text)) = 0 Then
    bOkToAdd = False
  End If
  
  If bOkToAdd Then
    
    ARQuery.Dirty = True
    
    If iDuplicateItemIndex >= 0 Then
      sKey = lvSearchArguments.ListItems(iDuplicateItemIndex).Text & _
        lvSearchArguments.ListItems(iDuplicateItemIndex).SubItems(2)
      lvSearchArguments.ListItems.Remove (iDuplicateItemIndex)
      ARQuery.Remove (sKey)
    End If
    
    iPosition = FindAddPosition(cboxProperties.Text, cboxValue.Text) + iLikeProperties
    
    If cboxProperties.Text = AR_EXECUTEON Then
      ARQuery.ExecuteOnANDORValue = sAndOr
      SetExecuteOnCondition (sAndOr)
      'cmdOr.Enabled = False
    End If

    Set liItem = Me.lvSearchArguments.ListItems.Add(iPosition)
  
    'Set the query as the users see it
    liItem.Text = cboxProperties.Text
    liItem.SubItems(1) = cboxConditions.Text
    liItem.SubItems(2) = cboxValue.Text
    liItem.SubItems(3) = sAndOr
    
    sKey = liItem.Text & liItem.SubItems(2)
    
    'Set the query as the ARcom object needs it
    sValueString = cboxValue.Text
    If cboxValue.tag = Empty Then
      sValueNum = 0
    Else
      sValueNum = cboxValue.tag
    End If
  
    ARQuery.Add cboxProperties.Text, cboxConditions.tag, sValueString, sValueNum, sAndOr, cboxConditions.Text, sKey
    
    liItem.tag = GetIdealPosition(liItem.Text, liItem.SubItems(2))
    
  Else
    Beep
  End If

End Sub


Private Sub ShowQueryInStatBar()
Dim i As Integer
Dim sQuery As String
Dim sCondition As String
Dim sValue As String
Dim liItem As ListItem
Dim sPreviousAndOr As String
Dim sPreviousProperty As String

  sPreviousAndOr = ""
  sPreviousProperty = ""
  sQuery = "("
  For i = 1 To lvSearchArguments.ListItems.Count
    Set liItem = lvSearchArguments.ListItems(i)
    
    If liItem.Text = sPreviousProperty Then
      If Len(sPreviousAndOr) > 0 Then
        sQuery = sQuery & " " & sPreviousAndOr & " "
      End If
    Else
      If Len(sPreviousProperty) > 0 Then
        sQuery = sQuery & ") " & sPreviousAndOr & " ("
      End If
    End If
    
    sQuery = sQuery & liItem.Text & " "
    sPreviousProperty = liItem.Text
    
    Select Case liItem.SubItems(1)
      Case AR_BEGINSWITH
        sCondition = "="
        sValue = "'" & Trim(liItem.SubItems(2)) & "*'"
      Case AR_CONTAINS
        sCondition = "="
        sValue = "'*" & Trim(liItem.SubItems(2)) & "*'"
      Case AR_ENDSWITH
        sCondition = "="
        sValue = "'*" & Trim(liItem.SubItems(2)) & "'"
      Case AR_GREATERTHAN
        sCondition = ">"
      Case AR_DATERANGE
        sCondition = " "
      Case AR_EXACTDATE
        sCondition = "="
      Case AR_EQUAL
        sCondition = "="
      Case AR_GREATERTHANOREQUAL
        sCondition = ">="
      Case AR_LESSTHAN
        sCondition = "<"
      Case AR_LESSTHANOREQUAL
        sCondition = "<="
    End Select
    sQuery = sQuery & sCondition & " "
    If Len(sValue) = 0 Then
      sQuery = sQuery & liItem.SubItems(2)
    Else
      sQuery = sQuery & sValue
    End If
    sPreviousAndOr = liItem.SubItems(3)
    'sQuery = sQuery & liItem.SubItems(3) & " "
    sValue = ""
    sCondition = ""
  Next i
  sQuery = sQuery & ")"
  
  SetStatusMessage (sQuery)

End Sub


Private Sub cmdAnd_Click()
  
  AddSearchItem ("AND")
  ShowQueryInStatBar
End Sub


Private Sub SetupFilterSearch()
Dim Column As ColumnHeader

  mnuListViewSortName.Caption = "Filter Name"
  mnuListViewSortEnabled.Visible = True
  mnuListViewSortMask.Visible = True
  mnuListViewSortMod.Visible = True
  mnuListViewSortExecute.Visible = True
  mnuListViewSortID.Visible = False
  mnuListViewSortDataType.Visible = False

  lvListView.ColumnHeaders.Clear
  
  Set Column = lvListView.ColumnHeaders.Add(1, "Name", "Filter Name", GetSetting(App.Title, "Settings", "FLColumn1", 2400))
  Set Column = lvListView.ColumnHeaders.Add(2, "FormName", "Form Name", GetSetting(App.Title, "Settings", "FLColumn2", 2400))
  Set Column = lvListView.ColumnHeaders.Add(3, "ModTime", "Modification Time", GetSetting(App.Title, "Settings", "FLColumn3", 2400))
  Set Column = lvListView.ColumnHeaders.Add(4, "ExecMask", "Execution Mask", GetSetting(App.Title, "Settings", "FLColumn4", 2400))
  Set Column = lvListView.ColumnHeaders.Add(5, "ExecOrder", "Execution Order", GetSetting(App.Title, "Settings", "FLColumn5", 2400))
  Set Column = lvListView.ColumnHeaders.Add(6, "Enabled", "Enabled", GetSetting(App.Title, "Settings", "FLColumn6", 2400))
  lvListView.ColumnHeaders(1).Position = GetSetting(App.Title, "Settings", "FLColumnPos1", 1)
  lvListView.ColumnHeaders(2).Position = GetSetting(App.Title, "Settings", "FLColumnPos2", 2)
  lvListView.ColumnHeaders(3).Position = GetSetting(App.Title, "Settings", "FLColumnPos3", 3)
  lvListView.ColumnHeaders(4).Position = GetSetting(App.Title, "Settings", "FLColumnPos4", 4)
  lvListView.ColumnHeaders(5).Position = GetSetting(App.Title, "Settings", "FLColumnPos5", 5)
  lvListView.ColumnHeaders(6).Position = GetSetting(App.Title, "Settings", "FLColumnPos6", 6)


  ckboxCaseSensitive.Enabled = True

  cboxProperties.Clear

  cboxProperties.AddItem (AR_FILTERNAME)
  cboxProperties.Text = AR_FILTERNAME
  cboxProperties.AddItem (AR_MODTIME)
  cboxProperties.AddItem (AR_ENABLEDDISABLED)
  cboxProperties.AddItem (AR_EXECUTIONORDER)
  cboxProperties.AddItem (AR_EXECUTEON)
  cboxProperties.AddItem (AR_RUNIFTEXT)


End Sub


Public Sub ResetModifyDialog()

  cboxEnabled.ListIndex = 0
  cboxSubmit.ListIndex = 0
  cboxEntryMode.ListIndex = 0
  cboxHelpText.ListIndex = 0
  tbDatabaseName.Text = ""
  tbFieldLabel.Text = ""
  tbName.Text = ""
  tbExecutionOrder.Text = ""
  cboxEnabled.ListIndex = 0
  lboxNoAccess.Clear
  lboxAccess.Clear
  cboxPermissionType.ListIndex = 0
  tbChangeHistory.Text = ""
  tbHelpText.Text = ""
  
End Sub


Public Sub PrefillPermissions()
Dim lCount As Long
Dim i As Long
Dim sMSG As String

  lboxNoAccess.Clear
  lboxAccess.Clear

  lCount = ARCom.LoadPermissions
  If lCount > 0 Then
    For i = 1 To lCount
      If ARCom.GroupOK = True Then
        lboxNoAccess.AddItem (ARCom.GetCurrentGroup)
      End If
      ARCom.GotoNextGroup
    Next i
  Else
    sMSG = "There was an error retreiving the permission group list."
    sMSG = sMSG & vbCrLf & ARCom.GetErrorText
    sMSG = sMSG & vbCrLf & "No permission modification(s) will be possible."
    i = MsgBox(sMSG, vbOKOnly + vbInformation, "Non Critical Error...")
  End If
  

End Sub


Public Sub SetupDialog(sMode As String)

  Select Case sMode
  Case "Search"
  
    tbMainToolbar.Buttons("ResetQuery").ToolTipText = "Reset Query"
    mnuModifyExecute.Enabled = False
  
    tbMainToolbar.Buttons("ResetQuery").Enabled = True
    mnuFileNew.Enabled = True
    tbMainToolbar.Buttons("SaveQuery").Enabled = True
    mnuFileSaveQuery.Enabled = True
    tbMainToolbar.Buttons("OpenQuery").Enabled = True
    mnuFileOpen.Enabled = True
    tbMainToolbar.Buttons("DeleteQuery").Enabled = True
    mnuFileDelete.Enabled = True
    'how to handle these two?
    tbMainToolbar.Buttons("PreviousQuery").Enabled = bPreviousEnabled
    tbMainToolbar.Buttons("NextQuery").Enabled = bNextEnabled
    
    cboxSavedQueries.Enabled = True
    mnuEditCut.Enabled = True
    mnuSearchPredefined.Enabled = True
    mnuSearchAssigned1.Enabled = True
    mnuSearchAssigned2.Enabled = True
    mnuSearchAssigned3.Enabled = True
    mnuSearchAssigned4.Enabled = True
    mnuSearchAssigned5.Enabled = True
    mnuSearchAssign.Enabled = True
    mnuSearchPerformSearch.Enabled = True
    
    tbMainToolbar.Buttons("ActionType").tag = "Search"
    tbMainToolbar.Buttons("ActionType").Image = icoPerformSearch
    tbMainToolbar.Buttons("ActionType").ToolTipText = "Search"
    tbMainToolbar.Buttons("ObjectType").ButtonMenus("ActiveLink").Enabled = True
    tbMainToolbar.Buttons("ObjectType").ButtonMenus("Filter").Enabled = True
    tbMainToolbar.Buttons("Execute").ToolTipText = "Execute Search."
  
    On Error Resume Next
    picCurrentView.Visible = False
    On Error GoTo 0
    picSearchOptions.Visible = True
    Set picCurrentView = picSearchOptions
  Case "Modify"
  
    ResetModifyDialog
  
    tbMainToolbar.Buttons("ResetQuery").ToolTipText = "Reset Modification Fields"
    
    mnuModifyExecute.Enabled = True
  
'    tbMainToolbar.Buttons("ResetQuery").Enabled = False
    mnuFileNew.Enabled = False
    tbMainToolbar.Buttons("SaveQuery").Enabled = False
    mnuFileSaveQuery.Enabled = False
    tbMainToolbar.Buttons("OpenQuery").Enabled = False
    mnuFileOpen.Enabled = False
    tbMainToolbar.Buttons("DeleteQuery").Enabled = False
    mnuFileDelete.Enabled = False
    
    bPreviousEnabled = tbMainToolbar.Buttons("PreviousQuery").Enabled
    bNextEnabled = tbMainToolbar.Buttons("NextQuery").Enabled
    tbMainToolbar.Buttons("PreviousQuery").Enabled = False
    tbMainToolbar.Buttons("NextQuery").Enabled = False
    
    cboxSavedQueries.Enabled = False
    mnuEditCut.Enabled = False
    mnuSearchPredefined.Enabled = False
    mnuSearchAssigned1.Enabled = False
    mnuSearchAssigned2.Enabled = False
    mnuSearchAssigned3.Enabled = False
    mnuSearchAssigned4.Enabled = False
    mnuSearchAssigned5.Enabled = False
    mnuSearchAssign.Enabled = False
    mnuSearchPerformSearch.Enabled = False
    
    tbMainToolbar.Buttons("ActionType").tag = "Modify"
    tbMainToolbar.Buttons("ActionType").Image = icoModify
    tbMainToolbar.Buttons("ActionType").ToolTipText = "Modify"
    'tbMainToolbar.Buttons("ObjectType").ButtonMenus("ActiveLink").Enabled = False
    'tbMainToolbar.Buttons("ObjectType").ButtonMenus("Filter").Enabled = False
    tbMainToolbar.Buttons("Execute").ToolTipText = "Execute Modification."
  
    On Error Resume Next
    picCurrentView.Visible = False
    On Error GoTo 0
    Select Case tbMainToolbar.Buttons("ObjectType").tag
    Case TYPE_AL
    
      If ARCom.IsConnected = True Then
        PrefillPermissions
      End If
      
      tabModify.TabVisible(0) = False
      tabModify.TabVisible(1) = True
      tabModify.TabVisible(2) = True
      tabModify.TabVisible(3) = True
      tabModify.TabVisible(4) = True
      lblObjectsModifying.Caption = "Active Links"
      Set picCurrentView = picModify
    Case TYPE_FILTER
      tabModify.TabVisible(0) = False
      tabModify.TabVisible(1) = True
      tabModify.TabVisible(2) = False
      tabModify.TabVisible(3) = True
      tabModify.TabVisible(4) = True
      lblObjectsModifying.Caption = "Filters"
      Set picCurrentView = picModify
    Case TYPE_FIELD
      tabModify.TabVisible(0) = True
      tabModify.TabVisible(1) = False
      tabModify.TabVisible(2) = False
      tabModify.TabVisible(3) = True
      tabModify.TabVisible(4) = True
      lblObjectsModifying.Caption = "Fields"
      Set picCurrentView = picModify
    End Select
    
  End Select
  
  SizeControls (imgVSplitter.Left)
  SwitchSearchView (True)

End Sub


Public Sub SetupSearchParams(sMode As String, bClearCurrentQuery As Boolean)
Dim i As Integer
Dim qryiTest As clsQueryItem
Dim sTest As String

  lvSearchArguments.ListItems.Clear
  
  'Don't want to lose the results
  If Not sMode = sCurrentSearchType Then
    lvListView.ListItems.Clear
  End If
  
  If bClearCurrentQuery = True Then
    ARQuery.ResetCollection
  End If
    
  ARQuery.SearchType = sMode
  
  SaveColumnHeaders (sCurrentSearchType)
  
  sCurrentSearchType = sMode
  
  SetupDialog (tbMainToolbar.Buttons("ActionType").tag)
  
  Select Case UCase(sMode)
    Case TYPE_AL
      tbMainToolbar.Buttons("ObjectType").ToolTipText = "Active Links"
      tbMainToolbar.Buttons("ObjectType").Image = icoActiveLink
      SetupALSearch
      lblObjectsSearching.Caption = "Active Links"
    Case TYPE_FILTER
      tbMainToolbar.Buttons("ObjectType").ToolTipText = "Filters"
      tbMainToolbar.Buttons("ObjectType").Image = icoFilter
      SetupFilterSearch
      lblObjectsSearching.Caption = "Filters"
    Case TYPE_FIELD
      tbMainToolbar.Buttons("ObjectType").ToolTipText = "Fields"
      tbMainToolbar.Buttons("ObjectType").Image = icoFields
      SetupFieldSearch
      lblObjectsSearching.Caption = "Fields"
  End Select
  
  tbMainToolbar.Buttons("ObjectType").tag = sMode
  'tbMainToolbar.Refresh

End Sub


Private Sub PrefillValue(sMode As String)

  cboxValue.tag = Empty
  cboxValue.Clear

  bEditOk = True
  
  Select Case sMode
    Case AR_MODTIME
      bEditOk = False
    Case AR_ENABLEDDISABLED
      cboxValue.AddItem (AR_ENABLED)
      cboxValue.Text = AR_ENABLED
        cboxValue.tag = 1
      cboxValue.AddItem (AR_DISABLED)
      bEditOk = False
    Case AR_EXECUTEON
      Select Case ARQuery.SearchType
        Case TYPE_AL
          cboxValue.AddItem (AR_ONBUTTON)
          cboxValue.Text = AR_ONBUTTON
            cboxValue.tag = 1
          cboxValue.AddItem (AR_ONRETURN)
          cboxValue.AddItem (AR_ONSUBMIT)
          cboxValue.AddItem (AR_ONMODIFY)
          cboxValue.AddItem (AR_ONDISPLAY)
    '      cboxValue.AddItem (AR_MODIFYALL)
    '      cboxValue.AddItem (AR_MENUOPEN)
          cboxValue.AddItem (AR_MENUCHOICE)
          cboxValue.AddItem (AR_LOSEFOCUS)
          cboxValue.AddItem (AR_SETDEFAULT)
          cboxValue.AddItem (AR_ONQUERY)
          cboxValue.AddItem (AR_AFTERMODIFY)
          cboxValue.AddItem (AR_AFTERSUBMIT)
          cboxValue.AddItem (AR_GAINFOCUS)
          cboxValue.AddItem (AR_WINDOWOPEN)
          cboxValue.AddItem (AR_WINDOWCLOSE)
          cboxValue.AddItem (AR_NONE)
        Case TYPE_FILTER
          cboxValue.AddItem (AR_ONSUBMIT)
          cboxValue.Text = AR_ONSUBMIT
          cboxValue.tag = 4
          cboxValue.AddItem (AR_ONMODIFY)
          cboxValue.AddItem (AR_GET)
          cboxValue.AddItem (AR_DELETE)
          cboxValue.AddItem (AR_MERGE)
        End Select
      bEditOk = False
    Case AR_FOCUSFIELDNAME, AR_BUTTONNAME
      bEditOk = False
'    Case AR_BUTTONNAME
'      bEditOk = False
    Case AR_FIELDTYPE
      cboxValue.AddItem (AR_INTEGER)
      cboxValue.Text = AR_INTEGER
      cboxValue.tag = 0
      cboxValue.AddItem (AR_REAL)
      cboxValue.AddItem (AR_CHAR)
      cboxValue.AddItem (AR_DIARY)
      cboxValue.AddItem (AR_SELECTION)
      cboxValue.AddItem (AR_DATE)
      cboxValue.AddItem (AR_FIXEDDECIMAL)
      cboxValue.AddItem (AR_ATTACHMENT)
      cboxValue.AddItem (AR_TRIM)
      cboxValue.AddItem (AR_CONTROL)
      cboxValue.AddItem (AR_TABLE)
      cboxValue.AddItem (AR_COLUMN)
      cboxValue.AddItem (AR_PAGE)
      cboxValue.AddItem (AR_PAGEHOLDER)
      bEditOk = False
  End Select

End Sub


Private Sub SetupFieldSearch()
Dim Column As ColumnHeader
  
  mnuListViewSortName.Caption = "Field Name"
  mnuListViewSortEnabled.Visible = False
  mnuListViewSortMask.Visible = False
  mnuListViewSortMod.Visible = False
  mnuListViewSortExecute.Visible = False
  mnuListViewSortID.Visible = True
  mnuListViewSortDataType.Visible = True

  lvListView.ColumnHeaders.Clear
  
  Set Column = lvListView.ColumnHeaders.Add(1, "Name", "Field Name", GetSetting(App.Title, "Settings", "FieldColumn1", 2400))
  Set Column = lvListView.ColumnHeaders.Add(2, "FormName", "Form Name", GetSetting(App.Title, "Settings", "FieldColumn2", 2400))
  Set Column = lvListView.ColumnHeaders.Add(3, "ID", "ID", GetSetting(App.Title, "Settings", "FieldColumn3", 2400))
  Set Column = lvListView.ColumnHeaders.Add(4, "Type", "Data Type", GetSetting(App.Title, "Settings", "FieldColumn4", 2400))
  lvListView.ColumnHeaders(1).Position = GetSetting(App.Title, "Settings", "FieldColumnPos1", 1)
  lvListView.ColumnHeaders(2).Position = GetSetting(App.Title, "Settings", "FieldColumnPos2", 2)
  lvListView.ColumnHeaders(3).Position = GetSetting(App.Title, "Settings", "FieldColumnPos3", 3)
  lvListView.ColumnHeaders(4).Position = GetSetting(App.Title, "Settings", "FieldColumnPos4", 4)

  ckboxCaseSensitive.Enabled = False

  cboxProperties.Clear

  cboxProperties.AddItem (AR_FIELDNAME)
  cboxProperties.Text = AR_FIELDNAME
  cboxProperties.AddItem (AR_FIELDID)
  cboxProperties.AddItem (AR_FIELDTYPE)

End Sub

Private Sub SaveColumnHeaders(sMode As String)

  Select Case sMode
  Case TYPE_AL
    SaveSetting App.Title, "Settings", "ALColumn1", lvListView.ColumnHeaders(1).Width
    SaveSetting App.Title, "Settings", "ALColumn2", lvListView.ColumnHeaders(2).Width
    SaveSetting App.Title, "Settings", "ALColumn3", lvListView.ColumnHeaders(3).Width
    SaveSetting App.Title, "Settings", "ALColumn4", lvListView.ColumnHeaders(4).Width
    SaveSetting App.Title, "Settings", "ALColumn5", lvListView.ColumnHeaders(5).Width
    SaveSetting App.Title, "Settings", "ALColumn6", lvListView.ColumnHeaders(6).Width
    SaveSetting App.Title, "Settings", "ALColumnPos1", lvListView.ColumnHeaders(1).Position
    SaveSetting App.Title, "Settings", "ALColumnPos2", lvListView.ColumnHeaders(2).Position
    SaveSetting App.Title, "Settings", "ALColumnPos3", lvListView.ColumnHeaders(3).Position
    SaveSetting App.Title, "Settings", "ALColumnPos4", lvListView.ColumnHeaders(4).Position
    SaveSetting App.Title, "Settings", "ALColumnPos5", lvListView.ColumnHeaders(5).Position
    SaveSetting App.Title, "Settings", "ALColumnPos6", lvListView.ColumnHeaders(6).Position
  Case TYPE_FILTER
    SaveSetting App.Title, "Settings", "FLColumn1", lvListView.ColumnHeaders(1).Width
    SaveSetting App.Title, "Settings", "FLColumn2", lvListView.ColumnHeaders(2).Width
    SaveSetting App.Title, "Settings", "FLColumn3", lvListView.ColumnHeaders(3).Width
    SaveSetting App.Title, "Settings", "FLColumn4", lvListView.ColumnHeaders(4).Width
    SaveSetting App.Title, "Settings", "FLColumn5", lvListView.ColumnHeaders(5).Width
    SaveSetting App.Title, "Settings", "FLColumn6", lvListView.ColumnHeaders(6).Width
    SaveSetting App.Title, "Settings", "FLColumnPos1", lvListView.ColumnHeaders(1).Position
    SaveSetting App.Title, "Settings", "FLColumnPos2", lvListView.ColumnHeaders(2).Position
    SaveSetting App.Title, "Settings", "FLColumnPos3", lvListView.ColumnHeaders(3).Position
    SaveSetting App.Title, "Settings", "FLColumnPos4", lvListView.ColumnHeaders(4).Position
    SaveSetting App.Title, "Settings", "FLColumnPos5", lvListView.ColumnHeaders(5).Position
    SaveSetting App.Title, "Settings", "FLColumnPos6", lvListView.ColumnHeaders(6).Position
  Case TYPE_FIELD
    SaveSetting App.Title, "Settings", "FieldColumn1", lvListView.ColumnHeaders(1).Width
    SaveSetting App.Title, "Settings", "FieldColumn2", lvListView.ColumnHeaders(2).Width
    SaveSetting App.Title, "Settings", "FieldColumn3", lvListView.ColumnHeaders(3).Width
    SaveSetting App.Title, "Settings", "FieldColumn4", lvListView.ColumnHeaders(4).Width
    SaveSetting App.Title, "Settings", "FieldColumnPos1", lvListView.ColumnHeaders(1).Position
    SaveSetting App.Title, "Settings", "FieldColumnPos2", lvListView.ColumnHeaders(2).Position
    SaveSetting App.Title, "Settings", "FieldColumnPos3", lvListView.ColumnHeaders(3).Position
    SaveSetting App.Title, "Settings", "FieldColumnPos4", lvListView.ColumnHeaders(4).Position
  End Select

End Sub


Private Sub SetupALSearch()
Dim Column As ColumnHeader

  mnuListViewSortName.Caption = "Active Link Name"
  mnuListViewSortEnabled.Visible = True
  mnuListViewSortMask.Visible = True
  mnuListViewSortMod.Visible = True
  mnuListViewSortExecute.Visible = True
  mnuListViewSortID.Visible = False
  mnuListViewSortDataType.Visible = False

  lvListView.ColumnHeaders.Clear
  Set Column = lvListView.ColumnHeaders.Add(1, "Name", "Active Link Name", GetSetting(App.Title, "Settings", "ALColumn1", 2400))
  Set Column = lvListView.ColumnHeaders.Add(2, "FormName", "Form Name", GetSetting(App.Title, "Settings", "ALColumn2", 2400))
  Set Column = lvListView.ColumnHeaders.Add(3, "ModTime", "Modification Time", GetSetting(App.Title, "Settings", "ALColumn3", 2400))
  Set Column = lvListView.ColumnHeaders.Add(4, "ExecMask", "Execution Mask", GetSetting(App.Title, "Settings", "ALColumn4", 2400))
  Set Column = lvListView.ColumnHeaders.Add(5, "ExecOrder", "Execution Order", GetSetting(App.Title, "Settings", "ALColumn5", 2400))
  Set Column = lvListView.ColumnHeaders.Add(6, "Enabled", "Enabled", GetSetting(App.Title, "Settings", "ALColumn6", 2400))
  lvListView.ColumnHeaders(1).Position = GetSetting(App.Title, "Settings", "ALColumnPos1", 1)
  lvListView.ColumnHeaders(2).Position = GetSetting(App.Title, "Settings", "ALColumnPos2", 2)
  lvListView.ColumnHeaders(3).Position = GetSetting(App.Title, "Settings", "ALColumnPos3", 3)
  lvListView.ColumnHeaders(4).Position = GetSetting(App.Title, "Settings", "ALColumnPos4", 4)
  lvListView.ColumnHeaders(5).Position = GetSetting(App.Title, "Settings", "ALColumnPos5", 5)
  lvListView.ColumnHeaders(6).Position = GetSetting(App.Title, "Settings", "ALColumnPos6", 6)


  ckboxCaseSensitive.Enabled = True
  
  cboxProperties.Clear

  cboxProperties.AddItem (AR_ALNAME)
  cboxProperties.Text = AR_ALNAME
  cboxProperties.AddItem (AR_MODTIME)
  cboxProperties.AddItem (AR_ENABLEDDISABLED)
  cboxProperties.AddItem (AR_EXECUTIONORDER)
  cboxProperties.AddItem (AR_EXECUTEON)
  cboxProperties.AddItem (AR_FOCUSFIELDNAME)
  cboxProperties.AddItem (AR_BUTTONNAME)
  cboxProperties.AddItem (AR_RUNIFTEXT)

End Sub


Private Sub cboxValue_Click()

  Select Case cboxValue.Text
    Case AR_ENABLED
      cboxValue.tag = 1
    Case AR_DISABLED
      cboxValue.tag = 0
  End Select

End Sub


Private Sub cboxConditions_Click()

  Select Case cboxProperties.Text
    Case AR_ALNAME, AR_FILTERNAME
    
      Select Case cboxConditions.Text
        Case AR_BEGINSWITH
          cboxConditions.tag = 1
        Case AR_CONTAINS
          cboxConditions.tag = 2
        Case AR_ENDSWITH
          cboxConditions.tag = 3
      End Select
      
    Case AR_MODTIME
    
      Select Case cboxConditions.Text
        Case AR_GREATERTHAN
          cboxConditions.tag = 2
        Case AR_LESSTHAN
          cboxConditions.tag = 5
        'the following two are not implemented yet.
        'case ar_daterange
        'case ar_exactdate
      End Select
      
    Case AR_ENABLEDDISABLED
      'There are no conditions for Enabled/Disabled
      cboxConditions.tag = Empty
    Case AR_EXECUTIONORDER
    
      Select Case cboxConditions.Text
        Case AR_EQUAL
          cboxConditions.tag = 1
        Case AR_GREATERTHAN
          cboxConditions.tag = 2
        Case AR_GREATERTHANOREQUAL
          cboxConditions.tag = 3
        Case AR_LESSTHAN
          cboxConditions.tag = 4
        Case AR_LESSTHANOREQUAL
          cboxConditions.tag = 5
      End Select
      
    Case AR_EXECUTEON
      'There are no conditions for Execute On
      cboxConditions.tag = Empty
    Case AR_FOCUSFIELDNAME
      'There are no conditions for Focus Field Name
      cboxConditions.tag = Empty
    Case AR_BUTTONNAME
      'There are no conditions for Button Name
      cboxConditions.tag = Empty
    Case AR_RUNIFTEXT
    
      Select Case cboxConditions.Text
        'Not implemented yet
        'case ar_beginswith
        '  cboxConditions.Tag = 1
        Case AR_CONTAINS
          cboxConditions.tag = 2
        'Not implemented yet
        'case ar_endswith
        '  cboxConditions.Tag = 3
      End Select
      
  End Select

End Sub



Private Sub PrefillCondition(sMode As String)

  bEditOk = True
  cboxValue.Clear
  cboxConditions.Clear
  cboxConditions.Enabled = True
  
  cmdOr.Visible = False
  cmdAnd.Caption = "&Add"
  cmdAnd.Enabled = True

  Select Case sMode
    Case AR_ALNAME
      cboxConditions.AddItem (AR_BEGINSWITH)
      cboxConditions.Text = AR_BEGINSWITH
      cboxConditions.tag = 1
      cboxConditions.AddItem (AR_CONTAINS)
      cboxConditions.AddItem (AR_ENDSWITH)
    Case AR_FILTERNAME
      cboxConditions.AddItem (AR_BEGINSWITH)
      cboxConditions.Text = AR_BEGINSWITH
      cboxConditions.tag = 1
      cboxConditions.AddItem (AR_CONTAINS)
      cboxConditions.AddItem (AR_ENDSWITH)
    Case AR_MODTIME
      cboxConditions.AddItem (AR_GREATERTHAN)
      cboxConditions.Text = AR_GREATERTHAN
      cboxConditions.tag = 2
      cboxConditions.AddItem (AR_LESSTHAN)
      'the following two are not implemented yet.
      'cboxConditions.AddItem (ar_daterange)
      'cboxConditions.AddItem (ar_exactdate)
    Case AR_ENABLEDDISABLED
      cboxConditions.AddItem (AR_EQUAL)
      cboxConditions.Text = AR_EQUAL
      cboxConditions.tag = 0
      cboxConditions.Enabled = False
      'PrefillValue (sMode)
    Case AR_EXECUTIONORDER
      cboxConditions.AddItem (AR_EQUAL)
      cboxConditions.Text = AR_EQUAL
      cboxConditions.tag = 1
      cboxConditions.AddItem (AR_GREATERTHAN)
      cboxConditions.AddItem (AR_GREATERTHANOREQUAL)
      cboxConditions.AddItem (AR_LESSTHAN)
      cboxConditions.AddItem (AR_LESSTHANOREQUAL)
    Case AR_EXECUTEON
      cboxConditions.AddItem (AR_EQUAL)
      cboxConditions.Text = AR_EQUAL
      cboxConditions.tag = 0
      cboxConditions.Enabled = False
'      PrefillValue (sMode)
      
      cmdOr.Visible = True
      cmdAnd.Caption = "&And"
'      If ARQuery.ExecuteOnANDORValue = "AND" Then
'        cmdAnd.Enabled = True
'        cmdOr.Enabled = False
'      ElseIf ARQuery.ExecuteOnANDORValue = "OR" Then
'        cmdAnd.Enabled = False
'        cmdOr.Enabled = True
'      Else
'        cmdAnd.Enabled = True
'        cmdOr.Enabled = True
'      End If
    Case AR_FOCUSFIELDNAME
      cboxConditions.AddItem (AR_EQUAL)
      cboxConditions.Text = AR_EQUAL
      cboxConditions.tag = 0
      cboxConditions.Enabled = False
    Case AR_BUTTONNAME
      cboxConditions.AddItem (AR_EQUAL)
      cboxConditions.Text = AR_EQUAL
      cboxConditions.tag = 0
      cboxConditions.Enabled = False
    Case AR_RUNIFTEXT
      'Not implemented yet
      'cboxConditions.AddItem (ar_beginswith)
      cboxConditions.AddItem (AR_CONTAINS)
      cboxConditions.Text = AR_CONTAINS
      cboxConditions.tag = 2
      cboxConditions.Enabled = False
      'Not implemented yet
      'cboxConditions.AddItem (ar_endswith)
    Case AR_FIELDNAME
      cboxConditions.AddItem (AR_BEGINSWITH)
      cboxConditions.Text = AR_BEGINSWITH
      cboxConditions.tag = 1
      cboxConditions.AddItem (AR_CONTAINS)
      cboxConditions.AddItem (AR_ENDSWITH)
      cmdOr.Visible = True
      cmdAnd.Caption = "&And"
    Case AR_FIELDID
      cboxConditions.AddItem (AR_EQUAL)
      cboxConditions.Text = AR_EQUAL
      cboxConditions.tag = 1
      cboxConditions.AddItem (AR_GREATERTHAN)
      cboxConditions.AddItem (AR_GREATERTHANOREQUAL)
      cboxConditions.AddItem (AR_LESSTHAN)
      cboxConditions.AddItem (AR_LESSTHANOREQUAL)
      cmdOr.Visible = True
      cmdAnd.Caption = "&And"
    Case AR_FIELDTYPE
      cboxConditions.AddItem (AR_EQUAL)
      cboxConditions.Text = AR_EQUAL
      cboxConditions.tag = 0
      cboxConditions.Enabled = False
      cmdOr.Visible = True
      cmdAnd.Caption = "&And"
  End Select
  
  PrefillValue (sMode)

End Sub


Private Sub cboxValue_DropDown()
Dim lCount As Long
Dim IsChecked As Boolean
Dim IsSelected As Boolean
Dim i As Long

  Select Case cboxProperties.Text
  Case AR_MODTIME
  'If cboxProperties.Text = AR_MODTIME Then
    'Popup Date/Time box
    frmDatePicker.Top = Me.Top + picSearchOptions.Top + cboxValue.Top + (cboxValue.Height * 3)
    frmDatePicker.Left = Me.Left + picSearchOptions.Left + cboxValue.Left + 40
    bEditOk = True
    frmDatePicker.Show vbModal
    cboxValue.Enabled = False
    cboxValue.Enabled = True
    bEditOk = False
  Case AR_FOCUSFIELDNAME, AR_BUTTONNAME, AR_RUNIFTEXT
  'ElseIf cboxProperties.Text = AR_FOCUSFIELDNAME Or cboxProperties.Text = AR_BUTTONNAME Then
    'Popup Field Name/ID box
    frmIDPicker.Top = Me.Top + picSearchOptions.Top + cboxValue.Top + (cboxValue.Height * 3)
    frmIDPicker.Left = Me.Left + picSearchOptions.Left + cboxValue.Left + 40
    
    lCount = tvTreeView.Nodes.Count
    
    frmProgress.Caption = "Updating Cache.."
    frmProgress.lblStatus = "Checking Form: "
    frmProgress.pbProgress.Max = lCount
    frmProgress.pbProgress2.Visible = True
    frmProgress.Show
    frmProgress.Refresh
  
    For i = 1 To lCount
      IsChecked = False
      
      frmProgress.pbProgress.Value = i
      
      IsChecked = tvTreeView.Nodes(i).Checked
      
      If IsChecked = True Then
      
        If tvTreeView.Nodes(i).Parent = ARFORMS Then
          frmProgress.lblStatus = "Checking Form: " & tvTreeView.Nodes(i).Text
          IsSelected = tvTreeView.Nodes(i).Selected
          'Check to see if Form modifaction time is new
          tvTreeView.Nodes(i).tag = UpdateFormCacheByModTime(tvTreeView.Nodes(i).Text, tvTreeView.Nodes(i).tag)
          Call frmIDPicker.AddForm(tvTreeView.Nodes(i).Text, IsSelected, cboxProperties.Text)
        End If
        
      End If
      
    Next i
    
    frmProgress.Hide
    
    If cboxProperties.Text = AR_RUNIFTEXT Then
      frmIDPicker.AddKeywords
    End If
    
    bEditOk = True
    frmIDPicker.Show vbModal
    cboxValue.Enabled = False
    cboxValue.Enabled = True
    
    If cboxProperties.Text = AR_RUNIFTEXT Then
      bEditOk = True
    Else
      bEditOk = False
    End If
    
  End Select
  'End If

End Sub


Public Function UpdateFormCacheByModTime(sFormName As String, lID As Long) As Long
Dim lCacheModTime As Long
Dim lARModTime As Long
Dim lFormID As Long
Dim lFieldCount As Long
Dim lFieldID As Long
Dim i As Long
Dim sName As String
Dim sType As String

  lFormID = lID

  lCacheModTime = modDatabase.GetFormCacheModTime(lFormID)
  lARModTime = ARCom.GetFormModTime(sFormName)
  
  If lARModTime > lCacheModTime Then
    If lCacheModTime > 0 Then
      modDatabase.DeleteFormFromCache (lFormID)
    End If
    lFieldCount = ARCom.SetFieldIDPairs(sFormName)
    lFormID = modDatabase.AddFormToCache(sFormName, lARModTime, lCurrentServerID)
    
    frmProgress.lblStatus.Caption = "Updating cache for Form: " & sFormName
    frmProgress.pbProgress2.Max = lFieldCount
    frmProgress.Refresh
    
    For i = 1 To lFieldCount
      frmProgress.pbProgress2.Value = i
      sName = ARCom.GetFieldName(i)
      lFieldID = ARCom.GetFieldID(i)
      sType = ARCom.GetFieldDataTypeString(i)
      Call modDatabase.AddFieldToCache(sName, lFieldID, sType, lFormID)
    Next i
      
  End If
  
  UpdateFormCacheByModTime = lFormID
  
End Function


Private Sub ModifyActiveLink()
Dim i As Long
Dim lresult As Long
Dim sMSG As String
Dim bErrorOccured As Boolean
Dim liListItem As ListItem

  bErrorOccured = False
  
  'Reset the field object
  ARCom.ResetActiveLinkModify
  
  'Populate formname/fieldID list
  For i = 1 To lvListView.ListItems.Count
    If lvListView.ListItems(i).Selected = True Then
      Set liListItem = lvListView.ListItems(i)
      lresult = ARCom.AddActiveLinkToModify(liListItem.Text)
      
      If lresult = arOK Then
      Else
        sMSG = "An error has occured setting the Active Link to modify." & vbCrLf
        sMSG = sMSG & "'" & ARCom.GetErrorText & "'"
        lresult = MsgBox(sMSG, vbOKOnly + vbCritical, "Error")
        bErrorOccured = True
      End If
    End If
  Next i
  
  If Len(tbName.Text) > 0 Then
    If Not ARCom.SetActiveLinkName(tbName.Text) = arOK Then
      sMSG = "An error has occured setting the Active Link Name property." & vbCrLf
      sMSG = sMSG & "'" & ARCom.GetErrorText & "'"
      lresult = MsgBox(sMSG, vbOKOnly + vbCritical, "Error")
      bErrorOccured = True
    End If
  End If
  
  If Len(tbExecutionOrder.Text) > 0 Then
    If Not ARCom.SetActiveLinkExecutionOrder(Val(tbExecutionOrder.Text)) = arOK Then
      sMSG = "An error has occured setting the Active Link Execution Order property." & vbCrLf
      sMSG = sMSG & "'" & ARCom.GetErrorText & "'"
      lresult = MsgBox(sMSG, vbOKOnly + vbCritical, "Error")
      bErrorOccured = True
    End If
  End If
  
  lresult = arOK
  
  Select Case cboxEnabled.Text
  Case "Yes"
    lresult = ARCom.SetActiveLinkEnabled(True)
  Case "No"
    lresult = ARCom.SetActiveLinkEnabled(False)
  End Select
  
  If Not lresult = arOK Then
    sMSG = "An error has occured setting the Active Link Enabled property." & vbCrLf
    sMSG = sMSG & "'" & ARCom.GetErrorText & "'"
    lresult = MsgBox(sMSG, vbOKOnly + vbCritical, "Error")
    bErrorOccured = True
  End If
  
  If Len(tbChangeHistory.Text) > 0 Then
    If Not ARCom.SetActiveLinkChangeHistory(tbChangeHistory.Text) = arOK Then
      sMSG = "An error has occured setting the Active Link Change History text." & vbCrLf
      sMSG = sMSG & "'" & ARCom.GetErrorText & "'"
      lresult = MsgBox(sMSG, vbOKOnly + vbCritical, "Error")
      bErrorOccured = True
    End If
  End If
  
  If Not cboxHelpText.Text = "<No Change>" Then
  
    If Len(tbHelpText.Text) > 0 Then
      If Not ARCom.SetActiveLinkHelpText(tbHelpText.Text) = arOK Then
        sMSG = "An error has occured setting the Active Link Help Text." & vbCrLf
        sMSG = sMSG & "'" & ARCom.GetErrorText & "'"
        lresult = MsgBox(sMSG, vbOKOnly + vbCritical, "Error")
        bErrorOccured = True
      End If
    End If
    
    ARCom.SetActiveLinkHelpTextAction (cboxHelpText.Text)

  End If

  If Not cboxPermissionType.Text = "<No Action>" Then
  
    If lboxAccess.ListCount > 0 Then
    
      For i = 0 To lboxAccess.ListCount - 1
      
        If Not ARCom.SetActiveLinkPermissionGroup(lboxAccess.List(i)) = arOK Then
          sMSG = "An error has occured setting an Active Link Permission Group (" & lboxAccess.List(i) & ")." & vbCrLf
          sMSG = sMSG & "'" & ARCom.GetErrorText & "'"
          lresult = MsgBox(sMSG, vbOKOnly + vbCritical, "Error")
          bErrorOccured = True
        End If
        
      Next i
      
    End If
  
    If Not ARCom.SetActiveLinkPermissionType(cboxPermissionType.Text) = arOK Then
      sMSG = "An error has occured setting the Active Link Permission Type." & vbCrLf
      sMSG = sMSG & "'" & ARCom.GetErrorText & "'"
      lresult = MsgBox(sMSG, vbOKOnly + vbCritical, "Error")
      bErrorOccured = True
    End If
  
  End If
  
  If bErrorOccured = False Then
    lresult = ARCom.ExecuteActiveLinkModification
    If lresult = arOK Then
      sMSG = "Modification of the selected Active Link(s) completed successfully."
    Else
      sMSG = "There was an error modifying the selected Active Link(s)."
      sMSG = sMSG & vbCrLf & ARCom.GetErrorText
      i = MsgBox(sMSG, vbOKOnly + vbCritical, "Critical Error:")
      sMSG = "Error modifying the selected Active Link(s)"
    End If
  Else
    sMSG = "Modification of the selected Active Link(s) incomplete, no Active Links were changed."
  End If
    
  SetStatusMessage (sMSG)

End Sub



Private Sub ModifyFilter()
Dim i As Long
Dim lresult As Long
Dim sMSG As String
Dim bErrorOccured As Boolean
Dim liListItem As ListItem

  bErrorOccured = False
  
  'Reset the field object
  ARCom.ResetFilterModify
  
  'Populate formname/fieldID list
  For i = 1 To lvListView.ListItems.Count
    If lvListView.ListItems(i).Selected = True Then
      Set liListItem = lvListView.ListItems(i)
      lresult = ARCom.AddFilterToModify(liListItem.Text)
      If lresult = arOK Then
      Else
        sMSG = "An error has occured setting the Filter to modify." & vbCrLf
        sMSG = sMSG & "'" & ARCom.GetErrorText & "'"
        lresult = MsgBox(sMSG, vbOKOnly + vbCritical, "Error")
        bErrorOccured = True
      End If
    End If
  Next i
  
  If Len(tbName.Text) > 0 Then
    If Not ARCom.SetFilterNameModify(tbName.Text) = arOK Then
      sMSG = "An error has occured setting the Filter Name property." & vbCrLf
      sMSG = sMSG & "'" & ARCom.GetErrorText & "'"
      lresult = MsgBox(sMSG, vbOKOnly + vbCritical, "Error")
      bErrorOccured = True
    End If
  End If
  
  If Len(tbExecutionOrder.Text) > 0 Then
    If Not ARCom.SetFilterExecutionOrder(Val(tbExecutionOrder.Text)) = arOK Then
      sMSG = "An error has occured setting the Filter Execution Order property." & vbCrLf
      sMSG = sMSG & "'" & ARCom.GetErrorText & "'"
      lresult = MsgBox(sMSG, vbOKOnly + vbCritical, "Error")
      bErrorOccured = True
    End If
  End If
  
  lresult = arOK
  
  Select Case cboxEnabled.Text
  Case "Yes"
    lresult = ARCom.SetFilterEnabled(True)
  Case "No"
    lresult = ARCom.SetFilterEnabled(False)
  End Select
  
  If Not lresult = arOK Then
    sMSG = "An error has occured setting the Filter Enabled property." & vbCrLf
    sMSG = sMSG & "'" & ARCom.GetErrorText & "'"
    lresult = MsgBox(sMSG, vbOKOnly + vbCritical, "Error")
    bErrorOccured = True
  End If
  
  If Len(tbChangeHistory.Text) > 0 Then
    If Not ARCom.SetFilterChangeHistory(tbChangeHistory.Text) = arOK Then
      sMSG = "An error has occured setting the Filter Change History text." & vbCrLf
      sMSG = sMSG & "'" & ARCom.GetErrorText & "'"
      lresult = MsgBox(sMSG, vbOKOnly + vbCritical, "Error")
      bErrorOccured = True
    End If
  End If
  
  If Not cboxHelpText.Text = "<No Change>" Then
  
    If Len(tbHelpText.Text) > 0 Then
      If Not ARCom.SetFilterHelpText(tbHelpText.Text) = arOK Then
        sMSG = "An error has occured setting the Filter Help Text." & vbCrLf
        sMSG = sMSG & "'" & ARCom.GetErrorText & "'"
        lresult = MsgBox(sMSG, vbOKOnly + vbCritical, "Error")
        bErrorOccured = True
      End If
    End If
    
    ARCom.SetFilterHelpTextAction (cboxHelpText.Text)
  
  End If
  
  If bErrorOccured = False Then
    lresult = ARCom.ExecuteFilterModification
    If lresult = arOK Then
      sMSG = "Modification of the selected Filter(s) completed successfully."
    Else
      sMSG = "There was an error modifying the selected Filter(s)."
      sMSG = sMSG & vbCrLf & ARCom.GetErrorText
      i = MsgBox(sMSG, vbOKOnly + vbCritical, "Critical Error:")
    End If
  Else
    sMSG = "Modification of the selected Filter(s) incomplete, no Filters were changed."
  End If
    
  SetStatusMessage (sMSG)

End Sub



'Main execution point for modifying Field properties
Private Sub ModifyField()
Dim i As Long
Dim lresult As Long
Dim sMSG As String
Dim bErrorOccured As Boolean
Dim liListItem As ListItem
Dim sDisplayOnly As String
Dim lDisplayCount As Long

  bErrorOccured = False
  
  'Reset the field object
  ARCom.ResetFieldObject
  
  Select Case cboxEntryMode.Text
  Case "Optional"
    ARCom.SetFieldOptionalProp (PropertyOptional)
  Case "Required"
    ARCom.SetFieldOptionalProp (PropertyRequired)
  End Select
  
'  If optEntryOptional.Value = True Then
'    ARCom.SetFieldOptionalProp (PropertyOptional)
'  Else
'
'    If optEntryRequired.Value = True Then
'      ARCom.SetFieldOptionalProp (PropertyRequired)
'    End If
'
'  End If

  Select Case cboxSubmit.Text
  Case "Yes"
    ARCom.SetFieldCreateMode (CreateModeOPEN)
  Case "No"
    ARCom.SetFieldCreateMode (CreateModePROTECTED)
  End Select
  
'  If optSubmitNo = True Then
'    ARCom.SetFieldCreateMode (CreateModePROTECTED)
'  Else
'    If optSubmitYes = True Then
'      ARCom.SetFieldCreateMode (CreateModeOPEN)
'    End If
'  End If

  If Len(tbFieldLabel.Text) > 0 Then
    lresult = ARCom.SetFieldLabel(tbFieldLabel.Text)
    If lresult <> arOK Then
      sMSG = "An error has occured:" & vbCrLf
      sMSG = sMSG & ARCom.GetErrorText
      lresult = MsgBox(sMSG, vbOKOnly + vbCritical, "Error")
      bErrorOccured = True
    End If
  End If

  If Len(tbDatabaseName.Text) > 0 Then
    lresult = ARCom.SetFieldDBName(tbDatabaseName.Text)
    If lresult <> arOK Then
      sMSG = "An error has occured:" & vbCrLf
      sMSG = sMSG & ARCom.GetErrorText
      lresult = MsgBox(sMSG, vbOKOnly + vbCritical, "Error")
      bErrorOccured = True
    End If
  End If


  If Not cboxHelpText.Text = "<No Change>" Then
  
    If Len(tbHelpText.Text) > 0 Then
      lresult = ARCom.SetFieldHelpText(tbHelpText.Text)
      If lresult <> arOK Then
        sMSG = "An error has occured:" & vbCrLf
        sMSG = sMSG & ARCom.GetErrorText
        lresult = MsgBox(sMSG, vbOKOnly + vbCritical, "Error")
        bErrorOccured = True
      End If
    End If

    ARCom.SetFieldHelpTextAction (cboxHelpText.Text)
    
  End If

  If Len(tbChangeHistory.Text) > 0 Then
    lresult = ARCom.SetFieldChangeHistory(tbChangeHistory.Text)
    If lresult <> arOK Then
      sMSG = "An error has occured:" & vbCrLf
      sMSG = sMSG & ARCom.GetErrorText
      lresult = MsgBox(sMSG, vbOKOnly + vbCritical, "Error")
      bErrorOccured = True
    End If
  End If

  'Populate formname/fieldID list
  For i = 1 To lvListView.ListItems.Count
    If lvListView.ListItems(i).Selected = True Then
      Set liListItem = lvListView.ListItems(i)
      lresult = ARCom.SetFormFieldPair(liListItem.SubItems(1), liListItem.SubItems(2))
      'lResult = ARCom.SetFormFieldPair(lvListView.ListItems(i).SubItems("FormName").Text, _
        lvListView.ListItems(i).SubItems("ID"))
      If lresult = arOK Then
      Else
        If lresult = 30001 Then '30001 = Display Only field
          sDisplayOnly = sDisplayOnly & "  " & lvListView.ListItems(i).Text & vbCrLf
          lDisplayCount = lDisplayCount + 1
        Else
          sMSG = "An error has occured setting the Form/FieldID pair." & vbCrLf
          sMSG = sMSG & "'" & ARCom.GetErrorText & "'"
          lresult = MsgBox(sMSG, vbOKOnly + vbCritical, "Error")
          bErrorOccured = True
        End If
      End If
    End If
  Next i
  
  If lDisplayCount > 0 Then
    sMSG = "The Entry mode property for the following fields "
    sMSG = sMSG & "could not be updated because they are "
    sMSG = sMSG & "Display Only or Core fields." & vbCrLf
    sMSG = sMSG & sDisplayOnly
    On Error Resume Next
    i = MsgBox(sMSG, vbOKOnly + vbInformation, "Please Note:")
    On Error GoTo 0
  End If
  
  If bErrorOccured = False Then
    sMSG = "Modification of the selected Field(s) completed successfully."
  Else
    sMSG = "Modification of the selected Field(s) incomplete, errors occured."
  End If
    
  SetStatusMessage (sMSG)

End Sub


Private Sub Modify()
Dim i As Long
Dim sMSG As String
Dim iCount As Long

  If bInDemoMode = False Then

    iCount = 0

    For i = 1 To lvListView.ListItems.Count
      If lvListView.ListItems(i).Selected = True Then
        iCount = iCount + 1
      End If
    Next i

    If iCount > 0 Then
      MousePointer = vbHourglass

      Select Case tbMainToolbar.Buttons("ObjectType").tag
        Case TYPE_AL
          sMSG = "You are about to change " & Trim(Str(iCount)) & " Active Link(s)."
          sMSG = sMSG & vbCrLf & "Are you sure you wish to commit the changes?"
          i = MsgBox(sMSG, vbYesNo + vbQuestion, "Confirm Modification")
          If i = vbYes Then
            ModifyActiveLink
          End If

        Case TYPE_FILTER
          sMSG = "You are about to change " & Trim(Str(iCount)) & " Filter(s)."
          sMSG = sMSG & vbCrLf & "Are you sure you wish to commit the changes?"
          i = MsgBox(sMSG, vbYesNo + vbQuestion, "Confirm Modification")
          If i = vbYes Then
            ModifyFilter
          End If

        Case TYPE_FIELD
          sMSG = "You are about to change " & Trim(Str(iCount)) & " Field(s)."
          sMSG = sMSG & vbCrLf & "Are you sure you wish to commit the changes?"
          i = MsgBox(sMSG, vbYesNo + vbQuestion, "Confirm Modification")
          If i = vbYes Then
            ModifyField
          End If

      End Select
      MousePointer = vbDefault
    Else
      sMSG = "There are no objects selected to modify." & vbCrLf
      sMSG = sMSG & "Please select the objects you wish to modify and try again."
      i = MsgBox(sMSG, vbOKOnly + vbInformation, "Non Critical Error..")
    End If

  Else
    
    sMSG = "Modifications are not allowed in the demo version of AR Explorer ™."
    sMSG = sMSG & vbCrLf & "Please contact sales@simpsons.arexperts.com to purchase a user license."
    sMSG = sMSG & vbCrLf
    sMSG = sMSG & vbCrLf & "Thank you,"
    sMSG = sMSG & vbCrLf
    sMSG = sMSG & vbCrLf & "AR Accelerators, Inc."
    i = MsgBox(sMSG, vbOKOnly + vbInformation, "Demo")
  
  End If
  
End Sub


Private Sub Search()

  Select Case ARQuery.SearchType
    Case TYPE_AL
      SearchActiveLink
    Case TYPE_FILTER
      SearchFilter
'    Case TYPE_FORM
'      SearchForm
    Case TYPE_FIELD
      SearchField
  End Select
  
  tbMainToolbar.Refresh

End Sub


'UPDATE:  Fixed.. need to add execute on code tho.
'Alrighty.. this needs to be changed AGAIN to support the new search format.
'I think this is a very good format to use and am surpised I didn't think of it
'sooner.  I WILL be moving this code over to the Collection side (or collection of
'collections rather).
Private Sub SearchFilter()
Dim i As Integer
Dim lNumFormsToSearch As Long
Dim iSearchCount As Integer
Dim sSearchType As String
Dim iReturnValue As Long
Dim bOkToExecuteSearch As Boolean
Dim sMSG As String
Dim lTime As Long
Dim QueryItem As clsQueryItem

  SetStatusMessage ("Searching...")

  iSearchCount = ARQuery.Count 'Me.lvSearchArguments.ListItems.Count
  
  Call ARCom.ResetFLSearchList

  lNumFormsToSearch = FlagFLFormsToSearch()

  bOkToExecuteSearch = True
  
  For i = 1 To iSearchCount
    Set QueryItem = ARQuery.Item(i)
    
    sSearchType = QueryItem.SearchType
    Select Case sSearchType
      Case AR_FILTERNAME
        If (Len(QueryItem.SearchValueString) > 0) Then
          iReturnValue = ARCom.SetFLNameSearchParam(QueryItem.SearchValueString, QueryItem.SearchParam)
        End If
      Case AR_MODTIME
        lTime = ARCom.ConvertDateToJulian(QueryItem.SearchValueString)
        iReturnValue = ARCom.SetFLModifiedSearchParam(lTime, QueryItem.SearchParam)
      Case AR_ENABLEDDISABLED
        iReturnValue = ARCom.SetFLEnabledDisabledParam(QueryItem.SearchValueNum)
      Case AR_EXECUTIONORDER
        iReturnValue = ARCom.SetFLExecutionOrderParam(Format(QueryItem.SearchValueString), QueryItem.SearchParam)
      Case AR_EXECUTEON
        If QueryItem.SearchCondition = "AND" Then
          iReturnValue = ARCom.SetFLExecuteOnParam(ARQuery.ExecuteOnValue, 1)
        Else
          iReturnValue = ARCom.SetFLExecuteOnParam(ARQuery.ExecuteOnValue, 2)
        End If
      Case AR_RUNIFTEXT
        iReturnValue = ARCom.SetFLRunIfSearchParam(QueryItem.SearchValueString, QueryItem.SearchParam)
    End Select
      
    If Not iReturnValue = arOK Then
      bOkToExecuteSearch = False
      lvSearchArguments.ListItems(i).Bold = True
      sMSG = ARCom.GetErrorText
      iReturnValue = MsgBox(sMSG, vbOKOnly + vbInformation)
    End If
    
  Next i

  If bOkToExecuteSearch Then
    Me.MousePointer = vbHourglass
    Call ARCom.SetFLCaseSensitive(ARQuery.CaseSensitive)
    iReturnValue = ARCom.SearchFilter()
    
    If iReturnValue = arOK Then
      lvListView.ListItems.Clear
      ShowFoundFilters
    Else
      sMSG = "An error occurred.  "
      sMSG = sMSG & ARCom.GetErrorText
      i = MsgBox(sMSG, vbOKOnly + vbCritical)
      ARCom.ErrorResolved (True)
    End If
    
    MousePointer = vbDefault
    
  Else
    sMSG = "There is a problem with an item in the Query List. "
    sMSG = sMSG & "Please correct and resubmit"
    i = MsgBox(sMSG & vbCrLf & ARCom.GetErrorText, vbOKOnly + vbInformation, "Query Error")
  End If
    
End Sub


''This should be a bit more simple then SearchActiveLink() or SearchFilter() as
''we are only searching the Cache and not using Remedy at all.
''But.. honestly.. who knows =)
'Private Sub SearchField()
'Dim rsResult As Recordset
'Dim IsChecked As Boolean
'Dim sSQL As String
'Dim iCount As Long
'Dim i As Long
'Dim liItem As ListItem
'Dim lFormCount As Long
'Dim sProperty As String
'Dim sCondition As String
'Dim sValue As String
'Dim sMSG As String
'Dim sPreviousAndOr As String
'Dim sPreviousProperty As String
'Dim sID As String
'
'  SetStatusMessage ("Searching...")
'
'  MousePointer = vbHourglass
'
'  'first we need to build the SQL string
'  sSQL = "SELECT * FROM [FieldProperties] "
'  sSQL = sSQL & "WHERE ("
'
'  iCount = tvTreeView.Nodes.Count
'  lFormCount = 0
'
'  frmProgress.Caption = "Updating Cache.."
'  frmProgress.lblStatus = "Checking Form: "
'  frmProgress.pbProgress.Max = iCount
'  frmProgress.pbProgress2.Visible = True
'  frmProgress.Show
'  frmProgress.Refresh
'
'
'  'here we need to get all the form ID's for the checked form(s)
'  For i = 1 To iCount
'    IsChecked = False
'
'    frmProgress.pbProgress.Value = i
'
'    IsChecked = tvTreeView.Nodes(i).Checked
'
'    If IsChecked = True Then
'
'      If tvTreeView.Nodes(i).Parent = ARFORMS Then
'        frmProgress.lblStatus = "Checking Form: " & tvTreeView.Nodes(i).Text
'        tvTreeView.Nodes(i).tag = UpdateFormCacheByModTime(tvTreeView.Nodes(i).Text, tvTreeView.Nodes(i).tag)
'        lFormCount = lFormCount + 1
'        'Add cache id to sql statement
'       If lFormCount = 1 Then
'          sSQL = sSQL & "[ParentFormID] = " & Trim(tvTreeView.Nodes(i).tag)
'        Else
'          'sSQL = sSQL & ", " & Trim(tvTreeView.Nodes(i).tag)
'          'sSQL = sSQL & " OR " & Trim(tvTreeView.Nodes(i).tag)
'          sSQL = sSQL & " OR [ParentFormID] = " & Trim(tvTreeView.Nodes(i).tag)
'        End If
'
'      End If
'
'    End If
'
'  Next i
'
'  frmProgress.Hide
'  Me.Refresh
'
'  sSQL = sSQL & ")"
'
'  If lFormCount > 0 Then
'    'Now we build the requested results
'    'note, if no search arguments exist then just return ALL fields
'    If lvSearchArguments.ListItems.Count > 0 Then
'      sSQL = sSQL & " AND "
'
'      sPreviousAndOr = ""
'      sPreviousProperty = ""
'      sSQL = sSQL & "("
'      For i = 1 To lvSearchArguments.ListItems.Count
'        Set liItem = lvSearchArguments.ListItems(i)
'
'        If liItem.Text = sPreviousProperty Then
'          If Len(sPreviousAndOr) > 0 Then
'            sSQL = sSQL & " " & sPreviousAndOr & " "
'          End If
'        Else
'          If Len(sPreviousProperty) > 0 Then
'            sSQL = sSQL & ") " & sPreviousAndOr & " ("
'          End If
'        End If
'
'        Select Case liItem.Text
'          Case AR_FIELDNAME
'            sSQL = sSQL & "[Name] "
'          Case AR_FIELDID
'            sSQL = sSQL & "[ARID] "
'          Case AR_FIELDTYPE
'            sSQL = sSQL & "[TYPE] "
'        End Select
'
'        sPreviousProperty = liItem.Text
'
'        Select Case liItem.SubItems(1)
'          Case AR_BEGINSWITH
'            If Me.ckboxCaseSensitive.Value = vbChecked Then
'              sCondition = "="
'            Else
'              sCondition = "LIKE"
'            End If
'            sValue = "'" & Trim(liItem.SubItems(2)) & "*'"
'          Case AR_CONTAINS
'            If Me.ckboxCaseSensitive.Value = vbChecked Then
'              sCondition = "="
'            Else
'              sCondition = "LIKE"
'            End If
'            sValue = "'*" & Trim(liItem.SubItems(2)) & "*'"
'          Case AR_ENDSWITH
'            If Me.ckboxCaseSensitive.Value = vbChecked Then
'              sCondition = "="
'            Else
'              sCondition = "LIKE"
'            End If
'            sValue = "'*" & Trim(liItem.SubItems(2)) & "'"
'          Case AR_GREATERTHAN
'            sCondition = ">"
'          Case AR_DATERANGE
'            sCondition = " "
'            sValue = "#" & Trim(liItem.SubItems(2)) & "#"
'          Case AR_EXACTDATE
'            sCondition = "="
'            sValue = "#" & Trim(liItem.SubItems(2)) & "#"
'          Case AR_EQUAL
'            If Me.cboxProperties.Text = AR_FIELDTYPE Then
'              sValue = "'" & Trim(liItem.SubItems(2)) & "'"
'            End If
'            sCondition = "="
'          Case AR_GREATERTHANOREQUAL
'            sCondition = ">="
'          Case AR_LESSTHAN
'            sCondition = "<"
'          Case AR_LESSTHANOREQUAL
'            sCondition = "<="
'        End Select
'        sSQL = sSQL & sCondition & " "
'        If Len(sValue) = 0 Then
'          sSQL = sSQL & Trim(liItem.SubItems(2))
'        Else
'          sSQL = sSQL & sValue
'        End If
'        sPreviousAndOr = liItem.SubItems(3)
'        'sQuery = sQuery & liItem.SubItems(3) & " "
'        sValue = ""
'        sCondition = ""
'      Next i
'      sSQL = sSQL & ")"
'
'    End If
'
'    sSQL = sSQL & ";"
'
''    frmGetText.tbText = sSQL
''    frmGetText.Show vbModal
'
'    'Then we call the database function to search it
'    Set rsResult = modDatabase.ExecuteCacheSQL(sSQL)
'
''    i = MsgBox("Executed SQL statement", vbOKOnly)
'
'    'Then we display results
'    On Error Resume Next
'    rsResult.MoveLast
'    rsResult.MoveFirst
'    On Error GoTo 0
'    lvListView.ListItems.Clear
'
''    i = MsgBox("Building Results", vbOKOnly)
'
'    For i = 1 To rsResult.RecordCount
'
''      iCount = MsgBox("Displaying result #" & Trim(Str(i)), vbOKOnly)
'      '  sExecutionOrder = String(4 - Len(sExecOrder), " ") & sExecOrder
'
'      sID = String(10 - Len(Trim(Str(rsResult(fldARID)))), " ") & Trim(Str(rsResult(fldARID)))
'
'      AddLVItemField rsResult(fldName), modDatabase.GetFormName(rsResult(fldParentFormID)), sID, rsResult(fldType)
'      rsResult.MoveNext
'
'    Next i
'
'    MousePointer = vbDefault
'
'    SetStatusMessage ("Found " & Trim(rsResult.RecordCount) & " Fields.")
'
'  Else
'    sMSG = "An error has occured: "
'    sMSG = sMSG & "No forms were selected."
'    i = MsgBox(sMSG, vbOKOnly + vbInformation, "Query Error")
'  End If
'
'  MousePointer = vbDefault
'
'End Sub
'This should be a bit more simple then SearchActiveLink() or SearchFilter() as
'we are only searching the Cache and not using Remedy at all.
'But.. honestly.. who knows =)
Private Sub SearchField()
Dim rsResult As Recordset
Dim IsChecked As Boolean
Dim sSQL As String
Dim iCount As Long
Dim i As Long
Dim j As Long
Dim liItem As ListItem
Dim lFormCount As Long
Dim sProperty As String
Dim sCondition As String
Dim sValue As String
Dim sMSG As String
Dim sPreviousAndOr As String
Dim sPreviousProperty As String
Dim sID As String
Dim lFoundFields As Long

  SetStatusMessage ("Searching...")
  
  MousePointer = vbHourglass
  
  'first we need to build the SQL string
  sSQL = "SELECT * FROM [FieldProperties] "
  sSQL = sSQL & "WHERE "
  
  iCount = tvTreeView.Nodes.Count
  lFormCount = 0
  lFoundFields = 0

  frmProgress.Caption = "Updating Cache.."
  frmProgress.lblStatus = "Checking Form: "
  frmProgress.pbProgress.Max = iCount
  frmProgress.pbProgress2.Visible = True
  frmProgress.Show
  frmProgress.Refresh


  'here we need to get all the form ID's for the checked form(s)
  For i = 1 To iCount
    IsChecked = False
    
    frmProgress.pbProgress.Value = i
    
    IsChecked = tvTreeView.Nodes(i).Checked
    
    If IsChecked = True Then
    
      If tvTreeView.Nodes(i).Parent = ARFORMS Then
        frmProgress.lblStatus = "Checking Form: " & tvTreeView.Nodes(i).Text
        tvTreeView.Nodes(i).tag = UpdateFormCacheByModTime(tvTreeView.Nodes(i).Text, tvTreeView.Nodes(i).tag)
        lFormCount = lFormCount + 1
        'Add cache id to sql statement
       'If lFormCount = 1 Then
          'sSQL = sSQL & "[ParentFormID] = " & Trim(tvTreeView.Nodes(i).tag)
        'Else
          'sSQL = sSQL & " OR [ParentFormID] = " & Trim(tvTreeView.Nodes(i).tag)
        'End If
        
      End If
      
    End If
    
  Next i
  
  frmProgress.Hide
  Me.Refresh
  
  lvListView.ListItems.Clear
  
  If lFormCount > 0 Then
  
    For j = 1 To iCount
    
      IsChecked = False
      IsChecked = tvTreeView.Nodes(j).Checked
      
      If IsChecked = True Then
        If tvTreeView.Nodes(j).Parent = ARFORMS Then
          sSQL = "SELECT * FROM [FieldProperties] "
          sSQL = sSQL & "WHERE "

          sSQL = sSQL & "[ParentFormID] = " & Trim(tvTreeView.Nodes(j).tag)
          'Now we build the requested results
          'note, if no search arguments exist then just return ALL fields
          If lvSearchArguments.ListItems.Count > 0 Then
            sSQL = sSQL & " AND "
      
            sPreviousAndOr = ""
            sPreviousProperty = ""
            sSQL = sSQL & "("
            For i = 1 To lvSearchArguments.ListItems.Count
              Set liItem = lvSearchArguments.ListItems(i)
              
              If liItem.Text = sPreviousProperty Then
                If Len(sPreviousAndOr) > 0 Then
                  sSQL = sSQL & " " & sPreviousAndOr & " "
                End If
              Else
                If Len(sPreviousProperty) > 0 Then
                  sSQL = sSQL & ") " & sPreviousAndOr & " ("
                End If
              End If
              
              Select Case liItem.Text
                Case AR_FIELDNAME
                  sSQL = sSQL & "[Name] "
                Case AR_FIELDID
                  sSQL = sSQL & "[ARID] "
                Case AR_FIELDTYPE
                  sSQL = sSQL & "[TYPE] "
              End Select
                
              sPreviousProperty = liItem.Text
              
              Select Case liItem.SubItems(1)
                Case AR_BEGINSWITH
                  If ckboxCaseSensitive.Value = vbChecked Then
                    sCondition = "="
                  Else
                    sCondition = "LIKE"
                  End If
                  sValue = "'" & Trim(liItem.SubItems(2)) & "*'"
                Case AR_CONTAINS
                  If ckboxCaseSensitive.Value = vbChecked Then
                    sCondition = "="
                  Else
                    sCondition = "LIKE"
                  End If
                  sValue = "'*" & Trim(liItem.SubItems(2)) & "*'"
                Case AR_ENDSWITH
                  If ckboxCaseSensitive.Value = vbChecked Then
                    sCondition = "="
                  Else
                    sCondition = "LIKE"
                  End If
                  sValue = "'*" & Trim(liItem.SubItems(2)) & "'"
                Case AR_GREATERTHAN
                  sCondition = ">"
                Case AR_DATERANGE
                  sCondition = " "
                  sValue = "#" & Trim(liItem.SubItems(2)) & "#"
                Case AR_EXACTDATE
                  sCondition = "="
                  sValue = "#" & Trim(liItem.SubItems(2)) & "#"
                Case AR_EQUAL
                  If liItem.Text = AR_FIELDTYPE Then
                  'If cboxProperties.Text = AR_FIELDTYPE Then
                    sValue = "'" & Trim(liItem.SubItems(2)) & "'"
                  End If
                  sCondition = "="
                Case AR_GREATERTHANOREQUAL
                  sCondition = ">="
                Case AR_LESSTHAN
                  sCondition = "<"
                Case AR_LESSTHANOREQUAL
                  sCondition = "<="
              End Select
              sSQL = sSQL & sCondition & " "
              If Len(sValue) = 0 Then
                sSQL = sSQL & Trim(liItem.SubItems(2))
              Else
                sSQL = sSQL & sValue
              End If
              sPreviousAndOr = liItem.SubItems(3)
              'sQuery = sQuery & liItem.SubItems(3) & " "
              sValue = ""
              sCondition = ""
            Next i
            sSQL = sSQL & ")"
                       
          End If
    
          sSQL = sSQL & ";"
              
      '    frmGetText.tbText = sSQL
      '    frmGetText.Show vbModal
              
          'Then we call the database function to search it
          Set rsResult = modDatabase.ExecuteCacheSQL(sSQL)
    
      '    i = MsgBox("Executed SQL statement", vbOKOnly)
          
          'Then we display results
          On Error Resume Next
          rsResult.MoveLast
          rsResult.MoveFirst
          On Error GoTo 0
          'lvListView.ListItems.Clear
          lFoundFields = lFoundFields + rsResult.RecordCount
          
      '    i = MsgBox("Building Results", vbOKOnly)
          
          For i = 1 To rsResult.RecordCount
          
      '      iCount = MsgBox("Displaying result #" & Trim(Str(i)), vbOKOnly)
            '  sExecutionOrder = String(4 - Len(sExecOrder), " ") & sExecOrder
      
            sID = String(10 - Len(Trim(Str(rsResult(fldARID)))), " ") & Trim(Str(rsResult(fldARID)))
            
            AddLVItemField rsResult(fldName), modDatabase.GetFormName(rsResult(fldParentFormID)), sID, rsResult(fldType)
            rsResult.MoveNext
          
          Next i
        End If
      End If
    Next j
    
    MousePointer = vbDefault
    
    SetStatusMessage ("Found " & Trim(lFoundFields) & " Fields.")
    
  Else
    sMSG = "An error has occured: "
    sMSG = sMSG & "No forms were selected."
    i = MsgBox(sMSG, vbOKOnly + vbInformation, "Query Error")
  End If
  
  MousePointer = vbDefault

End Sub



'UPDATE:  Fixed.. need to add execute on code tho.
'Alrighty.. this needs to be changed AGAIN to support the new search format.
'I think this is a very good format to use and am surpised I didn't think of it
'sooner.  I WILL be moving this code over to the Collection side (or collection of
'collections rather).
Private Sub SearchActiveLink()
'Const PROPERTY_VALUE = 4
'Const CONDITION_VALUE = 5
'Const SEARCH_VALUE = 6
'Const SEARCH_STRING = 2
Dim i As Integer
Dim lNumFormsToSearch As Long
Dim iSearchCount As Integer
Dim sSearchType As String
Dim iReturnValue As Long
Dim bOkToExecuteSearch As Boolean
Dim sMSG As String
Dim lTime As Long
Dim QueryItem As clsQueryItem

  SetStatusMessage ("Searching...")
  
  iSearchCount = ARQuery.Count 'Me.lvSearchArguments.ListItems.Count
  
  'If iSearchCount > 0 Then
  
    Call ARCom.ResetALSearchList 'Count how many search parameters were specified
  
    lNumFormsToSearch = FlagALFormsToSearch()
  
    bOkToExecuteSearch = True
    
    'Pass all search params over to AREServer.dll
    For i = 1 To iSearchCount
      Set QueryItem = ARQuery.Item(i)
      
      sSearchType = QueryItem.SearchType
      Select Case sSearchType
        Case AR_ALNAME
          If (Len(QueryItem.SearchValueString) > 0) Then
            iReturnValue = ARCom.SetALNameSearchParam(QueryItem.SearchValueString, QueryItem.SearchParam)
          End If
        Case AR_MODTIME
          lTime = ARCom.ConvertDateToJulian(QueryItem.SearchValueString)
          iReturnValue = ARCom.SetALModifiedSearchParam(lTime, QueryItem.SearchParam)
        Case AR_ENABLEDDISABLED
          iReturnValue = ARCom.SetALEnabledDisabledParam(QueryItem.SearchValueNum)
        Case AR_EXECUTIONORDER
          iReturnValue = ARCom.SetALExecutionOrderParam(Val(QueryItem.SearchValueString), QueryItem.SearchParam)
        Case AR_EXECUTEON
          If QueryItem.SearchCondition = "AND" Then
            iReturnValue = ARCom.SetALExecuteOnParam(ARQuery.ExecuteOnValue, 1)
          Else
            iReturnValue = ARCom.SetALExecuteOnParam(ARQuery.ExecuteOnValue, 2)
          End If
        Case AR_FOCUSFIELDNAME
          iReturnValue = ARCom.SetALFocusFieldIDParam(QueryItem.SearchValueNum)
        Case AR_BUTTONNAME
          iReturnValue = ARCom.SetALButtonIDParam(QueryItem.SearchValueNum)
        Case AR_RUNIFTEXT
          iReturnValue = ARCom.SetALRunIfSearchParam(QueryItem.SearchValueString, QueryItem.SearchParam)
      End Select
        
      If Not iReturnValue = arOK Then
        bOkToExecuteSearch = False
        lvSearchArguments.ListItems(i).Bold = True
      End If
      
    Next i
  
    If bOkToExecuteSearch Then
      MousePointer = vbHourglass
      Call ARCom.SetALCaseSensitive(ARQuery.CaseSensitive)
      iReturnValue = ARCom.SearchActiveLink()
      
      If iReturnValue = arOK Then
        lvListView.ListItems.Clear
        ShowFoundActiveLinks
        MousePointer = vbDefault
      Else
        MousePointer = vbDefault
        sMSG = "An error occurred.  "
        sMSG = sMSG & ARCom.GetErrorText
        i = MsgBox(sMSG, vbOKOnly + vbCritical)
        ARCom.ErrorResolved (True)
      End If
      
    Else
      sMSG = "There is a problem with an item in the Query List. "
      sMSG = sMSG & "Please correct and resubmit"
      i = MsgBox(sMSG & vbCrLf & ARCom.GetErrorText, vbOKOnly + vbInformation, "Query Error")
    End If
    
  'End If

End Sub


Private Sub ShowFoundFilters()
Dim lCount As Long
Dim i As Integer
Dim sFilterName As String

  lCount = ARCom.FoundNumberOfFilters
  
  lvListView.ListItems.Clear
  
  If lCount Then
    'loop and load each active link name into the List Box along with the form names
    i = 0
    
    Do While i < lCount
      
      sFilterName = ARCom.GetFoundFilterName
      Call ARCom.GetFilterProperty
      
      AddLVItemFL sFilterName, ARCom.sFormName, ARCom.sExecutionOrder, ARCom.sEnabled, ARCom.sExecuteMask, ARCom.sModifiedTime
      
  
      ARCom.GotoNextFoundFLName
      i = i + 1

    Loop
    
    SetStatusMessage ("Found " & Trim(Str(lCount)) & " Filters")
    
    'Me.sbMainStatusBar.Panels(1).Text = "Found " & lCount & " Filters."
  Else
    SetStatusMessage ("No matches were found.")
    'Me.sbMainStatusBar.Panels(1).Text = "No matches were found."
  End If

End Sub


Private Sub ShowFoundActiveLinks()
Dim lCount As Long
Dim i As Integer
Dim sALName As String

  lCount = ARCom.FoundNumberOfActiveLinks
  
  lvListView.ListItems.Clear
  
  If lCount Then
    'loop and load each active link name into the List Box along with the form names
    i = 0
    
    Do While i < lCount
      
      sALName = ARCom.GetFoundALName
      Call ARCom.GetALProperty
      
      AddLVItemAL sALName, ARCom.sFormName, ARCom.sExecutionOrder, ARCom.sEnabled, ARCom.sExecuteMask, ARCom.sModifiedTime
      
      ARCom.GotoNextFoundALName
      i = i + 1

    Loop
    
    SetStatusMessage ("Found " & Trim(Str(lCount)) & " Active Links.")
  Else
    SetStatusMessage ("No matches were found.")
  End If

End Sub


Private Function FlagFLFormsToSearch() As Long
Dim iCount As Integer
Dim i As Integer
Dim IsChecked As Boolean
Dim ndNode As Node

  'First loop and get all of the checked form names.
  iCount = tvTreeView.Nodes.Count

  For i = 1 To iCount
    IsChecked = False
    
    IsChecked = tvTreeView.Nodes(i).Checked
    
    If IsChecked = True Then
    
      If tvTreeView.Nodes(i).Parent = ARFORMS Then
        'Push the form name across to the COM dll.
        ARCom.AddFLFormSearchItem (tvTreeView.Nodes(i).Text)
      End If
      
    End If
    
  Next i
  
  FlagFLFormsToSearch = 1 'ARCom.GetFLNumFormsToSearch

End Function


Private Function FlagALFormsToSearch() As Long
Dim iCount As Integer
Dim i As Integer
Dim IsChecked As Boolean
Dim ndNode As Node

  'First loop and get all of the checked form names.
  iCount = tvTreeView.Nodes.Count

  For i = 1 To iCount
    IsChecked = False
    
    IsChecked = tvTreeView.Nodes(i).Checked
    
    If IsChecked = True Then
    
      If tvTreeView.Nodes(i).Parent = ARFORMS Then
        'Push the form name across to the COM dll.
        ARCom.AddALFormSearchItem (tvTreeView.Nodes(i).Text)
      End If
      
    End If
    
  Next i
  
  FlagALFormsToSearch = ARCom.GetALNumFormsToSearch

End Function

Private Sub tvTreeView_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

  If Button = 2 Then
    PopupMenu mnuTreeView
  End If

End Sub

Private Sub tvTreeView_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim i As Long
Dim nTempNode As MSComctlLib.Node
Dim bChecked As Boolean

  If Node.Text = ARFORMS Then
    bChecked = Node.Checked
    If bChecked = True Then
      lCheckedNodeCount = Node.Children + 1 '+1 for the parent node check
    Else
      lCheckedNodeCount = 0
    End If
    Set nTempNode = Node.Child
    For i = 1 To Node.Children
      nTempNode.Checked = bChecked
      Set nTempNode = nTempNode.Next
    Next i
  Else
    If Node.Checked = True Then
      lCheckedNodeCount = lCheckedNodeCount + 1
    Else
      lCheckedNodeCount = lCheckedNodeCount - 1
    End If
    Node.Selected = True
  End If
    
End Sub


Public Sub CreateCache(sServerName As String)
Dim lCount As Long
Dim lFieldCount As Long
Dim i As Long
Dim j As Long
Dim ndNode As Node
Dim sName As String
Dim lModTime As Long
Dim lID As Long
Dim sType As String
Dim lFormID As Long


  lCount = tvTreeView.Nodes.Count
  
  lCurrentServerID = modDatabase.AddServerToCache(sServerName)

  frmProgress.pbProgress.Max = lCount
  frmProgress.pbProgress2.Visible = True
  
  For i = 2 To lCount
    
    If Me.tvTreeView.Nodes(i).Parent = ARFORMS Then
      frmProgress.lblStatus.Caption = "Getting fields for form: " & tvTreeView.Nodes(i).Text
      frmProgress.pbProgress.Value = i
      frmProgress.Refresh
      lFieldCount = ARCom.SetFieldIDPairs(tvTreeView.Nodes(i).Text)
      lModTime = ARCom.GetFormModTime(tvTreeView.Nodes(i).Text)
      lFormID = modDatabase.AddFormToCache(tvTreeView.Nodes(i).Text, lModTime, lCurrentServerID)
      frmProgress.pbProgress2.Max = lFieldCount
      
      tvTreeView.Nodes(i).tag = lFormID

      For j = 1 To lFieldCount
        frmProgress.pbProgress2.Value = j
        frmProgress.Refresh
        sName = ARCom.GetFieldName(j)
        lID = ARCom.GetFieldID(j)
        sType = ARCom.GetFieldDataTypeString(j)
        Call modDatabase.AddFieldToCache(sName, lID, sType, lFormID)
      Next j
      
    End If
    
  Next i

End Sub


Private Sub PrintResults()
Dim i As Long
Dim sText As String
Dim sFieldText
Dim liListItem As ListItem
Dim sOldFontName As String
Dim sOldFontSize As Single
Dim j As Integer
Dim sMSG As String


  On Error GoTo ErrorHandler
  
  sOldFontName = Printer.Font.Name
  sOldFontSize = Printer.Font.Size
  
  Printer.Font.Name = "Courier New"
  Printer.Font.Size = 8.16
  
  If ARQuery.Count > 0 Then
    If ARQuery.SaveName = sEmptyString Then
      sText = "Search results for the current query:"
    Else
      sText = "Search results for query: " & ARQuery.SaveName
    End If
    
    Printer.Print sText
  
    For i = 1 To Me.lvSearchArguments.ListItems.Count
      Set liListItem = lvSearchArguments.ListItems(i)
      sText = liListItem.Text & vbTab & liListItem.SubItems(2) & vbTab & liListItem.SubItems(3)
      Printer.Print sText
    Next i
    
  End If
  
  sText = String(128, "-")
  Printer.Print sText

  For i = 1 To lvListView.ListItems.Count
    Set liListItem = lvListView.ListItems(i)
    sFieldText = liListItem.Text
    sText = sFieldText & String(FieldNameSize - Len(sFieldText), " ")
    sText = sText & String(SpacerSize, " ")
    
    sFieldText = liListItem.SubItems(1)
    sText = sText & sFieldText & String(FormNameSize - Len(sFieldText), " ")
    sText = sText & String(SpacerSize, " ")
    
    sFieldText = liListItem.SubItems(2)
    sText = sText & sFieldText & String(ModificationTimeSize - Len(sFieldText), " ")
    sText = sText & String(SpacerSize, " ")
    
    sFieldText = liListItem.SubItems(3)
    If Len(sFieldText) > 19 Then
      sText = sText & Left(sFieldText, ExecuteMaskSize)
    Else
      sText = sText & sFieldText & String(ExecuteMaskSize - Len(sFieldText), " ")
    End If
    sText = sText & String(SpacerSize, " ")
    
'    sFieldText = liListItem.SubItems(3)
'    sText = sText & sFieldText & String(ExecuteMaskSize - Len(sFieldText), " ")
'    sText = sText & String(SpacerSize, " ")
    
    sFieldText = liListItem.SubItems(4)
    sText = sText & sFieldText & String(ExecutionOrderSize - Len(sFieldText), " ")
    sText = sText & String(SpacerSize, " ")
    
    sFieldText = liListItem.SubItems(5)
    sText = sText & sFieldText & String(EnabledSize - Len(sFieldText), " ")
    
    
    'sText = liListItem.Text & liListItem.SubItems(4) & vbTab & liListItem.SubItems(5)
    Printer.Print sText
  Next i
  
  Printer.EndDoc
  
  Printer.Font.Name = sOldFontName
  Printer.Font.Size = sOldFontSize
  
  On Error GoTo 0
  
  Exit Sub
  
ErrorHandler:
  If Err.Number = 482 Then
    sMSG = Err.Description
    i = MsgBox(sMSG, vbOKOnly + vbCritical, "Printer")
    Exit Sub
  ElseIf Err.Number = 484 Then
    sMSG = "There is no printer available."
    i = MsgBox(sMSG, vbOKOnly + vbCritical, "Printer")
    Exit Sub
  Else
    sMSG = Err.Description
    i = MsgBox(sMSG, vbOKOnly + vbCritical, "Printer")
    Exit Sub
  End If

End Sub


Private Sub ExportResults(Optional sOveridePath As String)
Dim i As Long
Dim sFileName As String
Dim sText As String
Dim sFieldText As String
Dim iOutFile As Integer
Dim sPath As String
Dim liListItem As ListItem


On Error GoTo ErrHandler

  'used for printing
  If Len(sOveridePath) > 0 Then
    sFileName = sOveridePath
  Else
    sPath = GetSetting(App.Title, "Settings", "LastPath", App.Path)
    With dlgCommonDialog
      .Filter = "Text (*.txt)|*.txt|All Files (*.*)|*.*"
      .FilterIndex = 1
      .DefaultExt = ".txt"
      .DialogTitle = "Save Query Results"
      .CancelError = True
      .InitDir = sPath
      .ShowSave
      sFileName = .FileName
    End With
  End If
  
  On Error GoTo 0
  
'  If Len(Dir(sFileName)) > 0 Then
'    Kill (sFileName)
'  End If
  
  iOutFile = FreeFile
  Open sFileName For Output As iOutFile
  
  If ARQuery.Count > 0 Then
    If ARQuery.SaveName = sEmptyString Then
      sText = "Search results for the current query:"
    Else
      sText = "Search results for query: " & ARQuery.SaveName
    End If
    
    Print #iOutFile, sText
  
    For i = 1 To Me.lvSearchArguments.ListItems.Count
      Set liListItem = lvSearchArguments.ListItems(i)
      sText = liListItem.Text & vbTab & liListItem.SubItems(2) & vbTab & liListItem.SubItems(3)
      Print #iOutFile, sText
    Next i
    
  End If
  
  sText = String(120, "-")
  Print #iOutFile, sText

  For i = 1 To lvListView.ListItems.Count
    Set liListItem = lvListView.ListItems(i)
    sFieldText = liListItem.Text
    sText = sFieldText & String(FieldNameSize - Len(sFieldText), " ")
    sText = sText & String(SpacerSize, " ")
    
    sFieldText = liListItem.SubItems(1)
    sText = sText & sFieldText & String(FormNameSize - Len(sFieldText), " ")
    sText = sText & String(SpacerSize, " ")
    
    sFieldText = liListItem.SubItems(2)
    sText = sText & sFieldText & String(ModificationTimeSize - Len(sFieldText), " ")
    sText = sText & String(SpacerSize, " ")
    
    sFieldText = liListItem.SubItems(3)
    If Len(sFieldText) > 19 Then
      sText = sText & Left(sFieldText, ExecuteMaskSize)
    Else
      sText = sText & sFieldText & String(ExecuteMaskSize - Len(sFieldText), " ")
    End If
    sText = sText & String(SpacerSize, " ")
    
    sFieldText = liListItem.SubItems(4)
    sText = sText & sFieldText & String(ExecutionOrderSize - Len(sFieldText), " ")
    sText = sText & String(SpacerSize, " ")
    
    sFieldText = liListItem.SubItems(5)
    sText = sText & sFieldText & String(EnabledSize - Len(sFieldText), " ")
    
    'sText = liListItem.Text & liListItem.SubItems(4) & vbTab & liListItem.SubItems(5)
    Print #iOutFile, sText
  Next i
  
  Close iOutFile

  SaveSetting App.Title, "Settings", "LastPath", sPath
  
ErrHandler:
' User pressed Cancel button.
   Exit Sub
End Sub


Public Sub ResetStatusMessage()
  SetStatusMessage ("Ready")
End Sub


Public Function SetStatusMessage(sMessage As String)
  sbMainStatusBar.Panels(1).Text = sMessage
End Function

Public Function FillActionList(sObjectName As String, sObjectType)
    Dim liItem As ListItem 'var to hold AL Name as list view item
    Dim sActionType As String 'temp variable to hold Action Type returned from ARS Server
    
    'For each type, get the Actions from ARS Server and store in list
    Select Case sObjectType
    Case TYPE_AL
            'Set the Active Link Name to get Actions for
            If ARCom.ActiveLinkName(sObjectName) <> arOK Then
                'Do error processing here
            End If
            'Get the If Action List
            If ARCom.GetIfActionList() <> arOK Then
                'Error processing
            End If
            
            If ARCom.GetALActionCount > 0 Then
                'Insert "If Action" text before the if actions.
                Set liItem = lvListViewActions.ListItems.Add() 'Add a List View item and store the reference to it locally
                liItem.Text = "If Actions..." 'Insert spaces to indent action type
                
                'For each action in the list, populate the Action List View window
                'with the IF actions.
                Dim i As Integer
                i = 0
                Do While i < ARCom.GetALActionCount
                    Set liItem = lvListViewActions.ListItems.Add() 'Add a List View item and store the reference to it locally
                    sActionType = ARCom.GetALActionType() 'Store the AL name in the list item
                    liItem.Text = "     " 'Insert spaces to indent action type
                    liItem.Text = liItem.Text + sActionType 'Put the Action Type in the list view
                    ARCom.GotoNextALAction
                    i = i + 1
                Loop
            End If
            
            'Get the Else Action List
            If ARCom.GetElseActionList() <> arOK Then
                'Error processing
            End If
            If ARCom.GetALElseActionCount > 0 Then
                'Insert "Else Action" text before the if actions.
                Set liItem = lvListViewActions.ListItems.Add() 'Add a List View item and store the reference to it locally
                liItem.Text = "Else Actions..." 'Insert spaces to indent action type
                
                'For each action in the list, populate the Action List View window
                'with the IF actions.
                i = 0
                Do While i < ARCom.GetALElseActionCount
                    Set liItem = lvListViewActions.ListItems.Add() 'Add a List View item and store the reference to it locally
                    sActionType = ARCom.GetALElseActionType() 'Store the AL name in the list item
                    liItem.Text = "     " 'Insert spaces to indent action type
                    liItem.Text = liItem.Text + sActionType 'Put the Action Type in the list view
                    ARCom.GotoNextALElseAction
                    i = i + 1
                Loop
            End If

    Case TYPE_FILTER
            'Add a List View item and store the reference to it locally
            Set liItem = lvListViewActions.ListItems.Add()
            'Store the AL/FL name in the list item, created in the step above
            liItem.Text = sActionType
    Case TYPE_FIELD
    
    End Select

End Function
