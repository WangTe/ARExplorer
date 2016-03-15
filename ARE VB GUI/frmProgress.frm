VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgress 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5280
   Icon            =   "frmProgress.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar pbProgress2 
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   1440
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar pbProgress 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1020
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label2 
      Caption         =   "Status"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   1215
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   360
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   "Progress.."
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

