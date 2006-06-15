VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmProgress 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Progress"
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmProgress.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblInformation 
      Alignment       =   2  'Center
      Caption         =   "lblInformation"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4455
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Cancelled As Boolean

Private Sub cmdCancel_Click()
    Cancelled = True
End Sub
