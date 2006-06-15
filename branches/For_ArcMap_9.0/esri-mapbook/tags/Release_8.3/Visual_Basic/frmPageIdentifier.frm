VERSION 5.00
Begin VB.Form frmPageIdentifier 
   Caption         =   "Page Identifier"
   ClientHeight    =   2325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   3165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   345
      Left            =   1050
      TabIndex        =   4
      Top             =   1950
      Width           =   1005
   End
   Begin VB.Frame fraIdentifier 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1785
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   3045
      Begin VB.OptionButton optIdentifier 
         Caption         =   "Global"
         Height          =   225
         Index           =   1
         Left            =   1920
         TabIndex        =   3
         Top             =   1470
         Width           =   795
      End
      Begin VB.OptionButton optIdentifier 
         Caption         =   "Local"
         Height          =   225
         Index           =   0
         Left            =   390
         TabIndex        =   2
         Top             =   1470
         Width           =   765
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1395
         Left            =   30
         Picture         =   "frmPageIdentifier.frx":0000
         ScaleHeight     =   1395
         ScaleWidth      =   2985
         TabIndex        =   1
         Top             =   30
         Width           =   2985
      End
   End
End
Attribute VB_Name = "frmPageIdentifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bCancel As Boolean

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdOk_Click()
  bCancel = False
  Me.Hide
End Sub

Private Sub Form_Load()
  optIdentifier(0).Value = True
  bCancel = True
End Sub
