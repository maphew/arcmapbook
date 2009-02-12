VERSION 5.00
Begin VB.Form frmVisibleLayers 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Layers Visible in this Map Page"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   4455
      Begin VB.ComboBox cboVisibilityGroups 
         Height          =   315
         Left            =   2520
         TabIndex        =   6
         Text            =   "cboVisibilityGroups"
         Top             =   120
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Participate in visibility group"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   3120
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   2085
      ItemData        =   "VisibleLayers.frx":0000
      Left            =   120
      List            =   "VisibleLayers.frx":0016
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   960
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Select layers that will be visible in this map page."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmVisibleLayers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

