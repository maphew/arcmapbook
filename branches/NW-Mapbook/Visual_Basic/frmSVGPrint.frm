VERSION 5.00
Begin VB.Form frmSVGPrint 
   Caption         =   "Print to SVG Files"
   ClientHeight    =   2352
   ClientLeft      =   48
   ClientTop       =   360
   ClientWidth     =   5964
   LinkTopic       =   "Form1"
   ScaleHeight     =   2352
   ScaleWidth      =   5964
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBrowse 
      Height          =   315
      Left            =   5160
      Picture         =   "frmSVGPrint.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   345
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   372
      Left            =   4680
      TabIndex        =   3
      Top             =   1800
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   372
      Left            =   3360
      TabIndex        =   2
      Top             =   1800
      Width           =   972
   End
   Begin VB.TextBox txtSVGOutputPath 
      Height          =   288
      Left            =   1560
      TabIndex        =   1
      Top             =   840
      Width           =   3372
   End
   Begin VB.Label Label2 
      Caption         =   $"frmSVGPrint.frx":047A
      Height          =   612
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   5652
   End
   Begin VB.Label Label1 
      Caption         =   "Output Directory:"
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1332
   End
End
Attribute VB_Name = "frmSVGPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()

End Sub
