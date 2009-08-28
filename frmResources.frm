VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmResources 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "frmResources.frx":0000
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picCreateStripMap 
      Height          =   525
      Left            =   90
      Picture         =   "frmResources.frx":05C2
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   3
      Top             =   2100
      Width           =   1245
   End
   Begin VB.PictureBox picCreateGrid 
      Height          =   525
      Left            =   90
      Picture         =   "frmResources.frx":0B84
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   2
      Top             =   1440
      Width           =   1245
   End
   Begin VB.PictureBox picBook 
      Height          =   525
      Left            =   90
      Picture         =   "frmResources.frx":1146
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   1
      Top             =   750
      Width           =   1245
   End
   Begin VB.PictureBox picIdentifier 
      Height          =   525
      Left            =   150
      Picture         =   "frmResources.frx":1458
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   0
      Top             =   60
      Width           =   1245
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   1740
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResources.frx":1A1A
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmResources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Copyright 2008 ESRI
' 
' All rights reserved under the copyright laws of the United States
' and applicable international laws, treaties, and conventions.
' 
' You may freely redistribute and use this sample code, with or
' without modification, provided you include the original copyright
' notice and use restrictions.
' 
' See use restrictions at <your ArcGIS install location>/developerkit/userestrictions.txt.
' 




Option Explicit

