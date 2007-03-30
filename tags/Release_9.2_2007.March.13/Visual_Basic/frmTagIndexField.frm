VERSION 5.00
Begin VB.Form frmTagIndexField 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tag with Index layer field"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   315
      Left            =   2280
      TabIndex        =   4
      Top             =   2850
      Width           =   765
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   3090
      TabIndex        =   3
      Top             =   2850
      Width           =   765
   End
   Begin VB.ListBox lstFields 
      Height          =   1815
      Left            =   870
      TabIndex        =   2
      Top             =   930
      Width           =   2955
   End
   Begin VB.Label Label1 
      Caption         =   "Field:"
      Height          =   225
      Index           =   1
      Left            =   450
      TabIndex        =   1
      Top             =   960
      Width           =   435
   End
   Begin VB.Label Label1 
      Caption         =   "Choose the field you wish to use for tagging the selected text element (list shows field name and alias):"
      Height          =   585
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   3675
   End
End
Attribute VB_Name = "frmTagIndexField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Copyright 2006 ESRI
'
' All rights reserved under the copyright laws of the United States
' and applicable international laws, treaties, and conventions.
'
' You may freely redistribute and use this sample code, with or
' without modification, provided you include the original copyright
' notice and use restrictions.
'
' See use restrictions at /arcgis/developerkit/userestrictions.

Option Explicit

Public m_bCancel As Boolean    'Cancel flag

Private Sub cmdCancel_Click()
18:   Me.Hide
End Sub

Private Sub cmdOK_Click()
22:   m_bCancel = False
23:   Me.Hide
End Sub

Private Sub Form_Load()
27:   m_bCancel = True
End Sub

Public Sub InitializeList(pFields As IFields)
On Error GoTo ErrHand:
  Dim lLoop As Long, pField As IField
  
34:   lstFields.Clear
35:   For lLoop = 0 To pFields.FieldCount - 1
36:     Set pField = pFields.Field(lLoop)
37:     If pField.Type <> esriFieldTypeBlob And pField.Type <> esriFieldTypeGeometry Then
38:       lstFields.AddItem pField.Name & " - " & pField.AliasName
39:     End If
40:   Next lLoop
41:   If lstFields.ListCount > 0 Then lstFields.ListIndex = 0

  Exit Sub
ErrHand:
45:   MsgBox "InitializeList - " & Erl & " - " & Err.Description
End Sub
