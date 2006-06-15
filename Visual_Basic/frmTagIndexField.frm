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

' Copyright 1995-2004 ESRI

' All rights reserved under the copyright laws of the United States.

' You may freely redistribute and use this sample code, with or without modification.

' Disclaimer: THE SAMPLE CODE IS PROVIDED "AS IS" AND ANY EXPRESS OR IMPLIED 
' WARRANTIES, INCLUDING THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS 
' FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL ESRI OR 
' CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, 
' OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF 
' SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS 
' INTERRUPTION) SUSTAINED BY YOU OR A THIRD PARTY, HOWEVER CAUSED AND ON ANY 
' THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT ARISING IN ANY 
' WAY OUT OF THE USE OF THIS SAMPLE CODE, EVEN IF ADVISED OF THE POSSIBILITY OF 
' SUCH DAMAGE.

' For additional information contact: Environmental Systems Research Institute, Inc.

' Attn: Contracts Dept.

' 380 New York Street

' Redlands, California, U.S.A. 92373 

' Email: contracts@esri.com

Option Explicit

Public m_bCancel As Boolean    'Cancel flag

Private Sub cmdCancel_Click()
6:   Me.Hide
End Sub

Private Sub cmdOK_Click()
10:   m_bCancel = False
11:   Me.Hide
End Sub

Private Sub Form_Load()
15:   m_bCancel = True
End Sub

Public Sub InitializeList(pFields As IFields)
On Error GoTo ErrHand:
  Dim lLoop As Long, pField As IField
  
22:   lstFields.Clear
23:   For lLoop = 0 To pFields.FieldCount - 1
24:     Set pField = pFields.Field(lLoop)
25:     If pField.Type <> esriFieldTypeBlob And pField.Type <> esriFieldTypeGeometry Then
26:       lstFields.AddItem pField.Name & " - " & pField.AliasName
27:     End If
28:   Next lLoop
29:   If lstFields.ListCount > 0 Then lstFields.ListIndex = 0

  Exit Sub
ErrHand:
33:   MsgBox "InitializeList - " & Erl & " - " & Err.Description
End Sub
