VERSION 5.00
Begin VB.Form frmSelectAdjMapSymbol 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Adjacent Symbol"
   ClientHeight    =   2832
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   2904
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2832
   ScaleWidth      =   2904
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   2400
      Width           =   1095
   End
   Begin VB.ComboBox cboAdjacentSymbol 
      Height          =   1680
      Left            =   0
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Text            =   "cboAdjacentSymbol"
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label lblSelectSymbol 
      Caption         =   "Select the text symbol to label pages adjacent to this page."
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2895
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   2880
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   0
      X2              =   2880
      Y1              =   2280
      Y2              =   2280
   End
End
Attribute VB_Name = "frmSelectAdjMapSymbol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_sCurrentSymbol As String
Private m_sPrevSymbol As String
Private m_pNWMapSeriesOptions As INWMapSeriesOptions
' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "frmSelectAdjMapSymbol.frm"



Public Property Let CurrentSymbol(RHS As String)
12:   m_sCurrentSymbol = RHS
13:   m_sPrevSymbol = RHS
14:   cboAdjacentSymbol.Text = RHS
15:   If RHS = "" Then
16:     If Not m_pNWMapSeriesOptions Is Nothing Then
17:       Me.cboAdjacentSymbol.Text = m_pNWMapSeriesOptions.TextSymbolDefault
18:       m_sCurrentSymbol = m_pNWMapSeriesOptions.TextSymbolDefault
19:     End If
20:   End If
End Property

Public Property Get CurrentSymbol() As String
24:   CurrentSymbol = m_sCurrentSymbol
End Property



Public Property Set NWSeriesOptions(pNWSeriesOptions As INWMapSeriesOptions)
  On Error GoTo ErrorHandler

32:   Set m_pNWMapSeriesOptions = pNWSeriesOptions
  Dim i As Integer
  Dim lSymCount As Long
  Dim vSymNames As Variant
  
37:   Me.cboAdjacentSymbol.Clear
38:   With m_pNWMapSeriesOptions
39:     lSymCount = .TextSymbolCount
40:     vSymNames = m_pNWMapSeriesOptions.TextSymbolNames
    
42:     For i = 0 To (lSymCount - 1)
43:       cboAdjacentSymbol.AddItem vSymNames(i)
44:     Next i
    
46:     If m_sCurrentSymbol = "" Then
47:       cboAdjacentSymbol.Text = m_pNWMapSeriesOptions.TextSymbolDefault
48:     Else
49:       cboAdjacentSymbol.Text = m_sCurrentSymbol
50:     End If
51:   End With

  Exit Property
ErrorHandler:
  HandleError True, "NWSeriesOptions " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get GetNWOptionSymbols() As INWMapSeriesOptions
59:   Set GetNWOptionSymbols = m_pNWMapSeriesOptions
End Property



Private Sub cboAdjacentSymbol_Change()
  Dim lFindResult As Long
  
  With cboAdjacentSymbol
    lFindResult = FindControlString(cboAdjacentSymbol, .Text)
    If lFindResult = -1 Then
      .Text = m_sCurrentSymbol
    Else
      m_sCurrentSymbol = .Text
    End If
  End With
End Sub

Private Sub cmdCancel_Click()
65:   m_sCurrentSymbol = m_sPrevSymbol
66:   Me.Hide
End Sub

Private Sub cmdOK_Click()
70:   m_sCurrentSymbol = cboAdjacentSymbol.Text
71:   Me.Hide
End Sub

Private Sub Form_Terminate()
75:   m_sCurrentSymbol = m_sPrevSymbol
End Sub

Private Sub Form_Unload(Cancel As Integer)
79:   m_sCurrentSymbol = m_sPrevSymbol
End Sub
