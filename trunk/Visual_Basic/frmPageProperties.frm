VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmPageProperties 
   Caption         =   "Page Properties"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtUpdate 
      Height          =   315
      Left            =   4080
      TabIndex        =   5
      Top             =   3990
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   345
      Left            =   5280
      TabIndex        =   2
      Top             =   3990
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   345
      Left            =   6300
      TabIndex        =   1
      Top             =   4500
      Width           =   1125
   End
   Begin MSFlexGridLib.MSFlexGrid grdPages 
      Height          =   3795
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   6694
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
   End
   Begin VB.Label lblTile 
      Caption         =   "Tile:"
      Height          =   225
      Left            =   180
      TabIndex        =   4
      Top             =   4050
      Width           =   2625
   End
   Begin VB.Label lblType 
      Alignment       =   1  'Right Justify
      Caption         =   "Scale:"
      Height          =   225
      Left            =   2910
      TabIndex        =   3
      Top             =   4050
      Width           =   1035
   End
End
Attribute VB_Name = "frmPageProperties"
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

Public m_pApp As IApplication
Private bLoadFlag As Boolean

Private Sub cmdOK_Click()
19:   Unload Me
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo ErrHand:
  Dim pMapBook As IDSMapBook, dLastOutput As Date
  Dim pSeries As IDSMapSeries, lLoop As Long, pPage As IDSMapPage
26:   Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
29:   Set pSeries = pMapBook.ContentItem(0)
30:   Set pPage = pSeries.Page(grdPages.Row - 1)
  
  Select Case grdPages.Col
  Case 1   'Name
34:     grdPages.Text = txtUpdate.Text
35:     pPage.PageName = txtUpdate.Text
  Case 2   'Scale
37:     grdPages.Text = txtUpdate.Text
38:     pPage.PageScale = CDbl(txtUpdate.Text)
  Case 3   'Rotation
40:     grdPages.Text = txtUpdate.Text
41:     pPage.PageRotation = CDbl(txtUpdate.Text)
  Case 4   'Last Output
43:     dLastOutput = Format(txtUpdate.Text, "mm/dd/yyyy")
44:     grdPages.Text = dLastOutput
45:     pPage.LastOutputted = dLastOutput
46:   End Select

  Exit Sub
  
ErrHand:
51:   MsgBox "cmdUpdate_Click - " & Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo ErrHand:
  Dim pMapBook As IDSMapBook
  Dim pSeries As IDSMapSeries, lLoop As Long, pPage As IDSMapPage
58:   Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
61:   Set pSeries = pMapBook.ContentItem(0)
  
  'Load up the Grid
64:   bLoadFlag = True
65:   grdPages.Clear
66:   grdPages.Rows = 1
67:   grdPages.ColWidth(0) = 750
68:   grdPages.ColWidth(1) = 2500
69:   grdPages.ColWidth(2) = 1000
70:   grdPages.ColWidth(3) = 1000
71:   grdPages.ColWidth(4) = 1400
72:   grdPages.Row = 0
73:   grdPages.Col = 0
74:   grdPages.Text = "Number"
75:   grdPages.Col = 1
76:   grdPages.Text = "Tile Name"
77:   grdPages.Col = 2
78:   grdPages.Text = "Scale"
79:   grdPages.Col = 3
80:   grdPages.Text = "Rotation"
81:   grdPages.Col = 4
82:   grdPages.Text = "Last Output"
83:   grdPages.CellAlignment = 1
84:   For lLoop = 0 To pSeries.PageCount - 1
85:     Set pPage = pSeries.Page(lLoop)
86:     grdPages.AddItem lLoop + 1 & Chr(9) & pPage.PageName & Chr(9) & pPage.PageScale & _
     Chr(9) & pPage.PageRotation & Chr(9) & pPage.LastOutputted
88:   Next lLoop
89:   bLoadFlag = False
  
  'Set the Update properties
92:   lblType.Caption = ""
93:   txtUpdate.Text = ""
94:   lblTile.Caption = "Page:"
95:   cmdUpdate.Enabled = False
  
  'Make sure the wizard stays on top
98:   TopMost Me
  
  Exit Sub
  
ErrHand:
103:   MsgBox "frmPageProperties_Load - " & Err.Description
End Sub

Private Sub grdPages_EnterCell()
On Error GoTo ErrHand:
  Dim pMapBook As IDSMapBook, lCol As Long, lRow As Long
  Dim pSeries As IDSMapSeries, lLoop As Long, pPage As IDSMapPage
  
  'Exit the sub if we are loading up the grid
  If bLoadFlag Then Exit Sub
  
114:   Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
117:   Set pSeries = pMapBook.ContentItem(0)
  
119:   lCol = grdPages.Col
120:   lRow = grdPages.Row
  
122:   If lCol > 0 Then
123:     Set pPage = pSeries.Page(lRow - 1)
124:     lblTile.Caption = "Page: " & pPage.PageName
    Select Case lCol
    Case 1   'Page Name
127:       lblType.Caption = "Name:"
128:       txtUpdate.Text = pPage.PageName
    Case 2   'Scale
130:       lblType.Caption = "Scale:"
131:       txtUpdate.Text = CStr(pPage.PageScale)
    Case 3   'Rotation
133:       lblType.Caption = "Rotation:"
134:       txtUpdate.Text = CStr(pPage.PageRotation)
    Case 4   'Last Output
136:       lblType.Caption = "Last Output:"
137:       txtUpdate.Text = CStr(pPage.LastOutputted)
138:     End Select
139:     cmdUpdate.Enabled = True
140:   Else
141:     lblType.Caption = ""
142:     txtUpdate.Text = ""
143:     lblTile.Caption = "Page:"
144:     cmdUpdate.Enabled = False
145:   End If

  Exit Sub
ErrHand:
149:   MsgBox "grdPages_EnterCell - " & Err.Description
End Sub

Private Sub txtUpdate_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrHand:

155:   cmdUpdate.Enabled = False
  
  Select Case grdPages.Col
  Case 1   'Name
159:     If txtUpdate.Text <> "" Then
160:       cmdUpdate.Enabled = True
161:     End If
  Case 2, 3  'Scale and Rotation
163:     If txtUpdate.Text <> "" Then
164:       If IsNumeric(txtUpdate.Text) Then
165:         cmdUpdate.Enabled = True
166:       End If
167:     End If
  Case 4    'Last output
169:     If txtUpdate.Text <> "" Then
170:       If IsDate(txtUpdate.Text) Then
171:         cmdUpdate.Enabled = True
172:       End If
173:     End If
174:   End Select

  Exit Sub
ErrHand:
178:   MsgBox "txtUpdate_KeyUp - " & Err.Description
End Sub
