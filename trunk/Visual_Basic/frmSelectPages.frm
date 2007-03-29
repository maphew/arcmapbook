VERSION 5.00
Begin VB.Form frmSelectPages 
   Caption         =   "Select/Enable Pages"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   5235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   4020
      TabIndex        =   2
      Top             =   3330
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   345
      Left            =   2880
      TabIndex        =   1
      Top             =   3330
      Width           =   1125
   End
   Begin VB.Frame fraSelection 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3165
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   5025
      Begin VB.ComboBox cmbScale 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtScale 
         Height          =   285
         Left            =   2760
         TabIndex        =   12
         Top             =   2670
         Width           =   975
      End
      Begin VB.OptionButton optSelection 
         Caption         =   "Select with scale "
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   11
         Top             =   2700
         Width           =   1605
      End
      Begin VB.TextBox txtBefore 
         Height          =   285
         Left            =   1410
         TabIndex        =   10
         Top             =   2250
         Width           =   975
      End
      Begin VB.TextBox txtAfter 
         Height          =   285
         Left            =   2940
         TabIndex        =   9
         Top             =   2250
         Width           =   975
      End
      Begin VB.OptionButton optSelection 
         Caption         =   "Select by date last printed/exported (use format 01/01/02):"
         Height          =   345
         Index           =   2
         Left            =   360
         TabIndex        =   6
         Top             =   1770
         Width           =   2925
      End
      Begin VB.OptionButton optSelection 
         Caption         =   "Unselect all"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   1260
         Width           =   1695
      End
      Begin VB.OptionButton optSelection 
         Caption         =   "Select all"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   750
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "before:"
         Height          =   195
         Index           =   1
         Left            =   870
         TabIndex        =   8
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "after:"
         Height          =   195
         Index           =   0
         Left            =   2550
         TabIndex        =   7
         Top             =   2280
         Width           =   405
      End
      Begin VB.Label Label1 
         Caption         =   $"frmSelectPages.frx":0000
         Height          =   645
         Left            =   60
         TabIndex        =   3
         Top             =   60
         Width           =   4905
      End
   End
End
Attribute VB_Name = "frmSelectPages"
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
Private m_pMapSeries As IDSMapSeries

Private Sub cmdCancel_Click()
19:   Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrHand:
  Dim lLoop As Long, pMapPage As IDSMapPage
  'Check to see if a MapSeries already exists
  
27:   If optSelection(0).value Then      'Select all
28:     SelectAllPages True
29:   ElseIf optSelection(1).value Then  'Unselect all
30:     SelectAllPages False
31:   ElseIf optSelection(2).value Then  'Select by date last printed/exported
32:     SelectByDate
33:   ElseIf optSelection(3).value Then  'Select by scale value
34:     SelectByScale
35:   End If

37:   Unload Me
  
  Exit Sub
ErrHand:
41:   MsgBox "frmSelectPages_Click - " & Err.Description
End Sub

Private Sub SelectByDate()
On Error GoTo ErrHand:
  Dim lLoop As Long, pNode As Node, dDate As Date
  Dim pPage As IDSMapPage
  
  'Select pages by date last printed/exported
50:   For lLoop = 0 To m_pMapSeries.PageCount - 1
51:     Set pPage = m_pMapSeries.Page(lLoop)
52:     Set pNode = g_pFrmMapSeries.tvwMapBook.Nodes.Item(lLoop + 3)
53:     dDate = m_pMapSeries.Page(lLoop).LastOutputted
54:     If IsDate(txtBefore.Text) And txtAfter.Text = "" Then
55:       If dDate < txtBefore.Text Or dDate = #1/1/1900# Then
56:         pPage.EnablePage = True
57:         pNode.Image = 5
58:       Else
59:         pPage.EnablePage = False
60:         pNode.Image = 6
61:       End If
62:     ElseIf IsDate(txtBefore.Text) And IsDate(txtAfter.Text) Then
63:       If dDate >= txtBefore.Text And dDate <= txtAfter.Text Then
64:         pPage.EnablePage = True
65:         pNode.Image = 5
66:       Else
67:         pPage.EnablePage = False
68:         pNode.Image = 6
69:       End If
70:     Else
71:       If dDate > txtAfter.Text Then
72:         pPage.EnablePage = True
73:         pNode.Image = 5
74:       Else
75:         pPage.EnablePage = False
76:         pNode.Image = 6
77:       End If
78:     End If
79:   Next lLoop

  Exit Sub

ErrHand:
84:   MsgBox "SelectByDate - " & Err.Description
End Sub

Private Sub SelectByScale()
On Error GoTo ErrHand:
  Dim lLoop As Long, pNode As Node, dScale As Double
  Dim pPage As IDSMapPage, sExp As String
  
  'Select pages by Scale
93:   For lLoop = 0 To m_pMapSeries.PageCount - 1
94:     Set pPage = m_pMapSeries.Page(lLoop)
95:     Set pNode = g_pFrmMapSeries.tvwMapBook.Nodes.Item(lLoop + 3)
96:     dScale = m_pMapSeries.Page(lLoop).PageScale
97:     sExp = CStr(dScale) & " " & cmbScale.Text & " " & txtScale.Text
98:     If sExp Then
99:       pPage.EnablePage = True
100:       pNode.Image = 5
101:     Else
102:       pPage.EnablePage = False
103:       pNode.Image = 6
104:     End If
105:   Next lLoop

  Exit Sub

ErrHand:
110:   MsgBox "SelectByScale - " & Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo ErrHand:
  Dim pMapBook As IDSMapBook, pOpts As IDSMapSeriesOptions
116:   Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
119:   Set m_pMapSeries = pMapBook.ContentItem(0)
120:   Set pOpts = m_pMapSeries
121:   If pOpts.ExtentType = 2 Then
122:     optSelection(3).Enabled = True
123:   Else
124:     optSelection(3).Enabled = False
125:   End If

127:   optSelection(0).value = True
  
129:   cmbScale.Clear
130:   cmbScale.AddItem "="
131:   cmbScale.AddItem "<>"
132:   cmbScale.AddItem ">"
133:   cmbScale.AddItem ">="
134:   cmbScale.AddItem "<"
135:   cmbScale.AddItem "<="
136:   cmbScale.Text = "="
  
  Exit Sub
ErrHand:
140:   MsgBox "frmSelectPages_Load - " & Err.Description
End Sub

Private Sub SelectAllPages(bValue As Boolean)
On Error GoTo ErrHand:
  Dim lLoop As Long, pNode As Node
  
  'Loop through the pages turning them on or off
148:   For lLoop = 0 To m_pMapSeries.PageCount - 1
149:     Set pNode = g_pFrmMapSeries.tvwMapBook.Nodes.Item(lLoop + 3)
150:     m_pMapSeries.Page(lLoop).EnablePage = bValue
151:     If bValue Then
152:       pNode.Image = 5
153:     Else
154:       pNode.Image = 6
155:     End If
156:   Next lLoop
  
  Exit Sub
ErrHand:
160:   MsgBox "SelectAllPages - " & Err.Description
End Sub

Private Sub optSelection_Click(Index As Integer)
  Select Case Index
  Case 0    'Select all
166:     cmdOK.Enabled = True
  Case 1    'Unselect all
168:     cmdOK.Enabled = True
  Case 2    'Select by date last printed/exported
170:     If DateCheck Then
171:       cmdOK.Enabled = True
172:     Else
173:       cmdOK.Enabled = False
174:     End If
  Case 3    'Select by scale
176:     If ScaleCheck Then
177:       cmdOK.Enabled = True
178:     Else
179:       cmdOK.Enabled = False
180:     End If
181:   End Select
End Sub

Private Sub txtAfter_KeyUp(KeyCode As Integer, Shift As Integer)
185:   If DateCheck Then
186:     cmdOK.Enabled = True
187:   Else
188:     cmdOK.Enabled = False
189:   End If
End Sub

Private Sub txtBefore_KeyUp(KeyCode As Integer, Shift As Integer)
193:   If DateCheck Then
194:     cmdOK.Enabled = True
195:   Else
196:     cmdOK.Enabled = False
197:   End If
End Sub

Private Sub txtScale_KeyUp(KeyCode As Integer, Shift As Integer)
201:   If Not IsNumeric(txtScale.Text) Then
202:     txtScale.Text = ""
203:   End If
204:   If ScaleCheck Then
205:     cmdOK.Enabled = True
206:   Else
207:     cmdOK.Enabled = False
208:   End If
End Sub

Private Function ScaleCheck() As Boolean
212:   ScaleCheck = False
213:   If txtScale.Text <> "" Then
214:     If CDbl(txtScale.Text) >= 0 Then
215:       ScaleCheck = True
216:     End If
217:   End If
End Function

Private Function DateCheck() As Boolean
221:   If IsDate(txtBefore.Text) And txtAfter.Text = "" Then
222:     DateCheck = True
223:   ElseIf IsDate(txtBefore.Text) And IsDate(txtAfter.Text) Then
224:     DateCheck = True
225:   ElseIf txtBefore.Text = "" And IsDate(txtAfter.Text) Then
226:     DateCheck = True
227:   Else
228:     DateCheck = False
229:   End If
End Function
