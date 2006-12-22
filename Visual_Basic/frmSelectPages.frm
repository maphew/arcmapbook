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
7:   Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrHand:
  Dim lLoop As Long, pMapPage As IDSMapPage
  'Check to see if a MapSeries already exists
  
15:   If optSelection(0).Value Then      'Select all
16:     SelectAllPages True
17:   ElseIf optSelection(1).Value Then  'Unselect all
18:     SelectAllPages False
19:   ElseIf optSelection(2).Value Then  'Select by date last printed/exported
20:     SelectByDate
21:   ElseIf optSelection(3).Value Then  'Select by scale value
22:     SelectByScale
23:   End If

25:   Unload Me
  
  Exit Sub
ErrHand:
29:   MsgBox "frmSelectPages_Click - " & Err.Description
End Sub

Private Sub SelectByDate()
On Error GoTo ErrHand:
  Dim lLoop As Long, pNode As Node, dDate As Date
  Dim pPage As IDSMapPage
  
  'Select pages by date last printed/exported
38:   For lLoop = 0 To m_pMapSeries.PageCount - 1
39:     Set pPage = m_pMapSeries.Page(lLoop)
40:     Set pNode = g_pFrmMapSeries.tvwMapBook.Nodes.Item(lLoop + 3)
41:     dDate = m_pMapSeries.Page(lLoop).LastOutputted
42:     If IsDate(txtBefore.Text) And txtAfter.Text = "" Then
43:       If dDate < txtBefore.Text Or dDate = #1/1/1900# Then
44:         pPage.EnablePage = True
45:         pNode.Image = 5
46:       Else
47:         pPage.EnablePage = False
48:         pNode.Image = 6
49:       End If
50:     ElseIf IsDate(txtBefore.Text) And IsDate(txtAfter.Text) Then
51:       If dDate >= txtBefore.Text And dDate <= txtAfter.Text Then
52:         pPage.EnablePage = True
53:         pNode.Image = 5
54:       Else
55:         pPage.EnablePage = False
56:         pNode.Image = 6
57:       End If
58:     Else
59:       If dDate > txtAfter.Text Then
60:         pPage.EnablePage = True
61:         pNode.Image = 5
62:       Else
63:         pPage.EnablePage = False
64:         pNode.Image = 6
65:       End If
66:     End If
67:   Next lLoop

  Exit Sub

ErrHand:
72:   MsgBox "SelectByDate - " & Err.Description
End Sub

Private Sub SelectByScale()
On Error GoTo ErrHand:
  Dim lLoop As Long, pNode As Node, dScale As Double
  Dim pPage As IDSMapPage, sExp As String
  
  'Select pages by Scale
81:   For lLoop = 0 To m_pMapSeries.PageCount - 1
82:     Set pPage = m_pMapSeries.Page(lLoop)
83:     Set pNode = g_pFrmMapSeries.tvwMapBook.Nodes.Item(lLoop + 3)
84:     dScale = m_pMapSeries.Page(lLoop).PageScale
85:     sExp = CStr(dScale) & " " & cmbScale.Text & " " & txtScale.Text
86:     If sExp Then
87:       pPage.EnablePage = True
88:       pNode.Image = 5
89:     Else
90:       pPage.EnablePage = False
91:       pNode.Image = 6
92:     End If
93:   Next lLoop

  Exit Sub

ErrHand:
98:   MsgBox "SelectByScale - " & Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo ErrHand:
  Dim pMapBook As IDSMapBook, pOpts As IDSMapSeriesOptions
104:   Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
107:   Set m_pMapSeries = pMapBook.ContentItem(0)
108:   Set pOpts = m_pMapSeries
109:   If pOpts.ExtentType = 2 Then
110:     optSelection(3).Enabled = True
111:   Else
112:     optSelection(3).Enabled = False
113:   End If

115:   optSelection(0).Value = True
  
117:   cmbScale.Clear
118:   cmbScale.AddItem "="
119:   cmbScale.AddItem "<>"
120:   cmbScale.AddItem ">"
121:   cmbScale.AddItem ">="
122:   cmbScale.AddItem "<"
123:   cmbScale.AddItem "<="
124:   cmbScale.Text = "="
  
  Exit Sub
ErrHand:
128:   MsgBox "frmSelectPages_Load - " & Err.Description
End Sub

Private Sub SelectAllPages(bValue As Boolean)
On Error GoTo ErrHand:
  Dim lLoop As Long, pNode As Node
  
  'Loop through the pages turning them on or off
136:   For lLoop = 0 To m_pMapSeries.PageCount - 1
137:     Set pNode = g_pFrmMapSeries.tvwMapBook.Nodes.Item(lLoop + 3)
138:     m_pMapSeries.Page(lLoop).EnablePage = bValue
139:     If bValue Then
140:       pNode.Image = 5
141:     Else
142:       pNode.Image = 6
143:     End If
144:   Next lLoop
  
  Exit Sub
ErrHand:
148:   MsgBox "SelectAllPages - " & Err.Description
End Sub

Private Sub optSelection_Click(Index As Integer)
  Select Case Index
  Case 0    'Select all
154:     cmdOK.Enabled = True
  Case 1    'Unselect all
156:     cmdOK.Enabled = True
  Case 2    'Select by date last printed/exported
158:     If DateCheck Then
159:       cmdOK.Enabled = True
160:     Else
161:       cmdOK.Enabled = False
162:     End If
  Case 3    'Select by scale
164:     If ScaleCheck Then
165:       cmdOK.Enabled = True
166:     Else
167:       cmdOK.Enabled = False
168:     End If
169:   End Select
End Sub

Private Sub txtAfter_KeyUp(KeyCode As Integer, Shift As Integer)
173:   If DateCheck Then
174:     cmdOK.Enabled = True
175:   Else
176:     cmdOK.Enabled = False
177:   End If
End Sub

Private Sub txtBefore_KeyUp(KeyCode As Integer, Shift As Integer)
181:   If DateCheck Then
182:     cmdOK.Enabled = True
183:   Else
184:     cmdOK.Enabled = False
185:   End If
End Sub

Private Sub txtScale_KeyUp(KeyCode As Integer, Shift As Integer)
189:   If Not IsNumeric(txtScale.Text) Then
190:     txtScale.Text = ""
191:   End If
192:   If ScaleCheck Then
193:     cmdOK.Enabled = True
194:   Else
195:     cmdOK.Enabled = False
196:   End If
End Sub

Private Function ScaleCheck() As Boolean
200:   ScaleCheck = False
201:   If txtScale.Text <> "" Then
202:     If CDbl(txtScale.Text) >= 0 Then
203:       ScaleCheck = True
204:     End If
205:   End If
End Function

Private Function DateCheck() As Boolean
209:   If IsDate(txtBefore.Text) And txtAfter.Text = "" Then
210:     DateCheck = True
211:   ElseIf IsDate(txtBefore.Text) And IsDate(txtAfter.Text) Then
212:     DateCheck = True
213:   ElseIf txtBefore.Text = "" And IsDate(txtAfter.Text) Then
214:     DateCheck = True
215:   Else
216:     DateCheck = False
217:   End If
End Function
