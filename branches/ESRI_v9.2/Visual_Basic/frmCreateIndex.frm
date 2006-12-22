VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCreateIndex 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Create Index"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog codOutput 
      Left            =   90
      Top             =   4020
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1125
      Left            =   60
      TabIndex        =   10
      Top             =   2820
      Width           =   4515
      Begin VB.TextBox txtPageNumber 
         Height          =   315
         Left            =   3720
         TabIndex        =   15
         Top             =   660
         Width           =   465
      End
      Begin VB.OptionButton optIndex 
         Caption         =   "Page Number (Number shown on Series list)"
         Height          =   225
         Index           =   1
         Left            =   810
         TabIndex        =   13
         Top             =   420
         Width           =   3405
      End
      Begin VB.OptionButton optIndex 
         Caption         =   "Page Label"
         Height          =   225
         Index           =   0
         Left            =   810
         TabIndex        =   12
         Top             =   60
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Add this value to each page number:"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   14
         Top             =   690
         Width           =   2625
      End
      Begin VB.Label Label3 
         Caption         =   "Index by:"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   11
         Top             =   30
         Width           =   735
      End
   End
   Begin VB.ComboBox cmbFieldName 
      Height          =   315
      Left            =   870
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1860
      Width           =   3405
   End
   Begin VB.ComboBox cmbLayer 
      Height          =   315
      Left            =   870
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1410
      Width           =   3405
   End
   Begin VB.CommandButton cmdBrowse 
      Height          =   315
      Left            =   4350
      Picture         =   "frmCreateIndex.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2310
      Width           =   345
   End
   Begin VB.TextBox txtOutput 
      Enabled         =   0   'False
      Height          =   315
      Left            =   870
      TabIndex        =   2
      Top             =   2310
      Width           =   3375
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4050
      TabIndex        =   1
      Top             =   4110
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   4110
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Field:"
      Height          =   255
      Index           =   1
      Left            =   390
      TabIndex        =   9
      Top             =   1890
      Width           =   405
   End
   Begin VB.Label Label2 
      Caption         =   "Layer:"
      Height          =   255
      Index           =   0
      Left            =   330
      TabIndex        =   6
      Top             =   1440
      Width           =   465
   End
   Begin VB.Label Label1 
      Caption         =   $"frmCreateIndex.frx":047A
      Height          =   1215
      Left            =   60
      TabIndex        =   5
      Top             =   90
      Width           =   4665
   End
   Begin VB.Label lblExportTo 
      Caption         =   "Output to:"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   2340
      Width           =   735
   End
End
Attribute VB_Name = "frmCreateIndex"
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

Public m_pApp As IApplication
Private m_pMasterColl As Collection

Private Sub cmbLayer_Click()
On Error GoTo ErrHand:
  Dim pFeatLayer As IFeatureLayer, pFields As IFields, lLoop As Long
  
8:   Set pFeatLayer = FindFeatureLayerByName(cmbLayer.List(cmbLayer.ListIndex), m_pApp)
  If pFeatLayer Is Nothing Then Exit Sub
  
11:   cmbFieldName.Clear
12:   Set pFields = pFeatLayer.FeatureClass.Fields
13:   For lLoop = 0 To pFields.FieldCount - 1
14:     If pFields.Field(lLoop).Type = esriFieldTypeString Then
15:       cmbFieldName.AddItem pFields.Field(lLoop).Name
16:     End If
17:   Next lLoop
18:   If cmbFieldName.ListCount > 0 Then
19:     cmbFieldName.ListIndex = 0
20:   End If
  
22:   CheckSettings

  Exit Sub
ErrHand:
26:   MsgBox "cmbLayer_Click - " & Err.Description
End Sub

Private Sub cmdBrowse_Click()
On Error GoTo ErrHand:
31:   codOutput.DialogTitle = "Specify output file to create"
32:   codOutput.Filter = "(*.txt)|*.txt"
33:   codOutput.Flags = cdlOFNOverwritePrompt
34:   codOutput.ShowSave
35:   If codOutput.FileName = "" Then
36:     txtOutput.Text = ""
37:   Else
38:     txtOutput.Text = codOutput.FileName
39:   End If
  
41:   CheckSettings
  
  Exit Sub
ErrHand:
45:   MsgBox "cmdBrowse_Click - " & Err.Description
End Sub

Private Sub cmdCancel_Click()
49:   Set m_pApp = Nothing
50:   Unload Me
End Sub

Private Sub cmdOK_Click()
'This routine will create the index and write it out to the specified file.
On Error GoTo ErrHand:
  Dim pDoc As IMxDocument, pMap As IMap, lLoop As Long
  Dim pFeatLayer As IFeatureLayer, pMapBook As IDSMapBook, pSeries As IDSMapSeries
  Dim pPage As IDSMapPage, pColl As Collection
  Dim sPageId As String, sFieldName As String
  Dim sTempVal As String, sTempPage As String, lPos As Long
  Dim sPrev As String, sTemp As String, sOutput As String
  Dim pOutputPages As Collection, lLoop2 As Long, sOutputPages As String
63:   Set pMapBook = GetMapBookExtension(m_pApp)
64:   If pMapBook Is Nothing Then
65:     MsgBox "Map book was not found!!!"
    Exit Sub
67:   End If
  
  'Get the index layer
70:   Set pFeatLayer = FindFeatureLayerByName(cmbLayer.List(cmbLayer.ListIndex), m_pApp)
71:   If pFeatLayer Is Nothing Then
72:     MsgBox "Count not find the index layer for some reason!!!"
    Exit Sub
74:   End If
75:   sFieldName = cmbFieldName.List(cmbFieldName.ListIndex)
76:   Set pSeries = pMapBook.ContentItem(0)
  
  'Setup the progress bar
79:   Screen.MousePointer = vbHourglass
80:   With m_pApp.StatusBar.ProgressBar
81:     .Message = "Index Creation:"
82:     .MaxRange = pSeries.PageCount
83:     .StepValue = 1
84:     .Position = 1
85:     .Show
86:   End With
  
  'Loop through the pages returning a collection of the attribute values returned by the
  'features found on the page.  Add the page collections to the master.
90:   Set m_pMasterColl = New Collection
91:   For lLoop = 0 To pSeries.PageCount - 1
92:     Set pPage = pSeries.Page(lLoop)
93:     Set pColl = pPage.IndexPage(pFeatLayer, sFieldName)
94:     If optIndex(0).Value Then
95:       sPageId = pPage.PageName
96:     Else
97:       sPageId = CStr(lLoop + 1 + CLng(txtPageNumber.Text))
98:     End If
99:     AddPageToMasterCollection pColl, sPageId
    
101:     m_pApp.StatusBar.ProgressBar.Step
102:   Next lLoop
  
  'Dump the master collection out to the specified file
105:   sPrev = ""
106:   Open txtOutput.Text For Output As #1
107:   For lLoop = 1 To m_pMasterColl.count
108:     sTemp = m_pMasterColl.Item(lLoop)
109:     lPos = InStr(1, sTemp, "-$$$$-")
110:     sTempVal = Left(sTemp, lPos - 1)
111:     sTempPage = Mid(sTemp, lPos + 6)
112:     If sPrev = "" Then
113:       Set pOutputPages = New Collection
114:       sOutput = sTempVal
115:       pOutputPages.Add sTempPage, sTempPage
116:       sPrev = sTempVal
117:     ElseIf sPrev = sTempVal Then
118:       If optIndex(0).Value Then
119:         pOutputPages.Add sTempPage, sTempPage
120:       Else
121:         For lLoop2 = 1 To pOutputPages.count
122:           If CLng(sTempPage) < CLng(pOutputPages.Item(lLoop2)) Then
123:             pOutputPages.Add sTempPage, sTempPage, lLoop2
124:             Exit For
125:           End If
126:           If lLoop2 = pOutputPages.count Then
127:             pOutputPages.Add sTempPage, sTempPage
128:           End If
129:         Next lLoop2
130:       End If
131:     Else
132:       For lLoop2 = 1 To pOutputPages.count
133:         If lLoop2 = 1 Then
134:           sOutputPages = pOutputPages.Item(lLoop2)
135:         Else
136:           sOutputPages = sOutputPages & ", " & pOutputPages.Item(lLoop2)
137:         End If
138:       Next lLoop2
139:       Print #1, sOutput & ": " & sOutputPages
140:       sOutput = sTempVal
141:       Set pOutputPages = New Collection
142:       pOutputPages.Add sTempPage, sTempPage
143:       sPrev = sTempVal
144:     End If
145:     If lLoop = m_pMasterColl.count Then
146:       For lLoop2 = 1 To pOutputPages.count
147:         If lLoop2 = 1 Then
148:           sOutputPages = pOutputPages.Item(lLoop2)
149:         Else
150:           sOutputPages = sOutputPages & ", " & pOutputPages.Item(lLoop2)
151:         End If
152:       Next lLoop2
153:       Print #1, sOutput & ": " & sOutputPages
154:     End If
155:   Next lLoop
156:   Close #1
  
158:   m_pApp.StatusBar.ProgressBar.Hide
159:   Screen.MousePointer = vbNormal
160:   Unload Me
  
  Exit Sub
ErrHand:
164:   Screen.MousePointer = vbNormal
165:   MsgBox "cmdOK_Click - " & Erl & " - " & Err.Description
End Sub

Private Sub AddPageToMasterCollection(pColl As Collection, sPageId As String)
On Error GoTo ErrHand:
  Dim lLoop As Long, sValue As String, lLoop2 As Long, lStart As Long
171:   lStart = 1
172:   If m_pMasterColl.count = 0 Then
173:     For lLoop = 1 To pColl.count
174:       sValue = pColl.Item(lLoop) & "-$$$$-" & sPageId
175:       m_pMasterColl.Add sValue, sValue
176:     Next lLoop
177:   Else
178:     For lLoop = 1 To pColl.count
179:       sValue = pColl.Item(lLoop) & "-$$$$-" & sPageId
180:       For lLoop2 = lStart To m_pMasterColl.count
181:         If sValue < m_pMasterColl.Item(lLoop2) Then
182:           m_pMasterColl.Add sValue, sValue, lLoop2
183:           lStart = lLoop2
184:           Exit For
185:         End If
186:         If lLoop2 = m_pMasterColl.count Then
187:           m_pMasterColl.Add sValue, sValue
188:           lStart = lLoop2
189:         End If
190:       Next lLoop2
191:     Next lLoop
192:   End If

  Exit Sub
ErrHand:
196:   MsgBox "AddPageToMasterCollection - " & Erl & " - " & Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo ErrHand:
  Dim pDoc As IMxDocument, pMap As IMap, lLoop As Long
  Dim pFeatLayer As IFeatureLayer
  Dim pMapBook As IDSMapBook
  Dim pSeriesProps As IDSMapSeriesProps
205:   Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
208:   Set pSeriesProps = pMapBook.ContentItem(0)

210:   optIndex(0).Value = True
211:   txtPageNumber.Text = "0"
  
  'Populate the layer list box
214:   cmbLayer.Clear
215:   Set pDoc = m_pApp.Document
216:   Set pMap = pDoc.FocusMap
217:   For lLoop = 0 To pMap.LayerCount - 1
218:     If TypeOf pMap.Layer(lLoop) Is IFeatureLayer Then
219:       Set pFeatLayer = pMap.Layer(lLoop)
220:       If pFeatLayer.FeatureClass.FeatureType <> esriFTAnnotation And _
       pFeatLayer.FeatureClass.FeatureType <> esriFTDimension And _
       pFeatLayer.FeatureClass.FeatureType <> esriFTCoverageAnnotation Then
223:         If UCase(pFeatLayer.Name) <> UCase(pSeriesProps.IndexLayerName) Then
224:           cmbLayer.AddItem pFeatLayer.Name
225:         End If
226:       End If
227:     End If
228:   Next lLoop
229:   If cmbLayer.ListCount > 0 Then
230:     cmbLayer.ListIndex = 0
231:   End If
  
  'Make sure the wizard stays on top
234:   TopMost Me
  
  Exit Sub
ErrHand:
238:   MsgBox "frmCreateIndex_Load - " & Err.Description
End Sub

Private Sub optIndex_Click(Index As Integer)
242:   If Index = 0 Then
243:     txtPageNumber.Enabled = False
244:   Else
245:     txtPageNumber.Enabled = True
246:   End If
247:   CheckSettings
End Sub

Private Sub txtPageNumber_KeyUp(KeyCode As Integer, Shift As Integer)
251:   If txtPageNumber.Text = "" Then
252:     cmdOK.Enabled = False
253:   Else
254:     If Not IsNumeric(txtPageNumber.Text) Then
255:       txtPageNumber.Text = "0"
256:     End If
257:     CheckSettings
258:   End If
End Sub

Private Sub CheckSettings()
262:   If optIndex(0).Value = True Then
263:     If txtOutput.Text <> "" Then
264:       cmdOK.Enabled = True
265:     Else
266:       cmdOK.Enabled = False
267:     End If
268:   Else
269:     If txtOutput.Text <> "" And txtPageNumber.Text <> "" Then
270:       cmdOK.Enabled = True
271:     Else
272:       cmdOK.Enabled = False
273:     End If
274:   End If
End Sub
