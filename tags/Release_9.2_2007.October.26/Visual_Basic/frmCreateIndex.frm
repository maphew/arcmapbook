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
  
20:   Set pFeatLayer = FindFeatureLayerByName(cmbLayer.List(cmbLayer.ListIndex), m_pApp)
  If pFeatLayer Is Nothing Then Exit Sub
  
23:   cmbFieldName.Clear
24:   Set pFields = pFeatLayer.FeatureClass.Fields
25:   For lLoop = 0 To pFields.FieldCount - 1
26:     If pFields.Field(lLoop).Type = esriFieldTypeString Then
27:       cmbFieldName.AddItem pFields.Field(lLoop).Name
28:     End If
29:   Next lLoop
30:   If cmbFieldName.ListCount > 0 Then
31:     cmbFieldName.ListIndex = 0
32:   End If
  
34:   CheckSettings

  Exit Sub
ErrHand:
38:   MsgBox "cmbLayer_Click - " & Err.Description
End Sub

Private Sub cmdBrowse_Click()
On Error GoTo ErrHand:
43:   codOutput.DialogTitle = "Specify output file to create"
44:   codOutput.Filter = "(*.txt)|*.txt"
45:   codOutput.Flags = cdlOFNOverwritePrompt
46:   codOutput.ShowSave
47:   If codOutput.FileName = "" Then
48:     txtOutput.Text = ""
49:   Else
50:     txtOutput.Text = codOutput.FileName
51:   End If
  
53:   CheckSettings
  
  Exit Sub
ErrHand:
57:   MsgBox "cmdBrowse_Click - " & Err.Description
End Sub

Private Sub cmdCancel_Click()
61:   Set m_pApp = Nothing
62:   Unload Me
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
75:   Set pMapBook = GetMapBookExtension(m_pApp)
76:   If pMapBook Is Nothing Then
77:     MsgBox "Map book was not found!!!"
    Exit Sub
79:   End If
  
  'Get the index layer
82:   Set pFeatLayer = FindFeatureLayerByName(cmbLayer.List(cmbLayer.ListIndex), m_pApp)
83:   If pFeatLayer Is Nothing Then
84:     MsgBox "Count not find the index layer for some reason!!!"
    Exit Sub
86:   End If
87:   sFieldName = cmbFieldName.List(cmbFieldName.ListIndex)
88:   Set pSeries = pMapBook.ContentItem(0)
  
  'Setup the progress bar
91:   Screen.MousePointer = vbHourglass
92:   With m_pApp.StatusBar.ProgressBar
93:     .Message = "Index Creation:"
94:     .MaxRange = pSeries.PageCount
95:     .StepValue = 1
96:     .Position = 1
97:     .Show
98:   End With
  
  'Loop through the pages returning a collection of the attribute values returned by the
  'features found on the page.  Add the page collections to the master.
102:   Set m_pMasterColl = New Collection
103:   For lLoop = 0 To pSeries.PageCount - 1
104:     Set pPage = pSeries.Page(lLoop)
105:     Set pColl = pPage.IndexPage(pFeatLayer, sFieldName)
106:     If optIndex(0).value Then
107:       sPageId = pPage.PageName
108:     Else
109:       sPageId = CStr(lLoop + 1 + CLng(txtPageNumber.Text))
110:     End If
111:     AddPageToMasterCollection pColl, sPageId
    
113:     m_pApp.StatusBar.ProgressBar.Step
114:   Next lLoop
  
  'Dump the master collection out to the specified file
117:   sPrev = ""
118:   Open txtOutput.Text For Output As #1
119:   For lLoop = 1 To m_pMasterColl.count
120:     sTemp = m_pMasterColl.Item(lLoop)
121:     lPos = InStr(1, sTemp, "-$$$$-")
122:     sTempVal = Left(sTemp, lPos - 1)
123:     sTempPage = Mid(sTemp, lPos + 6)
124:     If sPrev = "" Then
125:       Set pOutputPages = New Collection
126:       sOutput = sTempVal
127:       pOutputPages.Add sTempPage, sTempPage
128:       sPrev = sTempVal
129:     ElseIf sPrev = sTempVal Then
130:       If optIndex(0).value Then
131:         pOutputPages.Add sTempPage, sTempPage
132:       Else
133:         For lLoop2 = 1 To pOutputPages.count
134:           If CLng(sTempPage) < CLng(pOutputPages.Item(lLoop2)) Then
135:             pOutputPages.Add sTempPage, sTempPage, lLoop2
136:             Exit For
137:           End If
138:           If lLoop2 = pOutputPages.count Then
139:             pOutputPages.Add sTempPage, sTempPage
140:           End If
141:         Next lLoop2
142:       End If
143:     Else
144:       For lLoop2 = 1 To pOutputPages.count
145:         If lLoop2 = 1 Then
146:           sOutputPages = pOutputPages.Item(lLoop2)
147:         Else
148:           sOutputPages = sOutputPages & ", " & pOutputPages.Item(lLoop2)
149:         End If
150:       Next lLoop2
151:       Print #1, sOutput & ": " & sOutputPages
152:       sOutput = sTempVal
153:       Set pOutputPages = New Collection
154:       pOutputPages.Add sTempPage, sTempPage
155:       sPrev = sTempVal
156:     End If
157:     If lLoop = m_pMasterColl.count Then
158:       For lLoop2 = 1 To pOutputPages.count
159:         If lLoop2 = 1 Then
160:           sOutputPages = pOutputPages.Item(lLoop2)
161:         Else
162:           sOutputPages = sOutputPages & ", " & pOutputPages.Item(lLoop2)
163:         End If
164:       Next lLoop2
165:       Print #1, sOutput & ": " & sOutputPages
166:     End If
167:   Next lLoop
168:   Close #1
  
170:   m_pApp.StatusBar.ProgressBar.Hide
171:   Screen.MousePointer = vbNormal
172:   Unload Me
  
  Exit Sub
ErrHand:
176:   Screen.MousePointer = vbNormal
177:   MsgBox "cmdOK_Click - " & Erl & " - " & Err.Description
End Sub

Private Sub AddPageToMasterCollection(pColl As Collection, sPageId As String)
On Error GoTo ErrHand:
  Dim lLoop As Long, sValue As String, lLoop2 As Long, lStart As Long
183:   lStart = 1
184:   If m_pMasterColl.count = 0 Then
185:     For lLoop = 1 To pColl.count
186:       sValue = pColl.Item(lLoop) & "-$$$$-" & sPageId
187:       m_pMasterColl.Add sValue, sValue
188:     Next lLoop
189:   Else
190:     For lLoop = 1 To pColl.count
191:       sValue = pColl.Item(lLoop) & "-$$$$-" & sPageId
192:       For lLoop2 = lStart To m_pMasterColl.count
193:         If sValue < m_pMasterColl.Item(lLoop2) Then
194:           m_pMasterColl.Add sValue, sValue, lLoop2
195:           lStart = lLoop2
196:           Exit For
197:         End If
198:         If lLoop2 = m_pMasterColl.count Then
199:           m_pMasterColl.Add sValue, sValue
200:           lStart = lLoop2
201:         End If
202:       Next lLoop2
203:     Next lLoop
204:   End If

  Exit Sub
ErrHand:
208:   MsgBox "AddPageToMasterCollection - " & Erl & " - " & Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo ErrHand:
  Dim pDoc As IMxDocument, pMap As IMap, lLoop As Long
  Dim pFeatLayer As IFeatureLayer
  Dim pMapBook As IDSMapBook
  Dim pSeriesProps As IDSMapSeriesProps
217:   Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
220:   Set pSeriesProps = pMapBook.ContentItem(0)

222:   optIndex(0).value = True
223:   txtPageNumber.Text = "0"
  
  'Populate the layer list box
226:   cmbLayer.Clear
227:   Set pDoc = m_pApp.Document
228:   Set pMap = pDoc.FocusMap
229:   For lLoop = 0 To pMap.LayerCount - 1
230:     If TypeOf pMap.Layer(lLoop) Is IFeatureLayer Then
231:       Set pFeatLayer = pMap.Layer(lLoop)
232:       If pFeatLayer.FeatureClass.FeatureType <> esriFTAnnotation And _
       pFeatLayer.FeatureClass.FeatureType <> esriFTDimension And _
       pFeatLayer.FeatureClass.FeatureType <> esriFTCoverageAnnotation Then
235:         If UCase(pFeatLayer.Name) <> UCase(pSeriesProps.IndexLayerName) Then
236:           cmbLayer.AddItem pFeatLayer.Name
237:         End If
238:       End If
239:     End If
240:   Next lLoop
241:   If cmbLayer.ListCount > 0 Then
242:     cmbLayer.ListIndex = 0
243:   End If
  
  'Make sure the wizard stays on top
246:   TopMost Me
  
  Exit Sub
ErrHand:
250:   MsgBox "frmCreateIndex_Load - " & Err.Description
End Sub

Private Sub optIndex_Click(Index As Integer)
254:   If Index = 0 Then
255:     txtPageNumber.Enabled = False
256:   Else
257:     txtPageNumber.Enabled = True
258:   End If
259:   CheckSettings
End Sub

Private Sub txtPageNumber_KeyUp(KeyCode As Integer, Shift As Integer)
263:   If txtPageNumber.Text = "" Then
264:     cmdOK.Enabled = False
265:   Else
266:     If Not IsNumeric(txtPageNumber.Text) Then
267:       txtPageNumber.Text = "0"
268:     End If
269:     CheckSettings
270:   End If
End Sub

Private Sub CheckSettings()
274:   If optIndex(0).value = True Then
275:     If txtOutput.Text <> "" Then
276:       cmdOK.Enabled = True
277:     Else
278:       cmdOK.Enabled = False
279:     End If
280:   Else
281:     If txtOutput.Text <> "" And txtPageNumber.Text <> "" Then
282:       cmdOK.Enabled = True
283:     Else
284:       cmdOK.Enabled = False
285:     End If
286:   End If
End Sub
