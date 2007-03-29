VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSeriesProperties 
   Caption         =   "Map Series Properties"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   7170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   5940
      TabIndex        =   40
      Top             =   4050
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   345
      Left            =   4800
      TabIndex        =   39
      Top             =   4050
      Width           =   1125
   End
   Begin TabDlg.SSTab tabProperties 
      Height          =   3885
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   6853
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Index Settings"
      TabPicture(0)   =   "frmSeriesProperties.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraPage1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tile Settings"
      TabPicture(1)   =   "frmSeriesProperties.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraPage2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Options"
      TabPicture(2)   =   "frmSeriesProperties.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraPage3"
      Tab(2).ControlCount=   1
      Begin VB.Frame fraPage1 
         BorderStyle     =   0  'None
         Height          =   3375
         Left            =   60
         TabIndex        =   30
         Top             =   420
         Width           =   6855
         Begin VB.ComboBox cmbDetailFrame 
            Enabled         =   0   'False
            Height          =   315
            Left            =   270
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   1140
            Width           =   2625
         End
         Begin VB.Frame fraIndexLayer 
            Caption         =   "Index Layer"
            Height          =   2415
            Left            =   3210
            TabIndex        =   31
            Top             =   900
            Width           =   3525
            Begin VB.ComboBox cmbIndexLayer 
               Enabled         =   0   'False
               Height          =   315
               Left            =   240
               Style           =   2  'Dropdown List
               TabIndex        =   33
               Top             =   510
               Width           =   3105
            End
            Begin VB.ComboBox cmbIndexField 
               Enabled         =   0   'False
               Height          =   315
               Left            =   240
               Style           =   2  'Dropdown List
               TabIndex        =   32
               Top             =   1200
               Width           =   3105
            End
            Begin VB.Label lblMapSheet 
               Caption         =   "Choose the index layer:"
               Height          =   225
               Index           =   3
               Left            =   240
               TabIndex        =   35
               Top             =   270
               Width           =   1725
            End
            Begin VB.Label lblMapSheet 
               Caption         =   "This field specifies the page name"
               Height          =   225
               Index           =   4
               Left            =   240
               TabIndex        =   34
               Top             =   960
               Width           =   2535
            End
         End
         Begin VB.Label Label1 
            Caption         =   $"frmSeriesProperties.frx":0054
            Height          =   615
            Index           =   0
            Left            =   30
            TabIndex        =   38
            Top             =   60
            Width           =   6705
         End
         Begin VB.Label lblMapSheet 
            Caption         =   "Choose the detail data frame:"
            Height          =   225
            Index           =   0
            Left            =   30
            TabIndex        =   37
            Top             =   870
            Width           =   2235
         End
      End
      Begin VB.Frame fraPage2 
         BorderStyle     =   0  'None
         Height          =   3375
         Left            =   -74910
         TabIndex        =   19
         Top             =   420
         Width           =   6765
         Begin VB.TextBox txtNumbering 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2520
            TabIndex        =   43
            Text            =   "1"
            Top             =   2610
            Width           =   525
         End
         Begin VB.Frame fraChooseTiles 
            Caption         =   "Choose tiles"
            Height          =   1485
            Left            =   60
            TabIndex        =   24
            Top             =   750
            Width           =   3255
            Begin VB.OptionButton optTiles 
               Caption         =   "Use all of the tiles"
               Enabled         =   0   'False
               Height          =   285
               Index           =   0
               Left            =   210
               TabIndex        =   27
               Top             =   270
               Width           =   1575
            End
            Begin VB.OptionButton optTiles 
               Caption         =   "Use the selected tiles"
               Enabled         =   0   'False
               Height          =   285
               Index           =   1
               Left            =   210
               TabIndex        =   26
               Top             =   660
               Width           =   1935
            End
            Begin VB.OptionButton optTiles 
               Caption         =   "Use the visible tiles"
               Enabled         =   0   'False
               Height          =   285
               Index           =   2
               Left            =   210
               TabIndex        =   25
               Top             =   1050
               Width           =   1995
            End
         End
         Begin VB.Frame fraSuppressTiles 
            Caption         =   "Suppress tiles"
            Height          =   2565
            Left            =   3420
            TabIndex        =   20
            Top             =   750
            Width           =   3255
            Begin VB.CheckBox chkSuppress 
               Caption         =   "Don't use empty tiles.  A tile is empty"
               Enabled         =   0   'False
               Height          =   225
               Left            =   180
               TabIndex        =   22
               Top             =   300
               Width           =   2865
            End
            Begin VB.ListBox lstSuppressTiles 
               Enabled         =   0   'False
               Height          =   1410
               Left            =   180
               Style           =   1  'Checkbox
               TabIndex        =   21
               Top             =   960
               Width           =   2955
            End
            Begin VB.Label Label2 
               Caption         =   "unless it contains data from at least one of the following selected layers:"
               Height          =   465
               Left            =   450
               TabIndex        =   23
               Top             =   480
               Width           =   2595
            End
         End
         Begin MSComctlLib.ListView lvwSheets 
            Height          =   1515
            Left            =   60
            TabIndex        =   28
            Top             =   4620
            Width           =   5205
            _ExtentX        =   9181
            _ExtentY        =   2672
            View            =   2
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.Label Label1 
            Caption         =   "Begin numbering tiles/pages at:"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   44
            Top             =   2640
            Width           =   2295
         End
         Begin VB.Label Label1 
            Caption         =   $"frmSeriesProperties.frx":0134
            Height          =   615
            Index           =   1
            Left            =   30
            TabIndex        =   29
            Top             =   60
            Width           =   6705
         End
      End
      Begin VB.Frame fraPage3 
         BorderStyle     =   0  'None
         Height          =   3345
         Left            =   -74940
         TabIndex        =   1
         Top             =   420
         Width           =   6825
         Begin VB.Frame fraExtent 
            Caption         =   "Extent"
            Height          =   2865
            Left            =   90
            TabIndex        =   8
            Top             =   450
            Width           =   3255
            Begin VB.OptionButton optExtent 
               Caption         =   "Data driven - The scale for each tile is"
               Height          =   225
               Index           =   2
               Left            =   120
               TabIndex        =   15
               Top             =   1710
               Width           =   3015
            End
            Begin VB.OptionButton optExtent 
               Caption         =   "Fixed - Always draw at this scale:"
               Height          =   285
               Index           =   1
               Left            =   120
               TabIndex        =   14
               Top             =   990
               Width           =   2925
            End
            Begin VB.OptionButton optExtent 
               Caption         =   "Variable - Fit the tiles to the data frame"
               Height          =   225
               Index           =   0
               Left            =   120
               TabIndex        =   13
               Top             =   270
               Width           =   3045
            End
            Begin VB.TextBox txtMargin 
               Height          =   315
               Left            =   1050
               TabIndex        =   12
               Top             =   540
               Width           =   855
            End
            Begin VB.ComboBox cmbMargin 
               Height          =   315
               Left            =   1950
               Style           =   2  'Dropdown List
               TabIndex        =   11
               Top             =   540
               Width           =   1215
            End
            Begin VB.TextBox txtFixed 
               Height          =   315
               Left            =   930
               TabIndex        =   10
               Top             =   1290
               Width           =   945
            End
            Begin VB.ComboBox cmbDataDriven 
               Height          =   315
               Left            =   420
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   2190
               Width           =   2655
            End
            Begin VB.Label Label4 
               Caption         =   "1:"
               Height          =   255
               Index           =   2
               Left            =   780
               TabIndex        =   41
               Top             =   1320
               Width           =   195
            End
            Begin VB.Label Label4 
               Caption         =   " specified in this index layer field:"
               Height          =   255
               Index           =   0
               Left            =   360
               TabIndex        =   17
               Top             =   1920
               Width           =   2595
            End
            Begin VB.Label Label4 
               Caption         =   "Margin"
               Height          =   255
               Index           =   1
               Left            =   480
               TabIndex        =   16
               Top             =   570
               Width           =   495
            End
         End
         Begin VB.Frame fraOptions 
            Caption         =   "Options"
            Height          =   2865
            Left            =   3450
            TabIndex        =   2
            Top             =   450
            Width           =   3255
            Begin VB.CheckBox chkOptions 
               Caption         =   "Select tile when drawing?"
               Height          =   225
               Index           =   4
               Left            =   90
               TabIndex        =   45
               Top             =   2520
               Width           =   2865
            End
            Begin VB.CheckBox chkOptions 
               Caption         =   "Cross-hatch data outside tile?"
               Height          =   225
               Index           =   3
               Left            =   360
               TabIndex        =   42
               Top             =   1260
               Width           =   2565
            End
            Begin VB.CheckBox chkOptions 
               Caption         =   "Rotate data using value from this field:"
               Height          =   225
               Index           =   0
               Left            =   120
               TabIndex        =   7
               Top             =   270
               Width           =   3045
            End
            Begin VB.ComboBox cmbRotateField 
               Height          =   315
               Left            =   390
               Style           =   2  'Dropdown List
               TabIndex        =   6
               Top             =   570
               Width           =   2655
            End
            Begin VB.CheckBox chkOptions 
               Caption         =   "Clip data to the outline of the tile"
               Height          =   225
               Index           =   1
               Left            =   90
               TabIndex        =   5
               Top             =   990
               Width           =   3045
            End
            Begin VB.CheckBox chkOptions 
               Caption         =   "Label neighboring tiles?"
               Height          =   225
               Index           =   2
               Left            =   90
               TabIndex        =   4
               Top             =   1710
               Width           =   3045
            End
            Begin VB.CommandButton cmdLabelProps 
               Caption         =   "Properties..."
               Height          =   345
               Left            =   420
               TabIndex        =   3
               Top             =   2010
               Width           =   1125
            End
         End
         Begin VB.Label Label1 
            Caption         =   "The Map Series provides several different options for fitting a tile to the data frame."
            Height          =   315
            Index           =   2
            Left            =   90
            TabIndex        =   18
            Top             =   180
            Width           =   5955
         End
      End
   End
End
Attribute VB_Name = "frmSeriesProperties"
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
Private m_pSeriesOptions As IDSMapSeriesOptions
Private m_pSeriesOptions2 As IDSMapSeriesOptions2
Private m_pSeriesOptions3 As IDSMapSeriesOptions3 'Added 11/23/04
Private m_bWasSelecting As Boolean                           'Added 11/23/04
Private m_pTextSym As ISimpleTextSymbol

Private Sub chkOptions_Click(Index As Integer)
  Select Case Index
  Case 0  'Rotate
25:     If chkOptions(0).value = 0 Then
26:       cmbRotateField.Enabled = False
27:     Else
28:       cmbRotateField.Enabled = True
29:     End If
  Case 1  'Clip to outline
31:     If chkOptions(1).value = 0 Then
32:       chkOptions(3).value = 0
33:       chkOptions(3).Enabled = False
34:     Else
35:       chkOptions(3).Enabled = True
36:     End If
  Case 2  'Label neighboring tiles
38:     If chkOptions(2).value = 0 Then
39:       cmdLabelProps.Enabled = False
40:     Else
41:       cmdLabelProps.Enabled = True
42:     End If
43:   End Select
End Sub

Private Sub cmdCancel_Click()
47:   Unload Me
End Sub

Private Sub cmdLabelProps_Click()
On Error GoTo ErrHand:
  Dim bChanged As Boolean, pTextSymEditor As ITextSymbolEditor
53:   Set pTextSymEditor = New TextSymbolEditor
54:   bChanged = pTextSymEditor.EditTextSymbol(m_pTextSym, m_pApp.hwnd)
55:   Me.SetFocus
  
  Exit Sub
ErrHand:
59:   MsgBox "cmdLabelProps_Click - " & Err.Description
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrHand:
  Dim pDoc As IMxDocument, pActive As IActiveView
  Dim pFeatSel As IFeatureSelection, pMap As IMap, pSeriesProps As IDSMapSeriesProps
  
  'Apply updates (only the Options can be updated, so we only need to look at those)
  'Set the clip and rotate properties
  'Update 6/18/03 to support cross hatching of clip area
70:   If chkOptions(1).value = 1 Then    'Clip
71:     If chkOptions(3).value = 0 Then   'clip without cross hatch
      'Make sure we don't leave the clip element
73:       If m_pSeriesOptions2.ClipData = 2 Then RemoveClipElement m_pApp.Document
74:       m_pSeriesOptions2.ClipData = 1
75:     Else
76:       m_pSeriesOptions2.ClipData = 2
77:       Set pDoc = m_pApp.Document
78:       pDoc.FocusMap.ClipGeometry = Nothing
79:     End If
'    m_pSeriesOptions.ClipData = True
81:   Else
    'Make sure we don't leave the clip element
83:     If m_pSeriesOptions2.ClipData = 2 Then RemoveClipElement m_pApp.Document
84:     m_pSeriesOptions2.ClipData = 0
'    m_pSeriesOptions.ClipData = False
    'Make sure clipping is turned off for the data frame
87:     Set pDoc = m_pApp.Document
88:     pDoc.FocusMap.ClipGeometry = Nothing
89:   End If
  
91:   If chkOptions(0).value = 1 Then     'Rotation
92:     If m_pSeriesOptions.RotateFrame = False Or m_pSeriesOptions.RotationField <> cmbRotateField.Text Then
93:       UpdatePageValues "ROTATION", cmbRotateField.Text
94:     End If
95:     m_pSeriesOptions.RotateFrame = True
96:     m_pSeriesOptions.RotationField = cmbRotateField.Text
97:   Else
98:     m_pSeriesOptions.RotateFrame = False
    'Make sure rotation is turned off for the data frame
100:     Set pDoc = m_pApp.Document
101:     Set pActive = pDoc.FocusMap
102:     If pActive.ScreenDisplay.DisplayTransformation.Rotation <> 0 Then
103:       pActive.ScreenDisplay.DisplayTransformation.Rotation = 0
104:       pActive.Refresh
105:     End If
106:   End If
107:   If chkOptions(2).value = 1 Then    'Label Neighbors
108:     m_pSeriesOptions.LabelNeighbors = True
109:   Else
110:     m_pSeriesOptions.LabelNeighbors = False
111:     RemoveLabels pDoc
112:     g_bLabelNeighbors = False
113:   End If
114:   Set m_pSeriesOptions.LabelSymbol = m_pTextSym
  
116:   If chkOptions(4).value = 1 Then  'Select tile when drawing
117:     m_pSeriesOptions3.SelectTile = True
118:   Else
119:     m_pSeriesOptions3.SelectTile = False
120:     If m_bWasSelecting Then   'If there were previously selecting tiles, then we need to clear the selection
121:       Set pSeriesProps = m_pSeriesOptions
122:       Set pMap = pActive
123:       Set pFeatSel = FindLayer(pSeriesProps.IndexLayerName, pMap)
124:       If Not pFeatSel Is Nothing Then
125:         pFeatSel.Clear
126:         pActive.PartialRefresh esriViewGeoSelection, Nothing, pActive.Extent
127:       End If
128:     End If
129:   End If
  
  'Set the extent properties
132:   If optExtent(0).value Then         'Variable
133:     m_pSeriesOptions.ExtentType = 0
134:     If txtMargin.Text = "" Then
135:       m_pSeriesOptions.Margin = 0
136:     Else
137:       m_pSeriesOptions.Margin = CDbl(txtMargin.Text)
138:     End If
139:     m_pSeriesOptions.MarginType = cmbMargin.ListIndex
140:   ElseIf optExtent(1).value Then    'Fixed
141:     m_pSeriesOptions.ExtentType = 1
142:     m_pSeriesOptions.FixedScale = txtFixed.Text
143:   Else                        'Data driven
144:     If m_pSeriesOptions.ExtentType <> 2 Or m_pSeriesOptions.RotationField <> cmbRotateField.Text Then
145:       UpdatePageValues "SCALE", cmbDataDriven.Text
146:     End If
147:     m_pSeriesOptions.ExtentType = 2
148:     m_pSeriesOptions.DataDrivenField = cmbDataDriven.Text
149:   End If
  
151:   Unload Me
  
  Exit Sub
  
ErrHand:
156:   MsgBox "cmdOK_Click - " & Err.Description
End Sub

Private Sub UpdatePageValues(sProperty As String, sFieldName As String)
On Error GoTo ErrHand:
  Dim lLoop As Long, pSeries As IDSMapSeries, pPage As IDSMapPage
  Dim pDoc As IMxDocument, pMap As IMap, pSeriesProps As IDSMapSeriesProps
  Dim pIndexLayer As IFeatureLayer, pDataset As IDataset, pWorkspace As IFeatureWorkspace
  Dim pQueryDef As IQueryDef, pCursor As ICursor, pRow As IRow, pColl As Collection
165:   Set pDoc = m_pApp.Document
166:   Set pSeries = m_pSeriesOptions
167:   Set pSeriesProps = pSeries
168:   Set pMap = FindDataFrame(pDoc, pSeriesProps.DataFrameName)
  If pMap Is Nothing Then Exit Sub
  
171:   Set pIndexLayer = FindLayer(pSeriesProps.IndexLayerName, pMap)
  If pIndexLayer Is Nothing Then Exit Sub
  
  'Loop through the features in the index layer creating a collection of the scales and tile names
175:   Set pDataset = pIndexLayer.FeatureClass
176:   Set pWorkspace = pDataset.Workspace
177:   Set pQueryDef = pWorkspace.CreateQueryDef
178:   pQueryDef.Tables = pDataset.Name
179:   pQueryDef.SubFields = sFieldName & "," & pSeriesProps.IndexFieldName
180:   Set pCursor = pQueryDef.Evaluate
181:   Set pColl = New Collection
182:   Set pRow = pCursor.NextRow
183:   Do While Not pRow Is Nothing
184:     If Not IsNull(pRow.value(0)) And Not IsNull(pRow.value(1)) Then
185:       pColl.Add pRow.value(0), pRow.value(1)
186:     End If
187:     Set pRow = pCursor.NextRow
188:   Loop
  
  'Now loop through the pages and try to find the corresponding tile name in the collection
  On Error GoTo ErrNoKey:
192:   For lLoop = 0 To pSeries.PageCount - 1
193:     Set pPage = pSeries.Page(lLoop)
194:     If sProperty = "ROTATION" Then
195:       pPage.PageRotation = pColl.Item(pPage.PageName)
196:     Else
197:       pPage.PageScale = pColl.Item(pPage.PageName)
198:     End If
199:   Next lLoop

  Exit Sub

ErrNoKey:
204:   Resume Next
ErrHand:
206:   MsgBox "UpdatePageValues - " & Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo ErrHand:
  Dim pMapBook As IDSMapBook
  Dim pSeriesProps As IDSMapSeriesProps
  Dim lLoop As Long
  'Check to see if a MapSeries already exists
215:   Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
218:   Set pSeriesProps = pMapBook.ContentItem(0)
219:   Set m_pSeriesOptions = pSeriesProps
220:   Set m_pSeriesOptions2 = m_pSeriesOptions
221:   Set m_pSeriesOptions3 = m_pSeriesOptions
  
  'Index Settings Tab
224:   cmbDetailFrame.Clear
225:   cmbDetailFrame.AddItem pSeriesProps.DataFrameName
226:   cmbDetailFrame.Text = pSeriesProps.DataFrameName
227:   cmbIndexLayer.Clear
228:   cmbIndexLayer.AddItem pSeriesProps.IndexLayerName
229:   cmbIndexLayer.Text = pSeriesProps.IndexLayerName
230:   cmbIndexField.Clear
231:   cmbIndexField.AddItem pSeriesProps.IndexFieldName
232:   cmbIndexField.Text = pSeriesProps.IndexFieldName
  
  'Tile Settings Tab
235:   optTiles(pSeriesProps.TileSelectionMethod) = True
236:   lstSuppressTiles.Clear
237:   If pSeriesProps.SuppressLayers Then
238:     chkSuppress.value = 1
239:     For lLoop = 0 To pSeriesProps.SuppressLayerCount - 1
240:       lstSuppressTiles.AddItem pSeriesProps.SuppressLayer(lLoop)
241:       lstSuppressTiles.Selected(lLoop) = True
242:     Next lLoop
243:   Else
244:     chkSuppress.value = 0
245:   End If
246:   txtNumbering.Text = CStr(pSeriesProps.StartNumber)  'Added 2/18/2004
  
  'Options tab
249:   PopulateFieldCombos
250:   cmbMargin.Clear
251:   cmbMargin.AddItem "percent"
252:   cmbMargin.AddItem "mapunits"
253:   cmbMargin.Text = "percent"
254:   optExtent(m_pSeriesOptions.ExtentType).value = True
255:   cmdOK.Enabled = True
  Select Case m_pSeriesOptions.ExtentType
  Case 0
258:     txtMargin.Text = m_pSeriesOptions.Margin
259:     If m_pSeriesOptions.MarginType = 0 Then
260:       cmbMargin.Text = "percent"
261:     Else
262:       cmbMargin.Text = "mapunits"
263:     End If
  Case 1
265:     txtFixed.Text = m_pSeriesOptions.FixedScale
  Case 2
267:     cmbDataDriven.Text = m_pSeriesOptions.DataDrivenField
268:   End Select
269:   If m_pSeriesOptions.RotateFrame Then
270:     chkOptions(0).value = 1
271:     cmbRotateField.Text = m_pSeriesOptions.RotationField
272:   Else
273:     chkOptions(0).value = 0
274:   End If
  
  'Update 6/18/03 to support cross hatching of clip area
  Select Case m_pSeriesOptions2.ClipData
  Case 0   'No clipping
279:     chkOptions(1).value = 0
280:     chkOptions(3).value = 0
281:     chkOptions(3).Enabled = False
  Case 1   'Clip only
283:     chkOptions(1).value = 1
284:     chkOptions(3).value = 0
285:     chkOptions(3).Enabled = True
  Case 2   'Clip with cross hatch outside clip area
287:     chkOptions(1).value = 1
288:     chkOptions(3).value = 1
289:     chkOptions(3).Enabled = True
290:   End Select
'  If m_pSeriesOptions.ClipData Then
'    chkOptions(1).Value = 1
'  Else
'    chkOptions(1).Value = 0
'  End If

297:   If m_pSeriesOptions.LabelNeighbors Then
298:     chkOptions(2).value = 1
299:     cmdLabelProps.Enabled = True
300:   Else
301:     chkOptions(2).value = 0
302:     cmdLabelProps.Enabled = False
303:   End If
304:   Set m_pTextSym = m_pSeriesOptions.LabelSymbol
  
306:   If m_pSeriesOptions3.SelectTile Then  'Added 11/23/04
307:     chkOptions(4).value = 1
308:     m_bWasSelecting = True
309:   Else
310:     chkOptions(4).value = 0
311:     m_bWasSelecting = False
312:   End If
  
  'Make sure the wizard stays on top
315:   TopMost Me
  
  Exit Sub
ErrHand:
319:   MsgBox "frmSeriesProperties_Load - " & Err.Description
End Sub

Private Sub PopulateFieldCombos()
On Error GoTo ErrHand:
  Dim pIndexLayer As IFeatureLayer, pMap As IMap, lLoop As Long
  Dim pFields As IFields, pDoc As IMxDocument
  
327:   Set pDoc = m_pApp.Document
328:   Set pMap = FindDataFrame(pDoc, cmbDetailFrame.Text)
329:   If pMap Is Nothing Then
330:     MsgBox "Could not find detail frame!!!"
    Exit Sub
332:   End If
  
334:   Set pIndexLayer = FindLayer(cmbIndexLayer.Text, pMap)
335:   If pIndexLayer Is Nothing Then
336:     MsgBox "Could not find specified layer!!!"
    Exit Sub
338:   End If
  
  'Populate the index layer combos
341:   Set pFields = pIndexLayer.FeatureClass.Fields
342:   cmbDataDriven.Clear
343:   cmbRotateField.Clear
344:   For lLoop = 0 To pFields.FieldCount - 1
    Select Case pFields.Field(lLoop).Type
    Case esriFieldTypeDouble, esriFieldTypeSingle, esriFieldTypeInteger
347:       If UCase(pFields.Field(lLoop).Name) <> "SHAPE_LENGTH" And _
       UCase(pFields.Field(lLoop).Name) <> "SHAPE_AREA" Then
349:         cmbDataDriven.AddItem pFields.Field(lLoop).Name
350:         cmbRotateField.AddItem pFields.Field(lLoop).Name
351:       End If
352:     End Select
353:   Next lLoop
354:   If cmbDataDriven.ListCount > 0 Then
355:     cmbDataDriven.ListIndex = 0
356:     cmbRotateField.ListIndex = 0
357:     optExtent.Item(2).Enabled = True
358:     chkOptions(0).Enabled = True
359:   Else
360:     optExtent.Item(2).Enabled = False
361:     chkOptions(0).Enabled = False
362:   End If
  
  Exit Sub
  
ErrHand:
367:   MsgBox "PopulateFieldCombos - " & Err.Description
End Sub

Private Sub Form_Terminate()
371:   Set m_pApp = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
375:   Set m_pApp = Nothing
End Sub

Private Sub optExtent_Click(Index As Integer)
On Error GoTo ErrHand:
  Select Case Index
  Case 0  'Variable
382:     txtMargin.Enabled = True
383:     cmbMargin.Enabled = True
384:     txtFixed.Enabled = False
385:     cmbDataDriven.Enabled = False
386:     If txtMargin.Text = "" Then
387:       cmdOK.Enabled = False
388:     Else
389:       cmdOK.Enabled = True
390:     End If
  Case 1  'Fixed
392:     txtMargin.Enabled = False
393:     cmbMargin.Enabled = False
394:     txtFixed.Enabled = True
395:     cmbDataDriven.Enabled = False
396:     If txtFixed.Text = "" Then
397:       cmdOK.Enabled = False
398:     Else
399:       cmdOK.Enabled = True
400:     End If
  Case 2  'Data driven
402:     txtMargin.Enabled = False
403:     cmbMargin.Enabled = False
404:     txtFixed.Enabled = False
405:     cmbDataDriven.Enabled = True
406:     cmdOK.Enabled = True
407:   End Select

  Exit Sub
ErrHand:
411:   MsgBox "optExtent_Click - " & Err.Description
End Sub

Private Sub txtFixed_KeyUp(KeyCode As Integer, Shift As Integer)
415:   If Not IsNumeric(txtFixed.Text) Then
416:     txtFixed.Text = ""
417:   End If
418:   If txtFixed.Text <> "" Then
419:     cmdOK.Enabled = True
420:   End If
End Sub

Private Sub txtMargin_KeyUp(KeyCode As Integer, Shift As Integer)
424:   If Not IsNumeric(txtMargin.Text) Then
425:     txtMargin.Text = ""
426:   End If
427:   If txtMargin.Text <> "" Then
428:     cmdOK.Enabled = True
429:   End If
End Sub
