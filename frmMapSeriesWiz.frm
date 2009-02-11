VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMapSeriesWiz 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Map Sheet Wizard"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "< Back"
      Height          =   345
      Left            =   3330
      TabIndex        =   2
      Top             =   3780
      Width           =   1125
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next >"
      Height          =   345
      Left            =   4470
      TabIndex        =   1
      Top             =   3780
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   5760
      TabIndex        =   0
      Top             =   3780
      Width           =   1125
   End
   Begin VB.Frame fraPage3 
      BorderStyle     =   0  'None
      Height          =   3525
      Left            =   60
      TabIndex        =   3
      Top             =   30
      Visible         =   0   'False
      Width           =   6855
      Begin VB.Frame fraOptions 
         Caption         =   "Options"
         Height          =   2865
         Left            =   3450
         TabIndex        =   28
         Top             =   450
         Width           =   3255
         Begin VB.CheckBox chkOptions 
            Caption         =   "Select tile when drawing?"
            Height          =   225
            Index           =   4
            Left            =   120
            TabIndex        =   45
            Top             =   2520
            Width           =   2865
         End
         Begin VB.CheckBox chkOptions 
            Caption         =   "Cross-hatch data outside tile?"
            Height          =   225
            Index           =   3
            Left            =   390
            TabIndex        =   42
            Top             =   1290
            Width           =   2565
         End
         Begin VB.CommandButton cmdLabelProps 
            Caption         =   "Properties..."
            Height          =   345
            Left            =   420
            TabIndex        =   40
            Top             =   2010
            Width           =   1125
         End
         Begin VB.CheckBox chkOptions 
            Caption         =   "Label neighboring tiles?"
            Height          =   225
            Index           =   2
            Left            =   120
            TabIndex        =   39
            Top             =   1710
            Width           =   3045
         End
         Begin VB.CheckBox chkOptions 
            Caption         =   "Clip data to the outline of the tile"
            Height          =   225
            Index           =   1
            Left            =   90
            TabIndex        =   38
            Top             =   990
            Width           =   3045
         End
         Begin VB.ComboBox cmbRotateField 
            Height          =   315
            Left            =   390
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   570
            Width           =   2655
         End
         Begin VB.CheckBox chkOptions 
            Caption         =   "Rotate data using value from this field:"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   29
            Top             =   270
            Width           =   3045
         End
      End
      Begin VB.Frame fraExtent 
         Caption         =   "Extent"
         Height          =   2865
         Left            =   90
         TabIndex        =   24
         Top             =   450
         Width           =   3255
         Begin VB.ComboBox cmbDataDriven 
            Height          =   315
            Left            =   420
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   2190
            Width           =   2655
         End
         Begin VB.TextBox txtFixed 
            Height          =   315
            Left            =   750
            TabIndex        =   35
            Top             =   1290
            Width           =   945
         End
         Begin VB.ComboBox cmbMargin 
            Height          =   315
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   540
            Width           =   1215
         End
         Begin VB.TextBox txtMargin 
            Height          =   315
            Left            =   1050
            TabIndex        =   33
            Top             =   540
            Width           =   855
         End
         Begin VB.OptionButton optExtent 
            Caption         =   "Variable - Fit the tiles to the data frame"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   27
            Top             =   270
            Width           =   3045
         End
         Begin VB.OptionButton optExtent 
            Caption         =   "Fixed - Always draw at this scale:"
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   26
            Top             =   990
            Width           =   2925
         End
         Begin VB.OptionButton optExtent 
            Caption         =   "Data driven - The scale for each tile is"
            Height          =   225
            Index           =   2
            Left            =   120
            TabIndex        =   25
            Top             =   1710
            Width           =   3015
         End
         Begin VB.Label Label4 
            Caption         =   "1:"
            Height          =   255
            Index           =   2
            Left            =   600
            TabIndex        =   41
            Top             =   1320
            Width           =   195
         End
         Begin VB.Label Label4 
            Caption         =   "Margin"
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   32
            Top             =   570
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   " specified in this index layer field:"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   31
            Top             =   1920
            Width           =   2595
         End
      End
      Begin VB.Label Label1 
         Caption         =   "The Map Series provides several different options for fitting a tile to the data frame."
         Height          =   315
         Index           =   2
         Left            =   90
         TabIndex        =   30
         Top             =   180
         Width           =   5955
      End
   End
   Begin VB.Frame fraPage2 
      BorderStyle     =   0  'None
      Height          =   3585
      Left            =   30
      TabIndex        =   5
      Top             =   30
      Visible         =   0   'False
      Width           =   6855
      Begin VB.TextBox txtNumbering 
         Height          =   315
         Left            =   2370
         TabIndex        =   43
         Text            =   "1"
         Top             =   2580
         Width           =   525
      End
      Begin VB.Frame fraSuppressTiles 
         Caption         =   "Suppress tiles"
         Height          =   2565
         Left            =   3420
         TabIndex        =   20
         Top             =   750
         Width           =   3255
         Begin VB.ListBox lstSuppressTiles 
            Height          =   1410
            Left            =   180
            Style           =   1  'Checkbox
            TabIndex        =   23
            Top             =   960
            Width           =   2955
         End
         Begin VB.CheckBox chkSuppress 
            Caption         =   "Don't use empty tiles.  A tile is empty"
            Height          =   225
            Left            =   180
            TabIndex        =   21
            Top             =   300
            Width           =   2865
         End
         Begin VB.Label Label2 
            Caption         =   "unless it contains data from at least one of the following selected layers:"
            Height          =   465
            Left            =   450
            TabIndex        =   22
            Top             =   480
            Width           =   2595
         End
      End
      Begin VB.Frame fraChooseTiles 
         Caption         =   "Choose tiles"
         Height          =   1515
         Left            =   60
         TabIndex        =   16
         Top             =   750
         Width           =   3255
         Begin VB.OptionButton optTiles 
            Caption         =   "Use the visible tiles"
            Height          =   285
            Index           =   2
            Left            =   210
            TabIndex        =   19
            Top             =   1050
            Width           =   1995
         End
         Begin VB.OptionButton optTiles 
            Caption         =   "Use the selected tiles"
            Height          =   285
            Index           =   1
            Left            =   210
            TabIndex        =   18
            Top             =   660
            Width           =   1935
         End
         Begin VB.OptionButton optTiles 
            Caption         =   "Use all of the tiles"
            Height          =   285
            Index           =   0
            Left            =   210
            TabIndex        =   17
            Top             =   270
            Width           =   1575
         End
      End
      Begin MSComctlLib.ListView lvwSheets 
         Height          =   1515
         Left            =   60
         TabIndex        =   6
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
         Left            =   90
         TabIndex        =   44
         Top             =   2610
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   $"frmMapSeriesWiz.frx":0000
         Height          =   615
         Index           =   1
         Left            =   30
         TabIndex        =   15
         Top             =   60
         Width           =   6705
      End
   End
   Begin VB.Frame fraPage1 
      BorderStyle     =   0  'None
      Height          =   3525
      Left            =   90
      TabIndex        =   4
      Top             =   90
      Visible         =   0   'False
      Width           =   6855
      Begin VB.Frame fraIndexLayer 
         Caption         =   "Index Layer"
         Height          =   2415
         Left            =   3210
         TabIndex        =   10
         Top             =   900
         Width           =   3525
         Begin VB.ComboBox cmbIndexField 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1200
            Width           =   3105
         End
         Begin VB.ComboBox cmbIndexLayer 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   510
            Width           =   3105
         End
         Begin VB.Label lblMapSheet 
            Caption         =   "This field specifies the page name"
            Height          =   225
            Index           =   4
            Left            =   240
            TabIndex        =   13
            Top             =   960
            Width           =   2535
         End
         Begin VB.Label lblMapSheet 
            Caption         =   "Choose the index layer:"
            Height          =   225
            Index           =   3
            Left            =   240
            TabIndex        =   11
            Top             =   270
            Width           =   1725
         End
      End
      Begin VB.ComboBox cmbDetailFrame 
         Height          =   315
         Left            =   270
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1140
         Width           =   2625
      End
      Begin VB.Label Label1 
         Caption         =   $"frmMapSeriesWiz.frx":00F8
         Height          =   615
         Index           =   0
         Left            =   60
         TabIndex        =   7
         Top             =   60
         Width           =   6705
      End
      Begin VB.Label lblMapSheet 
         Caption         =   "Choose the detail data frame:"
         Height          =   225
         Index           =   0
         Left            =   30
         TabIndex        =   8
         Top             =   870
         Width           =   2235
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   120
      X2              =   6780
      Y1              =   3580
      Y2              =   3580
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      Index           =   0
      X1              =   120
      X2              =   6780
      Y1              =   3570
      Y2              =   3570
   End
End
Attribute VB_Name = "frmMapSeriesWiz"
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

Private m_iPage As Integer
Public m_pApp As IApplication
Private m_pCurrentFrame As Frame
Private m_pMap As IMap
Private m_pIndexLayer As IFeatureLayer
Private m_bFormLoad As Boolean
Private m_pTextSym As ISimpleTextSymbol

Private Sub PositionFrame(pFrame As Frame)
On Error GoTo ErrHand:

26:   If Not m_pCurrentFrame Is Nothing Then m_pCurrentFrame.Visible = False
27:   pFrame.Visible = True
28:   pFrame.Height = 3495
29:   pFrame.Width = 6825
30:   pFrame.Left = 30
31:   pFrame.Top = 30
32:   Set m_pCurrentFrame = pFrame
33:   pFrame.Visible = True
     
  Exit Sub
ErrHand:
37:   MsgBox "PositionFrame - " & Err.Description
  Exit Sub
End Sub

Private Sub chkOptions_Click(Index As Integer)
  Select Case Index
  Case 0  'Rotate
44:     If chkOptions(0).value = 0 Then
45:       cmbRotateField.Enabled = False
46:     Else
47:       cmbRotateField.Enabled = True
48:     End If
  Case 1  'Clip to outline
50:     If chkOptions(1).value = 0 Then
51:       chkOptions(3).Enabled = False
52:       chkOptions(3).value = 0
53:     Else
54:       chkOptions(3).Enabled = True
55:     End If
  Case 2  'Label neighboring tiles
57:     If chkOptions(2).value = 0 Then
58:       cmdLabelProps.Enabled = False
59:     Else
60:       cmdLabelProps.Enabled = True
61:     End If
  Case 4  'Select tile when drawing - Added 11/23/04
    
64:   End Select
End Sub

Private Sub chkSuppress_Click()
68:   If chkSuppress.value = 0 Then
69:     lstSuppressTiles.Enabled = False
70:   Else
71:     lstSuppressTiles.Enabled = True
72:   End If
End Sub

Private Sub cmbDetailFrame_Click()
On Error GoTo ErrHand:
  Dim pDoc As IMxDocument, lLoop As Long
  Dim pFeatLayer As IFeatureLayer, pGroupLayer As ICompositeLayer
  
  'Set the Next button to false
81:   cmdNext.Enabled = False
  
  'Find the selected map
84:   cmbIndexLayer.Clear
85:   If cmbDetailFrame.Text = "" Then
86:     MsgBox "No detail frame selected!!!"
    Exit Sub
88:   End If
  
90:   Set pDoc = m_pApp.Document
91:   Set m_pMap = FindDataFrame(pDoc, cmbDetailFrame.Text)
92:   If m_pMap Is Nothing Then
93:     MsgBox "Could not find detail frame!!!"
    Exit Sub
95:   End If
  
  'Populate the index layer combo
98:   lstSuppressTiles.Clear
99:   cmbIndexLayer.Clear
100:   For lLoop = 0 To m_pMap.LayerCount - 1
101:     If TypeOf m_pMap.Layer(lLoop) Is ICompositeLayer Then
102:       CompositeLayer m_pMap.Layer(lLoop)
103:     Else
104:       LayerCheck m_pMap.Layer(lLoop)
105:     End If
106:   Next lLoop
107:   If cmbIndexLayer.ListCount = 0 Then
108:     MsgBox "You need at least one polygon layer in the detail frame to serve as the index layer!!!"
109:   Else
110:     cmbIndexLayer.ListIndex = 0
111:   End If
  
  Exit Sub
ErrHand:
115:   MsgBox "cmbDetailFrame_Click - " & Err.Description
End Sub

Private Sub CompositeLayer(pCompLayer As ICompositeLayer)
On Error GoTo ErrHand:
  Dim lLoop As Long
121:   For lLoop = 0 To pCompLayer.count - 1
122:     If TypeOf pCompLayer.Layer(lLoop) Is ICompositeLayer Then
123:       CompositeLayer pCompLayer.Layer(lLoop)
124:     Else
125:       LayerCheck pCompLayer.Layer(lLoop)
126:     End If
127:   Next lLoop

  Exit Sub
ErrHand:
131:   MsgBox "CompositeLayer - " & Err.Description
End Sub

Private Sub LayerCheck(pLayer As ILayer)
On Error GoTo ErrHand:
  Dim pFeatLayer As IFeatureLayer
  
138:   If TypeOf pLayer Is IFeatureLayer Then
139:     Set pFeatLayer = pLayer
140:     If pFeatLayer.FeatureClass.ShapeType = esriGeometryPolygon Then
141:       cmbIndexLayer.AddItem pFeatLayer.Name
142:     End If
143:     lstSuppressTiles.AddItem pFeatLayer.Name
144:   End If

  Exit Sub
ErrHand:
148:   MsgBox "LayerCheck - " & Err.Description
End Sub

Private Sub cmbIndexLayer_Click()
On Error GoTo ErrHand:
  Dim lLoop As Long, pFields As IFields, pField As IField
  
  'Set the Next button to false
156:   cmdNext.Enabled = False
  
  'Find the selected layer
159:   cmbIndexField.Clear
160:   If cmbIndexLayer.Text = "" Then
161:     MsgBox "No index layer selected!!!"
    Exit Sub
163:   End If
  
165:   Set m_pIndexLayer = FindLayer(cmbIndexLayer.Text, m_pMap)
166:   If m_pIndexLayer Is Nothing Then
167:     MsgBox "Could not find specified layer!!!"
    Exit Sub
169:   End If
  
  'Populate the index layer combos
172:   Set pFields = m_pIndexLayer.FeatureClass.Fields
173:   cmbDataDriven.Clear
174:   cmbRotateField.Clear
175:   For lLoop = 0 To pFields.FieldCount - 1
    Select Case pFields.Field(lLoop).Type
    Case esriFieldTypeString
178:       cmbIndexField.AddItem pFields.Field(lLoop).Name
    Case esriFieldTypeDouble, esriFieldTypeSingle, esriFieldTypeInteger
180:       If UCase(pFields.Field(lLoop).Name) <> "SHAPE_LENGTH" And _
       UCase(pFields.Field(lLoop).Name) <> "SHAPE_AREA" Then
182:         cmbDataDriven.AddItem pFields.Field(lLoop).Name
183:         cmbRotateField.AddItem pFields.Field(lLoop).Name
184:       End If
185:     End Select
186:   Next lLoop
187:   If cmbIndexField.ListCount = 0 Then
'    MsgBox "You need at least one string field in the layer for labeling the pages!!!"
189:   Else
190:     cmbIndexField.ListIndex = 0
191:     cmdNext.Enabled = True
192:   End If
193:   If cmbDataDriven.ListCount > 0 Then
194:     cmbDataDriven.ListIndex = 0
195:     cmbRotateField.ListIndex = 0
196:     optExtent.Item(2).Enabled = True
197:     chkOptions(0).Enabled = True
198:   Else
199:     optExtent.Item(2).Enabled = False
200:     chkOptions(0).Enabled = False
201:   End If

  Exit Sub
ErrHand:
205:   MsgBox "cmbIndexField_Click - " & Err.Description
End Sub

Private Sub cmdBack_Click()
209:   m_pCurrentFrame.Visible = False
  Select Case m_iPage
  Case 2
212:     PositionFrame fraPage1
213:     m_iPage = 1
  Case 3
215:     cmdNext.Caption = "Next >"
216:     PositionFrame fraPage2
217:     m_iPage = 2
218:   End Select
219:   cmdNext.Enabled = True
End Sub

Private Sub cmdCancel_Click()
223:   Unload Me
End Sub

Private Sub cmdLabelProps_Click()
On Error GoTo ErrHand:
  Dim bChanged As Boolean, pTextSymEditor As ITextSymbolEditor
229:   Set pTextSymEditor = New TextSymbolEditor
230:   bChanged = pTextSymEditor.EditTextSymbol(m_pTextSym, m_pApp.hwnd)
231:   Me.SetFocus
  
  Exit Sub
ErrHand:
235:   MsgBox "cmdLabelProps_Click - " & Err.Description
End Sub

Private Sub cmdNext_Click()
On Error GoTo ErrHand:
  Dim pMapSeries As IDSMapSeries
241:   m_pCurrentFrame.Visible = False
242:   cmdBack.Enabled = True
  Select Case m_iPage
  Case 1  'Done with date frame and index layer
245:     CheckForSelected    'Check index layer to see if there are selected features
246:     PositionFrame fraPage2
247:     m_iPage = 2
  Case 2  'Done with tile specification
249:     PositionFrame fraPage3
250:     m_iPage = 3
251:     cmdNext.Caption = "Finish"
252:     If optExtent(0).value Then
253:       If txtMargin.Text = "" Then
254:         cmdNext.Enabled = False
255:       Else
256:         cmdNext.Enabled = True
257:       End If
258:     ElseIf optExtent(1).value Then
259:       If txtFixed.Text = "" Then
260:         cmdNext.Enabled = False
261:       Else
262:         cmdNext.Enabled = True
263:       End If
264:     Else
265:       cmdNext.Enabled = True
266:     End If
  Case 3  'Finish button selected
268:     CreateSeries
269:     Unload Me
270:   End Select
  
  Exit Sub
ErrHand:
274:   MsgBox "cmdNext_click - " & Err.Description
  Exit Sub
End Sub

Private Sub CreateSeries()
On Error GoTo ErrHandler:
  Dim pMapSeries As IDSMapSeries, pSpatialQuery As ISpatialFilter
  Dim pTmpPage As tmpPageClass, pTmpColl As Collection, pClone As IClone
  Dim pSeriesOpt As IDSMapSeriesOptions, pFeatLayerSel As IFeatureSelection
  Dim pSeriesProps As IDSMapSeriesProps, pMapPage As IDSMapPage
  Dim pDoc As IMxDocument, pMap As IMap, lCount As Long, lLoop As Long
  Dim pFeatLayer As IFeatureLayer, pQuery As IQueryFilter, pCursor As IFeatureCursor
  Dim pFeature As IFeature, lIndex As Long, sName As String, sFieldName As String
  Dim pNode As Node, pMapBook As IDSMapBook
  Dim pActiveView As IActiveView, lRotIndex As Long, lScaleIndex As Long
  'Added 6/18/03 to support cross hatch outside clip area
  Dim pSeriesOpt2 As IDSMapSeriesOptions2
  Dim pSeriesOpt3 As IDSMapSeriesOptions3    'Added 11/23/04 to support tile selection
  'Add 2/18/04 to support the storing of page numbers
  Dim lPageNumber As Long
  
295:   Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
298:   pMapBook.EnableBook = True
299:   Set pDoc = m_pApp.Document
  
301:   Set pMapSeries = New DSMapSeries
302:   Set pSeriesOpt = pMapSeries
303:   Set pSeriesOpt2 = pSeriesOpt  'Added 6/18/03 to support cross hatch outside clip area
304:   Set pSeriesOpt3 = pSeriesOpt    'Added 11/23/04
305:   Set pSeriesProps = pMapSeries
306:   pMapSeries.EnableSeries = True
  
  'Find the detail frame
309:   Set pMap = FindDataFrame(pDoc, cmbDetailFrame.Text)
310:   If pMap Is Nothing Then
311:     MsgBox "Detail frame not found!!!"
    Exit Sub
313:   End If
314:   pSeriesProps.DataFrameName = pMap.Name
  
  'Find the layer
317:   Set pFeatLayer = FindLayer(cmbIndexLayer.Text, pMap)
318:   If pFeatLayer Is Nothing Then
319:     MsgBox "Index layer not found!!!"
    Exit Sub
321:   End If
322:   pSeriesProps.IndexLayerName = pFeatLayer.Name
323:   pSeriesProps.IndexFieldName = cmbIndexField.Text
    
  'Determine the tiles we are interested in
326:   Set pQuery = New QueryFilter
327:   sFieldName = cmbIndexField.Text
328:   pQuery.AddField sFieldName
'  pQuery.WhereClause = sFieldName & " <> ''"
330:   pQuery.WhereClause = sFieldName & " is not null"
331:   If optTiles(0).value Then
332:     Set pCursor = pFeatLayer.Search(pQuery, True)
333:     pSeriesProps.TileSelectionMethod = 0
334:   ElseIf optTiles(1).value Then
335:     Set pFeatLayerSel = pFeatLayer
336:     pFeatLayerSel.SelectionSet.Search pQuery, True, pCursor
337:     pSeriesProps.TileSelectionMethod = 1
338:   Else
339:     Set pActiveView = pMap
340:     Set pSpatialQuery = New SpatialFilter
341:     pSpatialQuery.AddField sFieldName
342:     pSpatialQuery.SpatialRel = esriSpatialRelIntersects
343:     Set pSpatialQuery.Geometry = pActiveView.Extent
344:     pSpatialQuery.WhereClause = sFieldName & " <> ''"
345:     pSpatialQuery.GeometryField = pFeatLayer.FeatureClass.shapeFieldName
346:     Set pCursor = pFeatLayer.Search(pSpatialQuery, True)
347:     pSeriesProps.TileSelectionMethod = 2
348:   End If
  
  'Add 2/18/04 to keep track of starting page number
351:   pSeriesProps.StartNumber = CLng(txtNumbering.Text)
  
  'Set the clip, label and rotate properties
  'Updated 6/18/03 to support cross hatch outside clip area
355:   If chkOptions(1).value = 1 Then
356:     If chkOptions(3).value = 1 Then
357:       pSeriesOpt2.ClipData = 2
358:     Else
359:       pSeriesOpt2.ClipData = 1
360:     End If
361:   Else
362:     pSeriesOpt2.ClipData = 0
363:   End If
'  If chkOptions(1).Value = 1 Then
'    pSeriesOpt.ClipData = True
'  Else
'    pSeriesOpt.ClipData = False
'  End If
  
370:   If chkOptions(0).value = 1 Then
371:     pSeriesOpt.RotateFrame = True
372:     pSeriesOpt.RotationField = cmbRotateField.Text
373:     lRotIndex = pFeatLayer.FeatureClass.FindField(cmbRotateField.Text)
374:   Else
375:     pSeriesOpt.RotateFrame = False
376:   End If
377:   If chkOptions(2).value = 1 Then
378:     pSeriesOpt.LabelNeighbors = True
379:   Else
380:     pSeriesOpt.LabelNeighbors = False
381:   End If
382:   Set pSeriesOpt.LabelSymbol = m_pTextSym
  
  'Set selection tile drawing property - Added 11/23/04
385:   If chkOptions(4).value = 1 Then
386:     pSeriesOpt3.SelectTile = True
387:   Else
388:     pSeriesOpt3.SelectTile = False
389:   End If
  
  'Set the extent properties
392:   If optExtent(0).value Then         'Variable
393:     pSeriesOpt.ExtentType = 0
394:     If txtMargin.Text = "" Then
395:       pSeriesOpt.Margin = 0
396:     Else
397:       pSeriesOpt.Margin = CDbl(txtMargin.Text)
398:     End If
399:     pSeriesOpt.MarginType = cmbMargin.ListIndex
400:   ElseIf optExtent(1).value Then    'Fixed
401:     pSeriesOpt.ExtentType = 1
402:     pSeriesOpt.FixedScale = txtFixed.Text
403:   Else                        'Data driven
404:     pSeriesOpt.ExtentType = 2
405:     pSeriesOpt.DataDrivenField = cmbDataDriven.Text
406:     lScaleIndex = pFeatLayer.FeatureClass.FindField(cmbDataDriven.Text)
407:   End If
  
  'Store suppression information
410:   If chkSuppress.value = 1 And lstSuppressTiles.SelCount > 0 Then
411:     pSeriesProps.SuppressLayers = True
412:     For lLoop = 0 To lstSuppressTiles.ListCount - 1
413:       If lstSuppressTiles.Selected(lLoop) Then
414:         pSeriesProps.AddLayerToSuppress lstSuppressTiles.List(lLoop)
415:       End If
416:     Next lLoop
417:   Else
418:     pSeriesProps.SuppressLayers = False
419:   End If
  
  'Create the pages and populate the treeview
422:   Set pTmpColl = New Collection
423:   lIndex = pFeatLayer.FeatureClass.FindField(sFieldName)
424:   Set pFeature = pCursor.NextFeature
425:   With g_pFrmMapSeries.tvwMapBook
426:     Set pNode = .Nodes.Add("MapBook", tvwChild, "MapSeries", "Map Series", 3)
427:     pNode.Tag = "MapSeries"
    
    'Add tile names to a listbox first for sort purposes
430:     g_pFrmMapSeries.lstSorter.Clear
431:     Do While Not pFeature Is Nothing
432:       sName = pFeature.value(lIndex)
433:       Set pTmpPage = New tmpPageClass
434:       pTmpPage.PageName = sName
435:       pTmpPage.PageRotation = 0
436:       pTmpPage.PageScale = 1
437:       Set pClone = pFeature.Shape
438:       Set pTmpPage.PageShape = pClone.Clone
      'Track the rotation and scale values (if we are going to use them) to the end
      'of the name, so we can assign them to the page when it is added without having
      'to query the index layer again.
442:       If chkOptions(0).value = 1 And lRotIndex >= 0 Then
443:         If Not IsNull(pFeature.value(lRotIndex)) Then
444:           pTmpPage.PageRotation = pFeature.value(lRotIndex)
445:         End If
446:       End If
447:       If optExtent(2).value And lScaleIndex >= 0 Then
448:         If Not IsNull(pFeature.value(lScaleIndex)) Then
449:           pTmpPage.PageScale = pFeature.value(lScaleIndex)
450:         End If
451:       End If
452:       If chkSuppress.value = 1 And lstSuppressTiles.SelCount > 0 Then
453:         If FeaturesInTile(pFeature, pMap) Then
454:           g_pFrmMapSeries.lstSorter.AddItem sName
455:           pTmpColl.Add pTmpPage, sName
456:         End If
457:       Else
458:         g_pFrmMapSeries.lstSorter.AddItem sName
459:         pTmpColl.Add pTmpPage, sName
460:       End If
461:       Set pFeature = pCursor.NextFeature
462:     Loop
    
    'Now loop back through the list and add the tile names as nodes in the tree
465:     For lLoop = 0 To g_pFrmMapSeries.lstSorter.ListCount - 1
466:       Set pMapPage = New DSMapPage
467:       lPageNumber = lLoop + CLng(txtNumbering.Text)
468:       sName = g_pFrmMapSeries.lstSorter.List(lLoop)
469:       Set pNode = .Nodes.Add("MapSeries", tvwChild, "a" & sName, lPageNumber & " - " & sName, 5)
470:       Set pTmpPage = pTmpColl.Item(sName)
471:       pNode.Tag = lLoop
472:       pMapPage.PageName = sName
473:       pMapPage.PageRotation = pTmpPage.PageRotation
474:       pMapPage.PageScale = pTmpPage.PageScale
475:       Set pMapPage.PageShape = pTmpPage.PageShape
476:       pMapPage.LastOutputted = #1/1/1900#
477:       pMapPage.EnablePage = True
478:       pMapPage.PageNumber = lPageNumber
479:       pMapSeries.AddPage pMapPage
480:     Next lLoop
481:     .Nodes.Item("MapBook").Expanded = True
482:     .Nodes.Item("MapSeries").Expanded = True
483:   End With
  
  'Add the series to the book
486:   pMapBook.AddContent pMapSeries

  Exit Sub
ErrHandler:
490:   MsgBox "CreateSeries - most likely you do not have unique names in your index layer!!!"
End Sub

Private Sub CheckForSelected()
On Error GoTo ErrHand:
  Dim pFeatSel As IFeatureSelection
  
  'Make sure there is something to check
498:   optTiles(1).Enabled = False
  If m_pIndexLayer Is Nothing Then Exit Sub
  
  'Check for selected features in the index layer
502:   Set pFeatSel = m_pIndexLayer
503:   If pFeatSel.SelectionSet.count <> 0 Then
504:     optTiles(1).Enabled = True
505:   End If

  Exit Sub
ErrHand:
509:   MsgBox "CheckForSelected - " & Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo ErrHand:
  Dim pDoc As IMxDocument, lLoop As Long
  'Get the extension
  If m_pApp Is Nothing Then Exit Sub
    
518:   m_bFormLoad = True
519:   Set m_pCurrentFrame = Nothing
520:   PositionFrame fraPage1
521:   cmdNext.Enabled = False
522:   cmdBack.Enabled = False
  
  'Initialize variables and controls
525:   m_iPage = 1
526:   chkOptions(0).value = 0
527:   chkOptions(1).value = 0
528:   chkOptions(2).value = 0
529:   chkOptions(4).value = 0
530:   chkSuppress.value = 0
531:   optTiles(0).value = True
532:   optExtent(0).value = True
533:   lstSuppressTiles.Enabled = False
534:   cmbRotateField.Enabled = False
535:   cmdLabelProps.Enabled = False
536:   chkOptions(3).Enabled = False
537:   txtNumbering.Text = "1"
  
  'Populate the data frame combo
540:   Set pDoc = m_pApp.Document
541:   cmbIndexField.Clear
542:   cmbDetailFrame.Clear
543:   For lLoop = 0 To pDoc.Maps.count - 1
544:     cmbDetailFrame.AddItem pDoc.Maps.Item(lLoop).Name
545:   Next lLoop
546:   cmbDetailFrame.ListIndex = 0
547:   m_bFormLoad = False
  
  'Populate the extent options
550:   cmbMargin.Clear
551:   cmbMargin.AddItem "percent"
552:   cmbMargin.AddItem "mapunits"
553:   cmbMargin.Text = "percent"
554:   txtMargin.Text = "0"
  
  'Set the initial Label symbol
557:   Set pDoc = m_pApp.Document
558:   Set m_pTextSym = New TextSymbol
559:   m_pTextSym.Font = pDoc.DefaultTextFont
560:   m_pTextSym.Size = pDoc.DefaultTextFontSize.Size
  
  'Make sure the wizard stays on top
563:   TopMost Me

  Exit Sub
  
ErrHand:
568:   MsgBox "frmMapSheetWiz Load - " & Err.Description
  Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
573:   Set m_pApp = Nothing
574:   Set m_pCurrentFrame = Nothing
575:   Set m_pMap = Nothing
576:   Set m_pIndexLayer = Nothing
End Sub

Private Sub optExtent_Click(Index As Integer)
On Error GoTo ErrHand:
  Select Case Index
  Case 0  'Variable
583:     txtMargin.Enabled = True
584:     cmbMargin.Enabled = True
585:     txtFixed.Enabled = False
586:     cmbDataDriven.Enabled = False
587:     If txtMargin.Text = "" Then
588:       cmdNext.Enabled = False
589:     Else
590:       cmdNext.Enabled = True
591:     End If
  Case 1  'Fixed
593:     txtMargin.Enabled = False
594:     cmbMargin.Enabled = False
595:     txtFixed.Enabled = True
596:     cmbDataDriven.Enabled = False
597:     If txtFixed.Text = "" Then
598:       cmdNext.Enabled = False
599:     Else
600:       cmdNext.Enabled = True
601:     End If
  Case 2  'Data driven
603:     txtMargin.Enabled = False
604:     cmbMargin.Enabled = False
605:     txtFixed.Enabled = False
606:     cmbDataDriven.Enabled = True
607:     cmdNext.Enabled = True
608:   End Select

  Exit Sub
ErrHand:
612:   MsgBox "optExtent_Click - " & Err.Description
End Sub

Private Sub txtFixed_KeyUp(KeyCode As Integer, Shift As Integer)
616:   If Not IsNumeric(txtFixed.Text) Then
617:     txtFixed.Text = ""
618:   End If
619:   If txtFixed.Text <> "" Then
620:     cmdNext.Enabled = True
621:   End If
End Sub

Private Sub txtMargin_KeyUp(KeyCode As Integer, Shift As Integer)
625:   If Not IsNumeric(txtMargin.Text) Then
626:     txtMargin.Text = ""
627:   End If
628:   If txtMargin.Text <> "" Then
629:     cmdNext.Enabled = True
630:   End If
End Sub

Private Function FeaturesInTile(pFeature As IFeature, pMap As IMap) As Boolean
'Routine for determining whether the specified tile feature (pFeature) should
'be suppressed.  Tiles are suppressed when there are no features from the checked
'layers in them.
On Error GoTo ErrHand:
  Dim lLoop As Long, pFeatLayer As IFeatureLayer, pSpatial As ISpatialFilter
  Dim pFeatCursor As IFeatureCursor, pSearchFeat As IFeature
  
641:   FeaturesInTile = False
  
643:   Set pSpatial = New SpatialFilter
644:   pSpatial.SpatialRel = esriSpatialRelIntersects
645:   Set pSpatial.Geometry = pFeature.Shape
646:   For lLoop = 0 To lstSuppressTiles.ListCount - 1
647:     If lstSuppressTiles.Selected(lLoop) Then
648:       Set pFeatLayer = FindLayer(lstSuppressTiles.List(lLoop), pMap)
649:       pSpatial.GeometryField = pFeatLayer.FeatureClass.shapeFieldName
650:       Set pFeatCursor = pFeatLayer.Search(pSpatial, True)
651:       Set pSearchFeat = pFeatCursor.NextFeature
652:       If Not pSearchFeat Is Nothing Then
653:         FeaturesInTile = True
        Exit Function
655:       End If
656:     End If
657:   Next lLoop

  Exit Function
  
ErrHand:
662:   MsgBox "FeaturesInTile - " & Err.Description
End Function

Private Sub txtNumbering_KeyUp(KeyCode As Integer, Shift As Integer)
666:   If Not IsNumeric(txtNumbering.Text) Then
667:     txtNumbering.Text = "1"
'  ElseIf CInt(txtNumbering.Text) < 0 Then
'    MsgBox "Can not use a number less than 0!!!"
'    txtNumbering.Text = "1"
671:   End If
End Sub
