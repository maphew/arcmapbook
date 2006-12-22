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

42:   If Not m_pCurrentFrame Is Nothing Then m_pCurrentFrame.Visible = False
43:   pFrame.Visible = True
44:   pFrame.Height = 3495
45:   pFrame.Width = 6825
46:   pFrame.Left = 30
47:   pFrame.Top = 30
48:   Set m_pCurrentFrame = pFrame
49:   pFrame.Visible = True
     
  Exit Sub
ErrHand:
53:   MsgBox "PositionFrame - " & Err.Description
  Exit Sub
End Sub

Private Sub chkOptions_Click(Index As Integer)
  Select Case Index
  Case 0  'Rotate
60:     If chkOptions(0).Value = 0 Then
61:       cmbRotateField.Enabled = False
62:     Else
63:       cmbRotateField.Enabled = True
64:     End If
  Case 1  'Clip to outline
66:     If chkOptions(1).Value = 0 Then
67:       chkOptions(3).Enabled = False
68:       chkOptions(3).Value = 0
69:     Else
70:       chkOptions(3).Enabled = True
71:     End If
  Case 2  'Label neighboring tiles
73:     If chkOptions(2).Value = 0 Then
74:       cmdLabelProps.Enabled = False
75:     Else
76:       cmdLabelProps.Enabled = True
77:     End If
  Case 4  'Select tile when drawing - Added 11/23/04
    
80:   End Select
End Sub

Private Sub chkSuppress_Click()
84:   If chkSuppress.Value = 0 Then
85:     lstSuppressTiles.Enabled = False
86:   Else
87:     lstSuppressTiles.Enabled = True
88:   End If
End Sub

Private Sub cmbDetailFrame_Click()
On Error GoTo ErrHand:
  Dim pDoc As IMxDocument, lLoop As Long
  Dim pFeatLayer As IFeatureLayer, pGroupLayer As ICompositeLayer
  
  'Set the Next button to false
97:   cmdNext.Enabled = False
  
  'Find the selected map
100:   cmbIndexLayer.Clear
101:   If cmbDetailFrame.Text = "" Then
102:     MsgBox "No detail frame selected!!!"
    Exit Sub
104:   End If
  
106:   Set pDoc = m_pApp.Document
107:   Set m_pMap = FindDataFrame(pDoc, cmbDetailFrame.Text)
108:   If m_pMap Is Nothing Then
109:     MsgBox "Could not find detail frame!!!"
    Exit Sub
111:   End If
  
  'Populate the index layer combo
114:   lstSuppressTiles.Clear
115:   cmbIndexLayer.Clear
116:   For lLoop = 0 To m_pMap.LayerCount - 1
117:     If TypeOf m_pMap.Layer(lLoop) Is ICompositeLayer Then
118:       CompositeLayer m_pMap.Layer(lLoop)
119:     Else
120:       LayerCheck m_pMap.Layer(lLoop)
121:     End If
122:   Next lLoop
123:   If cmbIndexLayer.ListCount = 0 Then
124:     MsgBox "You need at least one polygon layer in the detail frame to serve as the index layer!!!"
125:   Else
126:     cmbIndexLayer.ListIndex = 0
127:   End If
  
  Exit Sub
ErrHand:
131:   MsgBox "cmbDetailFrame_Click - " & Err.Description
End Sub

Private Sub CompositeLayer(pCompLayer As ICompositeLayer)
On Error GoTo ErrHand:
  Dim lLoop As Long
137:   For lLoop = 0 To pCompLayer.count - 1
138:     If TypeOf pCompLayer.Layer(lLoop) Is ICompositeLayer Then
139:       CompositeLayer pCompLayer.Layer(lLoop)
140:     Else
141:       LayerCheck pCompLayer.Layer(lLoop)
142:     End If
143:   Next lLoop

  Exit Sub
ErrHand:
147:   MsgBox "CompositeLayer - " & Err.Description
End Sub

Private Sub LayerCheck(pLayer As ILayer)
On Error GoTo ErrHand:
  Dim pFeatLayer As IFeatureLayer
  
154:   If TypeOf pLayer Is IFeatureLayer Then
155:     Set pFeatLayer = pLayer
156:     If pFeatLayer.FeatureClass.ShapeType = esriGeometryPolygon Then
157:       cmbIndexLayer.AddItem pFeatLayer.Name
158:     End If
159:     lstSuppressTiles.AddItem pFeatLayer.Name
160:   End If

  Exit Sub
ErrHand:
164:   MsgBox "LayerCheck - " & Err.Description
End Sub

Private Sub cmbIndexLayer_Click()
On Error GoTo ErrHand:
  Dim lLoop As Long, pFields As IFields, pField As IField
  
  'Set the Next button to false
172:   cmdNext.Enabled = False
  
  'Find the selected layer
175:   cmbIndexField.Clear
176:   If cmbIndexLayer.Text = "" Then
177:     MsgBox "No index layer selected!!!"
    Exit Sub
179:   End If
  
181:   Set m_pIndexLayer = FindLayer(cmbIndexLayer.Text, m_pMap)
182:   If m_pIndexLayer Is Nothing Then
183:     MsgBox "Could not find specified layer!!!"
    Exit Sub
185:   End If
  
  'Populate the index layer combos
188:   Set pFields = m_pIndexLayer.FeatureClass.Fields
189:   cmbDataDriven.Clear
190:   cmbRotateField.Clear
191:   For lLoop = 0 To pFields.FieldCount - 1
    Select Case pFields.Field(lLoop).Type
    Case esriFieldTypeString
194:       cmbIndexField.AddItem pFields.Field(lLoop).Name
    Case esriFieldTypeDouble, esriFieldTypeSingle, esriFieldTypeInteger
196:       If UCase(pFields.Field(lLoop).Name) <> "SHAPE_LENGTH" And _
       UCase(pFields.Field(lLoop).Name) <> "SHAPE_AREA" Then
198:         cmbDataDriven.AddItem pFields.Field(lLoop).Name
199:         cmbRotateField.AddItem pFields.Field(lLoop).Name
200:       End If
201:     End Select
202:   Next lLoop
203:   If cmbIndexField.ListCount = 0 Then
'    MsgBox "You need at least one string field in the layer for labeling the pages!!!"
205:   Else
206:     cmbIndexField.ListIndex = 0
207:     cmdNext.Enabled = True
208:   End If
209:   If cmbDataDriven.ListCount > 0 Then
210:     cmbDataDriven.ListIndex = 0
211:     cmbRotateField.ListIndex = 0
212:     optExtent.Item(2).Enabled = True
213:     chkOptions(0).Enabled = True
214:   Else
215:     optExtent.Item(2).Enabled = False
216:     chkOptions(0).Enabled = False
217:   End If

  Exit Sub
ErrHand:
221:   MsgBox "cmbIndexField_Click - " & Err.Description
End Sub

Private Sub cmdBack_Click()
225:   m_pCurrentFrame.Visible = False
  Select Case m_iPage
  Case 2
228:     PositionFrame fraPage1
229:     m_iPage = 1
  Case 3
231:     cmdNext.Caption = "Next >"
232:     PositionFrame fraPage2
233:     m_iPage = 2
234:   End Select
235:   cmdNext.Enabled = True
End Sub

Private Sub cmdCancel_Click()
239:   Unload Me
End Sub

Private Sub cmdLabelProps_Click()
On Error GoTo ErrHand:
  Dim bChanged As Boolean, pTextSymEditor As ITextSymbolEditor
245:   Set pTextSymEditor = New TextSymbolEditor
246:   bChanged = pTextSymEditor.EditTextSymbol(m_pTextSym, m_pApp.hwnd)
247:   Me.SetFocus
  
  Exit Sub
ErrHand:
251:   MsgBox "cmdLabelProps_Click - " & Err.Description
End Sub

Private Sub cmdNext_Click()
On Error GoTo ErrHand:
  Dim pMapSeries As IDSMapSeries
257:   m_pCurrentFrame.Visible = False
258:   cmdBack.Enabled = True
  Select Case m_iPage
  Case 1  'Done with date frame and index layer
261:     CheckForSelected    'Check index layer to see if there are selected features
262:     PositionFrame fraPage2
263:     m_iPage = 2
  Case 2  'Done with tile specification
265:     PositionFrame fraPage3
266:     m_iPage = 3
267:     cmdNext.Caption = "Finish"
268:     If optExtent(0).Value Then
269:       If txtMargin.Text = "" Then
270:         cmdNext.Enabled = False
271:       Else
272:         cmdNext.Enabled = True
273:       End If
274:     ElseIf optExtent(1).Value Then
275:       If txtFixed.Text = "" Then
276:         cmdNext.Enabled = False
277:       Else
278:         cmdNext.Enabled = True
279:       End If
280:     Else
281:       cmdNext.Enabled = True
282:     End If
  Case 3  'Finish button selected
284:     CreateSeries
285:     Unload Me
286:   End Select
  
  Exit Sub
ErrHand:
290:   MsgBox "cmdNext_click - " & Err.Description
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
  
311:   Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
314:   pMapBook.EnableBook = True
315:   Set pDoc = m_pApp.Document
  
317:   Set pMapSeries = New DSMapSeries
318:   Set pSeriesOpt = pMapSeries
319:   Set pSeriesOpt2 = pSeriesOpt  'Added 6/18/03 to support cross hatch outside clip area
320:   Set pSeriesOpt3 = pSeriesOpt    'Added 11/23/04
321:   Set pSeriesProps = pMapSeries
322:   pMapSeries.EnableSeries = True
  
  'Find the detail frame
325:   Set pMap = FindDataFrame(pDoc, cmbDetailFrame.Text)
326:   If pMap Is Nothing Then
327:     MsgBox "Detail frame not found!!!"
    Exit Sub
329:   End If
330:   pSeriesProps.DataFrameName = pMap.Name
  
  'Find the layer
333:   Set pFeatLayer = FindLayer(cmbIndexLayer.Text, pMap)
334:   If pFeatLayer Is Nothing Then
335:     MsgBox "Index layer not found!!!"
    Exit Sub
337:   End If
338:   pSeriesProps.IndexLayerName = pFeatLayer.Name
339:   pSeriesProps.IndexFieldName = cmbIndexField.Text
    
  'Determine the tiles we are interested in
342:   Set pQuery = New QueryFilter
343:   sFieldName = cmbIndexField.Text
344:   pQuery.AddField sFieldName
'  pQuery.WhereClause = sFieldName & " <> ''"
346:   pQuery.WhereClause = sFieldName & " is not null"
347:   If optTiles(0).Value Then
348:     Set pCursor = pFeatLayer.Search(pQuery, True)
349:     pSeriesProps.TileSelectionMethod = 0
350:   ElseIf optTiles(1).Value Then
351:     Set pFeatLayerSel = pFeatLayer
352:     pFeatLayerSel.SelectionSet.Search pQuery, True, pCursor
353:     pSeriesProps.TileSelectionMethod = 1
354:   Else
355:     Set pActiveView = pMap
356:     Set pSpatialQuery = New SpatialFilter
357:     pSpatialQuery.AddField sFieldName
358:     pSpatialQuery.SpatialRel = esriSpatialRelIntersects
359:     Set pSpatialQuery.Geometry = pActiveView.Extent
360:     pSpatialQuery.WhereClause = sFieldName & " <> ''"
361:     pSpatialQuery.GeometryField = pFeatLayer.FeatureClass.shapeFieldName
362:     Set pCursor = pFeatLayer.Search(pSpatialQuery, True)
363:     pSeriesProps.TileSelectionMethod = 2
364:   End If
  
  'Add 2/18/04 to keep track of starting page number
367:   pSeriesProps.StartNumber = CLng(txtNumbering.Text)
  
  'Set the clip, label and rotate properties
  'Updated 6/18/03 to support cross hatch outside clip area
371:   If chkOptions(1).Value = 1 Then
372:     If chkOptions(3).Value = 1 Then
373:       pSeriesOpt2.ClipData = 2
374:     Else
375:       pSeriesOpt2.ClipData = 1
376:     End If
377:   Else
378:     pSeriesOpt2.ClipData = 0
379:   End If
'  If chkOptions(1).Value = 1 Then
'    pSeriesOpt.ClipData = True
'  Else
'    pSeriesOpt.ClipData = False
'  End If
  
386:   If chkOptions(0).Value = 1 Then
387:     pSeriesOpt.RotateFrame = True
388:     pSeriesOpt.RotationField = cmbRotateField.Text
389:     lRotIndex = pFeatLayer.FeatureClass.FindField(cmbRotateField.Text)
390:   Else
391:     pSeriesOpt.RotateFrame = False
392:   End If
393:   If chkOptions(2).Value = 1 Then
394:     pSeriesOpt.LabelNeighbors = True
395:   Else
396:     pSeriesOpt.LabelNeighbors = False
397:   End If
398:   Set pSeriesOpt.LabelSymbol = m_pTextSym
  
  'Set selection tile drawing property - Added 11/23/04
401:   If chkOptions(4).Value = 1 Then
402:     pSeriesOpt3.SelectTile = True
403:   Else
404:     pSeriesOpt3.SelectTile = False
405:   End If
  
  'Set the extent properties
408:   If optExtent(0).Value Then         'Variable
409:     pSeriesOpt.ExtentType = 0
410:     If txtMargin.Text = "" Then
411:       pSeriesOpt.Margin = 0
412:     Else
413:       pSeriesOpt.Margin = CDbl(txtMargin.Text)
414:     End If
415:     pSeriesOpt.MarginType = cmbMargin.ListIndex
416:   ElseIf optExtent(1).Value Then    'Fixed
417:     pSeriesOpt.ExtentType = 1
418:     pSeriesOpt.FixedScale = txtFixed.Text
419:   Else                        'Data driven
420:     pSeriesOpt.ExtentType = 2
421:     pSeriesOpt.DataDrivenField = cmbDataDriven.Text
422:     lScaleIndex = pFeatLayer.FeatureClass.FindField(cmbDataDriven.Text)
423:   End If
  
  'Store suppression information
426:   If chkSuppress.Value = 1 And lstSuppressTiles.SelCount > 0 Then
427:     pSeriesProps.SuppressLayers = True
428:     For lLoop = 0 To lstSuppressTiles.ListCount - 1
429:       If lstSuppressTiles.Selected(lLoop) Then
430:         pSeriesProps.AddLayerToSuppress lstSuppressTiles.List(lLoop)
431:       End If
432:     Next lLoop
433:   Else
434:     pSeriesProps.SuppressLayers = False
435:   End If
  
  'Create the pages and populate the treeview
438:   Set pTmpColl = New Collection
439:   lIndex = pFeatLayer.FeatureClass.FindField(sFieldName)
440:   Set pFeature = pCursor.NextFeature
441:   With g_pFrmMapSeries.tvwMapBook
442:     Set pNode = .Nodes.Add("MapBook", tvwChild, "MapSeries", "Map Series", 3)
443:     pNode.Tag = "MapSeries"
    
    'Add tile names to a listbox first for sort purposes
446:     g_pFrmMapSeries.lstSorter.Clear
447:     Do While Not pFeature Is Nothing
448:       sName = pFeature.Value(lIndex)
449:       Set pTmpPage = New tmpPageClass
450:       pTmpPage.PageName = sName
451:       pTmpPage.PageRotation = 0
452:       pTmpPage.PageScale = 1
453:       Set pClone = pFeature.Shape
454:       Set pTmpPage.PageShape = pClone.Clone
      'Track the rotation and scale values (if we are going to use them) to the end
      'of the name, so we can assign them to the page when it is added without having
      'to query the index layer again.
458:       If chkOptions(0).Value = 1 And lRotIndex >= 0 Then
459:         If Not IsNull(pFeature.Value(lRotIndex)) Then
460:           pTmpPage.PageRotation = pFeature.Value(lRotIndex)
461:         End If
462:       End If
463:       If optExtent(2).Value And lScaleIndex >= 0 Then
464:         If Not IsNull(pFeature.Value(lScaleIndex)) Then
465:           pTmpPage.PageScale = pFeature.Value(lScaleIndex)
466:         End If
467:       End If
468:       If chkSuppress.Value = 1 And lstSuppressTiles.SelCount > 0 Then
469:         If FeaturesInTile(pFeature, pMap) Then
470:           g_pFrmMapSeries.lstSorter.AddItem sName
471:           pTmpColl.Add pTmpPage, sName
472:         End If
473:       Else
474:         g_pFrmMapSeries.lstSorter.AddItem sName
475:         pTmpColl.Add pTmpPage, sName
476:       End If
477:       Set pFeature = pCursor.NextFeature
478:     Loop
    
    'Now loop back through the list and add the tile names as nodes in the tree
481:     For lLoop = 0 To g_pFrmMapSeries.lstSorter.ListCount - 1
482:       Set pMapPage = New DSMapPage
483:       lPageNumber = lLoop + CLng(txtNumbering.Text)
484:       sName = g_pFrmMapSeries.lstSorter.List(lLoop)
485:       Set pNode = .Nodes.Add("MapSeries", tvwChild, "a" & sName, lPageNumber & " - " & sName, 5)
486:       Set pTmpPage = pTmpColl.Item(sName)
487:       pNode.Tag = lLoop
488:       pMapPage.PageName = sName
489:       pMapPage.PageRotation = pTmpPage.PageRotation
490:       pMapPage.PageScale = pTmpPage.PageScale
491:       Set pMapPage.PageShape = pTmpPage.PageShape
492:       pMapPage.LastOutputted = #1/1/1900#
493:       pMapPage.EnablePage = True
494:       pMapPage.PageNumber = lPageNumber
495:       pMapSeries.AddPage pMapPage
496:     Next lLoop
497:     .Nodes.Item("MapBook").Expanded = True
498:     .Nodes.Item("MapSeries").Expanded = True
499:   End With
  
  'Add the series to the book
502:   pMapBook.AddContent pMapSeries

  Exit Sub
ErrHandler:
506:   MsgBox "CreateSeries - most likely you do not have unique names in your index layer!!!"
End Sub

Private Sub CheckForSelected()
On Error GoTo ErrHand:
  Dim pFeatSel As IFeatureSelection
  
  'Make sure there is something to check
514:   optTiles(1).Enabled = False
  If m_pIndexLayer Is Nothing Then Exit Sub
  
  'Check for selected features in the index layer
518:   Set pFeatSel = m_pIndexLayer
519:   If pFeatSel.SelectionSet.count <> 0 Then
520:     optTiles(1).Enabled = True
521:   End If

  Exit Sub
ErrHand:
525:   MsgBox "CheckForSelected - " & Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo ErrHand:
  Dim pDoc As IMxDocument, lLoop As Long
  'Get the extension
  If m_pApp Is Nothing Then Exit Sub
    
534:   m_bFormLoad = True
535:   Set m_pCurrentFrame = Nothing
536:   PositionFrame fraPage1
537:   cmdNext.Enabled = False
538:   cmdBack.Enabled = False
  
  'Initialize variables and controls
541:   m_iPage = 1
542:   chkOptions(0).Value = 0
543:   chkOptions(1).Value = 0
544:   chkOptions(2).Value = 0
545:   chkOptions(4).Value = 0
546:   chkSuppress.Value = 0
547:   optTiles(0).Value = True
548:   optExtent(0).Value = True
549:   lstSuppressTiles.Enabled = False
550:   cmbRotateField.Enabled = False
551:   cmdLabelProps.Enabled = False
552:   chkOptions(3).Enabled = False
553:   txtNumbering.Text = "1"
  
  'Populate the data frame combo
556:   Set pDoc = m_pApp.Document
557:   cmbIndexField.Clear
558:   cmbDetailFrame.Clear
559:   For lLoop = 0 To pDoc.Maps.count - 1
560:     cmbDetailFrame.AddItem pDoc.Maps.Item(lLoop).Name
561:   Next lLoop
562:   cmbDetailFrame.ListIndex = 0
563:   m_bFormLoad = False
  
  'Populate the extent options
566:   cmbMargin.Clear
567:   cmbMargin.AddItem "percent"
568:   cmbMargin.AddItem "mapunits"
569:   cmbMargin.Text = "percent"
570:   txtMargin.Text = "0"
  
  'Set the initial Label symbol
573:   Set pDoc = m_pApp.Document
574:   Set m_pTextSym = New TextSymbol
575:   m_pTextSym.Font = pDoc.DefaultTextFont
576:   m_pTextSym.Size = pDoc.DefaultTextFontSize.Size
  
  'Make sure the wizard stays on top
579:   TopMost Me

  Exit Sub
  
ErrHand:
584:   MsgBox "frmMapSheetWiz Load - " & Err.Description
  Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
589:   Set m_pApp = Nothing
590:   Set m_pCurrentFrame = Nothing
591:   Set m_pMap = Nothing
592:   Set m_pIndexLayer = Nothing
End Sub

Private Sub optExtent_Click(Index As Integer)
On Error GoTo ErrHand:
  Select Case Index
  Case 0  'Variable
599:     txtMargin.Enabled = True
600:     cmbMargin.Enabled = True
601:     txtFixed.Enabled = False
602:     cmbDataDriven.Enabled = False
603:     If txtMargin.Text = "" Then
604:       cmdNext.Enabled = False
605:     Else
606:       cmdNext.Enabled = True
607:     End If
  Case 1  'Fixed
609:     txtMargin.Enabled = False
610:     cmbMargin.Enabled = False
611:     txtFixed.Enabled = True
612:     cmbDataDriven.Enabled = False
613:     If txtFixed.Text = "" Then
614:       cmdNext.Enabled = False
615:     Else
616:       cmdNext.Enabled = True
617:     End If
  Case 2  'Data driven
619:     txtMargin.Enabled = False
620:     cmbMargin.Enabled = False
621:     txtFixed.Enabled = False
622:     cmbDataDriven.Enabled = True
623:     cmdNext.Enabled = True
624:   End Select

  Exit Sub
ErrHand:
628:   MsgBox "optExtent_Click - " & Err.Description
End Sub

Private Sub txtFixed_KeyUp(KeyCode As Integer, Shift As Integer)
632:   If Not IsNumeric(txtFixed.Text) Then
633:     txtFixed.Text = ""
634:   End If
635:   If txtFixed.Text <> "" Then
636:     cmdNext.Enabled = True
637:   End If
End Sub

Private Sub txtMargin_KeyUp(KeyCode As Integer, Shift As Integer)
641:   If Not IsNumeric(txtMargin.Text) Then
642:     txtMargin.Text = ""
643:   End If
644:   If txtMargin.Text <> "" Then
645:     cmdNext.Enabled = True
646:   End If
End Sub

Private Function FeaturesInTile(pFeature As IFeature, pMap As IMap) As Boolean
'Routine for determining whether the specified tile feature (pFeature) should
'be suppressed.  Tiles are suppressed when there are no features from the checked
'layers in them.
On Error GoTo ErrHand:
  Dim lLoop As Long, pFeatLayer As IFeatureLayer, pSpatial As ISpatialFilter
  Dim pFeatCursor As IFeatureCursor, pSearchFeat As IFeature
  
657:   FeaturesInTile = False
  
659:   Set pSpatial = New SpatialFilter
660:   pSpatial.SpatialRel = esriSpatialRelIntersects
661:   Set pSpatial.Geometry = pFeature.Shape
662:   For lLoop = 0 To lstSuppressTiles.ListCount - 1
663:     If lstSuppressTiles.Selected(lLoop) Then
664:       Set pFeatLayer = FindLayer(lstSuppressTiles.List(lLoop), pMap)
665:       pSpatial.GeometryField = pFeatLayer.FeatureClass.shapeFieldName
666:       Set pFeatCursor = pFeatLayer.Search(pSpatial, True)
667:       Set pSearchFeat = pFeatCursor.NextFeature
668:       If Not pSearchFeat Is Nothing Then
669:         FeaturesInTile = True
        Exit Function
671:       End If
672:     End If
673:   Next lLoop

  Exit Function
  
ErrHand:
678:   MsgBox "FeaturesInTile - " & Err.Description
End Function

Private Sub txtNumbering_KeyUp(KeyCode As Integer, Shift As Integer)
682:   If Not IsNumeric(txtNumbering.Text) Then
683:     txtNumbering.Text = "1"
'  ElseIf CInt(txtNumbering.Text) < 0 Then
'    MsgBox "Can not use a number less than 0!!!"
'    txtNumbering.Text = "1"
687:   End If
End Sub
