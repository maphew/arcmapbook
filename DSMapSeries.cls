VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DSMapSeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Copyright 2008 ESRI
' 
' All rights reserved under the copyright laws of the United States
' and applicable international laws, treaties, and conventions.
' 
' You may freely redistribute and use this sample code, with or
' without modification, provided you include the original copyright
' notice and use restrictions.
' 
' See use restrictions at <your ArcGIS install location>/developerkit/userestrictions.txt.
' 




Option Explicit

Implements IDSMapSeries
Implements IDSMapSeriesProps
Implements IDSMapSeriesOptions
Implements IDSMapSeriesOptions2  'Added 6/18/03 to support cross hatching of clip
Implements IDSMapSeriesOptions3   'Added 11/23/2004 to support selection of the tile
Implements IPersistVariant

Private m_PageColl As Collection
Private m_bEnableSeries As Boolean
Private m_bClipData As Boolean
Private m_lClipData2 As Long  'Added 6/18/03 to support cross hatching of clip
Private m_sDataDrivenField As String
Private m_lExtentType As Long
Private m_dFixedScale As Double
Private m_bLabelNeighbors As Boolean
Private m_pLabelSymbol As ISymbol
Private m_dMargin As Double
Private m_sMarginType As String
Private m_bRotateFrame As Boolean
Private m_sRotationField As String
Private m_sDataFrameName As String
Private m_sIndexFieldName As String
Private m_sIndexLayerName As String
Private m_SuppressColl As Collection
Private m_bSupressLayers As Boolean
Private m_lTileSelection As Long
Private m_lStartNumber As Long
Private m_bSelectTile As Boolean   'Added 11/23/2004 to support selection of the tile

Private Sub Class_Initialize()
61:   Set m_PageColl = New Collection
62:   Set m_SuppressColl = New Collection
End Sub

Private Sub Class_Terminate()
66:   Set m_PageColl = Nothing
67:   Set m_SuppressColl = Nothing
End Sub

Private Sub IDSMapSeries_AddPage(ByVal Page As DSMapBookPrj.IDSMapPage)
71:   m_PageColl.Add Page
End Sub

Private Property Get IDSMapSeries_Page(Index As Long) As DSMapBookPrj.IDSMapPage
75:   If Index > -1 And Index < m_PageColl.Count Then
76:     Set IDSMapSeries_Page = m_PageColl.Item(Index + 1)
77:   Else
78:     Set IDSMapSeries_Page = Nothing
79:   End If
End Property

Private Property Get IDSMapSeries_PageCount() As Long
83:   IDSMapSeries_PageCount = m_PageColl.Count
End Property

Private Property Let IDSMapSeries_EnableSeries(ByVal RHS As Boolean)
87:   m_bEnableSeries = RHS
End Property

Private Property Get IDSMapSeries_EnableSeries() As Boolean
91:   IDSMapSeries_EnableSeries = m_bEnableSeries
End Property

Private Sub IDSMapSeries_RemovePage(Index As Long)
95:   If Index > -1 And Index < m_PageColl.Count Then
96:     m_PageColl.Remove Index + 1
97:   End If
End Sub

Private Property Let IDSMapSeriesOptions_ClipData(RHS As Boolean)
101:   m_bClipData = RHS
End Property

Private Property Get IDSMapSeriesOptions_ClipData() As Boolean
105:   IDSMapSeriesOptions_ClipData = m_bClipData
End Property

Private Property Let IDSMapSeriesOptions_DataDrivenField(RHS As String)
109:   m_sDataDrivenField = RHS
End Property

Private Property Get IDSMapSeriesOptions_DataDrivenField() As String
113:   IDSMapSeriesOptions_DataDrivenField = m_sDataDrivenField
End Property

Private Property Let IDSMapSeriesOptions_ExtentType(RHS As Long)
117:   m_lExtentType = RHS
End Property

Private Property Get IDSMapSeriesOptions_ExtentType() As Long
121:   IDSMapSeriesOptions_ExtentType = m_lExtentType
End Property

Private Property Let IDSMapSeriesOptions_FixedScale(RHS As Double)
125:   m_dFixedScale = RHS
End Property

Private Property Get IDSMapSeriesOptions_FixedScale() As Double
129:   IDSMapSeriesOptions_FixedScale = m_dFixedScale
End Property

Private Property Let IDSMapSeriesOptions_LabelNeighbors(RHS As Boolean)
133:   m_bLabelNeighbors = RHS
End Property

Private Property Get IDSMapSeriesOptions_LabelNeighbors() As Boolean
137:   IDSMapSeriesOptions_LabelNeighbors = m_bLabelNeighbors
End Property

Private Property Set IDSMapSeriesOptions_LabelSymbol(RHS As ISymbol)
141:   Set m_pLabelSymbol = RHS
End Property

Private Property Get IDSMapSeriesOptions_LabelSymbol() As ISymbol
145:   Set IDSMapSeriesOptions_LabelSymbol = m_pLabelSymbol
End Property

Private Property Let IDSMapSeriesOptions_Margin(RHS As Double)
149:   m_dMargin = RHS
End Property

Private Property Get IDSMapSeriesOptions_Margin() As Double
153:   IDSMapSeriesOptions_Margin = m_dMargin
End Property

Private Property Let IDSMapSeriesOptions_MarginType(RHS As String)
157:   m_sMarginType = RHS
End Property

Private Property Get IDSMapSeriesOptions_MarginType() As String
161:   IDSMapSeriesOptions_MarginType = m_sMarginType
End Property

Private Property Let IDSMapSeriesOptions_RotateFrame(RHS As Boolean)
165:   m_bRotateFrame = RHS
End Property

Private Property Get IDSMapSeriesOptions_RotateFrame() As Boolean
169:   IDSMapSeriesOptions_RotateFrame = m_bRotateFrame
End Property

Private Property Let IDSMapSeriesOptions_RotationField(RHS As String)
173:   m_sRotationField = RHS
End Property

Private Property Get IDSMapSeriesOptions_RotationField() As String
177:   IDSMapSeriesOptions_RotationField = m_sRotationField
End Property

Private Property Let IDSMapSeriesOptions2_ClipData(RHS As Long)
'Added 6/18/03 to support cross hatching of clip
182:   m_lClipData2 = RHS
End Property

Private Property Get IDSMapSeriesOptions2_ClipData() As Long
'Added 6/18/03 to support cross hatching of clip
187:   IDSMapSeriesOptions2_ClipData = m_lClipData2
End Property

Private Property Let IDSMapSeriesOptions3_SelectTile(RHS As Boolean)
'Added 11/12/04 to support the selection of the tile
192:   m_bSelectTile = RHS
End Property

Private Property Get IDSMapSeriesOptions3_SelectTile() As Boolean
'Added 11/12/04 to support the selection of the tile
197:   IDSMapSeriesOptions3_SelectTile = m_bSelectTile
End Property

Private Sub IDSMapSeriesProps_AddLayerToSuppress(ByVal LayerName As String)
201:   m_SuppressColl.Add LayerName
End Sub

Private Property Let IDSMapSeriesProps_DataFrameName(RHS As String)
205:   m_sDataFrameName = RHS
End Property

Private Property Get IDSMapSeriesProps_DataFrameName() As String
209:   IDSMapSeriesProps_DataFrameName = m_sDataFrameName
End Property

Private Property Let IDSMapSeriesProps_IndexFieldName(RHS As String)
213:   m_sIndexFieldName = RHS
End Property

Private Property Get IDSMapSeriesProps_IndexFieldName() As String
217:   IDSMapSeriesProps_IndexFieldName = m_sIndexFieldName
End Property

Private Property Let IDSMapSeriesProps_IndexLayerName(RHS As String)
221:   m_sIndexLayerName = RHS
End Property

Private Property Get IDSMapSeriesProps_IndexLayerName() As String
225:   IDSMapSeriesProps_IndexLayerName = m_sIndexLayerName
End Property

Private Sub IDSMapSeriesProps_RemoveLayerToSuppress(Index As Long)
229:   If Index > -1 And Index < m_SuppressColl.Count Then
230:     m_SuppressColl.Remove Index + 1
231:   End If
End Sub

Private Property Let IDSMapSeriesProps_StartNumber(RHS As Long)
235:   m_lStartNumber = RHS
End Property

Private Property Get IDSMapSeriesProps_StartNumber() As Long
239:   IDSMapSeriesProps_StartNumber = m_lStartNumber
End Property

Private Property Get IDSMapSeriesProps_SuppressLayer(Index As Long) As String
243:   If Index > -1 And Index < m_SuppressColl.Count Then
244:     IDSMapSeriesProps_SuppressLayer = m_SuppressColl.Item(Index + 1)
245:   Else
246:     IDSMapSeriesProps_SuppressLayer = ""
247:   End If
End Property

Private Property Get IDSMapSeriesProps_SuppressLayerCount() As Long
251:   IDSMapSeriesProps_SuppressLayerCount = m_SuppressColl.Count
End Property

Private Property Let IDSMapSeriesProps_SuppressLayers(ByVal RHS As Boolean)
255:   m_bSupressLayers = RHS
End Property

Private Property Get IDSMapSeriesProps_SuppressLayers() As Boolean
259:   IDSMapSeriesProps_SuppressLayers = m_bSupressLayers
End Property

Private Property Let IDSMapSeriesProps_TileSelectionMethod(RHS As Long)
263:   m_lTileSelection = RHS
End Property

Private Property Get IDSMapSeriesProps_TileSelectionMethod() As Long
267:   IDSMapSeriesProps_TileSelectionMethod = m_lTileSelection
End Property

Private Property Get IPersistVariant_ID() As esriSystem.IUID
  Dim id As New UID
272:   id = "DSMapBookPrj.DSMapSeries"
273:   Set IPersistVariant_ID = id
End Property

Private Sub IPersistVariant_Load(ByVal Stream As esriSystem.IVariantStream)
'Load the persisted parameters of the renderer
On Error GoTo ErrHand:
  Dim lLoop As Long, lCount As Long, pPage As IDSMapPage, vClip As Variant
  Dim bClip As Boolean, vCount As Variant, lPropCount As Long
  
  'Added 2/18/04 to make the list of persisted properties more dynamic
283:   vCount = Stream.Read
284:   If UCase(TypeName(vCount)) = "BOOLEAN" Then  'Version created before the 2/18/04 update.
285:     m_bEnableSeries = CBool(vCount)
    'Added 6/18/03 to support cross hatching of area outside clip
287:     vClip = Stream.Read
288:     If UCase(TypeName(vClip)) = "BOOLEAN" Then
289:       bClip = CBool(vClip)
290:       If bClip Then
291:         m_lClipData2 = 1
292:       Else
293:         m_lClipData2 = 0
294:       End If
295:     Else
296:       m_lClipData2 = CLng(vClip)
297:     End If
298:     lPropCount = 14
299:   Else
300:     lPropCount = CLng(Mid(CStr(vCount), 21)) - 2
301:     m_bEnableSeries = Stream.Read
302:     m_lClipData2 = Stream.Read
303:   End If

  'Original set of properties that every persisted map book will have.
306:   m_sDataDrivenField = Stream.Read
307:   m_lExtentType = Stream.Read
308:   m_dFixedScale = Stream.Read
309:   m_bLabelNeighbors = Stream.Read
310:   Set m_pLabelSymbol = Stream.Read
311:   m_dMargin = Stream.Read
312:   m_sMarginType = Stream.Read
313:   m_bRotateFrame = Stream.Read
314:   m_sRotationField = Stream.Read
315:   m_sDataFrameName = Stream.Read
316:   m_sIndexFieldName = Stream.Read
317:   m_sIndexLayerName = Stream.Read
318:   m_bSupressLayers = Stream.Read
319:   m_lTileSelection = Stream.Read
  
  'Additional properties added after 2/18/04
322:   If lPropCount > 14 Then    'Checking for start number property
323:     m_lStartNumber = Stream.Read
324:   Else
325:     m_lStartNumber = 1
326:   End If
  
  'Tile Selection added 11/23/04
329:   If lPropCount > 15 Then
330:     m_bSelectTile = Stream.Read
331:   Else
332:     m_bSelectTile = False
333:   End If
  
  'More original properties.  Writen out below the new properties because they are of variable length
336:   lCount = Stream.Read
337:   If lCount > 0 Then
338:     For lLoop = 1 To lCount
339:       m_SuppressColl.Add Stream.Read
340:     Next lLoop
341:   End If
  
343:   lCount = Stream.Read
344:   If lCount > 0 Then
345:     For lLoop = 1 To lCount
346:       Set pPage = Stream.Read
347:       If lPropCount <= 14 Then   'MapBook created before page numbers were added.  In this case we manually assign the page numbers.
348:         pPage.PageNumber = lLoop
349:       End If
350:       m_PageColl.Add pPage
351:     Next lLoop
352:   End If
    
  Exit Sub
ErrHand:
356:   MsgBox "MapSeries - IPersistVariant_Load - " & Erl & " - " & Err.Description
End Sub

Private Sub IPersistVariant_Save(ByVal Stream As esriSystem.IVariantStream)
'Write it all out
On Error GoTo ErrHand:
  Dim lLoop As Long
  
  'Added 2/18/04 to make the list of persisted properties more dynamic
  'Count changed from 17 to 18 on 11/23/2004 to support selection of tile
366:   Stream.Write "SERIESPROPERTYCOUNT-18"
    
368:   Stream.Write m_bEnableSeries
  
  'Added 6/18/03 to support cross hatching of area outside the clip
371:   Stream.Write m_lClipData2
372:   Stream.Write m_sDataDrivenField
373:   Stream.Write m_lExtentType
374:   Stream.Write m_dFixedScale
375:   Stream.Write m_bLabelNeighbors
376:   Stream.Write m_pLabelSymbol
377:   Stream.Write m_dMargin
378:   Stream.Write m_sMarginType
379:   Stream.Write m_bRotateFrame
380:   Stream.Write m_sRotationField
381:   Stream.Write m_sDataFrameName
382:   Stream.Write m_sIndexFieldName
383:   Stream.Write m_sIndexLayerName
384:   Stream.Write m_bSupressLayers
385:   Stream.Write m_lTileSelection
386:   Stream.Write m_lStartNumber   'Added 2/18/04
387:   Stream.Write m_bSelectTile       'Added 11/23/04 to support selection of tile

On Error GoTo ErrHand2:
390:   Stream.Write m_SuppressColl.Count
391:   If m_SuppressColl.Count > 0 Then
392:     For lLoop = 1 To m_SuppressColl.Count
393:       Stream.Write m_SuppressColl.Item(lLoop)
394:     Next lLoop
395:   End If
396:   Stream.Write m_PageColl.Count
397:   If m_PageColl.Count > 0 Then
398:     For lLoop = 1 To m_PageColl.Count
399:       Stream.Write m_PageColl.Item(lLoop)
400:     Next lLoop
401:   End If
    
  Exit Sub
ErrHand:
405:   MsgBox "MapSeries - IPersistVariant_Save - " & Err.Description
ErrHand2:
407:   MsgBox "MapSeriesCOLLS - IPersistVariant_Save - " & Err.Description
End Sub
