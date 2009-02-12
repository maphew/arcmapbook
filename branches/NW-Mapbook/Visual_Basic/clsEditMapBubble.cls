VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEditMapBubble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ICommand

Private m_pApp As IApplication
Private m_pMxApp As IMxApplication
Private m_pEditor As IEditor
Private m_pNWSeriesOptions As INWMapSeriesOptions
Private m_pMapBook As INWDSMapBook
Private m_pMapSeries As INWDSMapSeries
Const c_sModuleFileName As String = "clsEditMapBubble.cls"
Private m_pFeature As IFeature
Private m_pBubbleLayer  As ILayer

'graphics required to show feedback in the display
'for the current detail inset settings.
'-------------------------------------------------
Private m_pGraCont As IGraphicsContainer
Private m_pElemFillShp1 As IFillShapeElement
Private m_pElemFillShp2 As IFillShapeElement
Private m_pElemLine As ILineElement
Private m_bMouseIsDown As Boolean
Private m_pPnt1 As IPoint
Private m_pPnt2 As IPoint
Private m_pPnt3 As IPoint
Private m_pPnt4 As IPoint
Private m_sErrMessage As String  'used for debugging the Enabled event without messageboxes

Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
30:   ICommand_Bitmap = frmResources.picEditBubble.Picture.Handle
End Property

Private Property Get ICommand_Caption() As String
34:  ICommand_Caption = "EdBubble"
End Property

Private Property Get ICommand_Category() As String
38:   ICommand_Category = "NW MapBook"
End Property

Private Property Get ICommand_Checked() As Boolean
42:   ICommand_Checked = False
End Property

Private Property Get ICommand_Enabled() As Boolean
  On Error GoTo ErrorHandler

  Dim pMxDoc As IMxDocument
                                                  'confirm that I'm in data view
50:   Set pMxDoc = m_pApp.Document
  'If TypeOf pMxDoc.ActiveView Is IPageLayout Then
  '  ICommand_Enabled = False
  '  Exit Property
  'End If
                                                  'confirm that a bubble definitions layer
                                                  'has been identified, selected, maybe as
                                                  'a property of the map series
58:   If m_pNWSeriesOptions Is Nothing Then
59:     Set m_pMapBook = GetMapBookExtension(m_pApp)
60:     If m_pMapBook Is Nothing Then
61:       ICommand_Enabled = False
''''''''''''debug
'63: m_sErrMessage = "1"
      Exit Property
65:     End If
66:     Set m_pMapSeries = m_pMapBook.ContentItem(0)
67:     If m_pMapSeries Is Nothing Then
68:       ICommand_Enabled = False
''''''''''''debug
'70: m_sErrMessage = "2"
      Exit Property
72:     End If
73:     Set m_pNWSeriesOptions = m_pMapSeries
74:     If m_pNWSeriesOptions Is Nothing Then
75:       ICommand_Enabled = False
''''''''''''debug
'77: m_sErrMessage = "3"
      Exit Property
79:     End If
80:   End If
81:   If m_pNWSeriesOptions.BubbleLayer = "" Then
82:     ICommand_Enabled = False
''''''''''''debug
'84: m_sErrMessage = "4"
    Exit Property
86:   End If
                                                  'check that the editing environment
                                                  'is active
89:   If m_pEditor.EditState <> esriStateEditing Then
90:     ICommand_Enabled = False
''''''''''''debug
'92: m_sErrMessage = "5"
    Exit Property
94:   End If
                                                  'confirm that there is one selected feature, and
                                                  'that this feature is from the bubble layer
97:   If pMxDoc.FocusMap.SelectionCount <> 1 Then
98:     ICommand_Enabled = False
''''''''''''debug
'100: m_sErrMessage = "6"
    Exit Property
102:   End If
  
  Dim pLayer As ILayer, pLayers As IEnumLayer, pFeatLayer As IFeatureLayer
  Dim pEnumFeat As IEnumFeature, pFeature As IFeature, pFeatClass As IFeatureClass
  Dim pDataset1 As IDataset, pBubbleDataset As IDataset, pTable As ITable
  Dim pObjectClass As IObjectClass
  
  
110:   Set pLayers = pMxDoc.FocusMap.Layers
111:   pLayers.Reset
112:   Set pLayer = pLayers.Next
113:   Do While Not pLayer Is Nothing
114:     If pLayer.Name = m_pNWSeriesOptions.BubbleLayer Then
115:       Set pFeatLayer = pLayer
116:       Set pFeatClass = pFeatLayer.FeatureClass
117:     End If
118:     Set pLayer = pLayers.Next
119:   Loop
  
121:   Set pEnumFeat = pMxDoc.FocusMap.FeatureSelection
122:   pEnumFeat.Reset
123:   Set pFeature = pEnumFeat.Next
                                                  'disable the command if the feature isn't
                                                  'from the bubble feature class
126:   Set pTable = pFeature.Table
127:   Set pDataset1 = pTable
128:   If pDataset1.Name <> m_pNWSeriesOptions.BubbleLayer Then
129:     ICommand_Enabled = False
''''''''''''debug
'131: m_sErrMessage = "7 - pDataset1.Name is " & pDataset1.Name & ", m_pNWSeriesOptions.BubbleLayer is " & m_pNWSeriesOptions.BubbleLayer
    Exit Property
133:   End If
  
135:   Set m_pBubbleLayer = pFeatLayer
136:   Set m_pFeature = pFeature
137:   ICommand_Enabled = True

  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Enabled " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Property Get ICommand_HelpContextID() As Long

End Property

Private Property Get ICommand_HelpFile() As String

End Property

Private Property Get ICommand_Message() As String
153:   ICommand_Message = "Edit a circular map inset"
End Property

Private Property Get ICommand_Name() As String
157:   ICommand_Name = "NWMapBookUIPrj_clsEditMapBubble"
End Property

Private Sub ICommand_OnClick()
  On Error GoTo ErrorHandler

  If m_pFeature Is Nothing Then Exit Sub
  If m_pBubbleLayer Is Nothing Then Exit Sub
  
  Dim pFeatLayer As IFeatureLayer, pFeatClass As IFeatureClass
  
168:   Set pFeatLayer = m_pBubbleLayer
169:   If Not m_pNWSeriesOptions.IsBubbleLayer(pFeatLayer.FeatureClass) Then
170:     MsgBox "The selected feature is not from a valid detail insets layer."
    Exit Sub
172:   End If
  
  'grab the stats from the bubble feature that
  'is needed by the bubbleproperties dialog.
  
  'two options are available for determining the
  'attributes of a bubble feature: use either
  '  1 -- the actual spatial attrs of the polygon, or
  '  2 -- us the attr field values.
  'Either approach has its merits, but to allow the
  '  user to modify a detail inset using ArcGIS's
  '  extensive spatial tools, then allow this to be
  '  supported by this editor tool, the values will
  '  be accessed from the spatial properties of the
  '  polygon.
  
  Dim pPolygon As IPolygon, pEnvelope As IEnvelope, pField As IField
  Dim pFields As IFields, lFieldCount As Long, pRow As IRow
  
  Dim lXDestIdx As Long, lYDestIdx As Long, lWidthOrgIdx As Long
  Dim lXOrgIdx As Long, lYOrgIdx As Long, lLayersIdx As Long
  Dim lRadiusIdx As Long, lScaleidx As Long, lBubbleIdIdx As Long
  Dim i As Long, pMxDoc As IMxDocument, sDataFrameName As String
  Dim pDataFrameElement As IFrameElement, pMapFrame As IMapFrame
  Dim pElement As IElement, pDispTrans As IDisplayTransformation
  
198:   Set pPolygon = m_pFeature.Shape
199:   Set pEnvelope = pPolygon.envelope

201:   Set pFields = m_pFeature.Fields
202:   lFieldCount = pFields.FieldCount
203:   Set pRow = m_pFeature
  
205:   For i = 0 To lFieldCount - 1
206:     Set pField = pFields.Field(i)
    Select Case pField.Name
    Case "BUBBLEID"
209:       lBubbleIdIdx = i
    Case "XORG"
211:       lXOrgIdx = i
    Case "YORG"
213:       lYOrgIdx = i
    Case "XDEST"
215:       lXDestIdx = i
    Case "YDEST"
217:       lYDestIdx = i
    Case "RADIUS"
219:       lRadiusIdx = i
    Case "SCALE"
221:       lScaleidx = i
    Case "LAYERS"
223:       lLayersIdx = i
    Case "WIDTHORG"
225:       lWidthOrgIdx = i
226:     End Select
227:   Next i
  
  
230:   With frmBubbleProperties
231:     .Initializing = True
232:     .HeightWidthIsEnabled = (pEnvelope.Width = pEnvelope.Height)
233:     Set .Application = m_pApp
234:     .OriginX = ((pEnvelope.XMax - pEnvelope.XMin) / 2) + pEnvelope.XMin
235:     .OriginY = ((pEnvelope.YMax - pEnvelope.YMin) / 2) + pEnvelope.YMin
236:     .SourceWidth = pEnvelope.Width
    '.SourceHeight = pEnvelope.Height
238:     .Radius = (pEnvelope.Height / 2)
    
240:     .BubbleID = pRow.Value(lBubbleIdIdx)
241:     sDataFrameName = pRow.Value(lBubbleIdIdx)
242:     sDataFrameName = "BubbleID:" & sDataFrameName
243:     .DestinationX = pRow.Value(lXDestIdx)
244:     .DestinationY = pRow.Value(lYDestIdx)
245:     .ScaleFactor = pRow.Value(lScaleidx)
246:     If IsNull(pRow.Value(lLayersIdx)) Then
247:       .Layers = ""
248:     Else
249:       .Layers = pRow.Value(lLayersIdx)
250:     End If
251:     .Initializing = False
252:     .BubbleRadius = .Radius * .ScaleFactor
253:     .BubbleRadiusX = (.SourceWidth / 2) * .ScaleFactor
  
    'remove the data frame for the current detail inset
    '--------------------------------------------
257:     Set pMxDoc = m_pApp.Document
258:     m_pNWSeriesOptions.BubbleGraphicRemove pMxDoc, sDataFrameName
    
    
    'make feedback graphics available showing current inset location
    '--------------------------------------------
263:     Set m_pPnt1 = New Point
264:     Set m_pPnt2 = New Point
265:     Set m_pPnt3 = New Point
266:     Set m_pPnt4 = New Point
267:     m_pPnt1.x = .OriginX
268:     m_pPnt1.y = .OriginY
269:     m_pPnt2.x = .OriginX + .Radius
270:     m_pPnt2.y = .OriginY
271:     m_pPnt3.x = .DestinationX
272:     m_pPnt3.y = .DestinationY
273:     m_pPnt4.x = .DestinationX + .BubbleRadiusX
274:     m_pPnt4.y = .DestinationY
    
276:     Set .OriginShape = m_pElemFillShp1
277:     Set .DestinationShape = m_pElemFillShp2
278:     Set .BetweenLineElement = m_pElemLine
279:     Set .GraphicsContainer = m_pGraCont
280:     Set .PointOriginCenter = m_pPnt1
281:     Set .PointOriginEdge = m_pPnt2
282:     Set .PointDestCenter = m_pPnt3
283:     Set .PointDestEdge = m_pPnt4
284:     .DataFrameName = m_pNWSeriesOptions.DataFrameMainFrame
  
  
    'Get user input through the bubble properties dialog
    '---------------------------------------------------
289:     m_pNWSeriesOptions.RefreshEventShouldntRefresh = True
290:     .Show vbModal
  
292:     m_pEditor.StartOperation
293:     If .WasCancelled Then
294:       Set pElement = GetDataFrameElement(m_pNWSeriesOptions.DataFrameMainFrame, m_pApp)
295:       AddDataframe pElement, m_pApp, pRow, m_pNWSeriesOptions
296:       m_pNWSeriesOptions.RefreshEventShouldntRefresh = False
      Exit Sub
298:     End If

300:     m_pNWSeriesOptions.RefreshEventShouldntRefresh = False
301:     pRow.Value(lXOrgIdx) = .OriginX
302:     pRow.Value(lYOrgIdx) = .OriginY
303:     pRow.Value(lXDestIdx) = .DestinationX
304:     pRow.Value(lYDestIdx) = .DestinationY
305:     pRow.Value(lRadiusIdx) = .Radius
306:     pRow.Value(lScaleidx) = .ScaleFactor
307:     pRow.Value(lLayersIdx) = .Layers
308:     pRow.Value(lWidthOrgIdx) = .SourceWidth
309:     pRow.Store
     
    
312:     Set m_pFeature.Shape = CreateInsetPolygon(.OriginX, .OriginY, .SourceHeight, .SourceWidth)
313:     m_pFeature.Store
    
315:     m_pEditor.StopOperation "EditMapBubble"
316:   End With
                                                  'update the visible layers
  Dim pMaps As IMaps, pNewMap As IMap
  
320:   Set pMaps = pMxDoc.Maps
321:   For i = 0 To pMaps.count - 1
322:     Set pNewMap = pMaps.Item(i)
323:     If pNewMap.Name = sDataFrameName Then
324:       SetVisibleLayers pNewMap, frmBubbleProperties.Layers
325:       i = pMaps.count - 1
326:     End If
327:   Next i
  
  
  
  'create a new data frame after the editing
  'of the previous data frame's settings is complete.
  '--------------------------------------------------
334:   Set pElement = GetDataFrameElement(m_pNWSeriesOptions.DataFrameMainFrame, m_pApp)
335:   AddDataframe pElement, m_pApp, pRow, m_pNWSeriesOptions

337:   pMxDoc.ActiveView.Refresh


  Exit Sub
ErrorHandler:
342:   m_pNWSeriesOptions.RefreshEventShouldntRefresh = False
343:   m_pEditor.AbortOperation
  HandleError True, "ICommand_OnClick " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
  Dim pUID As New UID
  
350:   Set m_pApp = hook
351:   Set m_pMxApp = m_pApp
352:   pUID.Value = "esriEditor.Editor"
353:   Set m_pEditor = m_pApp.FindExtensionByCLSID(pUID)

355:   Set m_pMapBook = GetMapBookExtension(m_pApp)
  If m_pMapBook Is Nothing Then Exit Sub
357:   Set m_pMapSeries = m_pMapBook.ContentItem(0)
358:   Set m_pNWSeriesOptions = m_pMapSeries
End Sub

Private Property Get ICommand_Tooltip() As String
362:   ICommand_Tooltip = "Edit an existing circular map inset"
  '''''''''debug
'359:   ICommand_Tooltip = m_sErrMessage
End Property







''AddDataFrame
''
''This routine assumes that the pRow reference being passed in
''has already been confirmed as a valid bubble reference layer.
''
''Parameters:
''   pElementWithinBorders is the element reference to the main
''   map data frame.
''------------------------------------------------------------
Private Sub AddDataframe(pElementWithinBorders As IElement, pApp As IApplication, pRow As IRow, pNWSeriesOptions As INWMapSeriesOptions)
  On Error GoTo ErrorHandler

  Dim pFields As IFields, lFieldCount As Long, pField As IField
  Dim pNewMap As IMap, pMapFrame As IMapFrame, pEnv As IEnvelope
  Dim pGraphicsContainer As IGraphicsContainer, pElement As IElement
  Dim pFrameElement As IFrameElement, pNewMapAV As IActiveView
  Dim pNewEnv As IEnvelope

    
  Dim pMxDocDataViewAV As IActiveView, pScrDisplay As IScreenDisplay
  Dim pMxDocLayoutViewAV As IActiveView, pMxDoc As IMxDocument
  
  Dim lBubbleId As Long, dXOrigin As Double, dYOrigin As Double
  Dim dXDestination As Double, dYDestination As Double, dRadius As Double
  Dim dScale As Double, sLayers As String, dWidthOrigin As Double
  Dim bIsCircular As Boolean, i As Long, dPageToMapUnitRatio As Double
  Dim dMapScale As Double, dDataFrameWidth As Double, dDataFrameHeight As Double
  Dim sActiveDataFrameName As String, pMxDocFocusMapQIDataViewAV As IActiveView
  Dim pMxDocFocusMapQILayoutViewAV As IActiveView, pMainMapFrame As IMapFrame

  If pRow Is Nothing Then Exit Sub
  If pApp Is Nothing Then Exit Sub
405:   Set pMxDoc = pApp.Document
  
407:   Set pFields = pRow.Fields
408:   lFieldCount = pFields.FieldCount
    
410:   For i = 0 To lFieldCount - 1
411:     Set pField = pFields.Field(i)
    Select Case pField.Name
    Case "BUBBLEID"
414:       lBubbleId = pRow.Value(i)
    Case "XORG"
416:       dXOrigin = pRow.Value(i)
    Case "YORG"
418:       dYOrigin = pRow.Value(i)
    Case "XDEST"
420:       dXDestination = pRow.Value(i)
    Case "YDEST"
422:       dYDestination = pRow.Value(i)
    Case "RADIUS"
424:       dRadius = pRow.Value(i)
    Case "SCALE"
426:       dScale = pRow.Value(i)
    Case "LAYERS"
428:       sLayers = pRow.Value(i)
    Case "WIDTHORG"
430:       dWidthOrigin = pRow.Value(i)
431:     End Select
432:   Next i

434:   Set pNewMap = New Map
435:   pNewMap.Name = "BubbleID:" & lBubbleId
436:   pNewMap.Description = "Detail Inset " & lBubbleId
  
  
  ' clone all copies of layers so that modifying
  ' layers in one dataframe doesn't impact layers
  ' referenced in the other data frames.
  ''''''''''''''''''''''
  Dim pLayer As ILayer, pLayerSrc As ILayer
  Dim pFeatLyr As IFeatureLayer, pTinLyr As ITinLayer, pRastLyrSrc As IRasterLayer
  Dim pRastLyr As IRasterLayer, pFeatLyrSrc As IFeatureLayer, pTinLyrSrc As ITinLayer
  Dim pGeoFeatLyrSource As IGeoFeatureLayer, pGeoFeatLyrDestination As IGeoFeatureLayer
  
448:   If pMxDoc.FocusMap.LayerCount > 0 Then
449:     For i = (pMxDoc.FocusMap.LayerCount - 1) To 0 Step -1
450:       Set pLayerSrc = pMxDoc.FocusMap.Layer(i)
451:       If TypeOf pMxDoc.FocusMap.Layer(i) Is IFeatureLayer Then
452:         Set pFeatLyr = New FeatureLayer
453:         Set pFeatLyrSrc = pLayerSrc
454:         pFeatLyr.DataSourceType = pFeatLyrSrc.DataSourceType
455:         pFeatLyr.DisplayField = pFeatLyrSrc.DisplayField
456:         Set pFeatLyr.FeatureClass = pFeatLyrSrc.FeatureClass
457:         pFeatLyr.ScaleSymbols = pFeatLyrSrc.ScaleSymbols
458:         pFeatLyr.Selectable = pFeatLyrSrc.Selectable
459:         Set pGeoFeatLyrSource = pFeatLyrSrc
460:         Set pGeoFeatLyrDestination = pFeatLyr
461:         With pGeoFeatLyrSource
                                            'Set pGeoFeatLyrDestination.CurrentMapLevel = .CurrentMapLevel
                                            'pGeoFeatLyrDestination.DisplayFeatureClass = .DisplayFeatureClass
464:           pGeoFeatLyrDestination.AnnotationProperties = .AnnotationProperties
465:           pGeoFeatLyrDestination.AnnotationPropertiesID = .AnnotationPropertiesID
466:           pGeoFeatLyrDestination.DisplayAnnotation = .DisplayAnnotation
467:           Set pGeoFeatLyrDestination.ExclusionSet = .ExclusionSet
468:           Set pGeoFeatLyrDestination.Renderer = .Renderer
469:         End With
470:         Set pLayer = pFeatLyr
        
472:       ElseIf TypeOf pMxDoc.FocusMap.Layer(i) Is ITinLayer Then
                                            'pTinLyr.RendererCount = pTinLyrSrc.RendererCount
474:         Set pTinLyrSrc = pLayerSrc
475:         Set pTinLyr = New TinLayer
476:         Set pTinLyr.Dataset = pTinLyrSrc.Dataset
477:         pTinLyr.DisplayField = pTinLyrSrc.DisplayField
478:         pTinLyr.ScaleSymbols = pTinLyrSrc.ScaleSymbols
479:         Set pLayer = pTinLyr
480:       ElseIf TypeOf pMxDoc.FocusMap.Layer(i) Is IRasterLayer Then
                                            'pRastLyr.BandCount = pRastLyrSrc.BandCount
                                            'pRastLyr.ColumnCount = pRastLyrSrc.ColumnCount
                                            'pRastLyr.DataFrameExtent = pRastLyrSrc.DataFrameExtent
                                            'pRastLyr.FilePath = pRastLyrSrc.FilePath
                                            'pRastLyr.Raster = pRastLyrSrc.Raster
                                            'pRastLyr.RowCount = pRastLyrSrc.RowCount
487:         Set pRastLyr = New RasterLayer
488:         Set pRastLyrSrc = pLayerSrc
489:         pRastLyr.DisplayResolutionFactor = pRastLyrSrc.DisplayResolutionFactor
490:         pRastLyr.PrimaryField = pRastLyrSrc.PrimaryField
491:         pRastLyr.PyramidPresent = pRastLyrSrc.PyramidPresent
492:         Set pRastLyr.Renderer = pRastLyrSrc.Renderer
493:         pRastLyr.ShowResolution = pRastLyrSrc.ShowResolution
494:         pRastLyr.VisibleExtent = pRastLyrSrc.VisibleExtent
495:         Set pLayer = pRastLyr
496:       End If
                                            'Set pLayer.AreaOfInterest = pLayerSrc.AreaOfInterest
                                            'Set pLayer.SpatialReference = pLayerSrc.SpatialReference
                                            'pLayer.SupportedDrawPhases = pLayerSrc.SupportedDrawPhases
                                            'pLayer.TipText = pLayerSrc.TipText
                                            'pLayer.Valid = pLayerSrc.Valid
502:       pLayer.Cached = pLayerSrc.Cached
503:       pLayer.MaximumScale = pLayerSrc.MaximumScale
504:       pLayer.MinimumScale = pLayerSrc.MinimumScale
505:       pLayer.Name = pLayerSrc.Name
506:       pLayer.ShowTips = pLayerSrc.ShowTips
507:       pLayer.Visible = pLayerSrc.Visible
      
509:       pNewMap.AddLayer pLayer
510:     Next i
511:   End If
512:   SetVisibleLayers pNewMap, sLayers
                                                  'Create a new MapFrame and associate
                                                  'map with it
515:   Set pMapFrame = New MapFrame
516:   Set pMapFrame.Map = pNewMap
517:   pMapFrame.ExtentType = esriExtentDefault
  
519:   Set pGraphicsContainer = pMxDoc.PageLayout
                                                  'Set the position of the new map frame
521:   Set pElement = pMapFrame
522:   Set pEnv = New envelope
523:   Set pMxDocDataViewAV = pMxDoc.ActiveView
524:   Set pMxDocFocusMapQIDataViewAV = pMxDoc.FocusMap
  
526:   bIsCircular = ((2 * dRadius) = dWidthOrigin)
  
  
  Dim dBubbleRadius As Double, dBubbleWidth As Double
  Dim dBubbleLeft As Double, dBubbleRight As Double
  Dim dBubbleTop As Double, dBubbleBottom As Double
  Dim dScreenRightInches As Double, dScreenBottomInches As Double
  Dim lBubbleTop As Long, lBubbleBottom As Long
  Dim lBubbleLeft As Long, lBubbleRight As Long
  Dim lBubbleTopMapAV As Long, lBubbleBottomMapAV As Long
  Dim lBubbleLeftMapAV As Long, lBubbleRightMapAV As Long
  Dim pMapEnv As IEnvelope, pBubbleEnv As IEnvelope
  Dim pMap As IMap, pActiveView As IActiveView
  
  
541:   dBubbleRadius = dRadius * dScale
542:   dBubbleWidth = dWidthOrigin * dScale
                                                  'generate the top y
                                                  'generate the bottom y
                                                  'the left x and right x
                                                  'for detail insets
547:   dBubbleTop = dYDestination + dBubbleRadius
548:   dBubbleBottom = dYDestination - dBubbleRadius
549:   If bIsCircular Then
550:     dBubbleLeft = dXDestination - dBubbleRadius
551:     dBubbleRight = dXDestination + dBubbleRadius
552:   Else
553:     dBubbleLeft = dXDestination - (dBubbleWidth / 2)
554:     dBubbleRight = dXDestination + (dBubbleWidth / 2)
555:   End If
  
557:   Set pBubbleEnv = New envelope
558:   Set pMapEnv = pMxDoc.ActiveView.Extent 'sides of display in screen units
  
  
  
  'Shift to layout view
  '--------------------
564:   Set pMxDoc.ActiveView = pMxDoc.PageLayout
565:   Set pMapEnv = pMxDoc.ActiveView.Extent
566:   Set pMxDocLayoutViewAV = pMxDoc.ActiveView
567:   Set pMxDocFocusMapQILayoutViewAV = pMxDoc.FocusMap
  
  
  'Size and place the detail inset data frame
  '------------------------------------------
  Dim dDocWidthInches As Double, dDocHeightInches As Double
  Dim pPntScrBottomLeft As IPoint, pPntScrTopRight As IPoint
  Dim dScrMapUnitWidth As Double, dScrMapUnitHeight As Double
  Dim dScrMapUnitWidthWOBorders As Double, dScrMapUnitHeightWOBorders As Double
  Dim dRatioInchToMapUnitX As Double, dRatioInchToMapUnitY As Double
  Dim dBubbleLeftInches As Double, dBubbleRightInches As Double
  Dim dBubbleTopInches As Double, dBubbleBottomInches As Double
  Dim dBoundaryInchesXLeft As Double, dBoundaryInchesYTop As Double
  
581:   With pMapEnv
582:     dDocWidthInches = .XMax - (Abs(.XMin))
583:     dDocHeightInches = .YMax - (Abs(.YMin))
584:   End With
                                                  'get the data area width minus the
                                                  'buffer of empty space around the
                                                  'map display
588:   dDataFrameWidth = pElementWithinBorders.Geometry.envelope.Width
589:   dDataFrameHeight = pElementWithinBorders.Geometry.envelope.Height
590:   dBoundaryInchesXLeft = pElementWithinBorders.Geometry.envelope.XMin
591:   dBoundaryInchesYTop = pElementWithinBorders.Geometry.envelope.YMin
  
593:   With pMxDocFocusMapQILayoutViewAV.ScreenDisplay.DisplayTransformation
594:     Set pPntScrBottomLeft = .ToMapPoint(.DeviceFrame.Left, .DeviceFrame.bottom)
595:     Set pPntScrTopRight = .ToMapPoint(.DeviceFrame.Right, .DeviceFrame.Top)
596:   End With
  
598:   dScrMapUnitWidth = pPntScrTopRight.x - pPntScrBottomLeft.x
599:   dScrMapUnitHeight = pPntScrTopRight.y - pPntScrBottomLeft.y
600:   dRatioInchToMapUnitX = dDataFrameWidth / dScrMapUnitWidth
601:   dRatioInchToMapUnitY = dDataFrameHeight / dScrMapUnitHeight
  
603:   dBubbleLeftInches = ((dBubbleLeft - pPntScrBottomLeft.x) * dRatioInchToMapUnitX) + dBoundaryInchesXLeft
604:   dBubbleRightInches = ((dBubbleRight - pPntScrBottomLeft.x) * dRatioInchToMapUnitX) + dBoundaryInchesXLeft
605:   dBubbleTopInches = ((dBubbleTop - pPntScrBottomLeft.y) * dRatioInchToMapUnitY) + dBoundaryInchesYTop
606:   dBubbleBottomInches = ((dBubbleBottom - pPntScrBottomLeft.y) * dRatioInchToMapUnitY) + dBoundaryInchesYTop
  
608:   pBubbleEnv.XMin = dBubbleLeftInches
609:   pBubbleEnv.XMax = dBubbleRightInches
610:   pBubbleEnv.YMin = dBubbleBottomInches
611:   pBubbleEnv.YMax = dBubbleTopInches
  
  
  
615:   pElement.Geometry = pBubbleEnv
  
  '''''''''''''''''''''''
  ' data frame properties
  
  Dim pColor As IColor, pFillColor As IColor, pShadowColor As IColor
621:   Set pColor = New RgbColor
622:   Set pFillColor = New RgbColor
623:   Set pShadowColor = New RgbColor
624:   pColor.RGB = RGB(210, 210, 210)
625:   pFillColor.RGB = RGB(255, 255, 255)
626:   pShadowColor.RGB = RGB(128, 128, 128)
    
628:   Set pFrameElement = pElement
  
  ' create a border
  Dim pSymbolBorder As ISymbolBorder
  Dim pLineSymbol As ILineSymbol
  Dim pFrameDecoration As IFrameDecoration
  Dim pShadowFillSymbol As IFillSymbol
  Dim pSymbolShadow As ISymbolShadow
  Dim pFrameProperties As IFrameProperties
  
638:   Set pSymbolBorder = New SymbolBorder
639:   Set pLineSymbol = New SimpleLineSymbol
640:   pLineSymbol.Color = pColor
641:   pSymbolBorder.LineSymbol = pLineSymbol
642:   pSymbolBorder.LineSymbol.Color = pColor
643:   pSymbolBorder.CornerRounding = 100
644:   pFrameElement.Border = pSymbolBorder
  
  'modify the frame element background
647:   Set pFrameDecoration = New SymbolBackground
648:   pFrameDecoration.Color = pFillColor
649:   pFrameDecoration.CornerRounding = 100
  
651:   pFrameElement.Background = pFrameDecoration
  
  
  ' add shadow to detail inset
655:   Set pShadowFillSymbol = New SimpleFillSymbol
656:   pShadowFillSymbol.Color = pShadowColor
657:   pShadowFillSymbol.Outline.Color = pShadowColor
658:   Set pSymbolShadow = New SymbolShadow
659:   pSymbolShadow.FillSymbol = pShadowFillSymbol
660:   pSymbolShadow.HorizontalSpacing = -2
661:   pSymbolShadow.VerticalSpacing = -2
662:   pSymbolShadow.CornerRounding = 100
663:   Set pFrameProperties = pFrameElement
664:   pFrameProperties.Shadow = pSymbolShadow

  
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  'Build triangle and shadow triangle to point from the bubble inset to the
  '(presumably) circular polygon feature that represents the area in the inset
  '
  '  - Find the points for the triangle graphic to be created.
  '           - convert the origin and destination points to screen units.
  '           - convert the radius to screen units
  '           - take an angle 10 degrees from the line in each direction
  '       - create a triangle polygon from those points
  '       - add the triangle to the graphic layer
  
  
  Dim pLine As ILine, pFromPnt As IPoint, pToPnt As IPoint
  Dim pCircularArc As ICircularArc, pBaseLine As ILine
  Dim p1stPnt As IPoint, p2ndPnt As IPoint, p3rdPnt As IPoint
  Dim p1stShadowPnt As IPoint, p2ndShadowPnt As IPoint, p3rdShadowPnt As IPoint
  Dim pConstructPoint As IConstructPoint
  
685:   Set pLine = New esriGeometry.Line
686:   Set pFromPnt = New Point
687:   Set pToPnt = New Point
688:   pFromPnt.x = dXDestination
689:   pFromPnt.y = dYDestination
690:   pToPnt.x = dXOrigin
691:   pToPnt.y = dYOrigin
692:   pLine.PutCoords pFromPnt, pToPnt
  
  
695:   Set pCircularArc = New CircularArc
                                                  'angles are stored in radians,
                                                  'so calculate 10 degrees in radians
698:   pCircularArc.PutCoordsByAngle pFromPnt, _
                               (pLine.Angle - ((10 / 180) * 3.14159265358979)), _
                               ((20 / 180) * 3.14159265358979), _
                               dBubbleRadius
  
703:   Set pBaseLine = New esriGeometry.Line
704:   pBaseLine.PutCoords pCircularArc.FromPoint, pCircularArc.ToPoint
705:   Set p1stPnt = New Point
706:   Set p2ndPnt = New Point
707:   Set p3rdPnt = New Point
708:   Set p1stShadowPnt = New Point
709:   Set p2ndShadowPnt = New Point
710:   Set p3rdShadowPnt = New Point
711:   Set pConstructPoint = p3rdPnt
712:   pConstructPoint.ConstructDeflection pBaseLine, _
                                      pBaseLine.Length, _
                                      -((60 / 180) * 3.14159265358979)
                                                  '3 points are now available for triangle:
                                                  ' - pCircularArc.FromPoint
                                                  ' - pCircularArc.ToPoint,
                                                  ' - p3rdPnt
  Dim pTrianglePoly As IPolygon, pGeomColl As IGeometryCollection, pGeometry As IGeometry
  Dim pShadowTrianglePoly As IPolygon, pShadowPolygonElement As IPolygonElement
  Dim pPntColl As IPointCollection, pFeature As IFeature
  Dim pPolygonElement As IPolygonElement, pFillShapeElement As IFillShapeElement
  Dim pShadowPolyElement As IPolygonElement, pShadowFillShapeElement As IFillShapeElement
  Dim pArrowFillSymb As IFillSymbol, pElementPly As IElement, pElementShadowPly As IElement
  
  '''''''''''''''''''''''''''''''
  'triangle colors and dimensions
  
729:   Set pArrowFillSymb = New SimpleFillSymbol
730:   Set pPolygonElement = New PolygonElement
731:   Set pFillShapeElement = pPolygonElement
732:   pArrowFillSymb.Outline = pLineSymbol
733:   pArrowFillSymb.Color = pLineSymbol.Color
734:   pFillShapeElement.Symbol = pArrowFillSymb
735:   Set pElementPly = pPolygonElement
  
737:   Set pTrianglePoly = New esriGeometry.Polygon
738:   Set pPntColl = pTrianglePoly
739:   p1stPnt.x = ((pCircularArc.FromPoint.x - pPntScrBottomLeft.x) * dRatioInchToMapUnitX) + dBoundaryInchesXLeft
740:   p1stPnt.y = ((pCircularArc.FromPoint.y - pPntScrBottomLeft.y) * dRatioInchToMapUnitX) + dBoundaryInchesYTop
741:   p2ndPnt.x = ((pCircularArc.ToPoint.x - pPntScrBottomLeft.x) * dRatioInchToMapUnitX) + dBoundaryInchesXLeft
742:   p2ndPnt.y = ((pCircularArc.ToPoint.y - pPntScrBottomLeft.y) * dRatioInchToMapUnitX) + dBoundaryInchesYTop
743:   p3rdPnt.x = ((p3rdPnt.x - pPntScrBottomLeft.x) * dRatioInchToMapUnitX) + dBoundaryInchesXLeft
744:   p3rdPnt.y = ((p3rdPnt.y - pPntScrBottomLeft.y) * dRatioInchToMapUnitX) + dBoundaryInchesYTop
745:   pPntColl.AddPoint p1stPnt
746:   pPntColl.AddPoint p2ndPnt
747:   pPntColl.AddPoint p3rdPnt
748:   pTrianglePoly.Close
  
750:   Set pGeometry = pTrianglePoly
751:   pElementPly.Geometry = pGeometry
  
  ''''''''''''''''''''''''''''''''''''''
  'triangle shadow colors and dimensions
  
756:   Set pShadowPolyElement = New PolygonElement
757:   Set pShadowFillShapeElement = pShadowPolyElement
758:   Set pFillShapeElement = pShadowPolyElement
759:   pFillShapeElement.Symbol = pShadowFillSymbol
760:   Set pElementShadowPly = pShadowPolyElement
  
762:   Set pShadowTrianglePoly = New esriGeometry.Polygon
763:   Set pPntColl = pShadowTrianglePoly
                                                  'offset shadow triangle by 3 pixels
765:   p1stShadowPnt.x = p1stPnt.x - (ConvertPixelsToRW(2, pApp) * dRatioInchToMapUnitX)
766:   p1stShadowPnt.y = p1stPnt.y - (ConvertPixelsToRW(2, pApp) * dRatioInchToMapUnitY)
767:   p2ndShadowPnt.x = p2ndPnt.x - (ConvertPixelsToRW(2, pApp) * dRatioInchToMapUnitX)
768:   p2ndShadowPnt.y = p2ndPnt.y - (ConvertPixelsToRW(2, pApp) * dRatioInchToMapUnitY)
769:   p3rdShadowPnt.x = p3rdPnt.x - (ConvertPixelsToRW(2, pApp) * dRatioInchToMapUnitX)
770:   p3rdShadowPnt.y = p3rdPnt.y - (ConvertPixelsToRW(2, pApp) * dRatioInchToMapUnitY)
771:   pPntColl.AddPoint p1stShadowPnt
772:   pPntColl.AddPoint p2ndShadowPnt
773:   pPntColl.AddPoint p3rdShadowPnt
774:   pShadowTrianglePoly.Close
775:   Set pGeometry = pShadowTrianglePoly
776:   pElementShadowPly.Geometry = pGeometry
  
  'tag the graphic elements for later tracking
  Dim pElementProps As IElementProperties
780:   Set pElementProps = pElementShadowPly
781:   pElementProps.CustomProperty = "BubbleID:" & lBubbleId
782:   Set pElementProps = pElement
783:   pElementProps.CustomProperty = "BubbleID:" & lBubbleId
784:   Set pElementProps = pElementPly
785:   pElementProps.CustomProperty = "BubbleID:" & lBubbleId


  
  ''''''''''''''''''''''''''''''
  'Add triangle shadow to layout
791:   pGraphicsContainer.AddElement pElementShadowPly, 0
    
  '''''''''''''''''''''''''''
  'Add mapframe to the layout
795:   pGraphicsContainer.AddElement pElement, 0
  
  '''''''''''''''''''''''
  'Add triangle to layout
799:   pGraphicsContainer.AddElement pElementPly, 0
  
  
802:   Set pActiveView = pNewMap
  
  '''''''''''''''''''''''''''''''''''
  'Set the detail inset's zoom extent
806:   Set pNewEnv = New envelope
  
808:   If bIsCircular Then
809:     pNewEnv.XMin = dXOrigin - dRadius
810:     pNewEnv.XMax = dXOrigin + dRadius
811:   Else
812:     pNewEnv.XMin = dXOrigin - (dWidthOrigin / 2)
813:     pNewEnv.XMax = dXOrigin + (dWidthOrigin / 2)
814:   End If
815:   pNewEnv.YMin = dYOrigin - dRadius
816:   pNewEnv.YMax = dYOrigin + dRadius
    
818:   pActiveView.Extent = pNewEnv
819:   pActiveView.Refresh
  
  'Set pGraphicsLayer = pMxDoc.FocusMap.BasicGraphicsLayer
  'Set pCompositeGraphicsLayer = pGraphicsLayer
  'Set pGraphicsLayer.AssociatedLayer = pFeatureLayer
  
  'Set pGraphicsLayer = pCompositeGraphicsLayer.FindLayer("NW Detail Inset Arrows")
  'If pGraphicsLayer Is Nothing Then
  '  Set pGraphicsLayer = pCompositeGraphicsLayer.AddLayer("NW Detail Inset Arrows", pFeatureLayer)
  'End If

830:   Set pActiveView = pMxDoc.FocusMap
831:   pActiveView.PartialRefresh esriViewGraphics, Nothing, Nothing
832:   pMxDoc.CurrentContentsView.Refresh Nothing

834:   pNWSeriesOptions.BubbleGraphicAdd pElement, pElementPly, pElementShadowPly, pNewMap.Name

  Exit Sub
ErrorHandler:
  HandleError False, "AddDataframe " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub








Private Function ConvertPixelsToRW(pixelUnits As Double, pApp As IApplication) As Double
  On Error GoTo ErrorHandler
  
  Dim pMxDoc As IMxDocument
  Dim realWorldDisplayExtent As Double
  Dim pixelExtent As Long
  Dim sizeOfOnePixel As Double
  Dim pDT As IDisplayTransformation
  Dim deviceRECT As tagRECT
  Dim pEnv As IEnvelope
  Dim pActiveView As IActiveView
  
860:   Set pMxDoc = pApp.Document
861:   Set pActiveView = pMxDoc.FocusMap
862:   Set pDT = pActiveView.ScreenDisplay.DisplayTransformation
863:   deviceRECT = pDT.DeviceFrame
864:   pixelExtent = deviceRECT.Right - deviceRECT.Left
865:   Set pEnv = pDT.VisibleBounds
866:   realWorldDisplayExtent = pEnv.Width
867:   sizeOfOnePixel = realWorldDisplayExtent / pixelExtent
868:   ConvertPixelsToRW = pixelUnits * sizeOfOnePixel
  
  Exit Function
  
ErrorHandler:
873:    MsgBox "Error " & Err.Number & ": " & Err.Description & vbNewLine _
       & "In " & Err.Source & " at DrawSelectedArrows.ConvertPixelsToRW", vbCritical
End Function



