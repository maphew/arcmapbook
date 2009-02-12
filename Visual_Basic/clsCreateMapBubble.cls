VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCreateMapBubble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_pDoc As IMxDocument
Private m_pAV As IActiveView
Private m_pScrD As IScreenDisplay
Private m_pNewCircFeedback As INewCircleFeedback
Private m_p2ndCircFeedback As INewCircleFeedback
Private m_pLineFeedback As INewLineFeedback
Private m_pApp As IApplication
Private m_pMxApp As IMxApplication
Private Enum enumBubbleEditState
  BUBBLE_1_BEFORE_1ST = 1
  BUBBLE_2_1ST_CIRCLE
  BUBBLE_3_BETWEEN_LINE
  BUBBLE_4_2ND_CIRCLE
  BUBBLE_5_CREATING_BUBBLE
End Enum
Private m_eToolState As enumBubbleEditState
Private m_pEditor As IEditor
Private m_pNWSeriesOptions As INWMapSeriesOptions
Private m_pMapBook As INWDSMapBook
Private m_pMapSeries As INWDSMapSeries
Private m_pGraCont As IGraphicsContainer
Private m_pElemFillShp1 As IFillShapeElement
Private m_pElemFillShp2 As IFillShapeElement
Private m_pElemLine As ILineElement
Private m_bMouseIsDown As Boolean
Private m_pPnt1 As IPoint
Private m_pPnt2 As IPoint
Private m_pPnt3 As IPoint
Private m_pPnt4 As IPoint
Private m_pPntOrig As IPoint


''''''''''''''''''''''''----------------
'created to trap active view events so that
'the current map page is reloaded.  Without this code, refreshing
'the map will not cause the graphics for the current map page to
'be refreshed.

'Private WithEvents m_pActiveViewEvents As Map
'Private WithEvents m_pLayoutActiveViewEvents As PageLayout
''''''''''''''''''''''''----------------




Implements ICommand
Implements ITool
Const c_sModuleFileName As String = "clsCreateMapBubble.cls"





Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
57:   ICommand_Bitmap = frmResources.picCreateBubble.Picture.Handle
End Property

Private Property Get ICommand_Caption() As String
61:  ICommand_Caption = "MkBubble"
End Property

Private Property Get ICommand_Category() As String
65:   ICommand_Category = "NW MapBook"
End Property

Private Property Get ICommand_Checked() As Boolean
69:   ICommand_Checked = False
End Property





Private Function AddLineElement(pLine As IPolyline, pAV As IActiveView, pGraCont As IGraphicsContainer) As ILineElement
  On Error GoTo ErrorHandler

  Dim pElem As IElement, pLineSym As ISimpleLineSymbol, pRGB As IRgbColor
  Dim pElemLine As ILineElement
  Dim pGeom As IGeometry
  
83:   Set pElem = New LineElement
84:   Set pGeom = pLine
85:   pElem.Geometry = pGeom
86:   Set pElemLine = pElem
  
88:   Set pRGB = New RgbColor
89:   With pRGB
90:     .Red = 0
91:     .Green = 0
92:     .Blue = 0
93:   End With

95:   Set pLineSym = New SimpleLineSymbol
96:   pLineSym.Color = pRGB
97:   pLineSym.Style = esriSLSSolid
98:   pElemLine.Symbol = pLineSym
  
  ' QI for the IGraphicsContainer interface from the IActiveView, allows access to the BasicGraphicsLayer
101:   Set pGraCont = pAV
102:   Set AddLineElement = pElemLine
  
  Exit Function
ErrorHandler:
  HandleError False, "AddLineElement " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function


Private Function AddCircleElement(pCircArc As ICircularArc, pAV As IActiveView, pGraCont As IGraphicsContainer) As IFillShapeElement
  On Error GoTo ErrorHandler

  ' Takes an ICircularArc and IActiveView and creates a CircleElement in the ActiveView's BasicGraphicsLayer
  Dim pElemFillShp As IFillShapeElement
  Dim pElem As IElement
  Dim pSFSym As ISimpleFillSymbol
  Dim pRGB As IRgbColor
  Dim pSegColl As ISegmentCollection

  ' Create a new Polygon object and access the ISegmentCollection interface to add a segment
121:   Set pSegColl = New Polygon
122:   pSegColl.AddSegment pCircArc

  ' Create a new circleelement and use the IElement interface to set the its Geometry
125:   Set pElem = New CircleElement
126:   pElem.Geometry = pSegColl

  ' QI for the IFillShapeElement interface so that the Symbol property can be set
129:   Set pElemFillShp = pElem

  ' Create a new RGBColor
132:   Set pRGB = New RgbColor
133:   With pRGB
134:     .Red = 255
135:     .Green = 64
136:     .Blue = 64
137:   End With

  ' Create a new SimpleFillSymbol and set its Color and Style
140:   Set pSFSym = New SimpleFillSymbol
141:   pSFSym.Color = pRGB
142:   pSFSym.Outline.Color = pRGB
143:   pSFSym.Style = esriSFSHollow
144:   pElemFillShp.Symbol = pSFSym

  ' QI for the IGraphicsContainer interface from the IActiveView, allows access to the BasicGraphicsLayer
147:   Set pGraCont = pAV
148:   Set AddCircleElement = pElemFillShp

  Exit Function
ErrorHandler:
  HandleError False, "AddCircleElement " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function







Private Property Get ICommand_Enabled() As Boolean
  On Error GoTo ErrorHandler

  Dim pMxDoc As IMxDocument, pMap As IMap, pPageLayout As IPageLayout
165:   Static i As Long
                                                  'confirm that I'm in data view
167:   Set pMxDoc = m_pApp.Document
'  If TypeOf pMxDoc.ActiveView Is IPageLayout Then
'    ICommand_Enabled = False
'    Exit Property
'  End If
                                                  'confirm that a bubble definitions layer
                                                  'has been identified, selected, maybe as
                                                  'a property of the map series
175:   If m_pNWSeriesOptions Is Nothing Then
176:     Set m_pMapBook = GetMapBookExtension(m_pApp)
177:     If m_pMapBook Is Nothing Then
178:       ICommand_Enabled = False
      Exit Property
180:     End If
181:     Set m_pMapSeries = m_pMapBook.ContentItem(0)
182:     If m_pMapSeries Is Nothing Then
183:       ICommand_Enabled = False
      Exit Property
185:     End If
186:     Set m_pNWSeriesOptions = m_pMapSeries
187:     If m_pNWSeriesOptions Is Nothing Then
188:       ICommand_Enabled = False
      Exit Property
190:     End If
191:   End If
192:   If m_pNWSeriesOptions.BubbleLayer = "" Then
193:     ICommand_Enabled = False
    Exit Property
195:   End If
                                                  'check that the editing environment
                                                  'is active
198:   If m_pEditor.EditState <> esriStateEditing Then
199:     ICommand_Enabled = False
    Exit Property
201:   End If
  
203:   If m_pScrD Is Nothing Then
204:     Set m_pDoc = m_pApp.Document
    
    'guaranteeing that m_pAV is set to the data view.
207:     Set m_pAV = m_pDoc.FocusMap
208:     Set m_pScrD = m_pAV.ScreenDisplay
209:   End If
  
211:   ICommand_Enabled = True

  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Enabled " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Property Get ICommand_HelpContextID() As Long

End Property

Private Property Get ICommand_HelpFile() As String
  
End Property

Private Property Get ICommand_Message() As String
227:   ICommand_Message = "Create a circular map inset"
End Property

Private Property Get ICommand_Name() As String
231:   ICommand_Name = "NWMapBookUIPrj_clsCreateMapBubble"
End Property

Private Sub ICommand_OnClick()
235:   m_eToolState = BUBBLE_1_BEFORE_1ST
  
  'Shift to layout view
  '--------------------
  Dim pMxDoc As IMxDocument, pMxDocLayoutViewAV As IActiveView
  Dim pMapEnv As IEnvelope, pMxDocFocusMapQILayoutViewAV As IActiveView
241:   Set pMxDoc = m_pApp.Document
242:   Set pMxDoc.ActiveView = pMxDoc.PageLayout
243:   Set pMapEnv = pMxDoc.ActiveView.Extent
244:   Set pMxDocLayoutViewAV = pMxDoc.ActiveView
245:   Set pMxDocFocusMapQILayoutViewAV = pMxDoc.FocusMap
  
  'Set a reference to where dataframes will be inserted
248:   Set m_pGraCont = pMxDoc.PageLayout
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
  Dim pUID As New UID
  
254:   Set m_pApp = hook
255:   Set m_pMxApp = m_pApp
256:   pUID.Value = "esriEditor.Editor"
257:   Set m_pEditor = m_pApp.FindExtensionByCLSID(pUID)

259:   Set m_pMapBook = GetMapBookExtension(m_pApp)
  If m_pMapBook Is Nothing Then Exit Sub
261:   Set m_pMapSeries = m_pMapBook.ContentItem(0)
262:   Set m_pNWSeriesOptions = m_pMapSeries
End Sub

Private Property Get ICommand_Tooltip() As String
266:   ICommand_Tooltip = "Create a new circular map inset"
End Property

Private Property Get ITool_Cursor() As esriSystem.OLE_HANDLE
270:   ITool_Cursor = frmResources.picSelectCursor.Picture.Handle
End Property

Private Function ITool_Deactivate() As Boolean
  On Error GoTo ErrorHandler

  'clear all intermediate structures and states
277:   m_eToolState = BUBBLE_1_BEFORE_1ST
278:   If Not m_pElemFillShp1 Is Nothing Then
279:     m_pGraCont.DeleteElement m_pElemFillShp1
280:     Set m_pElemFillShp1 = Nothing
281:   End If
282:   If Not m_pElemFillShp1 Is Nothing Then
283:     m_pGraCont.DeleteElement m_pElemFillShp2
284:     Set m_pElemFillShp2 = Nothing
285:   End If
286:   If Not m_pElemLine Is Nothing Then
287:     m_pGraCont.DeleteElement m_pElemLine
288:     Set m_pElemLine = Nothing
289:   End If
290:   ITool_Deactivate = True

  Exit Function
ErrorHandler:
  HandleError True, "ITool_Deactivate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Function ITool_OnContextMenu(ByVal x As Long, ByVal y As Long) As Boolean
  'no right click functionality is part of the specification for this tool
End Function

Private Sub ITool_OnDblClick()
  
End Sub

Private Sub ITool_OnKeyDown(ByVal KeyCode As Long, ByVal Shift As Long)

End Sub

Private Sub ITool_OnKeyUp(ByVal KeyCode As Long, ByVal Shift As Long)
310:   If KeyCode = vbKeyEscape Then
311:     m_eToolState = BUBBLE_2_1ST_CIRCLE
                                                  'delete all temporary graphic objects
313:     If Not m_pElemFillShp1 Is Nothing Then
314:       m_pGraCont.DeleteElement m_pElemFillShp1
315:       Set m_pElemFillShp1 = Nothing
316:     End If
317:     If Not m_pElemFillShp2 Is Nothing Then
318:       m_pGraCont.DeleteElement m_pElemFillShp2
319:       Set m_pElemFillShp2 = Nothing
320:     End If
321:     If Not m_pElemLine Is Nothing Then
322:       m_pGraCont.DeleteElement m_pElemLine
323:       Set m_pElemLine = Nothing
324:     End If
325:     m_pAV.Refresh
326:   End If
End Sub




' ITool_OnMouseDown
'
''enable - editing, bubble defs layer defined
''Tool Click
''mouse down - (if in 1st circle state) start drawing circle
''mouse move - see if in start drawing state, make 1st circle
''mouse up - finish drawing 1st circle, start interim line
''mouse move - draw interim line
''mouse down - (if in 2nd circle state) start drawing 2nd circle,
''             stop and draw interim line
''mouse move - (if in 2nd circle state) make 2nd circle
''mouse up - finish drawing 2nd circle,
    ''make new bubble inset dialog appear,
    ''depending on options selected from dialog, generate
    ''new bubble definition, create a new "bubble definition" data
    ''frame.
    '-------------------------------
Private Sub ITool_OnMouseDown(ByVal button As Long, ByVal Shift As Long, ByVal x As Long, ByVal y As Long)
  On Error GoTo ErrorHandler
          
  Dim pPnt As IPoint, pPnt2 As IPoint, pPolyline As IPolyline
353:   m_bMouseIsDown = True
354:   If m_eToolState = BUBBLE_1_BEFORE_1ST Then
                                                  'Get the current mouse location in
                                                  'Map Units
357:     Set pPnt = m_pScrD.DisplayTransformation.ToMapPoint(x, y)
358:     Set m_pPnt1 = pPnt
                                                  'Check that user is not using an
                                                  'existing feedback
361:     If m_pNewCircFeedback Is Nothing Then
                                                  ' Create a NewCircleFeedback object
363:       Set m_pNewCircFeedback = New NewCircleFeedback
                                                  'Set the Feedback's Display and StartPoint
365:       Set m_pNewCircFeedback.Display = m_pScrD
366:       m_pNewCircFeedback.Start pPnt
                                                  'save the origin point for the
                                                  'feedback line
369:       Set m_pPntOrig = Nothing
370:       Set m_pPntOrig = pPnt
371:     End If
372:     m_eToolState = BUBBLE_2_1ST_CIRCLE
373:   ElseIf m_eToolState = BUBBLE_3_BETWEEN_LINE Then
                                                  'finish drawing the feedback line
375:     If Not m_pLineFeedback Is Nothing Then
376:       Set pPnt = m_pScrD.DisplayTransformation.ToMapPoint(x, y)
377:       Set m_pPnt3 = pPnt
378:       m_pLineFeedback.AddPoint pPnt
379:       Set pPolyline = m_pLineFeedback.Stop
380:       If Not pPolyline Is Nothing Then
381:         Set m_pElemLine = AddLineElement(pPolyline, m_pAV, m_pGraCont)
382:         m_pGraCont.AddElement m_pElemLine, 0
383:         m_pAV.Refresh
384:       End If
385:     End If
                                                  'start drawing the 2nd circle
387:     Set pPnt2 = m_pScrD.DisplayTransformation.ToMapPoint(x, y)
388:     If m_p2ndCircFeedback Is Nothing Then
389:       Set m_p2ndCircFeedback = New NewCircleFeedback
390:       Set m_p2ndCircFeedback.Display = m_pScrD
391:     End If
392:     m_p2ndCircFeedback.Start pPnt2
    
394:     m_eToolState = BUBBLE_4_2ND_CIRCLE
395:   End If


  Exit Sub
ErrorHandler:
  HandleError True, "ITool_OnMouseDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub ITool_OnMouseMove(ByVal button As Long, ByVal Shift As Long, ByVal x As Long, ByVal y As Long)
  On Error GoTo ErrorHandler

                                                  'adjust the circular or linear feedback based
                                                  'on the state that the tool is in
  Dim pPnt As IPoint
  
410:   If m_eToolState = BUBBLE_2_1ST_CIRCLE Then
411:     If m_bMouseIsDown Then
412:       If Not m_pNewCircFeedback Is Nothing Then
        'Move the Feedback to the current mouse location
414:         Set pPnt = m_pScrD.DisplayTransformation.ToMapPoint(x, y)
415:         m_pNewCircFeedback.MoveTo pPnt
416:       End If
417:     End If
418:   ElseIf m_eToolState = BUBBLE_3_BETWEEN_LINE Then
419:     If Not m_bMouseIsDown Then
      'extend the feedback line
421:       If Not m_pLineFeedback Is Nothing Then
422:         Set pPnt = m_pScrD.DisplayTransformation.ToMapPoint(x, y)
423:         m_pLineFeedback.MoveTo pPnt
424:       End If
425:     End If
426:   ElseIf m_eToolState = BUBBLE_4_2ND_CIRCLE Then
427:     If Not m_p2ndCircFeedback Is Nothing Then
428:       Set pPnt = m_pScrD.DisplayTransformation.ToMapPoint(x, y)
429:       m_p2ndCircFeedback.MoveTo pPnt
430:     End If
431:   End If
  Exit Sub
ErrorHandler:
  HandleError True, "ITool_OnMouseMove " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub


Private Sub ITool_OnMouseUp(ByVal button As Long, ByVal Shift As Long, ByVal x As Long, ByVal y As Long)
  On Error GoTo ErrorHandler
          
  Dim pCircArc As IGeometry, pPnt As IPoint, pTable As ITable
  Dim pLayers As IEnumLayer, pLayer As ILayer, lXDestIdx As Long, lYDestIdx As Long
  Dim lXOrgIdx As Long, lYOrgIdx As Long, pFields As IFields
  Dim lRadiusIdx As Long, lScaleidx As Long, lBubbleIdIdx As Long, pField As IField
  Dim bBubbleLayerIsFound As Boolean, lFieldCount As Long, i As Long
  Dim pDataset As IDataset, pFC As IFeatureClass, pFeatLayer As IFeatureLayer
  Dim pFeatClass As IFeatureClass, pRow As IRow, pBubbleLayer As ILayer
  Dim pFeature As IFeature, pRowBuffer As IRowBuffer, lLayersIdx As Long
  Dim lWidthOrgIdx As Long
  Dim sActiveDataFrameName As String, pElement As IElement
  Dim lFeatCount As Long, featCountQueryFilter As IQueryFilter
  
453:   m_bMouseIsDown = False
                                                  ' Check if the user is currently using the
                                                  'feedback
456:   If m_eToolState = BUBBLE_2_1ST_CIRCLE Then
457:     If Not m_pNewCircFeedback Is Nothing Then
458:       Set m_pPnt2 = m_pScrD.DisplayTransformation.ToMapPoint(x, y)
                                                  'Get the geometry (ICircularArc) returned
                                                  'from the feedback
461:       Set pCircArc = m_pNewCircFeedback.Stop
      
                                                  'If it is valid then create a CircleElement
                                                  'on the ActiveView using the 'AddCircleElement'
                                                  'procedure
466:       If Not pCircArc Is Nothing Then
467:         Set m_pElemFillShp1 = AddCircleElement(pCircArc, m_pAV, m_pGraCont)
468:         m_pGraCont.AddElement m_pElemFillShp1, 0
469:         m_pAV.Refresh
470:       End If
                                                  ' Set the feedback to nothing for
                                                  'the next use
473:       Set m_pNewCircFeedback = Nothing
474:     End If
                                                  'set the feedback line
476:     Set pPnt = m_pPntOrig
477:     If m_pLineFeedback Is Nothing Then
478:       Set m_pLineFeedback = New NewLineFeedback
479:     End If
480:     Set m_pLineFeedback.Display = m_pScrD
481:     m_pLineFeedback.Start pPnt
482:     m_eToolState = BUBBLE_3_BETWEEN_LINE
  
484:   ElseIf m_eToolState = BUBBLE_4_2ND_CIRCLE Then
485:     Set m_pPnt4 = m_pScrD.DisplayTransformation.ToMapPoint(x, y)
486:     If Not m_p2ndCircFeedback Is Nothing Then
487:       Set pCircArc = m_p2ndCircFeedback.Stop
488:       If Not pCircArc Is Nothing Then
489:         Set m_pElemFillShp2 = AddCircleElement(pCircArc, m_pAV, m_pGraCont)
490:         m_pGraCont.AddElement m_pElemFillShp2, 0
491:         m_pAV.Refresh
492:       End If
493:       Set m_p2ndCircFeedback = Nothing
494:     End If
  
496:     m_eToolState = BUBBLE_5_CREATING_BUBBLE
    
498:     Set pLayers = m_pDoc.FocusMap.Layers
                                                   'access the bubble layer
500:     Set pLayer = pLayers.Next
501:     bBubbleLayerIsFound = False
502:     Do While (Not pLayer Is Nothing) And (Not bBubbleLayerIsFound)
503:       If TypeOf pLayer Is IFeatureLayer Then
504:         Set pFeatLayer = pLayer
505:         If m_pNWSeriesOptions.IsBubbleLayer(pFeatLayer.FeatureClass) Then
506:           If StrComp(m_pNWSeriesOptions.BubbleLayer, pLayer.Name, vbTextCompare) = 0 Then
507:             bBubbleLayerIsFound = True
508:             Set pBubbleLayer = pLayer
509:           End If
510:         End If
511:       End If
512:       Set pLayer = pLayers.Next
513:     Loop
    
    If Not bBubbleLayerIsFound Then Exit Sub
                                                  ' using the list of layer names provided,  acquire the
                                                  ' field index where these values can be accessed for
                                                  ' either display to the user and/or knowing where to
                                                  ' write the results from the new bubble properties dialog
520:     Set pFeatLayer = pBubbleLayer
521:     Set pFields = pFeatLayer.FeatureClass.Fields
522:     lFieldCount = pFields.FieldCount
    
524:     For i = 0 To lFieldCount - 1
525:       Set pField = pFields.Field(i)
      Select Case pField.Name
      Case "BUBBLEID"
528:         lBubbleIdIdx = i
      Case "XORG"
530:         lXOrgIdx = i
      Case "YORG"
532:         lYOrgIdx = i
      Case "XDEST"
534:         lXDestIdx = i
      Case "YDEST"
536:         lYDestIdx = i
      Case "RADIUS"
538:         lRadiusIdx = i
      Case "SCALE"
540:         lScaleidx = i
      Case "LAYERS"
542:         lLayersIdx = i
      Case "WIDTHORG"
544:         lWidthOrgIdx = i
545:       End Select
546:     Next i
                        ' - radius - read/write (passed to form, can be modified)
                        '        --  acquire this value from the created feature
                        ' - scale factor - read/write
                        '        --  calculate this value from ratio of the origin vs
                        '            destination radii / map scale
                        '            Example: in a bubble with twice the radius of an
                        '            area represented, and the map is at the scale of
                        '            1:400, the bubble scale should be 1:200
                        ' - bubble radius - read/write
                        '        --  acquire this from the created graphic
                        ' - origin x/y - read/write
                        '        --  acquire this from the 1st click point
                        ' - destination x/y - read/write
                        '        --  acquire this from the 2nd click point
                        ' - destination height/width - read/write
                        '        --  set initial circular values from 2xradius
                        ' - application reference, (to add all layers) - read only
    
    ''''''''' Load the form '''''''''''
    '---------------------------------'
                                                  ' populate the bubble ID field,
                                                  ' pass that bubble ID field to the form
569:     Set pField = pFields.Field(lBubbleIdIdx)
570:     Set pFeatLayer = pBubbleLayer
571:     Set pFeatClass = pFeatLayer.FeatureClass
572:     Set featCountQueryFilter = New QueryFilter
573:     featCountQueryFilter.WhereClause = ""
574:     lFeatCount = pFeatClass.FeatureCount(featCountQueryFilter)
575:     Set pTable = pFeatClass
576:     Set pFeature = pFeatClass.CreateFeature
577:     Set pRow = pFeature
    'reed todo - this next line breaks if there are no pre-existing detail insets
579:     If lFeatCount > 0 Then
580:       pRow.Value(lBubbleIdIdx) = NextBubbleId(pBubbleLayer, "BUBBLEID")
581:     Else
582:       pRow.Value(lBubbleIdIdx) = 0
583:     End If
    
585:     With frmBubbleProperties
586:       .Initializing = True
587:       Set .Application = m_pApp
588:       .HeightWidthIsEnabled = False
589:       .Radius = m_pDoc.FocusMap.ComputeDistance(m_pPnt1, m_pPnt2)
      
591:       If .Radius = 0 Then GoTo GracefulExit
592:       .BubbleRadius = m_pDoc.FocusMap.ComputeDistance(m_pPnt3, m_pPnt4)
593:       .BubbleRadiusX = .BubbleRadius
      
595:       .BubbleID = pRow.Value(lBubbleIdIdx)
596:       .ScaleFactor = (.BubbleRadius / .Radius)
597:       .OriginX = m_pPnt1.x
598:       .OriginY = m_pPnt1.y
599:       .DestinationX = m_pPnt3.x
600:       .DestinationY = m_pPnt3.y
601:       .SourceHeight = .Radius + .Radius
602:       .SourceWidth = .SourceHeight
603:       .Initializing = False
    
      'add graphics components to form so that these graphics can be
      'updated based on user input in the dialog
607:       Set .OriginShape = m_pElemFillShp1
608:       Set .DestinationShape = m_pElemFillShp2
609:       Set .BetweenLineElement = m_pElemLine
610:       Set .GraphicsContainer = m_pGraCont
611:       Set .PointOriginCenter = m_pPnt1
612:       Set .PointOriginEdge = m_pPnt2
613:       Set .PointDestCenter = m_pPnt3
614:       Set .PointDestEdge = m_pPnt4
615:       .DataFrameName = m_pNWSeriesOptions.DataFrameMainFrame
616:     End With
    
                                                  'clean up all the temporary graphics
619:         If Not m_pElemFillShp1 Is Nothing Then
620:           m_pGraCont.DeleteElement m_pElemFillShp1
621:           Set m_pElemFillShp1 = Nothing
622:         End If
623:         If Not m_pElemFillShp2 Is Nothing Then
624:           m_pGraCont.DeleteElement m_pElemFillShp2
625:           Set m_pElemFillShp2 = Nothing
626:         End If
627:         If Not m_pElemLine Is Nothing Then
628:           m_pGraCont.DeleteElement m_pElemLine
629:           Set m_pElemLine = Nothing
630:         End If
    
                                                  'show the dialog
633:     m_pNWSeriesOptions.RefreshEventShouldntRefresh = True
634:     frmBubbleProperties.Show vbModal
                                                  'all these values from the form
                                                  'are stored in a data layer, not
                                                  'in any data structure to be stored
                                                  'in the map document.
    'idea - support the edit environment's undo/redo -- do
    '       this after other functions are implemented
    
                                                  'apply the results of the dialog
                                                  'user input
644:     With frmBubbleProperties
645:       If .WasCancelled Then
646:         Set pRow = Nothing
                                                  'clean up all the temporary graphics
648:         m_pAV.Refresh
649:         m_eToolState = BUBBLE_1_BEFORE_1ST
650:         m_pNWSeriesOptions.RefreshEventShouldntRefresh = False
        Exit Sub
652:       End If
      
654:       m_pEditor.StartOperation
      
656:       pRow.Value(lXOrgIdx) = .OriginX
657:       pRow.Value(lYOrgIdx) = .OriginY
658:       pRow.Value(lXDestIdx) = .DestinationX
659:       pRow.Value(lYDestIdx) = .DestinationY
660:       pRow.Value(lRadiusIdx) = .Radius
661:       pRow.Value(lScaleidx) = .ScaleFactor
662:       pRow.Value(lLayersIdx) = .Layers
663:       pRow.Value(lWidthOrgIdx) = .SourceWidth
664:     End With
    
                                                  'Create the polygon.  Use an oval shaped
                                                  'polygon to support possible enhancement
                                                  'of non-circular insets
669:     With frmBubbleProperties

671:       Set pFeature.Shape = CreateInsetPolygon(.OriginX, .OriginY, .SourceHeight, .SourceWidth)
672:       pFeature.Store
673:       m_pEditor.StopOperation "NWCreateBubble"
674:     End With
    
676:     Set pElement = GetDataFrameElement(m_pNWSeriesOptions.DataFrameMainFrame, m_pApp)
677:     AddDataframe pElement, m_pApp, pRow, m_pNWSeriesOptions
678:     m_pNWSeriesOptions.RefreshEventShouldntRefresh = False
    
'                                                  'clean up all the temporary graphics
'    If Not m_pElemFillShp1 Is Nothing Then
'      m_pGraCont.DeleteElement m_pElemFillShp1
'      Set m_pElemFillShp1 = Nothing
'    End If
'    If Not m_pElemFillShp2 Is Nothing Then
'      m_pGraCont.DeleteElement m_pElemFillShp2
'      Set m_pElemFillShp2 = Nothing
'    End If
'    If Not m_pElemLine Is Nothing Then
'      m_pGraCont.DeleteElement m_pElemLine
'      Set m_pElemLine = Nothing
'    End If
    
694:     m_pAV.Refresh
695:     m_eToolState = BUBBLE_1_BEFORE_1ST
696:   Else

698:     MsgBox "ITool_OnMouseUp -- code block that I suspect is unnecessary, " & vbNewLine _
         & "so therefore this message shouldn't appear.  If it does, please" & vbNewLine _
         & "pay attention to the steps necessary to reproduce this, and " & vbNewLine _
         & "insert code at the end of the clsCreateMapBubble.ITool_OnMouseUp " & vbNewLine _
         & "function to handle this scenario." & vbNewLine _
         & "The likely cause of this error is if the state of m_eToolState was" & vbNewLine _
         & "not handled correctly." & vbNewLine _
         & "m_eToolState is " & m_eToolState & vbNewLine _
         & "BUBBLE_1_BEFORE_1ST is " & BUBBLE_1_BEFORE_1ST & vbNewLine _
         & "BUBBLE_2_1ST_CIRCLE is " & BUBBLE_2_1ST_CIRCLE & vbNewLine _
         & "BUBBLE_3_BETWEEN_LINE is " & BUBBLE_2_1ST_CIRCLE & vbNewLine _
         & "BUBBLE_4_2ND_CIRCLE is " & BUBBLE_2_1ST_CIRCLE & vbNewLine _
         & "BUBBLE_5_CREATING_BUBBLE is " & BUBBLE_2_1ST_CIRCLE & vbNewLine
711:   End If

  Exit Sub
GracefulExit:
                                                  'clean up all the temporary graphics
716:     If Not m_pElemFillShp1 Is Nothing Then
717:       m_pGraCont.DeleteElement m_pElemFillShp1
718:       Set m_pElemFillShp1 = Nothing
719:     End If
720:     If Not m_pElemFillShp2 Is Nothing Then
721:       m_pGraCont.DeleteElement m_pElemFillShp2
722:       Set m_pElemFillShp2 = Nothing
723:     End If
724:     If Not m_pElemLine Is Nothing Then
725:       m_pGraCont.DeleteElement m_pElemLine
726:       Set m_pElemLine = Nothing
727:     End If
    
729:     m_pAV.Refresh
730:     m_eToolState = BUBBLE_1_BEFORE_1ST

  Exit Sub
ErrorHandler:
                                                  'clean up all the temporary graphics
735:     If Not m_pElemFillShp1 Is Nothing Then
736:       m_pGraCont.DeleteElement m_pElemFillShp1
737:       Set m_pElemFillShp1 = Nothing
738:     End If
739:     If Not m_pElemFillShp2 Is Nothing Then
740:       m_pGraCont.DeleteElement m_pElemFillShp2
741:       Set m_pElemFillShp2 = Nothing
742:     End If
743:     If Not m_pElemLine Is Nothing Then
744:       m_pGraCont.DeleteElement m_pElemLine
745:       Set m_pElemLine = Nothing
746:     End If
    
748:     m_pAV.Refresh
749:     m_eToolState = BUBBLE_1_BEFORE_1ST
750:     m_pNWSeriesOptions.RefreshEventShouldntRefresh = False
751:     m_pEditor.AbortOperation
  HandleError True, "ITool_OnMouseUp " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub ITool_Refresh(ByVal hdc As esriSystem.OLE_HANDLE)
  On Error GoTo ErrorHandler

  Exit Sub
ErrorHandler:
  HandleError True, "ITool_Refresh " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub













Private Function CloneMe(ByVal pCloner As esriSystem.IClone) As esriSystem.IClone
776:   If Not pCloner Is Nothing Then
777:     Set CloneMe = pCloner.Clone
778:   End If
End Function



Private Function CalculatePageToMapRatio(pApp As IApplication) As Double
    Dim pMx As IMxDocument
    Dim pPage As IPage
    Dim pPageUnits As esriUnits
    Dim pSR As ISpatialReference
    Dim pSRI As ISpatialReferenceInfo
    Dim pPCS As IProjectedCoordinateSystem
    Dim dMetersPerUnit As Double
    
    On Error GoTo eh
    
    ' Init
795:     Set pMx = pApp.Document
796:     Set pSR = pMx.FocusMap.SpatialReference
797:     If TypeOf pSR Is IProjectedCoordinateSystem Then
798:         Set pPCS = pSR
799:         dMetersPerUnit = pPCS.CoordinateUnit.MetersPerUnit
800:     Else
801:         dMetersPerUnit = 1
802:     End If
803:     Set pPage = pMx.PageLayout.Page
804:     pPageUnits = pPage.Units
    Select Case pPageUnits
        Case esriInches: CalculatePageToMapRatio = dMetersPerUnit / (1 / 12 * 0.304800609601219)
        Case esriFeet: CalculatePageToMapRatio = dMetersPerUnit / (0.304800609601219)
        Case esriCentimeters: CalculatePageToMapRatio = dMetersPerUnit / (1 / 100)
        Case esriMeters: CalculatePageToMapRatio = dMetersPerUnit / (1)
        Case Else:
811:             MsgBox "Warning: Only the following Page (Layout) Units are supported by this tool:" _
                & vbCrLf & " - Inches, Feet, Centimeters, Meters" _
                & vbCrLf & vbCrLf & "Calculating as though Page Units are in Inches..."
814:             CalculatePageToMapRatio = dMetersPerUnit / (1 / 12 * 0.304800609601219)
815:     End Select
    Exit Function
eh:
818:     CalculatePageToMapRatio = 1
819:     MsgBox "Error in CalculatePageToMapRatio" & vbCrLf & Err.Description
End Function







  






' NextBubbleId will query the field name from the layer provided,
' will acquire the numbers stored there, and will return the maximum
' value in that field + 1
'-----------------------------
Private Function NextBubbleId(pLayer As ILayer, sFieldName) As Long
  On Error GoTo ErrorHandler

  Dim pFeatLayer As IFeatureLayer, pFeatClass As IFeatureClass
  Dim pCursor As ICursor, lResult As Long, pStatResults As IStatisticsResults
  Dim pDataStats As IDataStatistics
  
  If pLayer Is Nothing Then Exit Function
847:   Set pFeatLayer = pLayer
848:   If Not TypeOf pLayer Is IFeatureLayer Then
849:     NextBubbleId = -1
    Exit Function
851:   End If
  
853:   Set pFeatClass = pFeatLayer.FeatureClass
854:   Set pCursor = pFeatClass.Search(Nothing, True)
855:   Set pDataStats = New DataStatistics
856:   pDataStats.Field = sFieldName
857:   Set pDataStats.Cursor = pCursor
858:   Set pStatResults = pDataStats.Statistics
859:   lResult = pStatResults.Maximum
860:   NextBubbleId = (lResult + 1)

  Exit Function
ErrorHandler:
  HandleError False, "NextBubbleId " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function













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
911:   Set pMxDoc = pApp.Document
  
913:   Set pFields = pRow.Fields
914:   lFieldCount = pFields.FieldCount
    
916:   For i = 0 To lFieldCount - 1
917:     Set pField = pFields.Field(i)
    Select Case pField.Name
    Case "BUBBLEID"
920:       lBubbleId = pRow.Value(i)
    Case "XORG"
922:       dXOrigin = pRow.Value(i)
    Case "YORG"
924:       dYOrigin = pRow.Value(i)
    Case "XDEST"
926:       dXDestination = pRow.Value(i)
    Case "YDEST"
928:       dYDestination = pRow.Value(i)
    Case "RADIUS"
930:       dRadius = pRow.Value(i)
    Case "SCALE"
932:       dScale = pRow.Value(i)
    Case "LAYERS"
934:       sLayers = pRow.Value(i)
    Case "WIDTHORG"
936:       dWidthOrigin = pRow.Value(i)
937:     End Select
938:   Next i

940:   Set pNewMap = New Map
941:   pNewMap.Name = "BubbleID:" & lBubbleId
942:   pNewMap.Description = "Detail Inset " & lBubbleId
  
  
  ' clone all copies of layers so that modifying
  ' layers in one dataframe doesn't impact layers
  ' referenced in the other data frames.
  ''''''''''''''''''''''
  Dim pLayer As ILayer, pLayerSrc As ILayer
  Dim pFeatLyr As IFeatureLayer, pTinLyr As ITinLayer, pRastLyrSrc As IRasterLayer
  Dim pRastLyr As IRasterLayer, pFeatLyrSrc As IFeatureLayer, pTinLyrSrc As ITinLayer
  Dim pGeoFeatLyrSource As IGeoFeatureLayer, pGeoFeatLyrDestination As IGeoFeatureLayer
  
954:   If pMxDoc.FocusMap.LayerCount > 0 Then
955:     For i = (pMxDoc.FocusMap.LayerCount - 1) To 0 Step -1
956:       Set pLayerSrc = pMxDoc.FocusMap.Layer(i)
957:       If TypeOf pMxDoc.FocusMap.Layer(i) Is IFeatureLayer Then
958:         Set pFeatLyr = New FeatureLayer
959:         Set pFeatLyrSrc = pLayerSrc
960:         pFeatLyr.DataSourceType = pFeatLyrSrc.DataSourceType
961:         pFeatLyr.DisplayField = pFeatLyrSrc.DisplayField
962:         Set pFeatLyr.FeatureClass = pFeatLyrSrc.FeatureClass
963:         pFeatLyr.ScaleSymbols = pFeatLyrSrc.ScaleSymbols
964:         pFeatLyr.Selectable = pFeatLyrSrc.Selectable
965:         Set pGeoFeatLyrSource = pFeatLyrSrc
966:         Set pGeoFeatLyrDestination = pFeatLyr
967:         With pGeoFeatLyrSource
                                            'Set pGeoFeatLyrDestination.CurrentMapLevel = .CurrentMapLevel
                                            'pGeoFeatLyrDestination.DisplayFeatureClass = .DisplayFeatureClass
970:           pGeoFeatLyrDestination.AnnotationProperties = .AnnotationProperties
971:           pGeoFeatLyrDestination.AnnotationPropertiesID = .AnnotationPropertiesID
972:           pGeoFeatLyrDestination.DisplayAnnotation = .DisplayAnnotation
973:           Set pGeoFeatLyrDestination.ExclusionSet = .ExclusionSet
974:           Set pGeoFeatLyrDestination.Renderer = .Renderer
975:         End With
976:         Set pLayer = pFeatLyr
        
978:       ElseIf TypeOf pMxDoc.FocusMap.Layer(i) Is ITinLayer Then
                                            'pTinLyr.RendererCount = pTinLyrSrc.RendererCount
980:         Set pTinLyrSrc = pLayerSrc
981:         Set pTinLyr = New TinLayer
982:         Set pTinLyr.Dataset = pTinLyrSrc.Dataset
983:         pTinLyr.DisplayField = pTinLyrSrc.DisplayField
984:         pTinLyr.ScaleSymbols = pTinLyrSrc.ScaleSymbols
985:         Set pLayer = pTinLyr
986:       ElseIf TypeOf pMxDoc.FocusMap.Layer(i) Is IRasterLayer Then
                                            'pRastLyr.BandCount = pRastLyrSrc.BandCount
                                            'pRastLyr.ColumnCount = pRastLyrSrc.ColumnCount
                                            'pRastLyr.DataFrameExtent = pRastLyrSrc.DataFrameExtent
                                            'pRastLyr.FilePath = pRastLyrSrc.FilePath
                                            'pRastLyr.Raster = pRastLyrSrc.Raster
                                            'pRastLyr.RowCount = pRastLyrSrc.RowCount
993:         Set pRastLyr = New RasterLayer
994:         Set pRastLyrSrc = pLayerSrc
995:         pRastLyr.DisplayResolutionFactor = pRastLyrSrc.DisplayResolutionFactor
996:         pRastLyr.PrimaryField = pRastLyrSrc.PrimaryField
997:         pRastLyr.PyramidPresent = pRastLyrSrc.PyramidPresent
998:         Set pRastLyr.Renderer = pRastLyrSrc.Renderer
999:         pRastLyr.ShowResolution = pRastLyrSrc.ShowResolution
1000:         pRastLyr.VisibleExtent = pRastLyrSrc.VisibleExtent
1001:         Set pLayer = pRastLyr
1002:       End If
                                            'Set pLayer.AreaOfInterest = pLayerSrc.AreaOfInterest
                                            'Set pLayer.SpatialReference = pLayerSrc.SpatialReference
                                            'pLayer.SupportedDrawPhases = pLayerSrc.SupportedDrawPhases
                                            'pLayer.TipText = pLayerSrc.TipText
                                            'pLayer.Valid = pLayerSrc.Valid
1008:       pLayer.Cached = pLayerSrc.Cached
1009:       pLayer.MaximumScale = pLayerSrc.MaximumScale
1010:       pLayer.MinimumScale = pLayerSrc.MinimumScale
1011:       pLayer.Name = pLayerSrc.Name
1012:       pLayer.ShowTips = pLayerSrc.ShowTips
1013:       pLayer.Visible = pLayerSrc.Visible
      
1015:       pNewMap.AddLayer pLayer
1016:     Next i
1017:   End If
1018:   SetVisibleLayers pNewMap, sLayers
                                                  'Create a new MapFrame and associate
                                                  'map with it
1021:   Set pMapFrame = New MapFrame
1022:   Set pMapFrame.Map = pNewMap
1023:   pMapFrame.ExtentType = esriExtentDefault
  
1025:   Set pGraphicsContainer = pMxDoc.PageLayout
                                                  'Set the position of the new map frame
1027:   Set pElement = pMapFrame
1028:   Set pEnv = New envelope
1029:   Set pMxDocDataViewAV = pMxDoc.ActiveView
1030:   Set pMxDocFocusMapQIDataViewAV = pMxDoc.FocusMap
  
1032:   bIsCircular = ((2 * dRadius) = dWidthOrigin)
  
  
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
  
  
1047:   dBubbleRadius = dRadius * dScale
1048:   dBubbleWidth = dWidthOrigin * dScale
                                                  'generate the top y
                                                  'generate the bottom y
                                                  'the left x and right x
                                                  'for detail insets
1053:   dBubbleTop = dYDestination + dBubbleRadius
1054:   dBubbleBottom = dYDestination - dBubbleRadius
1055:   If bIsCircular Then
1056:     dBubbleLeft = dXDestination - dBubbleRadius
1057:     dBubbleRight = dXDestination + dBubbleRadius
1058:   Else
1059:     dBubbleLeft = dXDestination - (dBubbleWidth / 2)
1060:     dBubbleRight = dXDestination + (dBubbleWidth / 2)
1061:   End If
  
1063:   Set pBubbleEnv = New envelope
1064:   Set pMapEnv = pMxDoc.ActiveView.Extent 'sides of display in screen units
  
  
  
  'Shift to layout view
  '--------------------
1070:   Set pMxDoc.ActiveView = pMxDoc.PageLayout
1071:   Set pMapEnv = pMxDoc.ActiveView.Extent
1072:   Set pMxDocLayoutViewAV = pMxDoc.ActiveView
1073:   Set pMxDocFocusMapQILayoutViewAV = pMxDoc.FocusMap
  
  
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
  
1087:   With pMapEnv
1088:     dDocWidthInches = .XMax - (Abs(.XMin))
1089:     dDocHeightInches = .YMax - (Abs(.YMin))
1090:   End With
                                                  'get the data area width minus the
                                                  'buffer of empty space around the
                                                  'map display
1094:   dDataFrameWidth = pElementWithinBorders.Geometry.envelope.Width
1095:   dDataFrameHeight = pElementWithinBorders.Geometry.envelope.Height
1096:   dBoundaryInchesXLeft = pElementWithinBorders.Geometry.envelope.XMin
1097:   dBoundaryInchesYTop = pElementWithinBorders.Geometry.envelope.YMin
  
1099:   With pMxDocFocusMapQILayoutViewAV.ScreenDisplay.DisplayTransformation
1100:     Set pPntScrBottomLeft = .ToMapPoint(.DeviceFrame.Left, .DeviceFrame.bottom)
1101:     Set pPntScrTopRight = .ToMapPoint(.DeviceFrame.Right, .DeviceFrame.Top)
1102:   End With
  
1104:   dScrMapUnitWidth = pPntScrTopRight.x - pPntScrBottomLeft.x
1105:   dScrMapUnitHeight = pPntScrTopRight.y - pPntScrBottomLeft.y
1106:   dRatioInchToMapUnitX = dDataFrameWidth / dScrMapUnitWidth
1107:   dRatioInchToMapUnitY = dDataFrameHeight / dScrMapUnitHeight
  
1109:   dBubbleLeftInches = ((dBubbleLeft - pPntScrBottomLeft.x) * dRatioInchToMapUnitX) + dBoundaryInchesXLeft
1110:   dBubbleRightInches = ((dBubbleRight - pPntScrBottomLeft.x) * dRatioInchToMapUnitX) + dBoundaryInchesXLeft
1111:   dBubbleTopInches = ((dBubbleTop - pPntScrBottomLeft.y) * dRatioInchToMapUnitY) + dBoundaryInchesYTop
1112:   dBubbleBottomInches = ((dBubbleBottom - pPntScrBottomLeft.y) * dRatioInchToMapUnitY) + dBoundaryInchesYTop
  
1114:   pBubbleEnv.XMin = dBubbleLeftInches
1115:   pBubbleEnv.XMax = dBubbleRightInches
1116:   pBubbleEnv.YMin = dBubbleBottomInches
1117:   pBubbleEnv.YMax = dBubbleTopInches
  
  
  
1121:   pElement.Geometry = pBubbleEnv
  
  '''''''''''''''''''''''
  ' data frame properties
  
  Dim pColor As IColor, pFillColor As IColor, pShadowColor As IColor
1127:   Set pColor = New RgbColor
1128:   Set pFillColor = New RgbColor
1129:   Set pShadowColor = New RgbColor
1130:   pColor.RGB = RGB(210, 210, 210)
1131:   pFillColor.RGB = RGB(255, 255, 255)
1132:   pShadowColor.RGB = RGB(128, 128, 128)
    
1134:   Set pFrameElement = pElement
  
  ' create a border
  Dim pSymbolBorder As ISymbolBorder
  Dim pLineSymbol As ILineSymbol
  Dim pFrameDecoration As IFrameDecoration
  Dim pShadowFillSymbol As IFillSymbol
  Dim pSymbolShadow As ISymbolShadow
  Dim pFrameProperties As IFrameProperties
  
1144:   Set pSymbolBorder = New SymbolBorder
1145:   Set pLineSymbol = New SimpleLineSymbol
1146:   pLineSymbol.Color = pColor
1147:   pSymbolBorder.LineSymbol = pLineSymbol
1148:   pSymbolBorder.LineSymbol.Color = pColor
1149:   pSymbolBorder.CornerRounding = 100
1150:   pFrameElement.Border = pSymbolBorder
  
  'modify the frame element background
1153:   Set pFrameDecoration = New SymbolBackground
1154:   pFrameDecoration.Color = pFillColor
1155:   pFrameDecoration.CornerRounding = 100
  
1157:   pFrameElement.Background = pFrameDecoration
  
  
  ' add shadow to detail inset
1161:   Set pShadowFillSymbol = New SimpleFillSymbol
1162:   pShadowFillSymbol.Color = pShadowColor
1163:   pShadowFillSymbol.Outline.Color = pShadowColor
1164:   Set pSymbolShadow = New SymbolShadow
1165:   pSymbolShadow.FillSymbol = pShadowFillSymbol
1166:   pSymbolShadow.HorizontalSpacing = -2
1167:   pSymbolShadow.VerticalSpacing = -2
1168:   pSymbolShadow.CornerRounding = 100
1169:   Set pFrameProperties = pFrameElement
1170:   pFrameProperties.Shadow = pSymbolShadow

  
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
  
1191:   Set pLine = New esriGeometry.Line
1192:   Set pFromPnt = New Point
1193:   Set pToPnt = New Point
1194:   pFromPnt.x = dXDestination
1195:   pFromPnt.y = dYDestination
1196:   pToPnt.x = dXOrigin
1197:   pToPnt.y = dYOrigin
1198:   pLine.PutCoords pFromPnt, pToPnt
  
  
1201:   Set pCircularArc = New CircularArc
                                                  'angles are stored in radians,
                                                  'so calculate 10 degrees in radians
1204:   pCircularArc.PutCoordsByAngle pFromPnt, _
                               (pLine.Angle - ((10 / 180) * 3.14159265358979)), _
                               ((20 / 180) * 3.14159265358979), _
                               dBubbleRadius
  
1209:   Set pBaseLine = New esriGeometry.Line
1210:   pBaseLine.PutCoords pCircularArc.FromPoint, pCircularArc.ToPoint
1211:   Set p1stPnt = New Point
1212:   Set p2ndPnt = New Point
1213:   Set p3rdPnt = New Point
1214:   Set p1stShadowPnt = New Point
1215:   Set p2ndShadowPnt = New Point
1216:   Set p3rdShadowPnt = New Point
1217:   Set pConstructPoint = p3rdPnt
1218:   pConstructPoint.ConstructDeflection pBaseLine, _
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
  
1235:   Set pArrowFillSymb = New SimpleFillSymbol
1236:   Set pPolygonElement = New PolygonElement
1237:   Set pFillShapeElement = pPolygonElement
1238:   pArrowFillSymb.Outline = pLineSymbol
1239:   pArrowFillSymb.Color = pLineSymbol.Color
1240:   pFillShapeElement.Symbol = pArrowFillSymb
1241:   Set pElementPly = pPolygonElement
  
1243:   Set pTrianglePoly = New esriGeometry.Polygon
1244:   Set pPntColl = pTrianglePoly
1245:   p1stPnt.x = ((pCircularArc.FromPoint.x - pPntScrBottomLeft.x) * dRatioInchToMapUnitX) + dBoundaryInchesXLeft
1246:   p1stPnt.y = ((pCircularArc.FromPoint.y - pPntScrBottomLeft.y) * dRatioInchToMapUnitX) + dBoundaryInchesYTop
1247:   p2ndPnt.x = ((pCircularArc.ToPoint.x - pPntScrBottomLeft.x) * dRatioInchToMapUnitX) + dBoundaryInchesXLeft
1248:   p2ndPnt.y = ((pCircularArc.ToPoint.y - pPntScrBottomLeft.y) * dRatioInchToMapUnitX) + dBoundaryInchesYTop
1249:   p3rdPnt.x = ((p3rdPnt.x - pPntScrBottomLeft.x) * dRatioInchToMapUnitX) + dBoundaryInchesXLeft
1250:   p3rdPnt.y = ((p3rdPnt.y - pPntScrBottomLeft.y) * dRatioInchToMapUnitX) + dBoundaryInchesYTop
1251:   pPntColl.AddPoint p1stPnt
1252:   pPntColl.AddPoint p2ndPnt
1253:   pPntColl.AddPoint p3rdPnt
1254:   pTrianglePoly.Close
  
1256:   Set pGeometry = pTrianglePoly
1257:   pElementPly.Geometry = pGeometry
  
  ''''''''''''''''''''''''''''''''''''''
  'triangle shadow colors and dimensions
  
1262:   Set pShadowPolyElement = New PolygonElement
1263:   Set pShadowFillShapeElement = pShadowPolyElement
1264:   Set pFillShapeElement = pShadowPolyElement
1265:   pFillShapeElement.Symbol = pShadowFillSymbol
1266:   Set pElementShadowPly = pShadowPolyElement
  
1268:   Set pShadowTrianglePoly = New esriGeometry.Polygon
1269:   Set pPntColl = pShadowTrianglePoly
                                                  'offset shadow triangle by 3 pixels
1271:   p1stShadowPnt.x = p1stPnt.x - (ConvertPixelsToRW(2, pApp) * dRatioInchToMapUnitX)
1272:   p1stShadowPnt.y = p1stPnt.y - (ConvertPixelsToRW(2, pApp) * dRatioInchToMapUnitY)
1273:   p2ndShadowPnt.x = p2ndPnt.x - (ConvertPixelsToRW(2, pApp) * dRatioInchToMapUnitX)
1274:   p2ndShadowPnt.y = p2ndPnt.y - (ConvertPixelsToRW(2, pApp) * dRatioInchToMapUnitY)
1275:   p3rdShadowPnt.x = p3rdPnt.x - (ConvertPixelsToRW(2, pApp) * dRatioInchToMapUnitX)
1276:   p3rdShadowPnt.y = p3rdPnt.y - (ConvertPixelsToRW(2, pApp) * dRatioInchToMapUnitY)
1277:   pPntColl.AddPoint p1stShadowPnt
1278:   pPntColl.AddPoint p2ndShadowPnt
1279:   pPntColl.AddPoint p3rdShadowPnt
1280:   pShadowTrianglePoly.Close
1281:   Set pGeometry = pShadowTrianglePoly
1282:   pElementShadowPly.Geometry = pGeometry
  
  'tag the graphic elements for later tracking
  Dim pElementProps As IElementProperties
1286:   Set pElementProps = pElementShadowPly
1287:   pElementProps.CustomProperty = "BubbleID:" & lBubbleId
1288:   Set pElementProps = pElement
1289:   pElementProps.CustomProperty = "BubbleID:" & lBubbleId
1290:   Set pElementProps = pElementPly
1291:   pElementProps.CustomProperty = "BubbleID:" & lBubbleId


  
  ''''''''''''''''''''''''''''''
  'Add triangle shadow to layout
1297:   pGraphicsContainer.AddElement pElementShadowPly, 0
    
  '''''''''''''''''''''''''''
  'Add mapframe to the layout
1301:   pGraphicsContainer.AddElement pElement, 0
  
  '''''''''''''''''''''''
  'Add triangle to layout
1305:   pGraphicsContainer.AddElement pElementPly, 0
  
  
1308:   Set pActiveView = pNewMap
  
  '''''''''''''''''''''''''''''''''''
  'Set the detail inset's zoom extent
1312:   Set pNewEnv = New envelope
  
1314:   If bIsCircular Then
1315:     pNewEnv.XMin = dXOrigin - dRadius
1316:     pNewEnv.XMax = dXOrigin + dRadius
1317:   Else
1318:     pNewEnv.XMin = dXOrigin - (dWidthOrigin / 2)
1319:     pNewEnv.XMax = dXOrigin + (dWidthOrigin / 2)
1320:   End If
1321:   pNewEnv.YMin = dYOrigin - dRadius
1322:   pNewEnv.YMax = dYOrigin + dRadius
    
1324:   pActiveView.Extent = pNewEnv
1325:   pActiveView.Refresh
  
  'Set pGraphicsLayer = pMxDoc.FocusMap.BasicGraphicsLayer
  'Set pCompositeGraphicsLayer = pGraphicsLayer
  'Set pGraphicsLayer.AssociatedLayer = pFeatureLayer
  
  'Set pGraphicsLayer = pCompositeGraphicsLayer.FindLayer("NW Detail Inset Arrows")
  'If pGraphicsLayer Is Nothing Then
  '  Set pGraphicsLayer = pCompositeGraphicsLayer.AddLayer("NW Detail Inset Arrows", pFeatureLayer)
  'End If

1336:   Set pActiveView = pMxDoc.FocusMap
1337:   pActiveView.PartialRefresh esriViewGraphics, Nothing, Nothing
1338:   pMxDoc.CurrentContentsView.Refresh Nothing

1340:   pNWSeriesOptions.BubbleGraphicAdd pElement, pElementPly, pElementShadowPly, pNewMap.Name

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
  
1366:   Set pMxDoc = pApp.Document
1367:   Set pActiveView = pMxDoc.FocusMap
1368:   Set pDT = pActiveView.ScreenDisplay.DisplayTransformation
1369:   deviceRECT = pDT.DeviceFrame
1370:   pixelExtent = deviceRECT.Right - deviceRECT.Left
1371:   Set pEnv = pDT.VisibleBounds
1372:   realWorldDisplayExtent = pEnv.Width
1373:   sizeOfOnePixel = realWorldDisplayExtent / pixelExtent
1374:   ConvertPixelsToRW = pixelUnits * sizeOfOnePixel
  
  Exit Function
  
ErrorHandler:
1379:    MsgBox "Error " & Err.Number & ": " & Err.Description & vbNewLine _
       & "In " & Err.Source & " at DrawSelectedArrows.ConvertPixelsToRW", vbCritical
End Function


