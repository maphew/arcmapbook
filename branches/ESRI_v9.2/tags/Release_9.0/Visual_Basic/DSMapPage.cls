VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DSMapPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Copyright 1995-2004 ESRI

' All rights reserved under the copyright laws of the United States.

' You may freely redistribute and use this sample code, with or without modification.

' Disclaimer: THE SAMPLE CODE IS PROVIDED "AS IS" AND ANY EXPRESS OR IMPLIED
' WARRANTIES, INCLUDING THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS
' FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL ESRI OR
' CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY,
' OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
' SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
' INTERRUPTION) SUSTAINED BY YOU OR A THIRD PARTY, HOWEVER CAUSED AND ON ANY
' THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT ARISING IN ANY
' WAY OUT OF THE USE OF THIS SAMPLE CODE, EVEN IF ADVISED OF THE POSSIBILITY OF
' SUCH DAMAGE.

' For additional information contact: Environmental Systems Research Institute, Inc.

' Attn: Contracts Dept.

' 380 New York Street

' Redlands, California, U.S.A. 92373

' Email: contracts@esri.com

Option Explicit

Implements IDSMapPage
Implements IPersistVariant

Dim m_PageItemColl As Collection
Dim m_sPageName As String
Dim m_bPrintPage As Boolean
Dim m_dPageRotation As Double
Dim m_dPageScale As Double
Dim m_dLastOutputted As Date
Dim m_pPageShape As IPolygon
Dim m_lPageNumber As Long

Private Sub Class_Initialize()
44:   Set m_PageItemColl = New Collection
End Sub

Private Sub Class_Terminate()
48:   Set m_PageItemColl = Nothing
End Sub

Private Sub IDSMapPage_AddPageItem(ByVal PageItem As IElement)
52:   m_PageItemColl.Add PageItem
End Sub

Private Sub IDSMapPage_DrawPage(pDoc As IMxDocument, pDSMapSeries As IDSMapSeries, bRefreshFlag As Boolean)
On Error GoTo ErrHand:
  Dim pMap As IMap, lLoop As Long, pEnv As IEnvelope, lIndex As Long
  Dim pFeatLayer As IFeatureLayer, pActive As IActiveView, pTempLayer As ILayer
  Dim pGraphicsCont As IGraphicsContainer, pTempColl As Collection, pElemProps As IElementProperties
  Dim pTextSym As ISimpleTextSymbol, pClone As IClone, pSeriesOpts2 As IDSMapSeriesOptions2
  Dim pSeriesProps As IDSMapSeriesProps, pSeriesOpts As IDSMapSeriesOptions
  Dim pSeriesOpts3 As IDSMapSeriesOptions3
63:   Set pSeriesProps = pDSMapSeries
  
  'Find the data frame
66:   Set pMap = FindDataFrame(pDoc, pSeriesProps.DataFrameName)
67:   If pMap Is Nothing Then
68:     MsgBox "Could not find map in DrawPage routine!!!"
    Exit Sub
70:   End If
  
  'Find the Index layer
73:   For lLoop = 0 To pMap.LayerCount - 1
74:     If TypeOf pMap.Layer(lLoop) Is ICompositeLayer Then
75:       Set pFeatLayer = CompositeLayer1(pMap.Layer(lLoop), pSeriesProps.IndexLayerName)
76:       If Not pFeatLayer Is Nothing Then Exit For
77:     Else
78:       If pMap.Layer(lLoop).Name = pSeriesProps.IndexLayerName Then
79:         Set pFeatLayer = pMap.Layer(lLoop)
80:         Exit For
81:       End If
82:     End If
83:   Next lLoop
84:   If pFeatLayer Is Nothing Then
85:     MsgBox "Could not find index layer in DrawPage routine!!!"
    Exit Sub
87:   End If
  
  'Switch to the Layout view if we are not already there
90:   If Not TypeOf pDoc.ActiveView Is IPageLayout Then
91:     Set pDoc.ActiveView = pDoc.PageLayout
92:   End If
  
  'Remove any previous neighbor labels.
95:   Set pGraphicsCont = pDoc.ActiveView
96:   pGraphicsCont.Reset
97:   Set pTempColl = New Collection
98:   Set pElemProps = pGraphicsCont.Next
99:   Do While Not pElemProps Is Nothing
100:     If pElemProps.Name = "DSMAPBOOK TEXT" Then
101:       pTempColl.Add pElemProps
102:     End If
103:     Set pElemProps = pGraphicsCont.Next
104:   Loop
105:   For lLoop = 1 To pTempColl.Count
106:     pGraphicsCont.DeleteElement pTempColl.Item(lLoop)
107:   Next lLoop
108:   Set pTempColl = Nothing
  
  'Rotate the frame if necessary
111:   Set pActive = pMap
112:   Set pSeriesOpts = pSeriesProps
113:   Set pSeriesOpts2 = pSeriesOpts
114:   If pSeriesOpts.RotateFrame Then
'    If m_dPageRotation > 0 Then
116:       pActive.ScreenDisplay.DisplayTransformation.Rotation = m_dPageRotation
'    End If
118:   End If
  
  'Set the extent and possibly the scale for the map
121:   SetMapExtent pSeriesOpts, pActive
    
  'Set the clip property
  'Updated 6/18/03 to support cross hatching of area outside the clip
  Select Case pSeriesOpts2.ClipData
  Case 0   'No clipping
'    pMap.ClipGeometry = Nothing
  Case 1   'Clipping only
129:     pMap.ClipGeometry = m_pPageShape
  Case 2   'clipping with cross hatching of area outside the clip
131:     pMap.ClipGeometry = Nothing
132:     CreateClipElement pDoc, pActive, pFeatLayer
133:   End Select
  
  'Check for indicator maps and update those also
136:   RefreshIndicators pDoc, pSeriesProps, bRefreshFlag
  
  'Check for Date and Title elements
139:   UpdateTaggedElements pDoc, m_sPageName, bRefreshFlag, pDSMapSeries
  
  'Label neighboring tiles if necessary
142:   If pSeriesOpts.LabelNeighbors Then
143:     Set pClone = pSeriesOpts.LabelSymbol
144:     Set pTextSym = pClone.Clone
145:     lIndex = pFeatLayer.FeatureClass.FindField(pSeriesProps.IndexFieldName)
146:     If lIndex >= 0 Then
147:       LabelNeighbors pDoc, pFeatLayer, pTextSym, lIndex, pSeriesProps.DataFrameName
148:     End If
149:   End If
  
   '--------------------------------
  ' DetailExtension:
  '
  ' Update detail maps by executing "Create maps" button.
  ' Added 6/17/2004 to support inset creation from a separate tool.
  '
  Dim pDocument As IDocument
  Dim pUID As IUID
  Dim pCommandItem As ICommandItem
  Dim pCommand As ICommand, bFlag As Boolean
  
162:   Set pUID = New UID
163:   bFlag = False
On Error GoTo NoDetails:
165:   pUID.Value = "DetailAreaExt.CreateDetailsCmd"
On Error GoTo ErrHand:
   
168:    If Not bFlag Then
169:     Set pDocument = pDoc
170:     Set pCommandItem = pDocument.CommandBars.Find(pUID)
171:     If Not pCommandItem Is Nothing Then
172:       If TypeOf pCommandItem Is ICommand Then
173:         Set pCommand = pCommandItem
174:         If pCommand.Enabled Then
175:           pCommandItem.Execute
176:         End If
177:       End If
178:     End If
179:   End If
  
  '  End of DetailExtension additions
  '-------------------------------
  
  'Select the tile if this option is selected - Added 11/23/2004 by LY
  Dim pFeatSel As IFeatureSelection, pQuery As IQueryFilter
186:   Set pSeriesOpts3 = pDSMapSeries
187:   If pSeriesOpts3.SelectTile Then
188:     Set pQuery = New QueryFilter
189:     pQuery.WhereClause = pSeriesProps.IndexFieldName & " = '" & m_sPageName & "'"
190:     Set pFeatSel = pFeatLayer
191:     pFeatSel.SelectFeatures pQuery, esriSelectionResultNew, True
192:   End If
  '-------------------------------------------------------------------------
  
195:   If bRefreshFlag Then
196:     pDoc.ActiveView.Refresh
197:   End If

  Exit Sub
  
NoDetails:
202:   bFlag = True
203:   Resume Next
  
ErrHand:
206:   MsgBox "IDSMapPage_DrawPage - " & Erl & " - " & Err.Description
End Sub

Private Function IDSMapPage_IndexPage(pIndexLayer As IFeatureLayer, sFieldName As String) As Collection
On Error GoTo ErrHand:
  Dim pFilter As ISpatialFilter, pIndex As Collection, lFieldIndex As Long
  Dim pCursor As IFeatureCursor, pFeat As IFeature, sValue As String, lLoop As Long
  
  'Check for a valid index layer
215:   Set IDSMapPage_IndexPage = Nothing
216:   If pIndexLayer Is Nothing Then
217:     MsgBox "You did not send a valid index layer to the IndexPage function!!"
    Exit Function
219:   End If
  
  'Check for a valid field name
222:   If sFieldName = "" Or sFieldName = " " Then
223:     MsgBox "You did not send a valid field name to the IndexPage function!!"
    Exit Function
225:   End If
  
  'Check for field name in the layer
228:   lFieldIndex = pIndexLayer.FeatureClass.FindField(sFieldName)
229:   If lFieldIndex < 0 Then
230:     MsgBox "Could not find field name in the index layer of the IndexPage function!!"
    Exit Function
232:   End If
  
  'Perform the query of the index layer using the page shape
235:   Set pFilter = New SpatialFilter
236:   pFilter.AddField sFieldName
237:   pFilter.WhereClause = sFieldName & " <> '' and " & sFieldName & " <> ' '"
238:   Set pFilter.Geometry = m_pPageShape
239:   pFilter.GeometryField = pIndexLayer.FeatureClass.ShapeFieldName
240:   pFilter.SpatialRel = esriSpatialRelIntersects
241:   Set pCursor = pIndexLayer.Search(pFilter, True)
  
  'Populate the collection with the results of the query
244:   Set pIndex = New Collection
245:   Set pFeat = pCursor.NextFeature
246:   Do While Not pFeat Is Nothing
247:     sValue = pFeat.Value(lFieldIndex)
248:     If pIndex.Count > 0 Then
249:       For lLoop = 1 To pIndex.Count
250:         If sValue < pIndex.Item(lLoop) Then
251:           pIndex.Add sValue, sValue, lLoop
252:           Exit For
253:         ElseIf sValue = pIndex.Item(lLoop) Then
254:           Exit For
255:         End If
256:         If lLoop = pIndex.Count Then
257:           pIndex.Add sValue, sValue
258:         End If
259:       Next lLoop
260:     Else
261:       pIndex.Add sValue, sValue
262:     End If
    
264:     Set pFeat = pCursor.NextFeature
265:   Loop
  
267:   Set IDSMapPage_IndexPage = pIndex

  Exit Function
  
ErrHand:
272:   MsgBox "IDSMapPage_IndexPage - " & Err.Description
End Function

Private Property Let IDSMapPage_LastOutputted(RHS As Date)
276:  m_dLastOutputted = RHS
End Property

Private Property Get IDSMapPage_LastOutputted() As Date
280:   IDSMapPage_LastOutputted = m_dLastOutputted
End Property

Private Property Get IDSMapPage_PageItem(Index As Long) As IElement
284:   If Index > -1 And Index < m_PageItemColl.Count Then
285:     Set IDSMapPage_PageItem = m_PageItemColl.Item(Index + 1)
286:   Else
287:     Set IDSMapPage_PageItem = Nothing
288:   End If
End Property

Private Property Get IDSMapPage_PageItemCount() As Long
292:   IDSMapPage_PageItemCount = m_PageItemColl.Count
End Property

Private Property Let IDSMapPage_PageName(RHS As String)
296:   m_sPageName = RHS
End Property

Private Property Get IDSMapPage_PageName() As String
300:   IDSMapPage_PageName = m_sPageName
End Property

Private Property Let IDSMapPage_EnablePage(RHS As Boolean)
304:   m_bPrintPage = RHS
End Property

Private Property Get IDSMapPage_EnablePage() As Boolean
308:   IDSMapPage_EnablePage = m_bPrintPage
End Property

Private Property Let IDSMapPage_PageNumber(RHS As Long)
312:   m_lPageNumber = RHS
End Property

Private Property Get IDSMapPage_PageNumber() As Long
316:   IDSMapPage_PageNumber = m_lPageNumber
End Property

Private Property Let IDSMapPage_PageRotation(RHS As Double)
320:   m_dPageRotation = RHS
End Property

Private Property Get IDSMapPage_PageRotation() As Double
324:   IDSMapPage_PageRotation = m_dPageRotation
End Property

Private Property Let IDSMapPage_PageScale(RHS As Double)
328:   m_dPageScale = RHS
End Property

Private Property Get IDSMapPage_PageScale() As Double
332:   IDSMapPage_PageScale = m_dPageScale
End Property

Private Property Set IDSMapPage_PageShape(RHS As IPolygon)
336:   Set m_pPageShape = RHS
End Property

Private Property Get IDSMapPage_PageShape() As IPolygon
340:   Set IDSMapPage_PageShape = m_pPageShape
End Property

Private Sub IDSMapPage_RemovePageItem(Index As Long)
344:   If Index > -1 And Index < m_PageItemColl.Count Then
345:     m_PageItemColl.Remove Index + 1
346:   End If
End Sub

Private Property Get IPersistVariant_ID() As esriSystem.IUID
  Dim id As New UID
351:   id = "DSMapBookPrj.DSMapPage"
352:   Set IPersistVariant_ID = id
End Property

Private Sub IPersistVariant_Load(ByVal Stream As esriSystem.IVariantStream)
'Load the persisted parameters of the renderer
On Error GoTo ErrHand:
  Dim lLoop As Long, lCount As Long, pElem As IElement, sFirstItem As String
  Dim lPropCount As Long
  
  'Added 2/18/04 to make the list of persisted properties more dynamic
362:   sFirstItem = Stream.Read
363:   If UCase(Left(sFirstItem, 18)) = "PAGEPROPERTYCOUNT-" Then
364:     lPropCount = Mid(sFirstItem, 19) - 1
365:     m_sPageName = Stream.Read
366:   Else
367:     lPropCount = 5
368:     m_sPageName = sFirstItem
369:   End If
    
  'Original page properties
372:   m_bPrintPage = Stream.Read
373:   m_dPageRotation = Stream.Read
374:   m_dPageScale = Stream.Read
375:   m_dLastOutputted = Stream.Read
376:   Set m_pPageShape = Stream.Read
  
  'Additional properties added after 2/18/04
379:   If lPropCount > 5 Then    'Checking for page number
380:     m_lPageNumber = Stream.Read
381:   Else
382:     m_lPageNumber = -1
383:   End If
  
  'More original properties.  Writen out below the new properties because they are of variable length
386:   lCount = Stream.Read
387:   If lCount > 0 Then
388:     For lLoop = 1 To lCount
389:       Set pElem = Stream.Read
390:       m_PageItemColl.Add pElem
391:     Next lLoop
392:   End If
    
  Exit Sub
ErrHand:
396:   MsgBox "MapPage - IPersistVariant_Load - " & Erl & " - " & Err.Description
End Sub

Private Sub IPersistVariant_Save(ByVal Stream As esriSystem.IVariantStream)
'Write it all out
On Error GoTo ErrHand:
  Dim lLoop As Long
  
  'Added 2/18/04 to make the list of persisted properties more dynamic
405:   Stream.Write "PAGEPROPERTYCOUNT-7"
    
407:   Stream.Write m_sPageName
408:   Stream.Write m_bPrintPage
409:   Stream.Write m_dPageRotation
410:   Stream.Write m_dPageScale
411:   Stream.Write m_dLastOutputted
412:   Stream.Write m_pPageShape
413:   Stream.Write m_lPageNumber   'Added 2/18/04
  
415:   Stream.Write m_PageItemColl.Count
416:   If m_PageItemColl.Count > 0 Then
417:     For lLoop = 1 To m_PageItemColl.Count
418:       Stream.Write m_PageItemColl.Item(lLoop)
419:     Next lLoop
420:   End If
    
  Exit Sub
ErrHand:
424:   MsgBox "MapPage - IPersistVariant_Save - " & Err.Description
End Sub

Private Sub LabelNeighbors(pDoc As IMxDocument, pFLayer As IFeatureLayer, pTextSym As ISimpleTextSymbol, _
 lIndex As Long, sFrameName As String)
'Routine for loop through the tiles that are touching are selected tile
On Error GoTo ErrHand:
  Dim pElem As IElement, pTextElem As ITextElement, pMap As IMap
  Dim pGraphSel As IGraphicsContainerSelect
  Dim pSpatial As ISpatialFilter, pFeatCursor As IFeatureCursor
  Dim pFeats As IFeature, pActive As IActiveView, sText As String
435:   Set pMap = pDoc.FocusMap
436:   Set pActive = pDoc.ActiveView
  
438:   Set pSpatial = New SpatialFilter
439:   Set pSpatial.Geometry = m_pPageShape
440:   pSpatial.GeometryField = pFLayer.FeatureClass.ShapeFieldName
441:   pSpatial.SpatialRel = esriSpatialRelTouches
442:   Set pFeatCursor = pFLayer.Search(pSpatial, False)
443:   Set pFeats = pFeatCursor.NextFeature
444:   Do While Not pFeats Is Nothing
    'If there is a value for the selected adjacent tile, then get it and label the feature
446:     If Not IsNull(pFeats.Value(lIndex)) Then
447:       sText = pFeats.Value(lIndex)
448:       LabelAdjacent pFeats, pMap, pActive, pTextSym, sText, sFrameName
449:     End If
450:     Set pFeats = pFeatCursor.NextFeature
451:   Loop
452:   Set pGraphSel = pActive
453:   pGraphSel.UnselectAllElements

  Exit Sub
ErrHand:
457:   MsgBox "LabelNeighbors - " & Err.Description
End Sub

Sub LabelAdjacent(pLabelFeat As IFeature, pMap As IMap, pActive As IActiveView, pTextSym As ITextSymbol, _
 sText As String, sFrameName As String)
'Routine for labeling the outside edge of our data frame based on the relative
'position of the tile feature being sent in (pLabelFeat) to the selected tile (pIndexFeat)
'The relative position is determined by finding the midpoint of the common line between
'the selected tile and the tile to be labeled.  The midpoint is then combined with the
'center of the extent of the map to create a line that is extended to the edge of the
'map extent.  The location along the map extent is then extrapolated out to a point on
'the boundary of our map frame.  This location is then labeled with the name of the tile.
On Error GoTo ErrHand:
  Dim pCommonGeom As IGeometry, pTopoOp As ITopologicalOperator, pPolyline2 As IPolyline
  Dim pMidPt As IPoint, pPolyLine As IPolyline, pEnv As IEnvelope, pTmpFrame As IMapFrame
  Dim pCenterPt As IPoint, pMapView As IActiveView, pMulti As IPointCollection
  Dim pGraph As IGraphicsContainer, lLoop As Long, pElem As IElement
  Dim pElemProps As IElementProperties, pTrans2D As ITransform2D
  Dim pPt As IConstructPoint, pNewPt As IPoint, pTextElem As ITextElement
  Dim pMapFrame As IMapFrame, pMapEnv As IEnvelope, pFramePoly As IPointCollection
  Dim pLine As ILine, dDist As Double, iSeg As Integer, pEndPt As IPoint
  Dim pProx As IProximityOperator, dTmpDist As Double, pCurve2 As ICurve
  Dim pOutPt As IPoint, dAlong As Double, dFrom As Double, bSide As Boolean
  Dim pPoints As IPointCollection, pPoints2 As IPointCollection, dAngle As Double
  Dim pIntPoints As IPointCollection, pIntTopo As ITopologicalOperator

483:   Set pTopoOp = m_pPageShape
484:   Set pCommonGeom = pTopoOp.Intersect(pLabelFeat.Shape, esriGeometry1Dimension)
485:   If pCommonGeom.IsEmpty Then
486:     Set pCommonGeom = pTopoOp.Intersect(pLabelFeat.Shape, esriGeometry0Dimension)
487:     Set pMulti = pCommonGeom
488:     Set pMidPt = pMulti.Point(0)
489:   Else
490:     Set pPolyLine = pCommonGeom
491:     Set pMidPt = New esriGeometry.Point
492:     pPolyLine.QueryPoint esriNoExtension, 0.5, True, pMidPt
493:   End If
    
  'Find center point of map frame
496:   Set pCenterPt = New esriGeometry.Point
497:   Set pMapView = pMap
498:   Set pEnv = pMapView.Extent
499:   pCenterPt.X = pEnv.XMin + ((pEnv.XMax - pEnv.XMin) / 2)
500:   pCenterPt.Y = pEnv.YMin + ((pEnv.YMax - pEnv.YMin) / 2)

  'Get the geometry of the map frame
503:   Set pGraph = pActive
504:   pGraph.Reset
505:   Set pElem = pGraph.Next
506:   Do While Not pElem Is Nothing
507:     If TypeOf pElem Is IMapFrame Then
508:       Set pTmpFrame = pElem
509:       If pTmpFrame.Map.Name = sFrameName Then
510:         Set pMapFrame = pElem
511:         Exit Do
512:       End If
513:     End If
514:     Set pElem = pGraph.Next
515:   Loop
  If pMapFrame Is Nothing Then Exit Sub
  
518:   Set pMapEnv = pMapFrame.MapBounds
519:   Set pFramePoly = pElem.Geometry
  
  'Create curves and intersect them
522:   Set pPoints = New Polyline
523:   pPoints.AddPoint pMapEnv.LowerLeft
524:   pPoints.AddPoint pMapEnv.LowerRight
525:   pPoints.AddPoint pMapEnv.UpperRight
526:   pPoints.AddPoint pMapEnv.UpperLeft
527:   pPoints.AddPoint pMapEnv.LowerLeft
  
529:   Set pPoints2 = RotatedAndExtendedLine(pCenterPt, pMidPt)
  'If for some reason nothing is returned, go with a rotation of 0
  If pPoints2 Is Nothing Then Exit Sub
532:   Set pPolyline2 = pPoints2
  If pPolyline2.IsEmpty Then Exit Sub
    
  'Find the intersection point of the line we created and the map extent boundary
536:   Set pIntTopo = pPoints2
537:   Set pIntPoints = pIntTopo.Intersect(pPoints, esriGeometry0Dimension)
  If pIntPoints.PointCount = 0 Then Exit Sub
  
540:   Set pEndPt = pIntPoints.Point(0)
  
  'Extrapolate the point on the extent to a point on the outside of the map frame
  'Figure out which segment we are closest to
544:   Set pProx = pEndPt
545:   dDist = 999999
546:   iSeg = -1
547:   For lLoop = 0 To 3
548:     Set pLine = New esriGeometry.Line
    Select Case lLoop
    Case 0
551:       pLine.PutCoords pMapEnv.LowerLeft, pMapEnv.UpperLeft
    Case 1
553:       pLine.PutCoords pMapEnv.UpperLeft, pMapEnv.UpperRight
    Case 2
555:       pLine.PutCoords pMapEnv.UpperRight, pMapEnv.LowerRight
    Case Else
557:       pLine.PutCoords pMapEnv.LowerRight, pMapEnv.LowerLeft
558:     End Select
    
560:     dTmpDist = pProx.ReturnDistance(pLine)
561:     If dTmpDist < dDist Then
562:       dDist = dTmpDist
563:       iSeg = lLoop
564:       Set pCurve2 = pLine
565:     End If
566:   Next lLoop
567:   Set pOutPt = New esriGeometry.Point
568:   pCurve2.QueryPointAndDistance esriNoExtension, pEndPt, True, pOutPt, dAlong, dFrom, bSide
  
  'We know have the segment and ratio length on that segment, so we can transfer that
  'information to the frame geometry and find the corresponding point there
572:   Set pPt = New esriGeometry.Point
573:   Set pLine = New esriGeometry.Line
  Select Case iSeg
  Case 0
576:     pLine.PutCoords pFramePoly.Point(0), pFramePoly.Point(1)
577:     pTextSym.HorizontalAlignment = esriTHACenter
578:     pTextSym.VerticalAlignment = esriTVABottom
  Case 1
580:     pLine.PutCoords pFramePoly.Point(1), pFramePoly.Point(2)
581:     pTextSym.HorizontalAlignment = esriTHACenter
582:     pTextSym.VerticalAlignment = esriTVABottom
  Case 2
584:     pLine.PutCoords pFramePoly.Point(2), pFramePoly.Point(3)
585:     pTextSym.HorizontalAlignment = esriTHACenter
586:     pTextSym.VerticalAlignment = esriTVABottom
  Case 3
588:     pLine.PutCoords pFramePoly.Point(3), pFramePoly.Point(0)
589:     pTextSym.HorizontalAlignment = esriTHACenter
590:     pTextSym.VerticalAlignment = esriTVATop
591:   End Select
592:   pPt.ConstructAlong pLine, esriNoExtension, dAlong, True
593:   Set pNewPt = pPt
  
  'Now that we have a point along the data frame, we can place the label based on
  'that point and which side of the frame it is on
597:   Set pTextElem = New TextElement
598:   Set pElem = pTextElem
599:   pTextElem.Symbol = pTextSym
600:   pElem.Geometry = pNewPt
601:   Set pElemProps = pElem
602:   pElemProps.Name = "DSMAPBOOK TEXT"
603:   pTextElem.Text = sText
604:   Set pTrans2D = pTextElem
  Select Case iSeg
  Case 0
607:     dAngle = 90 * (3.14159265358979 / 180)
  Case 1
609:     dAngle = 0
  Case 2
611:     dAngle = 270 * (3.14159265358979 / 180)
  Case 3
613:     dAngle = 0
614:   End Select
615:   pTrans2D.Rotate pNewPt, dAngle
616:   pGraph.AddElement pElem, 0
  
  Exit Sub
ErrHand:
620:   MsgBox "LabelAdjacent - " & Err.Description
End Sub

Private Function RotatedAndExtendedLine(pCenterPt As IPoint, pMidPt As IPoint) As IPolyline
On Error GoTo ErrHand:
  Dim pPoints As IPointCollection, pPolyLine As IPolyline, pLine As ILine, pNewPt As IConstructPoint
  Dim dOrigAngle As Double, dNewAngle As Double, dLength As Double, dRadAngle As Double
  Dim pNewPoints As IPointCollection, pNewPoint As IConstructPoint, dA As Double
  
  'Create a line so we can get the current angle and distance
630:   Set pLine = New esriGeometry.Line
631:   pLine.PutCoords pCenterPt, pMidPt
632:   dLength = pLine.Length
  
634:   If m_dPageRotation = 0 Then
    'Create another point at the same angle to make sure our line crosses the extent boundary
636:     Set pNewPt = New esriGeometry.Point
637:     pNewPt.ConstructAngleDistance pMidPt, pLine.Angle, dLength * 100
638:     Set pPoints = New Polyline
639:     pPoints.AddPoint pCenterPt
640:     pPoints.AddPoint pMidPt
641:     pPoints.AddPoint pNewPt
642:     Set RotatedAndExtendedLine = pPoints
    Exit Function
644:   End If
  
  'If the page is rotated, then we have to rotate the labeling of adjacent tiles also
647:   dOrigAngle = pLine.Angle * (180 / 3.14159265358979)
648:   dA = dOrigAngle
649:   If dOrigAngle < 0 Then
650:     dOrigAngle = 360 - Abs(dOrigAngle)
651:   End If
652:   dNewAngle = dOrigAngle + m_dPageRotation
653:   If dNewAngle >= 360 Then
654:     dNewAngle = dNewAngle - 360
655:   End If
656:   dRadAngle = dNewAngle * (3.14159265358979 / 180)
  
  'Make a new esrigeometry.line at the rotated angle we just calculated.  The new esrigeometry.line is made shorter than the original
  'to ensure the line does not extend past the map bounds we need to intersect it with in the next stage
660:   Set pNewPoint = New esriGeometry.Point
661:   Set pNewPoints = New Polyline
662:   pNewPoint.ConstructAngleDistance pCenterPt, dRadAngle, dLength * 100
663:   pNewPoints.AddPoint pCenterPt
664:   pNewPoints.AddPoint pNewPoint
665:   Set RotatedAndExtendedLine = pNewPoints
  
  Exit Function
ErrHand:
669:   MsgBox "RotatedLine - " & Err.Description
End Function

Private Sub RefreshIndicators(pDoc As IMxDocument, pSeriesProps As IDSMapSeriesProps, bRefreshFlag As Boolean)
'Routine for updating any identicator maps there might be
On Error GoTo ErrHand:
  Dim pGridLayer As IFeatureLayer, pGridSel As IFeatureSelection
  Dim lLoop As Long, pActive As IActiveView, pSpatial As ISpatialFilter
  Dim pFeature As IFeature, pCursor As IFeatureCursor, pEnv As IEnvelope
  Dim pQuery As IQueryFilter, lLoop2 As Long, pMap As IMap, pSelEvents As ISelectionEvents

  'Check for indicator maps and update those also
681:   For lLoop = 0 To pDoc.Maps.Count - 1
682:     If pDoc.Maps.Item(lLoop).Name = "Global Indicator" Or pDoc.Maps.Item(lLoop).Name = "Local Indicator" Then
683:       Set pMap = pDoc.Maps.Item(lLoop)
      'Find the Index layer
685:       For lLoop2 = 0 To pMap.LayerCount - 1
686:         If pMap.Layer(lLoop2).Name = "Identifier Layer" Then
687:           Set pGridLayer = pMap.Layer(lLoop2)
688:           Exit For
689:         End If
690:       Next lLoop2
691:       If pGridLayer Is Nothing Then
692:         MsgBox "Could not find layer called Identifier Layer, can not redraw " & pMap.Name & " frame!!!"
        Exit Sub
694:       End If
      
'      Set pGridLayer = pDoc.Maps.Item(lLoop).Layer(0)
697:       Set pGridSel = pGridLayer
698:       Set pQuery = New QueryFilter
699:       pQuery.WhereClause = pSeriesProps.IndexFieldName & " = '" & m_sPageName & "'"
700:       pGridSel.Clear
701:       pGridSel.SelectFeatures pQuery, esriSelectionResultNew, True
        
703:       If pMap.Name = "Global Indicator" Then
704:         Set pActive = pDoc.Maps.Item(lLoop)
705:         If bRefreshFlag Then pActive.Refresh
706:       ElseIf pMap.Name = "Local Indicator" Then
707:         Set pSpatial = New SpatialFilter
708:         Set pSpatial.Geometry = m_pPageShape
709:         pSpatial.GeometryField = pGridLayer.FeatureClass.ShapeFieldName
710:         pSpatial.SpatialRel = esriSpatialRelIntersects
711:         Set pCursor = pGridLayer.Search(pSpatial, False)
712:         Set pFeature = pCursor.NextFeature
713:         Do While Not pFeature Is Nothing
714:           If pEnv Is Nothing Then
715:             Set pEnv = pFeature.Shape.Envelope
716:           Else
717:             pEnv.Union pFeature.Shape.Envelope
718:           End If
719:           Set pFeature = pCursor.NextFeature
720:         Loop
721:         Set pActive = pMap
722:         pActive.Extent = pEnv
723:         If bRefreshFlag Then pActive.Refresh
724:       End If
725:       Set pSelEvents = pMap
726:       pSelEvents.SelectionChanged
      
728:       Set pGridLayer = Nothing
729:     End If
730:   Next lLoop

  Exit Sub
ErrHand:
734:   MsgBox "RefreshIndicators - " & Err.Description
End Sub

Private Sub UpdateTaggedElements(pDoc As IMxDocument, sTileName As String, bRefreshFlag As Boolean, _
 pDSMapSeries As IDSMapSeries)
'Routine for updating text elements tagged as Date or Title elements
On Error GoTo ErrHand:
  Dim pGraphicsCont As IGraphicsContainer, pElemProps As IElementProperties
  Dim pTextElement As ITextElement, pActive As IActiveView, pElem As IElement
  Dim pEnv As IEnvelope, pEnv2 As IEnvelope, sText As String, bUpdate As Boolean
744:   Set pGraphicsCont = pDoc.PageLayout
745:   Set pActive = pGraphicsCont
746:   pGraphicsCont.Reset
747:   Set pElemProps = pGraphicsCont.Next
748:   Do While Not pElemProps Is Nothing
749:     If TypeOf pElemProps Is ITextElement Then
750:       bUpdate = True
      Select Case pElemProps.Name
      Case "DSMAPBOOK - DATE"
753:         sText = Format(Date, "mmm dd, yyyy")
      Case "DSMAPBOOK - TITLE"
755:         sText = sTileName
      Case "DSMAPBOOK - PAGENUMBER"
757:         sText = CStr(m_lPageNumber)
      Case "DSMAPBOOK - EXTRAITEM"
759:         sText = GetExtraItemValue(pDoc, sTileName, pElemProps.Type, pDSMapSeries)
      Case Else
761:         bUpdate = False
762:       End Select
      
764:       If bUpdate Then
765:         Set pElem = pElemProps
766:         Set pEnv = New Envelope
767:         pElem.QueryBounds pActive.ScreenDisplay, pEnv
768:         Set pTextElement = pElemProps
769:         pTextElement.Text = sText
770:         pGraphicsCont.UpdateElement pTextElement
771:         Set pEnv2 = New Envelope
772:         pElem.QueryBounds pActive.ScreenDisplay, pEnv2
773:         pEnv.Union pEnv2
774:         If bRefreshFlag Then pActive.PartialRefresh esriViewGraphics, Nothing, pEnv
775:       End If
776:     End If
777:     Set pElemProps = pGraphicsCont.Next
778:   Loop

  Exit Sub
ErrHand:
782:   MsgBox "UpdateTaggedElements - " & Erl & " - " & Err.Description
End Sub

Private Function GetExtraItemValue(pDoc As IMxDocument, sTileName As String, sFieldName As String, _
 pSeriesProps As IDSMapSeriesProps) As String
On Error GoTo ErrHand:
  Dim pIndexLayer As IFeatureLayer, pQuery As IQueryFilter, pFCursor As IFeatureCursor
  Dim pFeat As IFeature, lIndex As Long, pMap As IMap, lIndex2 As Long
  
  'Find the data frame
792:   Set pMap = FindDataFrame(pDoc, pSeriesProps.DataFrameName)
793:   If pMap Is Nothing Then
794:     MsgBox "Could not find map in GetExtraItem routine!!!"
795:     GetExtraItemValue = "missing"
    Exit Function
797:   End If
  
  'Find the Index layer
800:   Set pIndexLayer = FindLayer(pSeriesProps.IndexLayerName, pMap)
801:   If pIndexLayer Is Nothing Then
802:     MsgBox "Could not find index layer (" & pSeriesProps.IndexLayerName & ") in GetExtraItemValue routine!!!"
803:     GetExtraItemValue = "missing"
    Exit Function
805:   End If
  
  'Find the field in the index layer
808:   lIndex = pIndexLayer.FeatureClass.FindField(sFieldName)
809:   If lIndex < 0 Then
810:     MsgBox "Could not find the field (" & sFieldName & ") you tagged the item with in the index layer!!!"
811:     GetExtraItemValue = "missing"
    Exit Function
813:   End If
  
  'Find the tile name field in the index layer
816:   lIndex2 = pIndexLayer.FeatureClass.FindField(pSeriesProps.IndexFieldName)
817:   If lIndex2 < 0 Then
818:     MsgBox "Could not find tile name field (" & pSeriesProps.IndexFieldName & ") in the index layer!!!"
819:     GetExtraItemValue = "missing"
    Exit Function
821:   End If
  
  'Create the query object then select the appropriate tile from the index layer
824:   Set pQuery = New QueryFilter
825:   pQuery.WhereClause = pSeriesProps.IndexFieldName & " = '" & sTileName & "'"
826:   Set pFCursor = pIndexLayer.Search(pQuery, False)
827:   Set pFeat = pFCursor.NextFeature
828:   If pFeat Is Nothing Then
829:     MsgBox "Could not select the tile from the index layer to tag with Extra Item!!!"
830:     GetExtraItemValue = "missing"
    Exit Function
832:   End If
  
  'Send back the value of the field
835:   If IsNull(pFeat.Value(lIndex)) Then
836:     GetExtraItemValue = " "
837:   Else
838:     GetExtraItemValue = pFeat.Value(lIndex)
839:   End If

  Exit Function
ErrHand:
843:   MsgBox "GetExtraItemValue - " & Erl & " - " & Err.Description
End Function

Private Sub CreateClipElement(pDoc As IMxDocument, pActive As IActiveView, _
 pFeatLayer As IFeatureLayer)
'Added 6/18/03 to support cross hatching of area outside the clip
On Error GoTo ErrHand:
  Dim pPoly As IPolygon, pTopoOp As ITopologicalOperator
  Dim pGraphs As IGraphicsContainer, pElem As IElement, pNewElem As IElement
  Dim pNewPoly As IPointCollection, pElemProps As IElementProperties
  Dim pFinalGeom As IPolygon, pPoly2 As IPolygon, lLoop As Long
  
  'Search for an existing clip element and delete it when found
'  Set pGraphs = pDoc.FocusMap
857:   Set pGraphs = pActive
858:   pGraphs.Reset
859:   Set pElemProps = pGraphs.Next
860:   Do While Not pElemProps Is Nothing
861:     If TypeOf pElemProps Is IPolygonElement Then
862:       If UCase(pElemProps.Name) = "DSMAPBOOK CLIP ELEMENT" Then
863:         pGraphs.DeleteElement pElemProps
864:         Exit Do
865:       End If
866:     End If
867:     Set pElemProps = pGraphs.Next
868:   Loop
  
870:   Set pElem = New PolygonElement
871:   Set pPoly = m_pPageShape
872:   Set pNewElem = New PolygonElement
873:   Set pNewPoly = New Polygon
874:   pNewPoly.AddPoint pFeatLayer.AreaOfInterest.LowerLeft
875:   pNewPoly.AddPoint pFeatLayer.AreaOfInterest.UpperLeft
876:   pNewPoly.AddPoint pFeatLayer.AreaOfInterest.UpperRight
877:   pNewPoly.AddPoint pFeatLayer.AreaOfInterest.LowerRight
878:   Set pPoly2 = pNewPoly
879:   pPoly2.Close
  
  Dim pLineSym As ISimpleLineSymbol, pLineFillSym As ILineFillSymbol
  Dim pFillShape As IFillShapeElement, pColor As IGrayColor
883:   Set pColor = New GrayColor
884:   pColor.Level = 150
885:   Set pLineSym = New SimpleLineSymbol
886:   pLineSym.Color = pColor
887:   Set pLineFillSym = New LineFillSymbol
888:   pLineFillSym.Angle = 45
889:   pLineFillSym.Color = pColor
890:   pLineFillSym.Outline = pLineSym
891:   Set pLineFillSym.LineSymbol = pLineSym
892:   pLineFillSym.Separation = 5
  
894:   Set pTopoOp = pPoly2
895:   Set pFinalGeom = pTopoOp.Difference(pPoly)
896:   pNewElem.Geometry = pFinalGeom
897:   Set pFillShape = pNewElem
898:   pFillShape.Symbol = pLineFillSym
899:   Set pElemProps = pFillShape
900:   pElemProps.Name = "DSMapBook Clip Element"
901:   pGraphs.AddElement pNewElem, 0

  Exit Sub
ErrHand:
905:   MsgBox "CreateClipElement - " & Erl & " - " & Err.Description
End Sub

Private Sub SetMapExtent(pSeriesOpts As IDSMapSeriesOptions, pActiveView As IActiveView)
On Error GoTo ErrHand:
'Routine for calculating the extent of the tile to be displayed in the layout
  Dim dMult As Double, pEnv As IEnvelope, pMap As IMap
  
913:   Set pMap = pActiveView
  Select Case pSeriesOpts.ExtentType
  Case 0  'Variable
916:     If pSeriesOpts.Margin > 0 Then
917:       Set pEnv = m_pPageShape.Envelope
      Select Case pSeriesOpts.MarginType
      Case 0  'Percent
920:         dMult = 1 + (pSeriesOpts.Margin / 100)
921:         pEnv.Expand dMult, dMult, True
      Case 1  'mapunits
923:         pEnv.Expand pSeriesOpts.Margin, pSeriesOpts.Margin, False
924:       End Select
925:       pActiveView.Extent = pEnv
926:     Else
927:       pActiveView.Extent = m_pPageShape.Envelope
928:     End If
  Case 1  'Fixed
930:     pActiveView.Extent = m_pPageShape.Envelope
931:     pMap.MapScale = pSeriesOpts.FixedScale
  Case 2  'DataDriven
933:     pActiveView.Extent = m_pPageShape.Envelope
934:     pMap.MapScale = m_dPageScale
935:   End Select

  Exit Sub
ErrHand:
939:   MsgBox "SetMapExtent - " & Err.Description
End Sub

Public Function FindDataFrame(pDoc As IMxDocument, sFrameName As String) As IMap
On Error GoTo ErrHand:
  Dim lLoop As Long, pMap As IMap
  
  'Find the data frame
947:   For lLoop = 0 To pDoc.Maps.Count - 1
948:     If pDoc.Maps.Item(lLoop).Name = sFrameName Then
949:       Set pMap = pDoc.Maps.Item(lLoop)
950:       Exit For
951:     End If
952:   Next lLoop
953:   If Not pMap Is Nothing Then
954:     Set FindDataFrame = pMap
955:   End If

  Exit Function
ErrHand:
959:   MsgBox "FindDataFrame - " & Err.Description
End Function

Private Function CompositeLayer1(pCompLayer As ICompositeLayer, sIndexName As String) As IFeatureLayer
On Error GoTo ErrHand:
  Dim lLoop As Long, pFeatLayer As IFeatureLayer
965:   For lLoop = 0 To pCompLayer.Count - 1
966:     If TypeOf pCompLayer.Layer(lLoop) Is ICompositeLayer Then
967:       Set pFeatLayer = CompositeLayer1(pCompLayer.Layer(lLoop), sIndexName)
968:       If Not pFeatLayer Is Nothing Then
969:         Set CompositeLayer1 = pFeatLayer
        Exit Function
971:       End If
972:     Else
973:       If pCompLayer.Layer(lLoop).Name = sIndexName Then
974:         Set CompositeLayer1 = pCompLayer.Layer(lLoop)
        Exit Function
976:       End If
977:     End If
978:   Next lLoop
  
980:   Set CompositeLayer1 = Nothing

  Exit Function
ErrHand:
984:   MsgBox "CompositeLayer - " & Err.Description
End Function

Private Function FindLayer(sLayerName As String, pMap As IMap) As IFeatureLayer
' Routine for finding a layer based on a name and then returning that layer as
' a IFeatureLayer
On Error GoTo ErrHand:
  Dim lLoop As Integer
  Dim pFLayer As IFeatureLayer

994:   For lLoop = 0 To pMap.LayerCount - 1
995:     If TypeOf pMap.Layer(lLoop) Is ICompositeLayer Then
996:       Set pFLayer = FindCompositeLayer(pMap.Layer(lLoop), sLayerName, pMap)
997:       If Not pFLayer Is Nothing Then
998:         Set FindLayer = pFLayer
        Exit Function
1000:       End If
1001:     ElseIf TypeOf pMap.Layer(lLoop) Is IFeatureLayer Then
1002:       Set pFLayer = pMap.Layer(lLoop)
1003:       If UCase(pFLayer.Name) = UCase(sLayerName) Then
1004:         Set FindLayer = pFLayer
        Exit Function
1006:       End If
1007:     End If
1008:   Next lLoop
  
1010:   Set FindLayer = Nothing
  
  Exit Function
  
ErrHand:
1015:   MsgBox "FindLayer - " & Erl & " - " & Err.Description
End Function

Private Function FindCompositeLayer(pCompLayer As ICompositeLayer, sLayerName As String, pMap As IMap) As IFeatureLayer
On Error GoTo ErrHand:
  Dim lLoop As Long, pFeatLayer As IFeatureLayer
1021:   For lLoop = 0 To pCompLayer.Count - 1
1022:     If TypeOf pCompLayer.Layer(lLoop) Is ICompositeLayer Then
1023:       Set pFeatLayer = FindCompositeLayer(pCompLayer.Layer(lLoop), sLayerName, pMap)
1024:       If Not pFeatLayer Is Nothing Then
1025:         Set FindCompositeLayer = pFeatLayer
        Exit Function
1027:       End If
1028:     Else
1029:       If TypeOf pCompLayer.Layer(lLoop) Is IFeatureLayer Then
1030:         If UCase(pCompLayer.Layer(lLoop).Name) = UCase(sLayerName) Then
1031:           Set FindCompositeLayer = pCompLayer.Layer(lLoop)
          Exit Function
1033:         End If
1034:       End If
1035:     End If
1036:   Next lLoop

  Exit Function
ErrHand:
1040:   MsgBox "CompositeLayer - " & Erl & " - " & Err.Description
End Function
