VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCreateStripMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

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

'----------------------------------------------
' Properties
' - DestinationFeatureClass (IFeatureClass, r/w)
' - MapScale (Double, r/w)
' - FrameWidthInPageUnits (Double, r/w)
' - FrameHeightInPageUnits (Double, r/w)
' - StripMapRoute (IPolyline, r/w)
'----------------------------------------------
' Methods
' - GenerateGrids(pApp as IApplication)
'    : Generates the grids using the values added.
' - RunStandardGUI(pApp as IApplication)
'    : To open the form without having the button
'      added to ArcMap's GUI.
'----------------------------------------------
' Local Global Declarations
Private m_DestFL As IFeatureLayer
Private m_DestFC As IFeatureClass
Private m_Polyline As IPolyline
Private m_dMapScale As Double
Private m_dFrameWidthInPageUnits As Double
Private m_dFrameHeightInPageUnits As Double
Private m_FldStripName As String
Private m_FldNumInSeries As String
Private m_FldMapAngle As String
Private m_FldScale As String
Private m_RemoveGrids As Boolean
Private m_Flip As Boolean
Private m_GridWidth As Double
Private m_GridHeight As Double
Private m_StripMapName As String
Private m_pProgress As IModelessFrame
'----------------------------------------------
' API call to keep form top most
Private Const GWL_HWNDPARENT = -8
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
  (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Property Set DestinationFeatureLayer(pFL As IFeatureLayer)
42:     If pFL.FeatureClass.ShapeType = esriGeometryPolygon Then
43:         Set m_DestFL = pFL
44:         Set m_DestFC = pFL.FeatureClass
45:     Else
46:         Err.Raise vbObjectError, "MapGridManager_Set_DestinationFeatureLayer", _
            "Not a polygon feature layer"
48:     End If
End Property

Public Property Get DestinationFeatureLayer() As IFeatureLayer
52:     Set DestinationFeatureLayer = m_DestFL
End Property

Public Property Set StripMapRoute(RoutePolyline As IPolyline)
56:     Set m_Polyline = RoutePolyline
End Property

Public Property Get StripMapRoute() As IPolyline
60:     Set StripMapRoute = m_Polyline
End Property

Public Property Let FrameWidthInPageUnits(dWidth As Double)
64:     m_dFrameWidthInPageUnits = dWidth
End Property

Public Property Get FrameWidthInPageUnits() As Double
68:     FrameWidthInPageUnits = m_dFrameWidthInPageUnits
End Property

Public Property Let FrameHeightInPageUnits(dHeight As Double)
72:     m_dFrameHeightInPageUnits = dHeight
End Property

Public Property Get FrameHeightInPageUnits() As Double
76:     FrameHeightInPageUnits = m_dFrameHeightInPageUnits
End Property

Public Property Let MapScale(dScale As Double)
80:     m_dMapScale = dScale
End Property

Public Property Get MapScale() As Double
84:     MapScale = m_dMapScale
End Property

Public Property Let StripMapName(MapName As String)
88:     m_StripMapName = MapName
End Property

Public Property Get StripMapName() As String
92:     StripMapName = m_StripMapName
End Property


Public Property Let FieldNameStripMapName(FieldName As String)
97:     m_FldStripName = FieldName
End Property

Public Property Get FieldNameStripMapName() As String
101:     FieldNameStripMapName = m_FldStripName
End Property

Public Property Let FieldNameNumberInSeries(FieldName As String)
105:     m_FldNumInSeries = FieldName
End Property

Public Property Get FieldNameNumberInSeries() As String
109:     FieldNameNumberInSeries = m_FldNumInSeries
End Property

Public Property Let FieldNameMapAngle(FieldName As String)
113:     m_FldMapAngle = FieldName
End Property

Public Property Get FieldNameMapAngle() As String
117:     FieldNameMapAngle = m_FldMapAngle
End Property

Public Property Let FieldNameScale(FieldName As String)
121:     m_FldScale = FieldName
End Property

Public Property Get FieldNameScale() As String
125:     FieldNameScale = m_FldScale
End Property

Public Property Let RemoveCurrentGrids(RemoveGrids As Boolean)
129:     m_RemoveGrids = RemoveGrids
End Property

Public Property Get RemoveCurrentGrids() As Boolean
133:     RemoveCurrentGrids = m_RemoveGrids
End Property

Public Property Let FlipPolyline(Flip As Boolean)
137:     m_Flip = Flip
End Property

Public Property Get FlipPolyline() As Boolean
141:     FlipPolyline = m_Flip
End Property

Private Sub Class_Initialize()
    ' Set the defaults
146:     Set m_pProgress = New ModelessFrame
End Sub

Public Sub RunStandardGUI(pApp As IApplication)
150:     Set frmSMapSettings.m_Application = pApp
151:     frmSMapSettings.Tickle
152:     SetWindowLong frmSMapSettings.hwnd, GWL_HWNDPARENT, pApp.hwnd
153:     frmSMapSettings.Show vbModeless
End Sub

Private Function CalculatePageToMapRatio(pApp As IApplication) As Double
    Dim pMx As IMxDocument
    Dim pPage As IPage
    Dim pPageUnits As esriUnits
    Dim pSR As ISpatialReference
    Dim pSRI As ISpatialReferenceInfo
    Dim pPCS As IProjectedCoordinateSystem
    Dim dMetersPerUnit As Double
    Dim dCurrScale As Double
    Dim pExtentEnv As IEnvelope
    Dim dEndX As Double, dEndY As Double
    Dim dStartX As Double, dStartY As Double
    
    On Error GoTo eh
    
    ' Init
172:     Set pMx = pApp.Document
173:     Set pSR = pMx.FocusMap.SpatialReference
    ' If a Projected coord system
175:     If TypeOf pSR Is IProjectedCoordinateSystem Then
        ' Use meters per unit as the conversion
177:         Set pPCS = pSR
178:         dMetersPerUnit = pPCS.CoordinateUnit.MetersPerUnit
        ' Now convert this into page (ie: paper) units
180:         Set pPage = pMx.PageLayout.Page
181:         pPageUnits = pPage.Units
        Select Case pPageUnits
            Case esriInches: CalculatePageToMapRatio = dMetersPerUnit / (1 / 12 * 0.304800609601219)
            Case esriFeet: CalculatePageToMapRatio = dMetersPerUnit / (0.304800609601219)
            Case esriCentimeters: CalculatePageToMapRatio = dMetersPerUnit / (1 / 100)
            Case esriMeters: CalculatePageToMapRatio = dMetersPerUnit / (1)
            Case Else:
188:                 MsgBox "Warning: Only the following Page (Layout) Units are supported by this tool:" _
                    & vbCrLf & " - Inches, Feet, Centimeters, Meters" _
                    & vbCrLf & vbCrLf & "Calculating as though Page Units are in Inches..."
191:                 CalculatePageToMapRatio = dMetersPerUnit / (1 / 12 * 0.304800609601219)
192:         End Select
    ' Otherwise
194:     Else
        ' If not projected, we can only do a "flat" conversion -> that is, use the current scale and extent
        '  as a ratio to be applied to the map grid scale.
        ' NOTE: We MUST be in Layout mode to make this calculation, as the scale in Map View and Layout View
        '  are not the same (as the extent envelope and data frame envelope can be different shapes).  The
        '  test for being in Layout Mode is made in the clsMapGridButton.ICommand_Enabled property.
200:         Set pExtentEnv = pMx.ActiveView.Extent
201:         dStartX = pExtentEnv.XMin
202:         dStartY = pExtentEnv.YMin
203:         dEndX = pExtentEnv.XMax
204:         dEndY = pExtentEnv.YMax
        
206:         dCurrScale = pMx.FocusMap.MapScale
207:         If ((dEndX - dStartX) / m_dFrameWidthInPageUnits) > ((dEndY - dStartY) / m_dFrameHeightInPageUnits) Then
208:             CalculatePageToMapRatio = m_dFrameWidthInPageUnits / ((dEndX - dStartX) / dCurrScale)
209:         Else
210:             CalculatePageToMapRatio = m_dFrameHeightInPageUnits / ((dEndY - dStartY) / dCurrScale)
211:         End If
212:     End If
    
    Exit Function
eh:
216:     CalculatePageToMapRatio = 1
217:     MsgBox "Error in CalculatePageToMapRatio" & vbCrLf & Err.Description
End Function

Private Sub Class_Terminate()
221:     Set m_DestFL = Nothing
222:     Set m_DestFC = Nothing
223:     Set m_pProgress = New ModelessFrame
End Sub

Public Sub GenerateStripMap(Application As IApplication)
    Dim pMx As IMxDocument
    Dim pPolyline As IPolyline
    Dim pCenterPoint As IPoint
    Dim pCirclePoly As IPolygon
    Dim pGridPoly As IPolygon
    Dim pCircularArc As IConstructCircularArc
    Dim pSegmentCollection As ISegmentCollection
    Dim pTopoOpt As ITopologicalOperator
    Dim pGeoCol As IGeometryCollection
    Dim pIntersectPoint As IPoint
    Dim pArc As ICurve
    Dim pIntersectPointPrev As IPoint
    Dim bFirstRun As Boolean
    Dim lLoop2 As Long
    Dim dHighest As Double, lHighestRef As Long
    Dim dHighestPrev As Double
    Dim pCurve As ICurve, pLine As ILine
    Dim pPLine As IPolyline
    Dim bContinue As Boolean
    Dim dGridAngle As Double
    Dim bReducedRadius As Boolean
    Dim lCounter As Long
    Dim dHighestThisTurn As Double
    
    Dim pWorkspaceEdit As IWorkspaceEdit
    Dim lLoop As Long
    Dim pFeatDataset As IFeatureDataset
    Dim pFeature As IFeature
    Dim pFeatCur As IFeatureCursor
    Dim pSourcePolygon As IPolygon
    'Dim pGridPolygon As IPolygon
    Dim pPointColl As IPointCollection
    Dim pStartingCoord As IPoint
    Dim pPoint As IPoint
    Dim lRow As Long
    Dim lCol As Long
    Dim lRowCount As Long
    Dim lColCount As Long
    Dim pClone As IClone
    Dim dGridSizeW As Double
    Dim dGridSizeH As Double
    Dim pTransform As ITransform2D
    Dim bOKToAdd As Boolean
    Dim iStringLengthRow As Integer
    Dim iStringLengthCol As Integer
    Dim pDataset As IDataset
    Dim lBase As Long
    Dim dDataFrameWidth As Double
    Dim dDataFrameHeight As Double
    Dim dConvertPageToMapUnits As Double
    Dim dIncrement As Double
    Dim pInsertFeatureBuffer As IFeatureBuffer
    Dim pInsertFeatureCursor As IFeatureCursor
    Dim pFL As IFeatureLayer
    Dim pFC As IFeatureClass
    Dim pProgress As frmProgress

    On Error GoTo eh

    ' Set mouse pointer
287:     Screen.MousePointer = vbArrowHourglass

    ' Init
290:     Set pMx = Application.Document
291:     For lLoop = 0 To pMx.FocusMap.LayerCount - 1
292:         If TypeOf pMx.FocusMap.Layer(lLoop) Is IFeatureLayer Then
293:             If UCase(pMx.FocusMap.Layer(lLoop).Name) = UCase(m_DestFL.Name) Then
294:                 Set pFL = pMx.FocusMap.Layer(lLoop)
295:                 Exit For
296:             End If
297:         End If
298:     Next
299:     If pFL Is Nothing Then
300:         MsgBox "No match in the Map for layer '" & m_DestFL.Name & "'."
        Exit Sub
302:     End If
303:     Set pFC = pFL.FeatureClass
    ' Check for required fields - that the field exists
    Dim bErrorWithFields As Boolean
306:     bErrorWithFields = (pFC.FindField(m_FldStripName) < 0)
307:     bErrorWithFields = bErrorWithFields Or (pFC.FindField(m_FldNumInSeries) < 0)
308:     bErrorWithFields = bErrorWithFields Or (pFC.FindField(m_FldMapAngle) < 0)
309:     If Len(m_FldScale) > 0 Then bErrorWithFields = bErrorWithFields Or (pFC.FindField(m_FldScale) < 0)
    ' If error
311:     If bErrorWithFields Then
312:         Err.Raise vbObjectError, "GenerateStripMap", "Could not find all the given field names in " & pFL.Name & "." _
            & vbCrLf & " - " & m_FldStripName & ", " & m_FldNumInSeries & ", " & m_FldMapAngle & ", " & m_FldScale
314:     End If
    ' Check the field types
316:     bErrorWithFields = (pFC.Fields.Field(pFC.FindField(m_FldStripName)).Type <> esriFieldTypeString)
317:     bErrorWithFields = bErrorWithFields Or _
       ((pFC.Fields.Field(pFC.FindField(m_FldNumInSeries)).Type <> esriFieldTypeDouble) And _
        (pFC.Fields.Field(pFC.FindField(m_FldNumInSeries)).Type <> esriFieldTypeInteger) And _
        (pFC.Fields.Field(pFC.FindField(m_FldNumInSeries)).Type <> esriFieldTypeSingle) And _
        (pFC.Fields.Field(pFC.FindField(m_FldNumInSeries)).Type <> esriFieldTypeSmallInteger))
322:     bErrorWithFields = bErrorWithFields Or _
       ((pFC.Fields.Field(pFC.FindField(m_FldMapAngle)).Type <> esriFieldTypeDouble) And _
        (pFC.Fields.Field(pFC.FindField(m_FldMapAngle)).Type <> esriFieldTypeInteger) And _
        (pFC.Fields.Field(pFC.FindField(m_FldMapAngle)).Type <> esriFieldTypeSingle) And _
        (pFC.Fields.Field(pFC.FindField(m_FldMapAngle)).Type <> esriFieldTypeSmallInteger))
327:     If Len(m_FldScale) > 0 Then
328:         bErrorWithFields = bErrorWithFields Or _
           ((pFC.Fields.Field(pFC.FindField(m_FldScale)).Type <> esriFieldTypeDouble) And _
            (pFC.Fields.Field(pFC.FindField(m_FldScale)).Type <> esriFieldTypeInteger) And _
            (pFC.Fields.Field(pFC.FindField(m_FldScale)).Type <> esriFieldTypeSingle) And _
            (pFC.Fields.Field(pFC.FindField(m_FldScale)).Type <> esriFieldTypeSmallInteger))
333:     End If
    ' if error
335:     If bErrorWithFields Then
336:         Err.Raise vbObjectError, "GenerateStripMap", "Given field names are not of the correct type." _
            & vbCrLf & "Strip Map Name field must be a Text field, all others must be numeric fields."
338:     End If
    ' Get the dataset and workspace (to start editing upon)
340:     Set pFeatDataset = pFC.FeatureDataset
341:     If Not pFeatDataset Is Nothing Then
342:         Set pWorkspaceEdit = pFeatDataset.Workspace
343:     Else
        ' Is a shapefile, go via just IDataset
345:         Set pDataset = pFC
346:         Set pWorkspaceEdit = pDataset.Workspace
347:     End If
    
'    ' If replacing, delete all existing polygons
350:     Set pProgress = New frmProgress
351:     m_pProgress.Create pProgress
352:     pProgress.ProgressBar1.Min = 0
353:     pProgress.ProgressBar1.Max = 100
354:     pProgress.ProgressBar1.Value = 0
355:     If m_RemoveGrids Then
        Dim pFCu As IFeatureCursor
        Dim pT As ITable
358:         Set pFCu = m_DestFL.Search(Nothing, False)
359:         Set pT = m_DestFL.FeatureClass
360:         If pT.RowCount(Nothing) = 0 Then
361:             dIncrement = 99
362:         Else
363:             dIncrement = 100 / pT.RowCount(Nothing)
364:         End If
365:         pProgress.lblInformation.Caption = "Deleting previous grids..."
366:         pProgress.cmdCancel.Visible = False        ' User cannot cancel this step
367:         m_pProgress.Visible = True
368:         Set pFeature = pFCu.NextFeature
369:         While Not pFeature Is Nothing
370:             pFeature.Delete
371:             If (pProgress.ProgressBar1.Value + dIncrement) <= pProgress.ProgressBar1.Max Then
372:                 pProgress.ProgressBar1.Value = pProgress.ProgressBar1.Value + dIncrement
373:             Else
374:                 pProgress.lblInformation.Caption = "Warning: Val > Max (" & pProgress.ProgressBar1.Max & ")"
375:             End If
376:             Set pFeature = pFCu.NextFeature
377:         Wend
378:         m_pProgress.Visible = False
379:     End If
    
        
    ' Init strip map stuff
383:     Set pPolyline = m_Polyline
    ' Flip, if required
385:     If m_Flip Then
386:         pPolyline.ReverseOrientation
387:     End If
388:     Set pCenterPoint = pPolyline.FromPoint
    ' Get the progress bar ready
390:     pProgress.ProgressBar1.Min = 0
391:     pProgress.ProgressBar1.Max = 101
392:     pProgress.ProgressBar1.Value = 0
393:     pProgress.lblInformation.Caption = "Creating strip map..."
394:     pProgress.cmdCancel.Visible = True        ' User cannot cancel this step
    ' Get map units size for grids
396:     dConvertPageToMapUnits = CalculatePageToMapRatio(Application)
397:     m_GridWidth = ((m_dMapScale * m_dFrameWidthInPageUnits) / dConvertPageToMapUnits)
398:     m_GridHeight = ((m_dMapScale * m_dFrameHeightInPageUnits) / dConvertPageToMapUnits)
    ' Init for processing
400:     dHighestPrev = -1
401:     bFirstRun = True
402:     Set pArc = pPolyline
403:     Set pInsertFeatureCursor = pFC.Insert(True)
404:     Set pInsertFeatureBuffer = pFC.CreateFeatureBuffer
405:     m_pProgress.Visible = True
406:     Do
        Dim dCircleRadius As Double, colIntersects As Collection, dIntersect As Double
408:         If bFirstRun Then
409:             dCircleRadius = m_GridWidth / 2
410:         Else
411:             dCircleRadius = m_GridWidth
412:         End If
413:         bReducedRadius = False
414:         Do
            ' Create the search circle
416:             Set pCircularArc = New CircularArc
417:             pCircularArc.ConstructCircle pCenterPoint, dCircleRadius, False 'make it clockwise
418:             Set pCirclePoly = New Polygon
419:             Set pSegmentCollection = pCirclePoly
420:             pSegmentCollection.AddSegment pCircularArc
            
            ' Intersect the polyline and the circle
423:             Set pTopoOpt = pPolyline
424:             Set pGeoCol = New GeometryBag
425:             Set pGeoCol = pTopoOpt.Intersect(pCirclePoly, esriGeometry0Dimension)
            
427:             If pGeoCol.GeometryCount = 0 Then
428:                 MsgBox "error - no geoms intersected"
                Exit Sub
430:             End If
431:             Set pArc = pPolyline
432:             lHighestRef = -1
433:             dHighestThisTurn = 102
434:             For lLoop2 = 0 To pGeoCol.GeometryCount - 1
435:                 Set pIntersectPoint = pGeoCol.Geometry(lLoop2)
436:                 dIntersect = ReturnPercentageAlong(pArc, pIntersectPoint)
437:                 If dIntersect > (dHighestPrev * 1.001) And dIntersect < dHighestThisTurn Then
438:                     dHighest = dIntersect
439:                     dHighestThisTurn = dIntersect
440:                     lHighestRef = lLoop2
441:                 End If
442:             Next
            ' If no intersection higher than our previous run, we are at the end.
444:             If lHighestRef < 0 Then
445:                 dHighest = 101
                ' Need to extend the end (tangent) to get intersection
447:                 Set pIntersectPoint = IntersectPointExtendedTo(pPolyline, pCirclePoly)
448:                 Set pIntersectPointPrev = pCenterPoint
            ' Otherwise, still in the middle somewhere
450:             Else
451:                 Set pIntersectPoint = pGeoCol.Geometry(lHighestRef)
                ' If just starting off (ie: first grid)
453:                 If bFirstRun Then
                    ' Set the grid so the polyline's starting point is in the
                    '  center of the first grid polygon we make
456:                     Set pIntersectPointPrev = New esrigeometry.Point
457:                     pIntersectPointPrev.PutCoords pCenterPoint.X - (pIntersectPoint.X - pCenterPoint.X), _
                                                  pCenterPoint.Y - (pIntersectPoint.Y - pCenterPoint.Y)
                ' Otherwise, we already have a previous point
460:                 Else
                    ' So use it
462:                     Set pIntersectPointPrev = pCenterPoint
463:                 End If
464:             End If
            ' Make our grid polygon, allowing for any 'shrunken' grids
466:             If bReducedRadius Then
                Dim pTmpPLine As IPolyline
                Dim pTmpCPoly As IPolygon
                Dim pTmpIntPoint As IPoint
470:                 Set pCircularArc = New CircularArc
471:                 If bFirstRun Then
472:                     pCircularArc.ConstructCircle pCenterPoint, m_GridWidth / 2, False 'make it clockwise
473:                 Else
474:                     pCircularArc.ConstructCircle pCenterPoint, m_GridWidth, False 'make it clockwise
475:                 End If
476:                 Set pTmpCPoly = New Polygon
477:                 Set pSegmentCollection = pTmpCPoly
478:                 pSegmentCollection.AddSegment pCircularArc
                
480:                 Set pTmpPLine = New Polyline
481:                 pTmpPLine.FromPoint = pIntersectPointPrev
482:                 pTmpPLine.ToPoint = pIntersectPoint
483:                 Set pTmpIntPoint = IntersectPointExtendedTo(pTmpPLine, pTmpCPoly)
484:                 CreateAngledGridPolygon pIntersectPointPrev, pTmpIntPoint, pGridPoly, dGridAngle
485:             Else
486:                 CreateAngledGridPolygon pIntersectPointPrev, pIntersectPoint, pGridPoly, dGridAngle
487:             End If
            ' Now, we potentially need to reprocess if the route dips out of our grid
489:             Set pTopoOpt = pGridPoly
490:             Set pGeoCol = New GeometryBag
491:             Set pGeoCol = pTopoOpt.Intersect(pPolyline, esriGeometry0Dimension)
492:             bContinue = True
493:             If pGeoCol.GeometryCount > 2 Then
494:                 Set colIntersects = New Collection
495:                 For lLoop2 = 0 To pGeoCol.GeometryCount - 1
496:                     colIntersects.Add ReturnPercentageAlong(pArc, pGeoCol.Geometry(lLoop2))
497:                 Next
498:                 For lLoop2 = 1 To colIntersects.count
499:                     If colIntersects.Item(lLoop2) > (dHighestPrev * 1.001) And colIntersects.Item(lLoop2) < (dHighest * 0.999) Then
500:                         bContinue = False
501:                         dHighest = dHighestPrev
502:                         dCircleRadius = dCircleRadius - (m_GridWidth * 0.1)
503:                         bReducedRadius = True
504:                         If dCircleRadius <= 0 Then
505:                             bContinue = True
506:                         End If
507:                         Exit For
508:                     End If
509:                 Next
510:             End If
            ' If all OK and a reduced radius, look for a quick jump ahead
512:             If bContinue And bReducedRadius Then 'And pGeoCol.GeometryCount <= 2 Then
                Dim dTmpHighest As Double
514:                 Set pArc = pPolyline
515:                 lHighestRef = -1
516:                 dTmpHighest = -1
517:                 For lLoop2 = 0 To pGeoCol.GeometryCount - 1
518:                     Set pIntersectPoint = pGeoCol.Geometry(lLoop2)
519:                     dIntersect = ReturnPercentageAlong(pArc, pIntersectPoint)
520:                     If dIntersect > dTmpHighest Then
521:                         dTmpHighest = dIntersect
522:                         lHighestRef = lLoop2
523:                     End If
524:                 Next
525:                 If lHighestRef >= 0 Then Set pIntersectPoint = pGeoCol.Geometry(lHighestRef)
526:                 dHighest = dTmpHighest
527:             End If
528:         Loop Until bContinue
        
530:         bFirstRun = False
531:         dHighestPrev = dHighest
        
        ' All OK to create our grid feature now (hopefully, anyway)
534:         lCounter = lCounter + 1
        'CreateGridFeaturesAsGraphics pGridPoly, lCounter, dGridAngle, Application 'AAA
        
        ' Create new grid cell feature
538:         Set pInsertFeatureBuffer.Shape = pGridPoly
539:         pInsertFeatureBuffer.Value(pFC.Fields.FindField(m_FldStripName)) = m_StripMapName & CStr(lCounter)
540:         pInsertFeatureBuffer.Value(pFC.Fields.FindField(m_FldNumInSeries)) = lCounter
541:         pInsertFeatureBuffer.Value(pFC.Fields.FindField(m_FldMapAngle)) = dGridAngle 'degrees
542:         If Len(m_FldScale) > 0 Then pInsertFeatureBuffer.Value(pFC.Fields.FindField(m_FldScale)) = m_dMapScale
543:         pInsertFeatureCursor.InsertFeature pInsertFeatureBuffer
544:         If dHighest <= pProgress.ProgressBar1.Max Then
545:             pProgress.ProgressBar1.Value = dHighest
546:         Else
547:             pProgress.lblInformation.Caption = "Warning: Val > Max (" & pProgress.ProgressBar1.Max & ")"
548:             pProgress.ProgressBar1.Value = pProgress.ProgressBar1.Max
549:         End If
550:         If (lCounter Mod 20 = 0) Then
551:             DoEvents
552:             pInsertFeatureCursor.Flush
553:         End If
554:         pProgress.Refresh
555:         If pProgress.Cancelled Then
            Dim vUserChoice
557:             pProgress.Cancelled = False       ' Reset the form
558:             vUserChoice = MsgBox("Operation cancelled." _
                & "  Save the edits made thus far?" & vbCrLf & vbCrLf _
                & "(Click Cancel to continue processing)", _
                            vbYesNoCancel, "Generate Strip Map")
562:             If vUserChoice <> vbCancel Then
563:                 GoTo CancelledGenerateGrids     'Sorry for GoTo usage - in a hurry
564:             End If
565:         End If
        ' For next time
567:         Set pCenterPoint = pIntersectPoint
568:     Loop While dHighest < 100
    ' Add remainder polys
570:     pInsertFeatureCursor.Flush
571:     m_pProgress.Visible = False

    ' Stop editing
574:     pWorkspaceEdit.StopEditOperation
575:     pWorkspaceEdit.StopEditing True
576:     pMx.ActiveView.Refresh
    
    Exit Sub
    
CancelledGenerateGrids:
581:     m_pProgress.Visible = False
582:     If vUserChoice = vbYes Then
583:         pInsertFeatureCursor.Flush
584:         pWorkspaceEdit.StopEditOperation
585:         pWorkspaceEdit.StopEditing True
586:     Else
587:         pWorkspaceEdit.StopEditOperation
588:         pWorkspaceEdit.StopEditing False
589:     End If
590:     Screen.MousePointer = vbDefault
591:     pMx.ActiveView.Refresh
    Exit Sub
    
594:     Resume
eh:
596:     MsgBox "Error in GenerateStripMap:" & vbCrLf & Err.Description
End Sub

Private Sub CreateGridFeaturesAsGraphics(pGridPolygon As IPolygon, lIndex As Long, dAngle As Double, pApp As IApplication)
    Dim pPntColl As IPointCollection
    Dim pArea As IArea
    Dim pCentroid As IPoint
    Dim lLoop As Long
    
    ' Create graphics (TEST PHASE)
606:     Set pPntColl = pGridPolygon
607:     For lLoop = 0 To pPntColl.PointCount - 2
608:         Perm_DrawLineFromPoints pPntColl.Point(lLoop), pPntColl.Point(lLoop + 1), pApp
609:     Next
610:     Perm_DrawLineFromPoints pPntColl.Point(0), pPntColl.Point(pPntColl.PointCount - 1), pApp
611:     Set pArea = pGridPolygon
612:     Set pCentroid = pArea.Centroid
613:     Perm_DrawTextFromPoint pCentroid, CStr(lIndex), pApp, , , , , 8
614:     pCentroid.Y = pCentroid.Y - (m_GridWidth / 3)
615:     Perm_DrawTextFromPoint pCentroid, Format(dAngle / cPI * 180, "(#0.0)"), pApp, , , , , 8

End Sub

Private Function ReturnPercentageAlong(ByVal pArc As ICurve, ByVal pPoint As IPoint) As Double
    Dim GeoCount As Long
    Dim pDistAlong As Double
    Dim pDist As Double
    Dim pRightSide As Boolean
    Dim pOutPt As IPoint
    Dim CompareDist As Double
    
    On Error GoTo ErrorHandler
    
629:     CompareDist = 0
    'Find the distance along curve
631:     Set pOutPt = New esrigeometry.Point
632:     pArc.QueryPointAndDistance esriNoExtension, pPoint, True, pOutPt, _
                        pDistAlong, pDist, pRightSide
    
635:     ReturnPercentageAlong = (pDistAlong * 100)
    Exit Function
ErrorHandler:
638:     Err.Raise Err.Number, "ReturnPercentageAlong", "Error in ReturnPercentageAlong." _
        & vbCrLf & "Err " & Err.Number & ": " & Err.Description
End Function

Private Sub CreateAngledGridPolygon(ByVal p1 As IPoint, ByVal p2 As IPoint, _
            ByRef ReturnedGrid As IPolygon, ByRef ReturnedAngleRadians As Double)
    Dim pPointColl As IPointCollection
    Dim pPointStart As IPoint
    Dim pPoint As IPoint
    Dim dAngleInRadians As Double
    Dim pLine As ILine
    
    On Error GoTo eh
    
    ' Init
653:     Set pLine = New esrigeometry.Line
654:     pLine.FromPoint = p1
655:     pLine.ToPoint = p2
656:     dAngleInRadians = pLine.Angle
657:     If dAngleInRadians = 0 Then
658:       ReturnedAngleRadians = 0
659:     ElseIf dAngleInRadians > 0 Then
660:       ReturnedAngleRadians = 360 - ((dAngleInRadians / cPI) * 180)
661:     Else
662:       ReturnedAngleRadians = Abs((dAngleInRadians / cPI) * 180)
663:     End If
664:     Set ReturnedGrid = New Polygon
665:     Set pPointColl = ReturnedGrid
    ' POINT 1 -------------------------------------------
667:     Set pPoint = New esrigeometry.Point
668:     pPoint.PutCoords p1.X + (Sin(dAngleInRadians) * (m_GridHeight / 2)), _
                     p1.Y - (Cos(dAngleInRadians) * (m_GridHeight / 2))
670:     pPointColl.AddPoint pPoint
671:     Set pPointStart = pPoint
    ' POINT 2 -------------------------------------------
673:     Set pPoint = New esrigeometry.Point
674:     pPoint.PutCoords p1.X - (Sin(dAngleInRadians) * (m_GridHeight / 2)), _
                     p1.Y + (Cos(dAngleInRadians) * (m_GridHeight / 2))
676:     pPointColl.AddPoint pPoint
    ' POINT 3 -------------------------------------------
678:     Set pPoint = New esrigeometry.Point
679:     pPoint.PutCoords p2.X - Sin(dAngleInRadians) * m_GridHeight / 2, _
                     p2.Y + Cos(dAngleInRadians) * m_GridHeight / 2
681:     pPointColl.AddPoint pPoint
    ' POINT 4 -------------------------------------------
683:     Set pPoint = New esrigeometry.Point
684:     pPoint.PutCoords p2.X + Sin(dAngleInRadians) * m_GridHeight / 2, _
                     p2.Y - Cos(dAngleInRadians) * m_GridHeight / 2
686:     pPointColl.AddPoint pPoint
    ' JOIN BACK TO FIRST (CLOSE POLYGON) ----------------
688:     pPointColl.AddPoint pPointStart
    
    Exit Sub
eh:
692:     Err.Raise Err.Number, Err.Source, "Error in CreateAngledGridPolygon." _
        & "Err " & Err.Number & ": " & Err.Description
End Sub

Public Sub Perm_DrawPoint(ByVal pPoint As IPoint, Application As IApplication, _
            Optional sElementName As String = "DEMO_TEMPORARY", _
            Optional dRed As Double = 255, Optional dGreen As Double = 0, _
            Optional dBlue As Double = 0, Optional dSize As Double = 6)
' Add a permanent graphic dot on the display at the given point location
    Dim pColor As IRgbColor
    Dim pMarker As ISimpleMarkerSymbol
    Dim pGLayer As IGraphicsLayer
    Dim pGCon As IGraphicsContainer
    Dim pElement As IElement
    Dim pMarkerElement As IMarkerElement
    Dim pElementProp As IElementProperties
    Dim pMx As IMxDocument
    
    ' Init
711:     Set pMx = Application.Document
712:     Set pGLayer = pMx.FocusMap.BasicGraphicsLayer
713:     Set pGCon = pGLayer
714:     Set pElement = New MarkerElement
715:     pElement.Geometry = pPoint
716:     Set pMarkerElement = pElement
    
    ' Set the symbol
719:     Set pColor = New RgbColor
720:     pColor.Red = dRed
721:     pColor.Green = dGreen
722:     pColor.Blue = dBlue
723:     Set pMarker = New SimpleMarkerSymbol
724:     With pMarker
725:         .Color = pColor
726:         .Size = dSize
727:     End With
728:     pMarkerElement.Symbol = pMarker
    
    ' Add the graphic
731:     Set pElementProp = pElement
732:     pElementProp.Name = sElementName
733:     pGCon.AddElement pElement, 0
End Sub

Public Sub Perm_DrawLineFromPoints(ByVal pFromPoint As IPoint, ByVal pToPoint As IPoint, _
            Application As IApplication, _
            Optional sElementName As String = "DEMO_TEMPORARY", _
            Optional dRed As Double = 0, Optional dGreen As Double = 0, _
            Optional dBlue As Double = 255, Optional dSize As Double = 1)
' Add a permanent graphic line on the display, using the From and To points supplied
    Dim pLnSym As ISimpleLineSymbol
    Dim pLine1 As ILine
    Dim pSeg1 As ISegment
    Dim pPolyline As ISegmentCollection
    Dim myColor As IRgbColor
    Dim pSym As ISymbol
    Dim pLineSym As ILineSymbol
    Dim pGLayer As IGraphicsLayer
    Dim pGCon As IGraphicsContainer
    Dim pElement As IElement
    Dim pLineElement As ILineElement
    Dim pElementProp As IElementProperties
    Dim pMx As IMxDocument
    
    ' Init
757:     Set pMx = Application.Document
758:     Set pGLayer = pMx.FocusMap.BasicGraphicsLayer
759:     Set pGCon = pGLayer
760:     Set pElement = New LineElement
    
    ' Set the line symbol
763:     Set pLnSym = New SimpleLineSymbol
764:     Set myColor = New RgbColor
765:     myColor.Red = dRed
766:     myColor.Green = dGreen
767:     myColor.Blue = dBlue
768:     pLnSym.Color = myColor
769:     pLnSym.Width = dSize
    
    ' Create a standard polyline (via 2 points)
772:     Set pLine1 = New esrigeometry.Line
773:     pLine1.PutCoords pFromPoint, pToPoint
774:     Set pSeg1 = pLine1
775:     Set pPolyline = New Polyline
776:     pPolyline.AddSegment pSeg1
777:     pElement.Geometry = pPolyline
778:     Set pLineElement = pElement
779:     pLineElement.Symbol = pLnSym
    
    ' Add the graphic
782:     Set pElementProp = pElement
783:     pElementProp.Name = sElementName
784:     pGCon.AddElement pElement, 0
End Sub

Public Sub Perm_DrawTextFromPoint(pPoint As IPoint, sText As String, Application As IApplication, _
            Optional sElementName As String = "DEMO_TEMPORARY", _
            Optional dRed As Double = 50, Optional dGreen As Double = 50, _
            Optional dBlue As Double = 50, Optional dSize As Double = 10)
' Add permanent graphic text on the display at the given point location
    Dim myTxtSym As ITextSymbol
    Dim myColor As IRgbColor
    Dim pGLayer As IGraphicsLayer
    Dim pGCon As IGraphicsContainer
    Dim pElement As IElement
    Dim pTextElement As ITextElement
    Dim pElementProp As IElementProperties
    Dim pMx As IMxDocument
    
    ' Init
802:     Set pMx = Application.Document
803:     Set pGLayer = pMx.FocusMap.BasicGraphicsLayer
804:     Set pGCon = pGLayer
805:     Set pElement = New TextElement
806:     pElement.Geometry = pPoint
807:     Set pTextElement = pElement
    
    ' Create the text symbol
810:     Set myTxtSym = New TextSymbol
811:     Set myColor = New RgbColor
812:     myColor.Red = dRed
813:     myColor.Green = dGreen
814:     myColor.Blue = dBlue
815:     myTxtSym.Color = myColor
816:     myTxtSym.Size = dSize
817:     myTxtSym.HorizontalAlignment = esriTHACenter
818:     pTextElement.Symbol = myTxtSym
819:     pTextElement.Text = sText
    
    ' Add the graphic
822:     Set pElementProp = pElement
823:     pElementProp.Name = sElementName
824:     pGCon.AddElement pElement, 0
End Sub

Public Sub RemoveGraphicsByName(pMxDoc As IMxDocument, _
            Optional sPrefix As String = "DEMO_TEMPORARY")
' Delete all graphics with our prefix from ArcScene
    Dim pElement As IElement
    Dim pElementProp As IElementProperties
    Dim sLocalPrefix As String
    Dim pGLayer As IGraphicsLayer
    Dim pGCon As IGraphicsContainer
    Dim lCount As Long
    
    On Error GoTo ErrorHandler
    
    ' Init and switch OFF the updating of the TOC
840:     pMxDoc.DelayUpdateContents = True
841:     Set pGLayer = pMxDoc.FocusMap.BasicGraphicsLayer
842:     Set pGCon = pGLayer
843:     pGCon.Next
    
    ' Delete all the graphic elements that we created (identify by the name prefix)
846:     pGCon.Reset
847:     Set pElement = pGCon.Next
848:     While Not pElement Is Nothing
849:         If TypeOf pElement Is IElement Then
850:             Set pElementProp = pElement
851:             If (Left(pElementProp.Name, Len(sPrefix)) = sPrefix) Then
852:                 pGCon.DeleteElement pElement
853:             End If
854:         End If
855:         Set pElement = pGCon.Next
856:     Wend
    
    ' Switch ON the updating of the TOC, refresh
859:     pMxDoc.DelayUpdateContents = False
860:     pMxDoc.ActiveView.Refresh
    
    Exit Sub
ErrorHandler:
864:     MsgBox "Error in RemoveGraphicsByName: " & Err.Description, , "RemoveGraphicsByName"
End Sub

Private Function IntersectPointExtendedTo(pPolyline As IPolyline, pCirclePoly As IPolygon) As IPoint
    Dim pCurve As ICurve
    Dim pLine As ILine
    Dim pPLine As IPolyline
    Dim pTopoOpt As ITopologicalOperator
    Dim pGeoCol As IGeometryCollection
    ' Need to extend the end (creates an ILine object)
874:     Set pCurve = pPolyline
875:     Set pLine = New esrigeometry.Line
876:     pCurve.QueryTangent esriExtendTangentAtTo, 1, True, _
                        CDbl(m_GridWidth) * 1.1, pLine
    ' Convert ILine to an IPolyline
879:     Set pPLine = New Polyline
880:     pPLine.FromPoint = pLine.FromPoint
881:     pPLine.ToPoint = pLine.ToPoint
    ' Intersect the polyline with the circle
883:     Set pTopoOpt = pPLine
884:     Set pGeoCol = New GeometryBag
885:     Set pGeoCol = pTopoOpt.Intersect(pCirclePoly, esriGeometry0Dimension)
886:     Set IntersectPointExtendedTo = pGeoCol.Geometry(0)
End Function

Private Function AddPathToPolyLine(pPolyline As IPolyline, pPath As IPath) As IPolyline
  Dim pGCol As IGeometryCollection
  Dim pGeom As IGeometry
    
893:   If pPolyline Is Nothing Then
894:     Set pPolyline = New Polyline
895:   End If
896:   Set pGCol = pPolyline
897:   Set pGeom = pPath
898:   pGCol.AddGeometry pGeom
899:   Set AddPathToPolyLine = pGCol

End Function


