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
54:     If pFL.FeatureClass.ShapeType = esriGeometryPolygon Then
55:         Set m_DestFL = pFL
56:         Set m_DestFC = pFL.FeatureClass
57:     Else
58:         Err.Raise vbObjectError, "MapGridManager_Set_DestinationFeatureLayer", _
            "Not a polygon feature layer"
60:     End If
End Property

Public Property Get DestinationFeatureLayer() As IFeatureLayer
64:     Set DestinationFeatureLayer = m_DestFL
End Property

Public Property Set StripMapRoute(RoutePolyline As IPolyline)
68:     Set m_Polyline = RoutePolyline
End Property

Public Property Get StripMapRoute() As IPolyline
72:     Set StripMapRoute = m_Polyline
End Property

Public Property Let FrameWidthInPageUnits(dWidth As Double)
76:     m_dFrameWidthInPageUnits = dWidth
End Property

Public Property Get FrameWidthInPageUnits() As Double
80:     FrameWidthInPageUnits = m_dFrameWidthInPageUnits
End Property

Public Property Let FrameHeightInPageUnits(dHeight As Double)
84:     m_dFrameHeightInPageUnits = dHeight
End Property

Public Property Get FrameHeightInPageUnits() As Double
88:     FrameHeightInPageUnits = m_dFrameHeightInPageUnits
End Property

Public Property Let MapScale(dScale As Double)
92:     m_dMapScale = dScale
End Property

Public Property Get MapScale() As Double
96:     MapScale = m_dMapScale
End Property

Public Property Let StripMapName(MapName As String)
100:     m_StripMapName = MapName
End Property

Public Property Get StripMapName() As String
104:     StripMapName = m_StripMapName
End Property


Public Property Let FieldNameStripMapName(FieldName As String)
109:     m_FldStripName = FieldName
End Property

Public Property Get FieldNameStripMapName() As String
113:     FieldNameStripMapName = m_FldStripName
End Property

Public Property Let FieldNameNumberInSeries(FieldName As String)
117:     m_FldNumInSeries = FieldName
End Property

Public Property Get FieldNameNumberInSeries() As String
121:     FieldNameNumberInSeries = m_FldNumInSeries
End Property

Public Property Let FieldNameMapAngle(FieldName As String)
125:     m_FldMapAngle = FieldName
End Property

Public Property Get FieldNameMapAngle() As String
129:     FieldNameMapAngle = m_FldMapAngle
End Property

Public Property Let FieldNameScale(FieldName As String)
133:     m_FldScale = FieldName
End Property

Public Property Get FieldNameScale() As String
137:     FieldNameScale = m_FldScale
End Property

Public Property Let RemoveCurrentGrids(RemoveGrids As Boolean)
141:     m_RemoveGrids = RemoveGrids
End Property

Public Property Get RemoveCurrentGrids() As Boolean
145:     RemoveCurrentGrids = m_RemoveGrids
End Property

Public Property Let FlipPolyline(Flip As Boolean)
149:     m_Flip = Flip
End Property

Public Property Get FlipPolyline() As Boolean
153:     FlipPolyline = m_Flip
End Property

Private Sub Class_Initialize()
    ' Set the defaults
158:     Set m_pProgress = New ModelessFrame
End Sub

Public Sub RunStandardGUI(pApp As IApplication)
162:     Set frmSMapSettings.m_Application = pApp
163:     frmSMapSettings.Tickle
164:     SetWindowLong frmSMapSettings.hwnd, GWL_HWNDPARENT, pApp.hwnd
165:     frmSMapSettings.Show vbModeless
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
184:     Set pMx = pApp.Document
185:     Set pSR = pMx.FocusMap.SpatialReference
    ' If a Projected coord system
187:     If TypeOf pSR Is IProjectedCoordinateSystem Then
        ' Use meters per unit as the conversion
189:         Set pPCS = pSR
190:         dMetersPerUnit = pPCS.CoordinateUnit.MetersPerUnit
        ' Now convert this into page (ie: paper) units
192:         Set pPage = pMx.PageLayout.Page
193:         pPageUnits = pPage.Units
        Select Case pPageUnits
            Case esriInches: CalculatePageToMapRatio = dMetersPerUnit / (1 / 12 * 0.304800609601219)
            Case esriFeet: CalculatePageToMapRatio = dMetersPerUnit / (0.304800609601219)
            Case esriCentimeters: CalculatePageToMapRatio = dMetersPerUnit / (1 / 100)
            Case esriMeters: CalculatePageToMapRatio = dMetersPerUnit / (1)
            Case Else:
200:                 MsgBox "Warning: Only the following Page (Layout) Units are supported by this tool:" _
                    & vbCrLf & " - Inches, Feet, Centimeters, Meters" _
                    & vbCrLf & vbCrLf & "Calculating as though Page Units are in Inches..."
203:                 CalculatePageToMapRatio = dMetersPerUnit / (1 / 12 * 0.304800609601219)
204:         End Select
    ' Otherwise
206:     Else
        ' If not projected, we can only do a "flat" conversion -> that is, use the current scale and extent
        '  as a ratio to be applied to the map grid scale.
        ' NOTE: We MUST be in Layout mode to make this calculation, as the scale in Map View and Layout View
        '  are not the same (as the extent envelope and data frame envelope can be different shapes).  The
        '  test for being in Layout Mode is made in the clsMapGridButton.ICommand_Enabled property.
212:         Set pExtentEnv = pMx.ActiveView.Extent
213:         dStartX = pExtentEnv.XMin
214:         dStartY = pExtentEnv.YMin
215:         dEndX = pExtentEnv.XMax
216:         dEndY = pExtentEnv.YMax
        
218:         dCurrScale = pMx.FocusMap.MapScale
219:         If ((dEndX - dStartX) / m_dFrameWidthInPageUnits) > ((dEndY - dStartY) / m_dFrameHeightInPageUnits) Then
220:             CalculatePageToMapRatio = m_dFrameWidthInPageUnits / ((dEndX - dStartX) / dCurrScale)
221:         Else
222:             CalculatePageToMapRatio = m_dFrameHeightInPageUnits / ((dEndY - dStartY) / dCurrScale)
223:         End If
224:     End If
    
    Exit Function
eh:
228:     CalculatePageToMapRatio = 1
229:     MsgBox "Error in CalculatePageToMapRatio" & vbCrLf & Err.Description
End Function

Private Sub Class_Terminate()
233:     Set m_DestFL = Nothing
234:     Set m_DestFC = Nothing
235:     Set m_pProgress = New ModelessFrame
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
299:     Screen.MousePointer = vbArrowHourglass

    ' Init
302:     Set pMx = Application.Document
303:     For lLoop = 0 To pMx.FocusMap.LayerCount - 1
304:         If TypeOf pMx.FocusMap.Layer(lLoop) Is IFeatureLayer Then
305:             If UCase(pMx.FocusMap.Layer(lLoop).Name) = UCase(m_DestFL.Name) Then
306:                 Set pFL = pMx.FocusMap.Layer(lLoop)
307:                 Exit For
308:             End If
309:         End If
310:     Next
311:     If pFL Is Nothing Then
312:         MsgBox "No match in the Map for layer '" & m_DestFL.Name & "'."
        Exit Sub
314:     End If
315:     Set pFC = pFL.FeatureClass
    ' Check for required fields - that the field exists
    Dim bErrorWithFields As Boolean
318:     bErrorWithFields = (pFC.FindField(m_FldStripName) < 0)
319:     bErrorWithFields = bErrorWithFields Or (pFC.FindField(m_FldNumInSeries) < 0)
320:     bErrorWithFields = bErrorWithFields Or (pFC.FindField(m_FldMapAngle) < 0)
321:     If Len(m_FldScale) > 0 Then bErrorWithFields = bErrorWithFields Or (pFC.FindField(m_FldScale) < 0)
    ' If error
323:     If bErrorWithFields Then
324:         Err.Raise vbObjectError, "GenerateStripMap", "Could not find all the given field names in " & pFL.Name & "." _
            & vbCrLf & " - " & m_FldStripName & ", " & m_FldNumInSeries & ", " & m_FldMapAngle & ", " & m_FldScale
326:     End If
    ' Check the field types
328:     bErrorWithFields = (pFC.Fields.Field(pFC.FindField(m_FldStripName)).Type <> esriFieldTypeString)
329:     bErrorWithFields = bErrorWithFields Or _
       ((pFC.Fields.Field(pFC.FindField(m_FldNumInSeries)).Type <> esriFieldTypeDouble) And _
        (pFC.Fields.Field(pFC.FindField(m_FldNumInSeries)).Type <> esriFieldTypeInteger) And _
        (pFC.Fields.Field(pFC.FindField(m_FldNumInSeries)).Type <> esriFieldTypeSingle) And _
        (pFC.Fields.Field(pFC.FindField(m_FldNumInSeries)).Type <> esriFieldTypeSmallInteger))
334:     bErrorWithFields = bErrorWithFields Or _
       ((pFC.Fields.Field(pFC.FindField(m_FldMapAngle)).Type <> esriFieldTypeDouble) And _
        (pFC.Fields.Field(pFC.FindField(m_FldMapAngle)).Type <> esriFieldTypeInteger) And _
        (pFC.Fields.Field(pFC.FindField(m_FldMapAngle)).Type <> esriFieldTypeSingle) And _
        (pFC.Fields.Field(pFC.FindField(m_FldMapAngle)).Type <> esriFieldTypeSmallInteger))
339:     If Len(m_FldScale) > 0 Then
340:         bErrorWithFields = bErrorWithFields Or _
           ((pFC.Fields.Field(pFC.FindField(m_FldScale)).Type <> esriFieldTypeDouble) And _
            (pFC.Fields.Field(pFC.FindField(m_FldScale)).Type <> esriFieldTypeInteger) And _
            (pFC.Fields.Field(pFC.FindField(m_FldScale)).Type <> esriFieldTypeSingle) And _
            (pFC.Fields.Field(pFC.FindField(m_FldScale)).Type <> esriFieldTypeSmallInteger))
345:     End If
    ' if error
347:     If bErrorWithFields Then
348:         Err.Raise vbObjectError, "GenerateStripMap", "Given field names are not of the correct type." _
            & vbCrLf & "Strip Map Name field must be a Text field, all others must be numeric fields."
350:     End If
    ' Get the dataset and workspace (to start editing upon)
352:     Set pFeatDataset = pFC.FeatureDataset
353:     If Not pFeatDataset Is Nothing Then
354:         Set pWorkspaceEdit = pFeatDataset.Workspace
355:     Else
        ' Is a shapefile, go via just IDataset
357:         Set pDataset = pFC
358:         Set pWorkspaceEdit = pDataset.Workspace
359:     End If
    
'    ' If replacing, delete all existing polygons
362:     Set pProgress = New frmProgress
363:     m_pProgress.Create pProgress
364:     pProgress.ProgressBar1.Min = 0
365:     pProgress.ProgressBar1.Max = 100
366:     pProgress.ProgressBar1.value = 0
367:     If m_RemoveGrids Then
        Dim pFCu As IFeatureCursor
        Dim pT As ITable
370:         Set pFCu = m_DestFL.Search(Nothing, False)
371:         Set pT = m_DestFL.FeatureClass
372:         If pT.RowCount(Nothing) = 0 Then
373:             dIncrement = 99
374:         Else
375:             dIncrement = 100 / pT.RowCount(Nothing)
376:         End If
377:         pProgress.lblInformation.Caption = "Deleting previous grids..."
378:         pProgress.cmdCancel.Visible = False        ' User cannot cancel this step
379:         m_pProgress.Visible = True
380:         Set pFeature = pFCu.NextFeature
381:         While Not pFeature Is Nothing
382:             pFeature.Delete
383:             If (pProgress.ProgressBar1.value + dIncrement) <= pProgress.ProgressBar1.Max Then
384:                 pProgress.ProgressBar1.value = pProgress.ProgressBar1.value + dIncrement
385:             Else
386:                 pProgress.lblInformation.Caption = "Warning: Val > Max (" & pProgress.ProgressBar1.Max & ")"
387:             End If
388:             Set pFeature = pFCu.NextFeature
389:         Wend
390:         m_pProgress.Visible = False
391:     End If
    
        
    ' Init strip map stuff
395:     Set pPolyline = m_Polyline
    ' Flip, if required
397:     If m_Flip Then
398:         pPolyline.ReverseOrientation
399:     End If
400:     Set pCenterPoint = pPolyline.FromPoint
    ' Get the progress bar ready
402:     pProgress.ProgressBar1.Min = 0
403:     pProgress.ProgressBar1.Max = 101
404:     pProgress.ProgressBar1.value = 0
405:     pProgress.lblInformation.Caption = "Creating strip map..."
406:     pProgress.cmdCancel.Visible = True        ' User cannot cancel this step
    ' Get map units size for grids
408:     dConvertPageToMapUnits = CalculatePageToMapRatio(Application)
409:     m_GridWidth = ((m_dMapScale * m_dFrameWidthInPageUnits) / dConvertPageToMapUnits)
410:     m_GridHeight = ((m_dMapScale * m_dFrameHeightInPageUnits) / dConvertPageToMapUnits)
    ' Init for processing
412:     dHighestPrev = -1
413:     bFirstRun = True
414:     Set pArc = pPolyline
415:     Set pInsertFeatureCursor = pFC.Insert(True)
416:     Set pInsertFeatureBuffer = pFC.CreateFeatureBuffer
417:     m_pProgress.Visible = True
418:     Do
        Dim dCircleRadius As Double, colIntersects As Collection, dIntersect As Double
420:         If bFirstRun Then
421:             dCircleRadius = m_GridWidth / 2
422:         Else
423:             dCircleRadius = m_GridWidth
424:         End If
425:         bReducedRadius = False
426:         Do
            ' Create the search circle
428:             Set pCircularArc = New CircularArc
429:             pCircularArc.ConstructCircle pCenterPoint, dCircleRadius, False 'make it clockwise
430:             Set pCirclePoly = New Polygon
431:             Set pSegmentCollection = pCirclePoly
432:             pSegmentCollection.AddSegment pCircularArc
            
            ' Intersect the polyline and the circle
435:             Set pTopoOpt = pPolyline
436:             Set pGeoCol = New GeometryBag
437:             Set pGeoCol = pTopoOpt.Intersect(pCirclePoly, esriGeometry0Dimension)
            
439:             If pGeoCol.GeometryCount = 0 Then
440:                 MsgBox "error - no geoms intersected"
                Exit Sub
442:             End If
443:             Set pArc = pPolyline
444:             lHighestRef = -1
445:             dHighestThisTurn = 102
446:             For lLoop2 = 0 To pGeoCol.GeometryCount - 1
447:                 Set pIntersectPoint = pGeoCol.Geometry(lLoop2)
448:                 dIntersect = ReturnPercentageAlong(pArc, pIntersectPoint)
449:                 If dIntersect > (dHighestPrev * 1.001) And dIntersect < dHighestThisTurn Then
450:                     dHighest = dIntersect
451:                     dHighestThisTurn = dIntersect
452:                     lHighestRef = lLoop2
453:                 End If
454:             Next
            ' If no intersection higher than our previous run, we are at the end.
456:             If lHighestRef < 0 Then
457:                 dHighest = 101
                ' Need to extend the end (tangent) to get intersection
459:                 Set pIntersectPoint = IntersectPointExtendedTo(pPolyline, pCirclePoly)
460:                 Set pIntersectPointPrev = pCenterPoint
            ' Otherwise, still in the middle somewhere
462:             Else
463:                 Set pIntersectPoint = pGeoCol.Geometry(lHighestRef)
                ' If just starting off (ie: first grid)
465:                 If bFirstRun Then
                    ' Set the grid so the polyline's starting point is in the
                    '  center of the first grid polygon we make
468:                     Set pIntersectPointPrev = New esrigeometry.Point
469:                     pIntersectPointPrev.PutCoords pCenterPoint.X - (pIntersectPoint.X - pCenterPoint.X), _
                                                  pCenterPoint.Y - (pIntersectPoint.Y - pCenterPoint.Y)
                ' Otherwise, we already have a previous point
472:                 Else
                    ' So use it
474:                     Set pIntersectPointPrev = pCenterPoint
475:                 End If
476:             End If
            ' Make our grid polygon, allowing for any 'shrunken' grids
478:             If bReducedRadius Then
                Dim pTmpPLine As IPolyline
                Dim pTmpCPoly As IPolygon
                Dim pTmpIntPoint As IPoint
482:                 Set pCircularArc = New CircularArc
483:                 If bFirstRun Then
484:                     pCircularArc.ConstructCircle pCenterPoint, m_GridWidth / 2, False 'make it clockwise
485:                 Else
486:                     pCircularArc.ConstructCircle pCenterPoint, m_GridWidth, False 'make it clockwise
487:                 End If
488:                 Set pTmpCPoly = New Polygon
489:                 Set pSegmentCollection = pTmpCPoly
490:                 pSegmentCollection.AddSegment pCircularArc
                
492:                 Set pTmpPLine = New Polyline
493:                 pTmpPLine.FromPoint = pIntersectPointPrev
494:                 pTmpPLine.ToPoint = pIntersectPoint
495:                 Set pTmpIntPoint = IntersectPointExtendedTo(pTmpPLine, pTmpCPoly)
496:                 CreateAngledGridPolygon pIntersectPointPrev, pTmpIntPoint, pGridPoly, dGridAngle
497:             Else
498:                 CreateAngledGridPolygon pIntersectPointPrev, pIntersectPoint, pGridPoly, dGridAngle
499:             End If
            ' Now, we potentially need to reprocess if the route dips out of our grid
501:             Set pTopoOpt = pGridPoly
502:             Set pGeoCol = New GeometryBag
503:             Set pGeoCol = pTopoOpt.Intersect(pPolyline, esriGeometry0Dimension)
504:             bContinue = True
505:             If pGeoCol.GeometryCount > 2 Then
506:                 Set colIntersects = New Collection
507:                 For lLoop2 = 0 To pGeoCol.GeometryCount - 1
508:                     colIntersects.Add ReturnPercentageAlong(pArc, pGeoCol.Geometry(lLoop2))
509:                 Next
510:                 For lLoop2 = 1 To colIntersects.count
511:                     If colIntersects.Item(lLoop2) > (dHighestPrev * 1.001) And colIntersects.Item(lLoop2) < (dHighest * 0.999) Then
512:                         bContinue = False
513:                         dHighest = dHighestPrev
514:                         dCircleRadius = dCircleRadius - (m_GridWidth * 0.1)
515:                         bReducedRadius = True
516:                         If dCircleRadius <= 0 Then
517:                             bContinue = True
518:                         End If
519:                         Exit For
520:                     End If
521:                 Next
522:             End If
            ' If all OK and a reduced radius, look for a quick jump ahead
524:             If bContinue And bReducedRadius Then 'And pGeoCol.GeometryCount <= 2 Then
                Dim dTmpHighest As Double
526:                 Set pArc = pPolyline
527:                 lHighestRef = -1
528:                 dTmpHighest = -1
529:                 For lLoop2 = 0 To pGeoCol.GeometryCount - 1
530:                     Set pIntersectPoint = pGeoCol.Geometry(lLoop2)
531:                     dIntersect = ReturnPercentageAlong(pArc, pIntersectPoint)
532:                     If dIntersect > dTmpHighest Then
533:                         dTmpHighest = dIntersect
534:                         lHighestRef = lLoop2
535:                     End If
536:                 Next
537:                 If lHighestRef >= 0 Then Set pIntersectPoint = pGeoCol.Geometry(lHighestRef)
538:                 dHighest = dTmpHighest
539:             End If
540:         Loop Until bContinue
        
542:         bFirstRun = False
543:         dHighestPrev = dHighest
        
        ' All OK to create our grid feature now (hopefully, anyway)
546:         lCounter = lCounter + 1
        'CreateGridFeaturesAsGraphics pGridPoly, lCounter, dGridAngle, Application 'AAA
        
        ' Create new grid cell feature
550:         Set pInsertFeatureBuffer.Shape = pGridPoly
551:         pInsertFeatureBuffer.value(pFC.Fields.FindField(m_FldStripName)) = m_StripMapName & CStr(lCounter)
552:         pInsertFeatureBuffer.value(pFC.Fields.FindField(m_FldNumInSeries)) = lCounter
553:         pInsertFeatureBuffer.value(pFC.Fields.FindField(m_FldMapAngle)) = dGridAngle 'degrees
554:         If Len(m_FldScale) > 0 Then pInsertFeatureBuffer.value(pFC.Fields.FindField(m_FldScale)) = m_dMapScale
555:         pInsertFeatureCursor.InsertFeature pInsertFeatureBuffer
556:         If dHighest <= pProgress.ProgressBar1.Max Then
557:             pProgress.ProgressBar1.value = dHighest
558:         Else
559:             pProgress.lblInformation.Caption = "Warning: Val > Max (" & pProgress.ProgressBar1.Max & ")"
560:             pProgress.ProgressBar1.value = pProgress.ProgressBar1.Max
561:         End If
562:         If (lCounter Mod 20 = 0) Then
563:             DoEvents
564:             pInsertFeatureCursor.Flush
565:         End If
566:         pProgress.Refresh
567:         If pProgress.Cancelled Then
            Dim vUserChoice
569:             pProgress.Cancelled = False       ' Reset the form
570:             vUserChoice = MsgBox("Operation cancelled." _
                & "  Save the edits made thus far?" & vbCrLf & vbCrLf _
                & "(Click Cancel to continue processing)", _
                            vbYesNoCancel, "Generate Strip Map")
574:             If vUserChoice <> vbCancel Then
575:                 GoTo CancelledGenerateGrids     'Sorry for GoTo usage - in a hurry
576:             End If
577:         End If
        ' For next time
579:         Set pCenterPoint = pIntersectPoint
580:     Loop While dHighest < 100
    ' Add remainder polys
582:     pInsertFeatureCursor.Flush
583:     m_pProgress.Visible = False

    ' Stop editing
586:     pWorkspaceEdit.StopEditOperation
587:     pWorkspaceEdit.StopEditing True
588:     pMx.ActiveView.Refresh
    
    Exit Sub
    
CancelledGenerateGrids:
593:     m_pProgress.Visible = False
594:     If vUserChoice = vbYes Then
595:         pInsertFeatureCursor.Flush
596:         pWorkspaceEdit.StopEditOperation
597:         pWorkspaceEdit.StopEditing True
598:     Else
599:         pWorkspaceEdit.StopEditOperation
600:         pWorkspaceEdit.StopEditing False
601:     End If
602:     Screen.MousePointer = vbDefault
603:     pMx.ActiveView.Refresh
    Exit Sub
    
606:     Resume
eh:
608:     MsgBox "Error in GenerateStripMap:" & vbCrLf & Err.Description
End Sub

Private Sub CreateGridFeaturesAsGraphics(pGridPolygon As IPolygon, lIndex As Long, dAngle As Double, pApp As IApplication)
    Dim pPntColl As IPointCollection
    Dim pArea As IArea
    Dim pCentroid As IPoint
    Dim lLoop As Long
    
    ' Create graphics (TEST PHASE)
618:     Set pPntColl = pGridPolygon
619:     For lLoop = 0 To pPntColl.PointCount - 2
620:         Perm_DrawLineFromPoints pPntColl.Point(lLoop), pPntColl.Point(lLoop + 1), pApp
621:     Next
622:     Perm_DrawLineFromPoints pPntColl.Point(0), pPntColl.Point(pPntColl.PointCount - 1), pApp
623:     Set pArea = pGridPolygon
624:     Set pCentroid = pArea.Centroid
625:     Perm_DrawTextFromPoint pCentroid, CStr(lIndex), pApp, , , , , 8
626:     pCentroid.Y = pCentroid.Y - (m_GridWidth / 3)
627:     Perm_DrawTextFromPoint pCentroid, Format(dAngle / cPI * 180, "(#0.0)"), pApp, , , , , 8

End Sub

Private Function ReturnPercentageAlong(ByVal pArc As ICurve, ByVal pPoint As IPoint) As Double
    Dim GeoCount As Long
    Dim pDistAlong As Double
    Dim pDist As Double
    Dim pRightSide As Boolean
    Dim pOutPt As IPoint
    Dim CompareDist As Double
    
    On Error GoTo ErrorHandler
    
641:     CompareDist = 0
    'Find the distance along curve
643:     Set pOutPt = New esrigeometry.Point
644:     pArc.QueryPointAndDistance esriNoExtension, pPoint, True, pOutPt, _
                        pDistAlong, pDist, pRightSide
    
647:     ReturnPercentageAlong = (pDistAlong * 100)
    Exit Function
ErrorHandler:
650:     Err.Raise Err.Number, "ReturnPercentageAlong", "Error in ReturnPercentageAlong." _
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
665:     Set pLine = New esrigeometry.Line
666:     pLine.FromPoint = p1
667:     pLine.ToPoint = p2
668:     dAngleInRadians = pLine.Angle
669:     If dAngleInRadians = 0 Then
670:       ReturnedAngleRadians = 0
671:     ElseIf dAngleInRadians > 0 Then
672:       ReturnedAngleRadians = 360 - ((dAngleInRadians / cPI) * 180)
673:     Else
674:       ReturnedAngleRadians = Abs((dAngleInRadians / cPI) * 180)
675:     End If
676:     Set ReturnedGrid = New Polygon
677:     Set pPointColl = ReturnedGrid
    ' POINT 1 -------------------------------------------
679:     Set pPoint = New esrigeometry.Point
680:     pPoint.PutCoords p1.X + (Sin(dAngleInRadians) * (m_GridHeight / 2)), _
                     p1.Y - (Cos(dAngleInRadians) * (m_GridHeight / 2))
682:     pPointColl.AddPoint pPoint
683:     Set pPointStart = pPoint
    ' POINT 2 -------------------------------------------
685:     Set pPoint = New esrigeometry.Point
686:     pPoint.PutCoords p1.X - (Sin(dAngleInRadians) * (m_GridHeight / 2)), _
                     p1.Y + (Cos(dAngleInRadians) * (m_GridHeight / 2))
688:     pPointColl.AddPoint pPoint
    ' POINT 3 -------------------------------------------
690:     Set pPoint = New esrigeometry.Point
691:     pPoint.PutCoords p2.X - Sin(dAngleInRadians) * m_GridHeight / 2, _
                     p2.Y + Cos(dAngleInRadians) * m_GridHeight / 2
693:     pPointColl.AddPoint pPoint
    ' POINT 4 -------------------------------------------
695:     Set pPoint = New esrigeometry.Point
696:     pPoint.PutCoords p2.X + Sin(dAngleInRadians) * m_GridHeight / 2, _
                     p2.Y - Cos(dAngleInRadians) * m_GridHeight / 2
698:     pPointColl.AddPoint pPoint
    ' JOIN BACK TO FIRST (CLOSE POLYGON) ----------------
700:     pPointColl.AddPoint pPointStart
    
    Exit Sub
eh:
704:     Err.Raise Err.Number, Err.source, "Error in CreateAngledGridPolygon." _
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
723:     Set pMx = Application.Document
724:     Set pGLayer = pMx.FocusMap.BasicGraphicsLayer
725:     Set pGCon = pGLayer
726:     Set pElement = New MarkerElement
727:     pElement.Geometry = pPoint
728:     Set pMarkerElement = pElement
    
    ' Set the symbol
731:     Set pColor = New RgbColor
732:     pColor.Red = dRed
733:     pColor.Green = dGreen
734:     pColor.Blue = dBlue
735:     Set pMarker = New SimpleMarkerSymbol
736:     With pMarker
737:         .Color = pColor
738:         .Size = dSize
739:     End With
740:     pMarkerElement.Symbol = pMarker
    
    ' Add the graphic
743:     Set pElementProp = pElement
744:     pElementProp.Name = sElementName
745:     pGCon.AddElement pElement, 0
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
769:     Set pMx = Application.Document
770:     Set pGLayer = pMx.FocusMap.BasicGraphicsLayer
771:     Set pGCon = pGLayer
772:     Set pElement = New LineElement
    
    ' Set the line symbol
775:     Set pLnSym = New SimpleLineSymbol
776:     Set myColor = New RgbColor
777:     myColor.Red = dRed
778:     myColor.Green = dGreen
779:     myColor.Blue = dBlue
780:     pLnSym.Color = myColor
781:     pLnSym.Width = dSize
    
    ' Create a standard polyline (via 2 points)
784:     Set pLine1 = New esrigeometry.Line
785:     pLine1.PutCoords pFromPoint, pToPoint
786:     Set pSeg1 = pLine1
787:     Set pPolyline = New Polyline
788:     pPolyline.AddSegment pSeg1
789:     pElement.Geometry = pPolyline
790:     Set pLineElement = pElement
791:     pLineElement.Symbol = pLnSym
    
    ' Add the graphic
794:     Set pElementProp = pElement
795:     pElementProp.Name = sElementName
796:     pGCon.AddElement pElement, 0
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
814:     Set pMx = Application.Document
815:     Set pGLayer = pMx.FocusMap.BasicGraphicsLayer
816:     Set pGCon = pGLayer
817:     Set pElement = New TextElement
818:     pElement.Geometry = pPoint
819:     Set pTextElement = pElement
    
    ' Create the text symbol
822:     Set myTxtSym = New TextSymbol
823:     Set myColor = New RgbColor
824:     myColor.Red = dRed
825:     myColor.Green = dGreen
826:     myColor.Blue = dBlue
827:     myTxtSym.Color = myColor
828:     myTxtSym.Size = dSize
829:     myTxtSym.HorizontalAlignment = esriTHACenter
830:     pTextElement.Symbol = myTxtSym
831:     pTextElement.Text = sText
    
    ' Add the graphic
834:     Set pElementProp = pElement
835:     pElementProp.Name = sElementName
836:     pGCon.AddElement pElement, 0
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
852:     pMxDoc.DelayUpdateContents = True
853:     Set pGLayer = pMxDoc.FocusMap.BasicGraphicsLayer
854:     Set pGCon = pGLayer
855:     pGCon.Next
    
    ' Delete all the graphic elements that we created (identify by the name prefix)
858:     pGCon.Reset
859:     Set pElement = pGCon.Next
860:     While Not pElement Is Nothing
861:         If TypeOf pElement Is IElement Then
862:             Set pElementProp = pElement
863:             If (Left(pElementProp.Name, Len(sPrefix)) = sPrefix) Then
864:                 pGCon.DeleteElement pElement
865:             End If
866:         End If
867:         Set pElement = pGCon.Next
868:     Wend
    
    ' Switch ON the updating of the TOC, refresh
871:     pMxDoc.DelayUpdateContents = False
872:     pMxDoc.ActiveView.Refresh
    
    Exit Sub
ErrorHandler:
876:     MsgBox "Error in RemoveGraphicsByName: " & Err.Description, , "RemoveGraphicsByName"
End Sub

Private Function IntersectPointExtendedTo(pPolyline As IPolyline, pCirclePoly As IPolygon) As IPoint
    Dim pCurve As ICurve
    Dim pLine As ILine
    Dim pPLine As IPolyline
    Dim pTopoOpt As ITopologicalOperator
    Dim pGeoCol As IGeometryCollection
    ' Need to extend the end (creates an ILine object)
886:     Set pCurve = pPolyline
887:     Set pLine = New esrigeometry.Line
888:     pCurve.QueryTangent esriExtendTangentAtTo, 1, True, _
                        CDbl(m_GridWidth) * 1.1, pLine
    ' Convert ILine to an IPolyline
891:     Set pPLine = New Polyline
892:     pPLine.FromPoint = pLine.FromPoint
893:     pPLine.ToPoint = pLine.ToPoint
    ' Intersect the polyline with the circle
895:     Set pTopoOpt = pPLine
896:     Set pGeoCol = New GeometryBag
897:     Set pGeoCol = pTopoOpt.Intersect(pCirclePoly, esriGeometry0Dimension)
898:     Set IntersectPointExtendedTo = pGeoCol.Geometry(0)
End Function

Private Function AddPathToPolyLine(pPolyline As IPolyline, pPath As IPath) As IPolyline
  Dim pGCol As IGeometryCollection
  Dim pGeom As IGeometry
    
905:   If pPolyline Is Nothing Then
906:     Set pPolyline = New Polyline
907:   End If
908:   Set pGCol = pPolyline
909:   Set pGeom = pPath
910:   pGCol.AddGeometry pGeom
911:   Set AddPathToPolyLine = pGCol

End Function


