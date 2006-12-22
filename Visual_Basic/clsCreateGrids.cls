VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCreateGrids"
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
' - IdentifierOrder (mgGridIdentifierOrderType, r/w)
' - RowIDType (mgGridIdentifierValueType, r/w)
' - ColIDType (mgGridIdentifierValueType, r/w)
' - StartingCoordinateLL_X (Double, r/w)
' - StartingCoordinateLL_Y (Double, r/w)
' - EndingCoordinateUR_X (Double, r/w)
' - EndingCoordinateUR_Y (Double, r/w)
' - UseUnderscore (Boolean, r/w)
'----------------------------------------------
' Methods
' - GenerateGrids(pApp as IApplication)
'    : Generates the grids using the values added.
' - RunStandardGUI(pApp as IApplication)
'    : To open the form without having the button
'      added to ArcMap's GUI.
'----------------------------------------------
' Enumerated Types
Public Enum mgGridIdentifierOrderType
    Row_Column = 0
    Column_Row = 1
End Enum
Public Enum mgGridIdentifierValueType
    Alphabetical = 0
    Numerical = 1
End Enum
Public Enum mgGridIdentifierStartPositionType
    TopLeft = 0
    LowerLeft = 1
End Enum
'----------------------------------------------
' Local Global Declarations
Private m_DestFL As IFeatureLayer
Private m_DestFC As IFeatureClass
Private m_dMapScale As Double
Private m_dFrameWidthInPageUnits As Double
Private m_dFrameHeightInPageUnits As Double
Private m_IDOrderType As mgGridIdentifierOrderType
Private m_RowIDType As mgGridIdentifierValueType
Private m_ColIDType As mgGridIdentifierValueType
Private m_StartIDType As mgGridIdentifierStartPositionType
Private m_StartX As Double
Private m_StartY As Double
Private m_EndX As Double
Private m_EndY As Double
Private m_UseUnderscore As Boolean
Private m_colLayerNamesForData As Collection
Private m_FldID As String
Private m_FldRowNum As String
Private m_FldColNum As String
Private m_FldScale As String
Private m_RemoveGrids As Boolean
Private m_NoEmptyGrids As Boolean
Private m_pProgress As IModelessFrame

'----------------------------------------------
' API call to keep form top most
Private Const GWL_HWNDPARENT = -8
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
  (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Property Set DestinationFeatureLayer(pFL As IFeatureLayer)
70:     If pFL.FeatureClass.ShapeType = esriGeometryPolygon Then
71:         Set m_DestFL = pFL
72:         Set m_DestFC = pFL.FeatureClass
73:     Else
74:         Err.Raise vbObjectError, "MapGridManager_Set_DestinationFeatureLayer", _
            "Not a polygon feature layer"
76:     End If
End Property

Public Property Get DestinationFeatureLayer() As IFeatureLayer
80:     Set DestinationFeatureLayer = m_DestFL
End Property

Public Property Let FrameWidthInPageUnits(dWidth As Double)
84:     m_dFrameWidthInPageUnits = dWidth
End Property

Public Property Get FrameWidthInPageUnits() As Double
88:     FrameWidthInPageUnits = m_dFrameWidthInPageUnits
End Property

Public Property Let FrameHeightInPageUnits(dHeight As Double)
92:     m_dFrameHeightInPageUnits = dHeight
End Property

Public Property Get FrameHeightInPageUnits() As Double
96:     FrameHeightInPageUnits = m_dFrameHeightInPageUnits
End Property

Public Property Let MapScale(dScale As Double)
100:     m_dMapScale = dScale
End Property

Public Property Get MapScale() As Double
104:     MapScale = m_dMapScale
End Property

Public Property Let IdentifierOrder(mgType As mgGridIdentifierOrderType)
108:     m_IDOrderType = mgType
End Property

Public Property Get IdentifierOrder() As mgGridIdentifierOrderType
112:     IdentifierOrder = m_IDOrderType
End Property

Public Property Let RowIDType(mgIDStyle As mgGridIdentifierValueType)
116:     m_RowIDType = mgIDStyle
End Property

Public Property Get RowIDType() As mgGridIdentifierValueType
120:     RowIDType = m_RowIDType
End Property

Public Property Let ColIDType(mgIDStyle As mgGridIdentifierValueType)
124:     m_ColIDType = mgIDStyle
End Property

Public Property Get ColIDType() As mgGridIdentifierValueType
128:     ColIDType = m_ColIDType
End Property

Public Property Let IDStartPositionType(mgStartPos As mgGridIdentifierStartPositionType)
132:     m_StartIDType = mgStartPos
End Property

Public Property Get IDStartPositionType() As mgGridIdentifierStartPositionType
136:     IDStartPositionType = m_StartIDType
End Property

Public Property Let StartingCoordinateLL_X(X As Double)
140:     m_StartX = X
End Property

Public Property Get StartingCoordinateLL_X() As Double
144:     StartingCoordinateLL_X = m_StartX
End Property

Public Property Let StartingCoordinateLL_Y(Y As Double)
148:     m_StartY = Y
End Property

Public Property Get StartingCoordinateLL_Y() As Double
152:     StartingCoordinateLL_Y = m_StartY
End Property

Public Property Let EndingCoordinateUR_X(X As Double)
156:     m_EndX = X
End Property

Public Property Get EndingCoordinateUR_X() As Double
160:     EndingCoordinateUR_X = m_EndX
End Property

Public Property Let EndingCoordinateUR_Y(Y As Double)
164:     m_EndY = Y
End Property

Public Property Get EndingCoordinateUR_Y() As Double
168:     EndingCoordinateUR_Y = m_EndY
End Property

Public Property Let FieldNameGridID(FieldName As String)
172:     m_FldID = FieldName
End Property

Public Property Let FieldNameRowNum(FieldName As String)
176:     m_FldRowNum = FieldName
End Property

Public Property Let FieldNameColNum(FieldName As String)
180:     m_FldColNum = FieldName
End Property

Public Property Let FieldNameScale(FieldName As String)
184:     m_FldScale = FieldName
End Property

Public Property Let UseUnderscore(UseUnderscoreInID As Boolean)
188:     m_UseUnderscore = UseUnderscoreInID
End Property

Public Property Get UseUnderscore() As Boolean
192:     UseUnderscore = m_UseUnderscore
End Property

Public Property Let RemoveCurrentGrids(RemoveGrids As Boolean)
196:     m_RemoveGrids = RemoveGrids
End Property

Public Property Get RemoveCurrentGrids() As Boolean
200:     RemoveCurrentGrids = m_RemoveGrids
End Property

Public Property Let NoEmptyGrids(NoEmptyGridPolygons As Boolean)
204:     m_NoEmptyGrids = NoEmptyGridPolygons
End Property

Public Property Get NoEmptyGrids() As Boolean
208:     NoEmptyGrids = m_NoEmptyGrids
End Property

Public Sub AddRequiredDataLayer(sLayerName As String)
212:     m_colLayerNamesForData.Add sLayerName
End Sub

Public Property Get RequiredDataLayer(LayerNumber As Long) As String
216:     If LayerNumber <= m_colLayerNamesForData.count Then
217:         RequiredDataLayer = m_colLayerNamesForData.Item(LayerNumber - 1)
218:     Else
219:         If m_colLayerNamesForData.count = 0 Then
220:             RequiredDataLayer = "There are no Required Data Layers."
221:         Else
222:             RequiredDataLayer = "Invalid layer number.  Valid range: 0 - " & (m_colLayerNamesForData.count - 1)
223:         End If
224:     End If
End Property

Public Property Get RequiredDataLayerCount() As Long
228:     RequiredDataLayerCount = m_colLayerNamesForData.count
End Property

Public Sub ClearRequiredDataLayers()
232:     Set m_colLayerNamesForData = New Collection
End Sub

Private Sub Class_Initialize()
    ' Set the default grid identifier (row-col, alpha/numeric => eg: "B02")
237:     m_IDOrderType = Row_Column
238:     m_RowIDType = Alphabetical
239:     m_ColIDType = Numerical
240:     Set m_colLayerNamesForData = New Collection
241:     Set m_pProgress = New ModelessFrame
End Sub

Private Function CalculateID(ByVal lRow As Long, ByVal lCol As Long, _
                             ByVal iRowIDLen As Integer, ByVal iColIDLen As Integer) As String
    Dim sRowID As String
    Dim sColID As String
    Dim sNumericFormat As String
    Dim lLoop As Long
    Dim lTmp As Long, lCalc As Long, lTmp2 As Long

    ' Row ---------------------------------------------
253:     sNumericFormat = ""
254:     For lLoop = 1 To iRowIDLen
255:         sNumericFormat = sNumericFormat & "0"
256:     Next
257:     If m_RowIDType = Alphabetical Then
258:         For lLoop = 1 To iRowIDLen
259:             lTmp = 26 ^ lLoop
260:             lTmp2 = (26 ^ (lLoop - 1))
261:             If lRow >= lTmp2 Then
262:                 lCalc = ((((lRow - lTmp2) / lTmp) * 26) + 1) Mod 26
263:                 sRowID = Chr(Asc("A") + lCalc) & sRowID
264:                 lRow = lRow - (lCalc * lTmp2)
265:             Else
266:                 sRowID = "A" & sRowID
267:             End If
268:         Next
269:     Else
270:         sRowID = Format(lRow + 1, sNumericFormat)
271:     End If
    ' Col ---------------------------------------------
273:     sNumericFormat = ""
274:     For lLoop = 1 To iColIDLen
275:         sNumericFormat = sNumericFormat & "0"
276:     Next
277:     If m_ColIDType = Alphabetical Then
278:         For lLoop = 1 To iColIDLen
279:             lTmp = 26 ^ lLoop
280:             lTmp2 = (26 ^ (lLoop - 1))
281:             If lCol >= lTmp2 Then
282:                 lCalc = ((((lCol - lTmp2) / lTmp) * 26) + 1) Mod 26
283:                 sColID = Chr(Asc("A") + lCalc) & sColID
284:                 lCol = lCol - (lCalc * lTmp2)
285:             Else
286:                 sColID = "A" & sColID
287:             End If
288:         Next
289:     Else
290:         sColID = Format(lCol + 1, sNumericFormat)
291:     End If
    ' Join --------------------------------------------
293:     If m_IDOrderType = Row_Column Then
294:         If m_UseUnderscore Then
295:             CalculateID = sRowID & "_" & sColID
296:         Else
297:             CalculateID = sRowID & sColID
298:         End If
299:     Else
300:         If m_UseUnderscore Then
301:             CalculateID = sColID & "_" & sRowID
302:         Else
303:             CalculateID = sColID & sRowID
304:         End If
305:     End If
End Function

Private Function GetMinimumStringLength(lValue As Long, lBase As Long) As Integer
    Dim lTmp As Long, lIndex As Long
    
    On Error GoTo eh
    ' ROW
313:     lTmp = lBase
314:     lIndex = 1
315:     While lValue > (lTmp - 1)
316:         lTmp = lTmp * lBase
317:         lIndex = lIndex + 1
318:     Wend
319:     GetMinimumStringLength = lIndex
    Exit Function
eh:
322:     Err.Raise Err.Number, "GetMinimumStringLength", "Error in GetMinimumStringLength: " & Err.Description
End Function

Public Sub RunStandardGUI(pApp As IApplication)
326:     Set frmGridSettings.m_Application = pApp
327:     frmGridSettings.Tickle
328:     SetWindowLong frmGridSettings.hwnd, GWL_HWNDPARENT, pApp.hwnd
329:     frmGridSettings.Show vbModeless
End Sub

'Private Function CreateGridPoly(pStartPoint As IPoint, lRow As Long, lCol As Long, _
'                                dGridWidth As Double, dGridHeight As Double) As IPolygon
'    Dim pPntColl As IPointCollection
'    Dim pPoint As IPoint
'    Dim dX As Double, dY As Double
'
'    Set CreateGridPoly = New Polygon
'    Set pPntColl = CreateGridPoly
'    Set pPoint = new esrigeometry.point
'    pPoint.PutCoords pStartPoint.X + (lCol * dGridWidth), pStartPoint.Y + (lRow * dGridHeight)
'    pPntColl.AddPoint pPoint
'    Set pPoint = new esrigeometry.point
'    pPoint.PutCoords pStartPoint.X + ((lCol + 1) * dGridWidth), pStartPoint.Y + (lRow * dGridHeight)
'    pPntColl.AddPoint pPoint
'    Set pPoint = new esrigeometry.point
'    pPoint.PutCoords pStartPoint.X + ((lCol + 1) * dGridWidth), pStartPoint.Y + ((lRow + 1) * dGridHeight)
'    pPntColl.AddPoint pPoint
'    Set pPoint = new esrigeometry.point
'    pPoint.PutCoords pStartPoint.X + (lCol * dGridWidth), pStartPoint.Y + ((lRow + 1) * dGridHeight)
'    pPntColl.AddPoint pPoint
'    Set pPoint = new esrigeometry.point
'    pPoint.PutCoords pStartPoint.X + (lCol * dGridWidth), pStartPoint.Y + (lRow * dGridHeight)
'    pPntColl.AddPoint pPoint
'
'    'Set CreateGridPoly = pPntColl
'    Debug.Print CreateGridPoly.IsClosed
'    'Debug.Print CreateGridPoly.Envelope.XMin & "," & CreateGridPoly.Envelope.YMin
'
'End Function

Public Sub GenerateGrids2(Application As IApplication) ', _
                         'Optional bRemoveEmptyGrids As Boolean = False, _
                         'Optional bReplaceExistingGrids As Boolean = False)
    Dim pEditor As IEditor
    Dim pUID As New UID
    Dim pWorkspaceEdit As IWorkspaceEdit
    Dim lLoop As Long
    Dim pFeatDataset As IFeatureDataset
    Dim pFeature As IFeature
    Dim pFeatCur As IFeatureCursor
    Dim pSourcePolygon As IPolygon
    Dim pGridPolygon As IPolygon
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
    Dim pMx As IMxDocument
    Dim pFL As IFeatureLayer
    Dim pFC As IFeatureClass
    Dim pProgress As frmProgress
    
    On Error GoTo eh
    
    ' Set mouse pointer
404:     Screen.MousePointer = vbArrowHourglass
    
    ' Init
407:     Set pMx = Application.Document
408:     For lLoop = 0 To pMx.FocusMap.LayerCount - 1
409:         If TypeOf pMx.FocusMap.Layer(lLoop) Is IFeatureLayer Then
410:             If UCase(pMx.FocusMap.Layer(lLoop).Name) = UCase(m_DestFL.Name) Then
411:                 Set pFL = pMx.FocusMap.Layer(lLoop)
412:                 Exit For
413:             End If
414:         End If
415:     Next
416:     If pFL Is Nothing Then
417:         MsgBox "No match in the Map for layer '" & m_DestFL.Name & "'."
        Exit Sub
419:     End If
420:     Set pFC = pFL.FeatureClass
    ' Check for required fields - that the field exists
    Dim bErrorWithFields As Boolean
423:     bErrorWithFields = (pFC.FindField(m_FldID) < 0)
424:     If Len(m_FldRowNum) > 0 Then bErrorWithFields = bErrorWithFields Or (pFC.FindField(m_FldRowNum) < 0)
425:     If Len(m_FldColNum) > 0 Then bErrorWithFields = bErrorWithFields Or (pFC.FindField(m_FldColNum) < 0)
426:     If Len(m_FldScale) > 0 Then bErrorWithFields = bErrorWithFields Or (pFC.FindField(m_FldScale) < 0)
    ' If error
428:     If bErrorWithFields Then
429:         Err.Raise vbObjectError, "GenerateGrids", "Could not find all the given field names in " & pFL.Name & "." _
            & vbCrLf & " - " & m_FldID & ", " & m_FldRowNum & ", " & m_FldColNum & ", " & m_FldScale
431:     End If
    ' Check the field types
433:     bErrorWithFields = (pFC.Fields.Field(pFC.FindField(m_FldID)).Type <> esriFieldTypeString)
434:     If Len(m_FldRowNum) > 0 Then
435:         bErrorWithFields = bErrorWithFields Or _
           ((pFC.Fields.Field(pFC.FindField(m_FldRowNum)).Type <> esriFieldTypeDouble) And _
            (pFC.Fields.Field(pFC.FindField(m_FldRowNum)).Type <> esriFieldTypeInteger) And _
            (pFC.Fields.Field(pFC.FindField(m_FldRowNum)).Type <> esriFieldTypeSingle) And _
            (pFC.Fields.Field(pFC.FindField(m_FldRowNum)).Type <> esriFieldTypeSmallInteger))
440:     End If
441:     If Len(m_FldColNum) > 0 Then
442:         bErrorWithFields = bErrorWithFields Or _
           ((pFC.Fields.Field(pFC.FindField(m_FldColNum)).Type <> esriFieldTypeDouble) And _
            (pFC.Fields.Field(pFC.FindField(m_FldColNum)).Type <> esriFieldTypeInteger) And _
            (pFC.Fields.Field(pFC.FindField(m_FldColNum)).Type <> esriFieldTypeSingle) And _
            (pFC.Fields.Field(pFC.FindField(m_FldColNum)).Type <> esriFieldTypeSmallInteger))
447:     End If
448:     If Len(m_FldScale) > 0 Then
449:         bErrorWithFields = bErrorWithFields Or _
           ((pFC.Fields.Field(pFC.FindField(m_FldScale)).Type <> esriFieldTypeDouble) And _
            (pFC.Fields.Field(pFC.FindField(m_FldScale)).Type <> esriFieldTypeInteger) And _
            (pFC.Fields.Field(pFC.FindField(m_FldScale)).Type <> esriFieldTypeSingle) And _
            (pFC.Fields.Field(pFC.FindField(m_FldScale)).Type <> esriFieldTypeSmallInteger))
454:     End If
    ' if error
456:     If bErrorWithFields Then
457:         Err.Raise vbObjectError, "GenerateGrids", "Given field names are not of the correct type." _
            & vbCrLf & "Grid ID field must be a Text field, all others must be numeric fields."
459:     End If
    ' Get the dataset and workspace (to start editing upon)
461:     Set pFeatDataset = pFC.FeatureDataset
462:     If Not pFeatDataset Is Nothing Then
463:         Set pWorkspaceEdit = pFeatDataset.Workspace
464:     Else
        ' Is a shapefile, go via just IDataset
466:         Set pDataset = pFC
467:         Set pWorkspaceEdit = pDataset.Workspace
468:     End If
469:     dDataFrameWidth = m_dFrameWidthInPageUnits
470:     dDataFrameHeight = m_dFrameHeightInPageUnits
    ' Start Editing
472:     pWorkspaceEdit.StartEditing False
473:     pWorkspaceEdit.StartEditOperation
    
    ' If replacing, delete all existing polygons
476:     Set pProgress = New frmProgress
477:     m_pProgress.Create pProgress
478:     If m_RemoveGrids Then
        Dim pFCu As IFeatureCursor
        Dim pT As ITable
481:         Set pFCu = m_DestFL.Search(Nothing, False)
482:         Set pT = m_DestFL.FeatureClass
483:         pProgress.ProgressBar1.Min = 0
484:         pProgress.ProgressBar1.Max = 100
485:         If pT.RowCount(Nothing) = 0 Then
486:             dIncrement = 99
487:         Else
488:             dIncrement = 100 / pT.RowCount(Nothing)
489:         End If
490:         pProgress.ProgressBar1.Value = 0
491:         pProgress.lblInformation.Caption = "Deleting previous grids..."
492:         pProgress.cmdCancel.Visible = False        ' User cannot cancel this step
        
494:         m_pProgress.Visible = True
495:         Set pFeature = pFCu.NextFeature
496:         While Not pFeature Is Nothing
497:             pFeature.Delete
498:             If (pProgress.ProgressBar1.Value + dIncrement) <= pProgress.ProgressBar1.Max Then
499:                 pProgress.ProgressBar1.Value = pProgress.ProgressBar1.Value + dIncrement
500:             Else
501:                 pProgress.lblInformation.Caption = "Warning: Val > Max (" & pProgress.ProgressBar1.Max & ")"
502:             End If
503:             Set pFeature = pFCu.NextFeature
504:         Wend
505:         m_pProgress.Visible = False
506:     End If
    
    ' Calc the row/column extents, grid size (map units), ID lengths and starting coordinate
509:     Set pStartingCoord = New esrigeometry.Point
510:     pStartingCoord.PutCoords m_StartX, m_StartY
511:     dConvertPageToMapUnits = CalculatePageToMapRatio(Application)
512:     dGridSizeW = ((m_dMapScale * dDataFrameWidth) / dConvertPageToMapUnits)
513:     dGridSizeH = ((m_dMapScale * dDataFrameHeight) / dConvertPageToMapUnits)
514:     If Not (pFL.FeatureClass.FeatureDataset Is Nothing) Then
515:         CalculateRowColCounts m_StartX, m_StartY, m_EndX, m_EndY, _
            dGridSizeW, dGridSizeH, GetValidExtentForLayer(m_DestFL), lColCount, lRowCount
517:     Else
518:         CalculateRowColCounts m_StartX, m_StartY, m_EndX, m_EndY, _
            dGridSizeW, dGridSizeH, Nothing, lColCount, lRowCount
520:     End If
521:     If lColCount = 0 Or lRowCount = 0 Then
522:         Err.Raise vbObjectError, "GenerateGrids", "CalculateRowColCounts returned zero row/columns"
523:     End If
524:     If m_ColIDType = Alphabetical Then
525:         lBase = 26
526:     Else
527:         lBase = 10
528:     End If
529:     iStringLengthCol = GetMinimumStringLength(lColCount, lBase)
530:     If m_RowIDType = Alphabetical Then
531:         lBase = 26
532:     Else
533:         lBase = 10
534:     End If
535:     iStringLengthRow = GetMinimumStringLength(lRowCount, lBase)
    
    ' Create and add the new Grid polygon features
538:     bOKToAdd = True
539:     pProgress.ProgressBar1.Min = 0
540:     pProgress.ProgressBar1.Max = 100
541:     pProgress.ProgressBar1.Value = 0
542:     pProgress.lblInformation.Caption = "Creating new Grid features..."
543:     pProgress.cmdCancel.Visible = True
544:     m_pProgress.Visible = True
545:     dIncrement = 99.9 / (lRowCount * lColCount)
546:     Set pInsertFeatureCursor = pFC.Insert(True)
547:     Set pInsertFeatureBuffer = pFC.CreateFeatureBuffer
548:     For lRow = 0 To lRowCount - 1
549:         For lCol = 0 To lColCount - 1
            ' Create the source polygon
551:             Set pGridPolygon = CreateGridPoly2(pStartingCoord, lRow, lCol, dGridSizeW, dGridSizeH)
            ' If required, check for containing features
553:             If m_NoEmptyGrids Then
554:                 bOKToAdd = HasFeatures(pGridPolygon, pMx.FocusMap)
555:             End If
556:             If bOKToAdd Then
                ' Create new grid cell feature
558:                 Set pInsertFeatureBuffer.Shape = pGridPolygon
559:                 If m_StartIDType = TopLeft Then
560:                     pInsertFeatureBuffer.Value(pFC.Fields.FindField(m_FldID)) = _
                            CalculateID((lRowCount - 1) - lRow, lCol, iStringLengthRow, iStringLengthCol)
562:                     If Len(m_FldRowNum) > 0 Then pInsertFeatureBuffer.Value(pFC.Fields.FindField(m_FldRowNum)) = (lRowCount - lRow)
563:                 Else
564:                     pInsertFeatureBuffer.Value(pFC.Fields.FindField(m_FldID)) = _
                            CalculateID(lRow, lCol, iStringLengthRow, iStringLengthCol)
566:                     If Len(m_FldRowNum) > 0 Then pInsertFeatureBuffer.Value(pFC.Fields.FindField(m_FldRowNum)) = (lRow + 1)
567:                 End If
568:                 If Len(m_FldColNum) > 0 Then pInsertFeatureBuffer.Value(pFC.Fields.FindField(m_FldColNum)) = (lCol + 1)
569:                 If Len(m_FldScale) > 0 Then pInsertFeatureBuffer.Value(pFC.Fields.FindField(m_FldScale)) = m_dMapScale
570:                 pInsertFeatureCursor.InsertFeature pInsertFeatureBuffer
571:             End If
572:             If pProgress.ProgressBar1.Value < pProgress.ProgressBar1.Max Then
573:                 pProgress.ProgressBar1.Value = pProgress.ProgressBar1.Value + dIncrement
574:             Else
575:                 pProgress.lblInformation.Caption = "Warning: Val > Max (" & pProgress.ProgressBar1.Max & ")"
576:             End If
577:             If ((lRow * lColCount) + lCol) Mod 20 = 0 Then
578:                 DoEvents
579:                 pInsertFeatureCursor.Flush
580:             End If
581:             pProgress.Refresh
582:             If pProgress.Cancelled Then
                Dim vUserChoice
584:                 pProgress.Cancelled = False       ' Reset the form
585:                 vUserChoice = MsgBox("Operation cancelled." _
                    & "  Save the edits made thus far?" & vbCrLf & vbCrLf _
                    & "(Click Cancel to continue processing)", _
                                vbYesNoCancel, "Generate Grids")
589:                 If vUserChoice <> vbCancel Then
590:                     GoTo CancelledGenerateGrids     'Sorry for GoTo usage - in a hurry
591:                 End If
592:             End If
593:         Next
594:     Next
595:     pInsertFeatureCursor.Flush
596:     m_pProgress.Visible = False
    
    ' Stop editing
599:     pWorkspaceEdit.StopEditOperation
600:     pWorkspaceEdit.StopEditing True
    ' ----------------------------------------------------
    
603:     Screen.MousePointer = vbDefault
604:     pMx.ActiveView.Refresh
    
    Exit Sub
    
CancelledGenerateGrids:
609:     m_pProgress.Visible = False
610:     If vUserChoice = vbYes Then
611:         pInsertFeatureCursor.Flush
612:         pWorkspaceEdit.StopEditOperation
613:         pWorkspaceEdit.StopEditing True
614:     Else
615:         pWorkspaceEdit.StopEditOperation
616:         pWorkspaceEdit.StopEditing False
617:     End If
618:     Screen.MousePointer = vbDefault
619:     pMx.ActiveView.Refresh
    Exit Sub
    
622:     Resume
eh:
624:     m_pProgress.Visible = False
625:     Screen.MousePointer = vbDefault
626:     If Not pWorkspaceEdit Is Nothing Then
627:         If pWorkspaceEdit.IsBeingEdited Then
628:             pWorkspaceEdit.StopEditOperation
629:             pWorkspaceEdit.StopEditing False
630:         End If
631:     End If
632:     MsgBox "GenerateGrids - " & Erl & " - " & Err.Description
End Sub


Private Function CreateGridPoly2(pStartPoint As IPoint, lRow As Long, lCol As Long, _
                                dGridWidth As Double, dGridHeight As Double) As IPolygon
    Dim pPntColl As IPointCollection
    Dim pPoint As IPoint
    Dim dX As Double, dY As Double
    
642:     Set CreateGridPoly2 = New Polygon
643:     Set pPntColl = CreateGridPoly2
644:     Set pPoint = New esrigeometry.Point
645:     pPoint.PutCoords pStartPoint.X + (lCol * dGridWidth), pStartPoint.Y + (lRow * dGridHeight)
646:     pPntColl.AddPoint pPoint
647:     Set pPoint = New esrigeometry.Point
648:     pPoint.PutCoords pStartPoint.X + (lCol * dGridWidth), pStartPoint.Y + ((lRow + 1) * dGridHeight)
649:     pPntColl.AddPoint pPoint
650:     Set pPoint = New esrigeometry.Point
651:     pPoint.PutCoords pStartPoint.X + ((lCol + 1) * dGridWidth), pStartPoint.Y + ((lRow + 1) * dGridHeight)
652:     pPntColl.AddPoint pPoint
653:     Set pPoint = New esrigeometry.Point
654:     pPoint.PutCoords pStartPoint.X + ((lCol + 1) * dGridWidth), pStartPoint.Y + (lRow * dGridHeight)
655:     pPntColl.AddPoint pPoint
656:     Set pPoint = New esrigeometry.Point
657:     pPoint.PutCoords pStartPoint.X + (lCol * dGridWidth), pStartPoint.Y + (lRow * dGridHeight)
658:     pPntColl.AddPoint pPoint
    
End Function

Private Sub CalculateRowColCounts(ByVal XStart As Double, ByVal YStart As Double, _
                                  ByVal XEnd As Double, ByVal YEnd As Double, _
                                  ByVal XGridWidth As Double, ByVal YGridHeight As Double, _
                                  ByVal LayerExtent As IEnvelope, _
                                  ByRef ReturnColCount As Long, ByRef ReturnRowCount As Long)
    Dim XRange As Double
    Dim YRange As Double
    Dim dTemp As Double
    Dim lMultiplierX As Long
    Dim lMultiplierY As Long
    
    On Error GoTo eh
    
    ' Protect against div by zero (where the grid width/height is less than 0.5 => ie: with Geographic Data)
676:     lMultiplierX = 1
677:     dTemp = XGridWidth
678:     While dTemp < 10
679:         dTemp = dTemp * 10
680:         lMultiplierX = lMultiplierX * 10
681:     Wend
682:     lMultiplierY = 1
683:     dTemp = YGridHeight
684:     While dTemp < 10
685:         dTemp = dTemp * 10
686:         lMultiplierY = lMultiplierY * 10
687:     Wend
    
    ' Init
690:     XRange = XEnd - XStart
691:     YRange = YEnd - YStart
    'X ------------------------------------------
693:     If Not (LayerExtent Is Nothing) Then
694:         If XStart < LayerExtent.XMin Then
695:             Err.Raise vbObjectError, "CalculateRowColCounts", _
                "Starting X (" & XStart & ") is outside the valid layer extent (" & LayerExtent.XMin & ")"
697:         ElseIf XEnd > LayerExtent.XMax Then
698:             Err.Raise vbObjectError, "CalculateRowColCounts", _
                "Ending X (" & XStart & ") is outside the valid layer extent (" & LayerExtent.XMax & ")"
700:         ElseIf (XStart + XGridWidth) > LayerExtent.XMax Then
701:             Err.Raise vbObjectError, "CalculateRowColCounts", _
                "Adding a single Grid (width " & XGridWidth & ") would exceed valid X range for layer (" & LayerExtent.XMax & ")"
703:         End If
704:     End If
705:     If XRange < XGridWidth Then
706:         Err.Raise vbObjectError, "CalculateRowColCounts", _
            "Grid width " & XGridWidth & " exceeds range specified (" & XStart & " - " & XEnd & ")"
708:     ElseIf (XStart + XGridWidth) > XEnd Then
709:         Err.Raise vbObjectError, "CalculateRowColCounts", _
            "Adding a single Grid (width " & XGridWidth & ") would exceed specified X range (" & XEnd & ")"
711:     End If
712:     ReturnColCount = (XRange * lMultiplierX) \ (XGridWidth * lMultiplierX)
    'Y ------------------------------------------
714:     If Not (LayerExtent Is Nothing) Then
715:         If YStart < LayerExtent.YMin Then
716:             Err.Raise vbObjectError, "CalculateRowColCounts", _
                "Starting Y (" & YStart & ") is outside the valid layer extent (" & LayerExtent.YMin & ")"
718:         ElseIf YEnd > LayerExtent.YMax Then
719:             Err.Raise vbObjectError, "CalculateRowColCounts", _
                "Ending Y (" & YStart & ") is outside the valid layer extent (" & LayerExtent.YMax & ")"
721:         ElseIf (YStart + YGridHeight) > LayerExtent.YMax Then
722:             Err.Raise vbObjectError, "CalculateRowColCounts", _
                "Adding a single Grid (height " & YGridHeight & ") would exceed valid Y range for layer (" & LayerExtent.YMax & ")"
724:         End If
725:     End If
726:     If YRange < YGridHeight Then
727:         Err.Raise vbObjectError, "CalculateRowColCounts", _
            "Grid height " & YGridHeight & " exceeds range specified (" & YStart & " - " & YEnd & ")"
729:     ElseIf (XStart + XGridWidth) > XEnd Then
730:         Err.Raise vbObjectError, "CalculateRowColCounts", _
            "Adding a single Grid (height " & YGridHeight & ") would exceed specified Y range (" & YEnd & ")"
732:     End If
733:     ReturnRowCount = (YRange * lMultiplierY) \ (YGridHeight * lMultiplierY)
    '--------------------------------------------
    
    Exit Sub
eh:
738:     Err.Raise Err.Number, Err.Source, "Error in CalculateRowColCounts: " & Err.Description
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
    
    On Error GoTo eh
    
    ' Init
755:     Set pMx = pApp.Document
756:     Set pSR = pMx.FocusMap.SpatialReference
    ' If a Projected coord system
758:     If TypeOf pSR Is IProjectedCoordinateSystem Then
        ' Use meters per unit as the conversion
760:         Set pPCS = pSR
761:         dMetersPerUnit = pPCS.CoordinateUnit.MetersPerUnit
        ' Now convert this into page (ie: paper) units
763:         Set pPage = pMx.PageLayout.Page
764:         pPageUnits = pPage.Units
        Select Case pPageUnits
            Case esriInches: CalculatePageToMapRatio = dMetersPerUnit / (1 / 12 * 0.304800609601219)
            Case esriFeet: CalculatePageToMapRatio = dMetersPerUnit / (0.304800609601219)
            Case esriCentimeters: CalculatePageToMapRatio = dMetersPerUnit / (1 / 100)
            Case esriMeters: CalculatePageToMapRatio = dMetersPerUnit / (1)
            Case Else:
771:                 MsgBox "Warning: Only the following Page (Layout) Units are supported by this tool:" _
                    & vbCrLf & " - Inches, Feet, Centimeters, Meters" _
                    & vbCrLf & vbCrLf & "Calculating as though Page Units are in Inches..."
774:                 CalculatePageToMapRatio = dMetersPerUnit / (1 / 12 * 0.304800609601219)
775:         End Select
    ' Otherwise
777:     Else
        ' If not projected, we can only do a "flat" conversion -> that is, use the current scale and extent
        '  as a ratio to be applied to the map grid scale.
        ' NOTE: We MUST be in Layout mode to make this calculation, as the scale in Map View and Layout View
        '  are not the same (as the extent envelope and data frame envelope can be different shapes).  The
        '  test for being in Layout Mode is made in the clsMapGridButton.ICommand_Enabled property.
783:         Set pExtentEnv = pMx.ActiveView.Extent
784:         dCurrScale = pMx.FocusMap.MapScale
785:         If ((m_EndX - m_StartX) / m_dFrameWidthInPageUnits) > ((m_EndY - m_StartY) / m_dFrameHeightInPageUnits) Then
786:             CalculatePageToMapRatio = m_dFrameWidthInPageUnits / ((m_EndX - m_StartX) / dCurrScale)
787:         Else
788:             CalculatePageToMapRatio = m_dFrameHeightInPageUnits / ((m_EndY - m_StartY) / dCurrScale)
789:         End If
790:     End If
    
    Exit Function
eh:
794:     CalculatePageToMapRatio = 1
795:     MsgBox "Error in CalculatePageToMapRatio" & vbCrLf & Err.Description
End Function

Private Function HasFeatures(pPolygon As IPolygon, pMap As IMap) As Boolean
    Dim lLoop As Long
    Dim pFL As IFeatureLayer
    Dim pSF As ISpatialFilter
    Dim pFC As IFeatureCursor
    Dim pF As IFeature
    
805:     HasFeatures = False
806:     Set pSF = New SpatialFilter
807:     Set pSF.Geometry = pPolygon
808:     pSF.SpatialRel = esriSpatialRelIntersects
809:     For lLoop = 0 To (pMap.LayerCount - 1)
810:         If TypeOf pMap.Layer(lLoop) Is IFeatureLayer Then
811:             Set pFL = pMap.Layer(lLoop)
812:             If pFL.Name <> m_DestFL.Name And IsARequiredLayer(pFL.Name) Then
813:                 pSF.GeometryField = pFL.FeatureClass.shapeFieldName
814:                 Set pFC = pFL.Search(pSF, False)
815:                 Set pF = pFC.NextFeature
816:                 If Not pF Is Nothing Then
817:                     HasFeatures = True
818:                     Exit For
819:                 End If
820:             End If
821:         End If
822:     Next
End Function

Private Function IsARequiredLayer(sLayerName As String) As Boolean
    Dim lLoop As Long
827:     For lLoop = 1 To m_colLayerNamesForData.count
828:         If UCase(m_colLayerNamesForData.Item(lLoop)) = UCase(sLayerName) Then
829:             IsARequiredLayer = True
830:             Exit For
831:         End If
832:     Next
End Function

Private Sub Class_Terminate()
836:     Set m_DestFL = Nothing
837:     Set m_DestFC = Nothing
838:     Set m_pProgress = Nothing
End Sub

