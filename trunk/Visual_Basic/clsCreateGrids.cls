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
82:     If pFL.FeatureClass.ShapeType = esriGeometryPolygon Then
83:         Set m_DestFL = pFL
84:         Set m_DestFC = pFL.FeatureClass
85:     Else
86:         Err.Raise vbObjectError, "MapGridManager_Set_DestinationFeatureLayer", _
            "Not a polygon feature layer"
88:     End If
End Property

Public Property Get DestinationFeatureLayer() As IFeatureLayer
92:     Set DestinationFeatureLayer = m_DestFL
End Property

Public Property Let FrameWidthInPageUnits(dWidth As Double)
96:     m_dFrameWidthInPageUnits = dWidth
End Property

Public Property Get FrameWidthInPageUnits() As Double
100:     FrameWidthInPageUnits = m_dFrameWidthInPageUnits
End Property

Public Property Let FrameHeightInPageUnits(dHeight As Double)
104:     m_dFrameHeightInPageUnits = dHeight
End Property

Public Property Get FrameHeightInPageUnits() As Double
108:     FrameHeightInPageUnits = m_dFrameHeightInPageUnits
End Property

Public Property Let MapScale(dScale As Double)
112:     m_dMapScale = dScale
End Property

Public Property Get MapScale() As Double
116:     MapScale = m_dMapScale
End Property

Public Property Let IdentifierOrder(mgType As mgGridIdentifierOrderType)
120:     m_IDOrderType = mgType
End Property

Public Property Get IdentifierOrder() As mgGridIdentifierOrderType
124:     IdentifierOrder = m_IDOrderType
End Property

Public Property Let RowIDType(mgIDStyle As mgGridIdentifierValueType)
128:     m_RowIDType = mgIDStyle
End Property

Public Property Get RowIDType() As mgGridIdentifierValueType
132:     RowIDType = m_RowIDType
End Property

Public Property Let ColIDType(mgIDStyle As mgGridIdentifierValueType)
136:     m_ColIDType = mgIDStyle
End Property

Public Property Get ColIDType() As mgGridIdentifierValueType
140:     ColIDType = m_ColIDType
End Property

Public Property Let IDStartPositionType(mgStartPos As mgGridIdentifierStartPositionType)
144:     m_StartIDType = mgStartPos
End Property

Public Property Get IDStartPositionType() As mgGridIdentifierStartPositionType
148:     IDStartPositionType = m_StartIDType
End Property

Public Property Let StartingCoordinateLL_X(X As Double)
152:     m_StartX = X
End Property

Public Property Get StartingCoordinateLL_X() As Double
156:     StartingCoordinateLL_X = m_StartX
End Property

Public Property Let StartingCoordinateLL_Y(Y As Double)
160:     m_StartY = Y
End Property

Public Property Get StartingCoordinateLL_Y() As Double
164:     StartingCoordinateLL_Y = m_StartY
End Property

Public Property Let EndingCoordinateUR_X(X As Double)
168:     m_EndX = X
End Property

Public Property Get EndingCoordinateUR_X() As Double
172:     EndingCoordinateUR_X = m_EndX
End Property

Public Property Let EndingCoordinateUR_Y(Y As Double)
176:     m_EndY = Y
End Property

Public Property Get EndingCoordinateUR_Y() As Double
180:     EndingCoordinateUR_Y = m_EndY
End Property

Public Property Let FieldNameGridID(FieldName As String)
184:     m_FldID = FieldName
End Property

Public Property Let FieldNameRowNum(FieldName As String)
188:     m_FldRowNum = FieldName
End Property

Public Property Let FieldNameColNum(FieldName As String)
192:     m_FldColNum = FieldName
End Property

Public Property Let FieldNameScale(FieldName As String)
196:     m_FldScale = FieldName
End Property

Public Property Let UseUnderscore(UseUnderscoreInID As Boolean)
200:     m_UseUnderscore = UseUnderscoreInID
End Property

Public Property Get UseUnderscore() As Boolean
204:     UseUnderscore = m_UseUnderscore
End Property

Public Property Let RemoveCurrentGrids(RemoveGrids As Boolean)
208:     m_RemoveGrids = RemoveGrids
End Property

Public Property Get RemoveCurrentGrids() As Boolean
212:     RemoveCurrentGrids = m_RemoveGrids
End Property

Public Property Let NoEmptyGrids(NoEmptyGridPolygons As Boolean)
216:     m_NoEmptyGrids = NoEmptyGridPolygons
End Property

Public Property Get NoEmptyGrids() As Boolean
220:     NoEmptyGrids = m_NoEmptyGrids
End Property

Public Sub AddRequiredDataLayer(sLayerName As String)
224:     m_colLayerNamesForData.Add sLayerName
End Sub

Public Property Get RequiredDataLayer(LayerNumber As Long) As String
228:     If LayerNumber <= m_colLayerNamesForData.count Then
229:         RequiredDataLayer = m_colLayerNamesForData.Item(LayerNumber - 1)
230:     Else
231:         If m_colLayerNamesForData.count = 0 Then
232:             RequiredDataLayer = "There are no Required Data Layers."
233:         Else
234:             RequiredDataLayer = "Invalid layer number.  Valid range: 0 - " & (m_colLayerNamesForData.count - 1)
235:         End If
236:     End If
End Property

Public Property Get RequiredDataLayerCount() As Long
240:     RequiredDataLayerCount = m_colLayerNamesForData.count
End Property

Public Sub ClearRequiredDataLayers()
244:     Set m_colLayerNamesForData = New Collection
End Sub

Private Sub Class_Initialize()
    ' Set the default grid identifier (row-col, alpha/numeric => eg: "B02")
249:     m_IDOrderType = Row_Column
250:     m_RowIDType = Alphabetical
251:     m_ColIDType = Numerical
252:     Set m_colLayerNamesForData = New Collection
253:     Set m_pProgress = New ModelessFrame
End Sub

Private Function CalculateID(ByVal lRow As Long, ByVal lCol As Long, _
                             ByVal iRowIDLen As Integer, ByVal iColIDLen As Integer) As String
    Dim sRowID As String
    Dim sColID As String
    Dim sNumericFormat As String
    Dim lLoop As Long
    Dim lTmp As Long, lCalc As Long, lTmp2 As Long

    ' Row ---------------------------------------------
265:     sNumericFormat = ""
266:     For lLoop = 1 To iRowIDLen
267:         sNumericFormat = sNumericFormat & "0"
268:     Next
269:     If m_RowIDType = Alphabetical Then
270:         For lLoop = 1 To iRowIDLen
271:             lTmp = 26 ^ lLoop
272:             lTmp2 = (26 ^ (lLoop - 1))
273:             If lRow >= lTmp2 Then
274:                 lCalc = ((((lRow - lTmp2) / lTmp) * 26) + 1) Mod 26
275:                 sRowID = Chr(Asc("A") + lCalc) & sRowID
276:                 lRow = lRow - (lCalc * lTmp2)
277:             Else
278:                 sRowID = "A" & sRowID
279:             End If
280:         Next
281:     Else
282:         sRowID = Format(lRow + 1, sNumericFormat)
283:     End If
    ' Col ---------------------------------------------
285:     sNumericFormat = ""
286:     For lLoop = 1 To iColIDLen
287:         sNumericFormat = sNumericFormat & "0"
288:     Next
289:     If m_ColIDType = Alphabetical Then
290:         For lLoop = 1 To iColIDLen
291:             lTmp = 26 ^ lLoop
292:             lTmp2 = (26 ^ (lLoop - 1))
293:             If lCol >= lTmp2 Then
294:                 lCalc = ((((lCol - lTmp2) / lTmp) * 26) + 1) Mod 26
295:                 sColID = Chr(Asc("A") + lCalc) & sColID
296:                 lCol = lCol - (lCalc * lTmp2)
297:             Else
298:                 sColID = "A" & sColID
299:             End If
300:         Next
301:     Else
302:         sColID = Format(lCol + 1, sNumericFormat)
303:     End If
    ' Join --------------------------------------------
305:     If m_IDOrderType = Row_Column Then
306:         If m_UseUnderscore Then
307:             CalculateID = sRowID & "_" & sColID
308:         Else
309:             CalculateID = sRowID & sColID
310:         End If
311:     Else
312:         If m_UseUnderscore Then
313:             CalculateID = sColID & "_" & sRowID
314:         Else
315:             CalculateID = sColID & sRowID
316:         End If
317:     End If
End Function

Private Function GetMinimumStringLength(lValue As Long, lBase As Long) As Integer
    Dim lTmp As Long, lIndex As Long
    
    On Error GoTo eh
    ' ROW
325:     lTmp = lBase
326:     lIndex = 1
327:     While lValue > (lTmp - 1)
328:         lTmp = lTmp * lBase
329:         lIndex = lIndex + 1
330:     Wend
331:     GetMinimumStringLength = lIndex
    Exit Function
eh:
334:     Err.Raise Err.Number, "GetMinimumStringLength", "Error in GetMinimumStringLength: " & Err.Description
End Function

Public Sub RunStandardGUI(pApp As IApplication)
338:     Set frmGridSettings.m_Application = pApp
339:     frmGridSettings.Tickle
340:     SetWindowLong frmGridSettings.hwnd, GWL_HWNDPARENT, pApp.hwnd
341:     frmGridSettings.Show vbModeless
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
416:     Screen.MousePointer = vbArrowHourglass
    
    ' Init
419:     Set pMx = Application.Document
420:     For lLoop = 0 To pMx.FocusMap.LayerCount - 1
421:         If TypeOf pMx.FocusMap.Layer(lLoop) Is IFeatureLayer Then
422:             If UCase(pMx.FocusMap.Layer(lLoop).Name) = UCase(m_DestFL.Name) Then
423:                 Set pFL = pMx.FocusMap.Layer(lLoop)
424:                 Exit For
425:             End If
426:         End If
427:     Next
428:     If pFL Is Nothing Then
429:         MsgBox "No match in the Map for layer '" & m_DestFL.Name & "'."
        Exit Sub
431:     End If
432:     Set pFC = pFL.FeatureClass
    ' Check for required fields - that the field exists
    Dim bErrorWithFields As Boolean
435:     bErrorWithFields = (pFC.FindField(m_FldID) < 0)
436:     If Len(m_FldRowNum) > 0 Then bErrorWithFields = bErrorWithFields Or (pFC.FindField(m_FldRowNum) < 0)
437:     If Len(m_FldColNum) > 0 Then bErrorWithFields = bErrorWithFields Or (pFC.FindField(m_FldColNum) < 0)
438:     If Len(m_FldScale) > 0 Then bErrorWithFields = bErrorWithFields Or (pFC.FindField(m_FldScale) < 0)
    ' If error
440:     If bErrorWithFields Then
441:         Err.Raise vbObjectError, "GenerateGrids", "Could not find all the given field names in " & pFL.Name & "." _
            & vbCrLf & " - " & m_FldID & ", " & m_FldRowNum & ", " & m_FldColNum & ", " & m_FldScale
443:     End If
    ' Check the field types
445:     bErrorWithFields = (pFC.Fields.Field(pFC.FindField(m_FldID)).Type <> esriFieldTypeString)
446:     If Len(m_FldRowNum) > 0 Then
447:         bErrorWithFields = bErrorWithFields Or _
           ((pFC.Fields.Field(pFC.FindField(m_FldRowNum)).Type <> esriFieldTypeDouble) And _
            (pFC.Fields.Field(pFC.FindField(m_FldRowNum)).Type <> esriFieldTypeInteger) And _
            (pFC.Fields.Field(pFC.FindField(m_FldRowNum)).Type <> esriFieldTypeSingle) And _
            (pFC.Fields.Field(pFC.FindField(m_FldRowNum)).Type <> esriFieldTypeSmallInteger))
452:     End If
453:     If Len(m_FldColNum) > 0 Then
454:         bErrorWithFields = bErrorWithFields Or _
           ((pFC.Fields.Field(pFC.FindField(m_FldColNum)).Type <> esriFieldTypeDouble) And _
            (pFC.Fields.Field(pFC.FindField(m_FldColNum)).Type <> esriFieldTypeInteger) And _
            (pFC.Fields.Field(pFC.FindField(m_FldColNum)).Type <> esriFieldTypeSingle) And _
            (pFC.Fields.Field(pFC.FindField(m_FldColNum)).Type <> esriFieldTypeSmallInteger))
459:     End If
460:     If Len(m_FldScale) > 0 Then
461:         bErrorWithFields = bErrorWithFields Or _
           ((pFC.Fields.Field(pFC.FindField(m_FldScale)).Type <> esriFieldTypeDouble) And _
            (pFC.Fields.Field(pFC.FindField(m_FldScale)).Type <> esriFieldTypeInteger) And _
            (pFC.Fields.Field(pFC.FindField(m_FldScale)).Type <> esriFieldTypeSingle) And _
            (pFC.Fields.Field(pFC.FindField(m_FldScale)).Type <> esriFieldTypeSmallInteger))
466:     End If
    ' if error
468:     If bErrorWithFields Then
469:         Err.Raise vbObjectError, "GenerateGrids", "Given field names are not of the correct type." _
            & vbCrLf & "Grid ID field must be a Text field, all others must be numeric fields."
471:     End If
    ' Get the dataset and workspace (to start editing upon)
473:     Set pFeatDataset = pFC.FeatureDataset
474:     If Not pFeatDataset Is Nothing Then
475:         Set pWorkspaceEdit = pFeatDataset.Workspace
476:     Else
        ' Is a shapefile, go via just IDataset
478:         Set pDataset = pFC
479:         Set pWorkspaceEdit = pDataset.Workspace
480:     End If
481:     dDataFrameWidth = m_dFrameWidthInPageUnits
482:     dDataFrameHeight = m_dFrameHeightInPageUnits
    ' Start Editing
484:     pWorkspaceEdit.StartEditing False
485:     pWorkspaceEdit.StartEditOperation
    
    ' If replacing, delete all existing polygons
488:     Set pProgress = New frmProgress
489:     m_pProgress.Create pProgress
490:     If m_RemoveGrids Then
        Dim pFCu As IFeatureCursor
        Dim pT As ITable
493:         Set pFCu = m_DestFL.Search(Nothing, False)
494:         Set pT = m_DestFL.FeatureClass
495:         pProgress.ProgressBar1.Min = 0
496:         pProgress.ProgressBar1.Max = 100
497:         If pT.RowCount(Nothing) = 0 Then
498:             dIncrement = 99
499:         Else
500:             dIncrement = 100 / pT.RowCount(Nothing)
501:         End If
502:         pProgress.ProgressBar1.value = 0
503:         pProgress.lblInformation.Caption = "Deleting previous grids..."
504:         pProgress.cmdCancel.Visible = False        ' User cannot cancel this step
        
506:         m_pProgress.Visible = True
507:         Set pFeature = pFCu.NextFeature
508:         While Not pFeature Is Nothing
509:             pFeature.Delete
510:             If (pProgress.ProgressBar1.value + dIncrement) <= pProgress.ProgressBar1.Max Then
511:                 pProgress.ProgressBar1.value = pProgress.ProgressBar1.value + dIncrement
512:             Else
513:                 pProgress.lblInformation.Caption = "Warning: Val > Max (" & pProgress.ProgressBar1.Max & ")"
514:             End If
515:             Set pFeature = pFCu.NextFeature
516:         Wend
517:         m_pProgress.Visible = False
518:     End If
    
    ' Calc the row/column extents, grid size (map units), ID lengths and starting coordinate
521:     Set pStartingCoord = New esrigeometry.Point
522:     pStartingCoord.PutCoords m_StartX, m_StartY
523:     dConvertPageToMapUnits = CalculatePageToMapRatio(Application)
524:     dGridSizeW = ((m_dMapScale * dDataFrameWidth) / dConvertPageToMapUnits)
525:     dGridSizeH = ((m_dMapScale * dDataFrameHeight) / dConvertPageToMapUnits)
526:     If Not (pFL.FeatureClass.FeatureDataset Is Nothing) Then
527:         CalculateRowColCounts m_StartX, m_StartY, m_EndX, m_EndY, _
            dGridSizeW, dGridSizeH, GetValidExtentForLayer(m_DestFL), lColCount, lRowCount
529:     Else
530:         CalculateRowColCounts m_StartX, m_StartY, m_EndX, m_EndY, _
            dGridSizeW, dGridSizeH, Nothing, lColCount, lRowCount
532:     End If
533:     If lColCount = 0 Or lRowCount = 0 Then
534:         Err.Raise vbObjectError, "GenerateGrids", "CalculateRowColCounts returned zero row/columns"
535:     End If
536:     If m_ColIDType = Alphabetical Then
537:         lBase = 26
538:     Else
539:         lBase = 10
540:     End If
541:     iStringLengthCol = GetMinimumStringLength(lColCount, lBase)
542:     If m_RowIDType = Alphabetical Then
543:         lBase = 26
544:     Else
545:         lBase = 10
546:     End If
547:     iStringLengthRow = GetMinimumStringLength(lRowCount, lBase)
    
    ' Create and add the new Grid polygon features
550:     bOKToAdd = True
551:     pProgress.ProgressBar1.Min = 0
552:     pProgress.ProgressBar1.Max = 100
553:     pProgress.ProgressBar1.value = 0
554:     pProgress.lblInformation.Caption = "Creating new Grid features..."
555:     pProgress.cmdCancel.Visible = True
556:     m_pProgress.Visible = True
557:     dIncrement = 99.9 / (lRowCount * lColCount)
558:     Set pInsertFeatureCursor = pFC.Insert(True)
559:     Set pInsertFeatureBuffer = pFC.CreateFeatureBuffer
560:     For lRow = 0 To lRowCount - 1
561:         For lCol = 0 To lColCount - 1
            ' Create the source polygon
563:             Set pGridPolygon = CreateGridPoly2(pStartingCoord, lRow, lCol, dGridSizeW, dGridSizeH)
            ' If required, check for containing features
565:             If m_NoEmptyGrids Then
566:                 bOKToAdd = HasFeatures(pGridPolygon, pMx.FocusMap)
567:             End If
568:             If bOKToAdd Then
                ' Create new grid cell feature
570:                 Set pInsertFeatureBuffer.Shape = pGridPolygon
571:                 If m_StartIDType = TopLeft Then
572:                     pInsertFeatureBuffer.value(pFC.Fields.FindField(m_FldID)) = _
                            CalculateID((lRowCount - 1) - lRow, lCol, iStringLengthRow, iStringLengthCol)
574:                     If Len(m_FldRowNum) > 0 Then pInsertFeatureBuffer.value(pFC.Fields.FindField(m_FldRowNum)) = (lRowCount - lRow)
575:                 Else
576:                     pInsertFeatureBuffer.value(pFC.Fields.FindField(m_FldID)) = _
                            CalculateID(lRow, lCol, iStringLengthRow, iStringLengthCol)
578:                     If Len(m_FldRowNum) > 0 Then pInsertFeatureBuffer.value(pFC.Fields.FindField(m_FldRowNum)) = (lRow + 1)
579:                 End If
580:                 If Len(m_FldColNum) > 0 Then pInsertFeatureBuffer.value(pFC.Fields.FindField(m_FldColNum)) = (lCol + 1)
581:                 If Len(m_FldScale) > 0 Then pInsertFeatureBuffer.value(pFC.Fields.FindField(m_FldScale)) = m_dMapScale
582:                 pInsertFeatureCursor.InsertFeature pInsertFeatureBuffer
583:             End If
584:             If pProgress.ProgressBar1.value < pProgress.ProgressBar1.Max Then
585:                 pProgress.ProgressBar1.value = pProgress.ProgressBar1.value + dIncrement
586:             Else
587:                 pProgress.lblInformation.Caption = "Warning: Val > Max (" & pProgress.ProgressBar1.Max & ")"
588:             End If
589:             If ((lRow * lColCount) + lCol) Mod 20 = 0 Then
590:                 DoEvents
591:                 pInsertFeatureCursor.Flush
592:             End If
593:             pProgress.Refresh
594:             If pProgress.Cancelled Then
                Dim vUserChoice
596:                 pProgress.Cancelled = False       ' Reset the form
597:                 vUserChoice = MsgBox("Operation cancelled." _
                    & "  Save the edits made thus far?" & vbCrLf & vbCrLf _
                    & "(Click Cancel to continue processing)", _
                                vbYesNoCancel, "Generate Grids")
601:                 If vUserChoice <> vbCancel Then
602:                     GoTo CancelledGenerateGrids     'Sorry for GoTo usage - in a hurry
603:                 End If
604:             End If
605:         Next
606:     Next
607:     pInsertFeatureCursor.Flush
608:     m_pProgress.Visible = False
    
    ' Stop editing
611:     pWorkspaceEdit.StopEditOperation
612:     pWorkspaceEdit.StopEditing True
    ' ----------------------------------------------------
    
615:     Screen.MousePointer = vbDefault
616:     pMx.ActiveView.Refresh
    
    Exit Sub
    
CancelledGenerateGrids:
621:     m_pProgress.Visible = False
622:     If vUserChoice = vbYes Then
623:         pInsertFeatureCursor.Flush
624:         pWorkspaceEdit.StopEditOperation
625:         pWorkspaceEdit.StopEditing True
626:     Else
627:         pWorkspaceEdit.StopEditOperation
628:         pWorkspaceEdit.StopEditing False
629:     End If
630:     Screen.MousePointer = vbDefault
631:     pMx.ActiveView.Refresh
    Exit Sub
    
634:     Resume
eh:
636:     m_pProgress.Visible = False
637:     Screen.MousePointer = vbDefault
638:     If Not pWorkspaceEdit Is Nothing Then
639:         If pWorkspaceEdit.IsBeingEdited Then
640:             pWorkspaceEdit.StopEditOperation
641:             pWorkspaceEdit.StopEditing False
642:         End If
643:     End If
644:     MsgBox "GenerateGrids - " & Erl & " - " & Err.Description
End Sub


Private Function CreateGridPoly2(pStartPoint As IPoint, lRow As Long, lCol As Long, _
                                dGridWidth As Double, dGridHeight As Double) As IPolygon
    Dim pPntColl As IPointCollection
    Dim pPoint As IPoint
    Dim dX As Double, dY As Double
    
654:     Set CreateGridPoly2 = New Polygon
655:     Set pPntColl = CreateGridPoly2
656:     Set pPoint = New esrigeometry.Point
657:     pPoint.PutCoords pStartPoint.X + (lCol * dGridWidth), pStartPoint.Y + (lRow * dGridHeight)
658:     pPntColl.AddPoint pPoint
659:     Set pPoint = New esrigeometry.Point
660:     pPoint.PutCoords pStartPoint.X + (lCol * dGridWidth), pStartPoint.Y + ((lRow + 1) * dGridHeight)
661:     pPntColl.AddPoint pPoint
662:     Set pPoint = New esrigeometry.Point
663:     pPoint.PutCoords pStartPoint.X + ((lCol + 1) * dGridWidth), pStartPoint.Y + ((lRow + 1) * dGridHeight)
664:     pPntColl.AddPoint pPoint
665:     Set pPoint = New esrigeometry.Point
666:     pPoint.PutCoords pStartPoint.X + ((lCol + 1) * dGridWidth), pStartPoint.Y + (lRow * dGridHeight)
667:     pPntColl.AddPoint pPoint
668:     Set pPoint = New esrigeometry.Point
669:     pPoint.PutCoords pStartPoint.X + (lCol * dGridWidth), pStartPoint.Y + (lRow * dGridHeight)
670:     pPntColl.AddPoint pPoint
    
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
688:     lMultiplierX = 1
689:     dTemp = XGridWidth
690:     While dTemp < 10
691:         dTemp = dTemp * 10
692:         lMultiplierX = lMultiplierX * 10
693:     Wend
694:     lMultiplierY = 1
695:     dTemp = YGridHeight
696:     While dTemp < 10
697:         dTemp = dTemp * 10
698:         lMultiplierY = lMultiplierY * 10
699:     Wend
    
    ' Init
702:     XRange = XEnd - XStart
703:     YRange = YEnd - YStart
    'X ------------------------------------------
705:     If Not (LayerExtent Is Nothing) Then
706:         If XStart < LayerExtent.XMin Then
707:             Err.Raise vbObjectError, "CalculateRowColCounts", _
                "Starting X (" & XStart & ") is outside the valid layer extent (" & LayerExtent.XMin & ")"
709:         ElseIf XEnd > LayerExtent.XMax Then
710:             Err.Raise vbObjectError, "CalculateRowColCounts", _
                "Ending X (" & XStart & ") is outside the valid layer extent (" & LayerExtent.XMax & ")"
712:         ElseIf (XStart + XGridWidth) > LayerExtent.XMax Then
713:             Err.Raise vbObjectError, "CalculateRowColCounts", _
                "Adding a single Grid (width " & XGridWidth & ") would exceed valid X range for layer (" & LayerExtent.XMax & ")"
715:         End If
716:     End If
717:     If XRange < XGridWidth Then
718:         Err.Raise vbObjectError, "CalculateRowColCounts", _
            "Grid width " & XGridWidth & " exceeds range specified (" & XStart & " - " & XEnd & ")"
720:     ElseIf (XStart + XGridWidth) > XEnd Then
721:         Err.Raise vbObjectError, "CalculateRowColCounts", _
            "Adding a single Grid (width " & XGridWidth & ") would exceed specified X range (" & XEnd & ")"
723:     End If
724:     ReturnColCount = (XRange * lMultiplierX) \ (XGridWidth * lMultiplierX)
    'Y ------------------------------------------
726:     If Not (LayerExtent Is Nothing) Then
727:         If YStart < LayerExtent.YMin Then
728:             Err.Raise vbObjectError, "CalculateRowColCounts", _
                "Starting Y (" & YStart & ") is outside the valid layer extent (" & LayerExtent.YMin & ")"
730:         ElseIf YEnd > LayerExtent.YMax Then
731:             Err.Raise vbObjectError, "CalculateRowColCounts", _
                "Ending Y (" & YStart & ") is outside the valid layer extent (" & LayerExtent.YMax & ")"
733:         ElseIf (YStart + YGridHeight) > LayerExtent.YMax Then
734:             Err.Raise vbObjectError, "CalculateRowColCounts", _
                "Adding a single Grid (height " & YGridHeight & ") would exceed valid Y range for layer (" & LayerExtent.YMax & ")"
736:         End If
737:     End If
738:     If YRange < YGridHeight Then
739:         Err.Raise vbObjectError, "CalculateRowColCounts", _
            "Grid height " & YGridHeight & " exceeds range specified (" & YStart & " - " & YEnd & ")"
741:     ElseIf (XStart + XGridWidth) > XEnd Then
742:         Err.Raise vbObjectError, "CalculateRowColCounts", _
            "Adding a single Grid (height " & YGridHeight & ") would exceed specified Y range (" & YEnd & ")"
744:     End If
745:     ReturnRowCount = (YRange * lMultiplierY) \ (YGridHeight * lMultiplierY)
    '--------------------------------------------
    
    Exit Sub
eh:
750:     Err.Raise Err.Number, Err.source, "Error in CalculateRowColCounts: " & Err.Description
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
767:     Set pMx = pApp.Document
768:     Set pSR = pMx.FocusMap.SpatialReference
    ' If a Projected coord system
770:     If TypeOf pSR Is IProjectedCoordinateSystem Then
        ' Use meters per unit as the conversion
772:         Set pPCS = pSR
773:         dMetersPerUnit = pPCS.CoordinateUnit.MetersPerUnit
        ' Now convert this into page (ie: paper) units
775:         Set pPage = pMx.PageLayout.Page
776:         pPageUnits = pPage.Units
        Select Case pPageUnits
            Case esriInches: CalculatePageToMapRatio = dMetersPerUnit / (1 / 12 * 0.304800609601219)
            Case esriFeet: CalculatePageToMapRatio = dMetersPerUnit / (0.304800609601219)
            Case esriCentimeters: CalculatePageToMapRatio = dMetersPerUnit / (1 / 100)
            Case esriMeters: CalculatePageToMapRatio = dMetersPerUnit / (1)
            Case Else:
783:                 MsgBox "Warning: Only the following Page (Layout) Units are supported by this tool:" _
                    & vbCrLf & " - Inches, Feet, Centimeters, Meters" _
                    & vbCrLf & vbCrLf & "Calculating as though Page Units are in Inches..."
786:                 CalculatePageToMapRatio = dMetersPerUnit / (1 / 12 * 0.304800609601219)
787:         End Select
    ' Otherwise
789:     Else
        ' If not projected, we can only do a "flat" conversion -> that is, use the current scale and extent
        '  as a ratio to be applied to the map grid scale.
        ' NOTE: We MUST be in Layout mode to make this calculation, as the scale in Map View and Layout View
        '  are not the same (as the extent envelope and data frame envelope can be different shapes).  The
        '  test for being in Layout Mode is made in the clsMapGridButton.ICommand_Enabled property.
795:         Set pExtentEnv = pMx.ActiveView.Extent
796:         dCurrScale = pMx.FocusMap.MapScale
797:         If ((m_EndX - m_StartX) / m_dFrameWidthInPageUnits) > ((m_EndY - m_StartY) / m_dFrameHeightInPageUnits) Then
798:             CalculatePageToMapRatio = m_dFrameWidthInPageUnits / ((m_EndX - m_StartX) / dCurrScale)
799:         Else
800:             CalculatePageToMapRatio = m_dFrameHeightInPageUnits / ((m_EndY - m_StartY) / dCurrScale)
801:         End If
802:     End If
    
    Exit Function
eh:
806:     CalculatePageToMapRatio = 1
807:     MsgBox "Error in CalculatePageToMapRatio" & vbCrLf & Err.Description
End Function

Private Function HasFeatures(pPolygon As IPolygon, pMap As IMap) As Boolean
    Dim lLoop As Long
    Dim pFL As IFeatureLayer
    Dim pSF As ISpatialFilter
    Dim pFC As IFeatureCursor
    Dim pF As IFeature
    
817:     HasFeatures = False
818:     Set pSF = New SpatialFilter
819:     Set pSF.Geometry = pPolygon
820:     pSF.SpatialRel = esriSpatialRelIntersects
821:     For lLoop = 0 To (pMap.LayerCount - 1)
822:         If TypeOf pMap.Layer(lLoop) Is IFeatureLayer Then
823:             Set pFL = pMap.Layer(lLoop)
824:             If pFL.Name <> m_DestFL.Name And IsARequiredLayer(pFL.Name) Then
825:                 pSF.GeometryField = pFL.FeatureClass.shapeFieldName
826:                 Set pFC = pFL.Search(pSF, False)
827:                 Set pF = pFC.NextFeature
828:                 If Not pF Is Nothing Then
829:                     HasFeatures = True
830:                     Exit For
831:                 End If
832:             End If
833:         End If
834:     Next
End Function

Private Function IsARequiredLayer(sLayerName As String) As Boolean
    Dim lLoop As Long
839:     For lLoop = 1 To m_colLayerNamesForData.count
840:         If UCase(m_colLayerNamesForData.Item(lLoop)) = UCase(sLayerName) Then
841:             IsARequiredLayer = True
842:             Exit For
843:         End If
844:     Next
End Function

Private Sub Class_Terminate()
848:     Set m_DestFL = Nothing
849:     Set m_DestFC = Nothing
850:     Set m_pProgress = Nothing
End Sub

