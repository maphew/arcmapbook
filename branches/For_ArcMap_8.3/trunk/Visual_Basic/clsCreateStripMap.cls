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
'----------------------------------------------
' API call to keep form top most
Private Const GWL_HWNDPARENT = -8
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
  (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Property Set DestinationFeatureLayer(pFL As IFeatureLayer)
    If pFL.FeatureClass.ShapeType = esriGeometryPolygon Then
        Set m_DestFL = pFL
        Set m_DestFC = pFL.FeatureClass
    Else
        Err.Raise vbObjectError, "MapGridManager_Set_DestinationFeatureLayer", _
            "Not a polygon feature layer"
    End If
End Property

Public Property Get DestinationFeatureLayer() As IFeatureLayer
    Set DestinationFeatureLayer = m_DestFL
End Property

Public Property Set StripMapRoute(RoutePolyline As IPolyline)
    Set m_Polyline = RoutePolyline
End Property

Public Property Get StripMapRoute() As IPolyline
    Set StripMapRoute = m_Polyline
End Property

Public Property Let FrameWidthInPageUnits(dWidth As Double)
    m_dFrameWidthInPageUnits = dWidth
End Property

Public Property Get FrameWidthInPageUnits() As Double
    FrameWidthInPageUnits = m_dFrameWidthInPageUnits
End Property

Public Property Let FrameHeightInPageUnits(dHeight As Double)
    m_dFrameHeightInPageUnits = dHeight
End Property

Public Property Get FrameHeightInPageUnits() As Double
    FrameHeightInPageUnits = m_dFrameHeightInPageUnits
End Property

Public Property Let MapScale(dScale As Double)
    m_dMapScale = dScale
End Property

Public Property Get MapScale() As Double
    MapScale = m_dMapScale
End Property

Public Property Let StripMapName(MapName As String)
    m_StripMapName = MapName
End Property

Public Property Get StripMapName() As String
    StripMapName = m_StripMapName
End Property


Public Property Let FieldNameStripMapName(FieldName As String)
    m_FldStripName = FieldName
End Property

Public Property Get FieldNameStripMapName() As String
    FieldNameStripMapName = m_FldStripName
End Property

Public Property Let FieldNameNumberInSeries(FieldName As String)
    m_FldNumInSeries = FieldName
End Property

Public Property Get FieldNameNumberInSeries() As String
    FieldNameNumberInSeries = m_FldNumInSeries
End Property

Public Property Let FieldNameMapAngle(FieldName As String)
    m_FldMapAngle = FieldName
End Property

Public Property Get FieldNameMapAngle() As String
    FieldNameMapAngle = m_FldMapAngle
End Property

Public Property Let FieldNameScale(FieldName As String)
    m_FldScale = FieldName
End Property

Public Property Get FieldNameScale() As String
    FieldNameScale = m_FldScale
End Property

Public Property Let RemoveCurrentGrids(RemoveGrids As Boolean)
    m_RemoveGrids = RemoveGrids
End Property

Public Property Get RemoveCurrentGrids() As Boolean
    RemoveCurrentGrids = m_RemoveGrids
End Property

Public Property Let FlipPolyline(Flip As Boolean)
    m_Flip = Flip
End Property

Public Property Get FlipPolyline() As Boolean
    FlipPolyline = m_Flip
End Property

Private Sub Class_Initialize()
    ' Set the defaults
End Sub

Public Sub RunStandardGUI(pApp As IApplication)
    Set frmSMapSettings.m_Application = pApp
    frmSMapSettings.Tickle
    SetWindowLong frmSMapSettings.hwnd, GWL_HWNDPARENT, pApp.hwnd
    frmSMapSettings.Show vbModeless
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
    Set pMx = pApp.Document
    Set pSR = pMx.FocusMap.SpatialReference
    ' If a Projected coord system
    If TypeOf pSR Is IProjectedCoordinateSystem Then
        ' Use meters per unit as the conversion
        Set pPCS = pSR
        dMetersPerUnit = pPCS.CoordinateUnit.MetersPerUnit
        ' Now convert this into page (ie: paper) units
        Set pPage = pMx.PageLayout.Page
        pPageUnits = pPage.Units
        Select Case pPageUnits
            Case esriInches: CalculatePageToMapRatio = dMetersPerUnit / (1 / 12 * 0.304800609601219)
            Case esriFeet: CalculatePageToMapRatio = dMetersPerUnit / (0.304800609601219)
            Case esriCentimeters: CalculatePageToMapRatio = dMetersPerUnit / (1 / 100)
            Case esriMeters: CalculatePageToMapRatio = dMetersPerUnit / (1)
            Case Else:
                MsgBox "Warning: Only the following Page (Layout) Units are supported by this tool:" _
                    & vbCrLf & " - Inches, Feet, Centimeters, Meters" _
                    & vbCrLf & vbCrLf & "Calculating as though Page Units are in Inches..."
                CalculatePageToMapRatio = dMetersPerUnit / (1 / 12 * 0.304800609601219)
        End Select
    ' Otherwise
    Else
        ' If not projected, we can only do a "flat" conversion -> that is, use the current scale and extent
        '  as a ratio to be applied to the map grid scale.
        ' NOTE: We MUST be in Layout mode to make this calculation, as the scale in Map View and Layout View
        '  are not the same (as the extent envelope and data frame envelope can be different shapes).  The
        '  test for being in Layout Mode is made in the clsMapGridButton.ICommand_Enabled property.
        Set pExtentEnv = pMx.ActiveView.Extent
        dStartX = pExtentEnv.XMin
        dStartY = pExtentEnv.YMin
        dEndX = pExtentEnv.XMax
        dEndY = pExtentEnv.YMax
        
        dCurrScale = pMx.FocusMap.MapScale
        If ((dEndX - dStartX) / m_dFrameWidthInPageUnits) > ((dEndY - dStartY) / m_dFrameHeightInPageUnits) Then
            CalculatePageToMapRatio = m_dFrameWidthInPageUnits / ((dEndX - dStartX) / dCurrScale)
        Else
            CalculatePageToMapRatio = m_dFrameHeightInPageUnits / ((dEndY - dStartY) / dCurrScale)
        End If
    End If
    
    Exit Function
eh:
    CalculatePageToMapRatio = 1
    MsgBox "Error in CalculatePageToMapRatio" & vbCrLf & Err.Description
End Function

Private Sub Class_Terminate()
    Set m_DestFL = Nothing
    Set m_DestFC = Nothing
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

    On Error GoTo eh

    ' Set mouse pointer
    Screen.MousePointer = vbArrowHourglass

    ' Init
    Set pMx = Application.Document
    For lLoop = 0 To pMx.FocusMap.LayerCount - 1
        If TypeOf pMx.FocusMap.Layer(lLoop) Is IFeatureLayer Then
            If UCase(pMx.FocusMap.Layer(lLoop).Name) = UCase(m_DestFL.Name) Then
                Set pFL = pMx.FocusMap.Layer(lLoop)
                Exit For
            End If
        End If
    Next
    If pFL Is Nothing Then
        MsgBox "No match in the Map for layer '" & m_DestFL.Name & "'."
        Exit Sub
    End If
    Set pFC = pFL.FeatureClass
    ' Check for required fields - that the field exists
    Dim bErrorWithFields As Boolean
    bErrorWithFields = (pFC.FindField(m_FldStripName) < 0)
    bErrorWithFields = bErrorWithFields Or (pFC.FindField(m_FldNumInSeries) < 0)
    bErrorWithFields = bErrorWithFields Or (pFC.FindField(m_FldMapAngle) < 0)
    If Len(m_FldScale) > 0 Then bErrorWithFields = bErrorWithFields Or (pFC.FindField(m_FldScale) < 0)
    ' If error
    If bErrorWithFields Then
        Err.Raise vbObjectError, "GenerateStripMap", "Could not find all the given field names in " & pFL.Name & "." _
            & vbCrLf & " - " & m_FldStripName & ", " & m_FldNumInSeries & ", " & m_FldMapAngle & ", " & m_FldScale
    End If
    ' Check the field types
    bErrorWithFields = (pFC.Fields.Field(pFC.FindField(m_FldStripName)).Type <> esriFieldTypeString)
    bErrorWithFields = bErrorWithFields Or _
       ((pFC.Fields.Field(pFC.FindField(m_FldNumInSeries)).Type <> esriFieldTypeDouble) And _
        (pFC.Fields.Field(pFC.FindField(m_FldNumInSeries)).Type <> esriFieldTypeInteger) And _
        (pFC.Fields.Field(pFC.FindField(m_FldNumInSeries)).Type <> esriFieldTypeSingle) And _
        (pFC.Fields.Field(pFC.FindField(m_FldNumInSeries)).Type <> esriFieldTypeSmallInteger))
    bErrorWithFields = bErrorWithFields Or _
       ((pFC.Fields.Field(pFC.FindField(m_FldMapAngle)).Type <> esriFieldTypeDouble) And _
        (pFC.Fields.Field(pFC.FindField(m_FldMapAngle)).Type <> esriFieldTypeInteger) And _
        (pFC.Fields.Field(pFC.FindField(m_FldMapAngle)).Type <> esriFieldTypeSingle) And _
        (pFC.Fields.Field(pFC.FindField(m_FldMapAngle)).Type <> esriFieldTypeSmallInteger))
    If Len(m_FldScale) > 0 Then
        bErrorWithFields = bErrorWithFields Or _
           ((pFC.Fields.Field(pFC.FindField(m_FldScale)).Type <> esriFieldTypeDouble) And _
            (pFC.Fields.Field(pFC.FindField(m_FldScale)).Type <> esriFieldTypeInteger) And _
            (pFC.Fields.Field(pFC.FindField(m_FldScale)).Type <> esriFieldTypeSingle) And _
            (pFC.Fields.Field(pFC.FindField(m_FldScale)).Type <> esriFieldTypeSmallInteger))
    End If
    ' if error
    If bErrorWithFields Then
        Err.Raise vbObjectError, "GenerateStripMap", "Given field names are not of the correct type." _
            & vbCrLf & "Strip Map Name field must be a Text field, all others must be numeric fields."
    End If
    ' Get the dataset and workspace (to start editing upon)
    Set pFeatDataset = pFC.FeatureDataset
    If Not pFeatDataset Is Nothing Then
        Set pWorkspaceEdit = pFeatDataset.Workspace
    Else
        ' Is a shapefile, go via just IDataset
        Set pDataset = pFC
        Set pWorkspaceEdit = pDataset.Workspace
    End If
    
'    ' If replacing, delete all existing polygons
    frmProgress.ProgressBar1.Min = 0
    frmProgress.ProgressBar1.Max = 100
    frmProgress.ProgressBar1.Value = 0
    If m_RemoveGrids Then
        Dim pFCu As IFeatureCursor
        Dim pT As ITable
        Set pFCu = m_DestFL.Search(Nothing, False)
        Set pT = m_DestFL.FeatureClass
        If pT.RowCount(Nothing) = 0 Then
            dIncrement = 99
        Else
            dIncrement = 100 / pT.RowCount(Nothing)
        End If
        frmProgress.lblInformation.Caption = "Deleting previous grids..."
        frmProgress.cmdCancel.Visible = False        ' User cannot cancel this step
        frmProgress.Show vbModeless
        Set pFeature = pFCu.NextFeature
        While Not pFeature Is Nothing
            pFeature.Delete
            If (frmProgress.ProgressBar1.Value + dIncrement) <= frmProgress.ProgressBar1.Max Then
                frmProgress.ProgressBar1.Value = frmProgress.ProgressBar1.Value + dIncrement
            Else
                frmProgress.lblInformation.Caption = "Warning: Val > Max (" & frmProgress.ProgressBar1.Max & ")"
            End If
            Set pFeature = pFCu.NextFeature
        Wend
        frmProgress.Hide
    End If
    
        
    ' Init strip map stuff
    Set pPolyline = m_Polyline
    ' Flip, if required
    If m_Flip Then
        pPolyline.ReverseOrientation
    End If
    Set pCenterPoint = pPolyline.FromPoint
    ' Get the progress bar ready
    frmProgress.ProgressBar1.Min = 0
    frmProgress.ProgressBar1.Max = 101
    frmProgress.ProgressBar1.Value = 0
    frmProgress.lblInformation.Caption = "Creating strip map..."
    frmProgress.cmdCancel.Visible = True        ' User cannot cancel this step
    ' Get map units size for grids
    dConvertPageToMapUnits = CalculatePageToMapRatio(Application)
    m_GridWidth = ((m_dMapScale * m_dFrameWidthInPageUnits) / dConvertPageToMapUnits)
    m_GridHeight = ((m_dMapScale * m_dFrameHeightInPageUnits) / dConvertPageToMapUnits)
    ' Init for processing
    dHighestPrev = -1
    bFirstRun = True
    Set pArc = pPolyline
    Set pInsertFeatureCursor = pFC.Insert(True)
    Set pInsertFeatureBuffer = pFC.CreateFeatureBuffer
    frmProgress.Show vbModeless
    Do
        Dim dCircleRadius As Double, colIntersects As Collection, dIntersect As Double
        If bFirstRun Then
            dCircleRadius = m_GridWidth / 2
        Else
            dCircleRadius = m_GridWidth
        End If
        bReducedRadius = False
        Do
            ' Create the search circle
            Set pCircularArc = New CircularArc
            pCircularArc.ConstructCircle pCenterPoint, dCircleRadius, False 'make it clockwise
            Set pCirclePoly = New esriCore.Polygon
            Set pSegmentCollection = pCirclePoly
            pSegmentCollection.AddSegment pCircularArc
            
            ' Intersect the polyline and the circle
            Set pTopoOpt = pPolyline
            Set pGeoCol = New GeometryBag
            Set pGeoCol = pTopoOpt.Intersect(pCirclePoly, esriGeometry0Dimension)
            
            If pGeoCol.GeometryCount = 0 Then
                MsgBox "error - no geoms intersected"
                Exit Sub
            End If
            Set pArc = pPolyline
            lHighestRef = -1
            dHighestThisTurn = 102
            For lLoop2 = 0 To pGeoCol.GeometryCount - 1
                Set pIntersectPoint = pGeoCol.Geometry(lLoop2)
                dIntersect = ReturnPercentageAlong(pArc, pIntersectPoint)
                If dIntersect > (dHighestPrev * 1.001) And dIntersect < dHighestThisTurn Then
                    dHighest = dIntersect
                    dHighestThisTurn = dIntersect
                    lHighestRef = lLoop2
                End If
            Next
            ' If no intersection higher than our previous run, we are at the end.
            If lHighestRef < 0 Then
                dHighest = 101
                ' Need to extend the end (tangent) to get intersection
                Set pIntersectPoint = IntersectPointExtendedTo(pPolyline, pCirclePoly)
                Set pIntersectPointPrev = pCenterPoint
            ' Otherwise, still in the middle somewhere
            Else
                Set pIntersectPoint = pGeoCol.Geometry(lHighestRef)
                ' If just starting off (ie: first grid)
                If bFirstRun Then
                    ' Set the grid so the polyline's starting point is in the
                    '  center of the first grid polygon we make
                    Set pIntersectPointPrev = New Point
                    pIntersectPointPrev.PutCoords pCenterPoint.X - (pIntersectPoint.X - pCenterPoint.X), _
                                                  pCenterPoint.Y - (pIntersectPoint.Y - pCenterPoint.Y)
                ' Otherwise, we already have a previous point
                Else
                    ' So use it
                    Set pIntersectPointPrev = pCenterPoint
                End If
            End If
            ' Make our grid polygon, allowing for any 'shrunken' grids
            If bReducedRadius Then
                Dim pTmpPLine As IPolyline
                Dim pTmpCPoly As IPolygon
                Dim pTmpIntPoint As IPoint
                Set pCircularArc = New CircularArc
                If bFirstRun Then
                    pCircularArc.ConstructCircle pCenterPoint, m_GridWidth / 2, False 'make it clockwise
                Else
                    pCircularArc.ConstructCircle pCenterPoint, m_GridWidth, False 'make it clockwise
                End If
                Set pTmpCPoly = New esriCore.Polygon
                Set pSegmentCollection = pTmpCPoly
                pSegmentCollection.AddSegment pCircularArc
                
                Set pTmpPLine = New Polyline
                pTmpPLine.FromPoint = pIntersectPointPrev
                pTmpPLine.ToPoint = pIntersectPoint
                Set pTmpIntPoint = IntersectPointExtendedTo(pTmpPLine, pTmpCPoly)
                CreateAngledGridPolygon pIntersectPointPrev, pTmpIntPoint, pGridPoly, dGridAngle
            Else
                CreateAngledGridPolygon pIntersectPointPrev, pIntersectPoint, pGridPoly, dGridAngle
            End If
            ' Now, we potentially need to reprocess if the route dips out of our grid
            Set pTopoOpt = pGridPoly
            Set pGeoCol = New GeometryBag
            Set pGeoCol = pTopoOpt.Intersect(pPolyline, esriGeometry0Dimension)
            bContinue = True
            If pGeoCol.GeometryCount > 2 Then
                Set colIntersects = New Collection
                For lLoop2 = 0 To pGeoCol.GeometryCount - 1
                    colIntersects.Add ReturnPercentageAlong(pArc, pGeoCol.Geometry(lLoop2))
                Next
                For lLoop2 = 1 To colIntersects.count
                    If colIntersects.Item(lLoop2) > (dHighestPrev * 1.001) And colIntersects.Item(lLoop2) < (dHighest * 0.999) Then
                        bContinue = False
                        dHighest = dHighestPrev
                        dCircleRadius = dCircleRadius - (m_GridWidth * 0.1)
                        bReducedRadius = True
                        If dCircleRadius <= 0 Then
                            bContinue = True
                        End If
                        Exit For
                    End If
                Next
            End If
            ' If all OK and a reduced radius, look for a quick jump ahead
            If bContinue And bReducedRadius Then 'And pGeoCol.GeometryCount <= 2 Then
                Dim dTmpHighest As Double
                Set pArc = pPolyline
                lHighestRef = -1
                dTmpHighest = -1
                For lLoop2 = 0 To pGeoCol.GeometryCount - 1
                    Set pIntersectPoint = pGeoCol.Geometry(lLoop2)
                    dIntersect = ReturnPercentageAlong(pArc, pIntersectPoint)
                    If dIntersect > dTmpHighest Then
                        dTmpHighest = dIntersect
                        lHighestRef = lLoop2
                    End If
                Next
                If lHighestRef >= 0 Then Set pIntersectPoint = pGeoCol.Geometry(lHighestRef)
                dHighest = dTmpHighest
            End If
        Loop Until bContinue
        
        bFirstRun = False
        dHighestPrev = dHighest
        
        ' All OK to create our grid feature now (hopefully, anyway)
        lCounter = lCounter + 1
        'CreateGridFeaturesAsGraphics pGridPoly, lCounter, dGridAngle, Application 'AAA
        
        ' Create new grid cell feature
        Set pInsertFeatureBuffer.Shape = pGridPoly
        pInsertFeatureBuffer.Value(pFC.Fields.FindField(m_FldStripName)) = m_StripMapName & CStr(lCounter)
        pInsertFeatureBuffer.Value(pFC.Fields.FindField(m_FldNumInSeries)) = lCounter
        pInsertFeatureBuffer.Value(pFC.Fields.FindField(m_FldMapAngle)) = dGridAngle 'degrees
        If Len(m_FldScale) > 0 Then pInsertFeatureBuffer.Value(pFC.Fields.FindField(m_FldScale)) = m_dMapScale
        pInsertFeatureCursor.InsertFeature pInsertFeatureBuffer
        If dHighest <= frmProgress.ProgressBar1.Max Then
            frmProgress.ProgressBar1.Value = dHighest
        Else
            frmProgress.lblInformation.Caption = "Warning: Val > Max (" & frmProgress.ProgressBar1.Max & ")"
            frmProgress.ProgressBar1.Value = frmProgress.ProgressBar1.Max
        End If
        If (lCounter Mod 20 = 0) Then
            DoEvents
            pInsertFeatureCursor.Flush
        End If
        frmProgress.Refresh
        If frmProgress.Cancelled Then
            Dim vUserChoice
            frmProgress.Cancelled = False       ' Reset the form
            vUserChoice = MsgBox("Operation cancelled." _
                & "  Save the edits made thus far?" & vbCrLf & vbCrLf _
                & "(Click Cancel to continue processing)", _
                            vbYesNoCancel, "Generate Strip Map")
            If vUserChoice <> vbCancel Then
                GoTo CancelledGenerateGrids     'Sorry for GoTo usage - in a hurry
            End If
        End If
        ' For next time
        Set pCenterPoint = pIntersectPoint
    Loop While dHighest < 100
    ' Add remainder polys
    pInsertFeatureCursor.Flush
    frmProgress.Hide

    ' Stop editing
    pWorkspaceEdit.StopEditOperation
    pWorkspaceEdit.StopEditing True
    pMx.ActiveView.Refresh
    
    Exit Sub
    
CancelledGenerateGrids:
    frmProgress.Hide
    If vUserChoice = vbYes Then
        pInsertFeatureCursor.Flush
        pWorkspaceEdit.StopEditOperation
        pWorkspaceEdit.StopEditing True
    Else
        pWorkspaceEdit.StopEditOperation
        pWorkspaceEdit.StopEditing False
    End If
    Screen.MousePointer = vbDefault
    pMx.ActiveView.Refresh
    Exit Sub
    
    Resume
eh:
    MsgBox "Error in GenerateStripMap:" & vbCrLf & Err.Description
End Sub

Private Sub CreateGridFeaturesAsGraphics(pGridPolygon As IPolygon, lIndex As Long, dAngle As Double, pApp As IApplication)
    Dim pPntColl As IPointCollection
    Dim pArea As IArea
    Dim pCentroid As IPoint
    Dim lLoop As Long
    
    ' Create graphics (TEST PHASE)
    Set pPntColl = pGridPolygon
    For lLoop = 0 To pPntColl.PointCount - 2
        Perm_DrawLineFromPoints pPntColl.Point(lLoop), pPntColl.Point(lLoop + 1), pApp
    Next
    Perm_DrawLineFromPoints pPntColl.Point(0), pPntColl.Point(pPntColl.PointCount - 1), pApp
    Set pArea = pGridPolygon
    Set pCentroid = pArea.Centroid
    Perm_DrawTextFromPoint pCentroid, CStr(lIndex), pApp, , , , , 8
    pCentroid.Y = pCentroid.Y - (m_GridWidth / 3)
    Perm_DrawTextFromPoint pCentroid, Format(dAngle / cPI * 180, "(#0.0)"), pApp, , , , , 8

End Sub

Private Function ReturnPercentageAlong(ByVal pArc As ICurve, ByVal pPoint As IPoint) As Double
    Dim GeoCount As Long
    Dim pDistAlong As Double
    Dim pDist As Double
    Dim pRightSide As Boolean
    Dim pOutPt As IPoint
    Dim CompareDist As Double
    
    On Error GoTo ErrorHandler
    
    CompareDist = 0
    'Find the distance along curve
    Set pOutPt = New esriCore.Point
    pArc.QueryPointAndDistance esriNoExtension, pPoint, True, pOutPt, _
                        pDistAlong, pDist, pRightSide
    
    ReturnPercentageAlong = (pDistAlong * 100)
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "ReturnPercentageAlong", "Error in ReturnPercentageAlong." _
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
    Set pLine = New esriCore.Line
    pLine.FromPoint = p1
    pLine.ToPoint = p2
    dAngleInRadians = pLine.Angle
    If dAngleInRadians = 0 Then
      ReturnedAngleRadians = 0
    ElseIf dAngleInRadians > 0 Then
      ReturnedAngleRadians = 360 - ((dAngleInRadians / cPI) * 180)
    Else
      ReturnedAngleRadians = Abs((dAngleInRadians / cPI) * 180)
    End If
    Set ReturnedGrid = New Polygon
    Set pPointColl = ReturnedGrid
    ' POINT 1 -------------------------------------------
    Set pPoint = New Point
    pPoint.PutCoords p1.X + (Sin(dAngleInRadians) * (m_GridHeight / 2)), _
                     p1.Y - (Cos(dAngleInRadians) * (m_GridHeight / 2))
    pPointColl.AddPoint pPoint
    Set pPointStart = pPoint
    ' POINT 2 -------------------------------------------
    Set pPoint = New Point
    pPoint.PutCoords p1.X - (Sin(dAngleInRadians) * (m_GridHeight / 2)), _
                     p1.Y + (Cos(dAngleInRadians) * (m_GridHeight / 2))
    pPointColl.AddPoint pPoint
    ' POINT 3 -------------------------------------------
    Set pPoint = New Point
    pPoint.PutCoords p2.X - Sin(dAngleInRadians) * m_GridHeight / 2, _
                     p2.Y + Cos(dAngleInRadians) * m_GridHeight / 2
    pPointColl.AddPoint pPoint
    ' POINT 4 -------------------------------------------
    Set pPoint = New Point
    pPoint.PutCoords p2.X + Sin(dAngleInRadians) * m_GridHeight / 2, _
                     p2.Y - Cos(dAngleInRadians) * m_GridHeight / 2
    pPointColl.AddPoint pPoint
    ' JOIN BACK TO FIRST (CLOSE POLYGON) ----------------
    pPointColl.AddPoint pPointStart
    
    Exit Sub
eh:
    Err.Raise Err.Number, Err.Source, "Error in CreateAngledGridPolygon." _
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
    Set pMx = Application.Document
    Set pGLayer = pMx.FocusMap.BasicGraphicsLayer
    Set pGCon = pGLayer
    Set pElement = New MarkerElement
    pElement.Geometry = pPoint
    Set pMarkerElement = pElement
    
    ' Set the symbol
    Set pColor = New RgbColor
    pColor.Red = dRed
    pColor.Green = dGreen
    pColor.Blue = dBlue
    Set pMarker = New SimpleMarkerSymbol
    With pMarker
        .Color = pColor
        .Size = dSize
    End With
    pMarkerElement.Symbol = pMarker
    
    ' Add the graphic
    Set pElementProp = pElement
    pElementProp.Name = sElementName
    pGCon.AddElement pElement, 0
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
    Set pMx = Application.Document
    Set pGLayer = pMx.FocusMap.BasicGraphicsLayer
    Set pGCon = pGLayer
    Set pElement = New LineElement
    
    ' Set the line symbol
    Set pLnSym = New SimpleLineSymbol
    Set myColor = New RgbColor
    myColor.Red = dRed
    myColor.Green = dGreen
    myColor.Blue = dBlue
    pLnSym.Color = myColor
    pLnSym.Width = dSize
    
    ' Create a standard polyline (via 2 points)
    Set pLine1 = New esriCore.Line
    pLine1.PutCoords pFromPoint, pToPoint
    Set pSeg1 = pLine1
    Set pPolyline = New Polyline
    pPolyline.AddSegment pSeg1
    pElement.Geometry = pPolyline
    Set pLineElement = pElement
    pLineElement.Symbol = pLnSym
    
    ' Add the graphic
    Set pElementProp = pElement
    pElementProp.Name = sElementName
    pGCon.AddElement pElement, 0
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
    Set pMx = Application.Document
    Set pGLayer = pMx.FocusMap.BasicGraphicsLayer
    Set pGCon = pGLayer
    Set pElement = New TextElement
    pElement.Geometry = pPoint
    Set pTextElement = pElement
    
    ' Create the text symbol
    Set myTxtSym = New TextSymbol
    Set myColor = New RgbColor
    myColor.Red = dRed
    myColor.Green = dGreen
    myColor.Blue = dBlue
    myTxtSym.Color = myColor
    myTxtSym.Size = dSize
    myTxtSym.HorizontalAlignment = esriTHACenter
    pTextElement.Symbol = myTxtSym
    pTextElement.Text = sText
    
    ' Add the graphic
    Set pElementProp = pElement
    pElementProp.Name = sElementName
    pGCon.AddElement pElement, 0
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
    pMxDoc.DelayUpdateContents = True
    Set pGLayer = pMxDoc.FocusMap.BasicGraphicsLayer
    Set pGCon = pGLayer
    pGCon.Next
    
    ' Delete all the graphic elements that we created (identify by the name prefix)
    pGCon.Reset
    Set pElement = pGCon.Next
    While Not pElement Is Nothing
        If TypeOf pElement Is IElement Then
            Set pElementProp = pElement
            If (Left(pElementProp.Name, Len(sPrefix)) = sPrefix) Then
                pGCon.DeleteElement pElement
            End If
        End If
        Set pElement = pGCon.Next
    Wend
    
    ' Switch ON the updating of the TOC, refresh
    pMxDoc.DelayUpdateContents = False
    pMxDoc.ActiveView.Refresh
    
    Exit Sub
ErrorHandler:
    MsgBox "Error in RemoveGraphicsByName: " & Err.Description, , "RemoveGraphicsByName"
End Sub

Private Function IntersectPointExtendedTo(pPolyline As IPolyline, pCirclePoly As IPolygon) As IPoint
    Dim pCurve As ICurve
    Dim pLine As ILine
    Dim pPLine As IPolyline
    Dim pTopoOpt As ITopologicalOperator
    Dim pGeoCol As IGeometryCollection
    ' Need to extend the end (creates an ILine object)
    Set pCurve = pPolyline
    Set pLine = New esriCore.Line
    pCurve.QueryTangent esriExtendTangentAtTo, 1, True, _
                        CDbl(m_GridWidth) * 1.1, pLine
    ' Convert ILine to an IPolyline
    Set pPLine = New Polyline
    pPLine.FromPoint = pLine.FromPoint
    pPLine.ToPoint = pLine.ToPoint
    ' Intersect the polyline with the circle
    Set pTopoOpt = pPLine
    Set pGeoCol = New GeometryBag
    Set pGeoCol = pTopoOpt.Intersect(pCirclePoly, esriGeometry0Dimension)
    Set IntersectPointExtendedTo = pGeoCol.Geometry(0)
End Function

Private Function AddPathToPolyLine(pPolyline As IPolyline, pPath As IPath) As IPolyline
  Dim pGCol As IGeometryCollection
  Dim pGeom As IGeometry
    
  If pPolyline Is Nothing Then
    Set pPolyline = New Polyline
  End If
  Set pGCol = pPolyline
  Set pGeom = pPath
  pGCol.AddGeometry pGeom
  Set AddPathToPolyLine = pGCol

End Function


