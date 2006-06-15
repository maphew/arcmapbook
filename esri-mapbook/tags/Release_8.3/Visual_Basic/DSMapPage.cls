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

Private Sub Class_Initialize()
  Set m_PageItemColl = New Collection
End Sub

Private Sub Class_Terminate()
  Set m_PageItemColl = Nothing
End Sub

Private Sub IDSMapPage_AddPageItem(ByVal PageItem As esriCore.IElement)
  m_PageItemColl.Add PageItem
End Sub

Private Sub IDSMapPage_DrawPage(pDoc As IMxDocument, pDSMapSeries As IDSMapSeries, bRefreshFlag As Boolean)
On Error GoTo ErrHand:
  Dim pMap As IMap, lLoop As Long, pEnv As IEnvelope, lIndex As Long
  Dim pFeatLayer As IFeatureLayer, pActive As IActiveView, pTempLayer As ILayer
  Dim pGraphicsCont As IGraphicsContainer, pTempColl As Collection, pElemProps As IElementProperties
  Dim pTextSym As ISimpleTextSymbol, pClone As IClone, pSeriesOpts2 As IDSMapSeriesOptions2
  Dim pSeriesProps As IDSMapSeriesProps, pSeriesOpts As IDSMapSeriesOptions
  Set pSeriesProps = pDSMapSeries
  
  'Find the data frame
  Set pMap = FindDataFrame(pDoc, pSeriesProps.DataFrameName)
  If pMap Is Nothing Then
    MsgBox "Could not find map in DrawPage routine!!!"
    Exit Sub
  End If
  
  'Find the Index layer
  For lLoop = 0 To pMap.LayerCount - 1
    If TypeOf pMap.Layer(lLoop) Is ICompositeLayer Then
      Set pFeatLayer = CompositeLayer1(pMap.Layer(lLoop), pSeriesProps.IndexLayerName)
      If Not pFeatLayer Is Nothing Then Exit For
    Else
      If pMap.Layer(lLoop).Name = pSeriesProps.IndexLayerName Then
        Set pFeatLayer = pMap.Layer(lLoop)
        Exit For
      End If
    End If
  Next lLoop
  If pFeatLayer Is Nothing Then
    MsgBox "Could not find index layer in DrawPage routine!!!"
    Exit Sub
  End If
  
  'Switch to the Layout view if we are not already there
  If Not TypeOf pDoc.ActiveView Is IPageLayout Then
    Set pDoc.ActiveView = pDoc.PageLayout
  End If
  
  'Remove any previous neighbor labels.
  Set pGraphicsCont = pDoc.ActiveView
  pGraphicsCont.Reset
  Set pTempColl = New Collection
  Set pElemProps = pGraphicsCont.Next
  Do While Not pElemProps Is Nothing
    If pElemProps.Name = "DSMAPBOOK TEXT" Then
      pTempColl.Add pElemProps
    End If
    Set pElemProps = pGraphicsCont.Next
  Loop
  For lLoop = 1 To pTempColl.Count
    pGraphicsCont.DeleteElement pTempColl.Item(lLoop)
  Next lLoop
  Set pTempColl = Nothing
  
  'Rotate the frame if necessary
  Set pActive = pMap
  Set pSeriesOpts = pSeriesProps
  Set pSeriesOpts2 = pSeriesOpts
  If pSeriesOpts.RotateFrame Then
'    If m_dPageRotation > 0 Then
      pActive.ScreenDisplay.DisplayTransformation.Rotation = m_dPageRotation
'    End If
  End If
  
  'Set the extent and possibly the scale for the map
  SetMapExtent pSeriesOpts, pActive
    
  'Set the clip property
  'Updated 6/18/03 to support cross hatching of area outside the clip
  Select Case pSeriesOpts2.ClipData
  Case 0   'No clipping
'    pMap.ClipGeometry = Nothing
  Case 1   'Clipping only
    pMap.ClipGeometry = m_pPageShape
  Case 2   'clipping with cross hatching of area outside the clip
'    pMap.ClipGeometry = Nothing
    CreateClipElement pDoc, pActive, pFeatLayer
  End Select
  
  'Check for indicator maps and update those also
  RefreshIndicators pDoc, pSeriesProps, bRefreshFlag
  
  'Check for Date and Title elements
  UpdateTaggedElements pDoc, m_sPageName, bRefreshFlag, pDSMapSeries
  
  'Label neighboring tiles if necessary
  If pSeriesOpts.LabelNeighbors Then
    Set pClone = pSeriesOpts.LabelSymbol
    Set pTextSym = pClone.Clone
    lIndex = pFeatLayer.FeatureClass.FindField(pSeriesProps.IndexFieldName)
    If lIndex >= 0 Then
      LabelNeighbors pDoc, pFeatLayer, pTextSym, lIndex, pSeriesProps.DataFrameName
    End If
  End If
  
  If bRefreshFlag Then
    pDoc.ActiveView.Refresh
  End If

  Exit Sub
ErrHand:
  MsgBox "IDSMapPage_DrawPage - " & Err.Description
End Sub

Private Function IDSMapPage_IndexPage(pIndexLayer As esriCore.IFeatureLayer, sFieldName As String) As Collection
On Error GoTo ErrHand:
  Dim pFilter As ISpatialFilter, pIndex As Collection, lFieldIndex As Long
  Dim pCursor As IFeatureCursor, pFeat As IFeature, sValue As String, lLoop As Long
  
  'Check for a valid index layer
  Set IDSMapPage_IndexPage = Nothing
  If pIndexLayer Is Nothing Then
    MsgBox "You did not send a valid index layer to the IndexPage function!!"
    Exit Function
  End If
  
  'Check for a valid field name
  If sFieldName = "" Or sFieldName = " " Then
    MsgBox "You did not send a valid field name to the IndexPage function!!"
    Exit Function
  End If
  
  'Check for field name in the layer
  lFieldIndex = pIndexLayer.FeatureClass.FindField(sFieldName)
  If lFieldIndex < 0 Then
    MsgBox "Could not find field name in the index layer of the IndexPage function!!"
    Exit Function
  End If
  
  'Perform the query of the index layer using the page shape
  Set pFilter = New SpatialFilter
  pFilter.AddField sFieldName
  pFilter.WhereClause = sFieldName & " <> '' and " & sFieldName & " <> ' '"
  Set pFilter.Geometry = m_pPageShape
  pFilter.GeometryField = pIndexLayer.FeatureClass.ShapeFieldName
  pFilter.SpatialRel = esriSpatialRelIntersects
  Set pCursor = pIndexLayer.Search(pFilter, True)
  
  'Populate the collection with the results of the query
  Set pIndex = New Collection
  Set pFeat = pCursor.NextFeature
  Do While Not pFeat Is Nothing
    sValue = pFeat.Value(lFieldIndex)
    If pIndex.Count > 0 Then
      For lLoop = 1 To pIndex.Count
        If sValue < pIndex.Item(lLoop) Then
          pIndex.Add sValue, sValue, lLoop
          Exit For
        ElseIf sValue = pIndex.Item(lLoop) Then
          Exit For
        End If
        If lLoop = pIndex.Count Then
          pIndex.Add sValue, sValue
        End If
      Next lLoop
    Else
      pIndex.Add sValue, sValue
    End If
    
    Set pFeat = pCursor.NextFeature
  Loop
  
  Set IDSMapPage_IndexPage = pIndex

  Exit Function
  
ErrHand:
  MsgBox "IDSMapPage_IndexPage - " & Err.Description
End Function

Private Property Let IDSMapPage_LastOutputted(RHS As Date)
 m_dLastOutputted = RHS
End Property

Private Property Get IDSMapPage_LastOutputted() As Date
  IDSMapPage_LastOutputted = m_dLastOutputted
End Property

Private Property Get IDSMapPage_PageItem(Index As Long) As esriCore.IElement
  If Index > -1 And Index < m_PageItemColl.Count Then
    Set IDSMapPage_PageItem = m_PageItemColl.Item(Index + 1)
  Else
    Set IDSMapPage_PageItem = Nothing
  End If
End Property

Private Property Get IDSMapPage_PageItemCount() As Long
  IDSMapPage_PageItemCount = m_PageItemColl.Count
End Property

Private Property Let IDSMapPage_PageName(RHS As String)
  m_sPageName = RHS
End Property

Private Property Get IDSMapPage_PageName() As String
  IDSMapPage_PageName = m_sPageName
End Property

Private Property Let IDSMapPage_EnablePage(RHS As Boolean)
  m_bPrintPage = RHS
End Property

Private Property Get IDSMapPage_EnablePage() As Boolean
  IDSMapPage_EnablePage = m_bPrintPage
End Property

Private Property Let IDSMapPage_PageRotation(RHS As Double)
  m_dPageRotation = RHS
End Property

Private Property Get IDSMapPage_PageRotation() As Double
  IDSMapPage_PageRotation = m_dPageRotation
End Property

Private Property Let IDSMapPage_PageScale(RHS As Double)
  m_dPageScale = RHS
End Property

Private Property Get IDSMapPage_PageScale() As Double
  IDSMapPage_PageScale = m_dPageScale
End Property

Private Property Set IDSMapPage_PageShape(RHS As esriCore.IPolygon)
  Set m_pPageShape = RHS
End Property

Private Property Get IDSMapPage_PageShape() As esriCore.IPolygon
  Set IDSMapPage_PageShape = m_pPageShape
End Property

Private Sub IDSMapPage_RemovePageItem(Index As Long)
  If Index > -1 And Index < m_PageItemColl.Count Then
    m_PageItemColl.Remove Index + 1
  End If
End Sub

Private Property Get IPersistVariant_ID() As esriCore.IUID
  Dim id As New UID
  id = "DSMapBookPrj.DSMapPage"
  Set IPersistVariant_ID = id
End Property

Private Sub IPersistVariant_Load(ByVal Stream As esriCore.IVariantStream)
'Load the persisted parameters of the renderer
On Error GoTo ErrHand:
  Dim lLoop As Long, lCount As Long, pElem As IElement
    
  m_sPageName = Stream.Read
  m_bPrintPage = Stream.Read
  m_dPageRotation = Stream.Read
  m_dPageScale = Stream.Read
  m_dLastOutputted = Stream.Read
  Set m_pPageShape = Stream.Read
  lCount = Stream.Read
  If lCount > 0 Then
    For lLoop = 1 To lCount
      Set pElem = Stream.Read
      m_PageItemColl.Add pElem
    Next lLoop
  End If
    
  Exit Sub
ErrHand:
  MsgBox "MapPage - IPersistVariant_Load - " & Err.Description
End Sub

Private Sub IPersistVariant_Save(ByVal Stream As esriCore.IVariantStream)
'Write it all out
On Error GoTo ErrHand:
  Dim lLoop As Long
    
  Stream.Write m_sPageName
  Stream.Write m_bPrintPage
  Stream.Write m_dPageRotation
  Stream.Write m_dPageScale
  Stream.Write m_dLastOutputted
  Stream.Write m_pPageShape
  Stream.Write m_PageItemColl.Count
  If m_PageItemColl.Count > 0 Then
    For lLoop = 1 To m_PageItemColl.Count
      Stream.Write m_PageItemColl.Item(lLoop)
    Next lLoop
  End If
    
  Exit Sub
ErrHand:
  MsgBox "MapPage - IPersistVariant_Save - " & Err.Description
End Sub

Private Sub LabelNeighbors(pDoc As IMxDocument, pFLayer As IFeatureLayer, pTextSym As ISimpleTextSymbol, _
 lIndex As Long, sFrameName As String)
'Routine for loop through the tiles that are touching are selected tile
On Error GoTo ErrHand:
  Dim pElem As IElement, pTextElem As ITextElement, pMap As IMap
  Dim pGraphSel As IGraphicsContainerSelect
  Dim pSpatial As ISpatialFilter, pFeatCursor As IFeatureCursor
  Dim pFeats As IFeature, pActive As IActiveView, sText As String
  Set pMap = pDoc.FocusMap
  Set pActive = pDoc.ActiveView
  
  Set pSpatial = New SpatialFilter
  Set pSpatial.Geometry = m_pPageShape
  pSpatial.GeometryField = pFLayer.FeatureClass.ShapeFieldName
  pSpatial.SpatialRel = esriSpatialRelTouches
  Set pFeatCursor = pFLayer.Search(pSpatial, False)
  Set pFeats = pFeatCursor.NextFeature
  Do While Not pFeats Is Nothing
    'If there is a value for the selected adjacent tile, then get it and label the feature
    If Not IsNull(pFeats.Value(lIndex)) Then
      sText = pFeats.Value(lIndex)
      LabelAdjacent pFeats, pMap, pActive, pTextSym, sText, sFrameName
    End If
    Set pFeats = pFeatCursor.NextFeature
  Loop
  Set pGraphSel = pActive
  pGraphSel.UnselectAllElements

  Exit Sub
ErrHand:
  MsgBox "LabelNeighbors - " & Err.Description
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

  Set pTopoOp = m_pPageShape
  Set pCommonGeom = pTopoOp.Intersect(pLabelFeat.Shape, esriGeometry1Dimension)
  If pCommonGeom.IsEmpty Then
    Set pCommonGeom = pTopoOp.Intersect(pLabelFeat.Shape, esriGeometry0Dimension)
    Set pMulti = pCommonGeom
    Set pMidPt = pMulti.Point(0)
  Else
    Set pPolyLine = pCommonGeom
    Set pMidPt = New Point
    pPolyLine.QueryPoint esriNoExtension, 0.5, True, pMidPt
  End If
    
  'Find center point of map frame
  Set pCenterPt = New esriCore.Point
  Set pMapView = pMap
  Set pEnv = pMapView.Extent
  pCenterPt.X = pEnv.XMin + ((pEnv.XMax - pEnv.XMin) / 2)
  pCenterPt.Y = pEnv.YMin + ((pEnv.YMax - pEnv.YMin) / 2)

  'Get the geometry of the map frame
  Set pGraph = pActive
  pGraph.Reset
  Set pElem = pGraph.Next
  Do While Not pElem Is Nothing
    If TypeOf pElem Is IMapFrame Then
      Set pTmpFrame = pElem
      If pTmpFrame.Map.Name = sFrameName Then
        Set pMapFrame = pElem
        Exit Do
      End If
    End If
    Set pElem = pGraph.Next
  Loop
  If pMapFrame Is Nothing Then Exit Sub
  
  Set pMapEnv = pMapFrame.MapBounds
  Set pFramePoly = pElem.Geometry
  
  'Create curves and intersect them
  Set pPoints = New Polyline
  pPoints.AddPoint pMapEnv.LowerLeft
  pPoints.AddPoint pMapEnv.LowerRight
  pPoints.AddPoint pMapEnv.UpperRight
  pPoints.AddPoint pMapEnv.UpperLeft
  pPoints.AddPoint pMapEnv.LowerLeft
  
  Set pPoints2 = RotatedAndExtendedLine(pCenterPt, pMidPt)
  'If for some reason nothing is returned, go with a rotation of 0
  If pPoints2 Is Nothing Then Exit Sub
  Set pPolyline2 = pPoints2
  If pPolyline2.IsEmpty Then Exit Sub
    
  'Find the intersection point of the line we created and the map extent boundary
  Set pIntTopo = pPoints2
  Set pIntPoints = pIntTopo.Intersect(pPoints, esriGeometry0Dimension)
  If pIntPoints.PointCount = 0 Then Exit Sub
  
  Set pEndPt = pIntPoints.Point(0)
  
  'Extrapolate the point on the extent to a point on the outside of the map frame
  'Figure out which segment we are closest to
  Set pProx = pEndPt
  dDist = 999999
  iSeg = -1
  For lLoop = 0 To 3
    Set pLine = New esriCore.Line
    Select Case lLoop
    Case 0
      pLine.PutCoords pMapEnv.LowerLeft, pMapEnv.UpperLeft
    Case 1
      pLine.PutCoords pMapEnv.UpperLeft, pMapEnv.UpperRight
    Case 2
      pLine.PutCoords pMapEnv.UpperRight, pMapEnv.LowerRight
    Case Else
      pLine.PutCoords pMapEnv.LowerRight, pMapEnv.LowerLeft
    End Select
    
    dTmpDist = pProx.ReturnDistance(pLine)
    If dTmpDist < dDist Then
      dDist = dTmpDist
      iSeg = lLoop
      Set pCurve2 = pLine
    End If
  Next lLoop
  Set pOutPt = New esriCore.Point
  pCurve2.QueryPointAndDistance esriNoExtension, pEndPt, True, pOutPt, dAlong, dFrom, bSide
  
  'We know have the segment and ratio length on that segment, so we can transfer that
  'information to the frame geometry and find the corresponding point there
  Set pPt = New esriCore.Point
  Set pLine = New esriCore.Line
  Select Case iSeg
  Case 0
    pLine.PutCoords pFramePoly.Point(0), pFramePoly.Point(1)
    pTextSym.HorizontalAlignment = esriTHACenter
    pTextSym.VerticalAlignment = esriTVABottom
  Case 1
    pLine.PutCoords pFramePoly.Point(1), pFramePoly.Point(2)
    pTextSym.HorizontalAlignment = esriTHACenter
    pTextSym.VerticalAlignment = esriTVABottom
  Case 2
    pLine.PutCoords pFramePoly.Point(2), pFramePoly.Point(3)
    pTextSym.HorizontalAlignment = esriTHACenter
    pTextSym.VerticalAlignment = esriTVABottom
  Case 3
    pLine.PutCoords pFramePoly.Point(3), pFramePoly.Point(0)
    pTextSym.HorizontalAlignment = esriTHACenter
    pTextSym.VerticalAlignment = esriTVATop
  End Select
  pPt.ConstructAlong pLine, esriNoExtension, dAlong, True
  Set pNewPt = pPt
  
  'Now that we have a point along the data frame, we can place the label based on
  'that point and which side of the frame it is on
  Set pTextElem = New TextElement
  Set pElem = pTextElem
  pTextElem.Symbol = pTextSym
  pElem.Geometry = pNewPt
  Set pElemProps = pElem
  pElemProps.Name = "DSMAPBOOK TEXT"
  pTextElem.Text = sText
  Set pTrans2D = pTextElem
  Select Case iSeg
  Case 0
    dAngle = 90 * (3.14159265358979 / 180)
  Case 1
    dAngle = 0
  Case 2
    dAngle = 270 * (3.14159265358979 / 180)
  Case 3
    dAngle = 0
  End Select
  pTrans2D.Rotate pNewPt, dAngle
  pGraph.AddElement pElem, 0
  
  Exit Sub
ErrHand:
  MsgBox "LabelAdjacent - " & Err.Description
End Sub

Private Function RotatedAndExtendedLine(pCenterPt As IPoint, pMidPt As IPoint) As IPolyline
On Error GoTo ErrHand:
  Dim pPoints As IPointCollection, pPolyLine As IPolyline, pLine As ILine, pNewPt As IConstructPoint
  Dim dOrigAngle As Double, dNewAngle As Double, dLength As Double, dRadAngle As Double
  Dim pNewPoints As IPointCollection, pNewPoint As IConstructPoint, dA As Double
  
  'Create a line so we can get the current angle and distance
  Set pLine = New esriCore.Line
  pLine.PutCoords pCenterPt, pMidPt
  dLength = pLine.Length
  
  If m_dPageRotation = 0 Then
    'Create another point at the same angle to make sure our line crosses the extent boundary
    Set pNewPt = New esriCore.Point
    pNewPt.ConstructAngleDistance pMidPt, pLine.Angle, dLength * 100
    Set pPoints = New Polyline
    pPoints.AddPoint pCenterPt
    pPoints.AddPoint pMidPt
    pPoints.AddPoint pNewPt
    Set RotatedAndExtendedLine = pPoints
    Exit Function
  End If
  
  'If the page is rotated, then we have to rotate the labeling of adjacent tiles also
  dOrigAngle = pLine.Angle * (180 / 3.14159265358979)
  dA = dOrigAngle
  If dOrigAngle < 0 Then
    dOrigAngle = 360 - Abs(dOrigAngle)
  End If
  dNewAngle = dOrigAngle + m_dPageRotation
  If dNewAngle >= 360 Then
    dNewAngle = dNewAngle - 360
  End If
  dRadAngle = dNewAngle * (3.14159265358979 / 180)
  
  'Make a new line at the rotated angle we just calculated.  The new line is made shorter than the original
  'to ensure the line does not extend past the map bounds we need to intersect it with in the next stage
  Set pNewPoint = New esriCore.Point
  Set pNewPoints = New Polyline
  pNewPoint.ConstructAngleDistance pCenterPt, dRadAngle, dLength * 100
  pNewPoints.AddPoint pCenterPt
  pNewPoints.AddPoint pNewPoint
  Set RotatedAndExtendedLine = pNewPoints
  
  Exit Function
ErrHand:
  MsgBox "RotatedLine - " & Err.Description
End Function

Private Sub RefreshIndicators(pDoc As IMxDocument, pSeriesProps As IDSMapSeriesProps, bRefreshFlag As Boolean)
'Routine for updating any identicator maps there might be
On Error GoTo ErrHand:
  Dim pGridLayer As IFeatureLayer, pGridSel As IFeatureSelection
  Dim lLoop As Long, pActive As IActiveView, pSpatial As ISpatialFilter
  Dim pFeature As IFeature, pCursor As IFeatureCursor, pEnv As IEnvelope
  Dim pQuery As IQueryFilter, lLoop2 As Long, pMap As IMap

  'Check for indicator maps and update those also
  For lLoop = 0 To pDoc.Maps.Count - 1
    If pDoc.Maps.Item(lLoop).Name = "Global Indicator" Or pDoc.Maps.Item(lLoop).Name = "Local Indicator" Then
      Set pMap = pDoc.Maps.Item(lLoop)
      'Find the Index layer
      For lLoop2 = 0 To pMap.LayerCount - 1
        If pMap.Layer(lLoop2).Name = "Identifier Layer" Then
          Set pGridLayer = pMap.Layer(lLoop2)
          Exit For
        End If
      Next lLoop2
      If pGridLayer Is Nothing Then
        MsgBox "Could not find layer called Identifier Layer, can not redraw " & pMap.Name & " frame!!!"
        Exit Sub
      End If
      
'      Set pGridLayer = pDoc.Maps.Item(lLoop).Layer(0)
      Set pGridSel = pGridLayer
      Set pQuery = New QueryFilter
      pQuery.WhereClause = pSeriesProps.IndexFieldName & " = '" & m_sPageName & "'"
      pGridSel.Clear
      pGridSel.SelectFeatures pQuery, esriSelectionResultNew, True
        
      If pMap.Name = "Global Indicator" Then
        Set pActive = pDoc.Maps.Item(lLoop)
        If bRefreshFlag Then pActive.Refresh
      ElseIf pMap.Name = "Local Indicator" Then
        Set pSpatial = New SpatialFilter
        Set pSpatial.Geometry = m_pPageShape
        pSpatial.GeometryField = pGridLayer.FeatureClass.ShapeFieldName
        pSpatial.SpatialRel = esriSpatialRelIntersects
        Set pCursor = pGridLayer.Search(pSpatial, False)
        Set pFeature = pCursor.NextFeature
        Do While Not pFeature Is Nothing
          If pEnv Is Nothing Then
            Set pEnv = pFeature.Shape.Envelope
          Else
            pEnv.Union pFeature.Shape.Envelope
          End If
          Set pFeature = pCursor.NextFeature
        Loop
        Set pActive = pMap
        pActive.Extent = pEnv
        If bRefreshFlag Then pActive.Refresh
      End If
      
      Set pGridLayer = Nothing
    End If
  Next lLoop

  Exit Sub
ErrHand:
  MsgBox "RefreshIndicators - " & Err.Description
End Sub

Private Sub UpdateTaggedElements(pDoc As IMxDocument, sTileName As String, bRefreshFlag As Boolean, _
 pDSMapSeries As IDSMapSeries)
'Routine for updating text elements tagged as Date or Title elements
On Error GoTo ErrHand:
  Dim pGraphicsCont As IGraphicsContainer, pElemProps As IElementProperties
  Dim pTextElement As ITextElement, pActive As IActiveView, pElem As IElement
  Dim pEnv As IEnvelope, pEnv2 As IEnvelope, sText As String, bUpdate As Boolean
  Set pGraphicsCont = pDoc.PageLayout
  Set pActive = pGraphicsCont
  pGraphicsCont.Reset
  Set pElemProps = pGraphicsCont.Next
  Do While Not pElemProps Is Nothing
    If TypeOf pElemProps Is ITextElement Then
      bUpdate = True
      Select Case pElemProps.Name
      Case "DSMAPBOOK - DATE"
        sText = Format(Date, "mmm dd, yyyy")
      Case "DSMAPBOOK - TITLE"
        sText = sTileName
      Case "DSMAPBOOK - EXTRAITEM"
        sText = GetExtraItemValue(pDoc, sTileName, pElemProps.Type, pDSMapSeries)
      Case Else
        bUpdate = False
      End Select
      
      If bUpdate Then
        Set pElem = pElemProps
        Set pEnv = New Envelope
        pElem.QueryBounds pActive.ScreenDisplay, pEnv
        Set pTextElement = pElemProps
        pTextElement.Text = sText
        pGraphicsCont.UpdateElement pTextElement
        Set pEnv2 = New Envelope
        pElem.QueryBounds pActive.ScreenDisplay, pEnv2
        pEnv.Union pEnv2
        If bRefreshFlag Then pActive.PartialRefresh esriViewGraphics, Nothing, pEnv
      End If
    End If
    Set pElemProps = pGraphicsCont.Next
  Loop

  Exit Sub
ErrHand:
  MsgBox "UpdateTaggedElements - " & Erl & " - " & Err.Description
End Sub

Private Function GetExtraItemValue(pDoc As IMxDocument, sTileName As String, sFieldName As String, _
 pSeriesProps As IDSMapSeriesProps) As String
On Error GoTo ErrHand:
  Dim pIndexLayer As IFeatureLayer, pQuery As IQueryFilter, pFCursor As IFeatureCursor
  Dim pFeat As IFeature, lIndex As Long, pMap As IMap, lIndex2 As Long
  
  'Find the data frame
  Set pMap = FindDataFrame(pDoc, pSeriesProps.DataFrameName)
  If pMap Is Nothing Then
    MsgBox "Could not find map in GetExtraItem routine!!!"
    GetExtraItemValue = "missing"
    Exit Function
  End If
  
  'Find the Index layer
  Set pIndexLayer = FindLayer(pSeriesProps.IndexLayerName, pMap)
  If pIndexLayer Is Nothing Then
    MsgBox "Could not find index layer (" & pSeriesProps.IndexLayerName & ") in GetExtraItemValue routine!!!"
    GetExtraItemValue = "missing"
    Exit Function
  End If
  
  'Find the field in the index layer
  lIndex = pIndexLayer.FeatureClass.FindField(sFieldName)
  If lIndex < 0 Then
    MsgBox "Could not find the field (" & sFieldName & ") you tagged the item with in the index layer!!!"
    GetExtraItemValue = "missing"
    Exit Function
  End If
  
  'Find the tile name field in the index layer
  lIndex2 = pIndexLayer.FeatureClass.FindField(pSeriesProps.IndexFieldName)
  If lIndex2 < 0 Then
    MsgBox "Could not find tile name field (" & pSeriesProps.IndexFieldName & ") in the index layer!!!"
    GetExtraItemValue = "missing"
    Exit Function
  End If
  
  'Create the query object then select the appropriate tile from the index layer
  Set pQuery = New QueryFilter
  pQuery.WhereClause = pSeriesProps.IndexFieldName & " = '" & sTileName & "'"
  Set pFCursor = pIndexLayer.Search(pQuery, False)
  Set pFeat = pFCursor.NextFeature
  If pFeat Is Nothing Then
    MsgBox "Could not select the tile from the index layer to tag with Extra Item!!!"
    GetExtraItemValue = "missing"
    Exit Function
  End If
  
  'Send back the value of the field
  If IsNull(pFeat.Value(lIndex)) Then
    GetExtraItemValue = " "
  Else
    GetExtraItemValue = pFeat.Value(lIndex)
  End If

  Exit Function
ErrHand:
  MsgBox "GetExtraItemValue - " & Erl & " - " & Err.Description
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
  Set pGraphs = pActive
  pGraphs.Reset
  Set pElemProps = pGraphs.Next
  Do While Not pElemProps Is Nothing
    If TypeOf pElemProps Is IPolygonElement Then
      If UCase(pElemProps.Name) = "DSMAPBOOK CLIP ELEMENT" Then
        pGraphs.DeleteElement pElemProps
        Exit Do
      End If
    End If
    Set pElemProps = pGraphs.Next
  Loop
  
  Set pElem = New PolygonElement
  Set pPoly = m_pPageShape
  Set pNewElem = New PolygonElement
  Set pNewPoly = New Polygon
  pNewPoly.AddPoint pFeatLayer.AreaOfInterest.LowerLeft
  pNewPoly.AddPoint pFeatLayer.AreaOfInterest.UpperLeft
  pNewPoly.AddPoint pFeatLayer.AreaOfInterest.UpperRight
  pNewPoly.AddPoint pFeatLayer.AreaOfInterest.LowerRight
  Set pPoly2 = pNewPoly
  pPoly2.Close
  
  Dim pLineSym As ISimpleLineSymbol, pLineFillSym As ILineFillSymbol
  Dim pFillShape As IFillShapeElement, pColor As IGrayColor
  Set pColor = New GrayColor
  pColor.Level = 150
  Set pLineSym = New SimpleLineSymbol
  pLineSym.Color = pColor
  Set pLineFillSym = New LineFillSymbol
  pLineFillSym.Angle = 45
  pLineFillSym.Color = pColor
  pLineFillSym.Outline = pLineSym
  Set pLineFillSym.LineSymbol = pLineSym
  pLineFillSym.Separation = 5
  
  Set pTopoOp = pPoly2
  Set pFinalGeom = pTopoOp.Difference(pPoly)
  pNewElem.Geometry = pFinalGeom
  Set pFillShape = pNewElem
  pFillShape.Symbol = pLineFillSym
  Set pElemProps = pFillShape
  pElemProps.Name = "DSMapBook Clip Element"
  pGraphs.AddElement pNewElem, 0

  Exit Sub
ErrHand:
  MsgBox "CreateClipElement - " & Erl & " - " & Err.Description
End Sub

Private Sub SetMapExtent(pSeriesOpts As IDSMapSeriesOptions, pActiveView As IActiveView)
On Error GoTo ErrHand:
'Routine for calculating the extent of the tile to be displayed in the layout
  Dim dMult As Double, pEnv As IEnvelope, pMap As IMap
  
  Set pMap = pActiveView
  Select Case pSeriesOpts.ExtentType
  Case 0  'Variable
    If pSeriesOpts.Margin > 0 Then
      Set pEnv = m_pPageShape.Envelope
      Select Case pSeriesOpts.MarginType
      Case 0  'Percent
        dMult = 1 + (pSeriesOpts.Margin / 100)
        pEnv.Expand dMult, dMult, True
      Case 1  'mapunits
        pEnv.Expand pSeriesOpts.Margin, pSeriesOpts.Margin, False
      End Select
      pActiveView.Extent = pEnv
    Else
      pActiveView.Extent = m_pPageShape.Envelope
    End If
  Case 1  'Fixed
    pActiveView.Extent = m_pPageShape.Envelope
    pMap.MapScale = pSeriesOpts.FixedScale
  Case 2  'DataDriven
    pActiveView.Extent = m_pPageShape.Envelope
    pMap.MapScale = m_dPageScale
  End Select

  Exit Sub
ErrHand:
  MsgBox "SetMapExtent - " & Err.Description
End Sub

Public Function FindDataFrame(pDoc As IMxDocument, sFrameName As String) As IMap
On Error GoTo ErrHand:
  Dim lLoop As Long, pMap As IMap
  
  'Find the data frame
  For lLoop = 0 To pDoc.Maps.Count - 1
    If pDoc.Maps.Item(lLoop).Name = sFrameName Then
      Set pMap = pDoc.Maps.Item(lLoop)
      Exit For
    End If
  Next lLoop
  If Not pMap Is Nothing Then
    Set FindDataFrame = pMap
  End If

  Exit Function
ErrHand:
  MsgBox "FindDataFrame - " & Err.Description
End Function

Private Function CompositeLayer1(pCompLayer As ICompositeLayer, sIndexName As String) As IFeatureLayer
On Error GoTo ErrHand:
  Dim lLoop As Long, pFeatLayer As IFeatureLayer
  For lLoop = 0 To pCompLayer.Count - 1
    If TypeOf pCompLayer.Layer(lLoop) Is ICompositeLayer Then
      Set pFeatLayer = CompositeLayer1(pCompLayer.Layer(lLoop), sIndexName)
      If Not pFeatLayer Is Nothing Then
        Set CompositeLayer1 = pFeatLayer
        Exit Function
      End If
    Else
      If pCompLayer.Layer(lLoop).Name = sIndexName Then
        Set CompositeLayer1 = pCompLayer.Layer(lLoop)
        Exit Function
      End If
    End If
  Next lLoop
  
  Set CompositeLayer1 = Nothing

  Exit Function
ErrHand:
  MsgBox "CompositeLayer - " & Err.Description
End Function

Private Function FindLayer(sLayerName As String, pMap As IMap) As IFeatureLayer
' Routine for finding a layer based on a name and then returning that layer as
' a IFeatureLayer
On Error GoTo ErrHand:
  Dim lLoop As Integer
  Dim pFLayer As IFeatureLayer

  For lLoop = 0 To pMap.LayerCount - 1
    If TypeOf pMap.Layer(lLoop) Is ICompositeLayer Then
      Set pFLayer = FindCompositeLayer(pMap.Layer(lLoop), sLayerName, pMap)
      If Not pFLayer Is Nothing Then
        Set FindLayer = pFLayer
        Exit Function
      End If
    ElseIf TypeOf pMap.Layer(lLoop) Is IFeatureLayer Then
      Set pFLayer = pMap.Layer(lLoop)
      If UCase(pFLayer.Name) = UCase(sLayerName) Then
        Set FindLayer = pFLayer
        Exit Function
      End If
    End If
  Next lLoop
  
  Set FindLayer = Nothing
  
  Exit Function
  
ErrHand:
  MsgBox "FindLayer - " & Erl & " - " & Err.Description
End Function

Private Function FindCompositeLayer(pCompLayer As ICompositeLayer, sLayerName As String, pMap As IMap) As IFeatureLayer
On Error GoTo ErrHand:
  Dim lLoop As Long, pFeatLayer As IFeatureLayer
  For lLoop = 0 To pCompLayer.Count - 1
    If TypeOf pCompLayer.Layer(lLoop) Is ICompositeLayer Then
      Set pFeatLayer = FindCompositeLayer(pCompLayer.Layer(lLoop), sLayerName, pMap)
      If Not pFeatLayer Is Nothing Then
        Set FindCompositeLayer = pFeatLayer
        Exit Function
      End If
    Else
      If TypeOf pCompLayer.Layer(lLoop) Is IFeatureLayer Then
        If UCase(pCompLayer.Layer(lLoop).Name) = UCase(sLayerName) Then
          Set FindCompositeLayer = pCompLayer.Layer(lLoop)
          Exit Function
        End If
      End If
    End If
  Next lLoop

  Exit Function
ErrHand:
  MsgBox "CompositeLayer - " & Erl & " - " & Err.Description
End Function
