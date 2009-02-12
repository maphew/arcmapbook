VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NWMapPage"
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

Implements INWDSMapPage
Implements INWMapPageAttribs
Implements IPersistVariant

Dim m_PageItemColl As Collection
Dim m_sPageName As String
Dim m_bPrintPage As Boolean
Dim m_dPageRotation As Double
Dim m_dPageScale As Double
Dim m_dLastOutputted As Date
Dim m_pPageShape As IPolygon
Dim m_lPageNumber As Long

'Added 9/3/04 to support NW Data
Private m_pUID_NWData1 As UID
Private m_pUID_FlexiblePersist As UID
Private m_sAdjacentPageTextSymbol As String
Private m_pInvisibleLayers As Scripting.Dictionary
Private m_sInvisibleLayersGroup As String
'Private m_pDetailInsets As Scripting.Dictionary
'Private m_pDetailInsetArrowShadows As Scripting.Dictionary
'Private m_pDetailInsetArrows As Scripting.Dictionary
'Private m_sPrevPageName As String


Const c_sModuleFileName As String = "NWMapPage.cls"


Private Sub Class_Initialize()
  Set m_PageItemColl = New Collection
  If m_pUID_FlexiblePersist Is Nothing Then
    Set m_pUID_FlexiblePersist = New UID
    m_pUID_FlexiblePersist.Value = "{839079CF-45B2-47ee-8901-988D5E05147B}"
  End If
  If m_pUID_NWData1 Is Nothing Then
    Set m_pUID_NWData1 = New UID
    m_pUID_NWData1.Value = "{AF401E4E-3AA4-4867-AA8E-87FACF67CB71}"
  End If
  Set m_pInvisibleLayers = New Scripting.Dictionary
'  Set m_pDetailInsets = New Scripting.Dictionary
'  Set m_pDetailInsetArrowShadows = New Scripting.Dictionary
'  Set m_pDetailInsetArrows = New Scripting.Dictionary
End Sub

Private Sub Class_Terminate()
  Set m_PageItemColl = Nothing
  Set m_pInvisibleLayers = Nothing
'  Set m_pDetailInsets = Nothing
'  Set m_pDetailInsetArrowShadows = Nothing
'  Set m_pDetailInsetArrows = Nothing
End Sub



'NW data properties
Private Property Get INWMapPageAttribs_AdjacentLabelSymbol() As String
  INWMapPageAttribs_AdjacentLabelSymbol = m_sAdjacentPageTextSymbol
End Property

Private Property Let INWMapPageAttribs_AdjacentLabelSymbol(ByVal RHS As String)
  m_sAdjacentPageTextSymbol = RHS
End Property









Private Sub INWMapPageAttribs_InvisibleLayerAdd(ByVal RHS As String)
  On Error GoTo ErrorHandler
  If Not m_pInvisibleLayers.Exists(RHS) Then
    m_pInvisibleLayers.Add RHS, Nothing
  End If

  Exit Sub
ErrorHandler:
  HandleError True, "INWMapPageAttribs_InvisibleLayerAdd " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Property Get INWMapPageAttribs_InvisibleLayerCount() As Long
  INWMapPageAttribs_InvisibleLayerCount = m_pInvisibleLayers.Count
End Property

Private Sub INWMapPageAttribs_InvisibleLayerRemove(ByVal sLayerName As String)
  If m_pInvisibleLayers.Exists(sLayerName) Then
    m_pInvisibleLayers.Remove (sLayerName)
  End If
End Sub

Private Property Get INWMapPageAttribs_InvisibleLayers() As Variant
  INWMapPageAttribs_InvisibleLayers = m_pInvisibleLayers.Keys
End Property

Private Property Let INWMapPageAttribs_LayerVisibilityGroup(ByVal RHS As String)
  m_sInvisibleLayersGroup = RHS
End Property

Private Property Get INWMapPageAttribs_LayerVisibilityGroup() As String
  INWMapPageAttribs_LayerVisibilityGroup = m_sInvisibleLayersGroup
End Property


'Private Property Let INWMapPageAttribs_PreviousPageName(RHS As String)
'702:   m_sPrevPageName = RHS
'End Property
'
'Private Property Get INWMapPageAttribs_PreviousPageName() As String
'706:   INWMapPageAttribs_PreviousPageName = m_sPrevPageName
'End Property

Private Sub INWDSMapPage_AddPageItem(ByVal PageItem As IElement)
  m_PageItemColl.Add PageItem
End Sub

Private Sub INWDSMapPage_DrawPage(pDoc As IMxDocument, pNWDSMapSeries As INWDSMapSeries, bRefreshFlag As Boolean)
On Error GoTo ErrHand:
  Dim pMap As IMap, lLoop As Long, pEnv As IEnvelope, lIndex As Long
  Dim pFeatLayer As IFeatureLayer, pActive As IActiveView, pTempLayer As ILayer
  Dim pGraphicsCont As IGraphicsContainer, pTempColl As Collection, pElemProps As IElementProperties
  Dim pTextSym As ISimpleTextSymbol, pClone As IClone, pSeriesOpts2 As INWDSMapSeriesOptions2
  Dim pSeriesProps As INWDSMapSeriesProps, pSeriesOpts As INWDSMapSeriesOptions
  Dim pNWSeriesOptions As INWMapSeriesOptions, pLayers As IEnumLayer, pLayer As ILayer
  Dim pLyrGroup As INWLayerVisibilityGroup
  Set pSeriesProps = pNWDSMapSeries
  Set pNWSeriesOptions = pSeriesProps
  
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
  
  'Set visible layers
  Set pLayers = pMap.Layers
  Set pLayer = pLayers.Next
  If m_sInvisibleLayersGroup = "" Then
    Do While Not pLayer Is Nothing
      pLayer.Visible = Not m_pInvisibleLayers.Exists(pLayer.Name)
      Set pLayer = pLayers.Next
    Loop
  Else
    If pNWSeriesOptions.LayerGroupExists(m_sInvisibleLayersGroup) Then
      Set pLyrGroup = pNWSeriesOptions.LayerGroupGet(m_sInvisibleLayersGroup)
      Do While Not pLayer Is Nothing
        pLayer.Visible = Not pLyrGroup.Exists(pLayer.Name)
        Set pLayer = pLayers.Next
      Loop
    End If
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
    If pElemProps.Name = "NWDSMAPBOOK TEXT" Then
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
    pMap.ClipGeometry = Nothing
    CreateClipElement pDoc, pActive, pFeatLayer
  End Select
  
  'Apply the dynamic definition query
  RefreshDynamicDefinitionQueries pDoc, pNWSeriesOptions
  'Set bubble detail insets
  Dim pBubbleLayer As ILayer
  Set pBubbleLayer = GetBubbleLayer(pNWSeriesOptions, pDoc)
  
  If Not pBubbleLayer Is Nothing Then
    pNWSeriesOptions.RefreshDetailInsets _
      pDoc, pBubbleLayer, m_pPageShape
  End If
  
  'find and update any indicator maps
  RefreshIndicators pDoc, pSeriesProps, bRefreshFlag
  
  'control data frame visibility
  pNWSeriesOptions.DataFrameDrawFrames pDoc, m_sPageName, pSeriesProps.DataFrameName
  
  'update the extent of visible data frames
  '(important to call after setting data frame visibility)
  pNWSeriesOptions.DataFrameToUpdateUpdateDataFrames pDoc, m_sPageName
  
  'update extents of other data frames
  pNWSeriesOptions.DataFrameToUpdateUpdateDataFrames pDoc, m_sPageName
  
  'Check for Date and Title elements
  UpdateTaggedElements pDoc, m_sPageName, bRefreshFlag, pNWDSMapSeries
  
  'control visibility of (most) layout elements
  pNWSeriesOptions.ElementsFilterLayoutElements pDoc, m_sPageName
  
  'Label neighboring tiles if necessary
  If pSeriesOpts.LabelNeighbors Then
                                                  'changed for NW Mapbook
                                                  'adjacent map labels
    With pNWSeriesOptions
      If m_sAdjacentPageTextSymbol = "" Then
        Set pClone = .TextSymbol(.TextSymbolDefault)
      Else
        Set pClone = .TextSymbol(m_sAdjacentPageTextSymbol)
      End If
    End With
    '------
    Set pTextSym = pClone.Clone
    lIndex = pFeatLayer.FeatureClass.FindField(pSeriesProps.IndexFieldName)
    If lIndex >= 0 Then
      Set pNWSeriesOptions = pSeriesOpts
      LabelNeighbors pDoc, pFeatLayer, pTextSym, lIndex, pSeriesProps.DataFrameName, pNWSeriesOptions.NeighborLabelIndent
    End If
  End If
  
  
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
  
  Set pUID = New UID
  bFlag = False
On Error GoTo NoDetails:
  pUID.Value = "DetailAreaExt.CreateDetailsCmd"
On Error GoTo ErrHand:
   
   If Not bFlag Then
    Set pDocument = pDoc
    Set pCommandItem = pDocument.CommandBars.Find(pUID)
    If Not pCommandItem Is Nothing Then
      If TypeOf pCommandItem Is ICommand Then
        Set pCommand = pCommandItem
        If pCommand.Enabled Then
          pCommandItem.Execute
        End If
      End If
    End If
  End If
  
  '  End of DetailExtension additions
  '-------------------------------
  
  If bRefreshFlag Then
    pDoc.ActiveView.Refresh
  End If

  Exit Sub
  
NoDetails:
  bFlag = True
  Resume Next
  
ErrHand:
  MsgBox "INWDSMapPage_DrawPage - " & Erl & " - " & Err.Description
End Sub

Private Function INWDSMapPage_IndexPage(pIndexLayer As IFeatureLayer, sFieldName As String) As Collection
On Error GoTo ErrHand:
  Dim pFilter As ISpatialFilter, pIndex As Collection, lFieldIndex As Long
  Dim pCursor As IFeatureCursor, pFeat As IFeature, sValue As String, lLoop As Long
  
  'Check for a valid index layer
  Set INWDSMapPage_IndexPage = Nothing
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
  
  Set INWDSMapPage_IndexPage = pIndex

  Exit Function
  
ErrHand:
  MsgBox "INWDSMapPage_IndexPage - " & Err.Description
End Function

Private Property Let INWDSMapPage_LastOutputted(RHS As Date)
 m_dLastOutputted = RHS
End Property

Private Property Get INWDSMapPage_LastOutputted() As Date
  INWDSMapPage_LastOutputted = m_dLastOutputted
End Property

Private Property Get INWDSMapPage_PageItem(Index As Long) As IElement
  If Index > -1 And Index < m_PageItemColl.Count Then
    Set INWDSMapPage_PageItem = m_PageItemColl.Item(Index + 1)
  Else
    Set INWDSMapPage_PageItem = Nothing
  End If
End Property

Private Property Get INWDSMapPage_PageItemCount() As Long
  INWDSMapPage_PageItemCount = m_PageItemColl.Count
End Property

Private Property Let INWDSMapPage_PageName(RHS As String)
  m_sPageName = RHS
End Property

Private Property Get INWDSMapPage_PageName() As String
  INWDSMapPage_PageName = m_sPageName
End Property

Private Property Let INWDSMapPage_EnablePage(RHS As Boolean)
  m_bPrintPage = RHS
End Property

Private Property Get INWDSMapPage_EnablePage() As Boolean
  INWDSMapPage_EnablePage = m_bPrintPage
End Property

Private Property Let INWDSMapPage_PageNumber(RHS As Long)
  m_lPageNumber = RHS
End Property

Private Property Get INWDSMapPage_PageNumber() As Long
  INWDSMapPage_PageNumber = m_lPageNumber
End Property

Private Property Let INWDSMapPage_PageRotation(RHS As Double)
  m_dPageRotation = RHS
End Property

Private Property Get INWDSMapPage_PageRotation() As Double
  INWDSMapPage_PageRotation = m_dPageRotation
End Property

Private Property Let INWDSMapPage_PageScale(RHS As Double)
  m_dPageScale = RHS
End Property

Private Property Get INWDSMapPage_PageScale() As Double
  INWDSMapPage_PageScale = m_dPageScale
End Property

Private Property Set INWDSMapPage_PageShape(RHS As IPolygon)
  Set m_pPageShape = RHS
End Property

Private Property Get INWDSMapPage_PageShape() As IPolygon
  Set INWDSMapPage_PageShape = m_pPageShape
End Property

Private Sub INWDSMapPage_RemovePageItem(Index As Long)
  If Index > -1 And Index < m_PageItemColl.Count Then
    m_PageItemColl.Remove Index + 1
  End If
End Sub

Private Property Get IPersistVariant_ID() As esriSystem.IUID
  Dim id As New UID
  id = "NWMapBookPrj.NWMapPage"
  Set IPersistVariant_ID = id
End Property

Private Sub IPersistVariant_Load(ByVal Stream As esriSystem.IVariantStream)
'Load the persisted parameters of the renderer
On Error GoTo ErrHand:
  Dim lLoop As Long, lCount As Long, pElem As IElement, sFirstItem As String
  Dim lPropCount As Long
  Dim bFoundUID As Boolean, bExitFlagFound As Boolean, vNWData As Variant
  Dim sUID1 As String, sUID2 As String, lInvisLyrCount As Long
  Dim vLyrItem As Variant, sLyrName As String, vDebug As Variant
  
  
  'Added 2/18/04 to make the list of persisted properties more dynamic
  sFirstItem = Stream.Read
  If UCase(Left(sFirstItem, 18)) = "PAGEPROPERTYCOUNT-" Then
    lPropCount = Mid(sFirstItem, 19) - 1
    m_sPageName = Stream.Read
  Else
    lPropCount = 5
    m_sPageName = sFirstItem
  End If
    
  'Original page properties
  m_bPrintPage = Stream.Read
  m_dPageRotation = Stream.Read
  m_dPageScale = Stream.Read
  m_dLastOutputted = Stream.Read
  Set m_pPageShape = Stream.Read
  
  'Additional properties added after 2/18/04
  If lPropCount > 5 Then    'Checking for page number
    m_lPageNumber = Stream.Read
  Else
    m_lPageNumber = -1
  End If
  
  'More original properties.  Writen out below the new properties because they are of variable length
  lCount = Stream.Read
  If lCount > 0 Then
    For lLoop = 1 To lCount
      Set pElem = Stream.Read
      m_PageItemColl.Add pElem
    Next lLoop
  End If
    
    
  'point of custom load code for NW Mapbook
  'customization properties.
  '{839079CF-45B2-47ee-8901-988D5E05147B}
  'The pattern is to have a nested persistence
  'structure of:
  '{guid 1}
  '  {guid 2}
  '  ... data ...
  '  {guid 2}
  '{guid 1}
  'This opens the door for any future amount of
  'data.  This code will loop through stream.read
  'until the closing GUID is detected, allowing
  'earlier versions of this code to work with
  'future files that have more data stored.
  
  bExitFlagFound = False
  vNWData = Stream.Read
  If Not IsEmpty(vNWData) Then
                                                  'for whatever reason, saved UIDs
                                                  'persist as strings
    If StrComp(UCase(TypeName(vNWData)), "String", vbTextCompare) = 0 Then
      sUID1 = vNWData
                                                  'if the GUID value is correct
      If StrComp(sUID1, m_pUID_FlexiblePersist.Value, vbTextCompare) = 0 Then
                                                  'then read the persisted data values
        bFoundUID = False
        
        '{guid 1}
        '  {guid 2} <--- search for this section
        '  ... data ...
        '  {guid 2}
        '{guid 1}
        Do While (Not bFoundUID) And (Not bExitFlagFound)
          vNWData = Stream.Read
          If StrComp(UCase(TypeName(vNWData)), "String", vbTextCompare) = 0 Then
            sUID2 = vNWData
            If StrComp(sUID2, m_pUID_NWData1.Value, vbTextCompare) = 0 Then
              bFoundUID = True
            ElseIf StrComp(sUID2, m_pUID_FlexiblePersist.Value, vbTextCompare) = 0 Then
              bExitFlagFound = True
            End If
          End If
        Loop
                                                  'once the bracketing UIDs
                                                  'have been detected, read the
                                                  'data for this version of the
                                                  'NW Mapbook application
        If bFoundUID And Not bExitFlagFound Then
          m_sAdjacentPageTextSymbol = Stream.Read
          
          m_pInvisibleLayers.RemoveAll
          lInvisLyrCount = Stream.Read
          For lLoop = 0 To lInvisLyrCount - 1
            sLyrName = Stream.Read
            'vLyrItem = Stream.Read
            m_pInvisibleLayers.Add sLyrName, Nothing
          Next lLoop
          m_sInvisibleLayersGroup = Stream.Read
        End If
        
        
        '{guid 1}
        '  {guid 2}
        '  ... data ...
        '  {guid 2}      <--- search for this section
        '{guid 1}
        bFoundUID = False
        Do While (Not bFoundUID) And (Not bExitFlagFound)
          vNWData = Stream.Read
          If StrComp(UCase(TypeName(vNWData)), "String", vbTextCompare) = 0 Then
            sUID2 = vNWData
            If StrComp(sUID2, m_pUID_NWData1.Value, vbTextCompare) = 0 Then
              bFoundUID = True
            ElseIf StrComp(sUID2, m_pUID_FlexiblePersist.Value, vbTextCompare) = 0 Then
              bExitFlagFound = True
            End If
          End If
        Loop
        
        '{guid 1}
        '  {guid 2}
        '  ... data ...
        '  {guid 2}
        '{guid 1}     <---- search for the closing section
        Do While Not bExitFlagFound
          vNWData = Stream.Read
          If StrComp(UCase(TypeName(vNWData)), "String", vbTextCompare) = 0 Then
            sUID2 = vNWData
            If StrComp(sUID2, m_pUID_FlexiblePersist.Value, vbTextCompare) = 0 Then
              bExitFlagFound = True
            End If
          End If
        Loop
        'The above structure is made so that future data can be saved ...
        '{guid 1}
        '  <-- future data, presumably enclosed within {guid3}..{guid3}, {guid4}.. etc. -->
        '  {guid 2}
        '    ... data ...
        '    <-- future data for the NW Mapbook component since stream reads keep going until {guid2} -->
        '  {guid 2}
        '  <-- future data, presumably enclosed within {guid5}..{guid5}, {guid6}.. etc. -->
        '{guid 1}     <---- search for the closing section
      
      
      End If
    Else
      'not sure if this will work:
      'Will too many stream.reads give me access to
      'whatever is loaded next?  Or will it give me
      'a stream.read error?  or will stream.read
      'return nothing?
      
          'assume that I read something from the next
          'section of load code.
      Stream.Write vNWData
    End If
  End If

''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''

    
  Exit Sub
ErrHand:
  MsgBox "MapPage - IPersistVariant_Load - " & Erl & " - " & Err.Description
End Sub

Private Sub IPersistVariant_Save(ByVal Stream As esriSystem.IVariantStream)
'Write it all out
On Error GoTo ErrHand:
  Dim lLoop As Long, lInvisLyrCount As Long, vKeys() As Variant
  Dim sLyrName As String
  
  'Added 2/18/04 to make the list of persisted properties more dynamic
  Stream.Write "PAGEPROPERTYCOUNT-7"
    
  Stream.Write m_sPageName
  Stream.Write m_bPrintPage
  Stream.Write m_dPageRotation
  Stream.Write m_dPageScale
  Stream.Write m_dLastOutputted
  Stream.Write m_pPageShape
  Stream.Write m_lPageNumber   'Added 2/18/04
  
  Stream.Write m_PageItemColl.Count
  If m_PageItemColl.Count > 0 Then
    For lLoop = 1 To m_PageItemColl.Count
      Stream.Write m_PageItemColl.Item(lLoop)
    Next lLoop
  End If
  
  
  'NW Mapbook point for inserting custom Saves code
  'Added 8/31/04
  If m_pUID_FlexiblePersist Is Nothing Then
    Set m_pUID_FlexiblePersist = New UID
    m_pUID_FlexiblePersist.Value = "{839079CF-45B2-47ee-8901-988D5E05147B}"
  End If
  If m_pUID_NWData1 Is Nothing Then
    Set m_pUID_NWData1 = New UID
    m_pUID_NWData1.Value = "{AF401E4E-3AA4-4867-AA8E-87FACF67CB71}"
  End If
  Stream.Write m_pUID_FlexiblePersist.Value
  Stream.Write m_pUID_NWData1.Value
  Stream.Write m_sAdjacentPageTextSymbol
  lInvisLyrCount = m_pInvisibleLayers.Count
  Stream.Write lInvisLyrCount
  If lInvisLyrCount > 0 Then
    vKeys = m_pInvisibleLayers.Keys
    For lLoop = 0 To (lInvisLyrCount - 1)
      sLyrName = vKeys(lLoop)
      Stream.Write sLyrName
      'Stream.Write m_pInvisibleLayers.Item(sLyrName)
    Next lLoop
  End If
  Stream.Write m_sInvisibleLayersGroup
  Stream.Write m_pUID_NWData1.Value
  Stream.Write m_pUID_FlexiblePersist.Value
  'End of NW Mapbook custom Save code
    
  Exit Sub
ErrHand:
  MsgBox "MapPage - IPersistVariant_Save - " & Err.Description
End Sub

Private Sub LabelNeighbors(pDoc As IMxDocument, pFLayer As IFeatureLayer, pTextSym As ISimpleTextSymbol, _
 lIndex As Long, sFrameName As String, dLabelIndent As Double)
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
      LabelAdjacent pFeats, pMap, pActive, pTextSym, sText, sFrameName, dLabelIndent
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
 sText As String, sFrameName As String, dLabelInset As Double)
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
    Set pMidPt = New esriGeometry.Point
    pPolyLine.QueryPoint esriNoExtension, 0.5, True, pMidPt
  End If
    
  'Find center point of map frame
  Set pCenterPt = New esriGeometry.Point
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
    Set pLine = New esriGeometry.Line
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
  Set pOutPt = New esriGeometry.Point
  pCurve2.QueryPointAndDistance esriNoExtension, pEndPt, True, pOutPt, dAlong, dFrom, bSide
  
  'We know have the segment and ratio length on that segment, so we can transfer that
  'information to the frame geometry and find the corresponding point there
  Set pPt = New esriGeometry.Point
  Set pLine = New esriGeometry.Line
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
  
  'move the point back from the data frame border
  Dim pLabelPt As IPoint
  Set pLabelPt = pPt
  Select Case iSeg
  Case 0   'left page edge
    pLabelPt.X = pLabelPt.X + dLabelInset
                                                  'see if Y is within label offset distance
                                                  'of top or bottom border
    If (pLabelPt.Y - pFramePoly.Point(0).Y) < dLabelInset Then
                                                  'enforce minimum dist from bottom
      pLabelPt.Y = pFramePoly.Point(0).Y + dLabelInset
    End If
    If (pFramePoly.Point(1).Y - pLabelPt.Y) < dLabelInset Then
                                                  'enforce minimum dist from top
      pLabelPt.Y = pFramePoly.Point(1).Y - dLabelInset
    End If
  
  Case 1   'top page edge
    pLabelPt.Y = pLabelPt.Y - dLabelInset
                                                  'see if X is within label offset distance
                                                  'of left or right border
    If (pLabelPt.X - pFramePoly.Point(1).X) < dLabelInset Then
                                                  'enforce minimum dist from left
      pLabelPt.X = pFramePoly.Point(1).X + dLabelInset
    End If
    If (pFramePoly.Point(2).X - pLabelPt.X) < dLabelInset Then
                                                  'enforce minimum dist from right
      pLabelPt.X = pFramePoly.Point(2).X - dLabelInset
    End If
    
  Case 2   'right page edge
    pLabelPt.X = pLabelPt.X - dLabelInset
                                                  'see if Y is within label offset distance
                                                  'of top or bottom border
    If (pFramePoly.Point(2).Y - pLabelPt.Y) < dLabelInset Then
                                                  'enforce minimum dist from bottom
      pLabelPt.Y = pFramePoly.Point(2).Y - dLabelInset
    End If
    If (pLabelPt.Y - pFramePoly.Point(3).Y) < dLabelInset Then
                                                  'enforce minimum dist from bottom
      pLabelPt.Y = pFramePoly.Point(3).Y + dLabelInset
    End If
    
  Case 3   'bottom page edge
    pLabelPt.Y = pLabelPt.Y + dLabelInset
    
                                                  'see if X is within label offset distance
                                                  'of left or right border
    If (pFramePoly.Point(3).X - pLabelPt.X) < dLabelInset Then
                                                  'enforce minimum dist from right
      pLabelPt.X = pFramePoly.Point(3).X - dLabelInset
    End If
    If (pLabelPt.X - pFramePoly.Point(0).X) < dLabelInset Then
                                                  'enforce minimum dist from left
      pLabelPt.X = pFramePoly.Point(0).X + dLabelInset
    End If
    
  End Select
  
  Set pNewPt = pPt
  
  'Now that we have a point along the data frame, we can place the label based on
  'that point and which side of the frame it is on
  Set pTextElem = New TextElement
  Set pElem = pTextElem
  pTextElem.Symbol = pTextSym
  pElem.Geometry = pNewPt
  Set pElemProps = pElem
  pElemProps.Name = "NWDSMAPBOOK TEXT"
  pTextElem.Text = sText
  Set pTrans2D = pTextElem
  Select Case iSeg
  Case 0
    'dAngle = 90 * (3.14159265358979 / 180)
                                                  'NW Mapbook maps require
                                                  'right reading adjacent labels
                                                  'on the left side of maps ...
    dAngle = 0
  Case 1
    dAngle = 0
  Case 2
                                                  'and on the right side of maps
    'dAngle = 270 * (3.14159265358979 / 180)
    dAngle = 0
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
  Set pLine = New esriGeometry.Line
  pLine.PutCoords pCenterPt, pMidPt
  dLength = pLine.Length
  
  If m_dPageRotation = 0 Then
    'Create another point at the same angle to make sure our line crosses the extent boundary
    Set pNewPt = New esriGeometry.Point
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
  
  'Make a new esrigeometry.line at the rotated angle we just calculated.  The new esrigeometry.line is made shorter than the original
  'to ensure the line does not extend past the map bounds we need to intersect it with in the next stage
  Set pNewPoint = New esriGeometry.Point
  Set pNewPoints = New Polyline
  pNewPoint.ConstructAngleDistance pCenterPt, dRadAngle, dLength * 100
  pNewPoints.AddPoint pCenterPt
  pNewPoints.AddPoint pNewPoint
  Set RotatedAndExtendedLine = pNewPoints
  
  Exit Function
ErrHand:
  MsgBox "RotatedLine - " & Err.Description
End Function





'RefreshDetailInsets
'  Package all the code per map page to work with detail insets into one subroutine
'  in order to keep the DrawPage routine simplified.
'----------------------------------------------------
'Private Sub RefreshDetailInsets(pMxDoc As IMxDocument, _
'                                pNWSeriesOpts As INWMapSeriesOptions)
''                                pDictArrowShadows As Scripting.Dictionary, _
''                                pDictArrows As Scripting.Dictionary, _
'                                pDictDetailInsets As Scripting.Dictionary)
'  On Error GoTo ErrorHandler
'
'  Dim sDataFrameName As String, pDataFrameElement As IElement, lArrowCount As Long
'  Dim lArrowShadowCount As Long, lDetailInsetCount As Long, i As Long
'  Dim pGraphicsContainer As IGraphicsContainer, pElement As IElement, sInsetID As String
'  Dim pLayers As IEnumLayer, pLayer As ILayer, vDetailInsetKeys() As Variant
'  Dim bLyrIsFound As Boolean, pInsetLyr As ILayer, pFeatureLyr As IFeatureLayer
'  Dim pFeatureClass As IFeatureClass, pSpatialFilter As ISpatialFilter
'  Dim pFeatCursor As IFeatureCursor, pBubbleFeature As IFeature
'  Dim pRow As esriGeoDatabase.IRow
'
'  If pMxDoc Is Nothing Then Exit Sub
'
'  sDataFrameName = GetActiveDataFrameName(pMxDoc)
'  Set pDataFrameElement = GetDataFrameElement(sDataFrameName, pMxDoc)
'
'  '-----------------------------
'  'remove all previous graphics,
'  'and detail inset dataframes
'
'  Set pGraphicsContainer = pMxDoc.PageLayout
'                                                  'make assumption that for every
'                                                  'detail inset, one arrow and one
'                                                  'arrow shadow exists
'  lDetailInsetCount = pDictDetailInsets.Count
'  vDetailInsetKeys = pDictDetailInsets.Keys
'  For i = 0 To lDetailInsetCount - 1
'    sInsetID = vDetailInsetKeys(i)
'    Set pElement = pDictArrows.Item(sInsetID)
'    pGraphicsContainer.DeleteElement pElement
'    Set pElement = pDictArrowShadows.Item(sInsetID)
'    pGraphicsContainer.DeleteElement pElement
'    Set pElement = pDictDetailInsets.Item(sInsetID)
'    pGraphicsContainer.DeleteElement pElement
'  Next i
'
'  'Delete the detail inset's data frame
'  'lMapCount = pMxDoc.Maps.Count
'  'For i = 0 To lMapCount - 1
'  '  Set pCurrMap = pMxDoc.Maps.Item(i)
'  '  If pCurrMap.Name = (GBL_BUBBLE_IDENTIFIER & lBubbleID) And (pCurrMap.Description = ("Detail Inset " & lBubbleID)) Then
'  '    pMxDoc.Maps.Remove (pCurrMap)
'  '    i = lMapCount - 1
'  '  End If
'  'Next i
'  'pOldMap.Name = GBL_BUBBLE_IDENTIFIER & lBubbleID
'  'pOldMap.Description = "Detail Inset " & lBubbleID
'
'
'  '-------------------------------------
'  'Acquire the selection of detail inset
'  'polygons from which to generate detail insets.
'
'                                                  'find the detail inset bubble layer
'  Set pLayers = pMxDoc.FocusMap.Layers
'  pLayers.Reset
'  Set pLayer = pLayers.Next
'  bLyrIsFound = False
'  Do While (Not pLayer Is Nothing) And (Not bLyrIsFound)
'    If pLayer.Name = pNWSeriesOpts.BubbleLayer Then
'      bLyrIsFound = True
'      Set pInsetLyr = pLayer
'    End If
'    Set pLayer = pLayers.Next
'  Loop
'  If Not bLyrIsFound Then
'    MsgBox "Could not locate detail inset layer " & pNWSeriesOpts.BubbleLayer & vbNewLine _
'         & "Detail insets will not be generated for this map page." & vbNewLine _
'         & "Adjusting the ''N@ Options'' tab of the map series properties" & vbNewLine _
'         & "is recommended.", vbExclamation, "Missing Detail Insets Layer"
'    Exit Sub
'  End If
'                                                  'QI to a feature layer,
'  Set pFeatureLyr = pInsetLyr
'  Set pFeatureClass = pFeatureLyr.FeatureClass
'  If Not pNWSeriesOpts.IsBubbleLayer(pFeatureClass) Then
'    MsgBox "Detail inset layer is supposed to be a ''feature'' layer." & vbNewLine _
'         & "If this layer is a raster layer or a TIN, please reselect a" & vbNewLine _
'         & "bubble detail inset layer in the map series properties.", vbExclamation, _
'         "Non-valid Detail Insets Layer"
'    Exit Sub
'  End If
'                                                  'select all features that are entirely
'                                                  'contained within the map extent.
'  Set pSpatialFilter = New SpatialFilter
'  pSpatialFilter.SpatialRel = esriSpatialRelContains
'  pSpatialFilter.GeometryField = pFeatureClass.ShapeFieldName
'  Set pSpatialFilter.Geometry = m_pPageShape
'  'Set pSpatialFilter.Geometry = pMxDoc.ActiveView.Extent.Envelope   'correct?  type mismatch?
'
'
'  'use pFeatureClass.Search pSpatialFilter to
'  'get a cursor of all detail insets within the
'  'current map
'  Set pFeatCursor = pFeatureClass.Search(pSpatialFilter, True)
'  Set pBubbleFeature = pFeatCursor.NextFeature
'
'
'  '------------------------------------------------
'  'for each result feature, generate a detail inset
'
'  Do While Not pBubbleFeature Is Nothing
'    Set pRow = pBubbleFeature
'    AddToArcMap pDataFrameElement, pMxDoc, pRow, pDictDetailInsets, pDictArrows, pDictArrowShadows
'    Set pBubbleFeature = pFeatCursor.NextFeature
'  Loop
'
'
'
''Private Sub DeleteFromArcMap(pMxDoc As esriArcMapUI.IMxDocument, pRow As esriGeoDatabase.IRow)
''  Dim pFields As IFields, lFieldCount As Long, pField As IField
''  Dim lBubbleID As Long, pOldMap As IMap, lMapCount As Long
''  Dim i As Long, pCurrMap As IMap
'
'  '  If pRow Is Nothing Then Exit Sub
'  '  If pMxDoc Is Nothing Then
'  '    Exit Sub
'  '  End If
'
'
'  '----------------------------------------
'  'grab the detail inset name from the pRow
'
'  '  Set pFields = pRow.Fields
'  '  lFieldCount = pFields.FieldCount
'  '
'  '  For i = 0 To lFieldCount - 1
'  '    Set pField = pFields.Field(i)
'  '    Select Case pField.Name
'  '    Case "BUBBLEID"
'  '      lBubbleID = pRow.Value(i)
'  '      i = lFieldCount - 1
'  '    End Select
'  '  Next i
'
'
'  '
'  'AddDataframe pElement, pMxDoc, pRow
'
'
'  Exit Sub
'ErrorHandler:
'  HandleError False, "RefreshDetailInsets " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
'End Sub



Private Sub RefreshDynamicDefinitionQueries(pMxDoc As esriArcMapUI.IMxDocument, pNWSeriesOpts As INWMapSeriesOptions)
  On Error GoTo ErrorHandler


  Dim sNewString As String, pMaps As IMaps, pMap As IMap, pLayers As IEnumLayer
  Dim pLayer As ILayer, pFeatLayerDefinition As IFeatureLayerDefinition
  Dim lMapCount As Long, i As Long, sOldString As String, sTemp As String
  Dim pMapFrame As IMapFrame, vStoredFrames As Variant, sStoredFrame As String
  Dim lDynLyrCount As Long
  
  If pMxDoc Is Nothing Then Exit Sub
  Set pMaps = pMxDoc.Maps
  lMapCount = pMaps.Count
  sOldString = pNWSeriesOpts.DynamicDefQueryPrevReplaceString
  sNewString = pNWSeriesOpts.DynamicDefQueryReplaceString
  
  If sOldString = "" Then
    'if there is one or more definition query enabled layers
    lDynLyrCount = pNWSeriesOpts.DynamicDefQueryTrackingLayerCount
    If lDynLyrCount > 0 Then
      MsgBox "Warning: The previous map page listing is ''''.  Without a previous map " & vbNewLine _
           & "page, no dynamic definition queries will be updated in the " & lDynLyrCount & " layers" & vbNewLine _
           & "currently being tracked for dynamic definition query updates.", vbCritical, "Dynamic Definition Queries"
      Exit Sub
    End If
  End If
                                                  'for each layer in each map
  For i = 0 To (lMapCount - 1)
    Set pMap = pMaps.Item(i)
    Set pLayers = pMap.Layers
    pLayers.Reset
    Set pLayer = pLayers.Next
    Do While Not pLayer Is Nothing
                                                  'if that layer is being tracked,
      If pNWSeriesOpts.DynamicDefQueryIsTrackingLayer(pMap.Name, pLayer.Name) Then
        Set pFeatLayerDefinition = pLayer
        If Not pFeatLayerDefinition Is Nothing Then
                                                  'then substitute the previous map page string
                                                  'for the current map page string in the
                                                  'definition query
          sTemp = Replace(pFeatLayerDefinition.DefinitionExpression, _
                          sOldString, _
                          sNewString, _
                          1, _
                          -1, _
                          vbTextCompare)
          pFeatLayerDefinition.DefinitionExpression = sTemp
        End If
      End If
      Set pLayer = pLayers.Next
    Loop
  Next i
                                                  'grab the list of stored data frames
  vStoredFrames = pNWSeriesOpts.DataFramesStored
  lMapCount = UBound(vStoredFrames) + 1
  For i = 0 To (lMapCount - 1)
    sStoredFrame = vStoredFrames(i)
    Set pMapFrame = pNWSeriesOpts.DataFrameStoredFrame(sStoredFrame)
    If Not pMapFrame Is Nothing Then
      Set pMap = pMapFrame.Map
      Set pLayers = pMap.Layers
      pLayers.Reset
      Set pLayer = pLayers.Next
                                                    'for each layer in each stored data frame
      Do While Not pLayer Is Nothing
                                                    'if that layer is being tracked,
        If pNWSeriesOpts.DynamicDefQueryIsTrackingLayer(pMap.Name, pLayer.Name) Then
          Set pFeatLayerDefinition = pLayer
          If Not pFeatLayerDefinition Is Nothing Then
                                                    'then substitute the previous map page string
                                                    'for the current map page string in the
                                                    'definition query
            sTemp = Replace(pFeatLayerDefinition.DefinitionExpression, _
                            sOldString, _
                            sNewString, _
                            1, _
                            -1, _
                            vbTextCompare)
            pFeatLayerDefinition.DefinitionExpression = sTemp
          End If
        End If
        Set pLayer = pLayers.Next
      Loop
    End If
  Next i
  
  

  Exit Sub
ErrorHandler:
  HandleError False, "RefreshDynamicDefinitionQueries " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub





Private Sub RefreshIndicators(pDoc As IMxDocument, pSeriesProps As INWDSMapSeriesProps, bRefreshFlag As Boolean)
'Routine for updating any identicator maps there might be
On Error GoTo ErrHand:
  Dim pGridLayer As IFeatureLayer, pGridSel As IFeatureSelection
  Dim lLoop As Long, pActive As IActiveView, pSpatial As ISpatialFilter
  Dim pFeature As IFeature, pCursor As IFeatureCursor, pEnv As IEnvelope
  Dim pQuery As IQueryFilter, lLoop2 As Long, pMap As IMap, pSelEvents As ISelectionEvents
  Dim lGlobalStrLen As Long, lLocalStrLen As Long, sMapName As String
  Dim lStrLen As Long
  
  'Check for indicator maps and update those also
  lGlobalStrLen = Len("Global Indicator")
  lLocalStrLen = Len("Local Indicator")
  
  
  For lLoop = 0 To pDoc.Maps.Count - 1
    sMapName = ""
    If (StrComp((Left$(pDoc.Maps.Item(lLoop).Name, lGlobalStrLen)), "Global Indicator", vbTextCompare) = 0) Then
      sMapName = "Global Indicator"
      lStrLen = lGlobalStrLen
    ElseIf (StrComp(Left$(pDoc.Maps.Item(lLoop).Name, lLocalStrLen), "Local Indicator", vbTextCompare) = 0) Then
      sMapName = "Local Indicator"
      lStrLen = lLocalStrLen
    End If
    'If (StrComp((Left$(pDoc.Maps.Item(lLoop), lGlobalStrLen)), "Glocal Indicator", vbTextCompare) = 0) Or (StrComp(Left$(pDoc.Maps.Item(lLoop).Name, lLocalStrLen), "Local Indicator", vbTextCompare) = 0) Then
    'If pDoc.Maps.Item(lLoop).Name = "Global Indicator" Or pDoc.Maps.Item(lLoop).Name = "Local Indicator" Then
    If sMapName = "Global Indicator" Or sMapName = "Local Indicator" Then
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
        
      'If pMap.Name = "Global Indicator" Then
      If sMapName = "Global Indicator" Then
        Set pActive = pDoc.Maps.Item(lLoop)
        If bRefreshFlag Then pActive.Refresh
      'ElseIf pMap.Name = "Local Indicator" Then
      ElseIf sMapName = "Local Indicator" Then
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
      Set pSelEvents = pMap
      pSelEvents.SelectionChanged
      
      Set pGridLayer = Nothing
    End If
  Next lLoop

  Exit Sub
ErrHand:
  MsgBox "RefreshIndicators - " & Err.Description
End Sub

Private Sub UpdateTaggedElements(pDoc As IMxDocument, sTileName As String, bRefreshFlag As Boolean, _
 pNWDSMapSeries As INWDSMapSeries)
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
      Case "NWDSMAPBOOK - DATE"
        sText = Format(Date, "mmm dd, yyyy")
      Case "NWDSMAPBOOK - TITLE"
        sText = sTileName
      Case "NWDSMAPBOOK - PAGENUMBER"
        sText = CStr(m_lPageNumber)
      Case "NWDSMAPBOOK - EXTRAITEM"
        sText = GetExtraItemValue(pDoc, sTileName, pElemProps.Type, pNWDSMapSeries)
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

Private Function GetBubbleLayer(pNWSeriesOpts As INWMapSeriesOptions, pMxDoc As IMxDocument) As ILayer
  On Error GoTo ErrorHandler

  Dim pLayers As IEnumLayer, pLayer As ILayer, pInsetLyr As ILayer, i As Long
  Dim bLyrIsFound As Boolean, pFeatureLyr As IFeatureLayer, pFeatureClass As IFeatureClass
  Dim pMap As IMap, bMapIsFound As Boolean, pMaps As IMaps, lMapCount As Long
  Dim sMainDataFrame As String
  
  If pNWSeriesOpts.BubbleLayer = "" Then
    Set GetBubbleLayer = Nothing
    Exit Function
  End If
                                                  'find the detail inset bubble layer
  sMainDataFrame = pNWSeriesOpts.DataFrameMainFrame
  bMapIsFound = False
  Set pMaps = pMxDoc.Maps
  lMapCount = pMaps.Count
  For i = 0 To (lMapCount - 1)
    Set pMap = pMaps.Item(i)
    If StrComp(pMap.Name, sMainDataFrame, vbTextCompare) = 0 Then
      bMapIsFound = True
      i = lMapCount - 1
    End If
  Next i
  If Not bMapIsFound Then
    MsgBox "Could not find the data frame " & sMainDataFrame & vbNewLine _
         & "Please confirm that the data frame containing the bubble detail" & vbNewLine _
         & "polygons is still visible." & vbNewLine, vbCritical, "Data frame with bubble polygons not found"
    Exit Function
  End If
  
  
  Set pLayers = pMap.Layers
  pLayers.Reset
  Set pLayer = pLayers.Next
  bLyrIsFound = False
  Do While (Not pLayer Is Nothing) And (Not bLyrIsFound)
    If pLayer.Name = pNWSeriesOpts.BubbleLayer Then
      bLyrIsFound = True
      Set pInsetLyr = pLayer
    End If
    Set pLayer = pLayers.Next
  Loop
  If Not bLyrIsFound Then
    MsgBox "Could not locate detail inset layer " & pNWSeriesOpts.BubbleLayer & vbNewLine _
         & "Detail insets will not be generated for this map page." & vbNewLine _
         & "Adjusting the ''NW Options'' tab of the map series properties" & vbNewLine _
         & "is recommended.", vbExclamation, "Missing Detail Insets Layer"
    Set GetBubbleLayer = Nothing
    Exit Function
  End If
                                                  'QI to a feature layer,
  Set pFeatureLyr = pInsetLyr
  Set pFeatureClass = pFeatureLyr.FeatureClass
  If Not pNWSeriesOpts.IsBubbleLayer(pFeatureClass) Then
    MsgBox "Detail inset layer is supposed to be a ''feature'' layer." & vbNewLine _
         & "If this layer is a raster layer or a TIN, please reselect a" & vbNewLine _
         & "bubble detail inset layer in the map series properties.", vbExclamation, _
         "Non-valid Detail Insets Layer"
    Set GetBubbleLayer = Nothing
    Exit Function
  End If

  Set GetBubbleLayer = pInsetLyr

  Exit Function
ErrorHandler:
  HandleError False, "GetBubbleLayer " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function


Private Function GetExtraItemValue(pDoc As IMxDocument, sTileName As String, sFieldName As String, _
 pSeriesProps As INWDSMapSeriesProps) As String
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
      If UCase(pElemProps.Name) = "NWDSMAPBOOK CLIP ELEMENT" Then
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
  pElemProps.Name = "NWDSMapBook Clip Element"
  pGraphs.AddElement pNewElem, 0

  Exit Sub
ErrHand:
  MsgBox "CreateClipElement - " & Erl & " - " & Err.Description
End Sub

Private Sub SetMapExtent(pSeriesOpts As INWDSMapSeriesOptions, pActiveView As IActiveView)
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






Private Function ConvertPixelsToRW(pixelUnits As Double, pMxDoc As IMxDocument) As Double
  Dim realWorldDisplayExtent As Double
  Dim pixelExtent As Long
  Dim sizeOfOnePixel As Double
  Dim pDT As IDisplayTransformation
  Dim deviceRECT As tagRECT
  Dim pEnv As IEnvelope
  Dim pActiveView As IActiveView

  Set pActiveView = pMxDoc.FocusMap
  Set pDT = pActiveView.ScreenDisplay.DisplayTransformation
  deviceRECT = pDT.DeviceFrame
  pixelExtent = deviceRECT.Right - deviceRECT.Left
  Set pEnv = pDT.VisibleBounds
  realWorldDisplayExtent = pEnv.Width
  sizeOfOnePixel = realWorldDisplayExtent / pixelExtent
  ConvertPixelsToRW = pixelUnits * sizeOfOnePixel
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





''control visibility of (most) layout elements
''-------------------------------
'Private Sub FilterLayoutElements(pMxDoc As IMxDocument, _
'                                 sPageName As String, _
'                                 pNWSeriesOptions As INWMapSeriesOptions)
'
'  Dim vInvisElems As Variant, sInvisElem As String, pGraphicsCont As IGraphicsContainer
'  Dim pElement As IElement, pPageLayout As IPageLayout, lElemCount As Long
'  Dim i As Long, bIsInvisible As Boolean
'  Dim pElemProps As IElementProperties, vCustProp As Variant
'  Dim sCustomProp As String, pDisplay As IDisplay, pAppDisplay As IAppDisplay
'
'  If pMxDoc Is Nothing Then
'    Err.Raise 0, "routine FilterLayoutElements in NWMapPage.cls", _
'              "The pMxDoc reference was detected to be equal to ''Nothing''."
'  End If
'  If pNWSeriesOptions Is Nothing Then
'    Err.Raise 0, "routine FilterLayoutElements in NWMapPage.cls", _
'              "The pNWSeriesOptions reference was detected to be equal to ''Nothing''."
'  End If
'
'  vInvisElems = pNWSeriesOptions.ElementsGetElementsInvisibleForMapPage(sPageName)
'  lElemCount = UBound(vInvisElems)
'
'                                                  'cycle through all layout elements
'  Set pPageLayout = pMxDoc.PageLayout
'  Set pGraphicsCont = pPageLayout
'  pGraphicsCont.Reset
'  Set pElement = pGraphicsCont.Next
'  Do While Not pElement Is Nothing
'                                                  'ignore data frames -- that's handled by
'                                                  'dataframe visibility management
'    If TypeOf pElement Is IMapFrame Then
'      GoTo nextLoop
'    End If
'                                                  'ignore text elements -- that's handled by
'                                                  'the various text element tags.
'    If TypeOf pElement Is ITextElement Then
'      GoTo nextLoop
'    End If
'    Set pElemProps = pElement
'
'    vCustProp = pElemProps.CustomProperty
'    If Not IsEmpty(vCustProp) Then
'      If StrComp(TypeName(vCustProp), "string", vbTextCompare) = 0 Then
'        sCustomProp = vCustProp
'                                                  'if the element is marked as visible
'                                                  'on this map page
'        If pNWSeriesOptions.ElementsElementIsVisibleOnMapPage(sPageName, sCustomProp) Then
'                                                  'make it visible
'          pElement.Activate pDisplay
'        Else
'                                                  'else make it not visible
'          pElement.Deactivate
'        End If
'      End If 'if vCustProp is a string
'    End If 'if not IsEmpty(vCustProp)
'
'nextLoop:
'    Set pElement = pGraphicsCont.Next
'  Loop 'do while not pElement is Nothing
'
'
'End Sub






'
'Public Function GetActiveDataFrameName(pMxDoc As IMxDocument) As String
'    Dim pFocusMap As IMap
'
'    Set pFocusMap = pMxDoc.FocusMap
'    GetActiveDataFrameName = pFocusMap.Name
'End Function
'
'
'
'Public Function GetDataFrameElement(sDataFrameName As String, pMxDoc As IMxDocument) As IElement
'' Get the data frame element by name
'    Dim pGraphicsContainer As IGraphicsContainer
'    Dim pElementProperties As IElementProperties
'    Dim pElement As IElement
'    Dim pFE As IFrameElement
'    Dim pElProps As IElementProperties
'
'    On Error GoTo ErrorHandler
'
'    ' Loop through the elements (in the layout)
'    Set pGraphicsContainer = pMxDoc.PageLayout
'    pGraphicsContainer.Reset
'    Set pElement = pGraphicsContainer.Next
'    While Not pElement Is Nothing
'        ' If type of element is an IFrameElement
'        If TypeOf pElement Is IFrameElement Then
'            Set pElProps = pElement
'            ' If Name matches
'            If UCase(pElProps.Name) = UCase(sDataFrameName) Then
'                ' Return element
'                Set GetDataFrameElement = pElement
'                Set pElement = Nothing
'            Else
'                Set pElement = pGraphicsContainer.Next
'            End If
'        Else
'            Set pElement = pGraphicsContainer.Next
'        End If
'    Wend
'
'    Exit Function
'ErrorHandler:
'    Err.Raise Err.Number, Err.Source, "Error in GetDataFrameElement:" _
'        & vbCrLf & Err.Description
'End Function
'
'
