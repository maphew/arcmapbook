Attribute VB_Name = "modGeneralFunctions"

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


Public Const c_DefaultFld_Shape = "SHAPE"
Public Const cPI = 3.14159265358979
Const c_sModuleFileName As String = "modGeneralFunctions.bas"


Public Function GetUnitsDescription(pUnits As esriUnits) As String
    Select Case pUnits
        Case esriInches: GetUnitsDescription = "Inches"
        Case esriPoints: GetUnitsDescription = "Points"
        Case esriFeet: GetUnitsDescription = "Feet"
        Case esriYards: GetUnitsDescription = "Yards"
        Case esriMiles: GetUnitsDescription = "Miles"
        Case esriNauticalMiles: GetUnitsDescription = "Nautical miles"
        Case esriMillimeters: GetUnitsDescription = "Millimeters"
        Case esriCentimeters: GetUnitsDescription = "Centimeters"
        Case esriMeters: GetUnitsDescription = "Meters"
        Case esriKilometers: GetUnitsDescription = "Kilometers"
        Case esriDecimalDegrees: GetUnitsDescription = "Decimal degrees"
        Case esriDecimeters: GetUnitsDescription = "Decimeters"
        Case esriUnknownUnits: GetUnitsDescription = "Unknown"
        Case Else: GetUnitsDescription = "Unknown"
    End Select
End Function

Public Function GetActiveDataFrameName(pApp As IApplication) As String
    Dim pMx As IMxDocument
    Dim pFocusMap As IMap
    
    Set pMx = pApp.Document
    Set pFocusMap = pMx.FocusMap
    
    GetActiveDataFrameName = pFocusMap.Name
End Function

Public Function GetDataFrameElement(sDataFramName As String, pApp As IApplication) As IElement
' Get the data frame element by name
    Dim pGraphicsContainer As IGraphicsContainer
    Dim pElementProperties As IElementProperties
    Dim pElement As IElement
    Dim pMx As IMxDocument
    Dim pFE As IFrameElement
    Dim pElProps As IElementProperties
    
    On Error GoTo ErrorHandler
    
    ' Init
    Set pMx = pApp.Document
    ' Loop through the elements (in the layout)
    Set pGraphicsContainer = pMx.PageLayout
    pGraphicsContainer.Reset
    Set pElement = pGraphicsContainer.Next
    While Not pElement Is Nothing
        ' If type of element is an IFrameElement
        If TypeOf pElement Is IFrameElement Then
            Set pElProps = pElement
            ' If Name matches
            If UCase(pElProps.Name) = UCase(sDataFramName) Then
                ' Return element
                Set GetDataFrameElement = pElement
                Set pElement = Nothing
            Else
                Set pElement = pGraphicsContainer.Next
            End If
        Else
            Set pElement = pGraphicsContainer.Next
        End If
    Wend
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, "Error in GetDataFrameElement:" _
        & vbCrLf & Err.Description
End Function

Public Function FindFeatureLayerByDS(DatasetName As String, pApp As IApplication) As IFeatureLayer
  
    On Error GoTo ErrorHandler
  
    Dim pMxDoc As IMxDocument
    Dim pMap As IMap
    Dim pFeatureLayer As IFeatureLayer
    Dim pDataset As IDataset
    Dim i As Integer
    
    Set pMxDoc = pApp.Document
    Set pMap = pMxDoc.FocusMap
  
    With pMap
        For i = 0 To .LayerCount - 1
            If TypeOf .Layer(i) Is IFeatureLayer Then
                Set pFeatureLayer = .Layer(i)
                Set pDataset = pFeatureLayer.FeatureClass
                If UCase(pDataset.Name) = UCase(DatasetName) Then
                    Set FindFeatureLayerByDS = pFeatureLayer
                    Exit For
                End If
            End If
        Next i
    End With
  
    If pFeatureLayer Is Nothing Then
        Err.Raise vbObjectError, "FindFeatureLayerByDS", "Error in " _
            & "FindFeatureLayerByDS:" & vbCrLf & "Could not locate " _
            & "layer with a dataset name of '" & DatasetName & "'."
    End If
  
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, "Error in routine: FindFeatureLayerByDS" _
        & vbCrLf & Err.Description
End Function

Public Function FindFeatureLayerByName(FLName As String, pApp As IApplication) As IFeatureLayer
  
    On Error GoTo ErrorHandler
  
    Dim pMxDoc As IMxDocument
    Dim pMap As IMap
    Dim pFeatureLayer As IFeatureLayer
    Dim pDataset As IDataset
    Dim i As Integer
    
    Set pMxDoc = pApp.Document
    Set pMap = pMxDoc.FocusMap
  
    With pMap
        For i = 0 To .LayerCount - 1
            If TypeOf .Layer(i) Is IFeatureLayer Then
                Set pFeatureLayer = .Layer(i)
                If UCase(pFeatureLayer.Name) = UCase(FLName) Then
                    Set FindFeatureLayerByName = pFeatureLayer
                    Exit For
                End If
            End If
        Next i
    End With
  
    If pFeatureLayer Is Nothing Then
        Err.Raise vbObjectError, "FindFeatureLayerByName", "Error in " _
            & "FindFeatureLayerByName:" & vbCrLf & "Could not locate " _
            & "layer with a Name of '" & FLName & "'."
    End If
  
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, "Error in routine: FindFeatureLayerByName" _
        & vbCrLf & Err.Description
End Function

Public Function GetValidExtentForLayer(pFL As IFeatureLayer) As IEnvelope
    Dim pGeoDataset As IGeoDataset
    Dim pMx As IMxDocument
    Dim pW As IWorkspace
    Dim pWSR As IWorkspaceSpatialReferenceInfo
    Dim pEnumSRI As IEnumSpatialReferenceInfo
    Dim pSR As ISpatialReference
    Dim dX1 As Double, dY1 As Double
    Dim dX2 As Double, dY2 As Double
    Dim pP As IPoint
    
    If Not pFL Is Nothing Then
        If Not pFL.FeatureClass Is Nothing Then
            If TypeOf pFL.FeatureClass Is IGeoDataset Then
                If pFL.FeatureClass.FeatureDataset Is Nothing Then
                    dX1 = -1000000000
                    dY1 = -1000000000
                    dX2 = 1000000000
                    dY2 = 1000000000
                Else
                    Set pW = pFL.FeatureClass.FeatureDataset.Workspace
                    Set pWSR = pW
                    Set pEnumSRI = pWSR.SpatialReferenceInfo
                    Set pSR = pEnumSRI.Next(0)
                    pSR.GetDomain dX1, dX2, dY1, dY2
                End If
                Set pP = New esriGeometry.Point
                Set GetValidExtentForLayer = New envelope
                pP.PutCoords dX1, dY1
                GetValidExtentForLayer.LowerLeft = pP
                pP.PutCoords dX2, dY2
                GetValidExtentForLayer.UpperRight = pP
            Else
                Err.Raise vbObjectError, "GetValidExtentForLayer", _
                    "The 'FeatureClass' property for the IFeatureLayer parameter is not an IGeoDataset"
            End If
        Else
            Err.Raise vbObjectError, "GetValidExtentForLayer", _
                "The IFeatureLayer parameter does not have a valid FeatureClass property"
        End If
    Else
        Err.Raise vbObjectError, "GetValidExtentForLayer", _
            "The IFeatureLayer parameter is set to Nothing"
    End If
End Function


Public Function CreateInsetPolygon(dOriginX As Double, _
                                    dOriginY As Double, _
                                    dHeight As Double, _
                                    dWidth As Double) As IPolygon
  On Error GoTo ErrorHandler


  Dim pCircularArc As ICircularArc, pPolygon As IPolygon
  Dim dFlatLength As Double, dRadius As Double, pCenterPnt As IPoint
  Dim dpi As Double, pLine As ILine, pLine2 As ILine, pConstCircArc As IConstructCircularArc
  Dim pPnt1 As IPoint, pPnt2 As IPoint, pPnt3 As IPoint, pPnt4 As IPoint
  Dim pSegment As ISegment, pConstructLine As IConstructLine
  Dim pSegment1 As ISegment, pSegment2 As ISegment, pSegment3 As ISegment, pSegment4 As ISegment
  
  Dim pSegmentColl As ISegmentCollection
  Set pPolygon = New Polygon
  Set pSegmentColl = pPolygon
  
  
  Set pCircularArc = New CircularArc
  Set pConstCircArc = pCircularArc
  Set pCenterPnt = New Point
  dpi = 3.14159265358979
  
  Set pPnt1 = New Point
  Set pPnt2 = New Point
  Set pPnt3 = New Point
  Set pPnt4 = New Point

  If dHeight <> dWidth Then
    
    If dWidth > dHeight Then
      dFlatLength = dWidth - dHeight
      dRadius = dHeight / 2
      pCenterPnt.x = dOriginX - (dFlatLength / 2)
      pCenterPnt.y = dOriginY
      pPnt1.x = pCenterPnt.x
      pPnt1.y = pCenterPnt.y - dRadius
      pPnt2.x = pCenterPnt.x
      pPnt2.y = pCenterPnt.y + dRadius
      pConstCircArc.ConstructEndPointsRadius pPnt1, pPnt2, False, dRadius, True
      Set pSegment1 = pCircularArc
      
      Set pLine = New esriGeometry.Line
      pCircularArc.QueryTangent esriExtendTangentAtTo, pCircularArc.Length, False, dFlatLength, pLine
      Set pSegment2 = pLine
      
      Set pCircularArc = New CircularArc
      Set pConstCircArc = pCircularArc
      pCenterPnt.x = dOriginX + (dFlatLength / 2)
      pPnt3.x = pCenterPnt.x
      pPnt3.y = pCenterPnt.y + dRadius
      pPnt4.x = pCenterPnt.x
      pPnt4.y = pCenterPnt.y - dRadius
      pConstCircArc.ConstructEndPointsRadius pPnt3, pPnt4, False, dRadius, True
      Set pSegment3 = pCircularArc
      
      Set pLine2 = New esriGeometry.Line
      pCircularArc.QueryTangent esriExtendTangentAtTo, pCircularArc.Length, False, dFlatLength, pLine2
      Set pSegment4 = pLine2
    Else 'less wide, taller inset
      dFlatLength = dHeight - dWidth
      dRadius = dWidth / 2
      pCenterPnt.x = dOriginX
      pCenterPnt.y = dOriginY + (dFlatLength / 2)
      pPnt1.x = pCenterPnt.x - dRadius
      pPnt1.y = pCenterPnt.y
      pPnt2.x = pCenterPnt.x + dRadius
      pPnt2.y = pCenterPnt.y
      pConstCircArc.ConstructEndPointsRadius pPnt1, pPnt2, False, dRadius, True
      Set pSegment1 = pCircularArc
      
      Set pLine = New esriGeometry.Line
      pCircularArc.QueryTangent esriExtendTangentAtTo, pCircularArc.Length, False, dFlatLength, pLine
      Set pSegment2 = pLine
      
      Set pCircularArc = New CircularArc
      Set pConstCircArc = pCircularArc
      pCenterPnt.y = dOriginY - (dFlatLength / 2)
      pPnt3.x = pCenterPnt.x + dRadius
      pPnt3.y = pCenterPnt.y
      pPnt4.x = pCenterPnt.x - dRadius
      pPnt4.y = pCenterPnt.y
      pConstCircArc.ConstructEndPointsRadius pPnt3, pPnt4, False, dRadius, True
      Set pSegment3 = pCircularArc
      
      Set pLine2 = New esriGeometry.Line
      pCircularArc.QueryTangent esriExtendTangentAtTo, pCircularArc.Length, False, dFlatLength, pLine2
      Set pSegment4 = pLine2
    End If
    
    pSegmentColl.AddSegment pSegment1
    pSegmentColl.AddSegment pSegment2
    pSegmentColl.AddSegment pSegment3
    pSegmentColl.AddSegment pSegment4
    
  Else 'width == height
    dRadius = dHeight / 2
    pCenterPnt.x = dOriginX
    pCenterPnt.y = dOriginY
    pConstCircArc.ConstructCircle pCenterPnt, dRadius, False
    Set pSegment1 = pCircularArc
    
    pSegmentColl.AddSegment pSegment1
  End If

  pPolygon.Close
  Set CreateInsetPolygon = pPolygon
  

  Exit Function
ErrorHandler:
  HandleError False, "CreateInsetPolygon " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function


Public Function DoesShapeFileExist(pPath As String) As Boolean
  Dim pTruncPath As String
  If InStr(1, pPath, ".shp") > 0 Then
    pTruncPath = Left(pPath, InStr(1, pPath, ".shp") - 1)
  Else
    pTruncPath = pPath
  End If
      
  'Make sure the specified file does not exist
  Dim fs As Object
  Set fs = CreateObject("Scripting.FileSystemObject")
  If fs.FileExists(pTruncPath & ".shp") Or fs.FileExists(pTruncPath & ".dbf") Or _
   fs.FileExists(pTruncPath & ".shx") Then
    DoesShapeFileExist = True
  Else
    DoesShapeFileExist = False
  End If
End Function

Private Function DoesFeatureClassExist(location As IGxObject, newObjectName As String) As Boolean
On Error GoTo ErrHand:
  Dim pFeatClass As IFeatureClass
  Dim pFeatDataset As IGxDataset
  Set pFeatDataset = location
  Dim pFeatClassCont As IFeatureClassContainer, pData As IFeatureDataset
  Set pData = pFeatDataset.Dataset
  Set pFeatClassCont = pData
  Dim pEnumClass As IEnumFeatureClass, pDataset As IDataset
  Set pEnumClass = pFeatClassCont.Classes
  Set pFeatClass = pEnumClass.Next
  While Not pFeatClass Is Nothing
    Set pDataset = pFeatClass
    If UCase(pDataset.Name) = UCase(newObjectName) Then
      DoesFeatureClassExist = True
      Exit Function
    End If
      
    Set pFeatClass = pEnumClass.Next
  Wend
  DoesFeatureClassExist = False
  
  Exit Function
ErrHand:
  MsgBox Err.Description
End Function

Public Function NewAccessFile(pDatabase As String, _
 pNewDataSet As String, pNewFile As String, Optional pMoreFields As IFields) As IFeatureClass
On Error GoTo ErrorHandler

    Dim pName As IName
    Dim pOutShpWspName As IWorkspaceName
    Dim pShapeWorkspace As IWorkspace
    Dim pOutputFields As IFields
    Dim pFieldChecker As IFieldChecker
    Dim pErrorEnum As IEnumFieldError
    Dim pNewFields As IFields, pField As IField
    Dim pClone As IClone, pCloneFields As IFields
    Dim pFeatureWorkspace As IFeatureWorkspace
    Dim pDataset As IFeatureDataset
    Dim shapeFieldName As String
    Dim pNewFeatClass As IFeatureClass
    Dim pFieldsEdit As IFieldsEdit
    Dim newFieldEdit As IFieldEdit
    Dim pGeomDef As IGeometryDef
    Dim pGeomDefEdit As IGeometryDefEdit
    Dim pGeoDataset As IGeoDataset
    Dim i As Integer
  
    Set pOutShpWspName = New WorkspaceName

    pOutShpWspName.PathName = pDatabase
    pOutShpWspName.WorkspaceFactoryProgID = "esriDataSourcesGDB.AccessWorkspaceFactory"
    Set pName = pOutShpWspName
    Set pShapeWorkspace = pName.Open
i = 1
    'Open the dataset
    Set pFeatureWorkspace = pShapeWorkspace
    Set pDataset = pFeatureWorkspace.OpenFeatureDataset(pNewDataSet)
i = 2
    ' Add the SHAPE field (based on the dataset)
    Set pFieldsEdit = pMoreFields
    Set pField = New Field
    Set newFieldEdit = pField
    With newFieldEdit
        .Name = c_DefaultFld_Shape
        .Type = esriFieldTypeGeometry
        .IsNullable = True
        .Editable = True
    End With
    Set pGeomDef = New GeometryDef
    Set pGeomDefEdit = pGeomDef
    With pGeomDefEdit
        .GeometryType = esriGeometryPolygon
        If TypeOf pDataset Is IGeoDataset Then
            Set pGeoDataset = pDataset
            Set .SpatialReference = pGeoDataset.SpatialReference
        Else
            Set .SpatialReference = New UnknownCoordinateSystem
        End If
        .GridCount = 1
        .GridSize(0) = 200
        .HasM = False
        .HasZ = False
        .AvgNumPoints = 4
    End With
    Set newFieldEdit.GeometryDef = pGeomDef
    pFieldsEdit.AddField pField
    ' Check the fields
    Set pFieldChecker = New FieldChecker
    Set pFieldChecker.ValidateWorkspace = pShapeWorkspace
    Set pNewFields = pMoreFields
i = 3
    Set pClone = pNewFields
    Set pCloneFields = pClone.Clone
    pFieldChecker.Validate pCloneFields, pErrorEnum, pOutputFields
      
  ' Create the output featureclass
  Dim pUID As New UID
  pUID = "{52353152-891A-11D0-BEC6-00805F7C4268}"
    shapeFieldName = c_DefaultFld_Shape
i = 4
    Set pNewFeatClass = pDataset.CreateFeatureClass(pNewFile, pOutputFields, pUID, Nothing, esriFTSimple, shapeFieldName, "")
i = 5
    Set NewAccessFile = pNewFeatClass
  
    Exit Function
  
ErrorHandler:
    MsgBox Err.Number & " " & Err.Description, vbCritical, "Error in NewAccessFile " & i
End Function

Public Function NewShapeFile(pNewFile As String, pMap As IMap, _
            Optional pMoreFields As IFields) As IFeatureClass
    On Error GoTo ErrorHandler

    Dim pOutShpWspName As IWorkspaceName
    Dim pName As IName
    Dim pShapeWorkspace As IWorkspace
    Dim pOutputFields As IFields
    Dim pFieldChecker As IFieldChecker
    Dim pErrorEnum As IEnumFieldError
    Dim pNewFields As IFields, pField As IField
    Dim pClone As IClone, pCloneFields As IFields
    Dim featureclassName As String, pNewFeatClass As IFeatureClass
    Dim pFeatureWorkspace As IFeatureWorkspace
    Dim pUID As IUID
    Dim shapeFieldName As String
    Dim pFieldsEdit As IFieldsEdit
    Dim newFieldEdit As IFieldEdit
    Dim pGeomDef As IGeometryDef
    Dim pGeomDefEdit As IGeometryDefEdit
    
    ' Open the workspace for the new shapefile
    Set pOutShpWspName = New WorkspaceName
    pOutShpWspName.PathName = EntryName(pNewFile)
    pOutShpWspName.WorkspaceFactoryProgID = "esriDataSourcesFile.ShapefileWorkspaceFactory.1"
    Set pName = pOutShpWspName
    Set pShapeWorkspace = pName.Open
    ' Add the SHAPE field (based on the Map)
    Set pFieldsEdit = pMoreFields
    Set pField = New Field
    Set newFieldEdit = pField
    newFieldEdit.Name = c_DefaultFld_Shape
    newFieldEdit.Type = esriFieldTypeGeometry
    Set pGeomDef = New GeometryDef
    Set pGeomDefEdit = pGeomDef
    With pGeomDefEdit
        .GeometryType = esriGeometryPolygon
        Set .SpatialReference = pMap.SpatialReference
    End With
    Set newFieldEdit.GeometryDef = pGeomDef
    pFieldsEdit.AddField pField
    ' Validate field names
    Set pFieldChecker = New FieldChecker
    Set pFieldChecker.ValidateWorkspace = pShapeWorkspace
    Set pNewFields = pMoreFields
    Set pClone = pNewFields
    Set pCloneFields = pClone.Clone
    pFieldChecker.Validate pCloneFields, pErrorEnum, pOutputFields
    ' Create the output featureclass
    shapeFieldName = c_DefaultFld_Shape
    featureclassName = Mid(pNewFile, Len(pOutShpWspName.PathName) + 2)
    Set pFeatureWorkspace = pShapeWorkspace
    Set pNewFeatClass = pFeatureWorkspace.CreateFeatureClass(featureclassName, pOutputFields, _
                            Nothing, Nothing, esriFTSimple, shapeFieldName, "")
    ' Return
    Set NewShapeFile = pNewFeatClass
  
    Exit Function
  
ErrorHandler:
    MsgBox "Error creating " & pNewFile & vbCrLf & Err.Number & ": " & Err.Description, _
        vbCritical, "Error in NewShapefile"
End Function

Public Function EntryName(sFile As String) As String
  ' work from the right side to the first file delimeter
  Dim iLength As Integer
  iLength = Len(sFile)
  Dim iCounter As Integer
  Dim sDelim As String
  sDelim = "\"
  Dim sRight As String
  
  For iCounter = iLength To 0 Step -1
    
    If Mid$(sFile, iCounter, 1) = sDelim Then
      EntryName = Mid$(sFile, 1, (iCounter - 1))
      Exit For
    End If
  
  Next
  
End Function


Public Sub SetVisibleLayers(pMap As IMap, sLayers As String)
On Error GoTo ErrorHandler
  Dim sLayerArr() As String, i As Long, j As Long
  Dim lVisLyrUBound As Long, pLayers As IEnumLayer
  Dim pLayer As ILayer
  
  
  sLayerArr = Split(sLayers, ",", -1, vbTextCompare)
  lVisLyrUBound = UBound(sLayerArr)
  Set pLayers = pMap.Layers
  pLayers.Reset
  Set pLayer = pLayers.Next
  
  If lVisLyrUBound = -1 Then
    Do While Not pLayer Is Nothing
      pLayer.Visible = False
      Set pLayer = pLayers.Next
    Loop
    Exit Sub
  Else
    Do While Not pLayer Is Nothing
      pLayer.Visible = False
      For i = 0 To lVisLyrUBound
        If pLayer.Name = sLayerArr(i) Then
          pLayer.Visible = True
        End If
      Next i
      Set pLayer = pLayers.Next
    Loop
  End If
  
  
  Exit Sub
ErrorHandler:
  MsgBox "SetVisibleLayers - " & Err.Description
End Sub



Public Sub TurnOffClipping(pSeriesProps As INWDSMapSeriesProps, pApp As IApplication)
On Error GoTo ErrHand:
  Dim pMap As IMap, pDoc As IMxDocument
  'Find the data frame
  Set pDoc = pApp.Document
  Set pMap = FindDataFrame(pDoc, pSeriesProps.DataFrameName)
  If pMap Is Nothing Then Exit Sub
  
  pMap.ClipGeometry = Nothing

  Exit Sub
ErrHand:
  MsgBox "TurnOffClipping - " & Err.Description
End Sub

Public Sub RemoveIndicators(pApp As IApplication)
On Error GoTo ErrHand:
  Dim lLoop As Long, pDoc As IMxDocument, pDelColl As Collection
  Dim pPage As IPageLayout, pGraphCont As IGraphicsContainer
  Dim pElem As IElement, pMapFrame As IMapFrame
  Set pDoc = pApp.Document
  Set pPage = pDoc.PageLayout
  Set pDelColl = New Collection
  Set pGraphCont = pPage
  pGraphCont.Reset
  Set pElem = pGraphCont.Next
  
  Dim lGlobalStrLen As Long, lLocalStrlen As Long, lStrlen As Long
  Dim sMapName As String
  
  lGlobalStrLen = Len("Global Indicator")
  lLocalStrlen = Len("Local Indicator")
  Do While Not pElem Is Nothing
    If TypeOf pElem Is IMapFrame Then
      Set pMapFrame = pElem
                                                  'if the left most characters of pMapFrame.Map.Name
                                                  'are "Local Indicator" or "Global Indicator" ...
      If (StrComp((Left$(pMapFrame.Map.Name, lGlobalStrLen)), "Glocal Indicator", vbTextCompare) = 0) _
         Or (StrComp(Left$(pMapFrame.Map.Name, lLocalStrlen), "Local Indicator", vbTextCompare) = 0) Then
        pDelColl.Add pMapFrame
      End If
'      If pMapFrame.Map.Name = "Local Indicator" Or _
'       pMapFrame.Map.Name = "Global Indicator" Then
'        pDelColl.Add pMapFrame
'      End If
    End If
    
    Set pElem = pGraphCont.Next
  Loop
  
  For lLoop = 1 To pDelColl.count
    pGraphCont.DeleteElement pDelColl.Item(lLoop)
  Next lLoop

  Exit Sub
ErrHand:
  MsgBox "RemoveIndicators - " & Err.Description
End Sub

Public Sub RemoveLabels(pDoc As IMxDocument)
On Error GoTo ErrHand:
  Dim pGraphicsCont As IGraphicsContainer
  Dim pTempColl As Collection, pElemProps As IElementProperties, lLoop As Long
  'Remove any previous neighbor labels.
  Set pGraphicsCont = pDoc.PageLayout
  pGraphicsCont.Reset
  Set pTempColl = New Collection
  Set pElemProps = pGraphicsCont.Next
  Do While Not pElemProps Is Nothing
    If pElemProps.Name = "NWDSMAPBOOK TEXT" Then
      pTempColl.Add pElemProps
    End If
    Set pElemProps = pGraphicsCont.Next
  Loop
  For lLoop = 1 To pTempColl.count
    pGraphicsCont.DeleteElement pTempColl.Item(lLoop)
  Next lLoop
  Set pTempColl = Nothing

  Exit Sub
ErrHand:
  MsgBox "RemoveLabels - " & Err.Description
End Sub

Public Function GetMapBookExtension(pApp As IApplication) As INWDSMapBook
On Error GoTo ErrHand:
  Dim pMapBookExt As NWMapBookExt, pMapBook As INWDSMapBook
  Set pMapBookExt = pApp.FindExtensionByName("NW_MapBook")
  If pMapBookExt Is Nothing Then
    MsgBox "Map Book code not installed properly!!  Make sure you can access the regsvr32 command" & vbCrLf & _
     "and rerun the _Install.bat batch file!!", , "Map Book Extension Not Found!!!"
    Set GetMapBookExtension = Nothing
    Exit Function
  End If
  
  Set GetMapBookExtension = pMapBookExt.MapBook

  Exit Function
ErrHand:
  MsgBox "GetMapBookExtension - " & Err.Description
End Function

Public Sub RemoveClipElement(pDoc As IMxDocument)
On Error GoTo ErrHand:
  Dim pGraphs As IGraphicsContainer, pElemProps As IElementProperties
  
  'Search for an existing clip element and delete it when found
  Set pGraphs = pDoc.FocusMap
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
  pDoc.ActiveView.PartialRefresh esriViewGraphics, Nothing, Nothing

  Exit Sub
ErrHand:
  MsgBox "RemoveClipElement - " & Erl & " - " & Err.Description
End Sub








'Function takes a data frame name and a layer name.  Function returns the
'reference to that layer object.  This function assumes that the named data
'frame and layer exist in the map document.
'----------------------------------
Public Function LayerFromDataFrame(sDataFrame As String, sLayer As String, pMxDoc As IMxDocument) As ILayer
  On Error GoTo ErrorHandler
  
  Dim pLayers As IEnumLayer, pGraphicsContainer As IGraphicsContainer, pMap As IMap
  Dim pMapFrame As IMapFrame, pPageLayout As IPageLayout, bFoundFrame As Boolean
  Dim bFoundLayer As Boolean, pLayer As ILayer, pElement As IElement
  
  If pMxDoc Is Nothing Then
    Set LayerFromDataFrame = Nothing
    Exit Function
  End If
  Set pPageLayout = pMxDoc.PageLayout
  Set pGraphicsContainer = pPageLayout
  pGraphicsContainer.Reset
  Set pElement = pGraphicsContainer.Next
  bFoundFrame = False
  Do While (Not pElement Is Nothing) And Not bFoundFrame
    If TypeOf pElement Is IMapFrame Then
      Set pMapFrame = pElement
      Set pMap = pMapFrame.Map
      If StrComp(pMap.Name, sDataFrame, vbTextCompare) = 0 Then
        bFoundFrame = True
      End If
    End If
    Set pElement = pGraphicsContainer.Next
  Loop
  
  If Not bFoundFrame Then
    Set LayerFromDataFrame = Nothing
    Exit Function
  End If
  
  'data frame is found by this point
  Set pLayers = pMap.Layers
  Set pLayer = pLayers.Next
  bFoundLayer = False
  Do While (Not pLayer Is Nothing) And Not bFoundLayer
    If StrComp(sLayer, pLayer.Name, vbTextCompare) = 0 Then
      bFoundLayer = True
    Else
      Set pLayer = pLayers.Next
    End If
  Loop
   
  If bFoundLayer Then
    Set LayerFromDataFrame = pLayer
  Else
    Set LayerFromDataFrame = Nothing
  End If
  

  Exit Function
ErrorHandler:
  HandleError False, "LayerFromDataFrame " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function





'This function takes X,Y coordinates in map units, and returns where in the
'layout view those coordinate show up relative to the top left corner of the
'layout view map display.
'---------------------------------------
Public Function LayoutUnitsFromMapUnits(sDataFrameName As String, ByRef x As Double, ByRef y As Double, pApp As IApplication) As IPoint
  On Error GoTo ErrorHandler
  
  ' Get the data frame element by name -- code copied from GetDataFrameElement routine
  Dim pGraphicsContainer As IGraphicsContainer
  Dim pElementProperties As IElementProperties
  Dim pElement As IElement, pResultElement As IElement
  Dim pMx As IMxDocument
  Dim pFE As IFrameElement
  Dim pElProps As IElementProperties
  
  
  Set pMx = pApp.Document
  Set pGraphicsContainer = pMx.PageLayout
  pGraphicsContainer.Reset
  Set pElement = pGraphicsContainer.Next
  While Not pElement Is Nothing
      ' If type of element is an IFrameElement
      If TypeOf pElement Is IFrameElement Then
          Set pElProps = pElement
          ' If Name matches
          If UCase(pElProps.Name) = UCase(sDataFrameName) Then
              ' Return element
              Set pResultElement = pElement
              Set pElement = Nothing
          Else
              Set pElement = pGraphicsContainer.Next
          End If
      Else
          Set pElement = pGraphicsContainer.Next
      End If
  Wend
  Set pElement = pResultElement
  '--------------
  'Now that we have the dataframe, take the X,Y coordinates from
  'map units to the layout units (usually inches)
  
  Dim dDocWidthInches As Double, dDocHeightInches As Double, pMxDoc As IMxDocument
  Dim pPntScrBottomLeft As IPoint, pPntScrTopRight As IPoint
  Dim dScrMapUnitWidth As Double, dScrMapUnitHeight As Double
  Dim dScrMapUnitWidthWOBorders As Double, dScrMapUnitHeightWOBorders As Double
  Dim dRatioInchToMapUnitX As Double, dRatioInchToMapUnitY As Double
  Dim dBubbleLeftInches As Double, dBubbleRightInches As Double
  Dim dBubbleTopInches As Double, dBubbleBottomInches As Double
  Dim dBoundaryInchesXLeft As Double, dBoundaryInchesYTop As Double
  Dim pMapEnv As IEnvelope, dDataFrameWidth As Double, dDataFrameHeight As Double
  Dim pMxDocFocusMapQILayoutViewAV As IActiveView, pPnt As IPoint
  
  Set pMxDoc = pApp.Document
  Set pMapEnv = pMxDoc.ActiveView.Extent 'sides of display in screen units
  With pMapEnv
    dDocWidthInches = .XMax - (Abs(.XMin))
    dDocHeightInches = .YMax - (Abs(.YMin))
  End With
                                                    'get the data area width minus the
                                                    'buffer of empty space around the
                                                    'map display
  dDataFrameWidth = pElement.Geometry.envelope.Width
  dDataFrameHeight = pElement.Geometry.envelope.Height
  dBoundaryInchesXLeft = pElement.Geometry.envelope.XMin
  dBoundaryInchesYTop = pElement.Geometry.envelope.YMin
  
  Set pMxDocFocusMapQILayoutViewAV = pMxDoc.FocusMap
  With pMxDocFocusMapQILayoutViewAV.ScreenDisplay.DisplayTransformation
    Set pPntScrBottomLeft = .ToMapPoint(.DeviceFrame.Left, .DeviceFrame.bottom)
    Set pPntScrTopRight = .ToMapPoint(.DeviceFrame.Right, .DeviceFrame.Top)
  End With
  
  dScrMapUnitWidth = pPntScrTopRight.x - pPntScrBottomLeft.x
  dScrMapUnitHeight = pPntScrTopRight.y - pPntScrBottomLeft.y
  dRatioInchToMapUnitX = dDataFrameWidth / dScrMapUnitWidth
  dRatioInchToMapUnitY = dDataFrameHeight / dScrMapUnitHeight
  
  Set pPnt = New Point
  pPnt.x = ((x - pPntScrBottomLeft.x) * dRatioInchToMapUnitX) + dBoundaryInchesXLeft
  pPnt.y = ((y - pPntScrBottomLeft.y) * dRatioInchToMapUnitY) + dBoundaryInchesYTop
  Set LayoutUnitsFromMapUnits = pPnt

  Exit Function
ErrorHandler:
  HandleError True, "LayoutUnitsFromMapUnits " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'
'
'
'
'Private Sub AddDataframe(pElementWithinBorders As IElement, pApp As IApplication, pRow As IRow, pNWSeriesOptions As INWMapSeriesOptions)
'  On Error GoTo ErrorHandler
'
'  Dim pFields As IFields, lFieldCount As Long, pField As IField
'  Dim pNewMap As IMap, pMapFrame As IMapFrame, pEnv As IEnvelope
'  Dim pGraphicsContainer As IGraphicsContainer, pElement As IElement
'  Dim pFrameElement As IFrameElement, pNewMapAV As IActiveView
'  Dim pNewEnv As IEnvelope
'
'
'  Dim pMxDocDataViewAV As IActiveView, pScrDisplay As IScreenDisplay
'  Dim pMxDocLayoutViewAV As IActiveView, pMxDoc As IMxDocument
'
'  Dim lBubbleId As Long, dXOrigin As Double, dYOrigin As Double
'  Dim dXDestination As Double, dYDestination As Double, dRadius As Double
'  Dim dScale As Double, sLayers As String, dWidthOrigin As Double
'  Dim bIsCircular As Boolean, i As Long, dPageToMapUnitRatio As Double
'  Dim dMapScale As Double, dDataFrameWidth As Double, dDataFrameHeight As Double
'  Dim sActiveDataFrameName As String, pMxDocFocusMapQIDataViewAV As IActiveView
'  Dim pMxDocFocusMapQILayoutViewAV As IActiveView, pMainMapFrame As IMapFrame
'
'  If pRow Is Nothing Then Exit Sub
'  If pApp Is Nothing Then Exit Sub
'  Set pMxDoc = pApp.Document
'
'  Set pFields = pRow.Fields
'  lFieldCount = pFields.FieldCount
'
'  For i = 0 To lFieldCount - 1
'    Set pField = pFields.Field(i)
'    Select Case pField.Name
'    Case "BUBBLEID"
'      lBubbleId = pRow.Value(i)
'    Case "XORG"
'      dXOrigin = pRow.Value(i)
'    Case "YORG"
'      dYOrigin = pRow.Value(i)
'    Case "XDEST"
'      dXDestination = pRow.Value(i)
'    Case "YDEST"
'      dYDestination = pRow.Value(i)
'    Case "RADIUS"
'      dRadius = pRow.Value(i)
'    Case "SCALE"
'      dScale = pRow.Value(i)
'    Case "LAYERS"
'      sLayers = pRow.Value(i)
'    Case "WIDTHORG"
'      dWidthOrigin = pRow.Value(i)
'    End Select
'  Next i
'
'  Set pNewMap = New Map
'  pNewMap.Name = "BubbleID:" & lBubbleId
'  pNewMap.Description = "Detail Inset " & lBubbleId
'
'
'  ' clone all copies of layers so that modifying
'  ' layers in one dataframe doesn't impact layers
'  ' referenced in the other data frames.
'  ''''''''''''''''''''''
'  Dim pLayer As ILayer, pLayerSrc As ILayer
'  Dim pFeatLyr As IFeatureLayer, pTinLyr As ITinLayer, pRastLyrSrc As IRasterLayer
'  Dim pRastLyr As IRasterLayer, pFeatLyrSrc As IFeatureLayer, pTinLyrSrc As ITinLayer
'  Dim pGeoFeatLyrSource As IGeoFeatureLayer, pGeoFeatLyrDestination As IGeoFeatureLayer
'
'  If pMxDoc.FocusMap.LayerCount > 0 Then
'    For i = (pMxDoc.FocusMap.LayerCount - 1) To 0 Step -1
'      Set pLayerSrc = pMxDoc.FocusMap.Layer(i)
'      If TypeOf pMxDoc.FocusMap.Layer(i) Is IFeatureLayer Then
'        Set pFeatLyr = New FeatureLayer
'        Set pFeatLyrSrc = pLayerSrc
'        pFeatLyr.DataSourceType = pFeatLyrSrc.DataSourceType
'        pFeatLyr.DisplayField = pFeatLyrSrc.DisplayField
'        Set pFeatLyr.FeatureClass = pFeatLyrSrc.FeatureClass
'        pFeatLyr.ScaleSymbols = pFeatLyrSrc.ScaleSymbols
'        pFeatLyr.Selectable = pFeatLyrSrc.Selectable
'        Set pGeoFeatLyrSource = pFeatLyrSrc
'        Set pGeoFeatLyrDestination = pFeatLyr
'        With pGeoFeatLyrSource
'                                            'Set pGeoFeatLyrDestination.CurrentMapLevel = .CurrentMapLevel
'                                            'pGeoFeatLyrDestination.DisplayFeatureClass = .DisplayFeatureClass
'          pGeoFeatLyrDestination.AnnotationProperties = .AnnotationProperties
'          pGeoFeatLyrDestination.AnnotationPropertiesID = .AnnotationPropertiesID
'          pGeoFeatLyrDestination.DisplayAnnotation = .DisplayAnnotation
'          Set pGeoFeatLyrDestination.ExclusionSet = .ExclusionSet
'          Set pGeoFeatLyrDestination.Renderer = .Renderer
'        End With
'        Set pLayer = pFeatLyr
'
'      ElseIf TypeOf pMxDoc.FocusMap.Layer(i) Is ITinLayer Then
'                                            'pTinLyr.RendererCount = pTinLyrSrc.RendererCount
'        Set pTinLyrSrc = pLayerSrc
'        Set pTinLyr = New TinLayer
'        Set pTinLyr.Dataset = pTinLyrSrc.Dataset
'        pTinLyr.DisplayField = pTinLyrSrc.DisplayField
'        pTinLyr.ScaleSymbols = pTinLyrSrc.ScaleSymbols
'        Set pLayer = pTinLyr
'      ElseIf TypeOf pMxDoc.FocusMap.Layer(i) Is IRasterLayer Then
'                                            'pRastLyr.BandCount = pRastLyrSrc.BandCount
'                                            'pRastLyr.ColumnCount = pRastLyrSrc.ColumnCount
'                                            'pRastLyr.DataFrameExtent = pRastLyrSrc.DataFrameExtent
'                                            'pRastLyr.FilePath = pRastLyrSrc.FilePath
'                                            'pRastLyr.Raster = pRastLyrSrc.Raster
'                                            'pRastLyr.RowCount = pRastLyrSrc.RowCount
'        Set pRastLyr = New RasterLayer
'        Set pRastLyrSrc = pLayerSrc
'        pRastLyr.DisplayResolutionFactor = pRastLyrSrc.DisplayResolutionFactor
'        pRastLyr.PrimaryField = pRastLyrSrc.PrimaryField
'        pRastLyr.PyramidPresent = pRastLyrSrc.PyramidPresent
'        Set pRastLyr.Renderer = pRastLyrSrc.Renderer
'        pRastLyr.ShowResolution = pRastLyrSrc.ShowResolution
'        pRastLyr.VisibleExtent = pRastLyrSrc.VisibleExtent
'        Set pLayer = pRastLyr
'      End If
'                                            'Set pLayer.AreaOfInterest = pLayerSrc.AreaOfInterest
'                                            'Set pLayer.SpatialReference = pLayerSrc.SpatialReference
'                                            'pLayer.SupportedDrawPhases = pLayerSrc.SupportedDrawPhases
'                                            'pLayer.TipText = pLayerSrc.TipText
'                                            'pLayer.Valid = pLayerSrc.Valid
'      pLayer.Cached = pLayerSrc.Cached
'      pLayer.MaximumScale = pLayerSrc.MaximumScale
'      pLayer.MinimumScale = pLayerSrc.MinimumScale
'      pLayer.Name = pLayerSrc.Name
'      pLayer.ShowTips = pLayerSrc.ShowTips
'      pLayer.Visible = pLayerSrc.Visible
'
'      pNewMap.AddLayer pLayer
'    Next i
'  End If
'  SetVisibleLayers pNewMap, sLayers
'                                                  'Create a new MapFrame and associate
'                                                  'map with it
'  Set pMapFrame = New MapFrame
'  Set pMapFrame.Map = pNewMap
'  pMapFrame.ExtentType = esriExtentDefault
'
'  Set pGraphicsContainer = pMxDoc.PageLayout
'                                                  'Set the position of the new map frame
'  Set pElement = pMapFrame
'  Set pEnv = New envelope
'  Set pMxDocDataViewAV = pMxDoc.ActiveView
'  Set pMxDocFocusMapQIDataViewAV = pMxDoc.FocusMap
'
'  bIsCircular = ((2 * dRadius) = dWidthOrigin)
'
'
'  Dim dBubbleRadius As Double, dBubbleWidth As Double
'  Dim dBubbleLeft As Double, dBubbleRight As Double
'  Dim dBubbleTop As Double, dBubbleBottom As Double
'  Dim dScreenRightInches As Double, dScreenBottomInches As Double
'  Dim lBubbleTop As Long, lBubbleBottom As Long
'  Dim lBubbleLeft As Long, lBubbleRight As Long
'  Dim lBubbleTopMapAV As Long, lBubbleBottomMapAV As Long
'  Dim lBubbleLeftMapAV As Long, lBubbleRightMapAV As Long
'  Dim pMapEnv As IEnvelope, pBubbleEnv As IEnvelope
'  Dim pMap As IMap, pActiveView As IActiveView
'
'
'  dBubbleRadius = dRadius * dScale
'  dBubbleWidth = dWidthOrigin * dScale
'                                                  'generate the top y
'                                                  'generate the bottom y
'                                                  'the left x and right x
'                                                  'for detail insets
'  dBubbleTop = dYDestination + dBubbleRadius
'  dBubbleBottom = dYDestination - dBubbleRadius
'  If bIsCircular Then
'    dBubbleLeft = dXDestination - dBubbleRadius
'    dBubbleRight = dXDestination + dBubbleRadius
'  Else
'    dBubbleLeft = dXDestination - (dBubbleWidth / 2)
'    dBubbleRight = dXDestination + (dBubbleWidth / 2)
'  End If
'
'  Set pBubbleEnv = New envelope
'  Set pMapEnv = pMxDoc.ActiveView.Extent 'sides of display in screen units
'
'
'
'  'Shift to layout view
'  '--------------------
'  Set pMxDoc.ActiveView = pMxDoc.PageLayout
'  Set pMapEnv = pMxDoc.ActiveView.Extent
'  Set pMxDocLayoutViewAV = pMxDoc.ActiveView
'  Set pMxDocFocusMapQILayoutViewAV = pMxDoc.FocusMap
'
'
'  'Size and place the detail inset data frame
'  '------------------------------------------
'  Dim dDocWidthInches As Double, dDocHeightInches As Double
'  Dim pPntScrBottomLeft As IPoint, pPntScrTopRight As IPoint
'  Dim dScrMapUnitWidth As Double, dScrMapUnitHeight As Double
'  Dim dScrMapUnitWidthWOBorders As Double, dScrMapUnitHeightWOBorders As Double
'  Dim dRatioInchToMapUnitX As Double, dRatioInchToMapUnitY As Double
'  Dim dBubbleLeftInches As Double, dBubbleRightInches As Double
'  Dim dBubbleTopInches As Double, dBubbleBottomInches As Double
'  Dim dBoundaryInchesXLeft As Double, dBoundaryInchesYTop As Double
'
'  With pMapEnv
'    dDocWidthInches = .XMax - (Abs(.XMin))
'    dDocHeightInches = .YMax - (Abs(.YMin))
'  End With
'                                                  'get the data area width minus the
'                                                  'buffer of empty space around the
'                                                  'map display
'  dDataFrameWidth = pElementWithinBorders.Geometry.envelope.Width
'  dDataFrameHeight = pElementWithinBorders.Geometry.envelope.Height
'  dBoundaryInchesXLeft = pElementWithinBorders.Geometry.envelope.XMin
'  dBoundaryInchesYTop = pElementWithinBorders.Geometry.envelope.YMin
'
'  With pMxDocFocusMapQILayoutViewAV.ScreenDisplay.DisplayTransformation
'    Set pPntScrBottomLeft = .ToMapPoint(.DeviceFrame.Left, .DeviceFrame.bottom)
'    Set pPntScrTopRight = .ToMapPoint(.DeviceFrame.Right, .DeviceFrame.Top)
'  End With
'
'  dScrMapUnitWidth = pPntScrTopRight.x - pPntScrBottomLeft.x
'  dScrMapUnitHeight = pPntScrTopRight.y - pPntScrBottomLeft.y
'  dRatioInchToMapUnitX = dDataFrameWidth / dScrMapUnitWidth
'  dRatioInchToMapUnitY = dDataFrameHeight / dScrMapUnitHeight
'
'  dBubbleLeftInches = ((dBubbleLeft - pPntScrBottomLeft.x) * dRatioInchToMapUnitX) + dBoundaryInchesXLeft
'  dBubbleRightInches = ((dBubbleRight - pPntScrBottomLeft.x) * dRatioInchToMapUnitX) + dBoundaryInchesXLeft
'  dBubbleTopInches = ((dBubbleTop - pPntScrBottomLeft.y) * dRatioInchToMapUnitY) + dBoundaryInchesYTop
'  dBubbleBottomInches = ((dBubbleBottom - pPntScrBottomLeft.y) * dRatioInchToMapUnitY) + dBoundaryInchesYTop
'
'  pBubbleEnv.XMin = dBubbleLeftInches
'  pBubbleEnv.XMax = dBubbleRightInches
'  pBubbleEnv.YMin = dBubbleBottomInches
'  pBubbleEnv.YMax = dBubbleTopInches
'
'
'
'  pElement.Geometry = pBubbleEnv
'
'  '''''''''''''''''''''''
'  ' data frame properties
'
'  Dim pColor As IColor, pFillColor As IColor, pShadowColor As IColor
'  Set pColor = New RgbColor
'  Set pFillColor = New RgbColor
'  Set pShadowColor = New RgbColor
'  pColor.RGB = RGB(210, 210, 210)
'  pFillColor.RGB = RGB(255, 255, 255)
'  pShadowColor.RGB = RGB(128, 128, 128)
'
'  Set pFrameElement = pElement
'
'  ' create a border
'  Dim pSymbolBorder As ISymbolBorder
'  Dim pLineSymbol As ILineSymbol
'  Dim pFrameDecoration As IFrameDecoration
'  Dim pShadowFillSymbol As IFillSymbol
'  Dim pSymbolShadow As ISymbolShadow
'  Dim pFrameProperties As IFrameProperties
'
'  Set pSymbolBorder = New SymbolBorder
'  Set pLineSymbol = New SimpleLineSymbol
'  pLineSymbol.Color = pColor
'  pSymbolBorder.LineSymbol = pLineSymbol
'  pSymbolBorder.LineSymbol.Color = pColor
'  pSymbolBorder.CornerRounding = 100
'  pFrameElement.Border = pSymbolBorder
'
'  'modify the frame element background
'  Set pFrameDecoration = New SymbolBackground
'  pFrameDecoration.Color = pFillColor
'  pFrameDecoration.CornerRounding = 100
'
'  pFrameElement.Background = pFrameDecoration
'
'
'  ' add shadow to detail inset
'  Set pShadowFillSymbol = New SimpleFillSymbol
'  pShadowFillSymbol.Color = pShadowColor
'  pShadowFillSymbol.Outline.Color = pShadowColor
'  Set pSymbolShadow = New SymbolShadow
'  pSymbolShadow.FillSymbol = pShadowFillSymbol
'  pSymbolShadow.HorizontalSpacing = -2
'  pSymbolShadow.VerticalSpacing = -2
'  pSymbolShadow.CornerRounding = 100
'  Set pFrameProperties = pFrameElement
'  pFrameProperties.Shadow = pSymbolShadow
'
'
'  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  'Build triangle and shadow triangle to point from the bubble inset to the
'  '(presumably) circular polygon feature that represents the area in the inset
'  '
'  '  - Find the points for the triangle graphic to be created.
'  '           - convert the origin and destination points to screen units.
'  '           - convert the radius to screen units
'  '           - take an angle 10 degrees from the line in each direction
'  '       - create a triangle polygon from those points
'  '       - add the triangle to the graphic layer
'
'
'  Dim pLine As ILine, pFromPnt As IPoint, pToPnt As IPoint
'  Dim pCircularArc As ICircularArc, pBaseLine As ILine
'  Dim p1stPnt As IPoint, p2ndPnt As IPoint, p3rdPnt As IPoint
'  Dim p1stShadowPnt As IPoint, p2ndShadowPnt As IPoint, p3rdShadowPnt As IPoint
'  Dim pConstructPoint As IConstructPoint
'
'  Set pLine = New esriGeometry.Line
'  Set pFromPnt = New Point
'  Set pToPnt = New Point
'  pFromPnt.x = dXDestination
'  pFromPnt.y = dYDestination
'  pToPnt.x = dXOrigin
'  pToPnt.y = dYOrigin
'  pLine.PutCoords pFromPnt, pToPnt
'
'
'  Set pCircularArc = New CircularArc
'                                                  'angles are stored in radians,
'                                                  'so calculate 10 degrees in radians
'  pCircularArc.PutCoordsByAngle pFromPnt, _
'                               (pLine.Angle - ((10 / 180) * 3.14159265358979)), _
'                               ((20 / 180) * 3.14159265358979), _
'                               dBubbleRadius
'
'  Set pBaseLine = New esriGeometry.Line
'  pBaseLine.PutCoords pCircularArc.FromPoint, pCircularArc.ToPoint
'  Set p1stPnt = New Point
'  Set p2ndPnt = New Point
'  Set p3rdPnt = New Point
'  Set p1stShadowPnt = New Point
'  Set p2ndShadowPnt = New Point
'  Set p3rdShadowPnt = New Point
'  Set pConstructPoint = p3rdPnt
'  pConstructPoint.ConstructDeflection pBaseLine, _
'                                      pBaseLine.Length, _
'                                      -((60 / 180) * 3.14159265358979)
'                                                  '3 points are now available for triangle:
'                                                  ' - pCircularArc.FromPoint
'                                                  ' - pCircularArc.ToPoint,
'                                                  ' - p3rdPnt
'  Dim pTrianglePoly As IPolygon, pGeomColl As IGeometryCollection, pGeometry As IGeometry
'  Dim pShadowTrianglePoly As IPolygon, pShadowPolygonElement As IPolygonElement
'  Dim pPntColl As IPointCollection, pFeature As IFeature
'  Dim pPolygonElement As IPolygonElement, pFillShapeElement As IFillShapeElement
'  Dim pShadowPolyElement As IPolygonElement, pShadowFillShapeElement As IFillShapeElement
'  Dim pArrowFillSymb As IFillSymbol, pElementPly As IElement, pElementShadowPly As IElement
'
'  '''''''''''''''''''''''''''''''
'  'triangle colors and dimensions
'
'  Set pArrowFillSymb = New SimpleFillSymbol
'  Set pPolygonElement = New PolygonElement
'  Set pFillShapeElement = pPolygonElement
'  pArrowFillSymb.Outline = pLineSymbol
'  pArrowFillSymb.Color = pLineSymbol.Color
'  pFillShapeElement.Symbol = pArrowFillSymb
'  Set pElementPly = pPolygonElement
'
'  Set pTrianglePoly = New esriGeometry.Polygon
'  Set pPntColl = pTrianglePoly
'  p1stPnt.x = ((pCircularArc.FromPoint.x - pPntScrBottomLeft.x) * dRatioInchToMapUnitX) + dBoundaryInchesXLeft
'  p1stPnt.y = ((pCircularArc.FromPoint.y - pPntScrBottomLeft.y) * dRatioInchToMapUnitX) + dBoundaryInchesYTop
'  p2ndPnt.x = ((pCircularArc.ToPoint.x - pPntScrBottomLeft.x) * dRatioInchToMapUnitX) + dBoundaryInchesXLeft
'  p2ndPnt.y = ((pCircularArc.ToPoint.y - pPntScrBottomLeft.y) * dRatioInchToMapUnitX) + dBoundaryInchesYTop
'  p3rdPnt.x = ((p3rdPnt.x - pPntScrBottomLeft.x) * dRatioInchToMapUnitX) + dBoundaryInchesXLeft
'  p3rdPnt.y = ((p3rdPnt.y - pPntScrBottomLeft.y) * dRatioInchToMapUnitX) + dBoundaryInchesYTop
'  pPntColl.AddPoint p1stPnt
'  pPntColl.AddPoint p2ndPnt
'  pPntColl.AddPoint p3rdPnt
'  pTrianglePoly.Close
'
'  Set pGeometry = pTrianglePoly
'  pElementPly.Geometry = pGeometry
'
'  ''''''''''''''''''''''''''''''''''''''
'  'triangle shadow colors and dimensions
'
'  Set pShadowPolyElement = New PolygonElement
'  Set pShadowFillShapeElement = pShadowPolyElement
'  Set pFillShapeElement = pShadowPolyElement
'  pFillShapeElement.Symbol = pShadowFillSymbol
'  Set pElementShadowPly = pShadowPolyElement
'
'  Set pShadowTrianglePoly = New esriGeometry.Polygon
'  Set pPntColl = pShadowTrianglePoly
'                                                  'offset shadow triangle by 3 pixels
'  p1stShadowPnt.x = p1stPnt.x - (ConvertPixelsToRW(2, pApp) * dRatioInchToMapUnitX)
'  p1stShadowPnt.y = p1stPnt.y - (ConvertPixelsToRW(2, pApp) * dRatioInchToMapUnitY)
'  p2ndShadowPnt.x = p2ndPnt.x - (ConvertPixelsToRW(2, pApp) * dRatioInchToMapUnitX)
'  p2ndShadowPnt.y = p2ndPnt.y - (ConvertPixelsToRW(2, pApp) * dRatioInchToMapUnitY)
'  p3rdShadowPnt.x = p3rdPnt.x - (ConvertPixelsToRW(2, pApp) * dRatioInchToMapUnitX)
'  p3rdShadowPnt.y = p3rdPnt.y - (ConvertPixelsToRW(2, pApp) * dRatioInchToMapUnitY)
'  pPntColl.AddPoint p1stShadowPnt
'  pPntColl.AddPoint p2ndShadowPnt
'  pPntColl.AddPoint p3rdShadowPnt
'  pShadowTrianglePoly.Close
'  Set pGeometry = pShadowTrianglePoly
'  pElementShadowPly.Geometry = pGeometry
'
'  'tag the graphic elements for later tracking
'  Dim pElementProps As IElementProperties
'  Set pElementProps = pElementShadowPly
'  pElementProps.CustomProperty = "BubbleID:" & lBubbleId
'  Set pElementProps = pElement
'  pElementProps.CustomProperty = "BubbleID:" & lBubbleId
'  Set pElementProps = pElementPly
'  pElementProps.CustomProperty = "BubbleID:" & lBubbleId
'
'
'
'  ''''''''''''''''''''''''''''''
'  'Add triangle shadow to layout
'  pGraphicsContainer.AddElement pElementShadowPly, 0
'
'  '''''''''''''''''''''''''''
'  'Add mapframe to the layout
'  pGraphicsContainer.AddElement pElement, 0
'
'  '''''''''''''''''''''''
'  'Add triangle to layout
'  pGraphicsContainer.AddElement pElementPly, 0
'
'
'  Set pActiveView = pNewMap
'
'  '''''''''''''''''''''''''''''''''''
'  'Set the detail inset's zoom extent
'  Set pNewEnv = New envelope
'
'  If bIsCircular Then
'    pNewEnv.XMin = dXOrigin - dRadius
'    pNewEnv.XMax = dXOrigin + dRadius
'  Else
'    pNewEnv.XMin = dXOrigin - (dWidthOrigin / 2)
'    pNewEnv.XMax = dXOrigin + (dWidthOrigin / 2)
'  End If
'  pNewEnv.YMin = dYOrigin - dRadius
'  pNewEnv.YMax = dYOrigin + dRadius
'
'  pActiveView.Extent = pNewEnv
'  pActiveView.Refresh
'
'  'Set pGraphicsLayer = pMxDoc.FocusMap.BasicGraphicsLayer
'  'Set pCompositeGraphicsLayer = pGraphicsLayer
'  'Set pGraphicsLayer.AssociatedLayer = pFeatureLayer
'
'  'Set pGraphicsLayer = pCompositeGraphicsLayer.FindLayer("NW Detail Inset Arrows")
'  'If pGraphicsLayer Is Nothing Then
'  '  Set pGraphicsLayer = pCompositeGraphicsLayer.AddLayer("NW Detail Inset Arrows", pFeatureLayer)
'  'End If
'
'  Set pActiveView = pMxDoc.FocusMap
'  pActiveView.PartialRefresh esriViewGraphics, Nothing, Nothing
'  pMxDoc.CurrentContentsView.Refresh Nothing
'
'  pNWSeriesOptions.BubbleGraphicAdd pElement, pElementPly, pElementShadowPly, pNewMap.Name
'
'  Exit Sub
'ErrorHandler:
'  HandleError False, "AddDataframe " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
'End Sub
'
'
'
'
'
'
'
'
'Private Function ConvertPixelsToRW(pixelUnits As Double, pApp As IApplication) As Double
'  On Error GoTo ErrorHandler
'
'  Dim pMxDoc As IMxDocument
'  Dim realWorldDisplayExtent As Double
'  Dim pixelExtent As Long
'  Dim sizeOfOnePixel As Double
'  Dim pDT As IDisplayTransformation
'  Dim deviceRECT As tagRECT
'  Dim pEnv As IEnvelope
'  Dim pActiveView As IActiveView
'
'  Set pMxDoc = pApp.Document
'  Set pActiveView = pMxDoc.FocusMap
'  Set pDT = pActiveView.ScreenDisplay.DisplayTransformation
'  deviceRECT = pDT.DeviceFrame
'  pixelExtent = deviceRECT.Right - deviceRECT.Left
'  Set pEnv = pDT.VisibleBounds
'  realWorldDisplayExtent = pEnv.Width
'  sizeOfOnePixel = realWorldDisplayExtent / pixelExtent
'  ConvertPixelsToRW = pixelUnits * sizeOfOnePixel
'
'  Exit Function
'
'ErrorHandler:
'   MsgBox "Error " & Err.Number & ": " & Err.Description & vbNewLine _
'       & "In " & Err.Source & " at DrawSelectedArrows.ConvertPixelsToRW", vbCritical
'End Function
'


