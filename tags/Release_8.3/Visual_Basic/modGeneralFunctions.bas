Attribute VB_Name = "modGeneralFunctions"
Option Explicit

Public Const c_DefaultFld_Shape = "SHAPE"
Public Const cPI = 3.14159265358979

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
                Set pP = New Point
                Set GetValidExtentForLayer = New Envelope
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
  If fs.fileexists(pTruncPath & ".shp") Or fs.fileexists(pTruncPath & ".dbf") Or _
   fs.fileexists(pTruncPath & ".shx") Then
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
    pOutShpWspName.WorkspaceFactoryProgID = "esriCore.AccessWorkspaceFactory"
    Set pName = pOutShpWspName
    Set pShapeWorkspace = pName.Open
i = 1
    'Open the dataset
    Set pFeatureWorkspace = pShapeWorkspace
    Set pDataset = pFeatureWorkspace.OpenFeatureDataset(pNewDataSet)
i = 2
    ' Add the SHAPE field (based on the dataset)
    Set pFieldsEdit = pMoreFields
    Set pField = New esriCore.Field
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
    pOutShpWspName.WorkspaceFactoryProgID = "esriCore.ShapefileWorkspaceFactory.1"
    Set pName = pOutShpWspName
    Set pShapeWorkspace = pName.Open
    ' Add the SHAPE field (based on the Map)
    Set pFieldsEdit = pMoreFields
    Set pField = New esriCore.Field
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

Public Sub TurnOffClipping(pSeriesProps As IDSMapSeriesProps, pApp As IApplication)
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
  Do While Not pElem Is Nothing
    If TypeOf pElem Is IMapFrame Then
      Set pMapFrame = pElem
      If pMapFrame.Map.Name = "Local Indicator" Or _
       pMapFrame.Map.Name = "Global Indicator" Then
        pDelColl.Add pMapFrame
      End If
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
    If pElemProps.Name = "DSMAPBOOK TEXT" Then
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

Public Function GetMapBookExtension(pApp As IApplication) As IDSMapBook
On Error GoTo ErrHand:
  Dim pMapBookExt As DSMapBookExt, pMapBook As IDSMapBook
  Set pMapBookExt = pApp.FindExtensionByName("DevSample_MapBook")
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
      If UCase(pElemProps.Name) = "DSMAPBOOK CLIP ELEMENT" Then
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

