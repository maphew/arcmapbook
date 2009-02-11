Attribute VB_Name = "modGeneralFunctions"

' Copyright 2008 ESRI
'
' All rights reserved under the copyright laws of the United States
' and applicable international laws, treaties, and conventions.
'
' You may freely redistribute and use this sample code, with or
' without modification, provided you include the original copyright
' notice and use restrictions.
'
' See use restrictions at <your ArcGIS install location>/developerkit/userestrictions.txt.


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
22:     End Select
End Function

Public Function GetActiveDataFrameName(pApp As IApplication) As String
    Dim pMx As IMxDocument
    Dim pFocusMap As IMap
    
29:     Set pMx = pApp.Document
30:     Set pFocusMap = pMx.FocusMap
    
32:     GetActiveDataFrameName = pFocusMap.Name
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
47:     Set pMx = pApp.Document
    ' Loop through the elements (in the layout)
49:     Set pGraphicsContainer = pMx.PageLayout
50:     pGraphicsContainer.Reset
51:     Set pElement = pGraphicsContainer.Next
52:     While Not pElement Is Nothing
        ' If type of element is an IFrameElement
54:         If TypeOf pElement Is IFrameElement Then
55:             Set pElProps = pElement
            ' If Name matches
57:             If UCase(pElProps.Name) = UCase(sDataFramName) Then
                ' Return element
59:                 Set GetDataFrameElement = pElement
60:                 Set pElement = Nothing
61:             Else
62:                 Set pElement = pGraphicsContainer.Next
63:             End If
64:         Else
65:             Set pElement = pGraphicsContainer.Next
66:         End If
67:     Wend
    
    Exit Function
ErrorHandler:
71:     Err.Raise Err.Number, Err.Source, "Error in GetDataFrameElement:" _
        & vbCrLf & Err.Description
End Function

Public Function FindFeatureLayerByDS(DatasetName As String, pApp As IApplication) As IFeatureLayer
  
    On Error GoTo ErrorHandler
  
    Dim pMxDoc As IMxDocument
    Dim pMap As IMap
    Dim pFeatureLayer As IFeatureLayer
    Dim pDataset As IDataset
    Dim i As Integer
    
85:     Set pMxDoc = pApp.Document
86:     Set pMap = pMxDoc.FocusMap
  
88:     With pMap
89:         For i = 0 To .LayerCount - 1
90:             If TypeOf .Layer(i) Is IFeatureLayer Then
91:                 Set pFeatureLayer = .Layer(i)
92:                 Set pDataset = pFeatureLayer.FeatureClass
93:                 If UCase(pDataset.Name) = UCase(DatasetName) Then
94:                     Set FindFeatureLayerByDS = pFeatureLayer
95:                     Exit For
96:                 End If
97:             End If
98:         Next i
99:     End With
  
101:     If pFeatureLayer Is Nothing Then
102:         Err.Raise vbObjectError, "FindFeatureLayerByDS", "Error in " _
            & "FindFeatureLayerByDS:" & vbCrLf & "Could not locate " _
            & "layer with a dataset name of '" & DatasetName & "'."
105:     End If
  
    Exit Function
ErrorHandler:
109:     Err.Raise Err.Number, Err.Source, "Error in routine: FindFeatureLayerByDS" _
        & vbCrLf & Err.Description
End Function

Public Function FindFeatureLayerByName(FLName As String, pApp As IApplication) As IFeatureLayer
  
    On Error GoTo ErrorHandler
  
    Dim pMxDoc As IMxDocument
    Dim pMap As IMap
    Dim pFeatureLayer As IFeatureLayer
    Dim pDataset As IDataset
    Dim i As Integer
    
123:     Set pMxDoc = pApp.Document
124:     Set pMap = pMxDoc.FocusMap
  
126:     With pMap
127:         For i = 0 To .LayerCount - 1
128:             If TypeOf .Layer(i) Is IFeatureLayer Then
129:                 Set pFeatureLayer = .Layer(i)
130:                 If UCase(pFeatureLayer.Name) = UCase(FLName) Then
131:                     Set FindFeatureLayerByName = pFeatureLayer
132:                     Exit For
133:                 End If
134:             End If
135:         Next i
136:     End With
  
138:     If pFeatureLayer Is Nothing Then
139:         Err.Raise vbObjectError, "FindFeatureLayerByName", "Error in " _
            & "FindFeatureLayerByName:" & vbCrLf & "Could not locate " _
            & "layer with a Name of '" & FLName & "'."
142:     End If
  
    Exit Function
ErrorHandler:
146:     Err.Raise Err.Number, Err.Source, "Error in routine: FindFeatureLayerByName" _
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
    
161:     If Not pFL Is Nothing Then
162:         If Not pFL.FeatureClass Is Nothing Then
163:             If TypeOf pFL.FeatureClass Is IGeoDataset Then
164:                 If pFL.FeatureClass.FeatureDataset Is Nothing Then
165:                     dX1 = -1000000000
166:                     dY1 = -1000000000
167:                     dX2 = 1000000000
168:                     dY2 = 1000000000
169:                 Else
170:                     Set pW = pFL.FeatureClass.FeatureDataset.Workspace
171:                     Set pWSR = pW
172:                     Set pEnumSRI = pWSR.SpatialReferenceInfo
173:                     Set pSR = pEnumSRI.Next(0)
174:                     pSR.GetDomain dX1, dX2, dY1, dY2
175:                 End If
176:                 Set pP = New esrigeometry.Point
177:                 Set GetValidExtentForLayer = New Envelope
178:                 pP.PutCoords dX1, dY1
179:                 GetValidExtentForLayer.LowerLeft = pP
180:                 pP.PutCoords dX2, dY2
181:                 GetValidExtentForLayer.UpperRight = pP
182:             Else
183:                 Err.Raise vbObjectError, "GetValidExtentForLayer", _
                    "The 'FeatureClass' property for the IFeatureLayer parameter is not an IGeoDataset"
185:             End If
186:         Else
187:             Err.Raise vbObjectError, "GetValidExtentForLayer", _
                "The IFeatureLayer parameter does not have a valid FeatureClass property"
189:         End If
190:     Else
191:         Err.Raise vbObjectError, "GetValidExtentForLayer", _
            "The IFeatureLayer parameter is set to Nothing"
193:     End If
End Function

Public Function DoesShapeFileExist(pPath As String) As Boolean
  Dim pTruncPath As String
198:   If InStr(1, pPath, ".shp") > 0 Then
199:     pTruncPath = Left(pPath, InStr(1, pPath, ".shp") - 1)
200:   Else
201:     pTruncPath = pPath
202:   End If
      
  'Make sure the specified file does not exist
  Dim fs As Object
206:   Set fs = CreateObject("Scripting.FileSystemObject")
207:   If fs.fileexists(pTruncPath & ".shp") Or fs.fileexists(pTruncPath & ".dbf") Or _
   fs.fileexists(pTruncPath & ".shx") Then
209:     DoesShapeFileExist = True
210:   Else
211:     DoesShapeFileExist = False
212:   End If
End Function

Private Function DoesFeatureClassExist(location As IGxObject, newObjectName As String) As Boolean
On Error GoTo ErrHand:
  Dim pFeatClass As IFeatureClass
  Dim pFeatDataset As IGxDataset
219:   Set pFeatDataset = location
  Dim pFeatClassCont As IFeatureClassContainer, pData As IFeatureDataset
221:   Set pData = pFeatDataset.Dataset
222:   Set pFeatClassCont = pData
  Dim pEnumClass As IEnumFeatureClass, pDataset As IDataset
224:   Set pEnumClass = pFeatClassCont.Classes
225:   Set pFeatClass = pEnumClass.Next
226:   While Not pFeatClass Is Nothing
227:     Set pDataset = pFeatClass
228:     If UCase(pDataset.Name) = UCase(newObjectName) Then
229:       DoesFeatureClassExist = True
      Exit Function
231:     End If
      
233:     Set pFeatClass = pEnumClass.Next
234:   Wend
235:   DoesFeatureClassExist = False
  
  Exit Function
ErrHand:
239:   MsgBox Err.Description
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
  
265:     Set pOutShpWspName = New WorkspaceName

267:     pOutShpWspName.PathName = pDatabase
268:     pOutShpWspName.WorkspaceFactoryProgID = "esriDataSourcesGDB.AccessWorkspaceFactory"
269:     Set pName = pOutShpWspName
270:     Set pShapeWorkspace = pName.Open
271: i = 1
    'Open the dataset
273:     Set pFeatureWorkspace = pShapeWorkspace
274:     Set pDataset = pFeatureWorkspace.OpenFeatureDataset(pNewDataSet)
275: i = 2
    ' Add the SHAPE field (based on the dataset)
277:     Set pFieldsEdit = pMoreFields
278:     Set pField = New Field
279:     Set newFieldEdit = pField
280:     With newFieldEdit
281:         .Name = c_DefaultFld_Shape
282:         .Type = esriFieldTypeGeometry
283:         .IsNullable = True
284:         .Editable = True
285:     End With
286:     Set pGeomDef = New GeometryDef
287:     Set pGeomDefEdit = pGeomDef
288:     With pGeomDefEdit
289:         .GeometryType = esriGeometryPolygon
290:         If TypeOf pDataset Is IGeoDataset Then
291:             Set pGeoDataset = pDataset
292:             Set .SpatialReference = pGeoDataset.SpatialReference
293:         Else
294:             Set .SpatialReference = New UnknownCoordinateSystem
295:         End If
296:         .GridCount = 1
297:         .GridSize(0) = 200
298:         .HasM = False
299:         .HasZ = False
300:         .AvgNumPoints = 4
301:     End With
302:     Set newFieldEdit.GeometryDef = pGeomDef
303:     pFieldsEdit.AddField pField
    ' Check the fields
305:     Set pFieldChecker = New FieldChecker
306:     Set pFieldChecker.ValidateWorkspace = pShapeWorkspace
307:     Set pNewFields = pMoreFields
308: i = 3
309:     Set pClone = pNewFields
310:     Set pCloneFields = pClone.Clone
311:     pFieldChecker.Validate pCloneFields, pErrorEnum, pOutputFields
      
  ' Create the output featureclass
  Dim pUID As New UID
315:   pUID = "{52353152-891A-11D0-BEC6-00805F7C4268}"
316:     shapeFieldName = c_DefaultFld_Shape
317: i = 4
318:     Set pNewFeatClass = pDataset.CreateFeatureClass(pNewFile, pOutputFields, pUID, Nothing, esriFTSimple, shapeFieldName, "")
319: i = 5
320:     Set NewAccessFile = pNewFeatClass
  
    Exit Function
  
ErrorHandler:
325:     MsgBox Err.Number & " " & Err.Description, vbCritical, "Error in NewAccessFile " & i
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
350:     Set pOutShpWspName = New WorkspaceName
351:     pOutShpWspName.PathName = EntryName(pNewFile)
352:     pOutShpWspName.WorkspaceFactoryProgID = "esriDataSourcesFile.ShapefileWorkspaceFactory.1"
353:     Set pName = pOutShpWspName
354:     Set pShapeWorkspace = pName.Open
    ' Add the SHAPE field (based on the Map)
356:     Set pFieldsEdit = pMoreFields
357:     Set pField = New Field
358:     Set newFieldEdit = pField
359:     newFieldEdit.Name = c_DefaultFld_Shape
360:     newFieldEdit.Type = esriFieldTypeGeometry
361:     Set pGeomDef = New GeometryDef
362:     Set pGeomDefEdit = pGeomDef
363:     With pGeomDefEdit
364:         .GeometryType = esriGeometryPolygon
365:         Set .SpatialReference = pMap.SpatialReference
366:     End With
367:     Set newFieldEdit.GeometryDef = pGeomDef
368:     pFieldsEdit.AddField pField
    ' Validate field names
370:     Set pFieldChecker = New FieldChecker
371:     Set pFieldChecker.ValidateWorkspace = pShapeWorkspace
372:     Set pNewFields = pMoreFields
373:     Set pClone = pNewFields
374:     Set pCloneFields = pClone.Clone
375:     pFieldChecker.Validate pCloneFields, pErrorEnum, pOutputFields
    ' Create the output featureclass
377:     shapeFieldName = c_DefaultFld_Shape
378:     featureclassName = Mid(pNewFile, Len(pOutShpWspName.PathName) + 2)
379:     Set pFeatureWorkspace = pShapeWorkspace
380:     Set pNewFeatClass = pFeatureWorkspace.CreateFeatureClass(featureclassName, pOutputFields, _
                            Nothing, Nothing, esriFTSimple, shapeFieldName, "")
    ' Return
383:     Set NewShapeFile = pNewFeatClass
  
    Exit Function
  
ErrorHandler:
388:     MsgBox "Error creating " & pNewFile & vbCrLf & Err.Number & ": " & Err.Description, _
        vbCritical, "Error in NewShapefile"
End Function

Public Function EntryName(sFile As String) As String
  ' work from the right side to the first file delimeter
  Dim iLength As Integer
395:   iLength = Len(sFile)
  Dim iCounter As Integer
  Dim sDelim As String
398:   sDelim = "\"
  Dim sRight As String
  
401:   For iCounter = iLength To 0 Step -1
    
403:     If Mid$(sFile, iCounter, 1) = sDelim Then
404:       EntryName = Mid$(sFile, 1, (iCounter - 1))
405:       Exit For
406:     End If
  
408:   Next
  
End Function

Public Sub TurnOffClipping(pSeriesProps As IDSMapSeriesProps, pApp As IApplication)
On Error GoTo ErrHand:
  Dim pMap As IMap, pDoc As IMxDocument
  'Find the data frame
416:   Set pDoc = pApp.Document
417:   Set pMap = FindDataFrame(pDoc, pSeriesProps.DataFrameName)
  If pMap Is Nothing Then Exit Sub
  
420:   pMap.ClipGeometry = Nothing

  Exit Sub
ErrHand:
424:   MsgBox "TurnOffClipping - " & Err.Description
End Sub

Public Sub RemoveIndicators(pApp As IApplication)
On Error GoTo ErrHand:
  Dim lLoop As Long, pDoc As IMxDocument, pDelColl As Collection
  Dim pPage As IPageLayout, pGraphCont As IGraphicsContainer
  Dim pElem As IElement, pMapFrame As IMapFrame
432:   Set pDoc = pApp.Document
433:   Set pPage = pDoc.PageLayout
434:   Set pDelColl = New Collection
435:   Set pGraphCont = pPage
436:   pGraphCont.Reset
437:   Set pElem = pGraphCont.Next
438:   Do While Not pElem Is Nothing
439:     If TypeOf pElem Is IMapFrame Then
440:       Set pMapFrame = pElem
441:       If pMapFrame.Map.Name = "Local Indicator" Or _
       pMapFrame.Map.Name = "Global Indicator" Then
443:         pDelColl.Add pMapFrame
444:       End If
445:     End If
    
447:     Set pElem = pGraphCont.Next
448:   Loop
  
450:   For lLoop = 1 To pDelColl.count
451:     pGraphCont.DeleteElement pDelColl.Item(lLoop)
452:   Next lLoop

  Exit Sub
ErrHand:
456:   MsgBox "RemoveIndicators - " & Err.Description
End Sub

Public Sub RemoveLabels(pDoc As IMxDocument)
On Error GoTo ErrHand:
  Dim pGraphicsCont As IGraphicsContainer
  Dim pTempColl As Collection, pElemProps As IElementProperties, lLoop As Long
  'Remove any previous neighbor labels.
464:   Set pGraphicsCont = pDoc.PageLayout
465:   pGraphicsCont.Reset
466:   Set pTempColl = New Collection
467:   Set pElemProps = pGraphicsCont.Next
468:   Do While Not pElemProps Is Nothing
469:     If pElemProps.Name = "DSMAPBOOK TEXT" Then
470:       pTempColl.Add pElemProps
471:     End If
472:     Set pElemProps = pGraphicsCont.Next
473:   Loop
474:   For lLoop = 1 To pTempColl.count
475:     pGraphicsCont.DeleteElement pTempColl.Item(lLoop)
476:   Next lLoop
477:   Set pTempColl = Nothing

  Exit Sub
ErrHand:
481:   MsgBox "RemoveLabels - " & Err.Description
End Sub

Public Function GetMapBookExtension(pApp As IApplication) As IDSMapBook
On Error GoTo ErrHand:
  Dim pMapBookExt As DSMapBookExt, pMapBook As IDSMapBook
487:   Set pMapBookExt = pApp.FindExtensionByName("DevSample_MapBook")
488:   If pMapBookExt Is Nothing Then
489:     MsgBox "Map Book code not installed properly!!  Make sure you can access the regsvr32 command" & vbCrLf & _
     "and rerun the _Install.bat batch file!!", , "Map Book Extension Not Found!!!"
491:     Set GetMapBookExtension = Nothing
    Exit Function
493:   End If
  
495:   Set GetMapBookExtension = pMapBookExt.MapBook

  Exit Function
ErrHand:
499:   MsgBox "GetMapBookExtension - " & Err.Description
End Function

Public Sub RemoveClipElement(pDoc As IMxDocument)
On Error GoTo ErrHand:
  Dim pGraphs As IGraphicsContainer, pElemProps As IElementProperties
  
  'Search for an existing clip element and delete it when found
507:   Set pGraphs = pDoc.FocusMap
508:   pGraphs.Reset
509:   Set pElemProps = pGraphs.Next
510:   Do While Not pElemProps Is Nothing
511:     If TypeOf pElemProps Is IPolygonElement Then
512:       If UCase(pElemProps.Name) = "DSMAPBOOK CLIP ELEMENT" Then
513:         pGraphs.DeleteElement pElemProps
514:         Exit Do
515:       End If
516:     End If
517:     Set pElemProps = pGraphs.Next
518:   Loop
519:   pDoc.ActiveView.PartialRefresh esriViewGraphics, Nothing, Nothing

  Exit Sub
ErrHand:
523:   MsgBox "RemoveClipElement - " & Erl & " - " & Err.Description
End Sub

