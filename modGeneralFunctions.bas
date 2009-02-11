Attribute VB_Name = "modGeneralFunctions"

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
35:     End Select
End Function

Public Function GetActiveDataFrameName(pApp As IApplication) As String
    Dim pMx As IMxDocument
    Dim pFocusMap As IMap
    
42:     Set pMx = pApp.Document
43:     Set pFocusMap = pMx.FocusMap
    
45:     GetActiveDataFrameName = pFocusMap.Name
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
60:     Set pMx = pApp.Document
    ' Loop through the elements (in the layout)
62:     Set pGraphicsContainer = pMx.PageLayout
63:     pGraphicsContainer.Reset
64:     Set pElement = pGraphicsContainer.Next
65:     While Not pElement Is Nothing
        ' If type of element is an IFrameElement
67:         If TypeOf pElement Is IFrameElement Then
68:             Set pElProps = pElement
            ' If Name matches
70:             If UCase(pElProps.Name) = UCase(sDataFramName) Then
                ' Return element
72:                 Set GetDataFrameElement = pElement
73:                 Set pElement = Nothing
74:             Else
75:                 Set pElement = pGraphicsContainer.Next
76:             End If
77:         Else
78:             Set pElement = pGraphicsContainer.Next
79:         End If
80:     Wend
    
    Exit Function
ErrorHandler:
84:     Err.Raise Err.Number, Err.source, "Error in GetDataFrameElement:" _
        & vbCrLf & Err.Description
End Function

Public Function FindFeatureLayerByDS(DatasetName As String, pApp As IApplication) As IFeatureLayer
  
    On Error GoTo ErrorHandler
  
    Dim pMxDoc As IMxDocument
    Dim pMap As IMap
    Dim pFeatureLayer As IFeatureLayer
    Dim pDataset As IDataset
    Dim i As Integer
    
98:     Set pMxDoc = pApp.Document
99:     Set pMap = pMxDoc.FocusMap
  
101:     With pMap
102:         For i = 0 To .LayerCount - 1
103:             If TypeOf .Layer(i) Is IFeatureLayer Then
104:                 Set pFeatureLayer = .Layer(i)
105:                 Set pDataset = pFeatureLayer.FeatureClass
106:                 If UCase(pDataset.Name) = UCase(DatasetName) Then
107:                     Set FindFeatureLayerByDS = pFeatureLayer
108:                     Exit For
109:                 End If
110:             End If
111:         Next i
112:     End With
  
114:     If pFeatureLayer Is Nothing Then
115:         Err.Raise vbObjectError, "FindFeatureLayerByDS", "Error in " _
            & "FindFeatureLayerByDS:" & vbCrLf & "Could not locate " _
            & "layer with a dataset name of '" & DatasetName & "'."
118:     End If
  
    Exit Function
ErrorHandler:
122:     Err.Raise Err.Number, Err.source, "Error in routine: FindFeatureLayerByDS" _
        & vbCrLf & Err.Description
End Function

Public Function FindFeatureLayerByName(FLName As String, pApp As IApplication) As IFeatureLayer
  
    On Error GoTo ErrorHandler
  
    Dim pMxDoc As IMxDocument
    Dim pMap As IMap
    Dim pFeatureLayer As IFeatureLayer
    Dim pDataset As IDataset
    Dim i As Integer
    
136:     Set pMxDoc = pApp.Document
137:     Set pMap = pMxDoc.FocusMap
  
139:     With pMap
140:         For i = 0 To .LayerCount - 1
141:             If TypeOf .Layer(i) Is IFeatureLayer Then
142:                 Set pFeatureLayer = .Layer(i)
143:                 If UCase(pFeatureLayer.Name) = UCase(FLName) Then
144:                     Set FindFeatureLayerByName = pFeatureLayer
145:                     Exit For
146:                 End If
147:             End If
148:         Next i
149:     End With
  
151:     If pFeatureLayer Is Nothing Then
152:         Err.Raise vbObjectError, "FindFeatureLayerByName", "Error in " _
            & "FindFeatureLayerByName:" & vbCrLf & "Could not locate " _
            & "layer with a Name of '" & FLName & "'."
155:     End If
  
    Exit Function
ErrorHandler:
159:     Err.Raise Err.Number, Err.source, "Error in routine: FindFeatureLayerByName" _
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
    
174:     If Not pFL Is Nothing Then
175:         If Not pFL.FeatureClass Is Nothing Then
176:             If TypeOf pFL.FeatureClass Is IGeoDataset Then
177:                 If pFL.FeatureClass.FeatureDataset Is Nothing Then
178:                     dX1 = -1000000000
179:                     dY1 = -1000000000
180:                     dX2 = 1000000000
181:                     dY2 = 1000000000
182:                 Else
183:                     Set pW = pFL.FeatureClass.FeatureDataset.Workspace
184:                     Set pWSR = pW
185:                     Set pEnumSRI = pWSR.SpatialReferenceInfo
186:                     Set pSR = pEnumSRI.Next(0)
187:                     pSR.GetDomain dX1, dX2, dY1, dY2
188:                 End If
189:                 Set pP = New esrigeometry.Point
190:                 Set GetValidExtentForLayer = New Envelope
191:                 pP.PutCoords dX1, dY1
192:                 GetValidExtentForLayer.LowerLeft = pP
193:                 pP.PutCoords dX2, dY2
194:                 GetValidExtentForLayer.UpperRight = pP
195:             Else
196:                 Err.Raise vbObjectError, "GetValidExtentForLayer", _
                    "The 'FeatureClass' property for the IFeatureLayer parameter is not an IGeoDataset"
198:             End If
199:         Else
200:             Err.Raise vbObjectError, "GetValidExtentForLayer", _
                "The IFeatureLayer parameter does not have a valid FeatureClass property"
202:         End If
203:     Else
204:         Err.Raise vbObjectError, "GetValidExtentForLayer", _
            "The IFeatureLayer parameter is set to Nothing"
206:     End If
End Function

Public Function DoesShapeFileExist(pPath As String) As Boolean
  Dim pTruncPath As String
211:   If InStr(1, pPath, ".shp") > 0 Then
212:     pTruncPath = Left(pPath, InStr(1, pPath, ".shp") - 1)
213:   Else
214:     pTruncPath = pPath
215:   End If
      
  'Make sure the specified file does not exist
  Dim fs As Object
219:   Set fs = CreateObject("Scripting.FileSystemObject")
220:   If fs.fileexists(pTruncPath & ".shp") Or fs.fileexists(pTruncPath & ".dbf") Or _
   fs.fileexists(pTruncPath & ".shx") Then
222:     DoesShapeFileExist = True
223:   Else
224:     DoesShapeFileExist = False
225:   End If
End Function

Private Function DoesFeatureClassExist(location As IGxObject, newObjectName As String) As Boolean
On Error GoTo ErrHand:
  Dim pFeatClass As IFeatureClass
  Dim pFeatDataset As IGxDataset
232:   Set pFeatDataset = location
  Dim pFeatClassCont As IFeatureClassContainer, pData As IFeatureDataset
234:   Set pData = pFeatDataset.Dataset
235:   Set pFeatClassCont = pData
  Dim pEnumClass As IEnumFeatureClass, pDataset As IDataset
237:   Set pEnumClass = pFeatClassCont.Classes
238:   Set pFeatClass = pEnumClass.Next
239:   While Not pFeatClass Is Nothing
240:     Set pDataset = pFeatClass
241:     If UCase(pDataset.Name) = UCase(newObjectName) Then
242:       DoesFeatureClassExist = True
      Exit Function
244:     End If
      
246:     Set pFeatClass = pEnumClass.Next
247:   Wend
248:   DoesFeatureClassExist = False
  
  Exit Function
ErrHand:
252:   MsgBox Err.Description
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
  
278:     Set pOutShpWspName = New WorkspaceName

280:     pOutShpWspName.PathName = pDatabase
281:     pOutShpWspName.WorkspaceFactoryProgID = "esriDataSourcesGDB.AccessWorkspaceFactory"
282:     Set pName = pOutShpWspName
283:     Set pShapeWorkspace = pName.Open
284: i = 1
    'Open the dataset
286:     Set pFeatureWorkspace = pShapeWorkspace
287:     Set pDataset = pFeatureWorkspace.OpenFeatureDataset(pNewDataSet)
288: i = 2
    ' Add the SHAPE field (based on the dataset)
290:     Set pFieldsEdit = pMoreFields
291:     Set pField = New Field
292:     Set newFieldEdit = pField
293:     With newFieldEdit
294:         .Name = c_DefaultFld_Shape
295:         .Type = esriFieldTypeGeometry
296:         .IsNullable = True
297:         .Editable = True
298:     End With
299:     Set pGeomDef = New GeometryDef
300:     Set pGeomDefEdit = pGeomDef
301:     With pGeomDefEdit
302:         .GeometryType = esriGeometryPolygon
303:         If TypeOf pDataset Is IGeoDataset Then
304:             Set pGeoDataset = pDataset
305:             Set .SpatialReference = pGeoDataset.SpatialReference
306:         Else
307:             Set .SpatialReference = New UnknownCoordinateSystem
308:         End If
309:         .GridCount = 1
310:         .GridSize(0) = 200
311:         .HasM = False
312:         .HasZ = False
313:         .AvgNumPoints = 4
314:     End With
315:     Set newFieldEdit.GeometryDef = pGeomDef
316:     pFieldsEdit.AddField pField
    ' Check the fields
318:     Set pFieldChecker = New FieldChecker
319:     Set pFieldChecker.ValidateWorkspace = pShapeWorkspace
320:     Set pNewFields = pMoreFields
321: i = 3
322:     Set pClone = pNewFields
323:     Set pCloneFields = pClone.Clone
324:     pFieldChecker.Validate pCloneFields, pErrorEnum, pOutputFields
      
  ' Create the output featureclass
  Dim pUID As New UID
328:   pUID = "{52353152-891A-11D0-BEC6-00805F7C4268}"
329:     shapeFieldName = c_DefaultFld_Shape
330: i = 4
331:     Set pNewFeatClass = pDataset.CreateFeatureClass(pNewFile, pOutputFields, pUID, Nothing, esriFTSimple, shapeFieldName, "")
332: i = 5
333:     Set NewAccessFile = pNewFeatClass
  
    Exit Function
  
ErrorHandler:
338:     MsgBox Err.Number & " " & Err.Description, vbCritical, "Error in NewAccessFile " & i
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
363:     Set pOutShpWspName = New WorkspaceName
364:     pOutShpWspName.PathName = EntryName(pNewFile)
365:     pOutShpWspName.WorkspaceFactoryProgID = "esriDataSourcesFile.ShapefileWorkspaceFactory.1"
366:     Set pName = pOutShpWspName
367:     Set pShapeWorkspace = pName.Open
    ' Add the SHAPE field (based on the Map)
369:     Set pFieldsEdit = pMoreFields
370:     Set pField = New Field
371:     Set newFieldEdit = pField
372:     newFieldEdit.Name = c_DefaultFld_Shape
373:     newFieldEdit.Type = esriFieldTypeGeometry
374:     Set pGeomDef = New GeometryDef
375:     Set pGeomDefEdit = pGeomDef
376:     With pGeomDefEdit
377:         .GeometryType = esriGeometryPolygon
378:         Set .SpatialReference = pMap.SpatialReference
379:     End With
380:     Set newFieldEdit.GeometryDef = pGeomDef
381:     pFieldsEdit.AddField pField
    ' Validate field names
383:     Set pFieldChecker = New FieldChecker
384:     Set pFieldChecker.ValidateWorkspace = pShapeWorkspace
385:     Set pNewFields = pMoreFields
386:     Set pClone = pNewFields
387:     Set pCloneFields = pClone.Clone
388:     pFieldChecker.Validate pCloneFields, pErrorEnum, pOutputFields
    ' Create the output featureclass
390:     shapeFieldName = c_DefaultFld_Shape
391:     featureclassName = Mid(pNewFile, Len(pOutShpWspName.PathName) + 2)
392:     Set pFeatureWorkspace = pShapeWorkspace
393:     Set pNewFeatClass = pFeatureWorkspace.CreateFeatureClass(featureclassName, pOutputFields, _
                            Nothing, Nothing, esriFTSimple, shapeFieldName, "")
    ' Return
396:     Set NewShapeFile = pNewFeatClass
  
    Exit Function
  
ErrorHandler:
401:     MsgBox "Error creating " & pNewFile & vbCrLf & Err.Number & ": " & Err.Description, _
        vbCritical, "Error in NewShapefile"
End Function

Public Function EntryName(sFile As String) As String
  ' work from the right side to the first file delimeter
  Dim iLength As Integer
408:   iLength = Len(sFile)
  Dim iCounter As Integer
  Dim sDelim As String
411:   sDelim = "\"
  Dim sRight As String
  
414:   For iCounter = iLength To 0 Step -1
    
416:     If Mid$(sFile, iCounter, 1) = sDelim Then
417:       EntryName = Mid$(sFile, 1, (iCounter - 1))
418:       Exit For
419:     End If
  
421:   Next
  
End Function

Public Sub TurnOffClipping(pSeriesProps As IDSMapSeriesProps, pApp As IApplication)
On Error GoTo ErrHand:
  Dim pMap As IMap, pDoc As IMxDocument
  'Find the data frame
429:   Set pDoc = pApp.Document
430:   Set pMap = FindDataFrame(pDoc, pSeriesProps.DataFrameName)
  If pMap Is Nothing Then Exit Sub
  
433:   pMap.ClipGeometry = Nothing

  Exit Sub
ErrHand:
437:   MsgBox "TurnOffClipping - " & Err.Description
End Sub

Public Sub RemoveIndicators(pApp As IApplication)
On Error GoTo ErrHand:
  Dim lLoop As Long, pDoc As IMxDocument, pDelColl As Collection
  Dim pPage As IPageLayout, pGraphCont As IGraphicsContainer
  Dim pElem As IElement, pMapFrame As IMapFrame
445:   Set pDoc = pApp.Document
446:   Set pPage = pDoc.PageLayout
447:   Set pDelColl = New Collection
448:   Set pGraphCont = pPage
449:   pGraphCont.Reset
450:   Set pElem = pGraphCont.Next
451:   Do While Not pElem Is Nothing
452:     If TypeOf pElem Is IMapFrame Then
453:       Set pMapFrame = pElem
454:       If pMapFrame.Map.Name = "Local Indicator" Or _
       pMapFrame.Map.Name = "Global Indicator" Then
456:         pDelColl.Add pMapFrame
457:       End If
458:     End If
    
460:     Set pElem = pGraphCont.Next
461:   Loop
  
463:   For lLoop = 1 To pDelColl.count
464:     pGraphCont.DeleteElement pDelColl.Item(lLoop)
465:   Next lLoop

  Exit Sub
ErrHand:
469:   MsgBox "RemoveIndicators - " & Err.Description
End Sub

Public Sub RemoveLabels(pDoc As IMxDocument)
On Error GoTo ErrHand:
  Dim pGraphicsCont As IGraphicsContainer
  Dim pTempColl As Collection, pElemProps As IElementProperties, lLoop As Long
  'Remove any previous neighbor labels.
477:   Set pGraphicsCont = pDoc.PageLayout
478:   pGraphicsCont.Reset
479:   Set pTempColl = New Collection
480:   Set pElemProps = pGraphicsCont.Next
481:   Do While Not pElemProps Is Nothing
482:     If pElemProps.Name = "DSMAPBOOK TEXT" Then
483:       pTempColl.Add pElemProps
484:     End If
485:     Set pElemProps = pGraphicsCont.Next
486:   Loop
487:   For lLoop = 1 To pTempColl.count
488:     pGraphicsCont.DeleteElement pTempColl.Item(lLoop)
489:   Next lLoop
490:   Set pTempColl = Nothing

  Exit Sub
ErrHand:
494:   MsgBox "RemoveLabels - " & Err.Description
End Sub

Public Function GetMapBookExtension(pApp As IApplication) As IDSMapBook
On Error GoTo ErrHand:
  Dim pMapBookExt As DSMapBookExt, pMapBook As IDSMapBook
500:   Set pMapBookExt = pApp.FindExtensionByName("DevSample_MapBook")
501:   If pMapBookExt Is Nothing Then
502:     MsgBox "Map Book code not installed properly!!  Make sure you can access the regsvr32 command" & vbCrLf & _
     "and rerun the _Install.bat batch file!!", , "Map Book Extension Not Found!!!"
504:     Set GetMapBookExtension = Nothing
    Exit Function
506:   End If
  
508:   Set GetMapBookExtension = pMapBookExt.MapBook

  Exit Function
ErrHand:
512:   MsgBox "GetMapBookExtension - " & Err.Description
End Function

Public Sub RemoveClipElement(pDoc As IMxDocument)
On Error GoTo ErrHand:
  Dim pGraphs As IGraphicsContainer, pElemProps As IElementProperties
  
  'Search for an existing clip element and delete it when found
520:   Set pGraphs = pDoc.FocusMap
521:   pGraphs.Reset
522:   Set pElemProps = pGraphs.Next
523:   Do While Not pElemProps Is Nothing
524:     If TypeOf pElemProps Is IPolygonElement Then
525:       If UCase(pElemProps.Name) = "DSMAPBOOK CLIP ELEMENT" Then
526:         pGraphs.DeleteElement pElemProps
527:         Exit Do
528:       End If
529:     End If
530:     Set pElemProps = pGraphs.Next
531:   Loop
532:   pDoc.ActiveView.PartialRefresh esriViewGraphics, Nothing, Nothing

  Exit Sub
ErrHand:
536:   MsgBox "RemoveClipElement - " & Erl & " - " & Err.Description
End Sub

