VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DSMapBookExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

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
'




Option Explicit

Implements IExtension
Implements IPersistVariant

Private m_pApp As IApplication
Private WithEvents m_pMxDoc As MxDocument
Attribute m_pMxDoc.VB_VarHelpID = -1
Private m_pMapBook As IDSMapBook

Private Property Get IExtension_Name() As String
23:   IExtension_Name = "DevSample_MapBook"
End Property

Private Sub IExtension_Shutdown()
27:   Set g_pFrmMapSeries = Nothing
28:   Set m_pMapBook = Nothing
29:   Set m_pMxDoc = Nothing
30:   Set m_pApp = Nothing
End Sub

Public Property Get MapBook() As IDSMapBook
34:   Set MapBook = m_pMapBook
End Property

Private Sub IExtension_Startup(initializationData As Variant)
On Error GoTo ErrHand:
  Dim pApp As IApplication
If g_pFrmMapSeries Is Nothing Then
40:   Set g_pFrmMapSeries = New frmMapSeries
End If
41:   Set g_pFrmMapSeries.m_pApp = initializationData
42:   Set m_pMapBook = New DSMapBook
43:   Set pApp = initializationData
44:   Set m_pApp = pApp
45:   Set m_pMxDoc = pApp.Document
46:   g_bClipFlag = False
47:   g_bRotateFlag = False
48:   g_bLabelNeighbors = False
  
  Exit Sub
ErrHand:
52:   MsgBox "DSMapBookExt_Startup - " & Erl & " - " & Err.Description
End Sub

Private Property Get IPersistVariant_ID() As IUID
On Error GoTo ErrHand:
  
  Dim pUID As New UID
59:   pUID = "DSMapBookUIPrj.DSMapBookExt"

61:   Set IPersistVariant_ID = pUID

63:   GoTo EndProc

ErrHand:
66:   MsgBox "DSMapBookExt_ID - " & Err.Description
  Exit Property
EndProc:
69:   Set pUID = Nothing
End Property

Private Sub IPersistVariant_Load(ByVal Stream As IVariantStream)
On Error GoTo ErrHand:
  Dim pNode As Node, pMapSeries As IDSMapSeries, lLoop As Long, sName As String
  Dim pPage As IDSMapPage
76:   Set m_pMapBook = Stream.Read
  
  'Make sure we have the tab form
  If g_pFrmMapSeries Is Nothing Then Exit Sub
  
  'Make sure our persisted map book has content
  If m_pMapBook.ContentCount = 0 Then Exit Sub
  
  'Put the content back on the form
85:   Set pMapSeries = m_pMapBook.ContentItem(0)
86:   With g_pFrmMapSeries.tvwMapBook
87:     Set pNode = .Nodes.Add("MapBook", tvwChild, "MapSeries", "Map Series", 3)
    
    'Now loop back through the list and add the tile names as nodes in the tree
90:     For lLoop = 0 To pMapSeries.PageCount - 1
91:       Set pPage = pMapSeries.Page(lLoop)
92:       sName = pPage.PageName
93:       If pPage.EnablePage Then
94:         Set pNode = .Nodes.Add("MapSeries", tvwChild, "a" & sName, pPage.PageNumber & " - " & sName, 5)
95:       Else
96:         Set pNode = .Nodes.Add("MapSeries", tvwChild, "a" & sName, pPage.PageNumber & " - " & sName, 6)
97:       End If
98:       pNode.Tag = lLoop
99:     Next lLoop
100:     .Nodes.Item("MapBook").Expanded = True
101:     .Nodes.Item("MapSeries").Expanded = True
102:   End With
  
  Exit Sub
ErrHand:
106:   MsgBox "DSMapBookExt_IPersistVariant_Load - " & Erl & " - " & Err.Description
End Sub

Private Sub IPersistVariant_Save(ByVal Stream As IVariantStream)
On Error GoTo ErrHand:
111:   Stream.Write m_pMapBook
  
  Exit Sub
ErrHand:
115:   MsgBox "DSMapBookExt_IPersistVariant_Save - " & Erl & " - " & Err.Description
End Sub

Private Function m_pMxDoc_ActiveViewChanged() As Boolean
On Error GoTo ErrHand:
'Check to see if the active view is being changed back to the data view after the Map Book code has
'set the clip shape.  The g_bClipFlag variable will tell us if the Map Book code has updated the
'Clip Shape.  If it has, then we want to clear the clip shape when the user switches back to a data view.
'If the clip shape was changed by some other method, then we don't want to update it.
  Dim pDoc As IMxDocument
  
126:   If g_bClipFlag Then
'    If pSeriesOpts2.ClipData = 1 Then
128:       Set pDoc = m_pMxDoc
129:       If TypeOf pDoc.ActiveView Is IMap Then
130:         pDoc.FocusMap.ClipGeometry = Nothing
131:         g_bClipFlag = False
132:       End If
'    Else
134:       RemoveClipElement m_pMxDoc
'      g_bClipFlag = False
'    End If
137:   End If
  
139:   If g_bRotateFlag Then
140:     Set pDoc = m_pMxDoc
141:     If TypeOf pDoc.ActiveView Is IMap Then
142:       pDoc.ActiveView.ScreenDisplay.DisplayTransformation.Rotation = 0
143:       g_bRotateFlag = False
144:     End If
145:   End If
  
147:   If g_bLabelNeighbors Then
148:     Set pDoc = m_pMxDoc
149:     If TypeOf pDoc.ActiveView Is IMap Then
150:       RemoveLabels pDoc
151:       g_bLabelNeighbors = False
152:     End If
153:   End If
  
  Exit Function
ErrHand:
157:   MsgBox "DSMapBookExt_ActiveViewChanged - " & Erl & " - " & Err.Description
End Function

Private Function m_pMxDoc_CloseDocument() As Boolean
161:   DeleteSeries
End Function

Private Function m_pMxDoc_NewDocument() As Boolean
165:   DeleteSeries
End Function

Public Sub DeleteSeries()
On Error GoTo ErrHand:
  Dim pDoc As IMxDocument, pActive As IActiveView, pMapSeries As IDSMapSeries
  
172:   Set pDoc = m_pMxDoc
  'Exit the routine if there is no active view.  This can happen when maps are being produced in batch.
  If pDoc Is Nothing Then Exit Sub
  If pDoc.ActiveView Is Nothing Then Exit Sub
  
177:   Set pActive = pDoc.FocusMap
178:   Set pMapSeries = m_pMapBook.ContentItem(0)
  If pMapSeries Is Nothing Then Exit Sub
  
181:   TurnOffClipping pMapSeries, m_pApp
182:   Set pMapSeries = Nothing
183:   m_pMapBook.RemoveContent 0
184:   g_pFrmMapSeries.tvwMapBook.Nodes.Clear
185:   g_pFrmMapSeries.tvwMapBook.Nodes.Add , , "MapBook", "Map Book", 1
186:   RemoveIndicators m_pApp
187:   pActive.Refresh

  Exit Sub
ErrHand:
191:   MsgBox "DSMapBookExt_DeleteSeries - " & Erl & " - " & Err.Description
End Sub

Public Sub ExportSeries(sInFileName As String, sExportType As String)
On Error GoTo ErrorHand:
  Dim pExport As IExport
  Dim pJpegExport As IExportJPEG
  Dim pActiveView As IActiveView
  Dim pMxDoc As IMxDocument
  Dim sFileExt As String
  Dim sFileName As String
  Dim aPath() As String
  
204:   If sInFileName = "" Then
'    MsgBox "You have not typed in a valid path!!!"
    Exit Sub
207:   End If
  
  Dim bValid As Boolean
210:   bValid = CheckForValidPath(sInFileName)
    
212:   If bValid = False Then
'    MsgBox "You have not typed in a valid path!!!"
    Exit Sub
215:   End If
  
  'Split up the file name
218:   aPath = Split(sInFileName, ".")
219:   sFileName = aPath(0)
220:   sFileExt = "." & aPath(1)

  '***Need to make sure it's a valid path
223:   Set pMxDoc = m_pApp.Document
    
225:   Set pExport = GetTheLastExport(sExportType)
226:   If pExport Is Nothing Then
'    MsgBox "No export object!!!"
    Exit Sub
229:   End If
  
  'Switch to the Layout view if we are not already there
232:   If Not TypeOf pMxDoc.ActiveView Is IPageLayout Then
233:     Set pMxDoc.ActiveView = pMxDoc.PageLayout
234:   End If

236:   Set pActiveView = pMxDoc.ActiveView
'  pActiveView.ScreenDisplay.DisplayTransformation.ZoomResolution = False
  'Need to include code here to create a collection of all of the map pages that you can
  'then loop through and print.
  Dim PagesToExport As Collection
  Dim i As Long
  Dim pMapSeries As IDSMapSeries
  Dim pMapPage As IDSMapPage, pSeriesOpts As IDSMapSeriesOptions
  Dim ExportFrame As tagRECT, pSeriesOpts2 As IDSMapSeriesOptions2
  Dim hdc As Long
  Dim dpi As Integer
  Dim sExportFile As String
248:   Set PagesToExport = New Collection
249:   Set pMapSeries = m_pMapBook.ContentItem(0)
250:   Set pSeriesOpts = pMapSeries
251:   Set pSeriesOpts2 = pSeriesOpts
    
253:   For i = 0 To pMapSeries.PageCount - 1
254:     If pMapSeries.Page(i).EnablePage Then
255:       PagesToExport.Add pMapSeries.Page(i)
256:     End If
257:   Next i
  
259:   If PagesToExport.Count > 0 Then
260:     If pSeriesOpts2.ClipData > 0 Then
261:       g_bClipFlag = True
262:     End If
263:     If pSeriesOpts.RotateFrame Then
264:       g_bRotateFlag = True
265:     End If
266:     If pSeriesOpts.LabelNeighbors Then
267:       g_bLabelNeighbors = True
268:     End If
269:     For i = 1 To PagesToExport.Count
270:       Set pMapPage = PagesToExport.Item(i)
271:       pMapPage.DrawPage pMxDoc, pMapSeries, False
          
273:       sExportFile = sFileName & "_" & pMapPage.PageName & sFileExt
'      lblStatus.Caption = "Exporting to " & m_sFileNameRoot & "_" & pMapPage.PageName & sFileExt & " ..."
275:       SetupToExport pExport, dpi, ExportFrame, pActiveView, sExportFile
      
      'Do the export
278:       hdc = pExport.StartExporting
279:         pActiveView.Output hdc, pExport.Resolution, ExportFrame, Nothing, Nothing
280:         pMapPage.LastOutputted = Format(Date, "mm/dd/yyyy")
281:       pExport.FinishExporting
282:       pExport.Cleanup
283:     Next i
284:   End If
  
  'Cleanup
287:   Set pMapPage = Nothing
288:   Set pMapSeries = Nothing
  
  Exit Sub
ErrorHand:
292:   MsgBox "DSMapBookExt_ExportSeries - " & Erl & " - " & Err.Description
End Sub

Private Function CheckForValidPath(sPathName As String) As Boolean
  On Error GoTo ErrorHand

298:   CheckForValidPath = False
  
  Dim aPath() As String
301:       aPath = Split(sPathName, ".")

303:   If UBound(aPath) = 0 Then
    Exit Function
305:   ElseIf UBound(aPath) = 1 Then
    
    Dim sPath As String
    Dim lPos As Long
    
310:       lPos = InStrRev(sPathName, "\")
311:       sPath = Left$(sPathName, (Len(sPathName) - (Len(sPathName) - lPos + 1)))
      
313:       If Dir(sPath, vbDirectory) <> "" Then
314:         CheckForValidPath = True
        Exit Function
316:       Else
        Exit Function
318:       End If
      
320:   ElseIf UBound(aPath) > 1 Then
    Exit Function
322:   End If
  
  Exit Function
ErrorHand:
326:   MsgBox "DSMapBookExt_CheckForValidPath - " & Erl & " - " & Err.Description
End Function

Private Function GetTheLastExport(sExportType As String) As IExport
  On Error GoTo ErrorHand
    
  Dim pTempExport As IExport
  Dim i As Integer
  Dim esriExportsCat As New UID
  Dim pCategoryFactory As ICategoryFactory
  Dim TempExportersCol As New Collection
  Dim pSettingsInRegistry As ISettingsInRegistry
  Dim sLastUsedExporterName As String
  Dim lLastUsedExporterPriority As Long
  
  'Use a Category Factory object to create one instance of every class registered
  ' in the "ESRI Exports" category.
   'Component Category: "ESRI Exports" = {66A7ECF7-9BE1-4E77-A8C7-42D3C62A2590}
344:   esriExportsCat.Value = "{66A7ECF7-9BE1-4E77-A8C7-42D3C62A2590}"
345:   Set pCategoryFactory = New CategoryFactory
346:   pCategoryFactory.CategoryID = esriExportsCat
  
  'As each exporter object is created, add it to a vb collection object for later use.
  ' Use each exporter object's Priority property as a unique static key for later
  ' access to each object in the collection.  Because some exporters change their file
  ' extension based on settings (eg. SVG), we should read and sync the registry values
  ' for each exporter after it is created.
353:   Set pTempExport = pCategoryFactory.CreateNext
354:   Do While Not pTempExport Is Nothing
355:     If pTempExport.Name = sExportType Then
356:       Exit Do
357:     End If
    
359:     Set pTempExport = pCategoryFactory.CreateNext
360:   Loop
  
  'Set the exporter properties to whatever they were for the last user
363:   Set pSettingsInRegistry = pTempExport
364:   If Not pSettingsInRegistry Is Nothing Then
365:     pSettingsInRegistry.RestoreForCurrentUser "Software\ESRI\Export\ExportObjectsParams"
366:   End If
  
368:   Set GetTheLastExport = pTempExport
  
  Exit Function
ErrorHand:
372:   MsgBox "DSMapBookExt_GetTheLastExport - " & Erl & " - " & Err.Description
End Function

Private Sub SetupToExport(ByRef pExport As IExport, ByRef dpi As Integer, ByRef ExportFrame As tagRECT, pActiveView As IActiveView, sExportFileName As String)
  On Error GoTo ErrorHand
  
  Dim pEnv As IEnvelope, pPageLayout As IPageLayout, pPage As IPage
  Dim dXmax As Double, dYmax As Double
  Dim pOutputRasterSettings As IOutputRasterSettings

382:    Set pEnv = New Envelope
'   pActiveView.ScreenDisplay.DisplayTransformation.Resolution = pExport.Resolution
  'Setup the Export
385:   ExportFrame = pActiveView.ExportFrame

387:   Set pPageLayout = pActiveView
388:   Set pPage = pPageLayout.Page
  
390:   If pPage.Units <> esriInches Then
391:     pPage.Units = esriInches
392:   End If
  
394:   pPage.QuerySize dXmax, dYmax
395:   pEnv.PutCoords 0, 0, dXmax * pExport.Resolution, dYmax * pExport.Resolution

'Commented out code removes a quarter of a unit, most likely an inch, from the extent to make it
'fit better on the page
'  ExportFrame.Top = pExport.Resolution * 0.25
'  ExportFrame.Right = (dXmax - 0.25) * pExport.Resolution
401:   ExportFrame.Right = dXmax * pExport.Resolution
402:   ExportFrame.bottom = dYmax * pExport.Resolution
  
404:   ExportFrame.Left = 0
405:   ExportFrame.Top = 0
            
407:   With pExport
408:     .PixelBounds = pEnv
409:     .ExportFileName = sExportFileName
410:   End With

  
  ' Output Image Quality of the export.  The value here will only be used if the export
  '  object is a format that allows setting of Output Image Quality, i.e. a vector exporter.
  '  The value assigned to ResampleRatio should be in the range 1 to 5.
  '  1 (esriRasterOutputBest) corresponds to "Best", 5 corresponds to "Fast"
417:   If TypeOf pExport Is IOutputRasterSettings Then
    ' for vector formats, get the ResampleRatio from the export object and call SetOutputQuality
    '   to control drawing of raster layers at export time
420:     Set pOutputRasterSettings = pExport
421:     SetOutputQuality pActiveView, pOutputRasterSettings.ResampleRatio
422:     Set pOutputRasterSettings = Nothing
423:   Else
    'always set the output quality of the display to 1 (esriRasterOutputBest) for image export formats
425:     SetOutputQuality pActiveView, esriRasterOutputBest
426:   End If
  
  Exit Sub
ErrorHand:
430:   MsgBox "DSMapBookExt_SetupToExport - " & Erl & " - " & Err.Description
End Sub

Private Sub SetOutputQuality(pActiveView As IActiveView, ByVal lOutputQuality As Long)
On Error GoTo ErrorHand
  Dim pMap As IMap
  Dim pGraphicsContainer As IGraphicsContainer
  Dim pElement As IElement
  Dim pOutputRasterSettings As IOutputRasterSettings
  Dim pMapFrame As IMapFrame
  Dim pTmpActiveView As IActiveView
  
  
443:   If TypeOf pActiveView Is IMap Then
444:     Set pOutputRasterSettings = pActiveView.ScreenDisplay.DisplayTransformation
445:     pOutputRasterSettings.ResampleRatio = lOutputQuality
446:   ElseIf TypeOf pActiveView Is IPageLayout Then
    
    'assign ResampleRatio for PageLayout
449:     Set pOutputRasterSettings = pActiveView.ScreenDisplay.DisplayTransformation
450:     pOutputRasterSettings.ResampleRatio = lOutputQuality
    
    'and assign ResampleRatio to the Maps in the PageLayout
453:     Set pGraphicsContainer = pActiveView
454:     pGraphicsContainer.Reset
455:     Set pElement = pGraphicsContainer.Next
456:     Do While Not pElement Is Nothing
457:       If TypeOf pElement Is IMapFrame Then
458:         Set pMapFrame = pElement
459:         Set pTmpActiveView = pMapFrame.Map
460:         Set pOutputRasterSettings = pTmpActiveView.ScreenDisplay.DisplayTransformation
461:         pOutputRasterSettings.ResampleRatio = lOutputQuality
462:       End If
463:       DoEvents
464:       Set pElement = pGraphicsContainer.Next
465:     Loop
466:     Set pMap = Nothing
467:     Set pMapFrame = Nothing
468:     Set pGraphicsContainer = Nothing
469:     Set pTmpActiveView = Nothing
470:   End If
471:   Set pOutputRasterSettings = Nothing
  
  Exit Sub
ErrorHand:
475:   MsgBox "DSMapBookExt_SetOutputQuality - " & Erl & " - " & Err.Description
End Sub

