VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmExport 
   Caption         =   "Export"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   5595
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFilename 
      Height          =   315
      Left            =   840
      TabIndex        =   12
      Top             =   30
      Width           =   3375
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   0
      Width           =   855
   End
   Begin VB.ComboBox cboSaveAsType 
      Height          =   315
      Left            =   1200
      TabIndex        =   10
      Top             =   870
      Width           =   3015
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "Options..."
      Height          =   375
      Left            =   4440
      TabIndex        =   9
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3870
      TabIndex        =   8
      Top             =   3300
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4770
      TabIndex        =   7
      Top             =   3300
      Width           =   735
   End
   Begin VB.Frame fraPageRange 
      Caption         =   "Page range"
      Height          =   2295
      Left            =   90
      TabIndex        =   0
      Top             =   1380
      Width           =   3615
      Begin VB.OptionButton optAll 
         Caption         =   "All"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton optCurrentPage 
         Caption         =   "Current page"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optPages 
         Caption         =   "Pages:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtPages 
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   1020
         Width           =   1575
      End
      Begin VB.CheckBox chkDisabled 
         Caption         =   "Don't output disabled pages"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label lblDescription 
         Caption         =   "Enter page number and/ or page ranges separated by commas.  For example, 1,2,5-12"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   3255
      End
   End
   Begin MSComDlg.CommonDialog dlgExport 
      Left            =   4800
      Top             =   1380
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Export"
   End
   Begin VB.Label Label1 
      Caption         =   "The name of the page will be appended to the specified file name."
      Height          =   225
      Left            =   300
      TabIndex        =   16
      Top             =   420
      Width           =   4755
   End
   Begin VB.Label lblStatus 
      Height          =   225
      Left            =   75
      TabIndex        =   15
      Top             =   3735
      Width           =   5445
   End
   Begin VB.Label lblExportTo 
      Caption         =   "Export to:"
      Height          =   255
      Left            =   60
      TabIndex        =   14
      Top             =   60
      Width           =   735
   End
   Begin VB.Label lblExportType 
      Caption         =   "Save as Type:"
      Height          =   255
      Left            =   60
      TabIndex        =   13
      Top             =   900
      Width           =   1095
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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




Private m_pMapPage As IDSMapPage
Private m_pMapSeries As IDSMapSeries
Private m_pMapBook As IDSMapBook
Private m_pApp As IApplication
Private m_pExport As IExport
Private m_pExportFrame As IModelessFrame

Private m_ExportersCol As New Collection
Private m_sFileExtension As String
Private m_sFileNameRoot As String
Private m_sPath As String


Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, _
    lpData As Any, lpcbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, _
    ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, _
    ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal _
    cbData As Long) As Long

Private Const HKEY_CURRENT_USER = &H80000001
Private Const REG_SZ = 1
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const REG_MULTI_SZ = 7
Private Const SYNCHRONIZE = &H100000
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_READ = ((READ_CONTROL Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const ERROR_SUCCESS = 0&


Public Property Get aDSMapPage() As IDSMapPage
56:     Set aDSMapPage = m_pMapPage
End Property

Public Property Let aDSMapPage(ByVal pMapPage As IDSMapPage)
60:     Set m_pMapPage = pMapPage
End Property

Public Property Let ExportFrame(ByVal pExportFrame As IModelessFrame)
64:     Set m_pExportFrame = pExportFrame
End Property

Public Property Get aDSMapSeries() As IDSMapSeries
68:     Set aDSMapSeries = m_pMapSeries
End Property

Public Property Let aDSMapSeries(ByVal pMapSeries As IDSMapSeries)
72:     Set m_pMapSeries = pMapSeries
End Property

Public Property Get aDSMapBook() As IDSMapBook
76:     Set aDSMapBook = m_pMapBook
End Property

Public Property Let aDSMapBook(ByVal pMapBook As IDSMapBook)
80:     Set m_pMapBook = pMapBook
End Property

Public Property Get Application() As IApplication
84:     Set Application = m_pApp
End Property

Public Property Let Application(ByVal pApp As IApplication)
88:     Set m_pApp = pApp
End Property

Public Sub SetupDialog()
  On Error GoTo ErrorHand
  
  Exit Sub
ErrorHand:
96:   MsgBox "SetupDialog - " & Erl & " - " & Err.Description
End Sub

Private Sub cmdBrowse_Click()
On Error GoTo ErrorHand
  Dim sFileName As String
  Dim pTempExport As IExport
  Dim sFileFilter As String
  Dim i As Integer
  
106:   For i = 1 To m_ExportersCol.count
107:     Set pTempExport = m_ExportersCol.Item(i)
108:     Debug.Print pTempExport.Name & ": " & pTempExport.Priority
109:     If pTempExport.Filter <> "" Then
110:       If sFileFilter <> "" Then sFileFilter = sFileFilter & "|"
111:       sFileFilter = sFileFilter & pTempExport.Filter
112:     End If
113:   Next
114:   Set pTempExport = Nothing
  
116:   Me.dlgExport.Filter = sFileFilter
    
118:   If Me.cboSaveAsType.ListIndex <> -1 Then
119:     Me.dlgExport.FilterIndex = Me.cboSaveAsType.ListIndex + 1
120:   Else
121:     Me.dlgExport.FilterIndex = 1
122:   End If
  
124:   Me.dlgExport.DialogTitle = "Export"
125:   Me.dlgExport.FileName = m_sFileNameRoot & m_sFileExtension
  
' Me.Hide
128:   m_pExportFrame.Visible = False

  
131:   Me.dlgExport.ShowSave
  
133:   If Me.dlgExport.FileName = "" Then
134:       Me.Show
      Exit Sub
136:   Else
137:       sFileName = Me.dlgExport.FileName
138:   End If
  
140:   Me.txtFilename.Text = sFileName
  
142:   m_pExportFrame.Visible = True
  
  Exit Sub
ErrorHand:
146:   MsgBox "cmdBrowse_Click - " & Erl & " - " & Err.Description
End Sub

Private Sub cmdCancel_Click()
150:     m_pExportFrame.Visible = False
151:     Unload Me
End Sub

Public Sub InitializeTheForm()
  On Error GoTo ErrorHand
    
  Dim pTempExport As IExport
  Dim i As Integer
  Dim esriExportsCat As New UID
  Dim pCategoryFactory As ICategoryFactory
  Dim TempExportersCol As New Collection
  Dim pSettingsInRegistry As ISettingsInRegistry
  Dim iHighest As Double
  Dim sLastUsedExporterName As String
  Dim lLastUsedExporterPriority As Long
  
  
  
  'Use a Category Factory object to create one instance of every class registered
  ' in the "ESRI Exports" category.
   'Component Category: "ESRI Exports" = {66A7ECF7-9BE1-4E77-A8C7-42D3C62A2590}
172:   esriExportsCat.value = "{66A7ECF7-9BE1-4E77-A8C7-42D3C62A2590}"
173:   Set pCategoryFactory = New CategoryFactory
174:   pCategoryFactory.CategoryID = esriExportsCat
  
  'As each exporter object is created, add it to a vb collection object for later use.
  ' Use each exporter object's Priority property as a unique static key for later
  ' access to each object in the collection.  Because some exporters change their file
  ' extension based on settings (eg. SVG), we should read and sync the registry values
  ' for each exporter after it is created.
181:   Set pTempExport = pCategoryFactory.CreateNext
182:   Do While Not pTempExport Is Nothing
    On Error Resume Next
184:     Set pSettingsInRegistry = pTempExport
    On Error GoTo 0
186:     If Not pSettingsInRegistry Is Nothing Then
187:       pSettingsInRegistry.RestoreForCurrentUser "Software\ESRI\Export\ExportObjectsParams"
188:       m_ExportersCol.Add pTempExport, CStr(pTempExport.Priority)
189:     End If
190:     Set pTempExport = pCategoryFactory.CreateNext
191:   Loop
192:   Set pTempExport = Nothing

  'Run a simple sort operation on the exporters collection, sorting by the exporter
  ' Priority property.  This property is primarily used only for determining the order in
  ' which the exporters are listed in the dialog listbox control.
197:   iHighest = -4294967296#
  Dim j As Integer
199:   Do While m_ExportersCol.count > 0
200:     For i = 1 To m_ExportersCol.count
201:       Set pTempExport = m_ExportersCol(i)
202:       If pTempExport.Priority > iHighest Then
203:         iHighest = pTempExport.Priority
204:       End If
205:     Next
206:     Set pTempExport = m_ExportersCol(CStr(iHighest))
207:     TempExportersCol.Add pTempExport, CStr(pTempExport.Priority)
208:     m_ExportersCol.Remove CStr(iHighest)
209:     iHighest = -4294967296#
210:   Loop
211:   Set m_ExportersCol = TempExportersCol
212:   Set TempExportersCol = Nothing
  
  'Populate the SaveAsType combo box.  VB combo box controls provide the ItemData property, in
  ' which the user to store a data value of type long.  Each value will be associated with each
  ' string entry in the list.  Assign the value of the Priority property to ItemData, so we
  ' can grab it at a later point to tie an exporter object to the selected string entry.
218:   For i = 1 To m_ExportersCol.count
219:     Set pTempExport = m_ExportersCol.Item(i)
220:     Debug.Print pTempExport.Name & ": " & pTempExport.Priority
221:     If pTempExport.Filter <> "" Then
222:       Me.cboSaveAsType.AddItem Split(pTempExport.Filter, "|")(0)
223:       cboSaveAsType.ItemData(cboSaveAsType.NewIndex) = pTempExport.Priority
224:     End If
225:   Next
  
  
  ' get the last used export type from the registry.
229:   If GetRegistryValue(HKEY_CURRENT_USER, "Software\ESRI\Export\ExportDlg", "LastExporter", REG_SZ) <> "" Then _
    sLastUsedExporterName = GetRegistryValue(HKEY_CURRENT_USER, "Software\ESRI\Export\ExportDlg", "LastExporter", REG_SZ)
  
232:   For i = 1 To m_ExportersCol.count
233:     Set pTempExport = m_ExportersCol.Item(i)
234:     If pTempExport.Name = sLastUsedExporterName Then
235:       Debug.Print pTempExport.Name & ": " & pTempExport.Priority
236:       lLastUsedExporterPriority = pTempExport.Priority
237:     End If
238:   Next
  
240:   For i = 0 To Me.cboSaveAsType.ListCount - 1
241:     If Me.cboSaveAsType.ItemData(i) = lLastUsedExporterPriority Then
242:         Me.cboSaveAsType.ListIndex = i
243:     End If
244:   Next
  
246:   If Me.cboSaveAsType.ListIndex = -1 Then
247:     Me.cboSaveAsType.ListIndex = 0
248:   End If
  
250:   Set pTempExport = Nothing

  'assign the last used export path to m_sPath.  Get the value from the registry.
253:   If GetRegistryValue(HKEY_CURRENT_USER, "Software\ESRI\Export\ExportDlg", "WorkingDirectory", REG_SZ) <> "" Then _
    m_sPath = GetRegistryValue(HKEY_CURRENT_USER, "Software\ESRI\Export\ExportDlg", "WorkingDirectory", REG_SZ)
255:   If Right(m_sPath, 1) <> "\" Then _
    m_sPath = m_sPath & "\"
    

259:   m_sFileNameRoot = Left(GetMxdName(), Len(GetMxdName()) - 4)
  
  ' Call the InitExporter procedure to QI the m_pExport onto the currently selected exporter class
262:   InitExporter
  
  Exit Sub
ErrorHand:
266:   MsgBox "InitializeTheForm - " & Erl & " - " & Err.Description
  
End Sub


Private Sub InitExporter()
On Error GoTo ErrorHand
  'Set the interface pointer for the global IExport variable.  The SaveAsType combo box's
  ' ItemData property will return the Priority value that we assigned in the Form_Load event.
  ' Use it as a key to return an exporter object from m_ExportersCol.
276:   Set m_pExport = m_ExportersCol(CStr(cboSaveAsType.ItemData(cboSaveAsType.ListIndex)))
  
  ' Build the file extension string and change the textbox string accordingly.  Resist
  '  temptation to set the exporter object's ExportFileName property here... better to
  '  do that step at the time of the export operation so it will accurately reflect any
  '  changes the user may make to the textbox contents.
282:   m_sFileExtension = Split(Split(cboSaveAsType.Text, "(")(1), ")")(0)
283:   m_sFileExtension = Right(m_sFileExtension, Len(m_sFileExtension) - 1)
  
285:   txtFilename.Text = m_sPath & m_sFileNameRoot & m_sFileExtension
  
  Exit Sub
ErrorHand:
289:   MsgBox "InitExporter - " & Erl & " - " & Err.Description
End Sub


Private Sub cboSaveAsType_Click()
  
295:   InitExporter
  
End Sub


Private Sub cmdOptions_Click()
  On Error GoTo ErrorHand

  'Set the Export property of the ExportPropDlg form, and then show the form modally.  You cannot
  ' show the ExportPropDlg form without first setting this property.
  'As users interact with the form, the properties of the assigned exporter object will change
  ' in real-time. When the form ExportPropDlg is dismissed, the exporter object will reflect any
  ' changes the user may have made.
308:   Set frmExportPropDlg.Export = m_pExport
309:   frmExportPropDlg.Show vbModal, Me
  
311:   Set frmExportPropDlg.Export = Nothing
312:   Unload frmExportPropDlg
  
  'The ExportSVG class has a Compression property that changes the value of the Filter property,
  ' and we must syncronize our file extension to account for the possible change.
316:   If TypeOf m_pExport Is IExportSVG Then
317:     cboSaveAsType.List(cboSaveAsType.ListIndex) = Split(m_pExport.Filter, "|")(0)
318:     m_sPath = GetPathFromPathAndFilename(txtFilename)
319:     m_sFileExtension = Split(Split(cboSaveAsType.Text, "(")(1), ")")(0)
320:     m_sFileExtension = Right(m_sFileExtension, Len(m_sFileExtension) - 1)
321:     txtFilename.Text = m_sPath & m_sFileNameRoot & m_sFileExtension
322:   End If
        
  Exit Sub
ErrorHand:
326:   MsgBox "cmdOptions_Click - " & Erl & " - " & Err.Description
End Sub


Private Sub txtFilename_Change()
331:   m_sFileNameRoot = GetRootNameFromPath(txtFilename)
332:   m_sPath = GetPathFromPathAndFilename(txtFilename)
End Sub

Private Sub txtFileName_GotFocus()
336:   txtFilename.SelStart = 0
337:   txtFilename.SelLength = Len(txtFilename.Text)
End Sub


Private Sub cmdExport_Click()
On Error GoTo ErrorHand:
  Dim sFileExt As String
  Dim pExport As IExport
  Dim pJpegExport As IExportJPEG
  Dim sFileName As String
  Dim pActiveView As IActiveView
  Dim pMxDoc As IMxDocument
  Dim pMouse As IMouseCursor
  Dim pOutputRasterSettings As IOutputRasterSettings
  Dim iPrevOutputImageQuality As Long
  
353:   If Me.txtFilename.Text = "" Then
354:     MsgBox "You have not typed in a valid path!!!"
    Exit Sub
356:   End If
  
  Dim bValid As Boolean
359:   bValid = CheckForValidPath(Me.txtFilename.Text)
    
361:   If bValid = False Then
362:     MsgBox "You have not typed in a valid path!!!"
    Exit Sub
364:   End If

  '***Need to make sure it's a valid path
  
368:   Set pMouse = New MouseCursor
369:   pMouse.SetCursor 2

371:   Set pMxDoc = m_pApp.Document
372:   sFileName = m_sPath & m_sFileNameRoot
373:   sFileExt = m_sFileExtension
    
375:   Set pExport = m_pExport
        
377:   If pExport Is Nothing Then
378:     MsgBox "No export object!!!"
    Exit Sub
380:   End If
   
382:   If GetRegistryValue(HKEY_CURRENT_USER, "Software\ESRI\Export\ExportDlg", "WorkingDirectory", REG_SZ) <> "" Then
383:     SetRegistryValue HKEY_CURRENT_USER, "Software\ESRI\Export\ExportDlg", "WorkingDirectory", REG_SZ, m_sPath
384:   End If
   
386:   If GetRegistryValue(HKEY_CURRENT_USER, "Software\ESRI\Export\ExportDlg", "LastExporter", REG_SZ) <> "" Then
387:     SetRegistryValue HKEY_CURRENT_USER, "Software\ESRI\Export\ExportDlg", "LastExporter", REG_SZ, pExport.Name
388:   End If
  
  'Switch to the Layout view if we are not already there
391:   If Not TypeOf pMxDoc.ActiveView Is IPageLayout Then
392:     Set pMxDoc.ActiveView = pMxDoc.PageLayout
393:   End If

395:   Set pActiveView = pMxDoc.ActiveView
'  pActiveView.ScreenDisplay.DisplayTransformation.ZoomResolution = False
  'Need to include code here to create a collection of all of the map pages that you can
  'then loop through and print.
  Dim PagesToExport As Collection
  Dim i As Long
  Dim pMapPage As IDSMapPage, pSeriesOpts As IDSMapSeriesOptions
  Dim ExportFrame As tagRECT, pSeriesOpts2 As IDSMapSeriesOptions2
  Dim hdc As Long
  Dim dpi As Integer
  Dim sExportFile As String
406:   Set PagesToExport = New Collection
407:   Set pSeriesOpts = m_pMapSeries
408:   Set pSeriesOpts2 = pSeriesOpts
  
410:   If Not m_pMapPage Is Nothing Then
411:       PagesToExport.Add m_pMapPage
412:   End If
  
414:   Set pOutputRasterSettings = pActiveView.ScreenDisplay.DisplayTransformation
415:   iPrevOutputImageQuality = pOutputRasterSettings.ResampleRatio
  
417:   If Not m_pMapSeries Is Nothing And m_pMapPage Is Nothing And m_pMapBook Is Nothing Then
418:     If Me.optAll.value = True Then
419:       For i = 0 To m_pMapSeries.PageCount - 1
420:         If Me.chkDisabled.value = 1 Then
421:           If m_pMapSeries.Page(i).EnablePage Then
422:             PagesToExport.Add m_pMapSeries.Page(i)
423:           End If
424:          Else
425:             PagesToExport.Add m_pMapSeries.Page(i)
426:         End If
427:       Next i
428:     ElseIf Me.optPages.value = True Then
      'parse out the pages to export
430:       If chkDisabled.value = 1 Then
431:         Set PagesToExport = ParseOutPages(Me.txtPages.Text, m_pMapSeries, True)
432:       Else
433:         Set PagesToExport = ParseOutPages(Me.txtPages.Text, m_pMapSeries, False)
434:       End If
      If PagesToExport.count = 0 Then Exit Sub
436:     End If
437:   End If
  
439:   If PagesToExport.count > 0 Then
440:     If pSeriesOpts2.ClipData > 0 Then
441:       g_bClipFlag = True
442:     End If
443:     If pSeriesOpts.RotateFrame Then
444:       g_bRotateFlag = True
445:     End If
446:     If pSeriesOpts.LabelNeighbors Then
447:       g_bLabelNeighbors = True
448:     End If
449:     For i = 1 To PagesToExport.count
450:       Set pMapPage = PagesToExport.Item(i)
451:       pMapPage.DrawPage pMxDoc, m_pMapSeries, False
          
453:       sExportFile = sFileName & "_" & pMapPage.PageName & sFileExt
454:       lblStatus.Caption = "Exporting to " & m_sFileNameRoot & "_" & pMapPage.PageName & sFileExt & " ..."
455:       SetupToExport pExport, dpi, ExportFrame, pActiveView, sExportFile
      
      'Do the export
458:       hdc = pExport.StartExporting
459:         pActiveView.Output hdc, pExport.Resolution, ExportFrame, Nothing, Nothing
460:         pMapPage.LastOutputted = Format(Date, "mm/dd/yyyy")
461:       pExport.FinishExporting
462:       pExport.Cleanup
463:     Next i
464:   End If
            
466:   If Not m_pMapBook Is Nothing Then
    Dim pMapSeries As IDSMapSeries
    Dim count As Long
469:     For i = 0 To m_pMapBook.ContentCount - 1
470:       Set PagesToExport = New Collection
471:       Set pMapSeries = m_pMapBook.ContentItem(i)
472:       Set pSeriesOpts = pMapSeries
    
474:       For count = 0 To pMapSeries.PageCount - 1
475:         If Me.chkDisabled.value = 1 Then
476:           If pMapSeries.Page(count).EnablePage Then
477:             PagesToExport.Add pMapSeries.Page(count)
478:           End If
479:         Else
480:             PagesToExport.Add pMapSeries.Page(count)
481:         End If
482:       Next count
        
484:       If pSeriesOpts2.ClipData > 0 Then
485:         g_bClipFlag = True
486:       End If
487:       If pSeriesOpts.RotateFrame Then
488:         g_bRotateFlag = True
489:       End If
490:       If pSeriesOpts.LabelNeighbors Then
491:         g_bLabelNeighbors = True
492:       End If
493:       For count = 1 To PagesToExport.count
        'now do export
495:         Set pMapPage = PagesToExport.Item(count)
496:         pMapPage.DrawPage pMxDoc, pMapSeries, False
      
498:         sExportFile = sFileName & "_series_" & i & "_" & pMapPage.PageName & sFileExt
 
500:         lblStatus.Caption = "Exporting to " & m_sFileNameRoot & "_series_" & i & "_" & pMapPage.PageName & sFileExt
501:         SetupToExport pExport, pExport.Resolution, ExportFrame, pActiveView, sExportFile
          
        'Do the export
504:         hdc = pExport.StartExporting
505:           pActiveView.Output hdc, pExport.Resolution, ExportFrame, Nothing, Nothing
506:           pMapPage.LastOutputted = Format(Date, "mm/dd/yyyy")
507:         pExport.FinishExporting
508:         pExport.Cleanup
509:       Next count
510:     Next i
511:   End If

'  pActiveView.ScreenDisplay.DisplayTransformation.ZoomResolution = True
514:   If TypeOf pExport Is IOutputCleanup Then
    Dim pCleanup As IOutputCleanup
516:     Set pCleanup = pExport
517:     pCleanup.Cleanup
518:   End If
  
520:   SetOutputQuality pActiveView, iPrevOutputImageQuality

522:   lblStatus.Caption = ""
523:   Set m_pMapBook = Nothing
524:   Set m_pMapPage = Nothing
525:   Set m_pMapSeries = Nothing
526:   m_pExportFrame.Visible = False
527:   Unload Me
  
  Exit Sub
ErrorHand:
531:   lblStatus.Caption = ""
532:   MsgBox "cmdExport_Click - " & Erl & " - " & Err.Description
End Sub



Public Sub SetupToExport(ByRef pExport As IExport, ByRef dpi As Integer, ByRef ExportFrame As tagRECT, pActiveView As IActiveView, sExportFileName As String)
  On Error GoTo ErrorHand
  
  Dim pEnv As IEnvelope, pPageLayout As IPageLayout, pPage As IPage
  Dim dXmax As Double, dYmax As Double
  Dim pOutputRasterSettings As IOutputRasterSettings

544:    Set pEnv = New Envelope
'   pActiveView.ScreenDisplay.DisplayTransformation.Resolution = pExport.Resolution
  'Setup the Export
547:   ExportFrame = pActiveView.ExportFrame

549:   Set pPageLayout = pActiveView
550:   Set pPage = pPageLayout.Page
  
552:   If pPage.Units <> esriInches Then
553:     pPage.Units = esriInches
554:   End If
  
556:   pPage.QuerySize dXmax, dYmax
557:   pEnv.PutCoords 0, 0, dXmax * pExport.Resolution, dYmax * pExport.Resolution

'Commented out code removes a quarter of a unit, most likely an inch, from the extent to make it
'fit better on the page
'  ExportFrame.Top = pExport.Resolution * 0.25
'  ExportFrame.Right = (dXmax - 0.25) * pExport.Resolution
563:   ExportFrame.Right = dXmax * pExport.Resolution
564:   ExportFrame.bottom = dYmax * pExport.Resolution
  
566:   ExportFrame.Left = 0
567:   ExportFrame.Top = 0
            
569:   With pExport
570:     .PixelBounds = pEnv
571:     .ExportFileName = sExportFileName
572:   End With

  
  ' Output Image Quality of the export.  The value here will only be used if the export
  '  object is a format that allows setting of Output Image Quality, i.e. a vector exporter.
  '  The value assigned to ResampleRatio should be in the range 1 to 5.
  '  1 (esriRasterOutputBest) corresponds to "Best", 5 corresponds to "Fast"
579:   If TypeOf pExport Is IOutputRasterSettings Then
    ' for vector formats, get the ResampleRatio from the export object and call SetOutputQuality
    '   to control drawing of raster layers at export time
582:     Set pOutputRasterSettings = pExport
583:     SetOutputQuality pActiveView, pOutputRasterSettings.ResampleRatio
584:     Set pOutputRasterSettings = Nothing
585:   Else
    'always set the output quality of the display to 1 (esriRasterOutputBest) for image export formats
587:     SetOutputQuality pActiveView, esriRasterOutputBest
588:   End If
  
  
  
  Exit Sub
ErrorHand:
594:   MsgBox "SetupToExport - " & Erl & " - " & Err.Description
End Sub


Public Function ConvertToPixels(sOrient As String, pExport As IExport) As Double
On Error GoTo ErrorHand:
  Dim pixelExtent As Long
  Dim pDT As IDisplayTransformation
  Dim deviceRECT As tagRECT
  Dim pMxDoc As IMxDocument
  
605:   Set pMxDoc = m_pApp.Document
606:   Set pDT = pMxDoc.ActiveView.ScreenDisplay.DisplayTransformation
607:   deviceRECT = pDT.DeviceFrame
  
609:   If sOrient = "Height" Then
610:     pixelExtent = Abs(deviceRECT.Top - deviceRECT.bottom)
611:   ElseIf sOrient = "Width" Then
612:     pixelExtent = Abs(deviceRECT.Top - deviceRECT.bottom)
613:   End If
  
615:   ConvertToPixels = (pExport.Resolution * (pixelExtent / pDT.Resolution))
  
  Exit Function
ErrorHand:
619:   MsgBox "ConvertToPixels - " & Erl & " - " & Err.Description
End Function

Private Sub Form_Load()
623:   chkDisabled.value = 1
End Sub

Private Function CheckForValidPath(sPathName As String) As Boolean
  On Error GoTo ErrorHand

629:   CheckForValidPath = False
  
  Dim aPath() As String
632:       aPath = Split(sPathName, ".")

634:   If UBound(aPath) = 0 Then
    Exit Function
636:   ElseIf UBound(aPath) = 1 Then
    
    Dim sPath As String
    Dim lPos As Long
    
641:       lPos = InStrRev(sPathName, "\")
642:       sPath = Left$(sPathName, (Len(sPathName) - (Len(sPathName) - lPos + 1)))
      
644:       If Dir(sPath, vbDirectory) <> "" Then
645:         CheckForValidPath = True
        Exit Function
647:       Else
        Exit Function
649:       End If
      
651:   ElseIf UBound(aPath) > 1 Then
    Exit Function
653:   End If
  
  Exit Function
ErrorHand:
657:   MsgBox "CheckForValidPath - " & Erl & " - " & Err.Description
End Function

Public Sub SetOutputQuality(pActiveView As IActiveView, ByVal lOutputQuality As Long)
On Error GoTo ErrorHand
  Dim pMap As IMap
  Dim pGraphicsContainer As IGraphicsContainer
  Dim pElement As IElement
  Dim pOutputRasterSettings As IOutputRasterSettings
  Dim pMapFrame As IMapFrame
  Dim pTmpActiveView As IActiveView
  
  
670:   If TypeOf pActiveView Is IMap Then
671:     Set pOutputRasterSettings = pActiveView.ScreenDisplay.DisplayTransformation
672:     pOutputRasterSettings.ResampleRatio = lOutputQuality
673:   ElseIf TypeOf pActiveView Is IPageLayout Then
    
    'assign ResampleRatio for PageLayout
676:     Set pOutputRasterSettings = pActiveView.ScreenDisplay.DisplayTransformation
677:     pOutputRasterSettings.ResampleRatio = lOutputQuality
    
    'and assign ResampleRatio to the Maps in the PageLayout
680:     Set pGraphicsContainer = pActiveView
681:     pGraphicsContainer.Reset
682:     Set pElement = pGraphicsContainer.Next
683:     Do While Not pElement Is Nothing
684:       If TypeOf pElement Is IMapFrame Then
685:         Set pMapFrame = pElement
686:         Set pTmpActiveView = pMapFrame.Map
687:         Set pOutputRasterSettings = pTmpActiveView.ScreenDisplay.DisplayTransformation
688:         pOutputRasterSettings.ResampleRatio = lOutputQuality
689:       End If
690:       DoEvents
691:       Set pElement = pGraphicsContainer.Next
692:     Loop
693:     Set pMap = Nothing
694:     Set pMapFrame = Nothing
695:     Set pGraphicsContainer = Nothing
696:     Set pTmpActiveView = Nothing
697:   End If
698:   Set pOutputRasterSettings = Nothing
  
  Exit Sub
ErrorHand:
702:   MsgBox "SetOutputQuality - " & Erl & " - " & Err.Description
End Sub


Private Sub Form_Unload(Cancel As Integer)
707:   Set m_pMapPage = Nothing
708:   Set m_pMapSeries = Nothing
709:   Set m_pMapBook = Nothing
710:   Set m_pApp = Nothing
711:   Set m_pExport = Nothing
712:   Set m_pExportFrame = Nothing
713:   Set m_ExportersCol = Nothing
End Sub

Public Function GetMxdName() As String
On Error GoTo ErrorHand
  Dim pTemplates As ITemplates
  Dim lTempCount As Long
  Dim strDocPath As String
  
722:   Set pTemplates = Application.Templates
723:   lTempCount = pTemplates.count
  
  ' The document is always the last item
726:   strDocPath = pTemplates.Item(lTempCount - 1)
727:   GetMxdName = Split(strDocPath, "\")(UBound(Split(strDocPath, "\")))
  Exit Function
ErrorHand:
730:   MsgBox "GetMxdName - " & Erl & " - " & Err.Description
End Function

Public Function GetRootNameFromPath(sPathAndFilename As String) As String
On Error GoTo ErrorHand

  Dim sRootName As String
737:   sRootName = Split(sPathAndFilename, "\")(UBound(Split(sPathAndFilename, "\")))
738:   sRootName = Split(sRootName, ".")(0)
739:   GetRootNameFromPath = sRootName
  Exit Function
ErrorHand:
742:   MsgBox "GetRootNameFromPath - " & Erl & " - " & Err.Description
End Function

Public Function GetPathFromPathAndFilename(sPathAndFilename As String) As String
On Error GoTo ErrorHand

  Dim sPathName As String
  Dim sRootName As String
750:   sRootName = Split(sPathAndFilename, "\")(UBound(Split(sPathAndFilename, "\")))
751:   sPathName = Left(sPathAndFilename, Len(sPathAndFilename) - Len(sRootName))

753:   GetPathFromPathAndFilename = sPathName
  Exit Function
ErrorHand:
756:   MsgBox "GetPathFromPathAndFilename - " & Erl & " - " & Err.Description
End Function


' Read a Registry value.
' Use KeyName = "" for the default value.
' Supports only DWORD, SZ, and BINARY value types.

Function GetRegistryValue(ByVal hKey As Long, ByVal KeyName As String, _
    ByVal ValueName As String, ByVal KeyType As Integer, _
    Optional DefaultValue As Variant = Empty) As Variant
On Error GoTo ErrorHand

    Dim handle As Long, resLong As Long
    Dim resString As String, length As Long
    Dim resBinary() As Byte
    
    ' Prepare the default result.
774:     GetRegistryValue = DefaultValue
    ' Open the key, exit if not found.
    If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then Exit Function
    
    Select Case KeyType
        Case REG_DWORD
            ' Read the value, use the default if not found.
781:             If RegQueryValueEx(handle, ValueName, 0, REG_DWORD, _
                resLong, 4) = 0 Then
783:                 GetRegistryValue = resLong
784:             End If
        Case REG_SZ
786:             length = 1024: resString = Space$(length)
787:             If RegQueryValueEx(handle, ValueName, 0, REG_SZ, _
                ByVal resString, length) = 0 Then
                ' If value is found, trim characters in excess.
790:                 GetRegistryValue = Left$(resString, length - 1)
791:             End If
        Case REG_BINARY
793:             length = 4096
            ReDim resBinary(length - 1) As Byte
795:             If RegQueryValueEx(handle, ValueName, 0, REG_BINARY, _
                resBinary(0), length) = 0 Then
797:                 GetRegistryValue = resBinary()
798:             End If
        Case Else
800:             Err.Raise 1001, , "Unsupported value type"
801:     End Select
    
803:     RegCloseKey handle
    
    Exit Function
ErrorHand:
807:   MsgBox "GetRegistryvalue - " & Erl & " - " & Err.Description
End Function

' Write / Create a Registry value.
' Use KeyName = "" for the default value.
' Supports only DWORD, SZ, REG_MULTI_SZ, and BINARY value types.

Sub SetRegistryValue(ByVal hKey As Long, ByVal KeyName As String, ByVal ValueName As String, ByVal KeyType As Integer, value As Variant)
On Error GoTo ErrorHand
    Dim handle As Long, lngValue As Long
    Dim strValue As String
    Dim binValue() As Byte, length As Long
    
    ' Open the key, exit if not found.
    If RegOpenKeyEx(hKey, KeyName, 0, KEY_WRITE, handle) Then Exit Sub
    
    Select Case KeyType
        Case REG_DWORD
825:             lngValue = value
826:             RegSetValueEx handle, ValueName, 0, KeyType, lngValue, 4
        Case REG_SZ
828:             strValue = value
829:             RegSetValueEx handle, ValueName, 0, KeyType, ByVal strValue, Len(strValue)
        Case REG_MULTI_SZ
831:             strValue = value
832:             RegSetValueEx handle, ValueName, 0, KeyType, ByVal strValue, Len(strValue)
        Case REG_BINARY
834:             binValue = value
835:             length = UBound(binValue) - LBound(binValue) + 1
836:             RegSetValueEx handle, ValueName, 0, KeyType, binValue(LBound(binValue)), length
837:     End Select
    
    ' Close the key.
840:     RegCloseKey handle
    
    Exit Sub
ErrorHand:
844:   MsgBox "SetRegistryValue - " & Erl & " - " & Err.Description
End Sub


