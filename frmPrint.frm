VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPrint 
   Caption         =   "Print"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   6945
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraMapSize 
      Caption         =   "Map Larger than Printer Paper"
      Height          =   2415
      Left            =   3840
      TabIndex        =   24
      Top             =   2280
      Width           =   3015
      Begin VB.Frame fraTileOptions 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Enabled         =   0   'False
         Height          =   975
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   2775
         Begin VB.TextBox txtTo 
            Height          =   285
            Left            =   2215
            TabIndex        =   34
            Text            =   "1"
            Top             =   465
            Width           =   375
         End
         Begin VB.TextBox txtFrom 
            Height          =   285
            Left            =   1500
            TabIndex        =   32
            Text            =   "1"
            Top             =   465
            Width           =   375
         End
         Begin VB.OptionButton optPages 
            Caption         =   "Pages"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   480
            Width           =   855
         End
         Begin VB.OptionButton optTileAll 
            Caption         =   "All"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "to:"
            Height          =   255
            Left            =   1960
            TabIndex        =   33
            Top             =   490
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "from:"
            Height          =   255
            Left            =   1080
            TabIndex        =   31
            Top             =   490
            Width           =   375
         End
      End
      Begin VB.OptionButton optProceed 
         Caption         =   "Proceed with printing, some clipping may occur"
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   1920
         Width           =   2775
      End
      Begin VB.OptionButton optScale 
         Caption         =   "Scale map to fit printer paper"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1560
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.OptionButton optTile 
         Caption         =   "Tile map to printer paper"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5520
      TabIndex        =   21
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4200
      TabIndex        =   20
      Top             =   5160
      Width           =   975
   End
   Begin VB.Frame fraCopies 
      Caption         =   "Copies"
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   3615
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   325
         Left            =   2056
         TabIndex        =   23
         Top             =   325
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtCopies"
         BuddyDispid     =   196623
         OrigLeft        =   2280
         OrigTop         =   240
         OrigRight       =   2520
         OrigBottom      =   615
         Max             =   100
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtCopies 
         Height          =   325
         Left            =   1680
         TabIndex        =   22
         Text            =   "1"
         Top             =   325
         Width           =   375
      End
      Begin VB.Label lblNumberofCopies 
         Caption         =   "Number of Copies:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame fraPageRange 
      Caption         =   "Page Range"
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   3615
      Begin VB.CheckBox chkDisabled 
         Caption         =   "Don't output disabled pages"
         Height          =   195
         Left            =   240
         TabIndex        =   35
         Top             =   2040
         Width           =   2385
      End
      Begin VB.TextBox txtPrintPages 
         Height          =   325
         Left            =   1200
         TabIndex        =   18
         Top             =   930
         Width           =   1695
      End
      Begin VB.OptionButton optPrintPages 
         Caption         =   "Pages:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton optPrintCurrentPage 
         Caption         =   "Current page"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optPrintAll 
         Caption         =   "All"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.Label lblPrintPagesDesc 
         Caption         =   "Enter page number and/ or page ranges separated by commas.  For example, 1,3,5-12"
         Height          =   435
         Left            =   240
         TabIndex        =   19
         Top             =   1440
         Width           =   3255
      End
   End
   Begin VB.Frame fraPrinter 
      Caption         =   "Printer"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin MSComDlg.CommonDialog dlgPrint 
         Left            =   4440
         Top             =   1200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CheckBox chkPrintToFile 
         Caption         =   "Print to File"
         Height          =   255
         Left            =   5280
         TabIndex        =   37
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdSetup 
         Caption         =   "Setup..."
         Height          =   375
         Left            =   5280
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblComment 
         Height          =   255
         Left            =   960
         TabIndex        =   14
         Top             =   1680
         Width           =   3495
      End
      Begin VB.Label lblLocation 
         Height          =   255
         Left            =   960
         TabIndex        =   13
         Top             =   1350
         Width           =   3495
      End
      Begin VB.Label lblType 
         Height          =   255
         Left            =   960
         TabIndex        =   12
         Top             =   1020
         Width           =   3495
      End
      Begin VB.Label lblStatus 
         Height          =   255
         Left            =   960
         TabIndex        =   11
         Top             =   690
         Width           =   3495
      End
      Begin VB.Label lblName 
         Height          =   255
         Left            =   960
         TabIndex        =   10
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label lblPrinterComment 
         Caption         =   "Comment:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblPrinterLocation 
         Caption         =   "Where:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1350
         Width           =   615
      End
      Begin VB.Label lblPrinterType 
         Caption         =   "Type:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1020
         Width           =   495
      End
      Begin VB.Label lblPrinterStatus 
         Caption         =   "Status:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   690
         Width           =   615
      End
      Begin VB.Label lblPrinterName 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Label lblPrintStatus 
      Height          =   225
      Left            =   0
      TabIndex        =   36
      Top             =   5880
      Width           =   6750
   End
End
Attribute VB_Name = "frmPrint"
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
Private m_pExportFrame As IModelessFrame

Private Sub chkPrintToFile_Click()
36:   If Me.chkPrintToFile.Value = 1 Then
37:     Me.txtCopies.Text = 1
38:     Me.fraCopies.Enabled = False
39:     Me.txtCopies.Enabled = False
40:     Me.UpDown1.Enabled = False
41:     Me.lblNumberofCopies.Enabled = False
42:   Else
43:     fraCopies.Enabled = True
44:     Me.txtCopies.Enabled = True
45:     Me.UpDown1.Enabled = True
46:     Me.lblNumberofCopies.Enabled = True
47:   End If
End Sub

Private Sub cmdCancel_Click()
51:     m_pExportFrame.Visible = False
52:     Unload Me
End Sub

Public Property Let ExportFrame(ByVal pExportFrame As IModelessFrame)
56:     Set m_pExportFrame = pExportFrame
End Property

Public Property Get aDSMapPage() As IDSMapPage
60:     Set aDSMapPage = m_pMapPage
End Property

Public Property Let aDSMapPage(ByVal pMapPage As IDSMapPage)
64:     Set m_pMapPage = pMapPage
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

Private Sub cmdOK_Click()
On Error GoTo ErrorHandler

  Dim pAView As IActiveView
  Dim pPrinter As IPrinter
  Dim pMxApp As IMxApplication
  Dim pMxDoc As IMxDocument
  Dim pLayout As IPageLayout
  Dim iNumPages As Integer
  Dim pPage As IPage
  Dim pMouse As IMouseCursor
  
95:   Set pMouse = New MouseCursor
96:   pMouse.SetCursor 2

98:   Set pMxApp = m_pApp
99:   Set pPrinter = pMxApp.Printer
100:   Set pMxDoc = m_pApp.Document
101:   Set pLayout = pMxDoc.PageLayout
102:   Set pPage = pLayout.Page
  
104:   If Me.chkPrintToFile.Value = 1 Then
'    If UCase(pPrinter.FileExtension) = "PS" Then
106:       Me.dlgPrint.Filter = "Postscript Files (*.ps,*.eps)|*.ps,*.eps"
'    Else
'      Me.dlgPrint.Filter = UCase(pPrinter.FileExtension) & " (*." & LCase(pPrinter.FileExtension) & ")" & "|*." & LCase(pPrinter.FileExtension)
'    End If
    
111:     Me.dlgPrint.DialogTitle = "Print to File"
'    Me.Hide
113:     m_pExportFrame.Visible = False
114:     Me.dlgPrint.ShowSave
    
    Dim sFileName As String, sPrefix As String, sExt As String, sSplit() As String
    
118:     sFileName = Me.dlgPrint.FileName
119:     If sFileName <> "" Then
120:       If InStr(1, sFileName, ".", vbTextCompare) > 0 Then
121:         sSplit = Split(sFileName, ".", , vbTextCompare)
122:         sPrefix = sSplit(0)
123:         sExt = sSplit(1)
124:       Else
125:         sPrefix = sFileName
126:         sExt = "ps"
127:         sFileName = sFileName & ".ps"
128:       End If
129:     Else
130:       MsgBox "Please specify a file name for the page(s)"
'      Me.Show
132:       m_pExportFrame.Visible = True
      Exit Sub
134:     End If
135:   End If
  
137:   If Me.optTile.Value = True Then
138:       pPage.PageToPrinterMapping = esriPageMappingTile
139:   ElseIf Me.optScale = True Then
140:       pPage.PageToPrinterMapping = esriPageMappingScale
141:   ElseIf Me.optProceed.Value = True Then
142:       pPage.PageToPrinterMapping = esriPageMappingCrop
143:   End If
  
145:   pPrinter.Paper.Orientation = pLayout.Page.Orientation
  
  Dim rectDeviceBounds As tagRECT
  Dim pVisBounds As IEnvelope
  Dim hdc As Long
  Dim lDPI As Long
  Dim devFrameEnvelope As IEnvelope
  Dim iCurrentPage As Integer, pSeriesOpts As IDSMapSeriesOptions
  Dim pSeriesOpts2 As IDSMapSeriesOptions2
  
  'Need to include code here to create a collection of all of the map pages that you can
  'then loop through and print.
  Dim PagesToPrint As Collection
  Dim i As Long
  Dim pMapPage As IDSMapPage
  Dim numPages As Long
  Dim a As Long
  
163:   Set PagesToPrint = New Collection
  
165:   If Not m_pMapPage Is Nothing Then
166:       PagesToPrint.Add m_pMapPage
167:   End If
  
169:   If m_pMapPage Is Nothing And m_pMapBook Is Nothing Then
170:     If Me.optPrintAll.Value = True Then
171:       For i = 0 To m_pMapSeries.PageCount - 1
172:         If chkDisabled.Value = 1 Then
173:           If m_pMapSeries.Page(i).EnablePage Then
174:             PagesToPrint.Add m_pMapSeries.Page(i)
175:           End If
176:         Else
177:           PagesToPrint.Add m_pMapSeries.Page(i)
178:         End If
179:       Next i
180:     ElseIf Me.optPrintPages.Value = True Then
      'parse out the pages to print
182:       If chkDisabled.Value = 1 Then
183:         Set PagesToPrint = ParseOutPages(Me.txtPrintPages.Text, m_pMapSeries, True)
184:       Else
185:         Set PagesToPrint = ParseOutPages(Me.txtPrintPages.Text, m_pMapSeries, False)
186:       End If
      If PagesToPrint.count = 0 Then Exit Sub
188:     End If
189:   End If
      
191:   numPages = CLng(Me.txtCopies.Text)
  
194:   If PagesToPrint.count > 0 Then
195:     Set pSeriesOpts = m_pMapSeries
196:     Set pSeriesOpts2 = pSeriesOpts
197:     If pSeriesOpts2.ClipData > 0 Then
198:       g_bClipFlag = True
199:     End If
200:     If pSeriesOpts.RotateFrame Then
201:       g_bRotateFlag = True
202:     End If
203:     If pSeriesOpts.LabelNeighbors Then
204:       g_bLabelNeighbors = True
205:     End If
206:     For i = 1 To PagesToPrint.count
207:       Set pMapPage = PagesToPrint.Item(i)
208:       pMapPage.DrawPage pMxDoc, m_pMapSeries, False
209:       CheckNumberOfPages pPage, pPrinter, iNumPages
210:       lblPrintStatus.Caption = "Printing page " & pMapPage.PageName & " ..."
        
212:       For iCurrentPage = 1 To iNumPages
213:         SetupToPrint pPrinter, pPage, iCurrentPage, lDPI, rectDeviceBounds, pVisBounds, devFrameEnvelope
214:         If Me.chkPrintToFile.Value = 1 Then
215:           If pPage.PageToPrinterMapping = esriPageMappingTile Then
216:             pPrinter.PrintToFile = sPrefix & "_" & pMapPage.PageName & "_" & iCurrentPage & "." & sExt
217:           Else
218:             pPrinter.PrintToFile = sPrefix & "_" & pMapPage.PageName & "." & sExt
219:           End If
220:         End If
221:         For a = 1 To numPages
222:           hdc = pPrinter.StartPrinting(devFrameEnvelope, 0)
223:             pMxDoc.ActiveView.Output hdc, lDPI, rectDeviceBounds, pVisBounds, Nothing
224:             pMapPage.LastOutputted = Format(Date, "mm/dd/yyyy")
225:           pPrinter.FinishPrinting
226:         Next a
227:       Next iCurrentPage
228:     Next i
229:   End If
  
231:   If Not m_pMapBook Is Nothing Then
    Dim pSeriesCount As Long
    Dim MapSeriesColl As Collection
    Dim pMapSeries As IDSMapSeries
    Dim count As Long
    
237:     pSeriesCount = m_pMapBook.ContentCount
    
239:     Set MapSeriesColl = New Collection
    
241:     For i = 0 To pSeriesCount - 1
242:         MapSeriesColl.Add m_pMapBook.ContentItem(i)
243:     Next i

    If MapSeriesColl.count = 0 Then Exit Sub
    
247:     For i = 1 To MapSeriesColl.count
248:       Set PagesToPrint = New Collection
249:       Set pMapSeries = MapSeriesColl.Item(i)
250:       Set pSeriesOpts = pMapSeries
251:       Set pSeriesOpts2 = pSeriesOpts
      
253:       If pSeriesOpts2.ClipData > 0 Then
254:         g_bClipFlag = True
255:       End If
256:       If pSeriesOpts.RotateFrame Then
257:         g_bRotateFlag = True
258:       End If
259:       If pSeriesOpts.LabelNeighbors Then
260:         g_bLabelNeighbors = True
261:       End If
        
263:       For count = 0 To pMapSeries.PageCount - 1
264:         If chkDisabled.Value = 1 Then
265:           If pMapSeries.Page(count).EnablePage Then
266:             PagesToPrint.Add pMapSeries.Page(count)
267:           End If
268:         Else
269:           PagesToPrint.Add pMapSeries.Page(count)
270:         End If
271:       Next count
      
273:       For count = 1 To PagesToPrint.count
      'now do printing
275:         Set pMapPage = PagesToPrint.Item(count)
276:         pMapPage.DrawPage pMxDoc, pMapSeries, False
        
278:         CheckNumberOfPages pPage, pPrinter, iNumPages
279:         lblPrintStatus.Caption = "Printing page " & pMapPage.PageName & " ..."
            
281:         For iCurrentPage = 1 To iNumPages
282:           SetupToPrint pPrinter, pPage, iCurrentPage, lDPI, rectDeviceBounds, pVisBounds, devFrameEnvelope
283:           If Me.chkPrintToFile.Value = 1 Then
284:             If pPage.PageToPrinterMapping = esriPageMappingTile Then
285:               pPrinter.PrintToFile = sPrefix & "_" & pMapPage.PageName & "_" & iCurrentPage & "." & sExt
286:             Else
287:               pPrinter.PrintToFile = sPrefix & "_" & pMapPage.PageName & "." & sExt
288:             End If
289:           End If
290:           For a = 1 To numPages
291:             hdc = pPrinter.StartPrinting(devFrameEnvelope, 0)
292:               pMxDoc.ActiveView.Output hdc, lDPI, rectDeviceBounds, pVisBounds, Nothing
293:               pMapPage.LastOutputted = Format(Date, "mm/dd/yyyy")
294:             pPrinter.FinishPrinting
295:           Next a
296:         Next iCurrentPage
      
298:       Next count
            
300:     Next i
301:   End If
                                   
303:   lblPrintStatus.Caption = ""
304:   Set m_pMapBook = Nothing
305:   Set m_pMapPage = Nothing
306:   Set m_pMapSeries = Nothing
307:   m_pExportFrame.Visible = False
308:   Unload Me

  Exit Sub
ErrorHandler:
312:   lblPrintStatus.Caption = ""
313:   MsgBox "cmdOK_Click - " & Err.Description
End Sub

Public Property Get Application() As IApplication
317:     Set Application = m_pApp
End Property

Public Property Let Application(ByVal pApp As IApplication)
321:     Set m_pApp = pApp
End Property

Private Sub cmdSetup_Click()
325:   If (Not m_pApp.IsDialogVisible(esriMxDlgPageSetup)) Then
    Dim bDialog As Boolean
    Dim pPrinter As IPrinter
    Dim pMxApp As IMxApplication
329:     m_pApp.ShowDialog esriMxDlgPageSetup, True
    
331:     m_pExportFrame.Visible = False
'    Me.Hide
333:     bDialog = True
    
335:     While bDialog = True
336:         bDialog = m_pApp.IsDialogVisible(esriMxDlgPageSetup)
337:         DoEvents
        
'            Sleep 1
    
341:     Wend
    
343:     Set pMxApp = m_pApp
344:     Set pPrinter = pMxApp.Printer
345:     Me.lblName.Caption = pPrinter.Paper.PrinterName
346:     Me.lblType.Caption = pPrinter.DriverName
347:     If TypeOf pPrinter Is IPsPrinter Then
348:       Me.chkPrintToFile.Enabled = True
349:     Else
350:       Me.chkPrintToFile.Value = 0
351:       Me.chkPrintToFile.Enabled = False
352:     End If
'    Me.Show
354:     m_pExportFrame.Visible = True
355:   End If
End Sub

Private Sub Form_Load()
359:   chkDisabled.Value = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
363:     Set m_pApp = Nothing
364:     Set m_pMapPage = Nothing
365:     Set m_pMapSeries = Nothing
366:     Set m_pMapBook = Nothing
367:     Set m_pExportFrame = Nothing
End Sub

Private Sub optProceed_Click()
371:     If optProceed.Value = True Then
372:         Me.fraTileOptions.Enabled = False
373:     End If
End Sub

Private Sub optScale_Click()
377:     If optScale.Value = True Then
378:         Me.fraTileOptions.Enabled = False
379:     End If
End Sub

Private Sub optTile_Click()
383:     If optTile.Value = True Then
384:         Me.fraTileOptions.Enabled = True
385:         Me.optTileAll.Value = True
386:     Else
387:         Me.fraTileOptions.Enabled = False
388:     End If
End Sub

Public Sub SetupToPrint(pPrinter As IPrinter, pPage As IPage, iCurrentPage As Integer, ByRef lDPI As Long, ByRef rectDeviceBounds As tagRECT, _
ByRef pVisBounds As IEnvelope, ByRef devFrameEnvelope As IEnvelope)
On Error GoTo ErrorHandler
  Dim idpi As Integer
  Dim pDeviceBounds As IEnvelope
  Dim paperWidthInch As Double
  Dim paperHeightInch As Double

399:   idpi = pPrinter.Resolution  'dots per inch
          
401:   Set pDeviceBounds = New Envelope
              
403:   pPage.GetDeviceBounds pPrinter, iCurrentPage, 0, idpi, pDeviceBounds
               
405:   rectDeviceBounds.Left = pDeviceBounds.XMin
406:   rectDeviceBounds.Top = pDeviceBounds.YMin
407:   rectDeviceBounds.Right = pDeviceBounds.XMax
408:   rectDeviceBounds.bottom = pDeviceBounds.YMax
  
  'Following block added 6/19/03 to fix problem with plots being cutoff
411:   If TypeOf pPrinter Is IEmfPrinter Then
    ' For emf printers we have to remove the top and left unprintable area
    ' from device coordinates so its origin is 0,0.
    '
415:     rectDeviceBounds.Right = rectDeviceBounds.Right - rectDeviceBounds.Left
416:     rectDeviceBounds.bottom = rectDeviceBounds.bottom - rectDeviceBounds.Top
417:     rectDeviceBounds.Left = 0
418:     rectDeviceBounds.Top = 0
419:   End If
  
421:   Set pVisBounds = New Envelope
422:   pPage.GetPageBounds pPrinter, iCurrentPage, 0, pVisBounds
423:   pPrinter.QueryPaperSize paperWidthInch, paperHeightInch
424:   Set devFrameEnvelope = New Envelope
425:   devFrameEnvelope.PutCoords 0, 0, paperWidthInch * idpi, paperHeightInch * idpi
  
427:   lDPI = CLng(idpi)

  Exit Sub
ErrorHandler:
431:   MsgBox "SetupToPrint - " & Err.Description
End Sub

Public Sub CheckNumberOfPages(pPage As IPage, pPrinter As IPrinter, ByRef iNumPages As Integer)
On Error GoTo ErrorHandler
436:   pPage.PrinterPageCount pPrinter, 0, iNumPages
      
438:   If Me.optTile.Value = True Then
439:     If Me.optPages.Value = True Then
      Dim iPageNo As Integer
      Dim sPageNo As String
442:       sPageNo = Me.txtTo.Text
      
444:       If sPageNo <> "" Then
445:           iPageNo = CInt(sPageNo)
446:       Else
          Exit Sub
448:       End If
      
450:       If iPageNo < iNumPages Then
451:           iNumPages = iPageNo
452:       End If
453:     End If
454:   End If
  
  Exit Sub
ErrorHandler:
458:   MsgBox "CheckNumberOfPages - " & Err.Description
End Sub
