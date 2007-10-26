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

Private m_pMapPage As IDSMapPage
Private m_pMapSeries As IDSMapSeries
Private m_pMapBook As IDSMapBook
Private m_pApp As IApplication
Private m_pExportFrame As IModelessFrame

Private Sub chkPrintToFile_Click()
20:   If Me.chkPrintToFile.value = 1 Then
21:     Me.txtCopies.Text = 1
22:     Me.fraCopies.Enabled = False
23:     Me.txtCopies.Enabled = False
24:     Me.UpDown1.Enabled = False
25:     Me.lblNumberofCopies.Enabled = False
26:   Else
27:     fraCopies.Enabled = True
28:     Me.txtCopies.Enabled = True
29:     Me.UpDown1.Enabled = True
30:     Me.lblNumberofCopies.Enabled = True
31:   End If
End Sub

Private Sub cmdCancel_Click()
35:     m_pExportFrame.Visible = False
36:     Unload Me
End Sub

Public Property Let ExportFrame(ByVal pExportFrame As IModelessFrame)
40:     Set m_pExportFrame = pExportFrame
End Property

Public Property Get aDSMapPage() As IDSMapPage
44:     Set aDSMapPage = m_pMapPage
End Property

Public Property Let aDSMapPage(ByVal pMapPage As IDSMapPage)
48:     Set m_pMapPage = pMapPage
End Property

Public Property Get aDSMapSeries() As IDSMapSeries
52:     Set aDSMapSeries = m_pMapSeries
End Property

Public Property Let aDSMapSeries(ByVal pMapSeries As IDSMapSeries)
56:     Set m_pMapSeries = pMapSeries
End Property

Public Property Get aDSMapBook() As IDSMapBook
60:     Set aDSMapBook = m_pMapBook
End Property

Public Property Let aDSMapBook(ByVal pMapBook As IDSMapBook)
64:     Set m_pMapBook = pMapBook
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
  
79:   Set pMouse = New MouseCursor
80:   pMouse.SetCursor 2

82:   Set pMxApp = m_pApp
83:   Set pPrinter = pMxApp.Printer
84:   Set pMxDoc = m_pApp.Document
85:   Set pLayout = pMxDoc.PageLayout
86:   Set pPage = pLayout.Page
  
88:   If Me.chkPrintToFile.value = 1 Then
'    If UCase(pPrinter.FileExtension) = "PS" Then
90:       Me.dlgPrint.Filter = "Postscript Files (*.ps,*.eps)|*.ps,*.eps"
'    Else
'      Me.dlgPrint.Filter = UCase(pPrinter.FileExtension) & " (*." & LCase(pPrinter.FileExtension) & ")" & "|*." & LCase(pPrinter.FileExtension)
'    End If
    
95:     Me.dlgPrint.DialogTitle = "Print to File"
'    Me.Hide
97:     m_pExportFrame.Visible = False
98:     Me.dlgPrint.ShowSave
    
    Dim sFileName As String, sPrefix As String, sExt As String, sSplit() As String
    
102:     sFileName = Me.dlgPrint.FileName
103:     If sFileName <> "" Then
104:       If InStr(1, sFileName, ".", vbTextCompare) > 0 Then
105:         sSplit = Split(sFileName, ".", , vbTextCompare)
106:         sPrefix = sSplit(0)
107:         sExt = sSplit(1)
108:       Else
109:         sPrefix = sFileName
110:         sExt = "ps"
111:         sFileName = sFileName & ".ps"
112:       End If
113:     Else
114:       MsgBox "Please specify a file name for the page(s)"
'      Me.Show
116:       m_pExportFrame.Visible = True
      Exit Sub
118:     End If
119:   End If
  
121:   If Me.optTile.value = True Then
122:       pPage.PageToPrinterMapping = esriPageMappingTile
123:   ElseIf Me.optScale = True Then
124:       pPage.PageToPrinterMapping = esriPageMappingScale
125:   ElseIf Me.optProceed.value = True Then
126:       pPage.PageToPrinterMapping = esriPageMappingCrop
127:   End If
  
129:   pPrinter.Paper.Orientation = pLayout.Page.Orientation
  
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
  
147:   Set PagesToPrint = New Collection
  
149:   If Not m_pMapPage Is Nothing Then
150:       PagesToPrint.Add m_pMapPage
151:   End If
  
153:   If m_pMapPage Is Nothing And m_pMapBook Is Nothing Then
154:     If Me.optPrintAll.value = True Then
155:       For i = 0 To m_pMapSeries.PageCount - 1
156:         If chkDisabled.value = 1 Then
157:           If m_pMapSeries.Page(i).EnablePage Then
158:             PagesToPrint.Add m_pMapSeries.Page(i)
159:           End If
160:         Else
161:           PagesToPrint.Add m_pMapSeries.Page(i)
162:         End If
163:       Next i
164:     ElseIf Me.optPrintPages.value = True Then
      'parse out the pages to print
166:       If chkDisabled.value = 1 Then
167:         Set PagesToPrint = ParseOutPages(Me.txtPrintPages.Text, m_pMapSeries, True)
168:       Else
169:         Set PagesToPrint = ParseOutPages(Me.txtPrintPages.Text, m_pMapSeries, False)
170:       End If
      If PagesToPrint.count = 0 Then Exit Sub
172:     End If
173:   End If
      
175:   numPages = CLng(Me.txtCopies.Text)
  
177:   If PagesToPrint.count > 0 Then
178:     Set pSeriesOpts = m_pMapSeries
179:     Set pSeriesOpts2 = pSeriesOpts
180:     If pSeriesOpts2.ClipData > 0 Then
181:       g_bClipFlag = True
182:     End If
183:     If pSeriesOpts.RotateFrame Then
184:       g_bRotateFlag = True
185:     End If
186:     If pSeriesOpts.LabelNeighbors Then
187:       g_bLabelNeighbors = True
188:     End If
189:     For i = 1 To PagesToPrint.count
190:       Set pMapPage = PagesToPrint.Item(i)
191:       pMapPage.DrawPage pMxDoc, m_pMapSeries, False
192:       CheckNumberOfPages pPage, pPrinter, iNumPages
193:       lblPrintStatus.Caption = "Printing page " & pMapPage.PageName & " ..."
        
195:       For iCurrentPage = 1 To iNumPages
196:         SetupToPrint pPrinter, pPage, iCurrentPage, lDPI, rectDeviceBounds, pVisBounds, devFrameEnvelope
197:         If Me.chkPrintToFile.value = 1 Then
198:           If pPage.PageToPrinterMapping = esriPageMappingTile Then
199:             pPrinter.PrintToFile = sPrefix & "_" & pMapPage.PageName & "_" & iCurrentPage & "." & sExt
200:           Else
201:             pPrinter.PrintToFile = sPrefix & "_" & pMapPage.PageName & "." & sExt
202:           End If
203:         End If
204:         For a = 1 To numPages
205:           hdc = pPrinter.StartPrinting(devFrameEnvelope, 0)
206:             pMxDoc.ActiveView.Output hdc, lDPI, rectDeviceBounds, pVisBounds, Nothing
207:             pMapPage.LastOutputted = Format(Date, "mm/dd/yyyy")
208:           pPrinter.FinishPrinting
209:         Next a
210:       Next iCurrentPage
211:     Next i
212:   End If
  
214:   If Not m_pMapBook Is Nothing Then
    Dim pSeriesCount As Long
    Dim MapSeriesColl As Collection
    Dim pMapSeries As IDSMapSeries
    Dim count As Long
    
220:     pSeriesCount = m_pMapBook.ContentCount
    
222:     Set MapSeriesColl = New Collection
    
224:     For i = 0 To pSeriesCount - 1
225:         MapSeriesColl.Add m_pMapBook.ContentItem(i)
226:     Next i

    If MapSeriesColl.count = 0 Then Exit Sub
    
230:     For i = 1 To MapSeriesColl.count
231:       Set PagesToPrint = New Collection
232:       Set pMapSeries = MapSeriesColl.Item(i)
233:       Set pSeriesOpts = pMapSeries
234:       Set pSeriesOpts2 = pSeriesOpts
      
236:       If pSeriesOpts2.ClipData > 0 Then
237:         g_bClipFlag = True
238:       End If
239:       If pSeriesOpts.RotateFrame Then
240:         g_bRotateFlag = True
241:       End If
242:       If pSeriesOpts.LabelNeighbors Then
243:         g_bLabelNeighbors = True
244:       End If
        
246:       For count = 0 To pMapSeries.PageCount - 1
247:         If chkDisabled.value = 1 Then
248:           If pMapSeries.Page(count).EnablePage Then
249:             PagesToPrint.Add pMapSeries.Page(count)
250:           End If
251:         Else
252:           PagesToPrint.Add pMapSeries.Page(count)
253:         End If
254:       Next count
      
256:       For count = 1 To PagesToPrint.count
      'now do printing
258:         Set pMapPage = PagesToPrint.Item(count)
259:         pMapPage.DrawPage pMxDoc, pMapSeries, False
        
261:         CheckNumberOfPages pPage, pPrinter, iNumPages
262:         lblPrintStatus.Caption = "Printing page " & pMapPage.PageName & " ..."
            
264:         For iCurrentPage = 1 To iNumPages
265:           SetupToPrint pPrinter, pPage, iCurrentPage, lDPI, rectDeviceBounds, pVisBounds, devFrameEnvelope
266:           If Me.chkPrintToFile.value = 1 Then
267:             If pPage.PageToPrinterMapping = esriPageMappingTile Then
268:               pPrinter.PrintToFile = sPrefix & "_" & pMapPage.PageName & "_" & iCurrentPage & "." & sExt
269:             Else
270:               pPrinter.PrintToFile = sPrefix & "_" & pMapPage.PageName & "." & sExt
271:             End If
272:           End If
273:           For a = 1 To numPages
274:             hdc = pPrinter.StartPrinting(devFrameEnvelope, 0)
275:               pMxDoc.ActiveView.Output hdc, lDPI, rectDeviceBounds, pVisBounds, Nothing
276:               pMapPage.LastOutputted = Format(Date, "mm/dd/yyyy")
277:             pPrinter.FinishPrinting
278:           Next a
279:         Next iCurrentPage
      
281:       Next count
            
283:     Next i
284:   End If
                                   
286:   lblPrintStatus.Caption = ""
287:   Set m_pMapBook = Nothing
288:   Set m_pMapPage = Nothing
289:   Set m_pMapSeries = Nothing
290:   m_pExportFrame.Visible = False
291:   Unload Me

  Exit Sub
ErrorHandler:
295:   lblPrintStatus.Caption = ""
296:   MsgBox "cmdOK_Click - " & Err.Description
End Sub

Public Property Get Application() As IApplication
300:     Set Application = m_pApp
End Property

Public Property Let Application(ByVal pApp As IApplication)
304:     Set m_pApp = pApp
End Property

Private Sub cmdSetup_Click()
308:   If (Not m_pApp.IsDialogVisible(esriMxDlgPageSetup)) Then
    Dim bDialog As Boolean
    Dim pPrinter As IPrinter
    Dim pMxApp As IMxApplication
312:     m_pApp.ShowDialog esriMxDlgPageSetup, True
    
314:     m_pExportFrame.Visible = False
'    Me.Hide
316:     bDialog = True
    
318:     While bDialog = True
319:         bDialog = m_pApp.IsDialogVisible(esriMxDlgPageSetup)
320:         DoEvents
        
'            Sleep 1
    
324:     Wend
    
326:     Set pMxApp = m_pApp
327:     Set pPrinter = pMxApp.Printer
328:     Me.lblName.Caption = pPrinter.Paper.PrinterName
329:     Me.lblType.Caption = pPrinter.DriverName
330:     If TypeOf pPrinter Is IPsPrinter Then
331:       Me.chkPrintToFile.Enabled = True
332:     Else
333:       Me.chkPrintToFile.value = 0
334:       Me.chkPrintToFile.Enabled = False
335:     End If
'    Me.Show
337:     m_pExportFrame.Visible = True
338:   End If
End Sub

Private Sub Form_Load()
342:   chkDisabled.value = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
346:     Set m_pApp = Nothing
347:     Set m_pMapPage = Nothing
348:     Set m_pMapSeries = Nothing
349:     Set m_pMapBook = Nothing
350:     Set m_pExportFrame = Nothing
End Sub

Private Sub optProceed_Click()
354:     If optProceed.value = True Then
355:         Me.fraTileOptions.Enabled = False
356:     End If
End Sub

Private Sub optScale_Click()
360:     If optScale.value = True Then
361:         Me.fraTileOptions.Enabled = False
362:     End If
End Sub

Private Sub optTile_Click()
366:     If optTile.value = True Then
367:         Me.fraTileOptions.Enabled = True
368:         Me.optTileAll.value = True
369:     Else
370:         Me.fraTileOptions.Enabled = False
371:     End If
End Sub

Public Sub SetupToPrint(pPrinter As IPrinter, pPage As IPage, iCurrentPage As Integer, ByRef lDPI As Long, ByRef rectDeviceBounds As tagRECT, _
ByRef pVisBounds As IEnvelope, ByRef devFrameEnvelope As IEnvelope)
On Error GoTo ErrorHandler
  Dim idpi As Integer
  Dim pDeviceBounds As IEnvelope
  Dim paperWidthInch As Double
  Dim paperHeightInch As Double

382:   idpi = pPrinter.Resolution  'dots per inch
          
384:   Set pDeviceBounds = New Envelope
              
386:   pPage.GetDeviceBounds pPrinter, iCurrentPage, 0, idpi, pDeviceBounds
               
388:   rectDeviceBounds.Left = pDeviceBounds.XMin
389:   rectDeviceBounds.Top = pDeviceBounds.YMin
390:   rectDeviceBounds.Right = pDeviceBounds.XMax
391:   rectDeviceBounds.bottom = pDeviceBounds.YMax
  
  'Following block added 6/19/03 to fix problem with plots being cutoff
394:   If TypeOf pPrinter Is IEmfPrinter Then
    ' For emf printers we have to remove the top and left unprintable area
    ' from device coordinates so its origin is 0,0.
    '
398:     rectDeviceBounds.Right = rectDeviceBounds.Right - rectDeviceBounds.Left
399:     rectDeviceBounds.bottom = rectDeviceBounds.bottom - rectDeviceBounds.Top
400:     rectDeviceBounds.Left = 0
401:     rectDeviceBounds.Top = 0
402:   End If
  
404:   Set pVisBounds = New Envelope
405:   pPage.GetPageBounds pPrinter, iCurrentPage, 0, pVisBounds
406:   pPrinter.QueryPaperSize paperWidthInch, paperHeightInch
407:   Set devFrameEnvelope = New Envelope
408:   devFrameEnvelope.PutCoords 0, 0, paperWidthInch * idpi, paperHeightInch * idpi
  
410:   lDPI = CLng(idpi)

  Exit Sub
ErrorHandler:
414:   MsgBox "SetupToPrint - " & Err.Description
End Sub

Public Sub CheckNumberOfPages(pPage As IPage, pPrinter As IPrinter, ByRef iNumPages As Integer)
On Error GoTo ErrorHandler
419:   pPage.PrinterPageCount pPrinter, 0, iNumPages
      
421:   If Me.optTile.value = True Then
422:     If Me.optPages.value = True Then
      Dim iPageNo As Integer
      Dim sPageNo As String
425:       sPageNo = Me.txtTo.Text
      
427:       If sPageNo <> "" Then
428:           iPageNo = CInt(sPageNo)
429:       Else
          Exit Sub
431:       End If
      
433:       If iPageNo < iNumPages Then
434:           iNumPages = iPageNo
435:       End If
436:     End If
437:   End If
  
  Exit Sub
ErrorHandler:
441:   MsgBox "CheckNumberOfPages - " & Err.Description
End Sub
