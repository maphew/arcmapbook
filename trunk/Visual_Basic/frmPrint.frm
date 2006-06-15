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

Private m_pMapPage As IDSMapPage
Private m_pMapSeries As IDSMapSeries
Private m_pMapBook As IDSMapBook
Private m_pApp As IApplication
Private m_pExportFrame As IModelessFrame

Private Sub chkPrintToFile_Click()
8:   If Me.chkPrintToFile.Value = 1 Then
9:     Me.txtCopies.Text = 1
10:     Me.fraCopies.Enabled = False
11:     Me.txtCopies.Enabled = False
12:     Me.UpDown1.Enabled = False
13:     Me.lblNumberofCopies.Enabled = False
14:   Else
15:     fraCopies.Enabled = True
16:     Me.txtCopies.Enabled = True
17:     Me.UpDown1.Enabled = True
18:     Me.lblNumberofCopies.Enabled = True
19:   End If
End Sub

Private Sub cmdCancel_Click()
23:     m_pExportFrame.Visible = False
24:     Unload Me
End Sub

Public Property Let ExportFrame(ByVal pExportFrame As IModelessFrame)
28:     Set m_pExportFrame = pExportFrame
End Property

Public Property Get aDSMapPage() As IDSMapPage
32:     Set aDSMapPage = m_pMapPage
End Property

Public Property Let aDSMapPage(ByVal pMapPage As IDSMapPage)
36:     Set m_pMapPage = pMapPage
End Property

Public Property Get aDSMapSeries() As IDSMapSeries
40:     Set aDSMapSeries = m_pMapSeries
End Property

Public Property Let aDSMapSeries(ByVal pMapSeries As IDSMapSeries)
44:     Set m_pMapSeries = pMapSeries
End Property

Public Property Get aDSMapBook() As IDSMapBook
48:     Set aDSMapBook = m_pMapBook
End Property

Public Property Let aDSMapBook(ByVal pMapBook As IDSMapBook)
52:     Set m_pMapBook = pMapBook
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
  
67:   Set pMouse = New MouseCursor
68:   pMouse.SetCursor 2

70:   Set pMxApp = m_pApp
71:   Set pPrinter = pMxApp.Printer
72:   Set pMxDoc = m_pApp.Document
73:   Set pLayout = pMxDoc.PageLayout
74:   Set pPage = pLayout.Page
  
76:   If Me.chkPrintToFile.Value = 1 Then
'    If UCase(pPrinter.FileExtension) = "PS" Then
78:       Me.dlgPrint.Filter = "Postscript Files (*.ps,*.eps)|*.ps,*.eps"
'    Else
'      Me.dlgPrint.Filter = UCase(pPrinter.FileExtension) & " (*." & LCase(pPrinter.FileExtension) & ")" & "|*." & LCase(pPrinter.FileExtension)
'    End If
    
83:     Me.dlgPrint.DialogTitle = "Print to File"
'    Me.Hide
85:     m_pExportFrame.Visible = False
86:     Me.dlgPrint.ShowSave
    
    Dim sFileName As String, sPrefix As String, sExt As String, sSplit() As String
    
90:     sFileName = Me.dlgPrint.FileName
91:     If sFileName <> "" Then
92:       If InStr(1, sFileName, ".", vbTextCompare) > 0 Then
93:         sSplit = Split(sFileName, ".", , vbTextCompare)
94:         sPrefix = sSplit(0)
95:         sExt = sSplit(1)
96:       Else
97:         sPrefix = sFileName
98:         sExt = "ps"
99:         sFileName = sFileName & ".ps"
100:       End If
101:     Else
102:       MsgBox "Please specify a file name for the page(s)"
'      Me.Show
104:       m_pExportFrame.Visible = True
      Exit Sub
106:     End If
107:   End If
  
109:   If Me.optTile.Value = True Then
110:       pPage.PageToPrinterMapping = esriPageMappingTile
111:   ElseIf Me.optScale = True Then
112:       pPage.PageToPrinterMapping = esriPageMappingScale
113:   ElseIf Me.optProceed.Value = True Then
114:       pPage.PageToPrinterMapping = esriPageMappingCrop
115:   End If
  
117:   pPrinter.Paper.Orientation = pLayout.Page.Orientation
  
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
  
135:   Set PagesToPrint = New Collection
  
137:   If Not m_pMapPage Is Nothing Then
138:       PagesToPrint.Add m_pMapPage
139:   End If
  
141:   If m_pMapPage Is Nothing And m_pMapBook Is Nothing Then
142:     If frmPrint.optPrintAll.Value = True Then
143:       For i = 0 To m_pMapSeries.PageCount - 1
144:         If chkDisabled.Value = 1 Then
145:           If m_pMapSeries.Page(i).EnablePage Then
146:             PagesToPrint.Add m_pMapSeries.Page(i)
147:           End If
148:         Else
149:           PagesToPrint.Add m_pMapSeries.Page(i)
150:         End If
151:       Next i
152:     ElseIf frmPrint.optPrintPages.Value = True Then
      'parse out the pages to print
154:       If chkDisabled.Value = 1 Then
155:         Set PagesToPrint = ParseOutPages(Me.txtPrintPages.Text, m_pMapSeries, True)
156:       Else
157:         Set PagesToPrint = ParseOutPages(Me.txtPrintPages.Text, m_pMapSeries, False)
158:       End If
      If PagesToPrint.count = 0 Then Exit Sub
160:     End If
161:   End If
      
163:   numPages = CLng(Me.txtCopies.Text)
  
165:   If PagesToPrint.count > 0 Then
166:     Set pSeriesOpts = m_pMapSeries
167:     Set pSeriesOpts2 = pSeriesOpts
168:     If pSeriesOpts2.ClipData > 0 Then
169:       g_bClipFlag = True
170:     End If
171:     If pSeriesOpts.RotateFrame Then
172:       g_bRotateFlag = True
173:     End If
174:     If pSeriesOpts.LabelNeighbors Then
175:       g_bLabelNeighbors = True
176:     End If
177:     For i = 1 To PagesToPrint.count
178:       Set pMapPage = PagesToPrint.Item(i)
179:       pMapPage.DrawPage pMxDoc, m_pMapSeries, False
180:       CheckNumberOfPages pPage, pPrinter, iNumPages
181:       lblPrintStatus.Caption = "Printing page " & pMapPage.PageName & " ..."
        
183:       For iCurrentPage = 1 To iNumPages
184:         SetupToPrint pPrinter, pPage, iCurrentPage, lDPI, rectDeviceBounds, pVisBounds, devFrameEnvelope
185:         If Me.chkPrintToFile.Value = 1 Then
186:           If pPage.PageToPrinterMapping = esriPageMappingTile Then
187:             pPrinter.PrintToFile = sPrefix & "_" & pMapPage.PageName & "_" & iCurrentPage & "." & sExt
188:           Else
189:             pPrinter.PrintToFile = sPrefix & "_" & pMapPage.PageName & "." & sExt
190:           End If
191:         End If
192:         For a = 1 To numPages
193:           hdc = pPrinter.StartPrinting(devFrameEnvelope, 0)
194:             pMxDoc.ActiveView.Output hdc, lDPI, rectDeviceBounds, pVisBounds, Nothing
195:             pMapPage.LastOutputted = Format(Date, "mm/dd/yyyy")
196:           pPrinter.FinishPrinting
197:         Next a
198:       Next iCurrentPage
199:     Next i
200:   End If
  
202:   If Not m_pMapBook Is Nothing Then
    Dim pSeriesCount As Long
    Dim MapSeriesColl As Collection
    Dim pMapSeries As IDSMapSeries
    Dim count As Long
    
208:     pSeriesCount = m_pMapBook.ContentCount
    
210:     Set MapSeriesColl = New Collection
    
212:     For i = 0 To pSeriesCount - 1
213:         MapSeriesColl.Add m_pMapBook.ContentItem(i)
214:     Next i

    If MapSeriesColl.count = 0 Then Exit Sub
    
218:     For i = 1 To MapSeriesColl.count
219:       Set PagesToPrint = New Collection
220:       Set pMapSeries = MapSeriesColl.Item(i)
221:       Set pSeriesOpts = pMapSeries
222:       Set pSeriesOpts2 = pSeriesOpts
      
224:       If pSeriesOpts2.ClipData > 0 Then
225:         g_bClipFlag = True
226:       End If
227:       If pSeriesOpts.RotateFrame Then
228:         g_bRotateFlag = True
229:       End If
230:       If pSeriesOpts.LabelNeighbors Then
231:         g_bLabelNeighbors = True
232:       End If
        
234:       For count = 0 To pMapSeries.PageCount - 1
235:         If chkDisabled.Value = 1 Then
236:           If pMapSeries.Page(count).EnablePage Then
237:             PagesToPrint.Add pMapSeries.Page(count)
238:           End If
239:         Else
240:           PagesToPrint.Add pMapSeries.Page(count)
241:         End If
242:       Next count
      
244:       For count = 1 To PagesToPrint.count
      'now do printing
246:         Set pMapPage = PagesToPrint.Item(count)
247:         pMapPage.DrawPage pMxDoc, pMapSeries, False
        
249:         CheckNumberOfPages pPage, pPrinter, iNumPages
250:         lblPrintStatus.Caption = "Printing page " & pMapPage.PageName & " ..."
            
252:         For iCurrentPage = 1 To iNumPages
253:           SetupToPrint pPrinter, pPage, iCurrentPage, lDPI, rectDeviceBounds, pVisBounds, devFrameEnvelope
254:           If Me.chkPrintToFile.Value = 1 Then
255:             If pPage.PageToPrinterMapping = esriPageMappingTile Then
256:               pPrinter.PrintToFile = sPrefix & "_" & pMapPage.PageName & "_" & iCurrentPage & "." & sExt
257:             Else
258:               pPrinter.PrintToFile = sPrefix & "_" & pMapPage.PageName & "." & sExt
259:             End If
260:           End If
261:           For a = 1 To numPages
262:             hdc = pPrinter.StartPrinting(devFrameEnvelope, 0)
263:               pMxDoc.ActiveView.Output hdc, lDPI, rectDeviceBounds, pVisBounds, Nothing
264:               pMapPage.LastOutputted = Format(Date, "mm/dd/yyyy")
265:             pPrinter.FinishPrinting
266:           Next a
267:         Next iCurrentPage
      
269:       Next count
            
271:     Next i
272:   End If
                                   
274:   lblPrintStatus.Caption = ""
275:   Set m_pMapBook = Nothing
276:   Set m_pMapPage = Nothing
277:   Set m_pMapSeries = Nothing
278:   m_pExportFrame.Visible = False
279:   Unload Me

  Exit Sub
ErrorHandler:
283:   lblPrintStatus.Caption = ""
284:   MsgBox "cmdOK_Click - " & Err.Description
End Sub

Public Property Get Application() As IApplication
288:     Set Application = m_pApp
End Property

Public Property Let Application(ByVal pApp As IApplication)
292:     Set m_pApp = pApp
End Property

Private Sub cmdSetup_Click()
296:   If (Not m_pApp.IsDialogVisible(esriMxDlgPageSetup)) Then
    Dim bDialog As Boolean
    Dim pPrinter As IPrinter
    Dim pMxApp As IMxApplication
300:     m_pApp.ShowDialog esriMxDlgPageSetup, True
    
302:     m_pExportFrame.Visible = False
'    Me.Hide
304:     bDialog = True
    
306:     While bDialog = True
307:         bDialog = m_pApp.IsDialogVisible(esriMxDlgPageSetup)
308:         DoEvents
        
'            Sleep 1
    
312:     Wend
    
314:     Set pMxApp = m_pApp
315:     Set pPrinter = pMxApp.Printer
316:     frmPrint.lblName.Caption = pPrinter.Paper.PrinterName
317:     frmPrint.lblType.Caption = pPrinter.DriverName
318:     If TypeOf pPrinter Is IPsPrinter Then
319:       Me.chkPrintToFile.Enabled = True
320:     Else
321:       Me.chkPrintToFile.Value = 0
322:       Me.chkPrintToFile.Enabled = False
323:     End If
'    Me.Show
325:     m_pExportFrame.Visible = True
326:   End If
End Sub

Private Sub Form_Load()
330:   chkDisabled.Value = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
334:     Set m_pApp = Nothing
335:     Set m_pMapPage = Nothing
336:     Set m_pMapSeries = Nothing
337:     Set m_pMapBook = Nothing
338:     Set m_pExportFrame = Nothing
End Sub

Private Sub optProceed_Click()
342:     If optProceed.Value = True Then
343:         Me.fraTileOptions.Enabled = False
344:     End If
End Sub

Private Sub optScale_Click()
348:     If optScale.Value = True Then
349:         Me.fraTileOptions.Enabled = False
350:     End If
End Sub

Private Sub optTile_Click()
354:     If optTile.Value = True Then
355:         Me.fraTileOptions.Enabled = True
356:         Me.optTileAll.Value = True
357:     Else
358:         Me.fraTileOptions.Enabled = False
359:     End If
End Sub

Public Sub SetupToPrint(pPrinter As IPrinter, pPage As IPage, iCurrentPage As Integer, ByRef lDPI As Long, ByRef rectDeviceBounds As tagRECT, _
ByRef pVisBounds As IEnvelope, ByRef devFrameEnvelope As IEnvelope)
On Error GoTo ErrorHandler
  Dim idpi As Integer
  Dim pDeviceBounds As IEnvelope
  Dim paperWidthInch As Double
  Dim paperHeightInch As Double

370:   idpi = pPrinter.Resolution  'dots per inch
          
372:   Set pDeviceBounds = New Envelope
              
374:   pPage.GetDeviceBounds pPrinter, iCurrentPage, 0, idpi, pDeviceBounds
               
376:   rectDeviceBounds.Left = pDeviceBounds.XMin
377:   rectDeviceBounds.Top = pDeviceBounds.YMin
378:   rectDeviceBounds.Right = pDeviceBounds.XMax
379:   rectDeviceBounds.bottom = pDeviceBounds.YMax
  
  'Following block added 6/19/03 to fix problem with plots being cutoff
382:   If TypeOf pPrinter Is IEmfPrinter Then
    ' For emf printers we have to remove the top and left unprintable area
    ' from device coordinates so its origin is 0,0.
    '
386:     rectDeviceBounds.Right = rectDeviceBounds.Right - rectDeviceBounds.Left
387:     rectDeviceBounds.bottom = rectDeviceBounds.bottom - rectDeviceBounds.Top
388:     rectDeviceBounds.Left = 0
389:     rectDeviceBounds.Top = 0
390:   End If
  
392:   Set pVisBounds = New Envelope
393:   pPage.GetPageBounds pPrinter, iCurrentPage, 0, pVisBounds
394:   pPrinter.QueryPaperSize paperWidthInch, paperHeightInch
395:   Set devFrameEnvelope = New Envelope
396:   devFrameEnvelope.PutCoords 0, 0, paperWidthInch * idpi, paperHeightInch * idpi
  
398:   lDPI = CLng(idpi)

  Exit Sub
ErrorHandler:
402:   MsgBox "SetupToPrint - " & Err.Description
End Sub

Public Sub CheckNumberOfPages(pPage As IPage, pPrinter As IPrinter, ByRef iNumPages As Integer)
On Error GoTo ErrorHandler
407:   pPage.PrinterPageCount pPrinter, 0, iNumPages
      
409:   If frmPrint.optTile.Value = True Then
410:     If frmPrint.optPages.Value = True Then
      Dim iPageNo As Integer
      Dim sPageNo As String
413:       sPageNo = frmPrint.txtTo.Text
      
415:       If sPageNo <> "" Then
416:           iPageNo = CInt(sPageNo)
417:       Else
          Exit Sub
419:       End If
      
421:       If iPageNo < iNumPages Then
422:           iNumPages = iPageNo
423:       End If
424:     End If
425:   End If
  
  Exit Sub
ErrorHandler:
429:   MsgBox "CheckNumberOfPages - " & Err.Description
End Sub
