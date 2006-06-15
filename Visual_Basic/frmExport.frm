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
   Begin VB.TextBox txtPath 
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
   Begin VB.ComboBox cmbExportType 
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
Private m_pExport As IExport
Private m_pExportFrame As IModelessFrame

Public Property Get aDSMapPage() As IDSMapPage
9:     Set aDSMapPage = m_pMapPage
End Property

Public Property Let aDSMapPage(ByVal pMapPage As IDSMapPage)
13:     Set m_pMapPage = pMapPage
End Property

Public Property Let ExportFrame(ByVal pExportFrame As IModelessFrame)
17:     Set m_pExportFrame = pExportFrame
End Property

Public Property Get aDSMapSeries() As IDSMapSeries
21:     Set aDSMapSeries = m_pMapSeries
End Property

Public Property Let aDSMapSeries(ByVal pMapSeries As IDSMapSeries)
25:     Set m_pMapSeries = pMapSeries
End Property

Public Property Get aDSMapBook() As IDSMapBook
29:     Set aDSMapBook = m_pMapBook
End Property

Public Property Let aDSMapBook(ByVal pMapBook As IDSMapBook)
33:     Set m_pMapBook = pMapBook
End Property

Public Property Get Application() As IApplication
37:     Set Application = m_pApp
End Property

Public Property Let Application(ByVal pApp As IApplication)
41:     Set m_pApp = pApp
End Property

Public Sub SetupDialog()
  On Error GoTo ErrorHandler
  
  Exit Sub
ErrorHandler:
49:   MsgBox "SetupDialog - " & Err.Description
End Sub


Private Sub cmbExportType_Click()

55: Set m_pExport = Nothing

If Me.txtPath.Text = "" Then Exit Sub

Dim sExt As String
60:     sExt = Me.cmbExportType.Text

62:     ChangeFileExtension sExt

End Sub

Private Sub cmdBrowse_Click()
Dim sFileExt As String
Dim sFileName As String

'    Me.dlgExport.Filter = "EMF (*.emf)|*.emf|CGM (*.cgm)|*.cgm|EPS (*.eps)|*.eps|AI (*.ai)|*.ai|PDF (*.pdf)|*.pdf|BMP (*.bmp)|*.bmp|TIFF (*.tif)|*.tif|JPEG (*.jpg)|*.jpg"
    
72:     Me.dlgExport.Filter = "BMP (*.bmp)|*.bmp|EPS (*.eps)|*.eps|JPEG (*.jpg)|*.jpg|PDF (*.pdf)|*.pdf|TIFF (*.tif)|*.tif"
   
74:     If Me.cmbExportType.ListIndex <> -1 Then
75:         Me.dlgExport.FilterIndex = Me.cmbExportType.ListIndex + 1
76:     Else
77:         Me.dlgExport.FilterIndex = 4
78:     End If
    
80:     Me.dlgExport.DialogTitle = "Export"
    
'    Me.Hide
83:     m_pExportFrame.Visible = False
    
85:     Me.dlgExport.ShowSave
    
87:     If Me.dlgExport.FileName = "" Then
88:         Me.Show
        Exit Sub
90:     Else
91:         sFileName = Me.dlgExport.FileName
92:     End If
    
94:      sFileExt = Right(sFileName, 3)
     
    Select Case sFileExt
        Case "emf"
98:             Me.cmbExportType.Text = "EMF (*.emf)"
'        Case "cgm"
'            Me.cmbExportType.Text = "CGM (*.cgm)"
        Case "eps"
102:             Me.cmbExportType.Text = "EPS (*.eps)"
        Case ".ai"
104:             Me.cmbExportType.Text = "AI (*.ai)"
        Case "pdf"
106:             Me.cmbExportType.Text = "PDF (*.pdf)"
        Case "bmp"
108:             Me.cmbExportType.Text = "BMP (*.bmp)"
        Case "tif"
110:             Me.cmbExportType.Text = "TIFF (*.tif)"
        Case "jpg"
112:             Me.cmbExportType.Text = "JPEG (*.jpg)"
113:     End Select
    
115:     Me.txtPath.Text = sFileName
    
'   Me.Show
118:   m_pExportFrame.Visible = True
  
End Sub

Private Sub cmdCancel_Click()
123:     m_pExportFrame.Visible = False
124:     Unload Me
End Sub

Private Sub cmdExport_Click()
On Error GoTo ErrHand:
  Dim sFileExt As String
  Dim pExport As IExport
  Dim pJpegExport As IExportJPEG
  Dim sFileName As String
  Dim pActiveView As IActiveView
  Dim pMxDoc As IMxDocument
  Dim pMouse As IMouseCursor
  
137:   If Me.txtPath.Text = "" Then
138:     MsgBox "You have not typed in a valid path!!!"
    Exit Sub
140:   End If
  
  Dim bValid As Boolean
143:   bValid = CheckForValidPath(Me.txtPath.Text)
    
145:   If bValid = False Then
146:     MsgBox "You have not typed in a valid path!!!"
    Exit Sub
148:   End If

  '***Need to make sure it's a valid path
  
152:   Set pMouse = New MouseCursor
153:   pMouse.SetCursor 2

155:   Set pMxDoc = m_pApp.Document
156:   sFileName = Left(Me.txtPath.Text, Len(Me.txtPath.Text) - 4)
157:   sFileExt = Right(Me.txtPath.Text, 3)
    
159:   If m_pExport Is Nothing Then
    Select Case sFileExt
    Case "emf"
162:       Set pExport = New ExportEMF
'    Case "cgm"
'      MsgBox "CGMExporter not supported at 9.0, need to change this code to the replacement."
'      Exit Sub
'      Set pExport = New CGMExporter
    Case "eps"
168:       Set pExport = New ExportPS
    Case ".ai"
170:       Set pExport = New ExportAI
    Case "pdf"
172:       Set pExport = New ExportPDF
      'Map the basic fonts
'174:       MapFonts pExport
    Case "bmp"
176:       Set pExport = New ExportBMP
    Case "tif"
178:       Set pExport = New ExportTIFF
    Case "jpg"
180:       Set pExport = New ExportJPEG
181:     End Select
182:   Else
183:     Set pExport = m_pExport
184:   End If
        
186:   If pExport Is Nothing Then
187:     MsgBox "No export object!!!"
    Exit Sub
189:   End If
   
  'Switch to the Layout view if we are not already there
192:   If Not TypeOf pMxDoc.ActiveView Is IPageLayout Then
193:     Set pMxDoc.ActiveView = pMxDoc.PageLayout
194:   End If

196:   Set pActiveView = pMxDoc.ActiveView
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
207:   Set PagesToExport = New Collection
208:   Set pSeriesOpts = m_pMapSeries
209:   Set pSeriesOpts2 = pSeriesOpts
  
211:   If Not m_pMapPage Is Nothing Then
212:       PagesToExport.Add m_pMapPage
213:   End If
  
215:   If Not m_pMapSeries Is Nothing And m_pMapPage Is Nothing And m_pMapBook Is Nothing Then
216:     If Me.optAll.Value = True Then
217:       For i = 0 To m_pMapSeries.PageCount - 1
218:         If Me.chkDisabled.Value = 1 Then
219:           If m_pMapSeries.Page(i).EnablePage Then
220:             PagesToExport.Add m_pMapSeries.Page(i)
221:           End If
222:          Else
223:             PagesToExport.Add m_pMapSeries.Page(i)
224:         End If
225:       Next i
226:     ElseIf Me.optPages.Value = True Then
      'parse out the pages to export
228:       If chkDisabled.Value = 1 Then
229:         Set PagesToExport = ParseOutPages(Me.txtPages.Text, m_pMapSeries, True)
230:       Else
231:         Set PagesToExport = ParseOutPages(Me.txtPages.Text, m_pMapSeries, False)
232:       End If
      If PagesToExport.count = 0 Then Exit Sub
234:     End If
235:   End If
  
237:   If PagesToExport.count > 0 Then
238:     If pSeriesOpts2.ClipData > 0 Then
239:       g_bClipFlag = True
240:     End If
241:     If pSeriesOpts.RotateFrame Then
242:       g_bRotateFlag = True
243:     End If
244:     If pSeriesOpts.LabelNeighbors Then
245:       g_bLabelNeighbors = True
246:     End If
247:     For i = 1 To PagesToExport.count
248:       Set pMapPage = PagesToExport.Item(i)
249:       pMapPage.DrawPage pMxDoc, m_pMapSeries, False
          
251:       If sFileExt = ".ai" Then
252:         sExportFile = sFileName & "_" & pMapPage.PageName & sFileExt
253:       Else
254:         sExportFile = sFileName & "_" & pMapPage.PageName & "." & sFileExt
255:       End If
256:       lblStatus.Caption = "Exporting to " & sExportFile & " ..."
257:       SetupToExport pExport, dpi, ExportFrame, pActiveView, sExportFile
      
      'Do the export
260:       hdc = pExport.StartExporting
261:         pActiveView.Output hdc, pExport.Resolution, ExportFrame, Nothing, Nothing
262:         pMapPage.LastOutputted = Format(Date, "mm/dd/yyyy")
263:       pExport.FinishExporting
264:     Next i
265:   End If
            
267:   If Not m_pMapBook Is Nothing Then
    Dim pMapSeries As IDSMapSeries
    Dim count As Long
270:     For i = 0 To m_pMapBook.ContentCount - 1
271:       Set PagesToExport = New Collection
272:       Set pMapSeries = m_pMapBook.ContentItem(i)
273:       Set pSeriesOpts = pMapSeries
    
275:       For count = 0 To pMapSeries.PageCount - 1
276:         If Me.chkDisabled.Value = 1 Then
277:           If pMapSeries.Page(count).EnablePage Then
278:             PagesToExport.Add pMapSeries.Page(count)
279:           End If
280:         Else
281:             PagesToExport.Add pMapSeries.Page(count)
282:         End If
283:       Next count
        
285:       If pSeriesOpts2.ClipData > 0 Then
286:         g_bClipFlag = True
287:       End If
288:       If pSeriesOpts.RotateFrame Then
289:         g_bRotateFlag = True
290:       End If
291:       If pSeriesOpts.LabelNeighbors Then
292:         g_bLabelNeighbors = True
293:       End If
294:       For count = 1 To PagesToExport.count
        'now do export
296:         Set pMapPage = PagesToExport.Item(count)
297:         pMapPage.DrawPage pMxDoc, pMapSeries, False
      
299:         If sFileExt = ".ai" Then
300:             sExportFile = sFileName & "_series_" & i & "_" & pMapPage.PageName & sFileExt
301:         Else
302:             sExportFile = sFileName & "_series_" & i & "_" & pMapPage.PageName & "." & sFileExt
303:         End If
304:         lblStatus.Caption = "Exporting to " & sExportFile & " ..."
305:         SetupToExport pExport, pExport.Resolution, ExportFrame, pActiveView, sExportFile
          
        'Do the export
308:         hdc = pExport.StartExporting
309:           pActiveView.Output hdc, pExport.Resolution, ExportFrame, Nothing, Nothing
310:           pMapPage.LastOutputted = Format(Date, "mm/dd/yyyy")
311:         pExport.FinishExporting
312:       Next count
313:     Next i
314:   End If

'  pActiveView.ScreenDisplay.DisplayTransformation.ZoomResolution = True
317:   If TypeOf pExport Is IOutputCleanup Then
    Dim pCleanup As IOutputCleanup
319:     Set pCleanup = pExport
320:     pCleanup.Cleanup
321:   End If
  
323:   lblStatus.Caption = ""
324:   Set m_pMapBook = Nothing
325:   Set m_pMapPage = Nothing
326:   Set m_pMapSeries = Nothing
327:   m_pExportFrame.Visible = False
328:   Unload Me
  
  Exit Sub
ErrHand:
332:   lblStatus.Caption = ""
333:   MsgBox "cmdExport_Click - " & Err.Description
End Sub

Private Sub MapFonts(pExport As IExport)
On Error GoTo ErrHand:
  If Not TypeOf pExport Is IFontMapEnvironment Then Exit Sub
  
  Dim pFontMapEnv As IFontMapEnvironment, pFontMapColl As IFontMapCollection
  Dim pFontMap As IFontMap2
342:   Set pFontMapEnv = pExport
343:   Set pFontMapColl = pFontMapEnv.FontMapCollection
344:   Set pFontMap = New FontMap
345:   pFontMap.SetMapping "Arial", "Helvetica"
346:   pFontMapColl.Add pFontMap
347:   Set pFontMap = New FontMap
348:   pFontMap.SetMapping "Arial Bold", "Helvetica-Bold"
349:   pFontMapColl.Add pFontMap
350:   Set pFontMap = New FontMap
351:   pFontMap.SetMapping "Arial Bold Italic", "Helvetica-BoldOblique"
352:   pFontMapColl.Add pFontMap
353:   Set pFontMap = New FontMap
354:   pFontMap.SetMapping "Arial Italic", "Helvetica-Oblique"
355:   pFontMapColl.Add pFontMap
356:   Set pFontMap = New FontMap
357:   pFontMap.SetMapping "Courier New", "Courier"
358:   pFontMapColl.Add pFontMap
359:   Set pFontMap = New FontMap
360:   pFontMap.SetMapping "Courier New Bold", "Courier-Bold"
361:   pFontMapColl.Add pFontMap
362:   Set pFontMap = New FontMap
363:   pFontMap.SetMapping "Courier New Bold Italic", "Courier-BoldOblique"
364:   pFontMapColl.Add pFontMap
365:   Set pFontMap = New FontMap
366:   pFontMap.SetMapping "Courier New Italic", "Courier-Oblique"
367:   pFontMapColl.Add pFontMap
368:   Set pFontMap = New FontMap
369:   pFontMap.SetMapping "Symbol", "Symbol"
370:   pFontMapColl.Add pFontMap
371:   Set pFontMap = New FontMap
372:   pFontMap.SetMapping "Times New Roman", "Times-Roman"
373:   pFontMapColl.Add pFontMap
374:   Set pFontMap = New FontMap
375:   pFontMap.SetMapping "Times New Roman Bold", "Times-Bold"
376:   pFontMapColl.Add pFontMap
377:   Set pFontMap = New FontMap
378:   pFontMap.SetMapping "Times New Roman Bold Italic", "Times-BoldItalic"
379:   pFontMapColl.Add pFontMap
380:   Set pFontMap = New FontMap
381:   pFontMap.SetMapping "Times New Roman Italic", "Times-Italic"
382:   pFontMapColl.Add pFontMap
  
  Exit Sub
ErrHand:
386:   MsgBox "MapFonts - " & Err.Description
End Sub

Public Sub InitializeTheForm()
    
391:     Me.cmbExportType.Clear
'    Me.cmbExportType.AddItem "EMF (*.emf)"
'    Me.cmbExportType.AddItem "CGM (*.cgm)"
'    Me.cmbExportType.AddItem "EPS (*.eps)"
'    Me.cmbExportType.AddItem "AI (*.ai)"
396:     Me.cmbExportType.AddItem "BMP (*.bmp)"
397:     Me.cmbExportType.AddItem "EPS (*.eps)"
398:     Me.cmbExportType.AddItem "JPEG (*.jpg)"
399:     Me.cmbExportType.AddItem "PDF (*.pdf)"
400:     Me.cmbExportType.AddItem "TIFF (*.tif)"
    
'    Me.cmbExportType.Text = "JPEG (*.jpg)"
    
404:     Me.cmbExportType.ListIndex = 2
    
End Sub

Private Sub ChangeFileExtension(sFileType As String)

Dim sExt As String
411:     sExt = Right(sFileType, 4)
412:     sExt = Left(sExt, 3)
    
Dim sFileName As String
Dim sFileNameExt As String

417:     sFileName = Me.txtPath.Text
418:     sFileNameExt = Right(sFileName, 3)
    
420:     If sExt <> sFileNameExt Then
        Dim aFileName() As String
        
423:         aFileName = Split(sFileName, ".")
        
425:         If sExt <> ".ai" Then
426:             Me.txtPath.Text = aFileName(0) & "." & sExt
427:         Else
428:             Me.txtPath.Text = aFileName(0) & sExt
429:         End If
    
431:     End If
    
End Sub

'Private Sub cmdOptions_Click()
'  On Error GoTo ErrorHandler
'
'  Dim sFileExt As String
'  sFileExt = Me.cmbExportType.Text
'
'  Dim pExportSet As ISet
'  Dim sTitle As String
'  Dim pMyPage As IComPropertyPage   'build the property page
'  Dim pMyPage2 As IComPropertyPage
'
'  'Set m_pExport = Nothing
'
'  Set pExportSet = New esriSystem.Set
'
'  Select Case sFileExt
'  Case "EMF (*.emf)"
'    If m_pExport Is Nothing Then
'      Set m_pExport = New ExportEMF
'    Else
'      If Not TypeOf m_pExport Is IExportEMF Then
'        Set m_pExport = New ExportEMF
'      End If
'    End If
'    sTitle = "EMF Options"
'    Set pMyPage = New EmfExporterPropertyPage
''CGM is no longer supported at 9.0
''  Case "CGM (*.cgm)"
''    If m_pExporter Is Nothing Then
''      Set m_pExporter = New CGMExporter
''    Else
''      If Not TypeOf m_pExport Is ICGMExporter Then
''        Set m_pExport = New CGMExporter
''      End If
''    End If
''    sTitle = "CGM Options"
''    Set pMyPage = New CGMExporterPropertyPage
''  Case "AI (*.ai)"
''    If m_pExport Is Nothing Then
''      Set m_pExport = New exportai
''    Else
''      If Not TypeOf m_pExport Is IExportAI Then
''        Set m_pExport = New Exportai
''      End If
''    End If
''    sTitle = "AI Options"
''    Set pMyPage = New AIExporterPropertyPage
'  Case "EPS (*.eps)"
'    If m_pExport Is Nothing Then
'      Set m_pExport = New ExportPS
'    Else
'      If Not TypeOf m_pExport Is IExportPS Then
'        Set m_pExport = New ExportPS
'      End If
'    End If
'    sTitle = "EPS Options"
'    Set pMyPage = New PDFExporterPropertyPage
'  Case "PDF (*.pdf)"
'    If m_pExport Is Nothing Then
'      Set m_pExport = New ExportPDF
'    Else
'      If Not TypeOf m_pExport Is IExportPDF Then
'        Set m_pExport = New ExportPDF
'      End If
'    End If
'    sTitle = "PDF Options"
'    Set pMyPage = New PDFExporterPropertyPage
'    Set pMyPage2 = New FontMappingPropertyPage
'  Case "BMP (*.bmp)"
'    If m_pExport Is Nothing Then
'      Set m_pExport = New ExportBMP
'    Else
'      If Not TypeOf m_pExport Is IExportBMP Then
'        Set m_pExport = New ExportBMP
'      End If
'    End If
'    sTitle = "BMP Options"
'    Set pMyPage = New DibExporterPropertyPage
'  Case "TIFF (*.tif)"
'    If m_pExport Is Nothing Then
'      Set m_pExport = New ExportTIFF
'    Else
'      If Not TypeOf m_pExport Is IExportTIFF Then
'        Set m_pExport = New ExportTIFF
'      End If
'    End If
'    sTitle = "TIFF Options"
'    Set pMyPage = New TiffExporterPropertyPage
'  Case "JPEG (*.jpg)"
'    If m_pExport Is Nothing Then
'      Set m_pExport = New ExportJPEG
'    Else
'      If Not TypeOf m_pExport Is IExportJPEG Then
'        Set m_pExport = New ExportJPEG
'      End If
'    End If
'    sTitle = "JPEG Options"
'    Set pMyPage = New JpegExporterPropertyPage
'  End Select
'
'  If m_pExport Is Nothing Then Exit Sub
'
'  pExportSet.Add m_pExport
'
'  Dim pPS As IComPropertySheet
'
'  Set pPS = New ComPropertySheet
'
'  If Not pMyPage Is Nothing Then
'    pPS.AddPage pMyPage
'  End If
'
'  If Not pMyPage2 Is Nothing Then
'    pPS.AddPage pMyPage2
'  End If
'
''  Me.Hide
'  m_pExportFrame.Visible = False
'
'  If pPS.CanEdit(pExportSet) = True Then
'    pPS.Title = sTitle
'    pPS.EditProperties pExportSet, m_pApp.hwnd 'show the property sheet
'  End If
'
''  Me.Show
'  m_pExportFrame.Visible = True
'
'
''  If pMyPage.IsPageDirty = True Then
'    pMyPage.Apply
''  End If
'
'  Exit Sub
'ErrorHandler:
'  MsgBox "cmdOptions_Click - " & Err.Description
'End Sub

Private Sub cmdOptions_Click()
  On Error GoTo ErrorHandler

  Dim sFileExt As String
576:   sFileExt = Me.cmbExportType.Text
      
  Dim sTitle As String

  Select Case sFileExt
  Case "EMF (*.emf)"
582:     If m_pExport Is Nothing Then
583:       Set m_pExport = New ExportEMF
584:     Else
585:       If Not TypeOf m_pExport Is IExportEMF Then
586:         Set m_pExport = New ExportEMF
587:       End If
588:     End If
589:     sTitle = "EMF Options"
'  Case "AI (*.ai)"
'    If m_pExport Is Nothing Then
'      Set m_pExport = New exportai
'    Else
'      If Not TypeOf m_pExport Is IExportAI Then
'        Set m_pExport = New Exportai
'      End If
'    End If
'    sTitle = "AI Options"
'    Set pMyPage = New AIExporterPropertyPage
  Case "EPS (*.eps)"
601:     If m_pExport Is Nothing Then
602:       Set m_pExport = New ExportPS
603:     Else
604:       If Not TypeOf m_pExport Is IExportPS Then
605:         Set m_pExport = New ExportPS
606:       End If
607:     End If
608:     sTitle = "EPS Options"
  Case "PDF (*.pdf)"
610:     If m_pExport Is Nothing Then
611:       Set m_pExport = New ExportPDF
612:     Else
613:       If Not TypeOf m_pExport Is IExportPDF Then
614:         Set m_pExport = New ExportPDF
615:       End If
616:     End If
617:     sTitle = "PDF Options"
  Case "BMP (*.bmp)"
619:     If m_pExport Is Nothing Then
620:       Set m_pExport = New ExportBMP
621:     Else
622:       If Not TypeOf m_pExport Is IExportBMP Then
623:         Set m_pExport = New ExportBMP
624:       End If
625:     End If
626:     sTitle = "BMP Options"
  Case "TIFF (*.tif)"
628:     If m_pExport Is Nothing Then
629:       Set m_pExport = New ExportTIFF
630:     Else
631:       If Not TypeOf m_pExport Is IExportTIFF Then
632:         Set m_pExport = New ExportTIFF
633:       End If
634:     End If
635:     sTitle = "TIFF Options"
  Case "JPEG (*.jpg)"
637:     If m_pExport Is Nothing Then
638:       Set m_pExport = New ExportJPEG
639:     Else
640:       If Not TypeOf m_pExport Is IExportJPEG Then
641:         Set m_pExport = New ExportJPEG
642:       End If
643:     End If
644:     sTitle = "JPEG Options"
645:   End Select

  If m_pExport Is Nothing Then Exit Sub
  
'  Me.Hide
650:   m_pExportFrame.Visible = False
              
652:   Set frmExportPropDlg.Export = m_pExport
653:   frmExportPropDlg.Caption = sTitle
654:   frmExportPropDlg.Show vbModal, Me
  
  'The ExportSVG class has a Compression property that changes the value of the Filter property,
  ' and we must syncronize our file extension to account for the possible change.
658:   If TypeOf m_pExport Is IExportSVG Then
659:     cboSaveAsType.List(cboSaveAsType.ListIndex) = Split(m_pExport.Filter, "|")(0)
660:     m_sFileExtension = Split(Split(cboSaveAsType.Text, "(")(1), ")")(0)
661:     m_sFileExtension = Right(m_sFileExtension, Len(m_sFileExtension) - 1)
662:     txtFileName.Text = "Unititled" & m_sFileExtension
663:   End If
              
'  Me.Show
666:   m_pExportFrame.Visible = True
        
  Exit Sub
ErrorHandler:
670:   MsgBox "cmdOptions_Click - " & Err.Description
End Sub

Public Sub SetupToExport(ByRef pExport As IExport, ByRef dpi As Integer, ByRef ExportFrame As tagRECT, pActiveView As IActiveView, sExportFileName As String)
  On Error GoTo ErrorHandler
  
  Dim pEnv As IEnvelope, pPageLayout As IPageLayout, pPage As IPage
  Dim dXmax As Double, dYmax As Double
  
679:    Set pEnv = New Envelope
'   pActiveView.ScreenDisplay.DisplayTransformation.Resolution = pExport.Resolution
  'Setup the Export
682:   ExportFrame = pActiveView.ExportFrame

684:   Set pPageLayout = pActiveView
685:   Set pPage = pPageLayout.Page
  
687:   If pPage.Units <> esriInches Then
688:     pPage.Units = esriInches
689:   End If
  
691:   pPage.QuerySize dXmax, dYmax
692:   pEnv.PutCoords 0, 0, dXmax * pExport.Resolution, dYmax * pExport.Resolution

'Commented out code removes a quarter of a unit, most likely an inch, from the extent to make it
'fit better on the page
'  ExportFrame.Top = pExport.Resolution * 0.25
'  ExportFrame.Right = (dXmax - 0.25) * pExport.Resolution
698:   ExportFrame.Right = dXmax * pExport.Resolution
699:   ExportFrame.bottom = dYmax * pExport.Resolution
  
701:   ExportFrame.Left = 0
702:   ExportFrame.Top = 0
            
704:   With pExport
705:     .PixelBounds = pEnv
706:     .ExportFileName = sExportFileName
707:   End With

  
  Exit Sub
ErrorHandler:
712:   MsgBox "SetupToExport - " & Err.Description
End Sub


Public Function ConvertToPixels(sOrient As String, pExport As IExport) As Double
On Error GoTo ErrHand:
  Dim pixelExtent As Long
  Dim pDT As IDisplayTransformation
  Dim deviceRECT As tagRECT
  Dim pMxDoc As IMxDocument
  
723:   Set pMxDoc = m_pApp.Document
724:   Set pDT = pMxDoc.ActiveView.ScreenDisplay.DisplayTransformation
725:   deviceRECT = pDT.DeviceFrame
  
727:   If sOrient = "Height" Then
728:     pixelExtent = Abs(deviceRECT.Top - deviceRECT.bottom)
729:   ElseIf sOrient = "Width" Then
730:     pixelExtent = Abs(deviceRECT.Top - deviceRECT.bottom)
731:   End If
  
733:   ConvertToPixels = (pExport.Resolution * (pixelExtent / pDT.Resolution))
  
  Exit Function
ErrHand:
737:   MsgBox "ConvertToPixels - " & Err.Description
End Function

Private Sub Form_Load()
741:   chkDisabled.Value = 1
End Sub

Private Function CheckForValidPath(sPathName As String) As Boolean
  On Error GoTo ErrorHandler

747:   CheckForValidPath = False
  
  Dim aPath() As String
750:       aPath = Split(sPathName, ".")

752:   If UBound(aPath) = 0 Then
    Exit Function
754:   ElseIf UBound(aPath) = 1 Then
    
    Dim sPath As String
    Dim lPos As Long
    
759:       lPos = InStrRev(sPathName, "\")
760:       sPath = Left$(sPathName, (Len(sPathName) - (Len(sPathName) - lPos + 1)))
      
762:       If Dir(sPath, vbDirectory) <> "" Then
763:         CheckForValidPath = True
        Exit Function
765:       Else
        Exit Function
767:       End If
      
769:   ElseIf UBound(aPath) > 1 Then
    Exit Function
771:   End If
  
  Exit Function
ErrorHandler:
775:   MsgBox "CheckForValidPath - " & Err.Description
End Function

Private Sub Form_Unload(Cancel As Integer)
779:   Set m_pMapPage = Nothing
780:   Set m_pMapSeries = Nothing
781:   Set m_pMapBook = Nothing
782:   Set m_pApp = Nothing
783:   Set m_pExport = Nothing
784:   Set m_pExportFrame = Nothing
End Sub
