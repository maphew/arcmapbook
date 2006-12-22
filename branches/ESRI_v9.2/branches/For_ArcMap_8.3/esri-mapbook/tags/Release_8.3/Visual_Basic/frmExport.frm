VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
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
Private m_pMapPage As IDSMapPage
Private m_pMapSeries As IDSMapSeries
Private m_pMapBook As IDSMapBook
Private m_pApp As IApplication
Private m_pExporter As IExporter

Public Property Get aDSMapPage() As IDSMapPage
    Set aDSMapPage = m_pMapPage
End Property

Public Property Let aDSMapPage(ByVal pMapPage As IDSMapPage)
    Set m_pMapPage = pMapPage
End Property

Public Property Get aDSMapSeries() As IDSMapSeries
    Set aDSMapSeries = m_pMapSeries
End Property

Public Property Let aDSMapSeries(ByVal pMapSeries As IDSMapSeries)
    Set m_pMapSeries = pMapSeries
End Property

Public Property Get aDSMapBook() As IDSMapBook
    Set aDSMapBook = m_pMapBook
End Property

Public Property Let aDSMapBook(ByVal pMapBook As IDSMapBook)
    Set m_pMapBook = pMapBook
End Property

Public Property Get Application() As IApplication
    Set Application = m_pApp
End Property

Public Property Let Application(ByVal pApp As IApplication)
    Set m_pApp = pApp
End Property

Public Sub SetupDialog()
  On Error GoTo ErrorHandler
  
  Exit Sub
ErrorHandler:
  MsgBox "SetupDialog - " & Err.Description
End Sub


Private Sub cmbExportType_Click()

Set m_pExporter = Nothing

If Me.txtPath.Text = "" Then Exit Sub

Dim sExt As String
    sExt = Me.cmbExportType.Text

    ChangeFileExtension sExt

End Sub

Private Sub cmdBrowse_Click()
Dim sFileExt As String
Dim sFileName As String

'    Me.dlgExport.Filter = "EMF (*.emf)|*.emf|CGM (*.cgm)|*.cgm|EPS (*.eps)|*.eps|AI (*.ai)|*.ai|PDF (*.pdf)|*.pdf|BMP (*.bmp)|*.bmp|TIFF (*.tif)|*.tif|JPEG (*.jpg)|*.jpg"
    
    Me.dlgExport.Filter = "PDF (*.pdf)|*.pdf|BMP (*.bmp)|*.bmp|TIFF (*.tif)|*.tif|JPEG (*.jpg)|*.jpg"
   
    If Me.cmbExportType.ListIndex <> -1 Then
        Me.dlgExport.FilterIndex = Me.cmbExportType.ListIndex + 1
    Else
        Me.dlgExport.FilterIndex = 4
    End If
    
    Me.dlgExport.DialogTitle = "Export"
    
    Me.Hide
    
    Me.dlgExport.ShowSave
    
    If Me.dlgExport.FileName = "" Then
        Me.Show
        Exit Sub
    Else
        sFileName = Me.dlgExport.FileName
    End If
    
     sFileExt = Right(sFileName, 3)
     
    Select Case sFileExt
        Case "emf"
            Me.cmbExportType.Text = "EMF (*.emf)"
        Case "cgm"
            Me.cmbExportType.Text = "CGM (*.cgm)"
        Case "eps"
            Me.cmbExportType.Text = "EPS (*.eps)"
        Case ".ai"
            Me.cmbExportType.Text = "AI (*.ai)"
        Case "pdf"
            Me.cmbExportType.Text = "PDF (*.pdf)"
        Case "bmp"
            Me.cmbExportType.Text = "BMP (*.bmp)"
        Case "tif"
            Me.cmbExportType.Text = "TIFF (*.tif)"
        Case "jpg"
            Me.cmbExportType.Text = "JPEG (*.jpg)"
    End Select
    
    Me.txtPath.Text = sFileName
    
   Me.Show
  
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdExport_Click()
On Error GoTo ErrHand:
  Dim sFileExt As String
  Dim pExporter As IExporter
  Dim pJpegExporter As IJpegExporter
  Dim sFileName As String
  Dim pActiveView As IActiveView
  Dim pMxDoc As IMxDocument
  Dim pMouse As IMouseCursor
  
  If Me.txtPath.Text = "" Then
    MsgBox "You have not typed in a valid path!!!"
    Exit Sub
  End If
  
  Dim bValid As Boolean
  bValid = CheckForValidPath(Me.txtPath.Text)
    
  If bValid = False Then
    MsgBox "You have not typed in a valid path!!!"
    Exit Sub
  End If

  '***Need to make sure it's a valid path
  
  Set pMouse = New MouseCursor
  pMouse.SetCursor 2

  Set pMxDoc = m_pApp.Document
  sFileName = Left(Me.txtPath.Text, Len(Me.txtPath.Text) - 4)
  sFileExt = Right(Me.txtPath.Text, 3)
    
  If m_pExporter Is Nothing Then
    Select Case sFileExt
    Case "emf"
      Set pExporter = New EmfExporter
    Case "cgm"
      Set pExporter = New CGMExporter
    Case "eps"
        'what exporter do I use here
    Case ".ai"
      Set pExporter = New AIExporter
    Case "pdf"
      Set pExporter = New PDFExporter
      'Map the basic fonts
      MapFonts pExporter
    Case "bmp"
      Set pExporter = New DibExporter
    Case "tif"
      Set pExporter = New TiffExporter
    Case "jpg"
      Set pExporter = New JpegExporter
    End Select
  Else
    Set pExporter = m_pExporter
  End If
        
  If pExporter Is Nothing Then
    MsgBox "No exporter object!!!"
    Exit Sub
  End If
   
  'Switch to the Layout view if we are not already there
  If Not TypeOf pMxDoc.ActiveView Is IPageLayout Then
    Set pMxDoc.ActiveView = pMxDoc.PageLayout
  End If

  Set pActiveView = pMxDoc.ActiveView
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
  Set PagesToExport = New Collection
  Set pSeriesOpts = m_pMapSeries
  Set pSeriesOpts2 = pSeriesOpts
  
  If Not m_pMapPage Is Nothing Then
      PagesToExport.Add m_pMapPage
  End If
  
  If Not m_pMapSeries Is Nothing And m_pMapPage Is Nothing And m_pMapBook Is Nothing Then
    If Me.optAll.Value = True Then
      For i = 0 To m_pMapSeries.PageCount - 1
        If Me.chkDisabled.Value = 1 Then
          If m_pMapSeries.Page(i).EnablePage Then
            PagesToExport.Add m_pMapSeries.Page(i)
          End If
         Else
            PagesToExport.Add m_pMapSeries.Page(i)
        End If
      Next i
    ElseIf Me.optPages.Value = True Then
      'parse out the pages to export
      If chkDisabled.Value = 1 Then
        Set PagesToExport = ParseOutPages(Me.txtPages.Text, m_pMapSeries, True)
      Else
        Set PagesToExport = ParseOutPages(Me.txtPages.Text, m_pMapSeries, False)
      End If
      If PagesToExport.count = 0 Then Exit Sub
    End If
  End If
  
  If PagesToExport.count > 0 Then
    If pSeriesOpts2.ClipData > 0 Then
      g_bClipFlag = True
    End If
    If pSeriesOpts.RotateFrame Then
      g_bRotateFlag = True
    End If
    If pSeriesOpts.LabelNeighbors Then
      g_bLabelNeighbors = True
    End If
    For i = 1 To PagesToExport.count
      Set pMapPage = PagesToExport.Item(i)
      pMapPage.DrawPage pMxDoc, m_pMapSeries, False
          
      If sFileExt = ".ai" Then
        sExportFile = sFileName & "_" & pMapPage.PageName & sFileExt
      Else
        sExportFile = sFileName & "_" & pMapPage.PageName & "." & sFileExt
      End If
      lblStatus.Caption = "Exporting to " & sExportFile & " ..."
      SetupToExport pExporter, dpi, ExportFrame, pActiveView, sExportFile
      
      'Do the export
      hdc = pExporter.StartExporting
        pActiveView.Output hdc, pExporter.Resolution, ExportFrame, Nothing, Nothing
        pMapPage.LastOutputted = Format(Date, "mm/dd/yyyy")
      pExporter.FinishExporting
    Next i
  End If
            
  If Not m_pMapBook Is Nothing Then
    Dim pMapSeries As IDSMapSeries
    Dim count As Long
    For i = 0 To m_pMapBook.ContentCount - 1
      Set PagesToExport = New Collection
      Set pMapSeries = m_pMapBook.ContentItem(i)
      Set pSeriesOpts = pMapSeries
    
      For count = 0 To pMapSeries.PageCount - 1
        If Me.chkDisabled.Value = 1 Then
          If pMapSeries.Page(count).EnablePage Then
            PagesToExport.Add pMapSeries.Page(count)
          End If
        Else
            PagesToExport.Add pMapSeries.Page(count)
        End If
      Next count
        
      If pSeriesOpts2.ClipData > 0 Then
        g_bClipFlag = True
      End If
      If pSeriesOpts.RotateFrame Then
        g_bRotateFlag = True
      End If
      If pSeriesOpts.LabelNeighbors Then
        g_bLabelNeighbors = True
      End If
      For count = 1 To PagesToExport.count
        'now do export
        Set pMapPage = PagesToExport.Item(count)
        pMapPage.DrawPage pMxDoc, pMapSeries, False
      
        If sFileExt = ".ai" Then
            sExportFile = sFileName & "_series_" & i & "_" & pMapPage.PageName & sFileExt
        Else
            sExportFile = sFileName & "_series_" & i & "_" & pMapPage.PageName & "." & sFileExt
        End If
        lblStatus.Caption = "Exporting to " & sExportFile & " ..."
        SetupToExport pExporter, pExporter.Resolution, ExportFrame, pActiveView, sExportFile
          
        'Do the export
        hdc = pExporter.StartExporting
          pActiveView.Output hdc, pExporter.Resolution, ExportFrame, Nothing, Nothing
          pMapPage.LastOutputted = Format(Date, "mm/dd/yyyy")
        pExporter.FinishExporting
      Next count
    Next i
  End If

'  pActiveView.ScreenDisplay.DisplayTransformation.ZoomResolution = True
  If TypeOf pExporter Is IOutputCleanup Then
    Dim pCleanup As IOutputCleanup
    Set pCleanup = pExporter
    pCleanup.Cleanup
  End If
  
  lblStatus.Caption = ""
  Set m_pMapBook = Nothing
  Set m_pMapPage = Nothing
  Set m_pMapSeries = Nothing
  Unload Me
  
  Exit Sub
ErrHand:
  lblStatus.Caption = ""
  MsgBox "cmdExport_Click - " & Err.Description
End Sub

Private Sub MapFonts(pExporter As IExporter)
On Error GoTo ErrHand:
  If Not TypeOf pExporter Is IFontMapEnvironment Then Exit Sub
  
  Dim pFontMapEnv As IFontMapEnvironment, pFontMapColl As IFontMapCollection
  Dim pFontMap As IFontMap2
  Set pFontMapEnv = pExporter
  Set pFontMapColl = pFontMapEnv.FontMapCollection
  Set pFontMap = New FontMap
  pFontMap.SetMapping "Arial", "Helvetica"
  pFontMapColl.Add pFontMap
  Set pFontMap = New FontMap
  pFontMap.SetMapping "Arial Bold", "Helvetica-Bold"
  pFontMapColl.Add pFontMap
  Set pFontMap = New FontMap
  pFontMap.SetMapping "Arial Bold Italic", "Helvetica-BoldOblique"
  pFontMapColl.Add pFontMap
  Set pFontMap = New FontMap
  pFontMap.SetMapping "Arial Italic", "Helvetica-Oblique"
  pFontMapColl.Add pFontMap
  Set pFontMap = New FontMap
  pFontMap.SetMapping "Courier New", "Courier"
  pFontMapColl.Add pFontMap
  Set pFontMap = New FontMap
  pFontMap.SetMapping "Courier New Bold", "Courier-Bold"
  pFontMapColl.Add pFontMap
  Set pFontMap = New FontMap
  pFontMap.SetMapping "Courier New Bold Italic", "Courier-BoldOblique"
  pFontMapColl.Add pFontMap
  Set pFontMap = New FontMap
  pFontMap.SetMapping "Courier New Italic", "Courier-Oblique"
  pFontMapColl.Add pFontMap
  Set pFontMap = New FontMap
  pFontMap.SetMapping "Symbol", "Symbol"
  pFontMapColl.Add pFontMap
  Set pFontMap = New FontMap
  pFontMap.SetMapping "Times New Roman", "Times-Roman"
  pFontMapColl.Add pFontMap
  Set pFontMap = New FontMap
  pFontMap.SetMapping "Times New Roman Bold", "Times-Bold"
  pFontMapColl.Add pFontMap
  Set pFontMap = New FontMap
  pFontMap.SetMapping "Times New Roman Bold Italic", "Times-BoldItalic"
  pFontMapColl.Add pFontMap
  Set pFontMap = New FontMap
  pFontMap.SetMapping "Times New Roman Italic", "Times-Italic"
  pFontMapColl.Add pFontMap
  
  Exit Sub
ErrHand:
  MsgBox "MapFonts - " & Err.Description
End Sub

Public Sub InitializeTheForm()
    
    Me.cmbExportType.Clear
'    Me.cmbExportType.AddItem "EMF (*.emf)"
'    Me.cmbExportType.AddItem "CGM (*.cgm)"
'    Me.cmbExportType.AddItem "EPS (*.eps)"
'    Me.cmbExportType.AddItem "AI (*.ai)"
    Me.cmbExportType.AddItem "PDF (*.pdf)"
    Me.cmbExportType.AddItem "BMP (*.bmp)"
    Me.cmbExportType.AddItem "TIFF (*.tif)"
    Me.cmbExportType.AddItem "JPEG (*.jpg)"
    
'    Me.cmbExportType.Text = "JPEG (*.jpg)"
    
    Me.cmbExportType.ListIndex = 3
    
End Sub

Private Sub ChangeFileExtension(sFileType As String)

Dim sExt As String
    sExt = Right(sFileType, 4)
    sExt = Left(sExt, 3)
    
Dim sFileName As String
Dim sFileNameExt As String

    sFileName = Me.txtPath.Text
    sFileNameExt = Right(sFileName, 3)
    
    If sExt <> sFileNameExt Then
        Dim aFileName() As String
        
        aFileName = Split(sFileName, ".")
        
        If sExt <> ".ai" Then
            Me.txtPath.Text = aFileName(0) & "." & sExt
        Else
            Me.txtPath.Text = aFileName(0) & sExt
        End If
    
    End If
    
End Sub

Private Sub cmdOptions_Click()
  On Error GoTo ErrorHandler

  Dim sFileExt As String
      sFileExt = Me.cmbExportType.Text
      
  Dim pExporterSet As ISet
  Dim sTitle As String
  Dim pMyPage As IComPropertyPage   'build the property page
  Dim pMyPage2 As IComPropertyPage
  
  'Set m_pExporter = Nothing
  
  Set pExporterSet = New esriCore.Set

  Select Case sFileExt
  Case "EMF (*.emf)"
    If m_pExporter Is Nothing Then
      Set m_pExporter = New EmfExporter
    Else
      If Not TypeOf m_pExporter Is IEmfExporter Then
        Set m_pExporter = New EmfExporter
      End If
    End If
    sTitle = "EMF Options"
    Set pMyPage = New EmfExporterPropertyPage
  Case "CGM (*.cgm)"
    If m_pExporter Is Nothing Then
      Set m_pExporter = New CGMExporter
    Else
      If Not TypeOf m_pExporter Is ICGMExporter Then
        Set m_pExporter = New CGMExporter
      End If
    End If
    sTitle = "CGM Options"
    Set pMyPage = New CGMExporterPropertyPage
  Case "EPS (*.eps)"
      'What do I do here?
  Case "AI (*.ai)"
    If m_pExporter Is Nothing Then
      Set m_pExporter = New AIExporter
    Else
      If Not TypeOf m_pExporter Is IAIExporter Then
        Set m_pExporter = New AIExporter
      End If
    End If
    sTitle = "AI Options"
    Set pMyPage = New AIExporterPropertyPage
  Case "PDF (*.pdf)"
    If m_pExporter Is Nothing Then
      Set m_pExporter = New PDFExporter
    Else
      If Not TypeOf m_pExporter Is IPDFExporter Then
        Set m_pExporter = New PDFExporter
      End If
    End If
    sTitle = "PDF Options"
    Set pMyPage = New PDFExporterPropertyPage
    Set pMyPage2 = New FontMappingPropertyPage
  Case "BMP (*.bmp)"
    If m_pExporter Is Nothing Then
      Set m_pExporter = New DibExporter
    Else
      If Not TypeOf m_pExporter Is IDibExporter Then
        Set m_pExporter = New DibExporter
      End If
    End If
    sTitle = "BMP Options"
    Set pMyPage = New DibExporterPropertyPage
  Case "TIFF (*.tif)"
    If m_pExporter Is Nothing Then
      Set m_pExporter = New TiffExporter
    Else
      If Not TypeOf m_pExporter Is ITiffExporter Then
        Set m_pExporter = New TiffExporter
      End If
    End If
    sTitle = "TIFF Options"
    Set pMyPage = New TiffExporterPropertyPage
  Case "JPEG (*.jpg)"
    If m_pExporter Is Nothing Then
      Set m_pExporter = New JpegExporter
    Else
      If Not TypeOf m_pExporter Is IJpegExporter Then
        Set m_pExporter = New JpegExporter
      End If
    End If
    sTitle = "JPEG Options"
    Set pMyPage = New JpegExporterPropertyPage
  End Select

  If m_pExporter Is Nothing Then Exit Sub

  pExporterSet.Add m_pExporter
    
  Dim pPS As IComPropertySheet
   
  Set pPS = New ComPropertySheet

  If Not pMyPage Is Nothing Then
    pPS.AddPage pMyPage
  End If
  
  If Not pMyPage2 Is Nothing Then
    pPS.AddPage pMyPage2
  End If
  
  Me.Hide
  
  If pPS.CanEdit(pExporterSet) = True Then
    pPS.Title = sTitle
    pPS.EditProperties pExporterSet, m_pApp.hwnd 'show the property sheet
  End If
              
  Me.Show
    
    
'  If pMyPage.IsPageDirty = True Then
    pMyPage.Apply
'  End If
    
  Exit Sub
ErrorHandler:
  MsgBox "cmdOptions_Click - " & Err.Description
End Sub

'Public Sub SetupToExport(ByRef pExporter As IExporter, ByRef dpi As Integer, ByRef ExportFrame As tagRECT, pActiveView As IActiveView, sExportFileName As String)
''  On Error GoTo ErrorHandler
''
''  Dim pEnv As IEnvelope
''
''    Set pEnv = New Envelope
''
''   pActiveView.ScreenDisplay.DisplayTransformation.Resolution = pExporter.Resolution
''
''  'Setup the exporter
''  ExportFrame = pActiveView.ExportFrame
''
''  pEnv.PutCoords ExportFrame.Left, ExportFrame.Top, ExportFrame.Right, ExportFrame.bottom
''  pEnv.Expand 2, 2, True
'''  dpi = pExporter.Resolution 'default screen resolution is usually 96
''
''  With pExporter
''    .PixelBounds = pEnv
''    .ExportFileName = sExportFileName
'''    .Resolution = dpi
''  End With
''
''
''  Exit Sub
''ErrorHandler:
''  MsgBox "SetupToExport - " & Err.Description
'On Error GoTo ErrorHandler
'
'  Dim pEnv As IEnvelope
'
'  Set pEnv = New Envelope
'
'  ExportFrame.Top = 0
'  ExportFrame.Left = 0
'  ExportFrame.Right = ConvertToPixels("Width", pExporter)
'  ExportFrame.bottom = ConvertToPixels("Height", pExporter)
'
'  pEnv.PutCoords ExportFrame.Left, ExportFrame.bottom, ExportFrame.Right, ExportFrame.Top
'
'  With pExporter
'    .PixelBounds = pEnv
'    .exportFileName = sExportFileName
'  End With
'
'  Exit Sub
'ErrorHandler:
'  MsgBox "SetupToExport - " & Err.Description
'End Sub

Public Sub SetupToExport(ByRef pExporter As IExporter, ByRef dpi As Integer, ByRef ExportFrame As tagRECT, pActiveView As IActiveView, sExportFileName As String)
  On Error GoTo ErrorHandler
  
  Dim pEnv As IEnvelope, pPageLayout As IPageLayout, pPage As IPage
  Dim dXmax As Double, dYmax As Double
  
   Set pEnv = New Envelope
'   pActiveView.ScreenDisplay.DisplayTransformation.Resolution = pExporter.Resolution
  'Setup the exporter
  ExportFrame = pActiveView.ExportFrame

  Set pPageLayout = pActiveView
  Set pPage = pPageLayout.Page
  
  If pPage.Units <> esriInches Then
    pPage.Units = esriInches
  End If
  
  pPage.QuerySize dXmax, dYmax
  pEnv.PutCoords 0, 0, dXmax * pExporter.Resolution, dYmax * pExporter.Resolution

'Commented out code removes a quarter of a unit, most likely an inch, from the extent to make it
'fit better on the page
'  ExportFrame.Top = pExporter.Resolution * 0.25
'  ExportFrame.Right = (dXmax - 0.25) * pExporter.Resolution
  ExportFrame.Right = dXmax * pExporter.Resolution
  ExportFrame.bottom = dYmax * pExporter.Resolution
  
  ExportFrame.Left = 0
  ExportFrame.Top = 0
            
  With pExporter
    .PixelBounds = pEnv
    .ExportFileName = sExportFileName
  End With

  
  Exit Sub
ErrorHandler:
  MsgBox "SetupToExport - " & Err.Description
End Sub


Public Function ConvertToPixels(sOrient As String, pExporter As IExporter) As Double
On Error GoTo ErrHand:
  Dim pixelExtent As Long
  Dim pDT As IDisplayTransformation
  Dim deviceRECT As tagRECT
  Dim pMxDoc As IMxDocument
  
  Set pMxDoc = m_pApp.Document
  Set pDT = pMxDoc.ActiveView.ScreenDisplay.DisplayTransformation
  deviceRECT = pDT.DeviceFrame
  
  If sOrient = "Height" Then
    pixelExtent = Abs(deviceRECT.Top - deviceRECT.bottom)
  ElseIf sOrient = "Width" Then
    pixelExtent = Abs(deviceRECT.Top - deviceRECT.bottom)
  End If
  
  ConvertToPixels = (pExporter.Resolution * (pixelExtent / pDT.Resolution))
  
  Exit Function
ErrHand:
  MsgBox "ConvertToPixels - " & Err.Description
End Function

Private Sub Form_Load()
  chkDisabled.Value = 1
End Sub

Private Function CheckForValidPath(sPathName As String) As Boolean
  On Error GoTo ErrorHandler

  CheckForValidPath = False
  
  Dim aPath() As String
      aPath = Split(sPathName, ".")

  If UBound(aPath) = 0 Then
    Exit Function
  ElseIf UBound(aPath) = 1 Then
    
    Dim sPath As String
    Dim lPos As Long
    
      lPos = InStrRev(sPathName, "\")
      sPath = Left$(sPathName, (Len(sPathName) - (Len(sPathName) - lPos + 1)))
      
      If Dir(sPath, vbDirectory) <> "" Then
        CheckForValidPath = True
        Exit Function
      Else
        Exit Function
      End If
      
  ElseIf UBound(aPath) > 1 Then
    Exit Function
  End If
  
  Exit Function
ErrorHandler:
  MsgBox "CheckForValidPath - " & Err.Description
End Function

Private Sub Form_Unload(Cancel As Integer)
Set m_pMapPage = Nothing
Set m_pMapSeries = Nothing
Set m_pMapBook = Nothing
Set m_pApp = Nothing
Set m_pExporter = Nothing
End Sub
