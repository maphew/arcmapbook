VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmExport 
   Caption         =   "Export"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   5592
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   5592
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

Private m_pMapPage As INWDSMapPage
Private m_pMapSeries As INWDSMapSeries
Private m_pMapBook As INWDSMapBook
Private m_pApp As IApplication
Private m_pExport As IExport
Private m_pExportFrame As IModelessFrame

Public Property Get aNWDSMapPage() As INWDSMapPage
    Set aNWDSMapPage = m_pMapPage
End Property

Public Property Let aNWDSMapPage(ByVal pMapPage As INWDSMapPage)
    Set m_pMapPage = pMapPage
End Property

Public Property Let ExportFrame(ByVal pExportFrame As IModelessFrame)
    Set m_pExportFrame = pExportFrame
End Property

Public Property Get aNWDSMapSeries() As INWDSMapSeries
    Set aNWDSMapSeries = m_pMapSeries
End Property

Public Property Let aNWDSMapSeries(ByVal pMapSeries As INWDSMapSeries)
    Set m_pMapSeries = pMapSeries
End Property

Public Property Get aNWDSMapBook() As INWDSMapBook
    Set aNWDSMapBook = m_pMapBook
End Property

Public Property Let aNWDSMapBook(ByVal pMapBook As INWDSMapBook)
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

Set m_pExport = Nothing

If Me.txtPath.Text = "" Then Exit Sub

Dim sExt As String
    sExt = Me.cmbExportType.Text

    ChangeFileExtension sExt

End Sub

Private Sub cmdBrowse_Click()
Dim sFileExt As String
Dim sFileName As String

'    Me.dlgExport.Filter = "EMF (*.emf)|*.emf|CGM (*.cgm)|*.cgm|EPS (*.eps)|*.eps|AI (*.ai)|*.ai|PDF (*.pdf)|*.pdf|BMP (*.bmp)|*.bmp|TIFF (*.tif)|*.tif|JPEG (*.jpg)|*.jpg"
    
    Me.dlgExport.Filter = "BMP (*.bmp)|*.bmp|EPS (*.eps)|*.eps|JPEG (*.jpg)|*.jpg|PDF (*.pdf)|*.pdf|TIFF (*.tif)|*.tif"
   
    If Me.cmbExportType.ListIndex <> -1 Then
        Me.dlgExport.FilterIndex = Me.cmbExportType.ListIndex + 1
    Else
        Me.dlgExport.FilterIndex = 4
    End If
    
    Me.dlgExport.DialogTitle = "Export"
    
'    Me.Hide
    m_pExportFrame.Visible = False
    
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
'        Case "cgm"
'            Me.cmbExportType.Text = "CGM (*.cgm)"
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
    
'   Me.Show
  m_pExportFrame.Visible = True
  
End Sub

Private Sub cmdCancel_Click()
    m_pExportFrame.Visible = False
    Unload Me
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
    
  If m_pExport Is Nothing Then
    Select Case sFileExt
    Case "emf"
      Set pExport = New ExportEMF
'    Case "cgm"
'      MsgBox "CGMExporter not supported at 9.0, need to change this code to the replacement."
'      Exit Sub
'      Set pExport = New CGMExporter
    Case "eps"
      Set pExport = New ExportPS
    Case ".ai"
      Set pExport = New ExportAI
    Case "pdf"
      Set pExport = New ExportPDF
      'Map the basic fonts
      MapFonts pExport
    Case "bmp"
      Set pExport = New ExportBMP
    Case "tif"
      Set pExport = New ExportTIFF
    Case "jpg"
      Set pExport = New ExportJPEG
    End Select
  Else
    Set pExport = m_pExport
  End If
        
  If pExport Is Nothing Then
    MsgBox "No export object!!!"
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
  Dim pMapPage As INWDSMapPage, pSeriesOpts As INWDSMapSeriesOptions
  Dim ExportFrame As tagRECT, pSeriesOpts2 As INWDSMapSeriesOptions2
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
      SetupToExport pExport, dpi, ExportFrame, pActiveView, sExportFile
      
      'Do the export
      hdc = pExport.StartExporting
        pActiveView.Output hdc, pExport.Resolution, ExportFrame, Nothing, Nothing
        pMapPage.LastOutputted = Format(Date, "mm/dd/yyyy")
      pExport.FinishExporting
    Next i
  End If
            
  If Not m_pMapBook Is Nothing Then
    Dim pMapSeries As INWDSMapSeries
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
        SetupToExport pExport, pExport.Resolution, ExportFrame, pActiveView, sExportFile
          
        'Do the export
        hdc = pExport.StartExporting
          pActiveView.Output hdc, pExport.Resolution, ExportFrame, Nothing, Nothing
          pMapPage.LastOutputted = Format(Date, "mm/dd/yyyy")
        pExport.FinishExporting
      Next count
    Next i
  End If

'  pActiveView.ScreenDisplay.DisplayTransformation.ZoomResolution = True
  If TypeOf pExport Is IOutputCleanup Then
    Dim pCleanup As IOutputCleanup
    Set pCleanup = pExport
    pCleanup.Cleanup
  End If
  
  lblStatus.Caption = ""
  Set m_pMapBook = Nothing
  Set m_pMapPage = Nothing
  Set m_pMapSeries = Nothing
  m_pExportFrame.Visible = False
  Unload Me
  
  Exit Sub
ErrHand:
  lblStatus.Caption = ""
  MsgBox "cmdExport_Click - " & Err.Description
End Sub

Private Sub MapFonts(pExport As IExport)
On Error GoTo ErrHand:
  If Not TypeOf pExport Is IFontMapEnvironment Then Exit Sub
  
  Dim pFontMapEnv As IFontMapEnvironment, pFontMapColl As IFontMapCollection
  Dim pFontMap As IFontMap2
  Set pFontMapEnv = pExport
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
    Me.cmbExportType.AddItem "BMP (*.bmp)"
    Me.cmbExportType.AddItem "EPS (*.eps)"
    Me.cmbExportType.AddItem "JPEG (*.jpg)"
    Me.cmbExportType.AddItem "PDF (*.pdf)"
    Me.cmbExportType.AddItem "TIFF (*.tif)"
    
'    Me.cmbExportType.Text = "JPEG (*.jpg)"
    
    Me.cmbExportType.ListIndex = 2
    
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
  sFileExt = Me.cmbExportType.Text
      
  Dim sTitle As String

  Select Case sFileExt
  Case "EMF (*.emf)"
    If m_pExport Is Nothing Then
      Set m_pExport = New ExportEMF
    Else
      If Not TypeOf m_pExport Is IExportEMF Then
        Set m_pExport = New ExportEMF
      End If
    End If
    sTitle = "EMF Options"
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
    If m_pExport Is Nothing Then
      Set m_pExport = New ExportPS
    Else
      If Not TypeOf m_pExport Is IExportPS Then
        Set m_pExport = New ExportPS
      End If
    End If
    sTitle = "EPS Options"
  Case "PDF (*.pdf)"
    If m_pExport Is Nothing Then
      Set m_pExport = New ExportPDF
    Else
      If Not TypeOf m_pExport Is IExportPDF Then
        Set m_pExport = New ExportPDF
      End If
    End If
    sTitle = "PDF Options"
  Case "BMP (*.bmp)"
    If m_pExport Is Nothing Then
      Set m_pExport = New ExportBMP
    Else
      If Not TypeOf m_pExport Is IExportBMP Then
        Set m_pExport = New ExportBMP
      End If
    End If
    sTitle = "BMP Options"
  Case "TIFF (*.tif)"
    If m_pExport Is Nothing Then
      Set m_pExport = New ExportTIFF
    Else
      If Not TypeOf m_pExport Is IExportTIFF Then
        Set m_pExport = New ExportTIFF
      End If
    End If
    sTitle = "TIFF Options"
  Case "JPEG (*.jpg)"
    If m_pExport Is Nothing Then
      Set m_pExport = New ExportJPEG
    Else
      If Not TypeOf m_pExport Is IExportJPEG Then
        Set m_pExport = New ExportJPEG
      End If
    End If
    sTitle = "JPEG Options"
  End Select

  If m_pExport Is Nothing Then Exit Sub
  
'  Me.Hide
  m_pExportFrame.Visible = False
              
  Set frmExportPropDlg.Export = m_pExport
  frmExportPropDlg.Caption = sTitle
  frmExportPropDlg.Show vbModal, Me
  
  'The ExportSVG class has a Compression property that changes the value of the Filter property,
  ' and we must syncronize our file extension to account for the possible change.
  If TypeOf m_pExport Is IExportSVG Then
    cboSaveAsType.List(cboSaveAsType.ListIndex) = Split(m_pExport.Filter, "|")(0)
    m_sFileExtension = Split(Split(cboSaveAsType.Text, "(")(1), ")")(0)
    m_sFileExtension = Right(m_sFileExtension, Len(m_sFileExtension) - 1)
    txtFileName.Text = "Unititled" & m_sFileExtension
  End If
              
'  Me.Show
  m_pExportFrame.Visible = True
        
  Exit Sub
ErrorHandler:
  MsgBox "cmdOptions_Click - " & Err.Description
End Sub

Public Sub SetupToExport(ByRef pExport As IExport, ByRef dpi As Integer, ByRef ExportFrame As tagRECT, pActiveView As IActiveView, sExportFileName As String)
  On Error GoTo ErrorHandler
  
  Dim pEnv As IEnvelope, pPageLayout As IPageLayout, pPage As IPage
  Dim dXmax As Double, dYmax As Double
  
   Set pEnv = New envelope
'   pActiveView.ScreenDisplay.DisplayTransformation.Resolution = pExport.Resolution
  'Setup the Export
  ExportFrame = pActiveView.ExportFrame

  Set pPageLayout = pActiveView
  Set pPage = pPageLayout.Page
  
  If pPage.Units <> esriInches Then
    pPage.Units = esriInches
  End If
  
  pPage.QuerySize dXmax, dYmax
  pEnv.PutCoords 0, 0, dXmax * pExport.Resolution, dYmax * pExport.Resolution

'Commented out code removes a quarter of a unit, most likely an inch, from the extent to make it
'fit better on the page
'  ExportFrame.Top = pExport.Resolution * 0.25
'  ExportFrame.Right = (dXmax - 0.25) * pExport.Resolution
  ExportFrame.Right = dXmax * pExport.Resolution
  ExportFrame.bottom = dYmax * pExport.Resolution
  
  ExportFrame.Left = 0
  ExportFrame.Top = 0
            
  With pExport
    .PixelBounds = pEnv
    .ExportFileName = sExportFileName
  End With

  
  Exit Sub
ErrorHandler:
  MsgBox "SetupToExport - " & Err.Description
End Sub


Public Function ConvertToPixels(sOrient As String, pExport As IExport) As Double
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
  
  ConvertToPixels = (pExport.Resolution * (pixelExtent / pDT.Resolution))
  
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
  Set m_pExport = Nothing
  Set m_pExportFrame = Nothing
End Sub
