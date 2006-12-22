VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
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
Private m_pMapPage As IDSMapPage
Private m_pMapSeries As IDSMapSeries
Private m_pMapBook As IDSMapBook
Private m_pApp As IApplication

Private Sub chkPrintToFile_Click()
  If Me.chkPrintToFile.Value = 1 Then
    Me.txtCopies.Text = 1
    Me.fraCopies.Enabled = False
    Me.txtCopies.Enabled = False
    Me.UpDown1.Enabled = False
    Me.lblNumberofCopies.Enabled = False
  Else
    fraCopies.Enabled = True
    Me.txtCopies.Enabled = True
    Me.UpDown1.Enabled = True
    Me.lblNumberofCopies.Enabled = True
  End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

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

Private Sub cmdOk_Click()
On Error GoTo ErrorHandler

  Dim pAView As IActiveView
  Dim pPrinter As IPrinter
  Dim pMxApp As IMxApplication
  Dim pMxDoc As IMxDocument
  Dim pLayout As IPageLayout
  Dim iNumPages As Integer
  Dim pPage As IPage
  Dim pMouse As IMouseCursor
  
  Set pMouse = New MouseCursor
  pMouse.SetCursor 2

  Set pMxApp = m_pApp
  Set pPrinter = pMxApp.Printer
  Set pMxDoc = m_pApp.Document
  Set pLayout = pMxDoc.PageLayout
  Set pPage = pLayout.Page
  
  If Me.chkPrintToFile.Value = 1 Then
'    If UCase(pPrinter.FileExtension) = "PS" Then
      Me.dlgPrint.Filter = "Postscript Files (*.ps,*.eps)|*.ps,*.eps"
'    Else
'      Me.dlgPrint.Filter = UCase(pPrinter.FileExtension) & " (*." & LCase(pPrinter.FileExtension) & ")" & "|*." & LCase(pPrinter.FileExtension)
'    End If
    
    Me.dlgPrint.DialogTitle = "Print to File"
    Me.Hide
    Me.dlgPrint.ShowSave
    
    Dim sFileName As String, sPrefix As String, sExt As String, sSplit() As String
    
    sFileName = Me.dlgPrint.FileName
    If sFileName <> "" Then
      If InStr(1, sFileName, ".", vbTextCompare) > 0 Then
        sSplit = Split(sFileName, ".", , vbTextCompare)
        sPrefix = sSplit(0)
        sExt = sSplit(1)
      Else
        sPrefix = sFileName
        sExt = "ps"
        sFileName = sFileName & ".ps"
      End If
    Else
      MsgBox "Please specify a file name for the page(s)"
      Me.Show
      Exit Sub
    End If
  End If
  
  If Me.optTile.Value = True Then
      pPage.PageToPrinterMapping = esriPageMappingTile
  ElseIf Me.optScale = True Then
      pPage.PageToPrinterMapping = esriPageMappingScale
  ElseIf Me.optProceed.Value = True Then
      pPage.PageToPrinterMapping = esriPageMappingCrop
  End If
  
  pPrinter.Paper.Orientation = pLayout.Page.Orientation
  
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
  
  Set PagesToPrint = New Collection
  
  If Not m_pMapPage Is Nothing Then
      PagesToPrint.Add m_pMapPage
  End If
  
  If m_pMapPage Is Nothing And m_pMapBook Is Nothing Then
    If frmPrint.optPrintAll.Value = True Then
      For i = 0 To m_pMapSeries.PageCount - 1
        If chkDisabled.Value = 1 Then
          If m_pMapSeries.Page(i).EnablePage Then
            PagesToPrint.Add m_pMapSeries.Page(i)
          End If
        Else
          PagesToPrint.Add m_pMapSeries.Page(i)
        End If
      Next i
    ElseIf frmPrint.optPrintPages.Value = True Then
      'parse out the pages to print
      If chkDisabled.Value = 1 Then
        Set PagesToPrint = ParseOutPages(Me.txtPrintPages.Text, m_pMapSeries, True)
      Else
        Set PagesToPrint = ParseOutPages(Me.txtPrintPages.Text, m_pMapSeries, False)
      End If
      If PagesToPrint.count = 0 Then Exit Sub
    End If
  End If
      
  numPages = CLng(Me.txtCopies.Text)
  
  If PagesToPrint.count > 0 Then
    Set pSeriesOpts = m_pMapSeries
    Set pSeriesOpts2 = pSeriesOpts
    If pSeriesOpts2.ClipData > 0 Then
      g_bClipFlag = True
    End If
    If pSeriesOpts.RotateFrame Then
      g_bRotateFlag = True
    End If
    If pSeriesOpts.LabelNeighbors Then
      g_bLabelNeighbors = True
    End If
    For i = 1 To PagesToPrint.count
      Set pMapPage = PagesToPrint.Item(i)
      pMapPage.DrawPage pMxDoc, m_pMapSeries, False
      CheckNumberOfPages pPage, pPrinter, iNumPages
      lblPrintStatus.Caption = "Printing page " & pMapPage.PageName & " ..."
        
      For iCurrentPage = 1 To iNumPages
        SetupToPrint pPrinter, pPage, iCurrentPage, lDPI, rectDeviceBounds, pVisBounds, devFrameEnvelope
        If Me.chkPrintToFile.Value = 1 Then
          If pPage.PageToPrinterMapping = esriPageMappingTile Then
            pPrinter.PrintToFile = sPrefix & "_" & pMapPage.PageName & "_" & iCurrentPage & "." & sExt
          Else
            pPrinter.PrintToFile = sPrefix & "_" & pMapPage.PageName & "." & sExt
          End If
        End If
        For a = 1 To numPages
          hdc = pPrinter.StartPrinting(devFrameEnvelope, 0)
            pMxDoc.ActiveView.Output hdc, lDPI, rectDeviceBounds, pVisBounds, Nothing
            pMapPage.LastOutputted = Format(Date, "mm/dd/yyyy")
          pPrinter.FinishPrinting
        Next a
      Next iCurrentPage
    Next i
  End If
  
  If Not m_pMapBook Is Nothing Then
    Dim pSeriesCount As Long
    Dim MapSeriesColl As Collection
    Dim pMapSeries As IDSMapSeries
    Dim count As Long
    
    pSeriesCount = m_pMapBook.ContentCount
    
    Set MapSeriesColl = New Collection
    
    For i = 0 To pSeriesCount - 1
        MapSeriesColl.Add m_pMapBook.ContentItem(i)
    Next i

    If MapSeriesColl.count = 0 Then Exit Sub
    
    For i = 1 To MapSeriesColl.count
      Set PagesToPrint = New Collection
      Set pMapSeries = MapSeriesColl.Item(i)
      Set pSeriesOpts = pMapSeries
      Set pSeriesOpts2 = pSeriesOpts
      
      If pSeriesOpts2.ClipData > 0 Then
        g_bClipFlag = True
      End If
      If pSeriesOpts.RotateFrame Then
        g_bRotateFlag = True
      End If
      If pSeriesOpts.LabelNeighbors Then
        g_bLabelNeighbors = True
      End If
        
      For count = 0 To pMapSeries.PageCount - 1
        If chkDisabled.Value = 1 Then
          If pMapSeries.Page(count).EnablePage Then
            PagesToPrint.Add pMapSeries.Page(count)
          End If
        Else
          PagesToPrint.Add pMapSeries.Page(count)
        End If
      Next count
      
      For count = 1 To PagesToPrint.count
      'now do printing
        Set pMapPage = PagesToPrint.Item(count)
        pMapPage.DrawPage pMxDoc, pMapSeries, False
        
        CheckNumberOfPages pPage, pPrinter, iNumPages
        lblPrintStatus.Caption = "Printing page " & pMapPage.PageName & " ..."
            
        For iCurrentPage = 1 To iNumPages
          SetupToPrint pPrinter, pPage, iCurrentPage, lDPI, rectDeviceBounds, pVisBounds, devFrameEnvelope
          If Me.chkPrintToFile.Value = 1 Then
            If pPage.PageToPrinterMapping = esriPageMappingTile Then
              pPrinter.PrintToFile = sPrefix & "_" & pMapPage.PageName & "_" & iCurrentPage & "." & sExt
            Else
              pPrinter.PrintToFile = sPrefix & "_" & pMapPage.PageName & "." & sExt
            End If
          End If
          For a = 1 To numPages
            hdc = pPrinter.StartPrinting(devFrameEnvelope, 0)
              pMxDoc.ActiveView.Output hdc, lDPI, rectDeviceBounds, pVisBounds, Nothing
              pMapPage.LastOutputted = Format(Date, "mm/dd/yyyy")
            pPrinter.FinishPrinting
          Next a
        Next iCurrentPage
      
      Next count
            
    Next i
  End If
                                   
  lblPrintStatus.Caption = ""
  Set m_pMapBook = Nothing
  Set m_pMapPage = Nothing
  Set m_pMapSeries = Nothing
  Unload Me

  Exit Sub
ErrorHandler:
  lblPrintStatus.Caption = ""
  MsgBox "cmdOK_Click - " & Err.Description
End Sub

Public Property Get Application() As IApplication
    Set Application = m_pApp
End Property

Public Property Let Application(ByVal pApp As IApplication)
    Set m_pApp = pApp
End Property

Private Sub cmdSetup_Click()
  If (Not m_pApp.IsDialogVisible(esriMxDlgPageSetup)) Then
    Dim bDialog As Boolean
    Dim pPrinter As IPrinter
    Dim pMxApp As IMxApplication
    m_pApp.ShowDialog esriMxDlgPageSetup, True
    
    Me.Hide
    bDialog = True
    
    While bDialog = True
        bDialog = m_pApp.IsDialogVisible(esriMxDlgPageSetup)
        DoEvents
        
'            Sleep 1
    
    Wend
    
    Set pMxApp = m_pApp
    Set pPrinter = pMxApp.Printer
    frmPrint.lblName.Caption = pPrinter.Paper.PrinterName
    frmPrint.lblType.Caption = pPrinter.DriverName
    If TypeOf pPrinter Is IPsPrinter Then
      Me.chkPrintToFile.Enabled = True
    Else
      Me.chkPrintToFile.Value = 0
      Me.chkPrintToFile.Enabled = False
    End If
    Me.Show
  End If
End Sub

Private Sub Form_Load()
  chkDisabled.Value = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_pApp = Nothing
    Set m_pMapPage = Nothing
    Set m_pMapSeries = Nothing
    Set m_pMapBook = Nothing
End Sub

Private Sub optProceed_Click()
    If optProceed.Value = True Then
        Me.fraTileOptions.Enabled = False
    End If
End Sub

Private Sub optScale_Click()
    If optScale.Value = True Then
        Me.fraTileOptions.Enabled = False
    End If
End Sub

Private Sub optTile_Click()
    If optTile.Value = True Then
        Me.fraTileOptions.Enabled = True
        Me.optTileAll.Value = True
    Else
        Me.fraTileOptions.Enabled = False
    End If
End Sub

Public Sub SetupToPrint(pPrinter As IPrinter, pPage As IPage, iCurrentPage As Integer, ByRef lDPI As Long, ByRef rectDeviceBounds As tagRECT, _
ByRef pVisBounds As IEnvelope, ByRef devFrameEnvelope As IEnvelope)
On Error GoTo ErrorHandler
  Dim idpi As Integer
  Dim pDeviceBounds As IEnvelope
  Dim paperWidthInch As Double
  Dim paperHeightInch As Double

  idpi = pPrinter.Resolution  'dots per inch
          
  Set pDeviceBounds = New Envelope
              
  pPage.GetDeviceBounds pPrinter, iCurrentPage, 0, idpi, pDeviceBounds
               
  rectDeviceBounds.Left = pDeviceBounds.XMin
  rectDeviceBounds.Top = pDeviceBounds.YMin
  rectDeviceBounds.Right = pDeviceBounds.XMax
  rectDeviceBounds.bottom = pDeviceBounds.YMax
  
  'Following block added 6/19/03 to fix problem with plots being cutoff
  If TypeOf pPrinter Is IEmfPrinter Then
    ' For emf printers we have to remove the top and left unprintable area
    ' from device coordinates so its origin is 0,0.
    '
    rectDeviceBounds.Right = rectDeviceBounds.Right - rectDeviceBounds.Left
    rectDeviceBounds.bottom = rectDeviceBounds.bottom - rectDeviceBounds.Top
    rectDeviceBounds.Left = 0
    rectDeviceBounds.Top = 0
  End If
  
  Set pVisBounds = New Envelope
  pPage.GetPageBounds pPrinter, iCurrentPage, 0, pVisBounds
  pPrinter.QueryPaperSize paperWidthInch, paperHeightInch
  Set devFrameEnvelope = New Envelope
  devFrameEnvelope.PutCoords 0, 0, paperWidthInch * idpi, paperHeightInch * idpi
  
  lDPI = CLng(idpi)

  Exit Sub
ErrorHandler:
  MsgBox "SetupToPrint - " & Err.Description
End Sub

Public Sub CheckNumberOfPages(pPage As IPage, pPrinter As IPrinter, ByRef iNumPages As Integer)
On Error GoTo ErrorHandler
  pPage.PrinterPageCount pPrinter, 0, iNumPages
      
  If frmPrint.optTile.Value = True Then
    If frmPrint.optPages.Value = True Then
      Dim iPageNo As Integer
      Dim sPageNo As String
      sPageNo = frmPrint.txtTo.Text
      
      If sPageNo <> "" Then
          iPageNo = CInt(sPageNo)
      Else
          Exit Sub
      End If
      
      If iPageNo < iNumPages Then
          iNumPages = iPageNo
      End If
    End If
  End If
  
  Exit Sub
ErrorHandler:
  MsgBox "CheckNumberOfPages - " & Err.Description
End Sub
