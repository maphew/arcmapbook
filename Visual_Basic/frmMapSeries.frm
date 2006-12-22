VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMapSeries 
   BackColor       =   &H80000004&
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4275
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   4275
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBook 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   150
      ScaleHeight     =   4335
      ScaleWidth      =   3885
      TabIndex        =   1
      Top             =   630
      Width           =   3885
      Begin MSComctlLib.TreeView tvwMapBook 
         Height          =   4725
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   8334
         _Version        =   393217
         Indentation     =   44
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   3
         FullRowSelect   =   -1  'True
         SingleSel       =   -1  'True
         ImageList       =   "ImageList1"
         Appearance      =   1
      End
   End
   Begin VB.ListBox lstSorter 
      Height          =   1230
      Left            =   2790
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   3420
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   33
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapSeries.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapSeries.frx":062E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapSeries.frx":0C5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapSeries.frx":11EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapSeries.frx":1780
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapSeries.frx":1BF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapSeries.frx":2064
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapSeries.frx":2656
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuHeadingBook 
      Caption         =   "Book"
      Begin VB.Menu mnuBook 
         Caption         =   "Add Map Series..."
         Index           =   0
      End
      Begin VB.Menu mnuBook 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuBook 
         Caption         =   "Print Map Book..."
         Index           =   2
      End
      Begin VB.Menu mnuBook 
         Caption         =   "Export Map Book..."
         Index           =   3
      End
   End
   Begin VB.Menu mnuHeadingSeries 
      Caption         =   "Series"
      Begin VB.Menu mnuSeries 
         Caption         =   "Select/Enable Pages..."
         Index           =   0
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "Tag as Date"
         Index           =   2
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "Tag as Title"
         Index           =   3
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "Tag with Index Layer Field..."
         Index           =   4
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "Clear Tag for Selected"
         Index           =   5
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "Delete Series"
         Index           =   7
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "Delete Disabled Pages"
         Index           =   8
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "Disable Series"
         Index           =   10
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "Print Series..."
         Index           =   12
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "Export Series..."
         Index           =   13
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "Create Series Index..."
         Index           =   14
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "Series Properties..."
         Index           =   16
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "Page Properties..."
         Index           =   17
      End
   End
   Begin VB.Menu mnuHeadingPage 
      Caption         =   "Page"
      Begin VB.Menu mnuPage 
         Caption         =   "View Page"
         Index           =   0
      End
      Begin VB.Menu mnuPage 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuPage 
         Caption         =   "Delete Page"
         Index           =   2
      End
      Begin VB.Menu mnuPage 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuPage 
         Caption         =   "Disable Page"
         Index           =   4
      End
      Begin VB.Menu mnuPage 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuPage 
         Caption         =   "Print Page..."
         Index           =   6
      End
      Begin VB.Menu mnuPage 
         Caption         =   "Export Page..."
         Index           =   7
      End
   End
End
Attribute VB_Name = "frmMapSeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_pApp As IApplication
Private m_lXClick As Single
Private m_lYClick As Single
Private m_lButton As Single
Private m_pCurrentNode As Node
Private m_bNodeFlag As Boolean
Private m_bClickFlag As Boolean
Private m_bLabelingChanged As Boolean

Private Sub Form_Load()
  tvwMapBook.Nodes.Clear
'  tvwMapBook.Nodes.Add , , "MapBook", "Map Book (0 pages)", 1
  tvwMapBook.Nodes.Add , , "MapBook", "Map Book", 1
  m_bNodeFlag = True
  m_bClickFlag = False
  m_bLabelingChanged = False
End Sub

Private Sub mnuBook_Click(Index As Integer)
On Error GoTo ErrHand:
  Dim pMapBook As IDSMapBook
  'Check to see if a MapSeries already exists
  Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
  Select Case Index
  Case 0  'Add Map Series
    If pMapBook.ContentCount > 0 Then
      MsgBox "You must remove the existing Map Series before adding another."
      Exit Sub
    End If
  
    'Call the wizard for setting parameters and creating the series
    Set frmMapSeriesWiz.m_pApp = m_pApp
    frmMapSeriesWiz.Show vbModal
  Case 1  'Separator
  Case 2  'Print Map Book
    ShowPrinterDialog m_pApp, , pMapBook
'    pMapBook.PrintBook
  Case 3  'Export Map Book
    ShowExporterDialog m_pApp, , pMapBook
'    pMapBook.ExportBook
  End Select
  
  Exit Sub
ErrHand:
  MsgBox "mnuBook_Click - " & Err.Description
End Sub

Private Sub mnuPage_Click(Index As Integer)
On Error GoTo ErrHand:
  Dim pMapBook As IDSMapBook, pMapSeries As IDSMapSeries
  Dim lPage As Long, sText As String, lPos As Long, pMapPage As IDSMapPage
  Dim pSeriesOpts As IDSMapSeriesOptions, pSeriesOpts2 As IDSMapSeriesOptions2
  'Check to see if a MapSeries already exists
  Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
  Set pMapSeries = pMapBook.ContentItem(0)
  Set pSeriesOpts = pMapSeries
  Set pSeriesOpts2 = pSeriesOpts
  lPage = m_pCurrentNode.Tag
  Select Case Index
  Case 0  'View Page
    Set pMapPage = pMapSeries.Page(lPage)
    pMapPage.DrawPage m_pApp.Document, pMapSeries, True
    If pSeriesOpts2.ClipData > 0 Then
      g_bClipFlag = True
    End If
    If pSeriesOpts.RotateFrame Then
      g_bRotateFlag = True
    End If
    If pSeriesOpts.LabelNeighbors Then
      g_bLabelNeighbors = True
    End If
  Case 1  'Separator
  Case 2  'Delete Page
    'Remove the page, then update the tags on all subsequent pages
    pMapSeries.RemovePage lPage
    tvwMapBook.Nodes.Remove lPage + 3
    RenumberPages pMapSeries
  Case 3  'Separator
  Case 4  'Disable Page
    'Get the index number from the tag of the node
    pMapSeries.Page(lPage).EnablePage = Not pMapSeries.Page(lPage).EnablePage
    If pMapSeries.Page(lPage).EnablePage Then
      m_pCurrentNode.Image = 5
    Else
      m_pCurrentNode.Image = 6
    End If
  Case 5  'Separator
  Case 6  'Print Page
    ShowPrinterDialog m_pApp, pMapSeries, pMapSeries.Page(lPage)
  Case 7  'Export Page
    ShowExporterDialog m_pApp, pMapSeries, pMapSeries.Page(lPage)
  End Select
  
  Exit Sub
ErrHand:
  MsgBox "mnuPage_Click - " & Err.Description
End Sub

Private Sub mnuSeries_Click(Index As Integer)
On Error GoTo ErrHand:
  Dim pMapBook As IDSMapBook, pMapSeries As IDSMapSeries, pSeriesProps As IDSMapSeriesProps
  Dim lLoop As Long, pDoc As IMxDocument, pActive As IActiveView, bFlag As Boolean
  Dim pGraphicsCont As IGraphicsContainer, pElemProps As IElementProperties
  Dim pEnv As IEnvelope, pElem As IElement, pTextElement As ITextElement, pEnv2 As IEnvelope
  Dim pGraphicsContSel As IGraphicsContainerSelect, pMap As IMap
  Dim pIndexLayer As IFeatureLayer, lIndex As Long, sName As String, sTemp As String
  'Check to see if a MapSeries already exists
  Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
  Set pMapSeries = pMapBook.ContentItem(0)
  Set pSeriesProps = pMapSeries
  Set pDoc = m_pApp.Document
  Select Case Index
  Case 0  'Select Pages
    Set frmSelectPages.m_pApp = m_pApp
    frmSelectPages.Show vbModal
  Case 1  'Separator
  Case 2  'Tag as Date
    bFlag = TagItem(pDoc, "DSMAPBOOK - DATE", "")
    If Not bFlag Then
      MsgBox "You must have one Text Element selected in the Page Layout for tagging!!!"
    End If
  Case 3  'Tag as Title
    bFlag = TagItem(pDoc, "DSMAPBOOK - TITLE", "")
    If Not bFlag Then
      MsgBox "You must have one Text Element selected in the Page Layout for tagging!!!"
    End If
  Case 4  'Tag with Index Layer Field...
    'Find the data frame
    Set pMap = FindDataFrame(pDoc, pSeriesProps.DataFrameName)
    If pMap Is Nothing Then
      MsgBox "Could not find map in mnuSeries_Click routine!!!"
      Exit Sub
    End If
    'Find the Index layer
    Set pIndexLayer = FindLayer(pSeriesProps.IndexLayerName, pMap)
    If pIndexLayer Is Nothing Then
      MsgBox "Could not find index layer in mnuSeries_Click routine!!!"
      Exit Sub
    End If
  
    frmTagIndexField.InitializeList pIndexLayer.FeatureClass.Fields
    frmTagIndexField.Show vbModal
    
    'Exit sub if Cancel was selected
    If frmTagIndexField.m_bCancel Then
      Unload frmTagIndexField
      Exit Sub
    End If
    
    lIndex = frmTagIndexField.lstFields.ListIndex
    If lIndex >= 0 Then
      sTemp = frmTagIndexField.lstFields.List(lIndex)
    Else
      MsgBox "You did not pick a field to tag with!!!"
      Unload frmTagIndexField
      Exit Sub
    End If
    Unload frmTagIndexField
    
    lIndex = InStr(1, sTemp, " - ")
    sName = Mid(sTemp, 1, lIndex - 1)
    bFlag = TagItem(pDoc, "DSMAPBOOK - EXTRAITEM", sName)
    If Not bFlag Then
      MsgBox "You must have one Text Element selected in the Page Layout for tagging!!!"
    End If
  Case 5  'Clear Tag for selected
    Set pGraphicsCont = pDoc.PageLayout
    Set pGraphicsContSel = pDoc.PageLayout
    For lLoop = 0 To pGraphicsContSel.ElementSelectionCount - 1
      Set pElemProps = pGraphicsContSel.SelectedElement(lLoop)
      If TypeOf pElemProps Is ITextElement Then
        pElemProps.Name = ""
        pElemProps.Type = ""
        pGraphicsCont.UpdateElement pTextElement
      End If
    Next lLoop
  Case 6  'Separator
  Case 7  'Delete Series
    Set pActive = pDoc.FocusMap
    TurnOffClipping pMapSeries, m_pApp
    Set pMapSeries = Nothing
    pMapBook.RemoveContent 0
    tvwMapBook.Nodes.Clear
    tvwMapBook.Nodes.Add , , "MapBook", "Map Book", 1
    RemoveIndicators m_pApp
    pActive.Refresh
  Case 8  'Delete Disabled pages
    'Loop in reverse order so we remove pages as we work up.  Doing it this way makes
    'sure numbering isn't messed up when a page/node is removed.
    For lLoop = pMapSeries.PageCount - 1 To 0 Step -1
      If Not pMapSeries.Page(lLoop).EnablePage Then
        pMapSeries.RemovePage lLoop
        tvwMapBook.Nodes.Remove lLoop + 3
      End If
    Next lLoop
    RenumberPages pMapSeries
  Case 9  'Separator
  Case 10  'Disable Series
    'Get the index number from the tag of the node
    pMapSeries.EnableSeries = Not pMapSeries.EnableSeries
    If pMapSeries.EnableSeries Then
      m_pCurrentNode.Image = 3
    Else
      m_pCurrentNode.Image = 4
    End If
  Case 11  'Separator
  Case 12  'Print Series
    ShowPrinterDialog m_pApp, pMapSeries, Nothing
'    pMapSeries.PrintSeries
  Case 13  'Export Series
    ShowExporterDialog m_pApp, pMapSeries, Nothing
'    pMapSeries.ExportSeries
  Case 14
    Set frmCreateIndex.m_pApp = m_pApp
    frmCreateIndex.Show vbModal
  Case 15  'Separator
  Case 16  'Series Properties...
    Set frmSeriesProperties.m_pApp = m_pApp
    frmSeriesProperties.Show vbModal
  Case 17  'Page Properties...
    Set frmPageProperties.m_pApp = m_pApp
    frmPageProperties.Show vbModal
  End Select
  
  Exit Sub
ErrHand:
  MsgBox "mnuSeries_Click - " & Erl & " - " & Err.Description
End Sub

Private Function TagItem(pDoc As IMxDocument, sName As String, sType As String) As Boolean
On Error GoTo ErrHand:
  Dim bFlag As Boolean, pGraphicsCont As IGraphicsContainer, pActive As IActiveView
  Dim pElemProps As IElementProperties, pElem As IElement, pTextElement As ITextElement
  Dim pEnv2 As IEnvelope, pGraphicsContSel As IGraphicsContainerSelect, pEnv As IEnvelope
  
  Set pGraphicsCont = pDoc.PageLayout
  Set pGraphicsContSel = pDoc.PageLayout
  bFlag = False
  If pGraphicsContSel.ElementSelectionCount = 1 Then
    Set pElemProps = pGraphicsContSel.SelectedElement(0)
    If TypeOf pElemProps Is ITextElement Then
      Set pActive = pDoc.PageLayout
      pElemProps.Name = sName
      Set pElem = pElemProps
      Set pEnv = New Envelope
      pElem.QueryBounds pActive.ScreenDisplay, pEnv
      Set pTextElement = pElemProps
      Select Case sName
      Case "DSMAPBOOK - DATE"
        pTextElement.Text = Format(Date, "mmm dd, yyyy")
      Case "DSMAPBOOK - TITLE"
        pTextElement.Text = "Title String"
      Case "DSMAPBOOK - EXTRAITEM"
        pTextElement.Text = sType
        pElemProps.Type = sType
      End Select
      pGraphicsCont.UpdateElement pTextElement
      Set pEnv2 = New Envelope
      pElem.QueryBounds pActive.ScreenDisplay, pEnv2
      pEnv.Union pEnv2
      pActive.PartialRefresh esriViewGraphics, Nothing, pEnv
      bFlag = True
    End If
  End If
  
  TagItem = bFlag

  Exit Function
ErrHand:
  MsgBox "TagItem - " & Erl & " - " & Err.Description
  TagItem = bFlag
End Function

Private Sub RenumberPages(pMapSeries As IDSMapSeries)
On Error GoTo ErrHand:
'Routine for renumber the pages after one is removed
  Dim lLoop As Long, pNode As Node, sName As String
  For lLoop = 0 To pMapSeries.PageCount - 1
    Set pNode = tvwMapBook.Nodes.Item(lLoop + 3)
    sName = Mid(pNode.Key, 2)
    pNode.Tag = lLoop
    pNode.Key = "a" & sName
    pNode.Text = lLoop + 1 & " - " & sName
  Next lLoop
  tvwMapBook.Refresh
  
  Exit Sub
ErrHand:
  MsgBox "RenumberPages - " & Err.Description
End Sub

Private Sub picBook_Resize()
  tvwMapBook.Width = picBook.Width
  tvwMapBook.Height = picBook.Height
End Sub

Private Sub tvwMapBook_DblClick()
On Error GoTo ErrHand:
  Dim lPos As String, sText As String, pMapPage As IDSMapPage, lPage As Long
  Dim pSeriesOpts As IDSMapSeriesOptions, pSeriesOpts2 As IDSMapSeriesOptions2
  Dim pMapBook As IDSMapBook, pMapSeries As IDSMapSeries
  Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
  Set pMapSeries = pMapBook.ContentItem(0)
  Set pSeriesOpts = pMapSeries
  Set pSeriesOpts2 = pSeriesOpts
  
  'There is no NodeDoubleClick event, so we have to use the DblClick event on the control
  'and check to make sure we are over a node.  In the event of a doubleclick on a node,
  'the order of events being fired are NodeClick, MouseUp, Click, DblClick, MouseUp.  To
  'make sure the doubleclick occurred over a node, we can set a flag in the NodeClick event
  'and then disable it in the MouseUp event after the Click event.
  If Not m_bNodeFlag Then Exit Sub
  
  Select Case m_pCurrentNode.Image
  Case 5, 6   'Enable and not Enabled options for a map page
    If m_lXClick > 1320 Then
      If m_lButton = 1 Then
        lPage = m_pCurrentNode.Tag
        Set pMapPage = pMapSeries.Page(lPage)
        pMapPage.DrawPage m_pApp.Document, pMapSeries, True
        If pSeriesOpts2.ClipData > 0 Then
          g_bClipFlag = True
        End If
        If pSeriesOpts.RotateFrame Then
          g_bRotateFlag = True
        End If
        If pSeriesOpts.LabelNeighbors Then
          g_bLabelNeighbors = True
        End If
      End If
    End If
  End Select

  Exit Sub
ErrHand:
  MsgBox "twvMapBook_NodeClick - " & Err.Description
End Sub

Private Sub tvwMapBook_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrHand:
  m_lXClick = X
  m_lYClick = Y
  
  m_lButton = Button
  
  Exit Sub
ErrHand:
  MsgBox "tvwMapBook_MouseDown - " & Err.Description
End Sub

Private Sub tvwMapBook_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrHand:
  If Not m_bClickFlag Then Exit Sub
  m_bClickFlag = False
  m_bNodeFlag = False
  
  Exit Sub
ErrHand:
  MsgBox "tvwMapBook_MouseUp - " & Err.Description
End Sub

Private Sub tvwMapBook_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo ErrHand:
  Dim lLoop As Long, pUID As New UID, lImage As Long
  Dim pItem As ICommandItem, lPos As Long, sText As String
  Dim pMapBook As IDSMapBook, pMapSeries As IDSMapSeries
  Dim lPage As Long
  'Check to see if a MapSeries already exists
  Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
  Set pMapSeries = pMapBook.ContentItem(0)
  
  Set m_pCurrentNode = Node
  Select Case Node.Image
  Case 1, 2   'Enable and not Enabled options for a map book
    If m_lXClick < 180 Then
      If Node.Image = 1 Then
        Node.Image = 2
        pMapBook.EnableBook = False
'        tvwMapBook.Nodes.Item("MapSeries").Image = 4
'        UpdatePages False
      Else
        Node.Image = 1
        pMapBook.EnableBook = True
'        tvwMapBook.Nodes.Item("MapSeries").Image = 3
'        UpdatePages True
      End If
    Else
      If m_lButton = 2 Then
        PopupMenu mnuHeadingBook
      End If
    End If
  Case 3, 4   'Enable and not Enabled options for a map series
    If m_lXClick > 510 And m_lXClick < 760 Then
      If Node.Image = 3 Then
        Node.Image = 4
        pMapSeries.EnableSeries = False
'        UpdatePages False
      Else
        Node.Image = 3
        pMapSeries.EnableSeries = True
'        UpdatePages True
      End If
    Else
      If m_lButton = 2 Then
        If Node.Image = 3 Then
          mnuSeries(8).Caption = "Disable Series"
        Else
          mnuSeries(8).Caption = "Enable Series"
        End If
        PopupMenu mnuHeadingSeries
      End If
    End If
  Case 5, 6   'Enable and not Enabled options for a map page
    If m_lXClick > 1320 Then
      If m_lButton = 2 Then
        If Node.Image = 5 Then
          mnuPage(4).Caption = "Disable Page"
        Else
          mnuPage(4).Caption = "Enable Page"
        End If
        PopupMenu mnuHeadingPage
      End If
    ElseIf m_lXClick > 1080 And m_lXClick <= 1320 Then
      lPage = Node.Tag
      If Node.Image = 5 Then
        Node.Image = 6
        pMapSeries.Page(lPage).EnablePage = False
      Else
        Node.Image = 5
        pMapSeries.Page(lPage).EnablePage = True
      End If
    End If
  End Select

  Exit Sub
ErrHand:
  MsgBox "twvMapBook_NodeClick - " & Err.Description
End Sub

Private Sub UpdatePages(bEnableFlag As Boolean)
On Error GoTo ErrHand:
  Dim lLoop As Long, pNode As Node
  For lLoop = 2 To tvwMapBook.Nodes.count
    Set pNode = tvwMapBook.Nodes.Item(lLoop)
    If pNode.Image = 5 Or pNode.Image = 6 Then
      If bEnableFlag = True Then
        pNode.Image = 5
      Else
        pNode.Image = 6
      End If
    End If
  Next lLoop

  Exit Sub
ErrHand:
  MsgBox "UpdatePages - " & Err.Description
End Sub

Public Sub ShowPrinterDialog(pMxApp As IMxApplication, Optional pMapSeries As IDSMapSeries, Optional pPrintMaterial As IUnknown)
  On Error GoTo ErrorHandler

  Dim pPrinter As IPrinter
  Dim pApp As IApplication
  Dim iNumPages As Integer
  Dim pPage As IPage
  Dim pDoc As IMxDocument
  Dim pLayout As IPageLayout
  
  Set pPrinter = pMxApp.Printer
  If pPrinter Is Nothing Then
    MsgBox "You must have at least one printer defined before using this command!!!"
    Exit Sub
  End If
  
  Set pApp = pMxApp
  
  Set pDoc = pApp.Document
  Set pLayout = pDoc.PageLayout
  Set pPage = pLayout.Page
          
  pPage.PrinterPageCount pPrinter, 0, iNumPages
  
  frmPrint.txtTo.Text = iNumPages
  
  frmPrint.Application = pApp
      
  frmPrint.lblName.Caption = pPrinter.Paper.PrinterName
  frmPrint.lblType.Caption = pPrinter.DriverName
  If TypeOf pPrinter Is IPsPrinter Then
    frmPrint.chkPrintToFile.Enabled = True
  Else
    frmPrint.chkPrintToFile.Value = 0
    frmPrint.chkPrintToFile.Enabled = False
  End If
  'If pprintmaterial is nothing then it means you are printing a map series
  
  If pPrintMaterial Is Nothing Then
      frmPrint.aDSMapSeries = pMapSeries
      frmPrint.optPrintCurrentPage.Enabled = False
      frmPrint.Show
      Exit Sub
  End If
  
  If TypeOf pPrintMaterial Is IDSMapBook Then
      frmPrint.aDSMapBook = pPrintMaterial
      frmPrint.optPrintCurrentPage.Enabled = False
      frmPrint.optPrintPages.Enabled = False
      frmPrint.txtPrintPages.Enabled = False
      frmPrint.Show
  ElseIf TypeOf pPrintMaterial Is IDSMapPage Then
      frmPrint.aDSMapPage = pPrintMaterial
      frmPrint.aDSMapSeries = pMapSeries
      frmPrint.optPrintCurrentPage.Value = True
      frmPrint.optPrintAll.Enabled = False
      frmPrint.optPrintPages.Enabled = False
      frmPrint.txtPrintPages.Enabled = False
      frmPrint.Show
  End If
  Set pPrintMaterial = Nothing
    
  Exit Sub
ErrorHandler:
  MsgBox "ShowPrinterDialog - " & Err.Description
End Sub
Public Sub ShowExporterDialog(pApp As IApplication, Optional pMapSeries As IDSMapSeries, Optional pExportMaterial As IUnknown)
  On Error GoTo ErrorHandler
    
  frmExport.Application = pApp
      
  'If pExportMaterial is nothing then it means you are printing a map series
  
  If pExportMaterial Is Nothing Then
      frmExport.aDSMapSeries = pMapSeries
      frmExport.optCurrentPage.Enabled = False
      frmExport.InitializeTheForm
      frmExport.Show
      Exit Sub
  End If
  
  If TypeOf pExportMaterial Is IDSMapBook Then
      frmExport.aDSMapBook = pExportMaterial
      frmExport.optCurrentPage.Enabled = False
      frmExport.optPages.Enabled = False
      frmExport.txtPages.Enabled = False
      frmExport.InitializeTheForm
      frmExport.Show
  ElseIf TypeOf pExportMaterial Is IDSMapPage Then
      frmExport.aDSMapPage = pExportMaterial
      frmExport.aDSMapSeries = pMapSeries
      frmExport.optCurrentPage.Value = True
      frmExport.optAll.Enabled = False
      frmExport.optPages.Enabled = False
      frmExport.txtPages.Enabled = False
      frmExport.InitializeTheForm
      frmExport.Show
  End If
  Set pExportMaterial = Nothing

  Exit Sub
ErrorHandler:
  MsgBox "ShowExporterDialog - " & Err.Description
End Sub



