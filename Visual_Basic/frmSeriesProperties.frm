VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmSeriesProperties 
   Caption         =   "Map Series Properties"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   7170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   5940
      TabIndex        =   40
      Top             =   4050
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   345
      Left            =   4800
      TabIndex        =   39
      Top             =   4050
      Width           =   1125
   End
   Begin TabDlg.SSTab tabProperties 
      Height          =   3885
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   6853
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Index Settings"
      TabPicture(0)   =   "frmSeriesProperties.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraPage1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tile Settings"
      TabPicture(1)   =   "frmSeriesProperties.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraPage2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Options"
      TabPicture(2)   =   "frmSeriesProperties.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "fraPage3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame fraPage1 
         BorderStyle     =   0  'None
         Height          =   3375
         Left            =   -74940
         TabIndex        =   30
         Top             =   420
         Width           =   6855
         Begin VB.ComboBox cmbDetailFrame 
            Enabled         =   0   'False
            Height          =   315
            Left            =   270
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   1140
            Width           =   2625
         End
         Begin VB.Frame fraIndexLayer 
            Caption         =   "Index Layer"
            Height          =   2415
            Left            =   3210
            TabIndex        =   31
            Top             =   900
            Width           =   3525
            Begin VB.ComboBox cmbIndexLayer 
               Enabled         =   0   'False
               Height          =   315
               Left            =   240
               Style           =   2  'Dropdown List
               TabIndex        =   33
               Top             =   510
               Width           =   3105
            End
            Begin VB.ComboBox cmbIndexField 
               Enabled         =   0   'False
               Height          =   315
               Left            =   240
               Style           =   2  'Dropdown List
               TabIndex        =   32
               Top             =   1200
               Width           =   3105
            End
            Begin VB.Label lblMapSheet 
               Caption         =   "Choose the index layer:"
               Height          =   225
               Index           =   3
               Left            =   240
               TabIndex        =   35
               Top             =   270
               Width           =   1725
            End
            Begin VB.Label lblMapSheet 
               Caption         =   "This field specifies the page name"
               Height          =   225
               Index           =   4
               Left            =   240
               TabIndex        =   34
               Top             =   960
               Width           =   2535
            End
         End
         Begin VB.Label Label1 
            Caption         =   $"frmSeriesProperties.frx":0054
            Height          =   615
            Index           =   0
            Left            =   30
            TabIndex        =   38
            Top             =   60
            Width           =   6705
         End
         Begin VB.Label lblMapSheet 
            Caption         =   "Choose the detail data frame:"
            Height          =   225
            Index           =   0
            Left            =   30
            TabIndex        =   37
            Top             =   870
            Width           =   2235
         End
      End
      Begin VB.Frame fraPage2 
         BorderStyle     =   0  'None
         Height          =   3375
         Left            =   -74910
         TabIndex        =   19
         Top             =   420
         Width           =   6765
         Begin VB.Frame fraChooseTiles 
            Caption         =   "Choose tiles"
            Height          =   2565
            Left            =   60
            TabIndex        =   24
            Top             =   750
            Width           =   3255
            Begin VB.OptionButton optTiles 
               Caption         =   "Use all of the tiles"
               Enabled         =   0   'False
               Height          =   285
               Index           =   0
               Left            =   210
               TabIndex        =   27
               Top             =   270
               Width           =   1575
            End
            Begin VB.OptionButton optTiles 
               Caption         =   "Use the selected tiles"
               Enabled         =   0   'False
               Height          =   285
               Index           =   1
               Left            =   210
               TabIndex        =   26
               Top             =   660
               Width           =   1935
            End
            Begin VB.OptionButton optTiles 
               Caption         =   "Use the visible tiles"
               Enabled         =   0   'False
               Height          =   285
               Index           =   2
               Left            =   210
               TabIndex        =   25
               Top             =   1050
               Width           =   1995
            End
         End
         Begin VB.Frame fraSuppressTiles 
            Caption         =   "Suppress tiles"
            Height          =   2565
            Left            =   3420
            TabIndex        =   20
            Top             =   750
            Width           =   3255
            Begin VB.CheckBox chkSuppress 
               Caption         =   "Don't use empty tiles.  A tile is empty"
               Enabled         =   0   'False
               Height          =   225
               Left            =   180
               TabIndex        =   22
               Top             =   300
               Width           =   2865
            End
            Begin VB.ListBox lstSuppressTiles 
               Enabled         =   0   'False
               Height          =   1410
               Left            =   180
               Style           =   1  'Checkbox
               TabIndex        =   21
               Top             =   960
               Width           =   2955
            End
            Begin VB.Label Label2 
               Caption         =   "unless it contains data from at least one of the following selected layers:"
               Height          =   465
               Left            =   450
               TabIndex        =   23
               Top             =   480
               Width           =   2595
            End
         End
         Begin MSComctlLib.ListView lvwSheets 
            Height          =   1515
            Left            =   60
            TabIndex        =   28
            Top             =   4620
            Width           =   5205
            _ExtentX        =   9181
            _ExtentY        =   2672
            View            =   2
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.Label Label1 
            Caption         =   $"frmSeriesProperties.frx":0134
            Height          =   615
            Index           =   1
            Left            =   30
            TabIndex        =   29
            Top             =   60
            Width           =   6705
         End
      End
      Begin VB.Frame fraPage3 
         BorderStyle     =   0  'None
         Height          =   3345
         Left            =   60
         TabIndex        =   1
         Top             =   420
         Width           =   6825
         Begin VB.Frame fraExtent 
            Caption         =   "Extent"
            Height          =   2565
            Left            =   90
            TabIndex        =   8
            Top             =   750
            Width           =   3255
            Begin VB.OptionButton optExtent 
               Caption         =   "Data driven - The scale for each tile is"
               Height          =   225
               Index           =   2
               Left            =   120
               TabIndex        =   15
               Top             =   1710
               Width           =   3015
            End
            Begin VB.OptionButton optExtent 
               Caption         =   "Fixed - Always draw at this scale:"
               Height          =   285
               Index           =   1
               Left            =   120
               TabIndex        =   14
               Top             =   990
               Width           =   2925
            End
            Begin VB.OptionButton optExtent 
               Caption         =   "Variable - Fit the tiles to the data frame"
               Height          =   225
               Index           =   0
               Left            =   120
               TabIndex        =   13
               Top             =   270
               Width           =   3045
            End
            Begin VB.TextBox txtMargin 
               Height          =   315
               Left            =   1050
               TabIndex        =   12
               Top             =   540
               Width           =   855
            End
            Begin VB.ComboBox cmbMargin 
               Height          =   315
               Left            =   1950
               Style           =   2  'Dropdown List
               TabIndex        =   11
               Top             =   540
               Width           =   1215
            End
            Begin VB.TextBox txtFixed 
               Height          =   315
               Left            =   930
               TabIndex        =   10
               Top             =   1290
               Width           =   945
            End
            Begin VB.ComboBox cmbDataDriven 
               Height          =   315
               Left            =   420
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   2190
               Width           =   2655
            End
            Begin VB.Label Label4 
               Caption         =   "1:"
               Height          =   255
               Index           =   2
               Left            =   780
               TabIndex        =   41
               Top             =   1320
               Width           =   195
            End
            Begin VB.Label Label4 
               Caption         =   " specified in this index layer field:"
               Height          =   255
               Index           =   0
               Left            =   360
               TabIndex        =   17
               Top             =   1920
               Width           =   2595
            End
            Begin VB.Label Label4 
               Caption         =   "Margin"
               Height          =   255
               Index           =   1
               Left            =   480
               TabIndex        =   16
               Top             =   570
               Width           =   495
            End
         End
         Begin VB.Frame fraOptions 
            Caption         =   "Options"
            Height          =   2565
            Left            =   3450
            TabIndex        =   2
            Top             =   750
            Width           =   3255
            Begin VB.CheckBox chkOptions 
               Caption         =   "Cross-hatch data outside tile?"
               Height          =   225
               Index           =   3
               Left            =   360
               TabIndex        =   42
               Top             =   1260
               Width           =   2565
            End
            Begin VB.CheckBox chkOptions 
               Caption         =   "Rotate data using value from this field:"
               Height          =   225
               Index           =   0
               Left            =   120
               TabIndex        =   7
               Top             =   270
               Width           =   3045
            End
            Begin VB.ComboBox cmbRotateField 
               Height          =   315
               Left            =   390
               Style           =   2  'Dropdown List
               TabIndex        =   6
               Top             =   570
               Width           =   2655
            End
            Begin VB.CheckBox chkOptions 
               Caption         =   "Clip data to the outline of the tile"
               Height          =   225
               Index           =   1
               Left            =   90
               TabIndex        =   5
               Top             =   990
               Width           =   3045
            End
            Begin VB.CheckBox chkOptions 
               Caption         =   "Label neighboring tiles?"
               Height          =   225
               Index           =   2
               Left            =   120
               TabIndex        =   4
               Top             =   1710
               Width           =   3045
            End
            Begin VB.CommandButton cmdLabelProps 
               Caption         =   "Properties..."
               Height          =   345
               Left            =   420
               TabIndex        =   3
               Top             =   2010
               Width           =   1125
            End
         End
         Begin VB.Label Label1 
            Caption         =   "The Map Series provides several different options for fitting a tile to the data frame."
            Height          =   315
            Index           =   2
            Left            =   90
            TabIndex        =   18
            Top             =   180
            Width           =   5955
         End
      End
   End
End
Attribute VB_Name = "frmSeriesProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_pApp As IApplication
Private m_pSeriesOptions As IDSMapSeriesOptions
Private m_pSeriesOptions2 As IDSMapSeriesOptions2
Private m_pTextSym As ISimpleTextSymbol

Private Sub chkOptions_Click(Index As Integer)
  Select Case Index
  Case 0  'Rotate
    If chkOptions(0).Value = 0 Then
      cmbRotateField.Enabled = False
    Else
      cmbRotateField.Enabled = True
    End If
  Case 1  'Clip to outline
    If chkOptions(1).Value = 0 Then
      chkOptions(3).Value = 0
      chkOptions(3).Enabled = False
    Else
      chkOptions(3).Enabled = True
    End If
  Case 2  'Label neighboring tiles
    If chkOptions(2).Value = 0 Then
      cmdLabelProps.Enabled = False
    Else
      cmdLabelProps.Enabled = True
    End If
  End Select
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdLabelProps_Click()
On Error GoTo ErrHand:
  Dim bChanged As Boolean, pTextSymEditor As ITextSymbolEditor
  Set pTextSymEditor = New TextSymbolEditor
  bChanged = pTextSymEditor.EditTextSymbol(m_pTextSym, m_pApp.hwnd)
  Me.SetFocus
  
  Exit Sub
ErrHand:
  MsgBox "cmdLabelProps_Click - " & Err.Description
End Sub

Private Sub cmdOk_Click()
On Error GoTo ErrHand:
  Dim pDoc As IMxDocument, pActive As IActiveView
  
  'Apply updates (only the Options can be updated, so we only need to look at those)
  'Set the clip and rotate properties
  'Update 6/18/03 to support cross hatching of clip area
  If chkOptions(1).Value = 1 Then    'Clip
    If chkOptions(3).Value = 0 Then   'clip without cross hatch
      'Make sure we don't leave the clip element
      If m_pSeriesOptions2.ClipData = 2 Then RemoveClipElement m_pApp.Document
      m_pSeriesOptions2.ClipData = 1
    Else
      m_pSeriesOptions2.ClipData = 2
      Set pDoc = m_pApp.Document
      pDoc.FocusMap.ClipGeometry = Nothing
    End If
'    m_pSeriesOptions.ClipData = True
  Else
    'Make sure we don't leave the clip element
    If m_pSeriesOptions2.ClipData = 2 Then RemoveClipElement m_pApp.Document
    m_pSeriesOptions2.ClipData = 0
'    m_pSeriesOptions.ClipData = False
    'Make sure clipping is turned off for the data frame
    Set pDoc = m_pApp.Document
    pDoc.FocusMap.ClipGeometry = Nothing
  End If
  
  If chkOptions(0).Value = 1 Then     'Rotation
    If m_pSeriesOptions.RotateFrame = False Or m_pSeriesOptions.RotationField <> cmbRotateField.Text Then
      UpdatePageValues "ROTATION", cmbRotateField.Text
    End If
    m_pSeriesOptions.RotateFrame = True
    m_pSeriesOptions.RotationField = cmbRotateField.Text
  Else
    m_pSeriesOptions.RotateFrame = False
    'Make sure rotation is turned off for the data frame
    Set pDoc = m_pApp.Document
    Set pActive = pDoc.FocusMap
    If pActive.ScreenDisplay.DisplayTransformation.Rotation <> 0 Then
      pActive.ScreenDisplay.DisplayTransformation.Rotation = 0
      pActive.Refresh
    End If
  End If
  If chkOptions(2).Value = 1 Then    'Label Neighbors
    m_pSeriesOptions.LabelNeighbors = True
  Else
    m_pSeriesOptions.LabelNeighbors = False
    RemoveLabels pDoc
    g_bLabelNeighbors = False
  End If
  Set m_pSeriesOptions.LabelSymbol = m_pTextSym
  
  'Set the extent properties
  If optExtent(0).Value Then         'Variable
    m_pSeriesOptions.ExtentType = 0
    If txtMargin.Text = "" Then
      m_pSeriesOptions.Margin = 0
    Else
      m_pSeriesOptions.Margin = CDbl(txtMargin.Text)
    End If
    m_pSeriesOptions.MarginType = cmbMargin.ListIndex
  ElseIf optExtent(1).Value Then    'Fixed
    m_pSeriesOptions.ExtentType = 1
    m_pSeriesOptions.FixedScale = txtFixed.Text
  Else                        'Data driven
    If m_pSeriesOptions.ExtentType <> 2 Or m_pSeriesOptions.RotationField <> cmbRotateField.Text Then
      UpdatePageValues "SCALE", cmbDataDriven.Text
    End If
    m_pSeriesOptions.ExtentType = 2
    m_pSeriesOptions.DataDrivenField = cmbDataDriven.Text
  End If
  
  Unload Me
  
  Exit Sub
  
ErrHand:
  MsgBox "cmdOK_Click - " & Err.Description
End Sub

Private Sub UpdatePageValues(sProperty As String, sFieldName As String)
On Error GoTo ErrHand:
  Dim lLoop As Long, pSeries As IDSMapSeries, pPage As IDSMapPage
  Dim pDoc As IMxDocument, pMap As IMap, pSeriesProps As IDSMapSeriesProps
  Dim pIndexLayer As IFeatureLayer, pDataset As IDataset, pWorkspace As IFeatureWorkspace
  Dim pQueryDef As IQueryDef, pCursor As ICursor, pRow As IRow, pColl As Collection
  Set pDoc = m_pApp.Document
  Set pSeries = m_pSeriesOptions
  Set pSeriesProps = pSeries
  Set pMap = FindDataFrame(pDoc, pSeriesProps.DataFrameName)
  If pMap Is Nothing Then Exit Sub
  
  Set pIndexLayer = FindLayer(pSeriesProps.IndexLayerName, pMap)
  If pIndexLayer Is Nothing Then Exit Sub
  
  'Loop through the features in the index layer creating a collection of the scales and tile names
  Set pDataset = pIndexLayer.FeatureClass
  Set pWorkspace = pDataset.Workspace
  Set pQueryDef = pWorkspace.CreateQueryDef
  pQueryDef.Tables = pDataset.Name
  pQueryDef.SubFields = sFieldName & "," & pSeriesProps.IndexFieldName
  Set pCursor = pQueryDef.Evaluate
  Set pColl = New Collection
  Set pRow = pCursor.NextRow
  Do While Not pRow Is Nothing
    If Not IsNull(pRow.Value(0)) And Not IsNull(pRow.Value(1)) Then
      pColl.Add pRow.Value(0), pRow.Value(1)
    End If
    Set pRow = pCursor.NextRow
  Loop
  
  'Now loop through the pages and try to find the corresponding tile name in the collection
  On Error GoTo ErrNoKey:
  For lLoop = 0 To pSeries.PageCount - 1
    Set pPage = pSeries.Page(lLoop)
    If sProperty = "ROTATION" Then
      pPage.PageRotation = pColl.Item(pPage.PageName)
    Else
      pPage.PageScale = pColl.Item(pPage.PageName)
    End If
  Next lLoop

  Exit Sub

ErrNoKey:
  Resume Next
ErrHand:
  MsgBox "UpdatePageValues - " & Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo ErrHand:
  Dim pMapBook As IDSMapBook
  Dim pSeriesProps As IDSMapSeriesProps
  Dim lLoop As Long
  'Check to see if a MapSeries already exists
  Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
  Set pSeriesProps = pMapBook.ContentItem(0)
  Set m_pSeriesOptions = pSeriesProps
  Set m_pSeriesOptions2 = m_pSeriesOptions
  
  'Index Settings Tab
  cmbDetailFrame.Clear
  cmbDetailFrame.AddItem pSeriesProps.DataFrameName
  cmbDetailFrame.Text = pSeriesProps.DataFrameName
  cmbIndexLayer.Clear
  cmbIndexLayer.AddItem pSeriesProps.IndexLayerName
  cmbIndexLayer.Text = pSeriesProps.IndexLayerName
  cmbIndexField.Clear
  cmbIndexField.AddItem pSeriesProps.IndexFieldName
  cmbIndexField.Text = pSeriesProps.IndexFieldName
  
  'Tile Settings Tab
  optTiles(pSeriesProps.TileSelectionMethod) = True
  lstSuppressTiles.Clear
  If pSeriesProps.SuppressLayers Then
    chkSuppress.Value = 1
    For lLoop = 0 To pSeriesProps.SuppressLayerCount - 1
      lstSuppressTiles.AddItem pSeriesProps.SuppressLayer(lLoop)
      lstSuppressTiles.Selected(lLoop) = True
    Next lLoop
  Else
    chkSuppress.Value = 0
  End If
  
  'Options tab
  PopulateFieldCombos
  cmbMargin.Clear
  cmbMargin.AddItem "percent"
  cmbMargin.AddItem "mapunits"
  cmbMargin.Text = "percent"
  optExtent(m_pSeriesOptions.ExtentType).Value = True
  cmdOK.Enabled = True
  Select Case m_pSeriesOptions.ExtentType
  Case 0
    txtMargin.Text = m_pSeriesOptions.Margin
    If m_pSeriesOptions.MarginType = 0 Then
      cmbMargin.Text = "percent"
    Else
      cmbMargin.Text = "mapunits"
    End If
  Case 1
    txtFixed.Text = m_pSeriesOptions.FixedScale
  Case 2
    cmbDataDriven.Text = m_pSeriesOptions.DataDrivenField
  End Select
  If m_pSeriesOptions.RotateFrame Then
    chkOptions(0).Value = 1
    cmbRotateField.Text = m_pSeriesOptions.RotationField
  Else
    chkOptions(0).Value = 0
  End If
  
  'Update 6/18/03 to support cross hatching of clip area
  Select Case m_pSeriesOptions2.ClipData
  Case 0   'No clipping
    chkOptions(1).Value = 0
    chkOptions(3).Value = 0
    chkOptions(3).Enabled = False
  Case 1   'Clip only
    chkOptions(1).Value = 1
    chkOptions(3).Value = 0
    chkOptions(3).Enabled = True
  Case 2   'Clip with cross hatch outside clip area
    chkOptions(1).Value = 1
    chkOptions(3).Value = 1
    chkOptions(3).Enabled = True
  End Select
'  If m_pSeriesOptions.ClipData Then
'    chkOptions(1).Value = 1
'  Else
'    chkOptions(1).Value = 0
'  End If

  If m_pSeriesOptions.LabelNeighbors Then
    chkOptions(2).Value = 1
    cmdLabelProps.Enabled = True
  Else
    chkOptions(2).Value = 0
    cmdLabelProps.Enabled = False
  End If
  Set m_pTextSym = m_pSeriesOptions.LabelSymbol
  
  'Make sure the wizard stays on top
  TopMost Me
  
  Exit Sub
ErrHand:
  MsgBox "frmSeriesProperties_Load - " & Err.Description
End Sub

Private Sub PopulateFieldCombos()
On Error GoTo ErrHand:
  Dim pIndexLayer As IFeatureLayer, pMap As IMap, lLoop As Long
  Dim pFields As IFields, pDoc As IMxDocument
  
  Set pDoc = m_pApp.Document
  Set pMap = FindDataFrame(pDoc, cmbDetailFrame.Text)
  If pMap Is Nothing Then
    MsgBox "Could not find detail frame!!!"
    Exit Sub
  End If
  
  Set pIndexLayer = FindLayer(cmbIndexLayer.Text, pMap)
  If pIndexLayer Is Nothing Then
    MsgBox "Could not find specified layer!!!"
    Exit Sub
  End If
  
  'Populate the index layer combos
  Set pFields = pIndexLayer.FeatureClass.Fields
  cmbDataDriven.Clear
  cmbRotateField.Clear
  For lLoop = 0 To pFields.FieldCount - 1
    Select Case pFields.Field(lLoop).Type
    Case esriFieldTypeDouble, esriFieldTypeSingle, esriFieldTypeInteger
      If UCase(pFields.Field(lLoop).Name) <> "SHAPE_LENGTH" And _
       UCase(pFields.Field(lLoop).Name) <> "SHAPE_AREA" Then
        cmbDataDriven.AddItem pFields.Field(lLoop).Name
        cmbRotateField.AddItem pFields.Field(lLoop).Name
      End If
    End Select
  Next lLoop
  If cmbDataDriven.ListCount > 0 Then
    cmbDataDriven.ListIndex = 0
    cmbRotateField.ListIndex = 0
    optExtent.Item(2).Enabled = True
    chkOptions(0).Enabled = True
  Else
    optExtent.Item(2).Enabled = False
    chkOptions(0).Enabled = False
  End If
  
  Exit Sub
  
ErrHand:
  MsgBox "PopulateFieldCombos - " & Err.Description
End Sub

Private Sub Form_Terminate()
  Set m_pApp = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set m_pApp = Nothing
End Sub

Private Sub optExtent_Click(Index As Integer)
On Error GoTo ErrHand:
  Select Case Index
  Case 0  'Variable
    txtMargin.Enabled = True
    cmbMargin.Enabled = True
    txtFixed.Enabled = False
    cmbDataDriven.Enabled = False
    If txtMargin.Text = "" Then
      cmdOK.Enabled = False
    Else
      cmdOK.Enabled = True
    End If
  Case 1  'Fixed
    txtMargin.Enabled = False
    cmbMargin.Enabled = False
    txtFixed.Enabled = True
    cmbDataDriven.Enabled = False
    If txtFixed.Text = "" Then
      cmdOK.Enabled = False
    Else
      cmdOK.Enabled = True
    End If
  Case 2  'Data driven
    txtMargin.Enabled = False
    cmbMargin.Enabled = False
    txtFixed.Enabled = False
    cmbDataDriven.Enabled = True
    cmdOK.Enabled = True
  End Select

  Exit Sub
ErrHand:
  MsgBox "optExtent_Click - " & Err.Description
End Sub

Private Sub txtFixed_KeyUp(KeyCode As Integer, Shift As Integer)
  If Not IsNumeric(txtFixed.Text) Then
    txtFixed.Text = ""
  End If
  If txtFixed.Text <> "" Then
    cmdOK.Enabled = True
  End If
End Sub

Private Sub txtMargin_KeyUp(KeyCode As Integer, Shift As Integer)
  If Not IsNumeric(txtMargin.Text) Then
    txtMargin.Text = ""
  End If
  If txtMargin.Text <> "" Then
    cmdOK.Enabled = True
  End If
End Sub
