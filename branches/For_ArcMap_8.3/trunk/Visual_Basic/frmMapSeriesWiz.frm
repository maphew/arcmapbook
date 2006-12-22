VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMapSeriesWiz 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Map Sheet Wizard"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraPage1 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   30
      TabIndex        =   4
      Top             =   30
      Visible         =   0   'False
      Width           =   6855
      Begin VB.Frame fraIndexLayer 
         Caption         =   "Index Layer"
         Height          =   2415
         Left            =   3210
         TabIndex        =   10
         Top             =   900
         Width           =   3525
         Begin VB.ComboBox cmbIndexField 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1200
            Width           =   3105
         End
         Begin VB.ComboBox cmbIndexLayer 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   510
            Width           =   3105
         End
         Begin VB.Label lblMapSheet 
            Caption         =   "This field specifies the page name"
            Height          =   225
            Index           =   4
            Left            =   240
            TabIndex        =   13
            Top             =   960
            Width           =   2535
         End
         Begin VB.Label lblMapSheet 
            Caption         =   "Choose the index layer:"
            Height          =   225
            Index           =   3
            Left            =   240
            TabIndex        =   11
            Top             =   270
            Width           =   1725
         End
      End
      Begin VB.ComboBox cmbDetailFrame 
         Height          =   315
         Left            =   270
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1140
         Width           =   2625
      End
      Begin VB.Label Label1 
         Caption         =   $"frmMapSeriesWiz.frx":0000
         Height          =   615
         Index           =   0
         Left            =   30
         TabIndex        =   7
         Top             =   60
         Width           =   6705
      End
      Begin VB.Label lblMapSheet 
         Caption         =   "Choose the detail data frame:"
         Height          =   225
         Index           =   0
         Left            =   30
         TabIndex        =   8
         Top             =   870
         Width           =   2235
      End
   End
   Begin VB.Frame fraPage3 
      BorderStyle     =   0  'None
      Height          =   3525
      Left            =   30
      TabIndex        =   3
      Top             =   90
      Visible         =   0   'False
      Width           =   6855
      Begin VB.Frame fraOptions 
         Caption         =   "Options"
         Height          =   2565
         Left            =   3450
         TabIndex        =   28
         Top             =   750
         Width           =   3255
         Begin VB.CheckBox chkOptions 
            Caption         =   "Cross-hatch data outside tile?"
            Height          =   225
            Index           =   3
            Left            =   390
            TabIndex        =   42
            Top             =   1290
            Width           =   2565
         End
         Begin VB.CommandButton cmdLabelProps 
            Caption         =   "Properties..."
            Height          =   345
            Left            =   420
            TabIndex        =   40
            Top             =   2010
            Width           =   1125
         End
         Begin VB.CheckBox chkOptions 
            Caption         =   "Label neighboring tiles?"
            Height          =   225
            Index           =   2
            Left            =   120
            TabIndex        =   39
            Top             =   1710
            Width           =   3045
         End
         Begin VB.CheckBox chkOptions 
            Caption         =   "Clip data to the outline of the tile"
            Height          =   225
            Index           =   1
            Left            =   90
            TabIndex        =   38
            Top             =   990
            Width           =   3045
         End
         Begin VB.ComboBox cmbRotateField 
            Height          =   315
            Left            =   390
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   570
            Width           =   2655
         End
         Begin VB.CheckBox chkOptions 
            Caption         =   "Rotate data using value from this field:"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   29
            Top             =   270
            Width           =   3045
         End
      End
      Begin VB.Frame fraExtent 
         Caption         =   "Extent"
         Height          =   2565
         Left            =   90
         TabIndex        =   24
         Top             =   750
         Width           =   3255
         Begin VB.ComboBox cmbDataDriven 
            Height          =   315
            Left            =   420
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   2190
            Width           =   2655
         End
         Begin VB.TextBox txtFixed 
            Height          =   315
            Left            =   750
            TabIndex        =   35
            Top             =   1290
            Width           =   945
         End
         Begin VB.ComboBox cmbMargin 
            Height          =   315
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   540
            Width           =   1215
         End
         Begin VB.TextBox txtMargin 
            Height          =   315
            Left            =   1050
            TabIndex        =   33
            Top             =   540
            Width           =   855
         End
         Begin VB.OptionButton optExtent 
            Caption         =   "Variable - Fit the tiles to the data frame"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   27
            Top             =   270
            Width           =   3045
         End
         Begin VB.OptionButton optExtent 
            Caption         =   "Fixed - Always draw at this scale:"
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   26
            Top             =   990
            Width           =   2925
         End
         Begin VB.OptionButton optExtent 
            Caption         =   "Data driven - The scale for each tile is"
            Height          =   225
            Index           =   2
            Left            =   120
            TabIndex        =   25
            Top             =   1710
            Width           =   3015
         End
         Begin VB.Label Label4 
            Caption         =   "1:"
            Height          =   255
            Index           =   2
            Left            =   600
            TabIndex        =   41
            Top             =   1320
            Width           =   195
         End
         Begin VB.Label Label4 
            Caption         =   "Margin"
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   32
            Top             =   570
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   " specified in this index layer field:"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   31
            Top             =   1920
            Width           =   2595
         End
      End
      Begin VB.Label Label1 
         Caption         =   "The Map Series provides several different options for fitting a tile to the data frame."
         Height          =   315
         Index           =   2
         Left            =   90
         TabIndex        =   30
         Top             =   180
         Width           =   5955
      End
   End
   Begin VB.Frame fraPage2 
      BorderStyle     =   0  'None
      Height          =   1965
      Left            =   360
      TabIndex        =   5
      Top             =   540
      Visible         =   0   'False
      Width           =   5805
      Begin VB.Frame fraSuppressTiles 
         Caption         =   "Suppress tiles"
         Height          =   2565
         Left            =   3420
         TabIndex        =   20
         Top             =   750
         Width           =   3255
         Begin VB.ListBox lstSuppressTiles 
            Height          =   1410
            Left            =   180
            Style           =   1  'Checkbox
            TabIndex        =   23
            Top             =   960
            Width           =   2955
         End
         Begin VB.CheckBox chkSuppress 
            Caption         =   "Don't use empty tiles.  A tile is empty"
            Height          =   225
            Left            =   180
            TabIndex        =   21
            Top             =   300
            Width           =   2865
         End
         Begin VB.Label Label2 
            Caption         =   "unless it contains data from at least one of the following selected layers:"
            Height          =   465
            Left            =   450
            TabIndex        =   22
            Top             =   480
            Width           =   2595
         End
      End
      Begin VB.Frame fraChooseTiles 
         Caption         =   "Choose tiles"
         Height          =   2565
         Left            =   60
         TabIndex        =   16
         Top             =   750
         Width           =   3255
         Begin VB.OptionButton optTiles 
            Caption         =   "Use the visible tiles"
            Height          =   285
            Index           =   2
            Left            =   210
            TabIndex        =   19
            Top             =   1050
            Width           =   1995
         End
         Begin VB.OptionButton optTiles 
            Caption         =   "Use the selected tiles"
            Height          =   285
            Index           =   1
            Left            =   210
            TabIndex        =   18
            Top             =   660
            Width           =   1935
         End
         Begin VB.OptionButton optTiles 
            Caption         =   "Use all of the tiles"
            Height          =   285
            Index           =   0
            Left            =   210
            TabIndex        =   17
            Top             =   270
            Width           =   1575
         End
      End
      Begin MSComctlLib.ListView lvwSheets 
         Height          =   1515
         Left            =   60
         TabIndex        =   6
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
         Caption         =   $"frmMapSeriesWiz.frx":00E0
         Height          =   615
         Index           =   1
         Left            =   30
         TabIndex        =   15
         Top             =   60
         Width           =   6705
      End
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "< Back"
      Height          =   345
      Left            =   3330
      TabIndex        =   2
      Top             =   3780
      Width           =   1125
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next >"
      Height          =   345
      Left            =   4470
      TabIndex        =   1
      Top             =   3780
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   5760
      TabIndex        =   0
      Top             =   3780
      Width           =   1125
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   120
      X2              =   6780
      Y1              =   3580
      Y2              =   3580
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      Index           =   0
      X1              =   120
      X2              =   6780
      Y1              =   3570
      Y2              =   3570
   End
End
Attribute VB_Name = "frmMapSeriesWiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_iPage As Integer
Public m_pApp As IApplication
Private m_pCurrentFrame As Frame
Private m_pMap As IMap
Private m_pIndexLayer As IFeatureLayer
Private m_bFormLoad As Boolean
Private m_pTextSym As ISimpleTextSymbol

Private Sub PositionFrame(pFrame As Frame)
On Error GoTo ErrHand:

  If Not m_pCurrentFrame Is Nothing Then m_pCurrentFrame.Visible = False
  pFrame.Visible = True
  pFrame.Height = 3495
  pFrame.Width = 6825
  pFrame.Left = 30
  pFrame.Top = 30
  Set m_pCurrentFrame = pFrame
  pFrame.Visible = True
     
  Exit Sub
ErrHand:
  MsgBox "PositionFrame - " & Err.Description
  Exit Sub
End Sub

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
      chkOptions(3).Enabled = False
      chkOptions(3).Value = 0
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

Private Sub chkSuppress_Click()
  If chkSuppress.Value = 0 Then
    lstSuppressTiles.Enabled = False
  Else
    lstSuppressTiles.Enabled = True
  End If
End Sub

Private Sub cmbDetailFrame_Click()
On Error GoTo ErrHand:
  Dim pDoc As IMxDocument, lLoop As Long
  Dim pFeatLayer As IFeatureLayer, pGroupLayer As ICompositeLayer
  
  'Set the Next button to false
  cmdNext.Enabled = False
  
  'Find the selected map
  cmbIndexLayer.Clear
  If cmbDetailFrame.Text = "" Then
    MsgBox "No detail frame selected!!!"
    Exit Sub
  End If
  
  Set pDoc = m_pApp.Document
  Set m_pMap = FindDataFrame(pDoc, cmbDetailFrame.Text)
  If m_pMap Is Nothing Then
    MsgBox "Could not find detail frame!!!"
    Exit Sub
  End If
  
  'Populate the index layer combo
  lstSuppressTiles.Clear
  cmbIndexLayer.Clear
  For lLoop = 0 To m_pMap.LayerCount - 1
    If TypeOf m_pMap.Layer(lLoop) Is ICompositeLayer Then
      CompositeLayer m_pMap.Layer(lLoop)
    Else
      LayerCheck m_pMap.Layer(lLoop)
    End If
  Next lLoop
  If cmbIndexLayer.ListCount = 0 Then
    MsgBox "You need at least one polygon layer in the detail frame to serve as the index layer!!!"
  Else
    cmbIndexLayer.ListIndex = 0
  End If
  
  Exit Sub
ErrHand:
  MsgBox "cmbDetailFrame_Click - " & Err.Description
End Sub

Private Sub CompositeLayer(pCompLayer As ICompositeLayer)
On Error GoTo ErrHand:
  Dim lLoop As Long
  For lLoop = 0 To pCompLayer.count - 1
    If TypeOf pCompLayer.Layer(lLoop) Is ICompositeLayer Then
      CompositeLayer pCompLayer.Layer(lLoop)
    Else
      LayerCheck pCompLayer.Layer(lLoop)
    End If
  Next lLoop

  Exit Sub
ErrHand:
  MsgBox "CompositeLayer - " & Err.Description
End Sub

Private Sub LayerCheck(pLayer As ILayer)
On Error GoTo ErrHand:
  Dim pFeatLayer As IFeatureLayer
  
  If TypeOf pLayer Is IFeatureLayer Then
    Set pFeatLayer = pLayer
    If pFeatLayer.FeatureClass.ShapeType = esriGeometryPolygon Then
      cmbIndexLayer.AddItem pFeatLayer.Name
    End If
    lstSuppressTiles.AddItem pFeatLayer.Name
  End If

  Exit Sub
ErrHand:
  MsgBox "LayerCheck - " & Err.Description
End Sub

Private Sub cmbIndexLayer_Click()
On Error GoTo ErrHand:
  Dim lLoop As Long, pFields As IFields, pField As IField
  
  'Set the Next button to false
  cmdNext.Enabled = False
  
  'Find the selected layer
  cmbIndexField.Clear
  If cmbIndexLayer.Text = "" Then
    MsgBox "No index layer selected!!!"
    Exit Sub
  End If
  
  Set m_pIndexLayer = FindLayer(cmbIndexLayer.Text, m_pMap)
  If m_pIndexLayer Is Nothing Then
    MsgBox "Could not find specified layer!!!"
    Exit Sub
  End If
  
  'Populate the index layer combos
  Set pFields = m_pIndexLayer.FeatureClass.Fields
  cmbDataDriven.Clear
  cmbRotateField.Clear
  For lLoop = 0 To pFields.FieldCount - 1
    Select Case pFields.Field(lLoop).Type
    Case esriFieldTypeString
      cmbIndexField.AddItem pFields.Field(lLoop).Name
    Case esriFieldTypeDouble, esriFieldTypeSingle, esriFieldTypeInteger
      If UCase(pFields.Field(lLoop).Name) <> "SHAPE_LENGTH" And _
       UCase(pFields.Field(lLoop).Name) <> "SHAPE_AREA" Then
        cmbDataDriven.AddItem pFields.Field(lLoop).Name
        cmbRotateField.AddItem pFields.Field(lLoop).Name
      End If
    End Select
  Next lLoop
  If cmbIndexField.ListCount = 0 Then
'    MsgBox "You need at least one string field in the layer for labeling the pages!!!"
  Else
    cmbIndexField.ListIndex = 0
    cmdNext.Enabled = True
  End If
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
  MsgBox "cmbIndexField_Click - " & Err.Description
End Sub

Private Sub cmdBack_Click()
  m_pCurrentFrame.Visible = False
  Select Case m_iPage
  Case 2
    PositionFrame fraPage1
    m_iPage = 1
  Case 3
    cmdNext.Caption = "Next >"
    PositionFrame fraPage2
    m_iPage = 2
  End Select
  cmdNext.Enabled = True
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

Private Sub cmdNext_Click()
On Error GoTo ErrHand:
  Dim pMapSeries As IDSMapSeries
  m_pCurrentFrame.Visible = False
  cmdBack.Enabled = True
  Select Case m_iPage
  Case 1  'Done with date frame and index layer
    CheckForSelected    'Check index layer to see if there are selected features
    PositionFrame fraPage2
    m_iPage = 2
  Case 2  'Done with tile specification
    PositionFrame fraPage3
    m_iPage = 3
    cmdNext.Caption = "Finish"
    If optExtent(0).Value Then
      If txtMargin.Text = "" Then
        cmdNext.Enabled = False
      Else
        cmdNext.Enabled = True
      End If
    ElseIf optExtent(1).Value Then
      If txtFixed.Text = "" Then
        cmdNext.Enabled = False
      Else
        cmdNext.Enabled = True
      End If
    Else
      cmdNext.Enabled = True
    End If
  Case 3  'Finish button selected
    CreateSeries
    Unload Me
  End Select
  
  Exit Sub
ErrHand:
  MsgBox "cmdNext_click - " & Err.Description
  Exit Sub
End Sub

Private Sub CreateSeries()
On Error GoTo ErrHandler:
  Dim pMapSeries As IDSMapSeries, pSpatialQuery As ISpatialFilter
  Dim pTmpPage As tmpPageClass, pTmpColl As Collection, pClone As IClone
  Dim pSeriesOpt As IDSMapSeriesOptions, pFeatLayerSel As IFeatureSelection
  Dim pSeriesProps As IDSMapSeriesProps, pMapPage As IDSMapPage
  Dim pDoc As IMxDocument, pMap As IMap, lCount As Long, lLoop As Long
  Dim pFeatLayer As IFeatureLayer, pQuery As IQueryFilter, pCursor As IFeatureCursor
  Dim pFeature As IFeature, lIndex As Long, sName As String, sFieldName As String
  Dim pNode As Node, pMapBook As IDSMapBook
  Dim pActiveView As IActiveView, lRotIndex As Long, lScaleIndex As Long
  'Added 6/18/03 to support cross hatch outside clip area
  Dim pSeriesOpt2 As IDSMapSeriesOptions2
  
  Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
  pMapBook.EnableBook = True
  Set pDoc = m_pApp.Document
  
  Set pMapSeries = New DSMapSeries
  Set pSeriesOpt = pMapSeries
  Set pSeriesOpt2 = pSeriesOpt  'Added 6/18/03 to support cross hatch outside clip area
  Set pSeriesProps = pMapSeries
  pMapSeries.EnableSeries = True
  
  'Find the detail frame
  Set pMap = FindDataFrame(pDoc, cmbDetailFrame.Text)
  If pMap Is Nothing Then
    MsgBox "Detail frame not found!!!"
    Exit Sub
  End If
  pSeriesProps.DataFrameName = pMap.Name
  
  'Find the layer
  Set pFeatLayer = FindLayer(cmbIndexLayer.Text, pMap)
  If pFeatLayer Is Nothing Then
    MsgBox "Index layer not found!!!"
    Exit Sub
  End If
  pSeriesProps.IndexLayerName = pFeatLayer.Name
  pSeriesProps.IndexFieldName = cmbIndexField.Text
    
  'Determine the tiles we are interested in
  Set pQuery = New QueryFilter
  sFieldName = cmbIndexField.Text
  pQuery.AddField sFieldName
'  pQuery.WhereClause = sFieldName & " <> ''"
  pQuery.WhereClause = sFieldName & " is not null"
  If optTiles(0).Value Then
    Set pCursor = pFeatLayer.Search(pQuery, True)
    pSeriesProps.TileSelectionMethod = 0
  ElseIf optTiles(1).Value Then
    Set pFeatLayerSel = pFeatLayer
    pFeatLayerSel.SelectionSet.Search pQuery, True, pCursor
    pSeriesProps.TileSelectionMethod = 1
  Else
    Set pActiveView = pMap
    Set pSpatialQuery = New SpatialFilter
    pSpatialQuery.AddField sFieldName
    pSpatialQuery.SpatialRel = esriSpatialRelIntersects
    Set pSpatialQuery.Geometry = pActiveView.Extent
    pSpatialQuery.WhereClause = sFieldName & " <> ''"
    pSpatialQuery.GeometryField = pFeatLayer.FeatureClass.shapeFieldName
    Set pCursor = pFeatLayer.Search(pSpatialQuery, True)
    pSeriesProps.TileSelectionMethod = 2
  End If
  
  'Set the clip, label and rotate properties
  'Updated 6/18/03 to support cross hatch outside clip area
  If chkOptions(1).Value = 1 Then
    If chkOptions(3).Value = 1 Then
      pSeriesOpt2.ClipData = 2
    Else
      pSeriesOpt2.ClipData = 1
    End If
  Else
    pSeriesOpt2.ClipData = 0
  End If
'  If chkOptions(1).Value = 1 Then
'    pSeriesOpt.ClipData = True
'  Else
'    pSeriesOpt.ClipData = False
'  End If
  
  If chkOptions(0).Value = 1 Then
    pSeriesOpt.RotateFrame = True
    pSeriesOpt.RotationField = cmbRotateField.Text
    lRotIndex = pFeatLayer.FeatureClass.FindField(cmbRotateField.Text)
  Else
    pSeriesOpt.RotateFrame = False
  End If
  If chkOptions(2).Value = 1 Then
    pSeriesOpt.LabelNeighbors = True
  Else
    pSeriesOpt.LabelNeighbors = False
  End If
  Set pSeriesOpt.LabelSymbol = m_pTextSym
  
  'Set the extent properties
  If optExtent(0).Value Then         'Variable
    pSeriesOpt.ExtentType = 0
    If txtMargin.Text = "" Then
      pSeriesOpt.Margin = 0
    Else
      pSeriesOpt.Margin = CDbl(txtMargin.Text)
    End If
    pSeriesOpt.MarginType = cmbMargin.ListIndex
  ElseIf optExtent(1).Value Then    'Fixed
    pSeriesOpt.ExtentType = 1
    pSeriesOpt.FixedScale = txtFixed.Text
  Else                        'Data driven
    pSeriesOpt.ExtentType = 2
    pSeriesOpt.DataDrivenField = cmbDataDriven.Text
    lScaleIndex = pFeatLayer.FeatureClass.FindField(cmbDataDriven.Text)
  End If
  
  'Store suppression information
  If chkSuppress.Value = 1 And lstSuppressTiles.SelCount > 0 Then
    pSeriesProps.SuppressLayers = True
    For lLoop = 0 To lstSuppressTiles.ListCount - 1
      If lstSuppressTiles.Selected(lLoop) Then
        pSeriesProps.AddLayerToSuppress lstSuppressTiles.List(lLoop)
      End If
    Next lLoop
  Else
    pSeriesProps.SuppressLayers = False
  End If
  
  'Create the pages and populate the treeview
  Set pTmpColl = New Collection
  lIndex = pFeatLayer.FeatureClass.FindField(sFieldName)
  Set pFeature = pCursor.NextFeature
  With g_pFrmMapSeries.tvwMapBook
    Set pNode = .Nodes.Add("MapBook", tvwChild, "MapSeries", "Map Series", 3)
    pNode.Tag = "MapSeries"
    
    'Add tile names to a listbox first for sort purposes
    g_pFrmMapSeries.lstSorter.Clear
    Do While Not pFeature Is Nothing
      sName = pFeature.Value(lIndex)
      Set pTmpPage = New tmpPageClass
      pTmpPage.PageName = sName
      pTmpPage.PageRotation = 0
      pTmpPage.PageScale = 1
      Set pClone = pFeature.Shape
      Set pTmpPage.PageShape = pClone.Clone
      'Track the rotation and scale values (if we are going to use them) to the end
      'of the name, so we can assign them to the page when it is added without having
      'to query the index layer again.
      If chkOptions(0).Value = 1 And lRotIndex >= 0 Then
        If Not IsNull(pFeature.Value(lRotIndex)) Then
          pTmpPage.PageRotation = pFeature.Value(lRotIndex)
        End If
      End If
      If optExtent(2).Value And lScaleIndex >= 0 Then
        If Not IsNull(pFeature.Value(lScaleIndex)) Then
          pTmpPage.PageScale = pFeature.Value(lScaleIndex)
        End If
      End If
      If chkSuppress.Value = 1 And lstSuppressTiles.SelCount > 0 Then
        If FeaturesInTile(pFeature, pMap) Then
          g_pFrmMapSeries.lstSorter.AddItem sName
          pTmpColl.Add pTmpPage, sName
        End If
      Else
        g_pFrmMapSeries.lstSorter.AddItem sName
        pTmpColl.Add pTmpPage, sName
      End If
      Set pFeature = pCursor.NextFeature
    Loop
    
    'Now loop back through the list and add the tile names as nodes in the tree
    For lLoop = 0 To g_pFrmMapSeries.lstSorter.ListCount - 1
      Set pMapPage = New DSMapPage
      sName = g_pFrmMapSeries.lstSorter.List(lLoop)
      Set pNode = .Nodes.Add("MapSeries", tvwChild, "a" & sName, lLoop + 1 & " - " & sName, 5)
      Set pTmpPage = pTmpColl.Item(sName)
      pNode.Tag = lLoop
      pMapPage.PageName = sName
      pMapPage.PageRotation = pTmpPage.PageRotation
      pMapPage.PageScale = pTmpPage.PageScale
      Set pMapPage.PageShape = pTmpPage.PageShape
      pMapPage.LastOutputted = #1/1/1900#
      pMapPage.EnablePage = True
      pMapSeries.AddPage pMapPage
    Next lLoop
    .Nodes.Item("MapBook").Expanded = True
    .Nodes.Item("MapSeries").Expanded = True
  End With
  
  'Add the series to the book
  pMapBook.AddContent pMapSeries

  Exit Sub
ErrHandler:
  MsgBox "CreateSeries - most likely you do not have unique names in your index layer!!!"
End Sub

Private Sub CheckForSelected()
On Error GoTo ErrHand:
  Dim pFeatSel As IFeatureSelection
  
  'Make sure there is something to check
  optTiles(1).Enabled = False
  If m_pIndexLayer Is Nothing Then Exit Sub
  
  'Check for selected features in the index layer
  Set pFeatSel = m_pIndexLayer
  If pFeatSel.SelectionSet.count <> 0 Then
    optTiles(1).Enabled = True
  End If

  Exit Sub
ErrHand:
  MsgBox "CheckForSelected - " & Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo ErrHand:
  Dim pDoc As IMxDocument, lLoop As Long
  'Get the extension
  If m_pApp Is Nothing Then Exit Sub
    
  m_bFormLoad = True
  Set m_pCurrentFrame = Nothing
  PositionFrame fraPage1
  cmdNext.Enabled = False
  cmdBack.Enabled = False
  
  'Initialize variables and controls
  m_iPage = 1
  chkOptions(0).Value = 0
  chkOptions(1).Value = 0
  chkOptions(2).Value = 0
  chkSuppress.Value = 0
  optTiles(0).Value = True
  optExtent(0).Value = True
  lstSuppressTiles.Enabled = False
  cmbRotateField.Enabled = False
  cmdLabelProps.Enabled = False
  chkOptions(3).Enabled = False
  
  'Populate the data frame combo
  Set pDoc = m_pApp.Document
  cmbIndexField.Clear
  cmbDetailFrame.Clear
  For lLoop = 0 To pDoc.Maps.count - 1
    cmbDetailFrame.AddItem pDoc.Maps.Item(lLoop).Name
  Next lLoop
  cmbDetailFrame.ListIndex = 0
  m_bFormLoad = False
  
  'Populate the extent options
  cmbMargin.Clear
  cmbMargin.AddItem "percent"
  cmbMargin.AddItem "mapunits"
  cmbMargin.Text = "percent"
  txtMargin.Text = "0"
  
  'Set the initial Label symbol
  Set pDoc = m_pApp.Document
  Set m_pTextSym = New TextSymbol
  m_pTextSym.Font = pDoc.DefaultTextFont
  m_pTextSym.Size = pDoc.DefaultTextFontSize.Size
  
  'Make sure the wizard stays on top
  TopMost Me

  Exit Sub
  
ErrHand:
  MsgBox "frmMapSheetWiz Load - " & Err.Description
  Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set m_pApp = Nothing
  Set m_pCurrentFrame = Nothing
  Set m_pMap = Nothing
  Set m_pIndexLayer = Nothing
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
      cmdNext.Enabled = False
    Else
      cmdNext.Enabled = True
    End If
  Case 1  'Fixed
    txtMargin.Enabled = False
    cmbMargin.Enabled = False
    txtFixed.Enabled = True
    cmbDataDriven.Enabled = False
    If txtFixed.Text = "" Then
      cmdNext.Enabled = False
    Else
      cmdNext.Enabled = True
    End If
  Case 2  'Data driven
    txtMargin.Enabled = False
    cmbMargin.Enabled = False
    txtFixed.Enabled = False
    cmbDataDriven.Enabled = True
    cmdNext.Enabled = True
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
    cmdNext.Enabled = True
  End If
End Sub

Private Sub txtMargin_KeyUp(KeyCode As Integer, Shift As Integer)
  If Not IsNumeric(txtMargin.Text) Then
    txtMargin.Text = ""
  End If
  If txtMargin.Text <> "" Then
    cmdNext.Enabled = True
  End If
End Sub

Private Function FeaturesInTile(pFeature As IFeature, pMap As IMap) As Boolean
'Routine for determining whether the specified tile feature (pFeature) should
'be suppressed.  Tiles are suppressed when there are no features from the checked
'layers in them.
On Error GoTo ErrHand:
  Dim lLoop As Long, pFeatLayer As IFeatureLayer, pSpatial As ISpatialFilter
  Dim pFeatCursor As IFeatureCursor, pSearchFeat As IFeature
  
  FeaturesInTile = False
  
  Set pSpatial = New SpatialFilter
  pSpatial.SpatialRel = esriSpatialRelIntersects
  Set pSpatial.Geometry = pFeature.Shape
  For lLoop = 0 To lstSuppressTiles.ListCount - 1
    If lstSuppressTiles.Selected(lLoop) Then
      Set pFeatLayer = FindLayer(lstSuppressTiles.List(lLoop), pMap)
      pSpatial.GeometryField = pFeatLayer.FeatureClass.shapeFieldName
      Set pFeatCursor = pFeatLayer.Search(pSpatial, True)
      Set pSearchFeat = pFeatCursor.NextFeature
      If Not pSearchFeat Is Nothing Then
        FeaturesInTile = True
        Exit Function
      End If
    End If
  Next lLoop

  Exit Function
  
ErrHand:
  MsgBox "FeaturesInTile - " & Err.Description
End Function


