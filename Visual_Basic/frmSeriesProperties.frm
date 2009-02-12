VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSeriesProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Map Series Properties"
   ClientHeight    =   4740
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7284
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   7284
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   5940
      TabIndex        =   40
      Top             =   4320
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   345
      Left            =   4680
      TabIndex        =   39
      Top             =   4320
      Width           =   1125
   End
   Begin TabDlg.SSTab tabProperties 
      Height          =   4245
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      _ExtentX        =   12510
      _ExtentY        =   7493
      _Version        =   393216
      Tabs            =   6
      TabHeight       =   520
      TabCaption(0)   =   "Index Settings"
      TabPicture(0)   =   "frmSeriesProperties.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraPage1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tile Settings"
      TabPicture(1)   =   "frmSeriesProperties.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraPage2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Options"
      TabPicture(2)   =   "frmSeriesProperties.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraPage3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Layer Groups"
      TabPicture(3)   =   "frmSeriesProperties.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraPage4"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Layer Filters"
      TabPicture(4)   =   "frmSeriesProperties.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraPage6"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "NW Options"
      TabPicture(5)   =   "frmSeriesProperties.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fraBubbleLayer"
      Tab(5).Control(1)=   "fraAdditionalDataFrames"
      Tab(5).ControlCount=   2
      Begin VB.Frame fraAdditionalDataFrames 
         Caption         =   "Additional Map Series Data Frames"
         Height          =   2055
         Left            =   -74880
         TabIndex        =   71
         Top             =   2040
         Width           =   6615
         Begin VB.ComboBox cboOtherSeriesExtentOption 
            Height          =   315
            Left            =   3240
            Style           =   2  'Dropdown List
            TabIndex        =   80
            Top             =   840
            Width           =   3255
         End
         Begin VB.ComboBox cboOtherSeriesDF_AttributeField 
            Height          =   315
            Left            =   3480
            Style           =   2  'Dropdown List
            TabIndex        =   78
            Top             =   1680
            Width           =   3015
         End
         Begin VB.ComboBox cboOtherSeriesDF_PolygonFC 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   76
            Top             =   1680
            Width           =   3135
         End
         Begin VB.ListBox lstOtherSeriesDF 
            Height          =   696
            Left            =   3240
            Style           =   1  'Checkbox
            TabIndex        =   73
            Top             =   120
            Width           =   3252
         End
         Begin VB.Label Label7 
            Caption         =   "Select a map page extent option :"
            Height          =   255
            Left            =   360
            TabIndex        =   81
            Top             =   840
            Width           =   2415
         End
         Begin VB.Label lblOtherDF_AttributeField 
            Caption         =   "Map Page ID attribute field :"
            Height          =   255
            Left            =   3600
            TabIndex        =   77
            Top             =   1440
            Width           =   2895
         End
         Begin VB.Label lblOtherDF_PolygonFC 
            Caption         =   "Map Series Polygon Feature Class :"
            Height          =   255
            Left            =   120
            TabIndex        =   75
            Top             =   1440
            Width           =   2775
         End
         Begin VB.Label lblMapSeriesPolyFCOtherDF 
            Caption         =   "For the selected data frame, select the map series polygon feature class and map ID field."
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   1200
            Width           =   6375
         End
         Begin VB.Label Label3 
            Caption         =   "Select the additional data frames where the current map page will be loaded."
            Height          =   375
            Left            =   360
            TabIndex        =   72
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.Frame fraBubbleLayer 
         Caption         =   "Map Bubble Layer"
         Height          =   1335
         Left            =   -74790
         TabIndex        =   61
         Top             =   720
         Width           =   6585
         Begin VB.CheckBox chkRefreshPage 
            Caption         =   "Map refresh command will update bubble attributes."
            Height          =   255
            Left            =   120
            TabIndex        =   79
            Top             =   960
            Width           =   4092
         End
         Begin VB.CheckBox chkBubbleLayer 
            Caption         =   "Use a polygon feature class to define circular map detail ""bubble"" insets."
            Height          =   435
            Left            =   105
            TabIndex        =   63
            Top             =   480
            Width           =   2970
         End
         Begin VB.ComboBox cboBubbleLayer 
            Height          =   315
            Left            =   3120
            Style           =   2  'Dropdown List
            TabIndex        =   62
            Top             =   480
            Width           =   3255
         End
         Begin VB.Label lblBubbleLayerWarning 
            Caption         =   "No circular inset layer was detected."
            Height          =   285
            Left            =   180
            TabIndex        =   65
            Top             =   240
            Width           =   2865
         End
         Begin VB.Label lblBubbleLayer 
            Caption         =   "Choose the bubble definitions layer :"
            Height          =   285
            Left            =   3240
            TabIndex        =   64
            Top             =   120
            Width           =   2835
         End
      End
      Begin VB.Frame fraPage6 
         Height          =   3255
         Left            =   -74880
         TabIndex        =   54
         Top             =   720
         Width           =   6735
         Begin VB.CommandButton cmdNextPageName 
            Caption         =   "Next >"
            Height          =   255
            Left            =   4920
            TabIndex        =   70
            Top             =   2760
            Width           =   975
         End
         Begin VB.CommandButton cmdPrevPageName 
            Caption         =   "< Prev"
            Height          =   255
            Left            =   2880
            TabIndex        =   69
            Top             =   2760
            Width           =   975
         End
         Begin VB.ComboBox cboDataFrameToFilter 
            Height          =   315
            Left            =   2880
            Style           =   2  'Dropdown List
            TabIndex        =   66
            Top             =   840
            Width           =   3735
         End
         Begin VB.TextBox txtDefinitionQuery 
            Height          =   1215
            HideSelection   =   0   'False
            Left            =   2880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   57
            Text            =   "frmSeriesProperties.frx":00A8
            Top             =   1440
            Width           =   3735
         End
         Begin VB.CommandButton cmdInsertPageName 
            Caption         =   "Insert"
            Height          =   255
            Left            =   3960
            TabIndex        =   56
            Top             =   2760
            Width           =   855
         End
         Begin VB.ListBox lstFilterLayers 
            Height          =   1128
            ItemData        =   "frmSeriesProperties.frx":00BB
            Left            =   120
            List            =   "frmSeriesProperties.frx":00BD
            Style           =   1  'Checkbox
            TabIndex        =   55
            Top             =   1440
            Width           =   2655
         End
         Begin VB.Label lblNoPageNameWarning 
            Caption         =   "Warning - No map page has been selected."
            Height          =   375
            Left            =   120
            TabIndex        =   68
            Top             =   2760
            Width           =   1815
         End
         Begin VB.Label lblLayerFiltersDataFrame 
            Caption         =   "Data Frame to be filtered"
            Height          =   255
            Left            =   960
            TabIndex        =   67
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label8 
            Caption         =   $"frmSeriesProperties.frx":00BF
            Height          =   615
            Left            =   240
            TabIndex        =   60
            Top             =   120
            Width           =   6255
         End
         Begin VB.Label lblLayers6 
            Caption         =   "Layers in that data frame"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   1200
            Width           =   2655
         End
         Begin VB.Label Label9 
            Caption         =   "Definition Query for that layer"
            Height          =   255
            Left            =   2880
            TabIndex        =   58
            Top             =   1200
            Width           =   3615
         End
      End
      Begin VB.Frame fraPage4 
         BorderStyle     =   0  'None
         Height          =   3375
         Left            =   -74880
         TabIndex        =   47
         Top             =   720
         Width           =   6735
         Begin VB.ListBox lstLyrGroups 
            Height          =   2160
            ItemData        =   "frmSeriesProperties.frx":01C7
            Left            =   120
            List            =   "frmSeriesProperties.frx":01D7
            TabIndex        =   51
            Top             =   480
            Width           =   1575
         End
         Begin VB.ListBox lstVisibleLayers 
            Height          =   2208
            ItemData        =   "frmSeriesProperties.frx":021D
            Left            =   1800
            List            =   "frmSeriesProperties.frx":023C
            Style           =   1  'Checkbox
            TabIndex        =   50
            Top             =   480
            Width           =   4815
         End
         Begin VB.CommandButton cmdAddGroup 
            Caption         =   "&Add Group"
            Height          =   375
            Left            =   120
            TabIndex        =   49
            Top             =   2880
            Width           =   975
         End
         Begin VB.CommandButton cmdRemoveGroup 
            Caption         =   "&Remove"
            Height          =   375
            Left            =   1200
            TabIndex        =   48
            Top             =   2880
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "Group Names"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label6 
            Caption         =   "Available Layers"
            Height          =   255
            Left            =   1800
            TabIndex        =   52
            Top             =   240
            Width           =   3375
         End
      End
      Begin VB.Frame fraPage1 
         BorderStyle     =   0  'None
         Height          =   3375
         Left            =   60
         TabIndex        =   30
         Top             =   720
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
            Caption         =   $"frmSeriesProperties.frx":02A7
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
         Top             =   720
         Width           =   6765
         Begin VB.TextBox txtNumbering 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2520
            TabIndex        =   43
            Text            =   "1"
            Top             =   2610
            Width           =   525
         End
         Begin VB.Frame fraChooseTiles 
            Caption         =   "Choose tiles"
            Height          =   1485
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
               Height          =   1344
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
            _ExtentX        =   9186
            _ExtentY        =   2667
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
            Caption         =   "Begin numbering tiles/pages at:"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   44
            Top             =   2640
            Width           =   2295
         End
         Begin VB.Label Label1 
            Caption         =   $"frmSeriesProperties.frx":0387
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
         Left            =   -74940
         TabIndex        =   1
         Top             =   720
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
            Begin VB.TextBox txtNeighborLabelIndent 
               Enabled         =   0   'False
               Height          =   285
               Left            =   2040
               TabIndex        =   46
               Text            =   "txtNeighborLabelIndent"
               Top             =   2160
               Width           =   975
            End
            Begin VB.CheckBox chkOptions 
               Caption         =   "Cross-hatch data outside tile?"
               Height          =   225
               Index           =   3
               Left            =   360
               TabIndex        =   42
               Top             =   1320
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
            Begin VB.Label lblIndentDistance 
               Caption         =   "Label Indent (in) :"
               Height          =   255
               Left            =   1680
               TabIndex        =   45
               Top             =   1920
               Width           =   1335
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

Option Explicit

Public m_pApp As IApplication
Private m_pSeriesOptions As INWDSMapSeriesOptions
Private m_pSeriesOptions2 As INWDSMapSeriesOptions2
Private m_pNWSeriesOptions As INWMapSeriesOptions
Private m_pMapBook As INWDSMapBook
Private m_pTextSym As ISimpleTextSymbol
Private m_sSelectedGroup As String
Private m_bIsRecursiveEntry_lstVisibleLayers As Boolean
Private m_bIsRecursiveEntrylstLyrGroups As Boolean
Private m_bInitializing As Boolean
Private m_sCurrentDataFrame As String
'dynamic definition query replacement string highlighting
'Private m_lStrStart As Long, m_lStrEnd As Long, m_sSearchStr As String
Private m_bDefQueryInitializing As Boolean  'required to block unwanted _Change event from running
Private m_bRecursiveTxtDefinitionQuery_Change As Boolean 'blocks recursive calls to this _Change event
Private m_pFeatLayerDefinition As IFeatureLayerDefinition

Private m_pDictDFsToUpdate_Layer As Scripting.Dictionary
Private m_pDictDFsToUpdate_Field As Scripting.Dictionary

Const c_sModuleFileName As String = "frmSeriesProperties.frm"





'  Assign to the NW map series options the selected bubble
'  definition layer
'-------------------------------
Private Sub cboBubbleLayer_Click()
  If m_bInitializing Then Exit Sub
62:   With cboBubbleLayer
    If m_pNWSeriesOptions Is Nothing Then Exit Sub
64:     m_pNWSeriesOptions.BubbleLayer = .List(.ListIndex)
65:   End With
End Sub






Private Sub cboDataFrameToFilter_Click()
74:   m_sCurrentDataFrame = cboDataFrameToFilter.List(cboDataFrameToFilter.ListIndex)
75:   LoadLayersUIFromDataFrame m_sCurrentDataFrame
76:   CleanOrphanedDynamicDefQueryDataFrames
End Sub







'  cboOtherSeriesDF_AttributeField
'    Change
'      - make a listindex selection based on any previous
'        selection, otherwise just select the first
'        qualified attribute field (triggers attr click
'        event)
'
'assumptions:
'   m_pDictDFsToUpdate_Layer and m_pDictDFsToUpdate_Field each
'   have been initialized.
'
'   lstOtherSeriesDF is populated with values, and one of them
'   has been selected with a check mark next to it.
'
'   cboOtherSeriesDF is populated with a list of layers from that
'   data frame, and that control's .listindex has the current layer.
'
'-------------------------------------------------
Private Sub cboOtherSeriesDF_AttributeField_Change()
  On Error GoTo ErrorHandler

106:   cboOtherSeriesDF_AttributeField_ChangeHandler
  
'  Dim sDataFrameName As String, sLayerName As String, sFieldToFind As String
'  Dim lFieldIdx As Long
'                                            'grab the data frame name and the
'                                            'layer name
'  With lstOtherSeriesDF
'    sDataFrameName = .List(.ListIndex)
'  End With
'  With cboOtherSeriesDF_PolygonFC
'    sLayerName = .List(.ListIndex)
'  End With
'                                            'first check for the existence of a
'                                            'local selection,
'  sFieldToFind = ""
'  With m_pDictDFsToUpdate_Layer
'    If .Exists(sDataFrameName) Then
'      If StrComp(sLayerName, .Item(sDataFrameName), vbTextCompare) = 0 Then
'        sFieldToFind = m_pDictDFsToUpdate_Field(sDataFrameName)
'      End If
'    End If
'  End With
'                                            'if that doesn't exist, check for the existence of
'                                            'an m_pNWSeriesOptions previous selection
'  If Len(sFieldToFind) = 0 Then
'    sFieldToFind = m_pNWSeriesOptions.DataFrameToUpdateGetPageNameField( _
'      sDataFrameName, sLayerName)
'  End If
'
'  With cboOtherSeriesDF_AttributeField
'                                            'if that doesn't exist, select the first
'                                            'listindex
'    If Len(sFieldToFind) = 0 Then
'      .ListIndex = 0
'      sFieldToFind = .List(.ListIndex)
'                                            'otherwise select the correct index for
'                                            'that attribute field name
'    Else
'      lFieldIdx = FindControlString(cboOtherSeriesDF_AttributeField, sFieldToFind)
'      .ListIndex = lFieldIdx
'    End If
'  End With
'

  Exit Sub
ErrorHandler:
  HandleError False, "cboOtherSeriesDF_AttributeField_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub



Private Sub cboOtherSeriesDF_AttributeField_ChangeHandler()
  On Error GoTo ErrorHandler

  Dim sDataFrameName As String, sLayerName As String, sFieldToFind As String
  Dim lFieldIdx As Long
                                            'grab the data frame name and the
                                            'layer name
164:   With lstOtherSeriesDF
165:     sDataFrameName = .List(.ListIndex)
166:   End With
167:   With cboOtherSeriesDF_PolygonFC
168:     sLayerName = .List(.ListIndex)
169:   End With
                                            'first check for the existence of a
                                            'local selection,
172:   sFieldToFind = ""
173:   With m_pDictDFsToUpdate_Layer
174:     If .Exists(sDataFrameName) Then
175:       If StrComp(sLayerName, .Item(sDataFrameName), vbTextCompare) = 0 Then
176:         sFieldToFind = m_pDictDFsToUpdate_Field(sDataFrameName)
177:       End If
178:     End If
179:   End With
                                            'if that doesn't exist, check for the existence of
                                            'an m_pNWSeriesOptions previous selection
182:   If Len(sFieldToFind) = 0 Then
183:     sFieldToFind = m_pNWSeriesOptions.DataFrameToUpdateGetPageNameField( _
      sDataFrameName, sLayerName)
185:   End If
  
187:   With cboOtherSeriesDF_AttributeField
                                            'if that doesn't exist, select the first
                                            'listindex
190:     If Len(sFieldToFind) = 0 Then
191:       .ListIndex = 0
      'cboOtherSeriesDF_AttributeField_ClickHandler
193:       .Text = .List(.ListIndex)
194:       sFieldToFind = .List(.ListIndex)
                                            'otherwise select the correct index for
                                            'that attribute field name
197:     Else
198:       lFieldIdx = FindControlString(cboOtherSeriesDF_AttributeField, sFieldToFind, , True)
199:       .ListIndex = lFieldIdx
200:       .Text = .List(.ListIndex)
      'cboOtherSeriesDF_AttributeField_ClickHandler
202:     End If
203:   End With


  Exit Sub
ErrorHandler:
  HandleError False, "cboOtherSeriesDF_AttributeField_ChangeHandler " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub




'  cboOtherSeriesDF_AttributeField
'    Click
'      - update the selection data structure with both the
'        current layer name and the current field name.
'        (this relies on the click event being trigger
'------------------------------------------------
Private Sub cboOtherSeriesDF_AttributeField_Click()
  On Error GoTo ErrorHandler

223:   cboOtherSeriesDF_AttributeField_ClickHandler
'  Dim sDataFrame As String, sLayerName As String, sFieldName As String
'
'  With lstOtherSeriesDF
'    sDataFrame = .List(.ListIndex)
'  End With
'  With cboOtherSeriesDF_PolygonFC
'    sLayerName = .List(.ListIndex)
'  End With
'  With cboOtherSeriesDF_AttributeField
'    sFieldName = .List(.ListIndex)
'  End With
'
'
'  With m_pDictDFsToUpdate_Layer
'    If .Exists(sDataFrame) Then
'      .Remove sDataFrame
'    End If
'    .Add sDataFrame, sLayerName
'  End With
'  With m_pDictDFsToUpdate_Field
'    If .Exists(sDataFrame) Then
'      .Remove sDataFrame
'    End If
'    .Add sDataFrame, sFieldName
'  End With


  Exit Sub
ErrorHandler:
  HandleError False, "cboOtherSeriesDF_AttributeField_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub



Private Sub cboOtherSeriesDF_AttributeField_ClickHandler()
  On Error GoTo ErrorHandler

  Dim sDataFrame As String, sLayerName As String, sFieldName As String
    
263:   With lstOtherSeriesDF
264:     sDataFrame = .List(.ListIndex)
265:   End With
266:   With cboOtherSeriesDF_PolygonFC
267:     sLayerName = .List(.ListIndex)
268:   End With
269:   With cboOtherSeriesDF_AttributeField
270:     sFieldName = .List(.ListIndex)
271:   End With
  
  
274:   With m_pDictDFsToUpdate_Layer
275:     If .Exists(sDataFrame) Then
276:       .Remove sDataFrame
277:     End If
278:     .Add sDataFrame, sLayerName
279:   End With
280:   With m_pDictDFsToUpdate_Field
281:     If .Exists(sDataFrame) Then
282:       .Remove sDataFrame
283:     End If
284:     .Add sDataFrame, sFieldName
285:   End With


  Exit Sub
ErrorHandler:
  HandleError False, "cboOtherSeriesDF_AttributeField_ClickHandler " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub


'  cboOtherSeriesDF_PolygonFC
'    Change event
'      This handles scenario of user never selecting an
'      attr field or a polygon FC, and clicking [OK]
'      - set listindex to previous selection, or to first
'        entry if previous selection doesn't exist.
'        (triggers cboOtherSeriesDF_PolygonFC click event)
'        click event will in turn populate the list of
'        attribute fields for that polygon feature class.
'----------------------------------------------
Private Sub cboOtherSeriesDF_PolygonFC_Change()
  On Error GoTo ErrorHandler

307:   cboOtherSeriesDF_PolygonFC_ChangeHandler
'  Dim sDataFrameName As String, pLayer As ILayer, pMxDoc As IMxDocument
'  Dim pMapFrame As IMapFrame, pLayers As IEnumLayer, sLayer As String
'  Dim lLyrIdx As Long, sLayerName As String
'
'
'  If m_pApp Is Nothing Then
'    Err.Raise vbObjectError + 1, "module: frmSeriesProperties" & vbNewLine _
'      & "routine: cboOtherSeriesDF_PolygonFC_Change", _
'      "An application reference was empty, so the list of fields could not" _
'      & " be updated for the selected data frame's map layers feature class." _
'      & " Try cancelling and reopening the Series Properties dialog."
'  End If
'
'  sDataFrameName = lstOtherSeriesDF.List(lstOtherSeriesDF.ListIndex)
'
'  If m_pDictDFsToUpdate_Layer Is Nothing Then
'    Set m_pDictDFsToUpdate_Layer = New Scripting.Dictionary
'    Exit Sub
'  End If
'  If m_pDictDFsToUpdate_Layer.Exists(sDataFrameName) Then
'    sLayer = m_pDictDFsToUpdate_Layer(sDataFrameName)
'    lLyrIdx = FindControlString(cboOtherSeriesDF_PolygonFC, sLayer)
'    If lLyrIdx >= 0 Then
'      cboOtherSeriesDF_PolygonFC.ListIndex = lLyrIdx 'trigger click event
'    End If
'  End If
'
'                                            'gather the parameters necessary to pass
'                                            'to the routine for populating the list
'                                            'of attribute fields of the current
'                                            'polygon feature class
'
'  'get the layer reference
'  Set pMxDoc = m_pApp.Document
'  If m_pNWSeriesOptions.DataFrameIsInStorage(sDataFrameName) Then
'    'get the data frame from storage
'    Set pMapFrame = m_pNWSeriesOptions.DataFrameStoredFrame(sDataFrameName)
'  Else
'    'get the data frame from the PageLayout
'    Set pMapFrame = GetDataFrameFromPageLayout(m_pApp, sDataFrameName)
'  End If
'
'  Set pLayers = pMapFrame.Map.Layers
'  pLayers.Reset
'  Set pLayer = pLayers.Next
'  If pLayer Is Nothing Then Exit Sub
'  sLayerName = pLayer.Name
'  Do While (Not pLayer Is Nothing) _
'       And (Not UCase(sLayerName) = UCase(cboOtherSeriesDF_PolygonFC.Text))
'    Set pLayer = pLayers.Next
'    If Not pLayer Is Nothing Then
'      sLayerName = pLayer.Name
'    End If
'  Loop
'
''  LoadCboOtherSeriesDF_Attributes pLayer, m_pNWSeriesOptions, sDataFrameName, _
'    cboOtherSeriesDF_PolygonFC.Text
  
  
  Exit Sub
ErrorHandler:
  HandleError False, "cboOtherSeriesDF_PolygonFC_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub




Private Sub cboOtherSeriesDF_PolygonFC_ChangeHandler()
  On Error GoTo ErrorHandler
  
  Dim sDataFrameName As String, pLayer As ILayer, pMxDoc As IMxDocument
  Dim pMapFrame As IMapFrame, pLayers As IEnumLayer, sLayer As String
  Dim lLyrIdx As Long, sLayerName As String
  
382:   If m_pApp Is Nothing Then
383:     Err.Raise vbObjectError + 1, "module: frmSeriesProperties" & vbNewLine _
      & "routine: cboOtherSeriesDF_PolygonFC_Change", _
      "An application reference was empty, so the list of fields could not" _
      & " be updated for the selected data frame's map layers feature class." _
      & " Try cancelling and reopening the Series Properties dialog."
388:   End If
  
390:   sDataFrameName = lstOtherSeriesDF.List(lstOtherSeriesDF.ListIndex)
  
392:   cboOtherSeriesDF_PolygonFC.Enabled = True
393:   If m_pDictDFsToUpdate_Layer Is Nothing Then
394:     Set m_pDictDFsToUpdate_Layer = New Scripting.Dictionary
    Exit Sub
396:   End If
397:   With cboOtherSeriesDF_PolygonFC
398:     If m_pDictDFsToUpdate_Layer.Exists(sDataFrameName) Then
399:       sLayer = m_pDictDFsToUpdate_Layer(sDataFrameName)
400:       lLyrIdx = FindControlString(cboOtherSeriesDF_PolygonFC, sLayer, , True)
401:       If lLyrIdx >= 0 Then
402:         .ListIndex = lLyrIdx 'trigger click event
403:       Else
404:         .ListIndex = 0
405:       End If
406:     Else
407:       .ListIndex = 0
408:     End If
    '.Text = .List(.ListIndex)
    'cboOtherSeriesDF_PolygonFC_ClickHandler
411:   End With
                                            'gather the parameters necessary to pass
                                            'to the routine for populating the list
                                            'of attribute fields of the current
                                            'polygon feature class
  
  'get the layer reference
'  Set pMxDoc = m_pApp.Document
'  If m_pNWSeriesOptions.DataFrameIsInStorage(sDataFrameName) Then
'    'get the data frame from storage
'    Set pMapFrame = m_pNWSeriesOptions.DataFrameStoredFrame(sDataFrameName)
'  Else
'    'get the data frame from the PageLayout
'    Set pMapFrame = GetDataFrameFromPageLayout(m_pApp, sDataFrameName)
'  End If
'
'  Set pLayers = pMapFrame.Map.Layers
'  pLayers.Reset
'  Set pLayer = pLayers.Next
'  If pLayer Is Nothing Then Exit Sub
'  sLayerName = pLayer.Name
'  Do While (Not pLayer Is Nothing) _
'       And (Not UCase(sLayerName) = UCase(cboOtherSeriesDF_PolygonFC.Text))
'    Set pLayer = pLayers.Next
'    If Not pLayer Is Nothing Then
'      sLayerName = pLayer.Name
'    End If
'  Loop
'  If cboOtherSeriesDF_PolygonFC.ListCount <= 0 Then
'    cboOtherSeriesDF_PolygonFC.Enabled = False
'  End If
  
'  LoadCboOtherSeriesDF_Attributes pLayer, m_pNWSeriesOptions, sDataFrameName, _
    cboOtherSeriesDF_PolygonFC.Text


  Exit Sub
ErrorHandler:
  HandleError False, "cboOtherSeriesDF_PolygonFC_ChangeHandler " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub



'  cboOtherSeriesDF_PolygonFC
'    Click event
'      - populate the attribute field list,
'        make a default selection from the attribute
'        field list in case user clicks [ok]
'        (triggers attr field change and click events)
'-------------------------------------------
Private Sub cboOtherSeriesDF_PolygonFC_Click()
  On Error GoTo ErrorHandler

464:   cboOtherSeriesDF_PolygonFC_ClickHandler
  
'  Dim sDataFrameName As String, pLayer As ILayer, pMxDoc As IMxDocument
'  Dim pMapFrame As IMapFrame, pLayers As IEnumLayer, sLayerName As String
'
'  If m_pApp Is Nothing Then
'    Err.Raise vbObjectError + 1, "module: frmSeriesProperties" & vbNewLine _
'      & "routine: cboOtherSeriesDF_PolygonFC_Change", _
'      "An application reference was empty, so the list of fields could not" _
'      & " be updated for the selected data frame's map layers feature class." _
'      & " Try cancelling and reopening the Series Properties dialog."
'  End If
'
'                                            'gather the parameters necessary to pass
'                                            'to the routine for populating the list
'                                            'of attribute fields of the current
'                                            'polygon feature class
'  sDataFrameName = lstOtherSeriesDF.List(lstOtherSeriesDF.ListIndex)
'
'  'get the layer reference
'  Set pMxDoc = m_pApp.Document
'  If m_pNWSeriesOptions.DataFrameIsInStorage(sDataFrameName) Then
'    'get the data frame from storage
'    Set pMapFrame = m_pNWSeriesOptions.DataFrameStoredFrame(sDataFrameName)
'  Else
'    'get the data frame from the PageLayout
'    Set pMapFrame = GetDataFrameFromPageLayout(m_pApp, sDataFrameName)
'  End If
'                                                  'get qualified bubble layers from
'                                                  'the data frame
'  Set pLayers = pMapFrame.Map.Layers
'  pLayers.Reset
'  Set pLayer = pLayers.Next
'  sLayerName = pLayer.Name
'  If pLayer Is Nothing Then Exit Sub
'  Do While (Not pLayer Is Nothing) _
'       And (Not UCase(sLayerName) = UCase(cboOtherSeriesDF_PolygonFC.Text))
'    Set pLayer = pLayers.Next
'    If Not pLayer Is Nothing Then
'      sLayerName = pLayer.Name
'    End If
'  Loop
'
''  LoadCboOtherSeriesDF_Attributes pLayer, m_pNWSeriesOptions, sDataFrameName, _
'    cboOtherSeriesDF_PolygonFC.Text
''''''''''''''''''''''''''''''''''''''
'
'  'load a list if attribute fields
'  'in cboOtherSeriesDF_AttributeField,
'
'  'rely on the change event for
'  'cboOtherSeriesDF_AttributeField to set
'  'the correct listindex
'
'
'  Dim pField As IField, pFields As IFields, pTable As ITable, lFieldCount As Long
'  Dim i As Long, sCurrentLayer As String, lFindResult As Long, sCurrentField As String
'  Dim sFieldName As String
'
'  With cboOtherSeriesDF_AttributeField
'    .Clear
'    If Not TypeOf pLayer Is ITable Then
'      .Enabled = False
'      Exit Sub
'    End If
'
'    Set pTable = pLayer
'    Set pFields = pTable.Fields
'    lFieldCount = pFields.FieldCount
'    For i = 0 To (lFieldCount - 1)
'      Set pField = pFields.Field(i)
'      If pField.Type = esriFieldTypeString Then
'        sFieldName = pField.Name
'        .AddItem sFieldName
'      End If
'    Next i
'
'    If .ListCount <= 0 Then
'      .Enabled = False
'    Else
'      .Enabled = True
'    End If
'
'  End With
'
'  'set the selected attribute field to a previous selection,
'  'assuming that one was previously made.  Otherwise, just
'  'leave the selected index equal to zero.
'                                          'when setting a listindex, prefer the
'                                          'local data structures
'                                          'that track previous selections over the
'                                          'ones stored in m_pNWSeriesOptions.  The
'                                          'local data structures will store more recent
'                                          'selections, and m_pNWSeriesOptions will
'                                          'store selections from a previous session.
'                                          'The local structures will be committed back
'                                          'to m_pNWSeriesOptions when the user clicks [OK]
'
'  'only grab the previous attribute field
'  'selection if the newly selected layer
'  'name matches the previous selection
'
'    '''''Put code to set the listindex within the change event of
'    '''''cboOtherSeriesDF_AttributeField instead.
'    '''''
'    '''''  sFieldName = ""
'    '''''  With m_pDictDFsToUpdate_Layer
'    '''''    If .Exists(sDataFrame) Then
'    '''''      sCurrentLayer = .Item(sDataFrame)
'    '''''      If StrComp(sCurrentLayer, sPolyFCName, vbTextCompare) = 0 Then
'    '''''        sFieldName = m_pDictDFsToUpdate_Field(sDataFrame)
'    '''''      End If
'    '''''    Else
'    '''''      'grab any settings from previous NWSeriesOptions
'    '''''      sCurrentLayer = m_pNWSeriesOptions.DataFrameToUpdateGetMapPageLayer(sDataFrame)
'    '''''      If Len(sCurrentLayer) > 0 Then
'    '''''        sFieldName = m_pNWSeriesOptions.DataFrameToUpdateGetPageNameField(sDataFrameName, sCurrentLayer)
'    '''''      End If
'    '''''    End If
'    '''''
'    '''''    If Len(sFieldName) > 0 Then
'    '''''      lFindResult = FindControlString(cboOtherSeriesDF_AttributeField, sCurrentLayer)
'    '''''      If lFindResult >= 0 Then
'    '''''        cboOtherSeriesDF_AttributeField.ListIndex = lFindResult
'    '''''      End If
'    '''''    Else
'    '''''      cboOtherSeriesDF_AttributeField.ListIndex = 0
'    '''''    End If
'    '''''  End With
'
'
'
'
'
''''''''''''''''''''''''''''''''''''''
  

  Exit Sub
ErrorHandler:
  HandleError False, "cboOtherSeriesDF_PolygonFC_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub




Private Sub cboOtherSeriesDF_PolygonFC_ClickHandler()
  On Error GoTo ErrorHandler

  Dim sDataFrameName As String, pLayer As ILayer, pMxDoc As IMxDocument
  Dim pMapFrame As IMapFrame, pLayers As IEnumLayer, sLayerName As String
  
  
616:   If m_pApp Is Nothing Then
617:     Err.Raise vbObjectError + 1, "module: frmSeriesProperties" & vbNewLine _
      & "routine: cboOtherSeriesDF_PolygonFC_Change", _
      "An application reference was empty, so the list of fields could not" _
      & " be updated for the selected data frame's map layers feature class." _
      & " Try cancelling and reopening the Series Properties dialog."
622:   End If
  
                                            'gather the parameters necessary to pass
                                            'to the routine for populating the list
                                            'of attribute fields of the current
                                            'polygon feature class
628:   sDataFrameName = lstOtherSeriesDF.List(lstOtherSeriesDF.ListIndex)
  
  'get the layer reference
631:   Set pMxDoc = m_pApp.Document
632:   If m_pNWSeriesOptions.DataFrameIsInStorage(sDataFrameName) Then
    'get the data frame from storage
634:     Set pMapFrame = m_pNWSeriesOptions.DataFrameStoredFrame(sDataFrameName)
635:   Else
    'get the data frame from the PageLayout
637:     Set pMapFrame = GetDataFrameFromPageLayout(m_pApp, sDataFrameName)
638:   End If
                                                  'get qualified bubble layers from
                                                  'the data frame
641:   Set pLayers = pMapFrame.Map.Layers
642:   pLayers.Reset
643:   Set pLayer = pLayers.Next
644:   sLayerName = pLayer.Name
  If pLayer Is Nothing Then Exit Sub
646:   Do While (Not pLayer Is Nothing) _
       And (Not UCase(sLayerName) = UCase(cboOtherSeriesDF_PolygonFC.Text))
648:     Set pLayer = pLayers.Next
649:     If Not pLayer Is Nothing Then
650:       sLayerName = pLayer.Name
651:     End If
652:   Loop
  

  'load a list if attribute fields
  'in cboOtherSeriesDF_AttributeField,
  
  'rely on the change event for
  'cboOtherSeriesDF_AttributeField to set
  'the correct listindex
  

  Dim pField As IField, pFields As IFields, pTable As ITable, lFieldCount As Long
  Dim i As Long, sCurrentLayer As String, lFindResult As Long, sCurrentField As String
  Dim sFieldName As String
  
667:   With cboOtherSeriesDF_AttributeField
668:     .Clear
669:     If Not TypeOf pLayer Is ITable Then
670:       .Enabled = False
      Exit Sub
672:     End If
  
674:     Set pTable = pLayer
675:     Set pFields = pTable.Fields
676:     lFieldCount = pFields.FieldCount
677:     For i = 0 To (lFieldCount - 1)
678:       Set pField = pFields.Field(i)
679:       If pField.Type = esriFieldTypeString Then
680:         sFieldName = pField.Name
681:         .AddItem sFieldName
682:       End If
683:     Next i
    
685:     If .ListCount <= 0 Then
686:       .Enabled = False
687:     Else
688:       .Enabled = True
689:     End If
690:     cboOtherSeriesDF_AttributeField_ChangeHandler
691:   End With


  Exit Sub
ErrorHandler:
  HandleError False, "cboOtherSeriesDF_PolygonFC_ClickHandler " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub



Private Sub cboOtherSeriesExtentOption_Change()
'    .AddItem "Use polygon extent (default)"
'    .AddItem "Use main dataframe's scale"
  
End Sub

Private Sub chkBubbleLayer_Click()
  If m_bInitializing Then Exit Sub
709:   With cboBubbleLayer
710:     If chkBubbleLayer.Value = 0 Then
711:       .Enabled = False
712:       m_pNWSeriesOptions.BubbleLayer = ""
713:     Else
714:       .Enabled = True
                                                  'assign the default value from the bubble
                                                  'layer combobox
717:       m_pNWSeriesOptions.BubbleLayer = .List(.ListIndex)
718:     End If
719:   End With
End Sub





Private Sub chkOptions_Click(Index As Integer)
  Select Case Index
  Case 0  'Rotate
729:     If chkOptions(0).Value = 0 Then
730:       cmbRotateField.Enabled = False
731:     Else
732:       cmbRotateField.Enabled = True
733:     End If
  Case 1  'Clip to outline
735:     If chkOptions(1).Value = 0 Then
736:       chkOptions(3).Value = 0
737:       chkOptions(3).Enabled = False
738:     Else
739:       chkOptions(3).Enabled = True
740:     End If
  Case 2  'Label neighboring tiles
742:     If chkOptions(2).Value = 0 Then
743:       cmdLabelProps.Enabled = False
744:       lblIndentDistance.Enabled = False
745:       txtNeighborLabelIndent.Enabled = False
746:     Else
747:       cmdLabelProps.Enabled = True
748:       lblIndentDistance.Enabled = True
749:       txtNeighborLabelIndent.Enabled = True
750:     End If
751:   End Select
End Sub







Private Sub cmdAddGroup_Click()
  On Error GoTo ErrorHandler

  Dim sGroupName As String, bBadInput As Boolean, pLyrGroup As INWLayerVisibilityGroup
  Dim lLyrCount As Long, i As Long, lIdx As Long
  
  
767:   bBadInput = True
768:   Do While bBadInput
769:     bBadInput = False
770:     sGroupName = InputBox("Please enter name of new layer visibility group.", "Layer Group Title")
                                                  'test for an empty group name
    If sGroupName = "" Then Exit Sub
                                                  'test for a duplicate group name
774:     If m_pNWSeriesOptions.LayerGroupExists(sGroupName) Then
775:       MsgBox "The layer group name ''" & sGroupName & "'' already exists." & vbNewLine _
           & "A new group layer with that same name will not be created." & vbNewLine _
           & "Please use a different name if you wish to create a new group layer.", vbOKOnly
778:       bBadInput = True
779:     End If
                                                  'limit the group name to 254 characters
                                                  '(an arbitrary number)
782:     If Len(sGroupName) > 254 Then
783:       MsgBox "Group names are limited to 254 characters in length.  The following" & vbNewLine _
           & "string was " & Len(sGroupName) & " characters long." & vbNewLine _
           & vbNewLine & sGroupName & vbNewLine _
           & "Please re-enter the group name.", vbOKOnly
787:       bBadInput = True
788:     End If
789:   Loop
                                                  'create a new group object
791:   Set pLyrGroup = New NWLayerVisibilityGroup
792:   m_pNWSeriesOptions.LayerGroupSet sGroupName, pLyrGroup
                                                  'set all layers to be visible in
                                                  'the interface
795:   lLyrCount = lstVisibleLayers.ListCount
796:   For i = 0 To lLyrCount - 1
797:     lstVisibleLayers.ItemData(i) = vbChecked
798:   Next i
                                                  'add the listing of that group
                                                  'name to the group listbox
801:   lstLyrGroups.AddItem sGroupName
                                                  'select that newly added group name
803:   lIdx = FindControlString(lstLyrGroups, sGroupName)
804:   lstLyrGroups.ListIndex = lIdx
  
  Exit Sub
ErrorHandler:
  HandleError True, "cmdAddGroup_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdCancel_Click()
812:   Unload Me
End Sub

Private Sub cmdInsertPageName_Click()
  On Error GoTo ErrorHandler

  Dim sReplaceString As String, sSource As String, sTemp As String
  
820:   sReplaceString = m_pNWSeriesOptions.DynamicDefQueryReplaceString
821:   If sReplaceString = "" Then
822:     lblNoPageNameWarning.Visible = True
823:     cmdInsertPageName.Enabled = False
824:     cmdPrevPageName.Enabled = False
825:     cmdNextPageName.Enabled = False
    Exit Sub
827:   Else
828:     lblNoPageNameWarning.Visible = False
    'no point in setting .Enabled = true
    'since they had to be enabled for this code to run
831:   End If
832:   With txtDefinitionQuery
833:     sSource = .Text
834:     If .SelLength < 1 Then
835:       sTemp = Left$(sSource, .SelStart)
836:       sTemp = sTemp & sReplaceString & Right$(sSource, (Len(sSource) - .SelStart))
837:       .Text = sTemp
838:     Else
839:       .SelText = sReplaceString
840:     End If
841:   End With


  Exit Sub
ErrorHandler:
  HandleError True, "cmdInsertPageName_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdLabelProps_Click()
  On Error GoTo ErrorHandler

852:   If m_pNWSeriesOptions Is Nothing Then
    Exit Sub
854:   End If
855:   Set frmAdjMapLabelSymbols.NWSeriesOptions = m_pNWSeriesOptions
856:   Set frmAdjMapLabelSymbols.Application = m_pApp
857:   frmAdjMapLabelSymbols.Show vbModal, frmSeriesProperties
858:   If Not frmAdjMapLabelSymbols.NWSeriesOptions Is Nothing Then
859:     Set m_pNWSeriesOptions = frmAdjMapLabelSymbols.NWSeriesOptions
860:   End If
861:   Set m_pTextSym = frmAdjMapLabelSymbols.TextSymbol
862:   Me.SetFocus

  Exit Sub
ErrorHandler:
  HandleError True, "cmdLabelProps_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

'routine highlights the next entry in the definition query
'that will be replaced on map page changes
Private Sub cmdNextPageName_Click()
  On Error GoTo ErrorHandler

  Dim sReplaceString As String, lNewStart As Long
  
876:   sReplaceString = m_pNWSeriesOptions.DynamicDefQueryReplaceString
877:   If sReplaceString = "" Then
878:     lblNoPageNameWarning.Visible = True
879:     cmdInsertPageName.Enabled = False
880:     cmdPrevPageName.Enabled = False
881:     cmdNextPageName.Enabled = False
    Exit Sub
883:   Else
884:     lblNoPageNameWarning.Visible = False
    'no point in setting .Enabled = true
    'since they had to be enabled for this code to run
887:   End If
  
889:   With txtDefinitionQuery
890:     If Len(sReplaceString) > Len(.Text) Then
891:       .SelLength = 0
      Exit Sub
893:     End If
894:     If .SelStart = 0 Then
895:       If StrComp(.SelText, sReplaceString, vbTextCompare) = 0 Then
896:         lNewStart = (InStr(Len(sReplaceString), _
                           .Text, _
                           sReplaceString, _
                           vbTextCompare) - 1)
900:         If lNewStart < 0 Then
901:           .SelStart = Len(.Text)
902:         Else
903:           .SelStart = lNewStart
904:         End If
905:       Else
906:         lNewStart = (InStr(1, _
                           txtDefinitionQuery.Text, _
                           sReplaceString, _
                           vbTextCompare) - 1)
910:         If lNewStart > -1 Then
911:           If lNewStart = 0 Then
912:             .SelStart = 0
913:           Else
914:             .SelStart = lNewStart
915:           End If
916:         End If
917:       End If
918:     Else
      'if the currently selected text is the search
      'string, then start searching at the end of the
      'selected text, else start searching at the
      'beginning of the selected text.
923:       If StrComp(.SelText, sReplaceString, vbTextCompare) = 0 Then
924:         lNewStart = InStr((.SelStart + .SelLength), _
                          txtDefinitionQuery.Text, _
                          sReplaceString, _
                          vbTextCompare) - 1
928:       Else
929:         lNewStart = InStr(.SelStart, _
                          .Text, _
                          sReplaceString, _
                          vbTextCompare) - 1
933:       End If
934:       If (lNewStart >= Len(.Text)) Or (lNewStart < 0) Then
935:         lNewStart = Len(.Text)
936:       End If
937:       .SelStart = lNewStart
938:     End If
939:     .SelLength = Len(sReplaceString)
940:   End With
  

  Exit Sub
ErrorHandler:
  HandleError True, "cmdNextPageName_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdOK_Click()
  On Error GoTo ErrorHandler

  
  Dim pDoc As IMxDocument, pActive As IActiveView, dMapRatio As Double
  
  'Apply updates (only the Options can be updated, so we only need to look at those)
  'Set the clip and rotate properties
  'Update 6/18/03 to support cross hatching of clip area
957:   If chkOptions(1).Value = 1 Then    'Clip
958:     If chkOptions(3).Value = 0 Then   'clip without cross hatch
      'Make sure we don't leave the clip element
960:       If m_pSeriesOptions2.ClipData = 2 Then RemoveClipElement m_pApp.Document
961:       m_pSeriesOptions2.ClipData = 1
962:     Else
963:       m_pSeriesOptions2.ClipData = 2
964:       Set pDoc = m_pApp.Document
965:       pDoc.FocusMap.ClipGeometry = Nothing
966:     End If
'    m_pSeriesOptions.ClipData = True
968:   Else
    'Make sure we don't leave the clip element
970:     If m_pSeriesOptions2.ClipData = 2 Then RemoveClipElement m_pApp.Document
971:     m_pSeriesOptions2.ClipData = 0
'    m_pSeriesOptions.ClipData = False
    'Make sure clipping is turned off for the data frame
974:     Set pDoc = m_pApp.Document
975:     pDoc.FocusMap.ClipGeometry = Nothing
976:   End If
  
978:   If chkOptions(0).Value = 1 Then     'Rotation
979:     If m_pSeriesOptions.RotateFrame = False Or m_pSeriesOptions.RotationField <> cmbRotateField.Text Then
980:       UpdatePageValues "ROTATION", cmbRotateField.Text
981:     End If
982:     m_pSeriesOptions.RotateFrame = True
983:     m_pSeriesOptions.RotationField = cmbRotateField.Text
984:   Else
985:     m_pSeriesOptions.RotateFrame = False
    'Make sure rotation is turned off for the data frame
987:     Set pDoc = m_pApp.Document
988:     Set pActive = pDoc.FocusMap
989:     If pActive.ScreenDisplay.DisplayTransformation.Rotation <> 0 Then
990:       pActive.ScreenDisplay.DisplayTransformation.Rotation = 0
991:       pActive.Refresh
992:     End If
993:   End If
994:   If chkOptions(2).Value = 1 Then    'Label Neighbors
995:     m_pSeriesOptions.LabelNeighbors = True
                                      'indent the neighbor labels
997:     m_pNWSeriesOptions.NeighborLabelIndent = CDbl(txtNeighborLabelIndent.Text)
998:   Else
999:     m_pSeriesOptions.LabelNeighbors = False
1000:     m_pNWSeriesOptions.NeighborLabelIndent = 0
1001:     RemoveLabels pDoc
1002:     g_bLabelNeighbors = False
1003:   End If
1004:   Set m_pSeriesOptions.LabelSymbol = m_pTextSym
  
  'Set the extent properties
1007:   If optExtent(0).Value Then         'Variable
1008:     m_pSeriesOptions.ExtentType = 0
1009:     If txtMargin.Text = "" Then
1010:       m_pSeriesOptions.Margin = 0
1011:     Else
1012:       m_pSeriesOptions.Margin = CDbl(txtMargin.Text)
1013:     End If
1014:     m_pSeriesOptions.MarginType = cmbMargin.ListIndex
1015:   ElseIf optExtent(1).Value Then    'Fixed
1016:     m_pSeriesOptions.ExtentType = 1
1017:     m_pSeriesOptions.FixedScale = txtFixed.Text
1018:   Else                        'Data driven
1019:     If m_pSeriesOptions.ExtentType <> 2 Or m_pSeriesOptions.RotationField <> cmbRotateField.Text Then
1020:       UpdatePageValues "SCALE", cmbDataDriven.Text
1021:     End If
1022:     m_pSeriesOptions.ExtentType = 2
1023:     m_pSeriesOptions.DataDrivenField = cmbDataDriven.Text
1024:   End If
  
  
  
1028:   If Not m_pNWSeriesOptions Is Nothing Then
1029:     If chkRefreshPage.Value = 1 Then
1030:       m_pNWSeriesOptions.RefreshEventLoadPage = True
1031:     Else
1032:       m_pNWSeriesOptions.RefreshEventLoadPage = False
1033:     End If
1034:   End If


  'mark other data frames that will have their
  'extent updated when map pages are loaded.
  Dim i As Integer, sOtherDF As String, vDataFrames As Variant, lDFCount As Long
  Dim sField As String, sLayer As String
1041:   m_pNWSeriesOptions.DataFrameToUpdateClearAllDataFrames
1042:   vDataFrames = m_pDictDFsToUpdate_Field.Keys
1043:   lDFCount = UBound(vDataFrames) + 1
1044:   For i = 0 To (lDFCount - 1)
1045:     sOtherDF = vDataFrames(i)
1046:     sLayer = m_pDictDFsToUpdate_Layer(sOtherDF)
1047:     sField = m_pDictDFsToUpdate_Field(sOtherDF)
1048:     m_pNWSeriesOptions.DataFrameToUpdateAdd sOtherDF, sLayer, sField
1049:   Next i
1050:   With cboOtherSeriesExtentOption
1051:     m_pNWSeriesOptions.DataFrameToUpdateExtentOption = .List(.ListIndex)
1052:   End With
  
1058:   Unload Me
  

  Exit Sub
ErrorHandler:
  HandleError True, "cmdOK_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub UpdatePageValues(sProperty As String, sFieldName As String)
On Error GoTo ErrHand:
  Dim lLoop As Long, pSeries As INWDSMapSeries, pPage As INWDSMapPage
  Dim pDoc As IMxDocument, pMap As IMap, pSeriesProps As INWDSMapSeriesProps
  Dim pIndexLayer As IFeatureLayer, pDataset As IDataset, pWorkspace As IFeatureWorkspace
  Dim pQueryDef As IQueryDef, pCursor As ICursor, pRow As IRow, pColl As Collection
1072:   Set pDoc = m_pApp.Document
1073:   Set pSeries = m_pSeriesOptions
1074:   Set pSeriesProps = pSeries
1075:   Set pMap = FindDataFrame(pDoc, pSeriesProps.DataFrameName)
  If pMap Is Nothing Then Exit Sub
  
1078:   Set pIndexLayer = FindLayer(pSeriesProps.IndexLayerName, pMap)
  If pIndexLayer Is Nothing Then Exit Sub
  
  'Loop through the features in the index layer creating a collection of the scales and tile names
1082:   Set pDataset = pIndexLayer.FeatureClass
1083:   Set pWorkspace = pDataset.Workspace
1084:   Set pQueryDef = pWorkspace.CreateQueryDef
1085:   pQueryDef.Tables = pDataset.Name
1086:   pQueryDef.SubFields = sFieldName & "," & pSeriesProps.IndexFieldName
1087:   Set pCursor = pQueryDef.Evaluate
1088:   Set pColl = New Collection
1089:   Set pRow = pCursor.NextRow
1090:   Do While Not pRow Is Nothing
1091:     If Not IsNull(pRow.Value(0)) And Not IsNull(pRow.Value(1)) Then
1092:       pColl.Add pRow.Value(0), pRow.Value(1)
1093:     End If
1094:     Set pRow = pCursor.NextRow
1095:   Loop
  
  'Now loop through the pages and try to find the corresponding tile name in the collection
  On Error GoTo ErrNoKey:
1099:   For lLoop = 0 To pSeries.PageCount - 1
1100:     Set pPage = pSeries.Page(lLoop)
1101:     If sProperty = "ROTATION" Then
1102:       pPage.PageRotation = pColl.Item(pPage.PageName)
1103:     Else
1104:       pPage.PageScale = pColl.Item(pPage.PageName)
1105:     End If
1106:   Next lLoop

  Exit Sub

ErrNoKey:
1111:   Resume Next
ErrHand:
1113:   MsgBox "UpdatePageValues - " & Err.Description
End Sub


'
''Function takes a data frame name and a layer name.  Function returns the
''reference to that layer object.  This function assumes that the named data
''frame and layer exist in the map document.
''----------------------------------
'Public Function LayerFromDataFrame(sDataFrame As String, sLayer As String, pMxDoc As IMxDocument) As ILayer
'  On Error GoTo ErrorHandler
'
'  Dim pLayers As IEnumLayer, pGraphicsContainer As IGraphicsContainer, pMap As IMap
'  Dim pMapFrame As IMapFrame, pPageLayout As IPageLayout, bFoundFrame As Boolean
'  Dim bFoundLayer As Boolean, pLayer As ILayer, pElement As IElement
'
'  If pMxDoc Is Nothing Then
'    Set LayerFromDataFrame = Nothing
'    Exit Function
'  End If
'
'  Set pPageLayout = pMxDoc.PageLayout
'  Set pGraphicsContainer = pPageLayout
'  pGraphicsContainer.Reset
'  Set pElement = pGraphicsContainer.Next
'  bFoundFrame = False
'  Do While (Not pElement Is Nothing) And Not bFoundFrame
'    If TypeOf pElement Is IMapFrame Then
'      Set pMapFrame = pElement
'      Set pMap = pMapFrame.Map
'      If StrComp(pMap.Name, sDataFrame, vbTextCompare) = 0 Then
'        bFoundFrame = True
'      End If
'    End If
'    Set pElement = pGraphicsContainer.Next
'  Loop
'
'  If Not bFoundFrame Then
'    Set LayerFromDataFrame = Nothing
'    Exit Function
'  End If
'
'  'data frame is found by this point
'  Set pLayers = pMap.Layers
'  Set pLayer = pLayers.Next
'  bFoundLayer = False
'  Do While (Not pLayer Is Nothing) And Not bFoundLayer
'    If StrComp(sLayer, pLayer.Name, vbTextCompare) = 0 Then
'      bFoundLayer = True
'    End If
'    Set pLayer = pLayers.Next
'  Loop
'
'  If bFoundLayer Then
'    Set LayerFromDataFrame = pLayer
'  Else
'    Set LayerFromDataFrame = Nothing
'  End If
'
'
'  Exit Function
'ErrorHandler:
'  HandleError False, "LayerFromDataFrame " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
'End Function

Private Sub cmdPrevPageName_Click()
  On Error GoTo ErrorHandler

  Dim sReplaceString As String, lNewStart As Long, sTemp As String
  Dim lReplaceStrLen As Long
  
1184:   sReplaceString = m_pNWSeriesOptions.DynamicDefQueryReplaceString
1185:   If sReplaceString = "" Then
1186:     lblNoPageNameWarning.Visible = True
1187:     cmdInsertPageName.Enabled = False
1188:     cmdPrevPageName.Enabled = False
1189:     cmdNextPageName.Enabled = False
    Exit Sub
1191:   Else
1192:     lblNoPageNameWarning.Visible = False
1193:     lReplaceStrLen = Len(sReplaceString)
    'no point in setting .Enabled = true
    'since they had to be enabled for this code to run
1196:   End If

1198:   With txtDefinitionQuery
1199:     If (InStr(1, .Text, sReplaceString, vbTextCompare) = 0) Then
1200:       MsgBox "The text ''" & sReplaceString & "'' was not found." & vbNewLine
      Exit Sub
1202:     End If
1203:     If lReplaceStrLen > Len(.Text) Then
1204:       .SelLength = 0
      Exit Sub
1206:     End If
    
1208:     If .SelStart = 0 Then
1209:       .SelLength = 0
1210:     Else
1211:       lNewStart = InStrRev(.Text, _
                           sReplaceString, _
                           .SelStart, _
                           vbTextCompare)
      '''''''if the searched text couldn't be found ...
1216:       If lNewStart = 0 Or lNewStart = -1 Then
1217:         .SelStart = 0
1218:         If StrComp(Left$(txtDefinitionQuery.Text, lReplaceStrLen), sReplaceString, vbTextCompare) = 0 Then
1219:           .SelLength = lReplaceStrLen
1220:         Else
1221:           .SelLength = 0
1222:         End If
      '''''''else the searched text was found
1224:       Else
1225:         sTemp = Left$(txtDefinitionQuery.Text, lNewStart + lReplaceStrLen - 1)
1226:         If StrComp(Right$(sTemp, lReplaceStrLen), sReplaceString, vbTextCompare) = 0 Then
1227:           .SelStart = lNewStart - 1
1228:           .SelLength = lReplaceStrLen
1229:         End If
1230:       End If
1231:     End If
1232:   End With


  Exit Sub
ErrorHandler:
  HandleError True, "cmdPrevPageName_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdRemoveGroup_Click()
  On Error GoTo ErrorHandler

  Dim lIdx As Long, pMxDoc As IMxDocument, i As Long, lLyrCount As Long
  Dim pLayer As ILayer, pLayers As IEnumLayer
  
1246:   With lstLyrGroups
1247:     If .ListIndex = -1 Then
1248:       If .ListCount > 0 Then
1249:         .ListIndex = 0
1250:       Else
        Exit Sub
1252:       End If
1253:     End If
    
    If .ListCount = 0 Then Exit Sub
    If .List(.ListIndex) = "" Then Exit Sub
    If m_pNWSeriesOptions Is Nothing Then Exit Sub
1258:     If Not m_pNWSeriesOptions.LayerGroupExists(.List(.ListIndex)) Then
1259:       .RemoveItem (.ListIndex)
      Exit Sub
1261:     End If
                                                  'remove the layer group
1263:     m_pNWSeriesOptions.LayerGroupSet .List(.ListIndex), Nothing
1264:     .RemoveItem (.ListIndex)
1265:     If .ListCount > 0 Then
1266:       .ListIndex = 0
1267:     Else
                                                  'if the last visible layers group
                                                  'was deleted, set the visible layers
                                                  'back to ArcMap's current settings
      If m_pApp Is Nothing Then Exit Sub
1272:       Set pMxDoc = m_pApp.Document
1273:       Set pLayers = pMxDoc.FocusMap.Layers
1274:       Set pLayer = pLayers.Next
1275:       Do While Not pLayer Is Nothing
1276:         lIdx = FindControlString(lstVisibleLayers, pLayer.Name)
1277:         If pLayer.Visible Then
1278:           lstVisibleLayers.ItemData(lIdx) = vbChecked
1279:         Else
1280:           lstVisibleLayers.ItemData(lIdx) = vbUnchecked
1281:         End If
1282:         Set pLayer = pLayers.Next
1283:       Loop
1284:     End If
1285:   End With

  Exit Sub
ErrorHandler:
  HandleError True, "cmdRemoveGroup_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub CleanOrphanedDynamicDefQueryDataFrames()
  On Error GoTo ErrorHandler

  Dim vManagedFrames As Variant, lManagedFramesCount As Long, sFrameName As String
  Dim i As Long
  
1298:   If m_pNWSeriesOptions Is Nothing Then
    Exit Sub
1300:   End If
  'for each data frame data structure,
    'does that structure have an existing data frame?
      'if not, then clean up that orphaned data structure
1304:   vManagedFrames = m_pNWSeriesOptions.DynamicDefQueryDataFrames
1305:   lManagedFramesCount = UBound(vManagedFrames) + 1
1306:   For i = 0 To (lManagedFramesCount - 1)
1307:     sFrameName = vManagedFrames(i)
1308:     If FindControlString(Me.cboDataFrameToFilter, sFrameName, -1, True) = -1 Then
1309:       m_pNWSeriesOptions.DynamicDefQueryRemoveDataFrame sFrameName
1310:     End If
1311:   Next i

  Exit Sub
ErrorHandler:
  HandleError False, "CleanOrphanedDynamicDefQueryDataFrames " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub


Private Sub Form_Activate()
  On Error GoTo ErrorHandler

1322:   If m_pNWSeriesOptions Is Nothing Then
    Exit Sub
1324:   End If
  If m_pApp Is Nothing Then Exit Sub
1326:   lblNoPageNameWarning.Visible = False
1327:   If m_pDictDFsToUpdate_Layer Is Nothing Then Set m_pDictDFsToUpdate_Layer = New Scripting.Dictionary
1328:   If m_pDictDFsToUpdate_Field Is Nothing Then Set m_pDictDFsToUpdate_Field = New Scripting.Dictionary
  
  'bubble inset layer can be set in two
  'places: in the UI of this form, or by
  'loading a previous setting from a map
  'document.  This code covers the
  'possibility that the UI must be updated
  'to reflect the result of a loaded MxDoc
  'setting.
  
  Dim lBubbleIdx As Long, pMxDoc As IMxDocument, pLayers As IEnumLayer
  Dim pLayer As ILayer, sBubbleLyrName As String, bBubbleLyrExists As Boolean
  Dim bTemp As Boolean
  
1342:   sBubbleLyrName = m_pNWSeriesOptions.BubbleLayer
1343:   If sBubbleLyrName <> "" Then
1344:     lBubbleIdx = FindControlString(cboBubbleLayer, sBubbleLyrName)
1345:     If lBubbleIdx = -1 Then
      'if layer doesn't exist,
1347:       Set pMxDoc = m_pApp.Document
1348:       Set pLayers = pMxDoc.FocusMap.Layers
      
1350:       bBubbleLyrExists = False
1351:       pLayers.Reset
1352:       Set pLayer = pLayers.Next
1353:       Do While Not pLayer Is Nothing
1354:         If pLayer.Name = sBubbleLyrName Then
1355:           bBubbleLyrExists = True
1356:         End If
1357:         Set pLayer = pLayers.Next
1358:       Loop
        
1360:       If Not bBubbleLyrExists Then
1361:         m_pNWSeriesOptions.BubbleLayer = ""
        Exit Sub
1363:       End If
      
      'if layer exists,
1366:       bTemp = m_bInitializing
1367:       m_bInitializing = True
1368:       cboBubbleLayer.AddItem sBubbleLyrName
1369:       m_bInitializing = bTemp
1370:       lBubbleIdx = FindControlString(cboBubbleLayer, sBubbleLyrName)
1371:     End If
1372:     bTemp = m_bInitializing
1373:     m_bInitializing = True
1374:     cboBubbleLayer.ListIndex = lBubbleIdx
1375:     chkBubbleLayer.Value = 1
1376:     m_bInitializing = bTemp
1377:   End If

  'reflect whether or not map pages are reloaded on
  'map refresh events
1381:   If m_pNWSeriesOptions.RefreshEventLoadPage Then
1382:     chkRefreshPage.Value = 1
1383:   Else
1384:     chkRefreshPage.Value = 0
1385:   End If
  
  
  Exit Sub
ErrorHandler:
  HandleError True, "Form_Activate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Form_Load()
On Error GoTo ErrHand:
  Dim pMapBook As INWDSMapBook
  Dim pSeriesProps As INWDSMapSeriesProps
  Dim lLoop As Long
  Dim pMxDoc As IMxDocument, pLayers As IEnumLayer, pLayer As ILayer, lIdx As Long
  Dim vGroupNames() As Variant, lGroupCount As Long, i As Long, vLyrNames() As Variant
  Dim pLyrGroup As INWLayerVisibilityGroup, pFeatClass As IFeatureClass
  Dim pFeatLayer As IFeatureLayer, sDFUpdateOption As String, lDFUpdateOption As Long
  Dim pMap As IMap
  
  
  'Check to see if a MapSeries already exists
1406:   Set m_pMapBook = GetMapBookExtension(m_pApp)
  'Set pMapBook = GetMapBookExtension(m_pApp)
1408:   Set pMapBook = m_pMapBook
  If pMapBook Is Nothing Then Exit Sub
  
1411:   Set pSeriesProps = pMapBook.ContentItem(0)
1412:   Set m_pSeriesOptions = pSeriesProps
1413:   Set m_pSeriesOptions2 = m_pSeriesOptions
1414:   Set m_pNWSeriesOptions = m_pSeriesOptions2
  
  'Index Settings Tab
1417:   cmbDetailFrame.Clear
1418:   cmbDetailFrame.AddItem pSeriesProps.DataFrameName
1419:   cmbDetailFrame.Text = pSeriesProps.DataFrameName
1420:   cmbIndexLayer.Clear
1421:   cmbIndexLayer.AddItem pSeriesProps.IndexLayerName
1422:   cmbIndexLayer.Text = pSeriesProps.IndexLayerName
1423:   cmbIndexField.Clear
1424:   cmbIndexField.AddItem pSeriesProps.IndexFieldName
1425:   cmbIndexField.Text = pSeriesProps.IndexFieldName
  
  'Tile Settings Tab
1428:   optTiles(pSeriesProps.TileSelectionMethod) = True
1429:   lstSuppressTiles.Clear
1430:   If pSeriesProps.SuppressLayers Then
1431:     chkSuppress.Value = 1
1432:     For lLoop = 0 To pSeriesProps.SuppressLayerCount - 1
1433:       lstSuppressTiles.AddItem pSeriesProps.SuppressLayer(lLoop)
1434:       lstSuppressTiles.Selected(lLoop) = True
1435:     Next lLoop
1436:   Else
1437:     chkSuppress.Value = 0
1438:   End If
1439:   txtNumbering.Text = CStr(pSeriesProps.StartNumber)  'Added 2/18/2004
  
  'Options tab
1442:   PopulateFieldCombos
1443:   cmbMargin.Clear
1444:   cmbMargin.AddItem "percent"
1445:   cmbMargin.AddItem "mapunits"
1446:   cmbMargin.Text = "percent"
1447:   optExtent(m_pSeriesOptions.ExtentType).Value = True
1448:   cmdOK.Enabled = True
  Select Case m_pSeriesOptions.ExtentType
  Case 0
1451:     txtMargin.Text = m_pSeriesOptions.Margin
1452:     If m_pSeriesOptions.MarginType = 0 Then
1453:       cmbMargin.Text = "percent"
1454:     Else
1455:       cmbMargin.Text = "mapunits"
1456:     End If
  Case 1
1458:     txtFixed.Text = m_pSeriesOptions.FixedScale
  Case 2
1460:     cmbDataDriven.Text = m_pSeriesOptions.DataDrivenField
1461:   End Select
1462:   If m_pSeriesOptions.RotateFrame Then
1463:     chkOptions(0).Value = 1
1464:     cmbRotateField.Text = m_pSeriesOptions.RotationField
1465:   Else
1466:     chkOptions(0).Value = 0
1467:   End If
  
1469:   txtNeighborLabelIndent.Text = m_pNWSeriesOptions.NeighborLabelIndent
  
  'Update 6/18/03 to support cross hatching of clip area
  Select Case m_pSeriesOptions2.ClipData
  Case 0   'No clipping
1474:     chkOptions(1).Value = 0
1475:     chkOptions(3).Value = 0
1476:     chkOptions(3).Enabled = False
  Case 1   'Clip only
1478:     chkOptions(1).Value = 1
1479:     chkOptions(3).Value = 0
1480:     chkOptions(3).Enabled = True
  Case 2   'Clip with cross hatch outside clip area
1482:     chkOptions(1).Value = 1
1483:     chkOptions(3).Value = 1
1484:     chkOptions(3).Enabled = True
1485:   End Select
'  If m_pSeriesOptions.ClipData Then
'    chkOptions(1).Value = 1
'  Else
'    chkOptions(1).Value = 0
'  End If

1492:   If m_pSeriesOptions.LabelNeighbors Then
1493:     chkOptions(2).Value = 1
1494:     cmdLabelProps.Enabled = True
1495:   Else
1496:     chkOptions(2).Value = 0
1497:     cmdLabelProps.Enabled = False
1498:   End If
1499:   Set m_pTextSym = m_pSeriesOptions.LabelSymbol
  
  'Layer Groups tab
1502:   lstLyrGroups.Clear
1503:   lstVisibleLayers.Clear
                                                  'access available layers, load
                                                  'into lstVisibleLayers
1506:   Set pMxDoc = m_pApp.Document
1507:   Set pMap = FindDataFrame(pMxDoc, m_pNWSeriesOptions.DataFrameMainFrame)
1508:   Set pLayers = pMap.Layers
1509:   Set pLayer = pLayers.Next
1510:   Do While Not pLayer Is Nothing
1511:     lstVisibleLayers.AddItem pLayer.Name
1512:     lIdx = FindControlString(lstVisibleLayers, pLayer.Name)
1513:     lstVisibleLayers.ItemData(lIdx) = pLayer.Visible
1514:     Set pLayer = pLayers.Next
1515:   Loop
                                                  'access layer groups, load into
                                                  'lstLyrGroups
1518:   lGroupCount = m_pNWSeriesOptions.LayerGroupCount
1519:   vGroupNames = m_pNWSeriesOptions.LayerGroups
1520:   For i = 0 To (lGroupCount - 1)
1521:     lstLyrGroups.AddItem (vGroupNames(i))
1522:   Next i
  
1524:   If lGroupCount > 0 Then
1525:     m_sSelectedGroup = lstLyrGroups.List(0)
1526:     Set pLyrGroup = m_pNWSeriesOptions.LayerGroupGet(m_sSelectedGroup)
1527:     For i = 0 To lstVisibleLayers.ListCount - 1
1528:       lstVisibleLayers.Selected(i) = Not (pLyrGroup.Exists(lstVisibleLayers.List(i)))
1529:     Next i
1530:     vLyrNames = pLyrGroup.InvisibleLayers
1531:   Else
1532:     m_sSelectedGroup = ""
1533:   End If
  
  
  'Layer Filters tab
1537:   LoadDataFrames
1538:   m_sCurrentDataFrame = cboDataFrameToFilter.List(cboDataFrameToFilter.ListIndex)
1539:   LoadLayersUIFromDataFrame m_sCurrentDataFrame
1540:   CleanOrphanedDynamicDefQueryDataFrames
1541:   lblNoPageNameWarning.Visible = False
  
  
  'NW Options tab
1545:   m_bInitializing = True
1546:   Set pLayers = pMxDoc.FocusMap.Layers
                                                  'sift through the list of loaded feature
                                                  'classes
1549:   Set pLayer = pLayers.Next
1550:   cboBubbleLayer.Clear
1551:   Do While Not pLayer Is Nothing
1552:     If TypeOf pLayer Is IFeatureLayer Then
1553:       Set pFeatLayer = pLayer
1554:       If m_pNWSeriesOptions.IsBubbleLayer(pFeatLayer.FeatureClass) Then
                                                  'add to the bubble def layers combobox
                                                  'all layers that match the requirements
                                                  'for a bubble definitions layer
1558:         cboBubbleLayer.AddItem pLayer.Name
1559:       End If
1560:     End If
1561:     Set pLayer = pLayers.Next
1562:   Loop
                                                  'disable the chkbox option to access this
                                                  'layer if none were detected, and display
                                                  'a message to the user that this was the case
                                                  'Setup the bubble layers UI section
1567: If cboBubbleLayer.ListCount > 0 Then
1568:     If m_pNWSeriesOptions.BubbleLayer <> "" Then
1569:       lIdx = FindControlString(cboBubbleLayer, m_pNWSeriesOptions.BubbleLayer)
1570:       If lIdx = -1 Then
1571:         lIdx = 0
1572:       Else
1573:         chkBubbleLayer.Value = 1
1574:       End If
1575:       cboBubbleLayer.ListIndex = lIdx
1576:     Else
1577:       cboBubbleLayer.ListIndex = 0
1578:     End If
1579:     lblBubbleLayerWarning.Visible = False
1580:     chkBubbleLayer.Enabled = True
1581:     lblBubbleLayer.Enabled = True
1582:     If chkBubbleLayer.Value = 1 Then
1583:       cboBubbleLayer.Enabled = True
1584:     Else
1585:       cboBubbleLayer.Enabled = False
1586:     End If
1587:   Else
1588:     chkBubbleLayer.Value = 0
1589:     lblBubbleLayerWarning.Visible = True
1590:     chkBubbleLayer.Enabled = False
1591:     lblBubbleLayer.Enabled = False
1592:     cboBubbleLayer.Enabled = False
1593:   End If
                                                  'populate the list of other data frames
                                                  'in ArcMap, minus those that are detail insets
1596:   lstOtherSeriesDF.Clear
1597:   cboOtherSeriesDF_AttributeField.Clear
1598:   cboOtherSeriesDF_PolygonFC.Clear
  
  Dim vStoredFrames As Variant, lStoredFramesCount As Long, pMapFrame As IMapFrame
  Dim pGraphicsContainer As IGraphicsContainer, pElement As IElement
  Dim lFindResult As Long, sDataFrame As String
  
  
1605:   If m_pDictDFsToUpdate_Layer Is Nothing Then Set m_pDictDFsToUpdate_Layer = New Scripting.Dictionary
1606:   If m_pDictDFsToUpdate_Field Is Nothing Then Set m_pDictDFsToUpdate_Field = New Scripting.Dictionary
  
1608:   Set pGraphicsContainer = pMxDoc.PageLayout
1609:   pGraphicsContainer.Reset
1610:   Set pElement = pGraphicsContainer.Next
  
1612:   Do While Not pElement Is Nothing
1613:     If TypeOf pElement Is IMapFrame Then
1614:       Set pMapFrame = pElement
                                                  'supposed to just list "other" data frames, not the
                                                  'main data frame containing the map series
1617:       If StrComp(pMapFrame.Map.Name, m_pNWSeriesOptions.DataFrameMainFrame) <> 0 Then
                                                  'ignore detail insets too
1619:         If StrComp(Left$(pMapFrame.Map.Name, Len("BubbleID:")), "BubbleID:", vbTextCompare) <> 0 Then
1620:           lFindResult = FindControlString(lstOtherSeriesDF, pMapFrame.Map.Name, 0, True)
1621:           If lFindResult = -1 Then
                                                  'add the data frame to the list
1623:             Me.lstOtherSeriesDF.AddItem pMapFrame.Map.Name
1624:             If m_pNWSeriesOptions.DataFrameToUpdateIsADataFrameToUpdate(pMapFrame.Map.Name) Then
                                                  'set the selected properties of the
                                                  'newly added lstOtherSeriesDF
1627:               lFindResult = FindControlString(lstOtherSeriesDF, pMapFrame.Map.Name, 0, True)
1628:               lstOtherSeriesDF.Selected(lFindResult) = True   'may trigger click event code
1629:             End If
1630:           End If
1631:         End If
1632:       End If
1633:     End If
1634:     Set pElement = pGraphicsContainer.Next
1635:   Loop
                                                  'also list data frames that are not currently visible
1637:   vStoredFrames = m_pNWSeriesOptions.DataFramesStored
1638:   lStoredFramesCount = UBound(vStoredFrames) + 1
1639:   For i = 0 To (lStoredFramesCount - 1)
1640:     sDataFrame = vStoredFrames(i)
1641:     lFindResult = FindControlString(lstOtherSeriesDF, sDataFrame, 0, True)
1642:     If lFindResult = -1 Then
1643:       lstOtherSeriesDF.AddItem sDataFrame
1644:       If m_pNWSeriesOptions.DataFrameToUpdateIsADataFrameToUpdate(pMapFrame.Map.Name) Then
                                                  'set the selected properties of the
                                                  'newly added lstOtherSeriesDF
1647:         lFindResult = FindControlString(lstOtherSeriesDF, pMapFrame.Map.Name, 0, True)
1648:         lstOtherSeriesDF.Selected(lFindResult) = True   'may trigger click event code
1649:       End If
1650:     End If
1651:   Next i
1652:   lstOtherSeriesDF.Enabled = (lstOtherSeriesDF.ListCount > 0)
1653:   cboOtherSeriesDF_PolygonFC.Enabled = (cboOtherSeriesDF_PolygonFC.ListCount > 0)
1654:   cboOtherSeriesDF_AttributeField.Enabled = (cboOtherSeriesDF_AttributeField.ListCount > 0)
  
  
1657:   m_bInitializing = False
1658:   If lstOtherSeriesDF.Enabled Then
1659:     lstOtherSeriesDF.ListIndex = 0
1660:   End If
  
                                            'When updating the extent in other data
                                            'frames on map page loads, options are
                                            'provided about how to update the extent.
                                            'Options include setting the extent based on
                                            'the envelope around the polygon,
                                            'or copying the extent of the main data frame.
                                            'A future version of this software may allow
                                            'varying which option is selected based on
                                            'the other data frame, this version uses the
                                            'same option for all other data frames.
1672:   With cboOtherSeriesExtentOption
1673:     .Clear
                                            'if these entries are altered, alter also
                                            'the strings in NWSeriesOptions.cls, in the
                                            'NWMapBookPrj project.
1677:     .AddItem "Use polygon extent (default)"
1678:     .AddItem "Use main dataframe's scale"
1679:     sDFUpdateOption = m_pNWSeriesOptions.DataFrameToUpdateExtentOption
1680:     lDFUpdateOption = FindControlString(cboOtherSeriesExtentOption, sDFUpdateOption)
1681:     If lDFUpdateOption >= 0 Then
1682:       .ListIndex = lDFUpdateOption
1683:     End If
1684:     .Refresh
1685:   End With
  'end of NW Options tab
  
  
  'Make sure the wizard stays on top
  'TopMost Me
  
  Exit Sub
ErrHand:
1694:   MsgBox "frmSeriesProperties_Load - " & Err.Description
End Sub






Public Function GetDataFrameFromPageLayout(pApp As IApplication, sDataFrameName) As IMapFrame
  On Error GoTo ErrorHandler

  Dim pMapFrame As IMapFrame, pMxDoc As IMxDocument, pValue As IMapFrame
  Dim pGraphicsContainer As IGraphicsContainer, pElement As IElement
  
1708:   Set pMxDoc = pApp.Document
1709:   Set pGraphicsContainer = pMxDoc.PageLayout
1710:   Set pValue = Nothing
1711:   pGraphicsContainer.Reset
1712:   Set pElement = pGraphicsContainer.Next
  
1714:   Do While (Not pElement Is Nothing) And (pValue Is Nothing)
1715:     If TypeOf pElement Is IMapFrame Then
1716:       Set pMapFrame = pElement
                                                  'supposed to just list "other" data frames, not the
                                                  'main data frame containing the map series
1719:       If StrComp(pMapFrame.Map.Name, sDataFrameName) = 0 Then
                                                    'add the data frame to the list
1721:         Set pValue = pMapFrame
1722:       End If
1723:     End If
1724:     Set pElement = pGraphicsContainer.Next
1725:   Loop
1726:   Set GetDataFrameFromPageLayout = pValue
  

  Exit Function
ErrorHandler:
  HandleError True, "GetDataFrameFromPageLayout " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function





'LoadCboOtherSeriesDF_PolygonFC
'
' Populate the combobox cboOtherSeriesDF_PolygonFC with the layer names from
' the passed data frame name that happen to qualify as bubble polygon layers.
' Also, populate the cboOtherSeriesDF_Attributes combobox for the default
' selected feature class.
'-----------------------------------------
Private Sub LoadCboOtherSeriesDF_PolygonFC( _
    sDataFrameName As String, _
    pNWOptions As INWMapSeriesOptions, _
    pApp As IApplication)
  On Error GoTo ErrorHandler
                                                  'select the first of those data frames,
                                                  'if that data frame has any bubble layers, list them
                                                  'in cboOtherSeriesDF_PolygonFC, and if at least one
                                                  'such bubble feature class exists, list the attribute
                                                  'field for the map page name in
                                                  'cboOtherSeriesDF_Attribute
  Dim pFeatureLayer As IFeatureLayer, pMapFrame As IMapFrame, pLayers As IEnumLayer
  Dim pLayer As ILayer, sLayer As String, sField As String, lLayerIdx As Long
  
  
1760:   If (sDataFrameName = "") Or (pNWOptions Is Nothing) Then
1761:     cboOtherSeriesDF_PolygonFC.Clear
1762:     cboOtherSeriesDF_AttributeField.Clear
1763:     cboOtherSeriesDF_PolygonFC.Enabled = False
1764:     cboOtherSeriesDF_AttributeField.Enabled = False
    Exit Sub
1766:   End If

1768:   If pNWOptions.DataFrameIsInStorage(sDataFrameName) Then
    'get the data frame from storage
1770:     Set pMapFrame = pNWOptions.DataFrameStoredFrame(sDataFrameName)
1771:   Else
    'get the data frame from the PageLayout
1773:     Set pMapFrame = GetDataFrameFromPageLayout(pApp, sDataFrameName)
1774:   End If
                                                  'get qualified bubble layers from
                                                  'the data frame
1777:   Set pLayers = pMapFrame.Map.Layers
1778:   pLayers.Reset
1779:   Set pLayer = pLayers.Next
1780:   Do While Not pLayer Is Nothing
1781:     If TypeOf pLayer Is IFeatureLayer Then
1782:       Set pFeatureLayer = pLayer
1783:       If pFeatureLayer.FeatureClass.ShapeType = esriGeometryPolygon Then
1784:         cboOtherSeriesDF_PolygonFC.AddItem pLayer.Name
1785:       End If
1786:     End If
1787:     Set pLayer = pLayers.Next
1788:   Loop
  
1790:   cboOtherSeriesDF_PolygonFC_ChangeHandler
'  With cboOtherSeriesDF_PolygonFC
'    If .ListCount = 0 Then
'      .Enabled = False
'      cboOtherSeriesDF_AttributeField.Clear
'      cboOtherSeriesDF_AttributeField.Enabled = False
'      Exit Sub
'    End If
'    .Enabled = True
'
'    'setting the listindex triggers the _Click
'    'event of cboOtherSeriesDF_PolygonFC.
'    'This in turn causes the list of attributes
'    'to be loaded for the feature class
'
'                                            'make a default selection
'    .ListIndex = 0
'                                            'if a selection was previously made,
'                                            'change the default selection to match
'                                            'the previous selection
'    If m_pDictDFsToUpdate_Layer.Exists(sDataFrameName) Then
'      sLayer = m_pDictDFsToUpdate_Layer(sDataFrameName)
'      lLayerIdx = FindControlString(cboOtherSeriesDF_PolygonFC, sLayer)
'      If lLayerIdx >= 0 Then
'        .ListIndex = lLayerIdx
'      Else
'        sLayer = .List(.ListIndex)
'      End If
'    End If
'
'  End With
  
  Exit Sub
ErrorHandler:
  HandleError False, "LoadCboOtherSeriesDF_PolygonFC " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub



'LoadCboOtherSeriesDF_Attributes
'
'Load the list of attribute fields into cboOtherSeriesDF.
'This combobox is part of the NW Options tab, and this
'list should be populated when the user wants to update the extent
'of data frames other than the main map frame whenever a new map page
'is loaded.  Those other data frames will need to have a map tile
'feature class, and the map tile name field will have to be identified
'for that feature class.  This routine helps populate the UI that will
'let the user indicate which attribute field of that tile feature class
'they want to use as the map tile name attribute field.
'
'todo - see if there is already a selected attribute field for this
'       data frame name and feature class name, and set that field as the
'       default selected field if it is
'------------------------------------------
'Private Sub LoadCboOtherSeriesDF_Attributes(pLayer As ILayer, pNWOptions As INWMapSeriesOptions, sDataFrame As String, sPolyFCName As String)
'  On Error GoTo ErrorHandler
'
'  Dim pField As IField, pFields As IFields, pTable As ITable, lFieldCount As Long
'  Dim i As Long, sCurrentLayer As String, lFindResult As Long, sCurrentField As String
'  Dim sFieldName As String
'
'  With cboOtherSeriesDF_AttributeField
'    .Clear
'    If Not TypeOf pLayer Is ITable Then
'      .Enabled = False
'      Exit Sub
'    End If
'
'    Set pTable = pLayer
'    Set pFields = pTable.Fields
'    lFieldCount = pFields.FieldCount
'    For i = 0 To (lFieldCount - 1)
'      Set pField = pFields.Field(i)
'      If pField.Type = esriFieldTypeString Then
'        sFieldName = pField.Name
'        .AddItem sFieldName
'      End If
'    Next i
'
'    If .ListCount = 0 Then
'      .Enabled = False
'      Exit Sub
'    End If
'    .Enabled = True
'
'  End With
'
'    'set the selected attribute field to a previous selection,
'    'assuming that one was previously made.  Otherwise, just
'    'leave the selected index equal to zero.
'
'  .ListIndex = 0  'set the default setting, and trigger the click event
'
'
''      sCurrentField = m_pNWSeriesOptions.DataFrameToUpdateGetPageNameField(sDataFrame, sPolyFCName)
''      lFindResult = FindControlString(cboOtherSeriesDF_AttributeField, sCurrentField)
''      If lFindResult >= 0 Then
''        cboOtherSeriesDF_AttributeField.ListIndex = lFindResult
''        sFieldName = sCurrentField
''      End If
'                                          'when setting a listindex other than
'                                          'just 0, prefer the local data structures
'                                          'that track previous selections over the
'                                          'ones stored in m_pNWSeriesOptions.  The
'                                          'local data structures will store more recent
'                                          'selections, and m_pNWSeriesOptions will
'                                          'store selections from a previous session.
'                                          'The local structures will be committed back
'                                          'to m_pNWSeriesOptions when the user clicks [OK]
'
'    'only grab the previous attribute field
'    'selection if the newly selected layer
'    'name matches the previous selection
'    sFieldName = ""
'    With m_pDictDFsToUpdate_Layer
'      If .Exists(sDataFrame) Then
'        sCurrentLayer = .Item(sDataFrame)
'        If StrComp(sCurrentLayer, sPolyFCName, vbTextCompare) = 0 Then
'          sFieldName = m_pDictDFsToUpdate_Field(sDataFrame)
''            lFindResult = FindControlString(cboOtherSeriesDF_PolygonFC, sCurrentLayer)
''            If lFindResult >= 0 Then
''              cboOtherSeriesDF_PolygonFC.ListIndex = lFindResult
''              sPolyFCName = sCurrentLayer
''            End If
'        End If
''          .Remove sDataFrame
'      Else
'        'grab any settings from previous NWSeriesOptions
'        sCurrentLayer = m_pNWSeriesOptions.DataFrameToUpdateGetMapPageLayer(sDataFrame)
'        If Len(sCurrentLayer) > 0 Then
'          sFieldName = m_pNWSeriesOptions.DataFrameToUpdateGetPageNameField(sDataFrameName, sCurrentLayer)
''            lFindResult = FindControlString(cboOtherSeriesDF_PolygonFC, sCurrentLayer)
''            If lFindResult >= 0 Then
''              cboOtherSeriesDF_PolygonFC.ListIndex = lFindResult
''              sPolyFCName = sCurrentLayer
''            End If
'        End If
'      End If
'
'      If Len(sFieldName) > 0 Then
'        lFindResult = FindControlString(cboOtherSeriesDF_AttributeField, sCurrentLayer)
'        If lFindResult >= 0 Then
'          cboOtherSeriesDF_AttributeField.ListIndex = lFindResult
'        End If
'      End If
'    End With
'
'    'update the local data structure with
'    'the updated selection
'''''This code became unnecessary when I put it into the click event for
'''''cboOtherSeriesDF_Attributes.
'''''
'''''    With m_pDictDFsToUpdate_Layer
'''''      If .Exists(sDataFrame) Then
'''''        .Remove sDataFrame
'''''      End If
'''''      .Add sDataFrame, sPolyFCName
'''''    End With
'''''    With m_pDictDFsToUpdate_Field
'''''      If .Exists(sDataFrame) Then
'''''        .Remove sDataFrame
'''''      End If
'''''      .Add sDataFrame, sFieldName
'''''    End With
'
'  End If
'
'  Exit Sub
'ErrorHandler:
'  HandleError False, "LoadCboOtherSeriesDF_Attributes " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
'End Sub






'This subroutine initializes the UI for the data frames combobox.
'It loads the list of data frames from the map layout,
'and it loads the list of stored data frames (if any), frames that
'are part of the map series, but night not be currently loaded in the
'map document.
'
'Also sets the list index to the current focus map
'---------------------------
Private Sub LoadDataFrames()
  
  Dim pPageLayout As IPageLayout, pGraphicsContainer As IGraphicsContainer
  Dim pElement As IElement, pMapFrame As IMapFrame, pMxDoc As IMxDocument
  Dim sStoredFrame As String, vStoredFrames As Variant, lStoredFramesCount As Long
  Dim sFocusMap As String, lFindResult As Long, i As Long
  
  'load list of data frames
  '''''''''''''''''''''''''
  If m_pApp Is Nothing Then Exit Sub
1986:   Set pMxDoc = m_pApp.Document
1987:   Set pPageLayout = pMxDoc.PageLayout
1988:   Set pGraphicsContainer = pPageLayout
1989:   pGraphicsContainer.Reset
1990:   Set pElement = pGraphicsContainer.Next
  
  'add all desired data frames from the map layout
1993:   cboDataFrameToFilter.Clear
1994:   Do While Not pElement Is Nothing
1995:     If TypeOf pElement Is IMapFrame Then
1996:       Set pMapFrame = pElement
1997:       If StrComp(Left$(pMapFrame.Map.Name, Len("BubbleID:")), "BubbleID:", vbTextCompare) <> 0 Then
1998:         lFindResult = FindControlString(cboDataFrameToFilter, pMapFrame.Map.Name, 0, True)
1999:         If lFindResult > -1 Then
                                                  'notify the user that more than one data
                                                  'frame has the same name
          '          If Not bWarningWasGiven Then
          '            bWarningWasGiven = True
          '            MsgBox "Warning: Duplicate data frame name " & pMapFrame.Map.Name & " was detected." & vbNewLine _
          '                 & "Only one entry will be accessible for creating dynamic definition queries." & vbNewLine
          '          End If
2007:         Else
                                                  'add the data frame to the list
2009:           cboDataFrameToFilter.AddItem pMapFrame.Map.Name
2010:         End If
2011:       End If
2012:     End If
2013:     Set pElement = pGraphicsContainer.Next
2014:   Loop
  
  'add stored data frames (those not
  'currently in the map layout)
2018:   vStoredFrames = m_pNWSeriesOptions.DataFramesStored
2019:   lStoredFramesCount = UBound(vStoredFrames) + 1
2020:   For i = 0 To (lStoredFramesCount - 1)
2021:     sStoredFrame = vStoredFrames(i)
2022:     If FindControlString(cboDataFrameToFilter, sStoredFrame, -1, True) > -1 Then
      '      MsgBox "Warning: Visibility of duplicate name data frames will not be " & vbNewLine _
      '           & "tracked by the NW Map Book application." & vbNewLine _
      '           & "More than one data frame called ''" & sStoredFrame & "''" & vbNewLine _
      '           & "was detected.", vbOKOnly
2027:     Else
2028:       cboDataFrameToFilter.AddItem sStoredFrame
2029:     End If
2030:   Next i
  
2032:   sFocusMap = pMxDoc.FocusMap.Name
2033:   cboDataFrameToFilter.ListIndex = FindControlString(cboDataFrameToFilter, sFocusMap, -1, True)

End Sub


'This routine will populate part of the dynamic definition query UI,
'and will load layers from the selected data frame.  This function
'was written with the intention to be called from the form
'initialization code, and from the event for when a new data frame selection
'has been made.
'
'Input variables:
'  sDataFrameName parameter
'Module level variables:
'  m_pApp, application reference
'  m_pNWSeriesOptions, custom object reference
'-------------------------------------
Private Sub LoadLayersUIFromDataFrame(sDataFrameName As String)
  On Error GoTo ErrorHandler

  
  Dim vLayerList As Variant, pElement As IElement, pMapFrame As IMapFrame
  Dim pMxDoc As IMxDocument, pPageLayout As IPageLayout, pMap As IMap
  Dim pGraphicsContainer As IGraphicsContainer, pLayers As IEnumLayer
  Dim pLayer As ILayer, bMapIsFound As Boolean, i As Long, lLayerCount As Long
  Dim lListCount As Long, sLayerName As String, sDataFrameToFilter As String
  Dim lSelectedIndex As Long
  
  If m_pNWSeriesOptions Is Nothing Then Exit Sub
2062:   m_bDefQueryInitializing = True
2063:   txtDefinitionQuery.Text = ""
2064:   m_bDefQueryInitializing = False
2065:   lstFilterLayers.Clear
                                                  'determine if the data frame is in the page
                                                  'layout or in storage.
2068:   If m_pNWSeriesOptions.DataFrameIsInStorage(sDataFrameName) Then
    'get the layer list from storage
2070:     vLayerList = m_pNWSeriesOptions.DataFrameStoredFrameLayerList(sDataFrameName)
2071:     lListCount = UBound(vLayerList) + 1
2072:     For i = 0 To (lListCount - 1)
2073:       sLayerName = vLayerList(i)
2074:       lstFilterLayers.AddItem sLayerName
2075:     Next i
2076:   Else
    'grab the layer list from the map
    If m_pApp Is Nothing Then Exit Sub
2079:     Set pMxDoc = m_pApp.Document
2080:     Set pPageLayout = pMxDoc.PageLayout
2081:     Set pGraphicsContainer = pPageLayout
2082:     pGraphicsContainer.Reset
2083:     Set pElement = pGraphicsContainer.Next
2084:     bMapIsFound = False
2085:     Do While (Not pElement Is Nothing) And (Not bMapIsFound)
2086:       If TypeOf pElement Is IMapFrame Then
2087:         Set pMapFrame = pElement
2088:         If StrComp(pMapFrame.Map.Name, sDataFrameName, vbTextCompare) = 0 Then
2089:           Set pMap = pMapFrame.Map
2090:           bMapIsFound = True
2091:         End If
2092:       End If
2093:       Set pElement = pGraphicsContainer.Next
2094:     Loop
    
2096:     If bMapIsFound Then
2097:       Set pLayers = pMap.Layers
2098:       Set pLayer = pLayers.Next
2099:       Do While Not pLayer Is Nothing
2100:         lstFilterLayers.AddItem pLayer.Name
2101:         Set pLayer = pLayers.Next
2102:       Loop
2103:     End If
2104:     If lstFilterLayers.ListCount > 0 Then
2105:       lstFilterLayers.ListIndex = 0
2106:       txtDefinitionQuery.Enabled = lstFilterLayers.Selected(0)
2107:     End If
2108:   End If
2109:   lSelectedIndex = 0
2110:   With lstFilterLayers
2111:     For i = (.ListCount - 1) To 0 Step -1
2112:       sDataFrameToFilter = cboDataFrameToFilter.List(cboDataFrameToFilter.ListIndex)
2113:       .Selected(i) = m_pNWSeriesOptions.DynamicDefQueryIsTrackingLayer(sDataFrameToFilter, .List(i))
2114:       If .Selected(i) Then
2115:         lSelectedIndex = i
2116:       End If
2117:     Next i
2118:   End With
  
2120:   sLayerName = lstFilterLayers.List(lSelectedIndex)
2121:   Set pLayer = GetDefinitionQueryLayer(m_pApp, m_pNWSeriesOptions, sLayerName)
2122:   If Not pLayer Is Nothing Then
2123:     Set m_pFeatLayerDefinition = pLayer
2124:     m_bDefQueryInitializing = True
2125:     txtDefinitionQuery.Text = m_pFeatLayerDefinition.DefinitionExpression
2126:     m_bDefQueryInitializing = False
2127:   End If

  Exit Sub
ErrorHandler:
  HandleError False, "LoadLayersUIFromDataFrame " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub



Private Sub PopulateFieldCombos()
On Error GoTo ErrHand:
  Dim pIndexLayer As IFeatureLayer, pMap As IMap, lLoop As Long
  Dim pFields As IFields, pDoc As IMxDocument
  
2141:   Set pDoc = m_pApp.Document
2142:   Set pMap = FindDataFrame(pDoc, cmbDetailFrame.Text)
2143:   If pMap Is Nothing Then
2144:     MsgBox "Could not find detail frame!!!"
    Exit Sub
2146:   End If
  
2148:   Set pIndexLayer = FindLayer(cmbIndexLayer.Text, pMap)
2149:   If pIndexLayer Is Nothing Then
2150:     MsgBox "Could not find specified layer!!!"
    Exit Sub
2152:   End If
  
  'Populate the index layer combos
2155:   Set pFields = pIndexLayer.FeatureClass.Fields
2156:   cmbDataDriven.Clear
2157:   cmbRotateField.Clear
2158:   For lLoop = 0 To pFields.FieldCount - 1
    Select Case pFields.Field(lLoop).Type
    Case esriFieldTypeDouble, esriFieldTypeSingle, esriFieldTypeInteger
2161:       If UCase(pFields.Field(lLoop).Name) <> "SHAPE_LENGTH" And _
       UCase(pFields.Field(lLoop).Name) <> "SHAPE_AREA" Then
2163:         cmbDataDriven.AddItem pFields.Field(lLoop).Name
2164:         cmbRotateField.AddItem pFields.Field(lLoop).Name
2165:       End If
2166:     End Select
2167:   Next lLoop
2168:   If cmbDataDriven.ListCount > 0 Then
2169:     cmbDataDriven.ListIndex = 0
2170:     cmbRotateField.ListIndex = 0
2171:     optExtent.Item(2).Enabled = True
2172:     chkOptions(0).Enabled = True
2173:   Else
2174:     optExtent.Item(2).Enabled = False
2175:     chkOptions(0).Enabled = False
2176:   End If
  
  Exit Sub
  
ErrHand:
2181:   MsgBox "PopulateFieldCombos - " & Err.Description
End Sub

Private Sub Form_Terminate()
2185:   Set m_pApp = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
2189:   Set m_pApp = Nothing
End Sub


Private Function GetDefinitionQueryLayer(pApp As IApplication, pNWSeriesOpts As INWMapSeriesOptions, sLayerName As String) As ILayer
  On Error GoTo ErrorHandler

  Dim pMxDoc As IMxDocument, pMap As IMap, pLayer As ILayer, pLayers As IEnumLayer
  Dim bLayerIsFound As Boolean
  
2199:   Set pMxDoc = m_pApp.Document
2200:   Set pMap = FindDataFrame(pMxDoc, m_sCurrentDataFrame)
  'sLayerName = lstFilterLayers.List(lstFilterLayers.ListIndex)
  
                                                  'access the feature class of the selected
                                                  'data frame
2205:   If pMap Is Nothing Then
2206:     Set pLayer = m_pNWSeriesOptions.DataFrameStoredFrameLayer(m_sCurrentDataFrame, sLayerName)
2207:   Else
    'grab the map from the stored data frames
2209:     Set pLayers = pMap.Layers
2210:     pLayers.Reset
2211:     Set pLayer = pLayers.Next
2212:     bLayerIsFound = False
2213:     Do While Not pLayer Is Nothing And Not bLayerIsFound
2214:       If (StrComp(pLayer.Name, sLayerName, vbTextCompare) = 0) Then
2215:         bLayerIsFound = True
2216:       Else
2217:         Set pLayer = pLayers.Next
2218:       End If
2219:     Loop
2220:   End If
2221:   Set GetDefinitionQueryLayer = pLayer

  Exit Function
ErrorHandler:
  HandleError False, "GetDefinitionQueryLayer " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Sub lstFilterLayers_Click()
  On Error GoTo ErrorHandler

  Dim pLayer As ILayer, sSearchString As String
  Dim sLayerName As String, pFeatLayerDefinition As IFeatureLayerDefinition
  Dim lStrStart As Long
  
2235:   sSearchString = m_pNWSeriesOptions.DynamicDefQueryReplaceString
2236:   If Len(sSearchString) = 0 Then
2237:     lblNoPageNameWarning.Visible = True
    Exit Sub
2239:   End If
2240:   lblNoPageNameWarning.Visible = False
  
  If m_pApp Is Nothing Then Exit Sub
  If m_pNWSeriesOptions Is Nothing Then Exit Sub
  
2245:   sLayerName = lstFilterLayers.List(lstFilterLayers.ListIndex)
2246:   Set pLayer = GetDefinitionQueryLayer(m_pApp, m_pNWSeriesOptions, sLayerName)
  If pLayer Is Nothing Then Exit Sub
  'pLayer should now point to what the user selected
  
  
                                                  'access the definition query for the selected
                                                  'feature class,
2253:   If Not TypeOf pLayer Is IFeatureLayerDefinition Then
2254:     m_bDefQueryInitializing = True
2255:     txtDefinitionQuery.Text = "Layer " & sLayerName & " does not support definition queries." & vbNewLine _
      & vbNewLine _
      & "Layer types that support definition queries for version 9 of ArgGIS are: " & vbNewLine _
      & "feature layers, coverage annotation layers, annotation layers, cad feature" & vbNewLine _
      & "layers, geodatabase raster catalog layers, tracking analyst temporal feature " & vbNewLine _
      & "layers, and dimension layers."
2261:     m_bDefQueryInitializing = False
2262:     txtDefinitionQuery.Enabled = False
    Exit Sub
2264:   Else
2265:     txtDefinitionQuery.Enabled = True
2266:     lstFilterLayers.Tag = lstFilterLayers.ListIndex
2267:     Set m_pFeatLayerDefinition = pLayer
                                                  'put layer's definition query into the textbox
2269:     Set pFeatLayerDefinition = pLayer
2270:     m_bDefQueryInitializing = True
2271:     txtDefinitionQuery.Text = pFeatLayerDefinition.DefinitionExpression
2272:     m_bDefQueryInitializing = False
2273:   End If
  
2275:   With lstFilterLayers
2276:     m_bDefQueryInitializing = True
2277:     If .Selected(.ListIndex) Then
                                                  'highlight the text for what will
                                                  'be replaced
2280:       lStrStart = InStr(1, txtDefinitionQuery.Text, sSearchString, vbTextCompare)
2281:       If lStrStart <= 0 Then
2282:         txtDefinitionQuery.SelStart = 0
2283:         txtDefinitionQuery.SelLength = 0
2284:       Else
2285:         txtDefinitionQuery.SelStart = lStrStart - 1
2286:         txtDefinitionQuery.SelLength = Len(sSearchString)
2287:       End If
                                                  'set the NWMapSeries data structure to
                                                  'track that this layer has a dynamic
                                                  'definition query
2291:       m_pNWSeriesOptions.DynamicDefQueryAddLayer m_sCurrentDataFrame, sLayerName
      
2293:     Else
                                                  'disable the textbox, turn off all
                                                  'highlighting
2296:       With txtDefinitionQuery
2297:         .SelLength = 0
2298:         .SelStart = 0
2299:         .Enabled = False
2300:       End With
                                                  'make sure this feature class is not
                                                  'being tracked for dynamic definition
                                                  'queries
2304:       m_pNWSeriesOptions.DynamicDefQueryRemoveLayer m_sCurrentDataFrame, sLayerName
2305:     End If
2306:     m_bDefQueryInitializing = False
2307:   End With


  Exit Sub
ErrorHandler:
  HandleError True, "lstFilterLayers_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub lstLyrGroups_Click()
  On Error GoTo ErrorHandler

  Dim sGroupName As String, pLyrGroup As INWLayerVisibilityGroup, i As Long
  
2320:   If Not m_bIsRecursiveEntrylstLyrGroups Then
2321:     m_bIsRecursiveEntrylstLyrGroups = True
    
2323:     With lstLyrGroups
2324:       If .ListCount = 0 Then
2325:         m_bIsRecursiveEntrylstLyrGroups = False
        Exit Sub
2327:       End If

2329:       If m_pNWSeriesOptions Is Nothing Then
2330:         m_bIsRecursiveEntrylstLyrGroups = False
        Exit Sub
2332:       End If
2333:       If StrComp(m_sSelectedGroup, .List(.ListIndex), vbTextCompare) = 0 Then
2334:         m_bIsRecursiveEntrylstLyrGroups = False
        Exit Sub
2336:       End If
      
2338:       m_sSelectedGroup = .List(.ListIndex)
2339:       Set pLyrGroup = m_pNWSeriesOptions.LayerGroupGet(m_sSelectedGroup)
2340:       For i = 0 To lstVisibleLayers.ListCount - 1
2341:         If pLyrGroup.Exists(lstVisibleLayers.List(i)) Then
2342:           lstVisibleLayers.ItemData(i) = vbUnchecked
2343:           lstVisibleLayers.Selected(i) = False
2344:         Else
2345:           lstVisibleLayers.ItemData(i) = vbChecked
2346:           lstVisibleLayers.Selected(i) = True
2347:         End If
2348:       Next i
2349:       lstVisibleLayers.Refresh
2350:     End With
    
2352:     m_bIsRecursiveEntrylstLyrGroups = False
2353:   End If

  Exit Sub
ErrorHandler:
  HandleError True, "lstLyrGroups_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub






'lstOtherSeriesDF_Click
'   This routine manages the UI for a list of data frames, allowing the
'   user to choose whether or not data frames other than detail insets and
'   the main data frame will have their extent updated when new map pages
'   are loaded.
'
'   Clicking a data frame name will cause two other controls to be populated
'   with data, but this has a cascading effect on the click and change events
'   of those two controls.  This cascading effect process is complex enough
'   to be listed here, from the routine that starts the process.
'
'  lstOtherSeriesDF
'    Change
'      - do nothing
'    Click
'      - populate other controls if selected (checked)
'        (triggers change event)
'
'
'  cboOtherSeriesDF_PolygonFC
'    Change
'      This handles scenario of user never selecting an
'      attr field or a polygon FC, and clicking [OK]
'      - cause list of fields to be populated for this
'        polygon feature class. (triggers attr field's
'        change event)
'      - set listindex to previous selection, or to first
'        entry if previous selection doesn't exist.
'       (triggers cboOtherSeriesDF_PolygonFC click event)
'    Click
'      This handles scenario of
'      - populate the attribute field list,
'        make a default selection from the attribute
'        field list in case user clicks [ok]
'        (triggers attr field change and click events)
'
'  cboOtherSeriesDF_AttributeField
'    Change
'      - make a listindex selection based on any previous
'        selection, otherwise just select the first
'        qualified attribute field (triggers attr click
'        event)
'    Click
'      - update the selection data structure with both the
'        current layer name and the current field name.
'        (this relies on the click event being trigger
'---------------------------------
Private Sub lstOtherSeriesDF_Click()
  On Error GoTo ErrorHandler

2415:   If m_bInitializing Then
    Exit Sub
2417:   End If
  
2419:   cboOtherSeriesDF_PolygonFC.Clear
2420:   cboOtherSeriesDF_AttributeField.Clear
2421:   cboOtherSeriesDF_PolygonFC.Enabled = False
2422:   cboOtherSeriesDF_AttributeField.Enabled = False
2423:   With lstOtherSeriesDF
2424:     If .ListIndex >= 0 Then
2425:       If .Selected(.ListIndex) Then
                                            'populate the comboboxes with information
                                            'about the selected data frame
2428:         LoadCboOtherSeriesDF_PolygonFC lstOtherSeriesDF.List(.ListIndex), _
            m_pNWSeriesOptions, _
            m_pApp
2431:       Else
                                            'user has un-checked the current item,
                                            'so data structures should be cleared
2434:         If m_pDictDFsToUpdate_Layer.Exists(.List(.ListIndex)) Then  'an object variable or with block variable is not set
2435:           m_pDictDFsToUpdate_Layer.Remove .List(.ListIndex)
2436:         End If
2437:         If m_pDictDFsToUpdate_Field.Exists(.List(.ListIndex)) Then
2438:           m_pDictDFsToUpdate_Field.Remove .List(.ListIndex)
2439:         End If
2440:       End If
2441:     End If
2442:   End With

  Exit Sub
ErrorHandler:
  HandleError True, "lstOtherSeriesDF_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub




Private Sub lstVisibleLayers_Click()
  On Error GoTo ErrorHandler

  Dim sSelectedLayer As String, sLyrName As String, bSelected As Boolean
  Dim pLyrGroup As INWLayerVisibilityGroup, lChkValue As Long
       
  If m_sSelectedGroup = "" Then Exit Sub
2459:   With lstLyrGroups
    If .ListCount = 0 Then Exit Sub
2461:     If .SelCount = 0 Then
2462:       If .ListIndex < 0 Or (.ListIndex >= .ListCount) Then
2463:         .ListIndex = 0
2464:       End If
2465:       .Selected(.ListIndex) = True
2466:     End If
2467:   End With
  
2469:   With lstVisibleLayers
2470:     If .ListIndex = -1 Then
2471:       If .ListCount > 0 Then
2472:         .ListIndex = 0
2473:       Else
        Exit Sub
2475:       End If
2476:     End If
2477:     If .ListIndex >= .ListCount Then .ListIndex = (.ListCount - 1)
2478:   End With
                                                  'if this code is not triggered
                                                  'by user input, but by the
                                                  'lstLyrGroup_Click handler, exit
  If m_bIsRecursiveEntrylstLyrGroups Then Exit Sub
  
2484:   If Not m_bIsRecursiveEntry_lstVisibleLayers Then
2485:     m_bIsRecursiveEntry_lstVisibleLayers = True
                                                  'confirm that a group name is selected
2487:     With lstVisibleLayers
2488:       sLyrName = .List(.ListIndex)
                                                  'grab the name of the selected
                                                  'group on the left.
2491:       Set pLyrGroup = m_pNWSeriesOptions.LayerGroupGet(m_sSelectedGroup)
                                                  'if this name is present in that group,
                                                  ' - remove this name from that group,
                                                  'otherwise,
                                                  ' - add this layer name to that group,
      'lChkValue = .ItemData(.ListIndex)
2497:       bSelected = .Selected(.ListIndex)
                                                  'testing revealed that this event is
                                                  'triggered right **BEFORE** the
                                                  'checked/unchecked value is changed.  the
                                                  'itemdata therefore reflects what the check
                                                  'mark was, not is.
2503:       If Not bSelected Then
2504:         If Not pLyrGroup.Exists(sLyrName) Then
2505:           pLyrGroup.AddLayer (sLyrName)
2506:         End If
2507:       Else 'If lChkValue = vbUnchecked Then
2508:         If pLyrGroup.Exists(sLyrName) Then
2509:           pLyrGroup.DeleteLayer (sLyrName)
2510:         End If
2511:       End If
2512:     End With
    
2514:     m_bIsRecursiveEntry_lstVisibleLayers = False
2515:   End If

  Exit Sub
ErrorHandler:
  HandleError True, "lstVisibleLayers_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub


Private Sub optExtent_Click(Index As Integer)
On Error GoTo ErrHand:
  Select Case Index
  Case 0  'Variable
2527:     txtMargin.Enabled = True
2528:     cmbMargin.Enabled = True
2529:     txtFixed.Enabled = False
2530:     cmbDataDriven.Enabled = False
2531:     If txtMargin.Text = "" Then
2532:       cmdOK.Enabled = False
2533:     Else
2534:       cmdOK.Enabled = True
2535:     End If
  Case 1  'Fixed
2537:     txtMargin.Enabled = False
2538:     cmbMargin.Enabled = False
2539:     txtFixed.Enabled = True
2540:     cmbDataDriven.Enabled = False
2541:     If txtFixed.Text = "" Then
2542:       cmdOK.Enabled = False
2543:     Else
2544:       cmdOK.Enabled = True
2545:     End If
  Case 2  'Data driven
2547:     txtMargin.Enabled = False
2548:     cmbMargin.Enabled = False
2549:     txtFixed.Enabled = False
2550:     cmbDataDriven.Enabled = True
2551:     cmdOK.Enabled = True
2552:   End Select

  Exit Sub
ErrHand:
2556:   MsgBox "optExtent_Click - " & Err.Description
End Sub





Private Sub txtDefinitionQuery_Change()
  On Error GoTo ErrorHandler

  Dim pLayer As ILayer, sSearchString As String
  Dim sLayerName As String, sDefQueryLayerName As String ', m_pFeatLayerDefinition As IFeatureLayerDefinition
  
2569:   If txtDefinitionQuery.Text = "" Then
2570:     cmdNextPageName.Enabled = False
2571:     cmdPrevPageName.Enabled = False
2572:   Else
2573:     cmdNextPageName.Enabled = True
2574:     cmdPrevPageName.Enabled = True
2575:   End If
  
2577:   If m_bDefQueryInitializing Or m_bRecursiveTxtDefinitionQuery_Change Then
    Exit Sub
2579:   End If
  
2581:   m_bRecursiveTxtDefinitionQuery_Change = True
  
  If m_pApp Is Nothing Then Exit Sub
  If m_pNWSeriesOptions Is Nothing Then Exit Sub
  
  'confirm that the layer being modified was
  'the last one selected in lstFilterLayers
2588:   Set pLayer = m_pFeatLayerDefinition
  If pLayer Is Nothing Then Exit Sub
2590:   sDefQueryLayerName = pLayer.Name
                                                  'put the definition query strings into the
                                                  'selected layer
2593:   sLayerName = lstFilterLayers.List(lstFilterLayers.ListIndex)
2594:   Set pLayer = GetDefinitionQueryLayer(m_pApp, m_pNWSeriesOptions, sLayerName)
  If pLayer Is Nothing Then Exit Sub
2596:   If StrComp(sDefQueryLayerName, sLayerName, vbTextCompare) <> 0 Then
    Exit Sub
2598:   End If
  
  'Set m_pFeatLayerDefinition = pLayer
2601:   m_bDefQueryInitializing = True
2602:   m_pFeatLayerDefinition.DefinitionExpression = txtDefinitionQuery.Text
2603:   m_bDefQueryInitializing = False
  
2605:   m_bRecursiveTxtDefinitionQuery_Change = False

  Exit Sub
ErrorHandler:
2609:   m_bDefQueryInitializing = False
2610:   m_bRecursiveTxtDefinitionQuery_Change = False
  HandleError True, "txtDefinitionQuery_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

'Assign definition query to the layer referenced in lstFilterLayers.
'-----------------------------------------
Private Sub txtDefinitionQuery_LostFocus()
  Dim pLayer As ILayer, sLayerName As String
  
2619:   If lstFilterLayers.Tag = "" Then  'code triggers when user hasn't selected a map page yet
    Exit Sub  'prevents a type mismatch error at the sLayerName = lstFilter.... line of code below
2621:   End If
2622:   sLayerName = lstFilterLayers.List(lstFilterLayers.Tag)
  
2624:   If Not m_pFeatLayerDefinition Is Nothing Then
2625:     Set pLayer = m_pFeatLayerDefinition
2626:     If StrComp(sLayerName, pLayer.Name, vbTextCompare) = 0 Then
2627:       m_pFeatLayerDefinition.DefinitionExpression = txtDefinitionQuery.Text
2628:     End If
2629:   End If
End Sub

Private Sub txtFixed_KeyUp(KeyCode As Integer, Shift As Integer)
2633:   If Not IsNumeric(txtFixed.Text) Then
2634:     txtFixed.Text = ""
2635:   End If
2636:   If txtFixed.Text <> "" Then
2637:     cmdOK.Enabled = True
2638:   End If
End Sub

Private Sub txtMargin_KeyUp(KeyCode As Integer, Shift As Integer)
2642:   If Not IsNumeric(txtMargin.Text) Then
2643:     txtMargin.Text = ""
2644:   End If
2645:   If txtMargin.Text <> "" Then
2646:     cmdOK.Enabled = True
2647:   End If
End Sub

Private Sub txtNeighborLabelIndent_Change()
2651:   With txtNeighborLabelIndent
2652:     If Not IsNumeric(.Text) Then
2653:       .BackColor = vbYellow
2654:     Else
2655:       .BackColor = vbWhite
2656:     End If
2657:   End With
End Sub

Private Sub txtNeighborLabelIndent_LostFocus()
2661:   With txtNeighborLabelIndent
2662:     If Not IsNumeric(.Text) Then
2663:       MsgBox "''" & .Text & "'' could not be interpreted as a number between" & vbNewLine _
           & "-3 and 3.  A default value of zero will be used instead.", vbOKOnly
2665:       .Text = "0"
2666:     Else
      
2668:     End If
2669:   End With
End Sub









Private Function CalculatePageToMapRatio(pApp As IApplication) As Double
    Dim pMx As IMxDocument
    Dim pPage As IPage
    Dim pPageUnits As esriUnits
    Dim pSR As ISpatialReference
    Dim pSRI As ISpatialReferenceInfo
    Dim pPCS As IProjectedCoordinateSystem
    Dim dMetersPerUnit As Double
    
    On Error GoTo eh
    
    ' Init
2692:     Set pMx = pApp.Document
2693:     Set pSR = pMx.FocusMap.SpatialReference
2694:     If TypeOf pSR Is IProjectedCoordinateSystem Then
2695:         Set pPCS = pSR
2696:         dMetersPerUnit = pPCS.CoordinateUnit.MetersPerUnit
2697:     Else
2698:         dMetersPerUnit = 1
2699:     End If
2700:     Set pPage = pMx.PageLayout.Page
2701:     pPageUnits = pPage.Units
    Select Case pPageUnits
        Case esriInches: CalculatePageToMapRatio = dMetersPerUnit / (1 / 12 * 0.304800609601219)
        Case esriFeet: CalculatePageToMapRatio = dMetersPerUnit / (0.304800609601219)
        Case esriCentimeters: CalculatePageToMapRatio = dMetersPerUnit / (1 / 100)
        Case esriMeters: CalculatePageToMapRatio = dMetersPerUnit / (1)
        Case Else:
2708:             MsgBox "Warning: Only the following Page (Layout) Units are supported by this tool:" _
                & vbCrLf & " - Inches, Feet, Centimeters, Meters" _
                & vbCrLf & vbCrLf & "Calculating as though Page Units are in Inches..."
2711:             CalculatePageToMapRatio = dMetersPerUnit / (1 / 12 * 0.304800609601219)
2712:     End Select
    Exit Function
eh:
2715:     CalculatePageToMapRatio = 1
2716:     MsgBox "Error in CalculatePageToMapRatio" & vbCrLf & Err.Description
End Function





