VERSION 5.00
Begin VB.Form frmGridSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Grid Generator Wizard"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8430
   Icon            =   "frmGridSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraDestinationFeatureClass 
      Height          =   4695
      Left            =   5040
      TabIndex        =   37
      Top             =   5280
      Width           =   4815
      Begin VB.CommandButton cmdSetNewGridLayer 
         Height          =   315
         Left            =   4320
         Picture         =   "frmGridSettings.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Set new Grid Layer"
         Top             =   1320
         Width           =   315
      End
      Begin VB.TextBox txtNewGridLayer 
         Height          =   315
         Left            =   2040
         TabIndex        =   74
         Top             =   1320
         Width           =   2295
      End
      Begin VB.OptionButton optLayerSource 
         Caption         =   "Create a new Layer:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   73
         Top             =   1320
         Width           =   1815
      End
      Begin VB.OptionButton optLayerSource 
         Caption         =   "Use existing Layer:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   72
         Top             =   960
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.ListBox lstRequiredDataLayers 
         Height          =   1410
         Left            =   600
         Style           =   1  'Checkbox
         TabIndex        =   45
         Top             =   3120
         Width           =   3855
      End
      Begin VB.ComboBox cmbPolygonLayers 
         Height          =   315
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   39
         Top             =   920
         Width           =   2535
      End
      Begin VB.CheckBox chkRemovePreviousGrids 
         Caption         =   "Clear existing grids.  This will delete all the current"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   1920
         Width           =   4335
      End
      Begin VB.CheckBox chkRemoveEmpties 
         Caption         =   "Don't create empty grids.  A grid is considered empty"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   2400
         Width           =   4095
      End
      Begin VB.Label Label12 
         Caption         =   "unless it contains features from at least one of the following selected layers:"
         Height          =   495
         Left            =   600
         TabIndex        =   44
         Top             =   2620
         Width           =   3735
      End
      Begin VB.Label Label11 
         Caption         =   "features in the feature class."
         Height          =   255
         Left            =   600
         TabIndex        =   43
         Top             =   2145
         Width           =   3855
      End
      Begin VB.Label Label5 
         Caption         =   $"frmGridSettings.frx":05C4
         Height          =   615
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Frame fraGridIDs 
      Height          =   4695
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   4815
      Begin VB.Frame Frame1 
         Caption         =   "Identifier Order"
         Height          =   975
         Left            =   120
         TabIndex        =   65
         Top             =   960
         Width           =   4575
         Begin VB.OptionButton optGridIDOrder 
            Caption         =   "Column - Row"
            Height          =   255
            Index           =   1
            Left            =   2520
            TabIndex        =   67
            Top             =   600
            Width           =   1575
         End
         Begin VB.OptionButton optGridIDOrder 
            Caption         =   "Row - Column"
            Height          =   255
            Index           =   0
            Left            =   2520
            TabIndex        =   66
            Top             =   240
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.Label Label4 
            Caption         =   "Set the order in which to construct the Identifier."
            Height          =   375
            Left            =   120
            TabIndex        =   68
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Row ID Style"
         Height          =   975
         Left            =   120
         TabIndex        =   33
         Top             =   3120
         Width           =   2235
         Begin VB.OptionButton optRowIDType 
            Caption         =   "Alpha (A, B, C, ...)"
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   35
            Top             =   240
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.OptionButton optRowIDType 
            Caption         =   "Numeric (0, 1, 2, ...)"
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   34
            Top             =   600
            Width           =   1860
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Column ID Style"
         Height          =   975
         Left            =   2460
         TabIndex        =   30
         Top             =   3120
         Width           =   2235
         Begin VB.OptionButton optColIDType 
            Caption         =   "Alpha (A, B, C, ...)"
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   32
            Top             =   240
            Width           =   1815
         End
         Begin VB.OptionButton optColIDType 
            Caption         =   "Numeric (0, 1, 2, ...)"
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   31
            Top             =   600
            Value           =   -1  'True
            Width           =   1935
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Starting ID Position"
         Height          =   975
         Left            =   120
         TabIndex        =   27
         Top             =   2040
         Width           =   4575
         Begin VB.OptionButton optStartingIDPosition 
            Caption         =   "Top Left"
            Height          =   255
            Index           =   0
            Left            =   2520
            TabIndex        =   29
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optStartingIDPosition 
            Caption         =   "Lower Left"
            Height          =   255
            Index           =   1
            Left            =   2520
            TabIndex        =   28
            Top             =   600
            Width           =   1160
         End
         Begin VB.Label Label26 
            Caption         =   "Set the position of the (1,1) or (A,A) grid."
            Height          =   495
            Left            =   120
            TabIndex        =   64
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.CheckBox chkBreak 
         Caption         =   "Use an Underscore as a Row/Column separator"
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   4200
         Width           =   4335
      End
      Begin VB.Label Label25 
         Caption         =   "Set the format for the Identifier.  The ID will be stored in the Text field specified earlier."
         Height          =   495
         Left            =   120
         TabIndex        =   63
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label lblExampleID 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "B3"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4320
         TabIndex        =   36
         Top             =   170
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   775
         Left            =   3660
         Picture         =   "frmGridSettings.frx":065F
         Stretch         =   -1  'True
         Top             =   160
         Width           =   1020
      End
   End
   Begin VB.Frame fraScaleStart 
      Height          =   4695
      Left            =   2400
      TabIndex        =   1
      Top             =   6000
      Width           =   4815
      Begin VB.CommandButton cmdLayersExtent 
         Caption         =   "All Layers"
         Height          =   315
         Left            =   3120
         TabIndex        =   71
         ToolTipText     =   "Extent of all Layers in Map"
         Top             =   3960
         Width           =   1215
      End
      Begin VB.OptionButton optScaleSource 
         Caption         =   "Manual Map Scale"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton optScaleSource 
         Caption         =   "Current Map Scale"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtManualMapScale 
         Height          =   285
         Left            =   2760
         TabIndex        =   8
         Text            =   "0"
         Top             =   940
         Width           =   1455
      End
      Begin VB.TextBox txtStartCoordX 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   600
         TabIndex        =   7
         Text            =   "0"
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox txtStartCoordY 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2520
         TabIndex        =   6
         Text            =   "0"
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox txtEndCoordX 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   600
         TabIndex        =   5
         Text            =   "0"
         Top             =   3480
         Width           =   1695
      End
      Begin VB.TextBox txtEndCoordY 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2520
         TabIndex        =   4
         Text            =   "0"
         Top             =   3480
         Width           =   1695
      End
      Begin VB.CommandButton cmdDatasetExtentLL 
         Caption         =   "FClass Extent"
         Height          =   315
         Left            =   480
         TabIndex        =   3
         ToolTipText     =   "Extent of the Grid featureclass"
         Top             =   3960
         Width           =   1215
      End
      Begin VB.CommandButton cmdMapExtentLL 
         Caption         =   "Current Extent"
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         ToolTipText     =   "Current Map Extent"
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label20 
         Caption         =   "Set the Extent within which to create grids."
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   1800
         Width           =   4575
      End
      Begin VB.Label Label8 
         Caption         =   "Set the Scale / Size for each of the grids."
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label lblCurrentMapScale 
         Caption         =   "5,000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   14
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "1 :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   13
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "Starting Coordinate (LowerLeft):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2160
         Width           =   3375
      End
      Begin VB.Label Label10 
         Caption         =   "Ending Coordinate (UpperRight):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   3120
         Width           =   3375
      End
      Begin VB.Label Label7 
         Caption         =   "1 :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   15
         Top             =   600
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "< Back"
      Height          =   375
      Left            =   1440
      TabIndex        =   70
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next >"
      Height          =   375
      Left            =   2540
      TabIndex        =   69
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Frame fraDataFrameSize 
      Height          =   4695
      Left            =   -1200
      TabIndex        =   16
      Top             =   5520
      Width           =   4815
      Begin VB.OptionButton optGridSize 
         Caption         =   "Use the size of the current Data Frame in the Layout"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Value           =   -1  'True
         Width           =   4335
      End
      Begin VB.OptionButton optGridSize 
         Caption         =   "Specify the Data Frame size (in Layout Units)"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   2160
         Width           =   4335
      End
      Begin VB.TextBox txtManualGridWidth 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   19
         Text            =   "0"
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox txtManualGridHeight 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   18
         Text            =   "0"
         Top             =   3720
         Width           =   1215
      End
      Begin VB.ComboBox cmbGridSizeUnits 
         Height          =   315
         ItemData        =   "frmGridSettings.frx":5E21
         Left            =   3480
         List            =   "frmGridSettings.frx":5E2B
         TabIndex        =   17
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "Units:"
         Height          =   255
         Left            =   2880
         TabIndex        =   62
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label Label23 
         Caption         =   "Note: If you use this option, for best results you should also update the Data Frame size in the Layout to match."
         Height          =   495
         Left            =   480
         TabIndex        =   61
         Top             =   2520
         Width           =   4095
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "Frame Name:"
         Height          =   255
         Left            =   480
         TabIndex        =   60
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   $"frmGridSettings.frx":5E3A
         Height          =   975
         Left            =   120
         TabIndex        =   59
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Width : "
         Height          =   255
         Left            =   480
         TabIndex        =   24
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Height : "
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label lblCurrFrameName 
         Caption         =   "Current Frame"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   22
         Top             =   1560
         Width           =   2895
      End
   End
   Begin VB.Frame fraAttributes 
      Height          =   4695
      Left            =   5040
      TabIndex        =   41
      Top             =   120
      Width           =   4815
      Begin VB.ComboBox cmbFieldColNum 
         Height          =   315
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   53
         Top             =   3960
         Width           =   2535
      End
      Begin VB.ComboBox cmbFieldRowNum 
         Height          =   315
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   52
         Top             =   3360
         Width           =   2535
      End
      Begin VB.ComboBox cmbFieldMapScale 
         Height          =   315
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   51
         Top             =   2760
         Width           =   2535
      End
      Begin VB.ComboBox cmbFieldID 
         Height          =   315
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   48
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Column Number:"
         Height          =   255
         Left            =   360
         TabIndex        =   56
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Row Number:"
         Height          =   255
         Left            =   360
         TabIndex        =   55
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Map Scale:"
         Height          =   255
         Left            =   360
         TabIndex        =   54
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label16 
         Caption         =   $"frmGridSettings.frx":5ED8
         Height          =   735
         Left            =   120
         TabIndex        =   50
         Top             =   1920
         Width           =   4575
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "ID:"
         Height          =   255
         Left            =   960
         TabIndex        =   49
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "REQUIRED: Each grid polygon feature requires an Identifier.  Select the Text field that will hold this ID."
         Height          =   495
         Left            =   120
         TabIndex        =   47
         Top             =   720
         Width           =   4575
      End
      Begin VB.Label Label13 
         Caption         =   "Assign roles to field names."
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   4800
      Width           =   1095
   End
End
Attribute VB_Name = "frmGridSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_Application As IApplication
Public GridSettings As clsCreateGrids

Public Enum intersectFileType
  ShapeFile = 0
  AccessFeatureClass = 1
  SDEFeatureClass = 2
End Enum

Private m_bIsGeoDatabase As Boolean
Private m_FileType As intersectFileType
Private m_OutputLayer As String
Private m_OutputDataset As String
Private m_OutputFClass As String
Private m_Step As Integer

Private Const c_DefaultFld_GridID = "GRID_ID"
Private Const c_DefaultFld_ColNum = "COL_NUM"
Private Const c_DefaultFld_RowNum = "ROW_NUM"
Private Const c_DefaultFld_Scale = "PLOTSCALE"

Private Sub SetControlsState()
    Dim dScale As Double
    Dim dGHeight As Double
    Dim dGWidth As Double
    Dim dStartX As Double
    Dim dStartY As Double
    Dim dEndX As Double
    Dim dEndY As Double
    Dim bValidScale As Boolean
    Dim bValidSize As Boolean
    Dim bValidTarget As Boolean
    Dim bValidIDField As Boolean
    Dim bNewFClassSet As Boolean
    Dim bValidReqdLayers As Boolean
    Dim bValidStart As Boolean
    Dim bValidEnd As Boolean
    Dim bCreatingNewFClass As Boolean
    Dim bDuplicateFieldsSelected As Boolean
    Dim pFL As IFeatureLayer
    Dim pDatasetExtent As IEnvelope
    Dim i As Integer
    
    On Error GoTo eh
    
    ' Protect against zero length string_to_double conversions
    If Len(lblCurrentMapScale.Caption) = 0 Then lblCurrentMapScale.Caption = "0"
    If Len(txtManualMapScale.Text) = 0 Then
        dScale = 0
    Else
        dScale = CDbl(txtManualMapScale.Text)
    End If
    If Len(txtManualGridHeight.Text) = 0 Then
        dGHeight = 0
    Else
        dGHeight = CDbl(txtManualGridHeight.Text)
    End If
    If Len(txtManualGridWidth.Text) = 0 Then
        dGWidth = 0
    Else
        dGWidth = CDbl(txtManualGridWidth.Text)
    End If
    If Len(txtStartCoordX.Text) = 0 Then
        dStartX = 0
    Else
        dStartX = CDbl(txtStartCoordX.Text)
    End If
    If Len(txtStartCoordY.Text) = 0 Then
        dStartY = 0
    Else
        dStartY = CDbl(txtStartCoordY.Text)
    End If
    If Len(txtEndCoordX.Text) = 0 Then
        dEndX = 0
    Else
        dEndX = CDbl(txtEndCoordX.Text)
    End If
    If Len(txtEndCoordY.Text) = 0 Then
        dEndY = 0
    Else
        dEndY = CDbl(txtEndCoordY.Text)
    End If
i = 1

    ' Calc values
    bValidScale = (optScaleSource(0).Value And CDbl(lblCurrentMapScale.Caption) > 0) Or _
                  (optScaleSource(1).Value And dScale > 0)
    bValidSize = (optGridSize(0).Value) Or _
                 (optGridSize(1).Value And dGHeight > 0 And dGWidth > 0)
    bCreatingNewFClass = optLayerSource(1).Value
    bNewFClassSet = (Len(txtNewGridLayer.Text) > 0)
    bValidTarget = (cmbPolygonLayers.ListIndex > 0) Or (bCreatingNewFClass And bNewFClassSet)
    bValidIDField = (cmbFieldID.ListIndex >= 0)
    bValidReqdLayers = (chkRemoveEmpties.Value = vbUnchecked) Or _
                       (chkRemoveEmpties.Value = vbChecked And lstRequiredDataLayers.SelCount > 0)
i = 2
    If bValidTarget And (Not bCreatingNewFClass) Then
        Set pFL = FindFeatureLayerByName(cmbPolygonLayers.List(cmbPolygonLayers.ListIndex), m_Application)
        If pFL.FeatureClass.FeatureDataset Is Nothing Then
            bValidStart = True
            bValidEnd = True
        Else
            Set pDatasetExtent = GetValidExtentForLayer(pFL)
            bValidStart = ((dStartX >= pDatasetExtent.XMin) And (dStartX <= pDatasetExtent.XMax)) _
                            And _
                          ((dStartY >= pDatasetExtent.YMin) And (dStartY <= pDatasetExtent.YMax))
            bValidEnd = ((dEndX >= pDatasetExtent.XMin) And (dEndX <= pDatasetExtent.XMax)) _
                            And _
                        ((dEndY >= pDatasetExtent.YMin) And (dEndY <= pDatasetExtent.YMax)) _
                            And _
                        ((dEndX > dStartX) And (dEndY > dStartY))
        End If
    ElseIf bValidTarget And bCreatingNewFClass Then
        bValidStart = True
        bValidEnd = True
    End If
    bDuplicateFieldsSelected = (cmbFieldRowNum.ListIndex > 0 And cmbFieldRowNum.ListIndex = cmbFieldColNum.ListIndex) _
                            Or (cmbFieldRowNum.ListIndex > 0 And cmbFieldRowNum.ListIndex = cmbFieldMapScale.ListIndex) _
                            Or (cmbFieldColNum.ListIndex > 0 And cmbFieldColNum.ListIndex = cmbFieldMapScale.ListIndex)
i = 3
    
    ' Set states
    Select Case m_Step
        Case 0:     ' Set the target feature layer
            cmdBack.Enabled = False
            cmdNext.Enabled = bValidTarget And bValidReqdLayers
            cmdNext.Caption = "Next >"
            cmbPolygonLayers.Enabled = Not bCreatingNewFClass
        Case 1:     ' Set the fields to populate
            cmdBack.Enabled = True
            cmdNext.Enabled = (bValidIDField And Not bDuplicateFieldsSelected)
            cmbFieldID.Enabled = Not bCreatingNewFClass
            cmbFieldRowNum.Enabled = Not bCreatingNewFClass
            cmbFieldColNum.Enabled = Not bCreatingNewFClass
            cmbFieldMapScale.Enabled = Not bCreatingNewFClass
        Case 2:     ' Set the scale / starting_coords
            cmdBack.Enabled = True
            cmdNext.Enabled = bValidScale And bValidStart And bValidEnd
            If Not bCreatingNewFClass Then
                cmdDatasetExtentLL.Enabled = Not (pFL.FeatureClass.FeatureDataset Is Nothing)
            Else
                cmdDatasetExtentLL.Enabled = False
            End If
        Case 3:     ' Set the dataframe properties
            cmdBack.Enabled = True
            cmdNext.Enabled = bValidSize
            cmdNext.Caption = "Next >"
        Case 4:     ' Set the ID values
            cmdBack.Enabled = True
            cmdNext.Enabled = True
            cmdNext.Caption = "Finish"
        Case Else:
            cmdBack.Enabled = False
            cmdNext.Enabled = False
    End Select
i = 4
    
    txtManualMapScale.Enabled = optScaleSource(1).Value
    txtManualGridWidth.Enabled = optGridSize(1).Value
    txtManualGridHeight.Enabled = optGridSize(1).Value
    cmbGridSizeUnits.Enabled = optGridSize(1).Value
    ' Set display
    If bValidStart Then
        txtStartCoordX.ForeColor = (&H0)    ' Black
        txtStartCoordY.ForeColor = (&H0)
    Else
        txtStartCoordX.ForeColor = (&HFF)   ' Red
        txtStartCoordY.ForeColor = (&HFF)
    End If
    If bValidEnd Then
        txtEndCoordX.ForeColor = (&H0)      ' Black
        txtEndCoordY.ForeColor = (&H0)
    Else
        txtEndCoordX.ForeColor = (&HFF)     ' Red
        txtEndCoordY.ForeColor = (&HFF)
    End If
    If optScaleSource(1).Value Then
        If bValidScale Then
            txtManualMapScale.ForeColor = (&H0)      ' Black
        Else
            txtManualMapScale.ForeColor = (&HFF)     ' Red
        End If
    End If
    If optGridSize(1).Value Then
        If bValidSize Then
            txtManualGridWidth.ForeColor = (&H0)      ' Black
            txtManualGridHeight.ForeColor = (&H0)
        Else
            txtManualGridWidth.ForeColor = (&HFF)     ' Red
            txtManualGridHeight.ForeColor = (&HFF)
        End If
    End If
    
    Exit Sub
    Resume
eh:
    MsgBox Err.Description, vbExclamation, "SetControlsState " & i
End Sub

Private Sub chkBreak_Click()
    lblExampleID.Caption = GenerateExampleID
End Sub

Private Sub chkRemoveEmpties_Click()
    SetControlsState
End Sub

Private Sub cmbFieldColNum_Click()
    SetControlsState
End Sub

Private Sub cmbFieldID_Click()
    SetControlsState
End Sub

Private Sub cmbFieldMapScale_Click()
    SetControlsState
End Sub

Private Sub cmbFieldRowNum_Click()
    SetControlsState
End Sub

Private Sub cmbPolygonLayers_Click()
    Dim pFL As IFeatureLayer
    Dim pFields As IFields
    Dim lLoop As Long
    ' Populate the fields combo boxes
    If cmbPolygonLayers.ListIndex > 0 Then
        Set pFL = FindFeatureLayerByName(cmbPolygonLayers.List(cmbPolygonLayers.ListIndex), m_Application)
        Set pFields = pFL.FeatureClass.Fields
        cmbFieldColNum.Clear
        cmbFieldID.Clear
        cmbFieldMapScale.Clear
        cmbFieldRowNum.Clear
        cmbFieldRowNum.AddItem "<None>"
        cmbFieldColNum.AddItem "<None>"
        cmbFieldMapScale.AddItem "<None>"
        For lLoop = 0 To pFields.FieldCount - 1
            If pFields.Field(lLoop).Type = esriFieldTypeString Then
                cmbFieldID.AddItem pFields.Field(lLoop).Name
            ElseIf pFields.Field(lLoop).Type = esriFieldTypeDouble Or _
                   pFields.Field(lLoop).Type = esriFieldTypeInteger Or _
                   pFields.Field(lLoop).Type = esriFieldTypeSmallInteger Or _
                   pFields.Field(lLoop).Type = esriFieldTypeSingle Then
                cmbFieldColNum.AddItem pFields.Field(lLoop).Name
                cmbFieldRowNum.AddItem pFields.Field(lLoop).Name
                cmbFieldMapScale.AddItem pFields.Field(lLoop).Name
            End If
        Next
        cmbFieldRowNum.ListIndex = 0
        cmbFieldColNum.ListIndex = 0
        cmbFieldMapScale.ListIndex = 0
    End If
    SetControlsState
End Sub

Private Sub cmdBack_Click()
    m_Step = m_Step - 1
    If m_Step < 0 Then
        m_Step = 0
    End If
    SetVisibleControls m_Step
    SetControlsState
End Sub

Private Sub cmdClose_Click()
    Set m_Application = Nothing
    Set Me.GridSettings = Nothing
    Me.Hide
End Sub

Private Sub CollateGridSettings()
    Dim pMx As IMxDocument
    Dim pCreateGrid As New clsCreateGrids
    Dim pFrameElement As IElement
    Dim sDestLayerName As String
    Dim lLoop As Long
    ' Populate class
    If (optGridIDOrder(0).Value) Then
        pCreateGrid.IdentifierOrder = Row_Column
    Else
        pCreateGrid.IdentifierOrder = Column_Row
    End If
    If (optRowIDType(0).Value) Then
        pCreateGrid.RowIDType = Alphabetical
    Else
        pCreateGrid.RowIDType = Numerical
    End If
    If (optColIDType(0).Value) Then
        pCreateGrid.ColIDType = Alphabetical
    Else
        pCreateGrid.ColIDType = Numerical
    End If
    If (optStartingIDPosition(0).Value) Then
        pCreateGrid.IDStartPositionType = TopLeft
    Else
        pCreateGrid.IDStartPositionType = LowerLeft
    End If
    If (optScaleSource(0).Value) Then
        pCreateGrid.MapScale = CDbl(lblCurrentMapScale.Caption)
    Else
        pCreateGrid.MapScale = CDbl(txtManualMapScale.Text)
    End If
    If (optGridSize(0).Value) Then
        Set pFrameElement = GetDataFrameElement(GetActiveDataFrameName(m_Application), m_Application)
        pCreateGrid.FrameWidthInPageUnits = pFrameElement.Geometry.Envelope.Width
        pCreateGrid.FrameHeightInPageUnits = pFrameElement.Geometry.Envelope.Height
    Else
        pCreateGrid.FrameWidthInPageUnits = CDbl(txtManualGridWidth.Text)
        pCreateGrid.FrameHeightInPageUnits = CDbl(txtManualGridHeight.Text)
    End If
    sDestLayerName = cmbPolygonLayers.List(cmbPolygonLayers.ListIndex)
    If optLayerSource(0).Value Then
        Set pCreateGrid.DestinationFeatureLayer = FindFeatureLayerByName(sDestLayerName, m_Application)
    End If
    pCreateGrid.StartingCoordinateLL_X = CDbl(txtStartCoordX.Text)
    pCreateGrid.StartingCoordinateLL_Y = CDbl(txtStartCoordY.Text)
    pCreateGrid.EndingCoordinateUR_X = CDbl(txtEndCoordX.Text)
    pCreateGrid.EndingCoordinateUR_Y = CDbl(txtEndCoordY.Text)
    pCreateGrid.UseUnderscore = (chkBreak.Value = vbChecked)
    pCreateGrid.FieldNameGridID = cmbFieldID.List(cmbFieldID.ListIndex)
    If cmbFieldRowNum.ListIndex > 0 Then pCreateGrid.FieldNameRowNum = cmbFieldRowNum.List(cmbFieldRowNum.ListIndex)
    If cmbFieldColNum.ListIndex > 0 Then pCreateGrid.FieldNameColNum = cmbFieldColNum.List(cmbFieldColNum.ListIndex)
    If cmbFieldMapScale.ListIndex > 0 Then pCreateGrid.FieldNameScale = cmbFieldMapScale.List(cmbFieldMapScale.ListIndex)
    pCreateGrid.NoEmptyGrids = (chkRemoveEmpties.Value = vbChecked)
    If pCreateGrid.NoEmptyGrids Then
        pCreateGrid.ClearRequiredDataLayers
        For lLoop = 0 To lstRequiredDataLayers.ListCount - 1
            If lstRequiredDataLayers.Selected(lLoop) Then
                pCreateGrid.AddRequiredDataLayer lstRequiredDataLayers.List(lLoop)
            End If
        Next
    End If
    pCreateGrid.RemoveCurrentGrids = (chkRemovePreviousGrids.Value = vbChecked)
    ' Place grid settings on Public form property (so calling function can use them)
    Set Me.GridSettings = pCreateGrid
End Sub

Private Sub cmdDatasetExtentLL_Click()
    Dim pFL As IFeatureLayer
    Dim pDatasetExtent As IEnvelope
    
    If cmbPolygonLayers.ListIndex > 0 Then
        Set pFL = FindFeatureLayerByName(cmbPolygonLayers.List(cmbPolygonLayers.ListIndex), m_Application)
        Set pDatasetExtent = GetValidExtentForLayer(pFL)
        txtStartCoordX.Text = Format(pDatasetExtent.XMin, "#,###,##0.00")
        txtStartCoordY.Text = Format(pDatasetExtent.YMin, "#,###,##0.00")
        txtEndCoordX.Text = Format(pDatasetExtent.XMax - 100, "#,###,##0.00")
        txtEndCoordY.Text = Format(pDatasetExtent.YMax - 100, "#,###,##0.00")
        SetControlsState
    End If
End Sub

Private Sub cmdLayersExtent_Click()
    Dim pMx As IMxDocument
    Dim pEnv As IEnvelope
    Dim pElement As IElement
    Dim pMapFrame As IMapFrame
    Dim pActiveView As IActiveView
    
    On Error GoTo eh
    
    Set pMx = m_Application.Document
    Set pActiveView = pMx.ActiveView
    If TypeOf pActiveView Is IPageLayout Then
        Set pElement = GetDataFrameElement(pMx.FocusMap.Name, m_Application)
        Set pMapFrame = pElement
        Set pEnv = pMapFrame.MapBounds
        Set pActiveView = pMapFrame.Map
        Set pEnv = pActiveView.FullExtent
    Else
        Set pEnv = pActiveView.FullExtent
    End If
    
    txtStartCoordX.Text = Format(pEnv.XMin, "#,###,##0.00")
    txtStartCoordY.Text = Format(pEnv.YMin, "#,###,##0.00")
    txtEndCoordX.Text = Format(pEnv.XMax, "#,###,##0.00")
    txtEndCoordY.Text = Format(pEnv.YMax, "#,###,##0.00")
    
    SetControlsState
    
    Exit Sub
eh:
    MsgBox Err.Description, , "cmdLayersExtent_Click"
End Sub

Private Sub cmdMapExtentLL_Click()
    Dim pMx As IMxDocument
    Dim pEnv As IEnvelope
    Dim pElement As IElement
    Dim pMapFrame As IMapFrame
    Dim pActiveView As IActiveView
    
    On Error GoTo eh
    
    Set pMx = m_Application.Document
    Set pActiveView = pMx.ActiveView
    If TypeOf pActiveView Is IPageLayout Then
        Set pElement = GetDataFrameElement(pMx.FocusMap.Name, m_Application)
        Set pMapFrame = pElement
        Set pEnv = pMapFrame.MapBounds
    Else
        Set pEnv = pActiveView.Extent
    End If
    
    txtStartCoordX.Text = Format(pEnv.XMin, "#,###,##0.00")
    txtStartCoordY.Text = Format(pEnv.YMin, "#,###,##0.00")
    txtEndCoordX.Text = Format(pEnv.XMax, "#,###,##0.00")
    txtEndCoordY.Text = Format(pEnv.YMax, "#,###,##0.00")
    
    SetControlsState
    
    Exit Sub
eh:
    MsgBox Err.Description, , "cmdMapExtentLL"
End Sub

Private Sub cmdNext_Click()
    Dim pMx As IMxDocument
    Dim pFeatureLayer As IFeatureLayer
    Dim pOutputFClass As IFeatureClass
    Dim pNewFields As IFields
    
    On Error GoTo eh
    ' Step
    m_Step = m_Step + 1
    ' If we're creating a new fclass, we can skip a step
    If m_Step = 1 And (optLayerSource(1).Value) Then
        m_Step = m_Step + 1
    End If
    ' If FINISH
    If m_Step >= 5 Then
        CollateGridSettings
        ' If creating a new layer
        If optLayerSource(1).Value Then
            ' Create the feature class
            Set pMx = m_Application.Document
            Set pNewFields = CreateTheFields
            Select Case m_FileType
                Case ShapeFile
                    Set pOutputFClass = NewShapeFile(m_OutputLayer, pMx.FocusMap, pNewFields)
                Case AccessFeatureClass
                    Set pOutputFClass = NewAccessFile(m_OutputLayer, _
                            m_OutputDataset, m_OutputFClass, pNewFields)
            End Select
            If pOutputFClass Is Nothing Then
                Err.Raise vbObjectError, "cmdNext", "Could not create the new output feature class."
            End If
            ' Create new layer
            Set pFeatureLayer = New FeatureLayer
            Set pFeatureLayer.FeatureClass = pOutputFClass
            pFeatureLayer.Name = pFeatureLayer.FeatureClass.AliasName
            ' Add the new layer to arcmap & reset the GridSettings object to point at it
            pMx.FocusMap.AddLayer pFeatureLayer
            Set GridSettings.DestinationFeatureLayer = pFeatureLayer
        End If
        Me.Hide
    Else
        SetVisibleControls m_Step
        SetControlsState
    End If
    
    Exit Sub
eh:
    MsgBox "Error: " & Err.Description, , "cmdNext_Click"
    m_Step = m_Step - 1
End Sub

Private Sub cmdSetNewGridLayer_Click()
  Dim pGxFilter As IGxObjectFilter
  Dim pGXBrow As IGxDialog, bFlag As Boolean
  Dim pSel As IEnumGxObject, pApp As IApplication
  
  Set pGxFilter = New GxFilter
  Set pApp = m_Application
  Set pGXBrow = New GxDialog
  Set pGXBrow.ObjectFilter = pGxFilter
  pGXBrow.Title = "Output feature class or shapefile"
  bFlag = pGXBrow.DoModalSave(pApp.hwnd)
  
  If bFlag Then
    Dim pObj As IGxObject
    Set pObj = pGXBrow.FinalLocation
    m_bIsGeoDatabase = True
    If UCase(pObj.Category) = "FOLDER" Then
      If InStr(1, pGXBrow.Name, ".shp") > 0 Then
        txtNewGridLayer.Text = pObj.FullName & "\" & pGXBrow.Name
      Else
        txtNewGridLayer.Text = pObj.FullName & "\" & pGXBrow.Name & ".shp"
      End If
      m_OutputLayer = txtNewGridLayer.Text
      m_bIsGeoDatabase = False
      m_FileType = ShapeFile
     CheckOutputFile
    Else
      Dim pLen As Long
      pLen = Len(pObj.FullName) - Len(pObj.BaseName) - 1
      txtNewGridLayer.Text = Left(pObj.FullName, pLen)
      m_OutputLayer = Left(pObj.FullName, pLen)
      m_OutputDataset = pObj.BaseName
      m_OutputFClass = pGXBrow.Name
      m_bIsGeoDatabase = True
      If UCase(pObj.Category) = "PERSONAL GEODATABASE FEATURE DATASET" Then
        m_FileType = AccessFeatureClass
      Else
        m_FileType = SDEFeatureClass
      End If
    End If
  Else
    txtNewGridLayer.Text = ""
    m_bIsGeoDatabase = False
  End If
  SetControlsState
End Sub

Private Sub Form_Load()
    Dim pMx As IMxDocument
    Dim bRenewCoordsX As Boolean
    Dim bRenewCoordsY As Boolean
    
    On Error GoTo eh
    
    Set pMx = m_Application.Document
    Me.Height = 5565
    Me.Width = 4935
    m_Step = 0
    LoadLayersComboBox
    LoadUnitsComboBox
    lblExampleID.Caption = GenerateExampleID
    lblCurrFrameName.Caption = GetActiveDataFrameName(m_Application)
    If pMx.FocusMap.MapUnits = esriUnknownUnits Then
        MsgBox "Error: The map has unknown units and therefore cannot calculate a Scale." _
            & vbCrLf & "Cannot create Map Grids at this time.", vbCritical, "Create Map Grids"
        Unload Me
        Exit Sub
    End If
    lblCurrentMapScale.Caption = Format(pMx.FocusMap.MapScale, "#,###,##0")
    Call cmdMapExtentLL_Click
    SetVisibleControls m_Step
    
    SetControlsState
    
    'Make sure the wizard stays on top
    TopMost Me
    
    Exit Sub
eh:
    MsgBox "Error loading the form: " & Err.Description & vbCrLf _
        & vbCrLf & "Attempting to continue the load...", , "MapGridManager: Form_Load "
    On Error Resume Next
    SetVisibleControls m_Step
    SetControlsState
End Sub

Private Sub LoadUnitsComboBox()
    Dim pMx As IMxDocument
    Dim sPageUnitsDesc As String
    Dim pPage As IPage
    
    On Error GoTo eh
    
    ' Init
    Set pMx = m_Application.Document
    Set pPage = pMx.PageLayout.Page
    sPageUnitsDesc = GetUnitsDescription(pPage.Units)
    cmbGridSizeUnits.Clear
    ' Add
    cmbGridSizeUnits.AddItem sPageUnitsDesc
    'cmbGridSizeUnits.AddItem "Map Units (" & sMapUnitsDesc & ")"
    ' Set page units as default
    cmbGridSizeUnits.ListIndex = 0
    
    Exit Sub
eh:
    Err.Raise vbObjectError, "LoadUnitsComboBox", "Error in LoadUnitsComboBox" & vbCrLf & Err.Description
End Sub

Private Sub LoadLayersComboBox()
    Dim pMx As IMxDocument
    Dim lLoop As Long
    Dim pFL As IFeatureLayer
    Dim pFC As IFeatureClass
    Dim sPreviousLayer  As String
    Dim lResetIndex As Long
    
    'Init
    Set pMx = m_Application.Document
'    If cmbPolygonLayers.ListCount > 0 Then
'        sPreviousLayer = cmbPolygonLayers.List(cmbPolygonLayers.ListIndex)
'    End If
    cmbPolygonLayers.Clear
    lstRequiredDataLayers.Clear
    cmbPolygonLayers.AddItem "<Not Set>"
    ' For all layers
    For lLoop = 0 To pMx.FocusMap.LayerCount - 1
        ' If a feature class
        If TypeOf pMx.FocusMap.Layer(lLoop) Is IFeatureLayer Then
            Set pFL = pMx.FocusMap.Layer(lLoop)
            Set pFC = pFL.FeatureClass
            ' If a polygon layer
            If pFC.ShapeType = esriGeometryPolygon Then
                ' Add to combo box
                cmbPolygonLayers.AddItem pFL.Name
'                If pFL.Name = sPreviousLayer Then
'                    lResetIndex = (cmbPolygonLayers.ListCount - 1)
'                End If
            End If
            lstRequiredDataLayers.AddItem pFL.Name
        End If
    Next
    'cmbPolygonLayers.ListIndex = lResetIndex
End Sub

Private Sub SetCurrentMapScaleCaption()
    Dim pMx As IMxDocument
    On Error GoTo eh
    Set pMx = m_Application.Document
    lblCurrentMapScale.Caption = Format(pMx.FocusMap.MapScale, "#,###,##0")
    Exit Sub
eh:
    lblCurrentMapScale.Caption = "<Scale Unknown>"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set m_Application = Nothing
    Set GridSettings = Nothing
End Sub


Private Sub lstRequiredDataLayers_Click()
    SetControlsState
End Sub

Private Sub optColIDType_Click(Index As Integer)
    lblExampleID.Caption = GenerateExampleID
    SetControlsState
End Sub

Private Sub optGridIDOrder_Click(Index As Integer)
    lblExampleID.Caption = GenerateExampleID
    SetControlsState
End Sub

Private Sub optGridSize_Click(Index As Integer)
    Dim pMx As IMxDocument
    Set pMx = m_Application.Document
    lblCurrFrameName.Caption = pMx.FocusMap.Name
    SetControlsState
End Sub

Private Sub optLayerSource_Click(Index As Integer)
    ' If creating a new fclass to hold the grids
    If Index = 1 Then
        ' Set the field names (will be created automatically)
        cmbFieldID.Clear
        cmbFieldRowNum.Clear
        cmbFieldColNum.Clear
        cmbFieldMapScale.Clear
        cmbFieldID.AddItem "<None>"
        cmbFieldRowNum.AddItem "<None>"
        cmbFieldColNum.AddItem "<None>"
        cmbFieldMapScale.AddItem "<None>"
        cmbFieldID.AddItem c_DefaultFld_GridID
        cmbFieldRowNum.AddItem c_DefaultFld_RowNum
        cmbFieldColNum.AddItem c_DefaultFld_ColNum
        cmbFieldMapScale.AddItem c_DefaultFld_Scale
        cmbFieldID.ListIndex = 1
        cmbFieldRowNum.ListIndex = 1
        cmbFieldColNum.ListIndex = 1
        cmbFieldMapScale.ListIndex = 1
    End If
    SetControlsState
End Sub

Private Sub optRowIDType_Click(Index As Integer)
    lblExampleID.Caption = GenerateExampleID
    SetControlsState
End Sub

Private Function GenerateExampleID() As String
    Dim sRow As String, sCol As String
    If optStartingIDPosition(0).Value Then  'Top left
        If (optRowIDType(0).Value) Then
            sRow = "A"
        Else
            sRow = "1"
        End If
        If (optColIDType(0).Value) Then
            sCol = "C"
        Else
            sCol = "3"
        End If
    Else                                    ' Lower left
        If (optRowIDType(0).Value) Then
            sRow = "C"
        Else
            sRow = "3"
        End If
        If (optColIDType(0).Value) Then
            sCol = "C"
        Else
            sCol = "3"
        End If
    End If
    If (optGridIDOrder(0).Value) Then
        If chkBreak.Value = vbChecked Then
            GenerateExampleID = sRow & "_" & sCol
        Else
            GenerateExampleID = sRow & sCol
        End If
    Else
        If chkBreak.Value = vbChecked Then
            GenerateExampleID = sCol & "_" & sRow
        Else
            GenerateExampleID = sCol & sRow
        End If
    End If
End Function

Private Sub optScaleSource_Click(Index As Integer)
    If Index = 0 Then
        SetCurrentMapScaleCaption
    End If
    SetControlsState
End Sub

Private Sub optStartingIDPosition_Click(Index As Integer)
    lblExampleID.Caption = GenerateExampleID
    SetControlsState
End Sub

Private Sub txtEndCoordX_Change()
    SetControlsState
End Sub

Private Sub txtEndCoordX_KeyPress(KeyAscii As Integer)
    ' If a non-numeric (that is not a decimal point)
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
     And KeyAscii <> Asc(".") _
     And Chr(KeyAscii) <> vbBack Then
        ' Do not allow this button to work
        KeyAscii = 0
    ' If a decimal point, make sure we only ever get one of them
    ElseIf KeyAscii = Asc(".") Then
        If InStr(txtEndCoordX.Text, ".") > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtEndCoordY_Change()
    SetControlsState
End Sub

Private Sub txtEndCoordY_KeyPress(KeyAscii As Integer)
    ' If a non-numeric (that is not a decimal point)
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
     And KeyAscii <> Asc(".") _
     And Chr(KeyAscii) <> vbBack Then
        ' Do not allow this button to work
        KeyAscii = 0
    ' If a decimal point, make sure we only ever get one of them
    ElseIf KeyAscii = Asc(".") Then
        If InStr(txtEndCoordY.Text, ".") > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtManualGridHeight_Change()
    SetControlsState
End Sub

Private Sub txtManualGridHeight_KeyPress(KeyAscii As Integer)
    ' If a non-numeric (that is not a decimal point)
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
     And KeyAscii <> Asc(".") _
     And Chr(KeyAscii) <> vbBack Then
        ' Do not allow this button to work
        KeyAscii = 0
    ' If a decimal point, make sure we only ever get one of them
    ElseIf KeyAscii = Asc(".") Then
        If InStr(txtManualGridHeight.Text, ".") > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtManualGridWidth_Change()
    SetControlsState
End Sub

Private Sub txtManualGridWidth_KeyPress(KeyAscii As Integer)
    ' If a non-numeric (that is not a decimal point)
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
     And KeyAscii <> Asc(".") _
     And Chr(KeyAscii) <> vbBack Then
        ' Do not allow this button to work
        KeyAscii = 0
    ' If a decimal point, make sure we only ever get one of them
    ElseIf KeyAscii = Asc(".") Then
        If InStr(txtManualGridWidth.Text, ".") > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtManualMapScale_Change()
    SetControlsState
End Sub

Private Sub txtManualMapScale_KeyPress(KeyAscii As Integer)
    ' If a non-numeric (that is not a decimal point)
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
     And KeyAscii <> Asc(".") _
     And Chr(KeyAscii) <> vbBack Then
        ' Do not allow this button to work
        KeyAscii = 0
    ' If a decimal point, make sure we only ever get one of them
    ElseIf KeyAscii = Asc(".") Then
        If InStr(txtManualMapScale.Text, ".") > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Public Sub Tickle()
    Call Form_Load
End Sub

Private Sub txtStartCoordX_Change()
    SetControlsState
End Sub

Private Sub txtStartCoordX_KeyPress(KeyAscii As Integer)
    ' If a non-numeric (that is not a decimal point)
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
     And KeyAscii <> Asc(".") _
     And Chr(KeyAscii) <> vbBack Then
        ' Do not allow this button to work
        KeyAscii = 0
    ' If a decimal point, make sure we only ever get one of them
    ElseIf KeyAscii = Asc(".") Then
        If InStr(txtStartCoordX.Text, ".") > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtStartCoordY_Change()
    SetControlsState
End Sub

Private Sub txtStartCoordY_KeyPress(KeyAscii As Integer)
    ' If a non-numeric (that is not a decimal point)
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
     And KeyAscii <> Asc(".") _
     And Chr(KeyAscii) <> vbBack Then
        ' Do not allow this button to work
        KeyAscii = 0
    ' If a decimal point, make sure we only ever get one of them
    ElseIf KeyAscii = Asc(".") Then
        If InStr(txtStartCoordY.Text, ".") > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub SetVisibleControls(iStep As Integer)
    ' Hide all
    fraAttributes.Visible = False
    fraDataFrameSize.Visible = False
    fraDestinationFeatureClass.Visible = False
    fraGridIDs.Visible = False
    fraScaleStart.Visible = False
    ' Show applicable frame, set top/left
    Select Case iStep
        Case 0:
            fraDestinationFeatureClass.Visible = True
            fraDestinationFeatureClass.Top = 0
            fraDestinationFeatureClass.Left = 0
        Case 1:
            fraAttributes.Visible = True
            fraAttributes.Top = 0
            fraAttributes.Left = 0
        Case 2:
            fraScaleStart.Visible = True
            fraScaleStart.Top = 0
            fraScaleStart.Left = 0
        Case 3:
            fraDataFrameSize.Visible = True
            fraDataFrameSize.Top = 0
            fraDataFrameSize.Left = 0
        Case 4:
            fraGridIDs.Visible = True
            fraGridIDs.Top = 0
            fraGridIDs.Left = 0
        Case Else:
            MsgBox "Invalid Step Value."
    End Select
End Sub

Private Sub CheckOutputFile()
    'Check the output option
    If txtNewGridLayer.Text <> "" Then
        If DoesShapeFileExist(txtNewGridLayer.Text) Then
            MsgBox "Shape file name already being used!!!"
            txtNewGridLayer.Text = ""
        End If
    End If
End Sub

Private Function CreateTheFields() As IFields
    Dim newField As IField
    Dim newFieldEdit As IFieldEdit
    Dim pNewFields As IFields
    Dim pFieldsEdit As IFieldsEdit
    Dim pGeomDef As IGeometryDef
    Dim pGeomDefEdit As IGeometryDefEdit
    Dim pMx As IMxDocument
    
    ' Init
    Set pNewFields = New Fields
    Set pFieldsEdit = pNewFields
    Set pMx = m_Application.Document
    ' Field: OID
    Set newField = New Field
    Set newFieldEdit = newField
    With newFieldEdit
        .Name = "OID"
        .Type = esriFieldTypeOID
        .AliasName = "Object ID"
        .IsNullable = False
    End With
    pFieldsEdit.AddField newField
    'Set pFieldsEdit.Field(0) = pFieldEdit
    
'    ' Field: SHAPE
'    Set newField = New esriCore.Field
'    Set newFieldEdit = newField
'    newFieldEdit.Name = c_DefaultFld_Shape
'    newFieldEdit.Type = esriFieldTypeGeometry
'    Set pGeomDef = New GeometryDef
'    Set pGeomDefEdit = pGeomDef
'    With pGeomDefEdit
'        .GeometryType = esriGeometryPolygon
'        Set .SpatialReference = pMx.FocusMap.SpatialReference ' New UnknownCoordinateSystem
'    End With
'    Set newFieldEdit.GeometryDef = pGeomDef
'    pFieldsEdit.AddField newField
    ' Field: GRID IDENTIFIER
    Set newField = New Field
    Set newFieldEdit = newField
    With newFieldEdit
      .Name = c_DefaultFld_GridID
      .AliasName = "GridIdentifier"
      .Type = esriFieldTypeString
      .IsNullable = True
      .Length = 50
    End With
    pFieldsEdit.AddField newField
    ' Field: ROW NUMBER
    Set newField = New Field
    Set newFieldEdit = newField
    With newFieldEdit
      .Name = c_DefaultFld_RowNum
      .AliasName = "Row Number"
      .Type = esriFieldTypeInteger
      .IsNullable = True
    End With
    pFieldsEdit.AddField newField
    ' Field: COLUMN NUMBER
    Set newField = New Field
    Set newFieldEdit = newField
    With newFieldEdit
      .Name = c_DefaultFld_ColNum
      .AliasName = "Column Number"
      .Type = esriFieldTypeInteger
      .IsNullable = True
    End With
    pFieldsEdit.AddField newField
    ' Field: SCALE
    Set newField = New Field
    Set newFieldEdit = newField
    With newFieldEdit
      .Name = c_DefaultFld_Scale
      .AliasName = "Plot Scale"
      .Type = esriFieldTypeDouble
      .IsNullable = True
      .Precision = 18
      .Scale = 11
    End With
    pFieldsEdit.AddField newField
    ' Return
    Set CreateTheFields = pFieldsEdit
End Function

