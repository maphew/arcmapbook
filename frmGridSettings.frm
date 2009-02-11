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

' Copyright 2006 ESRI
'
' All rights reserved under the copyright laws of the United States
' and applicable international laws, treaties, and conventions.
'
' You may freely redistribute and use this sample code, with or
' without modification, provided you include the original copyright
' notice and use restrictions.
'
' See use restrictions at /arcgis/developerkit/userestrictions.

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
61:     If Len(lblCurrentMapScale.Caption) = 0 Then lblCurrentMapScale.Caption = "0"
62:     If Len(txtManualMapScale.Text) = 0 Then
63:         dScale = 0
64:     Else
65:         dScale = CDbl(txtManualMapScale.Text)
66:     End If
67:     If Len(txtManualGridHeight.Text) = 0 Then
68:         dGHeight = 0
69:     Else
70:         dGHeight = CDbl(txtManualGridHeight.Text)
71:     End If
72:     If Len(txtManualGridWidth.Text) = 0 Then
73:         dGWidth = 0
74:     Else
75:         dGWidth = CDbl(txtManualGridWidth.Text)
76:     End If
77:     If Len(txtStartCoordX.Text) = 0 Then
78:         dStartX = 0
79:     Else
80:         dStartX = CDbl(txtStartCoordX.Text)
81:     End If
82:     If Len(txtStartCoordY.Text) = 0 Then
83:         dStartY = 0
84:     Else
85:         dStartY = CDbl(txtStartCoordY.Text)
86:     End If
87:     If Len(txtEndCoordX.Text) = 0 Then
88:         dEndX = 0
89:     Else
90:         dEndX = CDbl(txtEndCoordX.Text)
91:     End If
92:     If Len(txtEndCoordY.Text) = 0 Then
93:         dEndY = 0
94:     Else
95:         dEndY = CDbl(txtEndCoordY.Text)
96:     End If
97: i = 1

    ' Calc values
100:     bValidScale = (optScaleSource(0).value And CDbl(lblCurrentMapScale.Caption) > 0) Or _
                  (optScaleSource(1).value And dScale > 0)
102:     bValidSize = (optGridSize(0).value) Or _
                 (optGridSize(1).value And dGHeight > 0 And dGWidth > 0)
104:     bCreatingNewFClass = optLayerSource(1).value
105:     bNewFClassSet = (Len(txtNewGridLayer.Text) > 0)
106:     bValidTarget = (cmbPolygonLayers.ListIndex > 0) Or (bCreatingNewFClass And bNewFClassSet)
107:     bValidIDField = (cmbFieldID.ListIndex >= 0)
108:     bValidReqdLayers = (chkRemoveEmpties.value = vbUnchecked) Or _
                       (chkRemoveEmpties.value = vbChecked And lstRequiredDataLayers.SelCount > 0)
110: i = 2
111:     If bValidTarget And (Not bCreatingNewFClass) Then
112:         Set pFL = FindFeatureLayerByName(cmbPolygonLayers.List(cmbPolygonLayers.ListIndex), m_Application)
113:         If pFL.FeatureClass.FeatureDataset Is Nothing Then
114:             bValidStart = True
115:             bValidEnd = True
116:         Else
117:             Set pDatasetExtent = GetValidExtentForLayer(pFL)
118:             bValidStart = ((dStartX >= pDatasetExtent.XMin) And (dStartX <= pDatasetExtent.XMax)) _
                            And _
                          ((dStartY >= pDatasetExtent.YMin) And (dStartY <= pDatasetExtent.YMax))
121:             bValidEnd = ((dEndX >= pDatasetExtent.XMin) And (dEndX <= pDatasetExtent.XMax)) _
                            And _
                        ((dEndY >= pDatasetExtent.YMin) And (dEndY <= pDatasetExtent.YMax)) _
                            And _
                        ((dEndX > dStartX) And (dEndY > dStartY))
126:         End If
127:     ElseIf bValidTarget And bCreatingNewFClass Then
128:         bValidStart = True
129:         bValidEnd = True
130:     End If
131:     bDuplicateFieldsSelected = (cmbFieldRowNum.ListIndex > 0 And cmbFieldRowNum.ListIndex = cmbFieldColNum.ListIndex) _
                            Or (cmbFieldRowNum.ListIndex > 0 And cmbFieldRowNum.ListIndex = cmbFieldMapScale.ListIndex) _
                            Or (cmbFieldColNum.ListIndex > 0 And cmbFieldColNum.ListIndex = cmbFieldMapScale.ListIndex)
134: i = 3
    
    ' Set states
    Select Case m_Step
        Case 0:     ' Set the target feature layer
139:             cmdBack.Enabled = False
140:             cmdNext.Enabled = bValidTarget And bValidReqdLayers
141:             cmdNext.Caption = "Next >"
142:             cmbPolygonLayers.Enabled = Not bCreatingNewFClass
        Case 1:     ' Set the fields to populate
144:             cmdBack.Enabled = True
145:             cmdNext.Enabled = (bValidIDField And Not bDuplicateFieldsSelected)
146:             cmbFieldID.Enabled = Not bCreatingNewFClass
147:             cmbFieldRowNum.Enabled = Not bCreatingNewFClass
148:             cmbFieldColNum.Enabled = Not bCreatingNewFClass
149:             cmbFieldMapScale.Enabled = Not bCreatingNewFClass
        Case 2:     ' Set the scale / starting_coords
151:             cmdBack.Enabled = True
152:             cmdNext.Enabled = bValidScale And bValidStart And bValidEnd
153:             If Not bCreatingNewFClass Then
154:                 cmdDatasetExtentLL.Enabled = Not (pFL.FeatureClass.FeatureDataset Is Nothing)
155:             Else
156:                 cmdDatasetExtentLL.Enabled = False
157:             End If
        Case 3:     ' Set the dataframe properties
159:             cmdBack.Enabled = True
160:             cmdNext.Enabled = bValidSize
161:             cmdNext.Caption = "Next >"
        Case 4:     ' Set the ID values
163:             cmdBack.Enabled = True
164:             cmdNext.Enabled = True
165:             cmdNext.Caption = "Finish"
        Case Else:
167:             cmdBack.Enabled = False
168:             cmdNext.Enabled = False
169:     End Select
170: i = 4
    
172:     txtManualMapScale.Enabled = optScaleSource(1).value
173:     txtManualGridWidth.Enabled = optGridSize(1).value
174:     txtManualGridHeight.Enabled = optGridSize(1).value
175:     cmbGridSizeUnits.Enabled = optGridSize(1).value
    ' Set display
177:     If bValidStart Then
178:         txtStartCoordX.ForeColor = (&H0)    ' Black
179:         txtStartCoordY.ForeColor = (&H0)
180:     Else
181:         txtStartCoordX.ForeColor = (&HFF)   ' Red
182:         txtStartCoordY.ForeColor = (&HFF)
183:     End If
184:     If bValidEnd Then
185:         txtEndCoordX.ForeColor = (&H0)      ' Black
186:         txtEndCoordY.ForeColor = (&H0)
187:     Else
188:         txtEndCoordX.ForeColor = (&HFF)     ' Red
189:         txtEndCoordY.ForeColor = (&HFF)
190:     End If
191:     If optScaleSource(1).value Then
192:         If bValidScale Then
193:             txtManualMapScale.ForeColor = (&H0)      ' Black
194:         Else
195:             txtManualMapScale.ForeColor = (&HFF)     ' Red
196:         End If
197:     End If
198:     If optGridSize(1).value Then
199:         If bValidSize Then
200:             txtManualGridWidth.ForeColor = (&H0)      ' Black
201:             txtManualGridHeight.ForeColor = (&H0)
202:         Else
203:             txtManualGridWidth.ForeColor = (&HFF)     ' Red
204:             txtManualGridHeight.ForeColor = (&HFF)
205:         End If
206:     End If
    
    Exit Sub
209:     Resume
eh:
211:     MsgBox Err.Description, vbExclamation, "SetControlsState " & i
End Sub

Private Sub chkBreak_Click()
215:     lblExampleID.Caption = GenerateExampleID
End Sub

Private Sub chkRemoveEmpties_Click()
219:     SetControlsState
End Sub

Private Sub cmbFieldColNum_Click()
223:     SetControlsState
End Sub

Private Sub cmbFieldID_Click()
227:     SetControlsState
End Sub

Private Sub cmbFieldMapScale_Click()
231:     SetControlsState
End Sub

Private Sub cmbFieldRowNum_Click()
235:     SetControlsState
End Sub

Private Sub cmbPolygonLayers_Click()
    Dim pFL As IFeatureLayer
    Dim pFields As IFields
    Dim lLoop As Long
    ' Populate the fields combo boxes
243:     If cmbPolygonLayers.ListIndex > 0 Then
244:         Set pFL = FindFeatureLayerByName(cmbPolygonLayers.List(cmbPolygonLayers.ListIndex), m_Application)
245:         Set pFields = pFL.FeatureClass.Fields
246:         cmbFieldColNum.Clear
247:         cmbFieldID.Clear
248:         cmbFieldMapScale.Clear
249:         cmbFieldRowNum.Clear
250:         cmbFieldRowNum.AddItem "<None>"
251:         cmbFieldColNum.AddItem "<None>"
252:         cmbFieldMapScale.AddItem "<None>"
253:         For lLoop = 0 To pFields.FieldCount - 1
254:             If pFields.Field(lLoop).Type = esriFieldTypeString Then
255:                 cmbFieldID.AddItem pFields.Field(lLoop).Name
256:             ElseIf pFields.Field(lLoop).Type = esriFieldTypeDouble Or _
                   pFields.Field(lLoop).Type = esriFieldTypeInteger Or _
                   pFields.Field(lLoop).Type = esriFieldTypeSmallInteger Or _
                   pFields.Field(lLoop).Type = esriFieldTypeSingle Then
260:                 cmbFieldColNum.AddItem pFields.Field(lLoop).Name
261:                 cmbFieldRowNum.AddItem pFields.Field(lLoop).Name
262:                 cmbFieldMapScale.AddItem pFields.Field(lLoop).Name
263:             End If
264:         Next
265:         cmbFieldRowNum.ListIndex = 0
266:         cmbFieldColNum.ListIndex = 0
267:         cmbFieldMapScale.ListIndex = 0
268:     End If
269:     SetControlsState
End Sub

Private Sub cmdBack_Click()
273:     m_Step = m_Step - 1
274:     If m_Step < 0 Then
275:         m_Step = 0
276:     End If
277:     SetVisibleControls m_Step
278:     SetControlsState
End Sub

Private Sub cmdClose_Click()
282:     Set m_Application = Nothing
283:     Set Me.GridSettings = Nothing
284:     Me.Hide
End Sub

Private Sub CollateGridSettings()
    Dim pMx As IMxDocument
    Dim pCreateGrid As New clsCreateGrids
    Dim pFrameElement As IElement
    Dim sDestLayerName As String
    Dim lLoop As Long
    ' Populate class
294:     If (optGridIDOrder(0).value) Then
295:         pCreateGrid.IdentifierOrder = Row_Column
296:     Else
297:         pCreateGrid.IdentifierOrder = Column_Row
298:     End If
299:     If (optRowIDType(0).value) Then
300:         pCreateGrid.RowIDType = Alphabetical
301:     Else
302:         pCreateGrid.RowIDType = Numerical
303:     End If
304:     If (optColIDType(0).value) Then
305:         pCreateGrid.ColIDType = Alphabetical
306:     Else
307:         pCreateGrid.ColIDType = Numerical
308:     End If
309:     If (optStartingIDPosition(0).value) Then
310:         pCreateGrid.IDStartPositionType = TopLeft
311:     Else
312:         pCreateGrid.IDStartPositionType = LowerLeft
313:     End If
314:     If (optScaleSource(0).value) Then
315:         pCreateGrid.MapScale = CDbl(lblCurrentMapScale.Caption)
316:     Else
317:         pCreateGrid.MapScale = CDbl(txtManualMapScale.Text)
318:     End If
319:     If (optGridSize(0).value) Then
320:         Set pFrameElement = GetDataFrameElement(GetActiveDataFrameName(m_Application), m_Application)
321:         pCreateGrid.FrameWidthInPageUnits = pFrameElement.Geometry.Envelope.Width
322:         pCreateGrid.FrameHeightInPageUnits = pFrameElement.Geometry.Envelope.Height
323:     Else
324:         pCreateGrid.FrameWidthInPageUnits = CDbl(txtManualGridWidth.Text)
325:         pCreateGrid.FrameHeightInPageUnits = CDbl(txtManualGridHeight.Text)
326:     End If
327:     sDestLayerName = cmbPolygonLayers.List(cmbPolygonLayers.ListIndex)
328:     If optLayerSource(0).value Then
329:         Set pCreateGrid.DestinationFeatureLayer = FindFeatureLayerByName(sDestLayerName, m_Application)
330:     End If
331:     pCreateGrid.StartingCoordinateLL_X = CDbl(txtStartCoordX.Text)
332:     pCreateGrid.StartingCoordinateLL_Y = CDbl(txtStartCoordY.Text)
333:     pCreateGrid.EndingCoordinateUR_X = CDbl(txtEndCoordX.Text)
334:     pCreateGrid.EndingCoordinateUR_Y = CDbl(txtEndCoordY.Text)
335:     pCreateGrid.UseUnderscore = (chkBreak.value = vbChecked)
336:     pCreateGrid.FieldNameGridID = cmbFieldID.List(cmbFieldID.ListIndex)
337:     If cmbFieldRowNum.ListIndex > 0 Then pCreateGrid.FieldNameRowNum = cmbFieldRowNum.List(cmbFieldRowNum.ListIndex)
338:     If cmbFieldColNum.ListIndex > 0 Then pCreateGrid.FieldNameColNum = cmbFieldColNum.List(cmbFieldColNum.ListIndex)
339:     If cmbFieldMapScale.ListIndex > 0 Then pCreateGrid.FieldNameScale = cmbFieldMapScale.List(cmbFieldMapScale.ListIndex)
340:     pCreateGrid.NoEmptyGrids = (chkRemoveEmpties.value = vbChecked)
341:     If pCreateGrid.NoEmptyGrids Then
342:         pCreateGrid.ClearRequiredDataLayers
343:         For lLoop = 0 To lstRequiredDataLayers.ListCount - 1
344:             If lstRequiredDataLayers.Selected(lLoop) Then
345:                 pCreateGrid.AddRequiredDataLayer lstRequiredDataLayers.List(lLoop)
346:             End If
347:         Next
348:     End If
349:     pCreateGrid.RemoveCurrentGrids = (chkRemovePreviousGrids.value = vbChecked)
    ' Place grid settings on Public form property (so calling function can use them)
351:     Set Me.GridSettings = pCreateGrid
End Sub

Private Sub cmdDatasetExtentLL_Click()
    Dim pFL As IFeatureLayer
    Dim pDatasetExtent As IEnvelope
    
358:     If cmbPolygonLayers.ListIndex > 0 Then
359:         Set pFL = FindFeatureLayerByName(cmbPolygonLayers.List(cmbPolygonLayers.ListIndex), m_Application)
360:         Set pDatasetExtent = GetValidExtentForLayer(pFL)
361:         txtStartCoordX.Text = Format(pDatasetExtent.XMin, "#,###,##0.00")
362:         txtStartCoordY.Text = Format(pDatasetExtent.YMin, "#,###,##0.00")
363:         txtEndCoordX.Text = Format(pDatasetExtent.XMax - 100, "#,###,##0.00")
364:         txtEndCoordY.Text = Format(pDatasetExtent.YMax - 100, "#,###,##0.00")
365:         SetControlsState
366:     End If
End Sub

Private Sub cmdLayersExtent_Click()
    Dim pMx As IMxDocument
    Dim pEnv As IEnvelope
    Dim pElement As IElement
    Dim pMapFrame As IMapFrame
    Dim pActiveView As IActiveView
    
    On Error GoTo eh
    
378:     Set pMx = m_Application.Document
379:     Set pActiveView = pMx.ActiveView
380:     If TypeOf pActiveView Is IPageLayout Then
381:         Set pElement = GetDataFrameElement(pMx.FocusMap.Name, m_Application)
382:         Set pMapFrame = pElement
383:         Set pEnv = pMapFrame.MapBounds
384:         Set pActiveView = pMapFrame.Map
385:         Set pEnv = pActiveView.FullExtent
386:     Else
387:         Set pEnv = pActiveView.FullExtent
388:     End If
    
390:     txtStartCoordX.Text = Format(pEnv.XMin, "#,###,##0.00")
391:     txtStartCoordY.Text = Format(pEnv.YMin, "#,###,##0.00")
392:     txtEndCoordX.Text = Format(pEnv.XMax, "#,###,##0.00")
393:     txtEndCoordY.Text = Format(pEnv.YMax, "#,###,##0.00")
    
395:     SetControlsState
    
    Exit Sub
eh:
399:     MsgBox Err.Description, , "cmdLayersExtent_Click"
End Sub

Private Sub cmdMapExtentLL_Click()
    Dim pMx As IMxDocument
    Dim pEnv As IEnvelope
    Dim pElement As IElement
    Dim pMapFrame As IMapFrame
    Dim pActiveView As IActiveView
    
    On Error GoTo eh
    
411:     Set pMx = m_Application.Document
412:     Set pActiveView = pMx.ActiveView
413:     If TypeOf pActiveView Is IPageLayout Then
414:         Set pElement = GetDataFrameElement(pMx.FocusMap.Name, m_Application)
415:         Set pMapFrame = pElement
416:         Set pEnv = pMapFrame.MapBounds
417:     Else
418:         Set pEnv = pActiveView.Extent
419:     End If
    
421:     txtStartCoordX.Text = Format(pEnv.XMin, "#,###,##0.00")
422:     txtStartCoordY.Text = Format(pEnv.YMin, "#,###,##0.00")
423:     txtEndCoordX.Text = Format(pEnv.XMax, "#,###,##0.00")
424:     txtEndCoordY.Text = Format(pEnv.YMax, "#,###,##0.00")
    
426:     SetControlsState
    
    Exit Sub
eh:
430:     MsgBox Err.Description, , "cmdMapExtentLL"
End Sub

Private Sub cmdNext_Click()
    Dim pMx As IMxDocument
    Dim pFeatureLayer As IFeatureLayer
    Dim pOutputFClass As IFeatureClass
    Dim pNewFields As IFields
    
    On Error GoTo eh
    ' Step
441:     m_Step = m_Step + 1
    ' If we're creating a new fclass, we can skip a step
443:     If m_Step = 1 And (optLayerSource(1).value) Then
444:         m_Step = m_Step + 1
445:     End If
    ' If FINISH
447:     If m_Step >= 5 Then
448:         CollateGridSettings
        ' If creating a new layer
450:         If optLayerSource(1).value Then
            ' Create the feature class
452:             Set pMx = m_Application.Document
453:             Set pNewFields = CreateTheFields
            Select Case m_FileType
                Case ShapeFile
456:                     Set pOutputFClass = NewShapeFile(m_OutputLayer, pMx.FocusMap, pNewFields)
                Case AccessFeatureClass
458:                     Set pOutputFClass = NewAccessFile(m_OutputLayer, _
                            m_OutputDataset, m_OutputFClass, pNewFields)
460:             End Select
461:             If pOutputFClass Is Nothing Then
462:                 Err.Raise vbObjectError, "cmdNext", "Could not create the new output feature class."
463:             End If
            ' Create new layer
465:             Set pFeatureLayer = New FeatureLayer
466:             Set pFeatureLayer.FeatureClass = pOutputFClass
467:             pFeatureLayer.Name = pFeatureLayer.FeatureClass.AliasName
            ' Add the new layer to arcmap & reset the GridSettings object to point at it
469:             pMx.FocusMap.AddLayer pFeatureLayer
470:             Set GridSettings.DestinationFeatureLayer = pFeatureLayer
471:         End If
472:         Me.Hide
473:     Else
474:         SetVisibleControls m_Step
475:         SetControlsState
476:     End If
    
    Exit Sub
eh:
480:     MsgBox "cmdNext_Click - " & Erl & " - " & Err.Description
481:     m_Step = m_Step - 1
End Sub

Private Sub cmdSetNewGridLayer_Click()
On Error GoTo ErrHand:
  Dim pGxFilter As IGxObjectFilter
  Dim pGXBrow As IGxDialog, bFlag As Boolean
  Dim pSel As IEnumGxObject, pApp As IApplication
  
490:   Set pGxFilter = New GxFilter
491:   Set pApp = m_Application
492:   Set pGXBrow = New GxDialog
493:   Set pGXBrow.ObjectFilter = pGxFilter
494:   pGXBrow.Title = "Output feature class or shapefile"
495:   bFlag = pGXBrow.DoModalSave(pApp.hwnd)
  
497:   If bFlag Then
    Dim pObj As IGxObject
499:     Set pObj = pGXBrow.FinalLocation
500:     m_bIsGeoDatabase = True
501:     If UCase(pObj.Category) = "FOLDER" Then
502:       If InStr(1, pGXBrow.Name, ".shp") > 0 Then
503:         txtNewGridLayer.Text = pObj.FullName & "\" & pGXBrow.Name
504:       Else
505:         txtNewGridLayer.Text = pObj.FullName & "\" & pGXBrow.Name & ".shp"
506:       End If
507:       m_OutputLayer = txtNewGridLayer.Text
508:       m_bIsGeoDatabase = False
509:       m_FileType = ShapeFile
510:      CheckOutputFile
511:     Else
      Dim pLen As Long
513:       pLen = Len(pObj.FullName) - Len(pObj.BaseName) - 1
514:       txtNewGridLayer.Text = Left(pObj.FullName, pLen)
515:       m_OutputLayer = Left(pObj.FullName, pLen)
516:       m_OutputDataset = pObj.BaseName
517:       m_OutputFClass = pGXBrow.Name
518:       m_bIsGeoDatabase = True
519:       If UCase(pObj.Category) = "PERSONAL GEODATABASE FEATURE DATASET" Then
520:         m_FileType = AccessFeatureClass
521:       Else
522:         m_FileType = SDEFeatureClass
523:       End If
524:     End If
525:   Else
526:     txtNewGridLayer.Text = ""
527:     m_bIsGeoDatabase = False
528:   End If
529:   SetControlsState
  
  Exit Sub
ErrHand:
533:   MsgBox "cmdSetNewGridLayer_Click - " & Erl & " - " & Err.Description
End Sub

Private Sub Form_Load()
    Dim pMx As IMxDocument
    Dim bRenewCoordsX As Boolean
    Dim bRenewCoordsY As Boolean
    
    On Error GoTo eh
    
543:     Set pMx = m_Application.Document
544:     Me.Height = 5665
545:     Me.Width = 4935
546:     m_Step = 0
547:     LoadLayersComboBox
548:     LoadUnitsComboBox
549:     lblExampleID.Caption = GenerateExampleID
550:     lblCurrFrameName.Caption = GetActiveDataFrameName(m_Application)
551:     If pMx.FocusMap.MapUnits = esriUnknownUnits Then
552:         MsgBox "Error: The map has unknown units and therefore cannot calculate a Scale." _
            & vbCrLf & "Cannot create Map Grids at this time.", vbCritical, "Create Map Grids"
554:         Unload Me
        Exit Sub
556:     End If
557:     lblCurrentMapScale.Caption = Format(pMx.FocusMap.MapScale, "#,###,##0")
558:     Call cmdMapExtentLL_Click
559:     SetVisibleControls m_Step
    
561:     SetControlsState
    
    'Make sure the wizard stays on top
564:     TopMost Me
    
    Exit Sub
eh:
568:     MsgBox "Error loading the form: " & Erl & " - " & Err.Description & vbCrLf _
        & vbCrLf & "Attempting to continue the load...", , "MapGridManager: Form_Load "
    On Error Resume Next
571:     SetVisibleControls m_Step
572:     SetControlsState
End Sub

Private Sub LoadUnitsComboBox()
    Dim pMx As IMxDocument
    Dim sPageUnitsDesc As String
    Dim pPage As IPage
    
    On Error GoTo eh
    
    ' Init
583:     Set pMx = m_Application.Document
584:     Set pPage = pMx.PageLayout.Page
585:     sPageUnitsDesc = GetUnitsDescription(pPage.Units)
586:     cmbGridSizeUnits.Clear
    ' Add
588:     cmbGridSizeUnits.AddItem sPageUnitsDesc
    'cmbGridSizeUnits.AddItem "Map Units (" & sMapUnitsDesc & ")"
    ' Set page units as default
591:     cmbGridSizeUnits.ListIndex = 0
    
    Exit Sub
eh:
595:     Err.Raise vbObjectError, "LoadUnitsComboBox", "Error in LoadUnitsComboBox" & vbCrLf & Err.Description
End Sub

Private Sub LoadLayersComboBox()
    Dim pMx As IMxDocument
    Dim lLoop As Long
    Dim pFL As IFeatureLayer
    Dim pFC As IFeatureClass
    Dim sPreviousLayer  As String
    Dim lResetIndex As Long
    
    'Init
607:     Set pMx = m_Application.Document
'    If cmbPolygonLayers.ListCount > 0 Then
'        sPreviousLayer = cmbPolygonLayers.List(cmbPolygonLayers.ListIndex)
'    End If
611:     cmbPolygonLayers.Clear
612:     lstRequiredDataLayers.Clear
613:     cmbPolygonLayers.AddItem "<Not Set>"
    ' For all layers
615:     For lLoop = 0 To pMx.FocusMap.LayerCount - 1
        ' If a feature class
617:         If TypeOf pMx.FocusMap.Layer(lLoop) Is IFeatureLayer Then
618:             Set pFL = pMx.FocusMap.Layer(lLoop)
619:             Set pFC = pFL.FeatureClass
            ' If a polygon layer
621:             If pFC.ShapeType = esriGeometryPolygon Then
                ' Add to combo box
623:                 cmbPolygonLayers.AddItem pFL.Name
'                If pFL.Name = sPreviousLayer Then
'                    lResetIndex = (cmbPolygonLayers.ListCount - 1)
'                End If
627:             End If
628:             lstRequiredDataLayers.AddItem pFL.Name
629:         End If
630:     Next
    'cmbPolygonLayers.ListIndex = lResetIndex
End Sub

Private Sub SetCurrentMapScaleCaption()
    Dim pMx As IMxDocument
    On Error GoTo eh
637:     Set pMx = m_Application.Document
638:     lblCurrentMapScale.Caption = Format(pMx.FocusMap.MapScale, "#,###,##0")
    Exit Sub
eh:
641:     lblCurrentMapScale.Caption = "<Scale Unknown>"
End Sub


Private Sub Form_Unload(Cancel As Integer)
646:     Set m_Application = Nothing
647:     Set GridSettings = Nothing
End Sub


Private Sub lstRequiredDataLayers_Click()
652:     SetControlsState
End Sub

Private Sub optColIDType_Click(Index As Integer)
656:     lblExampleID.Caption = GenerateExampleID
657:     SetControlsState
End Sub

Private Sub optGridIDOrder_Click(Index As Integer)
661:     lblExampleID.Caption = GenerateExampleID
662:     SetControlsState
End Sub

Private Sub optGridSize_Click(Index As Integer)
    Dim pMx As IMxDocument
667:     Set pMx = m_Application.Document
668:     lblCurrFrameName.Caption = pMx.FocusMap.Name
669:     SetControlsState
End Sub

Private Sub optLayerSource_Click(Index As Integer)
    ' If creating a new fclass to hold the grids
674:     If Index = 1 Then
        ' Set the field names (will be created automatically)
676:         cmbFieldID.Clear
677:         cmbFieldRowNum.Clear
678:         cmbFieldColNum.Clear
679:         cmbFieldMapScale.Clear
680:         cmbFieldID.AddItem "<None>"
681:         cmbFieldRowNum.AddItem "<None>"
682:         cmbFieldColNum.AddItem "<None>"
683:         cmbFieldMapScale.AddItem "<None>"
684:         cmbFieldID.AddItem c_DefaultFld_GridID
685:         cmbFieldRowNum.AddItem c_DefaultFld_RowNum
686:         cmbFieldColNum.AddItem c_DefaultFld_ColNum
687:         cmbFieldMapScale.AddItem c_DefaultFld_Scale
688:         cmbFieldID.ListIndex = 1
689:         cmbFieldRowNum.ListIndex = 1
690:         cmbFieldColNum.ListIndex = 1
691:         cmbFieldMapScale.ListIndex = 1
692:     End If
693:     SetControlsState
End Sub

Private Sub optRowIDType_Click(Index As Integer)
697:     lblExampleID.Caption = GenerateExampleID
698:     SetControlsState
End Sub

Private Function GenerateExampleID() As String
    Dim sRow As String, sCol As String
703:     If optStartingIDPosition(0).value Then  'Top left
704:         If (optRowIDType(0).value) Then
705:             sRow = "A"
706:         Else
707:             sRow = "1"
708:         End If
709:         If (optColIDType(0).value) Then
710:             sCol = "C"
711:         Else
712:             sCol = "3"
713:         End If
714:     Else                                    ' Lower left
715:         If (optRowIDType(0).value) Then
716:             sRow = "C"
717:         Else
718:             sRow = "3"
719:         End If
720:         If (optColIDType(0).value) Then
721:             sCol = "C"
722:         Else
723:             sCol = "3"
724:         End If
725:     End If
726:     If (optGridIDOrder(0).value) Then
727:         If chkBreak.value = vbChecked Then
728:             GenerateExampleID = sRow & "_" & sCol
729:         Else
730:             GenerateExampleID = sRow & sCol
731:         End If
732:     Else
733:         If chkBreak.value = vbChecked Then
734:             GenerateExampleID = sCol & "_" & sRow
735:         Else
736:             GenerateExampleID = sCol & sRow
737:         End If
738:     End If
End Function

Private Sub optScaleSource_Click(Index As Integer)
742:     If Index = 0 Then
743:         SetCurrentMapScaleCaption
744:     End If
745:     SetControlsState
End Sub

Private Sub optStartingIDPosition_Click(Index As Integer)
749:     lblExampleID.Caption = GenerateExampleID
750:     SetControlsState
End Sub

Private Sub txtEndCoordX_Change()
754:     SetControlsState
End Sub

Private Sub txtEndCoordX_KeyPress(KeyAscii As Integer)
    ' If a non-numeric (that is not a decimal point)
759:     If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
     And KeyAscii <> Asc(".") _
     And Chr(KeyAscii) <> vbBack Then
        ' Do not allow this button to work
763:         KeyAscii = 0
    ' If a decimal point, make sure we only ever get one of them
765:     ElseIf KeyAscii = Asc(".") Then
766:         If InStr(txtEndCoordX.Text, ".") > 0 Then
767:             KeyAscii = 0
768:         End If
769:     End If
End Sub

Private Sub txtEndCoordY_Change()
773:     SetControlsState
End Sub

Private Sub txtEndCoordY_KeyPress(KeyAscii As Integer)
    ' If a non-numeric (that is not a decimal point)
778:     If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
     And KeyAscii <> Asc(".") _
     And Chr(KeyAscii) <> vbBack Then
        ' Do not allow this button to work
782:         KeyAscii = 0
    ' If a decimal point, make sure we only ever get one of them
784:     ElseIf KeyAscii = Asc(".") Then
785:         If InStr(txtEndCoordY.Text, ".") > 0 Then
786:             KeyAscii = 0
787:         End If
788:     End If
End Sub

Private Sub txtManualGridHeight_Change()
792:     SetControlsState
End Sub

Private Sub txtManualGridHeight_KeyPress(KeyAscii As Integer)
    ' If a non-numeric (that is not a decimal point)
797:     If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
     And KeyAscii <> Asc(".") _
     And Chr(KeyAscii) <> vbBack Then
        ' Do not allow this button to work
801:         KeyAscii = 0
    ' If a decimal point, make sure we only ever get one of them
803:     ElseIf KeyAscii = Asc(".") Then
804:         If InStr(txtManualGridHeight.Text, ".") > 0 Then
805:             KeyAscii = 0
806:         End If
807:     End If
End Sub

Private Sub txtManualGridWidth_Change()
811:     SetControlsState
End Sub

Private Sub txtManualGridWidth_KeyPress(KeyAscii As Integer)
    ' If a non-numeric (that is not a decimal point)
816:     If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
     And KeyAscii <> Asc(".") _
     And Chr(KeyAscii) <> vbBack Then
        ' Do not allow this button to work
820:         KeyAscii = 0
    ' If a decimal point, make sure we only ever get one of them
822:     ElseIf KeyAscii = Asc(".") Then
823:         If InStr(txtManualGridWidth.Text, ".") > 0 Then
824:             KeyAscii = 0
825:         End If
826:     End If
End Sub

Private Sub txtManualMapScale_Change()
830:     SetControlsState
End Sub

Private Sub txtManualMapScale_KeyPress(KeyAscii As Integer)
    ' If a non-numeric (that is not a decimal point)
835:     If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
     And KeyAscii <> Asc(".") _
     And Chr(KeyAscii) <> vbBack Then
        ' Do not allow this button to work
839:         KeyAscii = 0
    ' If a decimal point, make sure we only ever get one of them
841:     ElseIf KeyAscii = Asc(".") Then
842:         If InStr(txtManualMapScale.Text, ".") > 0 Then
843:             KeyAscii = 0
844:         End If
845:     End If
End Sub

Public Sub Tickle()
849:     Call Form_Load
End Sub

Private Sub txtStartCoordX_Change()
853:     SetControlsState
End Sub

Private Sub txtStartCoordX_KeyPress(KeyAscii As Integer)
    ' If a non-numeric (that is not a decimal point)
858:     If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
     And KeyAscii <> Asc(".") _
     And Chr(KeyAscii) <> vbBack Then
        ' Do not allow this button to work
862:         KeyAscii = 0
    ' If a decimal point, make sure we only ever get one of them
864:     ElseIf KeyAscii = Asc(".") Then
865:         If InStr(txtStartCoordX.Text, ".") > 0 Then
866:             KeyAscii = 0
867:         End If
868:     End If
End Sub

Private Sub txtStartCoordY_Change()
872:     SetControlsState
End Sub

Private Sub txtStartCoordY_KeyPress(KeyAscii As Integer)
    ' If a non-numeric (that is not a decimal point)
877:     If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
     And KeyAscii <> Asc(".") _
     And Chr(KeyAscii) <> vbBack Then
        ' Do not allow this button to work
881:         KeyAscii = 0
    ' If a decimal point, make sure we only ever get one of them
883:     ElseIf KeyAscii = Asc(".") Then
884:         If InStr(txtStartCoordY.Text, ".") > 0 Then
885:             KeyAscii = 0
886:         End If
887:     End If
End Sub

Private Sub SetVisibleControls(iStep As Integer)
    ' Hide all
892:     fraAttributes.Visible = False
893:     fraDataFrameSize.Visible = False
894:     fraDestinationFeatureClass.Visible = False
895:     fraGridIDs.Visible = False
896:     fraScaleStart.Visible = False
    ' Show applicable frame, set top/left
    Select Case iStep
        Case 0:
900:             fraDestinationFeatureClass.Visible = True
901:             fraDestinationFeatureClass.Top = 0
902:             fraDestinationFeatureClass.Left = 0
        Case 1:
904:             fraAttributes.Visible = True
905:             fraAttributes.Top = 0
906:             fraAttributes.Left = 0
        Case 2:
908:             fraScaleStart.Visible = True
909:             fraScaleStart.Top = 0
910:             fraScaleStart.Left = 0
        Case 3:
912:             fraDataFrameSize.Visible = True
913:             fraDataFrameSize.Top = 0
914:             fraDataFrameSize.Left = 0
        Case 4:
916:             fraGridIDs.Visible = True
917:             fraGridIDs.Top = 0
918:             fraGridIDs.Left = 0
        Case Else:
920:             MsgBox "Invalid Step Value."
921:     End Select
End Sub

Private Sub CheckOutputFile()
    'Check the output option
926:     If txtNewGridLayer.Text <> "" Then
927:         If DoesShapeFileExist(txtNewGridLayer.Text) Then
928:             MsgBox "Shape file name already being used!!!"
929:             txtNewGridLayer.Text = ""
930:         End If
931:     End If
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
944:     Set pNewFields = New Fields
945:     Set pFieldsEdit = pNewFields
946:     Set pMx = m_Application.Document
    ' Field: OID
948:     Set newField = New Field
949:     Set newFieldEdit = newField
950:     With newFieldEdit
951:         .Name = "OID"
952:         .Type = esriFieldTypeOID
953:         .AliasName = "Object ID"
954:         .IsNullable = False
955:     End With
956:     pFieldsEdit.AddField newField
    'Set pFieldsEdit.Field(0) = pFieldEdit
    
'    ' Field: SHAPE
'    Set newField = New Field
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
973:     Set newField = New Field
974:     Set newFieldEdit = newField
975:     With newFieldEdit
976:       .Name = c_DefaultFld_GridID
977:       .AliasName = "GridIdentifier"
978:       .Type = esriFieldTypeString
979:       .IsNullable = True
980:       .length = 50
981:     End With
982:     pFieldsEdit.AddField newField
    ' Field: ROW NUMBER
984:     Set newField = New Field
985:     Set newFieldEdit = newField
986:     With newFieldEdit
987:       .Name = c_DefaultFld_RowNum
988:       .AliasName = "Row Number"
989:       .Type = esriFieldTypeInteger
990:       .IsNullable = True
991:     End With
992:     pFieldsEdit.AddField newField
    ' Field: COLUMN NUMBER
994:     Set newField = New Field
995:     Set newFieldEdit = newField
996:     With newFieldEdit
997:       .Name = c_DefaultFld_ColNum
998:       .AliasName = "Column Number"
999:       .Type = esriFieldTypeInteger
1000:       .IsNullable = True
1001:     End With
1002:     pFieldsEdit.AddField newField
    ' Field: SCALE
1004:     Set newField = New Field
1005:     Set newFieldEdit = newField
1006:     With newFieldEdit
1007:       .Name = c_DefaultFld_Scale
1008:       .AliasName = "Plot Scale"
1009:       .Type = esriFieldTypeDouble
1010:       .IsNullable = True
1011:       .Precision = 18
1012:       .Scale = 11
1013:     End With
1014:     pFieldsEdit.AddField newField
    ' Return
1016:     Set CreateTheFields = pFieldsEdit
End Function

