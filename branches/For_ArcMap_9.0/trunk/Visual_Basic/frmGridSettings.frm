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
49:     If Len(lblCurrentMapScale.Caption) = 0 Then lblCurrentMapScale.Caption = "0"
50:     If Len(txtManualMapScale.Text) = 0 Then
51:         dScale = 0
52:     Else
53:         dScale = CDbl(txtManualMapScale.Text)
54:     End If
55:     If Len(txtManualGridHeight.Text) = 0 Then
56:         dGHeight = 0
57:     Else
58:         dGHeight = CDbl(txtManualGridHeight.Text)
59:     End If
60:     If Len(txtManualGridWidth.Text) = 0 Then
61:         dGWidth = 0
62:     Else
63:         dGWidth = CDbl(txtManualGridWidth.Text)
64:     End If
65:     If Len(txtStartCoordX.Text) = 0 Then
66:         dStartX = 0
67:     Else
68:         dStartX = CDbl(txtStartCoordX.Text)
69:     End If
70:     If Len(txtStartCoordY.Text) = 0 Then
71:         dStartY = 0
72:     Else
73:         dStartY = CDbl(txtStartCoordY.Text)
74:     End If
75:     If Len(txtEndCoordX.Text) = 0 Then
76:         dEndX = 0
77:     Else
78:         dEndX = CDbl(txtEndCoordX.Text)
79:     End If
80:     If Len(txtEndCoordY.Text) = 0 Then
81:         dEndY = 0
82:     Else
83:         dEndY = CDbl(txtEndCoordY.Text)
84:     End If
85: i = 1

    ' Calc values
88:     bValidScale = (optScaleSource(0).Value And CDbl(lblCurrentMapScale.Caption) > 0) Or _
                  (optScaleSource(1).Value And dScale > 0)
90:     bValidSize = (optGridSize(0).Value) Or _
                 (optGridSize(1).Value And dGHeight > 0 And dGWidth > 0)
92:     bCreatingNewFClass = optLayerSource(1).Value
93:     bNewFClassSet = (Len(txtNewGridLayer.Text) > 0)
94:     bValidTarget = (cmbPolygonLayers.ListIndex > 0) Or (bCreatingNewFClass And bNewFClassSet)
95:     bValidIDField = (cmbFieldID.ListIndex >= 0)
96:     bValidReqdLayers = (chkRemoveEmpties.Value = vbUnchecked) Or _
                       (chkRemoveEmpties.Value = vbChecked And lstRequiredDataLayers.SelCount > 0)
98: i = 2
99:     If bValidTarget And (Not bCreatingNewFClass) Then
100:         Set pFL = FindFeatureLayerByName(cmbPolygonLayers.List(cmbPolygonLayers.ListIndex), m_Application)
101:         If pFL.FeatureClass.FeatureDataset Is Nothing Then
102:             bValidStart = True
103:             bValidEnd = True
104:         Else
105:             Set pDatasetExtent = GetValidExtentForLayer(pFL)
106:             bValidStart = ((dStartX >= pDatasetExtent.XMin) And (dStartX <= pDatasetExtent.XMax)) _
                            And _
                          ((dStartY >= pDatasetExtent.YMin) And (dStartY <= pDatasetExtent.YMax))
109:             bValidEnd = ((dEndX >= pDatasetExtent.XMin) And (dEndX <= pDatasetExtent.XMax)) _
                            And _
                        ((dEndY >= pDatasetExtent.YMin) And (dEndY <= pDatasetExtent.YMax)) _
                            And _
                        ((dEndX > dStartX) And (dEndY > dStartY))
114:         End If
115:     ElseIf bValidTarget And bCreatingNewFClass Then
116:         bValidStart = True
117:         bValidEnd = True
118:     End If
119:     bDuplicateFieldsSelected = (cmbFieldRowNum.ListIndex > 0 And cmbFieldRowNum.ListIndex = cmbFieldColNum.ListIndex) _
                            Or (cmbFieldRowNum.ListIndex > 0 And cmbFieldRowNum.ListIndex = cmbFieldMapScale.ListIndex) _
                            Or (cmbFieldColNum.ListIndex > 0 And cmbFieldColNum.ListIndex = cmbFieldMapScale.ListIndex)
122: i = 3
    
    ' Set states
    Select Case m_Step
        Case 0:     ' Set the target feature layer
127:             cmdBack.Enabled = False
128:             cmdNext.Enabled = bValidTarget And bValidReqdLayers
129:             cmdNext.Caption = "Next >"
130:             cmbPolygonLayers.Enabled = Not bCreatingNewFClass
        Case 1:     ' Set the fields to populate
132:             cmdBack.Enabled = True
133:             cmdNext.Enabled = (bValidIDField And Not bDuplicateFieldsSelected)
134:             cmbFieldID.Enabled = Not bCreatingNewFClass
135:             cmbFieldRowNum.Enabled = Not bCreatingNewFClass
136:             cmbFieldColNum.Enabled = Not bCreatingNewFClass
137:             cmbFieldMapScale.Enabled = Not bCreatingNewFClass
        Case 2:     ' Set the scale / starting_coords
139:             cmdBack.Enabled = True
140:             cmdNext.Enabled = bValidScale And bValidStart And bValidEnd
141:             If Not bCreatingNewFClass Then
142:                 cmdDatasetExtentLL.Enabled = Not (pFL.FeatureClass.FeatureDataset Is Nothing)
143:             Else
144:                 cmdDatasetExtentLL.Enabled = False
145:             End If
        Case 3:     ' Set the dataframe properties
147:             cmdBack.Enabled = True
148:             cmdNext.Enabled = bValidSize
149:             cmdNext.Caption = "Next >"
        Case 4:     ' Set the ID values
151:             cmdBack.Enabled = True
152:             cmdNext.Enabled = True
153:             cmdNext.Caption = "Finish"
        Case Else:
155:             cmdBack.Enabled = False
156:             cmdNext.Enabled = False
157:     End Select
158: i = 4
    
160:     txtManualMapScale.Enabled = optScaleSource(1).Value
161:     txtManualGridWidth.Enabled = optGridSize(1).Value
162:     txtManualGridHeight.Enabled = optGridSize(1).Value
163:     cmbGridSizeUnits.Enabled = optGridSize(1).Value
    ' Set display
165:     If bValidStart Then
166:         txtStartCoordX.ForeColor = (&H0)    ' Black
167:         txtStartCoordY.ForeColor = (&H0)
168:     Else
169:         txtStartCoordX.ForeColor = (&HFF)   ' Red
170:         txtStartCoordY.ForeColor = (&HFF)
171:     End If
172:     If bValidEnd Then
173:         txtEndCoordX.ForeColor = (&H0)      ' Black
174:         txtEndCoordY.ForeColor = (&H0)
175:     Else
176:         txtEndCoordX.ForeColor = (&HFF)     ' Red
177:         txtEndCoordY.ForeColor = (&HFF)
178:     End If
179:     If optScaleSource(1).Value Then
180:         If bValidScale Then
181:             txtManualMapScale.ForeColor = (&H0)      ' Black
182:         Else
183:             txtManualMapScale.ForeColor = (&HFF)     ' Red
184:         End If
185:     End If
186:     If optGridSize(1).Value Then
187:         If bValidSize Then
188:             txtManualGridWidth.ForeColor = (&H0)      ' Black
189:             txtManualGridHeight.ForeColor = (&H0)
190:         Else
191:             txtManualGridWidth.ForeColor = (&HFF)     ' Red
192:             txtManualGridHeight.ForeColor = (&HFF)
193:         End If
194:     End If
    
    Exit Sub
197:     Resume
eh:
199:     MsgBox Err.Description, vbExclamation, "SetControlsState " & i
End Sub

Private Sub chkBreak_Click()
203:     lblExampleID.Caption = GenerateExampleID
End Sub

Private Sub chkRemoveEmpties_Click()
207:     SetControlsState
End Sub

Private Sub cmbFieldColNum_Click()
211:     SetControlsState
End Sub

Private Sub cmbFieldID_Click()
215:     SetControlsState
End Sub

Private Sub cmbFieldMapScale_Click()
219:     SetControlsState
End Sub

Private Sub cmbFieldRowNum_Click()
223:     SetControlsState
End Sub

Private Sub cmbPolygonLayers_Click()
    Dim pFL As IFeatureLayer
    Dim pFields As IFields
    Dim lLoop As Long
    ' Populate the fields combo boxes
231:     If cmbPolygonLayers.ListIndex > 0 Then
232:         Set pFL = FindFeatureLayerByName(cmbPolygonLayers.List(cmbPolygonLayers.ListIndex), m_Application)
233:         Set pFields = pFL.FeatureClass.Fields
234:         cmbFieldColNum.Clear
235:         cmbFieldID.Clear
236:         cmbFieldMapScale.Clear
237:         cmbFieldRowNum.Clear
238:         cmbFieldRowNum.AddItem "<None>"
239:         cmbFieldColNum.AddItem "<None>"
240:         cmbFieldMapScale.AddItem "<None>"
241:         For lLoop = 0 To pFields.FieldCount - 1
242:             If pFields.Field(lLoop).Type = esriFieldTypeString Then
243:                 cmbFieldID.AddItem pFields.Field(lLoop).Name
244:             ElseIf pFields.Field(lLoop).Type = esriFieldTypeDouble Or _
                   pFields.Field(lLoop).Type = esriFieldTypeInteger Or _
                   pFields.Field(lLoop).Type = esriFieldTypeSmallInteger Or _
                   pFields.Field(lLoop).Type = esriFieldTypeSingle Then
248:                 cmbFieldColNum.AddItem pFields.Field(lLoop).Name
249:                 cmbFieldRowNum.AddItem pFields.Field(lLoop).Name
250:                 cmbFieldMapScale.AddItem pFields.Field(lLoop).Name
251:             End If
252:         Next
253:         cmbFieldRowNum.ListIndex = 0
254:         cmbFieldColNum.ListIndex = 0
255:         cmbFieldMapScale.ListIndex = 0
256:     End If
257:     SetControlsState
End Sub

Private Sub cmdBack_Click()
261:     m_Step = m_Step - 1
262:     If m_Step < 0 Then
263:         m_Step = 0
264:     End If
265:     SetVisibleControls m_Step
266:     SetControlsState
End Sub

Private Sub cmdClose_Click()
270:     Set m_Application = Nothing
271:     Set Me.GridSettings = Nothing
272:     Me.Hide
End Sub

Private Sub CollateGridSettings()
    Dim pMx As IMxDocument
    Dim pCreateGrid As New clsCreateGrids
    Dim pFrameElement As IElement
    Dim sDestLayerName As String
    Dim lLoop As Long
    ' Populate class
282:     If (optGridIDOrder(0).Value) Then
283:         pCreateGrid.IdentifierOrder = Row_Column
284:     Else
285:         pCreateGrid.IdentifierOrder = Column_Row
286:     End If
287:     If (optRowIDType(0).Value) Then
288:         pCreateGrid.RowIDType = Alphabetical
289:     Else
290:         pCreateGrid.RowIDType = Numerical
291:     End If
292:     If (optColIDType(0).Value) Then
293:         pCreateGrid.ColIDType = Alphabetical
294:     Else
295:         pCreateGrid.ColIDType = Numerical
296:     End If
297:     If (optStartingIDPosition(0).Value) Then
298:         pCreateGrid.IDStartPositionType = TopLeft
299:     Else
300:         pCreateGrid.IDStartPositionType = LowerLeft
301:     End If
302:     If (optScaleSource(0).Value) Then
303:         pCreateGrid.MapScale = CDbl(lblCurrentMapScale.Caption)
304:     Else
305:         pCreateGrid.MapScale = CDbl(txtManualMapScale.Text)
306:     End If
307:     If (optGridSize(0).Value) Then
308:         Set pFrameElement = GetDataFrameElement(GetActiveDataFrameName(m_Application), m_Application)
309:         pCreateGrid.FrameWidthInPageUnits = pFrameElement.Geometry.Envelope.Width
310:         pCreateGrid.FrameHeightInPageUnits = pFrameElement.Geometry.Envelope.Height
311:     Else
312:         pCreateGrid.FrameWidthInPageUnits = CDbl(txtManualGridWidth.Text)
313:         pCreateGrid.FrameHeightInPageUnits = CDbl(txtManualGridHeight.Text)
314:     End If
315:     sDestLayerName = cmbPolygonLayers.List(cmbPolygonLayers.ListIndex)
316:     If optLayerSource(0).Value Then
317:         Set pCreateGrid.DestinationFeatureLayer = FindFeatureLayerByName(sDestLayerName, m_Application)
318:     End If
319:     pCreateGrid.StartingCoordinateLL_X = CDbl(txtStartCoordX.Text)
320:     pCreateGrid.StartingCoordinateLL_Y = CDbl(txtStartCoordY.Text)
321:     pCreateGrid.EndingCoordinateUR_X = CDbl(txtEndCoordX.Text)
322:     pCreateGrid.EndingCoordinateUR_Y = CDbl(txtEndCoordY.Text)
323:     pCreateGrid.UseUnderscore = (chkBreak.Value = vbChecked)
324:     pCreateGrid.FieldNameGridID = cmbFieldID.List(cmbFieldID.ListIndex)
325:     If cmbFieldRowNum.ListIndex > 0 Then pCreateGrid.FieldNameRowNum = cmbFieldRowNum.List(cmbFieldRowNum.ListIndex)
326:     If cmbFieldColNum.ListIndex > 0 Then pCreateGrid.FieldNameColNum = cmbFieldColNum.List(cmbFieldColNum.ListIndex)
327:     If cmbFieldMapScale.ListIndex > 0 Then pCreateGrid.FieldNameScale = cmbFieldMapScale.List(cmbFieldMapScale.ListIndex)
328:     pCreateGrid.NoEmptyGrids = (chkRemoveEmpties.Value = vbChecked)
329:     If pCreateGrid.NoEmptyGrids Then
330:         pCreateGrid.ClearRequiredDataLayers
331:         For lLoop = 0 To lstRequiredDataLayers.ListCount - 1
332:             If lstRequiredDataLayers.Selected(lLoop) Then
333:                 pCreateGrid.AddRequiredDataLayer lstRequiredDataLayers.List(lLoop)
334:             End If
335:         Next
336:     End If
337:     pCreateGrid.RemoveCurrentGrids = (chkRemovePreviousGrids.Value = vbChecked)
    ' Place grid settings on Public form property (so calling function can use them)
339:     Set Me.GridSettings = pCreateGrid
End Sub

Private Sub cmdDatasetExtentLL_Click()
    Dim pFL As IFeatureLayer
    Dim pDatasetExtent As IEnvelope
    
346:     If cmbPolygonLayers.ListIndex > 0 Then
347:         Set pFL = FindFeatureLayerByName(cmbPolygonLayers.List(cmbPolygonLayers.ListIndex), m_Application)
348:         Set pDatasetExtent = GetValidExtentForLayer(pFL)
349:         txtStartCoordX.Text = Format(pDatasetExtent.XMin, "#,###,##0.00")
350:         txtStartCoordY.Text = Format(pDatasetExtent.YMin, "#,###,##0.00")
351:         txtEndCoordX.Text = Format(pDatasetExtent.XMax - 100, "#,###,##0.00")
352:         txtEndCoordY.Text = Format(pDatasetExtent.YMax - 100, "#,###,##0.00")
353:         SetControlsState
354:     End If
End Sub

Private Sub cmdLayersExtent_Click()
    Dim pMx As IMxDocument
    Dim pEnv As IEnvelope
    Dim pElement As IElement
    Dim pMapFrame As IMapFrame
    Dim pActiveView As IActiveView
    
    On Error GoTo eh
    
366:     Set pMx = m_Application.Document
367:     Set pActiveView = pMx.ActiveView
368:     If TypeOf pActiveView Is IPageLayout Then
369:         Set pElement = GetDataFrameElement(pMx.FocusMap.Name, m_Application)
370:         Set pMapFrame = pElement
371:         Set pEnv = pMapFrame.MapBounds
372:         Set pActiveView = pMapFrame.Map
373:         Set pEnv = pActiveView.FullExtent
374:     Else
375:         Set pEnv = pActiveView.FullExtent
376:     End If
    
378:     txtStartCoordX.Text = Format(pEnv.XMin, "#,###,##0.00")
379:     txtStartCoordY.Text = Format(pEnv.YMin, "#,###,##0.00")
380:     txtEndCoordX.Text = Format(pEnv.XMax, "#,###,##0.00")
381:     txtEndCoordY.Text = Format(pEnv.YMax, "#,###,##0.00")
    
383:     SetControlsState
    
    Exit Sub
eh:
387:     MsgBox Err.Description, , "cmdLayersExtent_Click"
End Sub

Private Sub cmdMapExtentLL_Click()
    Dim pMx As IMxDocument
    Dim pEnv As IEnvelope
    Dim pElement As IElement
    Dim pMapFrame As IMapFrame
    Dim pActiveView As IActiveView
    
    On Error GoTo eh
    
399:     Set pMx = m_Application.Document
400:     Set pActiveView = pMx.ActiveView
401:     If TypeOf pActiveView Is IPageLayout Then
402:         Set pElement = GetDataFrameElement(pMx.FocusMap.Name, m_Application)
403:         Set pMapFrame = pElement
404:         Set pEnv = pMapFrame.MapBounds
405:     Else
406:         Set pEnv = pActiveView.Extent
407:     End If
    
409:     txtStartCoordX.Text = Format(pEnv.XMin, "#,###,##0.00")
410:     txtStartCoordY.Text = Format(pEnv.YMin, "#,###,##0.00")
411:     txtEndCoordX.Text = Format(pEnv.XMax, "#,###,##0.00")
412:     txtEndCoordY.Text = Format(pEnv.YMax, "#,###,##0.00")
    
414:     SetControlsState
    
    Exit Sub
eh:
418:     MsgBox Err.Description, , "cmdMapExtentLL"
End Sub

Private Sub cmdNext_Click()
    Dim pMx As IMxDocument
    Dim pFeatureLayer As IFeatureLayer
    Dim pOutputFClass As IFeatureClass
    Dim pNewFields As IFields
    
    On Error GoTo eh
    ' Step
429:     m_Step = m_Step + 1
    ' If we're creating a new fclass, we can skip a step
431:     If m_Step = 1 And (optLayerSource(1).Value) Then
432:         m_Step = m_Step + 1
433:     End If
    ' If FINISH
435:     If m_Step >= 5 Then
436:         CollateGridSettings
        ' If creating a new layer
438:         If optLayerSource(1).Value Then
            ' Create the feature class
440:             Set pMx = m_Application.Document
441:             Set pNewFields = CreateTheFields
            Select Case m_FileType
                Case ShapeFile
444:                     Set pOutputFClass = NewShapeFile(m_OutputLayer, pMx.FocusMap, pNewFields)
                Case AccessFeatureClass
446:                     Set pOutputFClass = NewAccessFile(m_OutputLayer, _
                            m_OutputDataset, m_OutputFClass, pNewFields)
448:             End Select
449:             If pOutputFClass Is Nothing Then
450:                 Err.Raise vbObjectError, "cmdNext", "Could not create the new output feature class."
451:             End If
            ' Create new layer
453:             Set pFeatureLayer = New FeatureLayer
454:             Set pFeatureLayer.FeatureClass = pOutputFClass
455:             pFeatureLayer.Name = pFeatureLayer.FeatureClass.AliasName
            ' Add the new layer to arcmap & reset the GridSettings object to point at it
457:             pMx.FocusMap.AddLayer pFeatureLayer
458:             Set GridSettings.DestinationFeatureLayer = pFeatureLayer
459:         End If
460:         Me.Hide
461:     Else
462:         SetVisibleControls m_Step
463:         SetControlsState
464:     End If
    
    Exit Sub
eh:
468:     MsgBox "cmdNext_Click - " & Erl & " - " & Err.Description
469:     m_Step = m_Step - 1
End Sub

Private Sub cmdSetNewGridLayer_Click()
On Error GoTo ErrHand:
  Dim pGxFilter As IGxObjectFilter
  Dim pGXBrow As IGxDialog, bFlag As Boolean
  Dim pSel As IEnumGxObject, pApp As IApplication
  
478:   Set pGxFilter = New GxFilter
479:   Set pApp = m_Application
480:   Set pGXBrow = New GxDialog
481:   Set pGXBrow.ObjectFilter = pGxFilter
482:   pGXBrow.Title = "Output feature class or shapefile"
483:   bFlag = pGXBrow.DoModalSave(pApp.hwnd)
  
485:   If bFlag Then
    Dim pObj As IGxObject
487:     Set pObj = pGXBrow.FinalLocation
488:     m_bIsGeoDatabase = True
489:     If UCase(pObj.Category) = "FOLDER" Then
490:       If InStr(1, pGXBrow.Name, ".shp") > 0 Then
491:         txtNewGridLayer.Text = pObj.FullName & "\" & pGXBrow.Name
492:       Else
493:         txtNewGridLayer.Text = pObj.FullName & "\" & pGXBrow.Name & ".shp"
494:       End If
495:       m_OutputLayer = txtNewGridLayer.Text
496:       m_bIsGeoDatabase = False
497:       m_FileType = ShapeFile
498:      CheckOutputFile
499:     Else
      Dim pLen As Long
501:       pLen = Len(pObj.FullName) - Len(pObj.BaseName) - 1
502:       txtNewGridLayer.Text = Left(pObj.FullName, pLen)
503:       m_OutputLayer = Left(pObj.FullName, pLen)
504:       m_OutputDataset = pObj.BaseName
505:       m_OutputFClass = pGXBrow.Name
506:       m_bIsGeoDatabase = True
507:       If UCase(pObj.Category) = "PERSONAL GEODATABASE FEATURE DATASET" Then
508:         m_FileType = AccessFeatureClass
509:       Else
510:         m_FileType = SDEFeatureClass
511:       End If
512:     End If
513:   Else
514:     txtNewGridLayer.Text = ""
515:     m_bIsGeoDatabase = False
516:   End If
517:   SetControlsState
  
  Exit Sub
ErrHand:
521:   MsgBox "cmdSetNewGridLayer_Click - " & Erl & " - " & Err.Description
End Sub

Private Sub Form_Load()
    Dim pMx As IMxDocument
    Dim bRenewCoordsX As Boolean
    Dim bRenewCoordsY As Boolean
    
    On Error GoTo eh
    
531:     Set pMx = m_Application.Document
532:     Me.Height = 5665
533:     Me.Width = 4935
534:     m_Step = 0
535:     LoadLayersComboBox
536:     LoadUnitsComboBox
537:     lblExampleID.Caption = GenerateExampleID
538:     lblCurrFrameName.Caption = GetActiveDataFrameName(m_Application)
539:     If pMx.FocusMap.MapUnits = esriUnknownUnits Then
540:         MsgBox "Error: The map has unknown units and therefore cannot calculate a Scale." _
            & vbCrLf & "Cannot create Map Grids at this time.", vbCritical, "Create Map Grids"
542:         Unload Me
        Exit Sub
544:     End If
545:     lblCurrentMapScale.Caption = Format(pMx.FocusMap.MapScale, "#,###,##0")
546:     Call cmdMapExtentLL_Click
547:     SetVisibleControls m_Step
    
549:     SetControlsState
    
    'Make sure the wizard stays on top
552:     TopMost Me
    
    Exit Sub
eh:
556:     MsgBox "Error loading the form: " & Erl & " - " & Err.Description & vbCrLf _
        & vbCrLf & "Attempting to continue the load...", , "MapGridManager: Form_Load "
    On Error Resume Next
559:     SetVisibleControls m_Step
560:     SetControlsState
End Sub

Private Sub LoadUnitsComboBox()
    Dim pMx As IMxDocument
    Dim sPageUnitsDesc As String
    Dim pPage As IPage
    
    On Error GoTo eh
    
    ' Init
571:     Set pMx = m_Application.Document
572:     Set pPage = pMx.PageLayout.Page
573:     sPageUnitsDesc = GetUnitsDescription(pPage.Units)
574:     cmbGridSizeUnits.Clear
    ' Add
576:     cmbGridSizeUnits.AddItem sPageUnitsDesc
    'cmbGridSizeUnits.AddItem "Map Units (" & sMapUnitsDesc & ")"
    ' Set page units as default
579:     cmbGridSizeUnits.ListIndex = 0
    
    Exit Sub
eh:
583:     Err.Raise vbObjectError, "LoadUnitsComboBox", "Error in LoadUnitsComboBox" & vbCrLf & Err.Description
End Sub

Private Sub LoadLayersComboBox()
    Dim pMx As IMxDocument
    Dim lLoop As Long
    Dim pFL As IFeatureLayer
    Dim pFC As IFeatureClass
    Dim sPreviousLayer  As String
    Dim lResetIndex As Long
    
    'Init
595:     Set pMx = m_Application.Document
'    If cmbPolygonLayers.ListCount > 0 Then
'        sPreviousLayer = cmbPolygonLayers.List(cmbPolygonLayers.ListIndex)
'    End If
599:     cmbPolygonLayers.Clear
600:     lstRequiredDataLayers.Clear
601:     cmbPolygonLayers.AddItem "<Not Set>"
    ' For all layers
603:     For lLoop = 0 To pMx.FocusMap.LayerCount - 1
        ' If a feature class
605:         If TypeOf pMx.FocusMap.Layer(lLoop) Is IFeatureLayer Then
606:             Set pFL = pMx.FocusMap.Layer(lLoop)
607:             Set pFC = pFL.FeatureClass
            ' If a polygon layer
609:             If pFC.ShapeType = esriGeometryPolygon Then
                ' Add to combo box
611:                 cmbPolygonLayers.AddItem pFL.Name
'                If pFL.Name = sPreviousLayer Then
'                    lResetIndex = (cmbPolygonLayers.ListCount - 1)
'                End If
615:             End If
616:             lstRequiredDataLayers.AddItem pFL.Name
617:         End If
618:     Next
    'cmbPolygonLayers.ListIndex = lResetIndex
End Sub

Private Sub SetCurrentMapScaleCaption()
    Dim pMx As IMxDocument
    On Error GoTo eh
625:     Set pMx = m_Application.Document
626:     lblCurrentMapScale.Caption = Format(pMx.FocusMap.MapScale, "#,###,##0")
    Exit Sub
eh:
629:     lblCurrentMapScale.Caption = "<Scale Unknown>"
End Sub


Private Sub Form_Unload(Cancel As Integer)
634:     Set m_Application = Nothing
635:     Set GridSettings = Nothing
End Sub


Private Sub lstRequiredDataLayers_Click()
640:     SetControlsState
End Sub

Private Sub optColIDType_Click(Index As Integer)
644:     lblExampleID.Caption = GenerateExampleID
645:     SetControlsState
End Sub

Private Sub optGridIDOrder_Click(Index As Integer)
649:     lblExampleID.Caption = GenerateExampleID
650:     SetControlsState
End Sub

Private Sub optGridSize_Click(Index As Integer)
    Dim pMx As IMxDocument
655:     Set pMx = m_Application.Document
656:     lblCurrFrameName.Caption = pMx.FocusMap.Name
657:     SetControlsState
End Sub

Private Sub optLayerSource_Click(Index As Integer)
    ' If creating a new fclass to hold the grids
662:     If Index = 1 Then
        ' Set the field names (will be created automatically)
664:         cmbFieldID.Clear
665:         cmbFieldRowNum.Clear
666:         cmbFieldColNum.Clear
667:         cmbFieldMapScale.Clear
668:         cmbFieldID.AddItem "<None>"
669:         cmbFieldRowNum.AddItem "<None>"
670:         cmbFieldColNum.AddItem "<None>"
671:         cmbFieldMapScale.AddItem "<None>"
672:         cmbFieldID.AddItem c_DefaultFld_GridID
673:         cmbFieldRowNum.AddItem c_DefaultFld_RowNum
674:         cmbFieldColNum.AddItem c_DefaultFld_ColNum
675:         cmbFieldMapScale.AddItem c_DefaultFld_Scale
676:         cmbFieldID.ListIndex = 1
677:         cmbFieldRowNum.ListIndex = 1
678:         cmbFieldColNum.ListIndex = 1
679:         cmbFieldMapScale.ListIndex = 1
680:     End If
681:     SetControlsState
End Sub

Private Sub optRowIDType_Click(Index As Integer)
685:     lblExampleID.Caption = GenerateExampleID
686:     SetControlsState
End Sub

Private Function GenerateExampleID() As String
    Dim sRow As String, sCol As String
691:     If optStartingIDPosition(0).Value Then  'Top left
692:         If (optRowIDType(0).Value) Then
693:             sRow = "A"
694:         Else
695:             sRow = "1"
696:         End If
697:         If (optColIDType(0).Value) Then
698:             sCol = "C"
699:         Else
700:             sCol = "3"
701:         End If
702:     Else                                    ' Lower left
703:         If (optRowIDType(0).Value) Then
704:             sRow = "C"
705:         Else
706:             sRow = "3"
707:         End If
708:         If (optColIDType(0).Value) Then
709:             sCol = "C"
710:         Else
711:             sCol = "3"
712:         End If
713:     End If
714:     If (optGridIDOrder(0).Value) Then
715:         If chkBreak.Value = vbChecked Then
716:             GenerateExampleID = sRow & "_" & sCol
717:         Else
718:             GenerateExampleID = sRow & sCol
719:         End If
720:     Else
721:         If chkBreak.Value = vbChecked Then
722:             GenerateExampleID = sCol & "_" & sRow
723:         Else
724:             GenerateExampleID = sCol & sRow
725:         End If
726:     End If
End Function

Private Sub optScaleSource_Click(Index As Integer)
730:     If Index = 0 Then
731:         SetCurrentMapScaleCaption
732:     End If
733:     SetControlsState
End Sub

Private Sub optStartingIDPosition_Click(Index As Integer)
737:     lblExampleID.Caption = GenerateExampleID
738:     SetControlsState
End Sub

Private Sub txtEndCoordX_Change()
742:     SetControlsState
End Sub

Private Sub txtEndCoordX_KeyPress(KeyAscii As Integer)
    ' If a non-numeric (that is not a decimal point)
747:     If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
     And KeyAscii <> Asc(".") _
     And Chr(KeyAscii) <> vbBack Then
        ' Do not allow this button to work
751:         KeyAscii = 0
    ' If a decimal point, make sure we only ever get one of them
753:     ElseIf KeyAscii = Asc(".") Then
754:         If InStr(txtEndCoordX.Text, ".") > 0 Then
755:             KeyAscii = 0
756:         End If
757:     End If
End Sub

Private Sub txtEndCoordY_Change()
761:     SetControlsState
End Sub

Private Sub txtEndCoordY_KeyPress(KeyAscii As Integer)
    ' If a non-numeric (that is not a decimal point)
766:     If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
     And KeyAscii <> Asc(".") _
     And Chr(KeyAscii) <> vbBack Then
        ' Do not allow this button to work
770:         KeyAscii = 0
    ' If a decimal point, make sure we only ever get one of them
772:     ElseIf KeyAscii = Asc(".") Then
773:         If InStr(txtEndCoordY.Text, ".") > 0 Then
774:             KeyAscii = 0
775:         End If
776:     End If
End Sub

Private Sub txtManualGridHeight_Change()
780:     SetControlsState
End Sub

Private Sub txtManualGridHeight_KeyPress(KeyAscii As Integer)
    ' If a non-numeric (that is not a decimal point)
785:     If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
     And KeyAscii <> Asc(".") _
     And Chr(KeyAscii) <> vbBack Then
        ' Do not allow this button to work
789:         KeyAscii = 0
    ' If a decimal point, make sure we only ever get one of them
791:     ElseIf KeyAscii = Asc(".") Then
792:         If InStr(txtManualGridHeight.Text, ".") > 0 Then
793:             KeyAscii = 0
794:         End If
795:     End If
End Sub

Private Sub txtManualGridWidth_Change()
799:     SetControlsState
End Sub

Private Sub txtManualGridWidth_KeyPress(KeyAscii As Integer)
    ' If a non-numeric (that is not a decimal point)
804:     If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
     And KeyAscii <> Asc(".") _
     And Chr(KeyAscii) <> vbBack Then
        ' Do not allow this button to work
808:         KeyAscii = 0
    ' If a decimal point, make sure we only ever get one of them
810:     ElseIf KeyAscii = Asc(".") Then
811:         If InStr(txtManualGridWidth.Text, ".") > 0 Then
812:             KeyAscii = 0
813:         End If
814:     End If
End Sub

Private Sub txtManualMapScale_Change()
818:     SetControlsState
End Sub

Private Sub txtManualMapScale_KeyPress(KeyAscii As Integer)
    ' If a non-numeric (that is not a decimal point)
823:     If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
     And KeyAscii <> Asc(".") _
     And Chr(KeyAscii) <> vbBack Then
        ' Do not allow this button to work
827:         KeyAscii = 0
    ' If a decimal point, make sure we only ever get one of them
829:     ElseIf KeyAscii = Asc(".") Then
830:         If InStr(txtManualMapScale.Text, ".") > 0 Then
831:             KeyAscii = 0
832:         End If
833:     End If
End Sub

Public Sub Tickle()
837:     Call Form_Load
End Sub

Private Sub txtStartCoordX_Change()
841:     SetControlsState
End Sub

Private Sub txtStartCoordX_KeyPress(KeyAscii As Integer)
    ' If a non-numeric (that is not a decimal point)
846:     If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
     And KeyAscii <> Asc(".") _
     And Chr(KeyAscii) <> vbBack Then
        ' Do not allow this button to work
850:         KeyAscii = 0
    ' If a decimal point, make sure we only ever get one of them
852:     ElseIf KeyAscii = Asc(".") Then
853:         If InStr(txtStartCoordX.Text, ".") > 0 Then
854:             KeyAscii = 0
855:         End If
856:     End If
End Sub

Private Sub txtStartCoordY_Change()
860:     SetControlsState
End Sub

Private Sub txtStartCoordY_KeyPress(KeyAscii As Integer)
    ' If a non-numeric (that is not a decimal point)
865:     If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
     And KeyAscii <> Asc(".") _
     And Chr(KeyAscii) <> vbBack Then
        ' Do not allow this button to work
869:         KeyAscii = 0
    ' If a decimal point, make sure we only ever get one of them
871:     ElseIf KeyAscii = Asc(".") Then
872:         If InStr(txtStartCoordY.Text, ".") > 0 Then
873:             KeyAscii = 0
874:         End If
875:     End If
End Sub

Private Sub SetVisibleControls(iStep As Integer)
    ' Hide all
880:     fraAttributes.Visible = False
881:     fraDataFrameSize.Visible = False
882:     fraDestinationFeatureClass.Visible = False
883:     fraGridIDs.Visible = False
884:     fraScaleStart.Visible = False
    ' Show applicable frame, set top/left
    Select Case iStep
        Case 0:
888:             fraDestinationFeatureClass.Visible = True
889:             fraDestinationFeatureClass.Top = 0
890:             fraDestinationFeatureClass.Left = 0
        Case 1:
892:             fraAttributes.Visible = True
893:             fraAttributes.Top = 0
894:             fraAttributes.Left = 0
        Case 2:
896:             fraScaleStart.Visible = True
897:             fraScaleStart.Top = 0
898:             fraScaleStart.Left = 0
        Case 3:
900:             fraDataFrameSize.Visible = True
901:             fraDataFrameSize.Top = 0
902:             fraDataFrameSize.Left = 0
        Case 4:
904:             fraGridIDs.Visible = True
905:             fraGridIDs.Top = 0
906:             fraGridIDs.Left = 0
        Case Else:
908:             MsgBox "Invalid Step Value."
909:     End Select
End Sub

Private Sub CheckOutputFile()
    'Check the output option
914:     If txtNewGridLayer.Text <> "" Then
915:         If DoesShapeFileExist(txtNewGridLayer.Text) Then
916:             MsgBox "Shape file name already being used!!!"
917:             txtNewGridLayer.Text = ""
918:         End If
919:     End If
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
932:     Set pNewFields = New Fields
933:     Set pFieldsEdit = pNewFields
934:     Set pMx = m_Application.Document
    ' Field: OID
936:     Set newField = New Field
937:     Set newFieldEdit = newField
938:     With newFieldEdit
939:         .Name = "OID"
940:         .Type = esriFieldTypeOID
941:         .AliasName = "Object ID"
942:         .IsNullable = False
943:     End With
944:     pFieldsEdit.AddField newField
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
961:     Set newField = New Field
962:     Set newFieldEdit = newField
963:     With newFieldEdit
964:       .Name = c_DefaultFld_GridID
965:       .AliasName = "GridIdentifier"
966:       .Type = esriFieldTypeString
967:       .IsNullable = True
968:       .Length = 50
969:     End With
970:     pFieldsEdit.AddField newField
    ' Field: ROW NUMBER
972:     Set newField = New Field
973:     Set newFieldEdit = newField
974:     With newFieldEdit
975:       .Name = c_DefaultFld_RowNum
976:       .AliasName = "Row Number"
977:       .Type = esriFieldTypeInteger
978:       .IsNullable = True
979:     End With
980:     pFieldsEdit.AddField newField
    ' Field: COLUMN NUMBER
982:     Set newField = New Field
983:     Set newFieldEdit = newField
984:     With newFieldEdit
985:       .Name = c_DefaultFld_ColNum
986:       .AliasName = "Column Number"
987:       .Type = esriFieldTypeInteger
988:       .IsNullable = True
989:     End With
990:     pFieldsEdit.AddField newField
    ' Field: SCALE
992:     Set newField = New Field
993:     Set newFieldEdit = newField
994:     With newFieldEdit
995:       .Name = c_DefaultFld_Scale
996:       .AliasName = "Plot Scale"
997:       .Type = esriFieldTypeDouble
998:       .IsNullable = True
999:       .Precision = 18
1000:       .Scale = 11
1001:     End With
1002:     pFieldsEdit.AddField newField
    ' Return
1004:     Set CreateTheFields = pFieldsEdit
End Function

