VERSION 5.00
Begin VB.Form frmSMapSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Strip Map Wizard"
   ClientHeight    =   7116
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   10140
   Icon            =   "frmSMapSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7116
   ScaleWidth      =   10140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraDataFrameSize 
      Height          =   4695
      Left            =   5040
      TabIndex        =   37
      Top             =   1560
      Width           =   4815
      Begin VB.OptionButton optGridSize 
         Caption         =   "Use the current Data Frame"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   42
         Top             =   1200
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.OptionButton optGridSize 
         Caption         =   "Specify the Data Frame size (in Layout Units)"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   41
         Top             =   1800
         Width           =   4095
      End
      Begin VB.TextBox txtManualGridWidth 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   40
         Text            =   "0"
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox txtManualGridHeight 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   39
         Text            =   "0"
         Top             =   3240
         Width           =   1215
      End
      Begin VB.ComboBox cmbGridSizeUnits 
         Height          =   315
         ItemData        =   "frmSMapSettings.frx":014A
         Left            =   3360
         List            =   "frmSMapSettings.frx":0154
         TabIndex        =   38
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "Units:"
         Height          =   255
         Left            =   2760
         TabIndex        =   48
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label23 
         Caption         =   "Note: For best results you should update the Data Frame size in the Layout to match."
         Height          =   495
         Left            =   600
         TabIndex        =   47
         Top             =   2160
         Width           =   3735
      End
      Begin VB.Label Label21 
         Caption         =   $"frmSMapSettings.frx":0163
         Height          =   735
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Width : "
         Height          =   255
         Left            =   480
         TabIndex        =   45
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label lblFrameHeight 
         Alignment       =   1  'Right Justify
         Caption         =   "Height : "
         Height          =   255
         Left            =   360
         TabIndex        =   44
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label lblCurrFrameName 
         Caption         =   "Current Frame"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   43
         Top             =   1200
         Width           =   1575
      End
   End
   Begin VB.Frame fraScaleStart 
      Height          =   4695
      Left            =   0
      TabIndex        =   23
      Top             =   5400
      Width           =   4815
      Begin VB.TextBox txtAbsoluteGridHeight 
         Height          =   285
         Left            =   2760
         TabIndex        =   29
         Text            =   "0"
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtAbsoluteGridWidth 
         Height          =   285
         Left            =   2760
         TabIndex        =   28
         Text            =   "0"
         Top             =   1560
         Width           =   975
      End
      Begin VB.OptionButton optScaleSource 
         Caption         =   "Absolute Size"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   27
         Top             =   1560
         Width           =   1455
      End
      Begin VB.OptionButton optScaleSource 
         Caption         =   "Manual Map Scale"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   26
         Top             =   1080
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton optScaleSource 
         Caption         =   "Current Map Scale"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   25
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtManualMapScale 
         Height          =   285
         Left            =   2760
         TabIndex        =   24
         Text            =   "0"
         Top             =   1065
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   $"frmSMapSettings.frx":0201
         Height          =   1215
         Left            =   480
         TabIndex        =   49
         Top             =   2400
         Width           =   4095
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Height :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   36
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Width :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   35
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Set the Scale / Size for each of the grids."
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label lblCurrentMapScale 
         Caption         =   "5,000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   33
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "1 :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   32
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "1 :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   31
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblMapUnits 
         Caption         =   "esriUnits"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3750
         TabIndex        =   30
         Top             =   1560
         Width           =   735
      End
   End
   Begin VB.Frame fraAttributes 
      Height          =   4695
      Left            =   5040
      TabIndex        =   4
      Top             =   0
      Width           =   4815
      Begin VB.ComboBox cmbFieldMapScale 
         Height          =   315
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   51
         Top             =   3960
         Width           =   2535
      End
      Begin VB.ComboBox cmbFieldStripMapName 
         Height          =   315
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   22
         Top             =   1560
         Width           =   2535
      End
      Begin VB.ComboBox cmbFieldGridAngle 
         Height          =   315
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   2520
         Width           =   2535
      End
      Begin VB.ComboBox cmbFieldSeriesNumber 
         Height          =   315
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label Label16 
         Caption         =   $"frmSMapSettings.frx":0302
         Height          =   735
         Left            =   120
         TabIndex        =   52
         Top             =   3120
         Width           =   4455
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Map Scale:"
         Height          =   375
         Left            =   720
         TabIndex        =   50
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Strip Map Name:"
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Map Angle:"
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Number in the Series:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   $"frmSMapSettings.frx":03B1
         Height          =   735
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   4455
      End
      Begin VB.Label Label13 
         Caption         =   "Assign roles to field names."
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Frame fraDestinationFeatureClass 
      Height          =   4695
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4815
      Begin VB.TextBox txtStripMapSeriesName 
         Height          =   315
         Left            =   2040
         TabIndex        =   53
         Text            =   "Strip Map Name"
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton cmdSetNewGridLayer 
         Height          =   315
         Left            =   4320
         Picture         =   "frmSMapSettings.frx":0449
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Set new Grid Layer"
         Top             =   2400
         Width           =   315
      End
      Begin VB.TextBox txtNewGridLayer 
         Height          =   315
         Left            =   2040
         TabIndex        =   18
         Top             =   2400
         Width           =   2295
      End
      Begin VB.OptionButton optLayerSource 
         Caption         =   "Create a new Layer:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   2400
         Width           =   1815
      End
      Begin VB.OptionButton optLayerSource 
         Caption         =   "Use existing Layer:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   2040
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.ComboBox cmbPolygonLayers 
         Height          =   315
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1995
         Width           =   2535
      End
      Begin VB.CheckBox chkRemovePreviousGrids 
         Caption         =   "Clear existing grids.  This will delete all the current"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   3000
         Width           =   4335
      End
      Begin VB.CheckBox chkFlipLine 
         Caption         =   "Flip the line.  This will reverse the orientation of the"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   3720
         Width           =   3975
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Name:"
         Height          =   255
         Left            =   960
         TabIndex        =   55
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   $"frmSMapSettings.frx":08C3
         Height          =   400
         Left            =   120
         TabIndex        =   54
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label Label6 
         Caption         =   $"frmSMapSettings.frx":094A
         Height          =   615
         Left            =   600
         TabIndex        =   21
         Top             =   3945
         Width           =   4095
      End
      Begin VB.Label lblClearExistingGridsPart2 
         Caption         =   "features in the feature class."
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   3225
         Width           =   3855
      End
      Begin VB.Label Label5 
         Caption         =   $"frmSMapSettings.frx":09E3
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   4575
      End
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "< Back"
      Height          =   375
      Left            =   1440
      TabIndex        =   15
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next >"
      Height          =   375
      Left            =   2540
      TabIndex        =   14
      Top             =   4800
      Width           =   1095
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
Attribute VB_Name = "frmSMapSettings"
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
Public StripMapSettings As clsCreateStripMap

Private m_Polyline As IPolyline
Private m_bIsGeoDatabase As Boolean
Private m_FileType As intersectFileType
Private m_OutputLayer As String
Private m_OutputDataset As String
Private m_OutputFClass As String
Private m_Step As Integer

Private Const c_DefaultFld_StripMapName = "SMAP_NAME"
Private Const c_DefaultFld_SeriesNum = "SMAP_NUM"
Private Const c_DefaultFld_MapAngle = "SMAP_ANGLE"
Private Const c_DefaultFld_MapScale = "SMAP_SCALE"

Private Sub SetControlsState()
    Dim dScale As Double
    Dim dGHeight As Double
    Dim dGWidth As Double
    Dim dStartX As Double
    Dim dStartY As Double
    Dim dEndX As Double
    Dim dEndY As Double
    Dim bValidName As Boolean
    Dim bValidScale As Boolean
    Dim bValidSize As Boolean
    Dim bValidTarget As Boolean
    Dim bValidRequiredFields As Boolean
    Dim bPolylineWithinDataset As Boolean
    Dim bNewFClassSet As Boolean
    Dim bCreatingNewFClass As Boolean
    Dim bDuplicateFieldsSelected As Boolean
    Dim pFL As IFeatureLayer
    Dim pDatasetExtent As IEnvelope
    Dim dAWidth As Double
    Dim dAHeight As Double
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
    If Len(txtAbsoluteGridHeight.Text) = 0 Then
        dAHeight = 0
    Else
        dAHeight = CDbl(txtAbsoluteGridHeight.Text)
    End If
    If Len(txtAbsoluteGridWidth.Text) = 0 Then
        dAWidth = 0
    Else
        dAWidth = CDbl(txtAbsoluteGridWidth.Text)
    End If
i = 1

    ' Calc values
    bValidName = Len(txtStripMapSeriesName.Text) > 0
    bValidScale = (optScaleSource(0).Value And CDbl(lblCurrentMapScale.Caption) > 0) Or _
                  (optScaleSource(1).Value And dScale > 0) Or _
                  (optScaleSource(2).Value And dAHeight > 0 And dAWidth > 0)
    bValidSize = (optGridSize(0).Value) Or _
                 (optGridSize(1).Value And dGHeight > 0 And dGWidth > 0) Or _
                 (optScaleSource(2).Value And CDbl(txtManualGridWidth.Text) > 0)
    bCreatingNewFClass = optLayerSource(1).Value
    bNewFClassSet = (Len(txtNewGridLayer.Text) > 0)
    bValidTarget = (cmbPolygonLayers.ListIndex > 0) Or (bCreatingNewFClass And bNewFClassSet)
    bValidRequiredFields = (cmbFieldStripMapName.ListIndex > 0) And _
                           (cmbFieldGridAngle.ListIndex > 0) And _
                           (cmbFieldSeriesNumber.ListIndex > 0)
i = 2
    If bValidTarget And (Not bCreatingNewFClass) Then
        Set pFL = FindFeatureLayerByName(cmbPolygonLayers.List(cmbPolygonLayers.ListIndex), m_Application)
        If pFL.FeatureClass.FeatureDataset Is Nothing Then
            bPolylineWithinDataset = True
        Else
            Set pDatasetExtent = GetValidExtentForLayer(pFL)
            bPolylineWithinDataset = (m_Polyline.envelope.XMin >= pDatasetExtent.XMin And m_Polyline.envelope.XMax <= pDatasetExtent.XMax) _
                     And (m_Polyline.envelope.YMin >= pDatasetExtent.YMin And m_Polyline.envelope.YMax <= pDatasetExtent.YMax)
        End If
    ElseIf bValidTarget And bCreatingNewFClass Then
        bPolylineWithinDataset = True
    End If
    Dim a As Long, b As Long, c As Long
    a = cmbFieldGridAngle.ListIndex
    b = cmbFieldMapScale.ListIndex
    c = cmbFieldSeriesNumber.ListIndex
    bDuplicateFieldsSelected = (a > 0 And (a = b Or a = c)) _
                            Or (b > 0 And (b = c))
i = 3
    
    ' Set states
    Select Case m_Step
        Case 0:     ' Set the target feature layer
            cmdBack.Enabled = False
            cmdNext.Enabled = bValidTarget And bValidName
            cmdNext.Caption = "Next >"
            cmbPolygonLayers.Enabled = Not bCreatingNewFClass
            chkRemovePreviousGrids.Enabled = Not bCreatingNewFClass
            lblClearExistingGridsPart2.Enabled = Not bCreatingNewFClass
            cmdSetNewGridLayer.Enabled = bCreatingNewFClass
        Case 1:     ' Set the fields to populate
            cmdBack.Enabled = True
            cmdNext.Enabled = (bValidRequiredFields And Not bDuplicateFieldsSelected)
            cmbFieldStripMapName.Enabled = Not bCreatingNewFClass
            cmbFieldGridAngle.Enabled = Not bCreatingNewFClass
            cmbFieldMapScale.Enabled = Not bCreatingNewFClass
            cmbFieldSeriesNumber.Enabled = Not bCreatingNewFClass
        Case 2:     ' Set the scale / starting_coords
            cmdBack.Enabled = True
            cmdNext.Enabled = bValidScale And bPolylineWithinDataset
            cmdNext.Caption = "Next >"
        Case 3:     ' Set the dataframe properties
            cmdBack.Enabled = True
            cmdNext.Enabled = bValidSize
            cmdNext.Caption = "Finish"
            txtManualGridHeight.Enabled = Not (optScaleSource(2).Value)
            txtManualGridHeight.Locked = (optScaleSource(2).Value)
            lblFrameHeight.Enabled = Not (optScaleSource(2).Value)
            optGridSize(0).Enabled = Not (optScaleSource(2).Value)
        Case Else:
            cmdBack.Enabled = False
            cmdNext.Enabled = False
    End Select
i = 4
    
    txtManualMapScale.Enabled = optScaleSource(1).Value
    txtManualGridWidth.Enabled = optGridSize(1).Value
    txtManualGridHeight.Enabled = optGridSize(1).Value
    cmbGridSizeUnits.Enabled = optGridSize(1).Value
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

Private Sub cmbFieldStripMapName_Click()
    SetControlsState
End Sub

Private Sub cmbFieldMapScale_Click()
    SetControlsState
End Sub

Private Sub cmbFieldSeriesNumber_Click()
    SetControlsState
End Sub

Private Sub cmbFieldGridAngle_Click()
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
        cmbFieldMapScale.Clear
        cmbFieldStripMapName.Clear
        cmbFieldSeriesNumber.Clear
        cmbFieldGridAngle.Clear
        cmbFieldStripMapName.AddItem "<None>"
        cmbFieldGridAngle.AddItem "<None>"
        cmbFieldMapScale.AddItem "<None>"
        cmbFieldSeriesNumber.AddItem "<None>"
        For lLoop = 0 To pFields.FieldCount - 1
            If pFields.Field(lLoop).Type = esriFieldTypeString Then
                cmbFieldStripMapName.AddItem pFields.Field(lLoop).Name
            ElseIf pFields.Field(lLoop).Type = esriFieldTypeDouble Or _
                   pFields.Field(lLoop).Type = esriFieldTypeInteger Or _
                   pFields.Field(lLoop).Type = esriFieldTypeSmallInteger Or _
                   pFields.Field(lLoop).Type = esriFieldTypeSingle Then
                cmbFieldMapScale.AddItem pFields.Field(lLoop).Name
                cmbFieldGridAngle.AddItem pFields.Field(lLoop).Name
                cmbFieldSeriesNumber.AddItem pFields.Field(lLoop).Name
            End If
        Next
        cmbFieldStripMapName.ListIndex = 0
        cmbFieldGridAngle.ListIndex = 0
        cmbFieldMapScale.ListIndex = 0
        cmbFieldSeriesNumber.ListIndex = 0
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
    Set Me.StripMapSettings = Nothing
    Me.Hide
End Sub

Private Sub CollateStripMapSettings()
    Dim pMx As IMxDocument
    Dim pCreateSMap As New clsCreateStripMap
    Dim pFrameElement As IElement
    Dim sDestLayerName As String
    Dim lLoop As Long
    ' Populate class
    pCreateSMap.StripMapName = txtStripMapSeriesName.Text
    pCreateSMap.FlipPolyline = (chkFlipLine.Value = vbChecked)
    If (optScaleSource(0).Value) Then
        pCreateSMap.MapScale = CDbl(lblCurrentMapScale.Caption)
    ElseIf (optScaleSource(1).Value) Then
        pCreateSMap.MapScale = CDbl(txtManualMapScale.Text)
    End If
    If (optGridSize(0).Value) Then
        Set pFrameElement = GetDataFrameElement(GetActiveDataFrameName(m_Application), m_Application)
        pCreateSMap.FrameWidthInPageUnits = pFrameElement.Geometry.envelope.Width
        pCreateSMap.FrameHeightInPageUnits = pFrameElement.Geometry.envelope.Height
    Else
        pCreateSMap.FrameWidthInPageUnits = CDbl(txtManualGridWidth.Text)
        pCreateSMap.FrameHeightInPageUnits = CDbl(txtManualGridHeight.Text)
    End If
    If (optScaleSource(2).Value) Then
        Dim dConvertPageToMapUnits As Double, dGridToFrameRatio As Double
        dConvertPageToMapUnits = CalculatePageToMapRatio(m_Application) 'NATHAN FIX THIS
        pCreateSMap.FrameWidthInPageUnits = CDbl(txtManualGridWidth.Text)
        pCreateSMap.FrameHeightInPageUnits = CDbl(txtManualGridHeight.Text)
        If pCreateSMap.FrameWidthInPageUnits >= pCreateSMap.FrameHeightInPageUnits Then
            dGridToFrameRatio = CDbl(txtAbsoluteGridWidth.Text) / pCreateSMap.FrameWidthInPageUnits
        Else
            dGridToFrameRatio = CDbl(txtAbsoluteGridHeight.Text) / pCreateSMap.FrameHeightInPageUnits
        End If
        pCreateSMap.MapScale = dGridToFrameRatio * dConvertPageToMapUnits
    End If
    sDestLayerName = cmbPolygonLayers.List(cmbPolygonLayers.ListIndex)
    If optLayerSource(0).Value Then
        Set pCreateSMap.DestinationFeatureLayer = FindFeatureLayerByName(sDestLayerName, m_Application)
    End If
    pCreateSMap.FieldNameStripMapName = cmbFieldStripMapName.List(cmbFieldStripMapName.ListIndex)
    pCreateSMap.FieldNameMapAngle = cmbFieldGridAngle.List(cmbFieldGridAngle.ListIndex)
    pCreateSMap.FieldNameNumberInSeries = cmbFieldSeriesNumber.List(cmbFieldSeriesNumber.ListIndex)
    If cmbFieldMapScale.ListIndex > 0 Then pCreateSMap.FieldNameScale = cmbFieldMapScale.List(cmbFieldMapScale.ListIndex)
    pCreateSMap.RemoveCurrentGrids = (chkRemovePreviousGrids.Value = vbChecked)
    Set pCreateSMap.StripMapRoute = m_Polyline
    ' Place grid settings on Public form property (so calling function can use them)
    Set Me.StripMapSettings = pCreateSMap
End Sub

Private Sub cmdNext_Click()
    Dim pMx As IMxDocument
    Dim pFeatureLayer As IFeatureLayer
    Dim pOutputFClass As IFeatureClass
    Dim pNewFields As IFields
    
    On Error GoTo eh
    ' Step
    m_Step = m_Step + 1
    ' If we're creating a new fclass, we can skip a the 'Set Field Roles' step
    If m_Step = 1 And (optLayerSource(1).Value) Then
        m_Step = m_Step + 1
    End If
    ' If FINISH
    If m_Step >= 4 Then
        Set pMx = m_Application.Document
        RemoveGraphicsByName pMx
        CollateStripMapSettings
        ' If creating a new layer
        If optLayerSource(1).Value Then
            ' Create the feature class
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
            ' Add the new layer to arcmap & reset the StripMapSettings object to point at it
            pMx.FocusMap.AddLayer pFeatureLayer
            Set StripMapSettings.DestinationFeatureLayer = pFeatureLayer
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
    Dim sErrMsg As String
    On Error GoTo eh
    
    sErrMsg = CreateStripMapPolyline
    If Len(sErrMsg) > 0 Then
        MsgBox sErrMsg, vbCritical, "Create Map Grids"
        Unload Me
        Exit Sub
    End If
    Set pMx = m_Application.Document
    Me.Height = 5665
    Me.Width = 4935
    m_Step = 0
    LoadLayersComboBox
    LoadUnitsComboBox
    lblCurrFrameName.Caption = GetActiveDataFrameName(m_Application)
    If pMx.FocusMap.MapUnits = esriUnknownUnits Then
        MsgBox "Error: The map has unknown units and therefore cannot calculate a Scale." _
            & vbCrLf & "Cannot create Map Grids at this time.", vbCritical, "Create Map Grids"
        Unload Me
        Exit Sub
    End If
    lblMapUnits.Caption = GetUnitsDescription(pMx.FocusMap.MapUnits)
    lblCurrentMapScale.Caption = Format(pMx.FocusMap.MapScale, "#,###,##0")
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
    cmbPolygonLayers.Clear
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
            End If
        End If
    Next
    cmbPolygonLayers.ListIndex = 0
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
    Set StripMapSettings = Nothing
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
        cmbFieldStripMapName.Clear
        cmbFieldGridAngle.Clear
        cmbFieldSeriesNumber.Clear
        cmbFieldMapScale.Clear
        cmbFieldStripMapName.AddItem "<None>"
        cmbFieldGridAngle.AddItem "<None>"
        cmbFieldSeriesNumber.AddItem "<None>"
        cmbFieldMapScale.AddItem "<None>"
        cmbFieldStripMapName.AddItem c_DefaultFld_StripMapName
        cmbFieldGridAngle.AddItem c_DefaultFld_MapAngle
        cmbFieldSeriesNumber.AddItem c_DefaultFld_SeriesNum
        cmbFieldMapScale.AddItem c_DefaultFld_MapScale
        cmbFieldStripMapName.ListIndex = 1
        cmbFieldGridAngle.ListIndex = 1
        cmbFieldSeriesNumber.ListIndex = 1
        cmbFieldMapScale.ListIndex = 1
    End If
    SetControlsState
End Sub

Private Sub optScaleSource_Click(Index As Integer)
    If Index = 0 Then
        SetCurrentMapScaleCaption
    ElseIf Index = 2 Then
        optGridSize(1).Value = True
    End If
    SetControlsState
End Sub

Private Sub txtAbsoluteGridHeight_Change()
    SetControlsState
End Sub

Private Sub txtAbsoluteGridHeight_KeyPress(KeyAscii As Integer)
    ' If a non-numeric (that is not a decimal point)
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
     And KeyAscii <> Asc(".") _
     And Chr(KeyAscii) <> vbBack Then
        ' Do not allow this button to work
        KeyAscii = 0
    ' If a decimal point, make sure we only ever get one of them
    ElseIf KeyAscii = Asc(".") Then
        If InStr(txtAbsoluteGridHeight.Text, ".") > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtAbsoluteGridWidth_Change()
    SetControlsState
End Sub

Private Sub txtAbsoluteGridWidth_KeyPress(KeyAscii As Integer)
    ' If a non-numeric (that is not a decimal point)
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
     And KeyAscii <> Asc(".") _
     And Chr(KeyAscii) <> vbBack Then
        ' Do not allow this button to work
        KeyAscii = 0
    ' If a decimal point, make sure we only ever get one of them
    ElseIf KeyAscii = Asc(".") Then
        If InStr(txtAbsoluteGridWidth.Text, ".") > 0 Then
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
    If IsNumeric(txtManualGridWidth.Text) And optScaleSource(2).Value Then
        Dim dRatio As Double, dGridWidth As Double
        dGridWidth = CDbl(txtManualGridWidth.Text)
        dRatio = CDbl(txtAbsoluteGridHeight.Text) / CDbl(txtAbsoluteGridWidth.Text)
        txtManualGridHeight.Text = CStr(dRatio * dGridWidth)
    End If
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

Private Sub SetVisibleControls(iStep As Integer)
    ' Hide all
    fraAttributes.Visible = False
    fraDataFrameSize.Visible = False
    fraDestinationFeatureClass.Visible = False
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
        Case Else:
            MsgBox "Invalid Step Value : " & iStep
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
    ' Field: OID  -------------------------
    Set newField = New Field
    Set newFieldEdit = newField
    With newFieldEdit
        .Name = "OID"
        .Type = esriFieldTypeOID
        .AliasName = "Object ID"
        .IsNullable = False
    End With
    pFieldsEdit.AddField newField
    ' Field: STRIP MAP NAME -------------------------
    Set newField = New Field
    Set newFieldEdit = newField
    With newFieldEdit
      .Name = c_DefaultFld_StripMapName
      .AliasName = "StripMapName"
      .Type = esriFieldTypeString
      .IsNullable = True
      .Length = 50
    End With
    pFieldsEdit.AddField newField
    ' Field: MAP ANGLE -------------------------
    Set newField = New Field
    Set newFieldEdit = newField
    With newFieldEdit
      .Name = c_DefaultFld_MapAngle
      .AliasName = "Map Angle"
      .Type = esriFieldTypeInteger
      .IsNullable = True
    End With
    pFieldsEdit.AddField newField
    ' Field: GRID NUMBER -------------------------
    Set newField = New Field
    Set newFieldEdit = newField
    With newFieldEdit
      .Name = c_DefaultFld_SeriesNum
      .AliasName = "Number In Series"
      .Type = esriFieldTypeInteger
      .IsNullable = True
    End With
    pFieldsEdit.AddField newField
    ' Field: SCALE -------------------------
    Set newField = New Field
    Set newFieldEdit = newField
    With newFieldEdit
      .Name = c_DefaultFld_MapScale
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
    Set pMx = pApp.Document
    Set pSR = pMx.FocusMap.SpatialReference
    If TypeOf pSR Is IProjectedCoordinateSystem Then
        Set pPCS = pSR
        dMetersPerUnit = pPCS.CoordinateUnit.MetersPerUnit
    Else
        dMetersPerUnit = 1
    End If
    Set pPage = pMx.PageLayout.Page
    pPageUnits = pPage.Units
    Select Case pPageUnits
        Case esriInches: CalculatePageToMapRatio = dMetersPerUnit / (1 / 12 * 0.304800609601219)
        Case esriFeet: CalculatePageToMapRatio = dMetersPerUnit / (0.304800609601219)
        Case esriCentimeters: CalculatePageToMapRatio = dMetersPerUnit / (1 / 100)
        Case esriMeters: CalculatePageToMapRatio = dMetersPerUnit / (1)
        Case Else:
            MsgBox "Warning: Only the following Page (Layout) Units are supported by this tool:" _
                & vbCrLf & " - Inches, Feet, Centimeters, Meters" _
                & vbCrLf & vbCrLf & "Calculating as though Page Units are in Inches..."
            CalculatePageToMapRatio = dMetersPerUnit / (1 / 12 * 0.304800609601219)
    End Select
    Exit Function
eh:
    CalculatePageToMapRatio = 1
    MsgBox "Error in CalculatePageToMapRatio" & vbCrLf & Err.Description
End Function

Private Function ReturnMax(dDouble1 As Double, dDouble2 As Double) As Double
    If dDouble1 >= dDouble2 Then
        ReturnMax = dDouble1
    Else
        ReturnMax = dDouble2
    End If
End Function

Private Function CreateStripMapPolyline() As String
    Dim pMx As IMxDocument
    Dim pFL As IFeatureLayer
    Dim pFC As IEnumFeature
    Dim pF As IFeature
    Dim pPolyline As IPolyline
    Dim pTmpPolyline As IPolyline
    Dim pTopoSimplify As ITopologicalOperator
    Dim pTopoUnion As ITopologicalOperator
    Dim pGeoColl As IGeometryCollection
    
    On Error GoTo eh
    
    ' Init
    Set pMx = m_Application.Document
    Set pFC = pMx.FocusMap.FeatureSelection
    Set pF = pFC.Next
    If pF Is Nothing Then
        CreateStripMapPolyline = "Requires selected polyline features/s."
        Exit Function
    End If
    ' Make polyline
    Set pPolyline = New Polyline
    While Not pF Is Nothing
        If pF.Shape.GeometryType = esriGeometryPolyline Then
            Set pTmpPolyline = pF.ShapeCopy
            Set pTopoSimplify = pTmpPolyline
            pTopoSimplify.Simplify
            Set pTopoUnion = pPolyline
            Set pPolyline = pTopoUnion.Union(pTopoSimplify)
            Set pTopoSimplify = pPolyline
            pTopoSimplify.Simplify
        End If
        Set pF = pFC.Next
    Wend
    ' Check polyline for beinga single, connected polyline (Path)
    Set pGeoColl = pPolyline
    If pGeoColl.GeometryCount = 0 Then
        CreateStripMapPolyline = "Requires selected polyline features/s."
        Exit Function
    ElseIf pGeoColl.GeometryCount > 1 Then
        CreateStripMapPolyline = "Cannot process the StripMap - multi-part polyline created." _
            & vbCrLf & "Check for non-connected segments, overlaps or loops."
        Exit Function
    End If
    ' Give option to flip
    Perm_DrawPoint pPolyline.FromPoint, , 0, 255, 0, 20
    Perm_DrawTextFromPoint pPolyline.FromPoint, "START", , , , , 20
    Perm_DrawPoint pPolyline.ToPoint, , 255, 0, 0, 20
    Perm_DrawTextFromPoint pPolyline.ToPoint, "END", , , , , 20
    pMx.ActiveView.PartialRefresh esriViewGraphics, Nothing, Nothing
    
    Set m_Polyline = pPolyline
    
    CreateStripMapPolyline = ""
    
    Exit Function
    Resume
eh:
    CreateStripMapPolyline = "Error in CreateStripMapPolyline : " & Err.Description
End Function

Public Sub Perm_DrawPoint(ByVal pPoint As IPoint, _
            Optional sElementName As String = "DEMO_TEMPORARY", _
            Optional dRed As Double = 255, Optional dGreen As Double = 0, _
            Optional dBlue As Double = 0, Optional dSize As Double = 6)
' Add a permanent graphic dot on the display at the given point location
    Dim pColor As IRgbColor
    Dim pMarker As ISimpleMarkerSymbol
    Dim pGLayer As IGraphicsLayer
    Dim pGCon As IGraphicsContainer
    Dim pElement As IElement
    Dim pMarkerElement As IMarkerElement
    Dim pElementProp As IElementProperties
    Dim pMx As IMxDocument
    
    ' Init
    Set pMx = m_Application.Document
    Set pGLayer = pMx.FocusMap.BasicGraphicsLayer
    Set pGCon = pGLayer
    Set pElement = New MarkerElement
    pElement.Geometry = pPoint
    Set pMarkerElement = pElement
    
    ' Set the symbol
    Set pColor = New RgbColor
    pColor.Red = dRed
    pColor.Green = dGreen
    pColor.Blue = dBlue
    Set pMarker = New SimpleMarkerSymbol
    With pMarker
        .Color = pColor
        .Size = dSize
    End With
    pMarkerElement.Symbol = pMarker
    
    ' Add the graphic
    Set pElementProp = pElement
    pElementProp.Name = sElementName
    pGCon.AddElement pElement, 0
End Sub

Public Sub Perm_DrawLineFromPoints(ByVal pFromPoint As IPoint, ByVal pToPoint As IPoint, _
            Optional sElementName As String = "DEMO_TEMPORARY", _
            Optional dRed As Double = 0, Optional dGreen As Double = 0, _
            Optional dBlue As Double = 255, Optional dSize As Double = 1)
' Add a permanent graphic line on the display, using the From and To points supplied
    Dim pLnSym As ISimpleLineSymbol
    Dim pLine1 As ILine
    Dim pSeg1 As ISegment
    Dim pPolyline As ISegmentCollection
    Dim myColor As IRgbColor
    Dim pSym As ISymbol
    Dim pLineSym As ILineSymbol
    Dim pGLayer As IGraphicsLayer
    Dim pGCon As IGraphicsContainer
    Dim pElement As IElement
    Dim pLineElement As ILineElement
    Dim pElementProp As IElementProperties
    Dim pMx As IMxDocument
    
    ' Init
    Set pMx = m_Application.Document
    Set pGLayer = pMx.FocusMap.BasicGraphicsLayer
    Set pGCon = pGLayer
    Set pElement = New LineElement
    
    ' Set the line symbol
    Set pLnSym = New SimpleLineSymbol
    Set myColor = New RgbColor
    myColor.Red = dRed
    myColor.Green = dGreen
    myColor.Blue = dBlue
    pLnSym.Color = myColor
    pLnSym.Width = dSize
    
    ' Create a standard polyline (via 2 points)
    Set pLine1 = New esriGeometry.Line
    pLine1.PutCoords pFromPoint, pToPoint
    Set pSeg1 = pLine1
    Set pPolyline = New Polyline
    pPolyline.AddSegment pSeg1
    pElement.Geometry = pPolyline
    Set pLineElement = pElement
    pLineElement.Symbol = pLnSym
    
    ' Add the graphic
    Set pElementProp = pElement
    pElementProp.Name = sElementName
    pGCon.AddElement pElement, 0
End Sub

Public Sub Perm_DrawTextFromPoint(pPoint As IPoint, sText As String, _
            Optional sElementName As String = "DEMO_TEMPORARY", _
            Optional dRed As Double = 50, Optional dGreen As Double = 50, _
            Optional dBlue As Double = 50, Optional dSize As Double = 10)
' Add permanent graphic text on the display at the given point location
    Dim myTxtSym As ITextSymbol
    Dim myColor As IRgbColor
    Dim pGLayer As IGraphicsLayer
    Dim pGCon As IGraphicsContainer
    Dim pElement As IElement
    Dim pTextElement As ITextElement
    Dim pElementProp As IElementProperties
    Dim pMx As IMxDocument
    
    ' Init
    Set pMx = m_Application.Document
    Set pGLayer = pMx.FocusMap.BasicGraphicsLayer
    Set pGCon = pGLayer
    Set pElement = New TextElement
    pElement.Geometry = pPoint
    Set pTextElement = pElement
    
    ' Create the text symbol
    Set myTxtSym = New TextSymbol
    Set myColor = New RgbColor
    myColor.Red = dRed
    myColor.Green = dGreen
    myColor.Blue = dBlue
    myTxtSym.Color = myColor
    myTxtSym.Size = dSize
    myTxtSym.HorizontalAlignment = esriTHACenter
    pTextElement.Symbol = myTxtSym
    pTextElement.Text = sText
    
    ' Add the graphic
    Set pElementProp = pElement
    pElementProp.Name = sElementName
    pGCon.AddElement pElement, 0
End Sub

Public Sub RemoveGraphicsByName(pMxDoc As IMxDocument, _
            Optional sPrefix As String = "DEMO_TEMPORARY")
' Delete all graphics with our prefix from ArcScene
    Dim pElement As IElement
    Dim pElementProp As IElementProperties
    Dim sLocalPrefix As String
    Dim pGLayer As IGraphicsLayer
    Dim pGCon As IGraphicsContainer
    Dim lCount As Long
    
    On Error GoTo ErrorHandler
    
    ' Init and switch OFF the updating of the TOC
    pMxDoc.DelayUpdateContents = True
    Set pGLayer = pMxDoc.FocusMap.BasicGraphicsLayer
    Set pGCon = pGLayer
    pGCon.Next
    
    ' Delete all the graphic elements that we created (identify by the name prefix)
    pGCon.Reset
    Set pElement = pGCon.Next
    While Not pElement Is Nothing
        If TypeOf pElement Is IElement Then
            Set pElementProp = pElement
            If (Left(pElementProp.Name, Len(sPrefix)) = sPrefix) Then
                pGCon.DeleteElement pElement
            End If
        End If
        Set pElement = pGCon.Next
    Wend
    
    ' Switch ON the updating of the TOC, refresh
    pMxDoc.DelayUpdateContents = False
    pMxDoc.ActiveView.Refresh
    
    Exit Sub
ErrorHandler:
    MsgBox "Error in RemoveGraphicsByName: " & Err.Description, , "RemoveGraphicsByName"
End Sub

Private Sub txtStripMapSeriesName_Change()
    SetControlsState
End Sub
