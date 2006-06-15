VERSION 5.00
Begin VB.Form frmSMapSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Strip Map Wizard"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10140
   Icon            =   "frmSMapSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
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
            Size            =   8.25
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
            Size            =   8.25
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
            Size            =   8.25
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
            Size            =   8.25
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
            Size            =   8.25
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
            Size            =   8.25
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
            Size            =   8.25
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
45:     If Len(lblCurrentMapScale.Caption) = 0 Then lblCurrentMapScale.Caption = "0"
46:     If Len(txtManualMapScale.Text) = 0 Then
47:         dScale = 0
48:     Else
49:         dScale = CDbl(txtManualMapScale.Text)
50:     End If
51:     If Len(txtManualGridHeight.Text) = 0 Then
52:         dGHeight = 0
53:     Else
54:         dGHeight = CDbl(txtManualGridHeight.Text)
55:     End If
56:     If Len(txtManualGridWidth.Text) = 0 Then
57:         dGWidth = 0
58:     Else
59:         dGWidth = CDbl(txtManualGridWidth.Text)
60:     End If
61:     If Len(txtAbsoluteGridHeight.Text) = 0 Then
62:         dAHeight = 0
63:     Else
64:         dAHeight = CDbl(txtAbsoluteGridHeight.Text)
65:     End If
66:     If Len(txtAbsoluteGridWidth.Text) = 0 Then
67:         dAWidth = 0
68:     Else
69:         dAWidth = CDbl(txtAbsoluteGridWidth.Text)
70:     End If
71: i = 1

    ' Calc values
74:     bValidName = Len(txtStripMapSeriesName.Text) > 0
75:     bValidScale = (optScaleSource(0).Value And CDbl(lblCurrentMapScale.Caption) > 0) Or _
                  (optScaleSource(1).Value And dScale > 0) Or _
                  (optScaleSource(2).Value And dAHeight > 0 And dAWidth > 0)
78:     bValidSize = (optGridSize(0).Value) Or _
                 (optGridSize(1).Value And dGHeight > 0 And dGWidth > 0) Or _
                 (optScaleSource(2).Value And CDbl(txtManualGridWidth.Text) > 0)
81:     bCreatingNewFClass = optLayerSource(1).Value
82:     bNewFClassSet = (Len(txtNewGridLayer.Text) > 0)
83:     bValidTarget = (cmbPolygonLayers.ListIndex > 0) Or (bCreatingNewFClass And bNewFClassSet)
84:     bValidRequiredFields = (cmbFieldStripMapName.ListIndex > 0) And _
                           (cmbFieldGridAngle.ListIndex > 0) And _
                           (cmbFieldSeriesNumber.ListIndex > 0)
87: i = 2
88:     If bValidTarget And (Not bCreatingNewFClass) Then
89:         Set pFL = FindFeatureLayerByName(cmbPolygonLayers.List(cmbPolygonLayers.ListIndex), m_Application)
90:         If pFL.FeatureClass.FeatureDataset Is Nothing Then
91:             bPolylineWithinDataset = True
92:         Else
93:             Set pDatasetExtent = GetValidExtentForLayer(pFL)
94:             bPolylineWithinDataset = (m_Polyline.Envelope.XMin >= pDatasetExtent.XMin And m_Polyline.Envelope.XMax <= pDatasetExtent.XMax) _
                     And (m_Polyline.Envelope.YMin >= pDatasetExtent.YMin And m_Polyline.Envelope.YMax <= pDatasetExtent.YMax)
96:         End If
97:     ElseIf bValidTarget And bCreatingNewFClass Then
98:         bPolylineWithinDataset = True
99:     End If
    Dim a As Long, b As Long, c As Long
101:     a = cmbFieldGridAngle.ListIndex
102:     b = cmbFieldMapScale.ListIndex
103:     c = cmbFieldSeriesNumber.ListIndex
104:     bDuplicateFieldsSelected = (a > 0 And (a = b Or a = c)) _
                            Or (b > 0 And (b = c))
106: i = 3
    
    ' Set states
    Select Case m_Step
        Case 0:     ' Set the target feature layer
111:             cmdBack.Enabled = False
112:             cmdNext.Enabled = bValidTarget And bValidName
113:             cmdNext.Caption = "Next >"
114:             cmbPolygonLayers.Enabled = Not bCreatingNewFClass
115:             chkRemovePreviousGrids.Enabled = Not bCreatingNewFClass
116:             lblClearExistingGridsPart2.Enabled = Not bCreatingNewFClass
117:             cmdSetNewGridLayer.Enabled = bCreatingNewFClass
        Case 1:     ' Set the fields to populate
119:             cmdBack.Enabled = True
120:             cmdNext.Enabled = (bValidRequiredFields And Not bDuplicateFieldsSelected)
121:             cmbFieldStripMapName.Enabled = Not bCreatingNewFClass
122:             cmbFieldGridAngle.Enabled = Not bCreatingNewFClass
123:             cmbFieldMapScale.Enabled = Not bCreatingNewFClass
124:             cmbFieldSeriesNumber.Enabled = Not bCreatingNewFClass
        Case 2:     ' Set the scale / starting_coords
126:             cmdBack.Enabled = True
127:             cmdNext.Enabled = bValidScale And bPolylineWithinDataset
128:             cmdNext.Caption = "Next >"
        Case 3:     ' Set the dataframe properties
130:             cmdBack.Enabled = True
131:             cmdNext.Enabled = bValidSize
132:             cmdNext.Caption = "Finish"
133:             txtManualGridHeight.Enabled = Not (optScaleSource(2).Value)
134:             txtManualGridHeight.Locked = (optScaleSource(2).Value)
135:             lblFrameHeight.Enabled = Not (optScaleSource(2).Value)
136:             optGridSize(0).Enabled = Not (optScaleSource(2).Value)
        Case Else:
138:             cmdBack.Enabled = False
139:             cmdNext.Enabled = False
140:     End Select
141: i = 4
    
143:     txtManualMapScale.Enabled = optScaleSource(1).Value
144:     txtManualGridWidth.Enabled = optGridSize(1).Value
145:     txtManualGridHeight.Enabled = optGridSize(1).Value
146:     cmbGridSizeUnits.Enabled = optGridSize(1).Value
147:     If optScaleSource(1).Value Then
148:         If bValidScale Then
149:             txtManualMapScale.ForeColor = (&H0)      ' Black
150:         Else
151:             txtManualMapScale.ForeColor = (&HFF)     ' Red
152:         End If
153:     End If
154:     If optGridSize(1).Value Then
155:         If bValidSize Then
156:             txtManualGridWidth.ForeColor = (&H0)      ' Black
157:             txtManualGridHeight.ForeColor = (&H0)
158:         Else
159:             txtManualGridWidth.ForeColor = (&HFF)     ' Red
160:             txtManualGridHeight.ForeColor = (&HFF)
161:         End If
162:     End If
    
    Exit Sub
165:     Resume
eh:
167:     MsgBox Err.Description, vbExclamation, "SetControlsState " & i
End Sub

Private Sub cmbFieldStripMapName_Click()
171:     SetControlsState
End Sub

Private Sub cmbFieldMapScale_Click()
175:     SetControlsState
End Sub

Private Sub cmbFieldSeriesNumber_Click()
179:     SetControlsState
End Sub

Private Sub cmbFieldGridAngle_Click()
183:     SetControlsState
End Sub

Private Sub cmbPolygonLayers_Click()
    Dim pFL As IFeatureLayer
    Dim pFields As IFields
    Dim lLoop As Long
    ' Populate the fields combo boxes
191:     If cmbPolygonLayers.ListIndex > 0 Then
192:         Set pFL = FindFeatureLayerByName(cmbPolygonLayers.List(cmbPolygonLayers.ListIndex), m_Application)
193:         Set pFields = pFL.FeatureClass.Fields
194:         cmbFieldMapScale.Clear
195:         cmbFieldStripMapName.Clear
196:         cmbFieldSeriesNumber.Clear
197:         cmbFieldGridAngle.Clear
198:         cmbFieldStripMapName.AddItem "<None>"
199:         cmbFieldGridAngle.AddItem "<None>"
200:         cmbFieldMapScale.AddItem "<None>"
201:         cmbFieldSeriesNumber.AddItem "<None>"
202:         For lLoop = 0 To pFields.FieldCount - 1
203:             If pFields.Field(lLoop).Type = esriFieldTypeString Then
204:                 cmbFieldStripMapName.AddItem pFields.Field(lLoop).Name
205:             ElseIf pFields.Field(lLoop).Type = esriFieldTypeDouble Or _
                   pFields.Field(lLoop).Type = esriFieldTypeInteger Or _
                   pFields.Field(lLoop).Type = esriFieldTypeSmallInteger Or _
                   pFields.Field(lLoop).Type = esriFieldTypeSingle Then
209:                 cmbFieldMapScale.AddItem pFields.Field(lLoop).Name
210:                 cmbFieldGridAngle.AddItem pFields.Field(lLoop).Name
211:                 cmbFieldSeriesNumber.AddItem pFields.Field(lLoop).Name
212:             End If
213:         Next
214:         cmbFieldStripMapName.ListIndex = 0
215:         cmbFieldGridAngle.ListIndex = 0
216:         cmbFieldMapScale.ListIndex = 0
217:         cmbFieldSeriesNumber.ListIndex = 0
218:     End If
219:     SetControlsState
End Sub

Private Sub cmdBack_Click()
223:     m_Step = m_Step - 1
224:     If m_Step < 0 Then
225:         m_Step = 0
226:     End If
227:     SetVisibleControls m_Step
228:     SetControlsState
End Sub

Private Sub cmdClose_Click()
232:     Set m_Application = Nothing
233:     Set Me.StripMapSettings = Nothing
234:     Me.Hide
End Sub

Private Sub CollateStripMapSettings()
    Dim pMx As IMxDocument
    Dim pCreateSMap As New clsCreateStripMap
    Dim pFrameElement As IElement
    Dim sDestLayerName As String
    Dim lLoop As Long
    ' Populate class
244:     pCreateSMap.StripMapName = txtStripMapSeriesName.Text
245:     pCreateSMap.FlipPolyline = (chkFlipLine.Value = vbChecked)
246:     If (optScaleSource(0).Value) Then
247:         pCreateSMap.MapScale = CDbl(lblCurrentMapScale.Caption)
248:     ElseIf (optScaleSource(1).Value) Then
249:         pCreateSMap.MapScale = CDbl(txtManualMapScale.Text)
250:     End If
251:     If (optGridSize(0).Value) Then
252:         Set pFrameElement = GetDataFrameElement(GetActiveDataFrameName(m_Application), m_Application)
253:         pCreateSMap.FrameWidthInPageUnits = pFrameElement.Geometry.Envelope.Width
254:         pCreateSMap.FrameHeightInPageUnits = pFrameElement.Geometry.Envelope.Height
255:     Else
256:         pCreateSMap.FrameWidthInPageUnits = CDbl(txtManualGridWidth.Text)
257:         pCreateSMap.FrameHeightInPageUnits = CDbl(txtManualGridHeight.Text)
258:     End If
259:     If (optScaleSource(2).Value) Then
        Dim dConvertPageToMapUnits As Double, dGridToFrameRatio As Double
261:         dConvertPageToMapUnits = CalculatePageToMapRatio(m_Application) 'NATHAN FIX THIS
262:         pCreateSMap.FrameWidthInPageUnits = CDbl(txtManualGridWidth.Text)
263:         pCreateSMap.FrameHeightInPageUnits = CDbl(txtManualGridHeight.Text)
264:         If pCreateSMap.FrameWidthInPageUnits >= pCreateSMap.FrameHeightInPageUnits Then
265:             dGridToFrameRatio = CDbl(txtAbsoluteGridWidth.Text) / pCreateSMap.FrameWidthInPageUnits
266:         Else
267:             dGridToFrameRatio = CDbl(txtAbsoluteGridHeight.Text) / pCreateSMap.FrameHeightInPageUnits
268:         End If
269:         pCreateSMap.MapScale = dGridToFrameRatio * dConvertPageToMapUnits
270:     End If
271:     sDestLayerName = cmbPolygonLayers.List(cmbPolygonLayers.ListIndex)
272:     If optLayerSource(0).Value Then
273:         Set pCreateSMap.DestinationFeatureLayer = FindFeatureLayerByName(sDestLayerName, m_Application)
274:     End If
275:     pCreateSMap.FieldNameStripMapName = cmbFieldStripMapName.List(cmbFieldStripMapName.ListIndex)
276:     pCreateSMap.FieldNameMapAngle = cmbFieldGridAngle.List(cmbFieldGridAngle.ListIndex)
277:     pCreateSMap.FieldNameNumberInSeries = cmbFieldSeriesNumber.List(cmbFieldSeriesNumber.ListIndex)
278:     If cmbFieldMapScale.ListIndex > 0 Then pCreateSMap.FieldNameScale = cmbFieldMapScale.List(cmbFieldMapScale.ListIndex)
279:     pCreateSMap.RemoveCurrentGrids = (chkRemovePreviousGrids.Value = vbChecked)
280:     Set pCreateSMap.StripMapRoute = m_Polyline
    ' Place grid settings on Public form property (so calling function can use them)
282:     Set Me.StripMapSettings = pCreateSMap
End Sub

Private Sub cmdNext_Click()
    Dim pMx As IMxDocument
    Dim pFeatureLayer As IFeatureLayer
    Dim pOutputFClass As IFeatureClass
    Dim pNewFields As IFields
    
    On Error GoTo eh
    ' Step
293:     m_Step = m_Step + 1
    ' If we're creating a new fclass, we can skip a the 'Set Field Roles' step
295:     If m_Step = 1 And (optLayerSource(1).Value) Then
296:         m_Step = m_Step + 1
297:     End If
    ' If FINISH
299:     If m_Step >= 4 Then
300:         Set pMx = m_Application.Document
301:         RemoveGraphicsByName pMx
302:         CollateStripMapSettings
        ' If creating a new layer
304:         If optLayerSource(1).Value Then
            ' Create the feature class
306:             Set pNewFields = CreateTheFields
            Select Case m_FileType
                Case ShapeFile
309:                     Set pOutputFClass = NewShapeFile(m_OutputLayer, pMx.FocusMap, pNewFields)
                Case AccessFeatureClass
311:                     Set pOutputFClass = NewAccessFile(m_OutputLayer, _
                            m_OutputDataset, m_OutputFClass, pNewFields)
313:             End Select
314:             If pOutputFClass Is Nothing Then
315:                 Err.Raise vbObjectError, "cmdNext", "Could not create the new output feature class."
316:             End If
            ' Create new layer
318:             Set pFeatureLayer = New FeatureLayer
319:             Set pFeatureLayer.FeatureClass = pOutputFClass
320:             pFeatureLayer.Name = pFeatureLayer.FeatureClass.AliasName
            ' Add the new layer to arcmap & reset the StripMapSettings object to point at it
322:             pMx.FocusMap.AddLayer pFeatureLayer
323:             Set StripMapSettings.DestinationFeatureLayer = pFeatureLayer
324:         End If
325:         Me.Hide
326:     Else
327:         SetVisibleControls m_Step
328:         SetControlsState
329:     End If
    
    Exit Sub
eh:
333:     MsgBox "Error: " & Err.Description, , "cmdNext_Click"
334:     m_Step = m_Step - 1
End Sub

Private Sub cmdSetNewGridLayer_Click()
  Dim pGxFilter As IGxObjectFilter
  Dim pGXBrow As IGxDialog, bFlag As Boolean
  Dim pSel As IEnumGxObject, pApp As IApplication
  
342:   Set pGxFilter = New GxFilter
343:   Set pApp = m_Application
344:   Set pGXBrow = New GxDialog
345:   Set pGXBrow.ObjectFilter = pGxFilter
346:   pGXBrow.Title = "Output feature class or shapefile"
347:   bFlag = pGXBrow.DoModalSave(pApp.hwnd)
  
349:   If bFlag Then
    Dim pObj As IGxObject
351:     Set pObj = pGXBrow.FinalLocation
352:     m_bIsGeoDatabase = True
353:     If UCase(pObj.Category) = "FOLDER" Then
354:       If InStr(1, pGXBrow.Name, ".shp") > 0 Then
355:         txtNewGridLayer.Text = pObj.FullName & "\" & pGXBrow.Name
356:       Else
357:         txtNewGridLayer.Text = pObj.FullName & "\" & pGXBrow.Name & ".shp"
358:       End If
359:       m_OutputLayer = txtNewGridLayer.Text
360:       m_bIsGeoDatabase = False
361:       m_FileType = ShapeFile
362:      CheckOutputFile
363:     Else
      Dim pLen As Long
365:       pLen = Len(pObj.FullName) - Len(pObj.BaseName) - 1
366:       txtNewGridLayer.Text = Left(pObj.FullName, pLen)
367:       m_OutputLayer = Left(pObj.FullName, pLen)
368:       m_OutputDataset = pObj.BaseName
369:       m_OutputFClass = pGXBrow.Name
370:       m_bIsGeoDatabase = True
371:       If UCase(pObj.Category) = "PERSONAL GEODATABASE FEATURE DATASET" Then
372:         m_FileType = AccessFeatureClass
373:       Else
374:         m_FileType = SDEFeatureClass
375:       End If
376:     End If
377:   Else
378:     txtNewGridLayer.Text = ""
379:     m_bIsGeoDatabase = False
380:   End If
381:   SetControlsState
End Sub

Private Sub Form_Load()
    Dim pMx As IMxDocument
    Dim bRenewCoordsX As Boolean
    Dim bRenewCoordsY As Boolean
    Dim sErrMsg As String
    On Error GoTo eh
    
391:     sErrMsg = CreateStripMapPolyline
392:     If Len(sErrMsg) > 0 Then
393:         MsgBox sErrMsg, vbCritical, "Create Map Grids"
394:         Unload Me
        Exit Sub
396:     End If
397:     Set pMx = m_Application.Document
398:     Me.Height = 5665
399:     Me.Width = 4935
400:     m_Step = 0
401:     LoadLayersComboBox
402:     LoadUnitsComboBox
403:     lblCurrFrameName.Caption = GetActiveDataFrameName(m_Application)
404:     If pMx.FocusMap.MapUnits = esriUnknownUnits Then
405:         MsgBox "Error: The map has unknown units and therefore cannot calculate a Scale." _
            & vbCrLf & "Cannot create Map Grids at this time.", vbCritical, "Create Map Grids"
407:         Unload Me
        Exit Sub
409:     End If
410:     lblMapUnits.Caption = GetUnitsDescription(pMx.FocusMap.MapUnits)
411:     lblCurrentMapScale.Caption = Format(pMx.FocusMap.MapScale, "#,###,##0")
412:     SetVisibleControls m_Step
    
414:     SetControlsState
    
    'Make sure the wizard stays on top
417:     TopMost Me
    
    Exit Sub
eh:
421:     MsgBox "Error loading the form: " & Err.Description & vbCrLf _
        & vbCrLf & "Attempting to continue the load...", , "MapGridManager: Form_Load "
    On Error Resume Next
424:     SetVisibleControls m_Step
425:     SetControlsState
End Sub

Private Sub LoadUnitsComboBox()
    Dim pMx As IMxDocument
    Dim sPageUnitsDesc As String
    Dim pPage As IPage
    
    On Error GoTo eh
    
    ' Init
436:     Set pMx = m_Application.Document
437:     Set pPage = pMx.PageLayout.Page
438:     sPageUnitsDesc = GetUnitsDescription(pPage.Units)
439:     cmbGridSizeUnits.Clear
    ' Add
441:     cmbGridSizeUnits.AddItem sPageUnitsDesc
    'cmbGridSizeUnits.AddItem "Map Units (" & sMapUnitsDesc & ")"
    ' Set page units as default
444:     cmbGridSizeUnits.ListIndex = 0
    
    Exit Sub
eh:
448:     Err.Raise vbObjectError, "LoadUnitsComboBox", "Error in LoadUnitsComboBox" & vbCrLf & Err.Description
End Sub

Private Sub LoadLayersComboBox()
    Dim pMx As IMxDocument
    Dim lLoop As Long
    Dim pFL As IFeatureLayer
    Dim pFC As IFeatureClass
    Dim sPreviousLayer  As String
    Dim lResetIndex As Long
    
    'Init
460:     Set pMx = m_Application.Document
461:     cmbPolygonLayers.Clear
462:     cmbPolygonLayers.AddItem "<Not Set>"
    ' For all layers
464:     For lLoop = 0 To pMx.FocusMap.LayerCount - 1
        ' If a feature class
466:         If TypeOf pMx.FocusMap.Layer(lLoop) Is IFeatureLayer Then
467:             Set pFL = pMx.FocusMap.Layer(lLoop)
468:             Set pFC = pFL.FeatureClass
            ' If a polygon layer
470:             If pFC.ShapeType = esriGeometryPolygon Then
                ' Add to combo box
472:                 cmbPolygonLayers.AddItem pFL.Name
473:             End If
474:         End If
475:     Next
476:     cmbPolygonLayers.ListIndex = 0
End Sub

Private Sub SetCurrentMapScaleCaption()
    Dim pMx As IMxDocument
    On Error GoTo eh
482:     Set pMx = m_Application.Document
483:     lblCurrentMapScale.Caption = Format(pMx.FocusMap.MapScale, "#,###,##0")
    Exit Sub
eh:
486:     lblCurrentMapScale.Caption = "<Scale Unknown>"
End Sub


Private Sub Form_Unload(Cancel As Integer)
491:     Set m_Application = Nothing
492:     Set StripMapSettings = Nothing
End Sub


Private Sub optGridSize_Click(Index As Integer)
    Dim pMx As IMxDocument
498:     Set pMx = m_Application.Document
499:     lblCurrFrameName.Caption = pMx.FocusMap.Name
500:     SetControlsState
End Sub

Private Sub optLayerSource_Click(Index As Integer)
    ' If creating a new fclass to hold the grids
505:     If Index = 1 Then
        ' Set the field names (will be created automatically)
507:         cmbFieldStripMapName.Clear
508:         cmbFieldGridAngle.Clear
509:         cmbFieldSeriesNumber.Clear
510:         cmbFieldMapScale.Clear
511:         cmbFieldStripMapName.AddItem "<None>"
512:         cmbFieldGridAngle.AddItem "<None>"
513:         cmbFieldSeriesNumber.AddItem "<None>"
514:         cmbFieldMapScale.AddItem "<None>"
515:         cmbFieldStripMapName.AddItem c_DefaultFld_StripMapName
516:         cmbFieldGridAngle.AddItem c_DefaultFld_MapAngle
517:         cmbFieldSeriesNumber.AddItem c_DefaultFld_SeriesNum
518:         cmbFieldMapScale.AddItem c_DefaultFld_MapScale
519:         cmbFieldStripMapName.ListIndex = 1
520:         cmbFieldGridAngle.ListIndex = 1
521:         cmbFieldSeriesNumber.ListIndex = 1
522:         cmbFieldMapScale.ListIndex = 1
523:     End If
524:     SetControlsState
End Sub

Private Sub optScaleSource_Click(Index As Integer)
528:     If Index = 0 Then
529:         SetCurrentMapScaleCaption
530:     ElseIf Index = 2 Then
531:         optGridSize(1).Value = True
532:     End If
533:     SetControlsState
End Sub

Private Sub txtAbsoluteGridHeight_Change()
537:     SetControlsState
End Sub

Private Sub txtAbsoluteGridHeight_KeyPress(KeyAscii As Integer)
    ' If a non-numeric (that is not a decimal point)
542:     If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
     And KeyAscii <> Asc(".") _
     And Chr(KeyAscii) <> vbBack Then
        ' Do not allow this button to work
546:         KeyAscii = 0
    ' If a decimal point, make sure we only ever get one of them
548:     ElseIf KeyAscii = Asc(".") Then
549:         If InStr(txtAbsoluteGridHeight.Text, ".") > 0 Then
550:             KeyAscii = 0
551:         End If
552:     End If
End Sub

Private Sub txtAbsoluteGridWidth_Change()
556:     SetControlsState
End Sub

Private Sub txtAbsoluteGridWidth_KeyPress(KeyAscii As Integer)
    ' If a non-numeric (that is not a decimal point)
561:     If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
     And KeyAscii <> Asc(".") _
     And Chr(KeyAscii) <> vbBack Then
        ' Do not allow this button to work
565:         KeyAscii = 0
    ' If a decimal point, make sure we only ever get one of them
567:     ElseIf KeyAscii = Asc(".") Then
568:         If InStr(txtAbsoluteGridWidth.Text, ".") > 0 Then
569:             KeyAscii = 0
570:         End If
571:     End If
End Sub

Private Sub txtManualGridHeight_Change()
575:     SetControlsState
End Sub

Private Sub txtManualGridHeight_KeyPress(KeyAscii As Integer)
    ' If a non-numeric (that is not a decimal point)
580:     If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
     And KeyAscii <> Asc(".") _
     And Chr(KeyAscii) <> vbBack Then
        ' Do not allow this button to work
584:         KeyAscii = 0
    ' If a decimal point, make sure we only ever get one of them
586:     ElseIf KeyAscii = Asc(".") Then
587:         If InStr(txtManualGridHeight.Text, ".") > 0 Then
588:             KeyAscii = 0
589:         End If
590:     End If
End Sub

Private Sub txtManualGridWidth_Change()
594:     If IsNumeric(txtManualGridWidth.Text) And optScaleSource(2).Value Then
        Dim dRatio As Double, dGridWidth As Double
596:         dGridWidth = CDbl(txtManualGridWidth.Text)
597:         dRatio = CDbl(txtAbsoluteGridHeight.Text) / CDbl(txtAbsoluteGridWidth.Text)
598:         txtManualGridHeight.Text = CStr(dRatio * dGridWidth)
599:     End If
600:     SetControlsState
End Sub

Private Sub txtManualGridWidth_KeyPress(KeyAscii As Integer)
    ' If a non-numeric (that is not a decimal point)
605:     If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
     And KeyAscii <> Asc(".") _
     And Chr(KeyAscii) <> vbBack Then
        ' Do not allow this button to work
609:         KeyAscii = 0
    ' If a decimal point, make sure we only ever get one of them
611:     ElseIf KeyAscii = Asc(".") Then
612:         If InStr(txtManualGridWidth.Text, ".") > 0 Then
613:             KeyAscii = 0
614:         End If
615:     End If
End Sub

Private Sub txtManualMapScale_Change()
619:     SetControlsState
End Sub

Private Sub txtManualMapScale_KeyPress(KeyAscii As Integer)
    ' If a non-numeric (that is not a decimal point)
624:     If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
     And KeyAscii <> Asc(".") _
     And Chr(KeyAscii) <> vbBack Then
        ' Do not allow this button to work
628:         KeyAscii = 0
    ' If a decimal point, make sure we only ever get one of them
630:     ElseIf KeyAscii = Asc(".") Then
631:         If InStr(txtManualMapScale.Text, ".") > 0 Then
632:             KeyAscii = 0
633:         End If
634:     End If
End Sub

Public Sub Tickle()
638:     Call Form_Load
End Sub

Private Sub SetVisibleControls(iStep As Integer)
    ' Hide all
643:     fraAttributes.Visible = False
644:     fraDataFrameSize.Visible = False
645:     fraDestinationFeatureClass.Visible = False
646:     fraScaleStart.Visible = False
    ' Show applicable frame, set top/left
    Select Case iStep
        Case 0:
650:             fraDestinationFeatureClass.Visible = True
651:             fraDestinationFeatureClass.Top = 0
652:             fraDestinationFeatureClass.Left = 0
        Case 1:
654:             fraAttributes.Visible = True
655:             fraAttributes.Top = 0
656:             fraAttributes.Left = 0
        Case 2:
658:             fraScaleStart.Visible = True
659:             fraScaleStart.Top = 0
660:             fraScaleStart.Left = 0
        Case 3:
662:             fraDataFrameSize.Visible = True
663:             fraDataFrameSize.Top = 0
664:             fraDataFrameSize.Left = 0
        Case Else:
666:             MsgBox "Invalid Step Value : " & iStep
667:     End Select
End Sub

Private Sub CheckOutputFile()
    'Check the output option
672:     If txtNewGridLayer.Text <> "" Then
673:         If DoesShapeFileExist(txtNewGridLayer.Text) Then
674:             MsgBox "Shape file name already being used!!!"
675:             txtNewGridLayer.Text = ""
676:         End If
677:     End If
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
690:     Set pNewFields = New Fields
691:     Set pFieldsEdit = pNewFields
692:     Set pMx = m_Application.Document
    ' Field: OID  -------------------------
694:     Set newField = New Field
695:     Set newFieldEdit = newField
696:     With newFieldEdit
697:         .Name = "OID"
698:         .Type = esriFieldTypeOID
699:         .AliasName = "Object ID"
700:         .IsNullable = False
701:     End With
702:     pFieldsEdit.AddField newField
    ' Field: STRIP MAP NAME -------------------------
704:     Set newField = New Field
705:     Set newFieldEdit = newField
706:     With newFieldEdit
707:       .Name = c_DefaultFld_StripMapName
708:       .AliasName = "StripMapName"
709:       .Type = esriFieldTypeString
710:       .IsNullable = True
711:       .Length = 50
712:     End With
713:     pFieldsEdit.AddField newField
    ' Field: MAP ANGLE -------------------------
715:     Set newField = New Field
716:     Set newFieldEdit = newField
717:     With newFieldEdit
718:       .Name = c_DefaultFld_MapAngle
719:       .AliasName = "Map Angle"
720:       .Type = esriFieldTypeInteger
721:       .IsNullable = True
722:     End With
723:     pFieldsEdit.AddField newField
    ' Field: GRID NUMBER -------------------------
725:     Set newField = New Field
726:     Set newFieldEdit = newField
727:     With newFieldEdit
728:       .Name = c_DefaultFld_SeriesNum
729:       .AliasName = "Number In Series"
730:       .Type = esriFieldTypeInteger
731:       .IsNullable = True
732:     End With
733:     pFieldsEdit.AddField newField
    ' Field: SCALE -------------------------
735:     Set newField = New Field
736:     Set newFieldEdit = newField
737:     With newFieldEdit
738:       .Name = c_DefaultFld_MapScale
739:       .AliasName = "Plot Scale"
740:       .Type = esriFieldTypeDouble
741:       .IsNullable = True
742:       .Precision = 18
743:       .Scale = 11
744:     End With
745:     pFieldsEdit.AddField newField
    ' Return
747:     Set CreateTheFields = pFieldsEdit
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
762:     Set pMx = pApp.Document
763:     Set pSR = pMx.FocusMap.SpatialReference
764:     If TypeOf pSR Is IProjectedCoordinateSystem Then
765:         Set pPCS = pSR
766:         dMetersPerUnit = pPCS.CoordinateUnit.MetersPerUnit
767:     Else
768:         dMetersPerUnit = 1
769:     End If
770:     Set pPage = pMx.PageLayout.Page
771:     pPageUnits = pPage.Units
    Select Case pPageUnits
        Case esriInches: CalculatePageToMapRatio = dMetersPerUnit / (1 / 12 * 0.304800609601219)
        Case esriFeet: CalculatePageToMapRatio = dMetersPerUnit / (0.304800609601219)
        Case esriCentimeters: CalculatePageToMapRatio = dMetersPerUnit / (1 / 100)
        Case esriMeters: CalculatePageToMapRatio = dMetersPerUnit / (1)
        Case Else:
778:             MsgBox "Warning: Only the following Page (Layout) Units are supported by this tool:" _
                & vbCrLf & " - Inches, Feet, Centimeters, Meters" _
                & vbCrLf & vbCrLf & "Calculating as though Page Units are in Inches..."
781:             CalculatePageToMapRatio = dMetersPerUnit / (1 / 12 * 0.304800609601219)
782:     End Select
    Exit Function
eh:
785:     CalculatePageToMapRatio = 1
786:     MsgBox "Error in CalculatePageToMapRatio" & vbCrLf & Err.Description
End Function

Private Function ReturnMax(dDouble1 As Double, dDouble2 As Double) As Double
790:     If dDouble1 >= dDouble2 Then
791:         ReturnMax = dDouble1
792:     Else
793:         ReturnMax = dDouble2
794:     End If
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
811:     Set pMx = m_Application.Document
812:     Set pFC = pMx.FocusMap.FeatureSelection
813:     Set pF = pFC.Next
814:     If pF Is Nothing Then
815:         CreateStripMapPolyline = "Requires selected polyline features/s."
        Exit Function
817:     End If
    ' Make polyline
819:     Set pPolyline = New Polyline
820:     While Not pF Is Nothing
821:         If pF.Shape.GeometryType = esriGeometryPolyline Then
822:             Set pTmpPolyline = pF.ShapeCopy
823:             Set pTopoSimplify = pTmpPolyline
824:             pTopoSimplify.Simplify
825:             Set pTopoUnion = pPolyline
826:             Set pPolyline = pTopoUnion.Union(pTopoSimplify)
827:             Set pTopoSimplify = pPolyline
828:             pTopoSimplify.Simplify
829:         End If
830:         Set pF = pFC.Next
831:     Wend
    ' Check polyline for beinga single, connected polyline (Path)
833:     Set pGeoColl = pPolyline
834:     If pGeoColl.GeometryCount = 0 Then
835:         CreateStripMapPolyline = "Requires selected polyline features/s."
        Exit Function
837:     ElseIf pGeoColl.GeometryCount > 1 Then
838:         CreateStripMapPolyline = "Cannot process the StripMap - multi-part polyline created." _
            & vbCrLf & "Check for non-connected segments, overlaps or loops."
        Exit Function
841:     End If
    ' Give option to flip
843:     Perm_DrawPoint pPolyline.FromPoint, , 0, 255, 0, 20
844:     Perm_DrawTextFromPoint pPolyline.FromPoint, "START", , , , , 20
845:     Perm_DrawPoint pPolyline.ToPoint, , 255, 0, 0, 20
846:     Perm_DrawTextFromPoint pPolyline.ToPoint, "END", , , , , 20
847:     pMx.ActiveView.PartialRefresh esriViewGraphics, Nothing, Nothing
    
849:     Set m_Polyline = pPolyline
    
851:     CreateStripMapPolyline = ""
    
    Exit Function
854:     Resume
eh:
856:     CreateStripMapPolyline = "Error in CreateStripMapPolyline : " & Err.Description
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
874:     Set pMx = m_Application.Document
875:     Set pGLayer = pMx.FocusMap.BasicGraphicsLayer
876:     Set pGCon = pGLayer
877:     Set pElement = New MarkerElement
878:     pElement.Geometry = pPoint
879:     Set pMarkerElement = pElement
    
    ' Set the symbol
882:     Set pColor = New RgbColor
883:     pColor.Red = dRed
884:     pColor.Green = dGreen
885:     pColor.Blue = dBlue
886:     Set pMarker = New SimpleMarkerSymbol
887:     With pMarker
888:         .Color = pColor
889:         .Size = dSize
890:     End With
891:     pMarkerElement.Symbol = pMarker
    
    ' Add the graphic
894:     Set pElementProp = pElement
895:     pElementProp.Name = sElementName
896:     pGCon.AddElement pElement, 0
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
919:     Set pMx = m_Application.Document
920:     Set pGLayer = pMx.FocusMap.BasicGraphicsLayer
921:     Set pGCon = pGLayer
922:     Set pElement = New LineElement
    
    ' Set the line symbol
925:     Set pLnSym = New SimpleLineSymbol
926:     Set myColor = New RgbColor
927:     myColor.Red = dRed
928:     myColor.Green = dGreen
929:     myColor.Blue = dBlue
930:     pLnSym.Color = myColor
931:     pLnSym.Width = dSize
    
    ' Create a standard polyline (via 2 points)
934:     Set pLine1 = New esrigeometry.Line
935:     pLine1.PutCoords pFromPoint, pToPoint
936:     Set pSeg1 = pLine1
937:     Set pPolyline = New Polyline
938:     pPolyline.AddSegment pSeg1
939:     pElement.Geometry = pPolyline
940:     Set pLineElement = pElement
941:     pLineElement.Symbol = pLnSym
    
    ' Add the graphic
944:     Set pElementProp = pElement
945:     pElementProp.Name = sElementName
946:     pGCon.AddElement pElement, 0
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
964:     Set pMx = m_Application.Document
965:     Set pGLayer = pMx.FocusMap.BasicGraphicsLayer
966:     Set pGCon = pGLayer
967:     Set pElement = New TextElement
968:     pElement.Geometry = pPoint
969:     Set pTextElement = pElement
    
    ' Create the text symbol
972:     Set myTxtSym = New TextSymbol
973:     Set myColor = New RgbColor
974:     myColor.Red = dRed
975:     myColor.Green = dGreen
976:     myColor.Blue = dBlue
977:     myTxtSym.Color = myColor
978:     myTxtSym.Size = dSize
979:     myTxtSym.HorizontalAlignment = esriTHACenter
980:     pTextElement.Symbol = myTxtSym
981:     pTextElement.Text = sText
    
    ' Add the graphic
984:     Set pElementProp = pElement
985:     pElementProp.Name = sElementName
986:     pGCon.AddElement pElement, 0
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
1002:     pMxDoc.DelayUpdateContents = True
1003:     Set pGLayer = pMxDoc.FocusMap.BasicGraphicsLayer
1004:     Set pGCon = pGLayer
1005:     pGCon.Next
    
    ' Delete all the graphic elements that we created (identify by the name prefix)
1008:     pGCon.Reset
1009:     Set pElement = pGCon.Next
1010:     While Not pElement Is Nothing
1011:         If TypeOf pElement Is IElement Then
1012:             Set pElementProp = pElement
1013:             If (Left(pElementProp.Name, Len(sPrefix)) = sPrefix) Then
1014:                 pGCon.DeleteElement pElement
1015:             End If
1016:         End If
1017:         Set pElement = pGCon.Next
1018:     Wend
    
    ' Switch ON the updating of the TOC, refresh
1021:     pMxDoc.DelayUpdateContents = False
1022:     pMxDoc.ActiveView.Refresh
    
    Exit Sub
ErrorHandler:
1026:     MsgBox "Error in RemoveGraphicsByName: " & Err.Description, , "RemoveGraphicsByName"
End Sub

Private Sub txtStripMapSeriesName_Change()
1030:     SetControlsState
End Sub
