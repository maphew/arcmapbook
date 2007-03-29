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
57:     If Len(lblCurrentMapScale.Caption) = 0 Then lblCurrentMapScale.Caption = "0"
58:     If Len(txtManualMapScale.Text) = 0 Then
59:         dScale = 0
60:     Else
61:         dScale = CDbl(txtManualMapScale.Text)
62:     End If
63:     If Len(txtManualGridHeight.Text) = 0 Then
64:         dGHeight = 0
65:     Else
66:         dGHeight = CDbl(txtManualGridHeight.Text)
67:     End If
68:     If Len(txtManualGridWidth.Text) = 0 Then
69:         dGWidth = 0
70:     Else
71:         dGWidth = CDbl(txtManualGridWidth.Text)
72:     End If
73:     If Len(txtAbsoluteGridHeight.Text) = 0 Then
74:         dAHeight = 0
75:     Else
76:         dAHeight = CDbl(txtAbsoluteGridHeight.Text)
77:     End If
78:     If Len(txtAbsoluteGridWidth.Text) = 0 Then
79:         dAWidth = 0
80:     Else
81:         dAWidth = CDbl(txtAbsoluteGridWidth.Text)
82:     End If
83: i = 1

    ' Calc values
86:     bValidName = Len(txtStripMapSeriesName.Text) > 0
87:     bValidScale = (optScaleSource(0).value And CDbl(lblCurrentMapScale.Caption) > 0) Or _
                  (optScaleSource(1).value And dScale > 0) Or _
                  (optScaleSource(2).value And dAHeight > 0 And dAWidth > 0)
90:     bValidSize = (optGridSize(0).value) Or _
                 (optGridSize(1).value And dGHeight > 0 And dGWidth > 0) Or _
                 (optScaleSource(2).value And CDbl(txtManualGridWidth.Text) > 0)
93:     bCreatingNewFClass = optLayerSource(1).value
94:     bNewFClassSet = (Len(txtNewGridLayer.Text) > 0)
95:     bValidTarget = (cmbPolygonLayers.ListIndex > 0) Or (bCreatingNewFClass And bNewFClassSet)
96:     bValidRequiredFields = (cmbFieldStripMapName.ListIndex > 0) And _
                           (cmbFieldGridAngle.ListIndex > 0) And _
                           (cmbFieldSeriesNumber.ListIndex > 0)
99: i = 2
100:     If bValidTarget And (Not bCreatingNewFClass) Then
101:         Set pFL = FindFeatureLayerByName(cmbPolygonLayers.List(cmbPolygonLayers.ListIndex), m_Application)
102:         If pFL.FeatureClass.FeatureDataset Is Nothing Then
103:             bPolylineWithinDataset = True
104:         Else
105:             Set pDatasetExtent = GetValidExtentForLayer(pFL)
106:             bPolylineWithinDataset = (m_Polyline.Envelope.XMin >= pDatasetExtent.XMin And m_Polyline.Envelope.XMax <= pDatasetExtent.XMax) _
                     And (m_Polyline.Envelope.YMin >= pDatasetExtent.YMin And m_Polyline.Envelope.YMax <= pDatasetExtent.YMax)
108:         End If
109:     ElseIf bValidTarget And bCreatingNewFClass Then
110:         bPolylineWithinDataset = True
111:     End If
    Dim a As Long, b As Long, c As Long
113:     a = cmbFieldGridAngle.ListIndex
114:     b = cmbFieldMapScale.ListIndex
115:     c = cmbFieldSeriesNumber.ListIndex
116:     bDuplicateFieldsSelected = (a > 0 And (a = b Or a = c)) _
                            Or (b > 0 And (b = c))
118: i = 3
    
    ' Set states
    Select Case m_Step
        Case 0:     ' Set the target feature layer
123:             cmdBack.Enabled = False
124:             cmdNext.Enabled = bValidTarget And bValidName
125:             cmdNext.Caption = "Next >"
126:             cmbPolygonLayers.Enabled = Not bCreatingNewFClass
127:             chkRemovePreviousGrids.Enabled = Not bCreatingNewFClass
128:             lblClearExistingGridsPart2.Enabled = Not bCreatingNewFClass
129:             cmdSetNewGridLayer.Enabled = bCreatingNewFClass
        Case 1:     ' Set the fields to populate
131:             cmdBack.Enabled = True
132:             cmdNext.Enabled = (bValidRequiredFields And Not bDuplicateFieldsSelected)
133:             cmbFieldStripMapName.Enabled = Not bCreatingNewFClass
134:             cmbFieldGridAngle.Enabled = Not bCreatingNewFClass
135:             cmbFieldMapScale.Enabled = Not bCreatingNewFClass
136:             cmbFieldSeriesNumber.Enabled = Not bCreatingNewFClass
        Case 2:     ' Set the scale / starting_coords
138:             cmdBack.Enabled = True
139:             cmdNext.Enabled = bValidScale And bPolylineWithinDataset
140:             cmdNext.Caption = "Next >"
        Case 3:     ' Set the dataframe properties
142:             cmdBack.Enabled = True
143:             cmdNext.Enabled = bValidSize
144:             cmdNext.Caption = "Finish"
145:             txtManualGridHeight.Enabled = Not (optScaleSource(2).value)
146:             txtManualGridHeight.Locked = (optScaleSource(2).value)
147:             lblFrameHeight.Enabled = Not (optScaleSource(2).value)
148:             optGridSize(0).Enabled = Not (optScaleSource(2).value)
        Case Else:
150:             cmdBack.Enabled = False
151:             cmdNext.Enabled = False
152:     End Select
153: i = 4
    
155:     txtManualMapScale.Enabled = optScaleSource(1).value
156:     txtManualGridWidth.Enabled = optGridSize(1).value
157:     txtManualGridHeight.Enabled = optGridSize(1).value
158:     cmbGridSizeUnits.Enabled = optGridSize(1).value
159:     If optScaleSource(1).value Then
160:         If bValidScale Then
161:             txtManualMapScale.ForeColor = (&H0)      ' Black
162:         Else
163:             txtManualMapScale.ForeColor = (&HFF)     ' Red
164:         End If
165:     End If
166:     If optGridSize(1).value Then
167:         If bValidSize Then
168:             txtManualGridWidth.ForeColor = (&H0)      ' Black
169:             txtManualGridHeight.ForeColor = (&H0)
170:         Else
171:             txtManualGridWidth.ForeColor = (&HFF)     ' Red
172:             txtManualGridHeight.ForeColor = (&HFF)
173:         End If
174:     End If
    
    Exit Sub
177:     Resume
eh:
179:     MsgBox Err.Description, vbExclamation, "SetControlsState " & i
End Sub

Private Sub cmbFieldStripMapName_Click()
183:     SetControlsState
End Sub

Private Sub cmbFieldMapScale_Click()
187:     SetControlsState
End Sub

Private Sub cmbFieldSeriesNumber_Click()
191:     SetControlsState
End Sub

Private Sub cmbFieldGridAngle_Click()
195:     SetControlsState
End Sub

Private Sub cmbPolygonLayers_Click()
    Dim pFL As IFeatureLayer
    Dim pFields As IFields
    Dim lLoop As Long
    ' Populate the fields combo boxes
203:     If cmbPolygonLayers.ListIndex > 0 Then
204:         Set pFL = FindFeatureLayerByName(cmbPolygonLayers.List(cmbPolygonLayers.ListIndex), m_Application)
205:         Set pFields = pFL.FeatureClass.Fields
206:         cmbFieldMapScale.Clear
207:         cmbFieldStripMapName.Clear
208:         cmbFieldSeriesNumber.Clear
209:         cmbFieldGridAngle.Clear
210:         cmbFieldStripMapName.AddItem "<None>"
211:         cmbFieldGridAngle.AddItem "<None>"
212:         cmbFieldMapScale.AddItem "<None>"
213:         cmbFieldSeriesNumber.AddItem "<None>"
214:         For lLoop = 0 To pFields.FieldCount - 1
215:             If pFields.Field(lLoop).Type = esriFieldTypeString Then
216:                 cmbFieldStripMapName.AddItem pFields.Field(lLoop).Name
217:             ElseIf pFields.Field(lLoop).Type = esriFieldTypeDouble Or _
                   pFields.Field(lLoop).Type = esriFieldTypeInteger Or _
                   pFields.Field(lLoop).Type = esriFieldTypeSmallInteger Or _
                   pFields.Field(lLoop).Type = esriFieldTypeSingle Then
221:                 cmbFieldMapScale.AddItem pFields.Field(lLoop).Name
222:                 cmbFieldGridAngle.AddItem pFields.Field(lLoop).Name
223:                 cmbFieldSeriesNumber.AddItem pFields.Field(lLoop).Name
224:             End If
225:         Next
226:         cmbFieldStripMapName.ListIndex = 0
227:         cmbFieldGridAngle.ListIndex = 0
228:         cmbFieldMapScale.ListIndex = 0
229:         cmbFieldSeriesNumber.ListIndex = 0
230:     End If
231:     SetControlsState
End Sub

Private Sub cmdBack_Click()
235:     m_Step = m_Step - 1
236:     If m_Step < 0 Then
237:         m_Step = 0
238:     End If
239:     SetVisibleControls m_Step
240:     SetControlsState
End Sub

Private Sub cmdClose_Click()
244:     Set m_Application = Nothing
245:     Set Me.StripMapSettings = Nothing
246:     Me.Hide
End Sub

Private Sub CollateStripMapSettings()
    Dim pMx As IMxDocument
    Dim pCreateSMap As New clsCreateStripMap
    Dim pFrameElement As IElement
    Dim sDestLayerName As String
    Dim lLoop As Long
    ' Populate class
256:     pCreateSMap.StripMapName = txtStripMapSeriesName.Text
257:     pCreateSMap.FlipPolyline = (chkFlipLine.value = vbChecked)
258:     If (optScaleSource(0).value) Then
259:         pCreateSMap.MapScale = CDbl(lblCurrentMapScale.Caption)
260:     ElseIf (optScaleSource(1).value) Then
261:         pCreateSMap.MapScale = CDbl(txtManualMapScale.Text)
262:     End If
263:     If (optGridSize(0).value) Then
264:         Set pFrameElement = GetDataFrameElement(GetActiveDataFrameName(m_Application), m_Application)
265:         pCreateSMap.FrameWidthInPageUnits = pFrameElement.Geometry.Envelope.Width
266:         pCreateSMap.FrameHeightInPageUnits = pFrameElement.Geometry.Envelope.Height
267:     Else
268:         pCreateSMap.FrameWidthInPageUnits = CDbl(txtManualGridWidth.Text)
269:         pCreateSMap.FrameHeightInPageUnits = CDbl(txtManualGridHeight.Text)
270:     End If
271:     If (optScaleSource(2).value) Then
        Dim dConvertPageToMapUnits As Double, dGridToFrameRatio As Double
273:         dConvertPageToMapUnits = CalculatePageToMapRatio(m_Application) 'NATHAN FIX THIS
274:         pCreateSMap.FrameWidthInPageUnits = CDbl(txtManualGridWidth.Text)
275:         pCreateSMap.FrameHeightInPageUnits = CDbl(txtManualGridHeight.Text)
276:         If pCreateSMap.FrameWidthInPageUnits >= pCreateSMap.FrameHeightInPageUnits Then
277:             dGridToFrameRatio = CDbl(txtAbsoluteGridWidth.Text) / pCreateSMap.FrameWidthInPageUnits
278:         Else
279:             dGridToFrameRatio = CDbl(txtAbsoluteGridHeight.Text) / pCreateSMap.FrameHeightInPageUnits
280:         End If
281:         pCreateSMap.MapScale = dGridToFrameRatio * dConvertPageToMapUnits
282:     End If
283:     sDestLayerName = cmbPolygonLayers.List(cmbPolygonLayers.ListIndex)
284:     If optLayerSource(0).value Then
285:         Set pCreateSMap.DestinationFeatureLayer = FindFeatureLayerByName(sDestLayerName, m_Application)
286:     End If
287:     pCreateSMap.FieldNameStripMapName = cmbFieldStripMapName.List(cmbFieldStripMapName.ListIndex)
288:     pCreateSMap.FieldNameMapAngle = cmbFieldGridAngle.List(cmbFieldGridAngle.ListIndex)
289:     pCreateSMap.FieldNameNumberInSeries = cmbFieldSeriesNumber.List(cmbFieldSeriesNumber.ListIndex)
290:     If cmbFieldMapScale.ListIndex > 0 Then pCreateSMap.FieldNameScale = cmbFieldMapScale.List(cmbFieldMapScale.ListIndex)
291:     pCreateSMap.RemoveCurrentGrids = (chkRemovePreviousGrids.value = vbChecked)
292:     Set pCreateSMap.StripMapRoute = m_Polyline
    ' Place grid settings on Public form property (so calling function can use them)
294:     Set Me.StripMapSettings = pCreateSMap
End Sub

Private Sub cmdNext_Click()
    Dim pMx As IMxDocument
    Dim pFeatureLayer As IFeatureLayer
    Dim pOutputFClass As IFeatureClass
    Dim pNewFields As IFields
    
    On Error GoTo eh
    ' Step
305:     m_Step = m_Step + 1
    ' If we're creating a new fclass, we can skip a the 'Set Field Roles' step
307:     If m_Step = 1 And (optLayerSource(1).value) Then
308:         m_Step = m_Step + 1
309:     End If
    ' If FINISH
311:     If m_Step >= 4 Then
312:         Set pMx = m_Application.Document
313:         RemoveGraphicsByName pMx
314:         CollateStripMapSettings
        ' If creating a new layer
316:         If optLayerSource(1).value Then
            ' Create the feature class
318:             Set pNewFields = CreateTheFields
            Select Case m_FileType
                Case ShapeFile
321:                     Set pOutputFClass = NewShapeFile(m_OutputLayer, pMx.FocusMap, pNewFields)
                Case AccessFeatureClass
323:                     Set pOutputFClass = NewAccessFile(m_OutputLayer, _
                            m_OutputDataset, m_OutputFClass, pNewFields)
325:             End Select
326:             If pOutputFClass Is Nothing Then
327:                 Err.Raise vbObjectError, "cmdNext", "Could not create the new output feature class."
328:             End If
            ' Create new layer
330:             Set pFeatureLayer = New FeatureLayer
331:             Set pFeatureLayer.FeatureClass = pOutputFClass
332:             pFeatureLayer.Name = pFeatureLayer.FeatureClass.AliasName
            ' Add the new layer to arcmap & reset the StripMapSettings object to point at it
334:             pMx.FocusMap.AddLayer pFeatureLayer
335:             Set StripMapSettings.DestinationFeatureLayer = pFeatureLayer
336:         End If
337:         Me.Hide
338:     Else
339:         SetVisibleControls m_Step
340:         SetControlsState
341:     End If
    
    Exit Sub
eh:
345:     MsgBox "Error: " & Err.Description, , "cmdNext_Click"
346:     m_Step = m_Step - 1
End Sub

Private Sub cmdSetNewGridLayer_Click()
  Dim pGxFilter As IGxObjectFilter
  Dim pGXBrow As IGxDialog, bFlag As Boolean
  Dim pSel As IEnumGxObject, pApp As IApplication
  
354:   Set pGxFilter = New GxFilter
355:   Set pApp = m_Application
356:   Set pGXBrow = New GxDialog
357:   Set pGXBrow.ObjectFilter = pGxFilter
358:   pGXBrow.Title = "Output feature class or shapefile"
359:   bFlag = pGXBrow.DoModalSave(pApp.hwnd)
  
361:   If bFlag Then
    Dim pObj As IGxObject
363:     Set pObj = pGXBrow.FinalLocation
364:     m_bIsGeoDatabase = True
365:     If UCase(pObj.Category) = "FOLDER" Then
366:       If InStr(1, pGXBrow.Name, ".shp") > 0 Then
367:         txtNewGridLayer.Text = pObj.FullName & "\" & pGXBrow.Name
368:       Else
369:         txtNewGridLayer.Text = pObj.FullName & "\" & pGXBrow.Name & ".shp"
370:       End If
371:       m_OutputLayer = txtNewGridLayer.Text
372:       m_bIsGeoDatabase = False
373:       m_FileType = ShapeFile
374:      CheckOutputFile
375:     Else
      Dim pLen As Long
377:       pLen = Len(pObj.FullName) - Len(pObj.BaseName) - 1
378:       txtNewGridLayer.Text = Left(pObj.FullName, pLen)
379:       m_OutputLayer = Left(pObj.FullName, pLen)
380:       m_OutputDataset = pObj.BaseName
381:       m_OutputFClass = pGXBrow.Name
382:       m_bIsGeoDatabase = True
383:       If UCase(pObj.Category) = "PERSONAL GEODATABASE FEATURE DATASET" Then
384:         m_FileType = AccessFeatureClass
385:       Else
386:         m_FileType = SDEFeatureClass
387:       End If
388:     End If
389:   Else
390:     txtNewGridLayer.Text = ""
391:     m_bIsGeoDatabase = False
392:   End If
393:   SetControlsState
End Sub

Private Sub Form_Load()
    Dim pMx As IMxDocument
    Dim bRenewCoordsX As Boolean
    Dim bRenewCoordsY As Boolean
    Dim sErrMsg As String
    On Error GoTo eh
    
403:     sErrMsg = CreateStripMapPolyline
404:     If Len(sErrMsg) > 0 Then
405:         MsgBox sErrMsg, vbCritical, "Create Map Grids"
406:         Unload Me
        Exit Sub
408:     End If
409:     Set pMx = m_Application.Document
410:     Me.Height = 5665
411:     Me.Width = 4935
412:     m_Step = 0
413:     LoadLayersComboBox
414:     LoadUnitsComboBox
415:     lblCurrFrameName.Caption = GetActiveDataFrameName(m_Application)
416:     If pMx.FocusMap.MapUnits = esriUnknownUnits Then
417:         MsgBox "Error: The map has unknown units and therefore cannot calculate a Scale." _
            & vbCrLf & "Cannot create Map Grids at this time.", vbCritical, "Create Map Grids"
419:         Unload Me
        Exit Sub
421:     End If
422:     lblMapUnits.Caption = GetUnitsDescription(pMx.FocusMap.MapUnits)
423:     lblCurrentMapScale.Caption = Format(pMx.FocusMap.MapScale, "#,###,##0")
424:     SetVisibleControls m_Step
    
426:     SetControlsState
    
    'Make sure the wizard stays on top
429:     TopMost Me
    
    Exit Sub
eh:
433:     MsgBox "Error loading the form: " & Err.Description & vbCrLf _
        & vbCrLf & "Attempting to continue the load...", , "MapGridManager: Form_Load "
    On Error Resume Next
436:     SetVisibleControls m_Step
437:     SetControlsState
End Sub

Private Sub LoadUnitsComboBox()
    Dim pMx As IMxDocument
    Dim sPageUnitsDesc As String
    Dim pPage As IPage
    
    On Error GoTo eh
    
    ' Init
448:     Set pMx = m_Application.Document
449:     Set pPage = pMx.PageLayout.Page
450:     sPageUnitsDesc = GetUnitsDescription(pPage.Units)
451:     cmbGridSizeUnits.Clear
    ' Add
453:     cmbGridSizeUnits.AddItem sPageUnitsDesc
    'cmbGridSizeUnits.AddItem "Map Units (" & sMapUnitsDesc & ")"
    ' Set page units as default
456:     cmbGridSizeUnits.ListIndex = 0
    
    Exit Sub
eh:
460:     Err.Raise vbObjectError, "LoadUnitsComboBox", "Error in LoadUnitsComboBox" & vbCrLf & Err.Description
End Sub

Private Sub LoadLayersComboBox()
    Dim pMx As IMxDocument
    Dim lLoop As Long
    Dim pFL As IFeatureLayer
    Dim pFC As IFeatureClass
    Dim sPreviousLayer  As String
    Dim lResetIndex As Long
    
    'Init
472:     Set pMx = m_Application.Document
473:     cmbPolygonLayers.Clear
474:     cmbPolygonLayers.AddItem "<Not Set>"
    ' For all layers
476:     For lLoop = 0 To pMx.FocusMap.LayerCount - 1
        ' If a feature class
478:         If TypeOf pMx.FocusMap.Layer(lLoop) Is IFeatureLayer Then
479:             Set pFL = pMx.FocusMap.Layer(lLoop)
480:             Set pFC = pFL.FeatureClass
            ' If a polygon layer
482:             If pFC.ShapeType = esriGeometryPolygon Then
                ' Add to combo box
484:                 cmbPolygonLayers.AddItem pFL.Name
485:             End If
486:         End If
487:     Next
488:     cmbPolygonLayers.ListIndex = 0
End Sub

Private Sub SetCurrentMapScaleCaption()
    Dim pMx As IMxDocument
    On Error GoTo eh
494:     Set pMx = m_Application.Document
495:     lblCurrentMapScale.Caption = Format(pMx.FocusMap.MapScale, "#,###,##0")
    Exit Sub
eh:
498:     lblCurrentMapScale.Caption = "<Scale Unknown>"
End Sub


Private Sub Form_Unload(Cancel As Integer)
503:     Set m_Application = Nothing
504:     Set StripMapSettings = Nothing
End Sub


Private Sub optGridSize_Click(Index As Integer)
    Dim pMx As IMxDocument
510:     Set pMx = m_Application.Document
511:     lblCurrFrameName.Caption = pMx.FocusMap.Name
512:     SetControlsState
End Sub

Private Sub optLayerSource_Click(Index As Integer)
    ' If creating a new fclass to hold the grids
517:     If Index = 1 Then
        ' Set the field names (will be created automatically)
519:         cmbFieldStripMapName.Clear
520:         cmbFieldGridAngle.Clear
521:         cmbFieldSeriesNumber.Clear
522:         cmbFieldMapScale.Clear
523:         cmbFieldStripMapName.AddItem "<None>"
524:         cmbFieldGridAngle.AddItem "<None>"
525:         cmbFieldSeriesNumber.AddItem "<None>"
526:         cmbFieldMapScale.AddItem "<None>"
527:         cmbFieldStripMapName.AddItem c_DefaultFld_StripMapName
528:         cmbFieldGridAngle.AddItem c_DefaultFld_MapAngle
529:         cmbFieldSeriesNumber.AddItem c_DefaultFld_SeriesNum
530:         cmbFieldMapScale.AddItem c_DefaultFld_MapScale
531:         cmbFieldStripMapName.ListIndex = 1
532:         cmbFieldGridAngle.ListIndex = 1
533:         cmbFieldSeriesNumber.ListIndex = 1
534:         cmbFieldMapScale.ListIndex = 1
535:     End If
536:     SetControlsState
End Sub

Private Sub optScaleSource_Click(Index As Integer)
540:     If Index = 0 Then
541:         SetCurrentMapScaleCaption
542:     ElseIf Index = 2 Then
543:         optGridSize(1).value = True
544:     End If
545:     SetControlsState
End Sub

Private Sub txtAbsoluteGridHeight_Change()
549:     SetControlsState
End Sub

Private Sub txtAbsoluteGridHeight_KeyPress(KeyAscii As Integer)
    ' If a non-numeric (that is not a decimal point)
554:     If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
     And KeyAscii <> Asc(".") _
     And Chr(KeyAscii) <> vbBack Then
        ' Do not allow this button to work
558:         KeyAscii = 0
    ' If a decimal point, make sure we only ever get one of them
560:     ElseIf KeyAscii = Asc(".") Then
561:         If InStr(txtAbsoluteGridHeight.Text, ".") > 0 Then
562:             KeyAscii = 0
563:         End If
564:     End If
End Sub

Private Sub txtAbsoluteGridWidth_Change()
568:     SetControlsState
End Sub

Private Sub txtAbsoluteGridWidth_KeyPress(KeyAscii As Integer)
    ' If a non-numeric (that is not a decimal point)
573:     If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
     And KeyAscii <> Asc(".") _
     And Chr(KeyAscii) <> vbBack Then
        ' Do not allow this button to work
577:         KeyAscii = 0
    ' If a decimal point, make sure we only ever get one of them
579:     ElseIf KeyAscii = Asc(".") Then
580:         If InStr(txtAbsoluteGridWidth.Text, ".") > 0 Then
581:             KeyAscii = 0
582:         End If
583:     End If
End Sub

Private Sub txtManualGridHeight_Change()
587:     SetControlsState
End Sub

Private Sub txtManualGridHeight_KeyPress(KeyAscii As Integer)
    ' If a non-numeric (that is not a decimal point)
592:     If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
     And KeyAscii <> Asc(".") _
     And Chr(KeyAscii) <> vbBack Then
        ' Do not allow this button to work
596:         KeyAscii = 0
    ' If a decimal point, make sure we only ever get one of them
598:     ElseIf KeyAscii = Asc(".") Then
599:         If InStr(txtManualGridHeight.Text, ".") > 0 Then
600:             KeyAscii = 0
601:         End If
602:     End If
End Sub

Private Sub txtManualGridWidth_Change()
606:     If IsNumeric(txtManualGridWidth.Text) And optScaleSource(2).value Then
        Dim dRatio As Double, dGridWidth As Double
608:         dGridWidth = CDbl(txtManualGridWidth.Text)
609:         dRatio = CDbl(txtAbsoluteGridHeight.Text) / CDbl(txtAbsoluteGridWidth.Text)
610:         txtManualGridHeight.Text = CStr(dRatio * dGridWidth)
611:     End If
612:     SetControlsState
End Sub

Private Sub txtManualGridWidth_KeyPress(KeyAscii As Integer)
    ' If a non-numeric (that is not a decimal point)
617:     If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
     And KeyAscii <> Asc(".") _
     And Chr(KeyAscii) <> vbBack Then
        ' Do not allow this button to work
621:         KeyAscii = 0
    ' If a decimal point, make sure we only ever get one of them
623:     ElseIf KeyAscii = Asc(".") Then
624:         If InStr(txtManualGridWidth.Text, ".") > 0 Then
625:             KeyAscii = 0
626:         End If
627:     End If
End Sub

Private Sub txtManualMapScale_Change()
631:     SetControlsState
End Sub

Private Sub txtManualMapScale_KeyPress(KeyAscii As Integer)
    ' If a non-numeric (that is not a decimal point)
636:     If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
     And KeyAscii <> Asc(".") _
     And Chr(KeyAscii) <> vbBack Then
        ' Do not allow this button to work
640:         KeyAscii = 0
    ' If a decimal point, make sure we only ever get one of them
642:     ElseIf KeyAscii = Asc(".") Then
643:         If InStr(txtManualMapScale.Text, ".") > 0 Then
644:             KeyAscii = 0
645:         End If
646:     End If
End Sub

Public Sub Tickle()
650:     Call Form_Load
End Sub

Private Sub SetVisibleControls(iStep As Integer)
    ' Hide all
655:     fraAttributes.Visible = False
656:     fraDataFrameSize.Visible = False
657:     fraDestinationFeatureClass.Visible = False
658:     fraScaleStart.Visible = False
    ' Show applicable frame, set top/left
    Select Case iStep
        Case 0:
662:             fraDestinationFeatureClass.Visible = True
663:             fraDestinationFeatureClass.Top = 0
664:             fraDestinationFeatureClass.Left = 0
        Case 1:
666:             fraAttributes.Visible = True
667:             fraAttributes.Top = 0
668:             fraAttributes.Left = 0
        Case 2:
670:             fraScaleStart.Visible = True
671:             fraScaleStart.Top = 0
672:             fraScaleStart.Left = 0
        Case 3:
674:             fraDataFrameSize.Visible = True
675:             fraDataFrameSize.Top = 0
676:             fraDataFrameSize.Left = 0
        Case Else:
678:             MsgBox "Invalid Step Value : " & iStep
679:     End Select
End Sub

Private Sub CheckOutputFile()
    'Check the output option
684:     If txtNewGridLayer.Text <> "" Then
685:         If DoesShapeFileExist(txtNewGridLayer.Text) Then
686:             MsgBox "Shape file name already being used!!!"
687:             txtNewGridLayer.Text = ""
688:         End If
689:     End If
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
702:     Set pNewFields = New Fields
703:     Set pFieldsEdit = pNewFields
704:     Set pMx = m_Application.Document
    ' Field: OID  -------------------------
706:     Set newField = New Field
707:     Set newFieldEdit = newField
708:     With newFieldEdit
709:         .Name = "OID"
710:         .Type = esriFieldTypeOID
711:         .AliasName = "Object ID"
712:         .IsNullable = False
713:     End With
714:     pFieldsEdit.AddField newField
    ' Field: STRIP MAP NAME -------------------------
716:     Set newField = New Field
717:     Set newFieldEdit = newField
718:     With newFieldEdit
719:       .Name = c_DefaultFld_StripMapName
720:       .AliasName = "StripMapName"
721:       .Type = esriFieldTypeString
722:       .IsNullable = True
723:       .length = 50
724:     End With
725:     pFieldsEdit.AddField newField
    ' Field: MAP ANGLE -------------------------
727:     Set newField = New Field
728:     Set newFieldEdit = newField
729:     With newFieldEdit
730:       .Name = c_DefaultFld_MapAngle
731:       .AliasName = "Map Angle"
732:       .Type = esriFieldTypeInteger
733:       .IsNullable = True
734:     End With
735:     pFieldsEdit.AddField newField
    ' Field: GRID NUMBER -------------------------
737:     Set newField = New Field
738:     Set newFieldEdit = newField
739:     With newFieldEdit
740:       .Name = c_DefaultFld_SeriesNum
741:       .AliasName = "Number In Series"
742:       .Type = esriFieldTypeInteger
743:       .IsNullable = True
744:     End With
745:     pFieldsEdit.AddField newField
    ' Field: SCALE -------------------------
747:     Set newField = New Field
748:     Set newFieldEdit = newField
749:     With newFieldEdit
750:       .Name = c_DefaultFld_MapScale
751:       .AliasName = "Plot Scale"
752:       .Type = esriFieldTypeDouble
753:       .IsNullable = True
754:       .Precision = 18
755:       .Scale = 11
756:     End With
757:     pFieldsEdit.AddField newField
    ' Return
759:     Set CreateTheFields = pFieldsEdit
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
774:     Set pMx = pApp.Document
775:     Set pSR = pMx.FocusMap.SpatialReference
776:     If TypeOf pSR Is IProjectedCoordinateSystem Then
777:         Set pPCS = pSR
778:         dMetersPerUnit = pPCS.CoordinateUnit.MetersPerUnit
779:     Else
780:         dMetersPerUnit = 1
781:     End If
782:     Set pPage = pMx.PageLayout.Page
783:     pPageUnits = pPage.Units
    Select Case pPageUnits
        Case esriInches: CalculatePageToMapRatio = dMetersPerUnit / (1 / 12 * 0.304800609601219)
        Case esriFeet: CalculatePageToMapRatio = dMetersPerUnit / (0.304800609601219)
        Case esriCentimeters: CalculatePageToMapRatio = dMetersPerUnit / (1 / 100)
        Case esriMeters: CalculatePageToMapRatio = dMetersPerUnit / (1)
        Case Else:
790:             MsgBox "Warning: Only the following Page (Layout) Units are supported by this tool:" _
                & vbCrLf & " - Inches, Feet, Centimeters, Meters" _
                & vbCrLf & vbCrLf & "Calculating as though Page Units are in Inches..."
793:             CalculatePageToMapRatio = dMetersPerUnit / (1 / 12 * 0.304800609601219)
794:     End Select
    Exit Function
eh:
797:     CalculatePageToMapRatio = 1
798:     MsgBox "Error in CalculatePageToMapRatio" & vbCrLf & Err.Description
End Function

Private Function ReturnMax(dDouble1 As Double, dDouble2 As Double) As Double
802:     If dDouble1 >= dDouble2 Then
803:         ReturnMax = dDouble1
804:     Else
805:         ReturnMax = dDouble2
806:     End If
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
823:     Set pMx = m_Application.Document
824:     Set pFC = pMx.FocusMap.FeatureSelection
825:     Set pF = pFC.Next
826:     If pF Is Nothing Then
827:         CreateStripMapPolyline = "Requires selected polyline features/s."
        Exit Function
829:     End If
    ' Make polyline
831:     Set pPolyline = New Polyline
832:     While Not pF Is Nothing
833:         If pF.Shape.GeometryType = esriGeometryPolyline Then
834:             Set pTmpPolyline = pF.ShapeCopy
835:             Set pTopoSimplify = pTmpPolyline
836:             pTopoSimplify.Simplify
837:             Set pTopoUnion = pPolyline
838:             Set pPolyline = pTopoUnion.Union(pTopoSimplify)
839:             Set pTopoSimplify = pPolyline
840:             pTopoSimplify.Simplify
841:         End If
842:         Set pF = pFC.Next
843:     Wend
    ' Check polyline for beinga single, connected polyline (Path)
845:     Set pGeoColl = pPolyline
846:     If pGeoColl.GeometryCount = 0 Then
847:         CreateStripMapPolyline = "Requires selected polyline features/s."
        Exit Function
849:     ElseIf pGeoColl.GeometryCount > 1 Then
850:         CreateStripMapPolyline = "Cannot process the StripMap - multi-part polyline created." _
            & vbCrLf & "Check for non-connected segments, overlaps or loops."
        Exit Function
853:     End If
    ' Give option to flip
855:     Perm_DrawPoint pPolyline.FromPoint, , 0, 255, 0, 20
856:     Perm_DrawTextFromPoint pPolyline.FromPoint, "START", , , , , 20
857:     Perm_DrawPoint pPolyline.ToPoint, , 255, 0, 0, 20
858:     Perm_DrawTextFromPoint pPolyline.ToPoint, "END", , , , , 20
859:     pMx.ActiveView.PartialRefresh esriViewGraphics, Nothing, Nothing
    
861:     Set m_Polyline = pPolyline
    
863:     CreateStripMapPolyline = ""
    
    Exit Function
866:     Resume
eh:
868:     CreateStripMapPolyline = "Error in CreateStripMapPolyline : " & Err.Description
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
886:     Set pMx = m_Application.Document
887:     Set pGLayer = pMx.FocusMap.BasicGraphicsLayer
888:     Set pGCon = pGLayer
889:     Set pElement = New MarkerElement
890:     pElement.Geometry = pPoint
891:     Set pMarkerElement = pElement
    
    ' Set the symbol
894:     Set pColor = New RgbColor
895:     pColor.Red = dRed
896:     pColor.Green = dGreen
897:     pColor.Blue = dBlue
898:     Set pMarker = New SimpleMarkerSymbol
899:     With pMarker
900:         .Color = pColor
901:         .Size = dSize
902:     End With
903:     pMarkerElement.Symbol = pMarker
    
    ' Add the graphic
906:     Set pElementProp = pElement
907:     pElementProp.Name = sElementName
908:     pGCon.AddElement pElement, 0
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
931:     Set pMx = m_Application.Document
932:     Set pGLayer = pMx.FocusMap.BasicGraphicsLayer
933:     Set pGCon = pGLayer
934:     Set pElement = New LineElement
    
    ' Set the line symbol
937:     Set pLnSym = New SimpleLineSymbol
938:     Set myColor = New RgbColor
939:     myColor.Red = dRed
940:     myColor.Green = dGreen
941:     myColor.Blue = dBlue
942:     pLnSym.Color = myColor
943:     pLnSym.Width = dSize
    
    ' Create a standard polyline (via 2 points)
946:     Set pLine1 = New esrigeometry.Line
947:     pLine1.PutCoords pFromPoint, pToPoint
948:     Set pSeg1 = pLine1
949:     Set pPolyline = New Polyline
950:     pPolyline.AddSegment pSeg1
951:     pElement.Geometry = pPolyline
952:     Set pLineElement = pElement
953:     pLineElement.Symbol = pLnSym
    
    ' Add the graphic
956:     Set pElementProp = pElement
957:     pElementProp.Name = sElementName
958:     pGCon.AddElement pElement, 0
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
976:     Set pMx = m_Application.Document
977:     Set pGLayer = pMx.FocusMap.BasicGraphicsLayer
978:     Set pGCon = pGLayer
979:     Set pElement = New TextElement
980:     pElement.Geometry = pPoint
981:     Set pTextElement = pElement
    
    ' Create the text symbol
984:     Set myTxtSym = New TextSymbol
985:     Set myColor = New RgbColor
986:     myColor.Red = dRed
987:     myColor.Green = dGreen
988:     myColor.Blue = dBlue
989:     myTxtSym.Color = myColor
990:     myTxtSym.Size = dSize
991:     myTxtSym.HorizontalAlignment = esriTHACenter
992:     pTextElement.Symbol = myTxtSym
993:     pTextElement.Text = sText
    
    ' Add the graphic
996:     Set pElementProp = pElement
997:     pElementProp.Name = sElementName
998:     pGCon.AddElement pElement, 0
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
1014:     pMxDoc.DelayUpdateContents = True
1015:     Set pGLayer = pMxDoc.FocusMap.BasicGraphicsLayer
1016:     Set pGCon = pGLayer
1017:     pGCon.Next
    
    ' Delete all the graphic elements that we created (identify by the name prefix)
1020:     pGCon.Reset
1021:     Set pElement = pGCon.Next
1022:     While Not pElement Is Nothing
1023:         If TypeOf pElement Is IElement Then
1024:             Set pElementProp = pElement
1025:             If (Left(pElementProp.Name, Len(sPrefix)) = sPrefix) Then
1026:                 pGCon.DeleteElement pElement
1027:             End If
1028:         End If
1029:         Set pElement = pGCon.Next
1030:     Wend
    
    ' Switch ON the updating of the TOC, refresh
1033:     pMxDoc.DelayUpdateContents = False
1034:     pMxDoc.ActiveView.Refresh
    
    Exit Sub
ErrorHandler:
1038:     MsgBox "Error in RemoveGraphicsByName: " & Err.Description, , "RemoveGraphicsByName"
End Sub

Private Sub txtStripMapSeriesName_Change()
1042:     SetControlsState
End Sub
