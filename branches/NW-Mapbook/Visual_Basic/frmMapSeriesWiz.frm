VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMapSeriesWiz 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Map Sheet Wizard"
   ClientHeight    =   4152
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   6912
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4152
   ScaleWidth      =   6912
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraPage2 
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   6855
      Begin VB.TextBox txtNumbering 
         Height          =   315
         Left            =   2370
         TabIndex        =   43
         Text            =   "1"
         Top             =   2580
         Width           =   525
      End
      Begin VB.Frame fraSuppressTiles 
         Caption         =   "Suppress tiles"
         Height          =   2565
         Left            =   3420
         TabIndex        =   20
         Top             =   750
         Width           =   3255
         Begin VB.ListBox lstSuppressTiles 
            Height          =   1344
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
         Height          =   1515
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
         Left            =   90
         TabIndex        =   44
         Top             =   2610
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   $"frmMapSeriesWiz.frx":0000
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   0
         Width           =   6705
      End
   End
   Begin VB.Frame fraPage1 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   600
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
      Begin VB.Label lblMapSheet 
         Caption         =   "Choose the detail data frame:"
         Height          =   225
         Index           =   0
         Left            =   30
         TabIndex        =   8
         Top             =   870
         Width           =   2235
      End
      Begin VB.Label Label1 
         Caption         =   $"frmMapSeriesWiz.frx":00F8
         Height          =   615
         Index           =   0
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   6705
      End
   End
   Begin VB.Frame fraPage3 
      BorderStyle     =   0  'None
      Height          =   3525
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   6855
      Begin VB.Frame fraOptions 
         Caption         =   "Options"
         Height          =   2565
         Left            =   3450
         TabIndex        =   28
         Top             =   750
         Width           =   3255
         Begin VB.TextBox txtNeighborLabelIndent 
            Height          =   285
            Left            =   1800
            TabIndex        =   45
            Text            =   "txtNeighborLabelIndent"
            Top             =   2160
            Width           =   1215
         End
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
         Begin VB.Label lblIndentDistance 
            Caption         =   "Label Indent (in) :"
            Height          =   255
            Left            =   1800
            TabIndex        =   46
            Top             =   1920
            Width           =   1335
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
      Y1              =   3600
      Y2              =   3600
   End
End
Attribute VB_Name = "frmMapSeriesWiz"
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

Private m_iPage As Integer
Public m_pApp As IApplication
Private m_pCurrentFrame As Frame
Private m_pMap As IMap
Private m_pIndexLayer As IFeatureLayer
Private m_bFormLoad As Boolean
Private m_pTextSym As ISimpleTextSymbol
Private m_pNWSeriesOpt As INWMapSeriesOptions
Private m_pMapSeries As INWDSMapSeries
Const c_sModuleFileName As String = "frmMapSeriesWiz.frm"


'Private Declare Function SetForegroundWindow Lib "user32" (ByVal HWnd As Long) As Long

Private Sub PositionFrame(pFrame As Frame)
On Error GoTo ErrHand:

48:   If Not m_pCurrentFrame Is Nothing Then m_pCurrentFrame.Visible = False
49:   pFrame.Visible = True
50:   pFrame.Height = 3495
51:   pFrame.Width = 6825
52:   pFrame.Left = 30
53:   pFrame.Top = 30
54:   Set m_pCurrentFrame = pFrame
55:   pFrame.Visible = True
     
  Exit Sub
ErrHand:
59:   MsgBox "PositionFrame - " & Err.Description
  Exit Sub
End Sub

Private Sub chkOptions_Click(Index As Integer)
  Select Case Index
  Case 0  'Rotate
66:     If chkOptions(0).Value = 0 Then
67:       cmbRotateField.Enabled = False
68:     Else
69:       cmbRotateField.Enabled = True
70:     End If
  Case 1  'Clip to outline
72:     If chkOptions(1).Value = 0 Then
73:       chkOptions(3).Enabled = False
74:       chkOptions(3).Value = 0
75:     Else
76:       chkOptions(3).Enabled = True
77:     End If
  Case 2  'Label neighboring tiles
79:     If chkOptions(2).Value = 0 Then
80:       cmdLabelProps.Enabled = False
81:     Else
82:       cmdLabelProps.Enabled = True
83:     End If
84:   End Select
End Sub

Private Sub chkSuppress_Click()
88:   If chkSuppress.Value = 0 Then
89:     lstSuppressTiles.Enabled = False
90:   Else
91:     lstSuppressTiles.Enabled = True
92:   End If
End Sub

Private Sub cmbDetailFrame_Click()
On Error GoTo ErrHand:
  Dim pDoc As IMxDocument, lLoop As Long
  Dim pFeatLayer As IFeatureLayer, pGroupLayer As ICompositeLayer
  
  'Set the Next button to false
101:   cmdNext.Enabled = False
  
  'Find the selected map
104:   cmbIndexLayer.Clear
105:   If cmbDetailFrame.Text = "" Then
106:     MsgBox "No detail frame selected!!!"
    Exit Sub
108:   End If
  
110:   Set pDoc = m_pApp.Document
111:   Set m_pMap = FindDataFrame(pDoc, cmbDetailFrame.Text)
112:   If m_pMap Is Nothing Then
113:     MsgBox "Could not find detail frame!!!"
    Exit Sub
115:   End If
  
  'Populate the index layer combo
118:   lstSuppressTiles.Clear
119:   cmbIndexLayer.Clear
120:   For lLoop = 0 To m_pMap.LayerCount - 1
121:     If TypeOf m_pMap.Layer(lLoop) Is ICompositeLayer Then
122:       CompositeLayer m_pMap.Layer(lLoop)
123:     Else
124:       LayerCheck m_pMap.Layer(lLoop)
125:     End If
126:   Next lLoop
127:   If cmbIndexLayer.ListCount = 0 Then
128:     MsgBox "You need at least one polygon layer in the detail frame to serve as the index layer!!!"
129:   Else
130:     cmbIndexLayer.ListIndex = 0
131:   End If
  
  
  Exit Sub
ErrHand:
136:   MsgBox "cmbDetailFrame_Click - " & Err.Description
End Sub

Private Sub CompositeLayer(pCompLayer As ICompositeLayer)
On Error GoTo ErrHand:
  Dim lLoop As Long
142:   For lLoop = 0 To pCompLayer.count - 1
143:     If TypeOf pCompLayer.Layer(lLoop) Is ICompositeLayer Then
144:       CompositeLayer pCompLayer.Layer(lLoop)
145:     Else
146:       LayerCheck pCompLayer.Layer(lLoop)
147:     End If
148:   Next lLoop

  Exit Sub
ErrHand:
152:   MsgBox "CompositeLayer - " & Err.Description
End Sub

Private Sub LayerCheck(pLayer As ILayer)
On Error GoTo ErrHand:
  Dim pFeatLayer As IFeatureLayer
  
159:   If TypeOf pLayer Is IFeatureLayer Then
160:     Set pFeatLayer = pLayer
161:     If pFeatLayer.FeatureClass.ShapeType = esriGeometryPolygon Then
162:       cmbIndexLayer.AddItem pFeatLayer.Name
163:     End If
164:     lstSuppressTiles.AddItem pFeatLayer.Name
165:   End If

  Exit Sub
ErrHand:
169:   MsgBox "LayerCheck - " & Err.Description
End Sub

Private Sub cmbIndexLayer_Click()
On Error GoTo ErrHand:
  Dim lLoop As Long, pFields As IFields, pField As IField
  
  'Set the Next button to false
177:   cmdNext.Enabled = False
  
  'Find the selected layer
180:   cmbIndexField.Clear
181:   If cmbIndexLayer.Text = "" Then
182:     MsgBox "No index layer selected!!!"
    Exit Sub
184:   End If
  
186:   Set m_pIndexLayer = FindLayer(cmbIndexLayer.Text, m_pMap)
187:   If m_pIndexLayer Is Nothing Then
188:     MsgBox "Could not find specified layer!!!"
    Exit Sub
190:   End If
  
  'Populate the index layer combos
193:   Set pFields = m_pIndexLayer.FeatureClass.Fields
194:   cmbDataDriven.Clear
195:   cmbRotateField.Clear
196:   For lLoop = 0 To pFields.FieldCount - 1
    Select Case pFields.Field(lLoop).Type
    Case esriFieldTypeString
199:       cmbIndexField.AddItem pFields.Field(lLoop).Name
    Case esriFieldTypeDouble, esriFieldTypeSingle, esriFieldTypeInteger
201:       If UCase(pFields.Field(lLoop).Name) <> "SHAPE_LENGTH" And _
       UCase(pFields.Field(lLoop).Name) <> "SHAPE_AREA" Then
203:         cmbDataDriven.AddItem pFields.Field(lLoop).Name
204:         cmbRotateField.AddItem pFields.Field(lLoop).Name
205:       End If
206:     End Select
207:   Next lLoop
208:   If cmbIndexField.ListCount = 0 Then
'    MsgBox "You need at least one string field in the layer for labeling the pages!!!"
210:   Else
211:     cmbIndexField.ListIndex = 0
212:     cmdNext.Enabled = True
213:   End If
214:   If cmbDataDriven.ListCount > 0 Then
215:     cmbDataDriven.ListIndex = 0
216:     cmbRotateField.ListIndex = 0
217:     optExtent.Item(2).Enabled = True
218:     chkOptions(0).Enabled = True
219:   Else
220:     optExtent.Item(2).Enabled = False
221:     chkOptions(0).Enabled = False
222:   End If

  Exit Sub
ErrHand:
226:   MsgBox "cmbIndexField_Click - " & Err.Description
End Sub

Private Sub cmdBack_Click()
230:   m_pCurrentFrame.Visible = False
  Select Case m_iPage
  Case 2
233:     PositionFrame fraPage1
234:     m_iPage = 1
  Case 3
236:     cmdNext.Caption = "Next >"
237:     PositionFrame fraPage2
238:     m_iPage = 2
239:   End Select
240:   cmdNext.Enabled = True
End Sub

Private Sub cmdCancel_Click()
244:   Unload Me
End Sub

Private Sub cmdLabelProps_Click()
  On Error GoTo ErrorHandler

250:   If m_pNWSeriesOpt Is Nothing Then
251:     If m_pMapSeries Is Nothing Then
      'Set m_pMapSeries = New NWDSMapBook
253:       Set m_pMapSeries = New NWMapBook
254:     End If
255:     Set m_pNWSeriesOpt = m_pMapSeries
256:   End If
257:   Set frmAdjMapLabelSymbols.NWSeriesOptions = m_pNWSeriesOpt
258:   Set frmAdjMapLabelSymbols.Application = m_pApp
  
260:   frmAdjMapLabelSymbols.Show vbModal, Me
261:   Me.SetFocus
  
263:   If Not frmAdjMapLabelSymbols.NWSeriesOptions Is Nothing Then
264:     Set m_pNWSeriesOpt = frmAdjMapLabelSymbols.NWSeriesOptions
265:   End If
  Exit Sub


  Exit Sub
ErrorHandler:
  HandleError True, "cmdLabelProps_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdNext_Click()
On Error GoTo ErrHand:
  Dim pMapSeries As INWDSMapSeries
277:   m_pCurrentFrame.Visible = False
278:   cmdBack.Enabled = True
  Select Case m_iPage
  Case 1  'Done with date frame and index layer
281:     CheckForSelected    'Check index layer to see if there are selected features
282:     PositionFrame fraPage2
283:     m_iPage = 2
  Case 2  'Done with tile specification
285:     PositionFrame fraPage3
286:     m_iPage = 3
287:     cmdNext.Caption = "Finish"
288:     If optExtent(0).Value Then
289:       If txtMargin.Text = "" Then
290:         cmdNext.Enabled = False
291:       Else
292:         cmdNext.Enabled = True
293:       End If
294:     ElseIf optExtent(1).Value Then
295:       If txtFixed.Text = "" Then
296:         cmdNext.Enabled = False
297:       Else
298:         cmdNext.Enabled = True
299:       End If
300:     Else
301:       cmdNext.Enabled = True
302:     End If
  Case 3  'Finish button selected
304:     CreateSeries
305:     Unload Me
306:   End Select
  
  Exit Sub
ErrHand:
310:   MsgBox "cmdNext_click - " & Err.Description
  Exit Sub
End Sub

Private Sub CreateSeries()
  On Error GoTo ErrorHandler


  'Dim pMapSeries As INWDSMapSeries, pSpatialQuery As ISpatialFilter
  Dim pSpatialQuery As ISpatialFilter
  Dim pTmpPage As tmpPageClass, pTmpColl As Collection, pClone As IClone
  Dim pSeriesOpt As INWDSMapSeriesOptions, pFeatLayerSel As IFeatureSelection
  Dim pSeriesProps As INWDSMapSeriesProps, pMapPage As INWDSMapPage
  Dim pDoc As IMxDocument, pMap As IMap, lCount As Long, lLoop As Long
  Dim pFeatLayer As IFeatureLayer, pQuery As IQueryFilter, pCursor As IFeatureCursor
  Dim pFeature As IFeature, lIndex As Long, sName As String, sFieldName As String
  Dim pNode As Node, pMapBook As INWDSMapBook
  Dim pActiveView As IActiveView, lRotIndex As Long, lScaleIndex As Long
  'Added 6/18/03 to support cross hatch outside clip area
  Dim pSeriesOpt2 As INWDSMapSeriesOptions2
  'Add 2/18/04 to support the storing of page numbers
  Dim lPageNumber As Long
  
  
334:   Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
337:   pMapBook.EnableBook = True
338:   Set pDoc = m_pApp.Document
  
  
  
342:   If m_pMapSeries Is Nothing Then
    'Set m_pMapSeries = New NWDSMapSeries
344:     Set m_pMapSeries = New NWMapSeries
345:   End If
346:   Set pSeriesOpt = m_pMapSeries
347:   Set pSeriesOpt2 = pSeriesOpt  'Added 6/18/03 to support cross hatch outside clip area
348:   Set pSeriesProps = m_pMapSeries
349:   If m_pNWSeriesOpt Is Nothing Then
350:     Set m_pNWSeriesOpt = m_pMapSeries 'Added 9/1/04 to support NW Mapbook features
351:   End If
352:   m_pNWSeriesOpt.DataFrameMainFrame = cmbDetailFrame.Text
353:   m_pMapSeries.EnableSeries = True
  
  'Find the detail frame
356:   Set pMap = FindDataFrame(pDoc, cmbDetailFrame.Text)
357:   If pMap Is Nothing Then
358:     MsgBox "Detail frame not found!!!"
    Exit Sub
360:   End If
361:   pSeriesProps.DataFrameName = pMap.Name
  
  'Find the layer
364:   Set pFeatLayer = FindLayer(cmbIndexLayer.Text, pMap)
365:   If pFeatLayer Is Nothing Then
366:     MsgBox "Index layer not found!!!"
    Exit Sub
368:   End If
369:   pSeriesProps.IndexLayerName = pFeatLayer.Name
370:   pSeriesProps.IndexFieldName = cmbIndexField.Text
    
  'Determine the tiles we are interested in
373:   Set pQuery = New QueryFilter
374:   sFieldName = cmbIndexField.Text
375:   pQuery.AddField sFieldName
'  pQuery.WhereClause = sFieldName & " <> ''"
377:   pQuery.WhereClause = sFieldName & " is not null"
378:   If optTiles(0).Value Then
379:     Set pCursor = pFeatLayer.Search(pQuery, True)
380:     pSeriesProps.TileSelectionMethod = 0
381:   ElseIf optTiles(1).Value Then
382:     Set pFeatLayerSel = pFeatLayer
383:     pFeatLayerSel.SelectionSet.Search pQuery, True, pCursor
384:     pSeriesProps.TileSelectionMethod = 1
385:   Else
386:     Set pActiveView = pMap
387:     Set pSpatialQuery = New SpatialFilter
388:     pSpatialQuery.AddField sFieldName
389:     pSpatialQuery.SpatialRel = esriSpatialRelIntersects
390:     Set pSpatialQuery.Geometry = pActiveView.Extent
391:     pSpatialQuery.WhereClause = sFieldName & " <> ''"
392:     pSpatialQuery.GeometryField = pFeatLayer.FeatureClass.shapeFieldName
393:     Set pCursor = pFeatLayer.Search(pSpatialQuery, True)
394:     pSeriesProps.TileSelectionMethod = 2
395:   End If
  
  'Add 2/18/04 to keep track of starting page number
398:   pSeriesProps.StartNumber = CLng(txtNumbering.Text)
  
  'Set the clip, label and rotate properties
  'Updated 6/18/03 to support cross hatch outside clip area
402:   If chkOptions(1).Value = 1 Then
403:     If chkOptions(3).Value = 1 Then
404:       pSeriesOpt2.ClipData = 2
405:     Else
406:       pSeriesOpt2.ClipData = 1
407:     End If
408:   Else
409:     pSeriesOpt2.ClipData = 0
410:   End If
'  If chkOptions(1).Value = 1 Then
'    pSeriesOpt.ClipData = True
'  Else
'    pSeriesOpt.ClipData = False
'  End If
  
417:   If chkOptions(0).Value = 1 Then
418:     pSeriesOpt.RotateFrame = True
419:     pSeriesOpt.RotationField = cmbRotateField.Text
420:     lRotIndex = pFeatLayer.FeatureClass.FindField(cmbRotateField.Text)
421:   Else
422:     pSeriesOpt.RotateFrame = False
423:   End If
  
425:   If chkOptions(2).Value = 1 Then
426:     pSeriesOpt.LabelNeighbors = True
427:   Else
428:     pSeriesOpt.LabelNeighbors = False
429:   End If
430:   If IsNumeric(txtNeighborLabelIndent.Text) And pSeriesOpt.LabelNeighbors Then
431:     m_pNWSeriesOpt.NeighborLabelIndent = CDbl(txtNeighborLabelIndent.Text)
432:   End If
433:   Set pSeriesOpt.LabelSymbol = m_pTextSym
  
  
  'Set the extent properties
437:   If optExtent(0).Value Then         'Variable
438:     pSeriesOpt.ExtentType = 0
439:     If txtMargin.Text = "" Then
440:       pSeriesOpt.Margin = 0
441:     Else
442:       pSeriesOpt.Margin = CDbl(txtMargin.Text)
443:     End If
444:     pSeriesOpt.MarginType = cmbMargin.ListIndex
445:   ElseIf optExtent(1).Value Then    'Fixed
446:     pSeriesOpt.ExtentType = 1
447:     pSeriesOpt.FixedScale = txtFixed.Text
448:   Else                        'Data driven
449:     pSeriesOpt.ExtentType = 2
450:     pSeriesOpt.DataDrivenField = cmbDataDriven.Text
451:     lScaleIndex = pFeatLayer.FeatureClass.FindField(cmbDataDriven.Text)
452:   End If
  
  'Store suppression information
455:   If chkSuppress.Value = 1 And lstSuppressTiles.SelCount > 0 Then
456:     pSeriesProps.SuppressLayers = True
457:     For lLoop = 0 To lstSuppressTiles.ListCount - 1
458:       If lstSuppressTiles.Selected(lLoop) Then
459:         pSeriesProps.AddLayerToSuppress lstSuppressTiles.List(lLoop)
460:       End If
461:     Next lLoop
462:   Else
463:     pSeriesProps.SuppressLayers = False
464:   End If
  
  'Create the pages and populate the treeview
467:   Set pTmpColl = New Collection
468:   lIndex = pFeatLayer.FeatureClass.FindField(sFieldName)
469:   Set pFeature = pCursor.NextFeature
470:   With g_pFrmMapSeries.tvwMapBook
471:     Set pNode = .Nodes.Add("MapBook", tvwChild, "MapSeries", "Map Series", 3)
472:     pNode.Tag = "MapSeries"
    
    'Add tile names to a listbox first for sort purposes
475:     g_pFrmMapSeries.lstSorter.Clear
476:     Do While Not pFeature Is Nothing
477:       sName = pFeature.Value(lIndex)
478:       Set pTmpPage = New tmpPageClass
479:       pTmpPage.PageName = sName
480:       pTmpPage.PageRotation = 0
481:       pTmpPage.PageScale = 1
482:       Set pClone = pFeature.Shape
483:       Set pTmpPage.PageShape = pClone.Clone
      'Track the rotation and scale values (if we are going to use them) to the end
      'of the name, so we can assign them to the page when it is added without having
      'to query the index layer again.
487:       If chkOptions(0).Value = 1 And lRotIndex >= 0 Then
488:         If Not IsNull(pFeature.Value(lRotIndex)) Then
489:           pTmpPage.PageRotation = pFeature.Value(lRotIndex)
490:         End If
491:       End If
492:       If optExtent(2).Value And lScaleIndex >= 0 Then
493:         If Not IsNull(pFeature.Value(lScaleIndex)) Then
494:           pTmpPage.PageScale = pFeature.Value(lScaleIndex)
495:         End If
496:       End If
497:       If chkSuppress.Value = 1 And lstSuppressTiles.SelCount > 0 Then
498:         If FeaturesInTile(pFeature, pMap) Then
499:           g_pFrmMapSeries.lstSorter.AddItem sName
500:           pTmpColl.Add pTmpPage, sName
501:         End If
502:       Else
503:         g_pFrmMapSeries.lstSorter.AddItem sName
504:         pTmpColl.Add pTmpPage, sName
505:       End If
506:       Set pFeature = pCursor.NextFeature
507:     Loop
    
    'Now loop back through the list and add the tile names as nodes in the tree
510:     For lLoop = 0 To g_pFrmMapSeries.lstSorter.ListCount - 1
      'Set pMapPage = New NWDSMapPage
512:       Set pMapPage = New NWMapPage
513:       lPageNumber = lLoop + CLng(txtNumbering.Text)
514:       sName = g_pFrmMapSeries.lstSorter.List(lLoop)
515:       Set pNode = .Nodes.Add("MapSeries", tvwChild, "a" & sName, lPageNumber & " - " & sName, 5)
516:       Set pTmpPage = pTmpColl.Item(sName)
517:       pNode.Tag = lLoop
518:       pMapPage.PageName = sName
519:       pMapPage.PageRotation = pTmpPage.PageRotation
520:       pMapPage.PageScale = pTmpPage.PageScale
521:       Set pMapPage.PageShape = pTmpPage.PageShape
522:       pMapPage.LastOutputted = #1/1/1900#
523:       pMapPage.EnablePage = True
524:       pMapPage.PageNumber = lPageNumber
525:       m_pMapSeries.AddPage pMapPage
526:     Next lLoop
527:     .Nodes.Item("MapBook").Expanded = True
528:     .Nodes.Item("MapSeries").Expanded = True
529:   End With

  
  'Add the series to the book
533:   pMapBook.AddContent m_pMapSeries

  Exit Sub
'ErrHandler:
'  MsgBox "CreateSeries - most likely you do not have unique names in your index layer!!!"

  Exit Sub
ErrorHandler:
  HandleError False, "CreateSeries " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub CheckForSelected()
On Error GoTo ErrHand:
  Dim pFeatSel As IFeatureSelection
  
  'Make sure there is something to check
549:   optTiles(1).Enabled = False
  If m_pIndexLayer Is Nothing Then Exit Sub
  
  'Check for selected features in the index layer
553:   Set pFeatSel = m_pIndexLayer
554:   If pFeatSel.SelectionSet.count <> 0 Then
555:     optTiles(1).Enabled = True
556:   End If

  Exit Sub
ErrHand:
560:   MsgBox "CheckForSelected - " & Err.Description
End Sub


Private Sub Form_Load()
  On Error GoTo ErrorHandler

  Dim pDoc As IMxDocument, lLoop As Long
  'Get the extension
  If m_pApp Is Nothing Then Exit Sub
    
571:   m_bFormLoad = True
572:   Set m_pCurrentFrame = Nothing
573:   PositionFrame fraPage1
574:   cmdNext.Enabled = False
575:   cmdBack.Enabled = False
  
  'NW Map Book new code
578:   txtNeighborLabelIndent.Text = ""
  'Initialize variables and controls
580:   m_iPage = 1
581:   chkOptions(0).Value = 0
582:   chkOptions(1).Value = 0
583:   chkOptions(2).Value = 0
584:   chkSuppress.Value = 0
585:   optTiles(0).Value = True
586:   optExtent(0).Value = True
587:   lstSuppressTiles.Enabled = False
588:   cmbRotateField.Enabled = False
589:   cmdLabelProps.Enabled = False
590:   chkOptions(3).Enabled = False
591:   txtNumbering.Text = "1"
  
  'Populate the data frame combo
594:   Set pDoc = m_pApp.Document
595:   cmbIndexField.Clear
596:   cmbDetailFrame.Clear
597:   For lLoop = 0 To pDoc.Maps.count - 1
598:     cmbDetailFrame.AddItem pDoc.Maps.Item(lLoop).Name
599:   Next lLoop
600:   cmbDetailFrame.ListIndex = 0
601:   m_bFormLoad = False
  
  'Populate the extent options
604:   cmbMargin.Clear
605:   cmbMargin.AddItem "percent"
606:   cmbMargin.AddItem "mapunits"
607:   cmbMargin.Text = "percent"
608:   txtMargin.Text = "0"
  
  'Set the initial Label symbol
611:   Set pDoc = m_pApp.Document
612:   Set m_pTextSym = New TextSymbol
613:   m_pTextSym.Font = pDoc.DefaultTextFont
614:   m_pTextSym.Size = pDoc.DefaultTextFontSize.Size
  
616:   Set m_pMapSeries = Nothing
617:   Set m_pMapSeries = New NWMapSeries
618:   Set m_pNWSeriesOpt = m_pMapSeries
  'If m_pNWSeriesOpt Is Nothing Then
  '  If m_pMapSeries Is Nothing Then
  '    'Set m_pMapSeries = New NWDSMapSeries
  '    Set m_pMapSeries = New NWMapSeries
  '  End If
  '  Set m_pNWSeriesOpt = m_pMapSeries
  'End If
626:   If frmAdjMapLabelSymbols.NWSeriesOptions Is Nothing Then
627:     Set frmAdjMapLabelSymbols.NWSeriesOptions = m_pNWSeriesOpt
628:   End If
629:   If m_pNWSeriesOpt.TextSymbolCount = 0 Then
630:     m_pNWSeriesOpt.TextSymbolAdd m_pTextSym, "Initial Text Symbol"
631:     m_pNWSeriesOpt.TextSymbolDefault = "Initial Text Symbol"
632:   End If
  
  
  'Make sure the wizard stays on top
'623:   TopMost Me



  Exit Sub
ErrorHandler:
  HandleError True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Form_Unload(Cancel As Integer)
646:   Set m_pApp = Nothing
647:   Set m_pCurrentFrame = Nothing
648:   Set m_pMap = Nothing
649:   Set m_pIndexLayer = Nothing
End Sub

Private Sub optExtent_Click(Index As Integer)
On Error GoTo ErrHand:
  Select Case Index
  Case 0  'Variable
656:     txtMargin.Enabled = True
657:     cmbMargin.Enabled = True
658:     txtFixed.Enabled = False
659:     cmbDataDriven.Enabled = False
660:     If txtMargin.Text = "" Then
661:       cmdNext.Enabled = False
662:     Else
663:       cmdNext.Enabled = True
664:     End If
  Case 1  'Fixed
666:     txtMargin.Enabled = False
667:     cmbMargin.Enabled = False
668:     txtFixed.Enabled = True
669:     cmbDataDriven.Enabled = False
670:     If txtFixed.Text = "" Then
671:       cmdNext.Enabled = False
672:     Else
673:       cmdNext.Enabled = True
674:     End If
  Case 2  'Data driven
676:     txtMargin.Enabled = False
677:     cmbMargin.Enabled = False
678:     txtFixed.Enabled = False
679:     cmbDataDriven.Enabled = True
680:     cmdNext.Enabled = True
681:   End Select

  Exit Sub
ErrHand:
685:   MsgBox "optExtent_Click - " & Err.Description
End Sub

Private Sub txtFixed_KeyUp(KeyCode As Integer, Shift As Integer)
689:   If Not IsNumeric(txtFixed.Text) Then
690:     txtFixed.Text = ""
691:   End If
692:   If txtFixed.Text <> "" Then
693:     cmdNext.Enabled = True
694:   End If
End Sub

Private Sub txtMargin_KeyUp(KeyCode As Integer, Shift As Integer)
698:   If Not IsNumeric(txtMargin.Text) Then
699:     txtMargin.Text = ""
700:   End If
701:   If txtMargin.Text <> "" Then
702:     cmdNext.Enabled = True
703:   End If
End Sub

Private Function FeaturesInTile(pFeature As IFeature, pMap As IMap) As Boolean
'Routine for determining whether the specified tile feature (pFeature) should
'be suppressed.  Tiles are suppressed when there are no features from the checked
'layers in them.
On Error GoTo ErrHand:
  Dim lLoop As Long, pFeatLayer As IFeatureLayer, pSpatial As ISpatialFilter
  Dim pFeatCursor As IFeatureCursor, pSearchFeat As IFeature
  
714:   FeaturesInTile = False
  
716:   Set pSpatial = New SpatialFilter
717:   pSpatial.SpatialRel = esriSpatialRelIntersects
718:   Set pSpatial.Geometry = pFeature.Shape
719:   For lLoop = 0 To lstSuppressTiles.ListCount - 1
720:     If lstSuppressTiles.Selected(lLoop) Then
721:       Set pFeatLayer = FindLayer(lstSuppressTiles.List(lLoop), pMap)
722:       pSpatial.GeometryField = pFeatLayer.FeatureClass.shapeFieldName
723:       Set pFeatCursor = pFeatLayer.Search(pSpatial, True)
724:       Set pSearchFeat = pFeatCursor.NextFeature
725:       If Not pSearchFeat Is Nothing Then
726:         FeaturesInTile = True
        Exit Function
728:       End If
729:     End If
730:   Next lLoop

  Exit Function
  
ErrHand:
735:   MsgBox "FeaturesInTile - " & Err.Description
End Function

Private Sub txtNumbering_KeyUp(KeyCode As Integer, Shift As Integer)
739:   If Not IsNumeric(txtNumbering.Text) Then
740:     txtNumbering.Text = "1"
'  ElseIf CInt(txtNumbering.Text) < 0 Then
'    MsgBox "Can not use a number less than 0!!!"
'    txtNumbering.Text = "1"
744:   End If
End Sub
