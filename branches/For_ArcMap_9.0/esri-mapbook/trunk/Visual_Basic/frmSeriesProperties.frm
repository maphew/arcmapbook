VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
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
      Begin VB.Frame fraPage1 
         BorderStyle     =   0  'None
         Height          =   3375
         Left            =   60
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
            Caption         =   "Begin numbering tiles/pages at:"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   44
            Top             =   2640
            Width           =   2295
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
         Left            =   -74940
         TabIndex        =   1
         Top             =   420
         Width           =   6825
         Begin VB.Frame fraExtent 
            Caption         =   "Extent"
            Height          =   2865
            Left            =   90
            TabIndex        =   8
            Top             =   450
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
            Height          =   2865
            Left            =   3450
            TabIndex        =   2
            Top             =   450
            Width           =   3255
            Begin VB.CheckBox chkOptions 
               Caption         =   "Select tile when drawing?"
               Height          =   225
               Index           =   4
               Left            =   90
               TabIndex        =   45
               Top             =   2520
               Width           =   2865
            End
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
               Left            =   90
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
Private m_pSeriesOptions As IDSMapSeriesOptions
Private m_pSeriesOptions2 As IDSMapSeriesOptions2
Private m_pSeriesOptions3 As IDSMapSeriesOptions3 'Added 11/23/04
Private m_bWasSelecting As Boolean                           'Added 11/23/04
Private m_pTextSym As ISimpleTextSymbol

Private Sub chkOptions_Click(Index As Integer)
  Select Case Index
  Case 0  'Rotate
41:     If chkOptions(0).Value = 0 Then
42:       cmbRotateField.Enabled = False
43:     Else
44:       cmbRotateField.Enabled = True
45:     End If
  Case 1  'Clip to outline
47:     If chkOptions(1).Value = 0 Then
48:       chkOptions(3).Value = 0
49:       chkOptions(3).Enabled = False
50:     Else
51:       chkOptions(3).Enabled = True
52:     End If
  Case 2  'Label neighboring tiles
54:     If chkOptions(2).Value = 0 Then
55:       cmdLabelProps.Enabled = False
56:     Else
57:       cmdLabelProps.Enabled = True
58:     End If
59:   End Select
End Sub

Private Sub cmdCancel_Click()
63:   Unload Me
End Sub

Private Sub cmdLabelProps_Click()
On Error GoTo ErrHand:
  Dim bChanged As Boolean, pTextSymEditor As ITextSymbolEditor
69:   Set pTextSymEditor = New TextSymbolEditor
70:   bChanged = pTextSymEditor.EditTextSymbol(m_pTextSym, m_pApp.hwnd)
71:   Me.SetFocus
  
  Exit Sub
ErrHand:
75:   MsgBox "cmdLabelProps_Click - " & Err.Description
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrHand:
  Dim pDoc As IMxDocument, pActive As IActiveView
  Dim pFeatSel As IFeatureSelection, pMap As IMap, pSeriesProps As IDSMapSeriesProps
  
  'Apply updates (only the Options can be updated, so we only need to look at those)
  'Set the clip and rotate properties
  'Update 6/18/03 to support cross hatching of clip area
86:   If chkOptions(1).Value = 1 Then    'Clip
87:     If chkOptions(3).Value = 0 Then   'clip without cross hatch
      'Make sure we don't leave the clip element
89:       If m_pSeriesOptions2.ClipData = 2 Then RemoveClipElement m_pApp.Document
90:       m_pSeriesOptions2.ClipData = 1
91:     Else
92:       m_pSeriesOptions2.ClipData = 2
93:       Set pDoc = m_pApp.Document
94:       pDoc.FocusMap.ClipGeometry = Nothing
95:     End If
'    m_pSeriesOptions.ClipData = True
97:   Else
    'Make sure we don't leave the clip element
99:     If m_pSeriesOptions2.ClipData = 2 Then RemoveClipElement m_pApp.Document
100:     m_pSeriesOptions2.ClipData = 0
'    m_pSeriesOptions.ClipData = False
    'Make sure clipping is turned off for the data frame
103:     Set pDoc = m_pApp.Document
104:     pDoc.FocusMap.ClipGeometry = Nothing
105:   End If
  
107:   If chkOptions(0).Value = 1 Then     'Rotation
108:     If m_pSeriesOptions.RotateFrame = False Or m_pSeriesOptions.RotationField <> cmbRotateField.Text Then
109:       UpdatePageValues "ROTATION", cmbRotateField.Text
110:     End If
111:     m_pSeriesOptions.RotateFrame = True
112:     m_pSeriesOptions.RotationField = cmbRotateField.Text
113:   Else
114:     m_pSeriesOptions.RotateFrame = False
    'Make sure rotation is turned off for the data frame
116:     Set pDoc = m_pApp.Document
117:     Set pActive = pDoc.FocusMap
118:     If pActive.ScreenDisplay.DisplayTransformation.Rotation <> 0 Then
119:       pActive.ScreenDisplay.DisplayTransformation.Rotation = 0
120:       pActive.Refresh
121:     End If
122:   End If
123:   If chkOptions(2).Value = 1 Then    'Label Neighbors
124:     m_pSeriesOptions.LabelNeighbors = True
125:   Else
126:     m_pSeriesOptions.LabelNeighbors = False
127:     RemoveLabels pDoc
128:     g_bLabelNeighbors = False
129:   End If
130:   Set m_pSeriesOptions.LabelSymbol = m_pTextSym
  
132:   If chkOptions(4).Value = 1 Then  'Select tile when drawing
133:     m_pSeriesOptions3.SelectTile = True
134:   Else
135:     m_pSeriesOptions3.SelectTile = False
136:     If m_bWasSelecting Then   'If there were previously selecting tiles, then we need to clear the selection
137:       Set pSeriesProps = m_pSeriesOptions
138:       Set pMap = pActive
139:       Set pFeatSel = FindLayer(pSeriesProps.IndexLayerName, pMap)
140:       If Not pFeatSel Is Nothing Then
141:         pFeatSel.Clear
142:         pActive.PartialRefresh esriViewGeoSelection, Nothing, pActive.Extent
143:       End If
144:     End If
145:   End If
  
  'Set the extent properties
148:   If optExtent(0).Value Then         'Variable
149:     m_pSeriesOptions.ExtentType = 0
150:     If txtMargin.Text = "" Then
151:       m_pSeriesOptions.Margin = 0
152:     Else
153:       m_pSeriesOptions.Margin = CDbl(txtMargin.Text)
154:     End If
155:     m_pSeriesOptions.MarginType = cmbMargin.ListIndex
156:   ElseIf optExtent(1).Value Then    'Fixed
157:     m_pSeriesOptions.ExtentType = 1
158:     m_pSeriesOptions.FixedScale = txtFixed.Text
159:   Else                        'Data driven
160:     If m_pSeriesOptions.ExtentType <> 2 Or m_pSeriesOptions.RotationField <> cmbRotateField.Text Then
161:       UpdatePageValues "SCALE", cmbDataDriven.Text
162:     End If
163:     m_pSeriesOptions.ExtentType = 2
164:     m_pSeriesOptions.DataDrivenField = cmbDataDriven.Text
165:   End If
  
167:   Unload Me
  
  Exit Sub
  
ErrHand:
172:   MsgBox "cmdOK_Click - " & Err.Description
End Sub

Private Sub UpdatePageValues(sProperty As String, sFieldName As String)
On Error GoTo ErrHand:
  Dim lLoop As Long, pSeries As IDSMapSeries, pPage As IDSMapPage
  Dim pDoc As IMxDocument, pMap As IMap, pSeriesProps As IDSMapSeriesProps
  Dim pIndexLayer As IFeatureLayer, pDataset As IDataset, pWorkspace As IFeatureWorkspace
  Dim pQueryDef As IQueryDef, pCursor As ICursor, pRow As IRow, pColl As Collection
181:   Set pDoc = m_pApp.Document
182:   Set pSeries = m_pSeriesOptions
183:   Set pSeriesProps = pSeries
184:   Set pMap = FindDataFrame(pDoc, pSeriesProps.DataFrameName)
  If pMap Is Nothing Then Exit Sub
  
187:   Set pIndexLayer = FindLayer(pSeriesProps.IndexLayerName, pMap)
  If pIndexLayer Is Nothing Then Exit Sub
  
  'Loop through the features in the index layer creating a collection of the scales and tile names
191:   Set pDataset = pIndexLayer.FeatureClass
192:   Set pWorkspace = pDataset.Workspace
193:   Set pQueryDef = pWorkspace.CreateQueryDef
194:   pQueryDef.Tables = pDataset.Name
195:   pQueryDef.SubFields = sFieldName & "," & pSeriesProps.IndexFieldName
196:   Set pCursor = pQueryDef.Evaluate
197:   Set pColl = New Collection
198:   Set pRow = pCursor.NextRow
199:   Do While Not pRow Is Nothing
200:     If Not IsNull(pRow.Value(0)) And Not IsNull(pRow.Value(1)) Then
201:       pColl.Add pRow.Value(0), pRow.Value(1)
202:     End If
203:     Set pRow = pCursor.NextRow
204:   Loop
  
  'Now loop through the pages and try to find the corresponding tile name in the collection
  On Error GoTo ErrNoKey:
208:   For lLoop = 0 To pSeries.PageCount - 1
209:     Set pPage = pSeries.Page(lLoop)
210:     If sProperty = "ROTATION" Then
211:       pPage.PageRotation = pColl.Item(pPage.PageName)
212:     Else
213:       pPage.PageScale = pColl.Item(pPage.PageName)
214:     End If
215:   Next lLoop

  Exit Sub

ErrNoKey:
220:   Resume Next
ErrHand:
222:   MsgBox "UpdatePageValues - " & Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo ErrHand:
  Dim pMapBook As IDSMapBook
  Dim pSeriesProps As IDSMapSeriesProps
  Dim lLoop As Long
  'Check to see if a MapSeries already exists
231:   Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
234:   Set pSeriesProps = pMapBook.ContentItem(0)
235:   Set m_pSeriesOptions = pSeriesProps
236:   Set m_pSeriesOptions2 = m_pSeriesOptions
237:   Set m_pSeriesOptions3 = m_pSeriesOptions
  
  'Index Settings Tab
240:   cmbDetailFrame.Clear
241:   cmbDetailFrame.AddItem pSeriesProps.DataFrameName
242:   cmbDetailFrame.Text = pSeriesProps.DataFrameName
243:   cmbIndexLayer.Clear
244:   cmbIndexLayer.AddItem pSeriesProps.IndexLayerName
245:   cmbIndexLayer.Text = pSeriesProps.IndexLayerName
246:   cmbIndexField.Clear
247:   cmbIndexField.AddItem pSeriesProps.IndexFieldName
248:   cmbIndexField.Text = pSeriesProps.IndexFieldName
  
  'Tile Settings Tab
251:   optTiles(pSeriesProps.TileSelectionMethod) = True
252:   lstSuppressTiles.Clear
253:   If pSeriesProps.SuppressLayers Then
254:     chkSuppress.Value = 1
255:     For lLoop = 0 To pSeriesProps.SuppressLayerCount - 1
256:       lstSuppressTiles.AddItem pSeriesProps.SuppressLayer(lLoop)
257:       lstSuppressTiles.Selected(lLoop) = True
258:     Next lLoop
259:   Else
260:     chkSuppress.Value = 0
261:   End If
262:   txtNumbering.Text = CStr(pSeriesProps.StartNumber)  'Added 2/18/2004
  
  'Options tab
265:   PopulateFieldCombos
266:   cmbMargin.Clear
267:   cmbMargin.AddItem "percent"
268:   cmbMargin.AddItem "mapunits"
269:   cmbMargin.Text = "percent"
270:   optExtent(m_pSeriesOptions.ExtentType).Value = True
271:   cmdOK.Enabled = True
  Select Case m_pSeriesOptions.ExtentType
  Case 0
274:     txtMargin.Text = m_pSeriesOptions.Margin
275:     If m_pSeriesOptions.MarginType = 0 Then
276:       cmbMargin.Text = "percent"
277:     Else
278:       cmbMargin.Text = "mapunits"
279:     End If
  Case 1
281:     txtFixed.Text = m_pSeriesOptions.FixedScale
  Case 2
283:     cmbDataDriven.Text = m_pSeriesOptions.DataDrivenField
284:   End Select
285:   If m_pSeriesOptions.RotateFrame Then
286:     chkOptions(0).Value = 1
287:     cmbRotateField.Text = m_pSeriesOptions.RotationField
288:   Else
289:     chkOptions(0).Value = 0
290:   End If
  
  'Update 6/18/03 to support cross hatching of clip area
  Select Case m_pSeriesOptions2.ClipData
  Case 0   'No clipping
295:     chkOptions(1).Value = 0
296:     chkOptions(3).Value = 0
297:     chkOptions(3).Enabled = False
  Case 1   'Clip only
299:     chkOptions(1).Value = 1
300:     chkOptions(3).Value = 0
301:     chkOptions(3).Enabled = True
  Case 2   'Clip with cross hatch outside clip area
303:     chkOptions(1).Value = 1
304:     chkOptions(3).Value = 1
305:     chkOptions(3).Enabled = True
306:   End Select
'  If m_pSeriesOptions.ClipData Then
'    chkOptions(1).Value = 1
'  Else
'    chkOptions(1).Value = 0
'  End If

313:   If m_pSeriesOptions.LabelNeighbors Then
314:     chkOptions(2).Value = 1
315:     cmdLabelProps.Enabled = True
316:   Else
317:     chkOptions(2).Value = 0
318:     cmdLabelProps.Enabled = False
319:   End If
320:   Set m_pTextSym = m_pSeriesOptions.LabelSymbol
  
322:   If m_pSeriesOptions3.SelectTile Then  'Added 11/23/04
323:     chkOptions(4).Value = 1
324:     m_bWasSelecting = True
325:   Else
326:     chkOptions(4).Value = 0
327:     m_bWasSelecting = False
328:   End If
  
  'Make sure the wizard stays on top
331:   TopMost Me
  
  Exit Sub
ErrHand:
335:   MsgBox "frmSeriesProperties_Load - " & Err.Description
End Sub

Private Sub PopulateFieldCombos()
On Error GoTo ErrHand:
  Dim pIndexLayer As IFeatureLayer, pMap As IMap, lLoop As Long
  Dim pFields As IFields, pDoc As IMxDocument
  
343:   Set pDoc = m_pApp.Document
344:   Set pMap = FindDataFrame(pDoc, cmbDetailFrame.Text)
345:   If pMap Is Nothing Then
346:     MsgBox "Could not find detail frame!!!"
    Exit Sub
348:   End If
  
350:   Set pIndexLayer = FindLayer(cmbIndexLayer.Text, pMap)
351:   If pIndexLayer Is Nothing Then
352:     MsgBox "Could not find specified layer!!!"
    Exit Sub
354:   End If
  
  'Populate the index layer combos
357:   Set pFields = pIndexLayer.FeatureClass.Fields
358:   cmbDataDriven.Clear
359:   cmbRotateField.Clear
360:   For lLoop = 0 To pFields.FieldCount - 1
    Select Case pFields.Field(lLoop).Type
    Case esriFieldTypeDouble, esriFieldTypeSingle, esriFieldTypeInteger
363:       If UCase(pFields.Field(lLoop).Name) <> "SHAPE_LENGTH" And _
       UCase(pFields.Field(lLoop).Name) <> "SHAPE_AREA" Then
365:         cmbDataDriven.AddItem pFields.Field(lLoop).Name
366:         cmbRotateField.AddItem pFields.Field(lLoop).Name
367:       End If
368:     End Select
369:   Next lLoop
370:   If cmbDataDriven.ListCount > 0 Then
371:     cmbDataDriven.ListIndex = 0
372:     cmbRotateField.ListIndex = 0
373:     optExtent.Item(2).Enabled = True
374:     chkOptions(0).Enabled = True
375:   Else
376:     optExtent.Item(2).Enabled = False
377:     chkOptions(0).Enabled = False
378:   End If
  
  Exit Sub
  
ErrHand:
383:   MsgBox "PopulateFieldCombos - " & Err.Description
End Sub

Private Sub Form_Terminate()
387:   Set m_pApp = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
391:   Set m_pApp = Nothing
End Sub

Private Sub optExtent_Click(Index As Integer)
On Error GoTo ErrHand:
  Select Case Index
  Case 0  'Variable
398:     txtMargin.Enabled = True
399:     cmbMargin.Enabled = True
400:     txtFixed.Enabled = False
401:     cmbDataDriven.Enabled = False
402:     If txtMargin.Text = "" Then
403:       cmdOK.Enabled = False
404:     Else
405:       cmdOK.Enabled = True
406:     End If
  Case 1  'Fixed
408:     txtMargin.Enabled = False
409:     cmbMargin.Enabled = False
410:     txtFixed.Enabled = True
411:     cmbDataDriven.Enabled = False
412:     If txtFixed.Text = "" Then
413:       cmdOK.Enabled = False
414:     Else
415:       cmdOK.Enabled = True
416:     End If
  Case 2  'Data driven
418:     txtMargin.Enabled = False
419:     cmbMargin.Enabled = False
420:     txtFixed.Enabled = False
421:     cmbDataDriven.Enabled = True
422:     cmdOK.Enabled = True
423:   End Select

  Exit Sub
ErrHand:
427:   MsgBox "optExtent_Click - " & Err.Description
End Sub

Private Sub txtFixed_KeyUp(KeyCode As Integer, Shift As Integer)
431:   If Not IsNumeric(txtFixed.Text) Then
432:     txtFixed.Text = ""
433:   End If
434:   If txtFixed.Text <> "" Then
435:     cmdOK.Enabled = True
436:   End If
End Sub

Private Sub txtMargin_KeyUp(KeyCode As Integer, Shift As Integer)
440:   If Not IsNumeric(txtMargin.Text) Then
441:     txtMargin.Text = ""
442:   End If
443:   If txtMargin.Text <> "" Then
444:     cmdOK.Enabled = True
445:   End If
End Sub
