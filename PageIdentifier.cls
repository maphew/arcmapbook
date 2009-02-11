VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PageIdentifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

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

Implements ICommand
Implements ITool

Dim m_pApp As IApplication
Dim m_sName As String

Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
22:   ICommand_Bitmap = frmResources.picIdentifier.Picture.handle
End Property

Private Property Get ICommand_Caption() As String
26:   ICommand_Caption = "Add Identifier Frame"
End Property

Private Property Get ICommand_Category() As String
30:   ICommand_Category = "Developer Samples"
End Property

Private Property Get ICommand_Checked() As Boolean
34:   ICommand_Checked = False
End Property

Private Property Get ICommand_Enabled() As Boolean
38:   ICommand_Enabled = True
End Property

Private Property Get ICommand_HelpContextID() As Long

End Property

Private Property Get ICommand_HelpFile() As String

End Property

Private Property Get ICommand_Message() As String
50:   ICommand_Message = "Identifier Frame"
End Property

Private Property Get ICommand_Name() As String
54:   ICommand_Name = "DSMapBookUIPrj.PageIdentifier"
End Property

Private Sub ICommand_OnClick()
    'Get Identifer type
59:   frmPageIdentifier.Show vbModal
60:   If frmPageIdentifier.optIdentifier(0).value Then
61:     m_sName = "Local Indicator"
62:   Else
63:     m_sName = "Global Indicator"
64:   End If
65:   Unload frmPageIdentifier
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
69:   Set m_pApp = hook
End Sub

Private Property Get ICommand_Tooltip() As String
73:   ICommand_Tooltip = "Add Identifier Frame"
End Property

Private Property Get ITool_Cursor() As esriSystem.OLE_HANDLE
On Error GoTo ErrHand:
78:   ITool_Cursor = frmResources.imlIcons.ListImages(1).Picture
  
  Exit Property
ErrHand:
82:   MsgBox "ITool_Cursor - " & Err.Description
End Property

Private Function ITool_Deactivate() As Boolean
86:   ITool_Deactivate = True
End Function

Private Function ITool_OnContextMenu(ByVal X As Long, ByVal Y As Long) As Boolean

End Function

Private Sub ITool_OnDblClick()

End Sub

Private Sub ITool_OnKeyDown(ByVal KeyCode As Long, ByVal Shift As Long)

End Sub

Private Sub ITool_OnKeyUp(ByVal KeyCode As Long, ByVal Shift As Long)

End Sub

Private Sub ITool_OnMouseDown(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
On Error GoTo ErrHand:
  Dim pGraphicsContainer As IGraphicsContainer, pLineSym2 As ISimpleLineSymbol
  Dim pElement As IElement, pMxApp As IMxApplication
  Dim rMapFrame As IMapFrame, pFeatLayer As IFeatureLayer
  Dim pMap As IMap, pGridLayer As IFeatureLayer
  Dim pColor2 As IColor, pColor3 As IColor
  Dim pRubberBand As IRubberBand, pScreenDisplay As IScreenDisplay
  Dim pGeometry As IGeometry, lLoop As Long, pFeatSel As IFeatureSelection
  Dim pMxDoc As IMxDocument, pLayer As ILayer, pActive As IActiveView
  Dim pRend As ISimpleRenderer, pColor As IRgbColor, pFill As ISimpleFillSymbol
  Dim pLineSym As ISimpleLineSymbol, pGeoFeatLayer As IGeoFeatureLayer
  Dim pMapBook As IDSMapBook
  Dim pSeriesProps As IDSMapSeriesProps, pFill2 As ISimpleFillSymbol
    
120:   Set pMxApp = m_pApp
121:   Set pMxDoc = m_pApp.Document
122:   Set pGraphicsContainer = pMxDoc.PageLayout
123:   Set pRubberBand = New RubberEnvelope
124:   Set pScreenDisplay = pMxApp.Display
125:   Set pGeometry = pRubberBand.TrackNew(pScreenDisplay, Nothing)
  
  'Get the index layer from the current Map Series
128:   Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
131:   If pMapBook.ContentCount = 0 Then
132:     MsgBox "You need to create a Map Series before adding a Page Identifier!!!"
    Exit Sub
134:   End If
135:   Set pSeriesProps = pMapBook.ContentItem(0)
  'Find the data frame
137:   Set pMap = FindDataFrame(pMxDoc, pSeriesProps.DataFrameName)
138:   If pMap Is Nothing Then
139:     MsgBox "Could not find map in PageIdentifier_OnMouseDown routine!!!"
    Exit Sub
141:   End If
  
  'Find the Index layer
144:   Set pGridLayer = FindLayer(pSeriesProps.IndexLayerName, pMap)
145:   If pGridLayer Is Nothing Then
146:     MsgBox "Could not find index layer in PageIdentifier_OnMouseDown routine!!!"
    Exit Sub
148:   End If
  
  'Create a new map and layer
151:   Set pMap = New Map
152:   pMap.Name = m_sName
153:   Set pFeatLayer = New FeatureLayer
154:   Set pRend = New SimpleRenderer
155:   Set pColor = New RgbColor    'Fill Outline symbol
156:   Set pColor2 = New RgbColor   'Fill Symbol
157:   Set pColor3 = New RgbColor   'Selection Fill symbol
158:   pColor.RGB = RGB(0, 0, 0)
159:   pColor2.NullColor = True
160:   pColor3.RGB = RGB(180, 180, 180)
161:   Set pFill = New SimpleFillSymbol
162:   Set pFill2 = New SimpleFillSymbol
163:   Set pLineSym = New SimpleLineSymbol
164:   Set pLineSym2 = New SimpleLineSymbol
165:   pLineSym.Color = pColor
166:   pLineSym.Width = 1
167:   pFill.Color = pColor2
168:   pFill.Outline = pLineSym
169:   Set pRend.Symbol = pFill
170:   Set pFeatLayer.FeatureClass = pGridLayer.FeatureClass
171:   pFeatLayer.Name = "Identifier Layer"
172:   Set pGeoFeatLayer = pFeatLayer
173:   Set pGeoFeatLayer.Renderer = pRend
174:   pLineSym2.Color = pColor2
175:   pFill2.Color = pColor3
176:   pFill2.Outline = pLineSym
177:   pMap.AddLayer pFeatLayer
178:   Set pFeatSel = pFeatLayer
179:   Set pFeatSel.SelectionSymbol = pFill2
180:   pFeatSel.SetSelectionSymbol = True
    
  'Create a new MapFrame and associate map with it
  Dim pMapFrame As IMapFrame
184:   Set pMapFrame = New MapFrame
185:   Set pMapFrame.Map = pMap
  
  'Set the position of the new map frame
188:   Set pElement = pMapFrame
189:   pElement.Geometry = pGeometry
  
  'Add mapframe to the layout
192:   pGraphicsContainer.AddElement pMapFrame, 0
193:   Set pActive = pMap
194:   pActive.Refresh
  
  'Refresh ActiveView and TOC
197:   Set pActive = pMxDoc.FocusMap
198:   pActive.Refresh
199:   pMxDoc.CurrentContentsView.Refresh 0
  
  'Deactivate the tool
202:   Set m_pApp.CurrentTool = Nothing
  
  Exit Sub
ErrHand:
206:   MsgBox "PageIdentifier_OnMouseDown - " & Err.Description
End Sub

Private Sub ITool_OnMouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)

End Sub

Private Sub ITool_OnMouseUp(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)

End Sub

Private Sub ITool_Refresh(ByVal hdc As esriSystem.OLE_HANDLE)

End Sub
