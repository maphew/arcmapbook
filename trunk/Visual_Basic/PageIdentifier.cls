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
10:   ICommand_Bitmap = frmResources.picIdentifier.Picture.Handle
End Property

Private Property Get ICommand_Caption() As String
14:   ICommand_Caption = "Add Identifier Frame"
End Property

Private Property Get ICommand_Category() As String
18:   ICommand_Category = "Developer Samples"
End Property

Private Property Get ICommand_Checked() As Boolean
22:   ICommand_Checked = False
End Property

Private Property Get ICommand_Enabled() As Boolean
26:   ICommand_Enabled = True
End Property

Private Property Get ICommand_HelpContextID() As Long

End Property

Private Property Get ICommand_HelpFile() As String

End Property

Private Property Get ICommand_Message() As String
38:   ICommand_Message = "Identifier Frame"
End Property

Private Property Get ICommand_Name() As String
42:   ICommand_Name = "DSMapBookUIPrj.PageIdentifier"
End Property

Private Sub ICommand_OnClick()
    'Get Identifer type
47:   frmPageIdentifier.Show vbModal
48:   If frmPageIdentifier.optIdentifier(0).Value Then
49:     m_sName = "Local Indicator"
50:   Else
51:     m_sName = "Global Indicator"
52:   End If
53:   Unload frmPageIdentifier
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
57:   Set m_pApp = hook
End Sub

Private Property Get ICommand_Tooltip() As String
61:   ICommand_Tooltip = "Add Identifier Frame"
End Property

Private Property Get ITool_Cursor() As esriSystem.OLE_HANDLE
On Error GoTo ErrHand:
66:   ITool_Cursor = frmResources.imlIcons.ListImages(1).Picture
  
  Exit Property
ErrHand:
70:   MsgBox "ITool_Cursor - " & Err.Description
End Property

Private Function ITool_Deactivate() As Boolean
74:   ITool_Deactivate = True
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
    
108:   Set pMxApp = m_pApp
109:   Set pMxDoc = m_pApp.Document
110:   Set pGraphicsContainer = pMxDoc.PageLayout
111:   Set pRubberBand = New RubberEnvelope
112:   Set pScreenDisplay = pMxApp.Display
113:   Set pGeometry = pRubberBand.TrackNew(pScreenDisplay, Nothing)
  
  'Get the index layer from the current Map Series
116:   Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
119:   If pMapBook.ContentCount = 0 Then
120:     MsgBox "You need to create a Map Series before adding a Page Identifier!!!"
    Exit Sub
122:   End If
123:   Set pSeriesProps = pMapBook.ContentItem(0)
  'Find the data frame
125:   Set pMap = FindDataFrame(pMxDoc, pSeriesProps.DataFrameName)
126:   If pMap Is Nothing Then
127:     MsgBox "Could not find map in PageIdentifier_OnMouseDown routine!!!"
    Exit Sub
129:   End If
  
  'Find the Index layer
132:   Set pGridLayer = FindLayer(pSeriesProps.IndexLayerName, pMap)
133:   If pGridLayer Is Nothing Then
134:     MsgBox "Could not find index layer in PageIdentifier_OnMouseDown routine!!!"
    Exit Sub
136:   End If
  
  'Create a new map and layer
139:   Set pMap = New Map
140:   pMap.Name = m_sName
141:   Set pFeatLayer = New FeatureLayer
142:   Set pRend = New SimpleRenderer
143:   Set pColor = New RgbColor    'Fill Outline symbol
144:   Set pColor2 = New RgbColor   'Fill Symbol
145:   Set pColor3 = New RgbColor   'Selection Fill symbol
146:   pColor.RGB = RGB(0, 0, 0)
147:   pColor2.NullColor = True
148:   pColor3.RGB = RGB(180, 180, 180)
149:   Set pFill = New SimpleFillSymbol
150:   Set pFill2 = New SimpleFillSymbol
151:   Set pLineSym = New SimpleLineSymbol
152:   Set pLineSym2 = New SimpleLineSymbol
153:   pLineSym.Color = pColor
154:   pLineSym.Width = 1
155:   pFill.Color = pColor2
156:   pFill.Outline = pLineSym
157:   Set pRend.Symbol = pFill
158:   Set pFeatLayer.FeatureClass = pGridLayer.FeatureClass
159:   pFeatLayer.Name = "Identifier Layer"
160:   Set pGeoFeatLayer = pFeatLayer
161:   Set pGeoFeatLayer.Renderer = pRend
162:   pLineSym2.Color = pColor2
163:   pFill2.Color = pColor3
164:   pFill2.Outline = pLineSym
165:   pMap.AddLayer pFeatLayer
166:   Set pFeatSel = pFeatLayer
167:   Set pFeatSel.SelectionSymbol = pFill2
168:   pFeatSel.SetSelectionSymbol = True
    
  'Create a new MapFrame and associate map with it
  Dim pMapFrame As IMapFrame
172:   Set pMapFrame = New MapFrame
173:   Set pMapFrame.Map = pMap
  
  'Set the position of the new map frame
176:   Set pElement = pMapFrame
177:   pElement.Geometry = pGeometry
  
  'Add mapframe to the layout
180:   pGraphicsContainer.AddElement pMapFrame, 0
181:   Set pActive = pMap
182:   pActive.Refresh
  
  'Refresh ActiveView and TOC
185:   Set pActive = pMxDoc.FocusMap
186:   pActive.Refresh
187:   pMxDoc.CurrentContentsView.Refresh 0
  
  'Deactivate the tool
190:   Set m_pApp.CurrentTool = Nothing
  
  Exit Sub
ErrHand:
194:   MsgBox "PageIdentifier_OnMouseDown - " & Err.Description
End Sub

Private Sub ITool_OnMouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)

End Sub

Private Sub ITool_OnMouseUp(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)

End Sub

Private Sub ITool_Refresh(ByVal hdc As esriSystem.OLE_HANDLE)

End Sub
