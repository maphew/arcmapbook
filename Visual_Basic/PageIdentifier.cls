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

Implements ICommand
Implements ITool

Dim m_pApp As IApplication
Dim m_sName As String

Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
38:   ICommand_Bitmap = frmResources.picIdentifier.Picture.Handle
End Property

Private Property Get ICommand_Caption() As String
42:   ICommand_Caption = "Add Identifier Frame"
End Property

Private Property Get ICommand_Category() As String
46:   ICommand_Category = "Developer Samples"
End Property

Private Property Get ICommand_Checked() As Boolean
50:   ICommand_Checked = False
End Property

Private Property Get ICommand_Enabled() As Boolean
54:   ICommand_Enabled = True
End Property

Private Property Get ICommand_HelpContextID() As Long

End Property

Private Property Get ICommand_HelpFile() As String

End Property

Private Property Get ICommand_Message() As String
66:   ICommand_Message = "Identifier Frame"
End Property

Private Property Get ICommand_Name() As String
70:   ICommand_Name = "NWMapBookUIPrj.PageIdentifier"
End Property

Private Sub ICommand_OnClick()
    'Get Identifer type
75:   frmPageIdentifier.Show vbModal
76:   If frmPageIdentifier.optIdentifier(0).Value Then
77:     m_sName = "Local Indicator"
78:   Else
79:     m_sName = "Global Indicator"
80:   End If
81:   Unload frmPageIdentifier
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
85:   Set m_pApp = hook
End Sub

Private Property Get ICommand_Tooltip() As String
89:   ICommand_Tooltip = "Add Identifier Frame"
End Property

Private Property Get ITool_Cursor() As esriSystem.OLE_HANDLE
On Error GoTo ErrHand:
94:   ITool_Cursor = frmResources.imlIcons.ListImages(1).Picture
  
  Exit Property
ErrHand:
98:   MsgBox "ITool_Cursor - " & Err.Description
End Property

Private Function ITool_Deactivate() As Boolean
102:   ITool_Deactivate = True
End Function

Private Function ITool_OnContextMenu(ByVal x As Long, ByVal y As Long) As Boolean

End Function

Private Sub ITool_OnDblClick()

End Sub

Private Sub ITool_OnKeyDown(ByVal KeyCode As Long, ByVal Shift As Long)

End Sub

Private Sub ITool_OnKeyUp(ByVal KeyCode As Long, ByVal Shift As Long)

End Sub

Private Sub ITool_OnMouseDown(ByVal button As Long, ByVal Shift As Long, ByVal x As Long, ByVal y As Long)
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
  Dim pMapBook As INWDSMapBook
  Dim pSeriesProps As INWDSMapSeriesProps, pFill2 As ISimpleFillSymbol
    
136:   Set pMxApp = m_pApp
137:   Set pMxDoc = m_pApp.Document
138:   Set pGraphicsContainer = pMxDoc.PageLayout
139:   Set pRubberBand = New RubberEnvelope
140:   Set pScreenDisplay = pMxApp.Display
141:   Set pGeometry = pRubberBand.TrackNew(pScreenDisplay, Nothing)
  
  'Get the index layer from the current Map Series
144:   Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
147:   If pMapBook.ContentCount = 0 Then
148:     MsgBox "You need to create a Map Series before adding a Page Identifier!!!"
    Exit Sub
150:   End If
151:   Set pSeriesProps = pMapBook.ContentItem(0)
  'Find the data frame
153:   Set pMap = FindDataFrame(pMxDoc, pSeriesProps.DataFrameName)
154:   If pMap Is Nothing Then
155:     MsgBox "Could not find map in PageIdentifier_OnMouseDown routine!!!"
    Exit Sub
157:   End If
  
  'Find the Index layer
160:   Set pGridLayer = FindLayer(pSeriesProps.IndexLayerName, pMap)
161:   If pGridLayer Is Nothing Then
162:     MsgBox "Could not find index layer in PageIdentifier_OnMouseDown routine!!!"
    Exit Sub
164:   End If
  
  'Create a new map and layer
167:   Set pMap = New Map
                                                  'if a data frame already exists with m_sName,
                                                  'then increment the number with the name.
  Dim pTempMap As IMap, lPageIdNum As Long, sPageIDName As String, i As Long
171:   sPageIDName = m_sName
172:   Set pTempMap = FindDataFrame(pMxDoc, sPageIDName)
173:   If pTempMap Is Nothing Then
174:     pMap.Name = sPageIDName
175:   Else
176:     For i = 2 To 99
177:       sPageIDName = m_sName & i
178:       Set pTempMap = FindDataFrame(pMxDoc, sPageIDName)
179:       If pTempMap Is Nothing Then
180:         pMap.Name = sPageIDName
181:         Exit For
182:       End If
183:     Next i
184:     If i >= 99 And (pTempMap Is Nothing) Then
185:       MsgBox "Only 99 PageIdentifier maps can be in the map display at one time." & vbNewLine _
           & "An additional page identifier will not be created." & vbNewLine
      Exit Sub
188:     End If
189:   End If
  
191:   Set pFeatLayer = New FeatureLayer
192:   Set pRend = New SimpleRenderer
193:   Set pColor = New RgbColor    'Fill Outline symbol
194:   Set pColor2 = New RgbColor   'Fill Symbol
195:   Set pColor3 = New RgbColor   'Selection Fill symbol
196:   pColor.RGB = RGB(0, 0, 0)
197:   pColor2.NullColor = True
198:   pColor3.RGB = RGB(180, 180, 180)
199:   Set pFill = New SimpleFillSymbol
200:   Set pFill2 = New SimpleFillSymbol
201:   Set pLineSym = New SimpleLineSymbol
202:   Set pLineSym2 = New SimpleLineSymbol
203:   pLineSym.Color = pColor
204:   pLineSym.Width = 1
205:   pFill.Color = pColor2
206:   pFill.Outline = pLineSym
207:   Set pRend.Symbol = pFill
208:   Set pFeatLayer.FeatureClass = pGridLayer.FeatureClass
209:   pFeatLayer.Name = "Identifier Layer"
210:   Set pGeoFeatLayer = pFeatLayer
211:   Set pGeoFeatLayer.Renderer = pRend
212:   pLineSym2.Color = pColor2
213:   pFill2.Color = pColor3
214:   pFill2.Outline = pLineSym
215:   pMap.AddLayer pFeatLayer
216:   Set pFeatSel = pFeatLayer
217:   Set pFeatSel.SelectionSymbol = pFill2
218:   pFeatSel.SetSelectionSymbol = True
    
  'Create a new MapFrame and associate map with it
  Dim pMapFrame As IMapFrame
222:   Set pMapFrame = New MapFrame
223:   Set pMapFrame.Map = pMap
  
  'Set the position of the new map frame
226:   Set pElement = pMapFrame
227:   pElement.Geometry = pGeometry
  
  'Add mapframe to the layout
230:   pGraphicsContainer.AddElement pMapFrame, 0
231:   Set pActive = pMap
232:   pActive.Refresh
  
  'Refresh ActiveView and TOC
235:   Set pActive = pMxDoc.FocusMap
236:   pActive.Refresh
237:   pMxDoc.CurrentContentsView.Refresh 0
  
  'Deactivate the tool
240:   Set m_pApp.CurrentTool = Nothing
  
  Exit Sub
ErrHand:
244:   MsgBox "PageIdentifier_OnMouseDown - " & Err.Description
End Sub

Private Sub ITool_OnMouseMove(ByVal button As Long, ByVal Shift As Long, ByVal x As Long, ByVal y As Long)

End Sub

Private Sub ITool_OnMouseUp(ByVal button As Long, ByVal Shift As Long, ByVal x As Long, ByVal y As Long)

End Sub

Private Sub ITool_Refresh(ByVal hdc As esriSystem.OLE_HANDLE)

End Sub
