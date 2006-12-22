Attribute VB_Name = "modGlobals"

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

Public g_pFrmMapSeries As frmMapSeries
Public g_bClipFlag As Boolean
Public g_bRotateFlag As Boolean
Public g_bLabelNeighbors As Boolean

' modFunctions.bas
' Als try to remove "What's this?" window.
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private lContextmenuWindowProc As Long
Declare Function FindWindow% Lib "user32" Alias "FindWindowA" _
    (ByVal lpclassname As Any, ByVal lpCaption As Any)

Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const HWND_NOTOPMOST = -2

Const GWL_WNDPROC = (-4)

Public Function NoContextMenuWindowProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo ErrHand:
57:   Const WM_CONTEXTMENU = &H7B
58:   If Msg <> WM_CONTEXTMENU Then
59:     NoContextMenuWindowProc = CallWindowProc(lContextmenuWindowProc, hwnd, Msg, wParam, lParam)
60:   End If
  
  Exit Function
ErrHand:
64:   MsgBox "NoContextMenuWindowProc - " & Err.Description
End Function
' This function starts the "NoContextMenuWindowProc" message loop
Public Sub RemoveContextMenu(lhWnd As Long)
On Error GoTo ErrHand:
69:   lContextmenuWindowProc = SetWindowLong(lhWnd, GWL_WNDPROC, AddressOf NoContextMenuWindowProc)
  
  Exit Sub
ErrHand:
73:   MsgBox "RemoveContextMenu - " & Err.Description
End Sub

Function TopMost(f As Form)
    Dim i As Integer
78:     Call SetWindowPos(f.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Function

Function NoTopMost(f As Form)
82:     Call SetWindowPos(f.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Function

Public Sub RemoveContextMenuSink(lhWnd As Long)
On Error GoTo ErrHand:
  Dim lngReturnValue As Long
88:   lngReturnValue = SetWindowLong(lhWnd, GWL_WNDPROC, lContextmenuWindowProc)
  
  Exit Sub
ErrHand:
92:   MsgBox "RemoveContextMenuSink - " & Err.Description
End Sub

Public Function FindDataFrame(pDoc As IMxDocument, sFrameName As String) As IMap
On Error GoTo ErrHand:
  Dim lLoop As Long, pMap As IMap
  
  'Find the data frame
100:   For lLoop = 0 To pDoc.Maps.count - 1
101:     If pDoc.Maps.Item(lLoop).Name = sFrameName Then
102:       Set pMap = pDoc.Maps.Item(lLoop)
103:       Exit For
104:     End If
105:   Next lLoop
106:   If Not pMap Is Nothing Then
107:     Set FindDataFrame = pMap
108:   End If

  Exit Function
ErrHand:
112:   MsgBox "FindDataFrame - " & Err.Description
End Function

Public Function FindLayer(sLayerName As String, pMap As IMap) As IFeatureLayer
' Routine for finding a layer based on a name and then returning that layer as
' a IFeatureLayer
On Error GoTo ErrHand:
  Dim lLoop As Integer
  Dim pFLayer As IFeatureLayer

122:   For lLoop = 0 To pMap.LayerCount - 1
123:     If TypeOf pMap.Layer(lLoop) Is ICompositeLayer Then
124:       Set pFLayer = FindCompositeLayer(pMap.Layer(lLoop), sLayerName, pMap)
125:       If Not pFLayer Is Nothing Then
126:         Set FindLayer = pFLayer
        Exit Function
128:       End If
129:     ElseIf TypeOf pMap.Layer(lLoop) Is IFeatureLayer Then
130:       Set pFLayer = pMap.Layer(lLoop)
131:       If UCase(pFLayer.Name) = UCase(sLayerName) Then
132:         Set FindLayer = pFLayer
        Exit Function
134:       End If
135:     End If
136:   Next lLoop
  
138:   Set FindLayer = Nothing
  
  Exit Function
  
ErrHand:
143:   MsgBox "FindLayer - " & Err.Description
End Function

Private Function FindCompositeLayer(pCompLayer As ICompositeLayer, sLayerName As String, pMap As IMap) As IFeatureLayer
On Error GoTo ErrHand:
  Dim lLoop As Long, pFeatLayer As IFeatureLayer
149:   For lLoop = 0 To pCompLayer.count - 1
150:     If TypeOf pCompLayer.Layer(lLoop) Is ICompositeLayer Then
151:       Set pFeatLayer = FindCompositeLayer(pCompLayer.Layer(lLoop), sLayerName, pMap)
152:       If Not pFeatLayer Is Nothing Then
153:         Set FindCompositeLayer = pFeatLayer
        Exit Function
155:       End If
156:     Else
157:       If TypeOf pCompLayer.Layer(lLoop) Is IFeatureLayer Then
158:         If UCase(pCompLayer.Layer(lLoop).Name) = UCase(sLayerName) Then
159:           Set FindCompositeLayer = pCompLayer.Layer(lLoop)
          Exit Function
161:         End If
162:       End If
163:     End If
164:   Next lLoop

  Exit Function
ErrHand:
168:   MsgBox "CompositeLayer - " & Err.Description
End Function

Public Function ParseOutPages(sPagesToPrint As String, pMapSeries As IDSMapSeries, bDisabled As Boolean) As Collection
On Error GoTo ErrHand:
  If Len(sPagesToPrint) = 0 Then Exit Function
  
  Dim NoSpaces() As String
  Dim sTextToSplit As String
  Dim pSeriesProps As IDSMapSeriesProps, lAdjustment As Long
  
      'Get rid of any spaces
180:       NoSpaces = Split(sPagesToPrint)
181:       sTextToSplit = Join(NoSpaces, "") 'joined with no spaces
      
  Dim aPages() As String
184:       aPages = Split(sTextToSplit, ",")
      
  Dim aPages2() As String
  
  Dim i As Long
  Dim sPage As String
  Dim lLength As Long
  Dim count As Long
  
  Dim DSPagesCollection As New Collection
  Dim lStart As Long, lEnd As Long, lPage As Long
  
  'Get the series properties so we can extract the starting page number
197:   Set pSeriesProps = pMapSeries
198:   lAdjustment = pSeriesProps.StartNumber - 1

200:   For i = 0 To UBound(aPages)
201:      aPages2 = Split(aPages(i), "-")
          
203:       If UBound(aPages2) = 1 Then
204:           lStart = CInt(aPages2(0)) - lAdjustment
205:               count = count + 1
206:           lEnd = CInt(aPages2(1)) - lAdjustment
              
208:           While lStart <> (lEnd + 1)
209:             If bDisabled Then
210:               If pMapSeries.Page(lStart - 1).EnablePage Then
211:                 DSPagesCollection.Add pMapSeries.Page(lStart - 1)
212:               End If
213:             Else
214:               DSPagesCollection.Add pMapSeries.Page(lStart - 1)
215:             End If
216:             lStart = lStart + 1
217:           Wend
218:       ElseIf UBound(aPages2) < 1 Then
219:           lPage = CInt(aPages2(0)) - lAdjustment
220:           If bDisabled Then
221:             If pMapSeries.Page(lPage - 1).EnablePage Then
222:               DSPagesCollection.Add pMapSeries.Page(lPage - 1)
223:             End If
224:           Else
225:             DSPagesCollection.Add pMapSeries.Page(lPage - 1)
226:           End If
227:       End If
228:   Next i
      
  If DSPagesCollection.count = 0 Then Exit Function
  
232:   Set ParseOutPages = DSPagesCollection
    
  Exit Function
ErrHand:
236:   MsgBox "ParseOutPages - " & Err.Description
End Function



