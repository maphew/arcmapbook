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
41:   Const WM_CONTEXTMENU = &H7B
42:   If Msg <> WM_CONTEXTMENU Then
43:     NoContextMenuWindowProc = CallWindowProc(lContextmenuWindowProc, hwnd, Msg, wParam, lParam)
44:   End If
  
  Exit Function
ErrHand:
48:   MsgBox "NoContextMenuWindowProc - " & Err.Description
End Function
' This function starts the "NoContextMenuWindowProc" message loop
Public Sub RemoveContextMenu(lhWnd As Long)
On Error GoTo ErrHand:
53:   lContextmenuWindowProc = SetWindowLong(lhWnd, GWL_WNDPROC, AddressOf NoContextMenuWindowProc)
  
  Exit Sub
ErrHand:
57:   MsgBox "RemoveContextMenu - " & Err.Description
End Sub

Function TopMost(f As Form)
    Dim i As Integer
62:     Call SetWindowPos(f.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Function

Function NoTopMost(f As Form)
66:     Call SetWindowPos(f.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Function

Public Sub RemoveContextMenuSink(lhWnd As Long)
On Error GoTo ErrHand:
  Dim lngReturnValue As Long
72:   lngReturnValue = SetWindowLong(lhWnd, GWL_WNDPROC, lContextmenuWindowProc)
  
  Exit Sub
ErrHand:
76:   MsgBox "RemoveContextMenuSink - " & Err.Description
End Sub

Public Function FindDataFrame(pDoc As IMxDocument, sFrameName As String) As IMap
On Error GoTo ErrHand:
  Dim lLoop As Long, pMap As IMap
  
  'Find the data frame
84:   For lLoop = 0 To pDoc.Maps.count - 1
85:     If pDoc.Maps.Item(lLoop).Name = sFrameName Then
86:       Set pMap = pDoc.Maps.Item(lLoop)
87:       Exit For
88:     End If
89:   Next lLoop
90:   If Not pMap Is Nothing Then
91:     Set FindDataFrame = pMap
92:   End If

  Exit Function
ErrHand:
96:   MsgBox "FindDataFrame - " & Err.Description
End Function

Public Function FindLayer(sLayerName As String, pMap As IMap) As IFeatureLayer
' Routine for finding a layer based on a name and then returning that layer as
' a IFeatureLayer
On Error GoTo ErrHand:
  Dim lLoop As Integer
  Dim pFLayer As IFeatureLayer

106:   For lLoop = 0 To pMap.LayerCount - 1
107:     If TypeOf pMap.Layer(lLoop) Is ICompositeLayer Then
108:       Set pFLayer = FindCompositeLayer(pMap.Layer(lLoop), sLayerName, pMap)
109:       If Not pFLayer Is Nothing Then
110:         Set FindLayer = pFLayer
        Exit Function
112:       End If
113:     ElseIf TypeOf pMap.Layer(lLoop) Is IFeatureLayer Then
114:       Set pFLayer = pMap.Layer(lLoop)
115:       If UCase(pFLayer.Name) = UCase(sLayerName) Then
116:         Set FindLayer = pFLayer
        Exit Function
118:       End If
119:     End If
120:   Next lLoop
  
122:   Set FindLayer = Nothing
  
  Exit Function
  
ErrHand:
127:   MsgBox "FindLayer - " & Err.Description
End Function

Private Function FindCompositeLayer(pCompLayer As ICompositeLayer, sLayerName As String, pMap As IMap) As IFeatureLayer
On Error GoTo ErrHand:
  Dim lLoop As Long, pFeatLayer As IFeatureLayer
133:   For lLoop = 0 To pCompLayer.count - 1
134:     If TypeOf pCompLayer.Layer(lLoop) Is ICompositeLayer Then
135:       Set pFeatLayer = FindCompositeLayer(pCompLayer.Layer(lLoop), sLayerName, pMap)
136:       If Not pFeatLayer Is Nothing Then
137:         Set FindCompositeLayer = pFeatLayer
        Exit Function
139:       End If
140:     Else
141:       If TypeOf pCompLayer.Layer(lLoop) Is IFeatureLayer Then
142:         If UCase(pCompLayer.Layer(lLoop).Name) = UCase(sLayerName) Then
143:           Set FindCompositeLayer = pCompLayer.Layer(lLoop)
          Exit Function
145:         End If
146:       End If
147:     End If
148:   Next lLoop

  Exit Function
ErrHand:
152:   MsgBox "CompositeLayer - " & Err.Description
End Function

Public Function ParseOutPages(sPagesToPrint As String, pMapSeries As IDSMapSeries, bDisabled As Boolean) As Collection
On Error GoTo ErrHand:
  If Len(sPagesToPrint) = 0 Then Exit Function
  
  Dim NoSpaces() As String
  Dim sTextToSplit As String
  Dim pSeriesProps As IDSMapSeriesProps, lAdjustment As Long
  
      'Get rid of any spaces
164:       NoSpaces = Split(sPagesToPrint)
165:       sTextToSplit = Join(NoSpaces, "") 'joined with no spaces
      
  Dim aPages() As String
168:       aPages = Split(sTextToSplit, ",")
      
  Dim aPages2() As String
  
  Dim i As Long
  Dim sPage As String
  Dim lLength As Long
  Dim count As Long
  
  Dim DSPagesCollection As New Collection
  Dim lStart As Long, lEnd As Long, lPage As Long
  
  'Get the series properties so we can extract the starting page number
181:   Set pSeriesProps = pMapSeries
182:   lAdjustment = pSeriesProps.StartNumber - 1

184:   For i = 0 To UBound(aPages)
185:      aPages2 = Split(aPages(i), "-")
          
187:       If UBound(aPages2) = 1 Then
188:           lStart = CInt(aPages2(0)) - lAdjustment
189:               count = count + 1
190:           lEnd = CInt(aPages2(1)) - lAdjustment
              
192:           While lStart <> (lEnd + 1)
193:             If bDisabled Then
194:               If pMapSeries.Page(lStart - 1).EnablePage Then
195:                 DSPagesCollection.Add pMapSeries.Page(lStart - 1)
196:               End If
197:             Else
198:               DSPagesCollection.Add pMapSeries.Page(lStart - 1)
199:             End If
200:             lStart = lStart + 1
201:           Wend
202:       ElseIf UBound(aPages2) < 1 Then
203:           lPage = CInt(aPages2(0)) - lAdjustment
204:           If bDisabled Then
205:             If pMapSeries.Page(lPage - 1).EnablePage Then
206:               DSPagesCollection.Add pMapSeries.Page(lPage - 1)
207:             End If
208:           Else
209:             DSPagesCollection.Add pMapSeries.Page(lPage - 1)
210:           End If
211:       End If
212:   Next i
      
  If DSPagesCollection.count = 0 Then Exit Function
  
216:   Set ParseOutPages = DSPagesCollection
    
  Exit Function
ErrHand:
220:   MsgBox "ParseOutPages - " & Err.Description
End Function



