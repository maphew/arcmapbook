Attribute VB_Name = "modGlobals"

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
28:   Const WM_CONTEXTMENU = &H7B
29:   If Msg <> WM_CONTEXTMENU Then
30:     NoContextMenuWindowProc = CallWindowProc(lContextmenuWindowProc, hwnd, Msg, wParam, lParam)
31:   End If
  
  Exit Function
ErrHand:
35:   MsgBox "NoContextMenuWindowProc - " & Err.Description
End Function
' This function starts the "NoContextMenuWindowProc" message loop
Public Sub RemoveContextMenu(lhWnd As Long)
On Error GoTo ErrHand:
40:   lContextmenuWindowProc = SetWindowLong(lhWnd, GWL_WNDPROC, AddressOf NoContextMenuWindowProc)
  
  Exit Sub
ErrHand:
44:   MsgBox "RemoveContextMenu - " & Err.Description
End Sub

Function TopMost(f As Form)
    Dim i As Integer
49:     Call SetWindowPos(f.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Function

Function NoTopMost(f As Form)
53:     Call SetWindowPos(f.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Function

Public Sub RemoveContextMenuSink(lhWnd As Long)
On Error GoTo ErrHand:
  Dim lngReturnValue As Long
59:   lngReturnValue = SetWindowLong(lhWnd, GWL_WNDPROC, lContextmenuWindowProc)
  
  Exit Sub
ErrHand:
63:   MsgBox "RemoveContextMenuSink - " & Err.Description
End Sub

Public Function FindDataFrame(pDoc As IMxDocument, sFrameName As String) As IMap
On Error GoTo ErrHand:
  Dim lLoop As Long, pMap As IMap
  
  'Find the data frame
71:   For lLoop = 0 To pDoc.Maps.count - 1
72:     If pDoc.Maps.Item(lLoop).Name = sFrameName Then
73:       Set pMap = pDoc.Maps.Item(lLoop)
74:       Exit For
75:     End If
76:   Next lLoop
77:   If Not pMap Is Nothing Then
78:     Set FindDataFrame = pMap
79:   End If

  Exit Function
ErrHand:
83:   MsgBox "FindDataFrame - " & Err.Description
End Function

Public Function FindLayer(sLayerName As String, pMap As IMap) As IFeatureLayer
' Routine for finding a layer based on a name and then returning that layer as
' a IFeatureLayer
On Error GoTo ErrHand:
  Dim lLoop As Integer
  Dim pFLayer As IFeatureLayer

93:   For lLoop = 0 To pMap.LayerCount - 1
94:     If TypeOf pMap.Layer(lLoop) Is ICompositeLayer Then
95:       Set pFLayer = FindCompositeLayer(pMap.Layer(lLoop), sLayerName, pMap)
96:       If Not pFLayer Is Nothing Then
97:         Set FindLayer = pFLayer
        Exit Function
99:       End If
100:     ElseIf TypeOf pMap.Layer(lLoop) Is IFeatureLayer Then
101:       Set pFLayer = pMap.Layer(lLoop)
102:       If UCase(pFLayer.Name) = UCase(sLayerName) Then
103:         Set FindLayer = pFLayer
        Exit Function
105:       End If
106:     End If
107:   Next lLoop
  
109:   Set FindLayer = Nothing
  
  Exit Function
  
ErrHand:
114:   MsgBox "FindLayer - " & Err.Description
End Function

Private Function FindCompositeLayer(pCompLayer As ICompositeLayer, sLayerName As String, pMap As IMap) As IFeatureLayer
On Error GoTo ErrHand:
  Dim lLoop As Long, pFeatLayer As IFeatureLayer
120:   For lLoop = 0 To pCompLayer.count - 1
121:     If TypeOf pCompLayer.Layer(lLoop) Is ICompositeLayer Then
122:       Set pFeatLayer = FindCompositeLayer(pCompLayer.Layer(lLoop), sLayerName, pMap)
123:       If Not pFeatLayer Is Nothing Then
124:         Set FindCompositeLayer = pFeatLayer
        Exit Function
126:       End If
127:     Else
128:       If TypeOf pCompLayer.Layer(lLoop) Is IFeatureLayer Then
129:         If UCase(pCompLayer.Layer(lLoop).Name) = UCase(sLayerName) Then
130:           Set FindCompositeLayer = pCompLayer.Layer(lLoop)
          Exit Function
132:         End If
133:       End If
134:     End If
135:   Next lLoop

  Exit Function
ErrHand:
139:   MsgBox "CompositeLayer - " & Err.Description
End Function

Public Function ParseOutPages(sPagesToPrint As String, pMapSeries As IDSMapSeries, bDisabled As Boolean) As Collection
On Error GoTo ErrHand:
  If Len(sPagesToPrint) = 0 Then Exit Function
  
  Dim NoSpaces() As String
  Dim sTextToSplit As String
  
      'Get rid of any spaces
150:       NoSpaces = Split(sPagesToPrint)
151:       sTextToSplit = Join(NoSpaces, "") 'joined with no spaces
      
  Dim aPages() As String
154:       aPages = Split(sTextToSplit, ",")
      
  Dim aPages2() As String
  
  Dim i As Long
  Dim sPage As String
  Dim lLength As Long
  Dim count As Long
  
  Dim DSPagesCollection As New Collection
  Dim lStart As Long, lEnd As Long, lPage As Long

166:   For i = 0 To UBound(aPages)
167:      aPages2 = Split(aPages(i), "-")
          
169:       If UBound(aPages2) = 1 Then
170:           lStart = CInt(aPages2(0))
171:               count = count + 1
172:           lEnd = CInt(aPages2(1))
              
174:           While lStart <> (lEnd + 1)
175:             If bDisabled Then
176:               If pMapSeries.Page(lStart - 1).EnablePage Then
177:                 DSPagesCollection.Add pMapSeries.Page(lStart - 1)
178:               End If
179:             Else
180:               DSPagesCollection.Add pMapSeries.Page(lStart - 1)
181:             End If
182:             lStart = lStart + 1
183:           Wend
184:       ElseIf UBound(aPages2) < 1 Then
185:           lPage = CInt(aPages2(0))
186:           If bDisabled Then
187:             If pMapSeries.Page(lPage - 1).EnablePage Then
188:               DSPagesCollection.Add pMapSeries.Page(lPage - 1)
189:             End If
190:           Else
191:             DSPagesCollection.Add pMapSeries.Page(lPage - 1)
192:           End If
193:       End If
194:   Next i
      
  If DSPagesCollection.count = 0 Then Exit Function
  
198:   Set ParseOutPages = DSPagesCollection
    
  Exit Function
ErrHand:
202:   MsgBox "ParseOutPages - " & Err.Description
End Function



