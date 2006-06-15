Attribute VB_Name = "modGlobals"
Option Explicit

Public g_pFrmMapSeries As frmMapSeries
Public g_bClipFlag As Boolean
Public g_bRotateFlag As Boolean
Public g_bLabelNeighbors As Boolean

' modFunctions.bas
' Als try to remove "What's this?" window.
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long
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

Public Function NoContextMenuWindowProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long
On Error GoTo ErrHand:
  Const WM_CONTEXTMENU = &H7B
  If Msg <> WM_CONTEXTMENU Then
    NoContextMenuWindowProc = CallWindowProc(lContextmenuWindowProc, hwnd, Msg, wParam, lparam)
  End If
  
  Exit Function
ErrHand:
  MsgBox "NoContextMenuWindowProc - " & Err.Description
End Function
' This function starts the "NoContextMenuWindowProc" message loop
Public Sub RemoveContextMenu(lhWnd As Long)
On Error GoTo ErrHand:
  lContextmenuWindowProc = SetWindowLong(lhWnd, GWL_WNDPROC, AddressOf NoContextMenuWindowProc)
  
  Exit Sub
ErrHand:
  MsgBox "RemoveContextMenu - " & Err.Description
End Sub

Function TopMost(f As Form)
    Dim i As Integer
    Call SetWindowPos(f.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Function

Function NoTopMost(f As Form)
    Call SetWindowPos(f.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Function

Public Sub RemoveContextMenuSink(lhWnd As Long)
On Error GoTo ErrHand:
  Dim lngReturnValue As Long
  lngReturnValue = SetWindowLong(lhWnd, GWL_WNDPROC, lContextmenuWindowProc)
  
  Exit Sub
ErrHand:
  MsgBox "RemoveContextMenuSink - " & Err.Description
End Sub

Public Function FindDataFrame(pDoc As IMxDocument, sFrameName As String) As IMap
On Error GoTo ErrHand:
  Dim lLoop As Long, pMap As IMap
  
  'Find the data frame
  For lLoop = 0 To pDoc.Maps.count - 1
    If pDoc.Maps.Item(lLoop).Name = sFrameName Then
      Set pMap = pDoc.Maps.Item(lLoop)
      Exit For
    End If
  Next lLoop
  If Not pMap Is Nothing Then
    Set FindDataFrame = pMap
  End If

  Exit Function
ErrHand:
  MsgBox "FindDataFrame - " & Err.Description
End Function

Public Function FindLayer(sLayerName As String, pMap As IMap) As IFeatureLayer
' Routine for finding a layer based on a name and then returning that layer as
' a IFeatureLayer
On Error GoTo ErrHand:
  Dim lLoop As Integer
  Dim pFLayer As IFeatureLayer

  For lLoop = 0 To pMap.LayerCount - 1
    If TypeOf pMap.Layer(lLoop) Is ICompositeLayer Then
      Set pFLayer = FindCompositeLayer(pMap.Layer(lLoop), sLayerName, pMap)
      If Not pFLayer Is Nothing Then
        Set FindLayer = pFLayer
        Exit Function
      End If
    ElseIf TypeOf pMap.Layer(lLoop) Is IFeatureLayer Then
      Set pFLayer = pMap.Layer(lLoop)
      If UCase(pFLayer.Name) = UCase(sLayerName) Then
        Set FindLayer = pFLayer
        Exit Function
      End If
    End If
  Next lLoop
  
  Set FindLayer = Nothing
  
  Exit Function
  
ErrHand:
  MsgBox "FindLayer - " & Err.Description
End Function

Private Function FindCompositeLayer(pCompLayer As ICompositeLayer, sLayerName As String, pMap As IMap) As IFeatureLayer
On Error GoTo ErrHand:
  Dim lLoop As Long, pFeatLayer As IFeatureLayer
  For lLoop = 0 To pCompLayer.count - 1
    If TypeOf pCompLayer.Layer(lLoop) Is ICompositeLayer Then
      Set pFeatLayer = FindCompositeLayer(pCompLayer.Layer(lLoop), sLayerName, pMap)
      If Not pFeatLayer Is Nothing Then
        Set FindCompositeLayer = pFeatLayer
        Exit Function
      End If
    Else
      If TypeOf pCompLayer.Layer(lLoop) Is IFeatureLayer Then
        If UCase(pCompLayer.Layer(lLoop).Name) = UCase(sLayerName) Then
          Set FindCompositeLayer = pCompLayer.Layer(lLoop)
          Exit Function
        End If
      End If
    End If
  Next lLoop

  Exit Function
ErrHand:
  MsgBox "CompositeLayer - " & Err.Description
End Function

Public Function ParseOutPages(sPagesToPrint As String, pMapSeries As IDSMapSeries, bDisabled As Boolean) As Collection
On Error GoTo ErrHand:
  If Len(sPagesToPrint) = 0 Then Exit Function
  
  Dim NoSpaces() As String
  Dim sTextToSplit As String
  
      'Get rid of any spaces
      NoSpaces = Split(sPagesToPrint)
      sTextToSplit = Join(NoSpaces, "") 'joined with no spaces
      
  Dim aPages() As String
      aPages = Split(sTextToSplit, ",")
      
  Dim aPages2() As String
  
  Dim i As Long
  Dim sPage As String
  Dim lLength As Long
  Dim count As Long
  
  Dim DSPagesCollection As New Collection
  Dim lStart As Long, lEnd As Long, lPage As Long

  For i = 0 To UBound(aPages)
     aPages2 = Split(aPages(i), "-")
          
      If UBound(aPages2) = 1 Then
          lStart = CInt(aPages2(0))
              count = count + 1
          lEnd = CInt(aPages2(1))
              
          While lStart <> (lEnd + 1)
            If bDisabled Then
              If pMapSeries.Page(lStart - 1).EnablePage Then
                DSPagesCollection.Add pMapSeries.Page(lStart - 1)
              End If
            Else
              DSPagesCollection.Add pMapSeries.Page(lStart - 1)
            End If
            lStart = lStart + 1
          Wend
      ElseIf UBound(aPages2) < 1 Then
          lPage = CInt(aPages2(0))
          If bDisabled Then
            If pMapSeries.Page(lPage - 1).EnablePage Then
              DSPagesCollection.Add pMapSeries.Page(lPage - 1)
            End If
          Else
            DSPagesCollection.Add pMapSeries.Page(lPage - 1)
          End If
      End If
  Next i
      
  If DSPagesCollection.count = 0 Then Exit Function
  
  Set ParseOutPages = DSPagesCollection
    
  Exit Function
ErrHand:
  MsgBox "ParseOutPages - " & Err.Description
End Function



