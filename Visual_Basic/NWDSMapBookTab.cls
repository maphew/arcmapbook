VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NWDSMapBookTab"
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

Private m_pDoc As IMxDocument
Private m_pMap As IMap
Private m_bIsVisible As Boolean
Private m_ContextItem As Variant
Private m_SelectedItem As Variant
Private m_ShowLines As Boolean

Implements IContentsView

Private Sub Class_Initialize()
  m_bIsVisible = True
  Load g_pFrmMapSeries
End Sub

Private Sub Class_Terminate()
  Unload g_pFrmMapSeries
  Set g_pFrmMapSeries = Nothing
End Sub

Private Sub IContentsView_Activate(ByVal parentHWND As esriSystem.OLE_HANDLE, ByVal Document As IMxDocument)

  'From Al
  RemoveContextMenu g_pFrmMapSeries.tvwMapBook.hwnd
End Sub

Private Sub IContentsView_AddToSelectedItems(ByVal Item As Variant)

End Sub

Private Property Let IContentsView_ContextItem(ByVal RHS As Variant)
On Error GoTo ErrHand:
  m_ContextItem = RHS
  
  Exit Property
ErrHand:
  MsgBox "Let IContentsView_ContextItem - " & Err.Description
End Property

Private Property Get IContentsView_ContextItem() As Variant
On Error GoTo ErrHand:
  IContentsView_ContextItem = m_ContextItem
  
  Exit Property
ErrHand:
  MsgBox "Get IContentsView_ContextItem - " & Err.Description
End Property

Private Sub IContentsView_Deactivate()
  RemoveContextMenuSink g_pFrmMapSeries.tvwMapBook.hwnd
End Sub

Private Property Get IContentsView_hWnd() As esriSystem.OLE_HANDLE
  IContentsView_hWnd = g_pFrmMapSeries.tvwMapBook.hwnd
End Property

Private Property Get IContentsView_Name() As String
  IContentsView_Name = "NW MapBook"
End Property

Private Property Let IContentsView_ProcessEvents(ByVal RHS As Boolean)

End Property

Private Sub IContentsView_Refresh(ByVal Item As Variant)

End Sub

Private Sub IContentsView_RemoveFromSelectedItems(ByVal Item As Variant)

End Sub

Private Property Let IContentsView_SelectedItem(ByVal RHS As Variant)
On Error GoTo ErrHand:
  m_SelectedItem = RHS
  
  Exit Property
ErrHand:
  MsgBox "Let IContentsView_SelectedItem - " & Err.Description
End Property

Private Property Get IContentsView_SelectedItem() As Variant
On Error GoTo ErrHand:
  IContentsView_SelectedItem = m_SelectedItem
  
  Exit Property
ErrHand:
  MsgBox "Get IContentsView_SelectedItem - " & Err.Description
End Property

Private Property Let IContentsView_ShowLines(ByVal RHS As Boolean)
  m_ShowLines = RHS
End Property

Private Property Get IContentsView_ShowLines() As Boolean
  IContentsView_ShowLines = m_ShowLines
End Property

Private Property Let IContentsView_Visible(ByVal RHS As Boolean)
  m_bIsVisible = RHS
End Property

Private Property Get IContentsView_Visible() As Boolean
  IContentsView_Visible = m_bIsVisible
End Property

