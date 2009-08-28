VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DSMapBookTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Copyright 2008 ESRI
'
' All rights reserved under the copyright laws of the United States
' and applicable international laws, treaties, and conventions.
'
' You may freely redistribute and use this sample code, with or
' without modification, provided you include the original copyright
' notice and use restrictions.
'
' See use restrictions at <your ArcGIS install location>/developerkit/userestrictions.txt.
'




Option Explicit

Private m_pDoc As IMxDocument
Private m_pMap As IMap
Private m_bIsVisible As Boolean
Private m_ContextItem As Variant
Private m_SelectedItem As Variant
Private m_ShowLines As Boolean

Implements IContentsView

Private Sub Class_Initialize()
13:   m_bIsVisible = True
14:   'Load g_pFrmMapSeries
End Sub

Private Sub Class_Terminate()
18:   'Unload g_pFrmMapSeries
19:   Set g_pFrmMapSeries = Nothing
End Sub

Private Sub IContentsView_Activate(ByVal parentHWND As esriSystem.OLE_HANDLE, ByVal Document As IMxDocument)

  'From Al
25:   RemoveContextMenu g_pFrmMapSeries.tvwMapBook.hWnd
End Sub

Private Sub IContentsView_AddToSelectedItems(ByVal Item As Variant)

End Sub

Private Property Let IContentsView_ContextItem(ByVal RHS As Variant)
On Error GoTo ErrHand:
34:   m_ContextItem = RHS
  
  Exit Property
ErrHand:
38:   MsgBox "Let IContentsView_ContextItem - " & Err.Description
End Property

Private Property Get IContentsView_ContextItem() As Variant
On Error GoTo ErrHand:
43:   IContentsView_ContextItem = m_ContextItem
  
  Exit Property
ErrHand:
47:   MsgBox "Get IContentsView_ContextItem - " & Err.Description
End Property

Private Sub IContentsView_Deactivate()
51:   RemoveContextMenuSink g_pFrmMapSeries.tvwMapBook.hWnd
End Sub

Private Property Get IContentsView_hWnd() As esriSystem.OLE_HANDLE
55:   IContentsView_hWnd = g_pFrmMapSeries.tvwMapBook.hWnd
End Property

Private Property Get IContentsView_Name() As String
59:   IContentsView_Name = "Map Book"
End Property

Private Property Let IContentsView_ProcessEvents(ByVal RHS As Boolean)

End Property

Private Sub IContentsView_Refresh(ByVal Item As Variant)

End Sub

Private Sub IContentsView_RemoveFromSelectedItems(ByVal Item As Variant)

End Sub

Private Property Let IContentsView_SelectedItem(ByVal RHS As Variant)
On Error GoTo ErrHand:
76:   m_SelectedItem = RHS
  
  Exit Property
ErrHand:
80:   MsgBox "Let IContentsView_SelectedItem - " & Err.Description
End Property

Private Property Get IContentsView_SelectedItem() As Variant
On Error GoTo ErrHand:
85:   IContentsView_SelectedItem = m_SelectedItem
  
  Exit Property
ErrHand:
89:   MsgBox "Get IContentsView_SelectedItem - " & Err.Description
End Property

Private Property Let IContentsView_ShowLines(ByVal RHS As Boolean)
93:   m_ShowLines = RHS
End Property

Private Property Get IContentsView_ShowLines() As Boolean
97:   IContentsView_ShowLines = m_ShowLines
End Property

Private Property Let IContentsView_Visible(ByVal RHS As Boolean)
101:   m_bIsVisible = RHS
End Property

Private Property Get IContentsView_Visible() As Boolean
105:   IContentsView_Visible = m_bIsVisible
End Property

