VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DSMapBook"
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

Implements IDSMapBook
Implements IPersistVariant

Private m_ContentColl As Collection
Private m_bEnableBook As Boolean

Private Sub Class_Initialize()
10:   Set m_ContentColl = New Collection
End Sub

Private Sub Class_Terminate()
14:   Set m_ContentColl = Nothing
End Sub

Private Sub IDSMapBook_AddContent(ByVal Content As Object)
18:   m_ContentColl.Add Content
End Sub

Private Property Get IDSMapBook_ContentCount() As Long
22:   IDSMapBook_ContentCount = m_ContentColl.Count
End Property

Private Property Get IDSMapBook_ContentItem(Index As Long) As Object
26:   If Index > -1 And Index < m_ContentColl.Count Then
27:     Set IDSMapBook_ContentItem = m_ContentColl.Item(Index + 1)
28:   Else
29:     Set IDSMapBook_ContentItem = Nothing
30:   End If
End Property

Private Property Let IDSMapBook_EnableBook(ByVal RHS As Boolean)
34:   m_bEnableBook = RHS
End Property

Private Property Get IDSMapBook_EnableBook() As Boolean
38:   IDSMapBook_EnableBook = m_bEnableBook
End Property

Private Sub IDSMapBook_RemoveContent(Index As Long)
42:   If Index > -1 And Index < m_ContentColl.Count Then
43:     m_ContentColl.Remove Index + 1
44:   End If
End Sub

Private Property Get IPersistVariant_ID() As esriSystem.IUID
  Dim id As New UID
49:   id = "DSMapBookPrj.DSMapBook"
50:   Set IPersistVariant_ID = id
End Property

Private Sub IPersistVariant_Load(ByVal Stream As esriSystem.IVariantStream)
'Load the persisted parameters of the renderer
On Error GoTo ErrHand:
  Dim lLoop As Long, lCount As Long, pMapSeries As IDSMapSeries
    
58:   m_bEnableBook = Stream.Read
59:   lCount = Stream.Read
60:   Set m_ContentColl = New Collection
61:   If lCount > 0 Then
62:     For lLoop = 1 To lCount
63:       Set pMapSeries = Stream.Read
64:       m_ContentColl.Add pMapSeries
65:     Next lLoop
66:   End If
    
  Exit Sub
ErrHand:
70:   MsgBox "MapBook - IPersistVariant_Load - " & Err.Description
End Sub

Private Sub IPersistVariant_Save(ByVal Stream As esriSystem.IVariantStream)
'Write it all out
On Error GoTo ErrHand:
  Dim lLoop As Long
    
78:   Stream.Write m_bEnableBook
79:   Stream.Write m_ContentColl.Count
80:   If m_ContentColl.Count > 0 Then
81:     For lLoop = 1 To m_ContentColl.Count
82:       Stream.Write m_ContentColl.Item(lLoop)
83:     Next lLoop
84:   End If
    
  Exit Sub
ErrHand:
88:   MsgBox "MapBook - IPersistVariant_Save - " & Err.Description
End Sub
