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
Option Explicit

Implements IDSMapBook
Implements IPersistVariant

Private m_ContentColl As Collection
Private m_bEnableBook As Boolean

Private Sub Class_Initialize()
  Set m_ContentColl = New Collection
End Sub

Private Sub Class_Terminate()
  Set m_ContentColl = Nothing
End Sub

Private Sub IDSMapBook_AddContent(ByVal Content As Object)
  m_ContentColl.Add Content
End Sub

Private Property Get IDSMapBook_ContentCount() As Long
  IDSMapBook_ContentCount = m_ContentColl.Count
End Property

Private Property Get IDSMapBook_ContentItem(Index As Long) As Object
  If Index > -1 And Index < m_ContentColl.Count Then
    Set IDSMapBook_ContentItem = m_ContentColl.Item(Index + 1)
  Else
    Set IDSMapBook_ContentItem = Nothing
  End If
End Property

Private Property Let IDSMapBook_EnableBook(ByVal RHS As Boolean)
  m_bEnableBook = RHS
End Property

Private Property Get IDSMapBook_EnableBook() As Boolean
  IDSMapBook_EnableBook = m_bEnableBook
End Property

Private Sub IDSMapBook_RemoveContent(Index As Long)
  If Index > -1 And Index < m_ContentColl.Count Then
    m_ContentColl.Remove Index + 1
  End If
End Sub

Private Property Get IPersistVariant_ID() As esriCore.IUID
  Dim id As New UID
  id = "DSMapBookPrj.DSMapBook"
  Set IPersistVariant_ID = id
End Property

Private Sub IPersistVariant_Load(ByVal Stream As esriCore.IVariantStream)
'Load the persisted parameters of the renderer
On Error GoTo ErrHand:
  Dim lLoop As Long, lCount As Long, pMapSeries As IDSMapSeries
    
  m_bEnableBook = Stream.Read
  lCount = Stream.Read
  Set m_ContentColl = New Collection
  If lCount > 0 Then
    For lLoop = 1 To lCount
      Set pMapSeries = Stream.Read
      m_ContentColl.Add pMapSeries
    Next lLoop
  End If
    
  Exit Sub
ErrHand:
  MsgBox "MapBook - IPersistVariant_Load - " & Err.Description
End Sub

Private Sub IPersistVariant_Save(ByVal Stream As esriCore.IVariantStream)
'Write it all out
On Error GoTo ErrHand:
  Dim lLoop As Long
    
  Stream.Write m_bEnableBook
  Stream.Write m_ContentColl.Count
  If m_ContentColl.Count > 0 Then
    For lLoop = 1 To m_ContentColl.Count
      Stream.Write m_ContentColl.Item(lLoop)
    Next lLoop
  End If
    
  Exit Sub
ErrHand:
  MsgBox "MapBook - IPersistVariant_Save - " & Err.Description
End Sub
