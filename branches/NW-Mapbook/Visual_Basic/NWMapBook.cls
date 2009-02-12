VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NWMapBook"
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

Implements INWDSMapBook
Implements IPersistVariant

Private m_ContentColl As Collection
Private m_bEnableBook As Boolean

Private Sub Class_Initialize()
38:   Set m_ContentColl = New Collection
End Sub

Private Sub Class_Terminate()
42:   Set m_ContentColl = Nothing
End Sub

Private Sub INWDSMapBook_AddContent(ByVal Content As Object)
46:   m_ContentColl.Add Content
End Sub

Private Property Get INWDSMapBook_ContentCount() As Long
50:   INWDSMapBook_ContentCount = m_ContentColl.Count
End Property

Private Property Get INWDSMapBook_ContentItem(Index As Long) As Object
54:   If Index > -1 And Index < m_ContentColl.Count Then
55:     Set INWDSMapBook_ContentItem = m_ContentColl.Item(Index + 1)
56:   Else
57:     Set INWDSMapBook_ContentItem = Nothing
58:   End If
End Property

Private Property Let INWDSMapBook_EnableBook(ByVal RHS As Boolean)
62:   m_bEnableBook = RHS
End Property

Private Property Get INWDSMapBook_EnableBook() As Boolean
66:   INWDSMapBook_EnableBook = m_bEnableBook
End Property

Private Sub INWDSMapBook_RemoveContent(Index As Long)
70:   If Index > -1 And Index < m_ContentColl.Count Then
71:     m_ContentColl.Remove Index + 1
72:   End If
End Sub

Private Property Get IPersistVariant_ID() As esriSystem.IUID
  Dim id As New UID
77:   id = "NWMapBookPrj.NWMapBook"
78:   Set IPersistVariant_ID = id
End Property

Private Sub IPersistVariant_Load(ByVal Stream As esriSystem.IVariantStream)
'Load the persisted parameters of the renderer
On Error GoTo ErrHand:
  Dim lLoop As Long, lCount As Long, pMapSeries As INWDSMapSeries
    
86:   m_bEnableBook = Stream.Read
87:   lCount = Stream.Read
88:   Set m_ContentColl = New Collection
89:   If lCount > 0 Then
90:     For lLoop = 1 To lCount
91:       Set pMapSeries = Stream.Read
92:       m_ContentColl.Add pMapSeries
93:     Next lLoop
94:   End If
    
  Exit Sub
ErrHand:
98:   MsgBox "MapBook - IPersistVariant_Load - " & Erl & ": " & Err.Description & vbNewLine
End Sub

Private Sub IPersistVariant_Save(ByVal Stream As esriSystem.IVariantStream)
'Write it all out
On Error GoTo ErrHand:
  Dim lLoop As Long, pNWMapSeries As INWDSMapSeries
    
106:   Stream.Write m_bEnableBook
107:   Stream.Write m_ContentColl.Count
108:   If m_ContentColl.Count > 0 Then
111:     For lLoop = 1 To m_ContentColl.Count
113:       Set pNWMapSeries = m_ContentColl.Item(lLoop)
115:       Stream.Write pNWMapSeries
116:     Next lLoop
117:   End If
    
  Exit Sub
ErrHand:
121:   MsgBox "MapBook - IPersistVariant_Save - " & Erl & ": " & Err.Description
End Sub
