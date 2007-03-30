VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DSMapBookExt"
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

Implements IExtension
Implements IPersistVariant

Private m_pApp As IApplication
Private WithEvents m_pMxDoc As MxDocument
Attribute m_pMxDoc.VB_VarHelpID = -1
Private m_pMapBook As IDSMapBook

Private Property Get IExtension_Name() As String
23:   IExtension_Name = "DevSample_MapBook"
End Property

Private Sub IExtension_Shutdown()
27:   Set g_pFrmMapSeries = Nothing
28:   Set m_pMapBook = Nothing
29:   Set m_pMxDoc = Nothing
30:   Set m_pApp = Nothing
End Sub

Public Property Get MapBook() As IDSMapBook
34:   Set MapBook = m_pMapBook
End Property

Private Sub IExtension_Startup(initializationData As Variant)
On Error GoTo ErrHand:
  Dim pApp As IApplication
40:   Set g_pFrmMapSeries = New frmMapSeries
41:   Set g_pFrmMapSeries.m_pApp = initializationData
42:   Set m_pMapBook = New DSMapBook
43:   Set pApp = initializationData
44:   Set m_pApp = pApp
45:   Set m_pMxDoc = pApp.Document
46:   g_bClipFlag = False
47:   g_bRotateFlag = False
48:   g_bLabelNeighbors = False
  
  Exit Sub
ErrHand:
52:   MsgBox "DSMapBookExt_Startup - " & Erl & " - " & Err.Description
End Sub

Private Property Get IPersistVariant_ID() As IUID
On Error GoTo ErrHand:
  
  Dim pUID As New UID
59:   pUID = "DSMapBookUIPrj.DSMapBookExt"

61:   Set IPersistVariant_ID = pUID

63:   GoTo EndProc

ErrHand:
66:   MsgBox "DSMapBookExt_ID - " & Err.Description
  Exit Property
EndProc:
69:   Set pUID = Nothing
End Property

Private Sub IPersistVariant_Load(ByVal Stream As IVariantStream)
On Error GoTo ErrHand:
  Dim pNode As Node, pMapSeries As IDSMapSeries, lLoop As Long, sName As String
  Dim pPage As IDSMapPage
76:   Set m_pMapBook = Stream.Read
  
  'Make sure we have the tab form
  If g_pFrmMapSeries Is Nothing Then Exit Sub
  
  'Make sure our persisted map book has content
  If m_pMapBook.ContentCount = 0 Then Exit Sub
  
  'Put the content back on the form
85:   Set pMapSeries = m_pMapBook.ContentItem(0)
86:   With g_pFrmMapSeries.tvwMapBook
87:     Set pNode = .Nodes.Add("MapBook", tvwChild, "MapSeries", "Map Series", 3)
    
    'Now loop back through the list and add the tile names as nodes in the tree
90:     For lLoop = 0 To pMapSeries.PageCount - 1
91:       Set pPage = pMapSeries.Page(lLoop)
92:       sName = pPage.PageName
93:       If pPage.EnablePage Then
94:         Set pNode = .Nodes.Add("MapSeries", tvwChild, "a" & sName, pPage.PageNumber & " - " & sName, 5)
95:       Else
96:         Set pNode = .Nodes.Add("MapSeries", tvwChild, "a" & sName, pPage.PageNumber & " - " & sName, 6)
97:       End If
98:       pNode.Tag = lLoop
99:     Next lLoop
100:     .Nodes.Item("MapBook").Expanded = True
101:     .Nodes.Item("MapSeries").Expanded = True
102:   End With
  
  Exit Sub
ErrHand:
106:   MsgBox "DSMapBookExt_IPersistVariant_Load - " & Erl & " - " & Err.Description
End Sub

Private Sub IPersistVariant_Save(ByVal Stream As IVariantStream)
On Error GoTo ErrHand:
111:   Stream.Write m_pMapBook
  
  Exit Sub
ErrHand:
115:   MsgBox "DSMapBookExt_IPersistVariant_Save - " & Erl & " - " & Err.Description
End Sub

Private Function m_pMxDoc_ActiveViewChanged() As Boolean
On Error GoTo ErrHand:
'Check to see if the active view is being changed back to the data view after the Map Book code has
'set the clip shape.  The g_bClipFlag variable will tell us if the Map Book code has updated the
'Clip Shape.  If it has, then we want to clear the clip shape when the user switches back to a data view.
'If the clip shape was changed by some other method, then we don't want to update it.
  Dim pDoc As IMxDocument
  
126:   If g_bClipFlag Then
'    If pSeriesOpts2.ClipData = 1 Then
128:       Set pDoc = m_pMxDoc
129:       If TypeOf pDoc.ActiveView Is IMap Then
130:         pDoc.FocusMap.ClipGeometry = Nothing
131:         g_bClipFlag = False
132:       End If
'    Else
134:       RemoveClipElement m_pMxDoc
'      g_bClipFlag = False
'    End If
137:   End If
  
139:   If g_bRotateFlag Then
140:     Set pDoc = m_pMxDoc
141:     If TypeOf pDoc.ActiveView Is IMap Then
142:       pDoc.ActiveView.ScreenDisplay.DisplayTransformation.Rotation = 0
143:       g_bRotateFlag = False
144:     End If
145:   End If
  
147:   If g_bLabelNeighbors Then
148:     Set pDoc = m_pMxDoc
149:     If TypeOf pDoc.ActiveView Is IMap Then
150:       RemoveLabels pDoc
151:       g_bLabelNeighbors = False
152:     End If
153:   End If
  
  Exit Function
ErrHand:
157:   MsgBox "DSMapBookExt_ActiveViewChanged - " & Erl & " - " & Err.Description
End Function

Private Function m_pMxDoc_CloseDocument() As Boolean
161:   DeleteSeries
End Function

Private Function m_pMxDoc_NewDocument() As Boolean
165:   DeleteSeries
End Function

Public Sub DeleteSeries()
On Error GoTo ErrHand:
  Dim pDoc As IMxDocument, pActive As IActiveView, pMapSeries As IDSMapSeries
  
172:   Set pDoc = m_pMxDoc
173:   Set pActive = pDoc.FocusMap
174:   Set pMapSeries = m_pMapBook.ContentItem(0)
  If pMapSeries Is Nothing Then Exit Sub
  
177:   TurnOffClipping pMapSeries, m_pApp
178:   Set pMapSeries = Nothing
179:   m_pMapBook.RemoveContent 0
180:   g_pFrmMapSeries.tvwMapBook.Nodes.Clear
181:   g_pFrmMapSeries.tvwMapBook.Nodes.Add , , "MapBook", "Map Book", 1
182:   RemoveIndicators m_pApp
183:   pActive.Refresh

  Exit Sub
ErrHand:
187:   MsgBox "DSMapBookExt_DeleteSeries - " & Erl & " - " & Err.Description
End Sub
