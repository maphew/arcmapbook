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
11:   IExtension_Name = "DevSample_MapBook"
End Property

Private Sub IExtension_Shutdown()
15:   Set g_pFrmMapSeries = Nothing
16:   Set m_pMapBook = Nothing
17:   Set m_pMxDoc = Nothing
18:   Set m_pApp = Nothing
End Sub

Public Property Get MapBook() As IDSMapBook
22:   Set MapBook = m_pMapBook
End Property

Private Sub IExtension_Startup(initializationData As Variant)
On Error GoTo ErrHand:
  Dim pApp As IApplication
28:   Set g_pFrmMapSeries = New frmMapSeries
29:   Set g_pFrmMapSeries.m_pApp = initializationData
30:   Set m_pMapBook = New DSMapBook
31:   Set pApp = initializationData
32:   Set m_pApp = pApp
33:   Set m_pMxDoc = pApp.Document
34:   g_bClipFlag = False
35:   g_bRotateFlag = False
36:   g_bLabelNeighbors = False
  
  Exit Sub
ErrHand:
40:   MsgBox "DSMapBookExt_Startup - " & Erl & " - " & Err.Description
End Sub

Private Property Get IPersistVariant_ID() As IUID
On Error GoTo ErrHand:
  
  Dim pUID As New UID
47:   pUID = "DSMapBookUIPrj.DSMapBookExt"

49:   Set IPersistVariant_ID = pUID

51:   GoTo EndProc

ErrHand:
54:   MsgBox "DSMapBookExt_ID - " & Err.Description
  Exit Property
EndProc:
57:   Set pUID = Nothing
End Property

Private Sub IPersistVariant_Load(ByVal Stream As IVariantStream)
On Error GoTo ErrHand:
  Dim pNode As Node, pMapSeries As IDSMapSeries, lLoop As Long, sName As String
  Dim pPage As IDSMapPage
64:   Set m_pMapBook = Stream.Read
  
  'Make sure we have the tab form
  If g_pFrmMapSeries Is Nothing Then Exit Sub
  
  'Make sure our persisted map book has content
  If m_pMapBook.ContentCount = 0 Then Exit Sub
  
  'Put the content back on the form
73:   Set pMapSeries = m_pMapBook.ContentItem(0)
74:   With g_pFrmMapSeries.tvwMapBook
75:     Set pNode = .Nodes.Add("MapBook", tvwChild, "MapSeries", "Map Series", 3)
    
    'Now loop back through the list and add the tile names as nodes in the tree
78:     For lLoop = 0 To pMapSeries.PageCount - 1
79:       Set pPage = pMapSeries.Page(lLoop)
80:       sName = pPage.PageName
81:       If pPage.EnablePage Then
82:         Set pNode = .Nodes.Add("MapSeries", tvwChild, "a" & sName, pPage.PageNumber & " - " & sName, 5)
83:       Else
84:         Set pNode = .Nodes.Add("MapSeries", tvwChild, "a" & sName, pPage.PageNumber & " - " & sName, 6)
85:       End If
86:       pNode.Tag = lLoop
87:     Next lLoop
88:     .Nodes.Item("MapBook").Expanded = True
89:     .Nodes.Item("MapSeries").Expanded = True
90:   End With
  
  Exit Sub
ErrHand:
94:   MsgBox "DSMapBookExt_IPersistVariant_Load - " & Erl & " - " & Err.Description
End Sub

Private Sub IPersistVariant_Save(ByVal Stream As IVariantStream)
On Error GoTo ErrHand:
99:   Stream.Write m_pMapBook
  
  Exit Sub
ErrHand:
103:   MsgBox "DSMapBookExt_IPersistVariant_Save - " & Erl & " - " & Err.Description
End Sub

Private Function m_pMxDoc_ActiveViewChanged() As Boolean
On Error GoTo ErrHand:
'Check to see if the active view is being changed back to the data view after the Map Book code has
'set the clip shape.  The g_bClipFlag variable will tell us if the Map Book code has updated the
'Clip Shape.  If it has, then we want to clear the clip shape when the user switches back to a data view.
'If the clip shape was changed by some other method, then we don't want to update it.
  Dim pDoc As IMxDocument
  
114:   If g_bClipFlag Then
'    If pSeriesOpts2.ClipData = 1 Then
116:       Set pDoc = m_pMxDoc
117:       If TypeOf pDoc.ActiveView Is IMap Then
118:         pDoc.FocusMap.ClipGeometry = Nothing
119:         g_bClipFlag = False
120:       End If
'    Else
122:       RemoveClipElement m_pMxDoc
'      g_bClipFlag = False
'    End If
125:   End If
  
127:   If g_bRotateFlag Then
128:     Set pDoc = m_pMxDoc
129:     If TypeOf pDoc.ActiveView Is IMap Then
130:       pDoc.ActiveView.ScreenDisplay.DisplayTransformation.Rotation = 0
131:       g_bRotateFlag = False
132:     End If
133:   End If
  
135:   If g_bLabelNeighbors Then
136:     Set pDoc = m_pMxDoc
137:     If TypeOf pDoc.ActiveView Is IMap Then
138:       RemoveLabels pDoc
139:       g_bLabelNeighbors = False
140:     End If
141:   End If
  
  Exit Function
ErrHand:
145:   MsgBox "DSMapBookExt_ActiveViewChanged - " & Erl & " - " & Err.Description
End Function

Private Function m_pMxDoc_CloseDocument() As Boolean
149:   DeleteSeries
End Function

Private Function m_pMxDoc_NewDocument() As Boolean
153:   DeleteSeries
End Function

Public Sub DeleteSeries()
On Error GoTo ErrHand:
  Dim pDoc As IMxDocument, pActive As IActiveView, pMapSeries As IDSMapSeries
  
160:   Set pDoc = m_pMxDoc
161:   Set pActive = pDoc.FocusMap
162:   Set pMapSeries = m_pMapBook.ContentItem(0)
  If pMapSeries Is Nothing Then Exit Sub
  
165:   TurnOffClipping pMapSeries, m_pApp
166:   Set pMapSeries = Nothing
167:   m_pMapBook.RemoveContent 0
168:   g_pFrmMapSeries.tvwMapBook.Nodes.Clear
169:   g_pFrmMapSeries.tvwMapBook.Nodes.Add , , "MapBook", "Map Book", 1
170:   RemoveIndicators m_pApp
171:   pActive.Refresh

  Exit Sub
ErrHand:
175:   MsgBox "DSMapBookExt_DeleteSeries - " & Erl & " - " & Err.Description
End Sub
