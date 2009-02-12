VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NWMapBookExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Copyright 1995-2004 ESRI

' All rights reserved under the copyright laws of the United States.



Option Explicit

Implements IExtension
Implements IPersistVariant

Private m_pApp As IApplication
Private m_pMxDoc2 As IMxDocument
Private WithEvents m_pMxDoc As MxDocument
Attribute m_pMxDoc.VB_VarHelpID = -1
Private WithEvents m_pLayoutActiveViewEvents As PageLayout
Attribute m_pLayoutActiveViewEvents.VB_VarHelpID = -1
'Private WithEvents m_pMapActiveViewEvents As Map
Private m_pNWSeriesOpts As INWMapSeriesOptions


Private m_pMapBook As INWDSMapBook
Const c_sModuleFileName As String = "NWMapBookExt.cls"


Private Sub Class_Initialize()
  '''''''''debug
'  MsgBox "In NWMapBookUIPrj, NWMapBookExt, Class_Initialize called."
End Sub

Private Sub Class_Terminate()
  '''''''''debug
'  MsgBox "In NWMapBookUIPrj, NWMapBookExt, Class_Terminate called."
End Sub

Private Property Get IExtension_Name() As String
36:   IExtension_Name = "NW_MapBook"
End Property

Private Sub IExtension_Shutdown()
40:   Set g_pFrmMapSeries = Nothing
41:   Set m_pMapBook = Nothing
42:   Set m_pMxDoc = Nothing
43:   Set m_pMxDoc2 = Nothing
44:   Set m_pApp = Nothing
  ''''''''''''debug
'  MsgBox "In NWMapBookUIPrj, NWMapBookExt, IExtension_Shutdown," & vbNewLine _
       & "disabling m_pLayoutActiveViewEvents"

49:   Set m_pLayoutActiveViewEvents = Nothing
End Sub

Public Property Get MapBook() As INWDSMapBook
53:   Set MapBook = m_pMapBook
End Property

Private Sub IExtension_Startup(initializationData As Variant)
On Error GoTo ErrHand:
  Dim pApp As IApplication
59:   Set g_pFrmMapSeries = New frmMapSeries
60:   Set g_pFrmMapSeries.m_pApp = initializationData
  'Set m_pMapBook = New NWDSMapBook
62:   Set m_pMapBook = New NWMapBook
63:   Set pApp = initializationData
64:   Set m_pApp = pApp
65:   Set m_pMxDoc = pApp.Document
66:   Set m_pMxDoc2 = pApp.Document
67:   g_bClipFlag = False
68:   g_bRotateFlag = False
69:   g_bLabelNeighbors = False
  
  'enable handling of map page reloads when the map is refreshed
  
  Dim pMxDoc As IMxDocument
74:   Set pMxDoc = m_pApp.Document
  ''''''''''''debug
'  MsgBox "In NWMapBookUIPrj, NWMapBookExt, IExtension_Startup," & vbNewLine _
       & "enabling m_pLayoutActiveViewEvents"

79:   Set m_pLayoutActiveViewEvents = pMxDoc.PageLayout
  Exit Sub
ErrHand:
82:   MsgBox "NWMapBookExt_Startup - " & Erl & " - " & Err.Description
End Sub

Private Property Get IPersistVariant_ID() As IUID
On Error GoTo ErrHand:
  
  Dim pUID As New UID
89:   pUID = "NWMapBookUIPrj.NWMapBookExt"

91:   Set IPersistVariant_ID = pUID

93:   GoTo EndProc

ErrHand:
96:   MsgBox "NWMapBookExt_ID - " & Err.Description
  Exit Property
EndProc:
99:   Set pUID = Nothing
End Property

Private Sub IPersistVariant_Load(ByVal Stream As IVariantStream)
'On Error GoTo ExitSub
  On Error GoTo ErrHand
  Dim pNode As Node, pMapSeries As INWDSMapSeries, lLoop As Long, sName As String
  Dim pPage As INWDSMapPage, vStreamData As Variant, lMapBookCount As Long
  Dim pMxDoc As IMxDocument
  
  
  'lMapBookCount = Stream.Read
'On Error GoTo ErrHand
  'If lMapBookCount = 0 Then
  '  Exit Sub
  'End If
115:   Set m_pMapBook = Stream.Read
  
  'Make sure we have the tab form
  '''''''''''''debug
'  MsgBox "NWMapBookExt, IPersistVariant_Load" & vbNewLine _
       & "1"
  If g_pFrmMapSeries Is Nothing Then Exit Sub
  '''''''''''''debug
'  MsgBox "NWMapBookExt, IPersistVariant_Load" & vbNewLine _
       & "2"
  
  'Make sure our persisted map book has content
  If m_pMapBook.ContentCount = 0 Then Exit Sub
  '''''''''''''debug
'  MsgBox "NWMapBookExt, IPersistVariant_Load" & vbNewLine _
       & "3"
  
  'Put the content back on the form
133:   Set pMapSeries = m_pMapBook.ContentItem(0)
134:   With g_pFrmMapSeries.tvwMapBook
135:     Set pNode = .Nodes.Add("MapBook", tvwChild, "MapSeries", "Map Series", 3)
    
    'Now loop back through the list and add the tile names as nodes in the tree
138:     For lLoop = 0 To pMapSeries.PageCount - 1
139:       Set pPage = pMapSeries.Page(lLoop)
140:       sName = pPage.PageName
141:       If pPage.EnablePage Then
142:         Set pNode = .Nodes.Add("MapSeries", tvwChild, "a" & sName, pPage.PageNumber & " - " & sName, 5)
143:       Else
144:         Set pNode = .Nodes.Add("MapSeries", tvwChild, "a" & sName, pPage.PageNumber & " - " & sName, 6)
145:       End If
146:       pNode.Tag = lLoop
147:     Next lLoop
148:     .Nodes.Item("MapBook").Expanded = True
149:     .Nodes.Item("MapSeries").Expanded = True
150:   End With
  '''''''''''''debug
'  MsgBox "NWMapBookExt, IPersistVariant_Load" & vbNewLine _
       & "4"
154:     Set pMxDoc = m_pApp.Document
155:     Set m_pLayoutActiveViewEvents = pMxDoc.PageLayout
  
ExitSub:
  Exit Sub

ErrHand:
161:   MsgBox "NWMapBookExt_IPersistVariant_Load - " & Erl & " - " & Err.Description & vbNewLine _
      & "This error typically happens when the NW Mapbook application has been upgraded." & vbNewLine _
      & "Microsoft COM technology limitations require that whenever this application is " & vbNewLine _
      & "upgraded, a new unique number is created to identify that version of the application." & vbNewLine _
      & "When you have the NW Mapbook application installed, it saves its settings" & vbNewLine _
      & "in ArcMap's MXD documents along with this signature number." & vbNewLine _
      & "When you upgrade to a new mapbook version and attempt to load an older MXD, " & vbNewLine _
      & "the current application doesn't match the saved number, and this error is triggered." & vbNewLine _
      & "As a result, the previous mapbook settings could not be recovered, and you will need" & vbNewLine _
      & "to recreate the map series from scratch." & vbNewLine _
      & vbNewLine & "It is possible to create a map series export/import function to save settings to a" & vbNewLine _
      & "database schema, and import your previous work after an upgrade, but that feature" & vbNewLine _
      & "has not yet been implemented."
174:   Err.Clear
    'testing showed that errors during loads caused the iactiveviewevents to
    'stop being caught, so this code to refresh the event handling was
    'added [RAH, 10-18-05]
178:     Set pMxDoc = m_pApp.Document
179:     Set m_pLayoutActiveViewEvents = pMxDoc.PageLayout
End Sub

Private Sub IPersistVariant_Save(ByVal Stream As IVariantStream)
On Error GoTo ErrHand:
  'Dim lMapBooksCount As Long
  
  'If m_pMapBook Is Nothing Then
  '  lMapBooksCount = 0
  '  Stream.Write lMapBooksCount
  'Else
  '  lMapBooksCount = 1
  '  Stream.Write lMapBooksCount
192:     Stream.Write m_pMapBook
  'End If
  
  
  
  Exit Sub
ErrHand:
199:   MsgBox "NWMapBookExt_IPersistVariant_Save - " & Erl & " - " & Err.Description
200:   Err.Clear
End Sub





Private Function m_pMxDoc_ActiveViewChanged() As Boolean
On Error GoTo ErrHand:
'Check to see if the active view is being changed back to the data view after the Map Book code has
'set the clip shape.  The g_bClipFlag variable will tell us if the Map Book code has updated the
'Clip Shape.  If it has, then we want to clear the clip shape when the user switches back to a data view.
'If the clip shape was changed by some other method, then we don't want to update it.
  Dim pDoc As IMxDocument
  
215:   If g_bClipFlag Then
'    If pSeriesOpts2.ClipData = 1 Then
217:       Set pDoc = m_pMxDoc
218:       If TypeOf pDoc.ActiveView Is IMap Then
219:         pDoc.FocusMap.ClipGeometry = Nothing
220:         g_bClipFlag = False
221:       End If
'    Else
223:       RemoveClipElement m_pMxDoc
'      g_bClipFlag = False
'    End If
226:   End If
  
228:   If g_bRotateFlag Then
229:     Set pDoc = m_pMxDoc
230:     If TypeOf pDoc.ActiveView Is IMap Then
231:       pDoc.ActiveView.ScreenDisplay.DisplayTransformation.Rotation = 0
232:       g_bRotateFlag = False
233:     End If
234:   End If
  
236:   If g_bLabelNeighbors Then
237:     Set pDoc = m_pMxDoc
238:     If TypeOf pDoc.ActiveView Is IMap Then
239:       RemoveLabels pDoc
240:       g_bLabelNeighbors = False
241:     End If
242:   End If
  
  Exit Function
ErrHand:
246:   MsgBox "NWMapBookExt_ActiveViewChanged - " & Erl & " - " & Err.Description
247:   Err.Clear
End Function

Private Function m_pMxDoc_CloseDocument() As Boolean
251:   DeleteSeries
End Function

Private Function m_pMxDoc_NewDocument() As Boolean
255:   DeleteSeries
End Function

Public Sub DeleteSeries()
On Error GoTo ErrHand:
  Dim pDoc As IMxDocument, pActive As IActiveView, pMapSeries As INWDSMapSeries
  
262:   Set pDoc = m_pMxDoc
263:   Set pActive = pDoc.FocusMap
264:   Set pMapSeries = m_pMapBook.ContentItem(0)
  If pMapSeries Is Nothing Then Exit Sub
  
267:   TurnOffClipping pMapSeries, m_pApp
268:   Set pMapSeries = Nothing
269:   m_pMapBook.RemoveContent 0
270:   g_pFrmMapSeries.tvwMapBook.Nodes.Clear
271:   g_pFrmMapSeries.tvwMapBook.Nodes.Add , , "MapBook", "Map Book", 1
272:   RemoveIndicators m_pApp
273:   pActive.Refresh

  Exit Sub
ErrHand:
277:   MsgBox "NWMapBookExt_DeleteSeries - " & Erl & " - " & Err.Description
278:   Err.Clear
End Sub














''''''''''''''''''''''''''''''''''
''Map refresh event handling code
''
''These routines allow for page loading objects such as the bubble
''detail insets to be redrawn whenever a map page is refreshed.
''
'-----------------------------------


Private Sub m_pLayoutActiveViewEvents_ViewRefreshed(ByVal view As esriCarto.IActiveView, ByVal phase As esriCarto.esriViewDrawPhase, ByVal data As Variant, ByVal envelope As esriGeometry.IEnvelope)
  On Error GoTo ErrorHandler
  
  'This code will reload the current map page in the map
  'book when the map refresh command has been given.
  '
  'multiple conditions must be met first however.
  ' - ArcMap must be in layout view (handled when m_pLayoutActiveViewEvents
  '   was declared "as PageLayout", and was assigned to pMxDoc.PageLayout
  ' - A map book must have been created
  ' - A map series must exist
  ' - the user must have chosen to have map pages loaded
  ' - A map page must have been selected
  ' - when a map page is loaded, it triggers another refresh event, causing
  '   this code to call itself.  The bReentrant static variable catches this,
  '   preventing an infinite loop.
  
  Dim pMapSeries As INWDSMapSeries
  Dim pNWMapSeries As INWMapSeriesOptions
  
'  MsgBox "In NWMapBookExt, m_pLayoutActiveViewEvents_ViewRefreshed, " & vbNewLine _
       & "  phase is " & phase & vbNewLine _
       & vbNewLine _
       & "for reference, the draw phases are: " & vbNewLine _
       & "esriViewNone: 0" & vbNewLine _
       & "esriViewBackground: 1" & vbNewLine _
       & "esriViewGeography: 2" & vbNewLine _
       & "esriViewGeoSelection: 4" & vbNewLine _
       & "esriViewGraphics: 8" & vbNewLine _
       & "esriViewGraphicSelection: 16" & vbNewLine _
       & "esriViewForeground: 32" & vbNewLine

  'refresh if either a single bubble polygon is being edited,
  'or if a single layout element is selected

'if there is one or more layout elements selected, and if this is the graphics selection phase, then
'let the code keep running
  Dim pFeatLayer As IFeatureLayer, pFeatClass As IFeatureClass, pFeatSelection As IFeatureSelection
  Dim pFeatDataset As IFeatureDataset, pGraCont As IGraphicsContainer, pElement As IElement, bCanContinue As Boolean
  Dim pPageLayout As IPageLayout, pGraContSelect As IGraphicsContainerSelect
  
  If m_pMxDoc2 Is Nothing Then Exit Sub
''''''''''''debug
'MsgBox "0"
347:   Set pGraContSelect = m_pMxDoc2.PageLayout
  
                                            'if user is trying to move a detail inset, then don't
                                            'handle all of the refresh events -- only the refresh for
                                            'when graphics are selected.  Otherwise, handling the refresh
                                            'event will deselect the element immediately after selection
353:   If phase <> 65535 Then
    Exit Sub
355:   End If
'  If pGraContSelect.ElementSelectionCount > 0 Then
'    Set pElement = pGraContSelect.DominantElement
'    If TypeOf pElement Is IMapFrame Then
'      If phase <> esriViewGraphicSelection Then
'        Exit Sub
'      End If
'    End If
'  End If
  
''''''''''debug
'MsgBox "1"
  
  If m_pMapBook Is Nothing Then Exit Sub
''''''''''debug
'MsgBox "2"
371:   Set pMapSeries = m_pMapBook.ContentItem(0)
  If pMapSeries Is Nothing Then Exit Sub
''''''''''debug
'MsgBox "3"
  
376:   Set pNWMapSeries = pMapSeries
  If pNWMapSeries Is Nothing Then Exit Sub
                                            'sometimes a detail inset has been temporarily deleted, but
''''''''''debug
'MsgBox "4"
                                            'the refresh event is triggered before the replacement detail
                                            'inset is added.  The RefreshEventShouldntRefresh property is used
                                            'to prevent the refresh event from causing bugs because a detail
                                            'inset hasn't yet been added.  This happens when editing an
                                            'existing detail inset using the inset editing tool.
  If pNWMapSeries.RefreshEventShouldntRefresh Then Exit Sub
''''''''''debug
'MsgBox "5"
  If Len(pNWMapSeries.BubbleLayer) = 0 Then Exit Sub
''''''''''debug
'MsgBox "6"
  
  If Not pNWMapSeries.RefreshEventLoadPage Then Exit Sub 'let user decide not to load map pages
''''''''''debug
'MsgBox "7"
  If pMapSeries.PageCount = 0 Then Exit Sub
''''''''''debug
'MsgBox "8"
399:   If Len(pNWMapSeries.CurrentPageName) = 0 Then
''''''''''debug
'MsgBox "9"
    Exit Sub
403:   End If
  
  Dim pMapPage As INWDSMapPage
406:   Set pMapPage = pNWMapSeries.CurrentPage
  If pMapPage Is Nothing Then Exit Sub
''''''''''debug
'MsgBox "10"

  'pMapPage.DrawPage m_pApp.Document, pMapSeries, True  'this line would reload the entire map page instead of just detail insets
  'only redraw the detail insets
  Dim pMxDoc As IMxDocument, pLayer As esriCarto.ILayer
414:   Set pMxDoc = m_pApp.Document
  
416:   Set pLayer = LayerFromDataFrame(pNWMapSeries.DataFrameMainFrame, pNWMapSeries.BubbleLayer, pMxDoc)
  If pNWMapSeries.BubbleLayer = "" Then Exit Sub  'don't bother refreshing detail insets if no bubble layer was defined
  '''''''''''''''debug
'  MsgBox "In NWMapBookUIPrj, NWMapBookExt.cls, m_pLayoutActiveViewEvents_ViewRefreshed," & vbNewLine _
       & "About to refresh the detail insets"
421:   pNWMapSeries.RefreshDetailInsets pMxDoc, pLayer, pMapPage.PageShape, True
  
  Exit Sub
  
ErrorHandler:
  HandleError False, "m_pLayoutActiveViewEvents_ViewRefreshed " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub




