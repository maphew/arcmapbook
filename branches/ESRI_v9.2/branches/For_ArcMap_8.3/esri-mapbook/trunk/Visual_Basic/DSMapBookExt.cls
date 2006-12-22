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
Option Explicit

Implements IExtension
Implements IPersistVariant

Private m_pApp As IApplication
Private WithEvents m_pMxDoc As MxDocument
Attribute m_pMxDoc.VB_VarHelpID = -1
Private m_pMapBook As IDSMapBook

Private Property Get IExtension_Name() As String
  IExtension_Name = "DevSample_MapBook"
End Property

Private Sub IExtension_Shutdown()
  Set g_pFrmMapSeries = Nothing
  Set m_pMapBook = Nothing
  Set m_pMxDoc = Nothing
  Set m_pApp = Nothing
End Sub

Public Property Get MapBook() As IDSMapBook
  Set MapBook = m_pMapBook
End Property

Private Sub IExtension_Startup(initializationData As Variant)
  Dim pApp As IApplication
  Set g_pFrmMapSeries = New frmMapSeries
  Set g_pFrmMapSeries.m_pApp = initializationData
  Set m_pMapBook = New DSMapBook
  Set pApp = initializationData
  Set m_pApp = pApp
  Set m_pMxDoc = pApp.Document
  g_bClipFlag = False
  g_bRotateFlag = False
  g_bLabelNeighbors = False
End Sub

Private Property Get IPersistVariant_ID() As esriCore.IUID
On Error GoTo ErrHand:
  
  Dim pUID As New UID
  pUID = "DSMapBookUIPrj.DSMapBookExt"

  Set IPersistVariant_ID = pUID

  GoTo EndProc

ErrHand:
  MsgBox "ID - " & Err.Description
  Exit Property
EndProc:
  Set pUID = Nothing
End Property

Private Sub IPersistVariant_Load(ByVal Stream As esriCore.IVariantStream)
On Error GoTo ErrHand:
  Dim pNode As Node, pMapSeries As IDSMapSeries, lLoop As Long, sName As String
  Set m_pMapBook = Stream.Read
  
  'Make sure we have the tab form
  If g_pFrmMapSeries Is Nothing Then Exit Sub
  
  'Make sure our persisted map book has content
  If m_pMapBook.ContentCount = 0 Then Exit Sub
  
  'Put the content back on the form
  Set pMapSeries = m_pMapBook.ContentItem(0)
  With g_pFrmMapSeries.tvwMapBook
    Set pNode = .Nodes.Add("MapBook", tvwChild, "MapSeries", "Map Series", 3)
    
    'Now loop back through the list and add the tile names as nodes in the tree
    For lLoop = 0 To pMapSeries.PageCount - 1
      sName = pMapSeries.Page(lLoop).PageName
      If pMapSeries.Page(lLoop).EnablePage Then
        Set pNode = .Nodes.Add("MapSeries", tvwChild, "a" & sName, lLoop + 1 & " - " & sName, 5)
      Else
        Set pNode = .Nodes.Add("MapSeries", tvwChild, "a" & sName, lLoop + 1 & " - " & sName, 6)
      End If
      pNode.Tag = lLoop
    Next lLoop
    .Nodes.Item("MapBook").Expanded = True
    .Nodes.Item("MapSeries").Expanded = True
  End With
  
  Exit Sub
ErrHand:
  MsgBox "IPersistVariant_Load - " & Err.Description
End Sub

Private Sub IPersistVariant_Save(ByVal Stream As esriCore.IVariantStream)
On Error GoTo ErrHand:
  Stream.Write m_pMapBook
  
  Exit Sub
ErrHand:
  MsgBox "IPersistVariant_Save - " & Err.Description
End Sub

Private Function m_pMxDoc_ActiveViewChanged() As Boolean
On Error GoTo ErrHand:
'Check to see if the active view is being changed back to the data view after the Map Book code has
'set the clip shape.  The g_bClipFlag variable will tell us if the Map Book code has updated the
'Clip Shape.  If it has, then we want to clear the clip shape when the user switches back to a data view.
'If the clip shape was changed by some other method, then we don't want to update it.
  Dim pDoc As IMxDocument
  
  If g_bClipFlag Then
'    If pSeriesOpts2.ClipData = 1 Then
      Set pDoc = m_pMxDoc
      If TypeOf pDoc.ActiveView Is IMap Then
        pDoc.FocusMap.ClipGeometry = Nothing
        g_bClipFlag = False
      End If
'    Else
      RemoveClipElement m_pMxDoc
'      g_bClipFlag = False
'    End If
  End If
  
  If g_bRotateFlag Then
    Set pDoc = m_pMxDoc
    If TypeOf pDoc.ActiveView Is IMap Then
      pDoc.ActiveView.ScreenDisplay.DisplayTransformation.Rotation = 0
      g_bRotateFlag = False
    End If
  End If
  
  If g_bLabelNeighbors Then
    Set pDoc = m_pMxDoc
    If TypeOf pDoc.ActiveView Is IMap Then
      RemoveLabels pDoc
      g_bLabelNeighbors = False
    End If
  End If
  
  Exit Function
ErrHand:
  MsgBox "m_pMxDoc_ActiveViewChanged - " & Err.Description
End Function

Private Function m_pMxDoc_CloseDocument() As Boolean
  DeleteSeries
End Function

Private Function m_pMxDoc_NewDocument() As Boolean
  DeleteSeries
End Function

Public Sub DeleteSeries()
On Error GoTo ErrHand:
  Dim pDoc As IMxDocument, pActive As IActiveView, pMapSeries As IDSMapSeries
  
  Set pDoc = m_pMxDoc
  Set pActive = pDoc.FocusMap
  Set pMapSeries = m_pMapBook.ContentItem(0)
  If pMapSeries Is Nothing Then Exit Sub
  
  TurnOffClipping pMapSeries, m_pApp
  Set pMapSeries = Nothing
  m_pMapBook.RemoveContent 0
  g_pFrmMapSeries.tvwMapBook.Nodes.Clear
  g_pFrmMapSeries.tvwMapBook.Nodes.Add , , "MapBook", "Map Book", 1
  RemoveIndicators m_pApp
  pActive.Refresh

  Exit Sub
ErrHand:
  MsgBox "DeleteSeries - " & Err.Description
End Sub
