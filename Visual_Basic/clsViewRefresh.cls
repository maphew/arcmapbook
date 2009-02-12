VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsViewRefresh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

'This module was created to trap active view events so that
'the current map page is reloaded.  Without this code, refreshing
'the map will not cause the graphics for the current map page to
'be refreshed.

Private m_pApp As IApplication



Private WithEvents m_pActiveViewEvents As Map
Attribute m_pActiveViewEvents.VB_VarHelpID = -1





Public Sub SetEvents()
  Dim pMxDoc As IMxDocument
  Set pMxDoc = m_pApp.Document
  Set m_pActiveViewEvents = pMxDoc.FocusMap

  '''''''''''''
  '  Set pMapBook = GetMapBookExtension(m_pApp)
  '  If pMapBook Is Nothing Then Exit Sub
  '
  '  Set pMapSeries = pMapBook.ContentItem(0)
  '  Set pSeriesOpts = pMapSeries
  '  Set pSeriesOpts2 = pSeriesOpts
  '  Set pCOSSeriesOpts = pMapSeries
  '
  '  lPage = m_pCurrentNode.Tag
  '  Select Case Index
  '  Case 0  'View Page
  '    Set pMapPage = pMapSeries.Page(lPage)
  '    pMapPage.DrawPage m_pApp.Document, pMapSeries, True
  '''''''''''''
  
End Sub





Private Sub m_pActiveViewEvents_ViewRefreshed(ByVal view As esriCarto.IActiveView, ByVal phase As esriCarto.esriViewDrawPhase, ByVal data As Variant, ByVal envelope As esriGeometry.IEnvelope)
  'detect if a map series exists
  
  'see if a map page was selected
  
  'call the load page routine for that map page
End Sub
