VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PageIdentifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ICommand
Implements ITool

Dim m_pApp As IApplication
Dim m_sName As String

Private Property Get ICommand_Bitmap() As esriCore.OLE_HANDLE
  ICommand_Bitmap = frmResources.picIdentifier.Picture.Handle
End Property

Private Property Get ICommand_Caption() As String
  ICommand_Caption = "Add Identifier Frame"
End Property

Private Property Get ICommand_Category() As String
  ICommand_Category = "Developer Samples"
End Property

Private Property Get ICommand_Checked() As Boolean
  ICommand_Checked = False
End Property

Private Property Get ICommand_Enabled() As Boolean
  ICommand_Enabled = True
End Property

Private Property Get ICommand_HelpContextID() As Long

End Property

Private Property Get ICommand_HelpFile() As String

End Property

Private Property Get ICommand_Message() As String
  ICommand_Message = "Identifier Frame"
End Property

Private Property Get ICommand_Name() As String
  ICommand_Name = "DSMapBookUIPrj.PageIdentifier"
End Property

Private Sub ICommand_OnClick()
    'Get Identifer type
  frmPageIdentifier.Show vbModal
  If frmPageIdentifier.optIdentifier(0).Value Then
    m_sName = "Local Indicator"
  Else
    m_sName = "Global Indicator"
  End If
  Unload frmPageIdentifier
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
  Set m_pApp = hook
End Sub

Private Property Get ICommand_Tooltip() As String
  ICommand_Tooltip = "Add Identifier Frame"
End Property

Private Property Get ITool_Cursor() As esriCore.OLE_HANDLE
On Error GoTo ErrHand:
  ITool_Cursor = frmResources.imlIcons.ListImages(1).Picture
  
  Exit Property
ErrHand:
  MsgBox "ITool_Cursor - " & Err.Description
End Property

Private Function ITool_Deactivate() As Boolean
  ITool_Deactivate = True
End Function

Private Function ITool_OnContextMenu(ByVal X As Long, ByVal Y As Long) As Boolean

End Function

Private Sub ITool_OnDblClick()

End Sub

Private Sub ITool_OnKeyDown(ByVal KeyCode As Long, ByVal Shift As Long)

End Sub

Private Sub ITool_OnKeyUp(ByVal KeyCode As Long, ByVal Shift As Long)

End Sub

Private Sub ITool_OnMouseDown(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
On Error GoTo ErrHand:
  Dim pGraphicsContainer As IGraphicsContainer, pLineSym2 As ISimpleLineSymbol
  Dim pElement As IElement, pMxApp As IMxApplication
  Dim rMapFrame As IMapFrame, pFeatLayer As IFeatureLayer
  Dim pMap As IMap, pGridLayer As IFeatureLayer
  Dim pColor2 As IColor, pColor3 As IColor
  Dim pRubberBand As IRubberBand, pScreenDisplay As IScreenDisplay
  Dim pGeometry As IGeometry, lLoop As Long, pFeatSel As IFeatureSelection
  Dim pMxDoc As IMxDocument, pLayer As ILayer, pActive As IActiveView
  Dim pRend As ISimpleRenderer, pColor As IRgbColor, pFill As ISimpleFillSymbol
  Dim pLineSym As ISimpleLineSymbol, pGeoFeatLayer As IGeoFeatureLayer
  Dim pMapBook As IDSMapBook
  Dim pSeriesProps As IDSMapSeriesProps, pFill2 As ISimpleFillSymbol
    
  Set pMxApp = m_pApp
  Set pMxDoc = m_pApp.Document
  Set pGraphicsContainer = pMxDoc.PageLayout
  Set pRubberBand = New RubberEnvelope
  Set pScreenDisplay = pMxApp.Display
  Set pGeometry = pRubberBand.TrackNew(pScreenDisplay, Nothing)
  
  'Get the index layer from the current Map Series
  Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
  If pMapBook.ContentCount = 0 Then
    MsgBox "You need to create a Map Series before adding a Page Identifier!!!"
    Exit Sub
  End If
  Set pSeriesProps = pMapBook.ContentItem(0)
  'Find the data frame
  Set pMap = FindDataFrame(pMxDoc, pSeriesProps.DataFrameName)
  If pMap Is Nothing Then
    MsgBox "Could not find map in PageIdentifier_OnMouseDown routine!!!"
    Exit Sub
  End If
  
  'Find the Index layer
  Set pGridLayer = FindLayer(pSeriesProps.IndexLayerName, pMap)
  If pGridLayer Is Nothing Then
    MsgBox "Could not find index layer in PageIdentifier_OnMouseDown routine!!!"
    Exit Sub
  End If
  
  'Create a new map and layer
  Set pMap = New Map
  pMap.Name = m_sName
  Set pFeatLayer = New FeatureLayer
  Set pRend = New SimpleRenderer
  Set pColor = New RgbColor    'Fill Outline symbol
  Set pColor2 = New RgbColor   'Fill Symbol
  Set pColor3 = New RgbColor   'Selection Fill symbol
  pColor.RGB = RGB(0, 0, 0)
  pColor2.NullColor = True
  pColor3.RGB = RGB(180, 180, 180)
  Set pFill = New SimpleFillSymbol
  Set pFill2 = New SimpleFillSymbol
  Set pLineSym = New SimpleLineSymbol
  Set pLineSym2 = New SimpleLineSymbol
  pLineSym.Color = pColor
  pLineSym.Width = 1
  pFill.Color = pColor2
  pFill.Outline = pLineSym
  Set pRend.Symbol = pFill
  Set pFeatLayer.FeatureClass = pGridLayer.FeatureClass
  pFeatLayer.Name = "Identifier Layer"
  Set pGeoFeatLayer = pFeatLayer
  Set pGeoFeatLayer.Renderer = pRend
  pLineSym2.Color = pColor2
  pFill2.Color = pColor3
  pFill2.Outline = pLineSym
  pMap.AddLayer pFeatLayer
  Set pFeatSel = pFeatLayer
  Set pFeatSel.SelectionSymbol = pFill2
  pFeatSel.SetSelectionSymbol = True
    
  'Create a new MapFrame and associate map with it
  Dim pMapFrame As IMapFrame
  Set pMapFrame = New MapFrame
  Set pMapFrame.Map = pMap
  
  'Set the position of the new map frame
  Set pElement = pMapFrame
  pElement.Geometry = pGeometry
  
  'Add mapframe to the layout
  pGraphicsContainer.AddElement pMapFrame, 0
  Set pActive = pMap
  pActive.Refresh
  
  'Refresh ActiveView and TOC
  Set pActive = pMxDoc.FocusMap
  pActive.Refresh
  pMxDoc.CurrentContentsView.Refresh 0
  
  'Deactivate the tool
  Set m_pApp.CurrentTool = Nothing
  
  Exit Sub
ErrHand:
  MsgBox "PageIdentifier_OnMouseDown - " & Err.Description
End Sub

Private Sub ITool_OnMouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)

End Sub

Private Sub ITool_OnMouseUp(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)

End Sub

Private Sub ITool_Refresh(ByVal hdc As esriCore.OLE_HANDLE)

End Sub
