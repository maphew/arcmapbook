VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DSMapSeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IDSMapSeries
Implements IDSMapSeriesProps
Implements IDSMapSeriesOptions
Implements IDSMapSeriesOptions2  'Added 6/18/03 to support cross hatching of clip
Implements IPersistVariant

Private m_PageColl As Collection
Private m_bEnableSeries As Boolean
Private m_bClipData As Boolean
Private m_lClipData2 As Long  'Added 6/18/03 to support cross hatching of clip
Private m_sDataDrivenField As String
Private m_lExtentType As Long
Private m_dFixedScale As Double
Private m_bLabelNeighbors As Boolean
Private m_pLabelSymbol As ISymbol
Private m_dMargin As Double
Private m_sMarginType As String
Private m_bRotateFrame As Boolean
Private m_sRotationField As String
Private m_sDataFrameName As String
Private m_sIndexFieldName As String
Private m_sIndexLayerName As String
Private m_SuppressColl As Collection
Private m_bSupressLayers As Boolean
Private m_lTileSelection As Long

Private Sub Class_Initialize()
  Set m_PageColl = New Collection
  Set m_SuppressColl = New Collection
End Sub

Private Sub Class_Terminate()
  Set m_PageColl = Nothing
  Set m_SuppressColl = Nothing
End Sub

Private Sub IDSMapSeries_AddPage(ByVal Page As DSMapBookPrj.IDSMapPage)
  m_PageColl.Add Page
End Sub

Private Property Get IDSMapSeries_Page(Index As Long) As DSMapBookPrj.IDSMapPage
  If Index > -1 And Index < m_PageColl.Count Then
    Set IDSMapSeries_Page = m_PageColl.Item(Index + 1)
  Else
    Set IDSMapSeries_Page = Nothing
  End If
End Property

Private Property Get IDSMapSeries_PageCount() As Long
  IDSMapSeries_PageCount = m_PageColl.Count
End Property

Private Property Let IDSMapSeries_EnableSeries(ByVal RHS As Boolean)
  m_bEnableSeries = RHS
End Property

Private Property Get IDSMapSeries_EnableSeries() As Boolean
  IDSMapSeries_EnableSeries = m_bEnableSeries
End Property

Private Sub IDSMapSeries_RemovePage(Index As Long)
  If Index > -1 And Index < m_PageColl.Count Then
    m_PageColl.Remove Index + 1
  End If
End Sub

Private Property Let IDSMapSeriesOptions_ClipData(RHS As Boolean)
  m_bClipData = RHS
End Property

Private Property Get IDSMapSeriesOptions_ClipData() As Boolean
  IDSMapSeriesOptions_ClipData = m_bClipData
End Property

Private Property Let IDSMapSeriesOptions_DataDrivenField(RHS As String)
  m_sDataDrivenField = RHS
End Property

Private Property Get IDSMapSeriesOptions_DataDrivenField() As String
  IDSMapSeriesOptions_DataDrivenField = m_sDataDrivenField
End Property

Private Property Let IDSMapSeriesOptions_ExtentType(RHS As Long)
  m_lExtentType = RHS
End Property

Private Property Get IDSMapSeriesOptions_ExtentType() As Long
  IDSMapSeriesOptions_ExtentType = m_lExtentType
End Property

Private Property Let IDSMapSeriesOptions_FixedScale(RHS As Double)
  m_dFixedScale = RHS
End Property

Private Property Get IDSMapSeriesOptions_FixedScale() As Double
  IDSMapSeriesOptions_FixedScale = m_dFixedScale
End Property

Private Property Let IDSMapSeriesOptions_LabelNeighbors(RHS As Boolean)
  m_bLabelNeighbors = RHS
End Property

Private Property Get IDSMapSeriesOptions_LabelNeighbors() As Boolean
  IDSMapSeriesOptions_LabelNeighbors = m_bLabelNeighbors
End Property

Private Property Set IDSMapSeriesOptions_LabelSymbol(RHS As esriCore.ISymbol)
  Set m_pLabelSymbol = RHS
End Property

Private Property Get IDSMapSeriesOptions_LabelSymbol() As esriCore.ISymbol
  Set IDSMapSeriesOptions_LabelSymbol = m_pLabelSymbol
End Property

Private Property Let IDSMapSeriesOptions_Margin(RHS As Double)
  m_dMargin = RHS
End Property

Private Property Get IDSMapSeriesOptions_Margin() As Double
  IDSMapSeriesOptions_Margin = m_dMargin
End Property

Private Property Let IDSMapSeriesOptions_MarginType(RHS As String)
  m_sMarginType = RHS
End Property

Private Property Get IDSMapSeriesOptions_MarginType() As String
  IDSMapSeriesOptions_MarginType = m_sMarginType
End Property

Private Property Let IDSMapSeriesOptions_RotateFrame(RHS As Boolean)
  m_bRotateFrame = RHS
End Property

Private Property Get IDSMapSeriesOptions_RotateFrame() As Boolean
  IDSMapSeriesOptions_RotateFrame = m_bRotateFrame
End Property

Private Property Let IDSMapSeriesOptions_RotationField(RHS As String)
  m_sRotationField = RHS
End Property

Private Property Get IDSMapSeriesOptions_RotationField() As String
  IDSMapSeriesOptions_RotationField = m_sRotationField
End Property

Private Property Let IDSMapSeriesOptions2_ClipData(RHS As Long)
'Added 6/18/03 to support cross hatching of clip
  m_lClipData2 = RHS
End Property

Private Property Get IDSMapSeriesOptions2_ClipData() As Long
'Added 6/18/03 to support cross hatching of clip
  IDSMapSeriesOptions2_ClipData = m_lClipData2
End Property

Private Sub IDSMapSeriesProps_AddLayerToSuppress(ByVal LayerName As String)
  m_SuppressColl.Add LayerName
End Sub

Private Property Let IDSMapSeriesProps_DataFrameName(RHS As String)
  m_sDataFrameName = RHS
End Property

Private Property Get IDSMapSeriesProps_DataFrameName() As String
  IDSMapSeriesProps_DataFrameName = m_sDataFrameName
End Property

Private Property Let IDSMapSeriesProps_IndexFieldName(RHS As String)
  m_sIndexFieldName = RHS
End Property

Private Property Get IDSMapSeriesProps_IndexFieldName() As String
  IDSMapSeriesProps_IndexFieldName = m_sIndexFieldName
End Property

Private Property Let IDSMapSeriesProps_IndexLayerName(RHS As String)
  m_sIndexLayerName = RHS
End Property

Private Property Get IDSMapSeriesProps_IndexLayerName() As String
  IDSMapSeriesProps_IndexLayerName = m_sIndexLayerName
End Property

Private Sub IDSMapSeriesProps_RemoveLayerToSuppress(Index As Long)
  If Index > -1 And Index < m_SuppressColl.Count Then
    m_SuppressColl.Remove Index + 1
  End If
End Sub

Private Property Get IDSMapSeriesProps_SuppressLayer(Index As Long) As String
  If Index > -1 And Index < m_SuppressColl.Count Then
    IDSMapSeriesProps_SuppressLayer = m_SuppressColl.Item(Index + 1)
  Else
    IDSMapSeriesProps_SuppressLayer = ""
  End If
End Property

Private Property Get IDSMapSeriesProps_SuppressLayerCount() As Long
  IDSMapSeriesProps_SuppressLayerCount = m_SuppressColl.Count
End Property

Private Property Let IDSMapSeriesProps_SuppressLayers(ByVal RHS As Boolean)
  m_bSupressLayers = RHS
End Property

Private Property Get IDSMapSeriesProps_SuppressLayers() As Boolean
  IDSMapSeriesProps_SuppressLayers = m_bSupressLayers
End Property

Private Property Let IDSMapSeriesProps_TileSelectionMethod(RHS As Long)
  m_lTileSelection = RHS
End Property

Private Property Get IDSMapSeriesProps_TileSelectionMethod() As Long
  IDSMapSeriesProps_TileSelectionMethod = m_lTileSelection
End Property

Private Property Get IPersistVariant_ID() As esriCore.IUID
  Dim id As New UID
  id = "DSMapBookPrj.DSMapSeries"
  Set IPersistVariant_ID = id
End Property

Private Sub IPersistVariant_Load(ByVal Stream As esriCore.IVariantStream)
'Load the persisted parameters of the renderer
On Error GoTo ErrHand:
  Dim lLoop As Long, lCount As Long, pPage As IDSMapPage, vClip As Variant
  Dim bClip As Boolean
    
  m_bEnableSeries = Stream.Read
  
  'Added 6/18/03 to support cross hatching of area outside clip
  vClip = Stream.Read
  If UCase(TypeName(vClip)) = "BOOLEAN" Then
    bClip = CBool(vClip)
    If bClip Then
      m_lClipData2 = 1
    Else
      m_lClipData2 = 0
    End If
  Else
    m_lClipData2 = CLng(vClip)
  End If
'  m_bClipData = Stream.Read

  m_sDataDrivenField = Stream.Read
  m_lExtentType = Stream.Read
  m_dFixedScale = Stream.Read
  m_bLabelNeighbors = Stream.Read
  Set m_pLabelSymbol = Stream.Read
  m_dMargin = Stream.Read
  m_sMarginType = Stream.Read
  m_bRotateFrame = Stream.Read
  m_sRotationField = Stream.Read
  m_sDataFrameName = Stream.Read
  m_sIndexFieldName = Stream.Read
  m_sIndexLayerName = Stream.Read
  m_bSupressLayers = Stream.Read
  m_lTileSelection = Stream.Read
  lCount = Stream.Read
  If lCount > 0 Then
    For lLoop = 1 To lCount
      m_SuppressColl.Add Stream.Read
    Next lLoop
  End If
  
  lCount = Stream.Read
  If lCount > 0 Then
    For lLoop = 1 To lCount
      Set pPage = Stream.Read
      m_PageColl.Add pPage
    Next lLoop
  End If
    
  Exit Sub
ErrHand:
  MsgBox "MapSeries - IPersistVariant_Load - " & Err.Description
End Sub

Private Sub IPersistVariant_Save(ByVal Stream As esriCore.IVariantStream)
'Write it all out
On Error GoTo ErrHand:
  Dim lLoop As Long
    
  Stream.Write m_bEnableSeries
  
  'Added 6/18/03 to support cross hatching of area outside the clip
  Stream.Write m_lClipData2
'  Stream.Write m_bClipData
  Stream.Write m_sDataDrivenField
  Stream.Write m_lExtentType
  Stream.Write m_dFixedScale
  Stream.Write m_bLabelNeighbors
  Stream.Write m_pLabelSymbol
  Stream.Write m_dMargin
  Stream.Write m_sMarginType
  Stream.Write m_bRotateFrame
  Stream.Write m_sRotationField
  Stream.Write m_sDataFrameName
  Stream.Write m_sIndexFieldName
  Stream.Write m_sIndexLayerName
  Stream.Write m_bSupressLayers
  Stream.Write m_lTileSelection

On Error GoTo ErrHand2:
  Stream.Write m_SuppressColl.Count
  If m_SuppressColl.Count > 0 Then
    For lLoop = 1 To m_SuppressColl.Count
      Stream.Write m_SuppressColl.Item(lLoop)
    Next lLoop
  End If
  Stream.Write m_PageColl.Count
  If m_PageColl.Count > 0 Then
    For lLoop = 1 To m_PageColl.Count
      Stream.Write m_PageColl.Item(lLoop)
    Next lLoop
  End If
    
  Exit Sub
ErrHand:
  MsgBox "MapSeries - IPersistVariant_Save - " & Err.Description
ErrHand2:
  MsgBox "MapSeriesCOLLS - IPersistVariant_Save - " & Err.Description
End Sub
