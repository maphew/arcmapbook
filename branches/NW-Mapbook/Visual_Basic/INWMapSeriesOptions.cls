VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INWMapSeriesOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Property Let NeighborLabelIndent(RHS As Double)
End Property

Public Property Get NeighborLabelIndent() As Double
End Property


'''text symbol interfaces
Public Sub TextSymbolAdd(pTextSymbol As esriDisplay.ISimpleTextSymbol, sSymbolName As String)
End Sub

Public Sub TextSymbolRemove(sKey As String)
End Sub

Public Property Get TextSymbol(sKey As String) As ISimpleTextSymbol
End Property

Public Sub TextSymbolSet(sKey As String, pTextSymbol As ISimpleTextSymbol)
End Sub

Public Property Get TextSymbolCount() As Long
End Property

Public Property Let TextSymbolDefault(sSymbolName As String)
End Property

Public Property Get TextSymbolDefault() As String
End Property

Public Property Get TextSymbolNames() As Variant
End Property

Public Property Get TextSymbolExists(sSymbolName As String) As Boolean
End Property

Public Sub TextSymbolRename(sPrevName As String, sSymbolName As String)
End Sub


'----------------------------
' Layer visibility methods

Public Property Get LayerGroups() As Variant
End Property

Public Function LayerGroupGet(sGroupName As String) As INWLayerVisibilityGroup
End Function

Public Sub LayerGroupSet(sGroupName As String, pGroup As INWLayerVisibilityGroup)
End Sub

Public Function LayerGroupExists(sGroupName As String) As Boolean
End Function

Public Property Get LayerGroupCount() As Long
End Property



'---------------------------------
' Bubble Detail properties/methods

Public Property Get BubbleLayer() As String
End Property

Public Property Let BubbleLayer(sLyrName As String)
End Property

Public Function IsBubbleLayer(pFeatClass As IFeatureClass) As Boolean
End Function

Public Sub RefreshDetailInsets(pMxDoc As IMxDocument, pBubbleLayer As ILayer, pPageShape As IPolygon, Optional bUseBubbleGeometry As Boolean)
End Sub

Public Sub BubbleGraphicAdd(pInsetElement As IElement, _
                            pArrowElement As IElement, _
                            pArrowShadowElement As IElement, _
                            sInsetName As String)
End Sub

Public Sub BubbleGraphicRemove(pMxDoc As IMxDocument, sInsetName As String)
End Sub

'----------------------------------------
' Identifier map frame properties/methods
Public Sub DataFrameDrawFrames(pMxDoc As IMxDocument, sMapPageID As String, sMainDataFrameName As String)
End Sub

Public Property Get DataFrameIsInStorage(sDataFrameID As String) As Boolean
End Property

Public Function DataFrameIsVisibleInPage(sDataFrameID As String, sMapPageID As String) As Boolean
End Function

Public Sub DataFrameLoadFrames(pMxDoc As IMxDocument)
End Sub

Public Property Get DataFrameMainFrame() As String
End Property

Public Property Let DataFrameMainFrame(sMapFrame As String)
End Property

Public Property Get DataFramesManaged() As Variant
End Property

Public Property Get DataFramePagesWhereVisible(sDataFrameID As String) As Variant
End Property

Public Property Get DataFramePagesWhereVisibleCount(sDataFrameID As String) As Long
End Property

Public Sub DataFrameRemoveFrame(sDataFrameID As String)
End Sub

Public Sub DataFrameRemovePage(sDataFrameID As String, sMapPageID As String)
End Sub

Public Sub DataFrameSetInvisibleInPage(sDataFrameID As String, sMapPageID As String)
End Sub

Public Sub DataFrameSetVisibleInPage(sDataFrameID As String, sMapPageID As String)
End Sub

Public Function DataFrameStoredFrameLayer(sDataFrameID As String, sLayerName As String) As ILayer
End Function

Public Function DataFrameStoredFrame(sDataFrameID As String) As IMapFrame
End Function

Public Function DataFrameStoredFrameLayerList(sDataFrameID As String) As Variant
End Function

Public Property Get DataFramesStored() As Variant
End Property



'Updating Extent of Data Frames other than then main data frame
'when drawing a new map page.
Public Sub DataFrameToUpdateAdd(sDataFrame As String, sMapPagesLayer As String, sPageNameField As String)
End Sub

Public Sub DataFrameToUpdateRemove(sDataFrame As String)
End Sub

Public Function DataFrameToUpdateListDataFrames() As Variant
End Function

Public Function DataFrameToUpdateGetMapPageLayer(sDataFrameName As String) As String
End Function

Public Sub DataFrameToUpdateSetLayerAndField(sDataFrameName As String, sPageNameField As String, sMapPageLayer As String)
End Sub

Public Function DataFrameToUpdateGetPageNameField(sDataFrameName As String, sMapPageLayer As String) As String
End Function

Public Sub DataFrameToUpdateUpdateDataFrames(pMxDoc As IMxDocument, sPageName As String)
End Sub

Public Sub DataFrameToUpdateClearAllDataFrames()
End Sub

Public Sub DataFrameToUpdateClearDataFrame(sDataFrameName As String)
End Sub

Public Sub DataFrameToUpdateHousekeepAbandonedFrames(pMxDoc As IMxDocument)
End Sub

Public Function DataFrameToUpdateIsADataFrameToUpdate(sDataFrameName As String) As Boolean
End Function

Public Property Get DataFrameToUpdateExtentOption() As String
End Property

Public Property Let DataFrameToUpdateExtentOption(sOptionName As String)
End Property





'Dynamic Definition Queries
Public Sub DynamicDefQueryAddLayer(sDataFrameID As String, sLayerID As String)
End Sub

Public Sub DynamicDefQueryRemoveLayer(sDataFrameID As String, sLayerID As String)
End Sub

Public Sub DynamicDefQueryRemoveDataFrame(sDataFrameID As String)
End Sub


'Map Refresh event handler
Public Property Get CurrentPageName() As String
End Property

Public Property Get CurrentPage() As INWDSMapPage
End Property

Public Property Let RefreshEventLoadPage(RHS As Boolean)
End Property
Public Property Get RefreshEventLoadPage() As Boolean
End Property
Public Property Let RefreshEventShouldntRefresh(RHS As Boolean)
End Property
Public Property Get RefreshEventShouldntRefresh() As Boolean
End Property


  'housekeeping must be done at some stage after the user
  'removes a dataframe.  A list of all tracked data frames must
  'be available to find any that no longer exist in the map document,
  'or in storage with the data frame visibility feature.
Public Function DynamicDefQueryDataFrames() As Variant
End Function

Public Property Let DynamicDefQueryReplaceString(RHS As String)
End Property
Public Property Get DynamicDefQueryReplaceString() As String
End Property
Public Property Get DynamicDefQueryPrevReplaceString() As String
End Property
Public Property Get DynamicDefQueryTrackingLayerCount() As Long
End Property
Public Property Get DynamicDefQueryIsTrackingLayer(sDataFrameID As String, sLayerID As String)
End Property


'Layout Element Visibility
Public Function ElementsElementIsInStorage(sElementID As String) As Boolean
End Function
Public Function ElementsElementIsQualified(pElement As IElement) As Boolean
End Function
Public Function ElementsElementIsTagged(pElement As IElement) As Boolean
End Function
Public Function ElementsElementIsVisibleOnMapPage(sMapPageID As String, sElementID As String) As Boolean
End Function
Public Property Get ElementsStored() As Variant
End Property
Public Sub ElementsFilterLayoutElements(pMxDoc As IMxDocument, sMapPageID As String)
End Sub
Public Function ElementsGetElementsInvisibleForMapPage(sMapPageID As String) As Variant
End Function
Public Function ElementsGetMapPagesWhereInvisible(pElement As IElement) As Variant
End Function
Public Function ElementsGetTaggedElements() As Variant
End Function
Public Sub ElementsSetElementsInvisibleForMapPage(sElementIDs() As String, sMapPageID As String)
End Sub
Public Sub ElementsSetMapPagesWhereInvisible(sMapPages() As String, pElement As IElement)
End Sub
Public Function ElementsStoredElement(sElementID As String) As IElement
End Function
Public Sub ElementsStoreElement(pElement As IElement)
End Sub
Public Sub ElementsTagElement(pMxDoc As IMxDocument)
End Sub
Public Sub ElementsUntagElement(pMxDoc As IMxDocument)
End Sub
Public Sub ElementsUntagElementString(sElemName As String)
End Sub
