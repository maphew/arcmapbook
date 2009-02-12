VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INWMapPageAttribs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Adjacent map label symbol
'-------------------------
Public Property Let AdjacentLabelSymbol(ByVal sSymbName As String)
End Property

Public Property Get AdjacentLabelSymbol() As String
End Property




'Invisible Layers per map page
'-----------------------------
Public Property Get InvisibleLayers() As Variant
End Property

Public Sub InvisibleLayerAdd(ByVal sLayerName As String)
End Sub

Public Property Get InvisibleLayerCount() As Long
End Property

Public Property Get LayerVisibilityGroup() As String
End Property

Public Property Let LayerVisibilityGroup(ByVal sLayerName As String)
End Property

Public Sub InvisibleLayerRemove(ByVal sLayerName As String)
End Sub


'Detail Insets per map page
'--------------------------

'Public Property Get DetailInsets() As Variant
'End Property
'
'Public Sub DetailInsetAdd(ByVal pElementWithinBorders As IElement, pMxDoc As IMxDocument, pRow As IRow)
'End Sub
'
'Public Sub DetailInsetDelete(pMxDoc As IMxDocument, pRow As IRow)
'End Sub
'


