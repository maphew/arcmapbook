VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IDSMapSeriesProps"
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

Public Property Get DataFrameName() As String

End Property

Public Property Let DataFrameName(DataFrame As String)

End Property

Public Property Get IndexLayerName() As String

End Property

Public Property Let IndexLayerName(LayerName As String)

End Property

Public Property Get IndexFieldName() As String

End Property

Public Property Let IndexFieldName(FieldName As String)

End Property

Public Property Let SuppressLayers(ByVal vdata As Boolean)

End Property

Public Property Get SuppressLayers() As Boolean

End Property

Public Property Get SuppressLayerCount() As Long

End Property

Public Property Get SuppressLayer(Index As Long) As String

End Property

Public Sub AddLayerToSuppress(ByVal LayerName As String)

End Sub

Public Sub RemoveLayerToSuppress(Index As Long)

End Sub

Public Property Get TileSelectionMethod() As Long

End Property

Public Property Let TileSelectionMethod(TileSelection As Long)

End Property

Public Property Get StartNumber() As Long

End Property

Public Property Let StartNumber(lStart As Long)

End Property
