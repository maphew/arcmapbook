VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IDSMapPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Copyright 2008 ESRI
' 
' All rights reserved under the copyright laws of the United States
' and applicable international laws, treaties, and conventions.
' 
' You may freely redistribute and use this sample code, with or
' without modification, provided you include the original copyright
' notice and use restrictions.
' 
' See use restrictions at <your ArcGIS install location>/developerkit/userestrictions.txt.
' 




Option Explicit

Public Property Get PageName() As String

End Property

Public Property Let PageName(PageName As String)

End Property

Public Property Get PageNumber() As Long

End Property

Public Property Let PageNumber(PageNumber As Long)

End Property

Public Property Get PageRotation() As Double

End Property

Public Property Let PageRotation(PageRotation As Double)

End Property

Public Property Get PageScale() As Double

End Property

Public Property Let PageScale(PageScale As Double)

End Property

Public Property Get PageShape() As IPolygon

End Property

Public Property Set PageShape(PageShape As IPolygon)

End Property

Public Property Get LastOutputted() As Date

End Property

Public Property Let LastOutputted(LastOutputted As Date)

End Property

Public Property Get EnablePage() As Boolean

End Property

Public Property Let EnablePage(PrintPage As Boolean)

End Property

Public Property Get PageItemCount() As Long

End Property

Public Property Get PageItem(Index As Long) As IElement

End Property

Public Sub AddPageItem(ByVal PageItem As IElement)

End Sub

Public Sub RemovePageItem(Index As Long)

End Sub

Public Sub DrawPage(pDoc As IMxDocument, pDSMapSeries As IDSMapSeries, bRefreshFlag As Boolean)

End Sub

Public Function IndexPage(pIndexLayer As IFeatureLayer, sFieldName As String) As Collection

End Function


