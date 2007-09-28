VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IDSMapSeriesOptions"
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

Public Property Get ExtentType() As Long

End Property

Public Property Let ExtentType(ExtentType As Long)

End Property

Public Property Get Margin() As Double

End Property

Public Property Let Margin(Margin As Double)

End Property

Public Property Get MarginType() As String

End Property

Public Property Let MarginType(MarginType As String)

End Property

Public Property Get FixedScale() As Double

End Property

Public Property Let FixedScale(FixedScale As Double)

End Property

Public Property Get DataDrivenField() As String

End Property

Public Property Let DataDrivenField(DataDrivenField As String)

End Property

Public Property Get RotateFrame() As Boolean

End Property

Public Property Let RotateFrame(RotateFrame As Boolean)

End Property

Public Property Get RotationField() As String

End Property

Public Property Let RotationField(RotationField As String)

End Property

Public Property Get ClipData() As Boolean

End Property

Public Property Let ClipData(ClipData As Boolean)

End Property

Public Property Get LabelNeighbors() As Boolean

End Property

Public Property Let LabelNeighbors(LabelNeighbors As Boolean)

End Property

Public Property Get LabelSymbol() As ISymbol

End Property

Public Property Set LabelSymbol(LabelSymbol As ISymbol)

End Property
