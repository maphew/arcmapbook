VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tmpPageClass"
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

Dim m_sPageName As String
Dim m_dRotation As Double
Dim m_dScale As Double
Dim m_dShape As IPolygon

Public Property Get PageName() As String
21:   PageName = m_sPageName
End Property

Public Property Let PageName(RHS As String)
25:   m_sPageName = RHS
End Property

Public Property Get PageRotation() As Double
29:   PageRotation = m_dRotation
End Property

Public Property Let PageRotation(RHS As Double)
33:   m_dRotation = RHS
End Property

Public Property Get PageScale() As Double
37:   PageScale = m_dScale
End Property

Public Property Let PageScale(RHS As Double)
41:   m_dScale = RHS
End Property

Public Property Get PageShape() As IPolygon
45:   Set PageShape = m_dShape
End Property

Public Property Set PageShape(RHS As IPolygon)
49:   Set m_dShape = RHS
End Property
