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
Option Explicit

Dim m_sPageName As String
Dim m_dRotation As Double
Dim m_dScale As Double
Dim m_dShape As IPolygon

Public Property Get PageName() As String
  PageName = m_sPageName
End Property

Public Property Let PageName(RHS As String)
  m_sPageName = RHS
End Property

Public Property Get PageRotation() As Double
  PageRotation = m_dRotation
End Property

Public Property Let PageRotation(RHS As Double)
  m_dRotation = RHS
End Property

Public Property Get PageScale() As Double
  PageScale = m_dScale
End Property

Public Property Let PageScale(RHS As Double)
  m_dScale = RHS
End Property

Public Property Get PageShape() As IPolygon
  Set PageShape = m_dShape
End Property

Public Property Set PageShape(RHS As IPolygon)
  Set m_dShape = RHS
End Property
