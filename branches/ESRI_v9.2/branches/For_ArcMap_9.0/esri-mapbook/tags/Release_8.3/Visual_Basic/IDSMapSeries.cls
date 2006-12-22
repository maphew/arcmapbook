VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IDSMapSeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get PageCount() As Long

End Property

Public Property Get Page(Index As Long) As IDSMapPage

End Property

Public Property Let EnableSeries(ByVal vdata As Boolean)

End Property

Public Property Get EnableSeries() As Boolean

End Property

Public Sub AddPage(ByVal Page As IDSMapPage)

End Sub

Public Sub RemovePage(Index As Long)

End Sub

