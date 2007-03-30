VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IDSMapBook"
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

Public Property Get ContentCount() As Long

End Property

Public Property Get ContentItem(Index As Long) As Object

End Property

Public Property Let EnableBook(ByVal vdata As Boolean)

End Property

Public Property Get EnableBook() As Boolean

End Property

Public Sub AddContent(ByVal Content As Object)

End Sub

Public Sub RemoveContent(Index As Long)

End Sub
