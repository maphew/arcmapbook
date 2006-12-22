VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DSMapBookBar"
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

Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'
' Toolbar containing the Split and Combine commands.
'
Option Explicit

Implements IToolBarDef

Private Property Get IToolBarDef_Caption() As String
9:   IToolBarDef_Caption = "Map Series"
End Property

Private Sub IToolBarDef_GetItemInfo(ByVal pos As Long, ByVal itemDef As IItemDef)
  Select Case pos
  Case 0
15:     itemDef.ID = "DSMapBookUIPrj.CreateMapBook"
  Case 1
17:     itemDef.ID = "DSMapBookUIPrj.PageIdentifier"
  Case 2
19:     itemDef.ID = "DSMapBookUIPrj.clsMapGridButton"
  Case 3
21:     itemDef.ID = "DSMapBookUIPrj.clsStripMapButton"
22:   End Select
End Sub

Private Property Get IToolBarDef_ItemCount() As Long
26:   IToolBarDef_ItemCount = 4
End Property

Private Property Get IToolBarDef_name() As String
30:   IToolBarDef_name = "DSMapBookUIPrj.MapSeriesBar"
End Property
