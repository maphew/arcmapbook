VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMapGridButton"
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

Implements ICommand

Private m_pApp As IApplication

' API call to keep form top most
Private Const GWL_HWNDPARENT = -8
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
  (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Sub Class_Terminate()
25:     Set m_pApp = Nothing
End Sub

Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
29:     ICommand_Bitmap = frmResources.picCreateGrid.Picture.handle
End Property

Private Property Get ICommand_Caption() As String
33:     ICommand_Caption = "Manage Map Grids"
End Property

Private Property Get ICommand_Category() As String
37:     ICommand_Category = "Samples"
End Property

Private Property Get ICommand_Checked() As Boolean

End Property

Private Property Get ICommand_Enabled() As Boolean
    Dim pMx As IMxDocument
    Dim pSR As ISpatialReference
    Dim pActiveView As IActiveView
    ' Enable, if (in Layout Mode) or (using a ProjectedCoordSystem)
    ' - Refer to clsCreateGrids.CalculatePageToMapRatio for reason why
50:     Set pMx = m_pApp.Document
51:     Set pSR = pMx.FocusMap.SpatialReference
52:     Set pActiveView = pMx.ActiveView
53:     ICommand_Enabled = (TypeOf pActiveView Is IPageLayout) Or (TypeOf pSR Is IProjectedCoordinateSystem)
End Property

Private Property Get ICommand_HelpContextID() As Long

End Property

Private Property Get ICommand_HelpFile() As String

End Property

Private Property Get ICommand_Message() As String
65:     ICommand_Message = "Create/Update Map Grid polygon feature classes in ArcMap"
End Property

Private Property Get ICommand_Name() As String
69:     ICommand_Name = "MapGridManager_clsMapGridButton"
End Property

Private Sub ICommand_OnClick()
    Dim frm As New frmGridSettings
    Dim pCreateGrid As clsCreateGrids
    
    On Error GoTo eh
    
78:     Set frm.m_Application = m_pApp
79:     frm.Show vbModal
80:     Set pCreateGrid = frm.GridSettings
81:     If pCreateGrid Is Nothing Then
        Exit Sub
83:     End If
84:     pCreateGrid.GenerateGrids2 m_pApp
    ' Clean up
86:     Unload frm
87:     Set frm = Nothing
88:     Set pCreateGrid = Nothing
    
    Exit Sub
eh:
92:     Unload frm
93:     Set frm = Nothing
94:     Set pCreateGrid = Nothing
95:     If Err.Number <> 364 Then
96:         MsgBox "Error in clsMapGridButton_OnClick.  Error " & Err.Number & ": " & Erl & " - " & Err.Description
97:     End If
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
101:     Set m_pApp = hook
End Sub

Private Property Get ICommand_Tooltip() As String
105:     ICommand_Tooltip = "Create/Update Map Grids"
End Property
