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

Implements ICommand

Private m_pApp As IApplication

' API call to keep form top most
Private Const GWL_HWNDPARENT = -8
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
  (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Sub Class_Terminate()
13:     Set m_pApp = Nothing
End Sub

Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
17:     ICommand_Bitmap = frmResources.picCreateGrid.Picture.Handle
End Property

Private Property Get ICommand_Caption() As String
21:     ICommand_Caption = "Manage Map Grids"
End Property

Private Property Get ICommand_Category() As String
25:     ICommand_Category = "Samples"
End Property

Private Property Get ICommand_Checked() As Boolean

End Property

Private Property Get ICommand_Enabled() As Boolean
    Dim pMx As IMxDocument
    Dim pSR As ISpatialReference
    Dim pActiveView As IActiveView
    ' Enable, if (in Layout Mode) or (using a ProjectedCoordSystem)
    ' - Refer to clsCreateGrids.CalculatePageToMapRatio for reason why
38:     Set pMx = m_pApp.Document
39:     Set pSR = pMx.FocusMap.SpatialReference
40:     Set pActiveView = pMx.ActiveView
41:     ICommand_Enabled = (TypeOf pActiveView Is IPageLayout) Or (TypeOf pSR Is IProjectedCoordinateSystem)
End Property

Private Property Get ICommand_HelpContextID() As Long

End Property

Private Property Get ICommand_HelpFile() As String

End Property

Private Property Get ICommand_Message() As String
53:     ICommand_Message = "Create/Update Map Grid polygon feature classes in ArcMap"
End Property

Private Property Get ICommand_Name() As String
57:     ICommand_Name = "MapGridManager_clsMapGridButton"
End Property

Private Sub ICommand_OnClick()
    Dim frm As New frmGridSettings
    Dim pCreateGrid As clsCreateGrids
    
    On Error GoTo eh
    
66:     Set frm.m_Application = m_pApp
67:     frm.Show vbModal
68:     Set pCreateGrid = frm.GridSettings
69:     If pCreateGrid Is Nothing Then
        Exit Sub
71:     End If
72:     pCreateGrid.GenerateGrids2 m_pApp
    ' Clean up
74:     Unload frm
75:     Set frm = Nothing
76:     Set pCreateGrid = Nothing
    
    Exit Sub
eh:
80:     Unload frm
81:     Set frm = Nothing
82:     Set pCreateGrid = Nothing
83:     If Err.Number <> 364 Then
84:         MsgBox "Error in clsMapGridButton_OnClick.  Error " & Err.Number & ": " & Erl & " - " & Err.Description
85:     End If
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
89:     Set m_pApp = hook
End Sub

Private Property Get ICommand_Tooltip() As String
93:     ICommand_Tooltip = "Create/Update Map Grids"
End Property
