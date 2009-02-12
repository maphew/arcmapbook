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

' Copyright 1995-2004 ESRI

' All rights reserved under the copyright laws of the United States.

' You may freely redistribute and use this sample code, with or without modification.

' Disclaimer: THE SAMPLE CODE IS PROVIDED "AS IS" AND ANY EXPRESS OR IMPLIED
' WARRANTIES, INCLUDING THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS
' FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL ESRI OR
' CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY,
' OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
' SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
' INTERRUPTION) SUSTAINED BY YOU OR A THIRD PARTY, HOWEVER CAUSED AND ON ANY
' THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT ARISING IN ANY
' WAY OUT OF THE USE OF THIS SAMPLE CODE, EVEN IF ADVISED OF THE POSSIBILITY OF
' SUCH DAMAGE.

' For additional information contact: Environmental Systems Research Institute, Inc.

' Attn: Contracts Dept.

' 380 New York Street

' Redlands, California, U.S.A. 92373

' Email: contracts@esri.com

Option Explicit

Implements ICommand

Private m_pApp As IApplication

' API call to keep form top most
Private Const GWL_HWNDPARENT = -8
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
  (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Sub Class_Terminate()
    Set m_pApp = Nothing
End Sub

Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
    ICommand_Bitmap = frmResources.picCreateGrid.Picture.Handle
End Property

Private Property Get ICommand_Caption() As String
    ICommand_Caption = "Manage Map Grids"
End Property

Private Property Get ICommand_Category() As String
    ICommand_Category = "Samples"
End Property

Private Property Get ICommand_Checked() As Boolean

End Property

Private Property Get ICommand_Enabled() As Boolean
    Dim pMx As IMxDocument
    Dim pSR As ISpatialReference
    Dim pActiveView As IActiveView
    ' Enable, if (in Layout Mode) or (using a ProjectedCoordSystem)
    ' - Refer to clsCreateGrids.CalculatePageToMapRatio for reason why
    Set pMx = m_pApp.Document
    Set pSR = pMx.FocusMap.SpatialReference
    Set pActiveView = pMx.ActiveView
    ICommand_Enabled = (TypeOf pActiveView Is IPageLayout) Or (TypeOf pSR Is IProjectedCoordinateSystem)
End Property

Private Property Get ICommand_HelpContextID() As Long

End Property

Private Property Get ICommand_HelpFile() As String

End Property

Private Property Get ICommand_Message() As String
    ICommand_Message = "Create/Update Map Grid polygon feature classes in ArcMap"
End Property

Private Property Get ICommand_Name() As String
    ICommand_Name = "MapGridManager_clsMapGridButton"
End Property

Private Sub ICommand_OnClick()
    Dim frm As New frmGridSettings
    Dim pCreateGrid As clsCreateGrids
    
    On Error GoTo eh
    
    Set frm.m_Application = m_pApp
    frm.Show vbModal
    Set pCreateGrid = frm.GridSettings
    If pCreateGrid Is Nothing Then
        Exit Sub
    End If
    pCreateGrid.GenerateGrids2 m_pApp
    ' Clean up
    Unload frm
    Set frm = Nothing
    Set pCreateGrid = Nothing
    
    Exit Sub
eh:
    Unload frm
    Set frm = Nothing
    Set pCreateGrid = Nothing
    If Err.Number <> 364 Then
        MsgBox "Error in clsMapGridButton_OnClick.  Error " & Err.Number & ": " & Erl & " - " & Err.Description
    End If
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
    Set m_pApp = hook
End Sub

Private Property Get ICommand_Tooltip() As String
    ICommand_Tooltip = "Create/Update Map Grids"
End Property
