VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStripMapButton"
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
13:     Set m_pApp = Nothing
End Sub

Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
17:     ICommand_Bitmap = frmResources.picCreateStripMap.Picture.Handle
End Property

Private Property Get ICommand_Caption() As String
21:     ICommand_Caption = "Strip Map"
End Property

Private Property Get ICommand_Category() As String
25:     ICommand_Category = "Developer Samples"
End Property

Private Property Get ICommand_Checked() As Boolean

End Property

Private Property Get ICommand_Enabled() As Boolean
33:     ICommand_Enabled = True
End Property

Private Property Get ICommand_HelpContextID() As Long

End Property

Private Property Get ICommand_HelpFile() As String

End Property

Private Property Get ICommand_Message() As String
45:     ICommand_Message = "Create Strip Map Grid polygon features in ArcMap"
End Property

Private Property Get ICommand_Name() As String
49:     ICommand_Name = "DSMapBookUIPrj_clsStripMapButton"
End Property

Private Sub ICommand_OnClick()
    Dim frm As New frmSMapSettings
    Dim pCreateSMap As clsCreateStripMap
    
    On Error GoTo eh
    
    ' Get properties
59:     Set frm.m_Application = m_pApp
60:     frm.Show vbModal
61:     Set pCreateSMap = frm.StripMapSettings
62:     If pCreateSMap Is Nothing Then
        Exit Sub
64:     End If
    ' Create strip map
66:     pCreateSMap.GenerateStripMap m_pApp
    ' Clean up
68:     Unload frm
69:     Set frm = Nothing
70:     Set pCreateSMap = Nothing
    
    Exit Sub
eh:
74:     Unload frm
75:     Set frm = Nothing
76:     Set pCreateSMap = Nothing
77:     If Err.Number <> 364 Then
78:         MsgBox "Error in clsMapGridButton_OnClick.  Error " & Err.Number & ": " & Err.Description
79:     End If
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
83:     Set m_pApp = hook
End Sub

Private Property Get ICommand_Tooltip() As String
87:     ICommand_Tooltip = "Create StripMap Grids"
End Property
