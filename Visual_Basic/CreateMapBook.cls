VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CreateMapBook"
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

Dim m_pApp As IApplication
' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "CreateMapBook.cls"


Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
  ICommand_Bitmap = frmResources.picBook.Picture.Handle
End Property

Private Property Get ICommand_Caption() As String
  ICommand_Caption = "Create Map Book"
End Property

Private Property Get ICommand_Category() As String
  ICommand_Category = "Developer Samples"
End Property

Private Property Get ICommand_Checked() As Boolean
  ICommand_Checked = False
End Property

Private Property Get ICommand_Enabled() As Boolean
  ICommand_Enabled = True
End Property

Private Property Get ICommand_HelpContextID() As Long

End Property

Private Property Get ICommand_HelpFile() As String

End Property

Private Property Get ICommand_Message() As String
  ICommand_Message = "Map Book Creation"
End Property

Private Property Get ICommand_Name() As String
  ICommand_Name = "NWMapBookUIPrj.CreateMapBook"
End Property

Private Sub ICommand_OnClick()
  On Error GoTo ErrorHandler

  Dim pMapBook As INWDSMapBook
  'Check to see if a MapSeries already exists
  Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
  If pMapBook.ContentCount > 0 Then
    MsgBox "You must remove the existing Map Series before adding another."
    Exit Sub
  End If
  
  'Call the wizard for setting parameters and creating the series
  Set frmMapSeriesWiz.m_pApp = m_pApp
  frmMapSeriesWiz.Show vbModal


  Exit Sub
ErrorHandler:
  HandleError True, "ICommand_OnClick " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
  Set m_pApp = hook
End Sub

Private Property Get ICommand_Tooltip() As String
  ICommand_Tooltip = "Create Map Book"
End Property
