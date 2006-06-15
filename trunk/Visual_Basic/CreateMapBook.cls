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

Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
8:   ICommand_Bitmap = frmResources.picBook.Picture.Handle
End Property

Private Property Get ICommand_Caption() As String
12:   ICommand_Caption = "Create Map Book"
End Property

Private Property Get ICommand_Category() As String
16:   ICommand_Category = "Developer Samples"
End Property

Private Property Get ICommand_Checked() As Boolean
20:   ICommand_Checked = False
End Property

Private Property Get ICommand_Enabled() As Boolean
24:   ICommand_Enabled = True
End Property

Private Property Get ICommand_HelpContextID() As Long

End Property

Private Property Get ICommand_HelpFile() As String

End Property

Private Property Get ICommand_Message() As String
36:   ICommand_Message = "Map Book Creation"
End Property

Private Property Get ICommand_Name() As String
40:   ICommand_Name = "DSMapBookUIPrj.CreateMapBook"
End Property

Private Sub ICommand_OnClick()
On Error GoTo ErrHand:
  Dim pMapBook As IDSMapBook
  'Check to see if a MapSeries already exists
47:   Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
50:   If pMapBook.ContentCount > 0 Then
51:     MsgBox "You must remove the existing Map Series before adding another."
    Exit Sub
53:   End If
  
  'Call the wizard for setting parameters and creating the series
56:   Set frmMapSeriesWiz.m_pApp = m_pApp
57:   frmMapSeriesWiz.Show vbModal

  Exit Sub
ErrHand:
61:   MsgBox "CreateMapBook_OnClick - " & Err.Description
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
65:   Set m_pApp = hook
End Sub

Private Property Get ICommand_Tooltip() As String
69:   ICommand_Tooltip = "Create Map Book"
End Property
