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
Option Explicit

Implements ICommand

Dim m_pApp As IApplication

Private Property Get ICommand_Bitmap() As esriCore.OLE_HANDLE
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
  ICommand_Name = "DSMapBookUIPrj.CreateMapBook"
End Property

Private Sub ICommand_OnClick()
On Error GoTo ErrHand:
  Dim pMapBook As IDSMapBook
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
ErrHand:
  MsgBox "CreateMapBook_OnClick - " & Err.Description
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
  Set m_pApp = hook
End Sub

Private Property Get ICommand_Tooltip() As String
  ICommand_Tooltip = "Create Map Book"
End Property
