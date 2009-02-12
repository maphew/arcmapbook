VERSION 5.00
Begin VB.Form frmExportPropDlg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   3312
   ClientLeft      =   3228
   ClientTop       =   2928
   ClientWidth     =   5460
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3312
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4410
      TabIndex        =   1
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   315
      Left            =   3330
      TabIndex        =   0
      Top             =   2880
      Width           =   975
   End
End
Attribute VB_Name = "frmExportPropDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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


''''''''''''''''''''''''''''''''''''''''
' frmExportPropDlg
'
' References: ESRI Output Object Library
'             ESRI OutputUI Object Library
'             ESRI System Object Library
'             TypeLib Information
''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Declare Function SetParent Lib "user32" _
  (ByVal hWndChild As Long, _
   ByVal hWndNewParent As Long) As Long
Private Declare Function ShowWindow Lib "user32" _
  (ByVal hwnd As Long, _
   ByVal nCmdShow As Long) As Long
Private Declare Function DestroyWindow Lib "user32" _
  (ByVal hwnd As Long) As Long
Private Declare Sub PostQuitMessage Lib "user32" _
  (ByVal nExitCode As Long)
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" _
  (ByVal hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   ByVal lParam As Long) As Long
   
Private Const SW_SHOWNORMAL = 1
Private Const WM_DESTROY = &H2

Private m_PropDlgHWnd As Long
Private m_lWinStyle As Long
Private m_pExport As IExport
Private m_pExportPropertiesDialog As IExportPropertiesDialog
Private m_pSettingsInRegistry As ISettingsInRegistry

Public Property Set Export(pExport As IExport)
  Set m_pExport = pExport
  Set m_pSettingsInRegistry = m_pExport
End Property

Private Sub Form_Load()
  
  'Restore the exporter object's properties to the settings last saved in the registry.
  m_pSettingsInRegistry.RestoreForCurrentUser "Software\ESRI\Export\ExportObjectsParams"
  
  'Initialize the m_pExportPropertiesDialog pointer to the proper dialog.  The InitExportPropDlg()
  ' function will dynamically create an ExportPropertiesDialog appropriate for the current export
  ' object.  It creates the dialog based on the contents of the "ESRI Export Properties Dialogs"
  ' component category, and uses an ExportPropertiesRouter to provide the maximum flexibility when
  ' using a plug-in exporter object.  If you are using a default ESRI Export object and have not
  ' created a custom UI CoClass for it, this single line will perform the same as calling
  ' the InitExportPropDlg function:
  '       Set m_pExportPropertiesDialog = New DefaultExportPropertiesDialog
  Set m_pExportPropertiesDialog = InitExportPropDlg(m_pExport)
  
  Me.Caption = m_pExport.Name & " Export Properties"
  
  'Tell the ExportPropertiesDialog object to control the exporter object.
  m_pExportPropertiesDialog.SetObject m_pExport
  
  'Assign the HWnd of m_pExportPropertiesDialog to a global variable so the Win32 API
  ' functions can see it.  This line is actually calling a "get_property" inside the
  ' ExportPropertiesDialog, which internally creates the dialog at this point.
  m_PropDlgHWnd = m_pExportPropertiesDialog.hwnd
  
  'Use the Win32 API SetParent function to set frmExportPropDlg as the parent
  ' dialog of the pExportPropertiesDialog.
  SetParent m_PropDlgHWnd, Me.hwnd
  
  'Use the Win32 API ShowWindow funtion to show ExportPropertiesDialog with a style of SW_SHOWNORMAL
  ShowWindow m_PropDlgHWnd, SW_SHOWNORMAL
  
End Sub

Private Sub cmdOK_Click()
  m_pSettingsInRegistry.StoreForCurrentUser "Software\ESRI\Export\ExportObjectsParams"
  CloseWindow m_PropDlgHWnd
  Unload Me
End Sub

Private Sub cmdCancel_Click()
  m_pSettingsInRegistry.RestoreForCurrentUser "Software\ESRI\Export\ExportObjectsParams"
  CloseWindow m_PropDlgHWnd
  Unload Me
End Sub

Private Function InitExportPropDlg(pExport As IExport) As IExportPropertiesDialog

  Dim pExportPropertiesDialog As IExportPropertiesDialog
  Dim esriExportPropsDlgsCat As New UID
  Dim pRouter As IExportPropertiesRouter
  Dim pTypeLibInfo As TLI.TypeLibInfo
  Dim pTypeInfo As Object
  Dim sProgId As String
  
  'Use a Category Factory object to create one instance of every class registered
  ' in the "ESRI Export Properties Dialogs" category.  As we loop through the components,
  ' Use the boolean returned by the IsValidObject method to check if the current object
  ' is the correct dialog for the currently selected exporter object.
   'Component Category: "ESRI Export Properties Dialogs" = {AE54680B-8099-4A93-8C29-6D727FBCF11A}
  esriExportPropsDlgsCat.Value = "{AE54680B-8099-4A93-8C29-6D727FBCF11A}"
  Dim pCategoryFactory As ICategoryFactory
  Set pCategoryFactory = New CategoryFactory
  pCategoryFactory.CategoryID = esriExportPropsDlgsCat
  
  Set pExportPropertiesDialog = pCategoryFactory.CreateNext
  Do While Not pExportPropertiesDialog Is Nothing
    If pExportPropertiesDialog.IsValidObject(pExport) Then Exit Do
    Set pExportPropertiesDialog = pCategoryFactory.CreateNext
  Loop

  'Use the ExportPropertiesRouter object to provide the guid of an ExportPropertiesDialog
  ' that is pre-populated with ExportPropertiesPageDialog objects, which appear as tabbed
  ' panels in the dialog.  Use the guid provided by the router to look up the CoClass name in
  ' the registry, build a ProgId from the the name, and create the object. If we don't use the
  ' router to create this dialog object, the dialog we created above will have no
  ' ExportPropertiesPageDialog and will give an error on display.
  Set pRouter = pExportPropertiesDialog
  Set pTypeLibInfo = TypeLibInfoFromRegistry("{AE064D40-D6CE-11D0-867A-0000F8751720}", 1, 0, 0)
  For Each pTypeInfo In pTypeLibInfo.CoClasses
    If pTypeInfo.Guid = pRouter.ExportPropertiesClass.Value Then
      sProgId = pTypeLibInfo.Name & "." & pTypeInfo.Name
    End If
  Next
  Set pExportPropertiesDialog = CreateObject(sProgId)
  
  Set pRouter = Nothing
  Set esriExportPropsDlgsCat = Nothing
  Set InitExportPropDlg = pExportPropertiesDialog
  
End Function

Private Function WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
' Defines the Window Procedure for ExportPropertiesDialog and listens for the
' WM_DESTROY message to cleanly destroy the ExportPropertiesDialog.
  Select Case uMsg&
    Case WM_DESTROY:
      Call PostQuitMessage(0&)
  End Select
  WndProc = DefWindowProc(hwnd&, uMsg&, wParam&, lParam&)
End Function

Private Function CloseWindow(hwnd As Long) As Boolean
' Win32 API Cleanup function to cleanly destroy the ExportPropertiesDialog.
  Dim lSuccess As Long
  lSuccess = DestroyWindow(hwnd)
End Function
