VERSION 5.00
Begin VB.Form frmExportPropDlg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   3510
   ClientLeft      =   3225
   ClientTop       =   2925
   ClientWidth     =   5460
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4410
      TabIndex        =   1
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   315
      Left            =   3330
      TabIndex        =   0
      Top             =   3120
      Width           =   975
   End
End
Attribute VB_Name = "frmExportPropDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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





''''''''''''''''''''''''''''''''''''''''
' frmExportPropDlg
'
' References: ESRI Output Object Library
'             ESRI OutputUI Object Library
'             ESRI System Object Library
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
   
Private Declare Function ProgIDFromCLSID Lib "ole32.dll" (pCLSID As Any, _
    lpszProgID As Long) As Long
Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpszProgID As _
    Long, pCLSID As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As _
    Any, source As Any, ByVal bytes As Long)
 
Private Const SW_SHOWNORMAL = 1
Private Const WM_DESTROY = &H2

Private m_PropDlgHWnd As Long
Private m_lWinStyle As Long
Private m_pExport As IExport
Private m_pExportPropertiesDialog As IExportPropertiesDialog
Private m_pSettingsInRegistry As ISettingsInRegistry

Public Property Set Export(pExport As IExport)
57:   Set m_pExport = pExport
58:   Set m_pSettingsInRegistry = m_pExport
End Property

Private Sub Form_Load()
  
  'Restore the exporter object's properties to the settings last saved in the registry.
64:   m_pSettingsInRegistry.RestoreForCurrentUser "Software\ESRI\Export\ExportObjectsParams"
  
  'Initialize the m_pExportPropertiesDialog pointer to the proper dialog.  The InitExportPropDlg()
  ' function will dynamically create an ExportPropertiesDialog appropriate for the current export
  ' object.  It creates the dialog based on the contents of the "ESRI Export Properties Dialogs"
  ' component category, and uses an ExportPropertiesRouter to provide the maximum flexibility when
  ' using a plug-in exporter object.  If you are using a default ESRI Export object and have not
  ' created a custom UI CoClass for it, this single line will perform the same as calling
  ' the InitExportPropDlg function:
  '       Set m_pExportPropertiesDialog = New DefaultExportPropertiesDialog
74:   Set m_pExportPropertiesDialog = InitExportPropDlg(m_pExport)
  
76:   Me.Caption = m_pExport.Name & " Export Properties"
  
  'Tell the ExportPropertiesDialog object to control the exporter object.
79:   m_pExportPropertiesDialog.SetObject m_pExport
  
  'Assign the HWnd of m_pExportPropertiesDialog to a global variable so the Win32 API
  ' functions can see it.  This line is actually calling a "get_property" inside the
  ' ExportPropertiesDialog, which internally creates the dialog at this point.
84:   m_PropDlgHWnd = m_pExportPropertiesDialog.hwnd
  
  'Use the Win32 API SetParent function to set frmExportPropDlg as the parent
  ' dialog of the pExportPropertiesDialog.
88:   SetParent m_PropDlgHWnd, Me.hwnd
  
  'Use the Win32 API ShowWindow funtion to show ExportPropertiesDialog with a style of SW_SHOWNORMAL
91:   ShowWindow m_PropDlgHWnd, SW_SHOWNORMAL
  
End Sub

Private Sub cmdOK_Click()
96:   m_pSettingsInRegistry.StoreForCurrentUser "Software\ESRI\Export\ExportObjectsParams"
97:   CloseWindow m_PropDlgHWnd
98:   Unload Me
End Sub

Private Sub cmdCancel_Click()
102:   m_pSettingsInRegistry.RestoreForCurrentUser "Software\ESRI\Export\ExportObjectsParams"
103:   CloseWindow m_PropDlgHWnd
104:   Unload Me
End Sub

Private Function InitExportPropDlg(pExport As IExport) As IExportPropertiesDialog

  Dim pExportPropertiesDialog As IExportPropertiesDialog
  Dim esriExportPropsDlgsCat As New UID
  Dim pRouter As IExportPropertiesRouter
  Dim sProgId As String
  
  'Use a Category Factory object to create one instance of every class registered
  ' in the "ESRI Export Properties Dialogs" category.  As we loop through the components,
  ' Use the boolean returned by the IsValidObject method to check if the current object
  ' is the correct dialog for the currently selected exporter object.
   'Component Category: "ESRI Export Properties Dialogs" = {AE54680B-8099-4A93-8C29-6D727FBCF11A}
119:   esriExportPropsDlgsCat.value = "{AE54680B-8099-4A93-8C29-6D727FBCF11A}"
  Dim pCategoryFactory As ICategoryFactory
121:   Set pCategoryFactory = New CategoryFactory
122:   pCategoryFactory.CategoryID = esriExportPropsDlgsCat
  
124:   Set pExportPropertiesDialog = pCategoryFactory.CreateNext
125:   Do While Not pExportPropertiesDialog Is Nothing
126:     If pExportPropertiesDialog.IsValidObject(pExport) Then Exit Do
127:     Set pExportPropertiesDialog = pCategoryFactory.CreateNext
128:   Loop

  'Use the ExportPropertiesRouter object to provide the guid of an ExportPropertiesDialog
  ' that is pre-populated with ExportPropertiesPageDialog objects, which appear as tabbed
  ' panels in the dialog.  Use the guid provided by the router to build a ProgId from that
  ' name, and create the object. If we don't use the router to create this dialog object,
  ' the dialog we created above will have no ExportPropertiesPageDialog and will give an
  ' error on display.
136:   Set pRouter = pExportPropertiesDialog
137:   sProgId = CLSIDToProgID(pRouter.ExportPropertiesClass.value)
138:   Set pExportPropertiesDialog = CreateObject(sProgId)
  
140:   Set pRouter = Nothing
141:   Set esriExportPropsDlgsCat = Nothing
142:   Set InitExportPropDlg = pExportPropertiesDialog
  
End Function

Private Function WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
' Defines the Window Procedure for ExportPropertiesDialog and listens for the
' WM_DESTROY message to cleanly destroy the ExportPropertiesDialog.
  Select Case uMsg&
    Case WM_DESTROY:
151:       Call PostQuitMessage(0&)
152:   End Select
153:   WndProc = DefWindowProc(hwnd&, uMsg&, wParam&, lParam&)
End Function

Private Function CloseWindow(hwnd As Long) As Boolean
' Win32 API Cleanup function to cleanly destroy the ExportPropertiesDialog.
  Dim lSuccess As Long
159:   lSuccess = DestroyWindow(hwnd)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CLSIDToProgID
' Convert a string representation of a CLSID, including the
' surrounding brace brackets, into the corresponding ProgID.
'
' usage example:
'    Debug.Print CLSIDToProgID("{179CD501-DF82-4EE5-8D32-7235D2414DE9}")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CLSIDToProgID(ByVal CLSID As String) As String
    Dim pResult As Long, pChar As Long
    Dim char As Integer, length As Long

    Dim guid(15) As Byte

    ' convert from string to a binary CLSID
177:     CLSIDFromString StrPtr(CLSID), guid(0)
    ' convert to a string, get pointer to result
179:     ProgIDFromCLSID guid(0), pResult
    ' return a null string if not found
    If pResult = 0 Then Exit Function

    ' find the terminating null char
184:     pChar = pResult - 2
185:     Do
186:         pChar = pChar + 2
187:         CopyMemory char, ByVal pChar, 2
188:     Loop While char
    ' now get the entire string in one operation
190:     length = pChar - pResult
    ' no need for a temporary string
192:     CLSIDToProgID = Space$(length \ 2)
193:     CopyMemory ByVal StrPtr(CLSIDToProgID), ByVal pResult, length
End Function




