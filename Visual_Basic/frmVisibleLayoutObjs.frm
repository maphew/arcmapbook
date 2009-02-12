VERSION 5.00
Begin VB.Form frmVisibleElements 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Set Visible Objects"
   ClientHeight    =   4284
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   5052
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4284
   ScaleWidth      =   5052
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdUntag 
      Caption         =   "&Untag"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   1335
   End
   Begin VB.ListBox lstVisibleObjects 
      Height          =   2424
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   720
      Width           =   4815
   End
   Begin VB.Label lblMapPage 
      Caption         =   "lblMapPage"
      Height          =   495
      Left            =   2760
      TabIndex        =   6
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Mark those objects that should be visible on this map page."
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Untagging a layout object with Map Book will make that object visible on all map pages."
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   3360
      Width           =   3255
   End
End
Attribute VB_Name = "frmVisibleElements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_pApp As IApplication
Private m_pNWSeriesOpts As INWMapSeriesOptions
Private m_sMainDataFrame As String
Private m_sMapPageID As String
Const c_sModuleFileName As String = "frmVisibleLayoutObjs.frm"




Public Property Set App(pApp As IApplication)
  Set m_pApp = pApp
End Property

Public Property Get App() As IApplication
  Set App = m_pApp
End Property



Public Sub Initialize(pApp As IApplication, _
                      pNWSeriesOpts As INWMapSeriesOptions, _
                      sMapPageID As String)
  
  Dim pNWDSMapSeriesProps As INWDSMapSeriesProps

  Set m_pApp = pApp
  Set m_pNWSeriesOpts = pNWSeriesOpts
  m_sMapPageID = sMapPageID
  Set pNWDSMapSeriesProps = m_pNWSeriesOpts
  m_sMainDataFrame = pNWDSMapSeriesProps.DataFrameName
  
End Sub

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdOK_Click()
  On Error GoTo ErrorHandler

  
  Dim sInvisElems() As String, lInvisElemCount As Long, i As Long
  
  If m_pNWSeriesOpts Is Nothing Then
    MsgBox "A data structure necessary for applying the settings of this" & vbNewLine _
         & "form has unexpectantly become empty.  Please note the steps required" & vbNewLine _
         & "to trigger this error message.", vbCritical, "m_pNWSeriesOpts is ''Nothing'' error in cmdOK_Click ..."
    Exit Sub
  End If
  
  'build the list of layout elements that are
  '*not* visible with this map page.
  ''''''''''''''''''''''''''''''''''
  lInvisElemCount = 0
  ReDim sInvisElems(0) As String 'UBound(sInvisElems) fails with 'subscript out of range' if there was never a redim
  With lstVisibleObjects
    For i = 0 To (.ListCount - 1)
      If Not .Selected(i) Then
        lInvisElemCount = lInvisElemCount + 1
      End If
    Next i
    ReDim sInvisElems(lInvisElemCount) As String
    lInvisElemCount = 0
    For i = 0 To (.ListCount - 1)
      If Not .Selected(i) Then
        sInvisElems(lInvisElemCount) = .List(i)
        lInvisElemCount = lInvisElemCount + 1
      End If
    Next i
  End With
  
  
  'assign the element invisibility for this map page
  ''''''''''''''''''''''''''''''''''''''''''''''''''
  m_pNWSeriesOpts.ElementsSetElementsInvisibleForMapPage sInvisElems, m_sMapPageID

  Me.Hide
  


  Exit Sub
ErrorHandler:
  HandleError True, "cmdOK_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdUntag_Click()
  On Error GoTo ErrorHandler

  Dim sElement As String, pElement As IElement, i As Long, pElemProps As IElementProperties
  Dim vCustomProp As Variant
  
  'be sure that the user has selected one and only one element
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  If Me.lstVisibleObjects.SelCount <> 1 Then
    MsgBox "Exactly one element must be selected for untagging."
    Exit Sub
  End If
  
  'untag the selected string
  'remove the selected string from the interface
  ''''''''''''''''''''''''''''''''''''''''''''''
  If m_pNWSeriesOpts Is Nothing Then Exit Sub
  With lstVisibleObjects
    For i = 0 To (.ListCount - 1)
      If .Selected(i) Then 'invalid index error
        sElement = .List(i)
        Set pElement = GetNamedElement(sElement)
        If Not pElement Is Nothing Then
          Set pElemProps = pElement
          m_pNWSeriesOpts.ElementsUntagElementString sElement
          pElemProps.CustomProperty = vCustomProp
          .RemoveItem (i) 'important now to no longer loop
          i = .ListCount
        End If 'not pElement is nothing
      End If '.Selected(i)
    Next i
  End With

  Exit Sub
ErrorHandler:
  HandleError True, "cmdUntag_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub




Private Sub Form_Activate()
  On Error GoTo ErrorHandler

  Dim vElements As Variant, sElemName As String, lElemCount As Long
  Dim i As Long, lElemIdx As Long
  
  lstVisibleObjects.Clear
  If m_pNWSeriesOpts Is Nothing Then Exit Sub
  If Len(Trim$(m_sMapPageID)) = 0 Then
    MsgBox "An empty map page name was specified.  The list of elements will not be " & vbNewLine _
         & "available without specifying a map page name." & vbNewLine, vbOKOnly, "Empty Map Page Name"
    Exit Sub
  End If
  
  lblMapPage.Caption = "Map Page: " & m_sMapPageID
  
  'acquire the list of elements being tracked
  '''''''''''''''''''''''''''''''''''''''''''
  vElements = m_pNWSeriesOpts.ElementsGetTaggedElements
  If IsEmpty(vElements) Then
    Exit Sub
  End If
  lElemCount = UBound(vElements) + 1
  For i = 0 To (lElemCount - 1)
    sElemName = vElements(i)
    lstVisibleObjects.AddItem sElemName
  Next i
  
  
  'clear selections for all object listings
  '''''''''''''''''''''''''''''''''''''''''
  For i = 0 To (lstVisibleObjects.ListCount - 1)
    lstVisibleObjects.Selected(i) = True
  Next i
  
  'set as selected those elements that
  'are visible on this map page.
  ''''''''''''''''''''''''''''''''''''
  vElements = m_pNWSeriesOpts.ElementsGetElementsInvisibleForMapPage(m_sMapPageID)
  If IsEmpty(vElements) Then
    'all elements will remain visible
    Exit Sub
  End If
  
  lElemCount = UBound(vElements) + 1 'subscript out of range
  For i = 0 To (lElemCount - 1)
    sElemName = vElements(i)
    lElemIdx = FindControlString(lstVisibleObjects, sElemName, -1, True)
    If lElemIdx > -1 Then
      lstVisibleObjects.Selected(lElemIdx) = False
    End If
  Next i
  
  
  Exit Sub
ErrorHandler:
  HandleError True, "Form_Activate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub






'GetNamedElement
'
'This function returns the element referenced by the string passed.
'This function first checks for the element in storage, in case it
'has already been removed from the PageLayout as part of layout element
'visibility functionality.  It then searches through the PageLayout for
'an element with the matching name.
'-------------------------------
Private Function GetNamedElement(sElemName As String) As IElement
  On Error GoTo ErrorHandler

  Dim pElement As IElement, pPageLayout As IPageLayout
  Dim pGraphicsContSelect As IGraphicsContainerSelect, pMxDoc As IMxDocument
  Dim pElemProps As IElementProperties, vCustProp As Variant, sCustProp As String
  Dim pLoopElement As IElement, pGraphicsCont As IGraphicsContainer
  
  'first check in storage
  '''''''''''''''''''''''
  Set pElement = Nothing
  If Not m_pNWSeriesOpts Is Nothing Then
    If m_pNWSeriesOpts.ElementsElementIsInStorage(sElemName) Then
      Set pElement = m_pNWSeriesOpts.ElementsStoredElement(sElemName)
    End If
  End If
  
  'if not in storage, check in the pagelayout
  '''''''''''''''''''''''''''''''''''''''''''
  If pElement Is Nothing Then
    If Not m_pApp Is Nothing Then
      Set pMxDoc = m_pApp.Document
      Set pPageLayout = pMxDoc.PageLayout
      Set pGraphicsCont = pPageLayout
      pGraphicsCont.Reset
      Set pLoopElement = pGraphicsCont.Next
      Do While (Not pLoopElement Is Nothing) And (pElement Is Nothing)
        Set pElemProps = pLoopElement
        vCustProp = pElemProps.CustomProperty
        If Not IsEmpty(vCustProp) Then
          If StrComp(TypeName(vCustProp), "string", vbTextCompare) = 0 Then
            sCustProp = vCustProp
            If StrComp(sCustProp, sElemName, vbTextCompare) = 0 Then
              Set pElement = pLoopElement
            End If
          End If
        End If
        Set pLoopElement = pGraphicsCont.Next
      Loop
    End If 'Not m_pApp is nothing
  End If 'not pElement is nothing
  Set GetNamedElement = pElement
  

  Exit Function
ErrorHandler:
  HandleError False, "GetNamedElement " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Sub Form_Unload(Cancel As Integer)
  lstVisibleObjects.Clear
  Set m_pApp = Nothing
  Set m_pNWSeriesOpts = Nothing
End Sub


