VERSION 5.00
Begin VB.Form frmPagesWhereElemIsVisible 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Page Visibility of Element"
   ClientHeight    =   8772
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   5172
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8772
   ScaleWidth      =   5172
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDebug 
      Caption         =   "Debug"
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   7800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   8280
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   8280
      Width           =   1335
   End
   Begin VB.ListBox lstMapPagesWhereVisible 
      Height          =   6960
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   600
      Width           =   4935
   End
   Begin VB.CommandButton cmdUncheckAll 
      Caption         =   "&Uncheck All"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton cmdCheckAll 
      Caption         =   "Check &All"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   7800
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Pages in map series where this layout object is visible.  Uncheck those map pages where this object should not appear."
      Height          =   480
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmPagesWhereElemIsVisible"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_pApp As IApplication
Private m_pNWSeriesOpts As INWMapSeriesOptions
Private m_sMainMapFrame As String
Private m_pElement As IElement

Const c_sModuleFileName As String = "frmPagesWhereVisible.frm"


Public Sub Initialize(pApp As IApplication, _
                      pNWSeriesOpts As INWMapSeriesOptions, _
                      pElement As IElement)
  On Error GoTo ErrorHandler

16:   Set m_pApp = pApp
17:   Set m_pNWSeriesOpts = pNWSeriesOpts
18:   Set m_pElement = pElement
  Exit Sub
ErrorHandler:
  HandleError True, "Initialize " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub





Private Sub cmdCancel_Click()
29:   Me.Hide
End Sub

Private Sub cmdCheckAll_Click()
  Dim lPageCount As Long, i As Long
  
35:   With lstMapPagesWhereVisible
36:     lPageCount = .ListCount
37:     For i = 0 To (lPageCount - 1)
38:       .Selected(i) = True
39:     Next i
40:   End With
End Sub

Private Sub cmdDebug_Click()
  On Error GoTo ErrorHandler

  Dim vMapPages As Variant, sMapPage As String, sElemName As String
  Dim pElemProps As IElementProperties, vCustProp As Variant
  Dim lInvisPageCount As Long, i As Long, sOutput As String
  
  If m_pNWSeriesOpts Is Nothing Then Exit Sub
  If m_pElement Is Nothing Then Exit Sub
52:   Set pElemProps = m_pElement
53:   vCustProp = pElemProps.CustomProperty
54:   If IsEmpty(vCustProp) Then
55:     MsgBox "vCustProp is empty."
    Exit Sub
57:   End If
58:   If Not StrComp(TypeName(vCustProp), "string", vbTextCompare) = 0 Then
59:     MsgBox "vCustProp isn't a string."
    Exit Sub
61:   End If
62:   sElemName = vCustProp
63:   vMapPages = m_pNWSeriesOpts.ElementsGetMapPagesWhereInvisible(m_pElement)
64:   lInvisPageCount = UBound(vMapPages)
65:   sOutput = ""
66:   For i = 0 To (lInvisPageCount)
67:     sOutput = sOutput & "''" & vMapPages(i) & "''" & vbNewLine
68:   Next i
  
70:   MsgBox "Map pages where element " & sElemName & " is not visible are " & vbNewLine & sOutput
  
  

  Exit Sub
ErrorHandler:
  HandleError True, "cmdDebug_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdOK_Click()
  On Error GoTo ErrorHandler

  Dim sMapPageIDs() As String, i As Long, lNonSelCount As Long
  Dim lInvisPageCount As Long
  
  'acquire the list of map pages that are
  'not selected (and therefore in which this
  'element is not visible).
  
89:   With lstMapPagesWhereVisible
90:     lNonSelCount = 0
91:     For i = 0 To (.ListCount - 1)
92:       If Not .Selected(i) Then
93:         lNonSelCount = lNonSelCount + 1
94:       End If
95:     Next i
    ReDim sMapPageIDs(lNonSelCount)
97:     lInvisPageCount = 0
98:     For i = 0 To (.ListCount - 1)
99:       If Not .Selected(i) Then
100:         sMapPageIDs(lInvisPageCount) = .List(i)
101:         lInvisPageCount = lInvisPageCount + 1
102:       End If
103:     Next i
104:   End With
     
106:   If lNonSelCount > 0 Then
107:     m_pNWSeriesOpts.ElementsSetMapPagesWhereInvisible sMapPageIDs, m_pElement
108:   End If

110:   Me.Hide
  
  Exit Sub
ErrorHandler:
  HandleError True, "cmdOK_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdUncheckAll_Click()
  Dim lPageCount As Long, i As Long
  
120:   With lstMapPagesWhereVisible
121:     lPageCount = .ListCount
122:     For i = 0 To (lPageCount - 1)
123:       .Selected(i) = False
124:     Next i
125:   End With
End Sub

Private Sub Form_Activate()
  On Error GoTo ErrorHandler

131:   SetControlPositions
  
  Dim vMapPages As Variant, lPageCount As Long, i As Long, sMapPage As String
  Dim pNWDSMapSeries As INWDSMapSeries, pNWDSMapPage As INWDSMapPage, lPageIdx As Long
  
  'load list of map pages
  '''''''''''''''''''''''
138:   lstMapPagesWhereVisible.Clear
139:   Set pNWDSMapSeries = m_pNWSeriesOpts
140:   lPageCount = pNWDSMapSeries.PageCount
141:   For i = 0 To (lPageCount - 1)
142:     Set pNWDSMapPage = pNWDSMapSeries.Page(i)
143:     lstMapPagesWhereVisible.AddItem pNWDSMapPage.PageName
144:     lstMapPagesWhereVisible.Selected(i) = True
145:   Next i
  
  
  'make the map page selections match the
  'map pages where this element is visible
  ''''''''''''''''''''''''''''''''''''''''
155:   vMapPages = m_pNWSeriesOpts.ElementsGetMapPagesWhereInvisible(m_pElement)
156:   lPageCount = UBound(vMapPages) + 1
157:   For i = 0 To (lPageCount - 1)
158:     sMapPage = vMapPages(i) 'map page where
160:     lPageIdx = FindControlString(lstMapPagesWhereVisible, sMapPage, -1, True)
161:     If lPageIdx >= 0 Then
162:       lstMapPagesWhereVisible.Selected(lPageIdx) = False
163:     End If
164:   Next i
  

  Exit Sub
ErrorHandler:
  HandleError True, "Form_Activate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub


Private Sub Form_Resize()
230:   SetControlPositions
End Sub



'Controls the positioning and sizing of controls in the form.
'------------------------------
Private Sub SetControlPositions()
  On Error GoTo ErrorHandler

  Dim lEstimatedHeightOfFormTitleBar As Long, lGaps As Long, lMinWidth
241:   lEstimatedHeightOfFormTitleBar = 350
242:   lGaps = 75
243:   lMinWidth = (Label1.Width + (lGaps * 2))
  
  '''''overall form sizing
246:   With frmPagesWhereElemIsVisible
247:     If .Width < lMinWidth Then .Width = lMinWidth
248:     If .Height < 3500 Then .Height = 3500
249:   End With
  
  '''''button positioning
252:   cmdOK.Top = frmPagesWhereElemIsVisible.Height - lEstimatedHeightOfFormTitleBar - cmdOK.Height - lGaps
253:   cmdCheckAll.Top = cmdOK.Top - lGaps - cmdCheckAll.Height
254:   cmdUncheckAll.Top = cmdCheckAll.Top
255:   cmdOK.Left = frmPagesWhereElemIsVisible.Width / 2 - (cmdOK.Width / 2)
256:   cmdCancel.Left = cmdOK.Left + cmdOK.Width + lGaps
257:   cmdCancel.Top = cmdOK.Top
  
  '''''listbox sizing
260:   lstMapPagesWhereVisible.Top = Label1.Top + Label1.Height + 10
261:   lstMapPagesWhereVisible.Height = (cmdCheckAll.Top - lGaps) - lstMapPagesWhereVisible.Top
262:   lstMapPagesWhereVisible.Width = frmPagesWhereElemIsVisible.Width - (4 * lGaps)



  Exit Sub
ErrorHandler:
  HandleError False, "SetControlPositions " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

