VERSION 5.00
Begin VB.Form frmAdjMapLabelSymbols 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Adjacent Map Label Symbols"
   ClientHeight    =   4284
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4284
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4284
   ScaleWidth      =   4284
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRename 
      Caption         =   "&Rename"
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton cmdProperties 
      Caption         =   "&Properties"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "&Set as Default"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.ListBox lstTextSymbols 
      Height          =   2160
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label lblDefaultSymbol 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblDefaultSymbol"
      Height          =   255
      Left            =   1320
      TabIndex        =   9
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Default Symbol:"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   $"frmAdjMapLabelSymbols.frx":0000
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmAdjMapLabelSymbols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_pApp As IApplication
Private m_pTextSym As ISimpleTextSymbol
                                                  'use a parent form reference
                                                  'to prevent this dialog window
                                                  'from competing with the parent
                                                  'window for screen space.
Private m_pParentForm As Form
Private m_pNWSeriesOptions As INWMapSeriesOptions


Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Const c_sModuleFileName As String = "frmAdjMapLabelSymbols.frm"



Property Set NWSeriesOptions(RHS As INWMapSeriesOptions)
19:   Set m_pNWSeriesOptions = RHS
End Property
Property Get NWSeriesOptions() As INWMapSeriesOptions
22:   Set NWSeriesOptions = m_pNWSeriesOptions
End Property


'Property Set ParentForm(RHS As Form)
'  Set m_pParentForm = RHS
'End Property
'Property Get ParentForm() As Form
'  Set ParentForm = m_pParentForm
'End Property


Property Set Application(pApp As IApplication)
35:   If Not pApp Is Nothing Then
36:     If TypeOf pApp Is IApplication Then
37:       Set m_pApp = pApp
38:     End If
39:   End If
End Property


Property Get TextSymbol() As ISimpleTextSymbol
44:   Set TextSymbol = m_pTextSym
End Property


Private Sub cmdAdd_Click()
  On Error GoTo ErrorHandler

  Dim lNameSuffix As Long, sSymbolName As String, bSymbExists As Boolean
  Dim pTextSymbol As ISimpleTextSymbol, pMxDoc As IMxDocument
                                                  'create a new item in lstTextSymbols
54:   lNameSuffix = 2
55:   sSymbolName = "New Label Symbol"
                                                  'Add a new name, make sure that there is
                                                  'no duplicate of that name
58:   bSymbExists = False
59:   If m_pNWSeriesOptions.TextSymbolExists(sSymbolName) Then
60:     bSymbExists = True
61:   End If
62:   Do While (lNameSuffix < 100) And bSymbExists
63:     If Not (m_pNWSeriesOptions.TextSymbolExists(sSymbolName & lNameSuffix)) Then
64:       bSymbExists = False
65:       sSymbolName = sSymbolName & lNameSuffix
66:     Else
67:       lNameSuffix = lNameSuffix + 1
68:     End If
69:   Loop
70:   If bSymbExists Then
71:     MsgBox "This application does not support more than 99 ''New Label Symbol'' names." & vbNewLine _
         & "Rename some text symbols in order to add more symbol types.", vbOKOnly
    Exit Sub
74:   End If
  
76:   Set pMxDoc = m_pApp.Document
77:   Set pTextSymbol = New TextSymbol
78:   pTextSymbol.Font = pMxDoc.DefaultTextFont
79:   pTextSymbol.Size = pMxDoc.DefaultTextFontSize.Size
80:   m_pNWSeriesOptions.TextSymbolAdd pTextSymbol, sSymbolName
81:   Me.lstTextSymbols.AddItem sSymbolName
  
83:   If m_pNWSeriesOptions.TextSymbolCount = 1 Then
84:     m_pNWSeriesOptions.TextSymbolDefault = sSymbolName
85:     lblDefaultSymbol.Caption = sSymbolName
86:   End If
  

  Exit Sub
ErrorHandler:
  HandleError True, "cmdAdd_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdClose_Click()
  On Error GoTo ErrorHandler
96:   Me.Hide
  Exit Sub
ErrorHandler:
  HandleError True, "cmdClose_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdDefault_Click()
103:   With lstTextSymbols
104:     If .List(.ListIndex) = "" Then
105:       .ListIndex = 0
106:     End If
107:     If m_pNWSeriesOptions.TextSymbolExists(.List(.ListIndex)) Then
108:       m_pNWSeriesOptions.TextSymbolDefault = .List(.ListIndex)
109:       lblDefaultSymbol.Caption = .List(.ListIndex)
110:     Else
111:       MsgBox "Warning, text symbol not found, despite being listed." & vbNewLine _
           & "It is recommended that you remove the text symbol and" & vbNewLine _
           & "recreate it.", vbOKOnly
114:     End If
115:   End With
End Sub

Private Sub cmdDelete_Click()
  Dim sSymbName As String, sDefaultSymb As String, sMessage As String
  Dim i As Integer
  
122:   sDefaultSymb = m_pNWSeriesOptions.TextSymbolDefault
123:   With lstTextSymbols
                                                  'deal with the edge scenario of only one
                                                  'text symbol left
126:     If .ListCount = 1 Then
127:       If Not m_pNWSeriesOptions.TextSymbolExists(.List(.ListIndex)) Then
128:         sMessage = "Warning - last remaining text symbol is not stored." & vbNewLine
129:       End If
130:       If m_pNWSeriesOptions.TextSymbolCount <> 1 Then
131:         sMessage = sMessage & "Warning - 1 entry was present in list of text symbols," & vbNewLine _
                 & "         but the number of stored symbols is " & m_pNWSeriesOptions.TextSymbolCount & vbNewLine
133:       End If
134:       MsgBox sMessage & "At least one text symbol must exist if adjacent map labels are enabled." & vbNewLine _
           & "The command to delete the last remaining text symbol will therefore be" & vbNewLine _
           & "ignored by the application.", vbOKOnly
      Exit Sub
138:     End If
                                                  'if the symbol to be deleted is marked as
                                                  'default, then automatically selected another
                                                  'symbol as default
142:     If StrComp(sDefaultSymb, .List(.ListIndex), vbTextCompare) = 0 Then
                                                  'assumed that there is more than one
                                                  'text symbol
145:       If StrComp(sDefaultSymb, .List(0), vbTextCompare) = 0 Then
                                                  'if the first symbol (the one being deleted)
                                                  'was the default, make the 2nd item the default
148:         sDefaultSymb = .List(1)
149:       Else
150:         sDefaultSymb = .List(0)
151:       End If
152:       If Not m_pNWSeriesOptions.TextSymbolExists(sDefaultSymb) Then
153:         MsgBox "Error, the symbol being deleted is the default symbol." & vbNewLine _
             & "The application was therefore in the process of automatically" & vbNewLine _
             & "selecting a new default text symbol, ''" & sDefaultSymb & "''" & vbNewLine _
             & "when it was detected that this text symbol does not exist" & vbNewLine _
             & "in the application's memory storage.  Please delete the symbol" & vbNewLine _
             & "''" & sDefaultSymb & "''", vbOKOnly
        Exit Sub
160:       End If
      
162:       m_pNWSeriesOptions.TextSymbolDefault = sDefaultSymb
163:       lblDefaultSymbol.Caption = sDefaultSymb
164:     End If
165:     sSymbName = .List(.ListIndex)
166:     If Not m_pNWSeriesOptions.TextSymbolExists(sSymbName) Then
167:       MsgBox "Warning, text symbol did not exist in the underlying data." & vbNewLine _
           & "Closing and restarting ArcMap is recommended.", vbOKOnly
    
170:     End If
171:     .RemoveItem (.ListIndex)
172:     m_pNWSeriesOptions.TextSymbolRemove sSymbName
173:   End With
  
End Sub

Private Sub cmdProperties_Click()
  On Error GoTo ErrorHandler

  Dim bChanged As Boolean, pTextSymEditor As ITextSymbolEditor
  Dim sSymbName As String
  
183:   sSymbName = lstTextSymbols.List(lstTextSymbols.ListIndex)
184:   If sSymbName = "" Then
185:     sSymbName = lstTextSymbols.List(0)
186:   End If
187:   If m_pNWSeriesOptions.TextSymbolExists(sSymbName) Then
188:     Set pTextSymEditor = New TextSymbolEditor
189:     Set m_pTextSym = m_pNWSeriesOptions.TextSymbol(sSymbName)
190:     bChanged = pTextSymEditor.EditTextSymbol(m_pTextSym, m_pApp.hwnd)
191:     m_pNWSeriesOptions.TextSymbolSet sSymbName, m_pTextSym
192:   End If
193:   Me.SetFocus

  Exit Sub
ErrorHandler:
  HandleError True, "cmdProperties_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub


Private Sub cmdRename_Click()
  On Error GoTo ErrorHandler

  Dim sSymbName As String
  Dim sDefaultSymb As String
  
207:   sSymbName = InputBox("Enter new text symbol name: ")
  If sSymbName = "" Then Exit Sub
209:   If m_pNWSeriesOptions.TextSymbolExists(sSymbName) Then
210:     MsgBox "Symbol name already exists.", vbOKOnly
    Exit Sub
212:   End If
    
214:   sDefaultSymb = m_pNWSeriesOptions.TextSymbolDefault
215:   With lstTextSymbols
216:     If .List(.ListIndex) = "" Then
217:       .ListIndex = 0
218:     End If
219:     m_pNWSeriesOptions.TextSymbolRename .List(.ListIndex), sSymbName
                                                  'detect if the symbol being renamed is
                                                  'also the default symbol
222:     If StrComp(sDefaultSymb, .List(.ListIndex), vbTextCompare) = 0 Then
223:       m_pNWSeriesOptions.TextSymbolDefault = sSymbName
224:       lblDefaultSymbol.Caption = sSymbName
225:     End If
226:     .RemoveItem (.ListIndex)
227:     .AddItem sSymbName
228:   End With


  Exit Sub
ErrorHandler:
  HandleError True, "cmdRename_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Form_Load()
  On Error GoTo ErrorHandler

  Dim HRESULT As Long, i As Long, lSymbolCount As Long
  Dim sTextSymbName As String
  Dim vSymbolNames() As Variant
                                                  'Initialize UI
243:   HRESULT = SetForegroundWindow(Me.hwnd)
244:   If m_pNWSeriesOptions Is Nothing Then
    Exit Sub
246:   End If
247:   lstTextSymbols.Clear
248:   lSymbolCount = m_pNWSeriesOptions.TextSymbolCount
249:   vSymbolNames = m_pNWSeriesOptions.TextSymbolNames
250:   For i = 0 To (lSymbolCount - 1)
251:     lstTextSymbols.AddItem vSymbolNames(i)
252:   Next i
253:   lblDefaultSymbol.Caption = m_pNWSeriesOptions.TextSymbolDefault


  Exit Sub
ErrorHandler:
  HandleError True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub lstTextSymbols_ItemCheck(Item As Integer)
  On Error GoTo ErrorHandler

264:   With lstTextSymbols
265:     If StrComp(m_pNWSeriesOptions.TextSymbolDefault, .List(.ListIndex), vbTextCompare) = 0 Then
      
267:     End If
268:   End With

  Exit Sub
ErrorHandler:
  HandleError True, "lstTextSymbols_ItemCheck " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
