VERSION 5.00
Begin VB.Form frmVisibleLayers 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Layers Visible in this Map Page"
   ClientHeight    =   4020
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDebug 
      Caption         =   "Debug"
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdUnselectAll 
      Caption         =   "&Unselect All"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "&Select All"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   4455
      Begin VB.ComboBox cboVisibilityGroups 
         Height          =   315
         Left            =   2520
         TabIndex        =   1
         Text            =   "cboVisibilityGroups"
         Top             =   120
         Width           =   1815
      End
      Begin VB.CheckBox chkVisibilityGroup 
         Caption         =   "Participate in visibility group"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   3600
      Width           =   1095
   End
   Begin VB.ListBox lstVisibleLayers 
      Height          =   1992
      ItemData        =   "frmVisibleLayers.frx":0000
      Left            =   120
      List            =   "frmVisibleLayers.frx":0016
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   1440
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Select layers that will be visible in this map page."
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmVisibleLayers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Public m_pApp As IApplication
Public m_pMapPage As INWDSMapPage
Public m_pNWMapPage As INWMapPageAttribs

Private m_pSeriesOptions As INWDSMapSeriesOptions
Private m_pSeriesOptions2 As INWDSMapSeriesOptions2
Private m_pNWSeriesOptions As INWMapSeriesOptions
Private m_pLyrInvisGroupNames() As Variant
Private m_pDictLayers As Scripting.Dictionary
Private m_pMainMap As IMap

Private m_bInitializing As Boolean
Const c_sModuleFileName As String = "frmVisibleLayers.frm"







Private Sub cboVisibilityGroups_Change()
  On Error GoTo ErrorHandler

28:   With cboVisibilityGroups
29:     If FindControlString(cboVisibilityGroups, .Text) = -1 Then
30:       If .ListCount > 0 Then
31:         .ListIndex = 0
32:       ElseIf .ListIndex >= .ListCount Then
33:         .ListIndex = 0
34:       End If
35:       .Text = .List(.ListIndex)
36:     End If
37:   End With

  Exit Sub
ErrorHandler:
  HandleError True, "cboVisibilityGroups_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cboVisibilityGroups_Click()
  On Error GoTo ErrorHandler

  Dim pMapBook As INWDSMapBook, pLyrGroup As INWLayerVisibilityGroup, lLyrCount As Long
  Dim i As Long, lFindIdx As Long, vInvLayers() As Variant
  
50:   m_pNWMapPage.LayerVisibilityGroup = cboVisibilityGroups.Text
                                                  'apply the visibility settings of that
                                                  'group to the list box
  If m_pApp Is Nothing Then Exit Sub
54:   Set pMapBook = GetMapBookExtension(m_pApp)
55:   Set m_pNWSeriesOptions = pMapBook.ContentItem(0)
56:   For i = 0 To lstVisibleLayers.ListCount - 1
57:     lstVisibleLayers.Selected(i) = True
58:   Next i
                                                  'set/unset the visible layers
                                                  'based on the group settings.
61:   Set pLyrGroup = m_pNWSeriesOptions.LayerGroupGet(cboVisibilityGroups.Text)
62:   lLyrCount = pLyrGroup.InvisibleLayerCount
63:   vInvLayers = pLyrGroup.InvisibleLayers
64:   For i = 0 To lLyrCount - 1
65:     lFindIdx = FindControlString(lstVisibleLayers, vInvLayers(i))
66:     If pLyrGroup.Exists(vInvLayers(i)) Then
67:       If lFindIdx > -1 Then
68:         lstVisibleLayers.Selected(lFindIdx) = False
69:       End If
70:     End If
71:   Next i

  Exit Sub
ErrorHandler:
  HandleError True, "cboVisibilityGroups_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub chkVisibilityGroup_Click()
  On Error GoTo ErrorHandler

81:   If Not m_bInitializing Then
  Dim pMapBook As INWDSMapBook, pLyrGroup As INWLayerVisibilityGroup, lLyrCount As Long
  Dim i As Long, lFindIdx As Long, vInvLayers() As Variant
                                                  'check for exit conditions
85:   If m_pMapPage Is Nothing Then
    Exit Sub
87:   End If
88:   If Me.cboVisibilityGroups.ListCount = 0 Then
89:     MsgBox "Unable to find any layer groups to make available.  Please define" & vbNewLine _
         & "a new layer visibility group in the series properties.", vbOKOnly
91:     chkVisibilityGroup.Value = 0
    Exit Sub
93:   End If
94:   If m_pNWMapPage Is Nothing Then
95:     Set m_pNWMapPage = m_pMapPage
96:   End If
                                                  'respond to user's selection
98:   If chkVisibilityGroup.Value = vbChecked Then
99:     cboVisibilityGroups.Enabled = True
100:     m_pNWMapPage.LayerVisibilityGroup = cboVisibilityGroups.Text
                                                  'apply the visibility settings of that
                                                  'group to the list box
    If m_pApp Is Nothing Then Exit Sub
104:     Set pMapBook = GetMapBookExtension(m_pApp)
105:     Set m_pNWSeriesOptions = pMapBook.ContentItem(0)
106:     Set pLyrGroup = m_pNWSeriesOptions.LayerGroupGet(cboVisibilityGroups.Text)
107:     lLyrCount = pLyrGroup.InvisibleLayerCount
108:     vInvLayers = pLyrGroup.InvisibleLayers
109:     For i = 0 To lstVisibleLayers.ListCount - 1
110:       lstVisibleLayers.Selected(i) = True
111:     Next i
112:     For i = 0 To lLyrCount - 1
113:       lFindIdx = FindControlString(lstVisibleLayers, vInvLayers(i))
114:       If pLyrGroup.Exists(vInvLayers(i)) Then
115:         If lFindIdx > -1 Then
116:           lstVisibleLayers.Selected(lFindIdx) = False
117:         End If
118:       End If
119:     Next i
120:     lstVisibleLayers.Enabled = False
121:     cmdSelectAll.Enabled = False
122:     cmdUnselectAll.Enabled = False
123:   Else
124:     cboVisibilityGroups.Enabled = False
125:     lstVisibleLayers.Enabled = True
126:     cmdSelectAll.Enabled = True
127:     cmdUnselectAll.Enabled = True
128:     m_pNWMapPage.LayerVisibilityGroup = ""
129:   End If

131:   End If
  Exit Sub
ErrorHandler:
  HandleError True, "chkVisibilityGroup_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdCancel_Click()
138:   Set m_pDictLayers = Nothing
139:   Me.Hide
End Sub


'A function purely for debugging
Private Sub cmdDebug_Click()
  On Error GoTo ErrorHandler

  Dim sInfo As String, pLyrGroup As INWLayerVisibilityGroup, pNWMapSeries As INWMapSeriesOptions
  Dim pMapBook As INWDSMapBook, vGrpNames() As Variant, lGrpCount As Long, i As Long
  Dim j As Long, vInvLyrNames() As Variant, sGrpName As String
  
151:   If m_pNWSeriesOptions Is Nothing Then
    If m_pApp Is Nothing Then Exit Sub
153:     Set pMapBook = GetMapBookExtension(m_pApp)
154:     Set m_pNWSeriesOptions = pMapBook.ContentItem(0)
155:   End If
156:   If m_pNWMapPage Is Nothing Then
157:     MsgBox "m_pNWMapPage is nothing."
    Exit Sub
159:   End If
  
161:   vGrpNames = m_pNWSeriesOptions.LayerGroups
162:   lGrpCount = m_pNWSeriesOptions.LayerGroupCount
163:   For i = 0 To (lGrpCount - 1)
164:     sInfo = sInfo & "group " & i & " is " & vGrpNames(i) & vbNewLine
165:     sGrpName = vGrpNames(i)
166:     Set pLyrGroup = m_pNWSeriesOptions.LayerGroupGet(sGrpName)
167:     vInvLyrNames = pLyrGroup.InvisibleLayers
168:     For j = 0 To pLyrGroup.InvisibleLayerCount - 1
169:       sInfo = sInfo & "    invis layer " & j & " -- " & vInvLyrNames(j) & vbNewLine
170:     Next j
171:   Next i
  'Set pLyrGroup = m_pNWMapPage.LayerVisibilityGroup
  'pLyrGroup.Name
174:   MsgBox sInfo


  Exit Sub
ErrorHandler:
  HandleError True, "cmdDebug_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub



' If the layer groups checkbox is set, then assign this map page to a layer group.
' If no layer groups checkbox is set, then assign the invisible map pages based on
' user selections in lstVisibleLayers.
'------------------------
Private Sub cmdOK_Click()
  On Error GoTo ErrorHandler
                                                  
  Dim i As Long, lLyrCount As Long, pMxDoc As IMxDocument, pLayers As IEnumLayer
  Dim pLayer As ILayer, lLyrIdx As Long, pLyrGroup As INWLayerVisibilityGroup
  Dim vLayers() As Variant, sLyrName As String, lInvLyrCount As Long
  
195:   Set pMxDoc = m_pApp.Document
  'Set pLayers = pMxDoc.FocusMap.Layers
197:   Set pLayers = m_pMainMap.Layers
198:   If m_pNWMapPage Is Nothing Then
199:     MsgBox "frmVisibleLayers,cmdOK_Click -- m_pNWMapPage is nothing, so the OK button" & vbNewLine _
         & "therefore cannot function.  Close this dialog, and reopen."
    Exit Sub
202:   End If
  
204:   If chkVisibilityGroup.Value = 1 Then
                                                  'assign the name of the group to the
                                                  'map page
207:     m_pNWMapPage.LayerVisibilityGroup = cboVisibilityGroups.Text
                                                  'alter which map layers are visible/not
                                                  'visible based on the layer group settings
210:     Set pLyrGroup = m_pNWSeriesOptions.LayerGroupGet(cboVisibilityGroups.Text)
211:     Set pLayer = pLayers.Next
212:     Do While Not pLayer Is Nothing
213:       pLayer.Visible = Not pLyrGroup.Exists(pLayer.Name)
214:       Set pLayer = pLayers.Next
215:     Loop
216:   Else
                                                  'assign the visible/invisible layers to the
                                                  'map page, and alter the layer settings in
                                                  'ArcMap
220:     m_pNWMapPage.LayerVisibilityGroup = ""
221:     Set pLayer = pLayers.Next
222:     Do While Not pLayer Is Nothing
223:       lLyrIdx = FindControlString(lstVisibleLayers, pLayer.Name)
224:       If lLyrIdx > -1 Then
225:         pLayer.Visible = lstVisibleLayers.Selected(lLyrIdx)
226:         If pLayer.Visible Then
227:           m_pNWMapPage.InvisibleLayerRemove pLayer.Name
228:         Else
229:           m_pNWMapPage.InvisibleLayerAdd pLayer.Name
230:         End If
231:       End If
232:       Set pLayer = pLayers.Next
233:     Loop
234:   End If

236:   Set m_pMapPage = m_pNWMapPage
237:   pMxDoc.ActiveView.Refresh
238:   Me.Hide
  
  Exit Sub
ErrorHandler:
  HandleError True, "cmdOK_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdSelectAll_Click()
  Dim i As Long, lLyrCount As Long
247:   lLyrCount = lstVisibleLayers.ListCount
248:   For i = 0 To (lLyrCount - 1)
249:     lstVisibleLayers.Selected(i) = True
250:   Next i
End Sub

Private Sub cmdUnselectAll_Click()
  Dim i As Long, lLyrCount As Long
255:   lLyrCount = lstVisibleLayers.ListCount
256:   For i = 0 To (lLyrCount - 1)
257:     lstVisibleLayers.Selected(i) = False
258:   Next i
End Sub

Public Sub Init_Form()
  On Error GoTo ErrorHandler

  Dim pMapBook As INWDSMapBook, lGroupCount As Long, lLayerCount As Long
  Dim vGroupNames() As Variant, pLyrGroup As INWLayerVisibilityGroup
  Dim sLyrName As String, pMxDoc As IMxDocument, pLayers As IEnumLayer
  Dim vInvisLayerNames() As Variant, sLayerNames() As String
  Dim lInvLyrCount As Long, i As Long, lNewItemIdx As Long
  Dim pLayer As ILayer, lFindIdx As Long, sMainMap As String, pMap As IMap
  Dim pMaps As IMaps
  
                                                  'acquire the list of layers in ArcMap
  If m_pApp Is Nothing Then Exit Sub
  If m_pMapPage Is Nothing Then Exit Sub
                                                  'Initialize the dialog UI
276:   m_bInitializing = True
277:   lstVisibleLayers.Clear
278:   cboVisibilityGroups.Clear
  
280:   Set m_pDictLayers = New Scripting.Dictionary
281:   Set pMxDoc = m_pApp.Document
  
283:   Set pMapBook = GetMapBookExtension(m_pApp)
284:   Set m_pNWSeriesOptions = pMapBook.ContentItem(0)
285:   Set m_pNWMapPage = m_pMapPage
                                                  'must use the main data frame
                                                  'since layer visibility doesn't
                                                  'account for data frames other
                                                  'than the main data frame
290:   sMainMap = m_pNWSeriesOptions.DataFrameMainFrame
291:   Set pMaps = pMxDoc.Maps
292:   For i = 0 To (pMaps.count - 1)
293:     Set pMap = pMaps.Item(i)
294:     If StrComp(pMap.Name, sMainMap, vbTextCompare) = 0 Then
295:       Set pLayers = pMap.Layers
296:       i = pMaps.count
297:     End If
298:   Next i
299:   Set m_pMainMap = pMap
                                                  'build the list of layer groups
301:   vGroupNames = m_pNWSeriesOptions.LayerGroups
302:   lGroupCount = UBound(vGroupNames)
303:   With cboVisibilityGroups
304:     For i = 0 To lGroupCount
305:       .AddItem vGroupNames(i)
306:     Next i
307:     If .ListCount > 0 Then
308:       .Text = .List(0)
309:       chkVisibilityGroup.Enabled = True
310:     Else
311:       chkVisibilityGroup.Value = 0
312:       chkVisibilityGroup.Enabled = False
313:     End If
314:     .Enabled = False
315:   End With
                                                  'if this map page is in a layer invisibility
                                                  'group, grab that group's layers
318:   If m_pNWMapPage.LayerVisibilityGroup <> "" Then
319:     If m_pNWSeriesOptions.LayerGroupExists(m_pNWMapPage.LayerVisibilityGroup) Then
320:       Set pLyrGroup = m_pNWSeriesOptions.LayerGroupGet(m_pNWMapPage.LayerVisibilityGroup)
321:       vInvisLayerNames = pLyrGroup.InvisibleLayers
322:       chkVisibilityGroup.Value = 1
323:       cboVisibilityGroups.Enabled = True
324:       cboVisibilityGroups.Text = m_pNWMapPage.LayerVisibilityGroup
325:       lstVisibleLayers.Enabled = False
326:       cmdSelectAll.Enabled = False
327:       cmdUnselectAll.Enabled = False
328:     Else
329:       MsgBox "Warning, layer visibility group " & m_pNWMapPage.LayerVisibilityGroup _
           & " appears to have been deleted." & vbNewLine _
           & "This map page will no longer participate in that map visibility group."
332:       m_pNWMapPage.LayerVisibilityGroup = ""
                                                  'get that deleted group out of the UI
                                                  'if it's there
335:       lFindIdx = FindControlString(cboVisibilityGroups, m_pNWMapPage.LayerVisibilityGroup)
336:       If lFindIdx > -1 Then
337:         cboVisibilityGroups.RemoveItem lFindIdx
338:         If cboVisibilityGroups.ListCount = 0 Then
339:           chkVisibilityGroup.Value = 1
340:           chkVisibilityGroup.Enabled = False
341:           cboVisibilityGroups.Enabled = False
342:           lstVisibleLayers.Enabled = True
343:           cmdSelectAll.Enabled = True
344:           cmdUnselectAll.Enabled = True
345:         End If
346:       End If
                                                  'treat this map page as one outside of
                                                  'a visible layers group (as it now is)
349:       vInvisLayerNames = m_pNWMapPage.InvisibleLayers
350:       chkVisibilityGroup.Value = 0
351:     End If
                                                  'otherwise grab the map page's invisible
                                                  'layers
354:   Else
355:     chkVisibilityGroup.Value = 0
356:     vInvisLayerNames = m_pNWMapPage.InvisibleLayers
357:     cboVisibilityGroups.Enabled = False
358:     lstVisibleLayers.Enabled = True
359:     cmdSelectAll.Enabled = True
360:     cmdUnselectAll.Enabled = True
361:   End If
                                                  'load the UI with the list of all Arcmap layers
                                                  'put a checkmark on all list entries
                                                  'remove checkmarks for all invisible layers
365:   lInvLyrCount = UBound(vInvisLayerNames)
366:   Set pLayer = pLayers.Next
367:   With lstVisibleLayers
368:     Do While Not pLayer Is Nothing
369:       .AddItem pLayer.Name
                                                    'apply the visible layer settings
371:       .Selected(.NewIndex) = True
372:       For i = 0 To lInvLyrCount
373:         sLyrName = vInvisLayerNames(i)
374:         If StrComp(pLayer.Name, sLyrName, vbTextCompare) = 0 Then
375:           .Selected(.NewIndex) = False
376:           i = lInvLyrCount
377:         End If
378:       Next i
      
      ''''''''option 2 -- (mutually exclusive from option 1, will have to pick one)
      ''''''''         -- use the layer visibility setting to set the stored
      ''''''''         -- settings
      '''''''' conclusion -- better to honor the settings, rather than the layer
      ''''''''            -- state.  It enforces the idea that you control the
      '''''''''''''''''''''''map series settings only in the map series UI.
386:       Set pLayer = pLayers.Next
387:     Loop
388:   End With
                                                  'perform automatic QC by removing layer
                                                  'names that don't exist in the map
391:   For i = 0 To lInvLyrCount - 1
392:     lNewItemIdx = FindControlString(lstVisibleLayers, vInvisLayerNames(i))
393:     If lNewItemIdx = -1 Then
394:       If pLyrGroup Is Nothing Then
                                                  'delete layer from map page
396:         m_pNWMapPage.InvisibleLayerRemove vInvisLayerNames(i)
397:       Else
                                                  'delete layer from group
399:         pLyrGroup.DeleteLayer vInvisLayerNames(i)
400:       End If
401:     End If
402:   Next i
  

405:   m_bInitializing = False
  Exit Sub
ErrorHandler:
408:   m_bInitializing = False
  HandleError True, "Init_Form " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub


Private Sub Form_Unload(Cancel As Integer)
  Dim i As Long
415:   With lstVisibleLayers
416:     For i = 0 To .ListCount - 1
417:       .Selected(i) = False
418:     Next i
419:   End With
420:   Set m_pDictLayers = Nothing
End Sub


