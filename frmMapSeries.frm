VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMapSeries 
   BackColor       =   &H80000004&
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   165
   ClientTop       =   885
   ClientWidth     =   4275
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   4275
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBook 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   150
      ScaleHeight     =   4335
      ScaleWidth      =   3885
      TabIndex        =   1
      Top             =   630
      Width           =   3885
      Begin MSComctlLib.TreeView tvwMapBook 
         Height          =   4725
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   8334
         _Version        =   393217
         Indentation     =   44
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   3
         FullRowSelect   =   -1  'True
         SingleSel       =   -1  'True
         ImageList       =   "ImageList1"
         Appearance      =   1
      End
   End
   Begin VB.ListBox lstSorter 
      Height          =   1230
      Left            =   2790
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   3420
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   33
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapSeries.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapSeries.frx":062E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapSeries.frx":0C5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapSeries.frx":11EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapSeries.frx":1780
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapSeries.frx":1BF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapSeries.frx":2064
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapSeries.frx":2656
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuHeadingBook 
      Caption         =   "Book"
      Begin VB.Menu mnuBook 
         Caption         =   "Add Map Series..."
         Index           =   0
      End
      Begin VB.Menu mnuBook 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuBook 
         Caption         =   "Print Map Book..."
         Index           =   2
      End
      Begin VB.Menu mnuBook 
         Caption         =   "Export Map Book..."
         Index           =   3
      End
   End
   Begin VB.Menu mnuHeadingSeries 
      Caption         =   "Series"
      Begin VB.Menu mnuSeries 
         Caption         =   "Select/Enable Pages..."
         Index           =   0
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "Tag as Date"
         Index           =   2
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "Tag as Title"
         Index           =   3
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "Tag as Page Number"
         Index           =   4
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "Tag with Index Layer Field..."
         Index           =   5
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "Clear Tag for Selected"
         Index           =   6
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "Delete Series"
         Index           =   8
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "Delete Disabled Pages"
         Index           =   9
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "Disable Series"
         Index           =   11
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "Print Series..."
         Index           =   13
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "Export Series..."
         Index           =   14
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "Create Series Index..."
         Index           =   15
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "-"
         Index           =   16
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "Series Properties..."
         Index           =   17
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "Page Properties..."
         Index           =   18
      End
   End
   Begin VB.Menu mnuHeadingPage 
      Caption         =   "Page"
      Begin VB.Menu mnuPage 
         Caption         =   "View Page"
         Index           =   0
      End
      Begin VB.Menu mnuPage 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuPage 
         Caption         =   "Delete Page"
         Index           =   2
      End
      Begin VB.Menu mnuPage 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuPage 
         Caption         =   "Disable Page"
         Index           =   4
      End
      Begin VB.Menu mnuPage 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuPage 
         Caption         =   "Print Page..."
         Index           =   6
      End
      Begin VB.Menu mnuPage 
         Caption         =   "Export Page..."
         Index           =   7
      End
   End
End
Attribute VB_Name = "frmMapSeries"
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




Option Explicit

Public m_pApp As IApplication
Private m_lXClick As Single
Private m_lYClick As Single
Private m_lButton As Single
Private m_pCurrentNode As Node
Private m_bNodeFlag As Boolean
Private m_bClickFlag As Boolean
Private m_bLabelingChanged As Boolean
Private m_pExportFrame As IModelessFrame

Private Sub Form_Load()
14:   tvwMapBook.Nodes.Clear
'  tvwMapBook.Nodes.Add , , "MapBook", "Map Book (0 pages)", 1
16:   tvwMapBook.Nodes.Add , , "MapBook", "Map Book", 1
17:   m_bNodeFlag = True
18:   m_bClickFlag = False
19:   m_bLabelingChanged = False
20:   Set m_pExportFrame = New ModelessFrame
End Sub

Private Sub Form_Terminate()
24:   Set m_pExportFrame = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
28:   Set m_pExportFrame = Nothing
End Sub

Private Sub mnuBook_Click(Index As Integer)
On Error GoTo ErrHand:
  Dim pMapBook As IDSMapBook
  'Check to see if a MapSeries already exists
35:   Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
  Select Case Index
  Case 0  'Add Map Series
40:     If pMapBook.ContentCount > 0 Then
41:       MsgBox "You must remove the existing Map Series before adding another."
      Exit Sub
43:     End If
  
    'Call the wizard for setting parameters and creating the series
46:     Set frmMapSeriesWiz.m_pApp = m_pApp
47:     frmMapSeriesWiz.Show vbModal
  Case 1  'Separator
  Case 2  'Print Map Book
50:     ShowPrinterDialog m_pApp, , pMapBook
'    pMapBook.PrintBook
  Case 3  'Export Map Book
53:     ShowExporterDialog m_pApp, , pMapBook
'    pMapBook.ExportBook
55:   End Select
  
  Exit Sub
ErrHand:
59:   MsgBox "mnuBook_Click - " & Err.Description
End Sub

Private Sub mnuPage_Click(Index As Integer)
On Error GoTo ErrHand:
  Dim pMapBook As IDSMapBook, pMapSeries As IDSMapSeries
  Dim lPage As Long, sText As String, lPos As Long, pMapPage As IDSMapPage
  Dim pSeriesOpts As IDSMapSeriesOptions, pSeriesOpts2 As IDSMapSeriesOptions2
  'Check to see if a MapSeries already exists
68:   Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
71:   Set pMapSeries = pMapBook.ContentItem(0)
72:   Set pSeriesOpts = pMapSeries
73:   Set pSeriesOpts2 = pSeriesOpts
74:   lPage = m_pCurrentNode.Tag
  Select Case Index
  Case 0  'View Page
77:     Set pMapPage = pMapSeries.Page(lPage)
78:     pMapPage.DrawPage m_pApp.Document, pMapSeries, True
79:     If pSeriesOpts2.ClipData > 0 Then
80:       g_bClipFlag = True
81:     End If
82:     If pSeriesOpts.RotateFrame Then
83:       g_bRotateFlag = True
84:     End If
85:     If pSeriesOpts.LabelNeighbors Then
86:       g_bLabelNeighbors = True
87:     End If
  Case 1  'Separator
  Case 2  'Delete Page
    'Remove the page, then update the tags on all subsequent pages
91:     pMapSeries.RemovePage lPage
92:     tvwMapBook.Nodes.Remove lPage + 3
93:     RenumberPages pMapSeries
  Case 3  'Separator
  Case 4  'Disable Page
    'Get the index number from the tag of the node
97:     pMapSeries.Page(lPage).EnablePage = Not pMapSeries.Page(lPage).EnablePage
98:     If pMapSeries.Page(lPage).EnablePage Then
99:       m_pCurrentNode.Image = 5
100:     Else
101:       m_pCurrentNode.Image = 6
102:     End If
  Case 5  'Separator
  Case 6  'Print Page
105:     ShowPrinterDialog m_pApp, pMapSeries, pMapSeries.Page(lPage)
  Case 7  'Export Page
107:     ShowExporterDialog m_pApp, pMapSeries, pMapSeries.Page(lPage)
108:   End Select
  
  Exit Sub
ErrHand:
112:   MsgBox "mnuPage_Click - " & Erl & " - " & Err.Description
End Sub

Private Sub mnuSeries_Click(Index As Integer)
On Error GoTo ErrHand:
  Dim pMapBook As IDSMapBook, pMapSeries As IDSMapSeries, pSeriesProps As IDSMapSeriesProps
  Dim lLoop As Long, pDoc As IMxDocument, pActive As IActiveView, bFlag As Boolean
  Dim pGraphicsCont As IGraphicsContainer, pElemProps As IElementProperties
  Dim pEnv As IEnvelope, pElem As IElement, pTextElement As ITextElement, pEnv2 As IEnvelope
  Dim pGraphicsContSel As IGraphicsContainerSelect, pMap As IMap
  Dim pIndexLayer As IFeatureLayer, lIndex As Long, sName As String, sTemp As String
  'Check to see if a MapSeries already exists
124:   Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
127:   Set pMapSeries = pMapBook.ContentItem(0)
128:   Set pSeriesProps = pMapSeries
129:   Set pDoc = m_pApp.Document
  Select Case Index
  Case 0  'Select Pages
132:     Set frmSelectPages.m_pApp = m_pApp
133:     frmSelectPages.Show vbModal
  Case 1  'Separator
  Case 2  'Tag as Date
136:     bFlag = TagItem(pDoc, "DSMAPBOOK - DATE", "")
137:     If Not bFlag Then
138:       MsgBox "You must have one Text Element selected in the Page Layout for tagging!!!"
139:     End If
  Case 3  'Tag as Title
141:     bFlag = TagItem(pDoc, "DSMAPBOOK - TITLE", "")
142:     If Not bFlag Then
143:       MsgBox "You must have one Text Element selected in the Page Layout for tagging!!!"
144:     End If
  Case 4  'Tag as Page Number
146:     bFlag = TagItem(pDoc, "DSMAPBOOK - PAGENUMBER", "")
147:     If Not bFlag Then
148:       MsgBox "You must have one Text Element selected in the Page Layout for tagging!!!"
149:     End If
  Case 5  'Tag with Index Layer Field...
    'Find the data frame
152:     Set pMap = FindDataFrame(pDoc, pSeriesProps.DataFrameName)
153:     If pMap Is Nothing Then
154:       MsgBox "Could not find map in mnuSeries_Click routine!!!"
      Exit Sub
156:     End If
    'Find the Index layer
158:     Set pIndexLayer = FindLayer(pSeriesProps.IndexLayerName, pMap)
159:     If pIndexLayer Is Nothing Then
160:       MsgBox "Could not find index layer in mnuSeries_Click routine!!!"
      Exit Sub
162:     End If
  
164:     frmTagIndexField.InitializeList pIndexLayer.FeatureClass.Fields
165:     frmTagIndexField.Show vbModal
    
    'Exit sub if Cancel was selected
168:     If frmTagIndexField.m_bCancel Then
169:       Unload frmTagIndexField
      Exit Sub
171:     End If
    
173:     lIndex = frmTagIndexField.lstFields.ListIndex
174:     If lIndex >= 0 Then
175:       sTemp = frmTagIndexField.lstFields.List(lIndex)
176:     Else
177:       MsgBox "You did not pick a field to tag with!!!"
178:       Unload frmTagIndexField
      Exit Sub
180:     End If
181:     Unload frmTagIndexField
    
183:     lIndex = InStr(1, sTemp, " - ")
184:     sName = Mid(sTemp, 1, lIndex - 1)
185:     bFlag = TagItem(pDoc, "DSMAPBOOK - EXTRAITEM", sName)
186:     If Not bFlag Then
187:       MsgBox "You must have one Text Element selected in the Page Layout for tagging!!!"
188:     End If
  Case 6  'Clear Tag for selected
190:     Set pGraphicsCont = pDoc.PageLayout
191:     Set pGraphicsContSel = pDoc.PageLayout
192:     For lLoop = 0 To pGraphicsContSel.ElementSelectionCount - 1
193:       Set pElemProps = pGraphicsContSel.SelectedElement(lLoop)
194:       If TypeOf pElemProps Is ITextElement Then
195:         pElemProps.Name = ""
196:         pElemProps.Type = ""
197:         pGraphicsCont.UpdateElement pTextElement
198:       End If
199:     Next lLoop
  Case 7  'Separator
  Case 8  'Delete Series
202:     Set pActive = pDoc.FocusMap
203:     TurnOffClipping pMapSeries, m_pApp
204:     Set pMapSeries = Nothing
205:     pMapBook.RemoveContent 0
206:     tvwMapBook.Nodes.Clear
207:     tvwMapBook.Nodes.Add , , "MapBook", "Map Book", 1
208:     RemoveIndicators m_pApp
209:     pActive.Refresh
  Case 9  'Delete Disabled pages
    'Loop in reverse order so we remove pages as we work up.  Doing it this way makes
    'sure numbering isn't messed up when a page/node is removed.
213:     For lLoop = pMapSeries.PageCount - 1 To 0 Step -1
214:       If Not pMapSeries.Page(lLoop).EnablePage Then
215:         pMapSeries.RemovePage lLoop
216:         tvwMapBook.Nodes.Remove lLoop + 3
217:       End If
218:     Next lLoop
219:     RenumberPages pMapSeries
  Case 10  'Separator
  Case 11  'Disable Series
    'Get the index number from the tag of the node
223:     pMapSeries.EnableSeries = Not pMapSeries.EnableSeries
224:     If pMapSeries.EnableSeries Then
225:       m_pCurrentNode.Image = 3
226:     Else
227:       m_pCurrentNode.Image = 4
228:     End If
  Case 12  'Separator
  Case 13  'Print Series
231:     ShowPrinterDialog m_pApp, pMapSeries, Nothing
'    pMapSeries.PrintSeries
  Case 14  'Export Series
234:     ShowExporterDialog m_pApp, pMapSeries, Nothing
'    pMapSeries.ExportSeries
  Case 15
237:     Set frmCreateIndex.m_pApp = m_pApp
238:     frmCreateIndex.Show vbModal
  Case 16  'Separator
  Case 17  'Series Properties...
241:     Set frmSeriesProperties.m_pApp = m_pApp
242:     frmSeriesProperties.Show vbModal
  Case 18  'Page Properties...
244:     Set frmPageProperties.m_pApp = m_pApp
245:     frmPageProperties.Show vbModal
246:   End Select
  
  Exit Sub
ErrHand:
250:   MsgBox "mnuSeries_Click - " & Erl & " - " & Err.Description
End Sub

Private Function TagItem(pDoc As IMxDocument, sName As String, sType As String) As Boolean
On Error GoTo ErrHand:
  Dim bFlag As Boolean, pGraphicsCont As IGraphicsContainer, pActive As IActiveView
  Dim pElemProps As IElementProperties, pElem As IElement, pTextElement As ITextElement
  Dim pEnv2 As IEnvelope, pGraphicsContSel As IGraphicsContainerSelect, pEnv As IEnvelope
  
259:   Set pGraphicsCont = pDoc.PageLayout
260:   Set pGraphicsContSel = pDoc.PageLayout
261:   bFlag = False
262:   If pGraphicsContSel.ElementSelectionCount = 1 Then
263:     Set pElemProps = pGraphicsContSel.SelectedElement(0)
264:     If TypeOf pElemProps Is ITextElement Then
265:       Set pActive = pDoc.PageLayout
266:       pElemProps.Name = sName
267:       Set pElem = pElemProps
268:       Set pEnv = New Envelope
269:       pElem.QueryBounds pActive.ScreenDisplay, pEnv
270:       Set pTextElement = pElemProps
      Select Case sName
      Case "DSMAPBOOK - DATE"
273:         pTextElement.Text = Format(Date, "mmm dd, yyyy")
      Case "DSMAPBOOK - TITLE"
275:         pTextElement.Text = "Title String"
      Case "DSMAPBOOK - PAGENUMBER"
277:         pTextElement.Text = "PAGE #"
      Case "DSMAPBOOK - EXTRAITEM"
279:         pTextElement.Text = sType
280:         pElemProps.Type = sType
281:       End Select
282:       pGraphicsCont.UpdateElement pTextElement
283:       Set pEnv2 = New Envelope
284:       pElem.QueryBounds pActive.ScreenDisplay, pEnv2
285:       pEnv.Union pEnv2
286:       pActive.PartialRefresh esriViewGraphics, Nothing, pEnv
287:       bFlag = True
288:     End If
289:   End If
  
291:   TagItem = bFlag

  Exit Function
ErrHand:
295:   MsgBox "TagItem - " & Erl & " - " & Err.Description
296:   TagItem = bFlag
End Function

Private Sub RenumberPages(pMapSeries As IDSMapSeries)
On Error GoTo ErrHand:
'Routine for renumber the pages after one is removed
  Dim lLoop As Long, pNode As Node, sName As String, lPageNumber As Long
  Dim pPage As IDSMapPage, pSeriesProps As IDSMapSeriesProps
304:   Set pSeriesProps = pMapSeries
305:   For lLoop = 0 To pMapSeries.PageCount - 1
306:     lPageNumber = lLoop + pSeriesProps.StartNumber
307:     Set pPage = pMapSeries.Page(lLoop)
308:     Set pNode = tvwMapBook.Nodes.Item(lLoop + 3)
309:     sName = Mid(pNode.Key, 2)
310:     pNode.Tag = lLoop
311:     pNode.Key = "a" & sName
312:     pNode.Text = lPageNumber & " - " & sName
313:     pPage.PageNumber = lPageNumber
314:   Next lLoop
315:   tvwMapBook.Refresh
  
  Exit Sub
ErrHand:
319:   MsgBox "RenumberPages - " & Erl & " - " & Err.Description
End Sub

Private Sub picBook_Resize()
323:   tvwMapBook.Width = picBook.Width
324:   tvwMapBook.Height = picBook.Height
End Sub

Private Sub tvwMapBook_DblClick()
On Error GoTo ErrHand:
  Dim lPos As String, sText As String, pMapPage As IDSMapPage, lPage As Long
  Dim pSeriesOpts As IDSMapSeriesOptions, pSeriesOpts2 As IDSMapSeriesOptions2
  Dim pMapBook As IDSMapBook, pMapSeries As IDSMapSeries
332:   Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
335:   Set pMapSeries = pMapBook.ContentItem(0)
336:   Set pSeriesOpts = pMapSeries
337:   Set pSeriesOpts2 = pSeriesOpts
  
  'There is no NodeDoubleClick event, so we have to use the DblClick event on the control
  'and check to make sure we are over a node.  In the event of a doubleclick on a node,
  'the order of events being fired are NodeClick, MouseUp, Click, DblClick, MouseUp.  To
  'make sure the doubleclick occurred over a node, we can set a flag in the NodeClick event
  'and then disable it in the MouseUp event after the Click event.
  If Not m_bNodeFlag Then Exit Sub
  
  Select Case m_pCurrentNode.Image
  Case 5, 6   'Enable and not Enabled options for a map page
348:     If m_lXClick > 1320 Then
349:       If m_lButton = 1 Then
350:         lPage = m_pCurrentNode.Tag
351:         Set pMapPage = pMapSeries.Page(lPage)
352:         pMapPage.DrawPage m_pApp.Document, pMapSeries, True
353:         If pSeriesOpts2.ClipData > 0 Then
354:           g_bClipFlag = True
355:         End If
356:         If pSeriesOpts.RotateFrame Then
357:           g_bRotateFlag = True
358:         End If
359:         If pSeriesOpts.LabelNeighbors Then
360:           g_bLabelNeighbors = True
361:         End If
362:       End If
363:     End If
364:   End Select

  Exit Sub
ErrHand:
368:   MsgBox "twvMapBook_NodeClick - " & Err.Description
End Sub

Private Sub tvwMapBook_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrHand:
373:   m_lXClick = X
374:   m_lYClick = Y
  
376:   m_lButton = Button
  
  Exit Sub
ErrHand:
380:   MsgBox "tvwMapBook_MouseDown - " & Err.Description
End Sub

Private Sub tvwMapBook_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrHand:
  If Not m_bClickFlag Then Exit Sub
386:   m_bClickFlag = False
387:   m_bNodeFlag = False
  
  Exit Sub
ErrHand:
391:   MsgBox "tvwMapBook_MouseUp - " & Err.Description
End Sub

Private Sub tvwMapBook_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo ErrHand:
  Dim lLoop As Long, pUID As New UID, lImage As Long
  Dim pItem As ICommandItem, lPos As Long, sText As String
  Dim pMapBook As IDSMapBook, pMapSeries As IDSMapSeries
  Dim lPage As Long
  'Check to see if a MapSeries already exists
401:   Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
404:   Set pMapSeries = pMapBook.ContentItem(0)
  
406:   Set m_pCurrentNode = Node
  Select Case Node.Image
  Case 1, 2   'Enable and not Enabled options for a map book
409:     If m_lXClick < 180 Then
410:       If Node.Image = 1 Then
411:         Node.Image = 2
412:         pMapBook.EnableBook = False
'        tvwMapBook.Nodes.Item("MapSeries").Image = 4
'        UpdatePages False
415:       Else
416:         Node.Image = 1
417:         pMapBook.EnableBook = True
'        tvwMapBook.Nodes.Item("MapSeries").Image = 3
'        UpdatePages True
420:       End If
421:     Else
422:       If m_lButton = 2 Then
423:         PopupMenu mnuHeadingBook
424:       End If
425:     End If
  Case 3, 4   'Enable and not Enabled options for a map series
427:     If m_lXClick > 510 And m_lXClick < 760 Then
428:       If Node.Image = 3 Then
429:         Node.Image = 4
430:         pMapSeries.EnableSeries = False
'        UpdatePages False
432:       Else
433:         Node.Image = 3
434:         pMapSeries.EnableSeries = True
'        UpdatePages True
436:       End If
437:     Else
438:       If m_lButton = 2 Then
439:         If Node.Image = 3 Then
440:           mnuSeries(11).Caption = "Disable Series"
441:         Else
442:           mnuSeries(11).Caption = "Enable Series"
443:         End If
444:         PopupMenu mnuHeadingSeries
445:       End If
446:     End If
  Case 5, 6   'Enable and not Enabled options for a map page
448:     If m_lXClick > 1320 Then
449:       If m_lButton = 2 Then
450:         If Node.Image = 5 Then
451:           mnuPage(4).Caption = "Disable Page"
452:         Else
453:           mnuPage(4).Caption = "Enable Page"
454:         End If
455:         PopupMenu mnuHeadingPage
456:       End If
457:     ElseIf m_lXClick > 1080 And m_lXClick <= 1320 Then
458:       lPage = Node.Tag
459:       If Node.Image = 5 Then
460:         Node.Image = 6
461:         pMapSeries.Page(lPage).EnablePage = False
462:       Else
463:         Node.Image = 5
464:         pMapSeries.Page(lPage).EnablePage = True
465:       End If
466:     End If
467:   End Select

  Exit Sub
ErrHand:
471:   MsgBox "twvMapBook_NodeClick - " & Err.Description
End Sub

Private Sub UpdatePages(bEnableFlag As Boolean)
On Error GoTo ErrHand:
  Dim lLoop As Long, pNode As Node
477:   For lLoop = 2 To tvwMapBook.Nodes.count
478:     Set pNode = tvwMapBook.Nodes.Item(lLoop)
479:     If pNode.Image = 5 Or pNode.Image = 6 Then
480:       If bEnableFlag = True Then
481:         pNode.Image = 5
482:       Else
483:         pNode.Image = 6
484:       End If
485:     End If
486:   Next lLoop

  Exit Sub
ErrHand:
490:   MsgBox "UpdatePages - " & Err.Description
End Sub

Public Sub ShowPrinterDialog(pMxApp As IMxApplication, Optional pMapSeries As IDSMapSeries, Optional pPrintMaterial As IUnknown)
  On Error GoTo ErrorHandler
  Dim pFrm As frmPrint

  Dim pPrinter As IPrinter
  Dim pApp As IApplication
  Dim iNumPages As Integer
  Dim pPage As IPage
  Dim pDoc As IMxDocument
  Dim pLayout As IPageLayout
  
504:   Set pPrinter = pMxApp.Printer
505:   If pPrinter Is Nothing Then
506:     MsgBox "You must have at least one printer defined before using this command!!!"
    Exit Sub
508:   End If
  
510:   Set pApp = pMxApp
511:   Set pFrm = New frmPrint
512:   pFrm.Application = pApp
513:   pFrm.ExportFrame = m_pExportFrame
514:   m_pExportFrame.Create pFrm
  
516:   Set pDoc = pApp.Document
517:   Set pLayout = pDoc.PageLayout
518:   Set pPage = pLayout.Page
          
520:   pPage.PrinterPageCount pPrinter, 0, iNumPages
  
522:   pFrm.txtTo.Text = iNumPages
      
524:   pFrm.lblName.Caption = pPrinter.Paper.PrinterName
525:   pFrm.lblType.Caption = pPrinter.DriverName
526:   If TypeOf pPrinter Is IPsPrinter Then
527:     pFrm.chkPrintToFile.Enabled = True
528:   Else
529:     pFrm.chkPrintToFile.Value = 0
530:     pFrm.chkPrintToFile.Enabled = False
531:   End If
  'If pprintmaterial is nothing then it means you are printing a map series
  
534:   If pPrintMaterial Is Nothing Then
535:       pFrm.aDSMapSeries = pMapSeries
536:       pFrm.optPrintCurrentPage.Enabled = False
537:       m_pExportFrame.Visible = True
      Exit Sub
539:   End If
  
541:   If TypeOf pPrintMaterial Is IDSMapBook Then
542:       pFrm.aDSMapBook = pPrintMaterial
543:       pFrm.optPrintCurrentPage.Enabled = False
544:       pFrm.optPrintPages.Enabled = False
545:       pFrm.txtPrintPages.Enabled = False
546:   ElseIf TypeOf pPrintMaterial Is IDSMapPage Then
547:       pFrm.aDSMapPage = pPrintMaterial
548:       pFrm.aDSMapSeries = pMapSeries
549:       pFrm.optPrintCurrentPage.Value = True
550:       pFrm.optPrintAll.Enabled = False
551:       pFrm.optPrintPages.Enabled = False
552:       pFrm.txtPrintPages.Enabled = False
553:   End If
554:   m_pExportFrame.Visible = True
555:   Set pPrintMaterial = Nothing
    
  Exit Sub
ErrorHandler:
559:   MsgBox "ShowPrinterDialog - " & Err.Description
End Sub

Public Sub ShowExporterDialog(pApp As IApplication, Optional pMapSeries As IDSMapSeries, Optional pExportMaterial As IUnknown)
  On Error GoTo ErrorHandler
  Dim pFrm As frmExport
    
566:   Set pFrm = New frmExport
567:   pFrm.Application = pApp
568:   pFrm.ExportFrame = m_pExportFrame
569:   m_pExportFrame.Create pFrm
  
571:   If pExportMaterial Is Nothing Then
572:       pFrm.aDSMapSeries = pMapSeries
573:       pFrm.optCurrentPage.Enabled = False
574:       pFrm.InitializeTheForm
575:       m_pExportFrame.Visible = True
      Exit Sub
577:   End If
  
579:   If TypeOf pExportMaterial Is IDSMapBook Then
580:     pFrm.aDSMapBook = pExportMaterial
581:     pFrm.optCurrentPage.Enabled = False
582:     pFrm.optPages.Enabled = False
583:     pFrm.txtPages.Enabled = False
584:     pFrm.InitializeTheForm
585:   ElseIf TypeOf pExportMaterial Is IDSMapPage Then
586:     pFrm.aDSMapPage = pExportMaterial
587:     pFrm.aDSMapSeries = pMapSeries
588:     pFrm.optCurrentPage.Value = True
589:     pFrm.optAll.Enabled = False
590:     pFrm.optPages.Enabled = False
591:     pFrm.txtPages.Enabled = False
592:     pFrm.InitializeTheForm
593:   End If
594:   m_pExportFrame.Visible = True
595:   Set pExportMaterial = Nothing

  Exit Sub
ErrorHandler:
599:   MsgBox "ShowExporterDialog - " & Err.Description
End Sub



