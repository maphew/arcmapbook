VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMapSeries 
   BackColor       =   &H80000004&
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   165
   ClientTop       =   855
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

' Copyright 2006 ESRI
'
' All rights reserved under the copyright laws of the United States
' and applicable international laws, treaties, and conventions.
'
' You may freely redistribute and use this sample code, with or
' without modification, provided you include the original copyright
' notice and use restrictions.
'
' See use restrictions at /arcgis/developerkit/userestrictions.

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
26:   tvwMapBook.Nodes.Clear
'  tvwMapBook.Nodes.Add , , "MapBook", "Map Book (0 pages)", 1
28:   tvwMapBook.Nodes.Add , , "MapBook", "Map Book", 1
29:   m_bNodeFlag = True
30:   m_bClickFlag = False
31:   m_bLabelingChanged = False
32:   Set m_pExportFrame = New ModelessFrame
End Sub

Private Sub Form_Terminate()
36:   Set m_pExportFrame = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
40:   Set m_pExportFrame = Nothing
End Sub

Private Sub mnuBook_Click(Index As Integer)
On Error GoTo ErrHand:
  Dim pMapBook As IDSMapBook
  'Check to see if a MapSeries already exists
47:   Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
  Select Case Index
  Case 0  'Add Map Series
52:     If pMapBook.ContentCount > 0 Then
53:       MsgBox "You must remove the existing Map Series before adding another."
      Exit Sub
55:     End If
  
    'Call the wizard for setting parameters and creating the series
58:     Set frmMapSeriesWiz.m_pApp = m_pApp
59:     frmMapSeriesWiz.Show vbModal
  Case 1  'Separator
  Case 2  'Print Map Book
62:     ShowPrinterDialog m_pApp, , pMapBook
'    pMapBook.PrintBook
  Case 3  'Export Map Book
65:     ShowExporterDialog m_pApp, , pMapBook
'    pMapBook.ExportBook
67:   End Select
  
  Exit Sub
ErrHand:
71:   MsgBox "mnuBook_Click - " & Err.Description
End Sub

Private Sub mnuPage_Click(Index As Integer)
On Error GoTo ErrHand:
  Dim pMapBook As IDSMapBook, pMapSeries As IDSMapSeries
  Dim lPage As Long, sText As String, lPos As Long, pMapPage As IDSMapPage
  Dim pSeriesOpts As IDSMapSeriesOptions, pSeriesOpts2 As IDSMapSeriesOptions2
  'Check to see if a MapSeries already exists
80:   Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
83:   Set pMapSeries = pMapBook.ContentItem(0)
84:   Set pSeriesOpts = pMapSeries
85:   Set pSeriesOpts2 = pSeriesOpts
86:   lPage = m_pCurrentNode.Tag
  Select Case Index
  Case 0  'View Page
89:     Set pMapPage = pMapSeries.Page(lPage)
90:     pMapPage.DrawPage m_pApp.Document, pMapSeries, True
91:     If pSeriesOpts2.ClipData > 0 Then
92:       g_bClipFlag = True
93:     End If
94:     If pSeriesOpts.RotateFrame Then
95:       g_bRotateFlag = True
96:     End If
97:     If pSeriesOpts.LabelNeighbors Then
98:       g_bLabelNeighbors = True
99:     End If
  Case 1  'Separator
  Case 2  'Delete Page
    'Remove the page, then update the tags on all subsequent pages
103:     pMapSeries.RemovePage lPage
104:     tvwMapBook.Nodes.Remove lPage + 3
105:     RenumberPages pMapSeries
  Case 3  'Separator
  Case 4  'Disable Page
    'Get the index number from the tag of the node
109:     pMapSeries.Page(lPage).EnablePage = Not pMapSeries.Page(lPage).EnablePage
110:     If pMapSeries.Page(lPage).EnablePage Then
111:       m_pCurrentNode.Image = 5
112:     Else
113:       m_pCurrentNode.Image = 6
114:     End If
  Case 5  'Separator
  Case 6  'Print Page
117:     ShowPrinterDialog m_pApp, pMapSeries, pMapSeries.Page(lPage)
  Case 7  'Export Page
119:     ShowExporterDialog m_pApp, pMapSeries, pMapSeries.Page(lPage)
120:   End Select
  
  Exit Sub
ErrHand:
124:   MsgBox "mnuPage_Click - " & Erl & " - " & Err.Description
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
136:   Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
139:   Set pMapSeries = pMapBook.ContentItem(0)
140:   Set pSeriesProps = pMapSeries
141:   Set pDoc = m_pApp.Document
  Select Case Index
  Case 0  'Select Pages
144:     Set frmSelectPages.m_pApp = m_pApp
145:     frmSelectPages.Show vbModal
  Case 1  'Separator
  Case 2  'Tag as Date
148:     bFlag = TagItem(pDoc, "DSMAPBOOK - DATE", "")
149:     If Not bFlag Then
150:       MsgBox "You must have one Text Element selected in the Page Layout for tagging!!!"
151:     End If
  Case 3  'Tag as Title
153:     bFlag = TagItem(pDoc, "DSMAPBOOK - TITLE", "")
154:     If Not bFlag Then
155:       MsgBox "You must have one Text Element selected in the Page Layout for tagging!!!"
156:     End If
  Case 4  'Tag as Page Number
158:     bFlag = TagItem(pDoc, "DSMAPBOOK - PAGENUMBER", "")
159:     If Not bFlag Then
160:       MsgBox "You must have one Text Element selected in the Page Layout for tagging!!!"
161:     End If
  Case 5  'Tag with Index Layer Field...
    'Find the data frame
164:     Set pMap = FindDataFrame(pDoc, pSeriesProps.DataFrameName)
165:     If pMap Is Nothing Then
166:       MsgBox "Could not find map in mnuSeries_Click routine!!!"
      Exit Sub
168:     End If
    'Find the Index layer
170:     Set pIndexLayer = FindLayer(pSeriesProps.IndexLayerName, pMap)
171:     If pIndexLayer Is Nothing Then
172:       MsgBox "Could not find index layer in mnuSeries_Click routine!!!"
      Exit Sub
174:     End If
  
176:     frmTagIndexField.InitializeList pIndexLayer.FeatureClass.Fields
177:     frmTagIndexField.Show vbModal
    
    'Exit sub if Cancel was selected
180:     If frmTagIndexField.m_bCancel Then
181:       Unload frmTagIndexField
      Exit Sub
183:     End If
    
185:     lIndex = frmTagIndexField.lstFields.ListIndex
186:     If lIndex >= 0 Then
187:       sTemp = frmTagIndexField.lstFields.List(lIndex)
188:     Else
189:       MsgBox "You did not pick a field to tag with!!!"
190:       Unload frmTagIndexField
      Exit Sub
192:     End If
193:     Unload frmTagIndexField
    
195:     lIndex = InStr(1, sTemp, " - ")
196:     sName = Mid(sTemp, 1, lIndex - 1)
197:     bFlag = TagItem(pDoc, "DSMAPBOOK - EXTRAITEM", sName)
198:     If Not bFlag Then
199:       MsgBox "You must have one Text Element selected in the Page Layout for tagging!!!"
200:     End If
  Case 6  'Clear Tag for selected
202:     Set pGraphicsCont = pDoc.PageLayout
203:     Set pGraphicsContSel = pDoc.PageLayout
204:     For lLoop = 0 To pGraphicsContSel.ElementSelectionCount - 1
205:       Set pElemProps = pGraphicsContSel.SelectedElement(lLoop)
206:       If TypeOf pElemProps Is ITextElement Then
207:         pElemProps.Name = ""
208:         pElemProps.Type = ""
209:         pGraphicsCont.UpdateElement pTextElement
210:       End If
211:     Next lLoop
  Case 7  'Separator
  Case 8  'Delete Series
214:     Set pActive = pDoc.FocusMap
215:     TurnOffClipping pMapSeries, m_pApp
216:     Set pMapSeries = Nothing
217:     pMapBook.RemoveContent 0
218:     tvwMapBook.Nodes.Clear
219:     tvwMapBook.Nodes.Add , , "MapBook", "Map Book", 1
220:     RemoveIndicators m_pApp
221:     pActive.Refresh
  Case 9  'Delete Disabled pages
    'Loop in reverse order so we remove pages as we work up.  Doing it this way makes
    'sure numbering isn't messed up when a page/node is removed.
225:     For lLoop = pMapSeries.PageCount - 1 To 0 Step -1
226:       If Not pMapSeries.Page(lLoop).EnablePage Then
227:         pMapSeries.RemovePage lLoop
228:         tvwMapBook.Nodes.Remove lLoop + 3
229:       End If
230:     Next lLoop
231:     RenumberPages pMapSeries
  Case 10  'Separator
  Case 11  'Disable Series
    'Get the index number from the tag of the node
235:     pMapSeries.EnableSeries = Not pMapSeries.EnableSeries
236:     If pMapSeries.EnableSeries Then
237:       m_pCurrentNode.Image = 3
238:     Else
239:       m_pCurrentNode.Image = 4
240:     End If
  Case 12  'Separator
  Case 13  'Print Series
243:     ShowPrinterDialog m_pApp, pMapSeries, Nothing
'    pMapSeries.PrintSeries
  Case 14  'Export Series
246:     ShowExporterDialog m_pApp, pMapSeries, Nothing
'    pMapSeries.ExportSeries
  Case 15
249:     Set frmCreateIndex.m_pApp = m_pApp
250:     frmCreateIndex.Show vbModal
  Case 16  'Separator
  Case 17  'Series Properties...
253:     Set frmSeriesProperties.m_pApp = m_pApp
254:     frmSeriesProperties.Show vbModal
  Case 18  'Page Properties...
256:     Set frmPageProperties.m_pApp = m_pApp
257:     frmPageProperties.Show vbModal
258:   End Select
  
  Exit Sub
ErrHand:
262:   MsgBox "mnuSeries_Click - " & Erl & " - " & Err.Description
End Sub

Private Function TagItem(pDoc As IMxDocument, sName As String, sType As String) As Boolean
On Error GoTo ErrHand:
  Dim bFlag As Boolean, pGraphicsCont As IGraphicsContainer, pActive As IActiveView
  Dim pElemProps As IElementProperties, pElem As IElement, pTextElement As ITextElement
  Dim pEnv2 As IEnvelope, pGraphicsContSel As IGraphicsContainerSelect, pEnv As IEnvelope
  
271:   Set pGraphicsCont = pDoc.PageLayout
272:   Set pGraphicsContSel = pDoc.PageLayout
273:   bFlag = False
274:   If pGraphicsContSel.ElementSelectionCount = 1 Then
275:     Set pElemProps = pGraphicsContSel.SelectedElement(0)
276:     If TypeOf pElemProps Is ITextElement Then
277:       Set pActive = pDoc.PageLayout
278:       pElemProps.Name = sName
279:       Set pElem = pElemProps
280:       Set pEnv = New Envelope
281:       pElem.QueryBounds pActive.ScreenDisplay, pEnv
282:       Set pTextElement = pElemProps
      Select Case sName
      Case "DSMAPBOOK - DATE"
285:         pTextElement.Text = Format(Date, "mmm dd, yyyy")
      Case "DSMAPBOOK - TITLE"
287:         pTextElement.Text = "Title String"
      Case "DSMAPBOOK - PAGENUMBER"
289:         pTextElement.Text = "PAGE #"
      Case "DSMAPBOOK - EXTRAITEM"
291:         pTextElement.Text = sType
292:         pElemProps.Type = sType
293:       End Select
294:       pGraphicsCont.UpdateElement pTextElement
295:       Set pEnv2 = New Envelope
296:       pElem.QueryBounds pActive.ScreenDisplay, pEnv2
297:       pEnv.Union pEnv2
298:       pActive.PartialRefresh esriViewGraphics, Nothing, pEnv
299:       bFlag = True
300:     End If
301:   End If
  
303:   TagItem = bFlag

  Exit Function
ErrHand:
307:   MsgBox "TagItem - " & Erl & " - " & Err.Description
308:   TagItem = bFlag
End Function

Private Sub RenumberPages(pMapSeries As IDSMapSeries)
On Error GoTo ErrHand:
'Routine for renumber the pages after one is removed
  Dim lLoop As Long, pNode As Node, sName As String, lPageNumber As Long
  Dim pPage As IDSMapPage, pSeriesProps As IDSMapSeriesProps
316:   Set pSeriesProps = pMapSeries
317:   For lLoop = 0 To pMapSeries.PageCount - 1
318:     lPageNumber = lLoop + pSeriesProps.StartNumber
319:     Set pPage = pMapSeries.Page(lLoop)
320:     Set pNode = tvwMapBook.Nodes.Item(lLoop + 3)
321:     sName = Mid(pNode.Key, 2)
322:     pNode.Tag = lLoop
323:     pNode.Key = "a" & sName
324:     pNode.Text = lPageNumber & " - " & sName
325:     pPage.PageNumber = lPageNumber
326:   Next lLoop
327:   tvwMapBook.Refresh
  
  Exit Sub
ErrHand:
331:   MsgBox "RenumberPages - " & Erl & " - " & Err.Description
End Sub

Private Sub picBook_Resize()
335:   tvwMapBook.Width = picBook.Width
336:   tvwMapBook.Height = picBook.Height
End Sub

Private Sub tvwMapBook_DblClick()
On Error GoTo ErrHand:
  Dim lPos As String, sText As String, pMapPage As IDSMapPage, lPage As Long
  Dim pSeriesOpts As IDSMapSeriesOptions, pSeriesOpts2 As IDSMapSeriesOptions2
  Dim pMapBook As IDSMapBook, pMapSeries As IDSMapSeries
344:   Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
347:   Set pMapSeries = pMapBook.ContentItem(0)
348:   Set pSeriesOpts = pMapSeries
349:   Set pSeriesOpts2 = pSeriesOpts
  
  'There is no NodeDoubleClick event, so we have to use the DblClick event on the control
  'and check to make sure we are over a node.  In the event of a doubleclick on a node,
  'the order of events being fired are NodeClick, MouseUp, Click, DblClick, MouseUp.  To
  'make sure the doubleclick occurred over a node, we can set a flag in the NodeClick event
  'and then disable it in the MouseUp event after the Click event.
  If Not m_bNodeFlag Then Exit Sub
  
  Select Case m_pCurrentNode.Image
  Case 5, 6   'Enable and not Enabled options for a map page
360:     If m_lXClick > 1320 Then
361:       If m_lButton = 1 Then
362:         lPage = m_pCurrentNode.Tag
363:         Set pMapPage = pMapSeries.Page(lPage)
364:         pMapPage.DrawPage m_pApp.Document, pMapSeries, True
365:         If pSeriesOpts2.ClipData > 0 Then
366:           g_bClipFlag = True
367:         End If
368:         If pSeriesOpts.RotateFrame Then
369:           g_bRotateFlag = True
370:         End If
371:         If pSeriesOpts.LabelNeighbors Then
372:           g_bLabelNeighbors = True
373:         End If
374:       End If
375:     End If
376:   End Select

  Exit Sub
ErrHand:
380:   MsgBox "twvMapBook_NodeClick - " & Err.Description
End Sub

Private Sub tvwMapBook_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrHand:
385:   m_lXClick = X
386:   m_lYClick = Y
  
388:   m_lButton = Button
  
  Exit Sub
ErrHand:
392:   MsgBox "tvwMapBook_MouseDown - " & Err.Description
End Sub

Private Sub tvwMapBook_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrHand:
  If Not m_bClickFlag Then Exit Sub
398:   m_bClickFlag = False
399:   m_bNodeFlag = False
  
  Exit Sub
ErrHand:
403:   MsgBox "tvwMapBook_MouseUp - " & Err.Description
End Sub

Private Sub tvwMapBook_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo ErrHand:
  Dim lLoop As Long, pUID As New UID, lImage As Long
  Dim pItem As ICommandItem, lPos As Long, sText As String
  Dim pMapBook As IDSMapBook, pMapSeries As IDSMapSeries
  Dim lPage As Long
  'Check to see if a MapSeries already exists
413:   Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
416:   Set pMapSeries = pMapBook.ContentItem(0)
  
418:   Set m_pCurrentNode = Node
  Select Case Node.Image
  Case 1, 2   'Enable and not Enabled options for a map book
421:     If m_lXClick < 180 Then
422:       If Node.Image = 1 Then
423:         Node.Image = 2
424:         pMapBook.EnableBook = False
'        tvwMapBook.Nodes.Item("MapSeries").Image = 4
'        UpdatePages False
427:       Else
428:         Node.Image = 1
429:         pMapBook.EnableBook = True
'        tvwMapBook.Nodes.Item("MapSeries").Image = 3
'        UpdatePages True
432:       End If
433:     Else
434:       If m_lButton = 2 Then
435:         PopupMenu mnuHeadingBook
436:       End If
437:     End If
  Case 3, 4   'Enable and not Enabled options for a map series
439:     If m_lXClick > 510 And m_lXClick < 760 Then
440:       If Node.Image = 3 Then
441:         Node.Image = 4
442:         pMapSeries.EnableSeries = False
'        UpdatePages False
444:       Else
445:         Node.Image = 3
446:         pMapSeries.EnableSeries = True
'        UpdatePages True
448:       End If
449:     Else
450:       If m_lButton = 2 Then
451:         If Node.Image = 3 Then
452:           mnuSeries(11).Caption = "Disable Series"
453:         Else
454:           mnuSeries(11).Caption = "Enable Series"
455:         End If
456:         PopupMenu mnuHeadingSeries
457:       End If
458:     End If
  Case 5, 6   'Enable and not Enabled options for a map page
460:     If m_lXClick > 1320 Then
461:       If m_lButton = 2 Then
462:         If Node.Image = 5 Then
463:           mnuPage(4).Caption = "Disable Page"
464:         Else
465:           mnuPage(4).Caption = "Enable Page"
466:         End If
467:         PopupMenu mnuHeadingPage
468:       End If
469:     ElseIf m_lXClick > 1080 And m_lXClick <= 1320 Then
470:       lPage = Node.Tag
471:       If Node.Image = 5 Then
472:         Node.Image = 6
473:         pMapSeries.Page(lPage).EnablePage = False
474:       Else
475:         Node.Image = 5
476:         pMapSeries.Page(lPage).EnablePage = True
477:       End If
478:     End If
479:   End Select

  Exit Sub
ErrHand:
483:   MsgBox "twvMapBook_NodeClick - " & Err.Description
End Sub

Private Sub UpdatePages(bEnableFlag As Boolean)
On Error GoTo ErrHand:
  Dim lLoop As Long, pNode As Node
489:   For lLoop = 2 To tvwMapBook.Nodes.count
490:     Set pNode = tvwMapBook.Nodes.Item(lLoop)
491:     If pNode.Image = 5 Or pNode.Image = 6 Then
492:       If bEnableFlag = True Then
493:         pNode.Image = 5
494:       Else
495:         pNode.Image = 6
496:       End If
497:     End If
498:   Next lLoop

  Exit Sub
ErrHand:
502:   MsgBox "UpdatePages - " & Err.Description
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
  
516:   Set pPrinter = pMxApp.Printer
517:   If pPrinter Is Nothing Then
518:     MsgBox "You must have at least one printer defined before using this command!!!"
    Exit Sub
520:   End If
  
522:   Set pApp = pMxApp
523:   Set pFrm = New frmPrint
524:   pFrm.Application = pApp
525:   pFrm.ExportFrame = m_pExportFrame
526:   m_pExportFrame.Create pFrm
  
528:   Set pDoc = pApp.Document
529:   Set pLayout = pDoc.PageLayout
530:   Set pPage = pLayout.Page
          
532:   pPage.PrinterPageCount pPrinter, 0, iNumPages
  
534:   pFrm.txtTo.Text = iNumPages
      
536:   pFrm.lblName.Caption = pPrinter.Paper.PrinterName
537:   pFrm.lblType.Caption = pPrinter.DriverName
538:   If TypeOf pPrinter Is IPsPrinter Then
539:     pFrm.chkPrintToFile.Enabled = True
540:   Else
541:     pFrm.chkPrintToFile.value = 0
542:     pFrm.chkPrintToFile.Enabled = False
543:   End If
  'If pprintmaterial is nothing then it means you are printing a map series
  
546:   If pPrintMaterial Is Nothing Then
547:       pFrm.aDSMapSeries = pMapSeries
548:       pFrm.optPrintCurrentPage.Enabled = False
549:       m_pExportFrame.Visible = True
      Exit Sub
551:   End If
  
553:   If TypeOf pPrintMaterial Is IDSMapBook Then
554:       pFrm.aDSMapBook = pPrintMaterial
555:       pFrm.optPrintCurrentPage.Enabled = False
556:       pFrm.optPrintPages.Enabled = False
557:       pFrm.txtPrintPages.Enabled = False
558:   ElseIf TypeOf pPrintMaterial Is IDSMapPage Then
559:       pFrm.aDSMapPage = pPrintMaterial
560:       pFrm.aDSMapSeries = pMapSeries
561:       pFrm.optPrintCurrentPage.value = True
562:       pFrm.optPrintAll.Enabled = False
563:       pFrm.optPrintPages.Enabled = False
564:       pFrm.txtPrintPages.Enabled = False
565:   End If
566:   m_pExportFrame.Visible = True
567:   Set pPrintMaterial = Nothing
    
  Exit Sub
ErrorHandler:
571:   MsgBox "ShowPrinterDialog - " & Err.Description
End Sub

Public Sub ShowExporterDialog(pApp As IApplication, Optional pMapSeries As IDSMapSeries, Optional pExportMaterial As IUnknown)
  On Error GoTo ErrorHandler
  Dim pFrm As frmExport
    
578:   Set pFrm = New frmExport
579:   pFrm.Application = pApp
580:   pFrm.ExportFrame = m_pExportFrame
581:   m_pExportFrame.Create pFrm
  
583:   If pExportMaterial Is Nothing Then
584:       pFrm.aDSMapSeries = pMapSeries
585:       pFrm.optCurrentPage.Enabled = False
586:       pFrm.InitializeTheForm
587:       m_pExportFrame.Visible = True
      Exit Sub
589:   End If
  
591:   If TypeOf pExportMaterial Is IDSMapBook Then
592:     pFrm.aDSMapBook = pExportMaterial
593:     pFrm.optCurrentPage.Enabled = False
594:     pFrm.optPages.Enabled = False
595:     pFrm.txtPages.Enabled = False
596:     pFrm.InitializeTheForm
597:   ElseIf TypeOf pExportMaterial Is IDSMapPage Then
598:     pFrm.aDSMapPage = pExportMaterial
599:     pFrm.aDSMapSeries = pMapSeries
600:     pFrm.optCurrentPage.value = True
601:     pFrm.optAll.Enabled = False
602:     pFrm.optPages.Enabled = False
603:     pFrm.txtPages.Enabled = False
604:     pFrm.InitializeTheForm
605:   End If
606:   m_pExportFrame.Visible = True
607:   Set pExportMaterial = Nothing

  Exit Sub
ErrorHandler:
611:   MsgBox "ShowExporterDialog - " & Err.Description
End Sub



