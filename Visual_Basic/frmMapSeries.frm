VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMapSeries 
   BackColor       =   &H80000004&
   Caption         =   "Form1"
   ClientHeight    =   5064
   ClientLeft      =   132
   ClientTop       =   684
   ClientWidth     =   4272
   LinkTopic       =   "Form1"
   ScaleHeight     =   5064
   ScaleWidth      =   4272
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBook 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   150
      ScaleHeight     =   4332
      ScaleWidth      =   3888
      TabIndex        =   1
      Top             =   630
      Width           =   3885
      Begin MSComctlLib.TreeView tvwMapBook 
         Height          =   4725
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   3855
         _ExtentX        =   6795
         _ExtentY        =   8340
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
      Height          =   1200
      Left            =   2760
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   3420
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   995
      _ExtentY        =   995
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
         Caption         =   "Tag as Visibility Managed"
         Index           =   6
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "Clear Tag for Selected"
         Index           =   7
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "Delete Series"
         Index           =   9
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "Delete Disabled Pages"
         Index           =   10
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "Disable Series"
         Index           =   12
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "Print Series..."
         Index           =   14
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "Export Series..."
         Index           =   15
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "Create Series Index..."
         Index           =   16
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "Visible Dataframes"
         Index           =   17
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "Visibility of Element"
         Index           =   18
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "-"
         Index           =   19
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "Series Properties..."
         Index           =   20
      End
      Begin VB.Menu mnuSeries 
         Caption         =   "Page Properties..."
         Index           =   21
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
         Caption         =   "Adjacent Page Symbol"
         Index           =   6
         Begin VB.Menu mnuAdjacentSymbol 
            Caption         =   "Default Symbol"
            Index           =   0
         End
         Begin VB.Menu mnuAdjacentSymbol 
            Caption         =   "Select Other ..."
            Index           =   1
         End
      End
      Begin VB.Menu mnuPage 
         Caption         =   "Layer Visibility"
         Index           =   7
      End
      Begin VB.Menu mnuPage 
         Caption         =   "Visible Elements"
         Index           =   8
      End
      Begin VB.Menu mnuPage 
         Caption         =   "Print Page..."
         Index           =   9
      End
      Begin VB.Menu mnuPage 
         Caption         =   "Export Page..."
         Index           =   10
      End
   End
End
Attribute VB_Name = "frmMapSeries"
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
Const c_sModuleFileName As String = "frmMapSeries.frm"


Private Sub Form_Load()
  On Error GoTo ErrorHandler

44:   tvwMapBook.Nodes.Clear
'  tvwMapBook.Nodes.Add , , "MapBook", "Map Book (0 pages)", 1
46:   tvwMapBook.Nodes.Add , , "MapBook", "Map Book", 1
47:   m_bNodeFlag = True
48:   m_bClickFlag = False
49:   m_bLabelingChanged = False
50:   Set m_pExportFrame = New ModelessFrame

  Exit Sub
ErrorHandler:
  HandleError True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Form_Terminate()
54:   Set m_pExportFrame = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
58:   Set m_pExportFrame = Nothing
End Sub

Private Sub mnuAdjacentSymbol_Click(Index As Integer)
  On Error GoTo ErrorHandler

  Dim pMapBook As INWDSMapBook, pMapSeries As INWDSMapSeries
  Dim lPage As Long, sText As String, lPos As Long, pMapPage As INWDSMapPage
  Dim pSeriesOpts As INWDSMapSeriesOptions, pSeriesOpts2 As INWDSMapSeriesOptions2
  Dim pNWSeriesOpts As INWMapSeriesOptions, pNWMapPageAttrs As INWMapPageAttribs
  'Check to see if a MapSeries already exists
69:   Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
72:   Set pMapSeries = pMapBook.ContentItem(0)
73:   Set pSeriesOpts = pMapSeries
74:   Set pSeriesOpts2 = pSeriesOpts
75:   Set pNWSeriesOpts = pMapSeries
  
77:   lPage = m_pCurrentNode.Tag
78:   Set pMapPage = pMapSeries.Page(lPage)
79:   Set pNWMapPageAttrs = pMapPage
  
  Select Case Index
  Case 0  'Default Symbol
                                                  'blank string is interpreted
                                                  'as the default symbol
85:     pNWMapPageAttrs.AdjacentLabelSymbol = ""
  Case 1  'Select Other...
87:     Set frmSelectAdjMapSymbol.NWSeriesOptions = pNWSeriesOpts
88:     frmSelectAdjMapSymbol.CurrentSymbol = pNWMapPageAttrs.AdjacentLabelSymbol
89:     frmSelectAdjMapSymbol.Show vbModal
90:     pNWMapPageAttrs.AdjacentLabelSymbol = frmSelectAdjMapSymbol.CurrentSymbol
91:   End Select

  Exit Sub
ErrorHandler:
  HandleError True, "mnuAdjacentSymbol_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub mnuBook_Click(Index As Integer)
On Error GoTo ErrHand:
  Dim pMapBook As INWDSMapBook
  'Check to see if a MapSeries already exists
102:   Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
  Select Case Index
  Case 0  'Add Map Series
107:     If pMapBook.ContentCount > 0 Then
108:       MsgBox "You must remove the existing Map Series before adding another."
      Exit Sub
110:     End If
  
    'Call the wizard for setting parameters and creating the series
113:     Set frmMapSeriesWiz.m_pApp = m_pApp
114:     frmMapSeriesWiz.Show vbModal
  Case 1  'Separator
  Case 2  'Print Map Book
117:     ShowPrinterDialog m_pApp, , pMapBook
'    pMapBook.PrintBook
  Case 3  'Export Map Book
120:     ShowExporterDialog m_pApp, , pMapBook
'    pMapBook.ExportBook
122:   End Select
  
  Exit Sub
ErrHand:
126:   MsgBox "mnuBook_Click - " & Err.Description
End Sub







Private Sub mnuPage_Click(Index As Integer)
On Error GoTo ErrHand:
  Dim pMapBook As INWDSMapBook, pMapSeries As INWDSMapSeries
  Dim lPage As Long, sText As String, lPos As Long, pMapPage As INWDSMapPage
  Dim pSeriesOpts As INWDSMapSeriesOptions, pSeriesOpts2 As INWDSMapSeriesOptions2
  Dim pNWSeriesOpts As INWMapSeriesOptions
  'Check to see if a MapSeries already exists
142:   Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
145:   Set pMapSeries = pMapBook.ContentItem(0)
146:   Set pSeriesOpts = pMapSeries
147:   Set pSeriesOpts2 = pSeriesOpts
148:   Set pNWSeriesOpts = pMapSeries
  
150:   lPage = m_pCurrentNode.Tag
  Select Case Index
  Case 0  'View Page
153:     Set pMapPage = pMapSeries.Page(lPage)
154:     pMapPage.DrawPage m_pApp.Document, pMapSeries, True
155:     If pSeriesOpts2.ClipData > 0 Then
156:       g_bClipFlag = True
157:     End If
158:     If pSeriesOpts.RotateFrame Then
159:       g_bRotateFlag = True
160:     End If
161:     If pSeriesOpts.LabelNeighbors Then
162:       g_bLabelNeighbors = True
163:     End If
  Case 1  'Separator
  Case 2  'Delete Page
    'Remove the page, then update the tags on all subsequent pages
167:     pMapSeries.RemovePage lPage
168:     tvwMapBook.Nodes.Remove lPage + 3
169:     RenumberPages pMapSeries
  Case 3  'Separator
  Case 4  'Disable Page
    'Get the index number from the tag of the node
173:     pMapSeries.Page(lPage).EnablePage = Not pMapSeries.Page(lPage).EnablePage
174:     If pMapSeries.Page(lPage).EnablePage Then
175:       m_pCurrentNode.Image = 5
176:     Else
177:       m_pCurrentNode.Image = 6
178:     End If
  Case 5  'Separator
  Case 6  'Adjacent Page Symbol
  Case 7  'Layer Visibility
182:     Set frmVisibleLayers.m_pApp = m_pApp
183:     Set frmVisibleLayers.m_pMapPage = pMapSeries.Page(lPage)
184:     frmVisibleLayers.Init_Form
185:     frmVisibleLayers.Show vbModal
  Case 8  'Element Visibility
187:     Set pMapPage = pMapSeries.Page(lPage)
188:     frmVisibleElements.Initialize m_pApp, pNWSeriesOpts, pMapPage.PageName
189:     frmVisibleElements.Show vbModal
  Case 9  'Print Page
191:     ShowPrinterDialog m_pApp, pMapSeries, pMapSeries.Page(lPage)
  Case 10 'Export Page
193:     ShowExporterDialog m_pApp, pMapSeries, pMapSeries.Page(lPage)
194:   End Select
  
  Exit Sub
ErrHand:
198:   MsgBox "mnuPage_Click - " & Erl & " - " & Err.Description
End Sub







Private Sub mnuSeries_Click(Index As Integer)
On Error GoTo ErrHand:
  Dim pMapBook As INWDSMapBook, pMapSeries As INWDSMapSeries, pSeriesProps As INWDSMapSeriesProps
  Dim lLoop As Long, pDoc As IMxDocument, pActive As IActiveView, bFlag As Boolean
  Dim pGraphicsCont As IGraphicsContainer, pElemProps As IElementProperties
  Dim pEnv As IEnvelope, pElem As IElement, pTextElement As ITextElement, pEnv2 As IEnvelope
  Dim pGraphicsContSel As IGraphicsContainerSelect, pMap As IMap
  Dim pIndexLayer As IFeatureLayer, lIndex As Long, sName As String, sTemp As String
  Dim pNWSeriesOpts As INWMapSeriesOptions
  
  'Check to see if a MapSeries already exists
218:   Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
221:   Set pMapSeries = pMapBook.ContentItem(0)
222:   Set pSeriesProps = pMapSeries
223:   Set pNWSeriesOpts = pMapSeries
224:   Set pDoc = m_pApp.Document
  Select Case Index
  Case 0  'Select Pages
227:     Set frmSelectPages.m_pApp = m_pApp
228:     frmSelectPages.Show vbModal
  Case 1  'Separator
  Case 2  'Tag as Date
231:     bFlag = TagItem(pDoc, "NWDSMAPBOOK - DATE", "")
232:     If Not bFlag Then
233:       MsgBox "You must have one Text Element selected in the Page Layout for tagging!!!"
234:     End If
  Case 3  'Tag as Title
236:     bFlag = TagItem(pDoc, "NWDSMAPBOOK - TITLE", "")
237:     If Not bFlag Then
238:       MsgBox "You must have one Text Element selected in the Page Layout for tagging!!!"
239:     End If
  Case 4  'Tag as Page Number
241:     bFlag = TagItem(pDoc, "NWDSMAPBOOK - PAGENUMBER", "")
242:     If Not bFlag Then
243:       MsgBox "You must have one Text Element selected in the Page Layout for tagging!!!"
244:     End If
  Case 5  'Tag with Index Layer Field...
    'Find the data frame
247:     Set pMap = FindDataFrame(pDoc, pSeriesProps.DataFrameName)
248:     If pMap Is Nothing Then
249:       MsgBox "Could not find map in mnuSeries_Click routine!!!"
      Exit Sub
251:     End If
    'Find the Index layer
253:     Set pIndexLayer = FindLayer(pSeriesProps.IndexLayerName, pMap)
254:     If pIndexLayer Is Nothing Then
255:       MsgBox "Could not find index layer in mnuSeries_Click routine!!!"
      Exit Sub
257:     End If
  
259:     frmTagIndexField.InitializeList pIndexLayer.FeatureClass.Fields
260:     frmTagIndexField.Show vbModal
    
    'Exit sub if Cancel was selected
263:     If frmTagIndexField.m_bCancel Then
264:       Unload frmTagIndexField
      Exit Sub
266:     End If
    
268:     lIndex = frmTagIndexField.lstFields.ListIndex
269:     If lIndex >= 0 Then
270:       sTemp = frmTagIndexField.lstFields.List(lIndex)
271:     Else
272:       MsgBox "You did not pick a field to tag with!!!"
273:       Unload frmTagIndexField
      Exit Sub
275:     End If
276:     Unload frmTagIndexField
    
278:     lIndex = InStr(1, sTemp, " - ")
279:     sName = Mid(sTemp, 1, lIndex - 1)
280:     bFlag = TagItem(pDoc, "NWDSMAPBOOK - EXTRAITEM", sName)
281:     If Not bFlag Then
282:       MsgBox "You must have one Text Element selected in the Page Layout for tagging!!!"
283:     End If
  Case 6  'Tag as Visibility Managed
285:     pNWSeriesOpts.ElementsTagElement pDoc
  Case 7  'Clear Tag for selected
    'clear layout element visibility tags (doesn't disturb any other tags)
288:     pNWSeriesOpts.ElementsUntagElement pDoc
    'clear text element tags
290:     Set pGraphicsCont = pDoc.PageLayout
291:     Set pGraphicsContSel = pDoc.PageLayout
292:     For lLoop = 0 To pGraphicsContSel.ElementSelectionCount - 1
293:       Set pElemProps = pGraphicsContSel.SelectedElement(lLoop)
294:       If TypeOf pElemProps Is ITextElement Then
295:         pElemProps.Name = ""
296:         pElemProps.Type = ""
297:         pGraphicsCont.UpdateElement pTextElement
298:       End If
299:     Next lLoop
  Case 8  'Separator
  Case 9  'Delete Series
302:     Set pActive = pDoc.FocusMap
303:     TurnOffClipping pMapSeries, m_pApp
304:     Set pMapSeries = Nothing
305:     pMapBook.RemoveContent 0
306:     tvwMapBook.Nodes.Clear
307:     tvwMapBook.Nodes.Add , , "MapBook", "Map Book", 1
308:     RemoveIndicators m_pApp
309:     pActive.Refresh
  Case 10  'Delete Disabled pages
    'Loop in reverse order so we remove pages as we work up.  Doing it this way makes
    'sure numbering isn't messed up when a page/node is removed.
313:     For lLoop = pMapSeries.PageCount - 1 To 0 Step -1
314:       If Not pMapSeries.Page(lLoop).EnablePage Then
315:         pMapSeries.RemovePage lLoop
316:         tvwMapBook.Nodes.Remove lLoop + 3
317:       End If
318:     Next lLoop
319:     RenumberPages pMapSeries
  Case 11  'Separator
  Case 12  'Disable Series
    'Get the index number from the tag of the node
323:     pMapSeries.EnableSeries = Not pMapSeries.EnableSeries
324:     If pMapSeries.EnableSeries Then
325:       m_pCurrentNode.Image = 3
326:     Else
327:       m_pCurrentNode.Image = 4
328:     End If
  Case 13  'Separator
  Case 14  'Print Series
331:     ShowPrinterDialog m_pApp, pMapSeries, Nothing
'    pMapSeries.PrintSeries
  Case 15  'Export Series
334:     ShowExporterDialog m_pApp, pMapSeries, Nothing
'    pMapSeries.ExportSeries
  Case 16
337:     Set frmCreateIndex.m_pApp = m_pApp
338:     frmCreateIndex.Show vbModal
  Case 17  'Visible Dataframes
340:     Set frmManageDataFrames.App = m_pApp
341:     frmManageDataFrames.Initialize
342:     frmManageDataFrames.Show vbModal, Me
  Case 18  'Visibility of Element
    Dim pPageLayout As IPageLayout, pElements As IEnumElement
    Dim pGraphicsContSelect As IGraphicsContainerSelect, pElement As IElement
  
    'access the selected element
348:     pDoc.ActiveView.Refresh
349:     Set pPageLayout = pDoc.PageLayout
350:     Set pGraphicsCont = pPageLayout
351:     Set pGraphicsContSelect = pGraphicsCont
    
353:     If pGraphicsContSelect.ElementSelectionCount = 0 Then
354:       MsgBox "Warning: At least one layout element must be selected before it can be tagged." & vbNewLine
      Exit Sub
356:     End If
357:     If pGraphicsContSelect.ElementSelectionCount > 1 Then
358:       MsgBox "Warning: more than one layout element was selected.  Select only one layout element" & vbNewLine _
           & "before controlling visibility of that element." & vbNewLine, vbOKOnly, "Must select one element"
      Exit Sub
361:     End If
362:     Set pElements = pGraphicsContSelect.SelectedElements
363:     pElements.Reset
364:     Set pElement = pElements.Next
    
366:     If pElement Is Nothing Then 'redundant to selection count= 0, but good to be defensive
367:       MsgBox "Warning: A layout element must be selected before managing the visibility of a layout element.", _
           vbCritical, "No element was selected."
      Exit Sub
370:     End If
371:     If Not pNWSeriesOpts.ElementsElementIsTagged(pElement) Then
372:       MsgBox "Error: Element should be tagged for element visibility management" & vbNewLine _
           & "before accessing this function.  Please tag the element, then try again." & vbNewLine, _
           vbOKOnly, "Element was not tagged"
      Exit Sub
376:     End If
    
378:     frmPagesWhereElemIsVisible.Initialize m_pApp, pNWSeriesOpts, pElement
379:     frmPagesWhereElemIsVisible.Show vbModal, Me
  Case 19  'Separator
  Case 20  'Series Properties...
382:     Set frmSeriesProperties.m_pApp = m_pApp
383:     frmSeriesProperties.Show vbModal, Me
  Case 21  'Page Properties...
385:     Set frmPageProperties.m_pApp = m_pApp
386:     frmPageProperties.Show vbModal, Me
387:   End Select
  
  Exit Sub
ErrHand:
391:   MsgBox "mnuSeries_Click - " & Erl & " - " & Err.Description
End Sub

Private Function TagItem(pDoc As IMxDocument, sName As String, sType As String) As Boolean
On Error GoTo ErrHand:
  Dim bFlag As Boolean, pGraphicsCont As IGraphicsContainer, pActive As IActiveView
  Dim pElemProps As IElementProperties, pElem As IElement, pTextElement As ITextElement
  Dim pEnv2 As IEnvelope, pGraphicsContSel As IGraphicsContainerSelect, pEnv As IEnvelope
  
400:   Set pGraphicsCont = pDoc.PageLayout
401:   Set pGraphicsContSel = pDoc.PageLayout
402:   bFlag = False
403:   If pGraphicsContSel.ElementSelectionCount = 1 Then
404:     Set pElemProps = pGraphicsContSel.SelectedElement(0)
405:     If TypeOf pElemProps Is ITextElement Then
406:       Set pActive = pDoc.PageLayout
407:       pElemProps.Name = sName
408:       Set pElem = pElemProps
409:       Set pEnv = New envelope
410:       pElem.QueryBounds pActive.ScreenDisplay, pEnv
411:       Set pTextElement = pElemProps
      Select Case sName
      Case "NWDSMAPBOOK - DATE"
414:         pTextElement.Text = Format(Date, "mmm dd, yyyy")
      Case "NWDSMAPBOOK - TITLE"
416:         pTextElement.Text = "Title String"
      Case "NWDSMAPBOOK - PAGENUMBER"
418:         pTextElement.Text = "PAGE #"
      Case "NWDSMAPBOOK - EXTRAITEM"
420:         pTextElement.Text = sType
421:         pElemProps.Type = sType
422:       End Select
423:       pGraphicsCont.UpdateElement pTextElement
424:       Set pEnv2 = New envelope
425:       pElem.QueryBounds pActive.ScreenDisplay, pEnv2
426:       pEnv.Union pEnv2
427:       pActive.PartialRefresh esriViewGraphics, Nothing, pEnv
428:       bFlag = True
429:     End If
430:   End If
  
432:   TagItem = bFlag

  Exit Function
ErrHand:
436:   MsgBox "TagItem - " & Erl & " - " & Err.Description
437:   TagItem = bFlag
End Function


'
'Private Function TrackPreviousPage(pMapPage As INWDSMapPage, pNWSeriesOpts As INWMapSeriesOptions)
'  pMapPage.PageName
'End Function

Private Sub RenumberPages(pMapSeries As INWDSMapSeries)
On Error GoTo ErrHand:
'Routine for renumber the pages after one is removed
  Dim lLoop As Long, pNode As Node, sName As String, lPageNumber As Long
  Dim pPage As INWDSMapPage, pSeriesProps As INWDSMapSeriesProps
451:   Set pSeriesProps = pMapSeries
452:   For lLoop = 0 To pMapSeries.PageCount - 1
453:     lPageNumber = lLoop + pSeriesProps.StartNumber
454:     Set pPage = pMapSeries.Page(lLoop)
455:     Set pNode = tvwMapBook.Nodes.Item(lLoop + 3)
456:     sName = Mid(pNode.Key, 2)
457:     pNode.Tag = lLoop
458:     pNode.Key = "a" & sName
459:     pNode.Text = lPageNumber & " - " & sName
460:     pPage.PageNumber = lPageNumber
461:   Next lLoop
462:   tvwMapBook.Refresh
  
  Exit Sub
ErrHand:
466:   MsgBox "RenumberPages - " & Erl & " - " & Err.Description
End Sub

Private Sub picBook_Resize()
470:   tvwMapBook.Width = picBook.Width
471:   tvwMapBook.Height = picBook.Height
End Sub

Private Sub tvwMapBook_DblClick()
On Error GoTo ErrHand:
  Dim lPos As String, sText As String, pMapPage As INWDSMapPage, lPage As Long
  Dim pSeriesOpts As INWDSMapSeriesOptions, pSeriesOpts2 As INWDSMapSeriesOptions2
  Dim pMapBook As INWDSMapBook, pMapSeries As INWDSMapSeries, sPrevPage As String
  Dim pNWSeriesOpts As INWMapSeriesOptions
480:   Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
483:   Set pMapSeries = pMapBook.ContentItem(0)
484:   Set pSeriesOpts = pMapSeries
485:   Set pSeriesOpts2 = pSeriesOpts
486:   Set pNWSeriesOpts = pMapSeries
  
  
  'There is no NodeDoubleClick event, so we have to use the DblClick event on the control
  'and check to make sure we are over a node.  In the event of a doubleclick on a node,
  'the order of events being fired are NodeClick, MouseUp, Click, DblClick, MouseUp.  To
  'make sure the doubleclick occurred over a node, we can set a flag in the NodeClick event
  'and then disable it in the MouseUp event after the Click event.
  If Not m_bNodeFlag Then Exit Sub
  
  Select Case m_pCurrentNode.Image
  Case 5, 6   'Enable and not Enabled options for a map page
498:     If m_lXClick > 1320 Then
499:       If m_lButton = 1 Then
500:         lPage = m_pCurrentNode.Tag
        'added to support dynamic definition queries -- previous
        'and current page name must be tracked.
503:         If pNWSeriesOpts.DynamicDefQueryReplaceString = "" Then
504:           If Not pMapPage Is Nothing Then
505:             pNWSeriesOpts.DynamicDefQueryReplaceString = pMapPage.PageName
506:           End If
507:         End If
508:         Set pMapPage = pMapSeries.Page(lPage)
509:         pNWSeriesOpts.DynamicDefQueryReplaceString = pMapPage.PageName
510:         pMapPage.DrawPage m_pApp.Document, pMapSeries, True
511:         If pSeriesOpts2.ClipData > 0 Then
512:           g_bClipFlag = True
513:         End If
514:         If pSeriesOpts.RotateFrame Then
515:           g_bRotateFlag = True
516:         End If
517:         If pSeriesOpts.LabelNeighbors Then
518:           g_bLabelNeighbors = True
519:         End If
520:       End If
521:     End If
522:   End Select

  Exit Sub
ErrHand:
526:   MsgBox "twvMapBook_NodeClick - " & Erl & ": " & Err.Description
End Sub

Private Sub tvwMapBook_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo ErrHand:
531:   m_lXClick = x
532:   m_lYClick = y
  
534:   m_lButton = button
  
  Exit Sub
ErrHand:
538:   MsgBox "tvwMapBook_MouseDown - " & Err.Description
End Sub

Private Sub tvwMapBook_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo ErrHand:
  If Not m_bClickFlag Then Exit Sub
544:   m_bClickFlag = False
545:   m_bNodeFlag = False
  
  Exit Sub
ErrHand:
549:   MsgBox "tvwMapBook_MouseUp - " & Err.Description
End Sub

Private Sub tvwMapBook_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo ErrHand:
  Dim lLoop As Long, pUID As New UID, lImage As Long
  Dim pItem As ICommandItem, lPos As Long, sText As String
  Dim pMapBook As INWDSMapBook, pMapSeries As INWDSMapSeries
  Dim lPage As Long
  'Check to see if a MapSeries already exists
559:   Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
562:   Set pMapSeries = pMapBook.ContentItem(0)
  
564:   Set m_pCurrentNode = Node
  Select Case Node.Image
  Case 1, 2   'Enable and not Enabled options for a map book
567:     If m_lXClick < 180 Then
568:       If Node.Image = 1 Then
569:         Node.Image = 2
570:         pMapBook.EnableBook = False
'        tvwMapBook.Nodes.Item("MapSeries").Image = 4
'        UpdatePages False
573:       Else
574:         Node.Image = 1
575:         pMapBook.EnableBook = True
'        tvwMapBook.Nodes.Item("MapSeries").Image = 3
'        UpdatePages True
578:       End If
579:     Else
580:       If m_lButton = 2 Then
581:         PopupMenu mnuHeadingBook
582:       End If
583:     End If
  Case 3, 4   'Enable and not Enabled options for a map series
585:     If m_lXClick > 510 And m_lXClick < 760 Then
586:       If Node.Image = 3 Then
587:         Node.Image = 4
588:         pMapSeries.EnableSeries = False
'        UpdatePages False
590:       Else
591:         Node.Image = 3
592:         pMapSeries.EnableSeries = True
'        UpdatePages True
594:       End If
595:     Else
596:       If m_lButton = 2 Then
597:         If Node.Image = 3 Then
598:           mnuSeries(11).Caption = "Disable Series"
599:         Else
600:           mnuSeries(11).Caption = "Enable Series"
601:         End If
602:         PopupMenu mnuHeadingSeries
603:       End If
604:     End If
  Case 5, 6   'Enable and not Enabled options for a map page
606:     If m_lXClick > 1320 Then
607:       If m_lButton = 2 Then
608:         If Node.Image = 5 Then
609:           mnuPage(4).Caption = "Disable Page"
610:         Else
611:           mnuPage(4).Caption = "Enable Page"
612:         End If
613:         PopupMenu mnuHeadingPage
614:       End If
615:     ElseIf m_lXClick > 1080 And m_lXClick <= 1320 Then
616:       lPage = Node.Tag
617:       If Node.Image = 5 Then
618:         Node.Image = 6
619:         pMapSeries.Page(lPage).EnablePage = False
620:       Else
621:         Node.Image = 5
622:         pMapSeries.Page(lPage).EnablePage = True
623:       End If
624:     End If
625:   End Select

  Exit Sub
ErrHand:
629:   MsgBox "twvMapBook_NodeClick - " & Erl & ": " & Err.Description
End Sub

Private Sub UpdatePages(bEnableFlag As Boolean)
On Error GoTo ErrHand:
  Dim lLoop As Long, pNode As Node
635:   For lLoop = 2 To tvwMapBook.Nodes.count
636:     Set pNode = tvwMapBook.Nodes.Item(lLoop)
637:     If pNode.Image = 5 Or pNode.Image = 6 Then
638:       If bEnableFlag = True Then
639:         pNode.Image = 5
640:       Else
641:         pNode.Image = 6
642:       End If
643:     End If
644:   Next lLoop

  Exit Sub
ErrHand:
648:   MsgBox "UpdatePages - " & Err.Description
End Sub

Public Sub ShowPrinterDialog(pMxApp As IMxApplication, Optional pMapSeries As INWDSMapSeries, Optional pPrintMaterial As IUnknown)
  On Error GoTo ErrorHandler
  Dim pFrm As frmPrint

  Dim pPrinter As IPrinter
  Dim pApp As IApplication
  Dim iNumPages As Integer
  Dim pPage As IPage
  Dim pDoc As IMxDocument
  Dim pLayout As IPageLayout
  
662:   Set pPrinter = pMxApp.Printer
663:   If pPrinter Is Nothing Then
664:     MsgBox "You must have at least one printer defined before using this command!!!"
    Exit Sub
666:   End If
  
668:   Set pApp = pMxApp
669:   Set pFrm = New frmPrint
670:   pFrm.Application = pApp
671:   pFrm.ExportFrame = m_pExportFrame
672:   m_pExportFrame.Create pFrm
  
674:   Set pDoc = pApp.Document
675:   Set pLayout = pDoc.PageLayout
676:   Set pPage = pLayout.Page
          
678:   pPage.PrinterPageCount pPrinter, 0, iNumPages
  
680:   pFrm.txtTo.Text = iNumPages
      
682:   pFrm.lblName.Caption = pPrinter.Paper.PrinterName
683:   pFrm.lblType.Caption = pPrinter.DriverName
684:   If TypeOf pPrinter Is IPsPrinter Then
685:     pFrm.chkPrintToFile.Enabled = True
686:   Else
687:     pFrm.chkPrintToFile.Value = 0
688:     pFrm.chkPrintToFile.Enabled = False
689:   End If
  'If pprintmaterial is nothing then it means you are printing a map series
  
692:   If pPrintMaterial Is Nothing Then
693:       pFrm.aNWDSMapSeries = pMapSeries
694:       pFrm.optPrintCurrentPage.Enabled = False
695:       m_pExportFrame.Visible = True
      Exit Sub
697:   End If
  
699:   If TypeOf pPrintMaterial Is INWDSMapBook Then
700:       pFrm.aNWDSMapBook = pPrintMaterial
701:       pFrm.optPrintCurrentPage.Enabled = False
702:       pFrm.optPrintPages.Enabled = False
703:       pFrm.txtPrintPages.Enabled = False
704:   ElseIf TypeOf pPrintMaterial Is INWDSMapPage Then
705:       pFrm.aNWDSMapPage = pPrintMaterial
706:       pFrm.aNWDSMapSeries = pMapSeries
707:       pFrm.optPrintCurrentPage.Value = True
708:       pFrm.optPrintAll.Enabled = False
709:       pFrm.optPrintPages.Enabled = False
710:       pFrm.txtPrintPages.Enabled = False
711:   End If
712:   m_pExportFrame.Visible = True
713:   Set pPrintMaterial = Nothing
    
  Exit Sub
ErrorHandler:
717:   MsgBox "ShowPrinterDialog - " & Err.Description
End Sub

Public Sub ShowExporterDialog(pApp As IApplication, Optional pMapSeries As INWDSMapSeries, Optional pExportMaterial As IUnknown)
  On Error GoTo ErrorHandler
  Dim pFrm As frmExport
    
724:   Set pFrm = New frmExport
725:   pFrm.Application = pApp
726:   pFrm.ExportFrame = m_pExportFrame
727:   m_pExportFrame.Create pFrm
  
729:   If pExportMaterial Is Nothing Then
730:       pFrm.aNWDSMapSeries = pMapSeries
731:       pFrm.optCurrentPage.Enabled = False
732:       pFrm.InitializeTheForm
733:       m_pExportFrame.Visible = True
      Exit Sub
735:   End If
  
737:   If TypeOf pExportMaterial Is INWDSMapBook Then
738:     pFrm.aNWDSMapBook = pExportMaterial
739:     pFrm.optCurrentPage.Enabled = False
740:     pFrm.optPages.Enabled = False
741:     pFrm.txtPages.Enabled = False
742:     pFrm.InitializeTheForm
743:   ElseIf TypeOf pExportMaterial Is INWDSMapPage Then
744:     pFrm.aNWDSMapPage = pExportMaterial
745:     pFrm.aNWDSMapSeries = pMapSeries
746:     pFrm.optCurrentPage.Value = True
747:     pFrm.optAll.Enabled = False
748:     pFrm.optPages.Enabled = False
749:     pFrm.txtPages.Enabled = False
750:     pFrm.InitializeTheForm
751:   End If
752:   m_pExportFrame.Visible = True
753:   Set pExportMaterial = Nothing

  Exit Sub
ErrorHandler:
757:   MsgBox "ShowExporterDialog - " & Err.Description
End Sub



