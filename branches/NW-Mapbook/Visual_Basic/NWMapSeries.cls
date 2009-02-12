VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NWMapSeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

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

Implements INWDSMapSeries
Implements INWDSMapSeriesProps
Implements INWDSMapSeriesOptions
Implements INWDSMapSeriesOptions2  'Added 6/18/03 to support cross hatching of clip
Implements INWMapSeriesOptions  'Added 7/2/04 for NW Map Production project
Implements IPersistVariant

Private m_PageColl As Collection
'Private m_PageColl As Scripting.Dictionary
Private m_bEnableSeries As Boolean
Private m_bClipData As Boolean
Private m_lClipData2 As Long  'Added 6/18/03 to support cross hatching of clip
Private m_sDataDrivenField As String
Private m_lExtentType As Long
Private m_dFixedScale As Double
Private m_bLabelNeighbors As Boolean
Private m_pLabelSymbol As ISymbol
Private m_dMargin As Double
Private m_sMarginType As String
Private m_bRotateFrame As Boolean
Private m_sRotationField As String
Private m_sDataFrameName As String
Private m_sIndexFieldName As String
Private m_sIndexLayerName As String
Private m_SuppressColl As Collection
Private m_bSupressLayers As Boolean
Private m_lTileSelection As Long
Private m_lStartNumber As Long

''''properties added for NW Mapbook
Private m_dLabelIndent As Double
Private m_pUID_NWData1 As UID
Private m_pUID_FlexiblePersist As UID
Private m_pTextSymbolDict As Scripting.Dictionary
Private m_sDefaultTextSymbol As String

'layer (in)visibility groups
Private m_pLyrInvisGroups As Scripting.Dictionary

'bubble layer data
Private m_sBubbleLayer As String

'detail inset graphics
Private m_pDetailInsets As Scripting.Dictionary
Private m_pDetailInsetArrowShadows As Scripting.Dictionary
Private m_pDetailInsetArrows As Scripting.Dictionary
Private m_pMxDoc As IMxDocument  'needed for cleaning display on loading graphics

'data structure for data frame visibility management per map page
Private m_pDictDataFramePages As Scripting.Dictionary 'map pages where data frame is visible (not where is invisible)
Private m_pDictStoredDataFrames As Scripting.Dictionary
Private m_sMainDataFrame As String

'dynamic definition queries
Private m_sPrevPageName As String
Private m_pDictDynDefQueryPages As Scripting.Dictionary
Private m_sCurrentPageName As String

'layout element visibility
Private m_pDictPagesWhereElementIsInvisible As Scripting.Dictionary
Private m_pDictTaggedElements As Scripting.Dictionary
Private m_pDictStoredElements As Scripting.Dictionary

'reload map page on map refresh events
Private m_bHandleMapRefreshEvents As Boolean
Private m_bShouldntRefresh As Boolean

'Data frames (other than main DF) to update when drawing map pages
Private m_pDictDataFramesToUpdate_Layer As Scripting.Dictionary 'tracks data frames, and the map page layer
Private m_pDictDataFramesToUpdate_Field As Scripting.Dictionary 'tracks data frames, and the map name attr field of map page layer
Private m_lDataFrameToUpdate_ExtentOption As Long
'OTHER_DF_EXTENTOPTION_POLYEXTENT = 1  'option of using polygon extent to set extent of data frame
'OTHER_DF_EXTENTOPTION_MAINDFSCALE = 2 'option of using the scale of the main data frame for the other data frames
''''end of NW Mapbook properties


Const c_sModuleFileName As String = "NWMapSeries.cls"





Private Sub Class_Initialize()
114:   Set m_PageColl = New Collection
  'Set m_PageColl = New Scripting.Dictionary
116:   Set m_SuppressColl = New Collection
                                                  'NW Mapbook additional code
118:   If m_pUID_FlexiblePersist Is Nothing Then
119:     Set m_pUID_FlexiblePersist = New UID
120:     m_pUID_FlexiblePersist.Value = "{92091FB8-FC34-40fe-ADBD-C00C2E39022A}"
121:   End If
122:   If m_pUID_NWData1 Is Nothing Then
123:     Set m_pUID_NWData1 = New UID
124:     m_pUID_NWData1.Value = "{8D818E86-5C50-4c1a-BBC3-6EF7A5CD5275}"
125:   End If
126:   Set m_pTextSymbolDict = New Scripting.Dictionary
127:   Set m_pLyrInvisGroups = New Scripting.Dictionary
128:   Set m_pDetailInsets = New Scripting.Dictionary
129:   Set m_pDetailInsetArrowShadows = New Scripting.Dictionary
130:   Set m_pDetailInsetArrows = New Scripting.Dictionary
131:   Set m_pDictDataFramePages = New Scripting.Dictionary
132:   Set m_pDictStoredDataFrames = New Scripting.Dictionary
133:   Set m_pDictDynDefQueryPages = New Scripting.Dictionary
134:   Set m_pDictPagesWhereElementIsInvisible = New Scripting.Dictionary
135:   Set m_pDictTaggedElements = New Scripting.Dictionary
136:   Set m_pDictStoredElements = New Scripting.Dictionary
137:   Set m_pDictDataFramesToUpdate_Layer = New Scripting.Dictionary
138:   Set m_pDictDataFramesToUpdate_Field = New Scripting.Dictionary
  
140:   GBL_BUBBLE_IDENTIFIER = "BubbleID:"
141:   GBL_ELEMENT_IDENTIFIER = "Elem"
142:   GBL_ELEMENT_COUNT = 0
  
144:   m_bHandleMapRefreshEvents = False
145:   m_bShouldntRefresh = False
  
147:   OTHER_DF_EXTENTOPTION_POLYEXTENT = 1  'option of using polygon extent to set extent of data frame
148:   OTHER_DF_EXTENTOPTION_MAINDFSCALE = 2 'option of using the scale of the main data frame for the other data frames
149:   m_lDataFrameToUpdate_ExtentOption = OTHER_DF_EXTENTOPTION_POLYEXTENT

End Sub

Private Sub Class_Terminate()
154:   Set m_PageColl = Nothing
155:   Set m_SuppressColl = Nothing
156:   Set m_pTextSymbolDict = Nothing
157:   Set m_pLyrInvisGroups = Nothing
158:   Set m_pDetailInsets = Nothing
159:   Set m_pDetailInsetArrowShadows = Nothing
160:   Set m_pDetailInsetArrows = Nothing
161:   Set m_pDictDataFramePages = Nothing
162:   Set m_pDictStoredDataFrames = Nothing
163:   Set m_pDictDynDefQueryPages = Nothing
164:   Set m_pDictPagesWhereElementIsInvisible = Nothing
165:   Set m_pDictTaggedElements = Nothing
166:   Set m_pDictStoredElements = Nothing
167:   Set m_pDictDataFramesToUpdate_Layer = Nothing
168:   Set m_pDictDataFramesToUpdate_Field = Nothing
169:   GBL_ELEMENT_COUNT = 0
End Sub




Private Sub INWMapSeriesOptions_BubbleGraphicAdd( _
  pInsetElement As esriCarto.IElement, _
  pArrowElement As esriCarto.IElement, _
  pArrowShadowElement As esriCarto.IElement, _
  sInsetName As String)
  
  On Error GoTo ErrorHandler


184:   If sInsetName = "" Then
    Exit Sub
186:   End If
  If pInsetElement Is Nothing Then Exit Sub
  If pArrowElement Is Nothing Then Exit Sub
  If pArrowShadowElement Is Nothing Then Exit Sub
  
191:   If m_pDetailInsets Is Nothing Then
192:     Set m_pDetailInsets = New Scripting.Dictionary
193:   End If
194:   If m_pDetailInsetArrows Is Nothing Then
195:     Set m_pDetailInsetArrows = New Scripting.Dictionary
196:   End If
197:   If m_pDetailInsetArrowShadows Is Nothing Then
198:     Set m_pDetailInsetArrowShadows = New Scripting.Dictionary
199:   End If
  ''''''''''debug
'  If m_pDetailInsets.Exists(sInsetName) Then MsgBox "In NWMapSeries, INWMapSeriesOptions_BubbleGraphicAdd" & vbNewLine _
        & "About to insert duplicate element for inset " & sInsetName & " into m_pDetailInsets."
203:   m_pDetailInsets.Add sInsetName, pInsetElement
204:   m_pDetailInsetArrows.Add sInsetName, pArrowElement
205:   m_pDetailInsetArrowShadows.Add sInsetName, pArrowShadowElement
  

  Exit Sub
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_BubbleGraphicAdd " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub INWMapSeriesOptions_BubbleGraphicRemove(pMxDoc As esriArcMapUI.IMxDocument, sInsetName As String)
  On Error GoTo ErrorHandler

  Dim pGraphicsContainer As IGraphicsContainer, pElement As IElement, pAV As IActiveView
  Dim pPageLayout As IPageLayout, pMapFrame As IMapFrame
  
  
  If pMxDoc Is Nothing Then Exit Sub
  If sInsetName = "" Then Exit Sub
    
223:   If m_pDetailInsets Is Nothing Then
224:     Set m_pDetailInsets = New Scripting.Dictionary
225:   End If
226:   If m_pDetailInsetArrows Is Nothing Then
227:     Set m_pDetailInsetArrows = New Scripting.Dictionary
228:   End If
229:   If m_pDetailInsetArrowShadows Is Nothing Then
230:     Set m_pDetailInsetArrowShadows = New Scripting.Dictionary
231:   End If

  If Not m_pDetailInsets.Exists(sInsetName) Then Exit Sub
                                              'delete the three elements in storage
235:   Set pPageLayout = pMxDoc.PageLayout
236:   Set pGraphicsContainer = pPageLayout
  
  ''''''''manual graphics container search approach
  'remove the detail inset
240:   pGraphicsContainer.Reset
241:   Set pElement = pGraphicsContainer.Next
242:   Do While Not pElement Is Nothing
243:     If TypeOf pElement Is IMapFrame Then
244:       Set pMapFrame = pElement
245:       If StrComp(sInsetName, pMapFrame.Map.Name) = 0 Then
246:         pGraphicsContainer.DeleteElement pElement
247:         Set pElement = Nothing
248:       End If
249:     End If
250:     If Not pElement Is Nothing Then
251:       Set pElement = pGraphicsContainer.Next
252:     End If
253:   Loop
  
  '''''''''or, just assume you can call deleteelement on some separately stored
  '''''''''element. Apparently works for arrow and shadow arrow graphics, but not
  '''''''''for data frames
  'Set pElement = m_pDetailInsets.Item(sInsetName)
  'pGraphicsContainer.DeleteElement pElement
260:   Set pElement = m_pDetailInsetArrows.Item(sInsetName)
261:   pGraphicsContainer.DeleteElement pElement
262:   Set pElement = m_pDetailInsetArrowShadows.Item(sInsetName)
263:   pGraphicsContainer.DeleteElement pElement
  
  
266:   m_pDetailInsets.Remove sInsetName
267:   m_pDetailInsetArrows.Remove sInsetName
268:   m_pDetailInsetArrowShadows.Remove sInsetName


  Exit Sub
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_BubbleGraphicRemove " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub


Private Property Let INWMapSeriesOptions_BubbleLayer(RHS As String)
278:   m_sBubbleLayer = RHS
End Property

Private Property Get INWMapSeriesOptions_BubbleLayer() As String
282:   INWMapSeriesOptions_BubbleLayer = m_sBubbleLayer
End Property





'Given the maze of data structures associated with data frames and map pages,
'this function localizes in one place the code necessary to make a decision
'about whether or not a data frame should be displayed on a given map page.
Public Function DataFrameShouldBeDisplayed( _
     sDataFrame As String, _
     sMapPage As String, _
     pDictDataFrameList As Scripting.Dictionary, _
     sMainDataFrameName As String) As Boolean
  On Error GoTo ErrorHandler


  Dim bResult As Boolean, pScrDict As Scripting.Dictionary
301:   bResult = False
                                                  'should display main map frame
303:   If StrComp(sMainDataFrameName, sDataFrame) = 0 Then
304:     bResult = True
305:   Else
                                                  'should display bubble detail insets
307:     If StrComp(Left$(sDataFrame, Len(GBL_BUBBLE_IDENTIFIER)), _
               GBL_BUBBLE_IDENTIFIER, _
               vbTextCompare) = 0 Then
310:       bResult = True
311:     Else
                                                  'should display data frames if the structure
                                                  'that tracks what data frames are visible for
                                                  'which map pages has an entry for this data frame
                                                  'and this map page
316:       If pDictDataFrameList.Exists(sDataFrame) Then
317:         Set pScrDict = pDictDataFrameList.Item(sDataFrame)
318:         If pScrDict.Exists(sMapPage) Then
319:           bResult = True
320:         End If
321:       End If
322:     End If
323:   End If
  
325:   DataFrameShouldBeDisplayed = bResult
  

  Exit Function
ErrorHandler:
  HandleError True, "DataFrameShouldBeDisplayed " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function




Private Property Get INWMapSeriesOptions_CurrentPage() As INWDSMapPage
  On Error GoTo ErrorHandler

339:   If Not ItemExists(m_PageColl, m_sCurrentPageName) Then
  'If Not m_PageColl.Exists(m_sCurrentPageName) Then  'scripting.dictionary method
341:     INWMapSeriesOptions_CurrentPage = Nothing
342:   Else
343:     Set INWMapSeriesOptions_CurrentPage = m_PageColl.Item(m_sCurrentPageName)
344:   End If

  Exit Property
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_CurrentPage " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Public Property Get INWMapSeriesOptions_CurrentPageName() As String
352:   INWMapSeriesOptions_CurrentPageName = m_sCurrentPageName
End Property

'When a new map page is loaded, this routine can be called to
'make visible those data frames that should be visible, and to
'set invisible those that should not be visible.
'---------------------------------------------------
Private Sub INWMapSeriesOptions_DataFrameDrawFrames( _
  pMxDoc As esriArcMapUI.IMxDocument, _
  sMapPageID As String, sMainDataFrameName As String)
  
  On Error GoTo ErrorHandler

  Dim pPageLayout As IPageLayout, i As Long, pGraphicsContainer As IGraphicsContainer
  Dim pElement As IElement, lFinishedCount As Long, lMaxRemainingCount As Long
  Dim pScrDict As Scripting.Dictionary, pMapFrame As IMapFrame, sDataFrameName As String
  Dim vNameKeys As Variant, pMainMap As IMap, pActiveView As IActiveView
  Dim pScrAlreadySparedFrames As Scripting.Dictionary
  Dim pViewManager As IViewManager
  
372:   If m_pDictDataFramePages Is Nothing Then
    Exit Sub
374:   End If
375:   If m_pDictStoredDataFrames Is Nothing Then
376:     Set m_pDictStoredDataFrames = New Scripting.Dictionary
377:   End If
                                                  'activate the main data frame used
                                                  'by map series
380:   Set pActiveView = pMxDoc.ActiveView
381:   For i = 0 To pMxDoc.Maps.Count - 1
382:     If StrComp(pMxDoc.Maps.Item(i).Name, m_sMainDataFrame, vbTextCompare) = 0 Then
383:       Set pMainMap = pMxDoc.Maps.Item(i)
384:     End If
385:   Next i
386:   If TypeOf pActiveView Is IPageLayout Then
387:     Set pMxDoc.ActiveView.FocusMap = pMainMap
388:   Else
389:     Set pMxDoc.ActiveView = pMainMap
390:   End If
391:   Set pViewManager = pMxDoc.FocusMap
392:   Set pPageLayout = pMxDoc.PageLayout 'caused an error if the selected data frame wasn't the main data frame
393:   Set pGraphicsContainer = pPageLayout
394:   pGraphicsContainer.Reset
  
                                                  'data frames are in one of two locations:
                                                  'the map, and m_pDictStoredDataFrames.  This
                                                  'is where data frames go when they're not on
                                                  'the page layout.
                                                  'Bring out from storage those data frames
                                                  'to be added to this current map page.
402:   vNameKeys = m_pDictStoredDataFrames.Keys
403:   For i = 0 To (m_pDictStoredDataFrames.Count - 1)
404:     sDataFrameName = vNameKeys(i)
405:     If DataFrameShouldBeDisplayed(sDataFrameName, sMapPageID, m_pDictDataFramePages, sMainDataFrameName) Then
406:       Set pElement = m_pDictStoredDataFrames.Item(sDataFrameName)
407:       pGraphicsContainer.AddElement pElement, 0
408:       m_pDictStoredDataFrames.Remove sDataFrameName
    'else
      'data frame should be left alone, it's already the way it should be
    
    
413:     End If
414:   Next i
                                                  'clear selection of added dataframes
  'pViewManager.ElementSelection.Clear
  
418:   pGraphicsContainer.Reset
419:   Set pElement = pGraphicsContainer.Next
420:   Set pScrAlreadySparedFrames = New Scripting.Dictionary
421:   lFinishedCount = 0   'limit the loop in case of a multitude of non-imapframe elements
422:   lMaxRemainingCount = m_pDictDataFramePages.Count
  
424:   Do While ((Not pElement Is Nothing) And (lFinishedCount < lMaxRemainingCount))
425:     If TypeOf pElement Is IMapFrame Then
426:       Set pMapFrame = pElement
427:       sDataFrameName = pMapFrame.Map.Name
428:       If Not pScrAlreadySparedFrames.Exists(sDataFrameName) Then
429:         If DataFrameShouldBeDisplayed(sDataFrameName, sMapPageID, m_pDictDataFramePages, sMainDataFrameName) Then
          'pElement is already in map, so leave alone.
          'Possibility of the data frame having been invisible
          'was taken care of by previous loop
433:           pScrAlreadySparedFrames.Add sDataFrameName, Nothing
434:           lFinishedCount = lFinishedCount + 1
435:         Else
436:           m_pDictStoredDataFrames.Add sDataFrameName, pElement
437:           pGraphicsContainer.DeleteElement pElement
438:           pGraphicsContainer.Reset
439:           lFinishedCount = lFinishedCount + 1
440:         End If
441:       End If
    'else
      'not a mapframe -> not a data frame, not tracking
444:     End If
445:     Set pElement = pGraphicsContainer.Next
446:   Loop

448:   Set pScrAlreadySparedFrames = Nothing
  'pViewManager.ElementSelection.Clear

  Exit Sub
ErrorHandler:
453:   Set pScrAlreadySparedFrames = Nothing
  HandleError True, "INWMapSeriesOptions_DataFrameDrawFrames " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub



'DataFrameIsInStorage
'
'This routine is useful when dealing with a mixed list of data frames in
'storage and in the page layout.  This can let that code test if a list item
'is in storage, and access the data frame differently as a result.
'-------------------------------------------------------------
Private Property Get INWMapSeriesOptions_DataFrameIsInStorage(sDataFrameID As String) As Boolean
466:   If Not m_pDictStoredDataFrames Is Nothing Then
467:     INWMapSeriesOptions_DataFrameIsInStorage = m_pDictStoredDataFrames.Exists(sDataFrameID)
468:   Else
469:     INWMapSeriesOptions_DataFrameIsInStorage = False
470:   End If
End Property



Private Function INWMapSeriesOptions_DataFrameIsVisibleInPage( _
  sDataFrameID As String, _
  sMapPageID As String) As Boolean
  On Error GoTo ErrorHandler


  Dim pScrDict As Scripting.Dictionary
  
483:   If Not m_pDictDataFramePages.Exists(sDataFrameID) Then
484:     INWMapSeriesOptions_DataFrameIsVisibleInPage = False
    Exit Function
486:   End If
487:   Set pScrDict = m_pDictDataFramePages.Item(sDataFrameID)
488:   INWMapSeriesOptions_DataFrameIsVisibleInPage = pScrDict.Exists(sMapPageID)

  Exit Function
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_DataFrameIsVisibleInPage " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Sub INWMapSeriesOptions_DataFrameLoadFrames(pMxDoc As esriArcMapUI.IMxDocument)
  On Error GoTo ErrorHandler

  Dim i As Long, pPageLayout As IPageLayout, pGraphicsContainer As IGraphicsContainer
  Dim pElement As IElement, pMapFrame As IMapFrame, sDataFrameName As String
  Dim bHasBeenWarned As Boolean, pScrDict As Scripting.Dictionary
  
502:   If m_pDictDataFramePages Is Nothing Then
503:     Set m_pDictDataFramePages = New Scripting.Dictionary
504:   End If
  'loop through available data frames.
  
507:   Set pPageLayout = pMxDoc.PageLayout
508:   Set pGraphicsContainer = pPageLayout
509:   pGraphicsContainer.Reset
510:   Set pElement = pGraphicsContainer.Next
511:   Do While Not pElement Is Nothing
512:     If TypeOf pElement Is IMapFrame Then
513:       Set pMapFrame = pElement
514:       sDataFrameName = pMapFrame.Map.Name
                                                  'track visiblity per map page for
                                                  'any that don't happen to be detail insets
517:       If StrComp(Left$(sDataFrameName, Len(GBL_BUBBLE_IDENTIFIER)), GBL_BUBBLE_IDENTIFIER, vbTextCompare) <> 0 Then
518:         If (Not bHasBeenWarned) And m_pDictDataFramePages.Exists(sDataFrameName) Then
519:           MsgBox "Warning: More than one data frame has the name " & vbNewLine _
               & vbTab & sDataFrameName & vbNewLine _
               & "Map series tracking of data frame visibility will only work for" & vbNewLine _
               & "the first data frame with any given name.", vbOKOnly
523:           bHasBeenWarned = True
524:         Else
525:           Set pScrDict = New Scripting.Dictionary
526:           m_pDictDataFramePages.Add sDataFrameName, pScrDict
527:         End If
528:       End If
529:     End If
530:     Set pElement = pGraphicsContainer.Next
531:   Loop

  Exit Sub
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_DataFrameLoadFrames " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub


Private Property Let INWMapSeriesOptions_DataFrameMainFrame(RHS As String)
540:   m_sMainDataFrame = RHS
End Property

Private Property Get INWMapSeriesOptions_DataFrameMainFrame() As String
544:   INWMapSeriesOptions_DataFrameMainFrame = m_sMainDataFrame
End Property

Private Property Get INWMapSeriesOptions_DataFramePagesWhereVisible(sDataFrameID As String) As Variant
  On Error GoTo ErrorHandler

  Dim pScrDict As Scripting.Dictionary, vResult As Variant, sResult() As String
  
552:   If m_pDictDataFramePages Is Nothing Then
553:     Set m_pDictDataFramePages = New Scripting.Dictionary
554:   End If
555:   If Not m_pDictDataFramePages.Exists(sDataFrameID) Then
556:     INWMapSeriesOptions_DataFramePagesWhereVisible = sResult
    Exit Property
558:   End If
559:   Set pScrDict = m_pDictDataFramePages.Item(sDataFrameID)
560:   INWMapSeriesOptions_DataFramePagesWhereVisible = pScrDict.Keys

  Exit Property
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_DataFramePagesWhereVisible " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Property Get INWMapSeriesOptions_DataFramePagesWhereVisibleCount(sDataFrameID As String) As Long
  On Error GoTo ErrorHandler

  Dim pScrDict As Scripting.Dictionary
  
572:   If m_pDictDataFramePages Is Nothing Then
573:     Set m_pDictDataFramePages = New Scripting.Dictionary
574:   End If
575:   If Not m_pDictDataFramePages.Exists(sDataFrameID) Then
576:     INWMapSeriesOptions_DataFramePagesWhereVisibleCount = -1
    Exit Property
578:   End If
579:   Set pScrDict = m_pDictDataFramePages.Item(sDataFrameID)
580:   INWMapSeriesOptions_DataFramePagesWhereVisibleCount = pScrDict.Count
  
  Exit Property
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_DataFramePagesWhereVisibleCount " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Sub INWMapSeriesOptions_DataFrameRemoveFrame(sDataFrameID As String)
588:   If m_pDictDataFramePages.Exists(sDataFrameID) Then
589:     m_pDictDataFramePages.Remove (sDataFrameID)
590:   End If
End Sub

Private Sub INWMapSeriesOptions_DataFrameRemovePage(sDataFrameID As String, sMapPageID As String)
  On Error GoTo ErrorHandler

  Dim pScrDict As Scripting.Dictionary
  
598:   If m_pDictDataFramePages Is Nothing Then
599:     Set m_pDictDataFramePages = New Scripting.Dictionary
600:   End If
601:   If Not m_pDictDataFramePages.Exists(sDataFrameID) Then
    Exit Sub
603:   End If
604:   Set pScrDict = m_pDictDataFramePages.Item(sDataFrameID)
  If pScrDict Is Nothing Then Exit Sub
606:   If pScrDict.Exists(sMapPageID) Then
607:     pScrDict.Remove sMapPageID
608:   End If

  Exit Sub
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_DataFrameRemovePage " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub INWMapSeriesOptions_DataFrameSetInvisibleInPage(sDataFrameID As String, sMapPageID As String)
  On Error GoTo ErrorHandler

  Dim pScrDict As Scripting.Dictionary
                                                  'remove the map page from the collection
                                                  'for the indicated data frame.
621:   If m_pDictDataFramePages Is Nothing Then
622:     Set m_pDictDataFramePages = New Scripting.Dictionary
623:   End If
624:   If Not m_pDictDataFramePages.Exists(sDataFrameID) Then
    Exit Sub
626:   End If
627:   Set pScrDict = m_pDictDataFramePages.Item(sDataFrameID)
628:   If pScrDict Is Nothing Then
    Exit Sub
630:   End If
631:   If pScrDict.Exists(sMapPageID) Then
632:     pScrDict.Remove sMapPageID
633:   End If

  Exit Sub
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_DataFrameSetInvisibleInPage " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

'Add a reference to the map page in the collection for the
'indicated data frame.
Private Sub INWMapSeriesOptions_DataFrameSetVisibleInPage(sDataFrameID As String, sMapPageID As String)
  On Error GoTo ErrorHandler

  Dim pScrDict As Scripting.Dictionary
646:   If m_pDictDataFramePages Is Nothing Then
647:     Set m_pDictDataFramePages = New Scripting.Dictionary
648:   End If
649:   If Not m_pDictDataFramePages.Exists(sDataFrameID) Then
650:     Set pScrDict = New Scripting.Dictionary
651:     m_pDictDataFramePages.Add sDataFrameID, pScrDict
    'hopefully the calling code with also set the other
    'map pages, otherwise this data frame will be visible
    '*only* in this map page.
    'This function was designed to be called from the
    'form that manages the assignment of map page visibility.
657:   Else
658:     Set pScrDict = m_pDictDataFramePages.Item(sDataFrameID)
659:   End If
660:   If Not pScrDict.Exists(sMapPageID) Then
661:     pScrDict.Add sMapPageID, Nothing
662:   End If

  Exit Sub
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_DataFrameSetVisibleInPage " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub








Private Property Get INWMapSeriesOptions_DataFramesManaged() As Variant
677:   INWMapSeriesOptions_DataFramesManaged = m_pDictDataFramePages.Keys
End Property

Private Property Get INWMapSeriesOptions_DataFramesStored() As Variant
  Dim vEmpty As Variant
682:   If Not m_pDictStoredDataFrames Is Nothing Then
683:     INWMapSeriesOptions_DataFramesStored = m_pDictStoredDataFrames.Keys
684:   Else
685:     INWMapSeriesOptions_DataFramesStored = vEmpty
686:   End If
End Property

Private Function INWMapSeriesOptions_DataFrameStoredFrame(sDataFrameID As String) As esriCarto.IMapFrame
  On Error GoTo ErrorHandler

  Dim pResult As IElement, pMapFrame As IMapFrame
693:   If m_pDictStoredDataFrames.Exists(sDataFrameID) Then
694:     Set pResult = m_pDictStoredDataFrames.Item(sDataFrameID)
695:   Else
696:     Set pResult = Nothing
697:   End If
698:   Set pMapFrame = pResult
699:   Set INWMapSeriesOptions_DataFrameStoredFrame = pMapFrame

  Exit Function
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_DataFrameStoredFrame " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'Given a named data frame that happens to be in storage, and given the
'name of a layer from that data frame, return a reference to that layer
Private Function INWMapSeriesOptions_DataFrameStoredFrameLayer(sDataFrameID As String, sLayerName As String) As esriCarto.ILayer
  On Error GoTo ErrorHandler

  Dim pLayer As ILayer, pElement As IElement, pMapFrame As IMapFrame
  Dim pMap As IMap, pLayers As IEnumLayer, bLayerIsFound As Boolean
  
714:   If m_pDictStoredDataFrames.Exists(sDataFrameID) Then
715:     Set pElement = m_pDictStoredDataFrames.Item(sDataFrameID)
716:     If pElement Is Nothing Then
717:       Set pLayer = Nothing
718:     Else
719:       If Not TypeOf pElement Is IMapFrame Then
720:         Set pLayer = Nothing
721:       Else
722:         Set pMapFrame = pElement
723:         Set pMap = pMapFrame.Map
724:         Set pLayers = pMap.Layers
725:         pLayers.Reset
726:         Set pLayer = pLayers.Next
727:         bLayerIsFound = False
728:         Do While (Not pLayer Is Nothing) And Not bLayerIsFound
729:           If StrComp(pLayer.Name, sLayerName, vbTextCompare) = 0 Then
730:             bLayerIsFound = True
731:           Else
732:             Set pLayer = pLayers.Next
733:           End If
734:         Loop
735:       End If
736:     End If
737:   Else
738:     Set pLayer = Nothing
739:   End If

741:   Set INWMapSeriesOptions_DataFrameStoredFrameLayer = pLayer

  Exit Function
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_DataFrameStoredFrameLayer " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Function INWMapSeriesOptions_DataFrameStoredFrameLayerList(sDataFrameID As String) As Variant
  On Error GoTo ErrorHandler

  Dim pMapFrame As IMapFrame, pElement As IElement, pMap As IMap
  Dim pLayers As IEnumLayer, pLayer As ILayer, sLayerList() As String
  Dim lLayerCount As Long, i As Long
  
  'grab a stored data frame
756:   If m_pDictStoredDataFrames Is Nothing Then
757:     INWMapSeriesOptions_DataFrameStoredFrameLayerList = sLayerList
    Exit Function
759:   End If
760:   If Not m_pDictStoredDataFrames.Exists(sDataFrameID) Then
761:     INWMapSeriesOptions_DataFrameStoredFrameLayerList = sLayerList
    Exit Function
763:   End If
764:   Set pElement = m_pDictStoredDataFrames.Item(sDataFrameID)
765:   If (pElement Is Nothing) Or (Not TypeOf pElement Is IMapFrame) Then
766:     INWMapSeriesOptions_DataFrameStoredFrameLayerList = sLayerList
    Exit Function
768:   End If
  
  'return that data frame's list of map layers
771:   Set pMapFrame = pElement
772:   Set pMap = pMapFrame.Map
773:   lLayerCount = pMap.LayerCount
  ReDim sLayerList(lLayerCount) As String
775:   Set pLayers = pMap.Layers
776:   Set pLayer = pLayers.Next
777:   i = 0
778:   Do While Not pLayer Is Nothing
779:     sLayerList(i) = pLayer.Name
780:     i = i + 1
781:     Set pLayer = pLayers.Next
782:   Loop
  
784:   INWMapSeriesOptions_DataFrameStoredFrameLayerList = sLayerList
  

  Exit Function
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_DataFrameStoredFrameLayerList " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function


Private Sub ConfirmDataStructures_DataFrameToUpdate()
794:   If m_pDictDataFramesToUpdate_Layer Is Nothing Then
795:     Set m_pDictDataFramesToUpdate_Layer = New Scripting.Dictionary
796:   End If
797:   If m_pDictDataFramesToUpdate_Field Is Nothing Then
798:     Set m_pDictDataFramesToUpdate_Field = New Scripting.Dictionary
799:   End If
End Sub

Private Sub INWMapSeriesOptions_DataFrameToUpdateAdd(sDataFrame As String, sMapPagesLayer As String, sPageNameField As String)
  On Error GoTo ErrorHandler

805:   ConfirmDataStructures_DataFrameToUpdate
  If sDataFrame = "" Then Exit Sub
  If sMapPagesLayer = "" Then Exit Sub
  If sPageNameField = "" Then Exit Sub
  
810:   If Not m_pDictDataFramesToUpdate_Layer.Exists(sDataFrame) Then
811:     m_pDictDataFramesToUpdate_Layer.Add sDataFrame, sMapPagesLayer
812:     m_pDictDataFramesToUpdate_Field.Add sDataFrame, sPageNameField
813:   End If


  Exit Sub
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_DataFrameToUpdate_Add " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub INWMapSeriesOptions_DataFrameToUpdateClearAllDataFrames()
  On Error GoTo ErrorHandler
  
824:   ConfirmDataStructures_DataFrameToUpdate
825:   m_pDictDataFramesToUpdate_Layer.RemoveAll
826:   m_pDictDataFramesToUpdate_Field.RemoveAll

  Exit Sub
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_DataFrameToUpdateClearAllDataFrames " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub INWMapSeriesOptions_DataFrameToUpdateClearDataFrame(sDataFrameName As String)
  On Error GoTo ErrorHandler


837:   ConfirmDataStructures_DataFrameToUpdate
838:   If m_pDictDataFramesToUpdate_Layer.Exists(sDataFrameName) Then
839:     m_pDictDataFramesToUpdate_Layer.Remove sDataFrameName
840:   End If
841:   If m_pDictDataFramesToUpdate_Field.Exists(sDataFrameName) Then
842:     m_pDictDataFramesToUpdate_Field.Remove sDataFrameName
843:   End If

  Exit Sub
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_DataFrameToUpdateClearDataFrame " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Property Get INWMapSeriesOptions_DataFrameToUpdateExtentOption() As String
  On Error GoTo ErrorHandler

  Dim sValue As String
  
  Select Case m_lDataFrameToUpdate_ExtentOption
    Case OTHER_DF_EXTENTOPTION_POLYEXTENT
857:       sValue = "Use polygon extent (default)"
    Case OTHER_DF_EXTENTOPTION_MAINDFSCALE
859:       sValue = "Use main dataframe's scale"
    Case Else
861:       Err.Raise vbObjectError + 1, "INWMapSeriesOptions_DataFrameToUpdateExtentOption (get)", _
        "Extent option not returned by routine."
863:   End Select
  
865:   INWMapSeriesOptions_DataFrameToUpdateExtentOption = sValue
  

  Exit Property
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_DataFrameToUpdateExtentOption " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Property Let INWMapSeriesOptions_DataFrameToUpdateExtentOption(RHS As String)
  On Error GoTo ErrorHandler

876:   m_lDataFrameToUpdate_ExtentOption = OTHER_DF_EXTENTOPTION_POLYEXTENT
877:   If StrComp(RHS, "Use polygon extent (default)", vbTextCompare) = 0 Then
878:     m_lDataFrameToUpdate_ExtentOption = OTHER_DF_EXTENTOPTION_POLYEXTENT
879:   ElseIf StrComp(RHS, "Use main dataframe's scale", vbTextCompare) = 0 Then
880:     m_lDataFrameToUpdate_ExtentOption = OTHER_DF_EXTENTOPTION_MAINDFSCALE
881:   End If
  Exit Property
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_DataFrameToUpdateExtentOption " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Function INWMapSeriesOptions_DataFrameToUpdateGetMapPageLayer(sDataFrameName As String) As String
  On Error GoTo ErrorHandler

  Dim sResult As String
891:   sResult = ""
  
893:   ConfirmDataStructures_DataFrameToUpdate
894:   If sDataFrameName <> "" Then
895:     sResult = m_pDictDataFramesToUpdate_Layer(sDataFrameName)
896:   End If
897:   INWMapSeriesOptions_DataFrameToUpdateGetMapPageLayer = sResult
  
  Exit Function
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_DataFrameToUpdateGetMapPageLayer " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Function INWMapSeriesOptions_DataFrameToUpdateGetPageNameField(sDataFrameName As String, sMapPageLayer As String) As String
  On Error GoTo ErrorHandler

  Dim sResult As String
908:   sResult = ""
  
910:   ConfirmDataStructures_DataFrameToUpdate
911:   If sDataFrameName <> "" Then
912:     sResult = m_pDictDataFramesToUpdate_Field(sDataFrameName)
913:   End If
914:   INWMapSeriesOptions_DataFrameToUpdateGetPageNameField = sResult
  
  Exit Function
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_DataFrameToUpdateGetPageNameField " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'DataFrameToUpdateHousekeepAbandonedFrames
'   This routine prevents memory leaks from orphaned data frame settings.
'   Settings can be orphaned when a data frame that was being tracked is
'   renamed or deleted.
'   This routine goes through all data frames listed in the data structures
'   that track data frames to update, and removes any entries that are not
'   either in the layout view, or are in the data frame storage data structure.
'   The dataframe storage data structure is for
'-------------------------------------------------------------------------
Private Sub INWMapSeriesOptions_DataFrameToUpdateHousekeepAbandonedFrames(pMxDoc As esriArcMapUI.IMxDocument)
  On Error GoTo ErrorHandler

  
  Dim pPageLayout As IPageLayout, pGraCont As IGraphicsContainer
  Dim vStoredDFList As Variant, lStoredDFCount As Long, i As Long
  Dim vTrackedDFList As Variant, lTrackedDFCount As Long
  Dim sStoredDF As String, sTrackedDF As String
  Dim pElement As IElement, pMapFrame As IMapFrame
  Dim bElemFound As Boolean
  
  
  If pMxDoc Is Nothing Then Exit Sub
943:   Set pGraCont = pMxDoc.PageLayout
  
  'sort through list of tracked data frames
946:   vTrackedDFList = m_pDictDataFramesToUpdate_Layer.Keys
947:   lTrackedDFCount = UBound(vTrackedDFList) + 1

949:   For i = 0 To (lTrackedDFCount - 1)

    'if the tracked data frame isn't in the layout
952:     sTrackedDF = vTrackedDFList(i)
953:     pGraCont.Reset
954:     Set pElement = pGraCont.Next
955:     bElemFound = False
956:     Do While (Not pElement Is Nothing) And (Not bElemFound)
957:       If TypeOf pElement Is IMapFrame Then
958:         Set pMapFrame = pElement
959:         If (StrComp(pMapFrame.Map.Name, sTrackedDF, vbTextCompare) = 0) Then
960:           bElemFound = True
961:         End If
962:       End If
963:       Set pElement = pGraCont.Next
964:     Loop
    
966:     If Not bElemFound Then
      'nor is it in the list of stored data frames,
968:       If Not m_pDictStoredDataFrames.Exists(sTrackedDF) Then
        'then that data frame no longer exists to
        'be tracked.
971:         m_pDictDataFramesToUpdate_Layer.Remove sTrackedDF
972:         m_pDictDataFramesToUpdate_Field.Remove sTrackedDF
973:       End If
974:     End If
975:   Next i
  

  Exit Sub
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_DataFrameToUpdateHousekeepAbandonedFrames " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Function INWMapSeriesOptions_DataFrameToUpdateIsADataFrameToUpdate(sDataFrameName As String) As Boolean
  On Error GoTo ErrorHandler

986:   ConfirmDataStructures_DataFrameToUpdate
987:   INWMapSeriesOptions_DataFrameToUpdateIsADataFrameToUpdate _
   = m_pDictDataFramesToUpdate_Layer.Exists(sDataFrameName)

  Exit Function
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_DataFrameToUpdateIsADataFrameToUpdate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Function INWMapSeriesOptions_DataFrameToUpdateListDataFrames() As Variant
996:   ConfirmDataStructures_DataFrameToUpdate
997:   INWMapSeriesOptions_DataFrameToUpdateListDataFrames = m_pDictDataFramesToUpdate_Layer.Keys
End Function

Private Sub INWMapSeriesOptions_DataFrameToUpdateRemove(sDataFrame As String)
  On Error GoTo ErrorHandler

1003:   ConfirmDataStructures_DataFrameToUpdate
1004:   With m_pDictDataFramesToUpdate_Layer
1005:     If .Exists(sDataFrame) Then
1006:       .Remove sDataFrame
1007:       m_pDictDataFramesToUpdate_Field.Remove sDataFrame
1008:     End If
1009:   End With

  Exit Sub
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_DataFrameToUpdateRemove " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub





'If an entry does not exist, create an entry.  If an entry already exists,
'then update existing values with new values.
'-----------------------------------------------------------------
Private Sub INWMapSeriesOptions_DataFrameToUpdateSetLayerAndField(sDataFrameName As String, sPageNameField As String, sMapPageLayer As String)
  On Error GoTo ErrorHandler

1026:   ConfirmDataStructures_DataFrameToUpdate
1027:   With m_pDictDataFramesToUpdate_Layer
1028:     If .Exists(sDataFrameName) Then
1029:       .Remove sDataFrameName
1030:     End If
1031:     .Add sDataFrameName, sMapPageLayer
1032:   End With
1033:   With m_pDictDataFramesToUpdate_Field
1034:     If .Exists(sDataFrameName) Then
1035:       .Remove sDataFrameName
1036:     End If
1037:     .Add sDataFrameName, sPageNameField
1038:   End With

  Exit Sub
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_DataFrameToUpdateSetLayerAndField " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub





'Cycle through the list of data frames in ArcMap's page layout, and whenever
'a data frame referenced in
'  m_pDictDataFramesToUpdate_Layer
'and in
'  m_pDictDataFramesToUpdate_Field
'use the polygon layer and the map name field in that polygon feature
'class to set the extent of that data frame.
'-----------------------------------------------------------------
Private Sub INWMapSeriesOptions_DataFrameToUpdateUpdateDataFrames(pMxDoc As IMxDocument, sPageName As String)
  On Error GoTo ErrorHandler

1060:   If pMxDoc Is Nothing Then
1061:     Err.Raise vbObjectError + 1, "NWMapSeries.cls, INWMapSeriesOptions_DataFrameToUpdateUpdateDataFrames", _
      "The MxDocument reference is empty. " & vbNewLine _
      & "To return this tool to a stable state, please try saving your map document" & vbNewLine _
      & "to a new file, closing and reopening ArcMap, then reopening this tool." & vbNewLine _
      & "If you continue to experience problems, please make note of the data and" & vbNewLine _
      & "the steps necessary to reproduce this error, then contact your source of " & vbNewLine _
      & "support for this application."
1068:   End If
  
  Dim pGraCont As IGraphicsContainer, pElement As IElement, pMapFrame As IMapFrame
  Dim pMap As IMap, pLayers As IEnumLayer, pFeature As IFeature, pLayer As ILayer
  Dim sPageLayerName As String, sMapPageField As String, sMapName As String
  Dim pFields As IField, pFeatureLayer As IFeatureLayer, lFieldIdx As Long
  Dim pQueryFilter As IQueryFilter, pFeatCursor As IFeatureCursor
  Dim pEnvelope As IEnvelope, pField As IField, bLyrFound As Boolean
  Dim pMainMapFrame As IMapFrame, pPageLayer As ILayer
  
1078:   Set pGraCont = pMxDoc.PageLayout
1079:   pGraCont.Reset
1080:   Set pElement = pGraCont.Next
1081:   Set pQueryFilter = New QueryFilter
  
                                            'grab the main map frame in case
                                            'user wants to set other data frames
                                            'to that same extent.
1086:   Do While Not pElement Is Nothing
1087:     If TypeOf pElement Is IMapFrame Then
1088:       Set pMapFrame = pElement
1089:       sMapName = pMapFrame.Map.Name
1090:       If StrComp(m_sMainDataFrame, sMapName, vbTextCompare) = 0 Then
1091:         Set pMainMapFrame = pElement
1092:       End If
1093:     End If
1094:     Set pElement = pGraCont.Next
1095:   Loop
  
1097:   pGraCont.Reset
1098:   Set pElement = pGraCont.Next
  
1100:   Do While Not pElement Is Nothing
1101:     If TypeOf pElement Is IMapFrame Then
1102:       Set pMapFrame = pElement
1103:       sMapName = pMapFrame.Map.Name
1104:       If m_pDictDataFramesToUpdate_Layer.Exists(sMapName) Then
1105:         sPageLayerName = m_pDictDataFramesToUpdate_Layer.Item(sMapName)
1106:         sMapPageField = m_pDictDataFramesToUpdate_Field.Item(sMapName)
                                            'acquire the polygon layer and the
                                            'attribute field reference so that
                                            'the polygon to use for setting the
                                            'extent of the data frame can found
1111:         Set pLayers = pMapFrame.Map.Layers
1112:         pLayers.Reset
1113:         Set pLayer = pLayers.Next
1114:         bLyrFound = False
1115:         Do While (Not pLayer Is Nothing) And (Not bLyrFound)
1116:           If (StrComp(pLayer.Name, sPageLayerName, vbTextCompare) = 0) Then
1117:             bLyrFound = True
1118:             Set pPageLayer = pLayer
1119:           End If
1120:           Set pLayer = pLayers.Next
1121:         Loop
                                            'layer is confirmed to exist, then grab
                                            'the attribute field containing the name
                                            'of the map page
1125:         If (Not pPageLayer Is Nothing) Then
1126:           If TypeOf pPageLayer Is IFeatureLayer Then
1127:             Set pFeatureLayer = pPageLayer
1128:             lFieldIdx = pFeatureLayer.FeatureClass.FindField(sMapPageField)
1129:             If lFieldIdx >= 0 Then
1130:               Set pField = pFeatureLayer.FeatureClass.Fields.Field(lFieldIdx)
1131:               pQueryFilter.WhereClause = pField.Name & " = '" & sPageName & "'"
1132:               Set pFeatCursor = pFeatureLayer.Search(pQueryFilter, True)
1133:               Set pFeature = pFeatCursor.NextFeature
1134:               If Not pFeature Is Nothing Then
                'set the extent based on user'd option selection
                'default value: OTHER_DF_EXTENTOPTION_POLYEXTENT
                
1138:                 SetDataFrameExtent pMapFrame, pFeature.Shape.Envelope
1139:                 If m_lDataFrameToUpdate_ExtentOption = OTHER_DF_EXTENTOPTION_MAINDFSCALE Then
1140:                   If Not pMainMapFrame Is Nothing Then
1141:                     pMapFrame.ExtentType = esriExtentScale
1142:                     pMapFrame.MapScale = pMainMapFrame.MapScale
1143:                   End If
1144:                 End If
1145:               End If
1146:             End If
1147:           End If
1148:         End If
1149:       End If
1150:     End If
1151:     Set pElement = pGraCont.Next
1152:   Loop
  

  Exit Sub
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_DataFrameToUpdateUpdateDataFrames " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub




'Given a map frame and an extent, set the extent of the map.
'This routine has been farmed out to a separate routine in case
'the extent update process should be modified in a manner similar
'to what happens with the main map frame.
'------------------------------
Private Sub SetDataFrameExtent(pMapFrame As IMapFrame, pEnvelope As IEnvelope)
  On Error GoTo ErrorHandler

  Dim pMap As IMap, pActiveView As IActiveView
  
1173:   Set pMap = pMapFrame.Map
1174:   Set pActiveView = pMap
1175:   pActiveView.Extent = pEnvelope

  Exit Sub
ErrorHandler:
  HandleError False, "SetDataFrameExtent " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub



Private Sub INWMapSeriesOptions_DynamicDefQueryAddLayer(sDataFrameID As String, sLayerID As String)
  On Error GoTo ErrorHandler

  Dim pScrDict As Scripting.Dictionary, sAttrib As String
  
1189:   If m_pDictDynDefQueryPages Is Nothing Then
1190:     Set m_pDictDynDefQueryPages = New Scripting.Dictionary
1191:   End If
1192:   If Not m_pDictDynDefQueryPages.Exists(sDataFrameID) Then
1193:     Set pScrDict = New Scripting.Dictionary
1194:     m_pDictDynDefQueryPages.Add sDataFrameID, pScrDict
1195:   End If
1196:   Set pScrDict = m_pDictDynDefQueryPages.Item(sDataFrameID)
  
1198:   If pScrDict Is Nothing Then
    Exit Sub
1200:   End If
1201:   If Not pScrDict.Exists(sLayerID) Then
1202:     sAttrib = ""   'stub for potential future support of other attrib info
1203:     pScrDict.Add sLayerID, sAttrib
1204:   End If
  
  Exit Sub
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_DynamicDefQueryAddLayer " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Function INWMapSeriesOptions_DynamicDefQueryDataFrames() As Variant
1212:   INWMapSeriesOptions_DynamicDefQueryDataFrames = m_pDictDynDefQueryPages.Keys
End Function

Private Property Get INWMapSeriesOptions_DynamicDefQueryIsTrackingLayer(sDataFrameID As String, sLayerID As String) As Variant
  On Error GoTo ErrorHandler
  Dim bResult As Boolean, pScrDict As Scripting.Dictionary
                                                  'see if there is a data structure for that
                                                  'combination of data frame and layer
1220:   bResult = False
1221:   If m_pDictDynDefQueryPages.Exists(sDataFrameID) Then
1222:     Set pScrDict = m_pDictDynDefQueryPages.Item(sDataFrameID)
1223:     If pScrDict.Exists(sLayerID) Then
1224:       bResult = True
1225:     End If
1226:   End If
1227:   INWMapSeriesOptions_DynamicDefQueryIsTrackingLayer = bResult

  Exit Property
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_DynamicDefQueryIsTrackingLayer " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property



Private Property Get INWMapSeriesOptions_DynamicDefQueryTrackingLayerCount() As Long
  On Error GoTo ErrorHandler

  Dim lDynLayerCount As Long, vDataFrames As Variant, lDataFrameCount As Long
  Dim pScrDict As Scripting.Dictionary, i As Long, sDataFrameName As String
  
1242:   vDataFrames = m_pDictDynDefQueryPages.Keys
1243:   lDataFrameCount = UBound(vDataFrames) + 1
1244:   lDynLayerCount = 0
1245:   For i = 0 To (lDataFrameCount - 1)
1246:     sDataFrameName = vDataFrames(i)
1247:     Set pScrDict = m_pDictDynDefQueryPages.Item(sDataFrameName)
1248:     If Not pScrDict Is Nothing Then
1249:       lDynLayerCount = lDynLayerCount + pScrDict.Count
1250:     End If
1251:   Next i
1252:   INWMapSeriesOptions_DynamicDefQueryTrackingLayerCount = lDynLayerCount

  Exit Property
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_DynamicDefQueryTrackingLayerCount " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property


Private Property Get INWMapSeriesOptions_DynamicDefQueryPrevReplaceString() As String
1261:   INWMapSeriesOptions_DynamicDefQueryPrevReplaceString = m_sPrevPageName
End Property

Private Sub INWMapSeriesOptions_DynamicDefQueryRemoveDataFrame(sDataFrameID As String)
  On Error GoTo ErrorHandler

  Dim pScrDict As Scripting.Dictionary
  
1269:   If m_pDictDynDefQueryPages Is Nothing Then
    Exit Sub
1271:   End If
1272:   If Not m_pDictDynDefQueryPages.Exists(sDataFrameID) Then
    Exit Sub
1274:   End If
1275:   Set pScrDict = m_pDictDynDefQueryPages.Item(sDataFrameID)
1276:   pScrDict.RemoveAll
1277:   m_pDictDynDefQueryPages.Remove sDataFrameID


  Exit Sub
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_DynamicDefQueryRemoveDataFrame " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub INWMapSeriesOptions_DynamicDefQueryRemoveLayer(sDataFrameID As String, sLayerID As String)
  On Error GoTo ErrorHandler

  Dim pScrDict As Scripting.Dictionary
  
1290:   If m_pDictDynDefQueryPages Is Nothing Then
1291:     Set m_pDictDynDefQueryPages = New Scripting.Dictionary
    Exit Sub
1293:   End If
1294:   If Not m_pDictDynDefQueryPages.Exists(sDataFrameID) Then
    Exit Sub
1296:   End If
1297:   Set pScrDict = m_pDictDynDefQueryPages.Item(sDataFrameID)
1298:   If pScrDict.Exists(sLayerID) Then
1299:     pScrDict.Remove sLayerID
1300:   End If
1301:   If pScrDict.Count = 0 Then
1302:     m_pDictDynDefQueryPages.Remove sDataFrameID
1303:   End If
  

  Exit Sub
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_DynamicDefQueryRemoveLayer " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Property Let INWMapSeriesOptions_DynamicDefQueryReplaceString(RHS As String)
1312:   m_sPrevPageName = m_sCurrentPageName
1313:   m_sCurrentPageName = RHS
End Property

Private Property Get INWMapSeriesOptions_DynamicDefQueryReplaceString() As String
1317:   INWMapSeriesOptions_DynamicDefQueryReplaceString = m_sCurrentPageName
End Property




'ElementIsQualifiedForVisibilityManagement
'
'confirm that the selected element is qualified to be tagged for visibility control
'  Is it a data frame? (therefore not qualified, detail inset or not)
'  Is it something else already being tracked for visibility
'  Is the pElementProperties.CustomProperty already in use?
'---------------------------------------------------------
Private Function ElementIsQualifiedForVisibilityManagement(pElem As IElement, Optional vPermitTaggedElems As Variant) As Boolean
  On Error GoTo ErrorHandler

  Dim bResult As Boolean, pElemProps As IElementProperties, vCustProp As Variant
  Dim msgBoxResult As VbMsgBoxResult, sStartString As String, bPermitTaggedElems As Boolean
  
  
  'handle optional parameter
  ''''''''''''''''''''''''''
1339:   If IsMissing(vPermitTaggedElems) Then
1340:     bPermitTaggedElems = False
1341:   Else
1342:     If StrComp(TypeName(vPermitTaggedElems), "boolean", vbTextCompare) = 0 Then
1343:       bPermitTaggedElems = vPermitTaggedElems
1344:     Else
1345:       bPermitTaggedElems = False
1346:     End If
1347:   End If
1348:   If pElem Is Nothing Then
1349:     ElementIsQualifiedForVisibilityManagement = False
    Exit Function
1351:   End If
  
1353:   bResult = True
1354:   Set pElemProps = pElem
1355:   If TypeOf pElem Is IMapFrame Then
1356:     ElementIsQualifiedForVisibilityManagement = False
    Exit Function
1358:   End If
1359:   If TypeOf pElem Is ITextElement Then
1360:     ElementIsQualifiedForVisibilityManagement = False
    Exit Function
1362:   End If
  
1364:   vCustProp = pElemProps.CustomProperty
1365:   If bResult = True Then
    'If Not IsEmpty(vCustProp) And (StrComp(TypeName(vCustProp), "string", vbTextCompare) = 0) Then
1367:     If Not IsEmpty(vCustProp) Then
                                                  'test for graphic elements associated with
                                                  'detail insets
1370:       If (StrComp(TypeName(vCustProp), "string", vbTextCompare) = 0) Then
1371:         If Len(vCustProp) >= Len(GBL_BUBBLE_IDENTIFIER) Then
1372:           sStartString = Left$(vCustProp, Len(GBL_BUBBLE_IDENTIFIER))
1373:           If StrComp(sStartString, GBL_BUBBLE_IDENTIFIER, vbTextCompare) = 0 Then
            'this element was tagged for this function,
            'now detect if this element was tagged in a previous
            'map book, or this one
            
            'if this was tagged but never cleared from a previous map book,
            'then give the green light to retag this element
1380:             If Not m_pDictTaggedElements.Exists(vCustProp) Then
1381:               bResult = False
1382:             End If
1383:           End If
1384:         End If
1385:       End If
                                                  'element wasn't a mapframe or part of a
                                                  'detail inset
1388:       If bResult = True Then
        'Warn the user if the custom property is
        'already in use, and will be overwritten
        'if this element is tagged
        'unless this custom property happens to
        'be a tag used by the layout visibility
        'function of this program.  In that case
        'don't throw the warning at the user, just
        'prevent a redundant tagging of this
        'element by returning false.
        
        
1400:         If Not bPermitTaggedElems Then
1401:           If (StrComp(TypeName(vCustProp), "string", vbTextCompare) = 0) Then
1402:             If Len(vCustProp) >= Len(GBL_ELEMENT_IDENTIFIER) Then
1403:               sStartString = Left$(vCustProp, Len(GBL_ELEMENT_IDENTIFIER))
1404:               If StrComp(sStartString, GBL_ELEMENT_IDENTIFIER, vbTextCompare) = 0 Then
1405:                 msgBoxResult = MsgBox("Warning: Custom property slot used to tag this element is already in use.  Overwrite? " & vbNewLine, vbYesNo, "Element Property Slot In Use")
1406:                 If msgBoxResult = vbNo Then
1407:                   bResult = False
1408:                 Else
1409:                   bResult = True
1410:                 End If
1411:               End If
1412:             End If
1413:           End If
1414:         End If 'not bPermitTaggedElems
        
        
1417:       End If 'bResult = True <-- element wasn't a mapframe or part of a detail inset
1418:     End If 'not IsEmpty(vCustProp)
1419:   End If 'bResult = True
1420:   ElementIsQualifiedForVisibilityManagement = bResult

  Exit Function
ErrorHandler:
  HandleError False, "ElementIsQualifiedForVisibilityManagement " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function




Private Function INWMapSeriesOptions_ElementsElementIsInStorage(sElementID As String) As Boolean
1431:   If Not m_pDictStoredElements Is Nothing Then
1432:     INWMapSeriesOptions_ElementsElementIsInStorage = m_pDictStoredElements.Exists(sElementID)
1433:   Else
1434:     INWMapSeriesOptions_ElementsElementIsInStorage = False
1435:   End If
End Function

'INWMapSeriesOptions_ElementsElementIsQualified
'
'Note: This function will return false for an otherwise qualified
'      element if that element is already being tracked for element
'      visibility.  If the caller wants true to be returned for elements
'      that are being tracked, then also call ElementsElementIsTagged.
'---------------------------------------------------------------
Private Function INWMapSeriesOptions_ElementsElementIsQualified(pElement As esriCarto.IElement) As Boolean
1446:   INWMapSeriesOptions_ElementsElementIsQualified _
    = ElementIsQualifiedForVisibilityManagement(pElement, True)
End Function


'ElementIsTagged
'
'Code in multiple locations tests whether or not a layout element is being
'tracked for visibility management.  It's not enough to just access
'm_pDictTaggedElements.Exists, because code bookends that call to test the
'CustomProperties of the element in question.  This function just puts all
'that code in one location.
'--------------------------------------------------
Private Function ElementIsTagged(pElem As IElement) As Boolean
  On Error GoTo ErrorHandler

  Dim vCustProp As Variant, pElemProps As IElementProperties, sProperty As String

1464:   If pElem Is Nothing Or m_pDictTaggedElements Is Nothing Then
1465:     ElementIsTagged = False
    Exit Function
1467:   End If
1468:   Set pElemProps = pElem
1469:   vCustProp = pElemProps.CustomProperty
1470:   If IsEmpty(vCustProp) Then
1471:     ElementIsTagged = False
    Exit Function
1473:   End If
  
1475:   If (StrComp(TypeName(vCustProp), "string", vbTextCompare) = 0) Then
1476:     sProperty = vCustProp
1477:     ElementIsTagged = m_pDictTaggedElements.Exists(sProperty)
1478:   Else
1479:     ElementIsTagged = False
1480:   End If

  Exit Function
ErrorHandler:
  HandleError False, "ElementIsTagged " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function




Private Function INWMapSeriesOptions_ElementsElementIsTagged(pElement As esriCarto.IElement) As Boolean
1491:   INWMapSeriesOptions_ElementsElementIsTagged = ElementIsTagged(pElement)
End Function



'INWMapSeriesOptions_ElementsFilterLayoutElements
'
'This routine controls element visibility based on two sources, those
'elements already in the pagelayout, and those elements in storage in
'm_pDictStoredElements.  This dual storage procedure is based on the fact
'that no method appears to be available for making a layout element
'non-visible while still leaving it in the pagelayout.
'------------------------------------------------------------
Private Sub INWMapSeriesOptions_ElementsFilterLayoutElements( _
    pMxDoc As esriArcMapUI.IMxDocument, _
    sMapPageID As String)
  On Error GoTo ErrorHandler

  
  Dim pActiveView As IActiveView, pMainMap As IMap, vNameKeys As Variant
  Dim sStoredElemName As String, pElement As IElement, lFinishedCount As Long
  Dim pGraphicsContainer As IGraphicsContainer, lMaxRemainingCount As Long
  Dim pElemProps As IElementProperties, vCustProp As Variant, i As Long
  Dim sCustomProp As String, pDictAlreadySparedElements As Scripting.Dictionary
  Dim pPageLayout As IPageLayout
  
  'check for bad input
  ''''''''''''''''''''
  
1520:   If pMxDoc Is Nothing Then
1521:     Err.Raise 0, "routine FilterLayoutElements in NWMapSeries.cls", _
              "The pMxDoc reference was detected to be equal to ''Nothing''."
    Exit Sub
1524:   End If
1525:   If Len(Trim$(sMapPageID)) = 0 Then
1526:     Err.Raise 0, "INWMapSeriesOptions_ElementsFilterLayoutElements", _
              "The routine was expecting the name of the map page (sMapPageID) on " & vbNewLine _
              & "which to manage which layout elements were supposed to be visible." & vbNewLine _
              & "The map page name this routine received was ''" & sMapPageID & "''."
    Exit Sub
1531:   End If
1532:   If m_pDictStoredElements Is Nothing Then
1533:     Set m_pDictStoredElements = New Scripting.Dictionary
1534:   End If
1535:   If m_pDictTaggedElements Is Nothing Then
1536:     Set m_pDictTaggedElements = New Scripting.Dictionary
1537:   End If
  
  'set the active data frame to the main map frame
  ''''''''''''''''''''''''''''''''''''''''''''''''
  
                                                  'activate the main data frame used
                                                  'by map series
1544:   Set pActiveView = pMxDoc.ActiveView
1545:   For i = 0 To pMxDoc.Maps.Count - 1
1546:     If StrComp(pMxDoc.Maps.Item(i).Name, m_sMainDataFrame, vbTextCompare) = 0 Then
1547:       Set pMainMap = pMxDoc.Maps.Item(i)
1548:     End If
1549:   Next i
1550:   If TypeOf pActiveView Is IPageLayout Then
1551:     Set pMxDoc.ActiveView.FocusMap = pMainMap
1552:   Else
1553:     Set pMxDoc.ActiveView = pMainMap
1554:   End If
1555:   Set pPageLayout = pMxDoc.PageLayout 'caused an error if the selected data frame wasn't the main data frame
1556:   Set pGraphicsContainer = pPageLayout
1557:   pGraphicsContainer.Reset
  
  
  
  'cycle through the stored elements, and restore
  'any to the layout that are visible for this map page.
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''
  
1565:   vNameKeys = m_pDictStoredElements.Keys
1566:   For i = 0 To (m_pDictStoredElements.Count - 1)
1567:     sStoredElemName = vNameKeys(i)
1568:     If INWMapSeriesOptions_ElementsElementIsVisibleOnMapPage(sMapPageID, sStoredElemName) Then
1569:       Set pElement = m_pDictStoredElements.Item(sStoredElemName)
1570:       pGraphicsContainer.AddElement pElement, 0
1571:       m_pDictStoredElements.Remove sStoredElemName
    'else
      'data frame should be left alone, it's already the way it should be
1574:     End If
1575:   Next i
    
    
    
  'cycle through layout elements, and remove those
  'elements that are not visible on this map page
  ''''''''''''''''''''''''''''''''''''''''''''''''
  
1583:   pGraphicsContainer.Reset
1584:   Set pElement = pGraphicsContainer.Next
                            'useful because removing an element requires resetting
                            'pGraphicsContainer.  That means cycling through all visible
                            'elements from the beginning whenever one is removed.  The spared
                            'elements list should allow this chunk of code to go faster by
                            'allowing the code to skip most processing of elements that were
                            'already evaluated as being visible on this map page.
1591:   Set pDictAlreadySparedElements = New Scripting.Dictionary
1592:   lFinishedCount = 0   'limit the loop in case of a multitude of non-imapframe elements
1593:   lMaxRemainingCount = m_pDictTaggedElements.Count
1594:   Do While ((Not pElement Is Nothing) And (lFinishedCount < lMaxRemainingCount))
1595:     If TypeOf pElement Is IMapFrame Then
1596:       GoTo nextPElementLoop
1597:     End If
1598:     If TypeOf pElement Is ITextElement Then
1599:       GoTo nextPElementLoop
1600:     End If
    
1602:     Set pElemProps = pElement
1603:     vCustProp = pElemProps.CustomProperty
1604:     If Not IsEmpty(vCustProp) Then
1605:       If StrComp(TypeName(vCustProp), "string", vbTextCompare) = 0 Then
1606:         sCustomProp = vCustProp
1607:         If Not pDictAlreadySparedElements.Exists(sCustomProp) Then
        
                                                      'if the element is marked as visible
                                                      'on this map page
1611:           If INWMapSeriesOptions_ElementsElementIsVisibleOnMapPage(sMapPageID, sCustomProp) Then
                                                      'let the element remain visible
    
1614:             pDictAlreadySparedElements.Add sCustomProp, pElement
1615:             lFinishedCount = lFinishedCount + 1
1616:           Else
                                                      'else make it not visible
1618:             m_pDictStoredElements.Add sCustomProp, pElement
1619:             pGraphicsContainer.DeleteElement pElement
1620:             pGraphicsContainer.Reset
1621:             lFinishedCount = lFinishedCount + 1
            
1623:           End If
1624:         End If 'element already spared from being removed from the layout
1625:       End If 'if vCustProp is a string
1626:     End If 'if not IsEmpty(vCustProp)
      
nextPElementLoop:
1629:     Set pElement = pGraphicsContainer.Next
1630:   Loop
  


  Exit Sub
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_ElementsFilterLayoutElements " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

'INWMapSeriesOptions_ElementsGetElementsInvisibleForMapPage
'
'  The data structures tracking element visibility store map pages per element,
'  with all map pages listed per element being map pages where that element
'  will not be visible.  To retrieve from this structure a list of elements
'  that are not visible for a given map page, all element data structures need
'  to be traversed to see if that map page is present.
'---------------------------------------------------------------------------
Private Function INWMapSeriesOptions_ElementsGetElementsInvisibleForMapPage(sMapPageID As String) As Variant
  On Error GoTo ErrorHandler

  Dim vResult As Variant, sElementIDs() As String, sElementID As String
  Dim lReturnElemCount As Long, pDictResultElements As Scripting.Dictionary
  Dim i As Long, lInvisElemCount As Long, vInvisElems As Variant
  Dim sInvisElem As String, pScrDict As Scripting.Dictionary
  ReDim sElementIDs(0) As String
  
  'do some QA on the routine input
1657:   If m_pDictPagesWhereElementIsInvisible Is Nothing Then
1658:     Set m_pDictPagesWhereElementIsInvisible = New Scripting.Dictionary
1659:   End If
1660:   If Len(sMapPageID) = 0 Then
1661:     vResult = sElementIDs
1662:     INWMapSeriesOptions_ElementsGetElementsInvisibleForMapPage = vResult
    Exit Function
1664:   End If

1666:   lReturnElemCount = 0
1667:   With m_pDictPagesWhereElementIsInvisible
1668:     lInvisElemCount = .Count
1669:     vInvisElems = .Keys
1670:     For i = 0 To (lInvisElemCount - 1)
1671:       sInvisElem = vInvisElems(i)
1672:       Set pScrDict = .Item(sInvisElem)
1673:       If pScrDict.Exists(sMapPageID) Then
1674:         lReturnElemCount = lReturnElemCount + 1
        ReDim Preserve sElementIDs(lReturnElemCount) As String
1676:         sElementIDs(lReturnElemCount - 1) = sInvisElem
1677:       End If
1678:     Next i
1679:   End With
1680:   vResult = sElementIDs
1681:   INWMapSeriesOptions_ElementsGetElementsInvisibleForMapPage = vResult

  Exit Function
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_ElementsGetElementsInvisibleForMapPage " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'ElementsGetMapPagesWhereInvisible
'
'Function returns the list of map pages where the application will drop the
'selected layout element from being visible.
'----------------------------------------------------------------------
Private Function INWMapSeriesOptions_ElementsGetMapPagesWhereInvisible(pElement As esriCarto.IElement) As Variant
  On Error GoTo ErrorHandler

    
  Dim pElemProps As IElementProperties, vCustProp As Variant, sElemID As String
  Dim sResult() As String, lInvisPageCount As Long, pScrDict As Scripting.Dictionary
  Dim vResult As Variant
  
  
1702:   If pElement Is Nothing Then
1703:     lInvisPageCount = 0
    ReDim sResult(lInvisPageCount) As String
1705:     vResult = sResult
1706:     INWMapSeriesOptions_ElementsGetMapPagesWhereInvisible = vResult
    Exit Function
1708:   End If
  
  'confirm that the element is tagged
1711:   If Not ElementIsTagged(pElement) Then
1712:     lInvisPageCount = 0
    ReDim sResult(lInvisPageCount) As String
1714:     vResult = sResult
1715:     INWMapSeriesOptions_ElementsGetMapPagesWhereInvisible = vResult
    Exit Function
1717:   End If
  
  'confirm that the data structure is tracking
  'the element
  ''''''''''''''''''''''''''''''''''''''''''''
1722:   If m_pDictPagesWhereElementIsInvisible Is Nothing Then
1723:     Set m_pDictPagesWhereElementIsInvisible = New Scripting.Dictionary
1724:   End If
1725:   Set pElemProps = pElement
1726:   vCustProp = pElemProps.CustomProperty
1727:   sElemID = vCustProp
  
  
  'get the list of pages where this element is
  'not visible
  ''''''''''''''''''''''''''''''''''''''''''''
1733:   With m_pDictPagesWhereElementIsInvisible
1734:     If Not .Exists(sElemID) Then
1735:       Set pScrDict = New Scripting.Dictionary
1736:       .Add sElemID, pScrDict
1737:     Else
1738:       Set pScrDict = .Item(sElemID)
1739:     End If
1740:   End With
1741:   vResult = pScrDict.Keys
1742:   INWMapSeriesOptions_ElementsGetMapPagesWhereInvisible = vResult

  Exit Function
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_ElementsGetMapPagesWhereInvisible " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function



Private Function INWMapSeriesOptions_ElementsGetTaggedElements() As Variant
  On Error GoTo ErrorHandler

  Dim vResult As Variant, sResult() As String
  
1756:   If m_pDictTaggedElements Is Nothing Then
    ReDim sResult(0) As String
1758:     vResult = sResult
1759:     INWMapSeriesOptions_ElementsGetTaggedElements = vResult
    Exit Function
1761:   End If
1762:   INWMapSeriesOptions_ElementsGetTaggedElements = m_pDictTaggedElements.Keys
  

  Exit Function
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_ElementsGetTaggedElements " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function


'Given a map page and an element name, indicate that the element is not visible
'if that map page and element combination exist.  Data structures track map pages
'where this element will *not* appear.  An absence of a map page listing means
'that the layout element will be made visible.
'----------------------------------------------------------------------
Private Function INWMapSeriesOptions_ElementsElementIsVisibleOnMapPage(sMapPageID As String, sElementID As String) As Boolean
  On Error GoTo ErrorHandler
  
  Dim pScrDict As Scripting.Dictionary
  
  'if the tracking data structure is nothing, or
  'if the named element isn't tracked, return true
  'to indicate that the element is visible
1784:   If m_pDictPagesWhereElementIsInvisible Is Nothing Then
1785:     Set m_pDictPagesWhereElementIsInvisible = New Scripting.Dictionary
1786:     INWMapSeriesOptions_ElementsElementIsVisibleOnMapPage = True
    Exit Function
1788:   End If
1789:   If Not m_pDictPagesWhereElementIsInvisible.Exists(sElementID) Then
1790:     INWMapSeriesOptions_ElementsElementIsVisibleOnMapPage = True
    Exit Function
1792:   End If
  
1794:   Set pScrDict = m_pDictPagesWhereElementIsInvisible.Item(sElementID)
  
1796:   If pScrDict Is Nothing Then
1797:     INWMapSeriesOptions_ElementsElementIsVisibleOnMapPage = True
    Exit Function
1799:   End If
  
1801:   INWMapSeriesOptions_ElementsElementIsVisibleOnMapPage = _
    Not pScrDict.Exists(sMapPageID)
    
  Exit Function
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_ElementsElementIsVisibleOnMapPage " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function



'INWMapSeriesOptions_ElementsSetElementsInvisibleForMapPage
'
'  Underlying data structures track map pages per element, not elements
'  per map page.  This routine will walk through the various element
'  objects, and set whether or not the map page is present based on
'  whether that data object's element name was part of the list of
'  elements that was passed in.
'----------------------------------------------------------------------
Private Sub INWMapSeriesOptions_ElementsSetElementsInvisibleForMapPage(sElementIDs() As String, sMapPageID As String)
  On Error GoTo ErrorHandler

  Dim lElementCount As Long, i As Long, sElementID As String
  Dim pDictPagesWhereInvisible As Scripting.Dictionary, sPlaceHolder As String
  Dim sInvisElem As String, vInvisElems As Variant, pScrDict As Scripting.Dictionary
  Dim pDictNewElements As Scripting.Dictionary
  
  'do some QA on the routine input
1828:   If m_pDictPagesWhereElementIsInvisible Is Nothing Then
1829:     Set m_pDictPagesWhereElementIsInvisible = New Scripting.Dictionary
1830:   End If
1831:   If Len(sMapPageID) = 0 Then
1832:     MsgBox "Warning: Attempted to set layout elements visible for an empty map page name." & vbNewLine _
         & "(The map page name was ''''.)  Application cannot set element visibility for" & vbNewLine _
         & "an empty map page name.", vbCritical, "Error setting layout element visibility"
    Exit Sub
1836:   End If
  
  'keep better track of map ids caller wants to set
1839:   Set pDictNewElements = New Scripting.Dictionary
1840:   lElementCount = UBound(sElementIDs) + 1  'subscript out of range - called from frmVisibleElements, cmdOK_Click
1841:   For i = 0 To (lElementCount - 1)
    'populate local scripting.dictionary
1843:     sElementID = sElementIDs(i)
    
1845:     If Len(Trim$(sElementID)) > 0 Then
1846:       pDictNewElements.Add sElementID, Nothing
1847:     End If
1848:   Next i
  


  'set element invisibility for the given map page
  
1854:   sPlaceHolder = " "
1855:   vInvisElems = pDictNewElements.Keys
1856:   lElementCount = pDictNewElements.Count
1857:   For i = 0 To (lElementCount - 1)
1858:     sInvisElem = vInvisElems(i)
1859:     If Not m_pDictPagesWhereElementIsInvisible.Exists(sInvisElem) Then
1860:       Set pScrDict = New Scripting.Dictionary
1861:       pScrDict.Add sMapPageID, sPlaceHolder
1862:       m_pDictPagesWhereElementIsInvisible.Add sInvisElem, pScrDict
1863:     Else
1864:       Set pScrDict = m_pDictPagesWhereElementIsInvisible.Item(sInvisElem)
1865:       If Not pScrDict.Exists(sMapPageID) Then
1866:         pScrDict.Add sMapPageID, sPlaceHolder
1867:       End If
1868:     End If
1869:   Next i
  
  
'  vInvisElems = m_pDictPagesWhereElementIsInvisible.Keys
'  lElementCount = m_pDictPagesWhereElementIsInvisible.Count
'  'cycle through *all* elements, not just
'  'the ones passed in.
'  'For each element in the m_pdictpageswhereelementisinvisible,
'  ''''''''''''''''''''
'
'
'
'
'  For i = 0 To (lElementCount - 1)
'    sInvisElem = vInvisElems(i)
'    'if that element is in the passed list,
'    'then make sure that map page is present for that element
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Set pDictPagesWhereInvisible = m_pDictPagesWhereElementIsInvisible.Item(sInvisElem)
'    If pDictPagesWhereInvisible Is Nothing Then
'      Set pDictPagesWhereInvisible = New Scripting.Dictionary
'    End If
'
'    If pDictNewElements.Exists(sInvisElem) Then
'      If Not pDictPagesWhereInvisible.Exists(sMapPageID) Then
'        pDictPagesWhereInvisible.Add sMapPageID, sPlaceHolder
'      End If
'
'    'if that element is not present in the passed list,
'    'then make sure that map page is not present for that element
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Else
'      If pDictPagesWhereInvisible.Exists(sMapPageID) Then
'        pDictPagesWhereInvisible.Remove (sMapPageID)
'      End If
'    End If 'pDictNewElements.Exists(sInvisElem)
'  Next i


  Exit Sub
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_ElementsSetElementsInvisibleForMapPage " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

'ElementsSetMapPagesWhereInvisible
'
'Routine assumes that all map page names are valid, and just stores the names.
'pElement is confirmed by code to be valid.
'---------------------------------------------------------------
Private Sub INWMapSeriesOptions_ElementsSetMapPagesWhereInvisible( _
  sMapPages() As String, _
  pElement As IElement)
  On Error GoTo ErrorHandler
                                                  'sPlaceHolder holds a place in data structure
                                                  'for future modifications without breaking
                                                  'save/load code
  Dim sPlaceHolder As String
  
  Dim lPageCount As Long, pElemProps As IElementProperties, vCustProp As Variant
  Dim sElemID As String, i As Long, pScrDict As Scripting.Dictionary
  Dim sMapPage As String
  
  If pElement Is Nothing Then Exit Sub
1932:   If Not ElementIsTagged(pElement) Then
1933:     MsgBox "Warning: An attempt was made to control the visibility of a layout" & vbNewLine _
         & "element that had not been tagged for visibility control.", vbCritical, "Layout Element not Tagged"
    Exit Sub
1936:   End If
  
  'I can assume by here that element was tagged.
  'confirm that data structs tracking visibility are set
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''
1941:   If m_pDictPagesWhereElementIsInvisible Is Nothing Then
1942:     Set m_pDictPagesWhereElementIsInvisible = New Scripting.Dictionary
1943:   End If
1944:   Set pElemProps = pElement
1945:   vCustProp = pElemProps.CustomProperty
1946:   sElemID = vCustProp
  
1948:   With m_pDictPagesWhereElementIsInvisible
1949:     If Not .Exists(sElemID) Then
1950:       Set pScrDict = New Scripting.Dictionary
1951:       .Add sElemID, pScrDict
1952:     Else
1953:       Set pScrDict = .Item(sElemID)
1954:     End If
1955:   End With 'm_pDictPagesWhereElementIsInvisible
  
1957:   lPageCount = UBound(sMapPages) + 1
1958:   sPlaceHolder = " "
1959:   pScrDict.RemoveAll
1960:   For i = 0 To (lPageCount - 1)
1961:     sMapPage = sMapPages(i)
1962:     pScrDict.Add sMapPage, sPlaceHolder
1963:   Next i
  

  Exit Sub
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_ElementsSetMapPagesWhereInvisible " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Property Get INWMapSeriesOptions_ElementsStored() As Variant
  Dim vEmpty As Variant
1973:   If Not m_pDictStoredElements Is Nothing Then
1974:     INWMapSeriesOptions_ElementsStored = m_pDictStoredElements.Keys
1975:   Else
1976:     INWMapSeriesOptions_ElementsStored = vEmpty
1977:   End If
End Property


'Returns the stored element of the given name
'----------------------------------------------------------
Private Function INWMapSeriesOptions_ElementsStoredElement(sElementID As String) As esriCarto.IElement
  On Error GoTo ErrorHandler

1986:   If m_pDictStoredElements Is Nothing Then
1987:     Set m_pDictStoredElements = New Scripting.Dictionary
1988:     Set INWMapSeriesOptions_ElementsStoredElement = Nothing
1989:   Else
1990:     If Not m_pDictStoredElements.Exists(sElementID) Then
1991:       Set INWMapSeriesOptions_ElementsStoredElement = Nothing
1992:     Else
1993:       Set INWMapSeriesOptions_ElementsStoredElement = m_pDictStoredElements.Item(sElementID)
1994:     End If
1995:   End If
  

  Exit Function
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_ElementsStoredElement " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

'Put the passed element into element storage for those elements that
'are not visible in the current map page.
'----------------------------------------------------
Private Sub INWMapSeriesOptions_ElementsStoreElement(pElement As IElement)
  On Error GoTo ErrorHandler

  Dim vCustProp As Variant, sElemName As String, pElemProps As IElementProperties
  
  'Check for bad input
  ''''''''''''''''''''
  
  If pElement Is Nothing Then Exit Sub
2015:   If m_pDictStoredElements Is Nothing Then
2016:     Set m_pDictStoredElements = New Scripting.Dictionary
2017:   End If
2018:   If m_pDictTaggedElements Is Nothing Then
2019:     MsgBox "The data structure that keeps track of tagged elements, " & vbNewLine _
         & "m_pDictTaggedElements, turned out to be empty when it shouldn't have been." & vbNewLine _
         & "It has been reinitialized, but any previously tagged elements are no longer" & vbNewLine _
         & "being tracked.  It is recommended that you close ArcMap without saving this map" & vbNewLine _
         & "document, and reload.", vbCritical, "Tagged Element Structure is Empty"
2024:     Set m_pDictTaggedElements = New Scripting.Dictionary
    Exit Sub
2026:   End If
                                                  'all the screening for IMapFrame,
                                                  'ITextElement, an empty CustomProperty,
                                                  'and a custom property that isn't a string
                                                  'may be redundant to screenings that happen
                                                  'before this routine is called.  This routine will
                                                  'be more robust however so that this screening is
                                                  'something you don't have to think/worry about
                                                  'when deciding when/from where to call this routine.
  If TypeOf pElement Is IMapFrame Then Exit Sub
  If TypeOf pElement Is ITextElement Then Exit Sub
  
2038:   Set pElemProps = pElement
2039:   vCustProp = pElemProps.CustomProperty
2040:   If IsEmpty(vCustProp) Then
2041:     GoTo NoTagMessage
2042:   End If
2043:   If StrComp(TypeName(vCustProp), "string", vbTextCompare) <> 0 Then
2044:     GoTo NoTagMessage
2045:   End If
  
2047:   sElemName = vCustProp
2048:   If Not m_pDictTaggedElements.Exists(sElemName) Then
2049:     GoTo NoTagMessage
2050:   End If
  
  
  'store the layout element
  '''''''''''''''''''''''''
2055:   m_pDictStoredElements.Add sElemName, pElement
  
  
  
  Exit Sub
NoTagMessage:
2061:     MsgBox "Attempted to store an element that was not tagged." & vbNewLine _
         & "Please call the layout element storage function only for " & vbNewLine _
         & "tagged elements.  User Interface code should screen out the" & vbNewLine _
         & "possibility of the user ever seeing this message.  If you are" & vbNewLine _
         & "user and not a programmer, please take note of the steps to " & vbNewLine _
         & "produce this message.  ", vbOKOnly


  Exit Sub
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_ElementsStoreElement " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

'ElementsTagElement
'
'  Respond when the user tags a layout element to manage visibility of
'  that element.
'--------------------------------------------------
Private Sub INWMapSeriesOptions_ElementsTagElement(pMxDoc As esriArcMapUI.IMxDocument)
  On Error GoTo ErrorHandler

  Dim pPageLayout As IPageLayout, pElements As IEnumElement, pElement As IElement
  Dim pGraphicsContainer As IGraphicsContainer, pElemProps As IElementProperties
  Dim sNewProperty As String, pGraphicsContSelect As IGraphicsContainerSelect
  
  ''''get a reference to the selected element
2087:   Set pPageLayout = pMxDoc.PageLayout
2088:   Set pGraphicsContainer = pPageLayout
2089:   Set pGraphicsContSelect = pGraphicsContainer
2090:   Set pElements = pGraphicsContSelect.SelectedElements
2091:   If pGraphicsContSelect.ElementSelectionCount = 0 Then
2092:     MsgBox "Warning: At least one layout element must be selected before it can be tagged." & vbNewLine
    Exit Sub
2094:   End If
  
  
2097:   pElements.Reset
2098:   Set pElement = pElements.Next
  
2100:   Do While Not pElement Is Nothing
2101:     If ElementIsQualifiedForVisibilityManagement(pElement, False) Then
                                                  'set the CustomProperty of the element
2103:       Set pElemProps = pElement
2104:       GBL_ELEMENT_COUNT = GBL_ELEMENT_COUNT + 1
2105:       sNewProperty = GBL_ELEMENT_IDENTIFIER & "_" & GBL_ELEMENT_COUNT & "_" & pElemProps.Name
2106:       pElemProps.CustomProperty = sNewProperty
      'set the tracking data structure to know
      'that this element should be
      'tracked for visibility.
2110:       m_pDictTaggedElements.Add sNewProperty, pElement
    
2112:     End If
2113:     Set pElement = pElements.Next
2114:   Loop

  Exit Sub
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_ElementsTagElement " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub





'INWMapSeriesOptions_ElementsUntagElement
'
'  Routine releases a layout element from tracking for visibility
'  per map page.
'----------------------------------------------------
Private Sub INWMapSeriesOptions_ElementsUntagElement(pMxDoc As IMxDocument)
  On Error GoTo ErrorHandler

  Dim pElemProps As IElementProperties, vCustProp As Variant, sProperty As String
  Dim pScrDict As Scripting.Dictionary, sNewProperty As String
  Dim pPageLayout As IPageLayout, pElements As IEnumElement, pElement As IElement
  Dim pGraphicsContainer As IGraphicsContainer
  Dim pGraphicsContSelect As IGraphicsContainerSelect
  
  ''''check for bad input
  '''''''''''''''''''''''
2141:   If Not m_pDictPagesWhereElementIsInvisible Is Nothing Then
2142:     m_pDictTaggedElements.RemoveAll
2143:     Err.Raise 0, "NWMapSeries.cls, INWMapSeriesOptions_ElementsUntagElement", _
              "m_pDictPagesWhereElementIsInvisible, the data structure for tracking" & vbNewLine _
            & "which layout elements are visible on which map pages, was empty." & vbNewLine _
            & "Please take note of the steps to reproduce this error." & vbNewLine _
            & "All layout element tags have been cleared in an attempt" & vbNewLine _
            & "to keep the application in a stable state."
    Exit Sub
2150:   End If
  
  ''''get a reference to the selected elements,
  ''''let the user untag multiple elements at once.
  ''''''''''''''''''''''''''''''''''''''''''''''''
2155:   Set pPageLayout = pMxDoc.PageLayout
2156:   Set pGraphicsContainer = pPageLayout
2157:   Set pGraphicsContSelect = pGraphicsContainer
2158:   If pGraphicsContSelect.ElementSelectionCount = 0 Then
2159:     MsgBox "Warning: At least one layout element must be selected before it can be tagged." & vbNewLine
    Exit Sub
2161:   End If
2162:   Set pElements = pGraphicsContSelect.SelectedElements
  
2164:   pElements.Reset
2165:   Set pElement = pElements.Next

2167:   Do While Not pElement Is Nothing
    ''''Only untags those elements that are tagged
    ''''for visibility management.
    ''''''''''''''''''''''''''''''
2171:     If ElementIsTagged(pElement) Then
      ''''untag the element
      '''''''''''''''''''''
2174:       Set pElemProps = pElement
2175:       vCustProp = pElemProps.CustomProperty
2176:       sProperty = vCustProp
2177:       m_pDictTaggedElements.Remove sProperty
      
      ''''update element visibility tracking
      ''''''''''''''''''''''''''''''''''''''
2181:       With m_pDictPagesWhereElementIsInvisible
2182:         If .Exists(sProperty) Then
2183:           Set pScrDict = .Item(sProperty)
2184:           pScrDict.RemoveAll
2185:         End If
2186:         .Remove (sProperty)
2187:       End With
2188:       pElemProps.CustomProperty = ""
2189:     End If 'if the custom property is a string
    
2191:     Set pElement = pElements.Next
2192:   Loop
  
  Exit Sub
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_ElementsUntagElement " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

'This routine takes care of all functions for untagging an element, but one.
'That is this routine doesn't have access to the original element to undo
'the IElementProperties.CustomElement value.
'----------------------------------------------------------
Private Sub INWMapSeriesOptions_ElementsUntagElementString(sElemName As String)
  On Error GoTo ErrorHandler

  Dim pScrDict As Scripting.Dictionary
  
  'remove the element from the list of elements
  'and the map pages where they are not visible.
  ''''''''''''''''''''''''''''''''''''''''''''''
  If m_pDictPagesWhereElementIsInvisible Is Nothing Then Exit Sub
2212:   If m_pDictPagesWhereElementIsInvisible.Exists(sElemName) Then
2213:     Set pScrDict = m_pDictPagesWhereElementIsInvisible.Item(sElemName)
2214:     If Not pScrDict Is Nothing Then
2215:       pScrDict.RemoveAll
2216:     End If
2217:     m_pDictPagesWhereElementIsInvisible.Remove (sElemName)
2218:   End If

  
  'remove the element from the list of tagged elements
  ''''''''''''''''''''''''''''''''''''''''''''''''''''
2223:   If m_pDictTaggedElements Is Nothing Then
    Exit Sub
2225:   End If
2226:   If m_pDictTaggedElements.Exists(sElemName) Then
2227:     m_pDictTaggedElements.Remove (sElemName)
2228:   End If
  

  Exit Sub
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_ElementsUntagElementString " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub








Private Function INWMapSeriesOptions_IsBubbleLayer(pFeatClass As esriGeoDatabase.IFeatureClass) As Boolean
  On Error GoTo ErrorHandler

  Dim i As Long, lFieldCount As Long, pField As IField, pFields As IFields
  Dim lFieldIdx As Long
  
2249:   If pFeatClass Is Nothing Then
2250:     INWMapSeriesOptions_IsBubbleLayer = False
    Exit Function
2252:   End If
                                                  'confirm that this FC is a polygon feature class
2254:   If pFeatClass.ShapeType <> esriGeometryPolygon Then
2255:     INWMapSeriesOptions_IsBubbleLayer = False
    Exit Function
2257:   End If
                                                  'confirm that this FC contains all field names
                                                  'and corresponding field types used by the
                                                  'bubble details feature class
2261:   Set pFields = pFeatClass.Fields
2262:   lFieldCount = pFields.FieldCount
2263:   lFieldIdx = pFields.FindField("XDEST")
2264:   If lFieldIdx = -1 Then
2265:     INWMapSeriesOptions_IsBubbleLayer = False
    Exit Function
2267:   ElseIf (pFields.Field(lFieldIdx).Type <> esriFieldTypeDouble) Then
2268:     INWMapSeriesOptions_IsBubbleLayer = False
    Exit Function
2270:   End If
  
  
2273:   lFieldIdx = pFields.FindField("YDEST")
2274:   If lFieldIdx = -1 Then
2275:     INWMapSeriesOptions_IsBubbleLayer = False
    Exit Function
2277:   ElseIf (pFields.Field(lFieldIdx).Type <> esriFieldTypeDouble) Then
2278:     INWMapSeriesOptions_IsBubbleLayer = False
    Exit Function
2280:   End If
  
  
2283:   lFieldIdx = pFields.FindField("RADIUS")
2284:   If lFieldIdx = -1 Then
2285:     INWMapSeriesOptions_IsBubbleLayer = False
    Exit Function
2287:   ElseIf (pFields.Field(lFieldIdx).Type <> esriFieldTypeDouble) Then
2288:     INWMapSeriesOptions_IsBubbleLayer = False
    Exit Function
2290:   End If
  
  
2293:   lFieldIdx = pFields.FindField("SCALE")
2294:   If lFieldIdx = -1 Then
2295:     INWMapSeriesOptions_IsBubbleLayer = False
    Exit Function
2297:   ElseIf (pFields.Field(lFieldIdx).Type <> esriFieldTypeDouble) Then
2298:     INWMapSeriesOptions_IsBubbleLayer = False
    Exit Function
2300:   End If
  
  
2303:   lFieldIdx = pFields.FindField("XORG")
2304:   If lFieldIdx = -1 Then
2305:     INWMapSeriesOptions_IsBubbleLayer = False
    Exit Function
2307:   ElseIf (pFields.Field(lFieldIdx).Type <> esriFieldTypeDouble) Then
2308:     INWMapSeriesOptions_IsBubbleLayer = False
    Exit Function
2310:   End If
  
  
2313:   lFieldIdx = pFields.FindField("YORG")
2314:   If lFieldIdx = -1 Then
2315:     INWMapSeriesOptions_IsBubbleLayer = False
    Exit Function
2317:   ElseIf (pFields.Field(lFieldIdx).Type <> esriFieldTypeDouble) Then
2318:     INWMapSeriesOptions_IsBubbleLayer = False
    Exit Function
2320:   End If
  
  
2323:   lFieldIdx = pFields.FindField("BUBBLEID")
2324:   If lFieldIdx = -1 Then
2325:     INWMapSeriesOptions_IsBubbleLayer = False
    Exit Function
2327:   ElseIf (pFields.Field(lFieldIdx).Type <> esriFieldTypeInteger) Then
2328:     INWMapSeriesOptions_IsBubbleLayer = False
    Exit Function
2330:   End If
  
  
2333:   lFieldIdx = pFields.FindField("LAYERS")
2334:   If lFieldIdx = -1 Then
2335:     INWMapSeriesOptions_IsBubbleLayer = False
    Exit Function
2337:   ElseIf (pFields.Field(lFieldIdx).Type <> esriFieldTypeString) Then
2338:     INWMapSeriesOptions_IsBubbleLayer = False
    Exit Function
2340:   End If
  
  
2343:   lFieldIdx = pFields.FindField("WIDTHORG")
2344:   If lFieldIdx = -1 Then
2345:     INWMapSeriesOptions_IsBubbleLayer = False
    Exit Function
2347:   ElseIf (pFields.Field(lFieldIdx).Type <> esriFieldTypeDouble) Then
2348:     INWMapSeriesOptions_IsBubbleLayer = False
    Exit Function
2350:   End If
  
2352:   INWMapSeriesOptions_IsBubbleLayer = True
  

  Exit Function
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_IsBubbleLayer " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Property Get INWMapSeriesOptions_LayerGroupCount() As Long
2361:   INWMapSeriesOptions_LayerGroupCount = m_pLyrInvisGroups.Count
End Property

Private Function INWMapSeriesOptions_LayerGroupExists(sGroupName As String) As Boolean
2365:   INWMapSeriesOptions_LayerGroupExists = m_pLyrInvisGroups.Exists(sGroupName)
End Function

Private Function INWMapSeriesOptions_LayerGroupGet(sGroupName As String) As INWLayerVisibilityGroup
  On Error GoTo ErrorHandler
2370:   Set INWMapSeriesOptions_LayerGroupGet = m_pLyrInvisGroups.Item(sGroupName)

  Exit Function
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_LayerGroupGet " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function

Private Sub INWMapSeriesOptions_LayerGroupSet(sGroupName As String, pGroup As INWLayerVisibilityGroup)
  On Error GoTo ErrorHandler

2380:   If pGroup Is Nothing Then
2381:     If m_pLyrInvisGroups.Exists(sGroupName) Then
2382:       m_pLyrInvisGroups.Remove (sGroupName)
2383:     End If
2384:   Else
2385:     pGroup.Name = sGroupName
2386:     m_pLyrInvisGroups.Add sGroupName, pGroup
2387:   End If

  Exit Sub
ErrorHandler:
  HandleError True, "INWMapSeriesOptions_LayerGroupSet " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Property Get INWMapSeriesOptions_LayerGroups() As Variant
2395:   INWMapSeriesOptions_LayerGroups = m_pLyrInvisGroups.Keys
End Property






Private Property Let INWMapSeriesOptions_PreviousPageName(RHS As String)
2404:   m_sPrevPageName = RHS
End Property

Private Property Get INWMapSeriesOptions_PreviousPageName() As String
2408:   INWMapSeriesOptions_PreviousPageName = m_sPrevPageName
End Property


Private Property Get INWMapSeriesOptions_RefreshEventLoadPage() As Boolean
2413:   INWMapSeriesOptions_RefreshEventLoadPage = m_bHandleMapRefreshEvents
End Property

Private Property Let INWMapSeriesOptions_RefreshEventLoadPage(RHS As Boolean)
2417:   m_bHandleMapRefreshEvents = RHS
End Property

Private Property Let INWMapSeriesOptions_RefreshEventShouldntRefresh(RHS As Boolean)
2421:   m_bShouldntRefresh = RHS
End Property

                                            'sometimes a detail inset has been temporarily deleted, but
                                            'the refresh event is triggered before the replacement detail
                                            'inset is added.  The RefreshEventShouldntRefresh property is used
                                            'to prevent the refresh event from causing bugs because a detail
                                            'inset hasn't yet been added.  This happens when editing an
                                            'existing detail inset using the inset editing tool.
Private Property Get INWMapSeriesOptions_RefreshEventShouldntRefresh() As Boolean
2431:   INWMapSeriesOptions_RefreshEventShouldntRefresh = m_bShouldntRefresh
End Property

Private Sub INWMapSeriesOptions_TextSymbolAdd(pTextSymbol As esriDisplay.ISimpleTextSymbol, sSymbolName As String)
2435:   m_pTextSymbolDict.Add sSymbolName, pTextSymbol
End Sub

Private Property Let INWMapSeriesOptions_TextSymbolDefault(sSymbolName As String)
2439:   m_sDefaultTextSymbol = sSymbolName
End Property

Private Property Get INWMapSeriesOptions_TextSymbolDefault() As String
2443:   INWMapSeriesOptions_TextSymbolDefault = m_sDefaultTextSymbol
End Property

Private Property Get INWMapSeriesOptions_TextSymbolExists(sSymbolName As String) As Boolean
2447:   INWMapSeriesOptions_TextSymbolExists = m_pTextSymbolDict.Exists(sSymbolName)
End Property

Private Sub INWMapSeriesOptions_TextSymbolRemove(sKey As String)
2451:   m_pTextSymbolDict.Remove sKey
End Sub

Private Property Get INWMapSeriesOptions_TextSymbol(sKey As String) As esriDisplay.ISimpleTextSymbol
2455:   If sKey = "" Then
2456:     sKey = m_sDefaultTextSymbol
2457:     If m_sDefaultTextSymbol = "" Then
2458:       Set INWMapSeriesOptions_TextSymbol = Nothing
      Exit Property
2460:     End If
2461:   End If
2462:   Set INWMapSeriesOptions_TextSymbol = m_pTextSymbolDict.Item(sKey)
End Property

Private Property Get INWMapSeriesOptions_TextSymbolCount() As Long
2466:   INWMapSeriesOptions_TextSymbolCount = m_pTextSymbolDict.Count
End Property

Public Property Get INWMapSeriesOptions_TextSymbolNames() As Variant
2470:   INWMapSeriesOptions_TextSymbolNames = m_pTextSymbolDict.Keys
End Property

Private Sub INWMapSeriesOptions_TextSymbolRename(sPrevName As String, RHS As String)
2474:   m_pTextSymbolDict.Key(sPrevName) = RHS
End Sub

Private Sub INWMapSeriesOptions_TextSymbolSet(sKey As String, pTextSymbol As esriDisplay.ISimpleTextSymbol)
2478:   If Not m_pTextSymbolDict.Exists(sKey) Then
    Exit Sub
2480:   End If
2481:   If Not pTextSymbol Is Nothing Then
2482:     m_pTextSymbolDict.Item(sKey) = pTextSymbol
2483:   End If
End Sub








Private Sub INWDSMapSeries_AddPage(ByVal Page As NWMapBookPrj.INWDSMapPage)
2494:   m_PageColl.Add Page, Page.PageName  'RAH added ", Page.PageName" on 10/6/05 to support map refresh event handling
End Sub

Private Property Get INWDSMapSeries_Page(Index As Long) As NWMapBookPrj.INWDSMapPage
2498:   If Index > -1 And Index < m_PageColl.Count Then
2499:     Set INWDSMapSeries_Page = m_PageColl.Item(Index + 1)
2500:   Else
2501:     Set INWDSMapSeries_Page = Nothing
2502:   End If
End Property

Private Property Get INWDSMapSeries_PageCount() As Long
2506:   INWDSMapSeries_PageCount = m_PageColl.Count
End Property

Private Property Let INWDSMapSeries_EnableSeries(ByVal RHS As Boolean)
2510:   m_bEnableSeries = RHS
End Property

Private Property Get INWDSMapSeries_EnableSeries() As Boolean
2514:   INWDSMapSeries_EnableSeries = m_bEnableSeries
End Property

Private Sub INWDSMapSeries_RemovePage(Index As Long)
2518:   If Index > -1 And Index < m_PageColl.Count Then
2519:     m_PageColl.Remove Index + 1
2520:   End If
End Sub

Private Property Let INWDSMapSeriesOptions_ClipData(RHS As Boolean)
2524:   m_bClipData = RHS
End Property

Private Property Get INWDSMapSeriesOptions_ClipData() As Boolean
2528:   INWDSMapSeriesOptions_ClipData = m_bClipData
End Property

Private Property Let INWDSMapSeriesOptions_DataDrivenField(RHS As String)
2532:   m_sDataDrivenField = RHS
End Property

Private Property Get INWDSMapSeriesOptions_DataDrivenField() As String
2536:   INWDSMapSeriesOptions_DataDrivenField = m_sDataDrivenField
End Property

Private Property Let INWDSMapSeriesOptions_ExtentType(RHS As Long)
2540:   m_lExtentType = RHS
End Property

Private Property Get INWDSMapSeriesOptions_ExtentType() As Long
2544:   INWDSMapSeriesOptions_ExtentType = m_lExtentType
End Property

Private Property Let INWDSMapSeriesOptions_FixedScale(RHS As Double)
2548:   m_dFixedScale = RHS
End Property

Private Property Get INWDSMapSeriesOptions_FixedScale() As Double
2552:   INWDSMapSeriesOptions_FixedScale = m_dFixedScale
End Property

Private Property Let INWDSMapSeriesOptions_LabelNeighbors(RHS As Boolean)
2556:   m_bLabelNeighbors = RHS
End Property

Private Property Get INWDSMapSeriesOptions_LabelNeighbors() As Boolean
2560:   INWDSMapSeriesOptions_LabelNeighbors = m_bLabelNeighbors
End Property

Private Property Set INWDSMapSeriesOptions_LabelSymbol(RHS As ISymbol)
2564:   Set m_pLabelSymbol = RHS
End Property

Private Property Get INWDSMapSeriesOptions_LabelSymbol() As ISymbol
2568:   Set INWDSMapSeriesOptions_LabelSymbol = m_pLabelSymbol
End Property

Private Property Let INWDSMapSeriesOptions_Margin(RHS As Double)
2572:   m_dMargin = RHS
End Property

Private Property Get INWDSMapSeriesOptions_Margin() As Double
2576:   INWDSMapSeriesOptions_Margin = m_dMargin
End Property

Private Property Let INWDSMapSeriesOptions_MarginType(RHS As String)
2580:   m_sMarginType = RHS
End Property

Private Property Get INWDSMapSeriesOptions_MarginType() As String
2584:   INWDSMapSeriesOptions_MarginType = m_sMarginType
End Property

Private Property Let INWDSMapSeriesOptions_RotateFrame(RHS As Boolean)
2588:   m_bRotateFrame = RHS
End Property

Private Property Get INWDSMapSeriesOptions_RotateFrame() As Boolean
2592:   INWDSMapSeriesOptions_RotateFrame = m_bRotateFrame
End Property

Private Property Let INWDSMapSeriesOptions_RotationField(RHS As String)
2596:   m_sRotationField = RHS
End Property

Private Property Get INWDSMapSeriesOptions_RotationField() As String
2600:   INWDSMapSeriesOptions_RotationField = m_sRotationField
End Property

Private Property Let INWDSMapSeriesOptions2_ClipData(RHS As Long)
'Added 6/18/03 to support cross hatching of clip
2605:   m_lClipData2 = RHS
End Property

Private Property Get INWDSMapSeriesOptions2_ClipData() As Long
'Added 6/18/03 to support cross hatching of clip
2610:   INWDSMapSeriesOptions2_ClipData = m_lClipData2
End Property

Private Sub INWDSMapSeriesProps_AddLayerToSuppress(ByVal LayerName As String)
2614:   m_SuppressColl.Add LayerName
End Sub

Private Property Let INWDSMapSeriesProps_DataFrameName(RHS As String)
2618:   m_sDataFrameName = RHS
End Property

Private Property Get INWDSMapSeriesProps_DataFrameName() As String
2622:   INWDSMapSeriesProps_DataFrameName = m_sDataFrameName
End Property

Private Property Let INWDSMapSeriesProps_IndexFieldName(RHS As String)
2626:   m_sIndexFieldName = RHS
End Property

Private Property Get INWDSMapSeriesProps_IndexFieldName() As String
2630:   INWDSMapSeriesProps_IndexFieldName = m_sIndexFieldName
End Property

Private Property Let INWDSMapSeriesProps_IndexLayerName(RHS As String)
2634:   m_sIndexLayerName = RHS
End Property

Private Property Get INWDSMapSeriesProps_IndexLayerName() As String
2638:   INWDSMapSeriesProps_IndexLayerName = m_sIndexLayerName
End Property

Private Sub INWDSMapSeriesProps_RemoveLayerToSuppress(Index As Long)
2642:   If Index > -1 And Index < m_SuppressColl.Count Then
2643:     m_SuppressColl.Remove Index + 1
2644:   End If
End Sub

Private Property Let INWDSMapSeriesProps_StartNumber(RHS As Long)
2648:   m_lStartNumber = RHS
End Property

Private Property Get INWDSMapSeriesProps_StartNumber() As Long
2652:   INWDSMapSeriesProps_StartNumber = m_lStartNumber
End Property

Private Property Get INWDSMapSeriesProps_SuppressLayer(Index As Long) As String
2656:   If Index > -1 And Index < m_SuppressColl.Count Then
2657:     INWDSMapSeriesProps_SuppressLayer = m_SuppressColl.Item(Index + 1)
2658:   Else
2659:     INWDSMapSeriesProps_SuppressLayer = ""
2660:   End If
End Property

Private Property Get INWDSMapSeriesProps_SuppressLayerCount() As Long
2664:   INWDSMapSeriesProps_SuppressLayerCount = m_SuppressColl.Count
End Property

Private Property Let INWDSMapSeriesProps_SuppressLayers(ByVal RHS As Boolean)
2668:   m_bSupressLayers = RHS
End Property

Private Property Get INWDSMapSeriesProps_SuppressLayers() As Boolean
2672:   INWDSMapSeriesProps_SuppressLayers = m_bSupressLayers
End Property

Private Property Let INWDSMapSeriesProps_TileSelectionMethod(RHS As Long)
2676:   m_lTileSelection = RHS
End Property

Private Property Get INWDSMapSeriesProps_TileSelectionMethod() As Long
2680:   INWDSMapSeriesProps_TileSelectionMethod = m_lTileSelection
End Property

Private Property Get IPersistVariant_ID() As esriSystem.IUID
  Dim id As New UID
2685:   id = "NWMapBookPrj.NWMapSeries"
2686:   Set IPersistVariant_ID = id
End Property

Private Property Let INWMapSeriesOptions_NeighborLabelIndent(RHS As Double)
2690:   m_dLabelIndent = RHS
End Property

Private Property Get INWMapSeriesOptions_NeighborLabelIndent() As Double
2694:   INWMapSeriesOptions_NeighborLabelIndent = m_dLabelIndent
End Property

Private Sub IPersistVariant_Load(ByVal Stream As esriSystem.IVariantStream)
'Load the persisted parameters of the renderer
On Error GoTo ErrHand:
  Dim lLoop As Long, lCount As Long, pPage As INWDSMapPage, vClip As Variant
  Dim bClip As Boolean, vCount As Variant, lPropCount As Long, vNWData As Variant
  Dim bFoundUID As Boolean, bExitFlagFound As Boolean
  Dim sUID1 As String, sUID2 As String, lSymbolCount As Long, sTextSymbKey As String
  Dim pTextSymb As ITextSymbol, lLyrGroupCount As Long, sLyrGroupName As String
  Dim pNWLyrGroup As INWLayerVisibilityGroup
  Dim lDetailInsetCount As Long, sInsetName As String, pElement As esriCarto.IElement
  Dim sDataFrameName As String, pMapFrame As IMapFrame, lDataFrameCount As Long
  Dim pScrDict As Scripting.Dictionary, i As Long, lPageCount As Long
  Dim sPageName As String
  Dim sLayerName As String, lDynDefQueryFrameCount As Long, lLoop2 As Long
  Dim vDynDefQueryFrameKeys As Variant, lLayerCount As Long
  Dim sAttribName As String
  
  
  'Added 2/18/04 to make the list of persisted properties more dynamic
2716:   vCount = Stream.Read
2717:   If UCase(TypeName(vCount)) = "BOOLEAN" Then  'Version created before the 2/18/04 update.
2718:     m_bEnableSeries = CBool(vCount)
    'Added 6/18/03 to support cross hatching of area outside clip
2720:     vClip = Stream.Read
2721:     If UCase(TypeName(vClip)) = "BOOLEAN" Then
2722:       bClip = CBool(vClip)
2723:       If bClip Then
2724:         m_lClipData2 = 1
2725:       Else
2726:         m_lClipData2 = 0
2727:       End If
2728:     Else
2729:       m_lClipData2 = CLng(vClip)
2730:     End If
2731:     lPropCount = 14
2732:   Else
2733:     lPropCount = CLng(Mid(CStr(vCount), 21)) - 2
2734:     m_bEnableSeries = Stream.Read
2735:     m_lClipData2 = Stream.Read
2736:   End If

  'Original set of properties that every persisted map book will have.
2739:   m_sDataDrivenField = Stream.Read
2740:   m_lExtentType = Stream.Read
2741:   m_dFixedScale = Stream.Read
2742:   m_bLabelNeighbors = Stream.Read
2743:   Set m_pLabelSymbol = Stream.Read
2744:   m_dMargin = Stream.Read
2745:   m_sMarginType = Stream.Read
2746:   m_bRotateFrame = Stream.Read
2747:   m_sRotationField = Stream.Read
2748:   m_sDataFrameName = Stream.Read
2749:   m_sIndexFieldName = Stream.Read
2750:   m_sIndexLayerName = Stream.Read
2751:   m_bSupressLayers = Stream.Read
2752:   m_lTileSelection = Stream.Read
  
  'Additional properties added after 2/18/04
2755:   If lPropCount > 14 Then    'Checking for start number property
2756:     m_lStartNumber = Stream.Read
2757:   Else
2758:     m_lStartNumber = 1
2759:   End If
  
  'More original properties.  Writen out below the new properties because they are of variable length
2762:   lCount = Stream.Read
2763:   If lCount > 0 Then
2764:     For lLoop = 1 To lCount
2765:       m_SuppressColl.Add Stream.Read
2766:     Next lLoop
2767:   End If
  
2769:   lCount = Stream.Read
2770:   If lCount > 0 Then
2771:     For lLoop = 1 To lCount
2772:       Set pPage = Stream.Read
2773:       If lPropCount <= 14 Then   'MapBook created before page numbers were added.  In this case we manually assign the page numbers.
2774:         pPage.PageNumber = lLoop
2775:       End If
2776:       m_PageColl.Add pPage, pPage.PageName 'added the ", pPage.PageName" to support map refresh event handling [RAH, 10-18-05]
2777:     Next lLoop
2778:   End If
      
  'point of custom load code for NW
  'Mapbook customization properties.
  '{92091FB8-FC34-40fe-ADBD-C00C2E39022A}
  'The pattern is to have a nested persistence
  'structure of:
  '{guid 1}
  '  {guid 2}
  '  ... data ...
  '  {guid 2}
  '{guid 1}
  'This opens the door for any future amount of
  'data.  This code will loop through stream.read
  'until the closing GUID is detected, allowing
  'earlier versions of this code to work with
  'future files that have more data stored.
  
2796:   bExitFlagFound = False
2797:   vNWData = Stream.Read
2798:   If Not IsEmpty(vNWData) Then
                                                  'for whatever reason, saved UIDs
                                                  'persist as strings
2801:     If StrComp(UCase(TypeName(vNWData)), "String", vbTextCompare) = 0 Then
2802:       sUID1 = vNWData
                                                  'if the GUID value is correct
2804:       If StrComp(sUID1, m_pUID_FlexiblePersist.Value, vbTextCompare) = 0 Then
                                                  'then read the persisted data values
2806:         bFoundUID = False
        
        '{guid 1}
        '  {guid 2} <--- search for this section
        '  ... data ...
        '  {guid 2}
        '{guid 1}
2813:         Do While (Not bFoundUID) And (Not bExitFlagFound)
2814:           vNWData = Stream.Read
2815:           If StrComp(UCase(TypeName(vNWData)), "String", vbTextCompare) = 0 Then
2816:             sUID2 = vNWData
2817:             If StrComp(sUID2, m_pUID_NWData1.Value, vbTextCompare) = 0 Then
2818:               bFoundUID = True
2819:             ElseIf StrComp(sUID2, m_pUID_FlexiblePersist.Value, vbTextCompare) = 0 Then
2820:               bExitFlagFound = True
2821:             End If
2822:           End If
2823:         Loop
                                                  'once the bracketing UIDs
                                                  'have been detected, read the
                                                  'data for this version of the
                                                  'NW Mapbook application
2828:         If bFoundUID And Not bExitFlagFound Then
2829:           m_dLabelIndent = Stream.Read
2830:           lSymbolCount = Stream.Read                     'count of how many text symbols exist
2831:           m_pTextSymbolDict.RemoveAll
2832:           For lLoop = 0 To (lSymbolCount - 1)
2833:             vNWData = Stream.Read
2834:             sTextSymbKey = vNWData
2835:             Set pTextSymb = Stream.Read
2836:             m_pTextSymbolDict.Add sTextSymbKey, pTextSymb
2837:           Next lLoop
2838:           m_sDefaultTextSymbol = Stream.Read
        
          'load layer visibility groups
2841:           m_pLyrInvisGroups.RemoveAll
2842:           lLyrGroupCount = Stream.Read
2843:           For lLoop = 0 To (lLyrGroupCount - 1)
2844:             vNWData = Stream.Read
2845:             sLyrGroupName = vNWData
2846:             Set pNWLyrGroup = Stream.Read
2847:             m_pLyrInvisGroups.Add sLyrGroupName, pNWLyrGroup
2848:           Next lLoop
        
          'load bubble layer
2851:           m_sBubbleLayer = Stream.Read
        
          'load settings for enhanced refresh event handler
2854:           m_bHandleMapRefreshEvents = Stream.Read
          
          'load detail inset graphics
2857:           m_pDetailInsets.RemoveAll
2858:           m_pDetailInsetArrowShadows.RemoveAll
2859:           m_pDetailInsetArrows.RemoveAll
        
          'load data frame visibility
                'structure for associating data frames with map pages
2863:           m_pDictDataFramePages.RemoveAll
2864:           lDataFrameCount = Stream.Read
2865:           For lLoop = 0 To (lDataFrameCount - 1)
2866:             vNWData = Stream.Read
2867:             sDataFrameName = vNWData
2868:             Set pScrDict = New Scripting.Dictionary
2869:             lPageCount = Stream.Read
2870:             For i = 0 To (lPageCount - 1)
2871:               vNWData = Stream.Read
2872:               sPageName = vNWData
2873:               pScrDict.Add sPageName, Nothing
2874:             Next i
2875:             m_pDictDataFramePages.Add sDataFrameName, pScrDict
2876:           Next lLoop
                'structure for holding data frames not currently visible
2878:           m_pDictStoredDataFrames.RemoveAll
2879:           lDataFrameCount = Stream.Read
2880:           For lLoop = 0 To (lDataFrameCount - 1)
2881:             vNWData = Stream.Read
2882:             sDataFrameName = vNWData
2883:             Set pElement = Stream.Read
2884:             m_pDictStoredDataFrames.Add sDataFrameName, pElement
2885:           Next lLoop
2886:           m_sMainDataFrame = Stream.Read  'data frame used for map book
        
          
          'load the list of data frames to update when loading map pages
2890:           m_pDictDataFramesToUpdate_Layer.RemoveAll
2891:           m_pDictDataFramesToUpdate_Field.RemoveAll
2892:           lDataFrameCount = Stream.Read
2893:           For lLoop = 0 To (lDataFrameCount - 1)
2894:             vNWData = Stream.Read
2895:             sDataFrameName = vNWData
2896:             vNWData = Stream.Read
2897:             sLayerName = vNWData
2898:             vNWData = Stream.Read
2899:             sAttribName = vNWData
            
2901:             m_pDictDataFramesToUpdate_Layer.Add sDataFrameName, sLayerName
2902:             m_pDictDataFramesToUpdate_Field.Add sDataFrameName, sAttribName
2903:           Next lLoop
          
          'dynamic definition queries
2906:           vNWData = Stream.Read
2907:           m_sPrevPageName = vNWData
2908:           vNWData = Stream.Read
2909:           m_sCurrentPageName = vNWData
2910:           m_pDictDynDefQueryPages.RemoveAll
2911:           lDynDefQueryFrameCount = Stream.Read
2912:           For lLoop = 0 To (lDynDefQueryFrameCount - 1)
2913:             vNWData = Stream.Read
2914:             sDataFrameName = vNWData
2915:             Set pScrDict = New Scripting.Dictionary
2916:             m_pDictDynDefQueryPages.Add sDataFrameName, pScrDict
2917:             vNWData = Stream.Read
2918:             lLayerCount = vNWData
2919:             For lLoop2 = 0 To (lLayerCount - 1)
2920:               vNWData = Stream.Read
2921:               sLayerName = vNWData
2922:               vNWData = Stream.Read
2923:               sAttribName = vNWData   'should be "" in this version
2924:               pScrDict.Add sLayerName, sAttribName
2925:             Next lLoop2
2926:           Next lLoop
        
        
          'load layout element visibility
          Dim lElemCount As Long, vElemKeys As Variant, sElemName As String
          Dim sPlaceHolder As String
                                                  'flush data structures to be loaded
2933:           lElemCount = m_pDictPagesWhereElementIsInvisible.Count
2934:           vElemKeys = m_pDictPagesWhereElementIsInvisible.Keys
2935:           For i = 0 To (lElemCount - 1)
2936:             sElemName = vElemKeys(i)
2937:             Set pScrDict = m_pDictPagesWhereElementIsInvisible.Item(sElemName)
2938:             pScrDict.RemoveAll
2939:           Next i
2940:           m_pDictPagesWhereElementIsInvisible.RemoveAll
2941:           lElemCount = Stream.Read 'm_pDictPagesWhereElementIsInvisible.Count
2942:           sPlaceHolder = " "
2943:           For lLoop = 0 To (lElemCount - 1) '
2944:             vNWData = Stream.Read 'element name
2945:             sElemName = vNWData
2946:             Set pScrDict = New Scripting.Dictionary
2947:             lPageCount = Stream.Read 'pScrDict.Count
2948:             For i = 0 To (lPageCount - 1)
2949:               vNWData = Stream.Read 'page name
2950:               sPageName = vNWData
2951:               pScrDict.Add sPageName, sPlaceHolder
2952:             Next i
2953:             m_pDictPagesWhereElementIsInvisible.Add sElemName, pScrDict
2954:           Next lLoop

2956:           m_pDictStoredElements.RemoveAll
2957:           lElemCount = Stream.Read 'm_pDictStoredElements.Count
2958:           For lLoop = 0 To (lElemCount - 1)
2959:             vNWData = Stream.Read 'Element name
2960:             sElemName = vNWData
2961:             Set pElement = Stream.Read
2962:             m_pDictStoredElements.Add sElemName, pElement
2963:           Next lLoop
          
2965:           m_pDictTaggedElements.RemoveAll
2966:           lElemCount = Stream.Read 'm_pDictTaggedElements.Count
2967:           For lLoop = 0 To (lElemCount - 1)
2968:             vNWData = Stream.Read 'Element name
2969:             sElemName = vNWData
2970:             Set pElement = Stream.Read
2971:             m_pDictTaggedElements.Add sElemName, pElement
2972:           Next lLoop
          'end of layout element visibility load code
          
2975:         End If
        
        
        '{guid 1}
        '  {guid 2}
        '  ... data ...
        '  {guid 2}      <--- search for this section
        '{guid 1}
2983:         bFoundUID = False
2984:         Do While (Not bFoundUID) And (Not bExitFlagFound)
2985:           vNWData = Stream.Read
2986:           If StrComp(UCase(TypeName(vNWData)), "String", vbTextCompare) = 0 Then
2987:             sUID2 = vNWData
2988:             If StrComp(sUID2, m_pUID_NWData1.Value, vbTextCompare) = 0 Then
2989:               bFoundUID = True
2990:             ElseIf StrComp(sUID2, m_pUID_FlexiblePersist.Value, vbTextCompare) = 0 Then
2991:               bExitFlagFound = True
2992:             End If
2993:           End If
2994:         Loop
        
        '{guid 1}
        '  {guid 2}
        '  ... data ...
        '  {guid 2}
        '{guid 1}     <---- search for the closing section
3001:         Do While Not bExitFlagFound
3002:           vNWData = Stream.Read
3003:           If StrComp(UCase(TypeName(vNWData)), "String", vbTextCompare) = 0 Then
3004:             sUID2 = vNWData
3005:             If StrComp(sUID2, m_pUID_FlexiblePersist.Value, vbTextCompare) = 0 Then
3006:               bExitFlagFound = True
3007:             End If
3008:           End If
3009:         Loop
        'The above structure is made so that future data can be saved ...
        '{guid 1}
        '  <-- future data, presumably enclosed within {guid3}..{guid3}, {guid4}.. etc. -->
        '  {guid 2}
        '    ... data ...
        '    <-- future data for the NW Mapbook component since stream reads keep going until {guid2} -->
        '  {guid 2}
        '  <-- future data, presumably enclosed within {guid5}..{guid5}, {guid6}.. etc. -->
        '{guid 1}     <---- search for the closing section
      
      
3021:       End If
3022:     Else
      'not sure if this will work:
      'Will too many stream.reads give me access to
      'whatever is loaded next?  Or will it give me
      'a stream.read error?  or will stream.read
      'return nothing?
      
          'assume that I read something from the next
          'section of load code.
3031:       Stream.Write vNWData
3032:     End If
3033:   End If
    
    
    
    
    
  Exit Sub
ErrHand:
3041:   MsgBox "MapSeries - IPersistVariant_Load - " & Erl & " - " & Err.Description
End Sub

Private Sub IPersistVariant_Save(ByVal Stream As esriSystem.IVariantStream)
'Write it all out
On Error GoTo ErrHand:
  Dim lLoop As Long, vKeys() As Variant, vTextSymbols() As Variant
  Dim lSymbolCount As Long, sSymbName As String, lLyrGroupCount As Long
  Dim sLyrGroupName As String, pLyrGroup As INWLayerVisibilityGroup
  Dim lDetailInsetCount As Long, pElement As IElement, sInsetName As String
  
  
  'Added 2/18/04 to make the list of persisted properties more dynamic
3054:   Stream.Write "SERIESPROPERTYCOUNT-17"
    
3056:   Stream.Write m_bEnableSeries
  
  'Added 6/18/03 to support cross hatching of area outside the clip
3059:   Stream.Write m_lClipData2
  
3061:   Stream.Write m_sDataDrivenField
3062:   Stream.Write m_lExtentType
3063:   Stream.Write m_dFixedScale
3064:   Stream.Write m_bLabelNeighbors
3065:   Stream.Write m_pLabelSymbol
3066:   Stream.Write m_dMargin
3067:   Stream.Write m_sMarginType
3068:   Stream.Write m_bRotateFrame
3069:   Stream.Write m_sRotationField
3070:   Stream.Write m_sDataFrameName
3071:   Stream.Write m_sIndexFieldName
3072:   Stream.Write m_sIndexLayerName
3073:   Stream.Write m_bSupressLayers
3074:   Stream.Write m_lTileSelection
3075:   Stream.Write m_lStartNumber   'Added 2/18/04

On Error GoTo ErrHand2:
3078:   Stream.Write m_SuppressColl.Count
3079:   If m_SuppressColl.Count > 0 Then
3080:     For lLoop = 1 To m_SuppressColl.Count
3081:       Stream.Write m_SuppressColl.Item(lLoop)
3082:     Next lLoop
3083:   End If
3084:   Stream.Write m_PageColl.Count
3085:   If m_PageColl.Count > 0 Then
3086:     For lLoop = 1 To m_PageColl.Count
3087:       Stream.Write m_PageColl.Item(lLoop)
3088:     Next lLoop
3089:   End If
    
  'NW Mapbook point for inserting custom Saves code
  'Added 8/31/04
3093:   If m_pUID_FlexiblePersist Is Nothing Then
3094:     Set m_pUID_FlexiblePersist = New UID
3095:     m_pUID_FlexiblePersist.Value = "{92091FB8-FC34-40fe-ADBD-C00C2E39022A}"
3096:   End If
3097:   If m_pUID_NWData1 Is Nothing Then
3098:     Set m_pUID_NWData1 = New UID
3099:     m_pUID_NWData1.Value = "{8D818E86-5C50-4c1a-BBC3-6EF7A5CD5275}"
3100:   End If
3101:   Stream.Write m_pUID_FlexiblePersist.Value
3102:   Stream.Write m_pUID_NWData1.Value
3103:   Stream.Write m_dLabelIndent
  
  '''''''save symbols for adjacent map labels
3106:   lSymbolCount = m_pTextSymbolDict.Count
3107:   Stream.Write lSymbolCount
3108:   If lSymbolCount > 0 Then
3109:     vKeys = m_pTextSymbolDict.Keys
3110:     For lLoop = 0 To (lSymbolCount - 1)
3111:       sSymbName = vKeys(lLoop)
3112:       Stream.Write sSymbName
3113:       Stream.Write m_pTextSymbolDict.Item(sSymbName)
3114:     Next lLoop
3115:     Stream.Write m_sDefaultTextSymbol
3116:   End If
  
  ''''''''save layer visibility groups
3119:   lLyrGroupCount = m_pLyrInvisGroups.Count
3120:   Stream.Write lLyrGroupCount
3121:   If lLyrGroupCount > 0 Then
3122:     vKeys = m_pLyrInvisGroups.Keys
3123:     For lLoop = 0 To (lLyrGroupCount - 1)
3124:       sLyrGroupName = vKeys(lLoop)
3125:       Stream.Write sLyrGroupName
3126:       Set pLyrGroup = m_pLyrInvisGroups.Item(sLyrGroupName)
3127:       Stream.Write pLyrGroup
3128:     Next lLoop
3129:   End If
  
  'save bubble layer
3132:   Stream.Write m_sBubbleLayer
  
  'Remove detail inset graphics --
  Dim sInsetID As String, pGraphicsContainer As IGraphicsContainer, i As Long
  Dim vDetailInsetKeys As Variant, pActiveView As IActiveView
3137:   If Not m_pMxDoc Is Nothing Then
3138:     lDetailInsetCount = m_pDetailInsets.Count
3139:     vDetailInsetKeys = m_pDetailInsets.Keys
3140:     Set pGraphicsContainer = m_pMxDoc.PageLayout
    
3142:     For i = 0 To lDetailInsetCount - 1
3143:       sInsetID = vDetailInsetKeys(i)
3144:       Set pElement = m_pDetailInsetArrows.Item(sInsetID)
3145:       pGraphicsContainer.DeleteElement pElement
      
3147:       Set pElement = m_pDetailInsetArrowShadows.Item(sInsetID)
3148:       pGraphicsContainer.DeleteElement pElement
3149:       Set pElement = m_pDetailInsets.Item(sInsetID)
3150:       pGraphicsContainer.DeleteElement pElement
    
3152:       m_pDetailInsetArrows.Remove sInsetID
3153:       m_pDetailInsetArrowShadows.Remove sInsetID
3154:       m_pDetailInsets.Remove sInsetID
3155:     Next i
    'Refresh ActiveView and TOC
3157:     Set pActiveView = m_pMxDoc.FocusMap
3158:     pActiveView.Refresh
3159:     m_pMxDoc.CurrentContentsView.Refresh 0
  'Else
      'pMxDoc reference is only set within code to handle detail insets.
3162:   End If
    
  
  
  'save settings for enhanced refresh event handler
3167:   Stream.Write m_bHandleMapRefreshEvents
  
    
    
  Dim lDataFrameCount As Long, vDataFrameKeys As Variant, pScrDict As Scripting.Dictionary
  Dim sDataFrameName As String, vPageKeys As Variant, lPageCount As Long
  Dim j As Long, sPageName As String
    
    
  ''''''''save data frame visibility
3177:   lDataFrameCount = m_pDictDataFramePages.Count
3178:   vDataFrameKeys = m_pDictDataFramePages.Keys
                                                  'data structure tracking associations
                                                  'of data frames to map pages
3181:   Stream.Write lDataFrameCount
    
3183:   For i = 0 To (lDataFrameCount - 1)
3184:     sDataFrameName = vDataFrameKeys(i)
3185:     Set pScrDict = m_pDictDataFramePages.Item(sDataFrameName)
3186:     lPageCount = pScrDict.Count
3187:     vPageKeys = pScrDict.Keys
      
3189:     Stream.Write sDataFrameName
3190:     Stream.Write lPageCount
      
    
3193:     For j = 0 To (lPageCount - 1)
3194:       sPageName = vPageKeys(j)
3195:       Stream.Write sPageName
3196:     Next j
3197:   Next i
3198:   lDataFrameCount = m_pDictStoredDataFrames.Count
3199:   vDataFrameKeys = m_pDictStoredDataFrames.Keys
                                                  'data structure for storing data frames
                                                  'not currently visible, but that should
                                                  'be available for other map pages.
    
3204:   Stream.Write lDataFrameCount
3205:   For i = 0 To (lDataFrameCount - 1)
3206:     sDataFrameName = vDataFrameKeys(i)
3207:     Stream.Write sDataFrameName
3208:     Set pElement = m_pDictStoredDataFrames.Item(sDataFrameName)
3209:     Stream.Write pElement
3210:   Next i
3211:   Stream.Write m_sMainDataFrame
  
  
  
  
  
  '''''''save info on data frames that have their extent updated when drawing pages
  Dim sFieldName As String, sLayerName As String
  
3220:   lDataFrameCount = m_pDictDataFramesToUpdate_Layer.Count
3221:   vDataFrameKeys = m_pDictDataFramesToUpdate_Layer.Keys
  
3223:   Stream.Write lDataFrameCount
  
3225:   For i = 0 To (lDataFrameCount - 1)
3226:     sDataFrameName = vDataFrameKeys(i)
3227:     sLayerName = m_pDictDataFramesToUpdate_Layer(sDataFrameName)
3228:     sFieldName = m_pDictDataFramesToUpdate_Field(sDataFrameName)
    
3230:     Stream.Write sDataFrameName
3231:     Stream.Write sLayerName
3232:     Stream.Write sFieldName
3233:   Next i
  
  
  
  
  '''''''save dynamic definition query information
  Dim lDynDefQueryFrameCount As Long, vDynDefQueryFrameKeys As Variant
  Dim vLayerKeys As Variant, sAttribName As String
  Dim lLayerCount As Long
  
3243:   Stream.Write m_sPrevPageName
3244:   Stream.Write m_sCurrentPageName
3245:   lDynDefQueryFrameCount = m_pDictDynDefQueryPages.Count
3246:   vDynDefQueryFrameKeys = m_pDictDynDefQueryPages.Keys

3248:   Stream.Write lDynDefQueryFrameCount
3249:   For i = 0 To (lDynDefQueryFrameCount - 1)
3250:     sDataFrameName = vDynDefQueryFrameKeys(i)
3251:     Stream.Write sDataFrameName
3252:     Set pScrDict = m_pDictDynDefQueryPages.Item(sDataFrameName)
3253:     lLayerCount = pScrDict.Count
3254:     vLayerKeys = pScrDict.Keys

3256:     Stream.Write lLayerCount
3257:     For j = 0 To (lLayerCount - 1)
3258:       sLayerName = vLayerKeys(j)
3259:       Stream.Write sLayerName
3260:       sAttribName = pScrDict.Item(sLayerName)
3261:       Stream.Write sAttribName 'should be "" in this version
3262:     Next j
3263:   Next i
          
  
  
  
  ''''''''save layout element visibility information
  Dim lElemCount As Long, vElemKeys As Variant, sElemName As String
                                                  'm_pDictPagesWhereElementIsInvisible
3271:   lElemCount = m_pDictPagesWhereElementIsInvisible.Count
3272:   vElemKeys = m_pDictPagesWhereElementIsInvisible.Keys
  
3274:   Stream.Write lElemCount
3275:   For i = 0 To (lElemCount - 1)
3276:     sElemName = vElemKeys(i)
3277:     Stream.Write sElemName
3278:     Set pScrDict = m_pDictPagesWhereElementIsInvisible.Item(sElemName)
3279:     lPageCount = pScrDict.Count
3280:     vPageKeys = pScrDict.Keys
    
3282:     Stream.Write lPageCount
3283:     For j = 0 To (lPageCount - 1)
3284:       sPageName = vPageKeys(j)
3285:       Stream.Write sPageName
3286:     Next j
3287:   Next i
                                                  'm_pDictStoredElements
3289:   lElemCount = m_pDictStoredElements.Count
3290:   vElemKeys = m_pDictStoredElements.Keys
  
3292:   Stream.Write lElemCount
3293:   For i = 0 To (lElemCount - 1)
3294:     sElemName = vElemKeys(i)
3295:     Stream.Write sElemName
3296:     Set pElement = m_pDictStoredElements.Item(sElemName)
3297:     Stream.Write pElement
3298:   Next i
                                                  'm_pDictTaggedElements
3300:   lElemCount = m_pDictTaggedElements.Count
3301:   vElemKeys = m_pDictTaggedElements.Keys
  
3303:   Stream.Write lElemCount
3304:   For i = 0 To (lElemCount - 1)
3305:     sElemName = vElemKeys(i)
3306:     Stream.Write sElemName
3307:     Set pElement = m_pDictTaggedElements(sElemName)
3308:     Stream.Write pElement
3309:   Next i
  
  
  
3313:   Stream.Write m_pUID_NWData1.Value
3314:   Stream.Write m_pUID_FlexiblePersist.Value
  'End of NW Mapbook custom Save code
  
  Exit Sub
ErrHand:
3319:   MsgBox "MapSeries - IPersistVariant_Save - " & Erl & ": " & Err.Description
  'MsgBox "MapSeries - IPersistVariant_Save - " & Err.Description
  Exit Sub
ErrHand2:
3323:   MsgBox "MapSeriesCOLLS - IPersistVariant_Save - " & Erl & Err.Description
End Sub














'AddToArcMap
'
'  This routine is called by RefreshDetailInsets to add a new detail
'  inset to the ArcMap layout view.  This routine is called once per
'  bubble polygon, and is passed information about that polygon.
'--------------------------------------------------------
Private Sub AddToArcMap(pElementWithinBorders As IElement, _
                                   pMxDoc As esriArcMapUI.IMxDocument, _
                                   pRow As esriGeoDatabase.IRow, _
                                   pDictDetailInsets As Scripting.Dictionary, _
                                   pDictArrows As Scripting.Dictionary, _
                                   pDictArrowShadows As Scripting.Dictionary)
  On Error GoTo ErrorHandler

3353:   If m_pMxDoc Is Nothing Then
3354:     Set m_pMxDoc = pMxDoc
3355:   End If

  Dim pFields As IFields, lFieldCount As Long, pField As IField
  Dim pNewMap As IMap, pMapFrame As IMapFrame, pEnv As IEnvelope
  Dim pGraphicsContainer As IGraphicsContainer, pElement As IElement
  Dim pFrameElement As IFrameElement, pNewMapAV As IActiveView
  Dim pNewEnv As IEnvelope, pGraphicsContainerSelect As IGraphicsContainerSelect

  Dim pMxDocDataViewAV As IActiveView, pScrDisplay As IScreenDisplay
  Dim pMxDocLayoutViewAV As IActiveView

  Dim lBubbleId As Long, dXOrigin As Double, dYOrigin As Double
  Dim dXDestination As Double, dYDestination As Double, dRadius As Double
  Dim dScale As Double, sLayers As String, dWidthOrigin As Double
  Dim bIsCircular As Boolean, i As Long, dPageToMapUnitRatio As Double
  Dim dMapScale As Double, dDataFrameWidth As Double, dDataFrameHeight As Double
  Dim sActiveDataFrameName As String, pMxDocFocusMapQIDataViewAV As IActiveView
  Dim pMxDocFocusMapQILayoutViewAV As IActiveView, pMainMapFrame As IMapFrame

  If pRow Is Nothing Then Exit Sub
3375:   If pMxDoc Is Nothing Then
    Exit Sub
3377:   End If

3379:   Set pFields = pRow.Fields
3380:   lFieldCount = pFields.FieldCount

3382:   For i = 0 To lFieldCount - 1
3383:     Set pField = pFields.Field(i)
    Select Case pField.Name
    Case "BUBBLEID"
3386:       lBubbleId = pRow.Value(i)
    Case "XORG"
3388:       dXOrigin = pRow.Value(i)
    Case "YORG"
3390:       dYOrigin = pRow.Value(i)
    Case "XDEST"
3392:       dXDestination = pRow.Value(i)
    Case "YDEST"
3394:       dYDestination = pRow.Value(i)
    Case "RADIUS"
3396:       dRadius = pRow.Value(i)
    Case "SCALE"
3398:       dScale = pRow.Value(i)
    Case "LAYERS"
3400:       If IsNull(pRow.Value(i)) Then
3401:         sLayers = ""
3402:       Else
3403:         sLayers = pRow.Value(i)
3404:       End If
    Case "WIDTHORG"
3406:       If IsNull(pRow.Value(i)) Or (pRow.Value(i) = 0) Then
3407:         If dRadius > 0 Then
3408:           dWidthOrigin = dRadius * 2
3409:         Else
3410:           dWidthOrigin = 10
3411:         End If
3412:       Else
3413:         dWidthOrigin = pRow.Value(i)
3414:       End If
3415:     End Select
3416:   Next i

3418:   Set pNewMap = New Map
3419:   pNewMap.Name = GBL_BUBBLE_IDENTIFIER & lBubbleId
3420:   pNewMap.Description = "Detail Inset " & lBubbleId


3423:   If pDictDetailInsets.Exists(pNewMap.Name) Then
3424:     MsgBox "Warning: Map book application attempted to create a detail " & vbNewLine _
         & "inset that should already exist called " & pNewMap.Name & vbNewLine _
         & vbNewLine _
         & "Duplicate detail inset will not be created.  Please check for the " & vbNewLine _
         & "presence of a non-unique name for this bubble detail polygon." & vbNewLine
    Exit Sub
3430:   End If

  ' clone all copies of layers so that modifying
  ' layers in one dataframe doesn't impact layers
  ' referenced in the other data frames.
  ''''''''''''''''''''''
  Dim pLayer As ILayer, pLayerSrc As ILayer
  Dim pFeatLyr As IFeatureLayer, pTinLyr As ITinLayer, pRastLyrSrc As IRasterLayer
  Dim pRastLyr As IRasterLayer, pFeatLyrSrc As IFeatureLayer, pTinLyrSrc As ITinLayer
  Dim pGeoFeatLyrSource As IGeoFeatureLayer, pGeoFeatLyrDestination As IGeoFeatureLayer

3441:   If pMxDoc.FocusMap.LayerCount > 0 Then
3442:     For i = (pMxDoc.FocusMap.LayerCount - 1) To 0 Step -1
3443:       Set pLayerSrc = pMxDoc.FocusMap.Layer(i)
3444:       If TypeOf pMxDoc.FocusMap.Layer(i) Is IFeatureLayer Then
3445:         Set pFeatLyr = New FeatureLayer
3446:         Set pFeatLyrSrc = pLayerSrc
3447:         pFeatLyr.DataSourceType = pFeatLyrSrc.DataSourceType
3448:         pFeatLyr.DisplayField = pFeatLyrSrc.DisplayField
3449:         Set pFeatLyr.FeatureClass = pFeatLyrSrc.FeatureClass
3450:         pFeatLyr.ScaleSymbols = pFeatLyrSrc.ScaleSymbols
3451:         pFeatLyr.Selectable = pFeatLyrSrc.Selectable
3452:         Set pGeoFeatLyrSource = pFeatLyrSrc
3453:         Set pGeoFeatLyrDestination = pFeatLyr
3454:         With pGeoFeatLyrSource
                                            'Set pGeoFeatLyrDestination.CurrentMapLevel = .CurrentMapLevel
                                            'pGeoFeatLyrDestination.DisplayFeatureClass = .DisplayFeatureClass
3457:           pGeoFeatLyrDestination.AnnotationProperties = .AnnotationProperties
3458:           pGeoFeatLyrDestination.AnnotationPropertiesID = .AnnotationPropertiesID
3459:           pGeoFeatLyrDestination.DisplayAnnotation = .DisplayAnnotation
3460:           Set pGeoFeatLyrDestination.ExclusionSet = .ExclusionSet
3461:           Set pGeoFeatLyrDestination.Renderer = .Renderer
3462:         End With
3463:         Set pLayer = pFeatLyr

3465:       ElseIf TypeOf pMxDoc.FocusMap.Layer(i) Is ITinLayer Then
                                            'pTinLyr.RendererCount = pTinLyrSrc.RendererCount
3467:         Set pTinLyrSrc = pLayerSrc
3468:         Set pTinLyr = New TinLayer
3469:         Set pTinLyr.Dataset = pTinLyrSrc.Dataset
3470:         pTinLyr.DisplayField = pTinLyrSrc.DisplayField
3471:         pTinLyr.ScaleSymbols = pTinLyrSrc.ScaleSymbols
3472:         Set pLayer = pTinLyr
3473:       ElseIf TypeOf pMxDoc.FocusMap.Layer(i) Is IRasterLayer Then
                                            'pRastLyr.BandCount = pRastLyrSrc.BandCount
                                            'pRastLyr.ColumnCount = pRastLyrSrc.ColumnCount
                                            'pRastLyr.DataFrameExtent = pRastLyrSrc.DataFrameExtent
                                            'pRastLyr.FilePath = pRastLyrSrc.FilePath
                                            'pRastLyr.Raster = pRastLyrSrc.Raster
                                            'pRastLyr.RowCount = pRastLyrSrc.RowCount
3480:         Set pRastLyr = New RasterLayer
3481:         Set pRastLyrSrc = pLayerSrc
3482:         pRastLyr.DisplayResolutionFactor = pRastLyrSrc.DisplayResolutionFactor
3483:         pRastLyr.PrimaryField = pRastLyrSrc.PrimaryField
3484:         pRastLyr.PyramidPresent = pRastLyrSrc.PyramidPresent
3485:         Set pRastLyr.Renderer = pRastLyrSrc.Renderer
3486:         pRastLyr.ShowResolution = pRastLyrSrc.ShowResolution
3487:         pRastLyr.VisibleExtent = pRastLyrSrc.VisibleExtent
3488:         Set pLayer = pRastLyr
3489:       End If
                                            'Set pLayer.AreaOfInterest = pLayerSrc.AreaOfInterest
                                            'Set pLayer.SpatialReference = pLayerSrc.SpatialReference
                                            'pLayer.SupportedDrawPhases = pLayerSrc.SupportedDrawPhases
                                            'pLayer.TipText = pLayerSrc.TipText
                                            'pLayer.Valid = pLayerSrc.Valid
3495:       pLayer.Cached = pLayerSrc.Cached
3496:       pLayer.MaximumScale = pLayerSrc.MaximumScale
3497:       pLayer.MinimumScale = pLayerSrc.MinimumScale
3498:       pLayer.Name = pLayerSrc.Name
3499:       pLayer.ShowTips = pLayerSrc.ShowTips
3500:       pLayer.Visible = pLayerSrc.Visible

3502:       pNewMap.AddLayer pLayer
3503:     Next i
3504:   End If
3505:   SetVisibleLayers pNewMap, sLayers
                                                  'Create a new MapFrame and associate
                                                  'map with it
3508:   Set pMapFrame = New MapFrame
3509:   Set pMapFrame.Map = pNewMap
3510:   pMapFrame.ExtentType = esriExtentDefault

3512:   Set pGraphicsContainer = pMxDoc.PageLayout
3513:   Set pGraphicsContainerSelect = pMxDoc.PageLayout
                                                  'Set the position of the new map frame
3515:   Set pElement = pMapFrame
3516:   Set pEnv = New Envelope
3517:   Set pMxDocDataViewAV = pMxDoc.ActiveView
3518:   Set pMxDocFocusMapQIDataViewAV = pMxDoc.FocusMap

3520:   bIsCircular = ((2 * dRadius) = dWidthOrigin)

  Dim dBubbleRadius As Double, dBubbleWidth As Double
  Dim dBubbleLeft As Double, dBubbleRight As Double
  Dim dBubbleTop As Double, dBubbleBottom As Double
  Dim dScreenRightInches As Double, dScreenBottomInches As Double
  Dim lBubbleTop As Long, lBubbleBottom As Long
  Dim lBubbleLeft As Long, lBubbleRight As Long
  Dim lBubbleTopMapAV As Long, lBubbleBottomMapAV As Long
  Dim lBubbleLeftMapAV As Long, lBubbleRightMapAV As Long
  Dim pMapEnv As IEnvelope, pBubbleEnv As IEnvelope
  Dim pMap As IMap, pActiveView As IActiveView


3534:   dBubbleRadius = dRadius * dScale
3535:   dBubbleWidth = dWidthOrigin * dScale
                                                  'generate the top y
                                                  'generate the bottom y
                                                  'the left x and right x
                                                  'for detail insets
3540:   dBubbleTop = dYDestination + dBubbleRadius
3541:   dBubbleBottom = dYDestination - dBubbleRadius
3542:   If bIsCircular Then
3543:     dBubbleLeft = dXDestination - dBubbleRadius
3544:     dBubbleRight = dXDestination + dBubbleRadius
3545:   Else
3546:     dBubbleLeft = dXDestination - (dBubbleWidth / 2)
3547:     dBubbleRight = dXDestination + (dBubbleWidth / 2)
3548:   End If

3550:   Set pBubbleEnv = New Envelope
3551:   Set pMapEnv = pMxDoc.ActiveView.Extent 'sides of display in screen units



  'Shift to layout view
  '--------------------
3557:   Set pMxDoc.ActiveView = pMxDoc.PageLayout
3558:   Set pMapEnv = pMxDoc.ActiveView.Extent
3559:   Set pMxDocLayoutViewAV = pMxDoc.ActiveView
3560:   Set pMxDocFocusMapQILayoutViewAV = pMxDoc.FocusMap



  ' Size and place the detail inset
  '--------------------------------
  
  Dim dDocWidthInches As Double, dDocHeightInches As Double
  Dim pPntScrBottomLeft As IPoint, pPntScrTopRight As IPoint
  Dim dScrMapUnitWidth As Double, dScrMapUnitHeight As Double
  Dim dScrMapUnitWidthWOBorders As Double, dScrMapUnitHeightWOBorders As Double
  Dim dRatioInchToMapUnitX As Double, dRatioInchToMapUnitY As Double
  Dim dBubbleLeftInches As Double, dBubbleRightInches As Double
  Dim dBubbleTopInches As Double, dBubbleBottomInches As Double
  Dim dBoundaryInchesXLeft As Double, dBoundaryInchesYTop As Double

3576:   With pMapEnv
3577:     dDocWidthInches = .XMax - (Abs(.XMin))
3578:     dDocHeightInches = .YMax - (Abs(.YMin))
3579:   End With
                                                  'get the data area width minus the
                                                  'buffer of empty space around the
                                                  'map display
3583:   dDataFrameWidth = pElementWithinBorders.Geometry.Envelope.Width
3584:   dDataFrameHeight = pElementWithinBorders.Geometry.Envelope.Height
3585:   dBoundaryInchesXLeft = pElementWithinBorders.Geometry.Envelope.XMin
3586:   dBoundaryInchesYTop = pElementWithinBorders.Geometry.Envelope.YMin

3588:   With pMxDocFocusMapQILayoutViewAV.ScreenDisplay.DisplayTransformation
3589:     Set pPntScrBottomLeft = .ToMapPoint(.DeviceFrame.Left, .DeviceFrame.bottom)
3590:     Set pPntScrTopRight = .ToMapPoint(.DeviceFrame.Right, .DeviceFrame.Top)
3591:   End With

3593:   dScrMapUnitWidth = pPntScrTopRight.X - pPntScrBottomLeft.X
3594:   dScrMapUnitHeight = pPntScrTopRight.Y - pPntScrBottomLeft.Y
3595:   dRatioInchToMapUnitX = dDataFrameWidth / dScrMapUnitWidth
3596:   dRatioInchToMapUnitY = dDataFrameHeight / dScrMapUnitHeight
  
3598:   dBubbleLeftInches = ((dBubbleLeft - pPntScrBottomLeft.X) * dRatioInchToMapUnitX) + dBoundaryInchesXLeft
3599:   dBubbleRightInches = ((dBubbleRight - pPntScrBottomLeft.X) * dRatioInchToMapUnitX) + dBoundaryInchesXLeft
3600:   dBubbleTopInches = ((dBubbleTop - pPntScrBottomLeft.Y) * dRatioInchToMapUnitY) + dBoundaryInchesYTop
3601:   dBubbleBottomInches = ((dBubbleBottom - pPntScrBottomLeft.Y) * dRatioInchToMapUnitY) + dBoundaryInchesYTop
  
3603:   pBubbleEnv.XMin = dBubbleLeftInches
3604:   pBubbleEnv.XMax = dBubbleRightInches
3605:   pBubbleEnv.YMin = dBubbleBottomInches
3606:   pBubbleEnv.YMax = dBubbleTopInches



3610:   pElement.Geometry = pBubbleEnv

  '''''''''''''''''''''''''''''''''''''''
  ' Add code to set data frame properties



  Dim pColor As IColor, pFillColor As IColor, pShadowColor As IColor
3618:   Set pColor = New RgbColor
3619:   Set pFillColor = New RgbColor
3620:   Set pShadowColor = New RgbColor
3621:   pColor.RGB = RGB(210, 210, 210)
3622:   pFillColor.RGB = RGB(255, 255, 255)
3623:   pShadowColor.RGB = RGB(128, 128, 128)

3625:   Set pFrameElement = pElement

  ' create a border
  Dim pSymbolBorder As ISymbolBorder
  Dim pLineSymbol As ILineSymbol
  Dim pFrameDecoration As IFrameDecoration
  Dim pShadowFillSymbol As IFillSymbol
  Dim pSymbolShadow As ISymbolShadow
  Dim pFrameProperties As IFrameProperties

3635:   Set pSymbolBorder = New SymbolBorder
3636:   Set pLineSymbol = New SimpleLineSymbol
3637:   pLineSymbol.Color = pColor
3638:   pSymbolBorder.LineSymbol = pLineSymbol
3639:   pSymbolBorder.LineSymbol.Color = pColor
3640:   pSymbolBorder.CornerRounding = 100
3641:   pFrameElement.Border = pSymbolBorder

  'modify the frame element background
3644:   Set pFrameDecoration = New SymbolBackground
3645:   pFrameDecoration.Color = pFillColor
3646:   pFrameDecoration.CornerRounding = 100

3648:   pFrameElement.Background = pFrameDecoration


  ' add shadow to detail inset
3652:   Set pShadowFillSymbol = New SimpleFillSymbol
3653:   pShadowFillSymbol.Color = pShadowColor
3654:   pShadowFillSymbol.Outline.Color = pShadowColor
3655:   Set pSymbolShadow = New SymbolShadow
3656:   pSymbolShadow.FillSymbol = pShadowFillSymbol
3657:   pSymbolShadow.HorizontalSpacing = -3
3658:   pSymbolShadow.VerticalSpacing = -3
3659:   pSymbolShadow.CornerRounding = 100
3660:   Set pFrameProperties = pFrameElement
3661:   pFrameProperties.Shadow = pSymbolShadow


  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  'Build triangle and shadow triangle to point from the bubble inset to the
  '(presumably) circular polygon feature that represents the area in the inset


  Dim pLine As ILine, pFromPnt As IPoint, pToPnt As IPoint
  Dim pCircularArc As ICircularArc, pBaseLine As ILine
  Dim p1stPnt As IPoint, p2ndPnt As IPoint, p3rdPnt As IPoint
  Dim p1stShadowPnt As IPoint, p2ndShadowPnt As IPoint, p3rdShadowPnt As IPoint
  Dim pConstructPoint As IConstructPoint

3675:   Set pLine = New esriGeometry.Line
3676:   Set pFromPnt = New Point
3677:   Set pToPnt = New Point
3678:   pFromPnt.X = dXDestination
3679:   pFromPnt.Y = dYDestination
3680:   pToPnt.X = dXOrigin
3681:   pToPnt.Y = dYOrigin
3682:   pLine.PutCoords pFromPnt, pToPnt

3684:   Set pCircularArc = New CircularArc
                                                  'angles are stored in radians,
                                                  'so calculate 10 degrees in radians
3687:   pCircularArc.PutCoordsByAngle pFromPnt, _
                               (pLine.Angle - ((10 / 180) * 3.14159265358979)), _
                               ((20 / 180) * 3.14159265358979), _
                               dBubbleRadius

3692:   Set pBaseLine = New esriGeometry.Line
3693:   pBaseLine.PutCoords pCircularArc.FromPoint, pCircularArc.ToPoint
3694:   Set p1stPnt = New Point
3695:   Set p2ndPnt = New Point
3696:   Set p3rdPnt = New Point
3697:   Set p1stShadowPnt = New Point
3698:   Set p2ndShadowPnt = New Point
3699:   Set p3rdShadowPnt = New Point
3700:   Set pConstructPoint = p3rdPnt
3701:   pConstructPoint.ConstructDeflection pBaseLine, _
                                      pBaseLine.Length, _
                                      -((60 / 180) * 3.14159265358979)
                                                  '3 points are now available for triangle:
                                                  ' - pCircularArc.FromPoint
                                                  ' - pCircularArc.ToPoint,
                                                  ' - p3rdPnt
  Dim pTrianglePoly As IPolygon, pGeomColl As IGeometryCollection, pGeometry As IGeometry
  Dim pShadowTrianglePoly As IPolygon, pShadowPolygonElement As IPolygonElement
  Dim pPntColl As IPointCollection, pFeature As IFeature
  Dim pPolygonElement As IPolygonElement, pFillShapeElement As IFillShapeElement
  Dim pShadowPolyElement As IPolygonElement, pShadowFillShapeElement As IFillShapeElement
  Dim pArrowFillSymb As IFillSymbol, pElementPly As IElement, pElementShadowPly As IElement

  '''''''''''''''''''''''''''''''
  'triangle colors and dimensions

3718:   Set pArrowFillSymb = New SimpleFillSymbol
3719:   Set pPolygonElement = New PolygonElement
3720:   Set pFillShapeElement = pPolygonElement
3721:   pArrowFillSymb.Outline = pLineSymbol
3722:   pArrowFillSymb.Color = pLineSymbol.Color
3723:   pFillShapeElement.Symbol = pArrowFillSymb
3724:   Set pElementPly = pPolygonElement

3726:   Set pTrianglePoly = New esriGeometry.Polygon
3727:   Set pPntColl = pTrianglePoly
3728:   p1stPnt.X = ((pCircularArc.FromPoint.X - pPntScrBottomLeft.X) * dRatioInchToMapUnitX) + dBoundaryInchesXLeft
3729:   p1stPnt.Y = ((pCircularArc.FromPoint.Y - pPntScrBottomLeft.Y) * dRatioInchToMapUnitX) + dBoundaryInchesYTop
3730:   p2ndPnt.X = ((pCircularArc.ToPoint.X - pPntScrBottomLeft.X) * dRatioInchToMapUnitX) + dBoundaryInchesXLeft
3731:   p2ndPnt.Y = ((pCircularArc.ToPoint.Y - pPntScrBottomLeft.Y) * dRatioInchToMapUnitX) + dBoundaryInchesYTop
3732:   p3rdPnt.X = ((p3rdPnt.X - pPntScrBottomLeft.X) * dRatioInchToMapUnitX) + dBoundaryInchesXLeft
3733:   p3rdPnt.Y = ((p3rdPnt.Y - pPntScrBottomLeft.Y) * dRatioInchToMapUnitX) + dBoundaryInchesYTop
3734:   pPntColl.AddPoint p1stPnt
3735:   pPntColl.AddPoint p2ndPnt
3736:   pPntColl.AddPoint p3rdPnt
3737:   pTrianglePoly.Close

3739:   Set pGeometry = pTrianglePoly
3740:   pElementPly.Geometry = pGeometry

  ''''''''''''''''''''''''''''''''''''''
  'triangle shadow colors and dimensions

3745:   Set pShadowPolyElement = New PolygonElement
3746:   Set pShadowFillShapeElement = pShadowPolyElement
3747:   Set pFillShapeElement = pShadowPolyElement
3748:   pFillShapeElement.Symbol = pShadowFillSymbol
3749:   Set pElementShadowPly = pShadowPolyElement

3751:   Set pShadowTrianglePoly = New esriGeometry.Polygon
3752:   Set pPntColl = pShadowTrianglePoly
                                                  'offset shadow triangle by 2 pixels
3754:   p1stShadowPnt.X = p1stPnt.X - (ConvertPixelsToRW(2, pMxDoc) * dRatioInchToMapUnitX)
3755:   p1stShadowPnt.Y = p1stPnt.Y - (ConvertPixelsToRW(2, pMxDoc) * dRatioInchToMapUnitY)
3756:   p2ndShadowPnt.X = p2ndPnt.X - (ConvertPixelsToRW(2, pMxDoc) * dRatioInchToMapUnitX)
3757:   p2ndShadowPnt.Y = p2ndPnt.Y - (ConvertPixelsToRW(2, pMxDoc) * dRatioInchToMapUnitY)
3758:   p3rdShadowPnt.X = p3rdPnt.X - (ConvertPixelsToRW(2, pMxDoc) * dRatioInchToMapUnitX)
3759:   p3rdShadowPnt.Y = p3rdPnt.Y - (ConvertPixelsToRW(2, pMxDoc) * dRatioInchToMapUnitY)
  
  
'  p1stShadowPnt.X = p1stPnt.X
'  p1stShadowPnt.Y = p1stPnt.Y
'  p2ndShadowPnt.X = p2ndPnt.X
'  p2ndShadowPnt.Y = p2ndPnt.Y
'  p3rdShadowPnt.X = p3rdPnt.X
'  p3rdShadowPnt.Y = p3rdPnt.Y
3768:   pPntColl.AddPoint p1stShadowPnt
3769:   pPntColl.AddPoint p2ndShadowPnt
3770:   pPntColl.AddPoint p3rdShadowPnt
3771:   pShadowTrianglePoly.Close
3772:   Set pGeometry = pShadowTrianglePoly
3773:   pElementShadowPly.Geometry = pGeometry

  'tag the graphic elements for later tracking
  Dim pElementProps As IElementProperties
3777:   Set pElementProps = pElementShadowPly
3778:   pElementProps.CustomProperty = GBL_BUBBLE_IDENTIFIER & lBubbleId
3779:   Set pElementProps = pElement
3780:   pElementProps.CustomProperty = GBL_BUBBLE_IDENTIFIER & lBubbleId
3781:   Set pElementProps = pElementPly
3782:   pElementProps.CustomProperty = GBL_BUBBLE_IDENTIFIER & lBubbleId
  


  ''''''''''''''''''''''''''''''
  'Add triangle shadow to layout
3788:   pGraphicsContainer.AddElement pElementShadowPly, 0

  '''''''''''''''''''''''''''
  'Add mapframe to the layout
3792:   pGraphicsContainer.AddElement pElement, 0

  '''''''''''''''''''''''
  'Add triangle to layout
3796:   pGraphicsContainer.AddElement pElementPly, 0

                                                  'Created graphics are selected by default.
                                                  'Unselect them
3800:   pGraphicsContainerSelect.UnselectElement pElementShadowPly
3801:   pGraphicsContainerSelect.UnselectElement pElement
3802:   pGraphicsContainerSelect.UnselectElement pElementPly
  
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  'Track the elements so that map page changes can delete them
3806:   pDictDetailInsets.Add pNewMap.Name, pElement
3807:   pDictArrows.Add pNewMap.Name, pElementPly
3808:   pDictArrowShadows.Add pNewMap.Name, pElementShadowPly



3812:   Set pActiveView = pNewMap
3813:   Set pNewEnv = New Envelope
  
3815:   pNewEnv.XMin = dXOrigin - dRadius
3816:   pNewEnv.XMax = dXOrigin + dRadius
3817:   pNewEnv.YMin = dYOrigin - dRadius
3818:   pNewEnv.YMax = dYOrigin + dRadius
  
3820:   pActiveView.Extent = pNewEnv
3821:   pActiveView.Refresh

  'Set pGraphicsLayer = pMxDoc.FocusMap.BasicGraphicsLayer
  'Set pCompositeGraphicsLayer = pGraphicsLayer
  'Set pGraphicsLayer.AssociatedLayer = pFeatureLayer

  'Set pGraphicsLayer = pCompositeGraphicsLayer.FindLayer("NW Detail Inset Arrows")
  'If pGraphicsLayer Is Nothing Then
  '  Set pGraphicsLayer = pCompositeGraphicsLayer.AddLayer("NW Detail Inset Arrows", pFeatureLayer)
  'End If

3832:   Set pActiveView = pMxDoc.FocusMap
3833:   pActiveView.PartialRefresh esriViewGraphics, Nothing, Nothing


  Exit Sub
ErrorHandler:
  HandleError False, "AddToArcMap " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub DeleteFromArcMap(pMxDoc As esriArcMapUI.IMxDocument, pRow As esriGeoDatabase.IRow)
  Dim pFields As IFields, lFieldCount As Long, pField As IField
  Dim lBubbleId As Long, pOldMap As IMap, lMapCount As Long
  Dim i As Long, pCurrMap As IMap

  If pRow Is Nothing Then Exit Sub
3847:   If pMxDoc Is Nothing Then
    Exit Sub
3849:   End If

  '----------------------------------------
  'grab the detail inset name from the pRow

3854:   Set pFields = pRow.Fields
3855:   lFieldCount = pFields.FieldCount

3857:   For i = 0 To lFieldCount - 1
3858:     Set pField = pFields.Field(i)
    Select Case pField.Name
    Case "BUBBLEID"
3861:       lBubbleId = pRow.Value(i)
3862:       i = lFieldCount - 1
3863:     End Select
3864:   Next i

  '------------------------------------
  'remove graphic elements from the map

  'arrow shadow element
  'arrow element
  'map element -- redundant to code removing map?

  '------------------------------------
  'Delete the detail inset's data frame

3876:   lMapCount = pMxDoc.Maps.Count
3877:   For i = 0 To lMapCount - 1
3878:     Set pCurrMap = pMxDoc.Maps.Item(i)
3879:     If pCurrMap.Name = (GBL_BUBBLE_IDENTIFIER & lBubbleId) And (pCurrMap.Description = ("Detail Inset " & lBubbleId)) Then
3880:       pMxDoc.Maps.Remove pCurrMap
3881:       i = lMapCount - 1
3882:     End If
3883:   Next i
3884:   pOldMap.Name = GBL_BUBBLE_IDENTIFIER & lBubbleId
3885:   pOldMap.Description = "Detail Inset " & lBubbleId

End Sub




'This function takes two data frames, presumably one of them being the
'main data frame for the map, and the second data frame being a detail
'inset data frame.  The function finds the centerpoint of the inset
'data frame, determines the map units x/y of that center point in the
'main data frame, and returns a point with those values.
'---------------------------------------
Private Function GetDataframeCenterPoint(pMainElement As IElement, pBubbleElement As IElement, pMxDoc As IMxDocument) As IPoint
  On Error GoTo ErrorHandler


  Dim pMainMapFrame As IMapFrame, pPoint As IPoint, pEnv As IEnvelope
  Dim dDataFrameWidth As Double, dDataFrameHeight As Double
  Dim dBubbleLeftInches As Double, dBubbleRightInches As Double
  Dim dBubbleTopInches As Double, dBubbleBottomInches As Double
  Dim dBoundaryInchesXLeft As Double, dBoundaryInchesYTop As Double
  Dim dScrMapUnitWidth As Double, dScrMapUnitHeight As Double
  Dim pPntScrTopRight As IPoint, pPntScrBottomLeft As IPoint
  Dim pAV As IActiveView
  Dim dBubbleLeft As Double, dBubbleRight As Double, dBubbleTop As Double
  Dim dBubbleBottom As Double, dBubbleCenterX As Double, dBubbleCenterY As Double
  Dim dRatioInchToMapUnitX As Double, dRatioInchToMapUnitY As Double
  
  
3915:   Set pPoint = New Point
3916:   Set pMainMapFrame = pMainElement
3917:   Set pAV = pMainMapFrame.Map
  
  'get the bubble frame's screen location
  'in inches
  
  'Set pEnv = New Envelope
3923:   Set pEnv = pBubbleElement.Geometry.Envelope
3924:   dBubbleLeftInches = pEnv.XMin
3925:   dBubbleRightInches = pEnv.XMax
3926:   dBubbleTopInches = pEnv.YMax
3927:   dBubbleBottomInches = pEnv.YMin
  
3929:   dBoundaryInchesXLeft = pMainElement.Geometry.Envelope.XMin
3930:   dBoundaryInchesYTop = pMainElement.Geometry.Envelope.YMin
3931:   dDataFrameWidth = pMainElement.Geometry.Envelope.Width
3932:   dDataFrameHeight = pMainElement.Geometry.Envelope.Height
  
3934:   Set pPntScrBottomLeft = New Point
3935:   Set pPntScrTopRight = New Point
  
3937:   With pAV.ScreenDisplay.DisplayTransformation
3938:     Set pPntScrBottomLeft = .ToMapPoint(.DeviceFrame.Left, .DeviceFrame.bottom)
3939:     Set pPntScrTopRight = .ToMapPoint(.DeviceFrame.Right, .DeviceFrame.Top)
3940:   End With
  
3942:   dScrMapUnitWidth = pPntScrTopRight.X - pPntScrBottomLeft.X
3943:   dScrMapUnitHeight = pPntScrTopRight.Y - pPntScrBottomLeft.Y
3944:   dRatioInchToMapUnitX = dDataFrameWidth / dScrMapUnitWidth
3945:   dRatioInchToMapUnitY = dDataFrameHeight / dScrMapUnitHeight
  
3947:   dBubbleLeft = ((dBubbleLeftInches - dBoundaryInchesXLeft) / dRatioInchToMapUnitX) + pPntScrBottomLeft.X
3948:   dBubbleRight = ((dBubbleRightInches - dBoundaryInchesXLeft) / dRatioInchToMapUnitX) + pPntScrBottomLeft.X
3949:   dBubbleCenterX = ((dBubbleRight - dBubbleLeft) / 2) + dBubbleLeft
3950:   dBubbleTop = ((dBubbleTopInches - dBoundaryInchesYTop) / dRatioInchToMapUnitY) + pPntScrBottomLeft.Y
3951:   dBubbleBottom = ((dBubbleBottomInches - dBoundaryInchesYTop) / dRatioInchToMapUnitY) + pPntScrBottomLeft.Y
3952:   dBubbleCenterY = ((dBubbleTop - dBubbleBottom) / 2) + dBubbleBottom
3953:   pPoint.X = dBubbleCenterX
3954:   pPoint.Y = dBubbleCenterY
  
'''''''''''''''debug
Dim pMapFrame As IMapFrame
3958: Set pMapFrame = pBubbleElement
'3959: MsgBox "In NWMapSeries, GetDataFrameCenterPoint" & vbNewLine _
     & "dBubbleLeftInches is " & dBubbleLeftInches & vbNewLine _
     & "dBubbleRightInches is " & dBubbleRightInches & vbNewLine _
     & "dBubbleTopInches is " & dBubbleTopInches & vbNewLine _
     & "dBubbleBottomInches is " & dBubbleBottomInches & vbNewLine _
     & "dBoundaryInchesXLeft is " & dBoundaryInchesXLeft & vbNewLine _
     & "dBoundaryInchesYTop is " & dBoundaryInchesYTop & vbNewLine _
     & "dScrMapUnitWidth is " & dScrMapUnitWidth & vbNewLine _
     & "dScrMapUnitHeight is " & dScrMapUnitHeight & vbNewLine _
     & "dRatioInchToMapUnitX is " & dRatioInchToMapUnitX & vbNewLine _
     & "dRatioInchToMapUnitY is " & dRatioInchToMapUnitY & vbNewLine _
     & "dBubbleCenterX is " & dBubbleCenterX & vbNewLine _
     & "dBubbleCenterY is " & dBubbleCenterY & vbNewLine _
     & "pMainMapFrame.Map.Name is " & pMainMapFrame.Map.Name & vbNewLine _
     & "pBubbleFrame name is " & pMapFrame.Map.Name & vbNewLine
     
3975:   Set GetDataframeCenterPoint = pPoint

  Exit Function
ErrorHandler:
  HandleError False, "GetDataframeCenterPoint " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function




'RefreshDetailInsets
'  Package all the code per map page to work with detail insets into one subroutine
'  in order to keep the DrawPage routine simplified.
'----------------------------------------------------
Private Sub INWMapSeriesOptions_RefreshDetailInsets(pMxDoc As esriArcMapUI.IMxDocument, pBubbleLayer As esriCarto.ILayer, pPageShape As esriGeometry.IPolygon, Optional bUseBubbleGeometry As Boolean)


  On Error GoTo ErrorHandler

  Dim sDataFrameName As String, pDataFrameElement As IElement, lArrowCount As Long
  Dim lArrowShadowCount As Long, lDetailInsetCount As Long, i As Long
  Dim pGraphicsContainer As IGraphicsContainer, pElement As IElement, sInsetID As String
  Dim pLayers As IEnumLayer, pLayer As ILayer, vDetailInsetKeys() As Variant
  Dim bLyrIsFound As Boolean, pInsetLyr As ILayer, pFeatureLyr As IFeatureLayer
  Dim pFeatureClass As IFeatureClass, pSpatialFilter As ISpatialFilter
  Dim pFeatCursor As IFeatureCursor, pBubbleFeature As IFeature
  Dim pRow As esriGeoDatabase.IRow, pElemProps As IElementProperties
  Dim pDictInsetFrameCenterX As Scripting.Dictionary, dCenterX As Double
  Dim pDictInsetFrameCenterY As Scripting.Dictionary, dCenterY As Double
  Dim pMainDFElement As IElement
  
4006:   If IsMissing(bUseBubbleGeometry) Then
4007:     bUseBubbleGeometry = False
4008:   End If
  
  If pMxDoc Is Nothing Then Exit Sub
  
4012:   Set pDictInsetFrameCenterX = New Scripting.Dictionary
4013:   Set pDictInsetFrameCenterY = New Scripting.Dictionary
  
4015:   If Len(m_sMainDataFrame) = 0 Then
4016:     sDataFrameName = GetActiveDataFrameName(pMxDoc)
4017:   Else
4018:     sDataFrameName = m_sMainDataFrame
4019:   End If
  
  
4022:   Set pDataFrameElement = GetDataFrameElement(sDataFrameName, pMxDoc)
4023:   If Len(m_sMainDataFrame) > 0 Then
4024:     Set pMainDFElement = GetDataFrameElement(m_sMainDataFrame, pMxDoc)
4025:   End If
  '-----------------------------
  'remove all previous graphics,
  'and detail inset dataframes
  
  
4031:   Set pGraphicsContainer = pMxDoc.PageLayout
                                                  'make assumption that for every
                                                  'detail inset, one arrow and one
                                                  'arrow shadow exists
4035:   lDetailInsetCount = m_pDetailInsets.Count
4036:   vDetailInsetKeys = m_pDetailInsets.Keys

  Dim lLoop As Long, pDelColl As Collection, pEnv As IEnvelope
  Dim pPage As IPageLayout, pGraphCont As IGraphicsContainer
  Dim pElem As IElement, pMapFrame As IMapFrame, pAV As IActiveView
  Dim pMap As IMap, pPoint As IPoint
  
  
4044:   Set pPage = pMxDoc.PageLayout
4045:   Set pDelColl = New Collection
4046:   Set pGraphicsContainer = pPage
4047:   pGraphicsContainer.Reset
4048:   Set pElement = pGraphicsContainer.Next
4049:   Set pPoint = New Point
  
4051:   Do While Not pElement Is Nothing
4052:     If TypeOf pElement Is IMapFrame Then
4053:       Set pMapFrame = pElement
4054:       If StrComp(GBL_BUBBLE_IDENTIFIER, Left$(pMapFrame.Map.Name, Len(GBL_BUBBLE_IDENTIFIER))) = 0 Then
4055:         If bUseBubbleGeometry Then
          'Acquire and store the data frame's
          'center location, but in terms of a
          'map unit point x/y in the main data
          'frame's extent
4060:           Set pPoint = GetDataframeCenterPoint(pMainDFElement, pElement, pMxDoc)
4061:           pDictInsetFrameCenterX.Add pMapFrame.Map.Name, pPoint.X
4062:           pDictInsetFrameCenterY.Add pMapFrame.Map.Name, pPoint.Y
4063:         End If
4064:         pDelColl.Add pMapFrame
4065:       End If
4066:     End If
4067:     Set pElement = pGraphicsContainer.Next
4068:   Loop
  
4070:   For lLoop = 1 To pDelColl.Count
4071:     pGraphicsContainer.DeleteElement pDelColl.Item(lLoop)
4072:   Next lLoop
  
  
  
4076:   For i = 0 To lDetailInsetCount - 1
4077:     sInsetID = vDetailInsetKeys(i)
4078:     Set pElement = m_pDetailInsetArrows.Item(sInsetID)
4079:     pGraphicsContainer.DeleteElement pElement
    
4081:     Set pElement = m_pDetailInsetArrowShadows.Item(sInsetID)
4082:     pGraphicsContainer.DeleteElement pElement
4083:     Set pElement = m_pDetailInsets.Item(sInsetID)
4084:     pGraphicsContainer.DeleteElement pElement
  
4086:     m_pDetailInsetArrows.Remove sInsetID
4087:     m_pDetailInsetArrowShadows.Remove sInsetID
4088:     m_pDetailInsets.Remove sInsetID
4089:   Next i
  
  'Delete the detail inset's data frame
  'lMapCount = pMxDoc.Maps.Count
  'For i = 0 To lMapCount - 1
  '  Set pCurrMap = pMxDoc.Maps.Item(i)
  '  If pCurrMap.Name = (GBL_BUBBLE_IDENTIFIER & lBubbleID) And (pCurrMap.Description = ("Detail Inset " & lBubbleID)) Then
  '    pMxDoc.Maps.Remove (pCurrMap)
  '    i = lMapCount - 1
  '  End If
  'Next i
  'pOldMap.Name = GBL_BUBBLE_IDENTIFIER & lBubbleID
  'pOldMap.Description = "Detail Inset " & lBubbleID
  
  
  '-------------------------------------
  'Acquire the selection of detail inset
  'polygons from which to generate detail insets.
  
4108:   Set pInsetLyr = pBubbleLayer
4109:   Set pFeatureLyr = pInsetLyr
  If pFeatureLyr Is Nothing Then Exit Sub
4111:   Set pFeatureClass = pFeatureLyr.FeatureClass
                                                  'select all features that are entirely
                                                  'contained within the map extent.
4114:   Set pSpatialFilter = New SpatialFilter
4115:   pSpatialFilter.SpatialRel = esriSpatialRelEnvelopeIntersects
  'pSpatialFilter.SpatialRel = esriSpatialRelOverlaps   'catches only border features
  'pSpatialFilter.SpatialRel = esriSpatialRelContains  'doesn't catch border features
4118:   pSpatialFilter.GeometryField = pFeatureClass.ShapeFieldName
4119:   Set pSpatialFilter.Geometry = pPageShape
                                                  'use pFeatureClass.Search pSpatialFilter to
                                                  'get a cursor of all detail insets within the
                                                  'current map
4123:   Set pFeatCursor = pFeatureClass.Search(pSpatialFilter, True)
4124:   Set pBubbleFeature = pFeatCursor.NextFeature

  
  '------------------------------------------------
  'for each result feature, generate a detail inset
       
       
4131:   Do While Not pBubbleFeature Is Nothing
    'update the source X/Y attribute based
    'on the polygon's extent center
4134:     If bUseBubbleGeometry Then
4135:       UpdateBubbleXYAttributes pBubbleFeature, pDictInsetFrameCenterX, pDictInsetFrameCenterY
4136:     End If
4137:     Set pRow = pBubbleFeature
4138:     AddToArcMap pDataFrameElement, pMxDoc, pRow, m_pDetailInsets, m_pDetailInsetArrows, m_pDetailInsetArrowShadows
4139:     Set pBubbleFeature = pFeatCursor.NextFeature
4140:   Loop
                                                  'created detail insets are selected.
                                                  'clear that selection
  Dim pViewManager As IViewManager, pMaps As IMaps, lMapCount As Long
4144:   If Not bUseBubbleGeometry Then
4145:     Set pMaps = pMxDoc.Maps
4146:     lMapCount = pMaps.Count
4147:     For i = 0 To (lMapCount - 1)
4148:       Set pMap = pMaps.Item(i)
4149:       Set pViewManager = pMap
4150:       pViewManager.ElementSelection.Clear
4151:     Next i
4152:   End If
  
4154:   Set pDictInsetFrameCenterX = Nothing
4155:   Set pDictInsetFrameCenterY = Nothing
  Exit Sub
ErrorHandler:
4158:   Set pDictInsetFrameCenterX = Nothing
4159:   Set pDictInsetFrameCenterY = Nothing
  
  HandleError False, "DeleteFromArcMap " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub





'UpdateBubbleXYAttributes
'  Set the XORG and YORG attribute fields based on the centroid of the
'  polygon, and set the XDEST and YDEST fields based on the values passed
'  in with the scripting.dictionary variables.
'-----------------------------------
Private Sub UpdateBubbleXYAttributes(pBubbleFeature As IFeature, ByRef pDictXCenters As Scripting.Dictionary, ByRef pDictYCenters As Scripting.Dictionary)
  On Error GoTo ErrorHandler

  
  Dim pPolygon As IPolygon, pEnvelope As IEnvelope, pFields As IFields
  Dim lFieldIdx As Long, i As Long, pRow As IRow
  Dim lBubbleId As Long, dDestX As Double, dDestY As Double
  Dim lXOrgIdx As Long, lYOrgIdx As Long, lXDestIdx As Long, lYDestIdx As Long
  Dim lBubbleIdx As Long, sBubbleID As String
  
  
4184:   Set pPolygon = pBubbleFeature.Shape
4185:   Set pEnvelope = pPolygon.Envelope
  
4187:   Set pFields = pBubbleFeature.Fields
4188:   lXOrgIdx = pFields.FindField("XORG")
4189:   lYOrgIdx = pFields.FindField("YORG")
4190:   lXDestIdx = pFields.FindField("XDEST")
4191:   lYDestIdx = pFields.FindField("YDEST")
4192:   lBubbleIdx = pFields.FindField("BUBBLEID")
  
  If lXOrgIdx < 0 Then Exit Sub
  If lYOrgIdx < 0 Then Exit Sub
  If lXDestIdx < 0 Then Exit Sub
  If lYDestIdx < 0 Then Exit Sub
  If lBubbleIdx < 0 Then Exit Sub
  
4200:   Set pRow = pBubbleFeature
     
  'get the BubbleID
4203:   lBubbleId = pRow.Value(lBubbleIdx)
4204:   sBubbleID = GBL_BUBBLE_IDENTIFIER & lBubbleId
  
  If Not pDictYCenters.Exists(sBubbleID) Then Exit Sub
  If Not pDictXCenters.Exists(sBubbleID) Then Exit Sub
  If pDictXCenters(sBubbleID) = 0 Then Exit Sub
  If pDictYCenters(sBubbleID) = 0 Then Exit Sub
  
4211:   pRow.Value(lXOrgIdx) = ((pEnvelope.XMax - pEnvelope.XMin) / 2) + pEnvelope.XMin
4212:   pRow.Value(lYOrgIdx) = ((pEnvelope.YMax - pEnvelope.YMin) / 2) + pEnvelope.YMin
  
  
'''''''''''''''debug
'4216: MsgBox "In NWMappage, UpdateBubbleXYAttributes, " & vbNewLine _
     & " lXOrgIdx is " & lXOrgIdx & vbNewLine _
     & " lYOrgIdx is " & lYOrgIdx & vbNewLine _
     & " lXDestIdx is " & lXDestIdx & vbNewLine _
     & " lYDestIdx is " & lYDestIdx & vbNewLine _
     & " lBubbleIdx is " & lBubbleIdx & vbNewLine _
     & " pRow.Value(lXOrgIdx) is " & pRow.Value(lXOrgIdx) & vbNewLine _
     & " pRow.Value(lYOrgIdx) is " & pRow.Value(lYOrgIdx) & vbNewLine _
     & " pRow.Value(lXDestIdx) is " & pRow.Value(lXDestIdx) & vbNewLine _
     & " pRow.Value(lYDestIdx) is " & pRow.Value(lYDestIdx) & vbNewLine _
     & " pRow.Value(lBubbleIdx) is " & pRow.Value(lBubbleIdx) & vbNewLine _
     & " pDictYCenters(sBubbleID) is " & pDictYCenters(sBubbleID) & vbNewLine _
     & " pDictXCenters(sBubbleID) is " & pDictXCenters(sBubbleID) & vbNewLine
     
     
  
4232:   If pDictXCenters.Exists(sBubbleID) Then
4233:     dDestX = pDictXCenters(sBubbleID)
4234:     pDictXCenters.Remove sBubbleID
4235:   End If
4236:   If pDictYCenters.Exists(sBubbleID) Then
4237:     dDestY = pDictYCenters(sBubbleID)
4238:     pDictYCenters.Remove sBubbleID
4239:   End If
4240:   pRow.Value(lXDestIdx) = dDestX
4241:   pRow.Value(lYDestIdx) = dDestY
4242:   pRow.Store
  
  Exit Sub
ErrorHandler:
  HandleError False, "UpdateBubbleXYAttributes " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub



Private Function ConvertPixelsToRW(pixelUnits As Double, pMxDoc As IMxDocument) As Double
  Dim realWorldDisplayExtent As Double
  Dim pixelExtent As Long
  Dim sizeOfOnePixel As Double
  Dim pDT As IDisplayTransformation
  Dim deviceRECT As tagRECT
  Dim pEnv As IEnvelope
  Dim pActiveView As IActiveView

4260:   Set pActiveView = pMxDoc.FocusMap
4261:   Set pDT = pActiveView.ScreenDisplay.DisplayTransformation
4262:   deviceRECT = pDT.DeviceFrame
4263:   pixelExtent = deviceRECT.Right - deviceRECT.Left
4264:   Set pEnv = pDT.VisibleBounds
4265:   realWorldDisplayExtent = pEnv.Width
4266:   sizeOfOnePixel = realWorldDisplayExtent / pixelExtent
4267:   ConvertPixelsToRW = pixelUnits * sizeOfOnePixel
End Function








Public Function GetDataFrameElement(sDataFrameName As String, pMxDoc As IMxDocument) As IElement
' Get the data frame element by name
    Dim pGraphicsContainer As IGraphicsContainer
    Dim pElementProperties As IElementProperties
    Dim pElement As IElement
    Dim pFE As IFrameElement
    Dim pElProps As IElementProperties
    
    On Error GoTo ErrorHandler
    
    ' Loop through the elements (in the layout)
4288:     Set pGraphicsContainer = pMxDoc.PageLayout
4289:     pGraphicsContainer.Reset
4290:     Set pElement = pGraphicsContainer.Next
4291:     While Not pElement Is Nothing
        ' If type of element is an IFrameElement
4293:         If TypeOf pElement Is IFrameElement Then
4294:             Set pElProps = pElement
            ' If Name matches
4296:             If UCase(pElProps.Name) = UCase(sDataFrameName) Then
                ' Return element
4298:                 Set GetDataFrameElement = pElement
4299:                 Set pElement = Nothing
4300:             Else
4301:                 Set pElement = pGraphicsContainer.Next
4302:             End If
4303:         Else
4304:             Set pElement = pGraphicsContainer.Next
4305:         End If
4306:     Wend
    
    Exit Function
ErrorHandler:
4310:     Err.Raise Err.Number, Err.Source, "Error in GetDataFrameElement:" _
        & vbCrLf & Err.Description
End Function







Public Function GetActiveDataFrameName(pMxDoc As IMxDocument) As String
    Dim pFocusMap As IMap
    
4323:     Set pFocusMap = pMxDoc.FocusMap
4324:     GetActiveDataFrameName = pFocusMap.Name
End Function




''VisibleElementRemove
''
''Sets the layout element visibility data structures so that the element
''under this name will not show up on this map page.
''------------------------------
'Public Sub VisibleElementRemove(sMapPageID As String, sElementID As String)
'  On Error GoTo ErrorHandler
'
'  Dim pScrDict As Scripting.Dictionary, sPlaceHolder As String
'  sPlaceHolder = " " 'room for future expansion, will save/load under ipersistvariant
'
'  If m_pDictPagesWhereElementIsInvisible Is Nothing Then
'    Set m_pDictPagesWhereElementIsInvisible = New Scripting.Dictionary
'  End If
'
'  If Not m_pDictPagesWhereElementIsInvisible.Exists(sElementID) Then
'    Set pScrDict = New Scripting.Dictionary
'    m_pDictPagesWhereElementIsInvisible.Add sElementID, pScrDict
'  Else
'  Set pScrDict = m_pDictPagesWhereElementIsInvisible(sElementID)
'    If Not pScrDict.Exists(sMapPageID) Then
'      pScrDict.Add sMapPageID, sPlaceHolder
'    End If
'  End If
'
'  Exit Sub
'ErrorHandler:
'  HandleError True, "VisibleElementRemove " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
'End Sub







'VisibleElementAdd
'
'Adds a layout element as visible for the given map page.
'Since there are two data structures, one for tracking what map pages
'there are per element, and another tracking what elements there are per
'map page, both are updated.
'
'To avoid the problems of cyclic references, the two data structures do
'not reference one another.  This costs more memory, but should make the
'code more stable during un-tagging of elements.
'
'In initial version, experimenting with using just one of those two
'data structures for programming simplicity (and therefore low bug count)
'sake.
'---------------------------
Public Sub VisibleElementAdd(sMapPageID As String, sElementID As String, pElement As IElement)
  On Error GoTo ErrorHandler

  'In order to simplify initialization, and to make the program more
  'robust, the initial data structure setting of being empty will coincide
  'with the initial setting of layout element visibility of being visible
  'for all map pages.
  
  Dim pScrDict As Scripting.Dictionary
  
'  If m_pDictTaggedElements Is Nothing Then
'    Set m_pDictTaggedElements = New Scripting.Dictionary
'  End If
'  If Not m_pDictTaggedElements.Exists(sElementID) Then
'    m_pDictTaggedElements.Add sElementID, pElement
'  End If
4397:   If m_pDictPagesWhereElementIsInvisible Is Nothing Then
4398:     Set m_pDictPagesWhereElementIsInvisible = New Scripting.Dictionary
    Exit Sub
4400:   End If
4401:   If m_pDictPagesWhereElementIsInvisible.Exists(sElementID) Then
4402:     Set pScrDict = m_pDictPagesWhereElementIsInvisible(sElementID)
    If pScrDict Is Nothing Then Exit Sub
4404:     If pScrDict.Exists(sMapPageID) Then
4405:       pScrDict.Remove sMapPageID
4406:     End If
4407:   End If

  Exit Sub
ErrorHandler:
  HandleError True, "AddVisibleElement " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub







Public Sub SetVisibleLayers(pMap As IMap, sLayers As String)
  On Error GoTo ErrorHandler

  Dim sLayerArr() As String, i As Long, j As Long
  Dim lVisLyrUBound As Long, pLayers As IEnumLayer
  Dim pLayer As ILayer
  
  
4428:   sLayerArr = Split(sLayers, ",", -1, vbTextCompare)
4429:   lVisLyrUBound = UBound(sLayerArr)
  
4431:   Set pLayers = pMap.Layers
4432:   pLayers.Reset
4433:   Set pLayer = pLayers.Next
  
4435:   If lVisLyrUBound = -1 Then
4436:     Do While Not pLayer Is Nothing
4437:       pLayer.Visible = False
4438:       Set pLayer = pLayers.Next
4439:     Loop
    Exit Sub
4441:   Else
4442:     Do While Not pLayer Is Nothing
4443:       pLayer.Visible = False
4444:       For i = 0 To lVisLyrUBound
4445:         If pLayer.Name = sLayerArr(i) Then
4446:           pLayer.Visible = True
4447:         End If
4448:       Next i
4449:       Set pLayer = pLayers.Next
4450:     Loop
4451:   End If

  Exit Sub
ErrorHandler:
  HandleError True, "SetVisibleLayers " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub







