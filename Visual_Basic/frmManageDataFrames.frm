VERSION 5.00
Begin VB.Form frmManageDataFrames 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Data Frames"
   ClientHeight    =   6876
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5424
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6876
   ScaleWidth      =   5424
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUncheckAll 
      Caption         =   "&Uncheck All"
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdCheckAll 
      Caption         =   "&Check All"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   6360
      Width           =   1215
   End
   Begin VB.ListBox lstMapPages 
      Height          =   2424
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   3480
      Width           =   5175
   End
   Begin VB.ComboBox cboDataFrames 
      Height          =   1296
      Left            =   120
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   1
      Text            =   "cboDataFrames"
      Top             =   1560
      Width           =   5175
   End
   Begin VB.Label Label3 
      Caption         =   "Map Pages"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Available Data Frames"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   $"frmManageDataFrames.frx":0000
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frmManageDataFrames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_pApp As IApplication
Private m_pMxDoc As IMxDocument
Private m_bCancelled As Boolean
Private m_pNWMapBook As INWDSMapBook
Private m_pNWMapSeries As INWDSMapSeries
Private m_pNWSeriesOptions As INWMapSeriesOptions
Private m_sCurrentDataFrame As String
Private m_bLstMapPagesIsLocked As Boolean
Private m_sMainDataFrame As String

Const c_sModuleFileName As String = "frmManageDataFrames.frm"





Public Property Get WasCancelled() As Boolean
20:   WasCancelled = m_bCancelled
End Property


Public Property Get App() As IApplication
25:   Set App = m_pApp
End Property

Public Property Set App(pApp As IApplication)
29:   Set m_pApp = pApp
30:   If Not pApp Is Nothing Then
31:     Set m_pMxDoc = pApp.Document
32:   End If
End Property



Private Sub cboDataFrames_Change()
  'handle bad input
39:   With cboDataFrames
40:     If FindControlString(cboDataFrames, .Text, -1, True) = -1 Then
41:       If .ListCount > 0 Then
42:         .ListIndex = 0
43:       ElseIf .ListIndex >= .ListCount Then
44:         .ListIndex = 0
45:       End If
46:       .Text = .List(.ListIndex)
47:     End If
48:   End With
End Sub

Private Sub cboDataFrames_Click()
  On Error GoTo ErrorHandler

  Dim vPagesWhereVisible As Variant, lPageCount As Long, i As Long
  Dim lPageIdx As Long, sCurrentPage As String
  
  
  'load the selections of map pages based on the list
  'associated with this data frame
60:   If m_pNWSeriesOptions Is Nothing Then
    Exit Sub
62:   End If
63:   m_sCurrentDataFrame = cboDataFrames.List(cboDataFrames.ListIndex)
64:   If StrComp(m_sCurrentDataFrame, m_sMainDataFrame, vbTextCompare) = 0 Then
65:     With lstMapPages
66:       For i = 0 To (.ListCount - 1)
67:         .Selected(i) = True
68:         m_pNWSeriesOptions.DataFrameSetVisibleInPage m_sMainDataFrame, .List(i)
69:       Next i
70:       .Enabled = False
71:     End With
72:   Else
73:     With cboDataFrames
      
      
76:       lstMapPages.Enabled = True
      
78:       lPageCount = m_pNWSeriesOptions.DataFramePagesWhereVisibleCount(m_sCurrentDataFrame)
79:       If lPageCount = -1 Then   'new data frame
                                                  'set all map pages to visible
81:         m_bLstMapPagesIsLocked = True
82:         For i = 0 To (lstMapPages.ListCount - 1)
                                                  'in the ui
84:           lstMapPages.Selected(i) = True
                                                  'and in the data structures
86:           m_pNWSeriesOptions.DataFrameSetVisibleInPage m_sCurrentDataFrame, lstMapPages.List(i)
87:         Next i
88:         m_bLstMapPagesIsLocked = False
      
90:       Else
91:         vPagesWhereVisible = m_pNWSeriesOptions.DataFramePagesWhereVisible(m_sCurrentDataFrame)
92:         lPageCount = m_pNWSeriesOptions.DataFramePagesWhereVisibleCount(m_sCurrentDataFrame)
93:         m_bLstMapPagesIsLocked = True
94:         For i = 0 To lstMapPages.ListCount - 1
95:           lstMapPages.Selected(i) = False
96:         Next i
97:         m_bLstMapPagesIsLocked = False
    
99:         For i = 0 To (lPageCount - 1)
100:           sCurrentPage = vPagesWhereVisible(i)
101:           lPageIdx = FindControlString(lstMapPages, sCurrentPage, -1, True)
102:           If lPageIdx > -1 Then
                                                  'add checkmarks beside the appropriate map pages.
104:             m_bLstMapPagesIsLocked = True
105:             lstMapPages.Selected(lPageIdx) = True
106:             m_bLstMapPagesIsLocked = False
107:           Else
                                                  'remove from the series any map page that
                                                  'doesn't happen to exist in the listbox
110:             m_pNWSeriesOptions.DataFrameRemovePage (m_sCurrentDataFrame), sCurrentPage
111:           End If
112:         Next i
113:       End If
114:     End With
115:   End If

  Exit Sub
ErrorHandler:
119:   m_bLstMapPagesIsLocked = False
  HandleError True, "cboDataFrames_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdCheckAll_Click()
  Dim i As Long
  
126:   With Me.lstMapPages
127:     For i = 0 To .ListCount - 1
128:       .ItemData(i) = 1
'129:       .Selected(i) = True
130:     Next i
131:   End With
End Sub

'Private Sub cmdCancel_Click()
'124:   m_bCancelled = True
'125:   Me.Hide
'End Sub

Private Sub cmdOK_Click()
140:   m_bCancelled = False
141:   Me.Hide
End Sub


Public Sub Initialize()
  On Error GoTo ErrorHandler

  Dim i As Long, pGraphicsContainer As IGraphicsContainer, pPageLayout As IPageLayout
  Dim pMapFrame As IMapFrame, pElement As IElement, sBubbleStr As String
  Dim bWarningWasGiven As Boolean, lPageCount As Long, pNWDSMapPage As INWDSMapPage
  Dim lFindResult As Long, sStoredFrames() As Variant, lStoredFramesCount As Long
  Dim pNWDSMapSeriesProps As INWDSMapSeriesProps, lPageIdx As Long, vStoredFrames As Variant
  Dim vMapPages As Variant, sMapPage As String, sStoredFrame As String
  
  If m_pApp Is Nothing Then Exit Sub
  
157:   Set m_pNWMapBook = GetMapBookExtension(m_pApp)
158:   Set m_pNWMapSeries = m_pNWMapBook.ContentItem(0)
159:   Set m_pNWSeriesOptions = m_pNWMapSeries
160:   Set pNWDSMapSeriesProps = m_pNWSeriesOptions
161:   m_sMainDataFrame = pNWDSMapSeriesProps.DataFrameName
  
  
  
  
  'load list of data frames
  '''''''''''''''''''''''''
  
169:   Set pPageLayout = m_pMxDoc.PageLayout
170:   Set pGraphicsContainer = pPageLayout
171:   pGraphicsContainer.Reset
172:   Set pElement = pGraphicsContainer.Next
  
  'add all desired data frames from the map layout
175:   cboDataFrames.Clear
176:   Do While Not pElement Is Nothing
177:     If TypeOf pElement Is IMapFrame Then
178:       Set pMapFrame = pElement
179:       If StrComp(Left$(pMapFrame.Map.Name, Len("BubbleID:")), "BubbleID:", vbTextCompare) <> 0 Then
180:         lFindResult = FindControlString(cboDataFrames, pMapFrame.Map.Name, 0, True)
181:         If lFindResult > -1 Then
                                                  'notify the user that more than one data
                                                  'frame has the same name
184:           If Not bWarningWasGiven Then
185:             bWarningWasGiven = True
186:             MsgBox "Warning: Visibility of duplicate name data frames will not be " & vbNewLine _
                 & "tracked by the NW Map Book application." & vbNewLine _
                 & "More than one data frame called ''" & pMapFrame.Map.Name & "''" & vbNewLine _
                 & "was detected.", vbOKOnly
190:           End If
191:         Else
                                                  'add the data frame to the list
193:           cboDataFrames.AddItem pMapFrame.Map.Name
194:         End If
195:       End If
196:     End If
197:     Set pElement = pGraphicsContainer.Next
198:   Loop
  
  'add stored data frames (those not
  'currently in the map layout)
202:   vStoredFrames = m_pNWSeriesOptions.DataFramesStored
203:   lStoredFramesCount = UBound(vStoredFrames) + 1
204:   For i = 0 To (lStoredFramesCount - 1)
205:     sStoredFrame = vStoredFrames(i)
206:     If FindControlString(cboDataFrames, sStoredFrame, -1, True) > -1 Then
207:       MsgBox "Warning: Visibility of duplicate name data frames will not be " & vbNewLine _
           & "tracked by the NW Map Book application." & vbNewLine _
           & "More than one data frame called ''" & sStoredFrame & "''" & vbNewLine _
           & "was detected.", vbOKOnly
211:     Else
212:       cboDataFrames.AddItem sStoredFrame
213:     End If
214:   Next i
  
  
  'load list of map pages.
  ''''''''''''''''''''''''
219:   lstMapPages.Clear
220:   lPageCount = m_pNWMapSeries.PageCount
    
222:   For i = 0 To (lPageCount - 1)
223:     Set pNWDSMapPage = m_pNWMapSeries.Page(i)
224:     lstMapPages.AddItem pNWDSMapPage.PageName
225:   Next i

232:   CleanOrphanedDataFrameStructs
  
  'initialize all the data frames by triggering
  'the cboDataFrames_Click handler for each one
236:   For i = 0 To (cboDataFrames.ListCount - 1)
237:     cboDataFrames.ListIndex = i
238:   Next i
239:   cboDataFrames.ListIndex = 0
                                                  'freeze out the UI if the 1st dataframe
                                                  'is the one containing the index layer
'  With cboDataFrames
'    If StrComp(.List(0), m_sMainDataFrame, vbTextCompare) = 0 Then
'      m_bLstMapPagesIsLocked = True
'      For i = 0 To (lstMapPages.ListCount - 1)
'        lstMapPages.Selected(i) = True
'      Next i
'      m_bLstMapPagesIsLocked = False
'      lstMapPages.Enabled = False
'    Else
'                                                  'load previous settings for this data frame
'      m_bLstMapPagesIsLocked = True
'      For i = 0 To (lstMapPages.ListCount - 1)
'        lstMapPages.Selected(i) = False
'      Next i
'      m_bLstMapPagesIsLocked = False
'      lstMapPages.Enabled = True
'
'      vMapPages = m_pNWSeriesOptions.DataFramePagesWhereVisible(.List(0))
'      lPageCount = m_pNWSeriesOptions.DataFramePagesWhereVisibleCount(.List(0))
'
'      If lPageCount = -1 Then   'new data frame
'                                                  'set all map pages to visible
'        m_bLstMapPagesIsLocked = True
'        For i = 0 To (lstMapPages.ListCount - 1)
'                                                  'in the ui
'          lstMapPages.Selected(i) = True
'                                                  'and in the data structures
'          m_pNWSeriesOptions.DataFrameSetVisibleInPage .List(0), lstMapPages.List(i)
'        Next i
'        m_bLstMapPagesIsLocked = False
'      Else
'        For i = 0 To (lPageCount - 1)
'          sMapPage = vMapPages(i)
'          lPageIdx = FindControlString(lstMapPages, sMapPage, -1, True)
'          If lPageIdx > -1 Then
'            m_bLstMapPagesIsLocked = True
'            lstMapPages.Selected(lPageIdx) = True
'            m_bLstMapPagesIsLocked = False
'          Else
'                                                    'remove from the series any map page that
'                                                    'doesn't happen to exist in the listbox
'            m_pNWSeriesOptions.DataFrameRemovePage .List(0), sMapPage
'          End If
'        Next i
'      End If
'    End If
'  End With
  
  Exit Sub
ErrorHandler:
  HandleError True, "Initialize " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub


'Occasionally, the user may have set map page visibility for a
'data frame, then will have deleted that data frame entirely.
'This leaves the management data structure in its wake, never
'to be deleted, but for this routine.  This routine should be
'run only when the list of data frames has been completely
'populated.
'''''''''''''''''''''''''''''''''''''''''''
Private Sub CleanOrphanedDataFrameStructs()
  Dim vManagedFrames As Variant, lManagedFramesCount As Long, sFrameName As String
  Dim i As Long
  
312:   If m_pNWSeriesOptions Is Nothing Then
    Exit Sub
314:   End If
  
  'for each data frame data structure,
    'does that structure have an existing data frame?
      'if not, then clean up that orphaned data structure
319:   vManagedFrames = m_pNWSeriesOptions.DataFramesManaged
320:   lManagedFramesCount = UBound(vManagedFrames) + 1
321:   For i = 0 To (lManagedFramesCount - 1)
322:     sFrameName = vManagedFrames(i)
323:     If FindControlString(Me.cboDataFrames, sFrameName, -1, True) = -1 Then
324:       m_pNWSeriesOptions.DataFrameRemoveFrame sFrameName
325:     End If
326:   Next i
End Sub


Private Sub cmdUncheckAll_Click()
  Dim i As Long
  
333:   With lstMapPages
334:     For i = 0 To .ListCount - 1
335:       .Selected(i) = False
336:     Next i
337:   End With
End Sub

Private Sub Form_Terminate()
341:   m_bCancelled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
345:   m_bCancelled = True
End Sub



Private Sub lstMapPages_ItemCheck(Item As Integer)
  On Error GoTo ErrorHandler

  If m_bLstMapPagesIsLocked Then Exit Sub
  If m_pNWMapSeries Is Nothing Then Exit Sub
  If m_pNWSeriesOptions Is Nothing Then Exit Sub
  
357:   With lstMapPages
358:     If .Selected(Item) Then
359:       m_pNWSeriesOptions.DataFrameSetVisibleInPage m_sCurrentDataFrame, .List(Item)
360:     Else
361:       m_pNWSeriesOptions.DataFrameSetInvisibleInPage m_sCurrentDataFrame, .List(Item)
362:     End If
363:   End With
  

  Exit Sub
ErrorHandler:
  HandleError True, "lstMapPages_ItemCheck " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
              

