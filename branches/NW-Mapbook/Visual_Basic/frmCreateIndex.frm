VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCreateIndex 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Create Index"
   ClientHeight    =   4584
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   4824
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4584
   ScaleWidth      =   4824
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog codOutput 
      Left            =   90
      Top             =   4020
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1125
      Left            =   60
      TabIndex        =   10
      Top             =   2820
      Width           =   4515
      Begin VB.TextBox txtPageNumber 
         Height          =   315
         Left            =   3720
         TabIndex        =   15
         Top             =   660
         Width           =   465
      End
      Begin VB.OptionButton optIndex 
         Caption         =   "Page Number (Number shown on Series list)"
         Height          =   225
         Index           =   1
         Left            =   810
         TabIndex        =   13
         Top             =   420
         Width           =   3405
      End
      Begin VB.OptionButton optIndex 
         Caption         =   "Page Label"
         Height          =   225
         Index           =   0
         Left            =   810
         TabIndex        =   12
         Top             =   60
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Add this value to each page number:"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   14
         Top             =   690
         Width           =   2625
      End
      Begin VB.Label Label3 
         Caption         =   "Index by:"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   11
         Top             =   30
         Width           =   735
      End
   End
   Begin VB.ComboBox cmbFieldName 
      Height          =   315
      Left            =   870
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1860
      Width           =   3405
   End
   Begin VB.ComboBox cmbLayer 
      Height          =   315
      Left            =   870
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1410
      Width           =   3405
   End
   Begin VB.CommandButton cmdBrowse 
      Height          =   315
      Left            =   4350
      Picture         =   "frmCreateIndex.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2310
      Width           =   345
   End
   Begin VB.TextBox txtOutput 
      Enabled         =   0   'False
      Height          =   315
      Left            =   870
      TabIndex        =   2
      Top             =   2310
      Width           =   3375
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4050
      TabIndex        =   1
      Top             =   4110
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   4110
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Field:"
      Height          =   255
      Index           =   1
      Left            =   390
      TabIndex        =   9
      Top             =   1890
      Width           =   405
   End
   Begin VB.Label Label2 
      Caption         =   "Layer:"
      Height          =   255
      Index           =   0
      Left            =   330
      TabIndex        =   6
      Top             =   1440
      Width           =   465
   End
   Begin VB.Label Label1 
      Caption         =   $"frmCreateIndex.frx":047A
      Height          =   1215
      Left            =   60
      TabIndex        =   5
      Top             =   90
      Width           =   4665
   End
   Begin VB.Label lblExportTo 
      Caption         =   "Output to:"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   2340
      Width           =   735
   End
End
Attribute VB_Name = "frmCreateIndex"
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

Public m_pApp As IApplication
Private m_pMasterColl As Collection

Private Sub cmbLayer_Click()
On Error GoTo ErrHand:
  Dim pFeatLayer As IFeatureLayer, pFields As IFields, lLoop As Long
  
  Set pFeatLayer = FindFeatureLayerByName(cmbLayer.List(cmbLayer.ListIndex), m_pApp)
  If pFeatLayer Is Nothing Then Exit Sub
  
  cmbFieldName.Clear
  Set pFields = pFeatLayer.FeatureClass.Fields
  For lLoop = 0 To pFields.FieldCount - 1
    If pFields.Field(lLoop).Type = esriFieldTypeString Then
      cmbFieldName.AddItem pFields.Field(lLoop).Name
    End If
  Next lLoop
  If cmbFieldName.ListCount > 0 Then
    cmbFieldName.ListIndex = 0
  End If
  
  CheckSettings

  Exit Sub
ErrHand:
  MsgBox "cmbLayer_Click - " & Err.Description
End Sub

Private Sub cmdBrowse_Click()
On Error GoTo ErrHand:
  codOutput.DialogTitle = "Specify output file to create"
  codOutput.Filter = "(*.txt)|*.txt"
  codOutput.Flags = cdlOFNOverwritePrompt
  codOutput.ShowSave
  If codOutput.FileName = "" Then
    txtOutput.Text = ""
  Else
    txtOutput.Text = codOutput.FileName
  End If
  
  CheckSettings
  
  Exit Sub
ErrHand:
  MsgBox "cmdBrowse_Click - " & Err.Description
End Sub

Private Sub cmdCancel_Click()
  Set m_pApp = Nothing
  Unload Me
End Sub

Private Sub cmdOK_Click()
'This routine will create the index and write it out to the specified file.
On Error GoTo ErrHand:
  Dim pDoc As IMxDocument, pMap As IMap, lLoop As Long
  Dim pFeatLayer As IFeatureLayer, pMapBook As INWDSMapBook, pSeries As INWDSMapSeries
  Dim pPage As INWDSMapPage, pColl As Collection
  Dim sPageId As String, sFieldName As String
  Dim sTempVal As String, sTempPage As String, lPos As Long
  Dim sPrev As String, sTemp As String, sOutput As String
  Dim pOutputPages As Collection, lLoop2 As Long, sOutputPages As String
  Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then
    MsgBox "Map book was not found!!!"
    Exit Sub
  End If
  
  'Get the index layer
  Set pFeatLayer = FindFeatureLayerByName(cmbLayer.List(cmbLayer.ListIndex), m_pApp)
  If pFeatLayer Is Nothing Then
    MsgBox "Count not find the index layer for some reason!!!"
    Exit Sub
  End If
  sFieldName = cmbFieldName.List(cmbFieldName.ListIndex)
  Set pSeries = pMapBook.ContentItem(0)
  
  'Setup the progress bar
  Screen.MousePointer = vbHourglass
  With m_pApp.StatusBar.ProgressBar
    .Message = "Index Creation:"
    .MaxRange = pSeries.PageCount
    .StepValue = 1
    .Position = 1
    .Show
  End With
  
  'Loop through the pages returning a collection of the attribute values returned by the
  'features found on the page.  Add the page collections to the master.
  Set m_pMasterColl = New Collection
  For lLoop = 0 To pSeries.PageCount - 1
    Set pPage = pSeries.Page(lLoop)
    Set pColl = pPage.IndexPage(pFeatLayer, sFieldName)
    If optIndex(0).Value Then
      sPageId = pPage.PageName
    Else
      sPageId = CStr(lLoop + 1 + CLng(txtPageNumber.Text))
    End If
    AddPageToMasterCollection pColl, sPageId
    
    m_pApp.StatusBar.ProgressBar.Step
  Next lLoop
  
  'Dump the master collection out to the specified file
  sPrev = ""
  Open txtOutput.Text For Output As #1
  For lLoop = 1 To m_pMasterColl.count
    sTemp = m_pMasterColl.Item(lLoop)
    lPos = InStr(1, sTemp, "-$$$$-")
    sTempVal = Left(sTemp, lPos - 1)
    sTempPage = Mid(sTemp, lPos + 6)
    If sPrev = "" Then
      Set pOutputPages = New Collection
      sOutput = sTempVal
      pOutputPages.Add sTempPage, sTempPage
      sPrev = sTempVal
    ElseIf sPrev = sTempVal Then
      If optIndex(0).Value Then
        pOutputPages.Add sTempPage, sTempPage
      Else
        For lLoop2 = 1 To pOutputPages.count
          If CLng(sTempPage) < CLng(pOutputPages.Item(lLoop2)) Then
            pOutputPages.Add sTempPage, sTempPage, lLoop2
            Exit For
          End If
          If lLoop2 = pOutputPages.count Then
            pOutputPages.Add sTempPage, sTempPage
          End If
        Next lLoop2
      End If
    Else
      For lLoop2 = 1 To pOutputPages.count
        If lLoop2 = 1 Then
          sOutputPages = pOutputPages.Item(lLoop2)
        Else
          sOutputPages = sOutputPages & ", " & pOutputPages.Item(lLoop2)
        End If
      Next lLoop2
      Print #1, sOutput & ": " & sOutputPages
      sOutput = sTempVal
      Set pOutputPages = New Collection
      pOutputPages.Add sTempPage, sTempPage
      sPrev = sTempVal
    End If
    If lLoop = m_pMasterColl.count Then
      For lLoop2 = 1 To pOutputPages.count
        If lLoop2 = 1 Then
          sOutputPages = pOutputPages.Item(lLoop2)
        Else
          sOutputPages = sOutputPages & ", " & pOutputPages.Item(lLoop2)
        End If
      Next lLoop2
      Print #1, sOutput & ": " & sOutputPages
    End If
  Next lLoop
  Close #1
  
  m_pApp.StatusBar.ProgressBar.Hide
  Screen.MousePointer = vbNormal
  Unload Me
  
  Exit Sub
ErrHand:
  Screen.MousePointer = vbNormal
  MsgBox "cmdOK_Click - " & Erl & " - " & Err.Description
End Sub

Private Sub AddPageToMasterCollection(pColl As Collection, sPageId As String)
On Error GoTo ErrHand:
  Dim lLoop As Long, sValue As String, lLoop2 As Long, lStart As Long
  lStart = 1
  If m_pMasterColl.count = 0 Then
    For lLoop = 1 To pColl.count
      sValue = pColl.Item(lLoop) & "-$$$$-" & sPageId
      m_pMasterColl.Add sValue, sValue
    Next lLoop
  Else
    For lLoop = 1 To pColl.count
      sValue = pColl.Item(lLoop) & "-$$$$-" & sPageId
      For lLoop2 = lStart To m_pMasterColl.count
        If sValue < m_pMasterColl.Item(lLoop2) Then
          m_pMasterColl.Add sValue, sValue, lLoop2
          lStart = lLoop2
          Exit For
        End If
        If lLoop2 = m_pMasterColl.count Then
          m_pMasterColl.Add sValue, sValue
          lStart = lLoop2
        End If
      Next lLoop2
    Next lLoop
  End If

  Exit Sub
ErrHand:
  MsgBox "AddPageToMasterCollection - " & Erl & " - " & Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo ErrHand:
  Dim pDoc As IMxDocument, pMap As IMap, lLoop As Long
  Dim pFeatLayer As IFeatureLayer
  Dim pMapBook As INWDSMapBook
  Dim pSeriesProps As INWDSMapSeriesProps
  Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
  Set pSeriesProps = pMapBook.ContentItem(0)

  optIndex(0).Value = True
  txtPageNumber.Text = "0"
  
  'Populate the layer list box
  cmbLayer.Clear
  Set pDoc = m_pApp.Document
  Set pMap = pDoc.FocusMap
  For lLoop = 0 To pMap.LayerCount - 1
    If TypeOf pMap.Layer(lLoop) Is IFeatureLayer Then
      Set pFeatLayer = pMap.Layer(lLoop)
      If pFeatLayer.FeatureClass.FeatureType <> esriFTAnnotation And _
       pFeatLayer.FeatureClass.FeatureType <> esriFTDimension And _
       pFeatLayer.FeatureClass.FeatureType <> esriFTCoverageAnnotation Then
        If UCase(pFeatLayer.Name) <> UCase(pSeriesProps.IndexLayerName) Then
          cmbLayer.AddItem pFeatLayer.Name
        End If
      End If
    End If
  Next lLoop
  If cmbLayer.ListCount > 0 Then
    cmbLayer.ListIndex = 0
  End If
  
  'Make sure the wizard stays on top
  TopMost Me
  
  Exit Sub
ErrHand:
  MsgBox "frmCreateIndex_Load - " & Err.Description
End Sub

Private Sub optIndex_Click(Index As Integer)
  If Index = 0 Then
    txtPageNumber.Enabled = False
  Else
    txtPageNumber.Enabled = True
  End If
  CheckSettings
End Sub

Private Sub txtPageNumber_KeyUp(KeyCode As Integer, Shift As Integer)
  If txtPageNumber.Text = "" Then
    cmdOK.Enabled = False
  Else
    If Not IsNumeric(txtPageNumber.Text) Then
      txtPageNumber.Text = "0"
    End If
    CheckSettings
  End If
End Sub

Private Sub CheckSettings()
  If optIndex(0).Value = True Then
    If txtOutput.Text <> "" Then
      cmdOK.Enabled = True
    Else
      cmdOK.Enabled = False
    End If
  Else
    If txtOutput.Text <> "" And txtPageNumber.Text <> "" Then
      cmdOK.Enabled = True
    Else
      cmdOK.Enabled = False
    End If
  End If
End Sub
