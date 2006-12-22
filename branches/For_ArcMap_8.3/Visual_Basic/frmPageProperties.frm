VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPageProperties 
   Caption         =   "Page Properties"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtUpdate 
      Height          =   315
      Left            =   4080
      TabIndex        =   5
      Top             =   3990
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   345
      Left            =   5280
      TabIndex        =   2
      Top             =   3990
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   345
      Left            =   6300
      TabIndex        =   1
      Top             =   4500
      Width           =   1125
   End
   Begin MSFlexGridLib.MSFlexGrid grdPages 
      Height          =   3795
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   6694
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
   End
   Begin VB.Label lblTile 
      Caption         =   "Tile:"
      Height          =   225
      Left            =   180
      TabIndex        =   4
      Top             =   4050
      Width           =   2625
   End
   Begin VB.Label lblType 
      Alignment       =   1  'Right Justify
      Caption         =   "Scale:"
      Height          =   225
      Left            =   2910
      TabIndex        =   3
      Top             =   4050
      Width           =   1035
   End
End
Attribute VB_Name = "frmPageProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_pApp As IApplication
Private bLoadFlag As Boolean

Private Sub cmdOk_Click()
  Unload Me
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo ErrHand:
  Dim pMapBook As IDSMapBook, dLastOutput As Date
  Dim pSeries As IDSMapSeries, lLoop As Long, pPage As IDSMapPage
  Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
  Set pSeries = pMapBook.ContentItem(0)
  Set pPage = pSeries.Page(grdPages.Row - 1)
  
  Select Case grdPages.Col
  Case 1   'Name
    grdPages.Text = txtUpdate.Text
    pPage.PageName = txtUpdate.Text
  Case 2   'Scale
    grdPages.Text = txtUpdate.Text
    pPage.PageScale = CDbl(txtUpdate.Text)
  Case 3   'Rotation
    grdPages.Text = txtUpdate.Text
    pPage.PageRotation = CDbl(txtUpdate.Text)
  Case 4   'Last Output
    dLastOutput = Format(txtUpdate.Text, "mm/dd/yyyy")
    grdPages.Text = dLastOutput
    pPage.LastOutputted = dLastOutput
  End Select

  Exit Sub
  
ErrHand:
  MsgBox "cmdUpdate_Click - " & Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo ErrHand:
  Dim pMapBook As IDSMapBook
  Dim pSeries As IDSMapSeries, lLoop As Long, pPage As IDSMapPage
  Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
  Set pSeries = pMapBook.ContentItem(0)
  
  'Load up the Grid
  bLoadFlag = True
  grdPages.Clear
  grdPages.Rows = 1
  grdPages.ColWidth(0) = 750
  grdPages.ColWidth(1) = 2500
  grdPages.ColWidth(2) = 1000
  grdPages.ColWidth(3) = 1000
  grdPages.ColWidth(4) = 1400
  grdPages.Row = 0
  grdPages.Col = 0
  grdPages.Text = "Number"
  grdPages.Col = 1
  grdPages.Text = "Tile Name"
  grdPages.Col = 2
  grdPages.Text = "Scale"
  grdPages.Col = 3
  grdPages.Text = "Rotation"
  grdPages.Col = 4
  grdPages.Text = "Last Output"
  grdPages.CellAlignment = 1
  For lLoop = 0 To pSeries.PageCount - 1
    Set pPage = pSeries.Page(lLoop)
    grdPages.AddItem lLoop + 1 & Chr(9) & pPage.PageName & Chr(9) & pPage.PageScale & _
     Chr(9) & pPage.PageRotation & Chr(9) & pPage.LastOutputted
  Next lLoop
  bLoadFlag = False
  
  'Set the Update properties
  lblType.Caption = ""
  txtUpdate.Text = ""
  lblTile.Caption = "Page:"
  cmdUpdate.Enabled = False
  
  'Make sure the wizard stays on top
  TopMost Me
  
  Exit Sub
  
ErrHand:
  MsgBox "frmPageProperties_Load - " & Err.Description
End Sub

Private Sub grdPages_EnterCell()
On Error GoTo ErrHand:
  Dim pMapBook As IDSMapBook, lCol As Long, lRow As Long
  Dim pSeries As IDSMapSeries, lLoop As Long, pPage As IDSMapPage
  
  'Exit the sub if we are loading up the grid
  If bLoadFlag Then Exit Sub
  
  Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
  Set pSeries = pMapBook.ContentItem(0)
  
  lCol = grdPages.Col
  lRow = grdPages.Row
  
  If lCol > 0 Then
    Set pPage = pSeries.Page(lRow - 1)
    lblTile.Caption = "Page: " & pPage.PageName
    Select Case lCol
    Case 1   'Page Name
      lblType.Caption = "Name:"
      txtUpdate.Text = pPage.PageName
    Case 2   'Scale
      lblType.Caption = "Scale:"
      txtUpdate.Text = CStr(pPage.PageScale)
    Case 3   'Rotation
      lblType.Caption = "Rotation:"
      txtUpdate.Text = CStr(pPage.PageRotation)
    Case 4   'Last Output
      lblType.Caption = "Last Output:"
      txtUpdate.Text = CStr(pPage.LastOutputted)
    End Select
    cmdUpdate.Enabled = True
  Else
    lblType.Caption = ""
    txtUpdate.Text = ""
    lblTile.Caption = "Page:"
    cmdUpdate.Enabled = False
  End If

  Exit Sub
ErrHand:
  MsgBox "grdPages_EnterCell - " & Err.Description
End Sub

Private Sub txtUpdate_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrHand:

  cmdUpdate.Enabled = False
  
  Select Case grdPages.Col
  Case 1   'Name
    If txtUpdate.Text <> "" Then
      cmdUpdate.Enabled = True
    End If
  Case 2, 3  'Scale and Rotation
    If txtUpdate.Text <> "" Then
      If IsNumeric(txtUpdate.Text) Then
        cmdUpdate.Enabled = True
      End If
    End If
  Case 4    'Last output
    If txtUpdate.Text <> "" Then
      If IsDate(txtUpdate.Text) Then
        cmdUpdate.Enabled = True
      End If
    End If
  End Select

  Exit Sub
ErrHand:
  MsgBox "txtUpdate_KeyUp - " & Err.Description
End Sub
