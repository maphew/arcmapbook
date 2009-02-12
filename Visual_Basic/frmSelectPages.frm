VERSION 5.00
Begin VB.Form frmSelectPages 
   Caption         =   "Select/Enable Pages"
   ClientHeight    =   3708
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   5232
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3708
   ScaleWidth      =   5232
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   4020
      TabIndex        =   2
      Top             =   3330
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   345
      Left            =   2880
      TabIndex        =   1
      Top             =   3330
      Width           =   1125
   End
   Begin VB.Frame fraSelection 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3165
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   5025
      Begin VB.ComboBox cmbScale 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtScale 
         Height          =   285
         Left            =   2760
         TabIndex        =   12
         Top             =   2670
         Width           =   975
      End
      Begin VB.OptionButton optSelection 
         Caption         =   "Select with scale "
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   11
         Top             =   2700
         Width           =   1605
      End
      Begin VB.TextBox txtBefore 
         Height          =   285
         Left            =   1410
         TabIndex        =   10
         Top             =   2250
         Width           =   975
      End
      Begin VB.TextBox txtAfter 
         Height          =   285
         Left            =   2940
         TabIndex        =   9
         Top             =   2250
         Width           =   975
      End
      Begin VB.OptionButton optSelection 
         Caption         =   "Select by date last printed/exported (use format 01/01/02):"
         Height          =   345
         Index           =   2
         Left            =   360
         TabIndex        =   6
         Top             =   1770
         Width           =   2925
      End
      Begin VB.OptionButton optSelection 
         Caption         =   "Unselect all"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   1260
         Width           =   1695
      End
      Begin VB.OptionButton optSelection 
         Caption         =   "Select all"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   750
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "before:"
         Height          =   195
         Index           =   1
         Left            =   870
         TabIndex        =   8
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "after:"
         Height          =   195
         Index           =   0
         Left            =   2550
         TabIndex        =   7
         Top             =   2280
         Width           =   405
      End
      Begin VB.Label Label1 
         Caption         =   $"frmSelectPages.frx":0000
         Height          =   645
         Left            =   60
         TabIndex        =   3
         Top             =   60
         Width           =   4905
      End
   End
End
Attribute VB_Name = "frmSelectPages"
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
Private m_pMapSeries As INWDSMapSeries

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrHand:
  Dim lLoop As Long, pMapPage As INWDSMapPage
  'Check to see if a MapSeries already exists
  
  If optSelection(0).Value Then      'Select all
    SelectAllPages True
  ElseIf optSelection(1).Value Then  'Unselect all
    SelectAllPages False
  ElseIf optSelection(2).Value Then  'Select by date last printed/exported
    SelectByDate
  ElseIf optSelection(3).Value Then  'Select by scale value
    SelectByScale
  End If

  Unload Me
  
  Exit Sub
ErrHand:
  MsgBox "frmSelectPages_Click - " & Err.Description
End Sub

Private Sub SelectByDate()
On Error GoTo ErrHand:
  Dim lLoop As Long, pNode As Node, dDate As Date
  Dim pPage As INWDSMapPage
  
  'Select pages by date last printed/exported
  For lLoop = 0 To m_pMapSeries.PageCount - 1
    Set pPage = m_pMapSeries.Page(lLoop)
    Set pNode = g_pFrmMapSeries.tvwMapBook.Nodes.Item(lLoop + 3)
    dDate = m_pMapSeries.Page(lLoop).LastOutputted
    If IsDate(txtBefore.Text) And txtAfter.Text = "" Then
      If dDate < txtBefore.Text Or dDate = #1/1/1900# Then
        pPage.EnablePage = True
        pNode.Image = 5
      Else
        pPage.EnablePage = False
        pNode.Image = 6
      End If
    ElseIf IsDate(txtBefore.Text) And IsDate(txtAfter.Text) Then
      If dDate >= txtBefore.Text And dDate <= txtAfter.Text Then
        pPage.EnablePage = True
        pNode.Image = 5
      Else
        pPage.EnablePage = False
        pNode.Image = 6
      End If
    Else
      If dDate > txtAfter.Text Then
        pPage.EnablePage = True
        pNode.Image = 5
      Else
        pPage.EnablePage = False
        pNode.Image = 6
      End If
    End If
  Next lLoop

  Exit Sub

ErrHand:
  MsgBox "SelectByDate - " & Err.Description
End Sub

Private Sub SelectByScale()
On Error GoTo ErrHand:
  Dim lLoop As Long, pNode As Node, dScale As Double
  Dim pPage As INWDSMapPage, sExp As String
  
  'Select pages by Scale
  For lLoop = 0 To m_pMapSeries.PageCount - 1
    Set pPage = m_pMapSeries.Page(lLoop)
    Set pNode = g_pFrmMapSeries.tvwMapBook.Nodes.Item(lLoop + 3)
    dScale = m_pMapSeries.Page(lLoop).PageScale
    sExp = CStr(dScale) & " " & cmbScale.Text & " " & txtScale.Text
    If sExp Then
      pPage.EnablePage = True
      pNode.Image = 5
    Else
      pPage.EnablePage = False
      pNode.Image = 6
    End If
  Next lLoop

  Exit Sub

ErrHand:
  MsgBox "SelectByScale - " & Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo ErrHand:
  Dim pMapBook As INWDSMapBook, pOpts As INWDSMapSeriesOptions
  Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
  Set m_pMapSeries = pMapBook.ContentItem(0)
  Set pOpts = m_pMapSeries
  If pOpts.ExtentType = 2 Then
    optSelection(3).Enabled = True
  Else
    optSelection(3).Enabled = False
  End If

  optSelection(0).Value = True
  
  cmbScale.Clear
  cmbScale.AddItem "="
  cmbScale.AddItem "<>"
  cmbScale.AddItem ">"
  cmbScale.AddItem ">="
  cmbScale.AddItem "<"
  cmbScale.AddItem "<="
  cmbScale.Text = "="
  
  Exit Sub
ErrHand:
  MsgBox "frmSelectPages_Load - " & Err.Description
End Sub

Private Sub SelectAllPages(bValue As Boolean)
On Error GoTo ErrHand:
  Dim lLoop As Long, pNode As Node
  
  'Loop through the pages turning them on or off
  For lLoop = 0 To m_pMapSeries.PageCount - 1
    Set pNode = g_pFrmMapSeries.tvwMapBook.Nodes.Item(lLoop + 3)
    m_pMapSeries.Page(lLoop).EnablePage = bValue
    If bValue Then
      pNode.Image = 5
    Else
      pNode.Image = 6
    End If
  Next lLoop
  
  Exit Sub
ErrHand:
  MsgBox "SelectAllPages - " & Err.Description
End Sub

Private Sub optSelection_Click(Index As Integer)
  Select Case Index
  Case 0    'Select all
    cmdOK.Enabled = True
  Case 1    'Unselect all
    cmdOK.Enabled = True
  Case 2    'Select by date last printed/exported
    If DateCheck Then
      cmdOK.Enabled = True
    Else
      cmdOK.Enabled = False
    End If
  Case 3    'Select by scale
    If ScaleCheck Then
      cmdOK.Enabled = True
    Else
      cmdOK.Enabled = False
    End If
  End Select
End Sub

Private Sub txtAfter_KeyUp(KeyCode As Integer, Shift As Integer)
  If DateCheck Then
    cmdOK.Enabled = True
  Else
    cmdOK.Enabled = False
  End If
End Sub

Private Sub txtBefore_KeyUp(KeyCode As Integer, Shift As Integer)
  If DateCheck Then
    cmdOK.Enabled = True
  Else
    cmdOK.Enabled = False
  End If
End Sub

Private Sub txtScale_KeyUp(KeyCode As Integer, Shift As Integer)
  If Not IsNumeric(txtScale.Text) Then
    txtScale.Text = ""
  End If
  If ScaleCheck Then
    cmdOK.Enabled = True
  Else
    cmdOK.Enabled = False
  End If
End Sub

Private Function ScaleCheck() As Boolean
  ScaleCheck = False
  If txtScale.Text <> "" Then
    If CDbl(txtScale.Text) >= 0 Then
      ScaleCheck = True
    End If
  End If
End Function

Private Function DateCheck() As Boolean
  If IsDate(txtBefore.Text) And txtAfter.Text = "" Then
    DateCheck = True
  ElseIf IsDate(txtBefore.Text) And IsDate(txtAfter.Text) Then
    DateCheck = True
  ElseIf txtBefore.Text = "" And IsDate(txtAfter.Text) Then
    DateCheck = True
  Else
    DateCheck = False
  End If
End Function
