VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
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
Private bLoadFlag As Boolean

Private Sub cmdOK_Click()
7:   Unload Me
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo ErrHand:
  Dim pMapBook As IDSMapBook, dLastOutput As Date
  Dim pSeries As IDSMapSeries, lLoop As Long, pPage As IDSMapPage
14:   Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
17:   Set pSeries = pMapBook.ContentItem(0)
18:   Set pPage = pSeries.Page(grdPages.Row - 1)
  
  Select Case grdPages.Col
  Case 1   'Name
22:     grdPages.Text = txtUpdate.Text
23:     pPage.PageName = txtUpdate.Text
  Case 2   'Scale
25:     grdPages.Text = txtUpdate.Text
26:     pPage.PageScale = CDbl(txtUpdate.Text)
  Case 3   'Rotation
28:     grdPages.Text = txtUpdate.Text
29:     pPage.PageRotation = CDbl(txtUpdate.Text)
  Case 4   'Last Output
31:     dLastOutput = Format(txtUpdate.Text, "mm/dd/yyyy")
32:     grdPages.Text = dLastOutput
33:     pPage.LastOutputted = dLastOutput
34:   End Select

  Exit Sub
  
ErrHand:
39:   MsgBox "cmdUpdate_Click - " & Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo ErrHand:
  Dim pMapBook As IDSMapBook
  Dim pSeries As IDSMapSeries, lLoop As Long, pPage As IDSMapPage
46:   Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
49:   Set pSeries = pMapBook.ContentItem(0)
  
  'Load up the Grid
52:   bLoadFlag = True
53:   grdPages.Clear
54:   grdPages.Rows = 1
55:   grdPages.ColWidth(0) = 750
56:   grdPages.ColWidth(1) = 2500
57:   grdPages.ColWidth(2) = 1000
58:   grdPages.ColWidth(3) = 1000
59:   grdPages.ColWidth(4) = 1400
60:   grdPages.Row = 0
61:   grdPages.Col = 0
62:   grdPages.Text = "Number"
63:   grdPages.Col = 1
64:   grdPages.Text = "Tile Name"
65:   grdPages.Col = 2
66:   grdPages.Text = "Scale"
67:   grdPages.Col = 3
68:   grdPages.Text = "Rotation"
69:   grdPages.Col = 4
70:   grdPages.Text = "Last Output"
71:   grdPages.CellAlignment = 1
72:   For lLoop = 0 To pSeries.PageCount - 1
73:     Set pPage = pSeries.Page(lLoop)
74:     grdPages.AddItem lLoop + 1 & Chr(9) & pPage.PageName & Chr(9) & pPage.PageScale & _
     Chr(9) & pPage.PageRotation & Chr(9) & pPage.LastOutputted
76:   Next lLoop
77:   bLoadFlag = False
  
  'Set the Update properties
80:   lblType.Caption = ""
81:   txtUpdate.Text = ""
82:   lblTile.Caption = "Page:"
83:   cmdUpdate.Enabled = False
  
  'Make sure the wizard stays on top
86:   TopMost Me
  
  Exit Sub
  
ErrHand:
91:   MsgBox "frmPageProperties_Load - " & Err.Description
End Sub

Private Sub grdPages_EnterCell()
On Error GoTo ErrHand:
  Dim pMapBook As IDSMapBook, lCol As Long, lRow As Long
  Dim pSeries As IDSMapSeries, lLoop As Long, pPage As IDSMapPage
  
  'Exit the sub if we are loading up the grid
  If bLoadFlag Then Exit Sub
  
102:   Set pMapBook = GetMapBookExtension(m_pApp)
  If pMapBook Is Nothing Then Exit Sub
  
105:   Set pSeries = pMapBook.ContentItem(0)
  
107:   lCol = grdPages.Col
108:   lRow = grdPages.Row
  
110:   If lCol > 0 Then
111:     Set pPage = pSeries.Page(lRow - 1)
112:     lblTile.Caption = "Page: " & pPage.PageName
    Select Case lCol
    Case 1   'Page Name
115:       lblType.Caption = "Name:"
116:       txtUpdate.Text = pPage.PageName
    Case 2   'Scale
118:       lblType.Caption = "Scale:"
119:       txtUpdate.Text = CStr(pPage.PageScale)
    Case 3   'Rotation
121:       lblType.Caption = "Rotation:"
122:       txtUpdate.Text = CStr(pPage.PageRotation)
    Case 4   'Last Output
124:       lblType.Caption = "Last Output:"
125:       txtUpdate.Text = CStr(pPage.LastOutputted)
126:     End Select
127:     cmdUpdate.Enabled = True
128:   Else
129:     lblType.Caption = ""
130:     txtUpdate.Text = ""
131:     lblTile.Caption = "Page:"
132:     cmdUpdate.Enabled = False
133:   End If

  Exit Sub
ErrHand:
137:   MsgBox "grdPages_EnterCell - " & Err.Description
End Sub

Private Sub txtUpdate_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrHand:

143:   cmdUpdate.Enabled = False
  
  Select Case grdPages.Col
  Case 1   'Name
147:     If txtUpdate.Text <> "" Then
148:       cmdUpdate.Enabled = True
149:     End If
  Case 2, 3  'Scale and Rotation
151:     If txtUpdate.Text <> "" Then
152:       If IsNumeric(txtUpdate.Text) Then
153:         cmdUpdate.Enabled = True
154:       End If
155:     End If
  Case 4    'Last output
157:     If txtUpdate.Text <> "" Then
158:       If IsDate(txtUpdate.Text) Then
159:         cmdUpdate.Enabled = True
160:       End If
161:     End If
162:   End Select

  Exit Sub
ErrHand:
166:   MsgBox "txtUpdate_KeyUp - " & Err.Description
End Sub
