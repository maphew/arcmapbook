VERSION 5.00
Begin VB.Form frmBubbleProperties 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Circular Inset Map Properties"
   ClientHeight    =   7392
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7392
   ScaleWidth      =   5640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUncheckAll 
      Caption         =   "&Uncheck All"
      Height          =   375
      Left            =   1560
      TabIndex        =   17
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdCheckAll 
      Caption         =   "Check &All"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CheckBox chkEnableHeightWidth 
      Caption         =   "Separate Height and Width Radii"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   3015
   End
   Begin VB.TextBox txtBubbleRadiusX 
      Height          =   285
      Left            =   3960
      TabIndex        =   9
      Text            =   "txtBubbleRadiusX"
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox txtRadiusX 
      Height          =   285
      Left            =   720
      TabIndex        =   8
      Text            =   "txtRadiusX"
      Top             =   2760
      Width           =   1245
   End
   Begin VB.TextBox txtDestinationY 
      Height          =   285
      Left            =   3960
      TabIndex        =   13
      Text            =   "txtDestinationY"
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox txtOriginY 
      Height          =   285
      Left            =   3960
      TabIndex        =   11
      Text            =   "txtOriginY"
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox txtDestinationX 
      Height          =   285
      Left            =   1560
      TabIndex        =   12
      Text            =   "txtDestinationX"
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox txtOriginX 
      Height          =   285
      Left            =   1560
      TabIndex        =   10
      Text            =   "txtOriginX"
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   19
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2400
      TabIndex        =   18
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox txtBubbleRadiusY 
      Height          =   285
      Left            =   3960
      TabIndex        =   7
      Text            =   "txtBubbleRadiusY"
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox txtScaleFactor 
      Height          =   285
      Left            =   2280
      TabIndex        =   6
      Text            =   "txtScaleFactor"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtRadiusY 
      Height          =   285
      Left            =   720
      TabIndex        =   5
      Text            =   "txtRadiusY"
      Top             =   2280
      Width           =   1245
   End
   Begin VB.Frame fraAdvancedSizing 
      Caption         =   "Advanced Sizing Options"
      Height          =   975
      Left            =   120
      TabIndex        =   21
      Top             =   630
      Width           =   5415
      Begin VB.OptionButton optFixedBubbleRadius 
         Caption         =   "Fixed Bubble Radius"
         Height          =   255
         Left            =   3360
         TabIndex        =   3
         Top             =   600
         Value           =   -1  'True
         Width           =   1890
      End
      Begin VB.OptionButton optFixedScaleFactor 
         Caption         =   "Fixed Scale Factor"
         Height          =   255
         Left            =   1560
         TabIndex        =   2
         Top             =   600
         Width           =   1695
      End
      Begin VB.CheckBox chkFixedValue 
         Caption         =   "Update calculation based on one fixed value"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   3855
      End
      Begin VB.OptionButton optFixedRadius 
         Caption         =   "Fixed Radius"
         Height          =   255
         Left            =   135
         TabIndex        =   1
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.ListBox lstLayers 
      Height          =   1776
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   15
      Top             =   4440
      Width           =   5295
   End
   Begin VB.Label lblWidth 
      Caption         =   "Width"
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblHeight 
      Caption         =   "Height"
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Line lineScaleTop 
      Visible         =   0   'False
      X1              =   2280
      X2              =   3480
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line lineScaleBottom 
      Visible         =   0   'False
      X1              =   3480
      X2              =   2280
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label lblEquals 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   36
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label Label9 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   35
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label lblX 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   34
      Top             =   2520
      Width           =   255
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   5520
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000015&
      BorderWidth     =   2
      X1              =   0
      X2              =   5520
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   5520
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label7 
      Caption         =   "Destination :"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Y -"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3465
      TabIndex        =   32
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "X -"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   31
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "Y -"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3465
      TabIndex        =   30
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "X -"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   29
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label lblOrigin 
      Caption         =   "Origin :"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   3240
      Width           =   615
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   5520
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000015&
      BorderWidth     =   2
      X1              =   0
      X2              =   5520
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Label Label5 
      Caption         =   "Layers to be displayed in circular detail inset."
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   4200
      Width           =   3495
   End
   Begin VB.Label lblBubbleRadius 
      Caption         =   "Bubble Radius"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   26
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   25
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label lblScaleFactor 
      Caption         =   "Scale Factor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   24
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   23
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label llbRadius 
      Caption         =   "Radius"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   22
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblBubbleId 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "116"
      Height          =   255
      Left            =   1560
      TabIndex        =   20
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Bubble ID : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   240
      Width           =   1215
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000015&
      BorderWidth     =   2
      X1              =   0
      X2              =   5520
      Y1              =   3120
      Y2              =   3120
   End
End
Attribute VB_Name = "frmBubbleProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private m_pApp As IApplication
Private m_lBubbleId As Long
Private m_dRadius As Double
Private m_dRadiusX As Double
Private m_dScaleFactor As Double
Private m_dBubbleRadius As Double
Private m_dBubbleRadiusX As Double
Private m_dOriginX As Double
Private m_dOriginY As Double
Private m_dDestinationX As Double
Private m_dDestinationY As Double
Private m_dSourceHeight As Double
Private m_dSourceWidth As Double
Private m_bHeightWidthIsEnabled As Boolean


  'variables used for dynamically updating graphics
  'based on what the user inputs
Private m_sDataFrameName As String
Private m_pMxDoc As IMxDocument
Private m_pPnt1 As IPoint
Private m_pPnt2 As IPoint
Private m_pPnt3 As IPoint
Private m_pPnt4 As IPoint
Private m_pGraCont As IGraphicsContainer
Private m_pElemFillShp1 As IFillShapeElement
Private m_pElemFillShp2 As IFillShapeElement
Private m_pElemLine As ILineElement
Private m_pAV As IActiveView
Private m_pScrD As IScreenDisplay
Private m_pNewCircFeedback As INewCircleFeedback
Private m_p2ndCircFeedback As INewCircleFeedback
Private m_pLineFeedback As INewLineFeedback
Private m_bCancelled As Boolean

Private m_pGXDlg As IGxDialog
Private m_bInitializing As Boolean
Private m_bLocked As Boolean

Const c_sModuleFileName As String = "frmBubbleProperties.frm"



Public Property Get WasCancelled() As Boolean
48:   WasCancelled = m_bCancelled
End Property



Private Sub chkEnableHeightWidth_Click()
54:   If chkEnableHeightWidth.Value = vbChecked Then
                                                  'make the width radius and bubbleradius
                                                  'visible
57:     txtRadiusX.Visible = True
58:     txtBubbleRadiusX.Visible = True
                                                  'move the scalefactor text box down
60:     txtScaleFactor.Top = lineScaleBottom.Y1
                                                  'make the height and width label visible
62:     lblHeight.Visible = True
63:     lblWidth.Visible = True
64:     lblX.Top = lineScaleBottom.Y1
65:     lblEquals.Visible = True
66:   Else
67:     txtRadiusX.Visible = False
68:     txtBubbleRadiusX.Visible = False
69:     txtScaleFactor.Top = lineScaleTop.Y1
70:     lblHeight.Visible = False
71:     lblWidth.Visible = False
72:     lblX.Top = lineScaleTop.Y1
73:     lblEquals.Visible = False
74:   End If
End Sub

Private Sub chkFixedValue_Click()
  On Error GoTo ErrorHandler

80:   If chkFixedValue.Value = vbChecked Then
81:     optFixedRadius.Enabled = True
82:     optFixedScaleFactor.Enabled = True
83:     optFixedBubbleRadius.Enabled = True
                                                  'fixed radius
85:     txtRadiusY.BackColor = vbWindowBackground
86:     txtScaleFactor.BackColor = vbWindowBackground
87:     txtBubbleRadiusY.BackColor = vbWindowBackground
88:     If optFixedRadius Then
89:       txtRadiusY.Enabled = False
90:       txtScaleFactor.Enabled = True
91:       txtBubbleRadiusY.Enabled = True
92:       txtRadiusY.BackColor = vbInactiveBorder
                                                  'fixed scale factor
94:     ElseIf optFixedScaleFactor Then
95:       txtRadiusY.Enabled = True
96:       txtScaleFactor.Enabled = False
97:       txtBubbleRadiusY.Enabled = True
98:       txtScaleFactor.BackColor = vbInactiveBorder
                                                  'fixed bubble radius
100:     ElseIf optFixedBubbleRadius Then
101:       txtRadiusY.Enabled = True
102:       txtScaleFactor.Enabled = True
103:       txtBubbleRadiusY.Enabled = False
104:       txtBubbleRadiusY.BackColor = vbInactiveBorder
105:     End If
106:   Else
107:     optFixedRadius.Enabled = False
108:     optFixedScaleFactor.Enabled = False
109:     optFixedBubbleRadius.Enabled = False
110:     txtRadiusY.BackColor = vbWindowBackground
111:     txtRadiusX.BackColor = vbWindowBackground
112:     txtScaleFactor.BackColor = vbWindowBackground
113:     txtBubbleRadiusY.BackColor = vbWindowBackground
114:     txtBubbleRadiusX.BackColor = vbWindowBackground
115:     txtRadiusY.Enabled = True
116:     txtScaleFactor.Enabled = True
117:     txtBubbleRadiusY.Enabled = True
118:     txtRadiusX.Enabled = True
119:     txtBubbleRadiusX.Enabled = True
120:   End If

  Exit Sub
ErrorHandler:
  HandleError True, "chkFixedValue_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub





  





    '       all variables required to be displayed in
    '       frmBubbleProperties
    '         - bubble id       '--- m_lBubbleId
    '         - radius          '--- m_dRadius
    '         - scale factor    '--- m_dScaleFactor
    '         - bubble radius   '--- m_dBubbleRadius
    '         - origin x/y      '--- m_dOriginX, m_dOriginY
    '         - destination x/y '--- m_dDestinationX, m_dDestinationY
    '         - destination height/width '--- m_dSourceHeight, m_dSourceWidth
    '         - application reference, (to add all layers) '-- m_pApp

'----------
'attributes stored in the data -
'scale, radius, xdest, ydest, originx, originy, bubbleid
'----------
'new attributes not in data source:
' destHeight, destWidth -- might not be set as an attribute field, but
'         instead be a feature of the envelope of the origin circle.

Private Sub cmdCancel_Click()
                                                  'clean up all the temporary graphics
158:   EraseGraphics m_pElemLine, m_pElemFillShp1, m_pElemFillShp2, m_pApp

160:   m_bCancelled = True
161:   Me.Hide
End Sub



Public Property Set Application(pApp As IApplication)
  On Error GoTo ErrorHandler

169:   Set m_pApp = pApp
170:   If Not m_pApp Is Nothing Then
171:     Set m_pMxDoc = m_pApp.Document
172:     Set m_pAV = m_pMxDoc.ActiveView
173:     Set m_pScrD = m_pAV.ScreenDisplay
174:   End If

  Exit Property
ErrorHandler:
  HandleError True, "Application " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property


Public Property Let ActiveView(RHS As IActiveView)

End Property

Public Property Get Application() As IApplication
  On Error GoTo ErrorHandler

189:   Set Application = m_pApp

  Exit Property
ErrorHandler:
  HandleError True, "Application " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property



Public Property Let BubbleID(RHS As Long)
199:   m_lBubbleId = RHS
200:   lblBubbleId.Caption = RHS
End Property
Public Property Get BubbleID() As Long
203:   BubbleID = m_lBubbleId
End Property



Public Property Let Radius(RHS As Double)
209:   m_dRadius = RHS
210:   txtRadiusY.Text = Format(RHS, "#########0.000")
End Property
Public Property Get Radius() As Double
213:   Radius = m_dRadius
End Property



Public Property Let ScaleFactor(dScale As Double)
219:   m_dScaleFactor = dScale
220:   txtScaleFactor.Text = Format(dScale, "#########0.000")
End Property
Public Property Get ScaleFactor() As Double
223:   ScaleFactor = m_dScaleFactor
End Property



Public Property Let BubbleRadius(RHS As Double)
229:   m_dBubbleRadius = RHS
230:   txtBubbleRadiusY.Text = Format(RHS, "#########0.000")
231:   txtBubbleRadiusX.Text = txtBubbleRadiusY.Text
End Property
Public Property Get BubbleRadius() As Double
234:   BubbleRadius = m_dBubbleRadius
End Property
Public Property Let BubbleRadiusX(RHS As Double)
237:   m_dBubbleRadiusX = RHS
End Property
Public Property Get BubbleRadiusX() As Double
240:   BubbleRadiusX = m_dBubbleRadiusX
End Property


Public Property Let DataFrameName(RHS As String)
245:   m_sDataFrameName = RHS
End Property
Public Property Get DataFrameName() As String
248:   DataFrameName = m_sDataFrameName
End Property




Public Property Let OriginX(RHS As Double)
255:   m_dOriginX = RHS
256:   txtOriginX.Text = Format(RHS, "#########0.000")
End Property
Public Property Get OriginX() As Double
259:   OriginX = m_dOriginX
End Property
Public Property Let OriginY(RHS As Double)
262:   m_dOriginY = RHS
263:   txtOriginY.Text = Format(RHS, "#########0.000")
End Property
Public Property Get OriginY() As Double
266:   OriginY = m_dOriginY
End Property





Public Property Let DestinationX(RHS As Double)
274:   m_dDestinationX = RHS
275:   txtDestinationX.Text = Format(RHS, "#########0.000")
End Property
Public Property Get DestinationX() As Double
278:   DestinationX = m_dDestinationX
End Property
Public Property Let DestinationY(RHS As Double)
281:   m_dDestinationY = RHS
282:   txtDestinationY.Text = Format(RHS, "#########0.000")
End Property
Public Property Get DestinationY() As Double
285:   DestinationY = m_dDestinationY
End Property





Public Property Let SourceHeight(RHS As Double)
                                                  'update the height
294:   m_dSourceHeight = RHS
                                                  'update the radius
296:   m_dRadius = RHS / 2
297:   txtRadiusY.Text = Format((RHS / 2), "#########0.00")
298:   If Not m_bHeightWidthIsEnabled Then
299:     m_dSourceWidth = RHS
300:     m_dRadiusX = (RHS / 2)
301:   End If
End Property
Public Property Get SourceHeight() As Double
304:   m_dSourceHeight = m_dRadius * 2
305:   SourceHeight = m_dSourceHeight
End Property
Public Property Let SourceWidth(RHS As Double)
308:   m_dSourceWidth = RHS
  
310:   txtRadiusX.Text = Format((RHS / 2), "#########0.00")
311:   If Not m_bHeightWidthIsEnabled Then
312:     m_dSourceHeight = RHS
313:     m_dRadius = (RHS / 2)
314:   End If
End Property
Public Property Get SourceWidth() As Double
317:   m_dSourceWidth = m_dRadiusX * 2
318:   SourceWidth = m_dSourceWidth
End Property





Public Property Let HeightWidthIsEnabled(RHS As Boolean)
326:   m_bHeightWidthIsEnabled = RHS
End Property
Public Property Get HeightWidthIsEnabled() As Boolean
329:   HeightWidthIsEnabled = m_bHeightWidthIsEnabled
End Property




Public Property Let Initializing(RHS As Boolean)
336:   m_bInitializing = RHS
End Property



'dynamically generate a string of those
'layers in the list of layers that happen
'to be selected
Public Property Get Layers() As String
  On Error GoTo ErrorHandler

  Dim i As Long, lCount As Long, sReturnVal As String, bFirstTime As Boolean
  
349:   sReturnVal = ""
350:   bFirstTime = True
351:   With lstLayers
352:     lCount = .ListCount
353:     For i = 0 To lCount - 1
354:       If .Selected(i) Then
355:         If bFirstTime Then
356:           bFirstTime = False
357:         Else
358:           sReturnVal = sReturnVal & ","
359:         End If
360:         sReturnVal = sReturnVal & .List(i)
361:       End If
362:     Next i
363:   End With
364:   Layers = sReturnVal

  Exit Property
ErrorHandler:
  HandleError True, "Layers " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property



Public Property Let Layers(sLayers As String)
  On Error GoTo ErrorHandler


  Dim lLyrUBound As Long, sLyrArr() As String, i As Long
  Dim lListCount As Long, sLayer As String, lIndex As Long
  
380:   With lstLayers
381:     lListCount = lstLayers.ListCount
382:     For i = 0 To lListCount - 1
383:       .Selected(i) = False
384:       .ItemData(i) = vbUnchecked
385:     Next i
386:   End With
  
388:   sLyrArr = Split(sLayers, ",", -1, vbTextCompare)
389:   lLyrUBound = UBound(sLyrArr)

391:   For i = 0 To lLyrUBound
392:     sLayer = sLyrArr(i)
393:     sLayer = Trim$(sLayer)
394:     If Len(sLayer) > 0 Then
395:       lIndex = FindControlString(lstLayers, sLayer, -1, True)
396:       If lIndex >= 0 Then
397:         lstLayers.Selected(lIndex) = True
398:       End If
399:     End If
400:   Next i

  Exit Property
ErrorHandler:
  HandleError True, "Layers " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property








'''''''''''properties used for updating graphics

Public Property Set PointOriginCenter(RHS As IPoint)
417:   Set m_pPnt1 = RHS
End Property
Public Property Get PointOriginCenter() As IPoint
420:   Set PointOriginCenter = m_pPnt1
End Property
Public Property Set PointOriginEdge(RHS As IPoint)
423:   Set m_pPnt2 = RHS
End Property
Public Property Get PointOriginEdge() As IPoint
426:   Set PointOriginEdge = m_pPnt2
End Property


Public Property Set PointDestCenter(RHS As IPoint)
431:   Set m_pPnt3 = RHS
End Property
Public Property Get PointDestCenter() As IPoint
434:   Set PointDestCenter = m_pPnt3
End Property
Public Property Set PointDestEdge(RHS As IPoint)
437:   Set m_pPnt4 = RHS
End Property
Public Property Get PointDestEdge() As IPoint
440:   Set PointDestEdge = m_pPnt4
End Property


Public Property Set GraphicsContainer(RHS As IGraphicsContainer)
445:   Set m_pGraCont = RHS
End Property
Public Property Get GraphicsContainer() As IGraphicsContainer
448:   Set GraphicsContainer = m_pGraCont
End Property


Public Property Set OriginShape(RHS As IFillShapeElement)
453:   Set m_pElemFillShp1 = RHS
End Property
Public Property Get OriginShape() As IFillShapeElement
456:   Set OriginShape = m_pElemFillShp1
End Property


Public Property Set DestinationShape(RHS As IFillShapeElement)
461:   Set m_pElemFillShp2 = RHS
End Property
Public Property Get DestinationShape() As IFillShapeElement
464:   Set DestinationShape = m_pElemFillShp2
End Property


Public Property Set BetweenLineElement(RHS As ILineElement)
469:   Set m_pElemLine = RHS
End Property
Public Property Get BetweenLineElement() As ILineElement
472:   Set BetweenLineElement = m_pElemLine
End Property


Public Property Set OriginCircFeedback(RHS As INewCircleFeedback)
477:   Set m_pNewCircFeedback = RHS
End Property


Public Property Set DestCircleFeedback(RHS As INewCircleFeedback)
482:   Set m_p2ndCircFeedback = RHS
End Property


Public Property Set BetweenLineFeedback(RHS As INewLineFeedback)
487:   Set m_pLineFeedback = RHS
End Property

'''''''''''end of properties for maintaining graphics while displaying this form








Private Sub cmdCheckAll_Click()
  Dim i As Long
  
502:   With lstLayers
503:     For i = 0 To .ListCount - 1
504:       .ItemData(i) = 1
505:       .Selected(i) = True
506:     Next i
507:   End With
End Sub

Private Sub cmdOK_Click()
  On Error GoTo ErrorHandler

513:   m_bCancelled = False
514:   Me.Hide

                                                  'clean up all the temporary graphics
517:   EraseGraphics m_pElemLine, m_pElemFillShp1, m_pElemFillShp2, m_pApp

  Exit Sub
ErrorHandler:
  HandleError True, "cmdOK_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdUncheckAll_Click()
  Dim i As Long
  
527:   With lstLayers
528:     For i = 0 To .ListCount - 1
529:       .ItemData(i) = vbUnchecked
530:       .Selected(i) = False
531:     Next i
532:   End With
End Sub




Private Sub Form_Activate()
  On Error GoTo ErrorHandler

541:   UpdateGraphics m_pApp, m_pElemLine, m_pElemFillShp1, m_pElemFillShp2, _
    m_pPnt1, m_pPnt2, m_pPnt3, m_pPnt4, m_sDataFrameName, True


  Exit Sub
ErrorHandler:
  HandleError True, "Form_Activate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Form_Deactivate()
  On Error GoTo ErrorHandler
                                                  'clean up all the temporary graphics
553:   EraseGraphics m_pElemLine, m_pElemFillShp1, m_pElemFillShp2, m_pApp
  
  Exit Sub
ErrorHandler:
  HandleError True, "Form_Deactivate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Form_Load()
  On Error GoTo ErrorHandler


  Dim pLayers As IEnumLayer, pLayer As ILayer
  Dim lLyrIdx As Long
  
567:   If m_pApp Is Nothing Then
568:     MsgBox "Application reference wasn't set.  Closing the bubble properties dialog is recommended."
    Exit Sub
570:   End If
571:   m_bLocked = False
                         
573:   chkEnableHeightWidth.Value = vbUnchecked
574:   optFixedRadius.Enabled = False
575:   optFixedScaleFactor.Enabled = False
576:   optFixedBubbleRadius.Enabled = False
                                                  ' acquire list of layers from map document,
                                                  ' and make them visible/non-visible based on
                                                  ' whether or not they're visible in the map
580:   lstLayers.Clear
  If m_pApp Is Nothing Then Exit Sub
582:   m_pMxDoc.ActiveView.Refresh
583:   Set pLayers = m_pMxDoc.FocusMap.Layers
584:   pLayers.Reset
585:   Set pLayer = pLayers.Next
586:   Do While Not pLayer Is Nothing
587:     lstLayers.AddItem pLayer.Name
588:     lLyrIdx = FindControlString(lstLayers, pLayer.Name)
589:     lstLayers.Selected(lLyrIdx) = pLayer.Visible
590:     Set pLayer = pLayers.Next
591:   Loop

593:   txtRadiusX.Visible = False
594:   txtBubbleRadiusX.Visible = False
595:   lblX.Top = lineScaleTop.Y1
596:   lblEquals.Visible = False
597:   lblHeight.Visible = False
598:   lblWidth.Visible = False
599:   txtScaleFactor.Top = lineScaleTop.Y1

  Exit Sub
ErrorHandler:
  HandleError True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Form_Terminate()
                                                  'clean up all the temporary graphics
608:   EraseGraphics m_pElemLine, m_pElemFillShp1, m_pElemFillShp2, m_pApp


611:   m_bCancelled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error GoTo ErrorHandler
                                                  'clean up all the temporary graphics
617:   EraseGraphics m_pElemLine, m_pElemFillShp1, m_pElemFillShp2, m_pApp
618:   m_bCancelled = True
  
  Exit Sub
ErrorHandler:
  HandleError True, "Form_Unload " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub



Private Sub optFixedBubbleRadius_Click()
628:   txtRadiusY.Enabled = True
629:   txtRadiusX.Enabled = True
630:   txtScaleFactor.Enabled = True
631:   txtBubbleRadiusY.Enabled = False
632:   txtBubbleRadiusX.Enabled = False
633:   txtRadiusY.BackColor = vbWindowBackground
634:   txtRadiusX.BackColor = vbWindowBackground
635:   txtScaleFactor.BackColor = vbWindowBackground
636:   txtBubbleRadiusY.BackColor = vbInactiveBorder
637:   txtBubbleRadiusX.BackColor = vbInactiveBorder
End Sub

Private Sub optFixedRadius_Click()
641:   txtRadiusY.Enabled = False
642:   txtRadiusX.Enabled = False
643:   txtScaleFactor.Enabled = True
644:   txtBubbleRadiusY.Enabled = True
645:   txtBubbleRadiusX.Enabled = True
646:   txtRadiusY.BackColor = vbInactiveBorder
647:   txtRadiusX.BackColor = vbInactiveBorder
648:   txtScaleFactor.BackColor = vbWindowBackground
649:   txtBubbleRadiusY.BackColor = vbWindowBackground
650:   txtBubbleRadiusX.BackColor = vbWindowBackground
End Sub

Private Sub optFixedScaleFactor_Click()
654:   txtRadiusY.Enabled = True
655:   txtRadiusX.Enabled = True
656:   txtScaleFactor.Enabled = False
657:   txtBubbleRadiusY.Enabled = True
658:   txtBubbleRadiusX.Enabled = True
659:   txtRadiusY.BackColor = vbWindowBackground
660:   txtRadiusX.BackColor = vbWindowBackground
661:   txtScaleFactor.BackColor = vbInactiveBorder
662:   txtBubbleRadiusY.BackColor = vbWindowBackground
663:   txtBubbleRadiusX.BackColor = vbWindowBackground
End Sub








Private Sub txtBubbleRadiusY_Change()
  On Error GoTo ErrorHandler
  
  Dim dTemp As Double
  
  If m_bInitializing = True Then Exit Sub
  If m_bLocked Then Exit Sub
680:   m_bLocked = True
                                                  'confirm that the input values are numeric
                                                  'alter the background to be yellow if they aren't
683:   If IsNumeric(txtBubbleRadiusY.Text) Then
684:     txtBubbleRadiusY.BackColor = vbWhite
685:     dTemp = m_dBubbleRadius
686:     m_dBubbleRadius = CDbl(txtBubbleRadiusY.Text)
687:     If m_dBubbleRadius = 0 Then
688:       MsgBox "The bubble radius cannot be zero.  The previous value will be restored."
689:       m_dBubbleRadius = dTemp
690:       txtBubbleRadiusY.Text = Format(m_dBubbleRadius, "#########0.00")
691:       m_bLocked = False
      Exit Sub
693:     End If
694:   Else
695:     txtBubbleRadiusY.BackColor = vbYellow
696:     m_bLocked = False
    Exit Sub
698:   End If
                                                  'if using some fixed value,
700:   If chkFixedValue Then
701:     If optFixedRadius Then
                                                  'scale factor:
703:       m_dScaleFactor = m_dBubbleRadius / m_dRadius
704:       txtScaleFactor.Text = Format(m_dScaleFactor, "#########0.00")
                                                  'cascade the change to the width values
706:       m_dBubbleRadiusX = m_dRadiusX * m_dScaleFactor
707:       txtBubbleRadiusX = Format(m_dBubbleRadiusX, "#########0.00")
708:     ElseIf optFixedScaleFactor Then
                                                  'or update radius:
710:       m_dRadius = m_dBubbleRadius / m_dScaleFactor
711:       txtRadiusY.Text = Format(m_dRadius, "#########0.00")
712:     End If
713:   Else
                                                  'else without a fixed value, adjust
                                                  'the radius
716:     m_dRadius = m_dBubbleRadius / m_dScaleFactor
717:     txtRadiusY.Text = Format(m_dRadius, "#########0.00")
    
719:     If chkEnableHeightWidth.Value = vbUnchecked Then
720:       m_dBubbleRadiusX = m_dBubbleRadius
721:       m_dRadiusX = m_dRadius
722:       txtBubbleRadiusX.Text = Format(m_dBubbleRadiusX, "#########0.00")
723:       txtRadiusX.Text = Format(m_dRadiusX, "#########0.00")
724:     End If
725:   End If
726:   m_bLocked = False

  Exit Sub
ErrorHandler:
730:   m_bLocked = False
  HandleError True, "txtBubbleRadiusY_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub txtBubbleRadiusY_LostFocus()
  On Error GoTo ErrorHandler

737:   If IsNumeric(txtRadiusY.Text) Then
738:     m_dRadius = CDbl(txtRadiusY.Text)
                                                  'otherwise, restore the original value,
740:   Else
                                                  'and reset the bubbleradius and radius
742:     txtRadiusY.Text = Format(m_dRadius, "#########0.00")
743:     m_dBubbleRadius = m_dRadius * m_dScaleFactor
744:     txtBubbleRadiusY.Text = Format(m_dBubbleRadius, "#########0.00")
745:   End If

747:   UpdateGraphics m_pApp, m_pElemLine, m_pElemFillShp1, m_pElemFillShp2, _
    m_pPnt1, m_pPnt2, m_pPnt3, m_pPnt4, m_sDataFrameName, True
  
  Exit Sub
ErrorHandler:
  HandleError True, "txtBubbleRadiusY_LostFocus " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub








Private Sub txtBubbleRadiusX_Change()
  On Error GoTo ErrorHandler


  Dim dTemp As Double
  
  If m_bInitializing = True Then Exit Sub
  If m_bLocked Then Exit Sub
770:   m_bLocked = True
  
772:   If IsNumeric(txtBubbleRadiusX.Text) Then
773:     txtBubbleRadiusX.BackColor = vbWhite
774:     dTemp = m_dBubbleRadiusX
775:     m_dBubbleRadiusX = CDbl(txtBubbleRadiusX.Text)
776:     If m_dBubbleRadiusX = 0 Then
777:       MsgBox "The bubble width cannot be zero.  The previous value will be restored."
778:       m_dBubbleRadiusX = dTemp
779:       txtBubbleRadiusX.Text = Format(m_dBubbleRadiusX, "#########0.00")
780:       m_bLocked = False
      Exit Sub
782:     End If
783:   Else
784:     txtBubbleRadiusX.BackColor = vbYellow
785:     m_bLocked = False
    Exit Sub
787:   End If

789:   If chkFixedValue.Value = vbChecked Then
790:     If optFixedRadius Then
                                                  'update scalefactor
792:       m_dScaleFactor = m_dBubbleRadiusX / m_dRadiusX
793:       txtScaleFactor.Text = Format(m_dScaleFactor, "#########0.00")
                                                  'cascade change to height values
795:       m_dBubbleRadius = m_dRadius * m_dScaleFactor
796:       txtBubbleRadiusY.Text = Format(m_dBubbleRadius, "#########0.00")
797:     ElseIf optFixedScaleFactor Then
                                                  'update RadiusX
799:       m_dRadiusX = m_dBubbleRadiusX / m_dScaleFactor
800:       txtRadiusX.Text = Format(m_dRadiusX, "#########0.00")
801:     End If
802:   Else
                                                  'update radius width
804:     m_dRadiusX = m_dBubbleRadiusX / m_dScaleFactor
805:     txtRadiusX.Text = Format(m_dRadiusX, "#########0.00")
806:   End If
807:   m_bLocked = False


  Exit Sub
ErrorHandler:
  HandleError True, "txtBubbleRadiusX_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
Private Sub txtBubbleRadiusX_LostFocus()
  On Error GoTo ErrorHandler

817:   If IsNumeric(txtBubbleRadiusX.Text) Then
818:     m_dBubbleRadiusX = CDbl(txtBubbleRadiusX.Text)
819:   Else
820:     txtBubbleRadiusX.Text = Format(m_dBubbleRadiusX, "#########0.00")
821:     m_dRadiusX = m_dBubbleRadiusX / m_dScaleFactor
822:     txtRadiusX.Text = Format(m_dRadiusX, "#########0.00")
823:   End If

825:   UpdateGraphics m_pApp, m_pElemLine, m_pElemFillShp1, m_pElemFillShp2, _
    m_pPnt1, m_pPnt2, m_pPnt3, m_pPnt4, m_sDataFrameName, True
  
  Exit Sub
ErrorHandler:
  HandleError True, "txtBubbleRadiusX_LostFocus " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub








Private Sub txtDestinationX_Change()
  On Error GoTo ErrorHandler

  
  Dim dTemp As Double
  
  If m_bInitializing = True Then Exit Sub
  If m_bLocked Then Exit Sub
848:   m_bLocked = True
                                                  'confirm that the input values are numeric
                                                  'alter the background to be yellow if they aren't
851:   If IsNumeric(txtDestinationX.Text) Then
852:     txtDestinationX.BackColor = vbWhite
853:     dTemp = m_dDestinationX
854:     m_dDestinationX = CDbl(txtDestinationX.Text)
855:     If m_dDestinationY = 0 Then
856:       MsgBox "The Destination X cannot be zero.  The previous value will be restored."
857:       m_dDestinationX = dTemp
858:       txtDestinationX.Text = Format(m_dDestinationX, "#########0.00")
859:       m_bLocked = False
      Exit Sub
861:     End If
862:   Else
863:     txtDestinationX.BackColor = vbYellow
864:     m_bLocked = False
    Exit Sub
866:   End If
867:   m_bLocked = False



  Exit Sub
ErrorHandler:
  HandleError True, "txtDestinationX_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub txtDestinationX_LostFocus()
877:   UpdateGraphics m_pApp, m_pElemLine, m_pElemFillShp1, m_pElemFillShp2, _
    m_pPnt1, m_pPnt2, m_pPnt3, m_pPnt4, m_sDataFrameName, True
End Sub

Private Sub txtDestinationY_Change()
  On Error GoTo ErrorHandler

  
  Dim dTemp As Double
  
  If m_bInitializing = True Then Exit Sub
  If m_bLocked Then Exit Sub
889:   m_bLocked = True
                                                  'confirm that the input values are numeric
                                                  'alter the background to be yellow if they aren't
892:   If IsNumeric(txtDestinationY.Text) Then
893:     txtDestinationY.BackColor = vbWhite
894:     dTemp = m_dDestinationY
895:     m_dDestinationY = CDbl(txtDestinationY.Text)
896:     If m_dDestinationX = 0 Then
897:       MsgBox "The Destination Y cannot be zero.  The previous value will be restored."
898:       m_dDestinationY = dTemp
899:       txtDestinationY.Text = Format(m_dDestinationY, "#########0.00")
900:       m_bLocked = False
      Exit Sub
902:     End If
903:   Else
904:     txtDestinationY.BackColor = vbYellow
905:     m_bLocked = False
    Exit Sub
907:   End If
908:   m_bLocked = False



  Exit Sub
ErrorHandler:
  HandleError True, "txtDestinationY_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub txtDestinationY_LostFocus()
918:   UpdateGraphics m_pApp, m_pElemLine, m_pElemFillShp1, m_pElemFillShp2, _
    m_pPnt1, m_pPnt2, m_pPnt3, m_pPnt4, m_sDataFrameName, True
End Sub

Private Sub txtOriginX_Change()
  On Error GoTo ErrorHandler

  Dim dTemp As Double
  
  If m_bInitializing = True Then Exit Sub
  If m_bLocked Then Exit Sub
929:   m_bLocked = True
                                                  'confirm that the input values are numeric
                                                  'alter the background to be yellow if they aren't
932:   If IsNumeric(txtOriginX.Text) Then
933:     txtOriginX.BackColor = vbWhite
934:     dTemp = m_dOriginX
935:     m_dOriginX = CDbl(txtOriginX.Text)
936:     If m_dOriginX = 0 Then
937:       MsgBox "The Origin X cannot be zero.  The previous value will be restored."
938:       m_dOriginX = dTemp
939:       txtOriginX.Text = Format(m_dOriginX, "#########0.00")
940:       m_bLocked = False
      Exit Sub
942:     End If
943:   Else
944:     txtOriginX.BackColor = vbYellow
945:     m_bLocked = False
    Exit Sub
947:   End If
948:   m_bLocked = False



  Exit Sub
ErrorHandler:
  HandleError True, "txtOriginX_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub txtOriginX_LostFocus()
958:   UpdateGraphics m_pApp, m_pElemLine, m_pElemFillShp1, m_pElemFillShp2, _
    m_pPnt1, m_pPnt2, m_pPnt3, m_pPnt4, m_sDataFrameName, True
End Sub

Private Sub txtOriginY_Change()
  On Error GoTo ErrorHandler

  
  Dim dTemp As Double
  
  If m_bInitializing = True Then Exit Sub
  If m_bLocked Then Exit Sub
970:   m_bLocked = True
                                                  'confirm that the input values are numeric
                                                  'alter the background to be yellow if they aren't
973:   If IsNumeric(txtOriginY.Text) Then
974:     txtOriginY.BackColor = vbWhite
975:     dTemp = m_dOriginY
976:     m_dOriginY = CDbl(txtOriginY.Text)
977:     If m_dOriginY = 0 Then
978:       MsgBox "The Origin Y cannot be zero.  The previous value will be restored."
979:       m_dOriginY = dTemp
980:       txtOriginY.Text = Format(m_dOriginY, "#########0.00")
981:       m_bLocked = False
      Exit Sub
983:     End If
984:   Else
985:     txtOriginY.BackColor = vbYellow
986:     m_bLocked = False
    Exit Sub
988:   End If
989:   m_bLocked = False


  Exit Sub
ErrorHandler:
  HandleError True, "txtOriginY_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub txtOriginY_LostFocus()
998:   UpdateGraphics m_pApp, m_pElemLine, m_pElemFillShp1, m_pElemFillShp2, _
    m_pPnt1, m_pPnt2, m_pPnt3, m_pPnt4, m_sDataFrameName, True
End Sub

Private Sub txtRadiusY_Change()
  On Error GoTo ErrorHandler
  
  Dim dTemp As Double
  
  If m_bInitializing = True Then Exit Sub
  If m_bLocked Then Exit Sub
1009:   m_bLocked = True
                                                  'confirm that the input values are numeric
                                                  'alter the background to be yellow if they aren't
1012:   If IsNumeric(txtRadiusY.Text) Then
1013:     txtRadiusY.BackColor = vbWhite
1014:     dTemp = m_dRadius
1015:     m_dRadius = CDbl(txtRadiusY.Text)
1016:     If m_dRadius = 0 Then
1017:       MsgBox "The radius cannot be zero.  The previous value will be restored."
1018:       m_dRadius = dTemp
1019:       txtRadiusY.Text = Format(m_dRadius, "#########0.00")
1020:       m_bLocked = False
      Exit Sub
1022:     End If
1023:   Else
1024:     txtRadiusY.BackColor = vbYellow
1025:     m_bLocked = False
    Exit Sub
1027:   End If
                                                  'if using some fixed value,
1029:   If chkFixedValue.Value = vbChecked Then
1030:     If optFixedScaleFactor Then
                                                  'update bubble radius:
1032:       m_dBubbleRadius = m_dRadius * m_dScaleFactor
1033:       txtBubbleRadiusY.Text = Format(m_dBubbleRadius, "#########0.00")
1034:     ElseIf optFixedBubbleRadius Then
                                                  'update scalefactor
1036:       m_dScaleFactor = m_dBubbleRadius / m_dRadius
1037:       txtScaleFactor.Text = Format(m_dScaleFactor, "#########0.00")
                                                  'cascade the change to the width values
1039:       m_dRadiusX = m_dBubbleRadiusX / m_dScaleFactor
1040:       txtRadiusX.Text = Format(m_dRadiusX, "#########0.00")
1041:     End If
1042:   Else
                                                  'else without a fixed value, adjust
                                                  'the bubble radius
1045:     m_dBubbleRadius = m_dRadius * m_dScaleFactor
1046:     txtBubbleRadiusY.Text = Format(m_dBubbleRadius, "#########0.00")
    
1048:     If chkEnableHeightWidth.Value = vbUnchecked Then
1049:       m_dBubbleRadiusX = m_dBubbleRadius
1050:       m_dRadiusX = m_dRadius
1051:       txtBubbleRadiusX.Text = Format(m_dBubbleRadiusX, "#########0.00")
1052:       txtRadiusX.Text = Format(m_dRadiusX, "#########0.00")
1053:     End If
1054:   End If
1055:   m_bLocked = False

  Exit Sub
ErrorHandler:
1059:   m_bLocked = False
  HandleError True, "txtRadiusY_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub txtRadiusY_LostFocus()
  On Error GoTo ErrorHandler

1066:   If IsNumeric(txtRadiusY.Text) Then
1067:     m_dRadius = CDbl(txtRadiusY.Text)
                                                  'otherwise, restore the original value,
1069:   Else
                                                  'and reset the bubbleradius and radius
1071:     txtRadiusY.Text = Format(m_dRadius, "#########0.00")
1072:     m_dBubbleRadius = m_dRadius * m_dScaleFactor
1073:     txtBubbleRadiusY.Text = Format(m_dBubbleRadius, "#########0.00")
1074:   End If

1076:   UpdateGraphics m_pApp, m_pElemLine, m_pElemFillShp1, m_pElemFillShp2, _
    m_pPnt1, m_pPnt2, m_pPnt3, m_pPnt4, m_sDataFrameName, True
  
  Exit Sub
ErrorHandler:
  HandleError True, "txtRadiusY_LostFocus " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub











Private Sub txtRadiusX_Change()
  On Error GoTo ErrorHandler

  Dim dTemp As Double
  
  If m_bInitializing Then Exit Sub
  If m_bLocked Then Exit Sub
1101:   m_bLocked = True
                                                  'QC the numeric text input
1103:   If IsNumeric(txtRadiusX.Text) Then
1104:     txtRadiusX.BackColor = vbWhite
1105:     dTemp = m_dRadiusX
1106:     m_dRadiusX = CDbl(txtRadiusX.Text)
1107:     If m_dRadiusX = 0 Then
1108:       MsgBox "The width must not be zero.  The previous value will be restored."
1109:       m_dRadiusX = dTemp
1110:       txtRadiusX.Text = Format(m_dRadiusX, "#########0.00")
1111:       m_bLocked = False
      Exit Sub
1113:     End If
1114:   Else
1115:     txtRadiusX.BackColor = vbYellow
1116:     m_bLocked = False
    Exit Sub
1118:   End If
                                                  'if using a fixed value
1120:   If chkFixedValue.Value = vbChecked Then
                                                  'update the bubble radius
1122:     If optFixedScaleFactor Then
1123:       m_dBubbleRadiusX = m_dRadiusX * m_dScaleFactor
1124:       txtBubbleRadiusX = Format(m_dBubbleRadiusX, "#########0.00")
                                                  'update the scale factor
1126:     ElseIf optFixedBubbleRadius Then
1127:       m_dScaleFactor = m_dBubbleRadiusX / m_dRadiusX
1128:       txtScaleFactor.Text = Format(m_dScaleFactor, "#########0.00")
                                                  'update the radius
1130:       m_dRadius = m_dBubbleRadius / m_dScaleFactor
1131:       txtRadiusY.Text = Format(m_dRadius, "#########0.00")
1132:     End If
1133:   Else
1134:     m_dBubbleRadiusX = m_dRadiusX * m_dScaleFactor
1135:     txtBubbleRadiusX.Text = Format(m_dBubbleRadiusX, "#########0.00")
1136:   End If
  
1138:   m_bLocked = False

  Exit Sub
ErrorHandler:
  HandleError True, "txtRadiusX_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub txtRadiusX_LostFocus()
  On Error GoTo ErrorHandler

1148:   If IsNumeric(txtRadiusX.Text) Then
1149:     m_dRadiusX = CDbl(txtRadiusX.Text)
1150:   Else
1151:     txtRadiusX.Text = Format(m_dRadiusX, "#########0.00")
1152:     m_dBubbleRadiusX = m_dRadiusX * m_dScaleFactor
1153:     txtBubbleRadiusX.Text = Format(m_dBubbleRadiusX, "#########0.00")
1154:   End If

1156:   UpdateGraphics m_pApp, m_pElemLine, m_pElemFillShp1, m_pElemFillShp2, _
    m_pPnt1, m_pPnt2, m_pPnt3, m_pPnt4, m_sDataFrameName, True
  
  Exit Sub
ErrorHandler:
  HandleError True, "txtRadiusX_LostFocus " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub






Private Sub txtScaleFactor_Change()
  On Error GoTo ErrorHandler
  
  Dim dTemp As Double
  
  If m_bInitializing = True Then Exit Sub
  If m_bLocked Then Exit Sub
1176:   m_bLocked = True
                                                  'confirm that the input values are numeric
                                                  'alter the background to be yellow if they aren't
1179:   If IsNumeric(txtScaleFactor.Text) Then
1180:     txtScaleFactor.BackColor = vbWhite
1181:     dTemp = m_dScaleFactor
1182:     m_dScaleFactor = CDbl(txtScaleFactor.Text)
1183:     If m_dScaleFactor = 0 Then
1184:       MsgBox "The scale factor cannot be zero.  The previous value will be restored."
1185:       m_dScaleFactor = dTemp
1186:       txtBubbleRadiusY.Text = Format(m_dScaleFactor, "#########0.00")
1187:       m_bLocked = False
      Exit Sub
1189:     End If
1190:   Else
1191:     txtScaleFactor.BackColor = vbYellow
1192:     m_bLocked = False
    Exit Sub
1194:   End If
                                                  'if using a fixed value,
1196:   If chkFixedValue.Value = vbChecked Then
1197:     If optFixedRadius Then
                                                  'update bubble radius:
1199:       m_dBubbleRadius = m_dRadius * m_dScaleFactor
1200:       txtBubbleRadiusY.Text = Format(m_dBubbleRadius, "#########0.00")
                                                  'cascade change to width values
1202:       m_dBubbleRadiusX = m_dRadiusX * m_dScaleFactor
1203:       txtBubbleRadiusX.Text = Format(m_dBubbleRadiusX, "#########0.00")
1204:     ElseIf Me.optFixedBubbleRadius Then
                                                  'update radius
1206:       m_dRadius = m_dBubbleRadius / m_dScaleFactor
1207:       txtRadiusY.Text = Format(m_dRadius, "#########0.00")
                                                  'cascade change to width values
1209:       m_dRadiusX = m_dBubbleRadiusX / m_dScaleFactor
1210:       txtRadiusX.Text = Format(m_dRadiusX, "#########0.00")
1211:     End If
1212:   Else
                                                  'else without a fixed value, adjust
                                                  'the bubble radius
1215:     m_dBubbleRadius = m_dRadius * m_dScaleFactor
1216:     txtBubbleRadiusY.Text = Format(m_dBubbleRadius, "#########0.00")
                                                  'cascade change to bubbleRadiusX
1218:     m_dBubbleRadiusX = m_dRadiusX * m_dScaleFactor
1219:     txtBubbleRadiusX.Text = Format(m_dBubbleRadiusX, "#########0.00")
1220:   End If
1221:   m_bLocked = False
  
  Exit Sub
ErrorHandler:
1225:   m_bLocked = False
  HandleError True, "txtScaleFactor_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub txtScaleFactor_LostFocus()
  On Error GoTo ErrorHandler

1232:   If IsNumeric(txtScaleFactor.Text) Then
1233:     m_dScaleFactor = CDbl(txtScaleFactor.Text)
                                                  'otherwise, restore the original value,
1235:   Else
                                                  'and reset the bubbleradius and radius
1237:     txtScaleFactor.Text = Format(m_dScaleFactor, "#########0.00")
1238:     m_dBubbleRadius = m_dRadius * m_dScaleFactor
1239:     txtBubbleRadiusY.Text = Format(m_dBubbleRadius, "#########0.00")
1240:   End If

1242:   UpdateGraphics m_pApp, m_pElemLine, m_pElemFillShp1, m_pElemFillShp2, _
    m_pPnt1, m_pPnt2, m_pPnt3, m_pPnt4, m_sDataFrameName, True
  
  Exit Sub
ErrorHandler:
  HandleError True, "txtScaleFactor_LostFocus " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub






Private Function AddCircleElement(pCircArc As ICircularArc, pAV As IActiveView, pGraCont As IGraphicsContainer) As IFillShapeElement
  On Error GoTo ErrorHandler

  ' Takes an ICircularArc and IActiveView and creates a CircleElement in the ActiveView's BasicGraphicsLayer
  Dim pElemFillShp As IFillShapeElement
  Dim pElem As IElement
  Dim pSFSym As ISimpleFillSymbol
  Dim pRGB As IRgbColor
  Dim pSegColl As ISegmentCollection

  ' Create a new Polygon object and access the ISegmentCollection interface to add a segment
1266:   Set pSegColl = New Polygon
1267:   pSegColl.AddSegment pCircArc

  ' Create a new circleelement and use the IElement interface to set the its Geometry
1270:   Set pElem = New CircleElement
1271:   pElem.Geometry = pSegColl

  ' QI for the IFillShapeElement interface so that the Symbol property can be set
1274:   Set pElemFillShp = pElem

  ' Create a new RGBColor
1277:   Set pRGB = New RgbColor
1278:   With pRGB
1279:     .Red = 255
1280:     .Green = 64
1281:     .Blue = 64
1282:   End With

  ' Create a new SimpleFillSymbol and set its Color and Style
1285:   Set pSFSym = New SimpleFillSymbol
1286:   pSFSym.Color = pRGB
1287:   pSFSym.Outline.Color = pRGB
1288:   pSFSym.Style = esriSFSHollow
1289:   pElemFillShp.Symbol = pSFSym

  ' QI for the IGraphicsContainer interface from the IActiveView, allows access to the BasicGraphicsLayer
1292:   Set pGraCont = pAV
1293:   Set AddCircleElement = pElemFillShp

  Exit Function
ErrorHandler:
  HandleError False, "AddCircleElement " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function



Private Function AddLineElement(pLine As IPolyline, pAV As IActiveView, pGraCont As IGraphicsContainer) As ILineElement
  On Error GoTo ErrorHandler

  Dim pElem As IElement, pLineSym As ISimpleLineSymbol, pRGB As IRgbColor
  Dim pElemLine As ILineElement
  Dim pGeom As IGeometry
  
1309:   Set pElem = New LineElement
1310:   Set pGeom = pLine
1311:   pElem.Geometry = pGeom
1312:   Set pElemLine = pElem
  
1314:   Set pRGB = New RgbColor
1315:   With pRGB
1316:     .Red = 0
1317:     .Green = 0
1318:     .Blue = 0
1319:   End With

1321:   Set pLineSym = New SimpleLineSymbol
1322:   pLineSym.Color = pRGB
1323:   pLineSym.Style = esriSLSSolid
1324:   pElemLine.Symbol = pLineSym
  
  ' QI for the IGraphicsContainer interface from the IActiveView, allows access to the BasicGraphicsLayer
1327:   Set pGraCont = pAV
1328:   Set AddLineElement = pElemLine

  Exit Function
ErrorHandler:
  HandleError False, "AddLineElement " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function







'UpdateGraphics
'
'This routine attempts to update the two circle graphics along with the
'line connecting the centers of the two circles.  The idea is to graphically
'reflect to the user the status of the input values they have provided.
'
'This routine presumes that all of the input provided among the various
'controls of this form is valid.
'--------------------------
Public Sub UpdateGraphics(pApp As IApplication, _
                          ByRef pLineElem As ILineElement, _
                          ByRef pCirc1Elem As ICircleElement, _
                          ByRef pCirc2Elem As ICircleElement, _
                          pPntCircleOrigCenter As IPoint, _
                          pPntCircleOrigEdge As IPoint, _
                          pPntCircleDestCenter As IPoint, _
                          pPntCircleDestEdge As IPoint, _
                          sDataFrameName As String, _
                          Optional bUseFormVariables As Boolean)
  On Error GoTo ErrorHandler

  Dim pCircArc As IGeometry, pOriginCircularArc As ICircularArc
  Dim pDestinationCircularArc As ICircularArc, pAV As IActiveView
  Dim pLine As ILine, pPolyline As IPolyline, pGeomColl As IGeometryCollection
  Dim pSegmentColl As ISegmentCollection, pGrCont As IGraphicsContainer
  Dim pMxDoc As IMxDocument
  
1368:   If IsMissing(bUseFormVariables) Then
1369:     bUseFormVariables = True
1370:   End If
  
1372:   If bUseFormVariables Then
                                      
                                      'update the four points with variables
                                      'controlled and updated in this form
1376:     pPntCircleOrigCenter.x = m_dOriginX
1377:     pPntCircleOrigCenter.y = m_dOriginY
1378:     pPntCircleOrigEdge.x = m_dOriginX + m_dRadius
1379:     pPntCircleOrigEdge.y = m_dOriginY
1380:     pPntCircleDestCenter.x = m_dDestinationX
1381:     pPntCircleDestCenter.y = m_dDestinationY
1382:     pPntCircleDestEdge.x = m_dDestinationX + m_dBubbleRadius
1383:     pPntCircleDestEdge.y = m_dDestinationY
1384:   End If
  
  If pApp Is Nothing Then Exit Sub
1387:   Set pMxDoc = pApp.Document
1388:   If TypeOf pMxDoc.ActiveView Is IPageLayout Then
1389:     Set pGrCont = pMxDoc.PageLayout
1390:     Set pPntCircleOrigCenter = LayoutUnitsFromMapUnits(sDataFrameName, _
                                                      pPntCircleOrigCenter.x, _
                                                      pPntCircleOrigCenter.y, _
                                                      pApp)
1394:     Set pPntCircleOrigEdge = LayoutUnitsFromMapUnits(sDataFrameName, _
                                                      pPntCircleOrigEdge.x, _
                                                      pPntCircleOrigEdge.y, _
                                                      pApp)
1398:     Set pPntCircleDestCenter = LayoutUnitsFromMapUnits(sDataFrameName, _
                                                      pPntCircleDestCenter.x, _
                                                      pPntCircleDestCenter.y, _
                                                      pApp)
1402:     Set pPntCircleDestEdge = LayoutUnitsFromMapUnits(sDataFrameName, _
                                                      pPntCircleDestEdge.x, _
                                                      pPntCircleDestEdge.y, _
                                                      pApp)
1406:   Else
1407:     Set pGrCont = pMxDoc.FocusMap
1408:   End If
  
1410:   If Not pLineElem Is Nothing Then pGrCont.DeleteElement pLineElem
1411:   If Not pCirc1Elem Is Nothing Then pGrCont.DeleteElement pCirc1Elem
1412:   If Not pCirc2Elem Is Nothing Then pGrCont.DeleteElement pCirc2Elem
  
1414:   Set pOriginCircularArc = New CircularArc
1415:   pOriginCircularArc.PutCoords pPntCircleOrigCenter, pPntCircleOrigEdge, pPntCircleOrigEdge, esriArcClockwise
1416:   Set pDestinationCircularArc = New CircularArc
1417:   pDestinationCircularArc.PutCoords pPntCircleDestCenter, pPntCircleDestEdge, pPntCircleDestEdge, esriArcClockwise
  
  
1420:   Set pLine = New esriGeometry.Line
1421:   pLine.PutCoords pPntCircleOrigCenter, pPntCircleDestCenter
1422:   Set pSegmentColl = New Path
1423:   pSegmentColl.AddSegment pLine
1424:   Set pGeomColl = New Polyline
1425:   pGeomColl.AddGeometry pSegmentColl
1426:   Set pPolyline = pGeomColl
  
  
1429:   Set pAV = pMxDoc.ActiveView
1430:   Set pCirc1Elem = AddCircleElement(pOriginCircularArc, pAV, pGrCont)
1431:   Set pCirc2Elem = AddCircleElement(pDestinationCircularArc, pAV, pGrCont)
1432:   Set pLineElem = AddLineElement(pPolyline, pAV, pGrCont)
  
1434:   pGrCont.AddElement pCirc1Elem, 0
1435:   pGrCont.AddElement pCirc2Elem, 0
1436:   pGrCont.AddElement pLineElem, 0
1437:   pAV.PartialRefresh esriViewGraphics, Nothing, Nothing


  Exit Sub
ErrorHandler:
  HandleError False, "UpdateGraphics " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub




Public Sub EraseGraphics(pLineElem As ILineElement, _
                         pCirc1Elem As ICircleElement, _
                         pCirc2Elem As ICircleElement, _
                         pApp As IApplication)
  On Error GoTo ErrorHandler
  Dim pMxDoc As IMxDocument, pGrCont As IGraphicsContainer

  If pApp Is Nothing Then Exit Sub
1456:   Set pMxDoc = pApp.Document
1457:   Set pGrCont = pMxDoc.ActiveView
  
1459:   If Not pLineElem Is Nothing Then pGrCont.DeleteElement pLineElem
1460:   If Not pCirc1Elem Is Nothing Then pGrCont.DeleteElement pCirc1Elem
1461:   If Not pCirc2Elem Is Nothing Then pGrCont.DeleteElement pCirc2Elem
1462:   Set pLineElem = Nothing
1463:   Set pCirc1Elem = Nothing
1464:   Set pCirc2Elem = Nothing
                                        'refresh the view
1466:   pMxDoc.ActiveView.Refresh

  Exit Sub
ErrorHandler:
  HandleError True, "EraseGraphics " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

