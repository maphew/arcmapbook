VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_pApp As esriFramework.IApplication
Private m_pMxDoc As IMxDocument

Implements ICommand
' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "D:\DSMapbook\DataFrame2\dfmain.cls"



Private Sub Class_Terminate()
  On Error GoTo ErrorHandler

    
    Set m_pApp = Nothing

  Exit Sub
ErrorHandler:
  HandleError True, "Class_Terminate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
Private Sub DataFrameProperties(pMap As IMap)
'add background
        
'        Dim pFrameElement As IFrameElement
'        Set pFrameElement = pElement
'        Dim pColor As IColor
'        Set pColor = New RgbColor
'        pColor.RGB = RGB(255, 0, 0)
'
'        ' create a framedecoration to modify the frame element background
'        Dim pFrameDecoration As IFrameDecoration
'        Set pFrameDecoration = New SymbolBackground
'        With pFrameDecoration
'            .Color = pColor
'            .CornerRounding = 100
'            .HorizontalSpacing = 0
'            .VerticalSpacing = 0
'        End With
'        pFrameElement.Background = pFrameDecoration
'
'        'add border
'
'        Dim pBorder As IBorder
'        Set pBorder = pMapFrame.Border
'
'        Dim pSymBorder As ISymbolBorder
'        Set pSymBorder = pMapFrame.Border
'
'        Dim pLineSymbol As ILineSymbol
'        Set pLineSymbol = pSymBorder.LineSymbol
'
'        pLineSymbol.Width = 2
'        pLineSymbol.Color = pColor
'        pSymBorder.LineSymbol = pLineSymbol
'        pSymBorder.CornerRounding = 100
'        pFrameElement.Border = pSymBorder
'
'
'        ''''pGraphicsContainer.UpdateElement pElement
'
'        ' add a shadow
'        Dim pFillSymbol As IFillSymbol
'        Set pFillSymbol = New SimpleFillSymbol
'        'pColor.RGB = RGB(255, 0, 0)
'        pFillSymbol.Color = pColor
'        Dim pSymbolShadow As ISymbolShadow
'        Set pSymbolShadow = New SymbolShadow
'        With pSymbolShadow
'            .FillSymbol = pFillSymbol
'            .HorizontalSpacing = 0
'            .VerticalSpacing = 0
'            .CornerRounding = 100
'        End With
'
'        ' use iFrameProperties to set the shadow of the frame element
'        Dim pFrameProperties As IFrameProperties
'        Set pFrameProperties = pFrameElement
'        pFrameProperties.Background = pFrameElement.Background
'        pFrameProperties.Border = pFrameElement.Border
'        pFrameProperties.Shadow = pSymbolShadow
'
'        pMapFrame.Background = pFrameProperties.Background  ' pFrameDecoration
'        pMapFrame.Border = pFrameProperties.Border
End Sub

Private Sub FrameDeco(pGraphicsContainer As IGraphicsContainer)
    'Dim m_pMxDoc As IMxDocument
    'Set m_pMxDoc = ThisDocument
    'Dim pGraphicsContainer As IGraphicsContainer
    'Set pGraphicsContainer = pMapFrame
    Dim pGraphicsContainerSelect As IGraphicsContainerSelect
    Set pGraphicsContainerSelect = pGraphicsContainer
    Dim pEnumElement As IEnumElement
    pGraphicsContainerSelect.SelectAllElements
    Set pEnumElement = pGraphicsContainerSelect.SelectedElements
    pEnumElement.Reset
    Dim pElement As IElement
    Set pElement = pEnumElement.Next
    Dim pColor As IColor
    Set pColor = New RgbColor
    pColor.RGB = RGB(0, 255, 0)
    ' Loop on all the selected elements of the page layout
    Do While Not pElement Is Nothing
        If TypeOf pElement Is IFrameElement Then
            Dim pFrameElement As IFrameElement
            Set pFrameElement = pElement
            
            ' create a framedecoration to modify the frame element background
            Dim pFrameDecoration As IFrameDecoration
            Set pFrameDecoration = New SymbolBackground
            With pFrameDecoration
                .Color = pColor
                .CornerRounding = 100
                .HorizontalSpacing = 20
                .VerticalSpacing = 40
            End With
            pFrameElement.Background = pFrameDecoration
            pGraphicsContainer.UpdateElement pElement
            ' add a shadow
            Dim pFillSymbol As IFillSymbol
            Set pFillSymbol = New SimpleFillSymbol
            pFillSymbol.Color = pColor
            Dim pSymbolShadow As ISymbolShadow
            Set pSymbolShadow = New SymbolShadow
            With pSymbolShadow
                .FillSymbol = pFillSymbol
                .HorizontalSpacing = 20
                .VerticalSpacing = 20
                .CornerRounding = 100
            End With
            ' use iFrameProperties to set the shadow of the frame element
            Dim pFrameProperties As IFrameProperties
            Set pFrameProperties = pFrameElement
            pFrameProperties.Shadow = pSymbolShadow
          End If
          Set pElement = pEnumElement.Next
      Loop
      ' refresh the view
      'Dim pActiveView As IActiveView
     ' Set pActiveView = m_pMxDoc.PageLayout
     ' pActiveView.Refresh
End Sub


Private Sub AddDataframe()
  On Error GoTo ErrorHandler

    'Dim m_pMxDoc As IMxDocument
    If Not m_pMxDoc Is Nothing Then  'Application
        Set m_pMxDoc.ActiveView = m_pMxDoc.PageLayout
        Dim pMap As IMap
        Set pMap = New Map
        pMap.Name = "New Frame"
        pMap.Description = "New Frame"
                
        'Create a new MapFrame and associate map with it
        Dim pMapFrame As IMapFrame
        Set pMapFrame = New MapFrame
        Set pMapFrame.Map = pMap
        
        Dim pGraphicsContainer As IGraphicsContainer
        Set pGraphicsContainer = m_pMxDoc.PageLayout
        
        'Set the position of the new map frame
        Dim pElement As IElement
        
        Dim pEnv As IEnvelope
        Set pElement = pMapFrame
               
        Set pEnv = New Envelope
        pEnv.PutCoords 3, 3, 6.5, 6.5
        'pEnv.GeometryType = esriGeometryCircularArc
        pElement.Geometry = pEnv
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Add code to set data frame properties
         Dim pColor As IColor
        Set pColor = New RgbColor
        pColor.RGB = RGB(0, 255, 0)
        Dim pFrameElement As IFrameElement
        Set pFrameElement = pElement
        
        ' create a border
        Dim pSymbolBorder As ISymbolBorder
        Set pSymbolBorder = New SymbolBorder
        Dim pLineSymbol As ILineSymbol
        Set pLineSymbol = New SimpleLineSymbol
        pLineSymbol.Width = 2
        pLineSymbol.Color = pColor
        pSymbolBorder.LineSymbol = pLineSymbol
        pSymbolBorder.LineSymbol.Color = pColor
        pSymbolBorder.LineSymbol.Width = 2
        pSymbolBorder.CornerRounding = 100
        pFrameElement.Border = pSymbolBorder
        
        ' create a framedecoration to modify the frame element background
        Dim pFrameDecoration As IFrameDecoration
        Set pFrameDecoration = New SymbolBackground
        With pFrameDecoration
            .Color = pColor
            .CornerRounding = 100
            .HorizontalSpacing = 20
            .VerticalSpacing = 40
        End With
        
        pFrameElement.Background = pFrameDecoration
        
        
        ' add a shadow
        Dim pFillSymbol As IFillSymbol
        Set pFillSymbol = New SimpleFillSymbol
        pFillSymbol.Color = pColor
        Dim pSymbolShadow As ISymbolShadow
        Set pSymbolShadow = New SymbolShadow
        With pSymbolShadow
            .FillSymbol = pFillSymbol
            .HorizontalSpacing = 20
            .VerticalSpacing = 20
            .CornerRounding = 100
        End With
        
        ' use iFrameProperties to set the shadow of the frame element
        Dim pFrameProperties As IFrameProperties
        Set pFrameProperties = pFrameElement
        pFrameProperties.Shadow = pSymbolShadow

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Add mapframe to the layout
        'pGraphicsContainer.AddElement pMapFrame, 0
        pGraphicsContainer.AddElement pElement, 0
        
        
        
        'Make the newly added map the focus map
        Dim pActiveView As IActiveView
        Set pActiveView = m_pMxDoc.ActiveView
        Set pActiveView.FocusMap = pMap
        
        'Refresh ActiveView and TOC
        pActiveView.Refresh
        m_pMxDoc.CurrentContentsView.Refresh 0
    End If


  Exit Sub
ErrorHandler:
  HandleError False, "AddDataframe " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
  On Error GoTo ErrorHandler



  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Bitmap " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Property Get ICommand_Caption() As String
  On Error GoTo ErrorHandler

    ICommand_Caption = "Dataframe Command"

  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Caption " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Property Get ICommand_Category() As String
  On Error GoTo ErrorHandler

  ICommand_Category = "COS MapBook"

  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Category " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Property Get ICommand_Checked() As Boolean
  On Error GoTo ErrorHandler

    ICommand_Checked = False

  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Checked " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Property Get ICommand_Enabled() As Boolean
  On Error GoTo ErrorHandler

    ICommand_Enabled = True

  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Enabled " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Property Get ICommand_HelpContextID() As Long
  On Error GoTo ErrorHandler

    ICommand_HelpContextID = 0

  Exit Property
ErrorHandler:
  HandleError True, "ICommand_HelpContextID " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Property Get ICommand_HelpFile() As String
  On Error GoTo ErrorHandler

    ICommand_HelpFile = ""

  Exit Property
ErrorHandler:
  HandleError True, "ICommand_HelpFile " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Property Get ICommand_Message() As String
  On Error GoTo ErrorHandler

    ICommand_Message = "Dataframe Command Messsage"

  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Message " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Property Get ICommand_Name() As String
  On Error GoTo ErrorHandler

    ICommand_Name = "Dataframe Command"

  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Name " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property

Private Sub ICommand_OnClick()
  On Error GoTo ErrorHandler

    AddDataframe

  Exit Sub
ErrorHandler:
  HandleError True, "ICommand_OnClick " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
  On Error GoTo ErrorHandler

    Set m_pApp = hook
    If Not m_pApp Is Nothing Then  'Application
        Set m_pMxDoc = m_pApp.Document
    End If
  Exit Sub
ErrorHandler:
  HandleError True, "ICommand_OnCreate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Property Get ICommand_Tooltip() As String
  On Error GoTo ErrorHandler

    ICommand_Tooltip = "DataframeCommand"

  Exit Property
ErrorHandler:
  HandleError True, "ICommand_Tooltip " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Property
