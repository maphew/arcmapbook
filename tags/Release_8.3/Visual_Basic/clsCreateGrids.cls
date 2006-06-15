VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCreateGrids"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'----------------------------------------------
' Properties
' - DestinationFeatureClass (IFeatureClass, r/w)
' - MapScale (Double, r/w)
' - FrameWidthInPageUnits (Double, r/w)
' - FrameHeightInPageUnits (Double, r/w)
' - IdentifierOrder (mgGridIdentifierOrderType, r/w)
' - RowIDType (mgGridIdentifierValueType, r/w)
' - ColIDType (mgGridIdentifierValueType, r/w)
' - StartingCoordinateLL_X (Double, r/w)
' - StartingCoordinateLL_Y (Double, r/w)
' - EndingCoordinateUR_X (Double, r/w)
' - EndingCoordinateUR_Y (Double, r/w)
' - UseUnderscore (Boolean, r/w)
'----------------------------------------------
' Methods
' - GenerateGrids(pApp as IApplication)
'    : Generates the grids using the values added.
' - RunStandardGUI(pApp as IApplication)
'    : To open the form without having the button
'      added to ArcMap's GUI.
'----------------------------------------------
' Enumerated Types
Public Enum mgGridIdentifierOrderType
    Row_Column = 0
    Column_Row = 1
End Enum
Public Enum mgGridIdentifierValueType
    Alphabetical = 0
    Numerical = 1
End Enum
Public Enum mgGridIdentifierStartPositionType
    TopLeft = 0
    LowerLeft = 1
End Enum
'----------------------------------------------
' Local Global Declarations
Private m_DestFL As IFeatureLayer
Private m_DestFC As IFeatureClass
Private m_dMapScale As Double
Private m_dFrameWidthInPageUnits As Double
Private m_dFrameHeightInPageUnits As Double
Private m_IDOrderType As mgGridIdentifierOrderType
Private m_RowIDType As mgGridIdentifierValueType
Private m_ColIDType As mgGridIdentifierValueType
Private m_StartIDType As mgGridIdentifierStartPositionType
Private m_StartX As Double
Private m_StartY As Double
Private m_EndX As Double
Private m_EndY As Double
Private m_UseUnderscore As Boolean
Private m_colLayerNamesForData As Collection
Private m_FldID As String
Private m_FldRowNum As String
Private m_FldColNum As String
Private m_FldScale As String
Private m_RemoveGrids As Boolean
Private m_NoEmptyGrids As Boolean
'----------------------------------------------
' API call to keep form top most
Private Const GWL_HWNDPARENT = -8
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
  (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Property Set DestinationFeatureLayer(pFL As IFeatureLayer)
    If pFL.FeatureClass.ShapeType = esriGeometryPolygon Then
        Set m_DestFL = pFL
        Set m_DestFC = pFL.FeatureClass
    Else
        Err.Raise vbObjectError, "MapGridManager_Set_DestinationFeatureLayer", _
            "Not a polygon feature layer"
    End If
End Property

Public Property Get DestinationFeatureLayer() As IFeatureLayer
    Set DestinationFeatureLayer = m_DestFL
End Property

Public Property Let FrameWidthInPageUnits(dWidth As Double)
    m_dFrameWidthInPageUnits = dWidth
End Property

Public Property Get FrameWidthInPageUnits() As Double
    FrameWidthInPageUnits = m_dFrameWidthInPageUnits
End Property

Public Property Let FrameHeightInPageUnits(dHeight As Double)
    m_dFrameHeightInPageUnits = dHeight
End Property

Public Property Get FrameHeightInPageUnits() As Double
    FrameHeightInPageUnits = m_dFrameHeightInPageUnits
End Property

Public Property Let MapScale(dScale As Double)
    m_dMapScale = dScale
End Property

Public Property Get MapScale() As Double
    MapScale = m_dMapScale
End Property

Public Property Let IdentifierOrder(mgType As mgGridIdentifierOrderType)
    m_IDOrderType = mgType
End Property

Public Property Get IdentifierOrder() As mgGridIdentifierOrderType
    IdentifierOrder = m_IDOrderType
End Property

Public Property Let RowIDType(mgIDStyle As mgGridIdentifierValueType)
    m_RowIDType = mgIDStyle
End Property

Public Property Get RowIDType() As mgGridIdentifierValueType
    RowIDType = m_RowIDType
End Property

Public Property Let ColIDType(mgIDStyle As mgGridIdentifierValueType)
    m_ColIDType = mgIDStyle
End Property

Public Property Get ColIDType() As mgGridIdentifierValueType
    ColIDType = m_ColIDType
End Property

Public Property Let IDStartPositionType(mgStartPos As mgGridIdentifierStartPositionType)
    m_StartIDType = mgStartPos
End Property

Public Property Get IDStartPositionType() As mgGridIdentifierStartPositionType
    IDStartPositionType = m_StartIDType
End Property

Public Property Let StartingCoordinateLL_X(X As Double)
    m_StartX = X
End Property

Public Property Get StartingCoordinateLL_X() As Double
    StartingCoordinateLL_X = m_StartX
End Property

Public Property Let StartingCoordinateLL_Y(Y As Double)
    m_StartY = Y
End Property

Public Property Get StartingCoordinateLL_Y() As Double
    StartingCoordinateLL_Y = m_StartY
End Property

Public Property Let EndingCoordinateUR_X(X As Double)
    m_EndX = X
End Property

Public Property Get EndingCoordinateUR_X() As Double
    EndingCoordinateUR_X = m_EndX
End Property

Public Property Let EndingCoordinateUR_Y(Y As Double)
    m_EndY = Y
End Property

Public Property Get EndingCoordinateUR_Y() As Double
    EndingCoordinateUR_Y = m_EndY
End Property

Public Property Let FieldNameGridID(FieldName As String)
    m_FldID = FieldName
End Property

Public Property Let FieldNameRowNum(FieldName As String)
    m_FldRowNum = FieldName
End Property

Public Property Let FieldNameColNum(FieldName As String)
    m_FldColNum = FieldName
End Property

Public Property Let FieldNameScale(FieldName As String)
    m_FldScale = FieldName
End Property

Public Property Let UseUnderscore(UseUnderscoreInID As Boolean)
    m_UseUnderscore = UseUnderscoreInID
End Property

Public Property Get UseUnderscore() As Boolean
    UseUnderscore = m_UseUnderscore
End Property

Public Property Let RemoveCurrentGrids(RemoveGrids As Boolean)
    m_RemoveGrids = RemoveGrids
End Property

Public Property Get RemoveCurrentGrids() As Boolean
    RemoveCurrentGrids = m_RemoveGrids
End Property

Public Property Let NoEmptyGrids(NoEmptyGridPolygons As Boolean)
    m_NoEmptyGrids = NoEmptyGridPolygons
End Property

Public Property Get NoEmptyGrids() As Boolean
    NoEmptyGrids = m_NoEmptyGrids
End Property

Public Sub AddRequiredDataLayer(sLayerName As String)
    m_colLayerNamesForData.Add sLayerName
End Sub

Public Property Get RequiredDataLayer(LayerNumber As Long) As String
    If LayerNumber <= m_colLayerNamesForData.count Then
        RequiredDataLayer = m_colLayerNamesForData.Item(LayerNumber - 1)
    Else
        If m_colLayerNamesForData.count = 0 Then
            RequiredDataLayer = "There are no Required Data Layers."
        Else
            RequiredDataLayer = "Invalid layer number.  Valid range: 0 - " & (m_colLayerNamesForData.count - 1)
        End If
    End If
End Property

Public Property Get RequiredDataLayerCount() As Long
    RequiredDataLayerCount = m_colLayerNamesForData.count
End Property

Public Sub ClearRequiredDataLayers()
    Set m_colLayerNamesForData = New Collection
End Sub

Private Sub Class_Initialize()
    ' Set the default grid identifier (row-col, alpha/numeric => eg: "B02")
    m_IDOrderType = Row_Column
    m_RowIDType = Alphabetical
    m_ColIDType = Numerical
    Set m_colLayerNamesForData = New Collection
End Sub

Private Function CalculateID(ByVal lRow As Long, ByVal lCol As Long, _
                             ByVal iRowIDLen As Integer, ByVal iColIDLen As Integer) As String
    Dim sRowID As String
    Dim sColID As String
    Dim sNumericFormat As String
    Dim lLoop As Long
    Dim lTmp As Long, lCalc As Long, lTmp2 As Long

    ' Row ---------------------------------------------
    sNumericFormat = ""
    For lLoop = 1 To iRowIDLen
        sNumericFormat = sNumericFormat & "0"
    Next
    If m_RowIDType = Alphabetical Then
        For lLoop = 1 To iRowIDLen
            lTmp = 26 ^ lLoop
            lTmp2 = (26 ^ (lLoop - 1))
            If lRow >= lTmp2 Then
                lCalc = ((((lRow - lTmp2) / lTmp) * 26) + 1) Mod 26
                sRowID = Chr(Asc("A") + lCalc) & sRowID
                lRow = lRow - (lCalc * lTmp2)
            Else
                sRowID = "A" & sRowID
            End If
        Next
    Else
        sRowID = Format(lRow + 1, sNumericFormat)
    End If
    ' Col ---------------------------------------------
    sNumericFormat = ""
    For lLoop = 1 To iColIDLen
        sNumericFormat = sNumericFormat & "0"
    Next
    If m_ColIDType = Alphabetical Then
        For lLoop = 1 To iColIDLen
            lTmp = 26 ^ lLoop
            lTmp2 = (26 ^ (lLoop - 1))
            If lCol >= lTmp2 Then
                lCalc = ((((lCol - lTmp2) / lTmp) * 26) + 1) Mod 26
                sColID = Chr(Asc("A") + lCalc) & sColID
                lCol = lCol - (lCalc * lTmp2)
            Else
                sColID = "A" & sColID
            End If
        Next
    Else
        sColID = Format(lCol + 1, sNumericFormat)
    End If
    ' Join --------------------------------------------
    If m_IDOrderType = Row_Column Then
        If m_UseUnderscore Then
            CalculateID = sRowID & "_" & sColID
        Else
            CalculateID = sRowID & sColID
        End If
    Else
        If m_UseUnderscore Then
            CalculateID = sColID & "_" & sRowID
        Else
            CalculateID = sColID & sRowID
        End If
    End If
End Function

Private Function GetMinimumStringLength(lValue As Long, lBase As Long) As Integer
    Dim lTmp As Long, lIndex As Long
    
    On Error GoTo eh
    ' ROW
    lTmp = lBase
    lIndex = 1
    While lValue > (lTmp - 1)
        lTmp = lTmp * lBase
        lIndex = lIndex + 1
    Wend
    GetMinimumStringLength = lIndex
    Exit Function
eh:
    Err.Raise Err.Number, "GetMinimumStringLength", "Error in GetMinimumStringLength: " & Err.Description
End Function

Public Sub RunStandardGUI(pApp As IApplication)
    Set frmGridSettings.m_Application = pApp
    frmGridSettings.Tickle
    SetWindowLong frmGridSettings.hwnd, GWL_HWNDPARENT, pApp.hwnd
    frmGridSettings.Show vbModeless
End Sub

'Private Function CreateGridPoly(pStartPoint As IPoint, lRow As Long, lCol As Long, _
'                                dGridWidth As Double, dGridHeight As Double) As IPolygon
'    Dim pPntColl As IPointCollection
'    Dim pPoint As IPoint
'    Dim dX As Double, dY As Double
'
'    Set CreateGridPoly = New Polygon
'    Set pPntColl = CreateGridPoly
'    Set pPoint = New Point
'    pPoint.PutCoords pStartPoint.X + (lCol * dGridWidth), pStartPoint.Y + (lRow * dGridHeight)
'    pPntColl.AddPoint pPoint
'    Set pPoint = New Point
'    pPoint.PutCoords pStartPoint.X + ((lCol + 1) * dGridWidth), pStartPoint.Y + (lRow * dGridHeight)
'    pPntColl.AddPoint pPoint
'    Set pPoint = New Point
'    pPoint.PutCoords pStartPoint.X + ((lCol + 1) * dGridWidth), pStartPoint.Y + ((lRow + 1) * dGridHeight)
'    pPntColl.AddPoint pPoint
'    Set pPoint = New Point
'    pPoint.PutCoords pStartPoint.X + (lCol * dGridWidth), pStartPoint.Y + ((lRow + 1) * dGridHeight)
'    pPntColl.AddPoint pPoint
'    Set pPoint = New Point
'    pPoint.PutCoords pStartPoint.X + (lCol * dGridWidth), pStartPoint.Y + (lRow * dGridHeight)
'    pPntColl.AddPoint pPoint
'
'    'Set CreateGridPoly = pPntColl
'    Debug.Print CreateGridPoly.IsClosed
'    'Debug.Print CreateGridPoly.Envelope.XMin & "," & CreateGridPoly.Envelope.YMin
'
'End Function

Public Sub GenerateGrids2(Application As IApplication) ', _
                         'Optional bRemoveEmptyGrids As Boolean = False, _
                         'Optional bReplaceExistingGrids As Boolean = False)
    Dim pEditor As IEditor
    Dim pUID As New UID
    Dim pWorkspaceEdit As IWorkspaceEdit
    Dim lLoop As Long
    Dim pFeatDataset As IFeatureDataset
    Dim pFeature As IFeature
    Dim pFeatCur As IFeatureCursor
    Dim pSourcePolygon As IPolygon
    Dim pGridPolygon As IPolygon
    Dim pPointColl As IPointCollection
    Dim pStartingCoord As IPoint
    Dim pPoint As IPoint
    Dim lRow As Long
    Dim lCol As Long
    Dim lRowCount As Long
    Dim lColCount As Long
    Dim pClone As IClone
    Dim dGridSizeW As Double
    Dim dGridSizeH As Double
    Dim pTransform As ITransform2D
    Dim bOKToAdd As Boolean
    Dim iStringLengthRow As Integer
    Dim iStringLengthCol As Integer
    Dim pDataset As IDataset
    Dim lBase As Long
    Dim dDataFrameWidth As Double
    Dim dDataFrameHeight As Double
    Dim dConvertPageToMapUnits As Double
    Dim dIncrement As Double
    Dim pInsertFeatureBuffer As IFeatureBuffer
    Dim pInsertFeatureCursor As IFeatureCursor
    Dim pMx As IMxDocument
    Dim pFL As IFeatureLayer
    Dim pFC As IFeatureClass
    
    On Error GoTo eh
    
    ' Set mouse pointer
    Screen.MousePointer = vbArrowHourglass
    
    ' Init
    Set pMx = Application.Document
    For lLoop = 0 To pMx.FocusMap.LayerCount - 1
        If TypeOf pMx.FocusMap.Layer(lLoop) Is IFeatureLayer Then
            If UCase(pMx.FocusMap.Layer(lLoop).Name) = UCase(m_DestFL.Name) Then
                Set pFL = pMx.FocusMap.Layer(lLoop)
                Exit For
            End If
        End If
    Next
    If pFL Is Nothing Then
        MsgBox "No match in the Map for layer '" & m_DestFL.Name & "'."
        Exit Sub
    End If
    Set pFC = pFL.FeatureClass
    ' Check for required fields - that the field exists
    Dim bErrorWithFields As Boolean
    bErrorWithFields = (pFC.FindField(m_FldID) < 0)
    If Len(m_FldRowNum) > 0 Then bErrorWithFields = bErrorWithFields Or (pFC.FindField(m_FldRowNum) < 0)
    If Len(m_FldColNum) > 0 Then bErrorWithFields = bErrorWithFields Or (pFC.FindField(m_FldColNum) < 0)
    If Len(m_FldScale) > 0 Then bErrorWithFields = bErrorWithFields Or (pFC.FindField(m_FldScale) < 0)
    ' If error
    If bErrorWithFields Then
        Err.Raise vbObjectError, "GenerateGrids", "Could not find all the given field names in " & pFL.Name & "." _
            & vbCrLf & " - " & m_FldID & ", " & m_FldRowNum & ", " & m_FldColNum & ", " & m_FldScale
    End If
    ' Check the field types
    bErrorWithFields = (pFC.Fields.Field(pFC.FindField(m_FldID)).Type <> esriFieldTypeString)
    If Len(m_FldRowNum) > 0 Then
        bErrorWithFields = bErrorWithFields Or _
           ((pFC.Fields.Field(pFC.FindField(m_FldRowNum)).Type <> esriFieldTypeDouble) And _
            (pFC.Fields.Field(pFC.FindField(m_FldRowNum)).Type <> esriFieldTypeInteger) And _
            (pFC.Fields.Field(pFC.FindField(m_FldRowNum)).Type <> esriFieldTypeSingle) And _
            (pFC.Fields.Field(pFC.FindField(m_FldRowNum)).Type <> esriFieldTypeSmallInteger))
    End If
    If Len(m_FldColNum) > 0 Then
        bErrorWithFields = bErrorWithFields Or _
           ((pFC.Fields.Field(pFC.FindField(m_FldColNum)).Type <> esriFieldTypeDouble) And _
            (pFC.Fields.Field(pFC.FindField(m_FldColNum)).Type <> esriFieldTypeInteger) And _
            (pFC.Fields.Field(pFC.FindField(m_FldColNum)).Type <> esriFieldTypeSingle) And _
            (pFC.Fields.Field(pFC.FindField(m_FldColNum)).Type <> esriFieldTypeSmallInteger))
    End If
    If Len(m_FldScale) > 0 Then
        bErrorWithFields = bErrorWithFields Or _
           ((pFC.Fields.Field(pFC.FindField(m_FldScale)).Type <> esriFieldTypeDouble) And _
            (pFC.Fields.Field(pFC.FindField(m_FldScale)).Type <> esriFieldTypeInteger) And _
            (pFC.Fields.Field(pFC.FindField(m_FldScale)).Type <> esriFieldTypeSingle) And _
            (pFC.Fields.Field(pFC.FindField(m_FldScale)).Type <> esriFieldTypeSmallInteger))
    End If
    ' if error
    If bErrorWithFields Then
        Err.Raise vbObjectError, "GenerateGrids", "Given field names are not of the correct type." _
            & vbCrLf & "Grid ID field must be a Text field, all others must be numeric fields."
    End If
    ' Get the dataset and workspace (to start editing upon)
    Set pFeatDataset = pFC.FeatureDataset
    If Not pFeatDataset Is Nothing Then
        Set pWorkspaceEdit = pFeatDataset.Workspace
    Else
        ' Is a shapefile, go via just IDataset
        Set pDataset = pFC
        Set pWorkspaceEdit = pDataset.Workspace
    End If
    dDataFrameWidth = m_dFrameWidthInPageUnits
    dDataFrameHeight = m_dFrameHeightInPageUnits
    ' Start Editing
    pWorkspaceEdit.StartEditing False
    pWorkspaceEdit.StartEditOperation
    
    ' If replacing, delete all existing polygons
    If m_RemoveGrids Then
        Dim pFCu As IFeatureCursor
        Dim pT As ITable
        Set pFCu = m_DestFL.Search(Nothing, False)
        Set pT = m_DestFL.FeatureClass
        frmProgress.ProgressBar1.Min = 0
        frmProgress.ProgressBar1.Max = 100
        If pT.RowCount(Nothing) = 0 Then
            dIncrement = 99
        Else
            dIncrement = 100 / pT.RowCount(Nothing)
        End If
        frmProgress.ProgressBar1.Value = 0
        frmProgress.lblInformation.Caption = "Deleting previous grids..."
        frmProgress.cmdCancel.Visible = False        ' User cannot cancel this step
        frmProgress.Show vbModeless
        Set pFeature = pFCu.NextFeature
        While Not pFeature Is Nothing
            pFeature.Delete
            If (frmProgress.ProgressBar1.Value + dIncrement) <= frmProgress.ProgressBar1.Max Then
                frmProgress.ProgressBar1.Value = frmProgress.ProgressBar1.Value + dIncrement
            Else
                frmProgress.lblInformation.Caption = "Warning: Val > Max (" & frmProgress.ProgressBar1.Max & ")"
            End If
            Set pFeature = pFCu.NextFeature
        Wend
        frmProgress.Hide
    End If
    
    ' Calc the row/column extents, grid size (map units), ID lengths and starting coordinate
    Set pStartingCoord = New Point
    pStartingCoord.PutCoords m_StartX, m_StartY
    dConvertPageToMapUnits = CalculatePageToMapRatio(Application)
    dGridSizeW = ((m_dMapScale * dDataFrameWidth) / dConvertPageToMapUnits)
    dGridSizeH = ((m_dMapScale * dDataFrameHeight) / dConvertPageToMapUnits)
    If Not (pFL.FeatureClass.FeatureDataset Is Nothing) Then
        CalculateRowColCounts m_StartX, m_StartY, m_EndX, m_EndY, _
            dGridSizeW, dGridSizeH, GetValidExtentForLayer(m_DestFL), lColCount, lRowCount
    Else
        CalculateRowColCounts m_StartX, m_StartY, m_EndX, m_EndY, _
            dGridSizeW, dGridSizeH, Nothing, lColCount, lRowCount
    End If
    If lColCount = 0 Or lRowCount = 0 Then
        Err.Raise vbObjectError, "GenerateGrids", "CalculateRowColCounts returned zero row/columns"
    End If
    If m_ColIDType = Alphabetical Then
        lBase = 26
    Else
        lBase = 10
    End If
    iStringLengthCol = GetMinimumStringLength(lColCount, lBase)
    If m_RowIDType = Alphabetical Then
        lBase = 26
    Else
        lBase = 10
    End If
    iStringLengthRow = GetMinimumStringLength(lRowCount, lBase)
    
    ' Create and add the new Grid polygon features
    bOKToAdd = True
    frmProgress.ProgressBar1.Min = 0
    frmProgress.ProgressBar1.Max = 100
    frmProgress.ProgressBar1.Value = 0
    frmProgress.lblInformation.Caption = "Creating new Grid features..."
    frmProgress.cmdCancel.Visible = True
    frmProgress.Show vbModeless
    dIncrement = 99.9 / (lRowCount * lColCount)
    Set pInsertFeatureCursor = pFC.Insert(True)
    Set pInsertFeatureBuffer = pFC.CreateFeatureBuffer
    For lRow = 0 To lRowCount - 1
        For lCol = 0 To lColCount - 1
            ' Create the source polygon
            Set pGridPolygon = CreateGridPoly2(pStartingCoord, lRow, lCol, dGridSizeW, dGridSizeH)
            ' If required, check for containing features
            If m_NoEmptyGrids Then
                bOKToAdd = HasFeatures(pGridPolygon, pMx.FocusMap)
            End If
            If bOKToAdd Then
                ' Create new grid cell feature
                Set pInsertFeatureBuffer.Shape = pGridPolygon
                If m_StartIDType = TopLeft Then
                    pInsertFeatureBuffer.Value(pFC.Fields.FindField(m_FldID)) = _
                            CalculateID((lRowCount - 1) - lRow, lCol, iStringLengthRow, iStringLengthCol)
                    If Len(m_FldRowNum) > 0 Then pInsertFeatureBuffer.Value(pFC.Fields.FindField(m_FldRowNum)) = (lRowCount - lRow)
                Else
                    pInsertFeatureBuffer.Value(pFC.Fields.FindField(m_FldID)) = _
                            CalculateID(lRow, lCol, iStringLengthRow, iStringLengthCol)
                    If Len(m_FldRowNum) > 0 Then pInsertFeatureBuffer.Value(pFC.Fields.FindField(m_FldRowNum)) = (lRow + 1)
                End If
                If Len(m_FldColNum) > 0 Then pInsertFeatureBuffer.Value(pFC.Fields.FindField(m_FldColNum)) = (lCol + 1)
                If Len(m_FldScale) > 0 Then pInsertFeatureBuffer.Value(pFC.Fields.FindField(m_FldScale)) = m_dMapScale
                pInsertFeatureCursor.InsertFeature pInsertFeatureBuffer
            End If
            If frmProgress.ProgressBar1.Value < frmProgress.ProgressBar1.Max Then
                frmProgress.ProgressBar1.Value = frmProgress.ProgressBar1.Value + dIncrement
            Else
                frmProgress.lblInformation.Caption = "Warning: Val > Max (" & frmProgress.ProgressBar1.Max & ")"
            End If
            If ((lRow * lColCount) + lCol) Mod 20 = 0 Then
                DoEvents
                pInsertFeatureCursor.Flush
            End If
            frmProgress.Refresh
            If frmProgress.Cancelled Then
                Dim vUserChoice
                frmProgress.Cancelled = False       ' Reset the form
                vUserChoice = MsgBox("Operation cancelled." _
                    & "  Save the edits made thus far?" & vbCrLf & vbCrLf _
                    & "(Click Cancel to continue processing)", _
                                vbYesNoCancel, "Generate Grids")
                If vUserChoice <> vbCancel Then
                    GoTo CancelledGenerateGrids     'Sorry for GoTo usage - in a hurry
                End If
            End If
        Next
    Next
    pInsertFeatureCursor.Flush
    frmProgress.Hide
    
    ' Stop editing
    pWorkspaceEdit.StopEditOperation
    pWorkspaceEdit.StopEditing True
    ' ----------------------------------------------------
    
    Screen.MousePointer = vbDefault
    pMx.ActiveView.Refresh
    
    Exit Sub
    
CancelledGenerateGrids:
    frmProgress.Hide
    If vUserChoice = vbYes Then
        pInsertFeatureCursor.Flush
        pWorkspaceEdit.StopEditOperation
        pWorkspaceEdit.StopEditing True
    Else
        pWorkspaceEdit.StopEditOperation
        pWorkspaceEdit.StopEditing False
    End If
    Screen.MousePointer = vbDefault
    pMx.ActiveView.Refresh
    Exit Sub
    
    Resume
eh:
    frmProgress.Hide
    Screen.MousePointer = vbDefault
    If Not pWorkspaceEdit Is Nothing Then
        If pWorkspaceEdit.IsBeingEdited Then
            pWorkspaceEdit.StopEditOperation
            pWorkspaceEdit.StopEditing False
        End If
    End If
    MsgBox "Error " & Err.Number & " : " & Err.Description, vbExclamation, "GenerateGrids"
End Sub


Private Function CreateGridPoly2(pStartPoint As IPoint, lRow As Long, lCol As Long, _
                                dGridWidth As Double, dGridHeight As Double) As IPolygon
    Dim pPntColl As IPointCollection
    Dim pPoint As IPoint
    Dim dX As Double, dY As Double
    
    Set CreateGridPoly2 = New Polygon
    Set pPntColl = CreateGridPoly2
    Set pPoint = New Point
    pPoint.PutCoords pStartPoint.X + (lCol * dGridWidth), pStartPoint.Y + (lRow * dGridHeight)
    pPntColl.AddPoint pPoint
    Set pPoint = New Point
    pPoint.PutCoords pStartPoint.X + (lCol * dGridWidth), pStartPoint.Y + ((lRow + 1) * dGridHeight)
    pPntColl.AddPoint pPoint
    Set pPoint = New Point
    pPoint.PutCoords pStartPoint.X + ((lCol + 1) * dGridWidth), pStartPoint.Y + ((lRow + 1) * dGridHeight)
    pPntColl.AddPoint pPoint
    Set pPoint = New Point
    pPoint.PutCoords pStartPoint.X + ((lCol + 1) * dGridWidth), pStartPoint.Y + (lRow * dGridHeight)
    pPntColl.AddPoint pPoint
    Set pPoint = New Point
    pPoint.PutCoords pStartPoint.X + (lCol * dGridWidth), pStartPoint.Y + (lRow * dGridHeight)
    pPntColl.AddPoint pPoint
    
End Function

Private Sub CalculateRowColCounts(ByVal XStart As Double, ByVal YStart As Double, _
                                  ByVal XEnd As Double, ByVal YEnd As Double, _
                                  ByVal XGridWidth As Double, ByVal YGridHeight As Double, _
                                  ByVal LayerExtent As IEnvelope, _
                                  ByRef ReturnColCount As Long, ByRef ReturnRowCount As Long)
    Dim XRange As Double
    Dim YRange As Double
    Dim dTemp As Double
    Dim lMultiplierX As Long
    Dim lMultiplierY As Long
    
    On Error GoTo eh
    
    ' Protect against div by zero (where the grid width/height is less than 0.5 => ie: with Geographic Data)
    lMultiplierX = 1
    dTemp = XGridWidth
    While dTemp < 10
        dTemp = dTemp * 10
        lMultiplierX = lMultiplierX * 10
    Wend
    lMultiplierY = 1
    dTemp = YGridHeight
    While dTemp < 10
        dTemp = dTemp * 10
        lMultiplierY = lMultiplierY * 10
    Wend
    
    ' Init
    XRange = XEnd - XStart
    YRange = YEnd - YStart
    'X ------------------------------------------
    If Not (LayerExtent Is Nothing) Then
        If XStart < LayerExtent.XMin Then
            Err.Raise vbObjectError, "CalculateRowColCounts", _
                "Starting X (" & XStart & ") is outside the valid layer extent (" & LayerExtent.XMin & ")"
        ElseIf XEnd > LayerExtent.XMax Then
            Err.Raise vbObjectError, "CalculateRowColCounts", _
                "Ending X (" & XStart & ") is outside the valid layer extent (" & LayerExtent.XMax & ")"
        ElseIf (XStart + XGridWidth) > LayerExtent.XMax Then
            Err.Raise vbObjectError, "CalculateRowColCounts", _
                "Adding a single Grid (width " & XGridWidth & ") would exceed valid X range for layer (" & LayerExtent.XMax & ")"
        End If
    End If
    If XRange < XGridWidth Then
        Err.Raise vbObjectError, "CalculateRowColCounts", _
            "Grid width " & XGridWidth & " exceeds range specified (" & XStart & " - " & XEnd & ")"
    ElseIf (XStart + XGridWidth) > XEnd Then
        Err.Raise vbObjectError, "CalculateRowColCounts", _
            "Adding a single Grid (width " & XGridWidth & ") would exceed specified X range (" & XEnd & ")"
    End If
    ReturnColCount = (XRange * lMultiplierX) \ (XGridWidth * lMultiplierX)
    'Y ------------------------------------------
    If Not (LayerExtent Is Nothing) Then
        If YStart < LayerExtent.YMin Then
            Err.Raise vbObjectError, "CalculateRowColCounts", _
                "Starting Y (" & YStart & ") is outside the valid layer extent (" & LayerExtent.YMin & ")"
        ElseIf YEnd > LayerExtent.YMax Then
            Err.Raise vbObjectError, "CalculateRowColCounts", _
                "Ending Y (" & YStart & ") is outside the valid layer extent (" & LayerExtent.YMax & ")"
        ElseIf (YStart + YGridHeight) > LayerExtent.YMax Then
            Err.Raise vbObjectError, "CalculateRowColCounts", _
                "Adding a single Grid (height " & YGridHeight & ") would exceed valid Y range for layer (" & LayerExtent.YMax & ")"
        End If
    End If
    If YRange < YGridHeight Then
        Err.Raise vbObjectError, "CalculateRowColCounts", _
            "Grid height " & YGridHeight & " exceeds range specified (" & YStart & " - " & YEnd & ")"
    ElseIf (XStart + XGridWidth) > XEnd Then
        Err.Raise vbObjectError, "CalculateRowColCounts", _
            "Adding a single Grid (height " & YGridHeight & ") would exceed specified Y range (" & YEnd & ")"
    End If
    ReturnRowCount = (YRange * lMultiplierY) \ (YGridHeight * lMultiplierY)
    '--------------------------------------------
    
    Exit Sub
eh:
    Err.Raise Err.Number, Err.Source, "Error in CalculateRowColCounts: " & Err.Description
End Sub

Private Function CalculatePageToMapRatio(pApp As IApplication) As Double
    Dim pMx As IMxDocument
    Dim pPage As IPage
    Dim pPageUnits As esriUnits
    Dim pSR As ISpatialReference
    Dim pSRI As ISpatialReferenceInfo
    Dim pPCS As IProjectedCoordinateSystem
    Dim dMetersPerUnit As Double
    Dim dCurrScale As Double
    Dim pExtentEnv As IEnvelope
    
    On Error GoTo eh
    
    ' Init
    Set pMx = pApp.Document
    Set pSR = pMx.FocusMap.SpatialReference
    ' If a Projected coord system
    If TypeOf pSR Is IProjectedCoordinateSystem Then
        ' Use meters per unit as the conversion
        Set pPCS = pSR
        dMetersPerUnit = pPCS.CoordinateUnit.MetersPerUnit
        ' Now convert this into page (ie: paper) units
        Set pPage = pMx.PageLayout.Page
        pPageUnits = pPage.Units
        Select Case pPageUnits
            Case esriInches: CalculatePageToMapRatio = dMetersPerUnit / (1 / 12 * 0.304800609601219)
            Case esriFeet: CalculatePageToMapRatio = dMetersPerUnit / (0.304800609601219)
            Case esriCentimeters: CalculatePageToMapRatio = dMetersPerUnit / (1 / 100)
            Case esriMeters: CalculatePageToMapRatio = dMetersPerUnit / (1)
            Case Else:
                MsgBox "Warning: Only the following Page (Layout) Units are supported by this tool:" _
                    & vbCrLf & " - Inches, Feet, Centimeters, Meters" _
                    & vbCrLf & vbCrLf & "Calculating as though Page Units are in Inches..."
                CalculatePageToMapRatio = dMetersPerUnit / (1 / 12 * 0.304800609601219)
        End Select
    ' Otherwise
    Else
        ' If not projected, we can only do a "flat" conversion -> that is, use the current scale and extent
        '  as a ratio to be applied to the map grid scale.
        ' NOTE: We MUST be in Layout mode to make this calculation, as the scale in Map View and Layout View
        '  are not the same (as the extent envelope and data frame envelope can be different shapes).  The
        '  test for being in Layout Mode is made in the clsMapGridButton.ICommand_Enabled property.
        Set pExtentEnv = pMx.ActiveView.Extent
        dCurrScale = pMx.FocusMap.MapScale
        If ((m_EndX - m_StartX) / m_dFrameWidthInPageUnits) > ((m_EndY - m_StartY) / m_dFrameHeightInPageUnits) Then
            CalculatePageToMapRatio = m_dFrameWidthInPageUnits / ((m_EndX - m_StartX) / dCurrScale)
        Else
            CalculatePageToMapRatio = m_dFrameHeightInPageUnits / ((m_EndY - m_StartY) / dCurrScale)
        End If
    End If
    
    Exit Function
eh:
    CalculatePageToMapRatio = 1
    MsgBox "Error in CalculatePageToMapRatio" & vbCrLf & Err.Description
End Function

Private Function HasFeatures(pPolygon As IPolygon, pMap As IMap) As Boolean
    Dim lLoop As Long
    Dim pFL As IFeatureLayer
    Dim pSF As ISpatialFilter
    Dim pFC As IFeatureCursor
    Dim pF As IFeature
    
    HasFeatures = False
    Set pSF = New SpatialFilter
    Set pSF.Geometry = pPolygon
    pSF.SpatialRel = esriSpatialRelIntersects
    For lLoop = 0 To (pMap.LayerCount - 1)
        If TypeOf pMap.Layer(lLoop) Is IFeatureLayer Then
            Set pFL = pMap.Layer(lLoop)
            If pFL.Name <> m_DestFL.Name And IsARequiredLayer(pFL.Name) Then
                pSF.GeometryField = pFL.FeatureClass.shapeFieldName
                Set pFC = pFL.Search(pSF, False)
                Set pF = pFC.NextFeature
                If Not pF Is Nothing Then
                    HasFeatures = True
                    Exit For
                End If
            End If
        End If
    Next
End Function

Private Function IsARequiredLayer(sLayerName As String) As Boolean
    Dim lLoop As Long
    For lLoop = 1 To m_colLayerNamesForData.count
        If UCase(m_colLayerNamesForData.Item(lLoop)) = UCase(sLayerName) Then
            IsARequiredLayer = True
            Exit For
        End If
    Next
End Function

Private Sub Class_Terminate()
    Set m_DestFL = Nothing
    Set m_DestFC = Nothing
End Sub

