VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NWLayerVisibilityGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'Layer visibility groups.  This object holds and maintains lists
'of those layers that are not visible.  This should be less costly
'than lists of layers that are visible, because an initial condition
'of an empty list coincides with an initial condition of all layers
'being visible.



Implements INWLayerVisibilityGroup
Implements IPersistVariant


Private m_pDictInvisLayers As Scripting.Dictionary
Private m_pUID_FlexiblePersist As UID
Private m_pUID_NWData1 As UID
Private m_sName As String
Const c_sModuleFileName As String = "NWLayerVisibilityGroup.cls"




Private Sub Class_Initialize()
  Set m_pDictInvisLayers = New Scripting.Dictionary
  If m_pUID_FlexiblePersist Is Nothing Then
    Set m_pUID_FlexiblePersist = New UID
    m_pUID_FlexiblePersist.Value = "{42383075-B5C8-4c2a-85FB-DE9CDC535FA6}"
  End If
  If m_pUID_NWData1 Is Nothing Then
    Set m_pUID_NWData1 = New UID
    m_pUID_NWData1.Value = "{E0A55865-A58F-48d7-8DDF-B4DB5997322F}"
  End If
End Sub

Private Sub INWLayerVisibilityGroup_AddLayer(ByVal sLayerName As String)
  If Not m_pDictInvisLayers.Exists(sLayerName) Then
    m_pDictInvisLayers.Add sLayerName, Nothing
  End If
End Sub

Private Sub INWLayerVisibilityGroup_DeleteLayer(ByVal sLayerName As String)
  If m_pDictInvisLayers.Exists(sLayerName) Then
    m_pDictInvisLayers.Remove sLayerName
  End If
End Sub

Private Function INWLayerVisibilityGroup_Exists(ByVal sLayerName As String) As Boolean
  INWLayerVisibilityGroup_Exists = m_pDictInvisLayers.Exists(sLayerName)
End Function

Private Function INWLayerVisibilityGroup_InvisibleLayerCount() As Long
  INWLayerVisibilityGroup_InvisibleLayerCount = m_pDictInvisLayers.Count
End Function

Private Function INWLayerVisibilityGroup_InvisibleLayers() As Variant
  INWLayerVisibilityGroup_InvisibleLayers = m_pDictInvisLayers.Keys
End Function

Private Property Let INWLayerVisibilityGroup_Name(RHS As String)
  m_sName = RHS
End Property

Private Property Get INWLayerVisibilityGroup_Name() As String
  INWLayerVisibilityGroup_Name = m_sName
End Property

Private Sub INWLayerVisibilityGroup_SetLayerInvisible(ByVal sLayerName As String)
  If Not m_pDictInvisLayers.Exists(sLayerName) Then
    m_pDictInvisLayers.Add sLayerName, Nothing
  End If
End Sub

Private Sub INWLayerVisibilityGroup_SetLayerVisible(ByVal sLayerName As String)
  If m_pDictInvisLayers.Exists(sLayerName) Then
    m_pDictInvisLayers.Remove (sLayerName)
  End If
End Sub

Private Property Get IPersistVariant_ID() As esriSystem.IUID
  Dim id As New UID
  id = "NWMapBookPrj.NWLayerVisibilityGroup"
  Set IPersistVariant_ID = id
End Property





Private Sub IPersistVariant_Load(ByVal Stream As esriSystem.IVariantStream)
  On Error GoTo ErrorHandler

'load the visible layers data for this group

  Dim lLoop As Long, lCount As Long, bExitFlagFound As Boolean
  Dim vNWData As Variant, sUID1 As String, sUID2 As String
  Dim bFoundUID As Boolean, lInvisLayerCount As Long, sTextSymbKey As String
  
  'point of custom load code for NW
  'Mapbook customization properties.
  '{839079CF-45B2-47ee-8901-988D5E05147B}
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
  
  bExitFlagFound = False
  vNWData = Stream.Read
  If Not IsEmpty(vNWData) Then
                                                  'for whatever reason, saved UIDs
                                                  'persist as strings
    If StrComp(UCase(TypeName(vNWData)), "String", vbTextCompare) = 0 Then
      sUID1 = vNWData
                                                  'if the GUID value is correct
      If StrComp(sUID1, m_pUID_FlexiblePersist.Value, vbTextCompare) = 0 Then
                                                  'then read the persisted data values
        bFoundUID = False
        
        '{guid 1}
        '  {guid 2} <--- search for this section
        '  ... data ...
        '  {guid 2}
        '{guid 1}
        Do While (Not bFoundUID) And (Not bExitFlagFound)
          vNWData = Stream.Read
          If StrComp(UCase(TypeName(vNWData)), "String", vbTextCompare) = 0 Then
            sUID2 = vNWData
            If StrComp(sUID2, m_pUID_NWData1.Value, vbTextCompare) = 0 Then
              bFoundUID = True
            ElseIf StrComp(sUID2, m_pUID_FlexiblePersist.Value, vbTextCompare) = 0 Then
              bExitFlagFound = True
            End If
          End If
        Loop
                                                  'once the bracketing UIDs
                                                  'have been detected, read the
                                                  'data for this version of the
                                                  'NW Mapbook application
        If bFoundUID And Not bExitFlagFound Then
          lInvisLayerCount = Stream.Read                     'count of how many text symbols exist
          m_pDictInvisLayers.RemoveAll
          For lLoop = 0 To (lInvisLayerCount - 1)
            vNWData = Stream.Read
            sTextSymbKey = vNWData
                                                  'initially does not have items in
                                                  'the dict, only keys, but this code is
                                                  'included to support future changes
            Stream.Read
            m_pDictInvisLayers.Add sTextSymbKey, Nothing
          Next lLoop
        End If
        
        '{guid 1}
        '  {guid 2}
        '  ... data ...
        '  {guid 2}      <--- search for this section
        '{guid 1}
        bFoundUID = False
        Do While (Not bFoundUID) And (Not bExitFlagFound)
          vNWData = Stream.Read
          If StrComp(UCase(TypeName(vNWData)), "String", vbTextCompare) = 0 Then
            sUID2 = vNWData
            If StrComp(sUID2, m_pUID_NWData1.Value, vbTextCompare) = 0 Then
              bFoundUID = True
            ElseIf StrComp(sUID2, m_pUID_FlexiblePersist.Value, vbTextCompare) = 0 Then
              bExitFlagFound = True
            End If
          End If
        Loop
        
        '{guid 1}
        '  {guid 2}
        '  ... data ...
        '  {guid 2}
        '{guid 1}     <---- search for the closing section
        Do While Not bExitFlagFound
          vNWData = Stream.Read
          If StrComp(UCase(TypeName(vNWData)), "String", vbTextCompare) = 0 Then
            sUID2 = vNWData
            If StrComp(sUID2, m_pUID_FlexiblePersist.Value, vbTextCompare) = 0 Then
              bExitFlagFound = True
            End If
          End If
        Loop
        'The above structure is made so that future data can be saved ...
        '{guid 1}
        '  <-- future data, presumably enclosed within {guid3}..{guid3}, {guid4}.. etc. -->
        '  {guid 2}
        '    ... data ...
        '    <-- future data for the NW Mapbook component since stream reads keep going until {guid2} -->
        '  {guid 2}
        '  <-- future data, presumably enclosed within {guid5}..{guid5}, {guid6}.. etc. -->
        '{guid 1}     <---- search for the closing section
      
      End If
    Else
          'assume that I read something from the next
          'section of load code.
      Stream.Write vNWData
    End If
  End If
    

  Exit Sub
ErrorHandler:
  HandleError True, "IPersistVariant_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub


Private Sub IPersistVariant_Save(ByVal Stream As esriSystem.IVariantStream)
  On Error GoTo ErrorHandler

  Dim i As Long, lInvisLayerCount As Long, vNWData As Variant
  Dim sInvisLyrName As String
  
  If m_pUID_FlexiblePersist Is Nothing Then
    Set m_pUID_FlexiblePersist = New UID
    m_pUID_FlexiblePersist.Value = "{42383075-B5C8-4c2a-85FB-DE9CDC535FA6}"
  End If
  If m_pUID_NWData1 Is Nothing Then
    Set m_pUID_NWData1 = New UID
    m_pUID_NWData1.Value = "{E0A55865-A58F-48d7-8DDF-B4DB5997322F}"
  End If
  
  Stream.Write m_pUID_FlexiblePersist.Value
  Stream.Write m_pUID_NWData1.Value
  lInvisLayerCount = m_pDictInvisLayers.Count
  Stream.Write lInvisLayerCount
  vNWData = m_pDictInvisLayers.Keys
  For i = 0 To (lInvisLayerCount - 1)
    sInvisLyrName = vNWData(i)
    Stream.Write sInvisLyrName
    Stream.Write m_pDictInvisLayers(vNWData(i))
  Next i
  Stream.Write m_pUID_NWData1.Value
  Stream.Write m_pUID_FlexiblePersist.Value


  Exit Sub
ErrorHandler:
  HandleError True, "IPersistVariant_Save " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub
