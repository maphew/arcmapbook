VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GxFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IGxObjectFilter

Private Function IGxObjectFilter_CanChooseObject(ByVal Object As IGxObject, result As esriDoubleClickResult) As Boolean

End Function

Private Function IGxObjectFilter_CanDisplayObject(ByVal Object As IGxObject) As Boolean
  Select Case UCase(Object.Category)
  Case "COVERAGE"
    IGxObjectFilter_CanDisplayObject = False
  Case Else
    IGxObjectFilter_CanDisplayObject = True
  End Select

End Function

Private Function IGxObjectFilter_CanSaveObject(ByVal location As IGxObject, ByVal newObjectName As String, objectAlreadyExists As Boolean) As Boolean
On Error GoTo ErrHand:
  Select Case UCase(location.Category)
  Case "PERSONAL GEODATABASE FEATURE DATASET"   ', "SDE FEATURE DATASET"
    objectAlreadyExists = DoesFeatureClassExist(location, newObjectName)
    IGxObjectFilter_CanSaveObject = Not objectAlreadyExists
  Case "FOLDER"
    objectAlreadyExists = DoesShapeFileExist(location.FullName & "\" & newObjectName)
    IGxObjectFilter_CanSaveObject = Not objectAlreadyExists
  Case Else
    IGxObjectFilter_CanSaveObject = False
  End Select
  
  Exit Function
ErrHand:
  MsgBox Err.Description & " -" & newObjectName & "-"
End Function

Private Property Get IGxObjectFilter_Description() As String
  IGxObjectFilter_Description = "Shapefile / Feature Class"
End Property

Private Property Get IGxObjectFilter_Name() As String
  IGxObjectFilter_Name = "Overlay filter"
End Property

Private Function DoesShapeFileExist(pPath As String) As Boolean
  Dim pTruncPath As String
  If InStr(1, pPath, ".shp") > 0 Then
    pTruncPath = Left(pPath, InStr(1, pPath, ".shp") - 1)
  Else
    pTruncPath = pPath
  End If
      
  'Make sure the specified file does not exist
  Dim fs As Object
  Set fs = CreateObject("Scripting.FileSystemObject")
  If fs.fileexists(pTruncPath & ".shp") Or fs.fileexists(pTruncPath & ".dbf") Or _
   fs.fileexists(pTruncPath & ".shx") Then
    DoesShapeFileExist = True
  Else
    DoesShapeFileExist = False
  End If
End Function

Private Function DoesFeatureClassExist(location As IGxObject, newObjectName As String) As Boolean
On Error GoTo ErrHand:
  Dim pFeatClass As IFeatureClass
  Dim pFeatDataset As IGxDataset
  Set pFeatDataset = location
  Dim pFeatClassCont As IFeatureClassContainer, pData As IFeatureDataset
  Set pData = pFeatDataset.Dataset
  Set pFeatClassCont = pData
  Dim pEnumClass As IEnumFeatureClass, pDataset As IDataset
  Set pEnumClass = pFeatClassCont.Classes
  Set pFeatClass = pEnumClass.Next
  While Not pFeatClass Is Nothing
    Set pDataset = pFeatClass
    If UCase(pDataset.Name) = UCase(newObjectName) Then
      DoesFeatureClassExist = True
      Exit Function
    End If
      
    Set pFeatClass = pEnumClass.Next
  Wend
  DoesFeatureClassExist = False
  
  Exit Function
ErrHand:
  MsgBox Err.Description
End Function
