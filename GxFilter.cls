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

' Copyright 2006 ESRI
'
' All rights reserved under the copyright laws of the United States
' and applicable international laws, treaties, and conventions.
'
' You may freely redistribute and use this sample code, with or
' without modification, provided you include the original copyright
' notice and use restrictions.
'
' See use restrictions at /arcgis/developerkit/userestrictions.

Option Explicit

Implements IGxObjectFilter

Private Function IGxObjectFilter_CanChooseObject(ByVal Object As IGxObject, result As esriDoubleClickResult) As Boolean

End Function

Private Function IGxObjectFilter_CanDisplayObject(ByVal Object As IGxObject) As Boolean
  Select Case UCase(Object.Category)
  Case "COVERAGE"
24:     IGxObjectFilter_CanDisplayObject = False
  Case Else
26:     IGxObjectFilter_CanDisplayObject = True
27:   End Select

End Function

Private Function IGxObjectFilter_CanSaveObject(ByVal location As IGxObject, ByVal newObjectName As String, objectAlreadyExists As Boolean) As Boolean
On Error GoTo ErrHand:
  Select Case UCase(location.Category)
  Case "PERSONAL GEODATABASE FEATURE DATASET"   ', "SDE FEATURE DATASET"
35:     objectAlreadyExists = DoesFeatureClassExist(location, newObjectName)
36:     IGxObjectFilter_CanSaveObject = Not objectAlreadyExists
  Case "FOLDER"
38:     objectAlreadyExists = DoesShapeFileExist(location.FullName & "\" & newObjectName)
39:     IGxObjectFilter_CanSaveObject = Not objectAlreadyExists
  Case Else
41:     IGxObjectFilter_CanSaveObject = False
42:   End Select
  
  Exit Function
ErrHand:
46:   MsgBox Err.Description & " -" & newObjectName & "-"
End Function

Private Property Get IGxObjectFilter_Description() As String
50:   IGxObjectFilter_Description = "Shapefile / Feature Class"
End Property

Private Property Get IGxObjectFilter_Name() As String
54:   IGxObjectFilter_Name = "Overlay filter"
End Property

Private Function DoesShapeFileExist(pPath As String) As Boolean
  Dim pTruncPath As String
59:   If InStr(1, pPath, ".shp") > 0 Then
60:     pTruncPath = Left(pPath, InStr(1, pPath, ".shp") - 1)
61:   Else
62:     pTruncPath = pPath
63:   End If
      
  'Make sure the specified file does not exist
  Dim fs As Object
67:   Set fs = CreateObject("Scripting.FileSystemObject")
68:   If fs.fileexists(pTruncPath & ".shp") Or fs.fileexists(pTruncPath & ".dbf") Or _
   fs.fileexists(pTruncPath & ".shx") Then
70:     DoesShapeFileExist = True
71:   Else
72:     DoesShapeFileExist = False
73:   End If
End Function

Private Function DoesFeatureClassExist(location As IGxObject, newObjectName As String) As Boolean
On Error GoTo ErrHand:
  Dim pFeatClass As IFeatureClass
  Dim pFeatDataset As IGxDataset
80:   Set pFeatDataset = location
  Dim pFeatClassCont As IFeatureClassContainer, pData As IFeatureDataset
82:   Set pData = pFeatDataset.Dataset
83:   Set pFeatClassCont = pData
  Dim pEnumClass As IEnumFeatureClass, pDataset As IDataset
85:   Set pEnumClass = pFeatClassCont.Classes
86:   Set pFeatClass = pEnumClass.Next
87:   While Not pFeatClass Is Nothing
88:     Set pDataset = pFeatClass
89:     If UCase(pDataset.Name) = UCase(newObjectName) Then
90:       DoesFeatureClassExist = True
      Exit Function
92:     End If
      
94:     Set pFeatClass = pEnumClass.Next
95:   Wend
96:   DoesFeatureClassExist = False
  
  Exit Function
ErrHand:
100:   MsgBox Err.Description
End Function
