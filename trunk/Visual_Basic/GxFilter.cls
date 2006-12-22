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
12:     IGxObjectFilter_CanDisplayObject = False
  Case Else
14:     IGxObjectFilter_CanDisplayObject = True
15:   End Select

End Function

Private Function IGxObjectFilter_CanSaveObject(ByVal location As IGxObject, ByVal newObjectName As String, objectAlreadyExists As Boolean) As Boolean
On Error GoTo ErrHand:
  Select Case UCase(location.Category)
  Case "PERSONAL GEODATABASE FEATURE DATASET"   ', "SDE FEATURE DATASET"
23:     objectAlreadyExists = DoesFeatureClassExist(location, newObjectName)
24:     IGxObjectFilter_CanSaveObject = Not objectAlreadyExists
  Case "FOLDER"
26:     objectAlreadyExists = DoesShapeFileExist(location.FullName & "\" & newObjectName)
27:     IGxObjectFilter_CanSaveObject = Not objectAlreadyExists
  Case Else
29:     IGxObjectFilter_CanSaveObject = False
30:   End Select
  
  Exit Function
ErrHand:
34:   MsgBox Err.Description & " -" & newObjectName & "-"
End Function

Private Property Get IGxObjectFilter_Description() As String
38:   IGxObjectFilter_Description = "Shapefile / Feature Class"
End Property

Private Property Get IGxObjectFilter_Name() As String
42:   IGxObjectFilter_Name = "Overlay filter"
End Property

Private Function DoesShapeFileExist(pPath As String) As Boolean
  Dim pTruncPath As String
47:   If InStr(1, pPath, ".shp") > 0 Then
48:     pTruncPath = Left(pPath, InStr(1, pPath, ".shp") - 1)
49:   Else
50:     pTruncPath = pPath
51:   End If
      
  'Make sure the specified file does not exist
  Dim fs As Object
55:   Set fs = CreateObject("Scripting.FileSystemObject")
56:   If fs.fileexists(pTruncPath & ".shp") Or fs.fileexists(pTruncPath & ".dbf") Or _
   fs.fileexists(pTruncPath & ".shx") Then
58:     DoesShapeFileExist = True
59:   Else
60:     DoesShapeFileExist = False
61:   End If
End Function

Private Function DoesFeatureClassExist(location As IGxObject, newObjectName As String) As Boolean
On Error GoTo ErrHand:
  Dim pFeatClass As IFeatureClass
  Dim pFeatDataset As IGxDataset
68:   Set pFeatDataset = location
  Dim pFeatClassCont As IFeatureClassContainer, pData As IFeatureDataset
70:   Set pData = pFeatDataset.Dataset
71:   Set pFeatClassCont = pData
  Dim pEnumClass As IEnumFeatureClass, pDataset As IDataset
73:   Set pEnumClass = pFeatClassCont.Classes
74:   Set pFeatClass = pEnumClass.Next
75:   While Not pFeatClass Is Nothing
76:     Set pDataset = pFeatClass
77:     If UCase(pDataset.Name) = UCase(newObjectName) Then
78:       DoesFeatureClassExist = True
      Exit Function
80:     End If
      
82:     Set pFeatClass = pEnumClass.Next
83:   Wend
84:   DoesFeatureClassExist = False
  
  Exit Function
ErrHand:
88:   MsgBox Err.Description
End Function
