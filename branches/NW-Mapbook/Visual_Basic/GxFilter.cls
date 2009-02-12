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
  If fs.FileExists(pTruncPath & ".shp") Or fs.FileExists(pTruncPath & ".dbf") Or _
   fs.FileExists(pTruncPath & ".shx") Then
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
