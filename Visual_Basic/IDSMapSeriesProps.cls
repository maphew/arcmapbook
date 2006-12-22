VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IDSMapSeriesProps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

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

Public Property Get DataFrameName() As String

End Property

Public Property Let DataFrameName(DataFrame As String)

End Property

Public Property Get IndexLayerName() As String

End Property

Public Property Let IndexLayerName(LayerName As String)

End Property

Public Property Get IndexFieldName() As String

End Property

Public Property Let IndexFieldName(FieldName As String)

End Property

Public Property Let SuppressLayers(ByVal vdata As Boolean)

End Property

Public Property Get SuppressLayers() As Boolean

End Property

Public Property Get SuppressLayerCount() As Long

End Property

Public Property Get SuppressLayer(Index As Long) As String

End Property

Public Sub AddLayerToSuppress(ByVal LayerName As String)

End Sub

Public Sub RemoveLayerToSuppress(Index As Long)

End Sub

Public Property Get TileSelectionMethod() As Long

End Property

Public Property Let TileSelectionMethod(TileSelection As Long)

End Property

Public Property Get StartNumber() As Long

End Property

Public Property Let StartNumber(lStart As Long)

End Property
