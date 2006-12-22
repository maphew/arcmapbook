VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IDSMapPage"
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

Public Property Get PageName() As String

End Property

Public Property Let PageName(PageName As String)

End Property

Public Property Get PageNumber() As Long

End Property

Public Property Let PageNumber(PageNumber As Long)

End Property

Public Property Get PageRotation() As Double

End Property

Public Property Let PageRotation(PageRotation As Double)

End Property

Public Property Get PageScale() As Double

End Property

Public Property Let PageScale(PageScale As Double)

End Property

Public Property Get PageShape() As IPolygon

End Property

Public Property Set PageShape(PageShape As IPolygon)

End Property

Public Property Get LastOutputted() As Date

End Property

Public Property Let LastOutputted(LastOutputted As Date)

End Property

Public Property Get EnablePage() As Boolean

End Property

Public Property Let EnablePage(PrintPage As Boolean)

End Property

Public Property Get PageItemCount() As Long

End Property

Public Property Get PageItem(Index As Long) As IElement

End Property

Public Sub AddPageItem(ByVal PageItem As IElement)

End Sub

Public Sub RemovePageItem(Index As Long)

End Sub

Public Sub DrawPage(pDoc As IMxDocument, pDSMapSeries As IDSMapSeries, bRefreshFlag As Boolean)

End Sub

Public Function IndexPage(pIndexLayer As IFeatureLayer, sFieldName As String) As Collection

End Function


