VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IDSMapSeriesOptions"
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

Public Property Get ExtentType() As Long

End Property

Public Property Let ExtentType(ExtentType As Long)

End Property

Public Property Get Margin() As Double

End Property

Public Property Let Margin(Margin As Double)

End Property

Public Property Get MarginType() As String

End Property

Public Property Let MarginType(MarginType As String)

End Property

Public Property Get FixedScale() As Double

End Property

Public Property Let FixedScale(FixedScale As Double)

End Property

Public Property Get DataDrivenField() As String

End Property

Public Property Let DataDrivenField(DataDrivenField As String)

End Property

Public Property Get RotateFrame() As Boolean

End Property

Public Property Let RotateFrame(RotateFrame As Boolean)

End Property

Public Property Get RotationField() As String

End Property

Public Property Let RotationField(RotationField As String)

End Property

Public Property Get ClipData() As Boolean

End Property

Public Property Let ClipData(ClipData As Boolean)

End Property

Public Property Get LabelNeighbors() As Boolean

End Property

Public Property Let LabelNeighbors(LabelNeighbors As Boolean)

End Property

Public Property Get LabelSymbol() As ISymbol

End Property

Public Property Set LabelSymbol(LabelSymbol As ISymbol)

End Property
