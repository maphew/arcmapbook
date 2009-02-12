VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NWDSMapBookBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"

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

'
' Toolbar containing the Split and Combine commands.
'
Option Explicit

Implements IToolBarDef

Private Property Get IToolBarDef_Caption() As String
  IToolBarDef_Caption = "NW Map Series"
End Property

Private Sub IToolBarDef_GetItemInfo(ByVal pos As Long, ByVal itemDef As IItemDef)
  Select Case pos
  Case 0
    itemDef.ID = "NWMapBookUIPrj.CreateMapBook"
  Case 1
    itemDef.ID = "NWMapBookUIPrj.PageIdentifier"
  Case 2
    itemDef.ID = "NWMapBookUIPrj.clsMapGridButton"
  Case 3
    itemDef.ID = "NWMapBookUIPrj.clsStripMapButton"
  Case 4
    itemDef.ID = "NWMapBookUIPrj.clsCreateMapBubble"
  Case 5
    itemDef.ID = "NWMapBookUIPrj.clsEditMapBubble"
  End Select
End Sub

Private Property Get IToolBarDef_ItemCount() As Long
  IToolBarDef_ItemCount = 6
End Property

Private Property Get IToolBarDef_name() As String
  IToolBarDef_name = "NWMapBookUIPrj.MapSeriesBar"
End Property
