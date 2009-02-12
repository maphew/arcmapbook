VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmResources 
   Caption         =   "Form1"
   ClientHeight    =   3192
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "frmResources.frx":0000
   ScaleHeight     =   3192
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picEditBubble 
      Height          =   375
      Left            =   2520
      Picture         =   "frmResources.frx":05C2
      ScaleHeight     =   324
      ScaleWidth      =   684
      TabIndex        =   6
      Top             =   1560
      Width           =   735
   End
   Begin VB.PictureBox picSelectCursor 
      Height          =   735
      Left            =   3525
      Picture         =   "frmResources.frx":0B14
      ScaleHeight     =   684
      ScaleWidth      =   684
      TabIndex        =   5
      Top             =   1995
      Width           =   735
   End
   Begin VB.PictureBox picCreateBubble 
      Height          =   435
      Left            =   2565
      Picture         =   "frmResources.frx":0C66
      ScaleHeight     =   384
      ScaleWidth      =   732
      TabIndex        =   4
      Top             =   2010
      Width           =   780
   End
   Begin VB.PictureBox picCreateStripMap 
      Height          =   525
      Left            =   90
      Picture         =   "frmResources.frx":11B8
      ScaleHeight     =   480
      ScaleWidth      =   1200
      TabIndex        =   3
      Top             =   2100
      Width           =   1245
   End
   Begin VB.PictureBox picCreateGrid 
      Height          =   525
      Left            =   90
      Picture         =   "frmResources.frx":177A
      ScaleHeight     =   480
      ScaleWidth      =   1200
      TabIndex        =   2
      Top             =   1440
      Width           =   1245
   End
   Begin VB.PictureBox picBook 
      Height          =   525
      Left            =   90
      Picture         =   "frmResources.frx":1D3C
      ScaleHeight     =   480
      ScaleWidth      =   1200
      TabIndex        =   1
      Top             =   750
      Width           =   1245
   End
   Begin VB.PictureBox picIdentifier 
      Height          =   525
      Left            =   105
      Picture         =   "frmResources.frx":204E
      ScaleHeight     =   480
      ScaleWidth      =   1200
      TabIndex        =   0
      Top             =   90
      Width           =   1245
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   1740
      Top             =   105
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResources.frx":2610
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmResources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

