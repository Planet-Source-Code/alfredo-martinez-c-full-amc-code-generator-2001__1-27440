VERSION 5.00
Begin VB.Form frmSourceCode 
   Caption         =   "Source Code"
   ClientHeight    =   5955
   ClientLeft      =   2535
   ClientTop       =   2460
   ClientWidth     =   6690
   Icon            =   "frmSourceCode.frx":0000
   LinkTopic       =   "frmSourceCode"
   ScaleHeight     =   5955
   ScaleWidth      =   6690
   Begin GenObj.HighlightSintax txtEditor 
      Height          =   795
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1402
      Language        =   4
      KeywordColor    =   12582912
      OperatorColor   =   255
      DelimiterColor  =   10979685
      ForeColor       =   0
      FunctionColor   =   12583104
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmSourceCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub SetLanguaje(ByVal sLang As String)
   txtEditor.Language = sLang
End Sub

Private Sub Form_Resize()
   txtEditor.Move 10, 10, ScaleWidth - 20, ScaleHeight - 20
End Sub
