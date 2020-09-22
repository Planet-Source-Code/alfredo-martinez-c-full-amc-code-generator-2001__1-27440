VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Test"
   ClientHeight    =   6105
   ClientLeft      =   405
   ClientTop       =   990
   ClientWidth     =   10065
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "frmTest"
   ScaleHeight     =   6105
   ScaleWidth      =   10065
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox txtResult 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2715
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2820
      Width           =   8955
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate"
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox txtSource 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   300
      Width           =   8955
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oSCGen As New CSCGen

Private Sub cmdGenerate_Click()
Dim oTable As New CTable

   With oTable
      .Name = "Orders"
      .Owner = "dbo"
      .TypeTable = "User"
      .CreateDate = "01/01/2001"
      With .Fields
         .Add
         .Item(1).Name = "Order_ID"
         .Item(1).DefinedSize = 4
         .Item(1).IsParamProc = False
         .Item(1).IsIdentity = True
         .Item(1).IsPK = True
         .Item(1).NumericScale = 4
         .Item(1).StringType = "int"
         .Item(1).TypeParam = NoParam
         .Add
         .Item(2).Name = "Date"
         .Item(2).DefinedSize = 8
         .Item(2).IsParamProc = False
         .Item(2).IsPK = False
         .Item(2).NumericScale = 4
         .Item(2).StringType = "smalldatetime"
         .Item(2).TypeParam = NoParam
         .Item(2).IsNull = False
         .Add
         .Item(3).Name = "Notes"
         .Item(3).DefinedSize = 80
         .Item(3).IsParamProc = False
         .Item(3).IsPK = False
         .Item(3).NumericScale = 16
         .Item(3).StringType = "varchar"
         .Item(3).TypeParam = NoParam
         .Item(3).IsNull = True
      End With
   End With
   
   Screen.MousePointer = 11
   
   With oSCGen
      .TabWidth = 3
      .Author = "Alfredo Martinez C."
      .ConnectionString = "Provider=SQLOLEDB;Server=Lupe;Uid=sa;Pwd=;Database=Northwind"
      .DataBase = "Nothwind"
      .Legal = "Copyrigth @2000, Todos los derechos reservados"
      .Owner = "AMC"
      .Server = "Lupe"
      .Table = oTable
      .Script = Trim(txtSource.Text)
      .Generate
      txtResult.Text = .Code
   End With
   txtResult.SelStart = 0
   Set oSCGen = Nothing
   Set oTable = Nothing
   Screen.MousePointer = 0

End Sub

Private Sub cmdStop_Click()
   Call oSCGen.CancelTrad
   Set oSCGen = Nothing
   
End Sub

Private Sub Form_Resize()
   txtSource.Move 0, (cmdGenerate.Height + 10), ScaleWidth, (ScaleHeight - cmdGenerate.Height + 50) / 2
   txtResult.Move 0, (txtSource.Top + 25 + txtSource.Height), txtSource.Width, (ScaleHeight - cmdGenerate.Height - 100) / 2
End Sub
