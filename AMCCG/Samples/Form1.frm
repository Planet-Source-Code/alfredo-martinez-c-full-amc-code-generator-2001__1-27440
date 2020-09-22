VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   1950
   ClientTop       =   1545
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   6585
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert Record"
      Height          =   435
      Left            =   240
      TabIndex        =   5
      Top             =   1380
      Width           =   1635
   End
   Begin VB.TextBox txtFields 
      Height          =   315
      Index           =   1
      Left            =   1500
      TabIndex        =   4
      Top             =   540
      Width           =   2295
   End
   Begin VB.TextBox txtFields 
      Height          =   315
      Index           =   0
      Left            =   1500
      TabIndex        =   2
      Top             =   180
      Width           =   2295
   End
   Begin VB.CommandButton cmdGet 
      Caption         =   "Get Record"
      Height          =   495
      Left            =   4200
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "ID:"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As New ADODB.Connection
Dim oReg As New CReg

Private Sub cmdGet_Click()
   If Trim(txtFields(0)) = "" Then Exit Sub
   With oReg
      .RegionID = txtFields(0)
      .GetData
      If .RegionDescription <> "" Then
         txtFields(0) = .RegionID
         txtFields(1) = Trim(.RegionDescription)
      Else
         MsgBox "Record Not Exist", vbInformation, App.Title
         txtFields(0) = ""
         txtFields(1) = ""
      End If
   End With
End Sub

Private Sub cmdInsert_Click()
   If Trim(txtFields(0)) = "" Then Exit Sub
   With oReg
      .RegionID = txtFields(0)
      .RegionDescription = txtFields(1)
      .Insert
   End With
   
End Sub

Private Sub Form_Load()
   With db
      .Provider = "SQLOLEDB"
      .ConnectionString = "Server=Lupe;Uid=sa;Pwd=;Database=Northwind"
      .Open
   End With
   Set oReg.adoConn = db
End Sub
