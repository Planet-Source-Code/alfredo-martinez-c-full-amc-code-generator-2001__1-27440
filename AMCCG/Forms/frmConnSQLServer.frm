VERSION 5.00
Begin VB.Form frmConnSQLServer 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Connect to SQL Server"
   ClientHeight    =   3600
   ClientLeft      =   4755
   ClientTop       =   1995
   ClientWidth     =   4575
   Icon            =   "frmConnSQLServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCap 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   2520
      Width           =   2475
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3240
      TabIndex        =   11
      Top             =   3180
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Height          =   315
      Left            =   1920
      TabIndex        =   10
      Top             =   3180
      Width           =   1215
   End
   Begin VB.OptionButton optCnnOptions 
      Appearance      =   0  'Flat
      Caption         =   "Use S&QL Server authentication"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   540
      TabIndex        =   3
      Top             =   1260
      Value           =   -1  'True
      Width           =   2715
   End
   Begin VB.TextBox txtCap 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   2160
      Width           =   2475
   End
   Begin VB.TextBox txtCap 
      Height          =   315
      Index           =   1
      Left            =   1920
      TabIndex        =   5
      Top             =   1620
      Width           =   2475
   End
   Begin VB.TextBox txtCap 
      Height          =   315
      Index           =   0
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   2475
   End
   Begin VB.Label lblTitles 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      Height          =   195
      Index           =   4
      Left            =   540
      TabIndex        =   8
      Top             =   2580
      Width           =   1035
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   4440
      Y1              =   3030
      Y2              =   3030
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000C&
      X1              =   120
      X2              =   4440
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   4440
      Y1              =   870
      Y2              =   870
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      X1              =   120
      X2              =   4440
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblTitles 
      BackStyle       =   0  'Transparent
      Caption         =   "Connection Information:"
      Height          =   195
      Index           =   3
      Left            =   180
      TabIndex        =   2
      Top             =   960
      Width           =   1875
   End
   Begin VB.Label lblTitles 
      BackStyle       =   0  'Transparent
      Caption         =   "&Login Name:"
      Height          =   195
      Index           =   2
      Left            =   540
      TabIndex        =   6
      Top             =   2220
      Width           =   1035
   End
   Begin VB.Label lblTitles 
      BackStyle       =   0  'Transparent
      Caption         =   "&Database Name:"
      Height          =   195
      Index           =   1
      Left            =   540
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Image imgCnn 
      Height          =   480
      Left            =   120
      Picture         =   "frmConnSQLServer.frx":000C
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lblTitles 
      BackStyle       =   0  'Transparent
      Caption         =   "&SQL Server:"
      Height          =   195
      Index           =   0
      Left            =   780
      TabIndex        =   0
      Top             =   300
      Width           =   1035
   End
End
Attribute VB_Name = "frmConnSQLServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
Dim db As New ADODB.Connection
Dim sConnectionString As String, sKey As String
Dim bNodeExist As Boolean, lNodeCount As Long
On Error GoTo ErrorcmdOK_Click

   DoEvents
   sConnectionString = "Provider=SQLOLEDB;" & "SERVER=" & txtCap(0) & ";UID=" & txtCap(2) & ";PWD=" & txtCap(3) & ";DATABASE=" & txtCap(1)
   sKey = UCase(txtCap(0) & " - " & txtCap(1))
   oCnnTmp.DisplayInfo "Open connection", "Open Microsoft SQL Server Database [" & sKey & "] "
   bNodeExist = False
   With frmMain.tvDBs.Nodes
      DoEvents
      For lNodeCount = 1 To .Count
         If UCase(.Item(lNodeCount).Text) = UCase(sKey) Then
            bNodeExist = True
            oCnnTmp.DisplayInfo "Open connection", "Connection with Microsoft SQL Server Database already exists..."
            Exit For
         End If
      Next lNodeCount
   End With
   
   If bNodeExist = False Then
      DoEvents
      oCnnTmp.DisplayInfo "Open connection", "Establish connection..."
      db.Open sConnectionString
      oDBs.Add db, sKey
      oCnnTmp.DisplayInfo "Open connection", "Connection completed successfully..."
      Call oCnnTmp.DisplayConnection(sKey, frmMain.tvDBs, sKey, [SQL Server])
      oCnnTmp.DisplayInfo "Open connection", " - (0) error(s),  (0) warning(s) "
   Else
      MsgBox "Connection already exists...", vbExclamation, App.Title
   End If
   
   Screen.MousePointer = vbDefault
   Unload Me
Exit Sub
ErrorcmdOK_Click:
   Screen.MousePointer = vbDefault
   'oMPG.DisplayError 0, Err, Error, "Connection..."
   oCnnTmp.DisplayInfo "Open connection", "Error [" & Err & "] " & Error
   oCnnTmp.DisplayInfo "Open connection", " - (1) error(s),  (1) warning(s) "
   Set db = Nothing
End Sub

Private Sub Form_Activate()
   'MDIGen.Tag = "frmfrmConnSQLServer"
End Sub

Private Sub Form_Load()
   oMPG.CentrarForma Me
   FlatBorder txtCap(0).hwnd
   FlatBorder txtCap(1).hwnd
   FlatBorder txtCap(2).hwnd
   FlatBorder txtCap(3).hwnd
   
   FlatBorder cmdCancel.hwnd
   FlatBorder cmdOK.hwnd
   
End Sub

Private Sub txtCap_GotFocus(Index As Integer)
   With txtCap(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub
