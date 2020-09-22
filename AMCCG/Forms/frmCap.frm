VERSION 5.00
Begin VB.Form frmCap 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form Capture"
   ClientHeight    =   3090
   ClientLeft      =   1590
   ClientTop       =   2625
   ClientWidth     =   7485
   Icon            =   "frmCap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   Begin VB.Frame framLib 
      Height          =   2895
      Left            =   120
      TabIndex        =   13
      Top             =   3540
      Visible         =   0   'False
      Width           =   5955
      Begin VB.ComboBox cmbStatusLib 
         Height          =   315
         Left            =   2100
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   2400
         Width           =   1515
      End
      Begin VB.TextBox txtFields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   2100
         TabIndex        =   25
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtFields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   2100
         TabIndex        =   23
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   6
         Left            =   2100
         TabIndex        =   21
         Top             =   1320
         Width           =   3615
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   5
         Left            =   2100
         TabIndex        =   19
         Top             =   960
         Width           =   3615
      End
      Begin VB.ComboBox cmbLanguaje 
         Height          =   315
         Left            =   2100
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   600
         Width           =   3615
      End
      Begin VB.TextBox txtFields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2100
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblFields 
         Caption         =   "ID"
         Height          =   195
         Index           =   12
         Left            =   240
         TabIndex        =   26
         Top             =   2460
         Width           =   1755
      End
      Begin VB.Label lblFields 
         Caption         =   "ID"
         Height          =   195
         Index           =   11
         Left            =   240
         TabIndex        =   24
         Top             =   2100
         Width           =   1755
      End
      Begin VB.Label lblFields 
         Caption         =   "ID"
         Height          =   195
         Index           =   10
         Left            =   240
         TabIndex        =   22
         Top             =   1740
         Width           =   1755
      End
      Begin VB.Label lblFields 
         Caption         =   "ID"
         Height          =   195
         Index           =   9
         Left            =   240
         TabIndex        =   20
         Top             =   1380
         Width           =   1755
      End
      Begin VB.Label lblFields 
         Caption         =   "ID"
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   18
         Top             =   1020
         Width           =   1755
      End
      Begin VB.Label lblFields 
         Caption         =   "ID"
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   16
         Top             =   660
         Width           =   1755
      End
      Begin VB.Label lblFields 
         Caption         =   "ID"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   14
         Top             =   300
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6180
      TabIndex        =   29
      Top             =   660
      Width           =   1215
   End
   Begin VB.CommandButton cmdAcept 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6180
      TabIndex        =   28
      Top             =   180
      Width           =   1215
   End
   Begin VB.Frame framCap 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   5955
      Begin VB.ComboBox cmbStatusLang 
         Height          =   315
         Left            =   2100
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2160
         Width           =   1515
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   3
         Left            =   2100
         TabIndex        =   10
         Top             =   1800
         Width           =   3615
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   2
         Left            =   2100
         TabIndex        =   8
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   1
         Left            =   2100
         TabIndex        =   6
         Top             =   1080
         Width           =   3615
      End
      Begin VB.ComboBox cmbTypeLang 
         Height          =   315
         Left            =   2100
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   3615
      End
      Begin VB.TextBox txtFields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2100
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblFields 
         Caption         =   "ID"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   11
         Top             =   2220
         Width           =   1755
      End
      Begin VB.Label lblFields 
         Caption         =   "ID"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   9
         Top             =   1860
         Width           =   1755
      End
      Begin VB.Label lblFields 
         Caption         =   "ID"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   1500
         Width           =   1755
      End
      Begin VB.Label lblFields 
         Caption         =   "ID"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   1140
         Width           =   1755
      End
      Begin VB.Label lblFields 
         Caption         =   "ID"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   780
         Width           =   1755
      End
      Begin VB.Label lblFields 
         Caption         =   "ID"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   420
         Width           =   1755
      End
   End
End
Attribute VB_Name = "frmCap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private iCount As Integer
Private Accion As String
Private sOption As String
Private iItemLV As Integer
Private iItemTV As Integer



Private Sub cmdAcept_Click()
   If Accion = "A" Then
      If LCase(sOption) = "lang" Then
         If ValidData(1) = False Then Exit Sub
         If AffectRecordLang(1) = True Then
            Unload Me
         End If
      Else
         If ValidData(2) = False Then Exit Sub
         If AffectRecordLib(1) = True Then
            Unload Me
         End If
      End If
   Else
      If LCase(sOption) = "lang" Then
         If ValidData(1) = False Then Exit Sub
         If AffectRecordLang(2) = True Then
            Unload Me
         End If
      Else
         If ValidData(2) = False Then Exit Sub
         If AffectRecordLib(2) = True Then
            Unload Me
         End If
      End If
   End If
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   oMPG.CentrarForma Me
   oMPG.Explosion Me.hwnd, 200, Negro
   With txtFields
      For iCount = 0 To .Count - 1
         FlatBorder .Item(iCount).hwnd
      Next iCount
   End With
   FlatBorder cmdAcept.hwnd
   FlatBorder cmdCancel.hwnd
End Sub


Private Sub LoadTypeLang()
Dim RsAMC As New ADODB.Recordset
   With RsAMC
      .CursorLocation = adUseClient
      sSQLcmd = "SELECT * FROM TypeLang Order By cDescript "
      .Open sSQLcmd, sysDB, adOpenForwardOnly, adLockReadOnly
      If .RecordCount > 0 Then
         For iCount = 1 To .RecordCount
            cmbTypeLang.AddItem Trim(.Fields("cDescript"))
            cmbTypeLang.ItemData(iCount - 1) = .Fields("nTypeLangID")
            .MoveNext
         Next iCount
         cmbTypeLang.ListIndex = 0
      End If
      .Close
   End With
   Set RsAMC = Nothing
   
   lblFields(0).Caption = "ID Languaje"
   lblFields(1).Caption = "Type Languaje"
   lblFields(2).Caption = "Name"
   lblFields(3).Caption = "Prefix"
   lblFields(4).Caption = "Description"
   lblFields(5).Caption = "Status"
   
End Sub


Public Sub NewLanguaje()
   Accion = "A"
   sOption = "lang"
   Me.Caption = "New Languaje"
   Call LoadTypeLang
   With cmbStatusLang
      .AddItem "A) Active"
      .AddItem "C) Cancel"
      .ListIndex = 0
   End With
End Sub



Public Sub ModifyLanguaje(ByVal nItemLV As Integer, ByVal nItemTV As Integer)
   Accion = "M"
   sOption = "lang"
   Me.Caption = "Modify Languaje"
   iItemLV = nItemLV
   iItemTV = nItemTV
   
   Call LoadTypeLang
   With frmLibs.lvLibs.ListItems(nItemLV)
      txtFields(0) = .Text
      For iCount = 0 To cmbTypeLang.ListCount - 1
         If cmbTypeLang.ItemData(iCount) = .SubItems(1) Then
            cmbTypeLang.ListIndex = iCount: Exit For
         End If
      Next iCount
      
      txtFields(1) = .SubItems(2)
      txtFields(2) = .SubItems(3)
      txtFields(3) = .SubItems(4)
      
      cmbTypeLang.Locked = True
      
      With cmbStatusLang
         .AddItem "A) Active"
         .AddItem "C) Cancel"
         .ListIndex = 0
      End With
      
      If UCase(.SubItems(5)) = "A" Then
         cmbStatusLang.ListIndex = 0
      Else
         cmbStatusLang.ListIndex = 1
      End If
   End With
End Sub



Private Sub LoadLanguajes()
Dim RsAMC As New ADODB.Recordset
   With RsAMC
      .CursorLocation = adUseClient
      sSQLcmd = "SELECT nLanguajeID, cName FROM Languaje Order By cName "
      .Open sSQLcmd, sysDB, adOpenForwardOnly, adLockReadOnly
      If .RecordCount > 0 Then
         For iCount = 1 To .RecordCount
            cmbLanguaje.AddItem Trim(.Fields("cName"))
            cmbLanguaje.ItemData(iCount - 1) = .Fields("nLanguajeID")
            .MoveNext
         Next iCount
         cmbLanguaje.ListIndex = 0
      End If
      .Close
   End With
   Set RsAMC = Nothing
   
   lblFields(6).Caption = "ID Library"
   lblFields(7).Caption = "Languaje"
   lblFields(8).Caption = "Name"
   lblFields(9).Caption = "Author"
   lblFields(10).Caption = "Date Create"
   lblFields(11).Caption = "DateModify"
   lblFields(12).Caption = "Status"
   
   framCap.Visible = False
   framLib.Move framCap.Left, framCap.Top, framCap.Width, framCap.Height
   framLib.Visible = True

End Sub



Public Sub NewLibrary()
   Accion = "A"
   sOption = "lib"
   Me.Caption = "New Library"
   txtFields(7).Text = Format(CDate(Now), "mm/dd/yyyy")
   txtFields(8).Text = Format(CDate(Now), "mm/dd/yyyy")
   Call LoadLanguajes
   With cmbStatusLib
      .AddItem "A) Active"
      .AddItem "C) Cancel"
      .ListIndex = 0
   End With
End Sub


Public Sub ModifyLibrary(ByVal nItemLV As Integer, ByVal nItemTV As Integer)
   Accion = "M"
   sOption = "lib"
   Me.Caption = "Modify Library"
   iItemLV = nItemLV
   iItemTV = nItemTV
   
   Call LoadLanguajes
   With frmLibs.lvLibs.ListItems(nItemLV)
      txtFields(4) = .Text
      For iCount = 0 To cmbLanguaje.ListCount - 1
         If cmbLanguaje.ItemData(iCount) = .SubItems(1) Then
            cmbLanguaje.ListIndex = iCount: Exit For
         End If
      Next iCount
      
      txtFields(5) = .SubItems(2)
      txtFields(6) = .SubItems(3)
      txtFields(7) = .SubItems(4)
      txtFields(8) = .SubItems(5)
      
      cmbLanguaje.Locked = True
      
      With cmbStatusLib
         .AddItem "A) Active"
         .AddItem "C) Cancel"
         .ListIndex = 0
      End With
      
      If UCase(.SubItems(6)) = "A" Then
         cmbStatusLib.ListIndex = 0
      Else
         cmbStatusLib.ListIndex = 1
      End If
   End With
End Sub



Private Function ValidData(ByVal nMode As Integer) As Boolean
   If nMode = 1 Then
      If oMPG.CamposRequeridos(txtFields(1), txtFields(2)) = False Then
         ValidData = False: Exit Function
      Else
         If oMPG.CombosRequeridos(cmbTypeLang) = False Then
            ValidData = False: Exit Function
         Else
            ValidData = True
         End If
      End If
   ElseIf nMode = 2 Then
      If oMPG.CamposRequeridos(txtFields(5), txtFields(6), txtFields(7)) = False Then
         ValidData = False: Exit Function
      Else
         If oMPG.CombosRequeridos(cmbLanguaje) = False Then
            ValidData = False: Exit Function
         Else
            ValidData = True
         End If
      End If
   End If
End Function


'Function for Insert New Languaje Programming
Private Function AffectRecordLang(ByVal nMode As Single) As Boolean
Dim RsAMC As New ADODB.Recordset
Dim iRecID As Integer
On Error GoTo ErrorAffectRecordLang
   
   With RsAMC
      .CursorLocation = adUseClient
      If nMode = 1 Then
         sSQLcmd = "SELECT MAX(nLanguajeID) FROM Languaje "
         .Open sSQLcmd, sysDB, adOpenForwardOnly, adLockReadOnly
         If .RecordCount <= 0 Then
            iRecID = 1
         Else
            If IsNull(.Fields(0)) Then
               iRecID = 1
            Else
               iRecID = .Fields(0) + 1
            End If
         End If
         .Close
         sSQLcmd = "SELECT * FROM Languaje WHERE nLanguajeID Is Null"
         .Open sSQLcmd, sysDB, adOpenKeyset, adLockOptimistic
         .AddNew
         .Fields("nLanguajeID") = iRecID
      Else
         sSQLcmd = "SELECT * FROM Languaje WHERE nLanguajeID = " & txtFields(0)
         .Open sSQLcmd, sysDB, adOpenKeyset, adLockOptimistic
      End If
      
      .Fields("nTypeLanguajeID") = cmbTypeLang.ItemData(cmbTypeLang.ListIndex)
      .Fields("cName") = Trim(txtFields(1))
      .Fields("cPrefix") = Trim(txtFields(2))
      .Fields("cComments") = Trim(txtFields(3))
      .Fields("cStatus") = Mid(cmbStatusLang.Text, 1, 1)
      .Update
      If nMode = 1 Then
         Dim xNode As Node
         Set xNode = frmLibs.tvLibs.Nodes.Add("all", tvwChild, "Lan" & iRecID, Trim(txtFields(1)), Image:=3, SelectedImage:=3)
         xNode.Tag = "lang"
      Else
         With frmLibs.lvLibs.ListItems(iItemLV)
            .Text = txtFields(0)
            .SubItems(1) = cmbTypeLang.ItemData(cmbTypeLang.ListIndex)
            .SubItems(2) = txtFields(1)
            .SubItems(3) = txtFields(2)
            .SubItems(4) = txtFields(3)
            .SubItems(5) = Mid(cmbStatusLang.Text, 1, 1)
         End With
         If iItemTV > 0 Then
            frmLibs.tvLibs.Nodes(iItemTV).Text = Trim(txtFields(1))
         End If
      End If
      .Close
   End With
   
   AffectRecordLang = True
   
Exit Function
ErrorAffectRecordLang:
   MsgBox "Error: [" & Err & "] " & Error, vbCritical, App.Title
   AffectRecordLang = False
End Function


'Function for Insert New Script Library
Private Function AffectRecordLib(ByVal nMode As Integer) As Boolean
Dim RsAMC As New ADODB.Recordset
Dim iRecID As Integer
On Error GoTo ErrorAffectRecordLib
   
   With RsAMC
      .CursorLocation = adUseClient
      If nMode = 1 Then
         sSQLcmd = "SELECT MAX(nLibraryID) FROM Libraries "
         .Open sSQLcmd, sysDB, adOpenForwardOnly, adLockReadOnly
         If .RecordCount <= 0 Then
            iRecID = 1
         Else
            If IsNull(.Fields(0)) Then
               iRecID = 1
            Else
               iRecID = .Fields(0) + 1
            End If
         End If
         .Close
      
         sSQLcmd = "SELECT * FROM Libraries where nLibraryID Is Null"
         .Open sSQLcmd, sysDB, adOpenKeyset, adLockOptimistic
         .AddNew
         .Fields("nLibraryID") = iRecID
      Else
         sSQLcmd = "SELECT * FROM Libraries where nLibraryID = " & txtFields(4)
         .Open sSQLcmd, sysDB, adOpenKeyset, adLockOptimistic
      End If
      .Fields("nLanguajeID") = cmbLanguaje.ItemData(cmbLanguaje.ListIndex)
      .Fields("cName") = Trim(txtFields(5))
      .Fields("cAuthor") = Trim(txtFields(6))
      .Fields("dCreation") = Trim(txtFields(7))
      If nMode = 1 Then
         .Fields("dModify") = Trim(txtFields(8))
      Else
         .Fields("dModify") = Format(CDate(Now), "mm/dd/yyyy")
      End If
      
      .Fields("cStatus") = Mid(cmbStatusLang.Text, 1, 1)
      .Update
      If nMode = 1 Then
         Dim xNode As Node
         Set xNode = frmLibs.tvLibs.Nodes.Add("Lan" & .Fields("nLanguajeID"), tvwChild, "Lib" & .Fields("nLibraryID"), Trim(txtFields(5)), Image:=5, SelectedImage:=5)
         xNode.Tag = "lib"
      Else
         With frmLibs.lvLibs.ListItems(iItemLV)
            .Text = txtFields(4)
            .SubItems(1) = cmbLanguaje.ItemData(cmbLanguaje.ListIndex)
            .SubItems(2) = txtFields(5)
            .SubItems(3) = txtFields(6)
            .SubItems(4) = txtFields(7)
            .SubItems(5) = txtFields(8)
            .SubItems(6) = Mid(cmbStatusLib.Text, 1, 1)
         End With
         If iItemTV > 0 Then
            frmLibs.tvLibs.Nodes(iItemTV).Text = Trim(txtFields(5))
         End If
      End If
   End With
   
   AffectRecordLib = True
   
Exit Function
ErrorAffectRecordLib:
   MsgBox "Error: [" & Err & "] " & Error, vbCritical, App.Title
   AffectRecordLib = False
End Function


Private Sub txtFields_GotFocus(Index As Integer)
   With txtFields(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub


