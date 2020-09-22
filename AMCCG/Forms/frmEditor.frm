VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmEditor 
   Caption         =   "Editor"
   ClientHeight    =   5325
   ClientLeft      =   960
   ClientTop       =   1365
   ClientWidth     =   9120
   Icon            =   "frmEditor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   9120
   WindowState     =   2  'Maximized
   Begin GenObj.Split spEditor 
      Height          =   2535
      Left            =   2220
      TabIndex        =   9
      Top             =   2100
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   4471
      HorizontalSplit =   -1  'True
      SplitPercent    =   100
      Begin GenObj.HighlightSintax txtResult 
         Height          =   555
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   979
         Language        =   5
         KeywordColor    =   8421376
         OperatorColor   =   8632256
         DelimiterColor  =   8421504
         ForeColor       =   0
         FunctionColor   =   12583104
         HighlightCode   =   0
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
      Begin GenObj.HighlightSintax txtEditor 
         Height          =   555
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   979
         Language        =   5
         KeywordColor    =   8421376
         OperatorColor   =   8632256
         DelimiterColor  =   8421504
         ForeColor       =   0
         FunctionColor   =   12583104
         HighlightCode   =   0
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
   Begin ComCtl3.CoolBar cbEditor 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   688
      BandCount       =   2
      _CBWidth        =   9120
      _CBHeight       =   390
      _Version        =   "6.7.8862"
      Child1          =   "tbGen"
      MinHeight1      =   330
      Width1          =   2490
      NewRow1         =   0   'False
      Child2          =   "tbEditor"
      MinHeight2      =   330
      Width2          =   3120
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar tbGen 
         Height          =   330
         Left            =   165
         TabIndex        =   13
         Top             =   30
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imlEdit"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Save"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Print"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Exit"
               ImageIndex      =   13
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbEditor 
         Height          =   330
         Left            =   2685
         TabIndex        =   12
         Top             =   30
         Width           =   6345
         _ExtentX        =   11192
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imlEdit"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Undo"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Redo"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Cut"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Copy"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Paste"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Delete"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Find"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Find Next"
               ImageIndex      =   8
            EndProperty
         EndProperty
         Enabled         =   0   'False
      End
   End
   Begin MSComctlLib.ImageList imlEdit 
      Left            =   8400
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":02AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":040A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":056A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":06CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":082A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":098A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":0AEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":0C4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":0DAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":3562
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":36C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":3822
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":3B3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditor.frx":3F92
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbEditor 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   11
      Top             =   5025
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   10
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Listen..."
            TextSave        =   "Listen..."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Enabled         =   0   'False
            TextSave        =   "SCRL"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "12:00 p.m."
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "14/08/01"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   7
            Enabled         =   0   'False
            TextSave        =   "KANA"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame framScript 
      Height          =   915
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   9435
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   1
         Left            =   960
         TabIndex        =   8
         Top             =   540
         Width           =   5295
      End
      Begin VB.ComboBox cmbStatus 
         Height          =   315
         Left            =   7080
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   180
         Width           =   1935
      End
      Begin VB.ComboBox cmbLibrary 
         Height          =   315
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   180
         Width           =   2895
      End
      Begin VB.TextBox txtFields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   2
         Top             =   180
         Width           =   1575
      End
      Begin VB.Label lblFields 
         Caption         =   "&Name:"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   7
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblFields 
         Caption         =   "Sta&tus:"
         Height          =   195
         Index           =   2
         Left            =   6480
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblFields 
         Caption         =   "&Library:"
         Height          =   195
         Index           =   1
         Left            =   2700
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblFields 
         Caption         =   "&ID Script:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Menu mnuFilE 
      Caption         =   "&File"
      Begin VB.Menu mnuFilSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu filS2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFilPrintSetup 
         Caption         =   "Print &Setup"
      End
      Begin VB.Menu filS3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuFilUndo 
         Caption         =   "&Undo"
         Enabled         =   0   'False
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuFilRedo 
         Caption         =   "&Redo"
         Enabled         =   0   'False
         Shortcut        =   ^R
      End
      Begin VB.Menu ediS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdiCut 
         Caption         =   "Cu&t"
         Enabled         =   0   'False
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEdiCopy 
         Caption         =   "&Copy"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEdiPaste 
         Caption         =   "&Paste"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEdiDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu ediS2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTooFind 
         Caption         =   "&Find..."
         Enabled         =   0   'False
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuTooFindNext 
         Caption         =   "Find &Next"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu ediS3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdi 
         Caption         =   "&Select All      "
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuToolBars 
         Caption         =   "&Toolbars"
         Begin VB.Menu mnuViewTBStandar 
            Caption         =   "Standar"
            Checked         =   -1  'True
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuViewTBEdit 
            Caption         =   "Edit"
            Checked         =   -1  'True
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu VieS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOutWin 
         Caption         =   "Output Window"
      End
      Begin VB.Menu VieS2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVieSB 
         Caption         =   "Statusbar"
         Checked         =   -1  'True
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuScript 
      Caption         =   "&Script"
      Begin VB.Menu mnuScriptStart 
         Caption         =   "&Execute"
         Enabled         =   0   'False
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuScriptEnd 
         Caption         =   "&Cancel Executing"
         Enabled         =   0   'False
         Shortcut        =   ^{F12}
      End
      Begin VB.Menu ScriptS1 
         Caption         =   "-"
      End
      Begin VB.Menu ScriptS2 
         Caption         =   "&Options..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHlpContents 
         Caption         =   "&Contents..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuHlpIndex 
         Caption         =   "&Index..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuHlpSearch 
         Caption         =   "&Search..."
         Enabled         =   0   'False
      End
      Begin VB.Menu HlpS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHlpAbout 
         Caption         =   "About AMC-CodeGenerator..."
      End
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ItemLib As Integer
Private ItemScript As Long
Private Accion As String


Private Sub Form_Load()
   
   oMPG.CentrarForma Me
   
   With spEditor
      Set .Control1 = txtEditor
      Set .Control2 = txtResult
   End With
   
   txtEditor.Language = [AMC Script Languaje]
   
   With sbEditor
      .Panels(1).Width = 3000
      .Panels(2).Width = 600
      .Panels(3).Width = 550
      .Panels(4).Width = 400
      .Panels(5).Width = 600
      .Panels(6).Width = 1000
      .Panels(7).Width = 1000
      .Panels(8).Width = 600
      .Panels(9).Visible = False
   End With
   
   With cmbStatus
      .AddItem "A) Active"
      .AddItem "C) Cancel"
      .ListIndex = 0
   End With
   
End Sub


Private Sub Form_Resize()
Dim lHeightProcs As Long
On Error Resume Next
   framScript.Width = Me.Width - 100
   spEditor.Move 10, (cbEditor.Height + 50 + framScript.Height), ScaleWidth - 10, ScaleHeight - (sbEditor.Height + cbEditor.Height + 75 + framScript.Height)
End Sub

Private Sub mnuEdi_Click()
'   txtEditor.SelStart = 0
   txtEditor.SelText = Len(txtEditor.Text)
End Sub

Private Sub mnuFilExit_Click()
   Unload Me
End Sub



Private Sub mnuFilSave_Click()
Dim RsAMC As New ADODB.Recordset
Dim nLastScript As Long

   If Trim(txtFields(1)) = "" Then
      MsgBox "Information it's requiered", vbInformation, App.Title
      txtFields(1).SetFocus
   End If
   
   If Trim(txtEditor.Text) = "" Then
      MsgBox "Information it's requiered", vbInformation, App.Title
      txtEditor.SetFocus
   End If
   
   With RsAMC
      .CursorLocation = adUseClient
      
      If Accion = "A" Then
         sSQLcmd = "SELECT MAX(nScriptID) FROM Scripts "
         .Open sSQLcmd, sysDB, adOpenForwardOnly, adLockReadOnly
         If .RecordCount > 0 Then
            If IsNull(.Fields(0)) = True Then
               nLastScript = 1
            Else
               nLastScript = .Fields(0)
               nLastScript = nLastScript + 1
            End If
         Else
            nLastScript = 1
         End If
         .Close
         
         sSQLcmd = "SELECT * FROM Scripts WHERE nScriptID IS NULL "
         .Open sSQLcmd, sysDB, adOpenStatic, adLockOptimistic
         .AddNew
         .Fields("nScriptID") = nLastScript
      Else
         sSQLcmd = "SELECT * FROM Scripts WHERE nScriptID = " & txtFields(0)
         .Open sSQLcmd, sysDB, adOpenStatic, adLockOptimistic
      End If
      
      .Fields("nLibraryID") = ItemLib
      .Fields("cName") = Trim(txtFields(1))
      .Fields("cCodeScript") = Trim(txtEditor.Text)
      .Fields("cStatus") = Mid(cmbStatus.Text, 1, 1)
      .Update
      
      txtFields(0) = .Fields("nScriptID")
      
      If Accion = "A" Then
         frmLibs.lvLibs.ListItems.Add , , .Fields(0), Icon:=9, SmallIcon:=9
         ItemScript = frmLibs.lvLibs.ListItems(frmLibs.lvLibs.ListItems.Count).Index
         frmLibs.lvLibs.ListItems.Item(ItemScript).SubItems(1) = .Fields(1)
         ItemLib = .Fields(1)
         
         frmLibs.lvLibs.ListItems.Item(ItemScript).SubItems(2) = .Fields(2)
         frmLibs.lvLibs.ListItems.Item(ItemScript).SubItems(3) = .Fields(4)
      Else
         frmLibs.lvLibs.ListItems(ItemScript).Text = .Fields(0)
         frmLibs.lvLibs.ListItems.Item(ItemScript).SubItems(1) = .Fields(1)
         frmLibs.lvLibs.ListItems.Item(ItemScript).SubItems(2) = .Fields(2)
         frmLibs.lvLibs.ListItems.Item(ItemScript).SubItems(3) = .Fields(4)
      End If
      Accion = "M"
      .Close
         
   End With
   
   Set RsAMC = Nothing
   
End Sub

Private Sub mnuHlpAbout_Click()
   Load frmAbout
   frmAbout.Show 1
End Sub

Private Sub mnuViewOutWin_Click()
   If mnuViewOutWin.Checked = True Then
      mnuViewOutWin.Checked = False
      spEditor.SplitPercent = 100
   Else
      mnuViewOutWin.Checked = True
      spEditor.SplitPercent = 65
   End If
End Sub

Private Sub tbGen_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Index
      Case 1
         Call mnuFilSave_Click
      Case 5
         Call mnuFilExit_Click
   End Select
End Sub


Public Sub NewScript(ByVal nItemLib As Integer, ByVal cLib As String)
   Me.Caption = "New Script on " & cLib
   ItemLib = nItemLib
   cmbLibrary.AddItem cLib
   cmbLibrary.ItemData(0) = nItemLib
   cmbLibrary.ListIndex = 0
   Accion = "A"
End Sub


Public Sub ModifyScript(ByVal nItemLV As Integer, ByVal cLib As String)
Dim RsAMC As New ADODB.Recordset
Dim nScriptID As Long
   
   cmbLibrary.AddItem cLib
   cmbLibrary.ListIndex = 0
   ItemScript = nItemLV
   Accion = "M"
   nScriptID = frmLibs.lvLibs.ListItems(nItemLV).Text
   With RsAMC
      sSQLcmd = "SELECT * FROM Scripts WHERE nScriptID = " & nScriptID
      .CursorLocation = adUseClient
      .Open sSQLcmd, sysDB, adOpenForwardOnly, adLockReadOnly
      If .RecordCount > 0 Then
         txtFields(0) = .Fields(0)
         ItemLib = .Fields(1)
         txtFields(1) = Trim(.Fields(2))
         txtEditor.Text = Trim(.Fields(3))
         If .Fields(4) = "A" Then
            cmbStatus.ListIndex = 0
         Else
            cmbStatus.ListIndex = 1
         End If
      End If
      .Close
   End With
   Set RsAMC = Nothing

End Sub



Private Sub txtFields_Change(Index As Integer)
   Me.Tag = "Change"
End Sub
