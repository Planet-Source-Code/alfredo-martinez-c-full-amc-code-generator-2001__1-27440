VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmLibs 
   Caption         =   "Libraries Desing"
   ClientHeight    =   4875
   ClientLeft      =   1215
   ClientTop       =   3375
   ClientWidth     =   8670
   Icon            =   "frmLibs.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   8670
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imlLibs 
      Left            =   7800
      Top             =   3900
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibs.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibs.frx":0170
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibs.frx":02D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibs.frx":0430
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibs.frx":074C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibs.frx":0A68
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibs.frx":0BC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibs.frx":0D20
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibs.frx":34D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibs.frx":3FA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibs.frx":40FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibs.frx":4418
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibs.frx":4574
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibs.frx":46D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibs.frx":482C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibs.frx":4988
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibs.frx":4AE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibs.frx":4C40
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibs.frx":4D9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibs.frx":50B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibs.frx":5214
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin GenObj.Split spGen 
      Height          =   2835
      Left            =   900
      TabIndex        =   0
      Top             =   1080
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   5001
      SplitPercent    =   30
      Begin GenObj.Split spProps 
         Height          =   2655
         Left            =   1860
         TabIndex        =   6
         Top             =   60
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   4683
         HorizontalSplit =   -1  'True
         Begin VB.TextBox txtCodeScript 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   975
            Left            =   480
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   8
            Top             =   1500
            Width           =   2775
         End
         Begin MSComctlLib.ListView lvLibs 
            Height          =   1275
            Left            =   480
            TabIndex        =   7
            Top             =   60
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   2249
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            Icons           =   "imlLibs"
            SmallIcons      =   "imlLibs"
            ColHdrIcons     =   "imlLibs"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin MSComctlLib.TreeView tvLibs 
         Height          =   2655
         Left            =   60
         TabIndex        =   5
         Top             =   60
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   4683
         _Version        =   393217
         Indentation     =   706
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "imlLibs"
         Appearance      =   1
      End
   End
   Begin MSComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   4575
      Width           =   8670
      _ExtentX        =   15293
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
            TextSave        =   "8:10 AM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "04/07/01"
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
   Begin ComCtl3.CoolBar cbLibs 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8670
      _ExtentX        =   15293
      _ExtentY        =   688
      BandCount       =   2
      _CBWidth        =   8670
      _CBHeight       =   390
      _Version        =   "6.7.8988"
      Child1          =   "tbLibs"
      MinHeight1      =   330
      Width1          =   2805
      NewRow1         =   0   'False
      Child2          =   "tbEdit"
      MinHeight2      =   330
      Width2          =   1905
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar tbEdit 
         Height          =   330
         Left            =   3000
         TabIndex        =   4
         Top             =   30
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imlLibs"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "Cut"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "Copy"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "Paste"
               ImageIndex      =   14
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "Find"
               ImageIndex      =   16
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "Find Next"
               ImageIndex      =   17
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbLibs 
         Height          =   330
         Left            =   165
         TabIndex        =   3
         Top             =   30
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imlLibs"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "New"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Remove"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modify"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.ToolTipText     =   "Print"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Exit"
               ImageIndex      =   11
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFilNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu FilS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilRemove 
         Caption         =   "Re&move"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuFilModify 
         Caption         =   "&Modify"
         Shortcut        =   ^M
      End
      Begin VB.Menu FilS2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFilPrnCfg 
         Caption         =   "Print &Setup"
      End
      Begin VB.Menu FilS3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
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
      Begin VB.Menu EdiS2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdiSelAll 
         Caption         =   "Select &All"
         Enabled         =   0   'False
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuVieTB 
         Caption         =   "&Toolbars"
         Begin VB.Menu mnuVieTBStandar 
            Caption         =   "Standar"
            Checked         =   -1  'True
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuVieTBEdicion 
            Caption         =   "Edición"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu ViewS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVieDBExp 
         Caption         =   "Database Library Explorer"
         Checked         =   -1  'True
         Shortcut        =   ^R
      End
      Begin VB.Menu ViewS2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Statusbar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHlpContents 
         Caption         =   "&Contents..."
      End
      Begin VB.Menu mnuHlpIndex 
         Caption         =   "&Index..."
      End
      Begin VB.Menu mnuHlpSearch 
         Caption         =   "&Search..."
      End
      Begin VB.Menu HlpS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About AMC-CodeGenerator"
      End
   End
End
Attribute VB_Name = "frmLibs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Load()
   Call ConfigStatusBar
   Call ConfigControls
   mnuEdiCopy.Enabled = False
   Call RefreshTreeData
End Sub


Public Sub RefreshTreeData()
Dim xNode As Node
   tvLibs.Nodes.Clear
   lvLibs.ListItems.Clear
   mnuEdiCopy.Enabled = False
   Set xNode = tvLibs.Nodes.Add(, , "all", "Libraries", Image:=19, SelectedImage:=19)
   xNode.Expanded = True
   xNode.Tag = "all"
   Call ChargeLanguajes(1)
   Call tvLibs_NodeClick(xNode)
End Sub




Private Sub Form_Resize()
   spGen.Move 0, cbLibs.Height + 50, ScaleWidth, ScaleHeight - (cbLibs.Height + sbMain.Height + 100)
End Sub


Private Sub lvLibs_GotFocus()
   tbLibs.Buttons(1).Enabled = True
   mnuFilNew.Enabled = True
   If lvLibs.ListItems.Count > 0 Then
      mnuFilRemove.Enabled = True
      mnuFilModify.Enabled = True
      tbLibs.Buttons(2).Enabled = True
      tbLibs.Buttons(3).Enabled = True
   Else
      mnuFilRemove.Enabled = False
      mnuFilModify.Enabled = False
      tbLibs.Buttons(2).Enabled = False
      tbLibs.Buttons(3).Enabled = False
   End If
End Sub

Private Sub lvLibs_ItemClick(ByVal Item As MSComctlLib.ListItem)
   If tvLibs.SelectedItem Is Nothing Then Exit Sub
   If tvLibs.SelectedItem.Tag = "lib" Then
      tbLibs.Buttons(5).Enabled = True
      mnuFilPrint.Enabled = True
      Call DisplayScript(1)
   Else
      tbLibs.Buttons(5).Enabled = False
      mnuFilPrint.Enabled = False
   End If
End Sub

Private Sub lvLibs_LostFocus()
   mnuFilNew.Enabled = False
   mnuFilRemove.Enabled = False
   mnuFilModify.Enabled = False
   mnuFilPrint.Enabled = False
   
   tbLibs.Buttons(1).Enabled = False
   tbLibs.Buttons(2).Enabled = False
   tbLibs.Buttons(3).Enabled = False
   tbLibs.Buttons(5).Enabled = False
End Sub

Private Sub mnuEdiCopy_Click()
   Clipboard.SetText txtCodeScript.SelText
End Sub

Private Sub mnuEdiSelAll_Click()
   With txtCodeScript
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub

Private Sub mnuExit_Click()
   Unload Me
End Sub


Private Sub ConfigControls()
   FlatBorder tvLibs.hwnd
   FlatBorder lvLibs.hwnd
   
   With spGen
      Set .Control1 = tvLibs
      Set .Control2 = spProps
   End With
   
   With spProps
      Set .Control1 = lvLibs
      Set .Control2 = txtCodeScript
   End With

End Sub

Private Sub ConfigStatusBar()
   With sbMain
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
End Sub


'Carge Languajes
' 1) Chage in TreeView
' 2) Charge in ListView
Private Sub ChargeLanguajes(ByVal nMode As Integer)
Dim RsAMC As New ADODB.Recordset
Dim xNode As Node, iCount As Integer
   
   Screen.MousePointer = vbHourglass
   
   sSQLcmd = "select * from Languaje Order by nLanguajeID "
   With RsAMC
      .CursorLocation = adUseClient
      .Open sSQLcmd, sysDB, adOpenForwardOnly, adLockReadOnly
      If .RecordCount > 0 Then
         For iCount = 0 To .RecordCount - 1
            If nMode = 1 Then
               Set xNode = tvLibs.Nodes.Add("all", tvwChild, "Lan" & RsAMC("nLanguajeID"), Trim(RsAMC("cName")), Image:=3, SelectedImage:=3)
               xNode.Tag = "lang"
               xNode.Expanded = True
            Else
               With lvLibs.ListItems
                  .Add , , RsAMC(0), Icon:=3, SmallIcon:=3
                  .Item(iCount + 1).SubItems(1) = RsAMC(1)
                  .Item(iCount + 1).SubItems(2) = Trim(RsAMC(2))
                  .Item(iCount + 1).SubItems(3) = Trim(RsAMC(3))
                  .Item(iCount + 1).SubItems(4) = Trim(RsAMC(4))
                  .Item(iCount + 1).SubItems(5) = Trim(RsAMC(5))
                  .Item(iCount + 1).Tag = tvLibs.Nodes("Lan" & RsAMC("nLanguajeID")).Index
               End With
            End If
            .MoveNext
         Next iCount
         If nMode = 1 Then
            Call ChargeLibraries(1)
         End If
      End If
      .Close
   End With
   Set RsAMC = Nothing
   Screen.MousePointer = vbDefault
End Sub

'Carge libraries
' 1) Chage in TreeView
' 2) Charge in ListView
Public Sub ChargeLibraries(ByVal nMode As Integer)
Dim RsAMC As New ADODB.Recordset
Dim xNode As Node, iCount As Integer

   Screen.MousePointer = vbHourglass
   If nMode = 1 Then
      sSQLcmd = "select * from Libraries Order by nLibraryID "
   Else
      sSQLcmd = "select * from Libraries where nLanguajeID = " & Mid(tvLibs.SelectedItem.Key, 4, Len(tvLibs.SelectedItem.Key)) & " Order by nLibraryID "
   End If
   
   With RsAMC
      .CursorLocation = adUseClient
      .Open sSQLcmd, sysDB, adOpenForwardOnly, adLockReadOnly
      If .RecordCount > 0 Then
         For iCount = 0 To .RecordCount - 1
            If nMode = 1 Then
               Set xNode = tvLibs.Nodes.Add("Lan" & RsAMC("nLanguajeID"), tvwChild, "Lib" & RsAMC("nLibraryID"), Trim(RsAMC("cName")), Image:=5, SelectedImage:=5)
               xNode.Tag = "lib"
               xNode.Expanded = True
            Else
               With lvLibs.ListItems
                  .Add , , RsAMC(0), Icon:=5, SmallIcon:=5
                  .Item(iCount + 1).SubItems(1) = RsAMC(1)
                  .Item(iCount + 1).SubItems(2) = Trim(RsAMC(2))
                  .Item(iCount + 1).SubItems(3) = Trim(RsAMC(3))
                  .Item(iCount + 1).SubItems(4) = Format(RsAMC(4), "mm/dd/yyyy")
                  .Item(iCount + 1).SubItems(5) = Format(RsAMC(5), "mm/dd/yyyy")
                  .Item(iCount + 1).SubItems(6) = Trim(RsAMC(6))
                  .Item(iCount + 1).Tag = tvLibs.Nodes("Lib" & RsAMC("nLibraryID")).Index
               End With
            End If
            .MoveNext
         Next iCount
      End If
      .Close
   End With
   Set RsAMC = Nothing
   Screen.MousePointer = vbDefault
End Sub


'Charge Templates
Private Sub ChargeTemplates()
Dim RsAMC As New ADODB.Recordset
Dim xNode As Node, iCount As Integer
   
   Screen.MousePointer = vbHourglass
   sSQLcmd = "select * from Scripts where nLibraryID = " & Mid(tvLibs.SelectedItem.Key, 4, Len(tvLibs.SelectedItem.Key))
   With RsAMC
      .CursorLocation = adUseClient
      .Open sSQLcmd, sysDB, adOpenForwardOnly, adLockReadOnly
      If .RecordCount > 0 Then
         For iCount = 0 To .RecordCount - 1
            With lvLibs.ListItems
               .Add , , RsAMC(0), Icon:=9, SmallIcon:=9
               .Item(iCount + 1).SubItems(1) = RsAMC(1)
               .Item(iCount + 1).SubItems(2) = Trim(RsAMC(2))
               .Item(iCount + 1).SubItems(3) = Trim(RsAMC(4))
               .Item(iCount + 1).Tag = tvLibs.Nodes("Lib" & RsAMC.Fields("nLibraryID")).Index
            End With
            .MoveNext
         Next iCount
      End If
      .Close
   End With
   Set RsAMC = Nothing
   Screen.MousePointer = vbDefault
End Sub



'Procedure for display code script
Private Sub DisplayScript(ByVal IDScript As Long)
Dim RsAMC As New ADODB.Recordset

   Screen.MousePointer = vbHourglass
   sSQLcmd = "select cCodeScript from Scripts where nScriptID = " & lvLibs.SelectedItem.Text
   
   With RsAMC
      .CursorLocation = adUseClient
      .Open sSQLcmd, sysDB, adOpenForwardOnly, adLockReadOnly
      If .RecordCount > 0 Then
         txtCodeScript.Text = RsAMC(0)
         txtCodeScript.Tag = ""
      End If
      .Close
   End With
   
   Set RsAMC = Nothing
   Screen.MousePointer = vbDefault
End Sub




'Configured ListView
' 1) Configured for Libraries
' 2) Configured for Templates
Private Sub ConfigLV(ByVal nMode As Integer)
   
   txtCodeScript.Text = ""
   With lvLibs.ColumnHeaders
      Select Case nMode
         Case 1
            If spProps.SplitPercent < 100 Then spProps.SplitPercent = 100
            If lvLibs.ListItems.Count >= 1 Then lvLibs.ListItems.Clear
            If .Count <> 6 Then
               .Clear
               .Add , , "ID Languaje", 1200
               .Add , , "ID Type Languaje", 0
               .Add , , "Name", 2500
               .Add , , "Prefix", 1000
               .Add , , "Comments", 2500
               .Add , , "Status", 800
            End If
         Case 2
            If spProps.SplitPercent < 100 Then spProps.SplitPercent = 100
            If lvLibs.ListItems.Count >= 1 Then lvLibs.ListItems.Clear
            If .Count < 7 Then
               .Clear
               .Add , , "ID Library", 1000
               .Add , , "ID Languaje", 0
               .Add , , "Name", 2500
               .Add , , "Author", 2500
               .Add , , "Create Date", 1200
               .Add , , "Create Modify", 1200
               .Add , , "Status", 800
            End If
         Case 3
            spProps.SplitPercent = 50
            If lvLibs.ListItems.Count >= 1 Then lvLibs.ListItems.Clear
            If .Count <> 4 Then
               .Clear
               .Add , , "ID Script", 1000
               .Add , , "ID Library", 0
               .Add , , "Name", 5000
               .Add , , "Status", 800
            End If
      End Select
   End With

End Sub



Private Sub mnuFilModify_Click()
   If lvLibs.Tag = "lang" Then
      Load frmCap
      Call frmCap.ModifyLanguaje(lvLibs.SelectedItem.Index, lvLibs.SelectedItem.Tag)
      frmCap.Show 1
   ElseIf lvLibs.Tag = "libs" Then
      Load frmCap
      Call frmCap.ModifyLibrary(lvLibs.SelectedItem.Index, lvLibs.SelectedItem.Tag)
      frmCap.Show 1
   ElseIf lvLibs.Tag = "scripts" Then
      Load frmEditor
      Call frmEditor.ModifyScript(lvLibs.SelectedItem.Index, tvLibs.SelectedItem.Text)
      frmEditor.Show 1
   End If
End Sub



Private Sub mnuFilNew_Click()
   If lvLibs.Tag = "lang" Then
      Load frmCap
      frmCap.NewLanguaje
      frmCap.Show 1
   ElseIf lvLibs.Tag = "libs" Then
      Load frmCap
      frmCap.NewLibrary
      frmCap.Show 1
   ElseIf lvLibs.Tag = "scripts" Then
      Load frmEditor
      frmEditor.NewScript Mid(tvLibs.SelectedItem.Key, 4, Len(tvLibs.SelectedItem.Key)), tvLibs.SelectedItem.Text
      frmEditor.Show 1
   End If
End Sub

Private Sub mnuFilPrint_Click()
   MsgBox "Print Scritp", vbInformation, App.Title
End Sub

Private Sub mnuFilRemove_Click()
Dim RsAMC As New ADODB.Recordset
Dim nRecordsAffected As Integer
Dim sMsg As String, sSQL As String
   
   If lvLibs.ListItems.Count <= 0 Then
      MsgBox "Not Records exits", vbInformation, App.Title: Exit Sub
   End If
   
   'If lvLibs.SelectedItem Is Nothing <= 0 Then
   '   MsgBox "Not Record selected", vbInformation, App.Title: Exit Sub
   'End If
   
   Select Case LCase(lvLibs.Tag)
      Case "lang"
         sSQLcmd = "SELECT * from Libraries WHERE nLanguajeID = " & lvLibs.SelectedItem.Text
         sSQL = "DELETE FROM Languaje where nLanguajeID = " & lvLibs.SelectedItem.Text
         sMsg = "Not it's possibled deleted record be caused contents register libraries"
      Case "libs"
         sSQLcmd = "SELECT * from Scripts WHERE nLibraryID = " & lvLibs.SelectedItem.Text
         sMsg = "Not it's possibled deleted record be caused contents register scripts"
         sSQL = "DELETE FROM Libraries where nLibraryID = " & lvLibs.SelectedItem.Text
      Case "scripts"
         sSQL = "DELETE FROM SCRIPTS where nScriptID = " & lvLibs.SelectedItem.Text
   End Select
      
   With RsAMC
      If LCase(lvLibs.Tag) = "lang" Or LCase(lvLibs.Tag) = "libs" Then
         .CursorLocation = adUseClient
         .Open sSQLcmd, sysDB, adOpenForwardOnly, adLockReadOnly
         If .RecordCount <= 0 Then
            .Close
         Else
            MsgBox sMsg, vbInformation, App.Title: .Close: Exit Sub
         End If
         
      End If
   End With
   
   sysDB.Execute sSQL, nRecordsAffected
   
   If nRecordsAffected <= 0 Then
      MsgBox "Record Not as deleted", vbInformation, App.Title
   Else
      lvLibs.ListItems.Remove lvLibs.SelectedItem.Index
      MsgBox "Record as been deleted", vbInformation, App.Title
   End If
   
   
End Sub

Private Sub tbLibs_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Index
      Case 1
         Call mnuFilNew_Click
      Case 2
         Call mnuFilRemove_Click
      Case 3
         Call mnuFilModify_Click
      Case 5
         Call mnuFilPrint_Click
      Case 7
         Call mnuExit_Click
   End Select
End Sub

Private Sub tvLibs_NodeClick(ByVal Node As MSComctlLib.Node)
   Select Case LCase(Node.Tag)
      Case "all"
         Call ConfigLV(1)
         Call ChargeLanguajes(2)
         lvLibs.Tag = "lang"
      Case "lang"
         Call ConfigLV(2)
         Call ChargeLibraries(2)
         lvLibs.Tag = "libs"
      Case "lib"
         Call ConfigLV(3)
         Call ChargeTemplates
         lvLibs.Tag = "scripts"
   End Select
End Sub

Private Sub txtCodeScript_Change()
   txtCodeScript.Tag = "change"
   
End Sub

Private Sub txtCodeScript_GotFocus()
   mnuEdiCopy.Enabled = True
   mnuEdiSelAll.Enabled = True
   tbEdit.Buttons(2).Enabled = True
End Sub

Private Sub txtCodeScript_LostFocus()
   mnuEdiCopy.Enabled = False
   mnuEdiSelAll.Enabled = False
   tbEdit.Buttons(2).Enabled = False
End Sub

