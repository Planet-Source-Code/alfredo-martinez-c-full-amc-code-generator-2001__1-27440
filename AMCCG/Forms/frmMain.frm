VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "AMC - Code Generator 2001"
   ClientHeight    =   5715
   ClientLeft      =   1515
   ClientTop       =   1875
   ClientWidth     =   9330
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "frmMain"
   ScaleHeight     =   5715
   ScaleWidth      =   9330
   WindowState     =   2  'Maximized
   Begin GenObj.HighlightSintax txtInternalCode 
      Height          =   795
      Left            =   3360
      TabIndex        =   18
      Top             =   4380
      Visible         =   0   'False
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
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GenObj.HighlightSintax txtEdit 
      Height          =   1575
      Left            =   4920
      TabIndex        =   17
      Top             =   600
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   2778
      Language        =   5
      KeywordColor    =   12582912
      OperatorColor   =   12583104
      DelimiterColor  =   8421504
      ForeColor       =   0
      FunctionColor   =   8632256
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
   Begin VB.TextBox txtBuild 
      Height          =   855
      Left            =   5220
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   4380
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSComctlLib.ListView lvFields 
      Height          =   855
      Left            =   1500
      TabIndex        =   11
      Top             =   4380
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1508
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      Icons           =   "imgBDTmp"
      SmallIcons      =   "imgBDTmp"
      ColHdrIcons     =   "imgBDTmp"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComDlg.CommonDialog cdlgScripts 
      Left            =   8100
      Top             =   2700
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlMain 
      Left            =   8220
      Top             =   4020
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0ABA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C16
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":33CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5B7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8332
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AAE6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlTools 
      Left            =   8100
      Top             =   1020
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AE02
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AF5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B27A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B6D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BB2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BE46
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin GenObj.Split spDBObjs 
      Height          =   1575
      Left            =   4860
      TabIndex        =   8
      Top             =   2340
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   2778
      HorizontalSplit =   -1  'True
      SplitPercent    =   100
      Begin MSComctlLib.TabStrip tsProps 
         Height          =   375
         Left            =   60
         TabIndex        =   12
         Top             =   1080
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   661
         Style           =   2
         HotTracking     =   -1  'True
         Separators      =   -1  'True
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Fields"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Internal Code"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvObjsView 
         Height          =   975
         Left            =   60
         TabIndex        =   9
         Top             =   60
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1720
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "imgBDTmp"
         SmallIcons      =   "imgBDTmp"
         ColHdrIcons     =   "imgBDTmp"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin ComCtl3.CoolBar cbMain 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   688
      _CBWidth        =   9330
      _CBHeight       =   390
      _Version        =   "6.7.8862"
      Child1          =   "tbMain"
      MinHeight1      =   330
      Width1          =   3195
      NewRow1         =   0   'False
      Child2          =   "tbTools"
      MinHeight2      =   330
      Width2          =   2745
      NewRow2         =   0   'False
      Child3          =   "tbEdit"
      MinHeight3      =   330
      Width3          =   2355
      NewRow3         =   0   'False
      Begin MSComctlLib.Toolbar tbTools 
         Height          =   330
         Left            =   3390
         TabIndex        =   10
         Top             =   30
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imlTools"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.ToolTipText     =   "Display Table/View/Procedures Properties"
               ImageIndex      =   1
               Style           =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.ToolTipText     =   "Display Table Objects"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Generate Code"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.ToolTipText     =   "Generate Code Wizard"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.ToolTipText     =   "Generate Visual Basic Form"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.ToolTipText     =   "Generate Complete Visual Basic Proyect"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbMain 
         Height          =   330
         Left            =   165
         TabIndex        =   7
         Top             =   30
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imlMain"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "New Workspace"
               ImageIndex      =   1
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "New Workspace"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "New Template"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Open Workspace"
               ImageIndex      =   2
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Open Workspace"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Open Existing Template"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.ToolTipText     =   "Close Template"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Add New Connection"
               ImageIndex      =   4
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Microsoft SQL Server Connection"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Microsoft Access Database"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "ODBC Database Connection"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.ToolTipText     =   "Close Existing Connection"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Exit"
               ImageIndex      =   6
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ListView lvTemplates 
      Height          =   975
      Left            =   6240
      TabIndex        =   5
      Top             =   4080
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1720
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imgBDTmp"
      SmallIcons      =   "imgBDTmp"
      ColHdrIcons     =   "imgBDTmp"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "AMC-Code GeneratorTemplates"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
   End
   Begin GenObj.Split spMain 
      Height          =   3855
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   6800
      SplitPercent    =   25
      Begin GenObj.Split spTree 
         Height          =   3615
         Left            =   60
         TabIndex        =   13
         Top             =   60
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   6376
         HorizontalSplit =   -1  'True
         SplitPercent    =   70
         Begin MSComctlLib.TreeView tvLibs 
            Height          =   1695
            Left            =   60
            TabIndex        =   15
            Top             =   1800
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   2990
            _Version        =   393217
            Indentation     =   706
            LabelEdit       =   1
            LineStyle       =   1
            PathSeparator   =   "/"
            Style           =   7
            FullRowSelect   =   -1  'True
            ImageList       =   "imgBDTmp"
            Appearance      =   1
         End
         Begin MSComctlLib.TreeView tvDBs 
            Height          =   1695
            Left            =   60
            TabIndex        =   14
            Top             =   60
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   2990
            _Version        =   393217
            Indentation     =   706
            LabelEdit       =   1
            LineStyle       =   1
            PathSeparator   =   "/"
            Style           =   7
            FullRowSelect   =   -1  'True
            ImageList       =   "imgBDTmp"
            Appearance      =   1
         End
      End
      Begin GenObj.Split spView 
         Height          =   3435
         Left            =   2100
         TabIndex        =   2
         Top             =   0
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   6059
         HorizontalSplit =   -1  'True
         SplitPercent    =   70
         Begin MSComctlLib.TabStrip tsFiles 
            Height          =   1215
            Left            =   420
            TabIndex        =   4
            Top             =   120
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   2143
            HotTracking     =   -1  'True
            Placement       =   1
            Separators      =   -1  'True
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   1
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Templates"
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.TabStrip tsObjects 
            Height          =   3315
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   5847
            MultiRow        =   -1  'True
            HotTracking     =   -1  'True
            Placement       =   2
            Separators      =   -1  'True
            TabStyle        =   1
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   2
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Database Objects"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Source Code Script"
                  ImageVarType    =   2
               EndProperty
            EndProperty
            Enabled         =   0   'False
         End
      End
   End
   Begin MSComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   5415
      Width           =   9330
      _ExtentX        =   16457
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
            TextSave        =   "11:35 a.m."
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
   Begin MSComDlg.CommonDialog cdlgGen 
      Left            =   720
      Top             =   4380
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgBDTmp 
      Left            =   120
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C162
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C4B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C80A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CB5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CEB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D206
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D65A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D7B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D912
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DA6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DBCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DD26
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DE82
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E19E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EA7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EBD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EEF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F20E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F52A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFilAdd 
         Caption         =   "&Add"
         Begin VB.Menu mnuFilAddSQLServer 
            Caption         =   "Microsoft &SQL Server Connection"
         End
         Begin VB.Menu mnuFilAddDBAcces 
            Caption         =   "Microsoft &Acces Database"
         End
         Begin VB.Menu FilAddS1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFilAddNewTemplateLibrary 
            Caption         =   "New Script &Library"
         End
      End
      Begin VB.Menu mnuFilRemove 
         Caption         =   "Re&move"
         Enabled         =   0   'False
      End
      Begin VB.Menu filS3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFilPrintSetup 
         Caption         =   "Print Set&up..."
      End
      Begin VB.Menu filS5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolBars 
         Caption         =   "Toolbars"
         Begin VB.Menu mnuViewStandar 
            Caption         =   "Standar"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewTbTools 
            Caption         =   "Tools"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu ViewS5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewWks 
         Caption         =   "&Workspace Explorer"
         Checked         =   -1  'True
         Shortcut        =   ^R
      End
      Begin VB.Menu ViewS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewProps 
         Caption         =   "&Properties"
         Enabled         =   0   'False
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuViewTblTrigg 
         Caption         =   "&Table Triggers"
         Enabled         =   0   'False
         Shortcut        =   {F8}
      End
      Begin VB.Menu ViewS2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefreshDB 
         Caption         =   "Refresh &Database"
         Enabled         =   0   'False
         Shortcut        =   {F6}
      End
      Begin VB.Menu ViewS4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Statusbar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolGenCode 
         Caption         =   "&Generate Code"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuGenCodeWiz 
         Caption         =   "Generate Code &Wizard"
         Enabled         =   0   'False
      End
      Begin VB.Menu ToolS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolGenCompProj 
         Caption         =   "Generate Complete Project"
         Enabled         =   0   'False
         Shortcut        =   {F7}
      End
      Begin VB.Menu ToolS2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDataTypesTraslations 
         Caption         =   "Data Types Traslations..."
         Enabled         =   0   'False
      End
      Begin VB.Menu ToolS3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolOptions 
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
         Caption         =   "About AMC-Code Generator"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents oCG As CSCGen
Attribute oCG.VB_VarHelpID = -1
Private oPBarOut As CProgressBar


Private Sub cbMain_HeightChanged(ByVal NewHeight As Single)
   Call Form_Resize
End Sub

Private Sub Form_Load()
   Call ConfigTVGen
   Call ConfigTVLibs
   Call ConfigStatusBar
   Call ConfigSplit
   Call ConfigControls
   
End Sub


Public Sub ConfigTVGen()
Dim xNode As Node
   Set xNode = tvDBs.Nodes.Add(, , "WS", "Workspace", Image:=17, SelectedImage:=17)
   xNode.Tag = "WS"
   xNode.Expanded = True
   Set xNode = tvDBs.Nodes.Add("WS", tvwChild, "databases", "SQL Server Databases", Image:=7, SelectedImage:=7)
   xNode.Tag = "databases"
   xNode.Expanded = True
   Set xNode = tvDBs.Nodes.Add("WS", tvwChild, "vbforms", "Visual Basic Forms", Image:=7, SelectedImage:=7)
   xNode.Tag = "vbforms"
   xNode.Expanded = True
   Set xNode = tvDBs.Nodes.Add("WS", tvwChild, "vbclass", "Visual Basic Class Modules", Image:=7, SelectedImage:=7)
   xNode.Tag = "vbclass"
   xNode.Expanded = True
   Set xNode = tvDBs.Nodes.Add("WS", tvwChild, "vbmodules", "Visual Basic Modules", Image:=7, SelectedImage:=7)
   xNode.Tag = "vbmodules"
   xNode.Expanded = True
   
   Set xNode = tvDBs.Nodes.Add("WS", tvwChild, "sqlfiles", "SQL Files", Image:=7, SelectedImage:=7)
   xNode.Tag = "sqlfiles"
   xNode.Expanded = True
End Sub


Public Sub ConfigTVLibs()
Dim xNode As Node
   Set xNode = tvLibs.Nodes.Add(, , "sqlserver", "SQL Server Code Libraries", Image:=7, SelectedImage:=7)
   xNode.Tag = "sqlserver"
   xNode.Expanded = True
   Set xNode = tvLibs.Nodes.Add(, , "vbcodelibs", "Visual Basic Code Libraries", Image:=7, SelectedImage:=7)
   xNode.Tag = "vbcodelibs"
   xNode.Expanded = True
   Set xNode = tvLibs.Nodes.Add(, , "vbformslibs", "Visual Basic Code Forms Libraries", Image:=7, SelectedImage:=7)
   xNode.Tag = "vbformslibs"
   xNode.Expanded = True

End Sub


Public Sub ConfigControls()
   FlatBorder tvDBs.hwnd
   FlatBorder tvLibs.hwnd
   FlatBorder tsFiles.hwnd
   FlatBorder tsObjects.hwnd
   FlatBorder lvObjsView.hwnd
   FlatBorder lvTemplates.hwnd
   FlatBorder lvFields.hwnd
   FlatBorder tsProps.hwnd
   'FlatBorder txtInternalCode.hwnd
   FlatBorder txtBuild.hwnd
   
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


Private Sub ConfigSplit()
   With spView
      Set .Control1 = tsObjects
      Set .Control2 = tsFiles
   End With
   
   With spTree
      Set .Control1 = tvDBs
      Set .Control2 = tvLibs
   End With
   
   With spMain
      Set .Control1 = spTree
      Set .Control2 = spView
   End With
   With spDBObjs
      Set .Control1 = lvObjsView
      Set .Control2 = tsProps
   End With
   
End Sub


Private Sub Form_Resize()
Dim lHeighTB As Long, lHeighSB As Long
On Error Resume Next
   If cbMain.Visible Then lHeighTB = cbMain.Height + 50
   If sbMain.Visible Then lHeighSB = sbMain.Height + 25
   spMain.Move 25, lHeighTB, ScaleWidth - 25, ScaleHeight - (lHeighTB + lHeighSB)
End Sub



Private Sub lvObjsView_GotFocus()
   mnuFilRemove.Enabled = False
End Sub

Private Sub lvObjsView_ItemClick(ByVal Item As MSComctlLib.ListItem)
   Screen.MousePointer = vbHourglass
   If tbTools.Buttons(1).Value = tbrPressed Then
      Select Case UCase(Item.Tag)
         Case "TABLE"
            txtInternalCode.Text = "!! --   Not' Internal Code Display   -- !!"
            Call oCnnTmp.DisplayFields(tvDBs.SelectedItem.Parent.Key, Item.Text, lvFields, val(Mid(tvDBs.SelectedItem.Parent.Tag, 1, 1)), Table, sbMain)
         Case "VIEW"
            Call oCnnTmp.DisplayFields(tvDBs.SelectedItem.Parent.Key, Item.Text, lvFields, val(Mid(tvDBs.SelectedItem.Parent.Tag, 1, 1)), View, sbMain)
            If val(Mid(tvDBs.SelectedItem.Parent.Tag, 1, 1)) = [SQL Server] Then
               txtInternalCode.Text = oCnnTmp.DisplayInternalCodeSQLServer(tvDBs.SelectedItem.Parent.Key, lvObjsView.SelectedItem.Text, sbMain)
            Else
               txtInternalCode.Text = oCnnTmp.DisplayInternalCodeAccessODBC(tvDBs.SelectedItem.Parent.Key, lvObjsView.SelectedItem.Text)
            End If
         Case "PROCEDURE"
            If val(Mid(tvDBs.SelectedItem.Parent.Tag, 1, 1)) = [SQL Server] Then
               Call oCnnTmp.DisplayParamsSP(tvDBs.SelectedItem.Parent.Key, lvObjsView.SelectedItem.Text, lvFields, sbMain)
               txtInternalCode.Text = oCnnTmp.DisplayInternalCodeSQLServer(tvDBs.SelectedItem.Parent.Key, lvObjsView.SelectedItem.Text, sbMain)
            Else
               txtInternalCode.Text = "!! --   Not' Internal Code Display   -- !!"
            End If
      End Select
   End If
   Screen.MousePointer = vbDefault
End Sub


Private Sub lvObjsView_LostFocus()
   mnuFilRemove.Enabled = False
End Sub

Private Sub lvTemplates_DblClick()
Dim RsAMC As New ADODB.Recordset
   
   If lvTemplates.ListItems.Count <= 0 Then Exit Sub
   
   Screen.MousePointer = vbHourglass
   With RsAMC
      .CursorLocation = adUseClient
      sSQLcmd = "select cCodeScript from Scripts where nScriptID = " & lvTemplates.SelectedItem.Text
      .Open sSQLcmd, sysDB, adOpenForwardOnly, adLockReadOnly
      If .Fields.Count > 0 Then
         If Not IsNull(.Fields(0)) Then txtEdit.Text = .Fields(0)
         tsObjects.Enabled = True
      Else
         txtEdit.Text = ""
         tsObjects.Enabled = False
      End If
      .Close
   End With
   Set RsAMC = Nothing
   Screen.MousePointer = vbDefault
   
End Sub

Private Sub lvTemplates_GotFocus()
   If lvTemplates.ListItems.Count <= 0 Then
      mnuFilRemove.Enabled = False
      tbMain.Buttons(3).Enabled = False
   Else
      mnuFilRemove.Enabled = True
      tbMain.Buttons(3).Enabled = True
   End If
End Sub

Private Sub lvTemplates_LostFocus()
   mnuFilRemove.Enabled = False
   tbMain.Buttons(3).Enabled = False
End Sub

Private Sub mnuFilAddDBAcces_Click()
Dim sConnectionString As String
Dim db As New ADODB.Connection
Dim lNodeCount As Long, bNodeExist As Boolean
On Error GoTo ErrormnuFileAddOpenAccessDB_Click
   With cdlgGen
      .DialogTitle = "Open Microsoft Access Database"
      .Filter = "Microsoft Access Database (*.mdb)|*.mdb"
      .ShowOpen
      DoEvents
      If Trim(.FileName) <> "" Then
         DoEvents
         oCnnTmp.InfoClear
         oCnnTmp.DisplayInfo "Open connection", "Open Microsoft Access Database [" & .FileName & "] "
         sConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & .FileName
         bNodeExist = False
         With tvDBs.Nodes
            DoEvents
            For lNodeCount = 1 To .Count
               If UCase(.Item(lNodeCount).Key) = UCase("ACCS" & cdlgGen.FileName) Then
                  bNodeExist = True
                  oCnnTmp.DisplayInfo "Open connection", "Connection with Microsoft Access Database already exists..."
                  Exit For
               End If
            Next lNodeCount
         End With
         If bNodeExist = False Then
            DoEvents
            oCnnTmp.DisplayInfo "Open connection", "Establish connection..."
            db.Open sConnectionString
            oDBs.Add db, "ACCS" & .FileName
            oCnnTmp.DisplayInfo "Open connection", "Connection completed successfully..."
            Call oCnnTmp.DisplayConnection("ACCS" & .FileName, tvDBs, UCase(.FileTitle), [Microsoft Access])
         Else
            MsgBox "Connection already exists...", vbExclamation, "AMC-CodeAssist"
         End If
         oCnnTmp.DisplayInfo "Open connection", " - (0) error(s),  (0) warning(s) "
      End If
      .FileName = ""
   End With
   Call tvDBs_GotFocus
   tvDBs.SetFocus
   Screen.MousePointer = vbDefault
Exit Sub
ErrormnuFileAddOpenAccessDB_Click:
   Screen.MousePointer = vbDefault
   oMPG.DisplayError 0, Err, Error, "Connection..."
   oCnnTmp.DisplayInfo "Open connection", "Error [" & Err & "] " & Error
   oCnnTmp.DisplayInfo "Open connection", " - (1) error(s),  (1) warning(s) "
   Set db = Nothing
End Sub



Private Sub mnuFilAddNewTemplateLibrary_Click()
   Load frmLibs
   frmLibs.Show 1, Me
End Sub



Private Sub mnuFilAddSQLServer_Click()
   Me.MousePointer = vbHourglass
   Load frmConnSQLServer
   frmConnSQLServer.Caption = "Microsoft SQL Server Database connection"
   frmConnSQLServer.Show 1
   Me.MousePointer = vbDefault
End Sub



Private Sub mnuFilExit_Click()
   End
End Sub

Private Sub mnuFilNewWks_Click()
   MsgBox "New Workspace", vbInformation, App.Title
End Sub

Private Sub mnuFilOpenWks_Click()
   MsgBox "Open Workspace", vbInformation, App.Title
End Sub

Private Sub mnuFilRemove_Click()
Dim iCount As Integer
   Select Case Me.ActiveControl.Name
      Case "tvDBs"
         If tvDBs.Nodes.Count <= 0 Then Exit Sub
         If tvDBs.SelectedItem Is Nothing Then
            MsgBox "No selected connection to removing", vbInformation, App.Title
            Exit Sub
         End If
         Select Case UCase(tvDBs.SelectedItem.Tag)
            Case "TABLES", "VIEWS", "PROCEDURES"
               oDBs.Remove (tvDBs.SelectedItem.Parent.Key)
               tvDBs.Nodes.Remove (tvDBs.SelectedItem.Parent.Index)
               lvObjsView.ListItems.Clear
               lvFields.ListItems.Clear
               txtEdit.Text = ""
            Case 0, 1, 2, 3
               oDBs.Remove (tvDBs.SelectedItem.Key)
               tvDBs.Nodes.Remove (tvDBs.SelectedItem.Index)
               lvObjsView.ListItems.Clear
               lvFields.ListItems.Clear
               txtEdit.Text = ""
         End Select
         Call tvDBs_GotFocus
      Case "lvTemplates"
         If lvTemplates.ListItems.Count <= 0 Then Exit Sub
         For iCount = 1 To lvTemplates.ListItems.Count
            If UCase(lvTemplates.ListItems(iCount).SubItems(1)) = UCase(txtEdit.Tag) Then
               txtEdit.Text = ""
               txtEdit.Tag = ""
               Exit For
            End If
         Next iCount
         lvTemplates.ListItems.Remove (lvTemplates.SelectedItem.Index)
         If lvTemplates.ListItems.Count <= 0 Then
            tsObjects.Tabs(1).Selected = True
            tsObjects.Enabled = False
            Call lvTemplates_GotFocus
            lvTemplates.SetFocus
         End If
         
   End Select
   
End Sub




Private Sub mnuFilSave_Click()
'
End Sub

Private Sub mnuHlpAbout_Click()
   Load frmAbout
   frmAbout.Show 1
End Sub

Private Sub mnuToolGenCode_Click()
   Call GenerateSourceCode
End Sub

Private Sub mnuToolOptions_Click()
   MsgBox "Options", vbInformation, App.Title
End Sub

Private Sub mnuViewProps_Click()
   If tbTools.Buttons(1).Value = tbrUnpressed Then
      tbTools.Buttons(1).Value = tbrPressed
   Else
      tbTools.Buttons(1).Value = tbrUnpressed
   End If
   Call tbTools_ButtonClick(tbTools.Buttons(1))
End Sub

Private Sub mnuViewTblTrigg_Click()
   Call tbTools_ButtonClick(tbTools.Buttons(3))
End Sub

Private Sub mnuViewWks_Click()
   If mnuViewWks.Checked = True Then
      mnuViewWks.Checked = False
      spMain.SplitPercent = 0
   Else
      mnuViewWks.Checked = True
      spMain.SplitPercent = 25
   End If
End Sub

Private Sub oCG_Progress(ByVal nPos As Long)
   oPBarOut.Value = nPos
End Sub

Private Sub spDBObjs_Resize()
On Error Resume Next
   lvFields.Move tvDBs.Width + 425, spMain.Top + lvObjsView.Height + 450, lvObjsView.Width, tsProps.Height - 375
   txtInternalCode.Move lvFields.Left, lvFields.Top, lvFields.Width, lvFields.Height
End Sub

Private Sub spView_Resize()
On Error Resume Next
   spDBObjs.Move tvDBs.Width + 425, spMain.Top + 50, tsObjects.Width - 400, tsObjects.Height - 125
   txtEdit.Move spDBObjs.Left, spDBObjs.Top, spDBObjs.Width, spDBObjs.Height
   lvTemplates.Move tvDBs.Width + 125, tsObjects.Height + 575, tsFiles.Width - 150, tsFiles.Height - 450
   lvTemplates.ColumnHeaders(1).Width = 3500
   lvTemplates.ColumnHeaders(2).Width = 5000
   txtBuild.Move lvTemplates.Left, lvTemplates.Top, lvTemplates.Width, lvTemplates.Height
End Sub

Private Sub tbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Index
      Case 1
      Case 2
      Case 3
         Call mnuFilRemove_Click
      Case 5
         Call mnuFilAddSQLServer_Click
      Case 6
         Call mnuFilRemove_Click
      Case 8
         Call mnuFilExit_Click
         
   End Select
End Sub

Private Sub tbMain_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
   Select Case ButtonMenu.Parent.Index
      Case 2
         Select Case ButtonMenu.Index
            Case 1
         End Select
      Case 5
         Select Case ButtonMenu.Index
            Case 1
               Call mnuFilAddSQLServer_Click
            Case 2
               Call mnuFilAddDBAcces_Click
         End Select
   End Select
End Sub

Private Sub tbTools_ButtonClick(ByVal Button As MSComctlLib.Button)
   If Button.Index = 1 Then
      If tsObjects.SelectedItem.Index = 2 Then
         If spDBObjs.SplitPercent = 100 Then
            Button.Value = tbrUnpressed
         Else
            Button.Value = tbrPressed
         End If
         Exit Sub
      End If
   End If
   
   Select Case Button.Index
      Case 1
         If Button.Value = tbrUnpressed Then
            spDBObjs.SplitPercent = 100
            lvFields.Visible = False
            txtInternalCode.Visible = False
            lvFields.ListItems.Clear
            txtInternalCode.Text = ""
         Else
            spDBObjs.SplitPercent = 60
            tsProps.Tabs(1).Selected = True
         End If
      Case 5
         Call GenerateSourceCode
   End Select
End Sub


Private Sub tsFiles_Click()
   If tsFiles.SelectedItem.Index = 1 Then
      lvTemplates.Visible = True
      txtBuild.Visible = False
   Else
      txtBuild.Visible = True
      lvTemplates.Visible = False
   End If
End Sub

Private Sub tsObjects_Click()
   Select Case tsObjects.SelectedItem.Index
      Case 1
         txtEdit.Visible = False
         spDBObjs.Visible = True
         If tbTools.Buttons(1).Value = tbrPressed Then
            lvFields.Visible = True
            txtInternalCode.Visible = True
         End If
         'mnuEdit.Enabled = False
      Case 2
         txtEdit.Visible = True
         spDBObjs.Visible = False
         lvFields.Visible = False
         txtInternalCode.Visible = False
         'mnuEdit.Enabled = True
   End Select
End Sub

Private Sub tsProps_Click()
   Select Case tsProps.SelectedItem.Index
      Case 1
         lvFields.Visible = True
         txtInternalCode.Visible = False
      Case 2
         lvFields.Visible = False
         txtInternalCode.Visible = True
   End Select
End Sub

Private Sub tvDBs_Collapse(ByVal Node As MSComctlLib.Node)
   Select Case Node.Key
      Case "databases", "vbfiles", "sqlfiles": Node.SelectedImage = 7
   End Select
End Sub

Private Sub tvDBs_Expand(ByVal Node As MSComctlLib.Node)
   Select Case Node.Key
      Case "databases", "vbfiles", "sqlfiles": Node.SelectedImage = 8
   End Select
End Sub

Private Sub tvDBs_GotFocus()
   If tvDBs.Nodes.Count <= 0 Then
      mnuFilRemove.Enabled = False
      tbMain.Buttons(6).Enabled = False
   Else
      If tvDBs.Nodes.Count <= 0 Then Exit Sub
      mnuFilRemove.Enabled = True
      tbMain.Buttons(6).Enabled = True
   End If
End Sub

Private Sub tvDBs_LostFocus()
   mnuFilRemove.Enabled = False
   tbMain.Buttons(6).Enabled = False
End Sub

Private Sub tvDBs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
      PopupMenu mnuFilAdd
   End If
End Sub

Private Sub tvDBs_NodeClick(ByVal Node As MSComctlLib.Node)
   tbTools.Buttons(1).Enabled = True
   lvFields.ListItems.Clear
   Screen.MousePointer = vbHourglass
   Select Case UCase(Node.Tag)
      Case "WS", "LIBRARIES", "DATABASES", "VBCLASS", "VBMODULES", "VBFORMS", "SQLFILES"
         Call ClearLV
         tbTools.Buttons(3).Enabled = False
         tbTools.Buttons(1).Value = tbrUnpressed
         tbTools.Buttons(1).Enabled = False
         mnuViewProps.Enabled = False
         mnuViewTblTrigg.Enabled = False
         lvFields.Tag = ""
         lvObjsView.Tag = ""
      Case "TABLES"
         tbTools.Buttons(3).Enabled = True
         Call oCnnTmp.DisplayTables(Node.Parent.Key, lvObjsView, sbMain)
         mnuViewProps.Enabled = True
         mnuViewTblTrigg.Enabled = True
         lvFields.Tag = "TABLES"
         lvObjsView.Tag = "TABLES"
      Case "VIEWS"
         tbTools.Buttons(3).Enabled = False
         mnuViewTblTrigg.Enabled = False
         Call oCnnTmp.DisplayViews(Node.Parent.Key, lvObjsView, sbMain)
         lvFields.Tag = "VIEWS"
         lvObjsView.Tag = "VIEWS"
      Case "PROCEDURES"
         tbTools.Buttons(3).Enabled = False
         mnuViewTblTrigg.Enabled = False
         Call oCnnTmp.DisplayProcedures(Node.Parent.Key, lvObjsView, sbMain)
         lvFields.Tag = "PROCEDURES"
         lvObjsView.Tag = "PROCEDURES"
      Case 0, 1, 2, 3
         tbTools.Buttons(3).Enabled = False
         Call oCnnTmp.DisplayDBProperties(Node.Key, lvObjsView, sbMain)
         tbTools.Buttons(1).Value = tbrUnpressed
         Call tbTools_ButtonClick(tbTools.Buttons(1))
         tbTools.Buttons(1).Enabled = False
         lvObjsView.Tag = "DB"
   End Select
   Screen.MousePointer = vbDefault
End Sub



Private Sub ClearLV()
   lvObjsView.ListItems.Clear
   lvObjsView.ColumnHeaders.Clear
End Sub


Public Function GenerateSourceCode() As Boolean
Dim oFields As New CFields
Dim sCodeResult As String
Dim iCount As Integer, bSelect As Boolean
Dim iCountScr As Integer

On Error GoTo ErrorGenerateSourceCode
   
   bSelect = False
   With lvObjsView.ListItems
      If .Count <= 0 Then
         MsgBox "Unabled to generate code...", vbExclamation, App.Title: Exit Function
      End If
      For iCount = 1 To .Count
         If .Item(iCount).Checked = True Then
            bSelect = True: Exit For
         End If
      Next iCount
   End With
   
   If bSelect = False Then
      MsgBox "Not Tables, Views or Procedures selected...", vbExclamation, App.Title: Exit Function
   End If

   If UCase(lvObjsView.Tag) = "DB" Then
      MsgBox "Not Tables, Views or Procedures selected...", vbExclamation, App.Title: Exit Function
   End If

   bSelect = False

   With lvTemplates.ListItems
      If .Count <= 0 Then
         MsgBox "Not scripts encountered...", vbExclamation, App.Title: Exit Function
      End If
      
      For iCount = 1 To .Count
         If .Item(iCount).Checked = True Then bSelect = True: Exit For
      Next iCount
      
   End With

   If bSelect = False Then
      MsgBox "Not scripts selected...", vbExclamation, App.Title: Exit Function
   End If
   
   txtBuild.Text = "AMC Code Generator..." & vbCrLf
   txtBuild.Text = "Accesing Objects Database..." & vbCrLf
   tsFiles.Tabs.Add , , "Build"
   tsFiles.Tabs(2).Selected = True
   tsFiles.Enabled = False
   Screen.MousePointer = vbHourglass
   Load frmSourceCode
   
   For iCount = 1 To lvObjsView.ListItems.Count
      
      If lvObjsView.ListItems(iCount).Checked = True Then
         
         Select Case lvObjsView.SelectedItem.Tag
            Case 0, 1, 2, 3
               oMPG.DisplayError 1, "27365", "Unabled to Generate code...", "Code Generator"
               Exit Function
            Case "TABLE"
               Set oFields = oCnnTmp.DisplayFields(tvDBs.SelectedItem.Parent.Key, lvObjsView.ListItems(iCount).Text, lvFields, val(Mid(tvDBs.SelectedItem.Parent.Tag, 1, 1)), Table, sbMain)
            Case "VIEW"
               Set oFields = oCnnTmp.DisplayFields(tvDBs.SelectedItem.Parent.Key, lvObjsView.ListItems(iCount).Text, lvFields, val(Mid(tvDBs.SelectedItem.Parent.Tag, 1, 1)), View, sbMain)
            Case "PROCEDURE"
               Set oFields = oCnnTmp.DisplayParamsSP(tvDBs.SelectedItem.Parent.Key, lvObjsView.ListItems(iCount).Text, lvFields, sbMain)
         End Select
         
         If oFields Is Nothing Then
            oMPG.DisplayError 1, "27365", "Unabled to Generate code...", "Code Generator"
            Exit Function
         End If
         
              
         Dim dbObject As New CTable
         
         With dbObject
            .CreateDate = Format(CDate(Now), "mm/dd/yyyy")
            .Name = lvObjsView.ListItems(iCount).Text
            .Owner = lvObjsView.ListItems(iCount).SubItems(1)
            .TypeTable = lvObjsView.ListItems(iCount).SubItems(2)
            .Fields = oFields
         End With
         
         txtBuild.Text = txtBuild.Text & "Loading table ==" & dbObject.Name & "==   " & vbCrLf
         
         For iCountScr = 1 To lvTemplates.ListItems.Count
            If lvTemplates.ListItems(iCountScr).Checked = True Then
               Dim RsAMC As New ADODB.Recordset
               RsAMC.CursorLocation = adUseClient
               Set oCG = New CSCGen
               Set oPBarOut = New CProgressBar
               oPBarOut.CreateProgress sbMain, 5
               RsAMC.Open "SELECT * FROM Scripts where nScriptID = " & lvTemplates.ListItems(iCountScr).Text, sysDB, adOpenForwardOnly, adLockReadOnly
               If RsAMC.RecordCount > 0 Then
                  With oCG
                     .Script = RsAMC("cCodeScript")
                     sbMain.Panels(1).Text = "Script -> " & lvTemplates.ListItems(iCountScr).SubItems(1)
                     oPBarOut.Max = .LengthScript
                     .Author = "Alfredo Martínez C."
                     .ConnectionString = oDBs.Item(tvDBs.SelectedItem.Parent.Key).ConnectionString
                     .Database = oDBs.Item(tvDBs.SelectedItem.Parent.Key).DefaultDatabase
                     .Legal = "Copyright 2001 All Rights Reserved"
                     .Owner = "Owner"
                     .Server = "Lupe"
                     .Table = dbObject
                     txtBuild.Text = txtBuild.Text & "Generating code with script ==" & Trim(lvTemplates.ListItems(iCountScr).SubItems(1)) & "== "
                     .Generate
                     If Mid(tvLibs.SelectedItem.Parent.Key, 4, Len(tvLibs.SelectedItem.Parent.Key)) = 1 Then
                        frmSourceCode.txtEditor.Language = hlVisualBasic
                     Else
                        frmSourceCode.txtEditor.Language = [SQL Server]
                     End If
                     
                     frmSourceCode.txtEditor.Text = frmSourceCode.txtEditor.Text & .Code & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
                     txtBuild.Text = txtBuild.Text & "  Finished. " & vbCrLf
                     txtBuild.SelStart = Len(txtBuild.Text)
                     txtBuild.SelLength = 1
                     
                     .CancelTrad
                     
                  End With
                  
                  Set oCG = Nothing
                  Set oPBarOut = Nothing
                  sbMain.Panels(1).Text = "Listen..."
                  RsAMC.Close
               End If
            End If
         Next iCountScr
         
      End If
   
   Next iCount
   
   tsFiles.Tabs.Remove (2)
   tsFiles.Tabs(1).Selected = True
   tsFiles.Enabled = True
   
   Screen.MousePointer = vbDefault
   
   Screen.MousePointer = vbHourglass
   sCodeResult = ""
   Screen.MousePointer = vbDefault
   frmSourceCode.Show 1, Me
   sCodeResult = ""
   
Exit Function
ErrorGenerateSourceCode:
   oMPG.DisplayError 0, Err, Error, "Code Generator"
   sCodeResult = ""
   Resume Next
   Screen.MousePointer = vbDefault
End Function



Private Sub tvLibs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
      PopupMenu mnuFilAdd
   End If
End Sub

Private Sub tvLibs_NodeClick(ByVal Node As MSComctlLib.Node)
Dim iCount As Integer, RsAMC As New ADODB.Recordset

   lvTemplates.ListItems.Clear
   lvTemplates.ColumnHeaders.Clear
   
   Select Case LCase(Mid(Node.Key, 1, 3))
      Case "lib"
         lvTemplates.ColumnHeaders.Add , , "ID", 1000
         lvTemplates.ColumnHeaders.Add , , "Name", 3000
         lvTemplates.ColumnHeaders.Add , , "Status", 800
         With RsAMC
            .CursorLocation = adUseClient
            sSQLcmd = "SELECT * FROM Scripts WHERE nLibraryID = " & Mid(tvLibs.SelectedItem.Key, 4, Len(tvLibs.SelectedItem.Key))
            .Open sSQLcmd, sysDB, adOpenForwardOnly, adLockReadOnly
            For iCount = 1 To .RecordCount
               lvTemplates.ListItems.Add , , .Fields(0), , SmallIcon:=12
               lvTemplates.ListItems.Item(iCount).SubItems(1) = Trim(.Fields(2))
               lvTemplates.ListItems.Item(iCount).SubItems(2) = .Fields(4)
               .MoveNext
            Next iCount
            .Close
            
         End With
         txtEdit.Text = ""
         
      Case Else
         lvTemplates.ColumnHeaders.Add , , "ID", 1000
         lvTemplates.ColumnHeaders.Add , , "Name", 3000
         lvTemplates.ColumnHeaders.Add , , "Status", 800
         txtEdit.Text = ""
         tsObjects.Tabs(1).Selected = True
   End Select
   
End Sub




Public Sub LoadLanguajes()
Dim RsAMC As New ADODB.Recordset
Dim xNode As Node, iCount As Integer
   
   tvLibs.Nodes.Clear
   
   Set xNode = tvLibs.Nodes.Add(, , "all", "Libraries", Image:=16, SelectedImage:=16)
   xNode.Expanded = True
   Screen.MousePointer = vbHourglass
   
   sSQLcmd = "select * from Languaje Order by nLanguajeID "
   With RsAMC
      .CursorLocation = adUseClient
      .Open sSQLcmd, sysDB, adOpenForwardOnly, adLockReadOnly
      If .RecordCount > 0 Then
         For iCount = 0 To .RecordCount - 1
            Set xNode = tvLibs.Nodes.Add("all", tvwChild, "Lan" & RsAMC("nLanguajeID"), Trim(RsAMC("cName")), Image:=19, SelectedImage:=19)
            xNode.Tag = "lang"
            xNode.Expanded = True
            .MoveNext
         Next iCount
      End If
      .Close
   End With
   Set RsAMC = Nothing
   Screen.MousePointer = vbDefault

End Sub


Public Sub LoadLibraries()
Dim RsAMC As New ADODB.Recordset
Dim xNode As Node, iCount As Integer

   Screen.MousePointer = vbHourglass
   sSQLcmd = "select * from Libraries Order by nLibraryID "
   
   With RsAMC
      .CursorLocation = adUseClient
      .Open sSQLcmd, sysDB, adOpenForwardOnly, adLockReadOnly
      If .RecordCount > 0 Then
         For iCount = 0 To .RecordCount - 1
            Set xNode = tvLibs.Nodes.Add("Lan" & RsAMC("nLanguajeID"), tvwChild, "Lib" & RsAMC("nLibraryID"), Trim(RsAMC("cName")), Image:=13, SelectedImage:=13)
            xNode.Tag = "lib"
            xNode.Expanded = True
            .MoveNext
         Next iCount
      End If
      .Close
   End With
   Set RsAMC = Nothing
   Screen.MousePointer = vbDefault

End Sub

